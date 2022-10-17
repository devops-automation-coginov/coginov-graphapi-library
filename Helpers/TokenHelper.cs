using Coginov.GraphApi.Library.Models;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IdentityModel.Tokens.Jwt;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;

namespace Coginov.GraphApi.Library.Helpers
{

    public static class TokenHelper
    {
        public static AuthenticationToken InitializeFromTokenPath(AuthenticationConfig config)
        {
            try
            {
                var jsonText = File.ReadAllText(config.TokenPath);
                if (string.IsNullOrWhiteSpace(jsonText)) return null;
                
                return JsonConvert.DeserializeObject<AuthenticationToken>(jsonText);
            }
            catch (Exception)
            {
                return null;
            }
        }

        public static async Task<AuthenticationToken> GetValidToken(AuthenticationToken token, AuthenticationConfig config, int timeInMinutes = 30)
        {
            var jwtToken = new JwtSecurityToken(token.Access_Token);
            TimeSpan dateDiff = jwtToken.ValidTo - DateTime.UtcNow;
            if (dateDiff.TotalMinutes < timeInMinutes)
            {
                token = await RefreshToken(token, config);
            }
            return token;
        }

        private static async Task<AuthenticationToken> RefreshToken(AuthenticationToken token, AuthenticationConfig config)
        {
            try
            {
                string refreshUrl = $"https://login.microsoftonline.com/{config.Tenant}/oauth2/v2.0/token";
                Dictionary<string, string> data = new Dictionary<string, string>
            {
                { "grant_type", "refresh_token" },
                { "client_id", config.ClientId },
                { "refresh_token", token.Refresh_Token }
            };

                using var httpClient = new HttpClient();
                var response = await httpClient.PostAsync(refreshUrl, new FormUrlEncodedContent(data));

                var result = await response.Content.ReadAsStringAsync();

                token = JsonConvert.DeserializeObject<AuthenticationToken>(result);

                File.WriteAllText(config.TokenPath, result);

                return token;
            }
            catch (Exception)
            {
                return null;
            }
        }

        public static bool IsTokenCloseToExpiration(string token, int timeInMinutes = 10)
        {
            var jwtToken = new JwtSecurityToken(token);
            TimeSpan dateDiff = jwtToken.ValidTo - DateTime.UtcNow;
            if (dateDiff.TotalMinutes < timeInMinutes)
            {
                return true;
            }
            return false;
        }
    }
}
