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
        private static AuthenticationToken Token;
        private static string TokenPath;
        private static string TenantId;
        private static string ClientId;

        public static bool Initialize(string tokenPath, string tenantId, string clientId)
        {
            try
            {
                var jsonText = File.ReadAllText(tokenPath);
                if (string.IsNullOrWhiteSpace(jsonText)) return false;
                
                Token = JsonConvert.DeserializeObject<AuthenticationToken>(jsonText);
                TokenPath = tokenPath;
                TenantId = tenantId;
                ClientId = clientId;
                
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public static async Task<string> GetValidToken()
        {
            var jwtToken = new JwtSecurityToken(Token.Access_Token);
            TimeSpan dateDiff = jwtToken.ValidTo - DateTime.UtcNow;
            if (dateDiff.TotalMinutes < 30)
            {
                if (!await RefreshToken()) return null;
            }
            return jwtToken.RawData;
        }

        private static async Task<bool> RefreshToken()
        {
            try
            {
                string refreshUrl = $"https://login.microsoftonline.com/{TenantId}/oauth2/v2.0/token";
                Dictionary<string, string> data = new Dictionary<string, string>
            {
                { "grant_type", "refresh_token" },
                { "client_id", ClientId },
                { "refresh_token", Token.Refresh_Token }
            };

                using var httpClient = new HttpClient();
                var response = await httpClient.PostAsync(refreshUrl, new FormUrlEncodedContent(data));

                var result = await response.Content.ReadAsStringAsync();

                Token = JsonConvert.DeserializeObject<AuthenticationToken>(result);

                File.WriteAllText(TokenPath, result);

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
    }
}
