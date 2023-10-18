using Coginov.GraphApi.Library.Models;
using Microsoft.Extensions.Logging;
using Microsoft.Kiota.Abstractions.Authentication;
using Newtonsoft.Json;
using System;
using System.IO;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.IdentityModel.Tokens.Jwt;

namespace Coginov.GraphApi.Library.Helpers
{
    public class TokenProvider : IAccessTokenProvider
    {
        private string _token;
        private AuthenticationConfig _authConfig;
        private AuthenticationToken _authToken;
        private ILogger _logger;

        public TokenProvider(ILogger logger, AuthenticationConfig authenticationConfig = null, AuthenticationToken authenticationToken = null, string token = null)
        {
            _token = token;
            _authConfig = authenticationConfig;
            _authToken = authenticationToken;
            _logger = logger;
        }

        public async Task<string> GetAuthorizationTokenAsync(Uri uri, Dictionary<string, object> additionalAuthenticationContext = default,
            CancellationToken cancellationToken = default)
        {
            if (!string.IsNullOrWhiteSpace(_token))
                return _token;

            if (_authConfig == null)
                return null;

            return await GetValidToken();
        }

        public AllowedHostsValidator AllowedHostsValidator { get; }

        #region Private Methods

        private async Task<string> GetValidToken(int timeInMinutes = 30)
        {
            var jwtToken = new JwtSecurityToken(_authToken.Access_Token);
            TimeSpan dateDiff = jwtToken.ValidTo - DateTime.UtcNow;
            return dateDiff.TotalMinutes < timeInMinutes ? await RefreshToken() : _authToken.Access_Token;
        }

        private async Task<string> RefreshToken()
        {
            try
            {
                string refreshUrl = $"https://login.microsoftonline.com/{_authConfig.Tenant}/oauth2/v2.0/token";
                Dictionary<string, string> data = new Dictionary<string, string>
                {
                    { "grant_type", "refresh_token" },
                    { "client_id", _authConfig.ClientId },
                    { "refresh_token", _authToken.Refresh_Token }
                };

                using var httpClient = new HttpClient();
                var response = await httpClient.PostAsync(refreshUrl, new FormUrlEncodedContent(data));

                var result = await response.Content.ReadAsStringAsync();
                _logger.LogInformation($"{response.ReasonPhrase}. {result}");
                response.EnsureSuccessStatusCode();

                var authenticationToken = JsonConvert.DeserializeObject<AuthenticationToken>(result);

                File.WriteAllText(_authConfig.TokenPath, AesHelper.EncryptToString(result));

                return authenticationToken.Access_Token;
            }
            catch (Exception ex)
            {
                _logger.LogError($"{Resource.CannotGetRefreshToken}. {ex.Message}");
                return null;
            }
        }

        public static bool IsTokenCloseToExpiration(string token, int timeInMinutes = 10)
        {
            var jwtToken = new JwtSecurityToken(token);
            TimeSpan dateDiff = jwtToken.ValidTo - DateTime.UtcNow;
            if (dateDiff.TotalMinutes < timeInMinutes)
                return true;

            return false;
        }

        #endregion

    }
}
