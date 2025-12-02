using RestSharp;
using Newtonsoft.Json;
using System;
using System.Threading.Tasks;
using Zoho.Sheet.SDK.Core;
using Zoho.Sheet.SDK.Domain;

namespace Zoho.Sheet.SDK.Services.Auth
{
    public class ZohoAuthProvider : Core.IAuthProvider
    {
        private readonly ZohoConfig _config;
        private string _accessToken;
        private string _refreshToken;
        private DateTime _expiresAt;

        public ZohoAuthProvider(ZohoConfig config)
        {
            _config = config;
            _refreshToken = config.RefreshToken; // Load refresh token if already saved
        }

        /// <summary>
        /// Get a valid access token. Refreshes automatically if expired.
        /// </summary>
        public async Task<string> GetAccessTokenAsync()
        {
            if (!string.IsNullOrEmpty(_accessToken) && DateTime.UtcNow < _expiresAt)
                return _accessToken;

            if (string.IsNullOrEmpty(_refreshToken))
            {
                if (!string.IsNullOrEmpty(_config.DeviceCode))
                {
                    await ExchangeDeviceCodeAsync();
                }
                else
                {
                    throw new Exception("No DeviceCode or RefreshToken found. Call GetDeviceCodeAsync first.");
                }
            }
            else
            {
                await RefreshTokenAsync();
            }

            return _accessToken;
        }

        /// <summary>
        /// Request a device code from Zoho (Step 1 of Device OAuth Flow)
        /// </summary>
        public async Task<ZohoDeviceCodeResponse> GetDeviceCodeAsync()
        {
            string url = "https://accounts.zoho.com/oauth/v3/device/code";
            var client = new RestClient(url);
            var request = new RestRequest("",Method.Post);

            request.AddParameter("client_id", _config.ClientId);
            request.AddParameter("scope", "ZohoSheet.dataAPI.READ,ZohoSheet.dataAPI.UPDATE");
            request.AddParameter("grant_type", "device_request");
            request.AddParameter("access_type", "offline");

            var response = await client.ExecuteAsync(request);

            if (!response.IsSuccessful)
                throw new Exception($"Error getting device code: {response.Content}");

            dynamic json = JsonConvert.DeserializeObject(response.Content);

            return new ZohoDeviceCodeResponse
            {
                DeviceCode = json.device_code,
                UserCode = json.user_code,
                VerificationUrl = json.verification_url,
                ExpiresIn = json.expires_in,
                Interval = json.interval
            };
        }

        /// <summary>
        /// Exchange the device code for access + refresh token (Step 2)
        /// </summary>
        private async Task ExchangeDeviceCodeAsync()
        {
            var client = new RestClient(_config.DeviceTokenUrl);
            var request = new RestRequest();
            request.Method = Method.Post;

            request.AddParameter("client_id", _config.ClientId);
            request.AddParameter("client_secret", _config.ClientSecret);
            request.AddParameter("grant_type", "device_token");
            request.AddParameter("code", _config.DeviceCode);

            var response = await client.ExecuteAsync(request);
            if (!response.IsSuccessful)
                throw new Exception($"Zoho Device Token Error: {response.Content}");

            dynamic tokenResp = JsonConvert.DeserializeObject(response.Content);
            _accessToken = tokenResp.access_token;
            _refreshToken = tokenResp.refresh_token;
            _expiresAt = DateTime.UtcNow.AddSeconds((int)tokenResp.expires_in - 60);

            // Optional: persist refresh token for future use
            _config.RefreshToken = _refreshToken;
        }

        /// <summary>
        /// Refresh the access token using the stored refresh token
        /// </summary>
        private async Task RefreshTokenAsync()
        {
            var client = new RestClient(_config.TokenUrl);
            var request = new RestRequest();
            request.Method = Method.Post;

            request.AddParameter("client_id", _config.ClientId);
            request.AddParameter("client_secret", _config.ClientSecret);
            request.AddParameter("grant_type", "refresh_token");
            request.AddParameter("refresh_token", _refreshToken);

            var response = await client.ExecuteAsync(request);
            if (!response.IsSuccessful)
                throw new Exception($"Zoho Refresh Token Error: {response.Content}");

            dynamic tokenResp = JsonConvert.DeserializeObject(response.Content);
            _accessToken = tokenResp.access_token;
            _expiresAt = DateTime.UtcNow.AddSeconds((int)tokenResp.expires_in - 60);
        }

        public string GetRefreshToken() => _refreshToken;
    }

    /// <summary>
    /// Response from Zoho device code request
    /// </summary>
    public class ZohoDeviceCodeResponse
    {
        public string DeviceCode { get; set; }
        public string UserCode { get; set; }
        public string VerificationUrl { get; set; }
        public int ExpiresIn { get; set; }
        public int Interval { get; set; }
    }
}
