using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Zoho.Sheet.SDK.Domain
{
    public class ZohoConfig
    {
        /// <summary>
        /// Your Zoho client ID
        /// </summary>
        public string ClientId { get; set; }

        /// <summary>
        /// Your Zoho client secret
        /// </summary>
        public string ClientSecret { get; set; }

        /// <summary>
        /// Device code (from Zoho Self Client / Device Token flow)
        /// </summary>
        public string DeviceCode { get; set; }

        public string AccessToken { get; set; }
        /// <summary>
        /// Refresh token (optional, populated after first run)
        /// </summary>
        public string RefreshToken { get; set; }

        /// <summary>
        /// Zoho Data Center (default: com)
        /// </summary>
        public string DataCenter { get; set; } = "com";

        /// <summary>
        /// Base API URL for Zoho Sheet
        /// </summary>
        public string BaseApiUrl => $"https://sheet.zoho.{DataCenter}/api/v2/";

        /// <summary>
        /// Token URL for refresh token calls
        /// </summary>
        public string TokenUrl => $"https://accounts.zoho.{DataCenter}/oauth/v2/token";

        /// <summary>
        /// Device token URL for exchanging device code
        /// </summary>
        public string DeviceTokenUrl => $"https://accounts.zoho.{DataCenter}/oauth/v3/device/token";

        public string DeviceCodeUrl => $"https://accounts.zoho.{DataCenter}/oauth/v3/device/code";

    }
}
