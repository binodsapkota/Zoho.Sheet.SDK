using System.Threading.Tasks;

namespace Zoho.Sheet.SDK.Core
{
    /// <summary>
    /// Provides access tokens for Zoho API calls.
    /// Handles token refresh automatically.
    /// </summary>
    public interface IAuthProvider
    {
        /// <summary>
        /// Returns a valid access token. Automatically refreshes if expired.
        /// </summary>
        /// <returns>Access token string</returns>
        Task<string> GetAccessTokenAsync();

        /// <summary>
        /// Optional: Returns the refresh token if available
        /// </summary>
        /// <returns>Refresh token string</returns>
        string GetRefreshToken();
    }
}
