namespace Zoho.Sheet.SDK.Core
{
    public class ZohoApiException : Exception
    {
        public ZohoApiException(string message) : base(message) { }
        public ZohoApiException(string message, Exception inner) : base(message, inner) { }
    }
}
