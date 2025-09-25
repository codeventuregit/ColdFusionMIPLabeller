using System.Net;
using System.Net.Http;
using Microsoft.Identity.Client;

namespace ColdFusionMIPLabeller
{
    internal class ProxyHttpClientFactory : IMsalHttpClientFactory
    {
        private readonly IWebProxy _proxy;

        public ProxyHttpClientFactory(IWebProxy proxy)
        {
            _proxy = proxy;
        }

        public HttpClient GetHttpClient()
        {
            var handler = new HttpClientHandler()
            {
                Proxy = _proxy,
                UseProxy = true
            };
            return new HttpClient(handler);
        }
    }
}