using System;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Identity.Client;
using Microsoft.InformationProtection;

namespace ColdFusionMIPLabeller
{
    /// <summary>
    /// Authentication delegate using MSAL client credentials flow for headless authentication.
    /// </summary>
    public class ConfidentialAuthDelegate : IAuthDelegate
    {
        private IConfidentialClientApplication? _clientApp;

        public string AcquireToken(Identity identity, string authority, string resource, string claims)
        {
            return AcquireTokenAsync(identity, authority, resource, claims).GetAwaiter().GetResult();
        }

        public async Task<string> AcquireTokenAsync(Identity identity, string authority, string resource, string claims)
        {
            try
            {
                if (_clientApp == null)
                {
                    var tenantId = Labeler.GetTenantIdInternal();
                    var clientId = Labeler.GetClientIdInternal();
                    var clientSecret = Labeler.GetClientSecretInternal();
                    
                    // Validate configuration
                    if (string.IsNullOrEmpty(tenantId))
                        throw new InvalidOperationException("Tenant ID is not configured. Call Labeler.Configure() first.");
                    if (string.IsNullOrEmpty(clientId))
                        throw new InvalidOperationException("Client ID is not configured. Call Labeler.Configure() first.");
                    if (string.IsNullOrEmpty(clientSecret))
                        throw new InvalidOperationException("Client Secret is not configured. Call Labeler.Configure() first.");
                    
                    var authorityUri = $"https://login.microsoftonline.com/{tenantId}";
                    
                    var builder = ConfidentialClientApplicationBuilder
                        .Create(clientId)
                        .WithClientSecret(clientSecret)
                        .WithAuthority(authorityUri);
                    
                    // Use system proxy if available
                    var proxy = System.Net.WebRequest.GetSystemWebProxy();
                    if (proxy != null)
                    {
                        builder.WithHttpClientFactory(new ProxyHttpClientFactory(proxy));
                    }
                    
                    _clientApp = builder.Build();
                }

                var scopes = new[] { $"{resource}/.default" };
                var result = await _clientApp.AcquireTokenForClient(scopes).ExecuteAsync();
                
                return result.AccessToken;
            }
            catch (Exception ex)
            {
                var tenantId = Labeler.GetTenantIdInternal();
                var clientId = Labeler.GetClientIdInternal();
                throw new InvalidOperationException($"Failed to acquire token for resource {resource}: TenantId='{tenantId}', ClientId='{clientId}', Error: {ex.Message}", ex);
            }
        }
    }
}