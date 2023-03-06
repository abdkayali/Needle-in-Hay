using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;

namespace MAUIwithMSGRaph
{
    internal class GraphService
    {
        private readonly string[] _scopes = new[] { "User.Read" };
        private string ClientId = Environment.GetEnvironmentVariable("HackTogether2023.ClientId", EnvironmentVariableTarget.User) ?? "CLIENT_ID";
        private string TenantId = Environment.GetEnvironmentVariable("HackTogether2023.TenantId", EnvironmentVariableTarget.User) ?? "common";
        private GraphServiceClient _client;

        public GraphService()
        {
            Initialize();
        }

        private void Initialize()
        {
            // assume Windows for this sample
            if (OperatingSystem.IsWindows())
            {
                var options = new InteractiveBrowserCredentialOptions
                {
                    ClientId = ClientId,
                    TenantId = TenantId,
                    AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
                    RedirectUri = new Uri("https://login.microsoftonline.com/common/oauth2/nativeclient"),
                };

                InteractiveBrowserCredential interactiveCredential = new(options);
                _client = new GraphServiceClient(interactiveCredential, _scopes);
            }
            else
            {
                // TODO: Add iOS/Android support
            }


        }

        public async Task<User> GetMyDetailsAsync()
        {
            try
            {
                return await _client.Me.GetAsync();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading user details: {ex}");
                return null;
            }
        }
    }
}
