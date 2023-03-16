using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;

namespace MAUIwithMSGRaph
{
    internal class GraphService
    {
        private readonly string[] _scopes = new[] 
        { 
            "User.Read",
			"Chat.Read", // chatMessage
			"ChannelMessage.Read.All", // chatMessage
            "ChannelSettings.Read.All", // Channel Info
			"People.Read", 
            "Mail.Read", // message
            "Calendars.Read", // event
            "Files.Read.All", // driveItem
            "Sites.Read.All" // list, listItem, site

            // TODO consider adding beta entity types such as bookmark
		};
        private string ClientId = Environment.GetEnvironmentVariable("HackTogether2023.ClientId", EnvironmentVariableTarget.User) ?? string.Empty;
        private string TenantId = Environment.GetEnvironmentVariable("HackTogether2023.TenantId", EnvironmentVariableTarget.User) ?? string.Empty;
        private GraphServiceClient _client;

        public GraphServiceClient Client => _client;

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
