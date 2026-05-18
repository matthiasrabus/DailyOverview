using Azure.Identity;
using Microsoft.Graph;

namespace TeamsMessageFetcher;

public static class GraphClientFactory
{
    public static GraphServiceClient Build(AppConfig config)
    {
        // Uses interactive browser login - opens browser for Microsoft sign-in
        var options = new InteractiveBrowserCredentialOptions
        {
            TenantId = config.TenantId,
            ClientId = config.ClientId,
            RedirectUri = new Uri("http://localhost"),
        };

        var credential = new InteractiveBrowserCredential(options);
        return new GraphServiceClient(credential);
    }
}
