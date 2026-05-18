namespace TeamsMessageFetcher;

// ─── App Configuration ────────────────────────────────────────────────────────

public class AppConfig
{
    public string  TenantId           { get; set; } = string.Empty;
    public string  ClientId           { get; set; } = string.Empty;

    // Azure DevOps — optional; commits are skipped when either is absent
    public string? DevOpsOrganization { get; set; }
    public string? DevOpsPat          { get; set; }

    public static AppConfig Load()
    {
        // Start from environment variables
        var tenantId  = Environment.GetEnvironmentVariable("TEAMS_TENANT_ID");
        var clientId  = Environment.GetEnvironmentVariable("TEAMS_CLIENT_ID");
        string? devOpsOrg = null;
        string? devOpsPat = null;

        // Overlay with appsettings.json (always read it, not just when env vars are absent)
        var settingsPath = Path.Combine(AppContext.BaseDirectory, "appsettings.json");
        if (File.Exists(settingsPath))
        {
            var json     = File.ReadAllText(settingsPath);
            var settings = System.Text.Json.JsonSerializer.Deserialize<AppSettings>(json);

            tenantId  ??= settings?.TenantId;
            clientId  ??= settings?.ClientId;
            devOpsOrg ??= settings?.DevOpsOrganization;
            devOpsPat ??= settings?.DevOpsPat;
        }

        // Prompt for required AAD fields if still missing
        if (string.IsNullOrEmpty(tenantId))
        {
            Console.Write("Enter your Azure AD Tenant ID: ");
            tenantId = Console.ReadLine()?.Trim();
        }

        if (string.IsNullOrEmpty(clientId))
        {
            Console.Write("Enter your Azure AD Client ID (App Registration): ");
            clientId = Console.ReadLine()?.Trim();
        }

        if (string.IsNullOrEmpty(tenantId) || string.IsNullOrEmpty(clientId))
            throw new InvalidOperationException("TenantId and ClientId are required.");

        return new AppConfig
        {
            TenantId           = tenantId,
            ClientId           = clientId,
            DevOpsOrganization = devOpsOrg,
            DevOpsPat          = devOpsPat,
        };
    }

    private class AppSettings
    {
        public string? TenantId           { get; set; }
        public string? ClientId           { get; set; }
        public string? DevOpsOrganization { get; set; }
        public string? DevOpsPat          { get; set; }
    }
}
