namespace TeamsMessageFetcher;

// ─── App Configuration ────────────────────────────────────────────────────────

public class AppConfig
{
    public string  TenantId           { get; set; } = string.Empty;
    public string  ClientId           { get; set; } = string.Empty;

    // Azure DevOps — optional; commits are skipped when either is absent
    public string? DevOpsOrganization { get; set; }
    public string? DevOpsPat          { get; set; }

    // Microsoft To Do — optional; defaults to "Tasks" when absent
    public string  TodoListName       { get; set; } = "Tasks";

    // Component switches
    public bool EnableMessages { get; set; } = true;
    public bool EnableEvents   { get; set; } = true;
    public bool EnableEmails   { get; set; } = true;
    public bool EnableCommits  { get; set; } = true;
    public bool EnableTodo     { get; set; } = true;
    public bool ShowWarnings   { get; set; } = true;

    public void Save()
    {
        var settingsPath = Path.Combine(AppContext.BaseDirectory, "appsettings.json");
        var settings = new AppSettings
        {
            TenantId           = TenantId,
            ClientId           = ClientId,
            DevOpsOrganization = DevOpsOrganization,
            DevOpsPat          = DevOpsPat,
            TodoListName       = TodoListName,
            EnableMessages     = EnableMessages,
            EnableEvents       = EnableEvents,
            EnableEmails       = EnableEmails,
            EnableCommits      = EnableCommits,
            EnableTodo         = EnableTodo,
            ShowWarnings       = ShowWarnings,
        };
        var options = new System.Text.Json.JsonSerializerOptions { WriteIndented = true };
        File.WriteAllText(settingsPath, System.Text.Json.JsonSerializer.Serialize(settings, options));
    }

    public static AppConfig Load()
    {
        // Start from environment variables
        var tenantId  = Environment.GetEnvironmentVariable("TEAMS_TENANT_ID");
        var clientId  = Environment.GetEnvironmentVariable("TEAMS_CLIENT_ID");
        string? devOpsOrg = null;
        string? devOpsPat = null;
        string? todoListName = Environment.GetEnvironmentVariable("TEAMS_TODO_LIST");
        AppSettings? settings = null;

        // Overlay with appsettings.json (always read it, not just when env vars are absent)
        var settingsPath = Path.Combine(AppContext.BaseDirectory, "appsettings.json");
        if (File.Exists(settingsPath))
        {
            var json     = File.ReadAllText(settingsPath);
            settings = System.Text.Json.JsonSerializer.Deserialize<AppSettings>(json);

            tenantId  ??= settings?.TenantId;
            clientId  ??= settings?.ClientId;
            devOpsOrg ??= settings?.DevOpsOrganization;
            devOpsPat ??= settings?.DevOpsPat;
            todoListName ??= settings?.TodoListName;
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
            TodoListName       = string.IsNullOrWhiteSpace(todoListName) ? "Tasks" : todoListName,
            EnableMessages     = settings?.EnableMessages ?? true,
            EnableEvents       = settings?.EnableEvents   ?? true,
            EnableEmails       = settings?.EnableEmails   ?? true,
            EnableCommits      = settings?.EnableCommits  ?? true,
            EnableTodo         = settings?.EnableTodo     ?? true,
            ShowWarnings       = settings?.ShowWarnings   ?? true,
        };
    }

    private class AppSettings
    {
        public string? TenantId           { get; set; }
        public string? ClientId           { get; set; }
        public string? DevOpsOrganization { get; set; }
        public string? DevOpsPat          { get; set; }
        public string? TodoListName       { get; set; }
        public bool?   EnableMessages     { get; set; }
        public bool?   EnableEvents       { get; set; }
        public bool?   EnableEmails       { get; set; }
        public bool?   EnableCommits      { get; set; }
        public bool?   EnableTodo         { get; set; }
        public bool?   ShowWarnings       { get; set; }
    }
}
