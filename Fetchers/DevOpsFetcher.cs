using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace TeamsMessageFetcher;

public class DevOpsFetcher
{
    private readonly HttpClient _http;
    private readonly string _org;
    private readonly string _userEmail;

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true,
    };

    public DevOpsFetcher(AppConfig config, string userEmail)
    {
        _org       = config.DevOpsOrganization ?? string.Empty;
        _userEmail = userEmail;

        _http = new HttpClient();

        if (!string.IsNullOrEmpty(config.DevOpsPat))
        {
            var encoded = Convert.ToBase64String(Encoding.ASCII.GetBytes($":{config.DevOpsPat}"));
            _http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", encoded);
        }
    }

    public async Task<List<DevOpsCommit>> GetCommitsForDateAsync(DateOnly date, Action<string>? onWarning = null, CancellationToken ct = default)
    {
        if (string.IsNullOrEmpty(_org) || !_http.DefaultRequestHeaders.Contains("Authorization"))
        {
            const string message = "DevOps org or PAT not configured — skipping commits.";
            FetchLog.Warning(message);
            onWarning?.Invoke(message);
            return [];
        }

        var startUtc = date.ToDateTime(TimeOnly.MinValue, DateTimeKind.Local).ToUniversalTime();
        var endUtc   = startUtc.AddDays(1);

        var result = new List<DevOpsCommit>();

        try
        {
            var projects = await GetProjectsAsync(ct);

            foreach (var project in projects)
            {
                var repos = await GetRepositoriesAsync(project.Name, ct);

                foreach (var repo in repos)
                {
                    var commits = await GetCommitsAsync(project.Name, repo.Id, repo.Name, startUtc, endUtc, ct);
                    result.AddRange(commits);
                }
            }
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            var message = $"Could not fetch DevOps commits: {ex.Message}";
            FetchLog.Warning(message);
            onWarning?.Invoke(message);
        }

        return result;
    }

    public async Task<List<DevOpsWorkItem>> GetWorkItemsForDateAsync(DateOnly date, Action<string>? onWarning = null, CancellationToken ct = default)
    {
        if (string.IsNullOrEmpty(_org) || !_http.DefaultRequestHeaders.Contains("Authorization"))
        {
            const string message = "DevOps org or PAT not configured — skipping work items.";
            FetchLog.Warning(message);
            onWarning?.Invoke(message);
            return [];
        }

        try
        {
            var from = date.ToString("yyyy-MM-dd");
            var to   = date.AddDays(1).ToString("yyyy-MM-dd");
            var wiql = "SELECT [System.Id] FROM WorkItems " +
                       $"WHERE [System.ChangedDate] >= '{from}' AND [System.ChangedDate] < '{to}' " +
                       "AND ([System.ChangedBy] = @Me OR [System.AssignedTo] = @Me) " +
                       "ORDER BY [System.ChangedDate] ASC";

            var queryResult = await PostAsync<AdoWiqlResult>(
                $"https://dev.azure.com/{_org}/_apis/wit/wiql?$top=200&api-version=7.1",
                new { query = wiql }, ct);

            var ids = queryResult?.WorkItems.Select(w => w.Id).ToList() ?? [];
            if (ids.Count == 0) return [];

            var batch = await PostAsync<AdoPagedResult<AdoWorkItem>>(
                $"https://dev.azure.com/{_org}/_apis/wit/workitemsbatch?api-version=7.1",
                new
                {
                    ids,
                    fields = new[]
                    {
                        "System.Id", "System.Title", "System.WorkItemType", "System.State",
                        "System.TeamProject", "System.AssignedTo", "System.ChangedDate",
                    },
                }, ct);

            return (batch?.Value ?? []).Select(w =>
            {
                var project = GetField(w, "System.TeamProject");
                return new DevOpsWorkItem
                {
                    Id          = w.Id,
                    Title       = GetField(w, "System.Title"),
                    Type        = GetField(w, "System.WorkItemType"),
                    State       = GetField(w, "System.State"),
                    Project     = project,
                    AssignedTo  = GetAssignedTo(w),
                    ChangedDate = DateTime.TryParse(GetField(w, "System.ChangedDate"), null,
                                      System.Globalization.DateTimeStyles.AdjustToUniversal | System.Globalization.DateTimeStyles.AssumeUniversal,
                                      out var changed) ? changed.ToLocalTime() : default,
                    Url         = $"https://dev.azure.com/{_org}/{Uri.EscapeDataString(project)}/_workitems/edit/{w.Id}",
                };
            }).OrderBy(w => w.ChangedDate).ToList();
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            var message = $"Could not fetch DevOps work items: {ex.Message}";
            FetchLog.Warning(message);
            onWarning?.Invoke(message);
            return [];
        }
    }

    private static string GetField(AdoWorkItem w, string name)
        => w.Fields.TryGetValue(name, out var v) && v.ValueKind == JsonValueKind.String ? v.GetString() ?? string.Empty : string.Empty;

    private static string GetAssignedTo(AdoWorkItem w)
        => w.Fields.TryGetValue("System.AssignedTo", out var v) && v.ValueKind == JsonValueKind.Object
            && v.TryGetProperty("displayName", out var n) ? n.GetString() ?? string.Empty : string.Empty;

    // ─── API calls 

    private async Task<List<AdoProject>> GetProjectsAsync(CancellationToken ct)
    {
        var url      = $"https://dev.azure.com/{_org}/_apis/projects?api-version=7.1";
        var response = await GetAsync<AdoPagedResult<AdoProject>>(url, ct);
        return response?.Value ?? [];
    }

    private async Task<List<AdoRepository>> GetRepositoriesAsync(string project, CancellationToken ct)
    {
        var url      = $"https://dev.azure.com/{_org}/{Uri.EscapeDataString(project)}/_apis/git/repositories?api-version=7.1";
        var response = await GetAsync<AdoPagedResult<AdoRepository>>(url, ct);
        return response?.Value ?? [];
    }

    private async Task<List<DevOpsCommit>> GetCommitsAsync(
        string project, string repoId, string repoName,
        DateTime fromUtc, DateTime toUtc, CancellationToken ct)
    {
        var from = Uri.EscapeDataString(fromUtc.ToString("yyyy-MM-ddTHH:mm:ssZ"));
        var to   = Uri.EscapeDataString(toUtc.ToString("yyyy-MM-ddTHH:mm:ssZ"));
        var author = Uri.EscapeDataString(_userEmail);

        var url = $"https://dev.azure.com/{_org}/{Uri.EscapeDataString(project)}" +
                  $"/_apis/git/repositories/{repoId}/commits" +
                  $"?searchCriteria.fromDate={from}" +
                  $"&searchCriteria.toDate={to}" +
                  $"&searchCriteria.author={author}" +
                  $"&$top=1000&api-version=7.1";

        var response = await GetAsync<AdoPagedResult<AdoCommit>>(url, ct);
        if (response?.Value == null) return [];

        return response.Value.Select(c => new DevOpsCommit
        {
            CommitId   = c.CommitId,
            Message    = c.Comment.Split('\n')[0].Trim(), // first line only
            Timestamp  = c.Author.Date.ToLocalTime(),
            Project    = project,
            Repository = repoName,
            Url        = c.RemoteUrl,
        }).ToList();
    }

    private async Task<T?> GetAsync<T>(string url, CancellationToken ct)
    {
        using var response = await _http.GetAsync(url, ct);
        response.EnsureSuccessStatusCode();
        var json = await response.Content.ReadAsStringAsync(ct);
        return JsonSerializer.Deserialize<T>(json, JsonOptions);
    }

    private async Task<T?> PostAsync<T>(string url, object body, CancellationToken ct)
    {
        using var content  = new StringContent(JsonSerializer.Serialize(body), Encoding.UTF8, "application/json");
        using var response = await _http.PostAsync(url, content, ct);
        response.EnsureSuccessStatusCode();
        var json = await response.Content.ReadAsStringAsync(ct);
        return JsonSerializer.Deserialize<T>(json, JsonOptions);
    }

    // ─── ADO response models 

    private sealed class AdoPagedResult<T>
    {
        [JsonPropertyName("value")] public List<T> Value { get; set; } = [];
    }

    private sealed class AdoWiqlResult
    {
        [JsonPropertyName("workItems")] public List<AdoWorkItemRef> WorkItems { get; set; } = [];
    }

    private sealed class AdoWorkItemRef
    {
        [JsonPropertyName("id")] public int Id { get; set; }
    }

    private sealed class AdoWorkItem
    {
        [JsonPropertyName("id")]     public int Id { get; set; }
        [JsonPropertyName("fields")] public Dictionary<string, JsonElement> Fields { get; set; } = [];
    }

    private sealed class AdoProject
    {
        [JsonPropertyName("id")]   public string Id   { get; set; } = string.Empty;
        [JsonPropertyName("name")] public string Name { get; set; } = string.Empty;
    }

    private sealed class AdoRepository
    {
        [JsonPropertyName("id")]   public string Id   { get; set; } = string.Empty;
        [JsonPropertyName("name")] public string Name { get; set; } = string.Empty;
    }

    private sealed class AdoCommit
    {
        [JsonPropertyName("commitId")]  public string        CommitId  { get; set; } = string.Empty;
        [JsonPropertyName("comment")]   public string        Comment   { get; set; } = string.Empty;
        [JsonPropertyName("author")]    public AdoAuthor     Author    { get; set; } = new();
        [JsonPropertyName("remoteUrl")] public string        RemoteUrl { get; set; } = string.Empty;
    }

    private sealed class AdoAuthor
    {
        [JsonPropertyName("name")]  public string   Name  { get; set; } = string.Empty;
        [JsonPropertyName("email")] public string   Email { get; set; } = string.Empty;
        [JsonPropertyName("date")]  public DateTime Date  { get; set; }
    }
}
