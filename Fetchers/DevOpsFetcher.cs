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

    public async Task<List<DevOpsCommit>> GetCommitsForDateAsync(DateOnly date)
    {
        if (string.IsNullOrEmpty(_org) || !_http.DefaultRequestHeaders.Contains("Authorization"))
        {
            Console.WriteLine("   ⚠️  DevOps org or PAT not configured — skipping commits.");
            return [];
        }

        var startUtc = date.ToDateTime(TimeOnly.MinValue, DateTimeKind.Local).ToUniversalTime();
        var endUtc   = startUtc.AddDays(1);

        var result = new List<DevOpsCommit>();

        try
        {
            var projects = await GetProjectsAsync();

            foreach (var project in projects)
            {
                var repos = await GetRepositoriesAsync(project.Name);

                foreach (var repo in repos)
                {
                    var commits = await GetCommitsAsync(project.Name, repo.Id, repo.Name, startUtc, endUtc);
                    result.AddRange(commits);
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"   ⚠️  Could not fetch DevOps commits: {ex.Message}");
        }

        return result;
    }

    // ─── API calls ────────────────────────────────────────────────────────────

    private async Task<List<AdoProject>> GetProjectsAsync()
    {
        var url      = $"https://dev.azure.com/{_org}/_apis/projects?api-version=7.1";
        var response = await GetAsync<AdoPagedResult<AdoProject>>(url);
        return response?.Value ?? [];
    }

    private async Task<List<AdoRepository>> GetRepositoriesAsync(string project)
    {
        var url      = $"https://dev.azure.com/{_org}/{Uri.EscapeDataString(project)}/_apis/git/repositories?api-version=7.1";
        var response = await GetAsync<AdoPagedResult<AdoRepository>>(url);
        return response?.Value ?? [];
    }

    private async Task<List<DevOpsCommit>> GetCommitsAsync(
        string project, string repoId, string repoName,
        DateTime fromUtc, DateTime toUtc)
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

        var response = await GetAsync<AdoPagedResult<AdoCommit>>(url);
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

    private async Task<T?> GetAsync<T>(string url)
    {
        using var response = await _http.GetAsync(url);
        response.EnsureSuccessStatusCode();
        var json = await response.Content.ReadAsStringAsync();
        return JsonSerializer.Deserialize<T>(json, JsonOptions);
    }

    // ─── ADO response models ──────────────────────────────────────────────────

    private sealed class AdoPagedResult<T>
    {
        [JsonPropertyName("value")] public List<T> Value { get; set; } = [];
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
