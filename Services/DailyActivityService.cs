using Microsoft.Graph;

namespace TeamsMessageFetcher.Services;

public class DailyActivityService
{
    private readonly AppConfig _config;
    private GraphServiceClient? _graphClient;
    private string? _userEmail;
    private string? _userId;

    public DailyActivityService(AppConfig config) => _config = config;

    // ─── Authentication ───────────────────────────────────────────────────────

    public async Task<(GraphServiceClient Graph, string UserEmail, string UserId)> AuthenticateAsync(CancellationToken ct = default)
    {
        _graphClient ??= GraphClientFactory.Build(_config);

        if (_userEmail == null)
        {
            var me     = await _graphClient.Me.GetAsync(cancellationToken: ct);
            _userEmail = me?.Mail ?? me?.UserPrincipalName ?? string.Empty;
            _userId    = me?.Id   ?? string.Empty;
        }

        return (_graphClient, _userEmail, _userId!);
    }

    // ─── Individual fetchers ──────────────────────────────────────────────────

    // userId is passed in so MessageFetcher can detect "sent by me" without
    // making a redundant /me call of its own
    public Task<List<TeamMessage>> FetchMessagesAsync(GraphServiceClient graph, string userId, DateOnly date, Action<string>? onWarning = null, CancellationToken ct = default)
        => new MessageFetcher(graph, userId).GetMessagesForDateAsync(date, onWarning, ct);

    public Task<List<CalendarEvent>> FetchEventsAsync(GraphServiceClient graph, DateOnly date, Action<string>? onWarning = null, CancellationToken ct = default)
        => new CalendarFetcher(graph).GetEventsForDateAsync(date, onWarning, ct);

    public Task<List<MailMessage>> FetchEmailsAsync(GraphServiceClient graph, DateOnly date, Action<string>? onWarning = null, CancellationToken ct = default)
        => new MailFetcher(graph).GetMailForDateAsync(date, onWarning, ct);

    public Task<List<DevOpsCommit>> FetchCommitsAsync(string userEmail, DateOnly date, Action<string>? onWarning = null, CancellationToken ct = default)
        => new DevOpsFetcher(_config, userEmail).GetCommitsForDateAsync(date, onWarning, ct);

    public Task<List<DevOpsWorkItem>> FetchWorkItemsAsync(string userEmail, DateOnly date, Action<string>? onWarning = null, CancellationToken ct = default)
        => new DevOpsFetcher(_config, userEmail).GetWorkItemsForDateAsync(date, onWarning, ct);

    public Task<List<TodoTask>> FetchTodoTasksAsync(GraphServiceClient graph, DateOnly date, Action<string>? onWarning = null, CancellationToken ct = default)
        => new TodoFetcher(graph, _config.TodoListName).GetTasksForDateAsync(date, onWarning, ct);
}
