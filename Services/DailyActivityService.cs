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

    public async Task<(GraphServiceClient Graph, string UserEmail, string UserId)> AuthenticateAsync()
    {
        _graphClient ??= GraphClientFactory.Build(_config);

        if (_userEmail == null)
        {
            var me     = await _graphClient.Me.GetAsync();
            _userEmail = me?.Mail ?? me?.UserPrincipalName ?? string.Empty;
            _userId    = me?.Id   ?? string.Empty;
        }

        return (_graphClient, _userEmail, _userId!);
    }

    // ─── Individual fetchers ──────────────────────────────────────────────────

    // userId is passed in so MessageFetcher can detect "sent by me" without
    // making a redundant /me call of its own
    public Task<List<TeamMessage>> FetchMessagesAsync(GraphServiceClient graph, string userId, DateOnly date, Action<string>? onWarning = null)
        => new MessageFetcher(graph, userId).GetMessagesForDateAsync(date, onWarning);

    public Task<List<CalendarEvent>> FetchEventsAsync(GraphServiceClient graph, DateOnly date, Action<string>? onWarning = null)
        => new CalendarFetcher(graph).GetEventsForDateAsync(date, onWarning);

    public Task<List<MailMessage>> FetchEmailsAsync(GraphServiceClient graph, DateOnly date, Action<string>? onWarning = null)
        => new MailFetcher(graph).GetMailForDateAsync(date, onWarning);

    public Task<List<DevOpsCommit>> FetchCommitsAsync(string userEmail, DateOnly date, Action<string>? onWarning = null)
        => new DevOpsFetcher(_config, userEmail).GetCommitsForDateAsync(date, onWarning);

    public Task<List<TodoTask>> FetchTodoTasksAsync(GraphServiceClient graph, DateOnly date, Action<string>? onWarning = null)
        => new TodoFetcher(graph, _config.TodoListName).GetTasksForDateAsync(date, onWarning);
}
