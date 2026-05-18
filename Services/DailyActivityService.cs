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
    public Task<List<TeamMessage>> FetchMessagesAsync(GraphServiceClient graph, string userId, DateOnly date)
        => new MessageFetcher(graph, userId).GetMessagesForDateAsync(date);

    public Task<List<CalendarEvent>> FetchEventsAsync(GraphServiceClient graph, DateOnly date)
        => new CalendarFetcher(graph).GetEventsForDateAsync(date);

    public Task<List<MailMessage>> FetchEmailsAsync(GraphServiceClient graph, DateOnly date)
        => new MailFetcher(graph).GetMailForDateAsync(date);

    public Task<List<DevOpsCommit>> FetchCommitsAsync(string userEmail, DateOnly date)
        => new DevOpsFetcher(_config, userEmail).GetCommitsForDateAsync(date);
}
