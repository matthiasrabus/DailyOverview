using Microsoft.Graph;
using Microsoft.Graph.Models;

namespace TeamsMessageFetcher;

public class MailFetcher(GraphServiceClient graph)
{
    private readonly GraphServiceClient _graph = graph;

    private static readonly string[] SelectFields =
        ["subject", "from", "toRecipients", "receivedDateTime", "sentDateTime", "bodyPreview", "webLink"];

    public async Task<List<MailMessage>> GetMailForDateAsync(DateOnly date, Action<string>? onWarning = null)
    {
        var startUtc = date.ToDateTime(TimeOnly.MinValue, DateTimeKind.Local).ToUniversalTime();
        var endUtc   = startUtc.AddDays(1);

        var startStr = startUtc.ToString("yyyy-MM-ddTHH:mm:ssZ");
        var endStr   = endUtc.ToString("yyyy-MM-ddTHH:mm:ssZ");

        var received = await FetchReceivedAsync(startStr, endStr, onWarning);
        var sent     = await FetchSentAsync(startStr, endStr, onWarning);

        return [.. received, .. sent];
    }

    // ─── Received ─────────────────────────────────────────────────────────────

    private async Task<List<MailMessage>> FetchReceivedAsync(string startStr, string endStr, Action<string>? onWarning)
    {
        try
        {
            var response = await _graph.Me.Messages.GetAsync(config =>
            {
                config.QueryParameters.Filter = $"receivedDateTime ge {startStr} and receivedDateTime lt {endStr}";
                config.QueryParameters.Select = SelectFields;
                config.QueryParameters.Top    = 50;
            });

            var messages = await GraphPaginator.PaginateAsync<Message, MessageCollectionResponse>(_graph, response);
            return messages.Select(m => Map(m, isSent: false)).ToList();
        }
        catch (Exception ex)
        {
            var message = $"Could not fetch received mail: {ex.Message}";
            Console.WriteLine($"   ⚠️  {message}");
            onWarning?.Invoke(message);
            return [];
        }
    }

    // ─── Sent ─────────────────────────────────────────────────────────────────

    private async Task<List<MailMessage>> FetchSentAsync(string startStr, string endStr, Action<string>? onWarning)
    {
        try
        {
            var response = await _graph.Me.MailFolders["SentItems"].Messages.GetAsync(config =>
            {
                config.QueryParameters.Filter = $"sentDateTime ge {startStr} and sentDateTime lt {endStr}";
                config.QueryParameters.Select = SelectFields;
                config.QueryParameters.Top    = 50;
            });

            var messages = await GraphPaginator.PaginateAsync<Message, MessageCollectionResponse>(_graph, response);
            return messages.Select(m => Map(m, isSent: true)).ToList();
        }
        catch (Exception ex)
        {
            var message = $"Could not fetch sent mail: {ex.Message}";
            Console.WriteLine($"   ⚠️  {message}");
            onWarning?.Invoke(message);
            return [];
        }
    }

    // ─── Mapping ──────────────────────────────────────────────────────────────

    private static MailMessage Map(Message msg, bool isSent) => new()
    {
        Id          = msg.Id ?? string.Empty,
        Subject     = msg.Subject ?? "(No subject)",
        Timestamp   = isSent
                        ? msg.SentDateTime?.LocalDateTime ?? DateTime.MinValue
                        : msg.ReceivedDateTime?.LocalDateTime ?? DateTime.MinValue,
        From        = msg.From?.EmailAddress?.Name ?? msg.From?.EmailAddress?.Address ?? "Unknown",
        To          = msg.ToRecipients?
                        .Select(r => r.EmailAddress?.Name ?? r.EmailAddress?.Address ?? "?")
                        .ToList() ?? [],
        IsSent      = isSent,
        BodyPreview = msg.BodyPreview ?? string.Empty,
        WebLink     = msg.WebLink ?? string.Empty,
    };
}
