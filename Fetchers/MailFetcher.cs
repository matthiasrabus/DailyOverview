using Microsoft.Graph;
using Microsoft.Graph.Models;

namespace TeamsMessageFetcher;

public class MailFetcher(GraphServiceClient graph)
{
    private readonly GraphServiceClient _graph = graph;

    private static readonly string[] SelectFields =
        ["subject", "from", "toRecipients", "receivedDateTime", "sentDateTime", "bodyPreview"];

    public async Task<List<MailMessage>> GetMailForDateAsync(DateOnly date)
    {
        var startUtc = date.ToDateTime(TimeOnly.MinValue, DateTimeKind.Local).ToUniversalTime();
        var endUtc   = startUtc.AddDays(1);

        var startStr = startUtc.ToString("yyyy-MM-ddTHH:mm:ssZ");
        var endStr   = endUtc.ToString("yyyy-MM-ddTHH:mm:ssZ");

        var received = await FetchReceivedAsync(startStr, endStr);
        var sent     = await FetchSentAsync(startStr, endStr);

        return [.. received, .. sent];
    }

    // ─── Received ─────────────────────────────────────────────────────────────

    private async Task<List<MailMessage>> FetchReceivedAsync(string startStr, string endStr)
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
            Console.WriteLine($"   ⚠️  Could not fetch received mail: {ex.Message}");
            return [];
        }
    }

    // ─── Sent ─────────────────────────────────────────────────────────────────

    private async Task<List<MailMessage>> FetchSentAsync(string startStr, string endStr)
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
            Console.WriteLine($"   ⚠️  Could not fetch sent mail: {ex.Message}");
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
    };
}
