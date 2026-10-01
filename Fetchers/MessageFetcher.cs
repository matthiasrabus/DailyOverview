using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Kiota.Abstractions.Serialization;

namespace TeamsMessageFetcher;

public class MessageFetcher
{
    private readonly GraphServiceClient _graph;
    private readonly string?            _myUserId;

    // Caps concurrent Graph calls when fetching channel messages in parallel
    private readonly SemaphoreSlim _throttle = new(6, 6);

    public MessageFetcher(GraphServiceClient graph, string? myUserId = null)
    {
        _graph    = graph;
        _myUserId = myUserId;
    }

    public async Task<List<TeamMessage>> GetMessagesForDateAsync(DateOnly date, Action<string>? onWarning = null)
    {
        var startUtc = date.ToDateTime(TimeOnly.MinValue, DateTimeKind.Local).ToUniversalTime();
        var endUtc   = startUtc.AddDays(1);

        // Chats and channels are independent — start both at the same time
        var chatTask    = FetchChatMessagesAsync(startUtc, endUtc, onWarning);
        var channelTask = FetchChannelMessagesAsync(startUtc, endUtc, onWarning);
        await Task.WhenAll(chatTask, channelTask);

        return [.. chatTask.Result, .. channelTask.Result];
    }

    // ─── Chat messages (1:1 and group chats) ─────────────────────────────────
    // Graph does not support $filter by date on chat messages, so we page each
    // chat individually. Fetches run in parallel (capped at throttle limit).

    private async Task<List<TeamMessage>> FetchChatMessagesAsync(DateTime start, DateTime end, Action<string>? onWarning)
    {
        try
        {
            var chatsPage = await _graph.Me.Chats.GetAsync(config =>
            {
                config.QueryParameters.Expand = ["members"];
                config.QueryParameters.Top    = 50;
            });

            var chats = await PaginateAsync<Chat, ChatCollectionResponse>(chatsPage);

            var tasks   = chats.Select(chat =>
                Throttled(() => FetchMessagesFromChatAsync(chat.Id!, GetChatDisplayName(chat), start, end, onWarning)));
            var results = await Task.WhenAll(tasks);
            return [.. results.SelectMany(r => r)];
        }
        catch (Exception ex)
        {
            var message = $"Could not fetch chats: {ex.Message}";
            Console.WriteLine($"   ⚠️  {message}");
            onWarning?.Invoke(message);
            return [];
        }
    }

    private async Task<List<TeamMessage>> FetchMessagesFromChatAsync(
        string chatId, string chatName, DateTime start, DateTime end, Action<string>? onWarning)
    {
        try
        {
            // Graph does not support $filter by date on chat messages —
            // fetch in descending order and stop as soon as we pass the target window
            var page = await _graph.Me.Chats[chatId].Messages.GetAsync(config =>
            {
                config.QueryParameters.Top     = 50;
                config.QueryParameters.Orderby = ["createdDateTime desc"];
            });

            var messages = await PaginateAsync<ChatMessage, ChatMessageCollectionResponse>(
                page,
                stopCondition: msg => msg.CreatedDateTime < start);

            return messages
                .Where(m => m.CreatedDateTime?.UtcDateTime is { } t && t >= start && t < end)
                .Select(m => MapToTeamMessage(m, chatName))
                .ToList();
        }
        catch (Exception ex)
        {
            var message = $"Skipped chat '{chatName}': {ex.Message}";
            Console.WriteLine($"   ⚠️  {message}");
            onWarning?.Invoke(message);
            return [];
        }
    }

    // ─── Channel messages ─────────────────────────────────────────────────────
    // Channel messages support server-side $filter by date, so fetching them
    // in parallel is safe — each call returns only the matching day's messages.

    private async Task<List<TeamMessage>> FetchChannelMessagesAsync(DateTime start, DateTime end, Action<string>? onWarning)
    {
        var result = new List<TeamMessage>();
        try
        {
            var teamsPage = await _graph.Me.JoinedTeams.GetAsync();
            var teams     = await PaginateAsync<Team, TeamCollectionResponse>(teamsPage);

            // Fan out channel enumeration across teams, then message fetches across channels
            var channelRefLists = await Task.WhenAll(teams.Select(team =>
                Throttled(async () =>
                {
                    var channelsPage = await _graph.Teams[team.Id].Channels.GetAsync();
                    var channels = await PaginateAsync<Channel, ChannelCollectionResponse>(channelsPage);
                    return channels.Select(ch =>
                        (TeamId: team.Id!, ChannelId: ch.Id!, Source: $"{team.DisplayName} › #{ch.DisplayName}"));
                })));

            var channelRefs = channelRefLists.SelectMany(r => r);

            var tasks = channelRefs.Select(r =>
                Throttled(() => FetchMessagesFromChannelAsync(r.TeamId, r.ChannelId, r.Source, start, end, onWarning)));

            var results = await Task.WhenAll(tasks);
            result.AddRange(results.SelectMany(r => r));
        }
        catch (Exception ex)
        {
            var message = $"Could not fetch channel messages: {ex.Message}";
            Console.WriteLine($"   ⚠️  {message}");
            onWarning?.Invoke(message);
        }
        return result;
    }

    private async Task<List<TeamMessage>> FetchMessagesFromChannelAsync(
        string teamId, string channelId, string sourceName, DateTime start, DateTime end, Action<string>? onWarning)
    {
        try
        {
            var startStr = start.ToString("yyyy-MM-ddTHH:mm:ssZ");
            var endStr   = end.ToString("yyyy-MM-ddTHH:mm:ssZ");

            var page = await _graph.Teams[teamId].Channels[channelId].Messages.GetAsync(config =>
            {
                config.QueryParameters.Filter =
                    $"createdDateTime ge {startStr} and createdDateTime lt {endStr}";
                config.QueryParameters.Top = 50;
            });

            var messages = await PaginateAsync<ChatMessage, ChatMessageCollectionResponse>(page);
            return messages.Select(m => MapToTeamMessage(m, sourceName)).ToList();
        }
        catch (Exception ex)
        {
            var message = $"Skipped channel '{sourceName}': {ex.Message}";
            Console.WriteLine($"   ⚠️  {message}");
            onWarning?.Invoke(message);
            return [];
        }
    }

    // ─── Throttle helper ──────────────────────────────────────────────────────

    private async Task<T> Throttled<T>(Func<Task<T>> work)
    {
        await _throttle.WaitAsync();
        try     { return await work(); }
        finally { _throttle.Release(); }
    }

    // ─── Mapping ──────────────────────────────────────────────────────────────

    private TeamMessage MapToTeamMessage(ChatMessage msg, string source)
    {
        var senderId   = msg.From?.User?.Id;
        var senderName = msg.From?.User?.DisplayName
                      ?? msg.From?.Application?.DisplayName
                      ?? "Unknown";

        var body = msg.Body?.Content ?? string.Empty;
        if (msg.Body?.ContentType == BodyType.Html)
            body = StripHtml(body);

        var preview = body.Length > 120 ? body[..117] + "..." : body;

        return new TeamMessage
        {
            Id          = msg.Id ?? string.Empty,
            Source      = source,
            SenderName  = senderName,
            IsSentByMe  = senderId != null && senderId == _myUserId,
            Timestamp   = msg.CreatedDateTime?.LocalDateTime ?? DateTime.MinValue,
            BodyPreview = preview.Trim(),
            FullBody    = body.Trim(),
            MessageType = msg.MessageType?.ToString() ?? "message",
            WebLink     = !string.IsNullOrEmpty(msg.WebUrl) ? msg.WebUrl : BuildChatLink(msg),
        };
    }

    // Chat messages usually have no WebUrl; build a Teams deep link from chat and message id
    private static string BuildChatLink(ChatMessage msg)
    {
        if (string.IsNullOrEmpty(msg.ChatId) || string.IsNullOrEmpty(msg.Id))
            return string.Empty;

        return $"https://teams.microsoft.com/l/message/{Uri.EscapeDataString(msg.ChatId)}/{Uri.EscapeDataString(msg.Id)}"
             + "?context=%7B%22contextType%22%3A%22chat%22%7D";
    }

    private static string GetChatDisplayName(Chat chat)
    {
        if (!string.IsNullOrEmpty(chat.Topic))
            return chat.Topic;

        if (chat.Members?.Count > 0)
        {
            var names = chat.Members
                .Take(3)
                .OfType<AadUserConversationMember>()
                .Select(m => m.DisplayName ?? "?")
                .Where(n => n != "?");
            return string.Join(", ", names);
        }

        return $"Chat ({chat.ChatType})";
    }

    private static readonly System.Text.RegularExpressions.Regex HtmlTagRegex =
        new("<[^>]+>", System.Text.RegularExpressions.RegexOptions.Compiled);
    private static readonly System.Text.RegularExpressions.Regex WhitespaceRegex =
        new(@"\s{2,}", System.Text.RegularExpressions.RegexOptions.Compiled);

    private static string StripHtml(string html)
    {
        var result = HtmlTagRegex.Replace(html, " ");
        result = result
            .Replace("&nbsp;", " ").Replace("&amp;",  "&")
            .Replace("&lt;",   "<").Replace("&gt;",   ">")
            .Replace("&quot;", "\"").Replace("&#39;", "'");
        return WhitespaceRegex.Replace(result, " ").Trim();
    }

    private Task<List<T>> PaginateAsync<T, TCollection>(
        TCollection?   firstPage,
        Func<T, bool>? stopCondition = null)
        where T           : class
        where TCollection : class, IParsable, IAdditionalDataHolder, new()
        => GraphPaginator.PaginateAsync<T, TCollection>(_graph, firstPage, stopCondition);
}
