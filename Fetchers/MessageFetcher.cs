using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Kiota.Abstractions.Serialization;

namespace TeamsMessageFetcher;

public class MessageFetcher
{
    private readonly GraphServiceClient _graph;
    private readonly string?            _myUserId;

    // Caps concurrent Graph calls when fetching channel messages in parallel
    private readonly SemaphoreSlim _throttle = new(4, 4);

    public MessageFetcher(GraphServiceClient graph, string? myUserId = null)
    {
        _graph    = graph;
        _myUserId = myUserId;
    }

    public async Task<List<TeamMessage>> GetMessagesForDateAsync(DateOnly date)
    {
        var startUtc = date.ToDateTime(TimeOnly.MinValue, DateTimeKind.Local).ToUniversalTime();
        var endUtc   = startUtc.AddDays(1);

        // Chats and channels are independent — start both at the same time
        var chatTask    = FetchChatMessagesAsync(startUtc, endUtc);
        var channelTask = FetchChannelMessagesAsync(startUtc, endUtc);
        await Task.WhenAll(chatTask, channelTask);

        return [.. chatTask.Result, .. channelTask.Result];
    }

    // ─── Chat messages (1:1 and group chats) ─────────────────────────────────
    // Fetched sequentially — Graph does not support $filter on chat messages,
    // so we must page through each chat individually. Racing many concurrent
    // requests against the same endpoint triggers silent 429s.

    private async Task<List<TeamMessage>> FetchChatMessagesAsync(DateTime start, DateTime end)
    {
        var result = new List<TeamMessage>();
        try
        {
            var chatsPage = await _graph.Me.Chats.GetAsync(config =>
            {
                config.QueryParameters.Expand = ["members"];
                config.QueryParameters.Top    = 50;
            });

            var chats = await PaginateAsync<Chat, ChatCollectionResponse>(chatsPage);

            foreach (var chat in chats)
            {
                var messages = await FetchMessagesFromChatAsync(
                    chat.Id!, GetChatDisplayName(chat), start, end);
                result.AddRange(messages);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"   ⚠️  Could not fetch chats: {ex.Message}");
        }
        return result;
    }

    private async Task<List<TeamMessage>> FetchMessagesFromChatAsync(
        string chatId, string chatName, DateTime start, DateTime end)
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
            Console.WriteLine($"   ⚠️  Skipped chat '{chatName}': {ex.Message}");
            return [];
        }
    }

    // ─── Channel messages ─────────────────────────────────────────────────────
    // Channel messages support server-side $filter by date, so fetching them
    // in parallel is safe — each call returns only the matching day's messages.

    private async Task<List<TeamMessage>> FetchChannelMessagesAsync(DateTime start, DateTime end)
    {
        var result = new List<TeamMessage>();
        try
        {
            var teamsPage = await _graph.Me.JoinedTeams.GetAsync();
            var teams     = await PaginateAsync<Team, TeamCollectionResponse>(teamsPage);

            // Build the full list of (teamId, channelId, sourceName) tuples first,
            // then fan out the message fetches in parallel with a throttle
            var channelRefs = new List<(string TeamId, string ChannelId, string Source)>();

            foreach (var team in teams)
            {
                var channelsPage = await _graph.Teams[team.Id].Channels.GetAsync();
                var channels     = await PaginateAsync<Channel, ChannelCollectionResponse>(channelsPage);

                foreach (var ch in channels)
                    channelRefs.Add((team.Id!, ch.Id!, $"{team.DisplayName} › #{ch.DisplayName}"));
            }

            // Fetch messages for all channels concurrently (server-side filter makes this safe)
            var tasks = channelRefs.Select(r =>
                Throttled(() => FetchMessagesFromChannelAsync(r.TeamId, r.ChannelId, r.Source, start, end)));

            var results = await Task.WhenAll(tasks);
            result.AddRange(results.SelectMany(r => r));
        }
        catch (Exception ex)
        {
            Console.WriteLine($"   ⚠️  Could not fetch channel messages: {ex.Message}");
        }
        return result;
    }

    private async Task<List<TeamMessage>> FetchMessagesFromChannelAsync(
        string teamId, string channelId, string sourceName, DateTime start, DateTime end)
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
            Console.WriteLine($"   ⚠️  Skipped channel '{sourceName}': {ex.Message}");
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
        };
    }

    private static string GetChatDisplayName(Chat chat)
    {
        if (!string.IsNullOrEmpty(chat.Topic))
            return $"Chat: {chat.Topic}";

        if (chat.Members?.Count > 0)
        {
            var names = chat.Members
                .Take(3)
                .OfType<AadUserConversationMember>()
                .Select(m => m.DisplayName ?? "?")
                .Where(n => n != "?");
            return $"Chat: {string.Join(", ", names)}";
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
