using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ODataErrors;

namespace TeamsMessageFetcher;

public class TodoFetcher(GraphServiceClient graph, string listName)
{
    private readonly GraphServiceClient _graph    = graph;
    private readonly string             _listName = listName;

    public async Task<List<TodoTask>> GetTasksForDateAsync(DateOnly date)
    {
        var result = new List<TodoTask>();

        string? listId;
        try
        {
            listId = await FindListIdAsync();
        }
        catch (ODataError odataEx)
        {
            LogODataError("list To Do lists", odataEx);
            return result;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"   ⚠️  Could not list To Do lists: {ex.Message}");
            return result;
        }

        if (listId == null)
        {
            Console.WriteLine($"   ⚠️  Could not find To Do list \"{_listName}\".");
            return result;
        }

        try
        {
            var response = await _graph.Me.Todo.Lists[listId].Tasks.GetAsync();

            var tasks = await GraphPaginator.PaginateAsync<Microsoft.Graph.Models.TodoTask, TodoTaskCollectionResponse>(
                _graph, response);

            result.AddRange(tasks
                .Select(Map)
                .Where(t => WasCreatedOrCompletedOn(t, date)));
        }
        catch (ODataError odataEx)
        {
            LogODataError("fetch To Do tasks", odataEx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"   ⚠️  Could not fetch To Do tasks: {ex.Message}");
        }

        return result;
    }

    private static void LogODataError(string action, ODataError odataEx)
    {
        var detail = odataEx.Error?.Message ?? odataEx.Message;
        var code   = odataEx.Error?.Code;
        Console.WriteLine($"   ⚠️  Could not {action}: {detail} (code: {code}, status: {odataEx.ResponseStatusCode})");
    }

    // ─── List lookup ──────────────────────────────────────────────────────────

    private async Task<string?> FindListIdAsync()
    {
        var response = await _graph.Me.Todo.Lists.GetAsync();

        var lists = await GraphPaginator.PaginateAsync<TodoTaskList, TodoTaskListCollectionResponse>(
            _graph, response);

        // 1) Exact match on the configured list name
        var match = lists.FirstOrDefault(l =>
            string.Equals(l.DisplayName, _listName, StringComparison.OrdinalIgnoreCase));

        // 2) Fall back to the account's default list (e.g. "Tasks" / "Aufgaben"),
        //    identified by the wellKnownListName property rather than its localized display name
        match ??= lists.FirstOrDefault(l =>
            l.WellknownListName == WellknownListName.DefaultList);

        // 3) Fall back to the first list available, if any
        match ??= lists.FirstOrDefault();

        if (match == null)
        {
            Console.WriteLine("   ⚠️  No To Do lists found for this account.");
        }
        else if (!string.Equals(match.DisplayName, _listName, StringComparison.OrdinalIgnoreCase))
        {
            Console.WriteLine($"   ℹ️  To Do list \"{_listName}\" not found; using \"{match.DisplayName}\" instead. " +
                               $"Available lists: {string.Join(", ", lists.Select(l => l.DisplayName))}");
        }

        return match?.Id;
    }

    // ─── Filtering ────────────────────────────────────────────────────────────

    private static bool WasCreatedOrCompletedOn(TodoTask task, DateOnly date)
    {
        if (DateOnly.FromDateTime(task.CreatedDateTime) == date)
            return true;

        return task.CompletedDateTime.HasValue &&
               DateOnly.FromDateTime(task.CompletedDateTime.Value) == date;
    }

    // ─── Mapping ──────────────────────────────────────────────────────────────

    private static TodoTask Map(Microsoft.Graph.Models.TodoTask task) => new()
    {
        Id                = task.Id ?? string.Empty,
        Title             = task.Title ?? "(No title)",
        Status            = task.Status?.ToString() ?? string.Empty,
        Importance        = task.Importance?.ToString() ?? string.Empty,
        CreatedDateTime   = ParseDateTimeTimeZone(task.CreatedDateTime),
        CompletedDateTime = task.CompletedDateTime != null
                              ? ParseDateTimeTimeZone(task.CompletedDateTime.DateTime)
                              : null,
        DueDateTime       = task.DueDateTime != null
                              ? ParseDateTimeTimeZone(task.DueDateTime.DateTime)
                              : null,
    };

    private static DateTime ParseDateTimeTimeZone(DateTimeOffset? dt) =>
        dt?.LocalDateTime ?? DateTime.MinValue;

    private static DateTime ParseDateTimeTimeZone(string? dt)
    {
        if (string.IsNullOrEmpty(dt)) return DateTime.MinValue;

        var parsed = DateTime.Parse(dt,
            System.Globalization.CultureInfo.InvariantCulture,
            System.Globalization.DateTimeStyles.None);

        return DateTime.SpecifyKind(parsed, DateTimeKind.Utc).ToLocalTime();
    }
}
