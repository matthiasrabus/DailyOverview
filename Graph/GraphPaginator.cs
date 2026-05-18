using Microsoft.Graph;
using Microsoft.Kiota.Abstractions.Serialization;

namespace TeamsMessageFetcher;

// ─── Shared Graph pagination helper ──────────────────────────────────────────
// Wraps PageIterator so every fetcher can paginate without duplicating code.

internal static class GraphPaginator
{
    internal static async Task<List<T>> PaginateAsync<T, TCollection>(
        GraphServiceClient graph,
        TCollection? firstPage,
        Func<T, bool>? stopCondition = null)
        where T : class
        where TCollection : class, IParsable, IAdditionalDataHolder, new()
    {
        var result = new List<T>();
        if (firstPage == null) return result;

        var pageIterator = PageIterator<T, TCollection>.CreatePageIterator(
            graph,
            firstPage,
            item =>
            {
                if (stopCondition != null && stopCondition(item))
                    return false;
                result.Add(item);
                return true;
            }
        );

        await pageIterator.IterateAsync();
        return result;
    }
}
