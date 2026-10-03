using System.Reflection;
using System.Runtime.CompilerServices;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// The product members a <c>Writes=Nothing</c> live test may call - each one CHECKED, not trusted:
/// <see cref="WritePathAnalyzer"/> stops its walk at the product boundary and asks this list, and
/// <c>ReadOnlyLiveTestTests.EveryListedProductMember_ReachesNoWriteInsideTheProduct</c> walks INSIDE the
/// product from every entry here and fails if any of them can reach an audit-log append (the product
/// audits every write it makes), a late-bound COM mutation, a registry write or a file written, moved or
/// deleted.
/// <para>
/// <b>Two kinds of entry.</b> A whole TYPE, for pure data and pure functions - request and result
/// shapes, value objects, parsers, planners. And a single MEMBER (every overload of that name) of a type
/// that can also write, such as <c>MailService</c> or <c>OutlookComSession</c>. Starting or attaching to
/// Outlook counts as a read: it changes no data, and these tests have always been allowed to do it
/// (S7/D17). A member a <c>Writes=Nothing</c> test reaches that is on neither list is a finding - a write,
/// or a read nobody has listed yet; adding it here is a claim about product code, so it says why, and the
/// in-product walk checks the claim.
/// </para>
/// </summary>
internal static class ReadOnlyProductApi
{
    private const string Data = "a data shape or pure function: no member can reach Outlook, the index or the disk";

    /// <summary>Product types every member of which changes nothing.</summary>
    internal static readonly IReadOnlyDictionary<string, string> Types = new Dictionary<string, string>(StringComparer.Ordinal)
    {
        // Snapshots handed back by the COM layer, COM-free once returned.
        ["OutlookAI.Core.Com.ComFolderInfo"] = Data,
        ["OutlookAI.Core.Com.ComFolderTree"] = Data,
        ["OutlookAI.Core.Com.ComOpenResult"] = Data,
        ["OutlookAI.Core.Com.ComStoreDetail"] = Data,
        ["OutlookAI.Core.Com.ComStoreInfo"] = Data,
        ["OutlookAI.Core.Com.ComTableDateKindProbe"] = Data,
        ["OutlookAI.Core.Com.ComTableSortAttempt"] = Data,
        ["OutlookAI.Core.Com.ComTableSortProbe"] = Data,
        ["OutlookAI.Core.Com.ComWalkedItem"] = Data,
        ["OutlookAI.Core.Com.HitLocationResult"] = Data,
        ["OutlookAI.Core.Com.LocateFailureAdvice"] = Data,

        // The non-creating special-folder resolver (Q84): reads a store's folder ids, never creates one.
        ["OutlookAI.Core.Com.SpecialFolders"] = "the Q84 non-creating resolver: reads a store's default-folder ids and never creates a folder",
        ["OutlookAI.Core.Com.ComSpecialFolderStore"] = "the resolver's view of one store: reads its type and its folder ids",

        // The Windows Search tier: SELECT statements against the system index, and the shapes around them.
        ["OutlookAI.Core.IndexSearch.IndexSearchService"] = "queries the Windows Search index with SELECT statements; it cannot write anything",
        ["OutlookAI.Core.IndexSearch.IndexClientFactory"] = "opens an index connection (OLE DB or ADO) for SELECT statements",
        ["OutlookAI.Core.IndexSearch.IIndexClient"] = "runs a SELECT statement against the index",
        ["OutlookAI.Core.IndexSearch.IndexHit"] = Data,
        ["OutlookAI.Core.IndexSearch.IndexQuery"] = Data,
        ["OutlookAI.Core.IndexSearch.IndexRowFilter"] = Data,
        ["OutlookAI.Core.IndexSearch.IndexRowMapper"] = Data,
        ["OutlookAI.Core.IndexSearch.IndexSearchResult"] = Data,
        ["OutlookAI.Core.IndexSearch.IndexStalenessReport"] = Data,
        ["OutlookAI.Core.IndexSearch.StoreScopeInfo"] = Data,
        ["OutlookAI.Core.IndexSearch.WsSqlBuilder"] = Data,
        ["OutlookAI.Core.Mapi.MapiItemUrl"] = Data,

        // The service layer's request and result shapes.
        ["OutlookAI.Core.Services.AccountView"] = Data,
        ["OutlookAI.Core.Services.AccountsOutcome"] = Data,
        ["OutlookAI.Core.Services.AuditHealthView"] = Data,
        ["OutlookAI.Core.Services.ExhaustiveInfo"] = Data,
        ["OutlookAI.Core.Services.FolderScopeResolution"] = Data,
        ["OutlookAI.Core.Services.FolderScopeResolver"] = Data,
        ["OutlookAI.Core.Services.FolderView"] = Data,
        ["OutlookAI.Core.Services.FoldersOutcome"] = Data,
        ["OutlookAI.Core.Services.HealthOutcome"] = Data,
        ["OutlookAI.Core.Services.HitSummary"] = Data,
        ["OutlookAI.Core.Services.IndexHealthView"] = Data,
        ["OutlookAI.Core.Services.OutlookHealthView"] = Data,
        ["OutlookAI.Core.Services.ReadOutcome"] = Data,
        ["OutlookAI.Core.Services.ScanPositionInfo"] = Data,
        ["OutlookAI.Core.Services.SearchOutcome"] = Data,
        ["OutlookAI.Core.Services.SearchRequest"] = Data,
        ["OutlookAI.Core.Services.SearchScopeInfo"] = Data,
        ["OutlookAI.Core.Services.StoreFoldersView"] = Data,
        ["OutlookAI.Core.Services.StoreStaleness"] = Data,
        ["OutlookAI.Core.Services.StoreView"] = Data,
        ["OutlookAI.Core.Services.SweepInfo"] = Data,
        ["OutlookAI.Core.Services.ThreadOutcome"] = Data,
        ["OutlookAI.Core.Services.TuningHealthView"] = Data,

        // The corpus freshness check the Phase-1 fixture runs (LiveCorpusFreshness): parses a manifest
        // file and plans against it. Planning only - the corpus tool's write commands live elsewhere.
        ["OutlookAI.RemediationTools.CorpusFreshness"] = Data,
        ["OutlookAI.RemediationTools.CorpusMailboxOwner"] = Data,
        ["OutlookAI.RemediationTools.CorpusManifest"] = "parses and formats a manifest the caller read; writes nothing",
        ["OutlookAI.RemediationTools.CorpusManifestHeader"] = Data,
        ["OutlookAI.RemediationTools.CorpusPlan"] = "plans a corpus in memory; builds nothing",
        ["OutlookAI.RemediationTools.CorpusPlanOptions"] = Data,
        ["OutlookAI.RemediationTools.CorpusPlanReport"] = Data,
    };

    /// <summary>Single product members (every overload of the name) that change nothing.</summary>
    internal static readonly IReadOnlyDictionary<string, string> Members = new Dictionary<string, string>(StringComparer.Ordinal)
    {
        ["OutlookAI.Core.Com.OutlookComSession.Connect"] = "attaches to Outlook, starting it headless if it is not running (S7/D17)",
        ["OutlookAI.Core.Com.OutlookComSession.Dispose"] = "releases COM references; closes only the invisible pin Explorer this session itself opened (D49)",
        ["OutlookAI.Core.Com.OutlookComSession.IsComCallFailure"] = "classifies an exception",
        ["OutlookAI.Core.Com.OutlookComSession.GetStores"] = "reads the profile's stores",
        ["OutlookAI.Core.Com.OutlookComSession.GetStoreDetails"] = "reads the profile's stores and their Exchange type",
        ["OutlookAI.Core.Com.OutlookComSession.CountOutboxItems"] = "counts the items in each Outbox",
        ["OutlookAI.Core.Com.OutlookComSession.TryOpenItem"] = "opens an item by entry id (GetItemFromID) and reads it; never displays it",
        ["OutlookAI.Core.Com.OutlookComSession.TryGetAttachmentFileNames"] = "reads an item's attachment names",
        ["OutlookAI.Core.Com.OutlookComSession.WalkStoreMailItems"] = "walks a store's folders and reads their items",
        ["OutlookAI.Core.Com.OutlookComSession.FindFolderPathsByLeafName"] = "walks a store's folder tree by reading it",
        ["OutlookAI.Core.Com.OutlookComSession.ListFolders"] = "walks a folder tree by reading it (no folder is created since Q84)",
        ["OutlookAI.Core.Com.OutlookComSession.ProbeTableSort"] = "reads one folder through Folder.GetTable with each sort it is asked to try",
        ["OutlookAI.Core.Com.OutlookComSession.ProbeTableDateKind"] = "reads one folder's dates through GetTable and through the items",
        ["OutlookAI.Core.Com.OutlookLiveness.Probe"] = "asks Windows whether Outlook's window answers; touches no COM",
        ["OutlookAI.Core.Com.OutlookLiveness.Describe"] = "formats a liveness reading",
        ["OutlookAI.Core.Com.HitLocator.Locate"] = "finds the item an index row describes by reading folders and items",
        ["OutlookAI.RemediationTools.CorpusReanchor.DeriveAppliedShift"] = "computes a time shift from two dates; the re-anchor that writes is elsewhere",

        ["OutlookAI.Core.Services.MailService.CreateDefault"] = "builds the service and its gateway; attaches to nothing yet",
        ["OutlookAI.Core.Services.MailService.Dispose"] = "releases the gateway's COM session",
        ["OutlookAI.Core.Services.MailService.Search"] = "the search tool: index query plus a read-only sweep; its cache is in memory",
        ["OutlookAI.Core.Services.MailService.ClearSweepCache"] = "empties the in-memory sweep cache",
        ["OutlookAI.Core.Services.MailService.Read"] = "the read tool: reads one item; reads UnRead and never sets it",
        ["OutlookAI.Core.Services.MailService.Thread"] = "the thread tool: reads a conversation",
        ["OutlookAI.Core.Services.MailService.ListAccounts"] = "the list_accounts tool: reads the profile's accounts and stores",
        ["OutlookAI.Core.Services.MailService.ListFolders"] = "the list_folders tool: reads a folder tree",
        ["OutlookAI.Core.Services.MailService.Health"] = "the outlook_health tool: reads state; its audit check opens the log for append and writes nothing",
    };

    /// <summary>
    /// What the in-product walk may meet and still call a read - each a reviewed case, keyed by the
    /// source member the call sits in and the late-bound member it calls.
    /// </summary>
    internal static readonly IReadOnlyDictionary<string, string> InProductExemptions = new Dictionary<string, string>(StringComparer.Ordinal)
    {
        ["OutlookAI.Core.Com.OutlookComSession.Dispose|Close"] =
            "closes the invisible pin Explorer the session itself opened, and only when the session started Outlook (D49): "
            + "a window it owns, never an item, a folder or the user's own window",
    };

    /// <summary>
    /// Test-side methods that name a write-capable MCP tool ONLY to refuse it before anything is sent -
    /// keyed by the analyzer's own description of the method (<c>Type.Member</c>), each with the reason
    /// the literal is a refusal and never a call. <c>EveryRefusalOnlyToolName_StillPointsAtRealCode</c>
    /// keeps every key on real code.
    /// </summary>
    internal static readonly IReadOnlyDictionary<string, string> RefusalOnlyToolNames = new Dictionary<string, string>(StringComparer.Ordinal)
    {
        ["McpStdioClient.DescribeAuditLogContact"] =
            "compares the tool name with 'discard_draft' to REFUSE, before sending, a call the server would answer by "
            + "writing its real audit log (Q86); the client never sends it",
    };

    /// <summary>The key a product member is listed under: its type's full name, a dot, its name.</summary>
    public static string KeyOf(MethodBase method)
    {
        return TypeKeyOf(method.DeclaringType) + "." + method.Name;
    }

    /// <summary>True when <paramref name="method"/> is on the list, by its type or by its name.</summary>
    public static bool IsReadOnly(MethodBase method)
    {
        return Types.ContainsKey(TypeKeyOf(method.DeclaringType)) || Members.ContainsKey(KeyOf(method));
    }

    /// <summary>
    /// The source member a compiled method belongs to: the method itself, or - for a lambda, a local
    /// function or a state machine - the member the compiler generated it from.
    /// </summary>
    public static string SourceKeyOf(MethodBase method)
    {
        Type? type = method.DeclaringType;
        string name = method.Name;
        while (type != null && type.IsDefined(typeof(CompilerGeneratedAttribute), false) && type.DeclaringType != null)
        {
            // A state machine's type is named after its method; a display class is not, and then the
            // lambda's own name carries it.
            if (name == "MoveNext" && type.Name.StartsWith("<", StringComparison.Ordinal))
            {
                name = type.Name;
            }

            type = type.DeclaringType;
        }

        int open = name.IndexOf('<');
        int close = name.IndexOf('>');
        if (open == 0 && close > 1)
        {
            name = name.Substring(1, close - 1);
        }

        return TypeKeyOf(type) + "." + name;
    }

    /// <summary>A type's full name with generic arity and nesting normalised, so one entry covers every instantiation.</summary>
    internal static string TypeKeyOf(Type? type)
    {
        if (type == null)
        {
            return "?";
        }

        Type definition = type.IsGenericType ? type.GetGenericTypeDefinition() : type;
        string name = definition.FullName ?? definition.Name;
        int tick = name.IndexOf('`');
        return (tick >= 0 ? name.Substring(0, tick) : name).Replace('+', '.');
    }
}
