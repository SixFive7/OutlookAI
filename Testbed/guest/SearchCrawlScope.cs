#nullable disable
// =================================================================================================
// SearchCrawlScope.cs - hand-declared COM interop for the Windows Search Crawl Scope Manager: the
// documented writer of the crawl scope, and the API the Indexing Options dialog itself calls.
// =================================================================================================
//
// WHY IT EXISTS. The crawl scope lives under
// HKLM\SOFTWARE\Microsoft\Windows Search\CrawlScopeManager\Windows\SystemIndex, and on Windows 11
// 25H2 an elevated administrator holds ReadKey only there (measured 2026-09-24): the writer is the
// WSearch service, reached through ISearchManager -> ISearchCatalogManager ->
// ISearchCrawlScopeManager. These interfaces derive from IUnknown and ship no type library, so
// nothing can late-bind to them - not PowerShell, not VBScript - and they have to be declared by
// hand, IN EXACT VTABLE ORDER, with the right IIDs.
//
// THE RULE THIS FILE LIVES BY, AND WHY. In September 2026 this repository's MAPI interop declared
// IProfAdmin with an IID that named no interface at all, the QueryInterface failed with
// E_NOINTERFACE, and the failure was blamed on Office for a week (Testbed/guest/OutlookMapiInterop.ps1,
// REOPEN section). A wrong IID fails loudly; a wrong SLOT ORDER is worse, because it calls a
// DIFFERENT method with the same stack shape and can succeed at doing the wrong thing. So:
//
//   * every IID and CLSID below is copied from the Windows SDK's SearchAPI.h, version
//     10.0.26100.0, and carries the header line it came from;
//   * every interface declares EVERY method from the header, in header order, including the ones
//     nothing here calls - the CLR lays the vtable out in declaration order, so skipping one shifts
//     every method after it by a slot;
//   * each method carries its header line and its COM slot (0-2 are IUnknown);
//   * T1/SearchCrawlScopeInteropTests compiles THIS FILE (linked into the test project) and checks
//     every GUID and every slot by reflection against a table carrying the same header citations:
//     the CLR's own start and end slot for each interface, and each method's position in metadata
//     order, which is the order the CLR lays the vtable out in. Change an IID or reorder a method
//     and that test fails. Set-OutlookIndexingDisabled.ps1 -SelfTest repeats the check under
//     Windows PowerShell 5.1's compiler with Marshal.GetComSlotForMethodInfo, the per-method slot
//     the .NET Framework CLR will actually call.
//
// CONSTRAINTS. C# 5 ONLY: Windows PowerShell 5.1's Add-Type compiles this with the .NET Framework's
// own csc, which knows nothing newer - no string interpolation, no nameof, no expression bodies,
// no ?. , no out var. The loader strips the '#nullable' line above before compiling, because that
// compiler rejects the directive even inside an inactive #if (measured 2026-09-24); the test
// project compiles the file as it stands, where the directive keeps it out of nullable analysis.
// Every method is [PreserveSig] and returns the raw HRESULT, so S_FALSE is visible (enumerators end
// with it) and a failure is reported with the method's own name rather than as a bare COMException.
// BOOL is declared as int, because it is a 4-byte integer and saying so leaves nothing to a
// marshaller default.
//
// WHAT IT NEVER DOES. It touches no Outlook object, no MAPI and no mail item. The only state it can
// change is the crawl scope of the local SystemIndex catalog, and only through the helpers at the
// bottom, which the guest scripts call with -Execute.
// =================================================================================================

using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace OutlookAI.Testbed.SearchScope
{
    /// <summary>Every GUID this file uses, with its SearchAPI.h (10.0.26100.0) line.</summary>
    public static class SearchApiGuids
    {
        public const string CLSID_CSearchManager = "7D096C5F-AC08-4f1f-BEB7-5C22C517CE39";      // SearchAPI.h:6057
        public const string CLSID_CSearchRoot = "30766BD2-EA1C-4F28-BF27-0B44E2F68DB7";         // SearchAPI.h:6065
        public const string IID_ISearchManager = "AB310581-AC80-11D1-8DF3-00C04FB6EF69";        // SearchAPI.h:5428
        public const string IID_ISearchCatalogManager = "AB310581-AC80-11D1-8DF3-00C04FB6EF50"; // SearchAPI.h:3763
        public const string IID_ISearchCrawlScopeManager = "AB310581-AC80-11D1-8DF3-00C04FB6EF55"; // SearchAPI.h:2720
        public const string IID_ISearchRoot = "04C18CCF-1F57-4CBD-88CC-3900F5195CE3";           // SearchAPI.h:2018
        public const string IID_IEnumSearchRoots = "AB310581-AC80-11D1-8DF3-00C04FB6EF52";      // SearchAPI.h:2333
        public const string IID_ISearchScopeRule = "AB310581-AC80-11D1-8DF3-00C04FB6EF53";      // SearchAPI.h:2467
        public const string IID_IEnumSearchScopeRules = "AB310581-AC80-11D1-8DF3-00C04FB6EF54"; // SearchAPI.h:2584
    }

    /// <summary>SearchAPI.h:2445-2449 _FOLLOW_FLAGS.</summary>
    public static class FollowFlags
    {
        public const uint IndexComplexUrls = 0x1;   // FF_INDEXCOMPLEXURLS
        public const uint SuppressIndexing = 0x2;   // FF_SUPPRESSINDEXING
    }

    /// <summary>SearchAPI.h:2697-2702 CLUSION_REASON.</summary>
    public static class ClusionReason
    {
        public const int UnknownScope = 0;  // CLUSIONREASON_UNKNOWNSCOPE
        public const int Default = 1;       // CLUSIONREASON_DEFAULT
        public const int User = 2;          // CLUSIONREASON_USER
        public const int GroupPolicy = 3;   // CLUSIONREASON_GROUPPOLICY
    }

    /// <summary>ISearchManager, SearchAPI.h:5428-5479. Slots 3-15.</summary>
    [ComImport, Guid(SearchApiGuids.IID_ISearchManager), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface ISearchManager
    {
        [PreserveSig] int GetIndexerVersionStr([MarshalAs(UnmanagedType.LPWStr)] out string ppszVersionString);   // :5432 slot 3
        [PreserveSig] int GetIndexerVersion(out uint pdwMajor, out uint pdwMinor);                                // :5435 slot 4
        [PreserveSig] int GetParameter([MarshalAs(UnmanagedType.LPWStr)] string pszName, out IntPtr ppValue);    // :5439 slot 5 (PROPVARIANT**)
        [PreserveSig] int SetParameter([MarshalAs(UnmanagedType.LPWStr)] string pszName, IntPtr pValue);         // :5443 slot 6 (const PROPVARIANT*)
        [PreserveSig] int get_ProxyName([MarshalAs(UnmanagedType.LPWStr)] out string ppszProxyName);             // :5447 slot 7
        [PreserveSig] int get_BypassList([MarshalAs(UnmanagedType.LPWStr)] out string ppszBypassList);           // :5450 slot 8
        [PreserveSig] int SetProxy(int sUseProxy, int fLocalByPassProxy, uint dwPortNumber,
            [MarshalAs(UnmanagedType.LPWStr)] string pszProxyName, [MarshalAs(UnmanagedType.LPWStr)] string pszByPassList); // :5453 slot 9
        [PreserveSig] int GetCatalog([MarshalAs(UnmanagedType.LPWStr)] string pszCatalog, out ISearchCatalogManager ppCatalogManager); // :5460 slot 10
        [PreserveSig] int get_UserAgent([MarshalAs(UnmanagedType.LPWStr)] out string ppszUserAgent);             // :5464 slot 11
        [PreserveSig] int put_UserAgent([MarshalAs(UnmanagedType.LPWStr)] string pszUserAgent);                  // :5467 slot 12
        [PreserveSig] int get_UseProxy(out int pUseProxy);                                                        // :5470 slot 13
        [PreserveSig] int get_LocalBypass(out int pfLocalBypass);                                                 // :5473 slot 14
        [PreserveSig] int get_PortNumber(out uint pdwPortNumber);                                                 // :5476 slot 15
    }

    /// <summary>ISearchCatalogManager, SearchAPI.h:3763-3857. Slots 3-28.</summary>
    [ComImport, Guid(SearchApiGuids.IID_ISearchCatalogManager), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface ISearchCatalogManager
    {
        [PreserveSig] int get_Name([MarshalAs(UnmanagedType.LPWStr)] out string pszName);                        // :3767 slot 3
        [PreserveSig] int GetParameter([MarshalAs(UnmanagedType.LPWStr)] string pszName, out IntPtr ppValue);    // :3770 slot 4
        [PreserveSig] int SetParameter([MarshalAs(UnmanagedType.LPWStr)] string pszName, IntPtr pValue);         // :3774 slot 5
        [PreserveSig] int GetCatalogStatus(out int pStatus, out int pPausedReason);                               // :3778 slot 6
        [PreserveSig] int Reset();                                                                                // :3782 slot 7
        [PreserveSig] int Reindex();                                                                              // :3784 slot 8
        [PreserveSig] int ReindexMatchingURLs([MarshalAs(UnmanagedType.LPWStr)] string pszPattern);              // :3786 slot 9
        [PreserveSig] int ReindexSearchRoot([MarshalAs(UnmanagedType.LPWStr)] string pszRootURL);                // :3789 slot 10
        [PreserveSig] int put_ConnectTimeout(uint dwConnectTimeout);                                              // :3792 slot 11
        [PreserveSig] int get_ConnectTimeout(out uint pdwConnectTimeout);                                         // :3795 slot 12
        [PreserveSig] int put_DataTimeout(uint dwDataTimeout);                                                    // :3798 slot 13
        [PreserveSig] int get_DataTimeout(out uint pdwDataTimeout);                                               // :3801 slot 14
        [PreserveSig] int NumberOfItems(out int plCount);                                                         // :3804 slot 15
        [PreserveSig] int NumberOfItemsToIndex(out int plIncrementalCount, out int plNotificationQueue, out int plHighPriorityQueue); // :3807 slot 16
        [PreserveSig] int URLBeingIndexed([MarshalAs(UnmanagedType.LPWStr)] out string pszUrl);                  // :3812 slot 17
        [PreserveSig] int GetURLIndexingState([MarshalAs(UnmanagedType.LPWStr)] string pszURL, out uint pdwState); // :3815 slot 18
        [PreserveSig] int GetPersistentItemsChangedSink(out IntPtr ppISearchPersistentItemsChangedSink);         // :3819 slot 19
        [PreserveSig] int RegisterViewForNotification([MarshalAs(UnmanagedType.LPWStr)] string pszView, IntPtr pViewChangedSink, out uint pdwCookie); // :3822 slot 20
        [PreserveSig] int GetItemsChangedSink(IntPtr pISearchNotifyInlineSite, ref Guid riid, out IntPtr ppv,
            out Guid pGUIDCatalogResetSignature, out Guid pGUIDCheckPointSignature, out uint pdwLastCheckPointNumber); // :3827 slot 21
        [PreserveSig] int UnregisterViewForNotification(uint dwCookie);                                           // :3835 slot 22
        [PreserveSig] int SetExtensionClusion([MarshalAs(UnmanagedType.LPWStr)] string pszExtension, int fExclude); // :3838 slot 23
        [PreserveSig] int EnumerateExcludedExtensions(out IntPtr ppExtensions);                                   // :3842 slot 24 (IEnumString**)
        [PreserveSig] int GetQueryHelper(out IntPtr ppSearchQueryHelper);                                         // :3845 slot 25
        [PreserveSig] int put_DiacriticSensitivity(int fDiacriticSensitive);                                      // :3848 slot 26
        [PreserveSig] int get_DiacriticSensitivity(out int pfDiacriticSensitive);                                 // :3851 slot 27
        [PreserveSig] int GetCrawlScopeManager(out ISearchCrawlScopeManager ppCrawlScopeManager);                 // :3854 slot 28
    }

    /// <summary>ISearchCrawlScopeManager, SearchAPI.h:2720-2784. Slots 3-18.</summary>
    [ComImport, Guid(SearchApiGuids.IID_ISearchCrawlScopeManager), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface ISearchCrawlScopeManager
    {
        [PreserveSig] int AddDefaultScopeRule([MarshalAs(UnmanagedType.LPWStr)] string pszURL, int fInclude, uint fFollowFlags); // :2724 slot 3
        [PreserveSig] int AddRoot(ISearchRoot pSearchRoot);                                                       // :2729 slot 4
        [PreserveSig] int RemoveRoot([MarshalAs(UnmanagedType.LPWStr)] string pszURL);                           // :2732 slot 5
        [PreserveSig] int EnumerateRoots(out IEnumSearchRoots ppSearchRoots);                                     // :2735 slot 6
        [PreserveSig] int AddHierarchicalScope([MarshalAs(UnmanagedType.LPWStr)] string pszURL, int fInclude, int fDefault, int fOverrideChildren); // :2738 slot 7
        [PreserveSig] int AddUserScopeRule([MarshalAs(UnmanagedType.LPWStr)] string pszURL, int fInclude, int fOverrideChildren, uint fFollowFlags); // :2744 slot 8
        [PreserveSig] int RemoveScopeRule([MarshalAs(UnmanagedType.LPWStr)] string pszRule);                     // :2750 slot 9
        [PreserveSig] int EnumerateScopeRules(out IEnumSearchScopeRules ppSearchScopeRules);                      // :2753 slot 10
        [PreserveSig] int HasParentScopeRule([MarshalAs(UnmanagedType.LPWStr)] string pszURL, out int pfHasParentRule); // :2756 slot 11
        [PreserveSig] int HasChildScopeRule([MarshalAs(UnmanagedType.LPWStr)] string pszURL, out int pfHasChildRule);   // :2760 slot 12
        [PreserveSig] int IncludedInCrawlScope([MarshalAs(UnmanagedType.LPWStr)] string pszURL, out int pfIsIncluded);  // :2764 slot 13
        [PreserveSig] int IncludedInCrawlScopeEx([MarshalAs(UnmanagedType.LPWStr)] string pszURL, out int pfIsIncluded, out int pReason); // :2768 slot 14
        [PreserveSig] int RevertToDefaultScopes();                                                                // :2773 slot 15
        [PreserveSig] int SaveAll();                                                                              // :2775 slot 16
        [PreserveSig] int GetParentScopeVersionId([MarshalAs(UnmanagedType.LPWStr)] string pszURL, out int plScopeId); // :2777 slot 17
        [PreserveSig] int RemoveDefaultScopeRule([MarshalAs(UnmanagedType.LPWStr)] string pszURL);               // :2781 slot 18
    }

    /// <summary>ISearchRoot, SearchAPI.h:2018-2088. Slots 3-24.</summary>
    [ComImport, Guid(SearchApiGuids.IID_ISearchRoot), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface ISearchRoot
    {
        [PreserveSig] int put_Schedule([MarshalAs(UnmanagedType.LPWStr)] string pszTaskArg);                     // :2022 slot 3
        [PreserveSig] int get_Schedule([MarshalAs(UnmanagedType.LPWStr)] out string ppszTaskArg);                // :2025 slot 4
        [PreserveSig] int put_RootURL([MarshalAs(UnmanagedType.LPWStr)] string pszURL);                          // :2028 slot 5
        [PreserveSig] int get_RootURL([MarshalAs(UnmanagedType.LPWStr)] out string ppszURL);                     // :2031 slot 6
        [PreserveSig] int put_IsHierarchical(int fIsHierarchical);                                                // :2034 slot 7
        [PreserveSig] int get_IsHierarchical(out int pfIsHierarchical);                                           // :2037 slot 8
        [PreserveSig] int put_ProvidesNotifications(int fProvidesNotifications);                                  // :2040 slot 9
        [PreserveSig] int get_ProvidesNotifications(out int pfProvidesNotifications);                             // :2043 slot 10
        [PreserveSig] int put_UseNotificationsOnly(int fUseNotificationsOnly);                                    // :2046 slot 11
        [PreserveSig] int get_UseNotificationsOnly(out int pfUseNotificationsOnly);                               // :2049 slot 12
        [PreserveSig] int put_EnumerationDepth(uint dwDepth);                                                     // :2052 slot 13
        [PreserveSig] int get_EnumerationDepth(out uint pdwDepth);                                                // :2055 slot 14
        [PreserveSig] int put_HostDepth(uint dwDepth);                                                            // :2058 slot 15
        [PreserveSig] int get_HostDepth(out uint pdwDepth);                                                       // :2061 slot 16
        [PreserveSig] int put_FollowDirectories(int fFollowDirectories);                                          // :2064 slot 17
        [PreserveSig] int get_FollowDirectories(out int pfFollowDirectories);                                     // :2067 slot 18
        [PreserveSig] int put_AuthenticationType(int authType);                                                   // :2070 slot 19 (AUTH_TYPE)
        [PreserveSig] int get_AuthenticationType(out int pAuthType);                                              // :2073 slot 20
        [PreserveSig] int put_User([MarshalAs(UnmanagedType.LPWStr)] string pszUser);                            // :2076 slot 21
        [PreserveSig] int get_User([MarshalAs(UnmanagedType.LPWStr)] out string ppszUser);                       // :2079 slot 22
        [PreserveSig] int put_Password([MarshalAs(UnmanagedType.LPWStr)] string pszValue);                       // :2082 slot 23
        [PreserveSig] int get_Password([MarshalAs(UnmanagedType.LPWStr)] out string ppszValue);                  // :2085 slot 24
    }

    /// <summary>IEnumSearchRoots, SearchAPI.h:2333-2350. Slots 3-6. Next is only ever called with celt = 1.</summary>
    [ComImport, Guid(SearchApiGuids.IID_IEnumSearchRoots), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IEnumSearchRoots
    {
        [PreserveSig] int Next(uint celt, out ISearchRoot rgelt, out uint pceltFetched);                          // :2337 slot 3
        [PreserveSig] int Skip(uint celt);                                                                        // :2342 slot 4
        [PreserveSig] int Reset();                                                                                // :2345 slot 5
        [PreserveSig] int Clone(out IEnumSearchRoots ppenum);                                                     // :2347 slot 6
    }

    /// <summary>ISearchScopeRule, SearchAPI.h:2467-2483. Slots 3-6.</summary>
    [ComImport, Guid(SearchApiGuids.IID_ISearchScopeRule), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface ISearchScopeRule
    {
        [PreserveSig] int get_PatternOrURL([MarshalAs(UnmanagedType.LPWStr)] out string ppszPatternOrURL);       // :2471 slot 3
        [PreserveSig] int get_IsIncluded(out int pfIsIncluded);                                                   // :2474 slot 4
        [PreserveSig] int get_IsDefault(out int pfIsDefault);                                                     // :2477 slot 5
        [PreserveSig] int get_FollowFlags(out uint pFollowFlags);                                                 // :2480 slot 6 (documented "Not supported"; E_NOTIMPL measured - never called)
    }

    /// <summary>IEnumSearchScopeRules, SearchAPI.h:2584-2601. Slots 3-6. Next is only ever called with celt = 1.</summary>
    [ComImport, Guid(SearchApiGuids.IID_IEnumSearchScopeRules), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IEnumSearchScopeRules
    {
        [PreserveSig] int Next(uint celt, out ISearchScopeRule pprgelt, out uint pceltFetched);                   // :2588 slot 3
        [PreserveSig] int Skip(uint celt);                                                                        // :2593 slot 4
        [PreserveSig] int Reset();                                                                                // :2596 slot 5
        [PreserveSig] int Clone(out IEnumSearchScopeRules ppenum);                                                // :2598 slot 6
    }

    /// <summary>One crawl-scope rule as the service reports it. No follow flags: Microsoft documents
    /// ISearchScopeRule::get_FollowFlags as "Not supported", and on Windows 11 25H2 it returns
    /// E_NOTIMPL (measured 2026-09-24) - so it is declared, for the slot order, and never called.</summary>
    public sealed class ScopeRuleInfo
    {
        public string Url;
        public bool Included;
        public bool IsDefault;
    }

    /// <summary>One search root as the service reports it.</summary>
    public sealed class SearchRootInfo
    {
        public string Url;
        public bool ProvidesNotifications;
        public bool IsHierarchical;
        public bool UseNotificationsOnly;
    }

    /// <summary>What the catalog says about its own work: the documented "is it done yet" counters.</summary>
    public sealed class CatalogCounters
    {
        public int Status;              // _CatalogStatus, SearchAPI.h:3720-3729 (0 = IDLE)
        public int PausedReason;        // _CatalogPausedReason, SearchAPI.h:3732-3745
        public int NumberOfItems;       // ISearchCatalogManager::NumberOfItems
        public int IncrementalCount;    // ISearchCatalogManager::NumberOfItemsToIndex, first out
        public int NotificationQueue;   //   second out
        public int HighPriorityQueue;   //   third out
        public string IndexerVersion;   // ISearchManager::GetIndexerVersionStr
    }

    /// <summary>The whole surface the guest scripts use. Every call opens its own manager, and every
    /// change is committed with SaveAll inside the same call - nothing is left pending in a manager
    /// somebody forgot to save.</summary>
    public static class CrawlScope
    {
        public const string CatalogName = "SystemIndex";
        public const int S_OK = 0;
        public const int S_FALSE = 1;

        static void Check(int hr, string what)
        {
            if (hr < 0) { throw new COMException(what + " failed with HRESULT 0x" + hr.ToString("X8"), hr); }
        }

        static void Release(object o)
        {
            if (o != null && Marshal.IsComObject(o)) { Marshal.ReleaseComObject(o); }
        }

        static ISearchManager OpenManager()
        {
            Type t = Type.GetTypeFromCLSID(new Guid(SearchApiGuids.CLSID_CSearchManager), true);
            return (ISearchManager)Activator.CreateInstance(t);
        }

        static ISearchCatalogManager OpenCatalog(ISearchManager manager)
        {
            ISearchCatalogManager catalog;
            Check(manager.GetCatalog(CatalogName, out catalog), "ISearchManager::GetCatalog(SystemIndex)");
            return catalog;
        }

        static ISearchCrawlScopeManager OpenScope(ISearchCatalogManager catalog)
        {
            ISearchCrawlScopeManager scope;
            Check(catalog.GetCrawlScopeManager(out scope), "ISearchCatalogManager::GetCrawlScopeManager");
            return scope;
        }

        public static CatalogCounters ReadCounters()
        {
            ISearchManager manager = null;
            ISearchCatalogManager catalog = null;
            try
            {
                manager = OpenManager();
                CatalogCounters c = new CatalogCounters();
                string version;
                Check(manager.GetIndexerVersionStr(out version), "ISearchManager::GetIndexerVersionStr");
                c.IndexerVersion = version;
                catalog = OpenCatalog(manager);
                Check(catalog.GetCatalogStatus(out c.Status, out c.PausedReason), "ISearchCatalogManager::GetCatalogStatus");
                Check(catalog.NumberOfItems(out c.NumberOfItems), "ISearchCatalogManager::NumberOfItems");
                Check(catalog.NumberOfItemsToIndex(out c.IncrementalCount, out c.NotificationQueue, out c.HighPriorityQueue), "ISearchCatalogManager::NumberOfItemsToIndex");
                return c;
            }
            finally { Release(catalog); Release(manager); }
        }

        public static ScopeRuleInfo[] ReadRules()
        {
            ISearchManager manager = null;
            ISearchCatalogManager catalog = null;
            ISearchCrawlScopeManager scope = null;
            IEnumSearchScopeRules rules = null;
            List<ScopeRuleInfo> found = new List<ScopeRuleInfo>();
            try
            {
                manager = OpenManager();
                catalog = OpenCatalog(manager);
                scope = OpenScope(catalog);
                Check(scope.EnumerateScopeRules(out rules), "ISearchCrawlScopeManager::EnumerateScopeRules");
                while (true)
                {
                    ISearchScopeRule rule;
                    uint fetched;
                    int hr = rules.Next(1, out rule, out fetched);
                    Check(hr, "IEnumSearchScopeRules::Next");
                    if (hr != S_OK || fetched == 0 || rule == null) { break; }
                    try
                    {
                        ScopeRuleInfo info = new ScopeRuleInfo();
                        int included, isDefault;
                        Check(rule.get_PatternOrURL(out info.Url), "ISearchScopeRule::get_PatternOrURL");
                        Check(rule.get_IsIncluded(out included), "ISearchScopeRule::get_IsIncluded");
                        Check(rule.get_IsDefault(out isDefault), "ISearchScopeRule::get_IsDefault");
                        info.Included = included != 0;
                        info.IsDefault = isDefault != 0;
                        found.Add(info);
                    }
                    finally { Release(rule); }
                }
                return found.ToArray();
            }
            finally { Release(rules); Release(scope); Release(catalog); Release(manager); }
        }

        public static SearchRootInfo[] ReadRoots()
        {
            ISearchManager manager = null;
            ISearchCatalogManager catalog = null;
            ISearchCrawlScopeManager scope = null;
            IEnumSearchRoots roots = null;
            List<SearchRootInfo> found = new List<SearchRootInfo>();
            try
            {
                manager = OpenManager();
                catalog = OpenCatalog(manager);
                scope = OpenScope(catalog);
                Check(scope.EnumerateRoots(out roots), "ISearchCrawlScopeManager::EnumerateRoots");
                while (true)
                {
                    ISearchRoot root;
                    uint fetched;
                    int hr = roots.Next(1, out root, out fetched);
                    Check(hr, "IEnumSearchRoots::Next");
                    if (hr != S_OK || fetched == 0 || root == null) { break; }
                    try
                    {
                        SearchRootInfo info = new SearchRootInfo();
                        int notifications, hierarchical, notificationsOnly;
                        Check(root.get_RootURL(out info.Url), "ISearchRoot::get_RootURL");
                        Check(root.get_ProvidesNotifications(out notifications), "ISearchRoot::get_ProvidesNotifications");
                        Check(root.get_IsHierarchical(out hierarchical), "ISearchRoot::get_IsHierarchical");
                        Check(root.get_UseNotificationsOnly(out notificationsOnly), "ISearchRoot::get_UseNotificationsOnly");
                        info.ProvidesNotifications = notifications != 0;
                        info.IsHierarchical = hierarchical != 0;
                        info.UseNotificationsOnly = notificationsOnly != 0;
                        found.Add(info);
                    }
                    finally { Release(root); }
                }
                return found.ToArray();
            }
            finally { Release(roots); Release(scope); Release(catalog); Release(manager); }
        }

        /// <summary>{ included (0/1), CLUSION_REASON } for a URL, as the service decides it -
        /// policy, user and default rules all applied.</summary>
        public static int[] IncludedInCrawlScope(string url)
        {
            ISearchManager manager = null;
            ISearchCatalogManager catalog = null;
            ISearchCrawlScopeManager scope = null;
            try
            {
                manager = OpenManager();
                catalog = OpenCatalog(manager);
                scope = OpenScope(catalog);
                int included, reason;
                Check(scope.IncludedInCrawlScopeEx(url, out included, out reason), "ISearchCrawlScopeManager::IncludedInCrawlScopeEx");
                return new int[] { included, reason };
            }
            finally { Release(scope); Release(catalog); Release(manager); }
        }

        /// <summary>Adds a search root for the URL (when it has none) and a USER scope rule, then
        /// SaveAll. The shape a non-elevated Outlook leaves when it registers itself (measured on
        /// OAI-INDEXED, 2026-09-24): a root reporting ProvidesNotifications and IsHierarchical, and a
        /// user rule (IsDefault false) - the same registry values the maintainer's workstation holds
        /// for mapi16://{SID}/. include = false writes an EXCLUDE rule instead and adds no root.</summary>
        public static void SetUserRule(string url, bool include, bool addRootIfMissing, uint followFlags)
        {
            ISearchManager manager = null;
            ISearchCatalogManager catalog = null;
            ISearchCrawlScopeManager scope = null;
            ISearchRoot root = null;
            try
            {
                manager = OpenManager();
                catalog = OpenCatalog(manager);
                scope = OpenScope(catalog);
                if (include && addRootIfMissing && !HasRoot(scope, url))
                {
                    Type t = Type.GetTypeFromCLSID(new Guid(SearchApiGuids.CLSID_CSearchRoot), true);
                    root = (ISearchRoot)Activator.CreateInstance(t);
                    Check(root.put_RootURL(url), "ISearchRoot::put_RootURL");
                    Check(root.put_IsHierarchical(1), "ISearchRoot::put_IsHierarchical");
                    Check(root.put_ProvidesNotifications(1), "ISearchRoot::put_ProvidesNotifications");
                    Check(scope.AddRoot(root), "ISearchCrawlScopeManager::AddRoot");
                }
                Check(scope.AddUserScopeRule(url, include ? 1 : 0, 1, followFlags), "ISearchCrawlScopeManager::AddUserScopeRule");
                Check(scope.SaveAll(), "ISearchCrawlScopeManager::SaveAll");
            }
            finally { Release(root); Release(scope); Release(catalog); Release(manager); }
        }

        /// <summary>Removes the user rule for exactly this URL, then SaveAll. Returns the raw HRESULT
        /// of RemoveScopeRule so the caller can tell "removed" from "there was none" (S_FALSE).</summary>
        public static int RemoveUserRule(string url)
        {
            ISearchManager manager = null;
            ISearchCatalogManager catalog = null;
            ISearchCrawlScopeManager scope = null;
            try
            {
                manager = OpenManager();
                catalog = OpenCatalog(manager);
                scope = OpenScope(catalog);
                int hr = scope.RemoveScopeRule(url);
                Check(hr, "ISearchCrawlScopeManager::RemoveScopeRule");
                Check(scope.SaveAll(), "ISearchCrawlScopeManager::SaveAll");
                return hr;
            }
            finally { Release(scope); Release(catalog); Release(manager); }
        }

        /// <summary>ISearchCatalogManager::Reset - the documented "rebuild the index" call: the
        /// catalog is thrown away and everything in the crawl scope is crawled again.</summary>
        public static void ResetCatalog()
        {
            ISearchManager manager = null;
            ISearchCatalogManager catalog = null;
            try
            {
                manager = OpenManager();
                catalog = OpenCatalog(manager);
                Check(catalog.Reset(), "ISearchCatalogManager::Reset");
            }
            finally { Release(catalog); Release(manager); }
        }

        static bool HasRoot(ISearchCrawlScopeManager scope, string url)
        {
            IEnumSearchRoots roots = null;
            try
            {
                Check(scope.EnumerateRoots(out roots), "ISearchCrawlScopeManager::EnumerateRoots");
                while (true)
                {
                    ISearchRoot root;
                    uint fetched;
                    int hr = roots.Next(1, out root, out fetched);
                    Check(hr, "IEnumSearchRoots::Next");
                    if (hr != S_OK || fetched == 0 || root == null) { return false; }
                    try
                    {
                        string rootUrl;
                        Check(root.get_RootURL(out rootUrl), "ISearchRoot::get_RootURL");
                        if (string.Equals(rootUrl, url, StringComparison.OrdinalIgnoreCase)) { return true; }
                    }
                    finally { Release(root); }
                }
            }
            finally { Release(roots); }
        }
    }
}
