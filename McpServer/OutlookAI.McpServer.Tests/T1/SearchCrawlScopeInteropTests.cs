using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using OutlookAI.Testbed.SearchScope;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins every IID, CLSID and vtable slot of <c>Testbed/guest/SearchCrawlScope.cs</c> - the
/// hand-declared interop the testbed uses to put Outlook into (and take it out of) the Windows
/// Search crawl scope - against the Windows SDK's <c>SearchAPI.h</c>, version 10.0.26100.0.
///
/// <para>
/// <b>Why this exists.</b> These interfaces derive from IUnknown and ship no type library, so the
/// declaration IS the contract and nothing checks it at run time except the call itself. In
/// September 2026 this repository's MAPI interop declared <c>IProfAdmin</c> with an IID that
/// named no interface; the QueryInterface failed with E_NOINTERFACE and the failure was blamed on
/// Office for a week. A wrong IID at least fails. A wrong SLOT does not: the CLR calls whatever
/// method sits at that vtable offset, and a method with the same stack shape "works" while doing
/// something else - on an API whose job is to change what a machine indexes.
/// </para>
///
/// <para>
/// <b>What it checks, and against what.</b> The table below is the header, transcribed with
/// line numbers; it is deliberately NOT derived from the interop file, so an edit to one without
/// the other fails. The vtable is read the way the CLR builds it: <see cref="Marshal.GetStartComSlot"/>
/// and <see cref="Marshal.GetEndComSlot"/> say where the CLR starts and ends the interface's
/// slots (3 and 3+n-1 for an IUnknown interface), and a [ComImport] interface is laid out in
/// metadata order from there - so the methods are compared in MetadataToken order, never in the
/// order reflection happens to list them (which is not guaranteed). Every declared method must
/// appear in the table and every table entry must be declared, so a skipped method (which would
/// shift every later slot by one) fails as loudly as a reordered one.
/// (<c>Marshal.GetComSlotForMethodInfo</c>, the per-method call, exists only on .NET Framework;
/// Set-OutlookIndexingDisabled.ps1 -SelfTest uses it there, under Windows PowerShell 5.1.)
/// </para>
/// </summary>
public sealed class SearchCrawlScopeInteropTests
{
    /// <summary>One vtable slot: the COM slot number (IUnknown takes 0-2), the method name in
    /// SearchAPI.h, and the header line its declaration starts on.</summary>
    public sealed record Slot(int Number, string Name, int HeaderLine);

    /// <summary>One interface as SearchAPI.h 10.0.26100.0 declares it.</summary>
    public sealed record Iface(Type Type, string Iid, int IidLine, Slot[] Slots);

    private static Slot[] Seq(int first, params (string Name, int Line)[] methods)
        => methods.Select((m, i) => new Slot(first + i, m.Name, m.Line)).ToArray();

    // ------------------------------------------------------------------ the header, transcribed

    private static readonly Iface[] Header =
    {
        new(typeof(ISearchManager), "AB310581-AC80-11D1-8DF3-00C04FB6EF69", 5428, Seq(3,
            ("GetIndexerVersionStr", 5432), ("GetIndexerVersion", 5435), ("GetParameter", 5439),
            ("SetParameter", 5443), ("get_ProxyName", 5447), ("get_BypassList", 5450),
            ("SetProxy", 5453), ("GetCatalog", 5460), ("get_UserAgent", 5464),
            ("put_UserAgent", 5467), ("get_UseProxy", 5470), ("get_LocalBypass", 5473),
            ("get_PortNumber", 5476))),

        new(typeof(ISearchCatalogManager), "AB310581-AC80-11D1-8DF3-00C04FB6EF50", 3763, Seq(3,
            ("get_Name", 3767), ("GetParameter", 3770), ("SetParameter", 3774),
            ("GetCatalogStatus", 3778), ("Reset", 3782), ("Reindex", 3784),
            ("ReindexMatchingURLs", 3786), ("ReindexSearchRoot", 3789), ("put_ConnectTimeout", 3792),
            ("get_ConnectTimeout", 3795), ("put_DataTimeout", 3798), ("get_DataTimeout", 3801),
            ("NumberOfItems", 3804), ("NumberOfItemsToIndex", 3807), ("URLBeingIndexed", 3812),
            ("GetURLIndexingState", 3815), ("GetPersistentItemsChangedSink", 3819),
            ("RegisterViewForNotification", 3822), ("GetItemsChangedSink", 3827),
            ("UnregisterViewForNotification", 3835), ("SetExtensionClusion", 3838),
            ("EnumerateExcludedExtensions", 3842), ("GetQueryHelper", 3845),
            ("put_DiacriticSensitivity", 3848), ("get_DiacriticSensitivity", 3851),
            ("GetCrawlScopeManager", 3854))),

        new(typeof(ISearchCrawlScopeManager), "AB310581-AC80-11D1-8DF3-00C04FB6EF55", 2720, Seq(3,
            ("AddDefaultScopeRule", 2724), ("AddRoot", 2729), ("RemoveRoot", 2732),
            ("EnumerateRoots", 2735), ("AddHierarchicalScope", 2738), ("AddUserScopeRule", 2744),
            ("RemoveScopeRule", 2750), ("EnumerateScopeRules", 2753), ("HasParentScopeRule", 2756),
            ("HasChildScopeRule", 2760), ("IncludedInCrawlScope", 2764),
            ("IncludedInCrawlScopeEx", 2768), ("RevertToDefaultScopes", 2773), ("SaveAll", 2775),
            ("GetParentScopeVersionId", 2777), ("RemoveDefaultScopeRule", 2781))),

        new(typeof(ISearchRoot), "04C18CCF-1F57-4CBD-88CC-3900F5195CE3", 2018, Seq(3,
            ("put_Schedule", 2022), ("get_Schedule", 2025), ("put_RootURL", 2028),
            ("get_RootURL", 2031), ("put_IsHierarchical", 2034), ("get_IsHierarchical", 2037),
            ("put_ProvidesNotifications", 2040), ("get_ProvidesNotifications", 2043),
            ("put_UseNotificationsOnly", 2046), ("get_UseNotificationsOnly", 2049),
            ("put_EnumerationDepth", 2052), ("get_EnumerationDepth", 2055), ("put_HostDepth", 2058),
            ("get_HostDepth", 2061), ("put_FollowDirectories", 2064), ("get_FollowDirectories", 2067),
            ("put_AuthenticationType", 2070), ("get_AuthenticationType", 2073), ("put_User", 2076),
            ("get_User", 2079), ("put_Password", 2082), ("get_Password", 2085))),

        new(typeof(IEnumSearchRoots), "AB310581-AC80-11D1-8DF3-00C04FB6EF52", 2333, Seq(3,
            ("Next", 2337), ("Skip", 2342), ("Reset", 2345), ("Clone", 2347))),

        new(typeof(ISearchScopeRule), "AB310581-AC80-11D1-8DF3-00C04FB6EF53", 2467, Seq(3,
            ("get_PatternOrURL", 2471), ("get_IsIncluded", 2474), ("get_IsDefault", 2477),
            ("get_FollowFlags", 2480))),

        new(typeof(IEnumSearchScopeRules), "AB310581-AC80-11D1-8DF3-00C04FB6EF54", 2584, Seq(3,
            ("Next", 2588), ("Skip", 2593), ("Reset", 2596), ("Clone", 2598))),
    };

    public static TheoryData<string> InterfaceNames()
    {
        TheoryData<string> data = new();
        foreach (Iface i in Header) data.Add(i.Type.Name);
        return data;
    }

    private static Iface Find(string name) => Header.Single(i => i.Type.Name == name);

    // ------------------------------------------------------------------ identities

    [Fact]
    public void Clsids_AreTheOnesSearchApiHDeclares()
    {
        // SearchAPI.h:6057 class DECLSPEC_UUID("7D096C5F-AC08-4f1f-BEB7-5C22C517CE39") CSearchManager
        Assert.Equal(new Guid("7D096C5F-AC08-4F1F-BEB7-5C22C517CE39"), new Guid(SearchApiGuids.CLSID_CSearchManager));
        // SearchAPI.h:6065 class DECLSPEC_UUID("30766BD2-EA1C-4F28-BF27-0B44E2F68DB7") CSearchRoot
        Assert.Equal(new Guid("30766BD2-EA1C-4F28-BF27-0B44E2F68DB7"), new Guid(SearchApiGuids.CLSID_CSearchRoot));
    }

    [Theory]
    [MemberData(nameof(InterfaceNames))]
    public void Iid_IsTheOneSearchApiHDeclares(string name)
    {
        Iface expected = Find(name);
        Assert.True(expected.Type.GUID == new Guid(expected.Iid),
            $"{name} carries IID {expected.Type.GUID:D}; SearchAPI.h:{expected.IidLine} declares {expected.Iid}. "
            + "A wrong IID is exactly the IProfAdmin bug (Testbed/guest/OutlookMapiInterop.ps1): the QueryInterface fails, and the failure reads as the component's.");
    }

    [Theory]
    [MemberData(nameof(InterfaceNames))]
    public void Interface_IsAnIUnknownComImport(string name)
    {
        Type t = Find(name).Type;
        Assert.True(t.IsInterface, $"{name} is not an interface");
        Assert.True(t.IsImport, $"{name} lost [ComImport]");
        InterfaceTypeAttribute? kind = t.GetCustomAttribute<InterfaceTypeAttribute>();
        Assert.NotNull(kind);
        Assert.Equal(ComInterfaceType.InterfaceIsIUnknown, kind!.Value);
        // IUnknown-derived with no base interface in the declaration: IUnknown's three slots are
        // implicit, which is what makes the first declared method slot 3.
        Assert.Empty(t.GetInterfaces());
    }

    [Theory]
    [MemberData(nameof(InterfaceNames))]
    public void EveryMethod_SitsInTheSlotSearchApiHGivesIt(string name)
    {
        Iface expected = Find(name);
        MethodInfo[] declared = expected.Type.GetMethods(BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly);

        string[] expectedNames = expected.Slots.Select(s => s.Name).ToArray();
        string[] declaredNames = declared.Select(m => m.Name).ToArray();
        string[] missing = expectedNames.Except(declaredNames).ToArray();
        string[] extra = declaredNames.Except(expectedNames).ToArray();
        Assert.True(missing.Length == 0 && extra.Length == 0,
            $"{name}: declared methods do not match SearchAPI.h. Missing: [{string.Join(", ", missing)}]. "
            + $"Not in the header: [{string.Join(", ", extra)}]. A missing method shifts every later slot by one.");
        Assert.Equal(expected.Slots.Length, declared.Length);

        // What the CLR itself reports about the interface's vtable: an IUnknown-based interface
        // starts at slot 3 (an IDispatch one would start at 7) and ends exactly where the header's
        // last method does - so no slot was added or dropped.
        int start = Marshal.GetStartComSlot(expected.Type);
        int end = Marshal.GetEndComSlot(expected.Type);
        Assert.True(start == 3, $"{name}: the CLR puts the first method in slot {start}, not 3 - it is not being laid out as an IUnknown interface");
        Assert.True(end == expected.Slots[^1].Number,
            $"{name}: the CLR's last slot is {end}; SearchAPI.h:{expected.Slots[^1].HeaderLine} ends the vtable at slot {expected.Slots[^1].Number}");

        // And which method is in which slot. The CLR lays a [ComImport] interface's vtable out in
        // metadata (declaration) order from the start slot, so a method's slot is its MethodDef
        // position plus the start slot - the order reflection happens to LIST methods in is not
        // guaranteed, which is why this sorts by metadata token rather than trusting it.
        MethodInfo[] inVtableOrder = declared.OrderBy(m => m.MetadataToken).ToArray();
        List<string> wrong = new();
        for (int i = 0; i < inVtableOrder.Length; i++)
        {
            Slot expectedSlot = expected.Slots[i];
            if (inVtableOrder[i].Name != expectedSlot.Name)
            {
                wrong.Add($"slot {start + i} holds {inVtableOrder[i].Name}; SearchAPI.h:{expectedSlot.HeaderLine} puts {expectedSlot.Name} there");
            }
        }

        Assert.True(wrong.Count == 0, $"{name}: " + string.Join("; ", wrong));
    }

    [Theory]
    [MemberData(nameof(InterfaceNames))]
    public void EveryMethod_ReturnsTheRawHResult(string name)
    {
        // [PreserveSig] + int: S_FALSE ends an enumeration and must stay visible, and a failure
        // is reported by the helper with the method's own name.
        foreach (MethodInfo m in Find(name).Type.GetMethods(BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly))
        {
            Assert.True(m.ReturnType == typeof(int), $"{name}.{m.Name} returns {m.ReturnType.Name}, not the HRESULT");
            Assert.True((m.MethodImplementationFlags & MethodImplAttributes.PreserveSig) != 0, $"{name}.{m.Name} lost [PreserveSig]");
        }
    }

    [Fact]
    public void HeaderTable_IsContiguousFromSlotThree()
    {
        // Guards the table itself: a gap or a duplicate here would let a wrong declaration pass.
        foreach (Iface i in Header)
        {
            Assert.Equal(Enumerable.Range(3, i.Slots.Length), i.Slots.Select(s => s.Number));
            Assert.Equal(i.Slots.Length, i.Slots.Select(s => s.Name).Distinct().Count());
            Assert.True(i.Slots.Zip(i.Slots.Skip(1), (a, b) => a.HeaderLine < b.HeaderLine).All(x => x),
                $"{i.Type.Name}: header lines must increase with the slot number");
            Assert.True(i.IidLine < i.Slots[0].HeaderLine, $"{i.Type.Name}: the IID line precedes the first method");
        }
    }

    [Fact]
    public void Constants_AreTheOnesSearchApiHDeclares()
    {
        // SearchAPI.h:2447-2448 FF_INDEXCOMPLEXURLS = 0x1, FF_SUPPRESSINDEXING = 0x2
        Assert.Equal(0x1u, FollowFlags.IndexComplexUrls);
        Assert.Equal(0x2u, FollowFlags.SuppressIndexing);
        // SearchAPI.h:2698-2701 CLUSIONREASON_UNKNOWNSCOPE .. CLUSIONREASON_GROUPPOLICY
        Assert.Equal(0, ClusionReason.UnknownScope);
        Assert.Equal(1, ClusionReason.Default);
        Assert.Equal(2, ClusionReason.User);
        Assert.Equal(3, ClusionReason.GroupPolicy);
        Assert.Equal("SystemIndex", CrawlScope.CatalogName);
    }
}
