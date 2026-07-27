using OutlookAI.Core.IndexSearch;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// T1 shape tests for the WS-SQL builder, including the v3.MD section-12 anti-patterns as
/// NEGATIVE tests: no LIKE on System.ItemPathDisplay, never SELECT System.Message.MessageId
/// / System.Search.Contents / System.Search.EntryID, no CONTAINS('*'), mapi-only SCOPE,
/// index-backed sender/recipient shapes (per-column CONTAINS - Phase-1 probes measured
/// =/LIKE on address columns at 1-10 s property-scan cost).
/// </summary>
public sealed class WsSqlBuilderTests
{
    private const string SyntheticScope = "mapi16://{S-1-5-21-1111111111-2222222222-3333333333-1001}/alice@example.com($deadbeef)";

    private static IndexQuery BaseQuery() => new()
    {
        Scope = SyntheticScope,
        Terms = new[] { "factuur" },
    };

    /// <summary>Everything after FROM SystemIndex WHERE - System.Kind is a legal SELECT column.</summary>
    private static string WhereClause(string sql)
    {
        const string marker = " FROM SystemIndex WHERE ";
        int cut = sql.IndexOf(marker, StringComparison.Ordinal);
        return cut < 0 ? sql : sql.Substring(cut + marker.Length);
    }

    [Fact]
    public void Build_DefaultShape_MatchesProvenProbeQuery()
    {
        string sql = WsSqlBuilder.Build(BaseQuery());

        Assert.StartsWith("SELECT TOP 25 ", sql, StringComparison.Ordinal);
        Assert.Contains(" FROM SystemIndex WHERE ", sql, StringComparison.Ordinal);
        Assert.Contains($"SCOPE='{SyntheticScope}'", sql, StringComparison.Ordinal);

        // Soak fix 16: under a mapi SCOPE the default (attachment-bearing) shape emits NO
        // Kind predicate at all. An attachment-content row carries the ATTACHMENT's kind,
        // so 'document' dropped 22.6% of them (picture / communication / calendar / music
        // / video); IndexRowFilter decides admission on the URL after the rows come back.
        // (System.Kind stays in the SELECT list - it is the filter's input.)
        Assert.DoesNotContain("System.Kind=", WhereClause(sql), StringComparison.Ordinal);
        Assert.Contains(
            "(CONTAINS(System.Subject, '\"factuur\"') OR CONTAINS(System.Search.Contents, '\"factuur\"'))",
            sql,
            StringComparison.Ordinal);
        Assert.EndsWith(" ORDER BY System.Message.DateReceived DESC", sql, StringComparison.Ordinal);
    }

    // --------------------------------------------- search_in scopes (D40 / SF-6, user 2026-07-26)

    [Fact]
    public void Build_DefaultSearchIn_IsSubjectAndBody()
    {
        Assert.Equal(SearchIn.SubjectAndBody, new IndexQuery().SearchIn);
        Assert.Equal(SearchIn.SubjectAndBody, SearchInValues.Default);
    }

    [Fact]
    public void Build_NeverEmitsUnqualifiedContains_Sf6RecallBug()
    {
        // SF-6 (measured 2026-07-26): a bare CONTAINS('term') searches
        // System.Search.Contents alone - the contents stream carries no subject text, so
        // mail whose term appears only in the subject was invisible (~3.4% of items
        // store-wide; the 138-item HAProxy alert population was the discovery case).
        // Every term predicate must name its column(s).
        foreach (SearchIn scope in new[] { SearchIn.SubjectAndBody, SearchIn.SubjectOnly, SearchIn.BodyOnly })
        {
            var query = BaseQuery();
            query.SearchIn = scope;
            query.Terms = new[] { "factuur", "betaling" };

            string sql = WsSqlBuilder.Build(query);

            Assert.DoesNotContain("CONTAINS('", sql, StringComparison.Ordinal);
            Assert.Contains("CONTAINS(System.", sql, StringComparison.Ordinal);
        }
    }

    [Fact]
    public void Build_SubjectOnlySearchIn_QueriesSubjectColumnAlone()
    {
        var query = BaseQuery();
        query.SearchIn = SearchIn.SubjectOnly;

        string sql = WsSqlBuilder.Build(query);

        Assert.Contains("CONTAINS(System.Subject, '\"factuur\"')", sql, StringComparison.Ordinal);
        Assert.DoesNotContain("System.Search.Contents", sql, StringComparison.Ordinal);
    }

    [Fact]
    public void Build_BodyOnlySearchIn_QueriesContentsColumnAlone()
    {
        var query = BaseQuery();
        query.SearchIn = SearchIn.BodyOnly;

        string sql = WsSqlBuilder.Build(query);

        Assert.Contains("CONTAINS(System.Search.Contents, '\"factuur\"')", sql, StringComparison.Ordinal);
        Assert.DoesNotContain("CONTAINS(System.Subject", sql, StringComparison.Ordinal);

        // Contents stays query-only - never in the SELECT list (section 12).
        int selectEnd = sql.IndexOf(" FROM SystemIndex", StringComparison.Ordinal);
        Assert.DoesNotContain("System.Search.Contents", sql.Substring(0, selectEnd), StringComparison.OrdinalIgnoreCase);
    }

    // ------------------------------------- cross-column AND (soak fix 13, user 2026-07-26)

    [Fact]
    public void Build_MultipleTerms_AndAcrossTheColumns_NotInsideOne()
    {
        // The shipped shape until soak fix 13 ANDed the terms INSIDE each column and ORed
        // the columns, so mail with one term only in the subject and another only in the
        // body matched nothing. Each term now gets its own Subject-OR-Contents pair.
        var query = BaseQuery();
        query.Terms = new[] { "balans", "energie" };

        string sql = WsSqlBuilder.Build(query);

        Assert.Contains(
            "(CONTAINS(System.Subject, '\"balans\"') OR CONTAINS(System.Search.Contents, '\"balans\"')) "
            + "AND (CONTAINS(System.Subject, '\"energie\"') OR CONTAINS(System.Search.Contents, '\"energie\"'))",
            sql,
            StringComparison.Ordinal);

        // The regressed shape must not come back.
        Assert.DoesNotContain("\"balans\" AND \"energie\"", sql, StringComparison.Ordinal);
    }

    [Fact]
    public void Build_ThreeTerms_EmitOnePairPerTerm()
    {
        var query = BaseQuery();
        query.Terms = new[] { "factuur", "betaling", "bedrag" };

        string sql = WsSqlBuilder.Build(query);

        foreach (string term in new[] { "factuur", "betaling", "bedrag" })
        {
            Assert.Contains(
                "(CONTAINS(System.Subject, '\"" + term + "\"') OR CONTAINS(System.Search.Contents, '\"" + term + "\"'))",
                sql,
                StringComparison.Ordinal);
        }

        // Three pairs, ANDed: the term predicate contributes exactly two ' AND ' joins on
        // top of the WHERE-level ones (scope, kind).
        Assert.Equal(3, CountOccurrences(sql, "CONTAINS(System.Subject"));
        Assert.Equal(3, CountOccurrences(sql, "CONTAINS(System.Search.Contents"));
    }

    [Fact]
    public void Build_SingleTerm_ShapeUnchangedByTheCrossColumnFix()
    {
        string sql = WsSqlBuilder.Build(BaseQuery());

        Assert.Contains(
            "(CONTAINS(System.Subject, '\"factuur\"') OR CONTAINS(System.Search.Contents, '\"factuur\"'))",
            sql,
            StringComparison.Ordinal);
        Assert.Equal(1, CountOccurrences(sql, "CONTAINS(System.Subject"));
    }

    [Fact]
    public void Build_NarrowedSearchIn_StaysSingleColumn_WithMultipleTerms()
    {
        // Narrowed scopes need no OR pair: an in-column AND is equivalent and cheaper.
        var subject = BaseQuery();
        subject.SearchIn = SearchIn.SubjectOnly;
        subject.Terms = new[] { "balans", "energie" };
        string subjectSql = WsSqlBuilder.Build(subject);

        Assert.Contains("CONTAINS(System.Subject, '\"balans\" AND \"energie\"')", subjectSql, StringComparison.Ordinal);
        Assert.DoesNotContain("System.Search.Contents", subjectSql, StringComparison.Ordinal);

        var body = BaseQuery();
        body.SearchIn = SearchIn.BodyOnly;
        body.Terms = new[] { "balans", "energie" };
        string bodySql = WsSqlBuilder.Build(body);

        Assert.Contains("CONTAINS(System.Search.Contents, '\"balans\" AND \"energie\"')", bodySql, StringComparison.Ordinal);
        Assert.DoesNotContain("CONTAINS(System.Subject", bodySql, StringComparison.Ordinal);
    }

    [Fact]
    public void Build_PrefixStar_SurvivesInBothColumns()
    {
        var query = BaseQuery();
        query.Terms = new[] { "factu*" };

        string sql = WsSqlBuilder.Build(query);

        Assert.Contains("CONTAINS(System.Subject, '\"factu*\"')", sql, StringComparison.Ordinal);
        Assert.Contains("CONTAINS(System.Search.Contents, '\"factu*\"')", sql, StringComparison.Ordinal);
    }

    [Fact]
    public void Build_PrefixStar_WorksInEveryTermPosition()
    {
        var query = BaseQuery();
        query.Terms = new[] { "bala*", "energie", "fact*" };

        string sql = WsSqlBuilder.Build(query);

        foreach (string term in new[] { "bala*", "energie", "fact*" })
        {
            Assert.Contains(
                "(CONTAINS(System.Subject, '\"" + term + "\"') OR CONTAINS(System.Search.Contents, '\"" + term + "\"'))",
                sql,
                StringComparison.Ordinal);
        }

        // Still no bare CONTAINS('*') anywhere (section 12: 0x80041605).
        Assert.DoesNotContain("CONTAINS('", sql, StringComparison.Ordinal);
        Assert.DoesNotContain("'\"*\"'", sql, StringComparison.Ordinal);
    }

    private static int CountOccurrences(string haystack, string needle)
    {
        int count = 0;
        int index = haystack.IndexOf(needle, StringComparison.Ordinal);
        while (index >= 0)
        {
            count++;
            index = haystack.IndexOf(needle, index + needle.Length, StringComparison.Ordinal);
        }

        return count;
    }

    [Fact]
    public void Build_EscapesSingleQuotesInBothColumns()
    {
        var query = BaseQuery();
        query.Terms = new[] { "o'brien" };

        string sql = WsSqlBuilder.Build(query);

        Assert.Contains("CONTAINS(System.Subject, '\"o''brien\"')", sql, StringComparison.Ordinal);
        Assert.Contains("CONTAINS(System.Search.Contents, '\"o''brien\"')", sql, StringComparison.Ordinal);
    }

    [Fact]
    public void Build_RejectsUnknownSearchIn()
    {
        var query = BaseQuery();
        query.SearchIn = (SearchIn)99;

        Assert.Throws<ArgumentException>(() => WsSqlBuilder.Build(query));
    }

    [Theory]
    [InlineData(null, SearchIn.SubjectAndBody)]
    [InlineData("", SearchIn.SubjectAndBody)]
    [InlineData("   ", SearchIn.SubjectAndBody)]
    [InlineData("subject_and_body", SearchIn.SubjectAndBody)]
    [InlineData(" Subject_And_Body ", SearchIn.SubjectAndBody)]
    [InlineData("subject", SearchIn.SubjectOnly)]
    [InlineData("SUBJECT", SearchIn.SubjectOnly)]
    [InlineData("subject_only", SearchIn.SubjectOnly)]
    [InlineData("body", SearchIn.BodyOnly)]
    [InlineData("Body", SearchIn.BodyOnly)]
    [InlineData("body_only", SearchIn.BodyOnly)]
    public void SearchInValues_Parse_AcceptsWireNamesAndDefaultsWhenOmitted(string? value, SearchIn expected)
    {
        Assert.Equal(expected, SearchInValues.Parse(value));
    }

    [Theory]
    [InlineData("all")]
    [InlineData("sender")]
    [InlineData("all_properties")]
    [InlineData("subject and body")]
    public void SearchInValues_Parse_RejectsUnknownValuesWithTheValidOnesNamed(string value)
    {
        ArgumentException ex = Assert.Throws<ArgumentException>(() => SearchInValues.Parse(value));

        foreach (string wireName in SearchInValues.WireNames)
        {
            Assert.Contains(wireName, ex.Message, StringComparison.Ordinal);
        }
    }

    [Fact]
    public void SearchInValues_WireNames_RoundTrip()
    {
        Assert.Equal(new[] { "subject_and_body", "subject", "body" }, SearchInValues.WireNames);
        foreach (string wireName in SearchInValues.WireNames)
        {
            Assert.Equal(wireName, SearchInValues.ToWireName(SearchInValues.Parse(wireName)));
        }
    }

    [Fact]
    public void Build_EmailOnly_WhenAttachmentHitsExcluded()
    {
        var query = BaseQuery();
        query.Kinds = KindFilter.EmailOnly;

        string sql = WsSqlBuilder.Build(query);

        Assert.Contains("System.Kind='email'", sql, StringComparison.Ordinal);
        Assert.DoesNotContain("System.Kind='document'", sql, StringComparison.Ordinal);
    }

    [Fact]
    public void Build_DocumentsOnly_EmitsNoKindPredicateUnderAScope()
    {
        // attachment_hits_only used to mean Kind='document', which is exactly the filter
        // that dropped 709 of 3,139 attachment rows. It now means "every /at= row",
        // decided by IndexRowFilter.
        var query = BaseQuery();
        query.Kinds = KindFilter.DocumentsOnly;

        string sql = WsSqlBuilder.Build(query);

        Assert.DoesNotContain("System.Kind=", WhereClause(sql), StringComparison.Ordinal);
    }

    [Fact]
    public void Build_Unscoped_EnumeratesTheAttachmentKinds_SoTheFileSystemStaysOut()
    {
        // With no SCOPE there is no namespace fence, so the provider still gets a kind
        // list - the measured union of message-level and attachment-row kinds.
        var query = BaseQuery();
        query.Scope = null;

        string sql = WsSqlBuilder.Build(query);

        foreach (string kind in IndexRowFilter.UnscopedKinds)
        {
            Assert.Contains($"System.Kind='{kind}'", sql, StringComparison.Ordinal);
        }

        Assert.Contains("System.Kind='picture'", sql, StringComparison.Ordinal);
        Assert.Contains("System.Kind='communication'", sql, StringComparison.Ordinal);
        Assert.Contains("System.Kind='calendar'", sql, StringComparison.Ordinal);
    }

    [Fact]
    public void Build_TopOverride_EmitsTheOverFetchedTop_NotTheRequestedOne()
    {
        var query = BaseQuery();
        query.Top = 26;

        Assert.StartsWith("SELECT TOP 62 ", WsSqlBuilder.Build(query, 62), StringComparison.Ordinal);
        Assert.StartsWith("SELECT TOP 26 ", WsSqlBuilder.Build(query, null), StringComparison.Ordinal);
        Assert.Throws<ArgumentException>(() => WsSqlBuilder.Build(query, 0));
        Assert.Throws<ArgumentException>(() => WsSqlBuilder.Build(query, WsSqlBuilder.MaxTop + 1));
    }

    [Fact]
    public void Build_NeverSelectsForbiddenColumns()
    {
        string sql = WsSqlBuilder.Build(BaseQuery());
        string selectList = sql.Substring(0, sql.IndexOf(" FROM SystemIndex", StringComparison.Ordinal));

        foreach (string forbidden in WsSqlBuilder.ForbiddenSelectColumns)
        {
            Assert.DoesNotContain(forbidden, selectList, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain(forbidden, WsSqlBuilder.SelectColumns, StringComparer.OrdinalIgnoreCase);
        }

        // MessageId and Search.EntryID are unusable anywhere (0x80040E55 / not a MAPI id);
        // Search.Contents is the one forbidden SELECT column that IS a legal CONTAINS
        // target - and since D40 the default term predicate queries it by name.
        Assert.DoesNotContain("System.Message.MessageId", sql, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("System.Search.EntryID", sql, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Build_NeverEmitsLike_AntiPattern()
    {
        var query = new IndexQuery
        {
            Scope = SyntheticScope,
            Terms = new[] { "factuur", "betaling" },
            FromAddressContains = "billing@example.com",
            RecipientContains = "alice@example.com",
            ReceivedOnOrAfterUtc = new DateTime(2026, 7, 1, 0, 0, 0, DateTimeKind.Utc),
            ReceivedBeforeUtc = new DateTime(2026, 7, 20, 0, 0, 0, DateTimeKind.Utc),
            IsRead = false,
            HasAttachments = true,
        };

        string sql = WsSqlBuilder.Build(query);

        // The 9-10 s property-scan anti-patterns (v3.MD sections 5/12 + Phase-1 probes).
        Assert.DoesNotContain("LIKE", sql, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("System.ItemPathDisplay LIKE", sql, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("FromAddress =", sql, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("FromAddress='", sql, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("ToAddress =", sql, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("ToAddress='", sql, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Build_AllFilters_ComposeIntoOneStatement()
    {
        var query = new IndexQuery
        {
            Scope = SyntheticScope,
            Terms = new[] { "factuur", "betaling" },
            FromAddressContains = "billing@example.com",
            RecipientContains = "alice@example.com",
            ReceivedOnOrAfterUtc = new DateTime(2026, 7, 1, 12, 30, 45, DateTimeKind.Utc),
            ReceivedBeforeUtc = new DateTime(2026, 7, 20, 0, 0, 0, DateTimeKind.Utc),
            IsRead = false,
            HasAttachments = true,
            Top = 100,
        };

        string sql = WsSqlBuilder.Build(query);

        Assert.StartsWith("SELECT TOP 100 ", sql, StringComparison.Ordinal);
        Assert.Contains(
            "(CONTAINS(System.Subject, '\"factuur\"') OR CONTAINS(System.Search.Contents, '\"factuur\"')) "
            + "AND (CONTAINS(System.Subject, '\"betaling\"') OR CONTAINS(System.Search.Contents, '\"betaling\"'))",
            sql,
            StringComparison.Ordinal);
        Assert.Contains("CONTAINS(System.Message.FromAddress, '\"billing@example.com\"')", sql, StringComparison.Ordinal);
        Assert.Contains("CONTAINS(System.Message.ToAddress, '\"alice@example.com\"')", sql, StringComparison.Ordinal);
        Assert.Contains("CONTAINS(System.Message.CcAddress, '\"alice@example.com\"')", sql, StringComparison.Ordinal);
        Assert.Contains("System.Message.DateReceived >= '2026-07-01 12:30:45'", sql, StringComparison.Ordinal);
        Assert.Contains("System.Message.DateReceived < '2026-07-20 00:00:00'", sql, StringComparison.Ordinal);
        Assert.Contains("System.IsRead=FALSE", sql, StringComparison.Ordinal);
        Assert.Contains("System.Message.HasAttachments=TRUE", sql, StringComparison.Ordinal);
    }

    [Fact]
    public void Build_SubjectAndBodySearchIn_UsesColumnScopedContains()
    {
        var query = BaseQuery();
        query.SearchIn = SearchIn.SubjectAndBody;

        string sql = WsSqlBuilder.Build(query);

        Assert.Contains("(CONTAINS(System.Subject, '\"factuur\"') OR CONTAINS(System.Search.Contents, '\"factuur\"'))", sql, StringComparison.Ordinal);
        // Contents may be queried but never selected.
        int selectEnd = sql.IndexOf(" FROM SystemIndex", StringComparison.Ordinal);
        Assert.DoesNotContain("System.Search.Contents", sql.Substring(0, selectEnd), StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("*")]
    [InlineData("  *  ")]
    public void Build_RejectsBareStarTerm_ContainsStarIsInvalid(string term)
    {
        var query = BaseQuery();
        query.Terms = new[] { term };

        Assert.Throws<ArgumentException>(() => WsSqlBuilder.Build(query));
    }

    [Theory]
    [InlineData("fac*tuur")]
    [InlineData("*factuur")]
    [InlineData("fac\"tuur")]
    [InlineData("fac;DROP TABLE x")]
    [InlineData("fac(tuur")]
    [InlineData("")]
    [InlineData("   ")]
    public void Build_RejectsUnsafeTerms(string term)
    {
        var query = BaseQuery();
        query.Terms = new[] { term };

        Assert.Throws<ArgumentException>(() => WsSqlBuilder.Build(query));
    }

    [Theory]
    [InlineData("file:///c:/temp")]
    [InlineData("csc://{S-1-5-21-1}/x")]
    [InlineData("/alice@example.com/Inbox")]
    [InlineData("")]
    public void Build_RejectsNonMapiScope(string scope)
    {
        var query = BaseQuery();
        query.Scope = scope;

        Assert.Throws<ArgumentException>(() => WsSqlBuilder.Build(query));
    }

    // ------------------------------------------------- escaping (soak fix 15, item 3)

    [Fact]
    public void Build_EscapesSingleQuotesInTheScope_InsteadOfThrowing()
    {
        // Was: ValidateScope THREW on any scope containing an apostrophe, which made a
        // folder named O'Brien un-searchable by hard exception. Measured rule: a raw '
        // is a syntax error (0x80040E14) in BOTH literal positions and '' parses in
        // both, so doubling is the fix.
        var query = BaseQuery();
        query.Scope = "mapi16://{sid}/store($ab12)/0/O'Brien";

        string sql = WsSqlBuilder.Build(query);

        Assert.Contains("SCOPE='mapi16://{sid}/store($ab12)/0/O''Brien'", sql, StringComparison.Ordinal);
        // Exactly one doubled quote - the literal is not double-escaped.
        Assert.DoesNotContain("O'''Brien", sql, StringComparison.Ordinal);
    }

    [Fact]
    public void Build_EscapesSingleQuotesInTheFolderPathLiteral()
    {
        var query = BaseQuery();
        query.Scope = "mapi16://{sid}/store($ab12)/0/O'Brien";
        query.FolderPathsAnyOf = new[] { "/store/O'Brien" };

        string sql = WsSqlBuilder.Build(query);

        Assert.Contains("System.ItemFolderPathDisplay='/store/O''Brien'", sql, StringComparison.Ordinal);
    }

    [Fact]
    public void Build_LeavesEveryOtherCharacterLiteralUnderEquality()
    {
        // Proved live: '=' is not LIKE - '%', '_', '[', ']', '{', '}', '"' are literal,
        // and a space must stay a space (%20 matches nothing, because the MAPI handler
        // already URL-encoded its URLs at index time).
        var query = BaseQuery();
        query.Scope = "mapi16://{sid}/store($ab12)/0/Sent Items";
        query.FolderPathsAnyOf = new[] { "/store/A_b%c [d] {e} \"f\" Sent Items" };

        string sql = WsSqlBuilder.Build(query);

        Assert.Contains("System.ItemFolderPathDisplay='/store/A_b%c [d] {e} \"f\" Sent Items'", sql, StringComparison.Ordinal);
        Assert.DoesNotContain("%20", sql, StringComparison.Ordinal);
        Assert.Contains("SCOPE='mapi16://{sid}/store($ab12)/0/Sent Items'", sql, StringComparison.Ordinal);
    }

    [Fact]
    public void Build_RejectsControlCharactersInTheScope()
    {
        var query = BaseQuery();
        query.Scope = "mapi16://{sid}/store" + (char)7 + "bell";

        Assert.Throws<ArgumentException>(() => WsSqlBuilder.Build(query));
    }

    // ------------------------------- non-recursive folder predicate (soak fix 15, item 4)

    [Fact]
    public void Build_EmitsFolderPathEquality_NotDirectory_ForANonRecursiveScope()
    {
        var query = BaseQuery();
        query.Scope = "mapi16://{sid}/alice@example.com($ab12)/0/Inbox";
        query.FolderPathsAnyOf = new[] { "/alice@example.com/Inbox" };

        string sql = WsSqlBuilder.Build(query);

        Assert.Contains("SCOPE='mapi16://{sid}/alice@example.com($ab12)/0/Inbox'", sql, StringComparison.Ordinal);
        Assert.Contains("System.ItemFolderPathDisplay='/alice@example.com/Inbox'", sql, StringComparison.Ordinal);

        // DIRECTORY= is shallow AND fast but returns ZERO Kind='document' rows - it drops
        // every attachment-content hit (41% of one real folder's rows). Never emitted.
        Assert.DoesNotContain("DIRECTORY", sql, StringComparison.OrdinalIgnoreCase);
        // ...and no Kind filter either, so attachment rows of EVERY kind survive the
        // narrowing (they inherit the parent's folder display path).
        Assert.DoesNotContain("System.Kind=", WhereClause(sql), StringComparison.Ordinal);
    }

    [Fact]
    public void Build_OrsMultipleFolderPaths_ForAFlatDelegateSubtree()
    {
        var query = BaseQuery();
        query.Scope = "mapi16://{sid}/host@example.com($ab12)/1/Someone Else";
        query.FolderPathsAnyOf = new[]
        {
            "/host@example.com/Someone Else/Inbox",
            "/host@example.com/Someone Else/20251015",
        };

        string sql = WsSqlBuilder.Build(query);

        Assert.Contains(
            "(System.ItemFolderPathDisplay='/host@example.com/Someone Else/Inbox' "
            + "OR System.ItemFolderPathDisplay='/host@example.com/Someone Else/20251015')",
            sql,
            StringComparison.Ordinal);
    }

    [Fact]
    public void Build_CollapsesDuplicateFolderPaths()
    {
        var query = BaseQuery();
        query.Scope = "mapi16://{sid}/host@example.com($ab12)/1/Someone Else";
        query.FolderPathsAnyOf = new[]
        {
            "/host@example.com/Someone Else/Conflicts",
            "/host@example.com/Someone Else/conflicts",
        };

        string sql = WsSqlBuilder.Build(query);

        Assert.Contains("System.ItemFolderPathDisplay='/host@example.com/Someone Else/Conflicts'", sql, StringComparison.Ordinal);
        Assert.DoesNotContain(" OR System.ItemFolderPathDisplay=", sql, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("alice@example.com/Inbox")] // no leading slash
    [InlineData("")]
    [InlineData("   ")]
    public void Build_RejectsMalformedFolderPaths(string path)
    {
        var query = BaseQuery();
        query.Scope = "mapi16://{sid}/alice@example.com($ab12)/0/Inbox";
        query.FolderPathsAnyOf = new[] { path };

        Assert.Throws<ArgumentException>(() => WsSqlBuilder.Build(query));
    }

    [Fact]
    public void Build_RefusesAnOverlongFolderPathOrSet()
    {
        // MEASURED: the provider executes a 95-literal OR-set and FAILS OUTRIGHT
        // ("Catastrophic failure", 0x8000FFFF) at 100. The builder's ceiling is the
        // last-resort guard; callers widen long before it.
        Assert.Equal(64, WsSqlBuilder.MaxFolderPaths);

        var query = BaseQuery();
        query.Scope = "mapi16://{sid}/host@example.com($ab12)/1/Someone Else";
        query.FolderPathsAnyOf = Enumerable
            .Range(0, WsSqlBuilder.MaxFolderPaths + 1)
            .Select(i => "/host@example.com/Someone Else/f" + i)
            .ToArray();

        Assert.Throws<ArgumentException>(() => WsSqlBuilder.Build(query));
    }

    [Fact]
    public void FolderScopeExistenceProbe_CarriesNoSearchFilters()
    {
        // The C7 guard must answer "does this folder scope resolve", never "does the
        // search match" - so it carries no term, date or sender predicate.
        string sql = WsSqlBuilder.BuildFolderScopeExistenceProbe(
            "mapi16://{sid}/host@example.com($ab12)/1/Someone Else",
            new[] { "/host@example.com/Someone Else/Inbox" });

        Assert.Contains("SELECT TOP 1 System.ItemUrl", sql, StringComparison.Ordinal);
        Assert.Contains("SCOPE='mapi16://{sid}/host@example.com($ab12)/1/Someone Else'", sql, StringComparison.Ordinal);
        Assert.Contains("System.ItemFolderPathDisplay='/host@example.com/Someone Else/Inbox'", sql, StringComparison.Ordinal);

        // No kind filter: the question is "does this folder scope resolve", and a folder
        // holding only meeting requests resolves just as well as one holding mail.
        Assert.DoesNotContain("System.Kind=", WhereClause(sql), StringComparison.Ordinal);
        Assert.DoesNotContain("CONTAINS", sql, StringComparison.Ordinal);
        Assert.DoesNotContain("DateReceived", sql, StringComparison.Ordinal);
        Assert.DoesNotContain("ORDER BY", sql, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    [InlineData(5001)]
    public void Build_RejectsTopOutOfRange(int top)
    {
        var query = BaseQuery();
        query.Top = top;

        Assert.Throws<ArgumentException>(() => WsSqlBuilder.Build(query));
    }

    [Fact]
    public void Build_RejectsInvertedDateRange()
    {
        var query = BaseQuery();
        query.ReceivedOnOrAfterUtc = new DateTime(2026, 7, 20, 0, 0, 0, DateTimeKind.Utc);
        query.ReceivedBeforeUtc = new DateTime(2026, 7, 1, 0, 0, 0, DateTimeKind.Utc);

        Assert.Throws<ArgumentException>(() => WsSqlBuilder.Build(query));
    }

    [Fact]
    public void Build_RejectsUnsafeSenderAndRecipientFilters()
    {
        var withBadSender = BaseQuery();
        withBadSender.FromAddressContains = "*";
        Assert.Throws<ArgumentException>(() => WsSqlBuilder.Build(withBadSender));

        var withBadRecipient = BaseQuery();
        withBadRecipient.RecipientContains = "a\"b";
        Assert.Throws<ArgumentException>(() => WsSqlBuilder.Build(withBadRecipient));
    }

    [Fact]
    public void Build_NoAggregatesOrJoinsEverEmitted()
    {
        string sql = WsSqlBuilder.Build(BaseQuery());

        Assert.DoesNotContain("COUNT(", sql, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain(" JOIN ", sql, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("GROUP BY", sql, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("DISTINCT", sql, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void BuildNewestReceivedProbe_ShapesWithAndWithoutScope()
    {
        string unscoped = WsSqlBuilder.BuildNewestReceivedProbe(null);
        Assert.Equal(
            "SELECT TOP 1 System.Message.DateReceived FROM SystemIndex WHERE System.Kind='email' ORDER BY System.Message.DateReceived DESC",
            unscoped);

        string scoped = WsSqlBuilder.BuildNewestReceivedProbe(SyntheticScope);
        Assert.Contains($"SCOPE='{SyntheticScope}' AND System.Kind='email'", scoped, StringComparison.Ordinal);
    }

    [Fact]
    public void BuildScopeExistenceProbe_ValidatesScope()
    {
        Assert.Throws<ArgumentException>(() => WsSqlBuilder.BuildScopeExistenceProbe("file:///x"));
        Assert.Equal(
            $"SELECT TOP 1 System.ItemUrl FROM SystemIndex WHERE SCOPE='{SyntheticScope}/1'",
            WsSqlBuilder.BuildScopeExistenceProbe(SyntheticScope + "/1"));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(100001)]
    public void BuildStoreDiscoverySample_RejectsOutOfRange(int top)
    {
        Assert.Throws<ArgumentException>(() => WsSqlBuilder.BuildStoreDiscoverySample(top));
    }

    [Fact]
    public void Build_TreatsUnspecifiedDateKindAsUtc()
    {
        var query = BaseQuery();
        query.ReceivedOnOrAfterUtc = new DateTime(2026, 7, 1, 6, 0, 0, DateTimeKind.Unspecified);

        string sql = WsSqlBuilder.Build(query);

        Assert.Contains("System.Message.DateReceived >= '2026-07-01 06:00:00'", sql, StringComparison.Ordinal);
    }
}
