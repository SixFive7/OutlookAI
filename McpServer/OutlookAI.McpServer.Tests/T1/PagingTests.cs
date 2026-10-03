using System.Text;
using OutlookAI.Core.Services;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// T1 for the paging every tool shares (Q119, 2026-10-03, <c>Core/Services/Paging.cs</c>): the offset
/// window <c>list_folders</c> and <c>read</c> cut their pages with, the fingerprint <c>search</c> and
/// <c>audit_log</c> bind a continuation to, and the stateless token <c>audit_log</c> hands out. The tools'
/// own contracts are pinned where they always were (<c>FolderWalkReportingTests</c>,
/// <c>PayloadDisciplineTests</c>, <c>BodyCacheTests</c>, <c>ResumableScanTests</c>, <c>AuditLogToolTests</c>);
/// this pins the shared rules under them.
/// </summary>
public sealed class PagingTests
{
    // ================================================================== the window

    [Theory]
    [InlineData(10, 0, 4, 0, 4, true, 4)]
    [InlineData(10, 4, 4, 4, 4, true, 8)]
    [InlineData(10, 8, 4, 8, 2, false, null)]
    [InlineData(10, 10, 4, 10, 0, false, null)]
    [InlineData(10, 12, 4, 10, 0, false, null)]
    [InlineData(10, -3, 4, 0, 4, true, 4)]
    [InlineData(10, 6, -1, 6, 0, true, 6)]
    [InlineData(0, 0, 4, 0, 0, false, null)]
    [InlineData(5, 0, int.MaxValue, 0, 5, false, null)]
    public void AWindow_ClampsItsOffset_AndSaysExplicitlyWhetherMoreFollows(
        int total, int offset, int size, int start, int count, bool hasMore, int? nextOffset)
    {
        PageWindow window = PageWindow.Of(total, offset, size);

        Assert.Equal(start, window.Start);
        Assert.Equal(count, window.Count);
        Assert.Equal(start + count, window.End);
        Assert.Equal(total, window.Total);
        Assert.Equal(hasMore, window.HasMore);
        Assert.Equal(nextOffset, window.NextOffset);
    }

    [Fact]
    public void ReadsBodyWindow_IsTheSharedWindow()
    {
        // read's body_offset contract, cut by the shared window: offset clamped, more = beyond the window.
        Assert.Equal((2, "cdef", true), MailService.ComputeBodyWindow("abcdefgh", 2, 4));
        Assert.Equal((8, string.Empty, false), MailService.ComputeBodyWindow("abcdefgh", 50, 4));
        Assert.Equal((0, string.Empty, true), MailService.ComputeBodyWindow("abcdefgh", -1, 0));
    }

    // ================================================================== the fingerprint

    [Fact]
    public void AFingerprint_TellsAnAbsentArgumentFromAnEmptyOne()
    {
        Assert.NotEqual(new PagingFingerprint().Add("a", null).ToString(), new PagingFingerprint().Add("a", string.Empty).ToString());
        Assert.NotEqual(
            new PagingFingerprint().AddList("l", null).ToString(),
            new PagingFingerprint().AddList("l", Array.Empty<string>()).ToString());
        Assert.Equal("a=0\nb=1|x\nl=1|2|p|q\n", new PagingFingerprint().Add("a", null).Add("b", "x").AddList("l", new[] { "p", "q" }).ToString());
    }

    [Fact]
    public void DifferingArguments_NamesWhatChanged_InFingerprintOrder()
    {
        string before = new PagingFingerprint().Add("after", "1").Add("operation", "send").Add("entry_id", null).ToString();
        string after = new PagingFingerprint().Add("after", "1").Add("operation", "move").Add("entry_id", "00AB").ToString();

        Assert.Equal(new[] { "operation", "entry_id" }, PagingFingerprint.DifferingArguments(before, after));
        Assert.Empty(PagingFingerprint.DifferingArguments(before, before));

        // The compact form keeps that, while carrying no argument value.
        Assert.Equal(
            new[] { "operation", "entry_id" },
            PagingFingerprint.DifferingArguments(PagingFingerprint.Compact(before), PagingFingerprint.Compact(after)));
        Assert.DoesNotContain("send", PagingFingerprint.Compact(before), StringComparison.Ordinal);
        Assert.Contains("entry_id=0", PagingFingerprint.Compact(before), StringComparison.Ordinal);
    }

    [Fact]
    public void SearchsFingerprint_IsBuiltByTheSharedBuilder_InItsOldShape()
    {
        // ExhaustiveScanCursors now delegates; the canonical text is what it always was.
        string fingerprint = ExhaustiveScanCursors.FingerprintOf(new SearchRequest { Store = "s", Folder = "Inbox" }, new[] { "t" });
        Assert.StartsWith("terms=1|1|t\nsearchIn=1|", fingerprint, StringComparison.Ordinal);
        Assert.Contains("\nstore=1|s\nfolder=1|Inbox\nincludeSubfolders=1|1\nafter=0\n", fingerprint, StringComparison.Ordinal);
    }

    // ================================================================== the stateless token

    [Fact]
    public void AToken_RoundTrips_ItsPosition_ForTheSameKindAndArguments()
    {
        string fingerprint = new PagingFingerprint().Add("operation", "send").ToString();
        string token = Paging.IssueToken("audit_log", fingerprint, new[] { "v1-2", "123", "cafe0123" });

        Assert.StartsWith(Paging.TokenPrefix, token, StringComparison.Ordinal);
        Assert.Equal(PageTokenDecision.Valid, Paging.ResolveToken(token, "audit_log", fingerprint, out IReadOnlyList<string> position, out _));
        Assert.Equal(new[] { "v1-2", "123", "cafe0123" }, position);

        // Nothing is consumed: the same token works again.
        Assert.Equal(PageTokenDecision.Valid, Paging.ResolveToken(token, "audit_log", fingerprint, out _, out _));
    }

    [Fact]
    public void AToken_IsRefused_ForOtherArguments_AnotherKind_OrAnyDamage()
    {
        string fingerprint = new PagingFingerprint().Add("operation", "send").Add("entry_id", null).ToString();
        string token = Paging.IssueToken("audit_log", fingerprint, new[] { "a", "1" });

        string other = new PagingFingerprint().Add("operation", "send").Add("entry_id", "00AB").ToString();
        Assert.Equal(PageTokenDecision.RequestChanged, Paging.ResolveToken(token, "audit_log", other, out _, out IReadOnlyList<string> changed));
        Assert.Equal(new[] { "entry_id" }, changed);

        Assert.Equal(PageTokenDecision.OtherKind, Paging.ResolveToken(token, "list_folders", fingerprint, out _, out _));

        // Any character changed, the end cut off, or not a token at all.
        char[] chars = token.ToCharArray();
        int middle = Paging.TokenPrefix.Length + 3;
        chars[middle] = chars[middle] == 'A' ? 'B' : 'A';
        Assert.Equal(PageTokenDecision.Malformed, Paging.ResolveToken(new string(chars), "audit_log", fingerprint, out _, out _));
        Assert.Equal(PageTokenDecision.Malformed, Paging.ResolveToken(token.Substring(0, token.Length - 2), "audit_log", fingerprint, out _, out _));
        Assert.Equal(PageTokenDecision.Malformed, Paging.ResolveToken("scan-0123456789abcdef0123456789abcdef", "audit_log", fingerprint, out _, out _));
        Assert.Equal(PageTokenDecision.Malformed, Paging.ResolveToken(null, "audit_log", fingerprint, out _, out _));
    }

    [Fact]
    public void AToken_CarriesNoArgumentValue()
    {
        const string Secret = "00000000AA11BB22CC33DD44EE55FF66AA11BB22CC33DD44EE55FF66";
        string token = Paging.IssueToken("audit_log", new PagingFingerprint().Add("entry_id", Secret).ToString(), new[] { "x" });

        string payload = token.Substring(Paging.TokenPrefix.Length, token.LastIndexOf('.') - Paging.TokenPrefix.Length)
            .Replace('-', '+').Replace('_', '/');
        payload = payload.PadRight(payload.Length + ((4 - (payload.Length % 4)) % 4), '=');
        Assert.DoesNotContain(Secret, Encoding.UTF8.GetString(Convert.FromBase64String(payload)), StringComparison.OrdinalIgnoreCase);
    }
}
