using OutlookAI.Core.Com;
using OutlookAI.Core.IndexSearch;
using OutlookAI.Core.Mapi;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// T1 tests for the lookup-subject derivation (pure logic in HitLocator). Live probes on
/// this machine showed attachment (kind=document) entries carry COMBINED display
/// strings: System.Subject = "&lt;filename&gt; (&lt;parent subject&gt;)" and the
/// ItemPathDisplay tail = "&lt;parent subject&gt; : &lt;filename&gt;". Synthetic fixtures
/// only (S6).
/// </summary>
public sealed class HitLocatorSubjectTests
{
    private const string Sid = "{S-1-5-21-1111111111-2222222222-3333333333-1001}";
    private static readonly string StorePrefix = $"mapi16://{Sid}/alice@example.com($deadbeef)";

    private static IndexHit AttachmentHit(string? subject, string? itemPathDisplay, string fileName)
    {
        string parentUrl = $"{StorePrefix}/0/Inbox/{EntryIdCodecTests.SyntheticEncodedTail()}";
        string attachUrl = $"{parentUrl}/at={EntryIdCodec.EncodeBytes(new byte[] { 0x01, 0x02 })}:{fileName}";
        return IndexRowMapper.Map(new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
        {
            ["System.ItemUrl"] = attachUrl,
            ["System.Subject"] = subject,
            ["System.ItemPathDisplay"] = itemPathDisplay,
            ["System.Kind"] = new object[] { "document" },
        });
    }

    [Fact]
    public void DeriveLookupSubject_AttachmentHit_StripsPathTailDecoration()
    {
        IndexHit hit = AttachmentHit(
            subject: "report.pdf (RE: Project update: phase 2)",
            itemPathDisplay: "/alice@example.com/Inbox/RE: Project update: phase 2 : report.pdf",
            fileName: "report.pdf");

        Assert.Equal("RE: Project update: phase 2", HitLocator.DeriveLookupSubject(hit));
    }

    [Fact]
    public void DeriveLookupSubject_AttachmentHit_FallsBackToSubjectDecoration()
    {
        IndexHit hit = AttachmentHit(
            subject: "report.pdf (Quarterly figures)",
            itemPathDisplay: null,
            fileName: "report.pdf");

        Assert.Equal("Quarterly figures", HitLocator.DeriveLookupSubject(hit));
    }

    [Fact]
    public void DeriveLookupSubject_MessageHit_ReturnsSubjectVerbatim()
    {
        IndexHit hit = IndexRowMapper.Map(new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
        {
            ["System.ItemUrl"] = $"{StorePrefix}/0/Inbox/{EntryIdCodecTests.SyntheticEncodedTail()}",
            ["System.Subject"] = "plain subject (with parens) : and colon",
            ["System.Kind"] = new object[] { "email" },
        });

        Assert.Equal("plain subject (with parens) : and colon", HitLocator.DeriveLookupSubject(hit));
    }

    [Fact]
    public void DeriveLookupSubject_NoSubject_UsesItemPathDisplayTail()
    {
        IndexHit hit = IndexRowMapper.Map(new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
        {
            ["System.ItemUrl"] = $"{StorePrefix}/0/Inbox/{EntryIdCodecTests.SyntheticEncodedTail()}",
            ["System.ItemPathDisplay"] = "/alice@example.com/Inbox/derived name",
            ["System.Kind"] = new object[] { "email" },
        });

        Assert.Equal("derived name", HitLocator.DeriveLookupSubject(hit));
    }

    [Fact]
    public void TryMapUrlTarget_DelegateHit_TargetsDelegateStore()
    {
        IndexHit hit = IndexRowMapper.Map(new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
        {
            ["System.ItemUrl"] = $"{StorePrefix}/1/Bob Delegate/Inbox/{EntryIdCodecTests.SyntheticEncodedTail()}",
            ["System.Subject"] = "delegate mail",
            ["System.Kind"] = new object[] { "email" },
        });

        Assert.True(HitLocator.TryMapUrlTarget(hit, out string? store, out IReadOnlyList<string>? folders));
        Assert.Equal("Bob Delegate", store);
        Assert.Equal(new[] { "Inbox" }, folders);
    }
}
