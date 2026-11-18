using System.Text;
using OutlookAI.Core.Services;
using OutlookAI.Core.AddIn;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// list_signatures' answer (Q147, Q155 (c), Q159): every profile named, with the default and the
/// running one marked; each account's own and effective defaults and whether manage_signature may
/// change them; and, asked by name, one signature in full - every rendition and every picture.
/// </summary>
public sealed class SignatureListingTests : IDisposable
{
    private readonly TempSignatureFolder _folder = new TempSignatureFolder();

    public void Dispose() => _folder.Dispose();

    [Fact]
    public void EveryProfileIsListed_TheDefaultAndTheRunningOneMarked_EachWithItsOwnAccounts()
    {
        FakeProfileAccountRegistry accounts = new FakeProfileAccountRegistry()
            .Profile("Corpus", isDefault: true)
            .Add("Corpus", 1, FakeProfileAccountRegistry.DataFile("Corpus A"))
            .Add("Tier", 1, FakeProfileAccountRegistry.Pop3("tier@vm.invalid", newSignature: "Home", accountName: "OutlookAI tier sink"))
            .Add("Tier", 2, FakeProfileAccountRegistry.Exchange("jan@example.com"));
        _folder.Write("Home.htm", "<p>Home</p>", Encoding.ASCII);

        SignaturesOutcome outcome = SignatureListing.Build(Environment(accounts), "tier", null, null);

        Assert.Equal(new[] { "Corpus", "Tier" }, outcome.Profiles.Select(p => p.Name));
        Assert.True(outcome.Profiles[0].Default);
        Assert.Null(outcome.Profiles[0].Running);
        Assert.Empty(outcome.Profiles[0].Accounts);
        Assert.True(outcome.Profiles[1].Running);
        Assert.Equal("Tier", outcome.RunningProfile);
        Assert.Equal("Corpus", outcome.DefaultProfile);

        SignatureAccountView tier = outcome.Profiles[1].Accounts[0];
        Assert.Equal("tier@vm.invalid", tier.Account);
        Assert.Equal("pop3", tier.Kind);
        Assert.Equal("OutlookAI tier sink", tier.AccountName);
        Assert.Equal("Home", tier.NewMessage);
        Assert.Equal("Home", tier.EffectiveNewMessage);
        Assert.Equal("account", tier.DefaultsFrom);
        Assert.True(tier.DefaultsEditable);
        Assert.Null(outcome.Profiles[1].Accounts[1].AccountName);

        SignatureView home = Assert.Single(outcome.Signatures);
        Assert.Equal(new[] { "html" }, home.Formats);
        SignatureDefaultUseView use = Assert.Single(home.DefaultFor!);
        Assert.Equal(("Tier", "tier@vm.invalid", "new"), (use.Profile, use.Account, use.Scope));
        Assert.Contains("\"none\"", outcome.Note, StringComparison.Ordinal);
        Assert.Equal("default", outcome.Folder.Source);
        Assert.Null(outcome.Policy);
    }

    [Fact]
    public void WithOutlookNotRunning_TheNoteSaysSo_AndNoProfileIsMarkedRunning()
    {
        SignaturesOutcome outcome = SignatureListing.Build(
            Environment(new FakeProfileAccountRegistry().Add("Outlook", 1, FakeProfileAccountRegistry.Pop3("a@b.example"))),
            null, "Outlook is not running", null);

        Assert.Null(outcome.RunningProfile);
        Assert.Equal("Outlook is not running", outcome.RunningProfileNote);
        Assert.All(outcome.Profiles, p => Assert.Null(p.Running));
    }

    [Fact]
    public void PolicyAndRoaming_ShowWhyADefaultCannotBeChanged()
    {
        FakeProfileAccountRegistry accounts = new FakeProfileAccountRegistry()
            .Add("Outlook", 1, FakeProfileAccountRegistry.Pop3("pop@example.com", newSignature: "Own"))
            .Add("Outlook", 2, FakeProfileAccountRegistry.Exchange("jan@example.com"));

        SignaturesOutcome forced = SignatureListing.Build(
            Environment(accounts, new FakeSignatureSettings().Policy(SignaturePolicySource.GroupPolicy, newMessage: "Corporate")), "Outlook", null, null);
        SignatureAccountView pop = forced.Profiles[0].Accounts[0];
        Assert.Equal("Corporate", pop.EffectiveNewMessage);
        Assert.Equal("Own", pop.NewMessage);
        Assert.Equal("groupPolicy", pop.DefaultsFrom);
        Assert.False(pop.DefaultsEditable);
        Assert.Contains("Group Policy", pop.DefaultsLocked, StringComparison.Ordinal);
        Assert.Equal("Corporate", forced.Policy!.NewMessage);
        Assert.Equal("groupPolicy", forced.Policy.NewMessageFrom);

        SignaturesOutcome roaming = SignatureListing.Build(Environment(accounts, new FakeSignatureSettings().Microsoft365()), "Outlook", null, null);
        Assert.True(roaming.Profiles[0].Accounts[0].DefaultsEditable);
        Assert.False(roaming.Profiles[0].Accounts[1].DefaultsEditable);
        Assert.Contains("may roam", roaming.Profiles[0].Accounts[1].DefaultsLocked, StringComparison.Ordinal);
        Assert.Equal("possible", roaming.Roaming.State);

        // The testbed's LTSC 2024 with no copies of Outlook's own: the local files decide, every default editable.
        SignaturesOutcome ltsc = SignatureListing.Build(Environment(accounts, new FakeSignatureSettings().Ltsc2024()), "Outlook", null, null);
        Assert.All(ltsc.Profiles[0].Accounts, a => Assert.True(a.DefaultsEditable));
        Assert.Equal("off", ltsc.Roaming.State);
    }

    [Fact]
    public void ByName_OneSignatureInFull_EveryRenditionAndEveryPicture()
    {
        byte[] logo = TestPictures.Png(120, 40, 4);
        byte[] embedded = TestPictures.Png(2, 2, 6);
        Encoding w1252 = CodePage(1252);
        _folder.Write("My Sig.htm",
            "<html><head><meta charset=windows-1252><link rel=File-List href=\"My%20Sig_files/filelist.xml\"></head>"
            + "<body><p>Café</p><img src=\"My%20Sig_files/image001.png\"><img src=\"My%20Sig_files/gone.png\"></body></html>", w1252);
        _folder.WriteBytes(Path.Combine("My Sig_files", "image001.png"), logo);
        _folder.WriteBytes(Path.Combine("My Sig_files", "unused.png"), TestPictures.Png(1, 1));
        _folder.Write(Path.Combine("My Sig_files", "filelist.xml"), "<xml/>", Encoding.ASCII);
        _folder.Write("My Sig.rtf", "{\\rtf1 Caf\\'e9{\\pict\\picwgoal1800\\pichgoal600\\pngblip " + TestPictures.Hex(embedded) + "}}", Encoding.ASCII);
        _folder.Write("My Sig.txt", "Café", Encoding.Unicode);

        SignaturesOutcome outcome = SignatureListing.Build(Environment(new FakeProfileAccountRegistry()), null, null, "my sig");
        SignatureDetailView detail = outcome.Signature!;

        Assert.Equal("My Sig", detail.Name);
        Assert.Contains("Café", detail.Html, StringComparison.Ordinal);
        Assert.Equal("windows-1252", detail.HtmlEncoding);
        Assert.Equal("{\\rtf1 Caf\\'e9[[picture:1]]}", detail.Rtf);
        Assert.Equal("Café", detail.Text);
        Assert.Equal("utf-16", detail.TextEncoding);
        Assert.Equal("My%20Sig_files/", detail.PictureReference);

        Assert.Equal(new[] { "image001.png", "unused.png" }, detail.Pictures!.Select(p => p.File));
        SignaturePictureView picture = detail.Pictures![0];
        Assert.True(picture.ReferencedByHtml);
        Assert.Equal((120, 40), (picture.Width, picture.Height));
        Assert.Equal("image/png", picture.ContentType);
        Assert.Equal(logo.Length, picture.Bytes);
        Assert.True(File.Exists(picture.Path));
        Assert.False(detail.Pictures[1].ReferencedByHtml);
        Assert.Equal(new[] { "filelist.xml" }, detail.OtherFiles);
        Assert.Equal(new[] { "gone.png" }, detail.MissingPictures);

        SignatureRtfPictureView rtfPicture = Assert.Single(detail.RtfPictures!);
        Assert.Equal(1, rtfPicture.Number);
        Assert.Equal("png", rtfPicture.Kind);
        Assert.Equal(120, rtfPicture.Width);
        Assert.Equal(SignatureRtf.Sha256Hex(embedded), rtfPicture.Sha256);
    }

    [Fact]
    public void ByName_Unknown_IsRefused_ListingWhatIsInstalled()
    {
        _folder.Write("Only.txt", "x", Encoding.ASCII);

        ArgumentException refusal = Assert.Throws<ArgumentException>(
            () => SignatureListing.Build(Environment(new FakeProfileAccountRegistry()), null, null, "Missing"));

        Assert.Contains("Installed: Only.", refusal.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ARoamingCopy_IsNamedForItsAccount_AndMarkedManagedInTheMailbox_OnlyWhileRoamingIsPossible()
    {
        _folder.Write("Work (jan@example.com).htm", "<p>Work</p>", Encoding.ASCII);

        SignatureView roams = Assert.Single(SignatureListing.Build(
            Environment(new FakeProfileAccountRegistry(), new FakeSignatureSettings().Microsoft365()), null, null, null).Signatures);
        SignatureView off = Assert.Single(SignatureListing.Build(Environment(new FakeProfileAccountRegistry()), null, null, null).Signatures);

        Assert.Equal("jan@example.com", roams.Account);
        Assert.True(roams.ManagedInMailbox);
        Assert.Equal("jan@example.com", off.Account);
        Assert.Null(off.ManagedInMailbox);
    }

    private SignatureEnvironment Environment(FakeProfileAccountRegistry accounts, FakeSignatureSettings? settings = null)
    {
        return SignatureEnvironment.Read(settings ?? new FakeSignatureSettings(), accounts, "16.0", _folder.Root, _folder.Directory);
    }

    private static Encoding CodePage(int codePage)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        return Encoding.GetEncoding(codePage);
    }
}
