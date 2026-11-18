using System.Runtime.InteropServices;
using System.Text;
using OutlookAI.Core.AddIn;
using OutlookAI.Core.Services;
using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// The add-in's task pane, driven on a test guest (Q168 (b), decided 2026-11-18): the code-sharing refactor's
/// stage S4 moved the pane's reading and writing of a draft onto OutlookAI.Core's <see cref="ComposeDocument"/>
/// and its signatures onto Core's catalog and selection rule, and nothing else drives the pane. A draft the MCP
/// server opens in a compose window gets the pane, as any compose window does; the test types an instruction
/// into it and clicks "Edit current draft" through Win32 messages, and the pane's AI requests go to
/// <see cref="ClaudeCliStandIn"/>, which answers the body and a signature name the test chose. Then the open
/// window is read, read only, through <see cref="LiveOutlookTestMailer.ReadOpenDraft"/>.
/// <para>
/// What it measures, for Q165 (c): the pane still writes one paragraph a line (its layout, which Core's writer
/// keeps for it); the AI's own signature choice (Q162 (b)) follows the body it wrote and replaces the signature
/// the agent put in; the chosen signature's picture is embedded, where the pane used to leave a link to the
/// signature's folder (D47); and the signature request lists the signatures with Core's catalog.
/// </para>
/// </summary>
[Collection(LiveCollections.Phase4)]
[Trait("Category", "Live")]
public sealed class LivePaneTests
{
    private static readonly TimeSpan PaneTimeout = TimeSpan.FromSeconds(120);

    private readonly LivePhase4Fixture _fixture;
    private readonly ITestOutputHelper _output;

    public LivePaneTests(LivePhase4Fixture fixture, ITestOutputHelper output)
    {
        _fixture = fixture;
        _output = output;
    }

    private MailService Service => _fixture.Service;

    private string Hub => _fixture.Settings.TestHubStoreDisplayName;

    private string Marker => _fixture.RunMarker;

    [Fact]
    [Trait("Requires", "AddInRegistry")]
    [Trait("Requires", "MailAccount")]
    [Trait("Requires", "LocalSignatures")]
    public void ThePaneRewritesTheDraftInParagraphs_ThenItsAIChoosesTheSignature_WithThePictureEmbedded()
    {
        // A test machine's checkpoint holds the add-in of the day it was made; only a run that installed the tested
        // commit's (Invoke-LiveTierOnGuest.ps1 -WithAddIn) has the pane this test means to drive, and that run marks it.
        string mark = Path.Combine(RepoPaths.RepoRoot(), TestedAddInMark);
        IReadOnlyList<string> testedAddIn = LivePopulationCoverage.Require(
            _fixture.Settings,
            File.Exists(mark) ? new[] { File.ReadAllText(mark).Trim() } : Array.Empty<string>(),
            "the tested commit's add-in, installed on this machine for this run",
            "driving the task pane",
            "run Testbed/host/Invoke-LiveTierOnGuest.ps1 with -WithAddIn, which installs it before Outlook starts and marks it",
            _output.WriteLine);
        foreach (string commit in testedAddIn)
        {
            _output.WriteLine("the add-in under test: " + commit);
            DriveThePane();
        }
    }

    /// <summary>The file a run with -WithAddIn writes into the suite's source folder once the tested add-in is ready.</summary>
    private const string TestedAddInMark = ".tested-addin";

    private void DriveThePane()
    {
        using TestSignature agentsChoice = TestSignature.Create(Marker);
        using TestSignature aisChoice = TestSignature.Create(Marker, "B");
        string body = "Stand-in draft, first line " + Marker + "\nStand-in draft, second line.\n\nStand-in draft, a paragraph of its own.";
        using ClaudeCliStandIn standIn = ClaudeCliStandIn.Install(signatureAnswer: aisChoice.Name, bodyAnswer: body);

        string subject = _fixture.TaggedSubject("pane");
        DraftOutcome draft = Service.NewDraft(
            LiveStoreWriteGuard.Writable(Hub, StoreWriteKind.Draft, "new_draft"), Hub, cc: null, subject,
            "The agent's text " + Marker, display: true, signature: agentsChoice.Name);
        try
        {
            Assert.True(draft.SignatureApplied, $"the agent's signature must apply (error: {draft.SignatureError ?? "-"})");

            // The compose window and its pane come up; the agent's signature is in it.
            OpenDraftReading opened = WaitFor(
                "the draft's compose window with the agent's signature",
                () => LiveOutlookTestMailer.ReadOpenDraft(Hub, draft.EntryId, Marker),
                r => r.Window != IntPtr.Zero && r.SignatureText.Contains(agentsChoice.BodyMarker, StringComparison.Ordinal));
            _output.WriteLine($"opened: window 0x{opened.Window.ToInt64():X}, draft {opened.DraftParagraphs} paragraph(s), "
                + $"signature pictures {opened.SignaturePictures} ({opened.SignaturePicturesLinked} linked)");

            IntPtr prompt = WaitForWindow("the pane's instruction box", () => PaneControls.Find(opened.Window, "WindowsForms10.EDIT", null));
            IntPtr edit = WaitForWindow("the pane's Edit current draft button", () => PaneControls.Find(opened.Window, "WindowsForms10.BUTTON", "Edit current draft"));
            _output.WriteLine($"pane: instruction box {PaneControls.Describe(prompt)}, button {PaneControls.Describe(edit)}, "
                + $"status \"{PaneControls.ReadStatus(opened.Window)}\"");
            PaneControls.SetText(prompt, "Rewrite this mail. " + Marker);
            PaneControls.Click(edit);

            // What the pane says along the way is what tells a click that never landed from a request that failed.
            List<string> statuses = new List<string>();
            OpenDraftReading after;
            try
            {
                after = WaitFor(
                    "the stand-in's body and the AI's signature in the open window",
                    () =>
                    {
                        string now = PaneControls.ReadStatus(opened.Window);
                        if (statuses.Count == 0 || statuses[^1] != now)
                        {
                            statuses.Add(now);
                        }

                        return LiveOutlookTestMailer.ReadOpenDraft(Hub, draft.EntryId, Marker);
                    },
                    r => r.DraftText.Contains("Stand-in draft, first line", StringComparison.Ordinal)
                         && r.SignatureText.Contains(aisChoice.BodyMarker, StringComparison.Ordinal));
            }
            catch (TimeoutException ex)
            {
                throw new TimeoutException(
                    ex.Message + "\npane status, in order: " + string.Join(" -> ", statuses.Select(st => "\"" + st + "\""))
                    + "\ninstruction box now: " + PaneControls.Describe(prompt) + "; button now: " + PaneControls.Describe(edit)
                    + "\nprompts the stand-in received: " + standIn.Prompts().Count
                    + "; the stand-in is still at " + ClaudeCliStandIn.CliPath + ": " + ClaudeCliStandIn.IsInPlace, ex);
            }
            IReadOnlyList<string> prompts = standIn.Prompts();
            string status = PaneControls.ReadStatus(opened.Window);
            _output.WriteLine($"after: draft {after.DraftParagraphs} paragraph(s), signature pictures {after.SignaturePictures} "
                + $"({after.SignaturePicturesLinked} linked), {prompts.Count} prompt(s) to the stand-in, pane status \"{status}\"");
            _output.WriteLine("draft text: " + after.DraftText.Replace("\r", "\\r", StringComparison.Ordinal).Replace("\v", "\\v", StringComparison.Ordinal));

            // The pane's layout, kept through Core's writer: every line its own paragraph, no soft line breaks.
            Assert.DoesNotContain("\v", after.DraftText, StringComparison.Ordinal);
            Assert.True(after.DraftParagraphs >= 4, $"one paragraph a line: 4 lines, {after.DraftParagraphs} paragraph(s)");

            // Q162 (b): the AI chose after it wrote the body, and replaced the signature the agent put in.
            Assert.DoesNotContain(agentsChoice.BodyMarker, after.SignatureText, StringComparison.Ordinal);
            Assert.Contains(prompts, p => p.Contains("Rewrite this mail. " + Marker, StringComparison.Ordinal));
            Assert.Contains(prompts, p => p.Contains("## Available Signatures", StringComparison.Ordinal)
                                          && p.Contains(agentsChoice.Name, StringComparison.Ordinal)
                                          && p.Contains(aisChoice.Name, StringComparison.Ordinal));

            // D47: the chosen signature's picture is embedded, not a link to the signature's folder.
            Assert.True(after.SignaturePictures >= 1, "the chosen signature's picture must be in the draft");
            Assert.Equal(0, after.SignaturePicturesLinked);
        }
        finally
        {
            LiveOutlookTestMailer.CloseOpenDraftDiscarding(Hub, draft.EntryId, Marker);
            TryDeleteArtifact(draft.EntryId);
            LiveOutlookTestMailer.DeleteTaggedArtifactsUntilStableZero(Hub, Marker);
        }

        Assert.Equal(0, LiveOutlookTestMailer.CountTaggedArtifacts(Hub, Marker));
    }

    private static IntPtr WaitForWindow(string what, Func<IntPtr> find)
    {
        IntPtr? found = WaitFor<IntPtr?>(what, () => find() is IntPtr h && h != IntPtr.Zero ? h : null, h => h.HasValue);
        return found!.Value;
    }

    private static T WaitFor<T>(string what, Func<T?> read, Func<T, bool> done)
    {
        LiveWaitBudget wait = LiveWaitBudget.Of(PaneTimeout);
        T? last = default;
        while (true)
        {
            last = read();
            if (last != null && done(last))
            {
                return last;
            }

            if (!wait.HasTimeLeft)
            {
                throw new TimeoutException($"waited {PaneTimeout.TotalSeconds:F0} s for {what}; last seen: {last?.ToString() ?? "nothing"}");
            }

            Thread.Sleep(1000);
        }
    }

    private void TryDeleteArtifact(string entryId)
    {
        for (int attempt = 0; attempt < 3; attempt++)
        {
            try
            {
                LiveOutlookTestMailer.DeleteItemByEntryId(Hub, entryId, Marker);
                return;
            }
            catch (Exception) when (attempt < 2)
            {
                Thread.Sleep(1000);
            }
        }
    }

    /// <summary>
    /// The pane's controls, found under the compose window by their Win32 class and caption - WinForms
    /// controls are real windows - and worked with the messages a person's typing and clicking send.
    /// </summary>
    private static class PaneControls
    {
        private const uint WmSetText = 0x000C;
        private const uint BmClick = 0x00F5;

        /// <summary>The first child of <paramref name="window"/> whose class starts with <paramref name="classPrefix"/> and whose text is <paramref name="text"/> (any, when null).</summary>
        public static IntPtr Find(IntPtr window, string classPrefix, string? text)
        {
            IntPtr found = IntPtr.Zero;
            _ = EnumChildWindows(window, (child, _) =>
            {
                StringBuilder name = new StringBuilder(256);
                _ = GetClassNameW(child, name, name.Capacity);
                if (!name.ToString().StartsWith(classPrefix, StringComparison.Ordinal))
                {
                    return true;
                }

                if (text != null && !string.Equals(ReadText(child), text, StringComparison.Ordinal))
                {
                    return true;
                }

                found = child;
                return false;
            }, IntPtr.Zero);
            return found;
        }

        /// <summary>The pane's status line, or empty: the WinForms label that is neither the title nor the version.</summary>
        public static string ReadStatus(IntPtr window)
        {
            List<string> labels = new List<string>();
            _ = EnumChildWindows(window, (child, _) =>
            {
                StringBuilder name = new StringBuilder(256);
                _ = GetClassNameW(child, name, name.Capacity);
                if (name.ToString().StartsWith("WindowsForms10.STATIC", StringComparison.Ordinal))
                {
                    labels.Add(ReadText(child));
                }

                return true;
            }, IntPtr.Zero);
            return string.Join(" | ", labels.Where(l => l.Length > 0));
        }

        public static void SetText(IntPtr control, string text) => _ = SendMessageW(control, WmSetText, IntPtr.Zero, text);

        /// <summary>A control's class, text, and whether it is enabled and visible: for a test's output.</summary>
        public static string Describe(IntPtr control)
        {
            StringBuilder name = new StringBuilder(256);
            _ = GetClassNameW(control, name, name.Capacity);
            return $"0x{control.ToInt64():X} {name} \"{ReadText(control)}\" enabled={IsWindowEnabled(control)} visible={IsWindowVisible(control)}";
        }

        public static void Click(IntPtr button) => _ = PostMessageW(button, BmClick, IntPtr.Zero, IntPtr.Zero);

        private static string ReadText(IntPtr window)
        {
            StringBuilder text = new StringBuilder(512);
            _ = GetWindowTextW(window, text, text.Capacity);
            return text.ToString();
        }

        private delegate bool EnumChildProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool EnumChildWindows(IntPtr parent, EnumChildProc callback, IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetClassNameW(IntPtr hWnd, StringBuilder className, int maxCount);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetWindowTextW(IntPtr hWnd, StringBuilder text, int maxCount);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern IntPtr SendMessageW(IntPtr hWnd, uint msg, IntPtr wParam, string lParam);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool PostMessageW(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool IsWindowEnabled(IntPtr hWnd);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool IsWindowVisible(IntPtr hWnd);
    }
}
