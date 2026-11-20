using System.Text.RegularExpressions;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// The live tests read a draft's layout DIRECTLY, its inspector never activated, and promote one only where they say
/// why (the coordinator's direction of 2026-11-20, runbook section 4.15): re-opening the product's drafts and promoting
/// their inspectors is the step the mso40uiwin32client.dll crash needs - a loop doing that and nothing else crashed
/// Outlook on its 121st draft. These read the sources, so a promotion cannot creep back into the readers, nor appear
/// anywhere else in the live tests without a reason beside it.
/// </summary>
public sealed class DraftReadPromotionTests
{
    /// <summary>The marker every kept promotion carries, in its member's documentation or in a comment before the call.</summary>
    private const string KeptMarker = "PROMOTED ON PURPOSE";

    /// <summary>
    /// Every member of the live tests that may promote an inspector, by file: the readers of Outlook's own signature
    /// engine, which were measured through promoted windows; the one test of the promoted read and the hunt's
    /// reproduction of the crash; the hunt's shown windows; and the windowless cells that measure the promotion itself.
    /// A new one needs its reason beside it, under the marker, and its place here.
    /// </summary>
    private static readonly (string File, string Member)[] KeptPromotions =
    {
        ("LiveOutlookTestMailer.cs", "ReadSignatureEngine"),
        ("LiveOutlookTestMailer.cs", "ReadOutlookOwnSignature"),
        ("LiveOutlookTestMailer.cs", "ReadDraftLookPromoted"),
        ("LiveComposeDrafts.cs", "ReadLookOfPromoted"),
        ("LiveHuntDerivedDraftTests.cs", "ShowAndClose"),
        ("LiveQ149WindowlessComposeTests.cs", "PromoteOnce"),
    };

    [Theory]
    [InlineData("LiveComposeDrafts.cs", "internal static (List<WordParagraphLook>? Look, string? Failure) ReadLookOf(")]
    [InlineData("LiveComposeDrafts.cs", "private static (List<WordParagraphLook>? Look, string? Failure) ReadLookThrough(")]
    [InlineData("LiveOutlookTestMailer.cs", "public static IReadOnlyList<WordParagraphLook> ReadDraftLook(")]
    [InlineData("LiveOutlookTestMailer.cs", "private static IReadOnlyList<WordParagraphLook> ReadDraftLookThrough(")]
    public void TheLayoutReaders_NeverPromoteNorActivate(string file, string declaration)
    {
        string body = SourceText.CodeOnly(SourceText.MemberBody(SourceText.Lines(SourceText.LiveSource(file)), declaration, "    ", file));

        Assert.DoesNotContain("PromoteForWordEditor(", body, StringComparison.Ordinal);
        Assert.DoesNotContain(".Activate(", body, StringComparison.Ordinal);
        Assert.DoesNotContain("ObtainWordEditor(", body, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("LiveComposeDrafts.cs", "internal static (List<WordParagraphLook>? Look, string? Failure) ReadLookOf(")]
    [InlineData("LiveOutlookTestMailer.cs", "public static IReadOnlyList<WordParagraphLook> ReadDraftLook(")]
    public void TheLayoutReaders_ReadTheEditorDirectly_AsTheProductsDirectRouteDoes(string file, string declaration)
    {
        string body = SourceText.CodeOnly(SourceText.MemberBody(SourceText.Lines(SourceText.LiveSource(file)), declaration, "    ", file));

        Assert.Contains("ReadEditor(", body, StringComparison.Ordinal);
        Assert.Contains("ComposeSurface.DirectReadBudgetMilliseconds", body, StringComparison.Ordinal);
    }

    [Fact]
    public void EveryPromotionInTheLiveTests_IsAKeptOne_WithItsReasonBesideIt()
    {
        string dir = Path.Combine(RepoPaths.TestProjectDir(), "T2");
        List<string> found = new();
        foreach (string path in Directory.GetFiles(dir, "*.cs"))
        {
            string file = Path.GetFileName(path);
            string[] lines = SourceText.Lines(File.ReadAllText(path));
            for (int i = 0; i < lines.Length; i++)
            {
                string code = lines[i].TrimStart();
                if (code.StartsWith("//", StringComparison.Ordinal) || !code.Contains("PromoteForWordEditor(", StringComparison.Ordinal))
                {
                    continue;
                }

                int declaration = EnclosingMember(lines, i);
                Assert.True(declaration >= 0, file + " line " + (i + 1) + " promotes outside any member this test can name");
                string member = MemberName(lines[declaration]);
                found.Add(file + " " + member);
                Assert.True(
                    KeptPromotions.Contains((file, member)),
                    file + " line " + (i + 1) + ", in " + member + ", promotes an inspector: read the editor directly, or say why it must"
                        + " promote under '" + KeptMarker + "' and add the member to " + nameof(KeptPromotions) + " (runbook 4.15)");
                int docStart = declaration;
                while (docStart > 0 && lines[docStart - 1].TrimStart().StartsWith("///", StringComparison.Ordinal))
                {
                    docStart--;
                }

                string reasonText = string.Join("\n", lines[docStart..(i + 1)]);
                Assert.True(
                    reasonText.Contains(KeptMarker, StringComparison.Ordinal),
                    file + " line " + (i + 1) + ", in " + member + ": a kept promotion says why under '" + KeptMarker + "', in its member's documentation or before the call");
            }
        }

        // Every kept one is still there: a list that names a member which no longer promotes has stopped saying anything.
        foreach ((string file, string member) in KeptPromotions)
        {
            Assert.Contains(file + " " + member, found);
        }
    }

    /// <summary>The line of the declaration of the member line <paramref name="at"/> is in: the nearest member line above it at four spaces.</summary>
    private static int EnclosingMember(string[] lines, int at)
    {
        for (int i = at; i >= 0; i--)
        {
            string line = lines[i];
            if (line.StartsWith("    ", StringComparison.Ordinal) && !line.StartsWith("     ", StringComparison.Ordinal)
                && (line.StartsWith("    public ", StringComparison.Ordinal) || line.StartsWith("    internal ", StringComparison.Ordinal)
                    || line.StartsWith("    private ", StringComparison.Ordinal) || line.StartsWith("    protected ", StringComparison.Ordinal))
                && line.Contains('(', StringComparison.Ordinal))
            {
                return i;
            }
        }

        return -1;
    }

    /// <summary>
    /// The member's name: the first word its declaration line puts straight before a parenthesis - so not a tuple it
    /// returns, whose parenthesis follows a space (<c>internal static (List&lt;WordParagraphLook&gt;? Look, string? Failure) ReadLookOf(</c>).
    /// </summary>
    private static string MemberName(string declaration)
    {
        Match name = Regex.Match(declaration, @"\b([A-Za-z_]\w*)(?:<[^<>()]*>)?\(");
        Assert.True(name.Success, "no member name in: " + declaration.Trim());
        return name.Groups[1].Value;
    }
}
