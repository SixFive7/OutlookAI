using System.Globalization;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// The step of the typed-text tests a hunting run leaves out (the coordinator's direction of 2026-11-18, after the
/// variants): every crash the hunt caught struck in those tests, and none in the loops, which do less.
/// </summary>
public enum HuntAblation
{
    /// <summary>Nothing left out: the tests as they are.</summary>
    None,

    /// <summary>No new mail: only the reply and the forward are typed and made, never the new mail beside them.</summary>
    NewMail,

    /// <summary>No layout read: Word's paragraphs are never walked, in the typed window or in the product's drafts re-opened after.</summary>
    Layout,

    /// <summary>No signature: the product's second draft of each kind is made without the test signature, which is never made.</summary>
    Signature,

    /// <summary>One source mail: the tests of a leg share the first one's, where each sends itself its own.</summary>
    Source,
}

/// <summary>What makes a hunting variant's derived drafts (runbook section 4.15).</summary>
public enum HuntMaker
{
    /// <summary>The product's reply_draft and forward_draft, display off: its direct route in a windowed Outlook.</summary>
    Product,

    /// <summary>
    /// Reply() or Forward() by the object model, saved, re-opened by its EntryID, its editor read and edited and the
    /// inspector closed with olSave - the product's direct route without its "Message" form page.
    /// </summary>
    NoFormPage,

    /// <summary>Reply() or Forward() on the source, and what it returned released unsaved: Outlook's creator alone.</summary>
    Bare,

    /// <summary>No draft at all: the step only moves the main window (<see cref="HuntPlan.CalendarStep"/>).</summary>
    None,
}

/// <summary>
/// One variant of the hunt for the crash of OUTLOOK.EXE in mso40uiwin32client.dll (runbook section 4.15; the
/// coordinator's direction of 2026-11-18, option (b)): what makes the derived drafts, whether Outlook has a window,
/// and which folder its main window shows. Read from the two variables the hunt's suite script sets
/// (<c>Invoke-LiveTierOnGuest.ps1 -HuntFirstChance -HuntVariant</c>); pure, so T1 pins every variant.
/// </summary>
public sealed class HuntPlan
{
    /// <summary>The variable the hunt's suite script names the variant in. Unset in every other run.</summary>
    public const string VariantVariable = "OUTLOOKAI_HUNT_VARIANT";

    /// <summary>The variable it gives the minutes of one leg - one dotnet test - in.</summary>
    public const string LegMinutesVariable = "OUTLOOKAI_HUNT_LEG_MINUTES";

    /// <summary>
    /// The variable that, set to 1, has every draft come after a compose window shown and closed: a reply or forward the
    /// object model made, opened in an inspector the product's own way (parked, activated, hidden), closed discarding.
    /// Set to "typed", the window is displayed on screen and typed into at Outlook's insertion point, as the typed-text
    /// tests do (<c>Display</c>, Word's <c>Selection.TypeText</c>), then closed discarding.
    /// </summary>
    public const string ShowEachVariable = "OUTLOOKAI_HUNT_SHOW_EACH";

    /// <summary>The variable the hunt's suite script names the typed-text tests' ablation in (<see cref="HuntAblation"/>). Unset in every other run.</summary>
    public const string AblateVariable = "OUTLOOKAI_HUNT_ABLATE";

    /// <summary>PersonTyping: how many of the product's drafts are made while one compose window stays open and typed into.</summary>
    public const int DraftsPerOpenWindow = 5;

    /// <summary>The longest leg: two hours, beyond any run's own time limit.</summary>
    public const int MostLegMinutes = 120;

    /// <summary>
    /// The step the runner's LegsOtherFolder puts before the typed-text tests in every round: the main window moved to
    /// the Calendar, and no draft - so the tests the crashes struck in run with no mail of the Inbox on screen.
    /// </summary>
    public const string CalendarStep = "CalendarOnly";

    /// <summary>How many failed drafts in a row end a leg whose Outlook is still running.</summary>
    public const int MostFailuresInARow = 5;

    private HuntPlan(string variant, HuntMaker maker, bool windowless, bool showsCalendar, int legMinutes, bool showsEach, bool typesEach = false, bool keepsAWindowOpen = false)
    {
        TypesEach = typesEach;
        KeepsAWindowOpen = keepsAWindowOpen;
        Variant = variant;
        Maker = maker;
        Windowless = windowless;
        ShowsCalendar = showsCalendar;
        LegMinutes = legMinutes;
        ShowsEach = showsEach;
    }

    /// <summary>
    /// The variants this test runs, in the order the runner lists them; the runner's <c>-HuntVariant</c> takes these and
    /// its own default, Legs - the tests the crashes struck in, which are not this test's.
    /// </summary>
    public static IReadOnlyList<string> Variants { get; } = new[] { "ProductLoop", "OtherFolder", "NoWindow", "NoFormPage", "BareCreators", "PersonTyping" };

    public string Variant { get; }

    public HuntMaker Maker { get; }

    /// <summary>Whether the drafts are made in an Outlook that COM started, with no window: the windowed one is ended first.</summary>
    public bool Windowless { get; }

    /// <summary>Whether the main window is moved to the Calendar before the first draft, so no mail of the Inbox is on screen.</summary>
    public bool ShowsCalendar { get; }

    public int LegMinutes { get; }

    /// <summary>
    /// Whether every draft comes after a compose window shown and closed (<see cref="ShowEachVariable"/>): what the
    /// typed-text tests, where every hunted crash struck, do before each of the product's drafts - and what the loops
    /// without it, which never crashed, do not.
    /// </summary>
    public bool ShowsEach { get; }

    /// <summary>Whether the window shown before each draft is displayed on screen and typed into, rather than promoted.</summary>
    public bool TypesEach { get; }

    /// <summary>
    /// PersonTyping: a compose window stays open on screen - a new mail or a reply, by turns - and a line is typed into it
    /// before each of the product's drafts, <see cref="DraftsPerOpenWindow"/> of them, before it is closed discarding: a
    /// person writing in Outlook while an agent drafts replies and forwards.
    /// </summary>
    public bool KeepsAWindowOpen { get; }

    /// <summary>The cell of the variant's attempt lines - one a draft.</summary>
    public string Cell => "S5-hunt-" + Variant + (TypesEach ? "-typed" : ShowsEach ? "-shown" : string.Empty);

    /// <summary>
    /// The plan the two variables name; null with the reason when they name none - unset outside a hunting run, so
    /// the test writes a NOT RUN line and makes nothing.
    /// </summary>
    public static HuntPlan? Parse(string? variant, string? legMinutes, out string? why)
    {
        return Parse(variant, legMinutes, showEach: null, out why);
    }

    /// <summary><see cref="Parse(string?, string?, out string?)"/>, with <see cref="ShowEachVariable"/>: "1" shows a window before each draft, unset or "0" does not.</summary>
    public static HuntPlan? Parse(string? variant, string? legMinutes, string? showEach, out string? why)
    {
        why = null;
        if (string.IsNullOrWhiteSpace(variant))
        {
            why = "the hunt's variants run only in a hunting run that names one - " + VariantVariable + " is not set";
            return null;
        }

        if (!int.TryParse(legMinutes, NumberStyles.None, CultureInfo.InvariantCulture, out int minutes) || minutes < 1 || minutes > MostLegMinutes)
        {
            why = LegMinutesVariable + " must be a whole number of minutes from 1 to " + MostLegMinutes.ToString(CultureInfo.InvariantCulture) + ", not '" + legMinutes + "'";
            return null;
        }

        bool shows;
        bool types = false;
        switch (showEach)
        {
            case null:
            case "":
            case "0":
                shows = false;
                break;
            case "1":
                shows = true;
                break;
            case "typed":
                shows = true;
                types = true;
                break;
            default:
                why = ShowEachVariable + " must be 1, typed or 0, not '" + showEach + "'";
                return null;
        }

        switch (variant)
        {
            case "ProductLoop":
                return new HuntPlan(variant, HuntMaker.Product, windowless: false, showsCalendar: false, minutes, shows, types);
            case "OtherFolder":
                return new HuntPlan(variant, HuntMaker.Product, windowless: false, showsCalendar: true, minutes, shows, types);
            case "NoWindow":
                return new HuntPlan(variant, HuntMaker.Product, windowless: true, showsCalendar: false, minutes, shows, types);
            case "NoFormPage":
                return new HuntPlan(variant, HuntMaker.NoFormPage, windowless: false, showsCalendar: false, minutes, shows, types);
            case "BareCreators":
                return new HuntPlan(variant, HuntMaker.Bare, windowless: false, showsCalendar: false, minutes, shows, types);
            case "PersonTyping":
                return new HuntPlan(variant, HuntMaker.Product, windowless: false, showsCalendar: false, minutes, showsEach: false, typesEach: false, keepsAWindowOpen: true);
            case CalendarStep:
                return new HuntPlan(variant, HuntMaker.None, windowless: false, showsCalendar: true, minutes, showsEach: false);
            default:
                why = "'" + variant + "' is no hunting variant: " + string.Join(", ", Variants);
                return null;
        }
    }

    /// <summary>
    /// Whether a failure's own text says the Outlook it was made in is going: a COM call that failed in flight
    /// (RPC_S_CALL_FAILED), a server gone (RPC_S_SERVER_UNAVAILABLE) or a disconnect (RPC_E_DISCONNECTED) - the
    /// product's messages carry the HRESULT, the object model's exceptions their own. Such a failure waits for the
    /// process to end: a debugger writing its dump holds Outlook still for a while first.
    /// </summary>
    public static bool SaysOutlookIsGoing(string? failureText)
    {
        if (string.IsNullOrEmpty(failureText))
        {
            return false;
        }

        foreach (string hresult in new[] { "0x800706BE", "0x800706BA", "0x80010108" })
        {
            if (failureText!.Contains(hresult, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        return false;
    }

    /// <summary>The ablation <see cref="AblateVariable"/> names: unset or empty is none; null with the reason for anything unknown.</summary>
    public static HuntAblation? ParseAblation(string? value, out string? why)
    {
        why = null;
        switch (value)
        {
            case null:
            case "":
                return HuntAblation.None;
            case "new":
                return HuntAblation.NewMail;
            case "layout":
                return HuntAblation.Layout;
            case "signature":
                return HuntAblation.Signature;
            case "source":
                return HuntAblation.Source;
            default:
                why = AblateVariable + " must be new, layout, signature or source, not '" + value + "'";
                return null;
        }
    }

    /// <summary>Whether a leg stops after a draft: its Outlook ended, or <see cref="MostFailuresInARow"/> drafts in a row failed.</summary>
    public static bool StopsAfter(bool outlookAlive, int failuresInARow)
    {
        return !outlookAlive || failuresInARow >= MostFailuresInARow;
    }
}
