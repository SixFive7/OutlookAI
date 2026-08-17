using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using Microsoft.Win32;

namespace OutlookAI.Services
{
    /// <summary>
    /// The key/value operations <see cref="PromptStore"/> needs, and nothing else. Exists so the
    /// store's rules - override wins, absent means default, never persist a default, prune what
    /// no button points at - are pinned by the T1 suite against a fabricated backing store
    /// instead of against the developer's own HKCU. Same Func-injection idea as
    /// <c>HealthReporting.ReadTuningState</c>, widened to an interface because a store needs six
    /// operations rather than one read.
    ///
    /// <paramref name="subKey"/> is the empty string for the root Prompts key, or one of
    /// <see cref="PromptStore.ButtonPromptsSubKey"/> / <see cref="PromptStore.SectionsSubKey"/>.
    ///
    /// Contract, and both halves matter:
    ///  - READS must not throw. A key that is missing, unreadable or holding the wrong value
    ///    type is reported as false or as an empty list, because a settings read happens while
    ///    building the task pane and an exception there costs the user the pane, not a prompt.
    ///  - WRITES may throw. The store catches and reports the failure to the caller, so
    ///    "your prompts were saved" is never said about a write that did not happen.
    /// </summary>
    internal interface IPromptRegistry
    {
        /// <summary>
        /// Reads a string value. False when the value is absent or not a string, which is what
        /// makes "absent" distinguishable from an empty string - the difference between "use
        /// the six default buttons" and "the user deleted all of them".
        /// </summary>
        bool TryReadString(string subKey, string valueName, out string value);

        /// <summary>Reads a DWORD value. False when absent or not a DWORD.</summary>
        bool TryReadDword(string subKey, string valueName, out int value);

        /// <summary>Creates the key if needed and writes a string value.</summary>
        void WriteString(string subKey, string valueName, string value);

        /// <summary>Creates the key if needed and writes a DWORD value.</summary>
        void WriteDword(string subKey, string valueName, int value);

        /// <summary>Deletes a value. A value that is already absent is not an error.</summary>
        void DeleteValue(string subKey, string valueName);

        /// <summary>
        /// Value names present under <paramref name="subKey"/>, or an empty list when the key
        /// does not exist. A fresh list, so the caller may delete while iterating it.
        /// </summary>
        IList<string> ListValueNames(string subKey);
    }

    /// <summary>
    /// Outcome of a settings write: succeeded, or a list of messages written for the user.
    /// Returned rather than thrown - every caller is a click handler, and a rejected button
    /// name is an ordinary thing for a person to type, not an exceptional one. A registry write
    /// that fails arrives here too, so a caller that reports <see cref="Succeeded"/> honestly
    /// cannot claim a save that did not land.
    /// </summary>
    internal sealed class PromptValidationResult
    {
        private static readonly PromptValidationResult _ok =
            new PromptValidationResult(new List<string>());

        private readonly ReadOnlyCollection<string> _errors;

        private PromptValidationResult(IList<string> errors)
        {
            _errors = new ReadOnlyCollection<string>(new List<string>(errors));
        }

        internal static PromptValidationResult Ok()
        {
            return _ok;
        }

        /// <summary>
        /// A failure carrying at least one message. Called with nothing to say, it still fails
        /// (with a generic message) rather than quietly reporting success.
        /// </summary>
        internal static PromptValidationResult Failed(params string[] errors)
        {
            var list = new List<string>();
            if (errors != null)
            {
                foreach (string error in errors)
                {
                    if (!string.IsNullOrEmpty(error))
                        list.Add(error);
                }
            }
            if (list.Count == 0)
                list.Add("The prompt settings could not be saved.");
            return new PromptValidationResult(list);
        }

        /// <summary>True when nothing was rejected and everything asked for was written.</summary>
        internal bool Succeeded
        {
            get { return _errors.Count == 0; }
        }

        /// <summary>Messages fit to show a user. Empty when <see cref="Succeeded"/>.</summary>
        internal IList<string> Errors
        {
            get { return _errors; }
        }

        /// <summary>All messages on separate lines; empty string when there are none.</summary>
        internal string Message
        {
            get { return string.Join(Environment.NewLine, _errors); }
        }
    }

    /// <summary>
    /// USER-EDITED PROMPTS AND QUICK BUTTONS, STORED AS OVERRIDES ONLY.
    ///
    /// All the rules, with the backing store injected, so the T1 suite can pin them without a
    /// registry. <see cref="PromptStore"/> is the static facade over an instance of this bound
    /// to HKCU; everything below is the actual behaviour.
    ///
    /// The one idea the whole class is built around: <b>only differences are stored</b>. Text
    /// that matches what OutlookAI ships is not written, and is deleted if it is already there.
    /// So an absent value always means "use the text in PromptDefaults", and improving a
    /// default still reaches every user who never edited that block. A store that wrote its
    /// defaults out on first run would silently pin today's wording on every machine forever.
    ///
    /// Buttons follow from that, plus one deliberate simplification: a button IS its name.
    ///  - The ordered name list is authoritative WHEN PRESENT. Absent means the six shipped
    ///    buttons; present means exactly what it lists, so a default the user deleted stays
    ///    deleted with no tombstone to maintain, and an empty list means "no buttons".
    ///  - A prompt override is a value named after its button. Renaming a button therefore
    ///    creates a new one and orphans the old override, which is why every save prunes
    ///    overrides no listed button points at.
    ///  - Registry value names are case-insensitive, so button names are compared that way and
    ///    a save that would collide is rejected instead of silently merging two buttons.
    /// </summary>
    internal sealed class PromptStoreCore
    {
        private readonly IPromptRegistry _registry;
        private readonly object _gate = new object();

        /// <summary>
        /// Raised after a write that landed. Initialised to a no-op so raising it never needs a
        /// null check, and so a subscriber that unhooks itself cannot leave it null.
        /// </summary>
        internal event EventHandler Changed = delegate { };

        internal PromptStoreCore(IPromptRegistry registry)
        {
            if (registry == null)
                throw new ArgumentNullException(nameof(registry));
            _registry = registry;
        }

        // ===== Buttons =====

        /// <summary>
        /// The buttons to show, in order. The stored order decides WHICH buttons exist - the
        /// shipped six when there is no order value - and each prompt is then resolved the same
        /// way whether or not that value was there: the override if one is stored, otherwise the
        /// shipped text. Resolving prompts on one path is deliberate; an override that only
        /// applied once an order had also been written would be a trap.
        ///
        /// Never throws, and never returns a broken button: a listed name with neither an
        /// override nor a shipped prompt is a stale order entry, not a button, and is skipped -
        /// showing a button whose instruction is empty would send the model a blank action.
        /// </summary>
        internal IList<PromptButton> GetButtons()
        {
            try
            {
                string order;
                IList<string> names = _registry.TryReadString(string.Empty, PromptStore.ButtonsValueName, out order)
                    ? SplitNames(order)
                    : PromptDefaults.ButtonNames;

                var buttons = new List<PromptButton>(names.Count);
                foreach (string name in names)
                {
                    string prompt;
                    if (!TryReadButtonOverride(name, out prompt) &&
                        !PromptDefaults.TryGetButtonPrompt(name, out prompt))
                        continue;

                    buttons.Add(new PromptButton(name, prompt));
                }
                return buttons;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("PromptStore.GetButtons: " + ex.Message);
                return PromptDefaults.CreateButtons();
            }
        }

        /// <summary>
        /// Ordered names from the stored value: one per line, blanks dropped, and duplicates
        /// dropped case-insensitively because two registry values cannot differ by case alone.
        /// </summary>
        private static IList<string> SplitNames(string order)
        {
            var names = new List<string>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string rawName in order.Split('\n'))
            {
                string name = rawName.Trim();
                if (name.Length == 0)
                    continue;
                if (!seen.Add(name))
                    continue;
                names.Add(name);
            }
            return names;
        }

        /// <summary>
        /// Checks a button set without writing anything, so the settings dialog can report a
        /// bad name while the user is still typing it.
        /// </summary>
        internal PromptValidationResult ValidateButtons(IEnumerable<PromptButton> buttons)
        {
            List<PromptButton> list;
            return Validate(buttons, out list);
        }

        /// <summary>
        /// Writes the button set: the order, then the prompts that actually differ from the
        /// shipped text, then a prune of every override no listed button points at.
        ///
        /// All or nothing on validation - a rejected set writes nothing at all, so a typo in
        /// one name cannot half-apply the dialog. An empty set is legal and means the user
        /// removed every button.
        /// </summary>
        internal PromptValidationResult SaveButtons(IEnumerable<PromptButton> buttons)
        {
            List<PromptButton> list;
            PromptValidationResult validation = Validate(buttons, out list);
            if (!validation.Succeeded)
                return validation;

            lock (_gate)
            {
                try
                {
                    var names = new List<string>(list.Count);
                    foreach (PromptButton button in list)
                        names.Add(button.Name);

                    // Order of operations, and it matters if a write fails half way through:
                    // prompts first, the order next, the prune last. Every intermediate state is
                    // then still coherent - the buttons in force are either the old set or the
                    // new one, and none of them is left pointing at a prompt that is not there.
                    foreach (PromptButton button in list)
                    {
                        string shipped;
                        if (PromptDefaults.TryGetButtonPrompt(button.Name, out shipped) &&
                            PromptDefaults.SameText(button.Prompt, shipped))
                        {
                            // Identical to what we ship: drop the override so a future
                            // improvement to the default still reaches this user.
                            _registry.DeleteValue(PromptStore.ButtonPromptsSubKey, button.Name);
                        }
                        else
                        {
                            _registry.WriteString(PromptStore.ButtonPromptsSubKey, button.Name, button.Prompt);
                        }
                    }

                    _registry.WriteString(string.Empty, PromptStore.ButtonsValueName,
                        string.Join("\n", names));

                    var keep = new HashSet<string>(names, StringComparer.OrdinalIgnoreCase);
                    foreach (string existing in _registry.ListValueNames(PromptStore.ButtonPromptsSubKey))
                    {
                        if (!keep.Contains(existing))
                            _registry.DeleteValue(PromptStore.ButtonPromptsSubKey, existing);
                    }

                    EnsureSchemaVersion();
                }
                catch (Exception ex)
                {
                    return WriteFailure("buttons", ex);
                }
            }

            RaiseChanged();
            return PromptValidationResult.Ok();
        }

        /// <summary>
        /// Drops one button's prompt override, so it falls back to the shipped text. Leaves the
        /// button order alone; a custom button has no shipped text to fall back to, so its
        /// prompt is deleted and the name becomes a stale order entry until the user saves again.
        /// </summary>
        internal bool ResetButtonPrompt(string name)
        {
            if (string.IsNullOrEmpty(name))
                return false;
            return Write(delegate
            {
                _registry.DeleteValue(PromptStore.ButtonPromptsSubKey, name);
            }, "button prompt");
        }

        /// <summary>
        /// Back to the six shipped buttons: the order and every prompt override go, which also
        /// removes any custom button. Sections are untouched.
        /// </summary>
        internal bool RestoreDefaultButtons()
        {
            return Write(delegate
            {
                // Overrides first, order last: a failure part way through then leaves the
                // user's own button set with shipped prompts, never the reverse - which would
                // be the shipped button set still answering with edited prompts.
                foreach (string existing in _registry.ListValueNames(PromptStore.ButtonPromptsSubKey))
                    _registry.DeleteValue(PromptStore.ButtonPromptsSubKey, existing);
                _registry.DeleteValue(string.Empty, PromptStore.ButtonsValueName);
            }, "buttons");
        }

        /// <summary>
        /// Whether this button is not what OutlookAI ships: any name that is not built in, or a
        /// built-in name with a stored prompt that differs. Reads storage, so it answers for the
        /// saved state rather than for whatever the dialog currently shows.
        /// </summary>
        internal bool IsButtonCustomized(string name)
        {
            if (string.IsNullOrEmpty(name) || name.Trim().Length == 0)
                return false;
            if (!PromptDefaults.IsDefaultButtonName(name))
                return true;

            try
            {
                string stored;
                if (!TryReadButtonOverride(name, out stored))
                    return false;

                string shipped;
                PromptDefaults.TryGetButtonPrompt(name, out shipped);
                return !PromptDefaults.SameText(stored, shipped);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("PromptStore.IsButtonCustomized: " + ex.Message);
                return false;
            }
        }

        // ===== Sections =====

        /// <summary>
        /// The text to use for a section: the user's override if one is stored, otherwise the
        /// shipped text. An override that is present but empty is honoured as empty - a user
        /// who clears a section means to remove it, and only a genuinely absent value falls
        /// back to the default. Never throws.
        /// </summary>
        internal string GetSection(PromptSection section)
        {
            try
            {
                string stored;
                if (_registry.TryReadString(PromptStore.SectionsSubKey, SectionValueName(section), out stored))
                    return stored;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("PromptStore.GetSection: " + ex.Message);
            }
            return PromptDefaults.GetSection(section);
        }

        /// <summary>
        /// Stores an edited section. Text equal to the shipped default deletes the override
        /// instead of writing a copy of it, so returning a section to its original wording
        /// really does hand it back to the code default.
        /// </summary>
        internal bool SetSection(PromptSection section, string text)
        {
            string value = text == null ? string.Empty : text;
            if (PromptDefaults.SameText(value, PromptDefaults.GetSection(section)))
                return ResetSection(section);

            return Write(delegate
            {
                _registry.WriteString(PromptStore.SectionsSubKey, SectionValueName(section), value);
            }, "prompt section");
        }

        /// <summary>Drops a section override, restoring the shipped text.</summary>
        internal bool ResetSection(PromptSection section)
        {
            return Write(delegate
            {
                _registry.DeleteValue(PromptStore.SectionsSubKey, SectionValueName(section));
            }, "prompt section");
        }

        /// <summary>Whether a section has a stored override whose text differs from the shipped one.</summary>
        internal bool IsSectionCustomized(PromptSection section)
        {
            try
            {
                string stored;
                if (!_registry.TryReadString(PromptStore.SectionsSubKey, SectionValueName(section), out stored))
                    return false;
                return !PromptDefaults.SameText(stored, PromptDefaults.GetSection(section));
            }
            catch (Exception ex)
            {
                Debug.WriteLine("PromptStore.IsSectionCustomized: " + ex.Message);
                return false;
            }
        }

        // ===== Internals =====

        private bool TryReadButtonOverride(string name, out string prompt)
        {
            prompt = string.Empty;
            string stored;
            if (!_registry.TryReadString(PromptStore.ButtonPromptsSubKey, name, out stored))
                return false;
            // A blank override is not an instruction. Treat it as absent so a built-in button
            // keeps working instead of asking the model to perform an empty action.
            if (string.IsNullOrEmpty(stored) || stored.Trim().Length == 0)
                return false;
            prompt = stored;
            return true;
        }

        private PromptValidationResult Validate(IEnumerable<PromptButton> buttons, out List<PromptButton> accepted)
        {
            accepted = new List<PromptButton>();
            if (buttons == null)
                return PromptValidationResult.Failed("No buttons were supplied.");

            var errors = new List<string>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (PromptButton button in buttons)
            {
                if (button == null)
                {
                    errors.Add("A button was missing.");
                    continue;
                }

                string name = button.Name;
                if (string.IsNullOrEmpty(name) || name.Trim().Length == 0)
                {
                    errors.Add("A button name cannot be empty.");
                    continue;
                }
                if (name.IndexOf('\n') >= 0 || name.IndexOf('\r') >= 0 || name.IndexOf('\t') >= 0)
                {
                    // The order is one value with one name per line, so a line break in a name
                    // would silently split it into two buttons.
                    errors.Add("Button name \"" + Describe(name) + "\" cannot contain line breaks or tabs.");
                    continue;
                }
                if (name != name.Trim())
                {
                    errors.Add("Button name \"" + name + "\" cannot start or end with a space.");
                    continue;
                }
                if (name.Length > PromptDefaults.MaxButtonNameLength)
                {
                    errors.Add("Button name \"" + name + "\" is longer than " +
                        PromptDefaults.MaxButtonNameLength + " characters.");
                    continue;
                }
                if (!seen.Add(name))
                {
                    errors.Add("There is more than one button named \"" + name +
                        "\". Button names must be unique, and they are not case sensitive.");
                    continue;
                }
                if (string.IsNullOrEmpty(button.Prompt) || button.Prompt.Trim().Length == 0)
                {
                    errors.Add("The prompt for \"" + name + "\" cannot be empty.");
                    continue;
                }

                accepted.Add(button);
            }

            if (errors.Count > 0)
            {
                accepted.Clear();
                return PromptValidationResult.Failed(errors.ToArray());
            }
            return PromptValidationResult.Ok();
        }

        /// <summary>Value name for a section: the enum name, which is the registry layout.</summary>
        private static string SectionValueName(PromptSection section)
        {
            return section.ToString();
        }

        private static string Describe(string name)
        {
            return name.Replace("\r", " ").Replace("\n", " ").Replace("\t", " ").Trim();
        }

        /// <summary>
        /// Runs a write under the lock, stamps the schema version, reports failure instead of
        /// throwing into a click handler, and raises <see cref="Changed"/> only when it landed.
        /// </summary>
        private bool Write(Action write, string what)
        {
            lock (_gate)
            {
                try
                {
                    write();
                    EnsureSchemaVersion();
                }
                catch (Exception ex)
                {
                    Debug.WriteLine("PromptStore write (" + what + "): " + ex.Message);
                    return false;
                }
            }
            RaiseChanged();
            return true;
        }

        private PromptValidationResult WriteFailure(string what, Exception ex)
        {
            Debug.WriteLine("PromptStore write (" + what + "): " + ex.Message);
            return PromptValidationResult.Failed(
                "The prompt settings could not be saved: " + ex.Message);
        }

        /// <summary>
        /// Stamps the layout version, on the first write and after that only if something has
        /// removed or changed it. Nothing reads it yet; it exists so a future layout change can
        /// tell an old store from a new one without guessing.
        /// </summary>
        private void EnsureSchemaVersion()
        {
            int stored;
            if (_registry.TryReadDword(string.Empty, PromptStore.SchemaVersionValueName, out stored) &&
                stored == PromptStore.SchemaVersion)
                return;
            _registry.WriteDword(string.Empty, PromptStore.SchemaVersionValueName, PromptStore.SchemaVersion);
        }

        private void RaiseChanged()
        {
            try
            {
                Changed(this, EventArgs.Empty);
            }
            catch (Exception ex)
            {
                // A pane that failed to rebuild is not a reason to report a failed save.
                Debug.WriteLine("PromptStore.Changed handler: " + ex.Message);
            }
        }
    }

    /// <summary>
    /// <see cref="IPromptRegistry"/> over HKCU. Same discipline as
    /// <c>OutlookTuningService</c>'s registry helpers: current user only, no elevation, and a
    /// read that cannot be satisfied returns false rather than throwing.
    /// </summary>
    internal sealed class HkcuPromptRegistry : IPromptRegistry
    {
        public bool TryReadString(string subKey, string valueName, out string value)
        {
            value = string.Empty;
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(KeyPathFor(subKey)))
                {
                    if (key == null)
                        return false;
                    var stored = key.GetValue(valueName);
                    if (stored is string text)
                    {
                        value = text;
                        return true;
                    }
                    return false;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("HkcuPromptRegistry.TryReadString: " + ex.Message);
                return false;
            }
        }

        public bool TryReadDword(string subKey, string valueName, out int value)
        {
            value = 0;
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(KeyPathFor(subKey)))
                {
                    if (key == null)
                        return false;
                    var stored = key.GetValue(valueName);
                    if (stored is int number)
                    {
                        value = number;
                        return true;
                    }
                    return false;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("HkcuPromptRegistry.TryReadDword: " + ex.Message);
                return false;
            }
        }

        public void WriteString(string subKey, string valueName, string value)
        {
            string path = KeyPathFor(subKey);
            using (var key = Registry.CurrentUser.CreateSubKey(path))
            {
                if (key == null)
                    throw new InvalidOperationException("Could not open HKCU\\" + path + " for writing.");
                key.SetValue(valueName, value == null ? string.Empty : value, RegistryValueKind.String);
            }
        }

        public void WriteDword(string subKey, string valueName, int value)
        {
            string path = KeyPathFor(subKey);
            using (var key = Registry.CurrentUser.CreateSubKey(path))
            {
                if (key == null)
                    throw new InvalidOperationException("Could not open HKCU\\" + path + " for writing.");
                key.SetValue(valueName, value, RegistryValueKind.DWord);
            }
        }

        public void DeleteValue(string subKey, string valueName)
        {
            using (var key = Registry.CurrentUser.OpenSubKey(KeyPathFor(subKey), writable: true))
            {
                if (key == null)
                    return;
                key.DeleteValue(valueName, throwOnMissingValue: false);
            }
        }

        public IList<string> ListValueNames(string subKey)
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(KeyPathFor(subKey)))
                {
                    if (key == null)
                        return new List<string>();
                    return new List<string>(key.GetValueNames());
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("HkcuPromptRegistry.ListValueNames: " + ex.Message);
                return new List<string>();
            }
        }

        private static string KeyPathFor(string subKey)
        {
            return string.IsNullOrEmpty(subKey)
                ? PromptStore.KeyPath
                : PromptStore.KeyPath + "\\" + subKey;
        }
    }

    /// <summary>
    /// The add-in's prompt settings, bound to HKCU. Static because the task pane, the ribbon
    /// and the settings dialog all need the same answer and none of them owns the others.
    ///
    /// Registry layout, all under <c>HKCU\Software\OutlookAI\Prompts</c>, and all of it optional:
    /// <code>
    /// Prompts
    ///   Buttons        REG_SZ    ordered button names, one per line. Authoritative WHEN
    ///                            PRESENT; absent means the six shipped buttons. A shipped
    ///                            button the user deleted is simply not in the list.
    ///   SchemaVersion  REG_DWORD 1. Stamped on the first write.
    ///   ButtonPrompts\ <value name = button name>  REG_SZ  prompt override for that button
    ///   Sections\      <value name = PromptSection name> REG_SZ  override for that section
    /// </code>
    /// Nothing is created until the user changes something, and nothing that matches the
    /// shipped text is ever written - see <see cref="PromptStoreCore"/> for why that is the
    /// whole design rather than an optimisation.
    ///
    /// A section value under a name no <see cref="PromptSection"/> carries any more - a machine
    /// that saved the old <c>HumanVoice</c> block before it was folded into the preamble - is
    /// inert rather than a problem: every read asks for one named value, and the only key that
    /// is ever enumerated (and pruned) is <c>ButtonPrompts</c>. Nothing looks for it, so nothing
    /// trips over it, and it is left where it is rather than deleted behind the user's back.
    ///
    /// Nothing is cached: a read goes to the registry every time. The values are tiny, they
    /// change only when a human clicks Save, and a stale prompt would be far more confusing
    /// than the microseconds are worth.
    /// </summary>
    internal static class PromptStore
    {
        /// <summary>Root key, under HKCU.</summary>
        internal const string KeyPath = @"Software\OutlookAI\Prompts";

        /// <summary>Ordered button names, one per line. Absent means the shipped six.</summary>
        internal const string ButtonsValueName = "Buttons";

        /// <summary>Subkey of prompt overrides, one value per button name.</summary>
        internal const string ButtonPromptsSubKey = "ButtonPrompts";

        /// <summary>Subkey of section overrides, one value per <see cref="PromptSection"/>.</summary>
        internal const string SectionsSubKey = "Sections";

        /// <summary>Layout version value name.</summary>
        internal const string SchemaVersionValueName = "SchemaVersion";

        /// <summary>Current layout version.</summary>
        internal const int SchemaVersion = 1;

        /// <summary>
        /// Raised after any write that landed, so a pane showing buttons can rebuild them.
        /// Static and process-wide: several compose windows can be open at once and all of them
        /// are stale the moment the settings dialog saves. Initialised to a no-op, so raising it
        /// needs no null check.
        /// </summary>
        internal static event EventHandler Changed = delegate { };

        private static readonly PromptStoreCore _core = CreateCore();

        /// <summary>The buttons to show, in order. Never throws; never empty unless the user emptied it.</summary>
        internal static IList<PromptButton> GetButtons()
        {
            return _core.GetButtons();
        }

        /// <summary>Checks a button set without writing anything.</summary>
        internal static PromptValidationResult ValidateButtons(IEnumerable<PromptButton> buttons)
        {
            return _core.ValidateButtons(buttons);
        }

        /// <summary>Stores the button set: order, the prompts that differ, and a prune of the rest.</summary>
        internal static PromptValidationResult SaveButtons(IEnumerable<PromptButton> buttons)
        {
            return _core.SaveButtons(buttons);
        }

        /// <summary>Drops one button's prompt override.</summary>
        internal static bool ResetButtonPrompt(string name)
        {
            return _core.ResetButtonPrompt(name);
        }

        /// <summary>Back to the six shipped buttons, dropping every override and custom button.</summary>
        internal static bool RestoreDefaultButtons()
        {
            return _core.RestoreDefaultButtons();
        }

        /// <summary>Whether the saved state of this button differs from what OutlookAI ships.</summary>
        internal static bool IsButtonCustomized(string name)
        {
            return _core.IsButtonCustomized(name);
        }

        /// <summary>The text to use for a section: the override if stored, otherwise the shipped text.</summary>
        internal static string GetSection(PromptSection section)
        {
            return _core.GetSection(section);
        }

        /// <summary>Stores an edited section; text equal to the shipped default deletes the override.</summary>
        internal static bool SetSection(PromptSection section, string text)
        {
            return _core.SetSection(section, text);
        }

        /// <summary>Drops a section override, restoring the shipped text.</summary>
        internal static bool ResetSection(PromptSection section)
        {
            return _core.ResetSection(section);
        }

        /// <summary>Whether a section has a stored override that differs from the shipped text.</summary>
        internal static bool IsSectionCustomized(PromptSection section)
        {
            return _core.IsSectionCustomized(section);
        }

        private static PromptStoreCore CreateCore()
        {
            var core = new PromptStoreCore(new HkcuPromptRegistry());
            // Lambda rather than a method group: the parameters then take their nullability from
            // EventHandler itself, which keeps this compiling both in the add-in (no nullable
            // context) and in the nullable-enabled test project that links this file.
            core.Changed += (sender, e) => Changed(null, EventArgs.Empty);
            return core;
        }
    }
}
