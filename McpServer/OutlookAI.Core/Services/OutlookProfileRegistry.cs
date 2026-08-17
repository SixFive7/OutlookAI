namespace OutlookAI.Core.Services
{
    /// <summary>
    /// The Outlook profile registry addresses this product depends on, in one place.
    /// <para>
    /// The accounts GUID subkey in particular was written out three times - twice in the
    /// signature code and once in a live test fixture - as a bare 32-character literal with
    /// no name attached. A typo in one copy is invisible on inspection and produces "no
    /// accounts found" rather than an error, so it exists here instead, named.
    /// </para>
    /// </summary>
    public static class OutlookProfileRegistry
    {
        /// <summary>
        /// HKCU root of Outlook's own settings. Office 16.0 only: Outlook 2013 (15.0) and
        /// the "new Outlook" are not supported by this product, and nothing here probes for
        /// them (see Docs/magic-numbers.md).
        /// </summary>
        public const string OutlookRootKeyPath = @"Software\Microsoft\Office\16.0\Outlook";

        /// <summary>Profiles container, relative to <see cref="OutlookRootKeyPath"/>.</summary>
        public const string ProfilesSubKeyName = "Profiles";

        /// <summary>
        /// The per-account container inside a profile. An Outlook-internal GUID, stable
        /// across every Outlook version this product supports, and not documented anywhere
        /// but the profile itself.
        /// </summary>
        public const string AccountsSubKeyName = "9375CFF0413111d3B88A00104B2A6676";
    }
}
