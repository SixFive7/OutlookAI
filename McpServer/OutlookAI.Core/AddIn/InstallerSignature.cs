using System;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Security.Cryptography.X509Certificates;

namespace OutlookAI.Core.AddIn
{
    /// <summary>
    /// THE CHECK AN INSTALLER'S SIGNATURE HAS TO PASS before the updater runs it: the add-in's
    /// <c>UpdateService</c> makes it of every installer it downloads, and <c>Tools/Publish-Release.ps1</c>
    /// makes it of the installer it has just signed, so a release is refused unless every installed copy
    /// would accept what it publishes. One source both compile (Q171 (c), decided by the maintainer
    /// 2026-11-18): until then the release script restated the WinVerifyTrust call and the rule in a
    /// copy of its own.
    ///
    /// <para>
    /// The installer is Authenticode-signed with a SELF-SIGNED certificate (CN=OutlookAI). Two things
    /// are required: the signature is cryptographically valid, so the file was not changed, and the
    /// signer is exactly the pinned certificate. Because the certificate is self-signed, WinVerifyTrust
    /// answers CERT_E_UNTRUSTEDROOT for a perfectly valid file, so that one "valid hash, untrusted
    /// root" answer is accepted beside 0, and the signer is then pinned by thumbprint. Every other
    /// answer is refused. It fails closed: any error counts as unverified, which stops the automatic
    /// update and leaves a manual one working.
    /// </para>
    ///
    /// <para>
    /// C# 5, AND NOTHING ELSE OF CORE. The release script compiles this file alone with Add-Type, which
    /// in Windows PowerShell 5.1 - the shell its self-test runs under on the build VM - is the .NET
    /// Framework's own C# 5 compiler: no <c>=&gt;</c> members, no <c>$""</c>, no <c>nameof</c>, no
    /// <c>?.</c>, no <c>?</c> on a reference type, no <c>when</c> on a catch, and no type of Core's.
    /// </para>
    /// </summary>
    public static class InstallerSignature
    {
        // WinVerifyTrust's answer for a valid signature whose chain ends in a root this machine does
        // not trust - the self-signed release certificate's, on every machine.
        private const int CertEUntrustedRoot = unchecked((int)0x800B0109);

        // What the check answers when the call itself fails: unverified.
        private const int EFail = unchecked((int)0x80004005);

        private static readonly Guid WinTrustActionGenericVerifyV2 =
            new Guid("00AAC56B-CD44-11D0-8CC2-00C04FC295EE");

        /// <summary>
        /// The updater's whole check of the installer at <paramref name="path"/>: an empty string when
        /// it may run, and otherwise the reason it may not, in the words the update line shows.
        /// </summary>
        public static string Refusal(string path, string pinnedThumbprint)
        {
            string refusal = TrustRefusal(WinVerifyTrustFile(path));
            if (refusal.Length != 0)
            {
                return refusal;
            }

            try
            {
                // The signer, read the way every installed copy reads it. .NET 9 made the call obsolete
                // (SYSLIB0057) for a loader that reads no Authenticode signature, and Core's .NET 11 build
                // compiles this file too. 1692 first: Windows PowerShell 5.1's compiler takes only a
                // numbered warning, and calls SYSLIB0057 an invalid number - a warning, which Add-Type
                // makes an error.
#pragma warning disable 1692
#pragma warning disable SYSLIB0057
                using (X509Certificate2 cert = new X509Certificate2(X509Certificate.CreateFromSignedFile(path)))
#pragma warning restore SYSLIB0057
#pragma warning restore 1692
                {
                    return SignerRefusal(cert.Thumbprint, pinnedThumbprint);
                }
            }
            catch (Exception ex)
            {
                return "Could not read installer signature: " + ex.Message;
            }
        }

        /// <summary>
        /// The first half of the rule: an empty string for the two WinVerifyTrust answers the updater
        /// accepts, 0 and CERT_E_UNTRUSTEDROOT (0x800B0109), and the reason for any other.
        /// </summary>
        public static string TrustRefusal(int trustResult)
        {
            if (trustResult == 0 || trustResult == CertEUntrustedRoot)
            {
                return "";
            }

            return "Installer signature is invalid or missing (0x" + trustResult.ToString("X8", CultureInfo.InvariantCulture) + ").";
        }

        /// <summary>
        /// The second half: an empty string when the signer's thumbprint is the pinned one, in either
        /// case, and the reason when it is not.
        /// </summary>
        public static string SignerRefusal(string signerThumbprint, string pinnedThumbprint)
        {
            if (string.Equals(signerThumbprint, pinnedThumbprint, StringComparison.OrdinalIgnoreCase))
            {
                return "";
            }

            return "Installer was not signed by the expected OutlookAI certificate.";
        }

        /// <summary>
        /// WinVerifyTrust's answer for the file at <paramref name="path"/>, asked with the generic
        /// Authenticode action, no UI, no revocation check and WTD_SAFER_FLAG; E_FAIL (0x80004005)
        /// when the call itself fails.
        /// </summary>
        public static int WinVerifyTrustFile(string path)
        {
            WINTRUST_FILE_INFO fileInfo = new WINTRUST_FILE_INFO
            {
                cbStruct = (uint)Marshal.SizeOf(typeof(WINTRUST_FILE_INFO)),
                pcwszFilePath = path,
                hFile = IntPtr.Zero,
                pgKnownSubject = IntPtr.Zero
            };

            IntPtr pFile = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(WINTRUST_FILE_INFO)));
            IntPtr pData = IntPtr.Zero;
            try
            {
                Marshal.StructureToPtr(fileInfo, pFile, false);

                WINTRUST_DATA data = new WINTRUST_DATA
                {
                    cbStruct = (uint)Marshal.SizeOf(typeof(WINTRUST_DATA)),
                    pPolicyCallbackData = IntPtr.Zero,
                    pSIPClientData = IntPtr.Zero,
                    dwUIChoice = 2,            // WTD_UI_NONE
                    fdwRevocationChecks = 0,   // WTD_REVOKE_NONE
                    dwUnionChoice = 1,         // WTD_CHOICE_FILE
                    pFile = pFile,
                    dwStateAction = 0,         // WTD_STATEACTION_IGNORE
                    hWVTStateData = IntPtr.Zero,
                    pwszURLReference = IntPtr.Zero,
                    dwProvFlags = 0x00000100,  // WTD_SAFER_FLAG
                    dwUIContext = 0,
                    pSignatureSettings = IntPtr.Zero
                };

                pData = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(WINTRUST_DATA)));
                Marshal.StructureToPtr(data, pData, false);
                return WinVerifyTrust(IntPtr.Zero, WinTrustActionGenericVerifyV2, pData);
            }
            catch
            {
                // Fail closed: any marshalling or P/Invoke problem means "not verified".
                return EFail;
            }
            finally
            {
                if (pData != IntPtr.Zero)
                {
                    Marshal.FreeHGlobal(pData);
                }

                Marshal.DestroyStructure(pFile, typeof(WINTRUST_FILE_INFO));
                Marshal.FreeHGlobal(pFile);
            }
        }

        // From System32 and nowhere else, whichever assembly this is compiled into: Core and the add-in
        // say so for every call they make (CA5392), and an assembly Add-Type compiles says nothing.
        [DllImport("wintrust.dll", ExactSpelling = true, CharSet = CharSet.Unicode)]
        [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
        private static extern int WinVerifyTrust(IntPtr hwnd, [MarshalAs(UnmanagedType.LPStruct)] Guid pgActionID, IntPtr pWVTData);

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct WINTRUST_FILE_INFO
        {
            public uint cbStruct;
            [MarshalAs(UnmanagedType.LPWStr)] public string pcwszFilePath;
            public IntPtr hFile;
            public IntPtr pgKnownSubject;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct WINTRUST_DATA
        {
            public uint cbStruct;
            public IntPtr pPolicyCallbackData;
            public IntPtr pSIPClientData;
            public uint dwUIChoice;
            public uint fdwRevocationChecks;
            public uint dwUnionChoice;
            public IntPtr pFile;
            public uint dwStateAction;
            public IntPtr hWVTStateData;
            public IntPtr pwszURLReference;
            public uint dwProvFlags;
            public uint dwUIContext;
            public IntPtr pSignatureSettings;
        }
    }
}
