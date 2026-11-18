using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace OutlookAI.Core.Services
{
    /// <summary>
    /// list_signatures' answer, built from a <see cref="SignatureEnvironment"/> (Q147, Q155 (c), Q159):
    /// every profile with its accounts and their own and effective defaults, the folder, the policy,
    /// roaming, every signature - and, asked by name, one signature in full: every rendition and its
    /// pictures. Pure over the environment, apart from reading the one signature's files.
    /// </summary>
    public static class SignatureListing
    {
        /// <summary>The steering every listing carries: how the draft tools' signature argument reads.</summary>
        public const string ChoosingNote =
            "Draft tools: leave 'signature' out to get the sending account's effective default (effectiveNewMessage / "
            + "effectiveReplyForward of that account in the running profile), pass \"none\" for no signature, or pass a name "
            + "from 'signatures' that the account can use - pick by language and purpose from the excerpts. The draft's "
            + "result says which signature went in and why.";

        /// <summary>Builds the listing. <paramref name="name"/> asks for one signature in full.</summary>
        public static SignaturesOutcome Build(SignatureEnvironment environment, string? runningProfile, string? runningProfileNote, string? name)
        {
            if (environment == null)
            {
                throw new ArgumentNullException(nameof(environment));
            }

            string? running = environment.FindProfile(runningProfile) ?? runningProfile;
            List<SignatureProfileView> profiles = environment.ProfileNames
                .Select(p => ProfileView(environment, p, running))
                .ToList();

            Dictionary<string, List<SignatureDefaultUseView>> uses = DefaultUses(environment);
            List<SignatureView> signatures = environment.Signatures
                .Select(s => new SignatureView
                {
                    Name = s.Name,
                    Excerpt = s.Excerpt,
                    Formats = s.Formats,
                    Pictures = s.PictureCount > 0 ? s.PictureCount : (int?)null,
                    Account = s.ScopeAccount,
                    ManagedInMailbox = environment.IsManagedInMailbox(s) ? true : (bool?)null,
                    DefaultFor = uses.TryGetValue(s.Name, out List<SignatureDefaultUseView>? used) ? used : null,
                })
                .ToList();

            SignatureDetailView? detail = null;
            if (!string.IsNullOrWhiteSpace(name))
            {
                SignatureInfo signature = environment.FindSignature(name)
                    ?? throw new ArgumentException(
                        "signature '" + name!.Trim() + "' was not found in " + environment.Folder.Path
                        + (environment.Signatures.Count > 0
                            ? ". Installed: " + string.Join(", ", environment.Signatures.Select(s => s.Name)) + "."
                            : " - no signatures are installed."));
                detail = Detail(environment, signature, uses.TryGetValue(signature.Name, out List<SignatureDefaultUseView>? signatureUses) ? signatureUses : null);
            }

            return new SignaturesOutcome
            {
                Signatures = signatures,
                Profiles = profiles,
                DefaultProfile = environment.FindProfile(environment.Profiles.DefaultProfile) ?? environment.Profiles.DefaultProfile,
                RunningProfile = running,
                RunningProfileNote = running == null ? runningProfileNote : null,
                Folder = new SignatureFolderView
                {
                    Path = environment.Folder.Path,
                    Source = environment.Folder.FromRegistry ? "registry" : "default",
                    Exists = Directory.Exists(environment.Folder.Path),
                },
                Policy = PolicyView(environment.Policy),
                Roaming = new SignatureRoamingView { State = environment.Roaming.Token, Reason = environment.Roaming.Reason },
                Signature = detail,
                Note = ChoosingNote,
                Problems = environment.Problems.Count > 0 ? environment.Problems : null,
            };
        }

        private static SignatureProfileView ProfileView(SignatureEnvironment environment, string profile, string? running)
        {
            SignaturePolicy policy = environment.Policy;
            List<SignatureAccountView> accounts = new List<SignatureAccountView>();
            foreach (ProfileAccountEntry entry in environment.MailAccountsOf(profile).Where(e => e.Address != null))
            {
                string? locked = DefaultsLocked(environment, entry);
                string? effectiveNew = policy.EffectiveDefault(false, entry.NewSignature, out SignaturePolicySource? newFrom);
                string? effectiveReply = policy.EffectiveDefault(true, entry.ReplyForwardSignature, out SignaturePolicySource? replyFrom);
                accounts.Add(new SignatureAccountView
                {
                    Account = entry.Address!,
                    Kind = entry.AccountKind?.ToString().ToLowerInvariant(),
                    AccountName = entry.AccountName != null && !ProfileAccountEntries.AddressEquals(entry.AccountName, entry.Address)
                        ? entry.AccountName
                        : null,
                    NewMessage = entry.NewSignature,
                    ReplyForward = entry.ReplyForwardSignature,
                    EffectiveNewMessage = policy.SignaturesDisabled ? null : effectiveNew,
                    EffectiveReplyForward = policy.SignaturesDisabled ? null : effectiveReply,
                    DefaultsFrom = policy.SignaturesDisabled ? SignaturePolicy.Token(policy.DisabledBy)
                        : SignaturePolicy.Token(newFrom ?? replyFrom) ?? "account",
                    DefaultsEditable = locked == null,
                    DefaultsLocked = locked,
                });
            }

            return new SignatureProfileView
            {
                Name = profile,
                Default = ProfileAccountEntries.ProfileEquals(profile, environment.Profiles.DefaultProfile) ? true : (bool?)null,
                Running = ProfileAccountEntries.ProfileEquals(profile, running) ? true : (bool?)null,
                Accounts = accounts,
            };
        }

        /// <summary>Why manage_signature would refuse to change this account's defaults, or null - the same rules it applies.</summary>
        private static string? DefaultsLocked(SignatureEnvironment environment, ProfileAccountEntry entry)
        {
            SignaturePolicy policy = environment.Policy;
            if (policy.SignaturesDisabled)
            {
                return "signatures are disabled by " + SignaturePolicy.Describe(policy.DisabledBy);
            }

            if (policy.TakesDefault(false) || policy.TakesDefault(true))
            {
                return SignaturePolicy.Describe(policy.TakesDefault(false) ? policy.NewMessageSource : policy.ReplyForwardSource)
                    + " names a default signature, and OutlookAI does not change an account's default under a policy (Q156)";
            }

            if (entry.AccountKind == MailAccountKind.Exchange && environment.Roaming.MayRoamFor(entry.Address))
            {
                return "an Exchange account whose signatures may roam (" + environment.Roaming.Reason + "): change its defaults in Outlook";
            }

            return null;
        }

        private static SignaturePolicyView? PolicyView(SignaturePolicy policy)
        {
            if (!policy.SignaturesDisabled && !policy.AnyDefaultSet)
            {
                return null;
            }

            return new SignaturePolicyView
            {
                SignaturesDisabled = policy.SignaturesDisabled ? true : (bool?)null,
                DisabledBy = SignaturePolicy.Token(policy.DisabledBy),
                NewMessage = policy.NewMessage,
                NewMessageFrom = SignaturePolicy.Token(policy.NewMessageSource),
                ReplyForward = policy.ReplyForward,
                ReplyForwardFrom = SignaturePolicy.Token(policy.ReplyForwardSource),
            };
        }

        private static Dictionary<string, List<SignatureDefaultUseView>> DefaultUses(SignatureEnvironment environment)
        {
            Dictionary<string, List<SignatureDefaultUseView>> uses =
                new Dictionary<string, List<SignatureDefaultUseView>>(StringComparer.OrdinalIgnoreCase);
            foreach (ProfileAccountEntry entry in environment.Entries.Where(e => e.Kind == ProfileEntryKind.MailAccount && e.Address != null))
            {
                Add(uses, entry.NewSignature, entry, "new");
                Add(uses, entry.ReplyForwardSignature, entry, "reply");
            }

            return uses;
        }

        private static void Add(Dictionary<string, List<SignatureDefaultUseView>> uses, string? name, ProfileAccountEntry entry, string scope)
        {
            if (name == null)
            {
                return;
            }

            if (!uses.TryGetValue(name, out List<SignatureDefaultUseView>? list))
            {
                list = new List<SignatureDefaultUseView>();
                uses[name] = list;
            }

            list.Add(new SignatureDefaultUseView { Profile = entry.ProfileName ?? string.Empty, Account = entry.Address!, Scope = scope });
        }

        /// <summary>One signature in full: every rendition, decoded as stored, and every picture (Q159).</summary>
        public static SignatureDetailView Detail(SignatureEnvironment environment, SignatureInfo signature, IReadOnlyList<SignatureDefaultUseView>? uses)
        {
            if (signature == null)
            {
                throw new ArgumentNullException(nameof(signature));
            }

            List<string> truncated = new List<string>();
            DecodedSignatureText? html = signature.HtmlPath != null ? SignatureText.ReadHtml(signature.HtmlPath) : null;
            DecodedSignatureText? text = signature.TextPath != null ? SignatureText.ReadText(signature.TextPath) : null;
            string? rtfSource = signature.RtfPath != null ? SignatureText.ReadRtf(signature.RtfPath).Text : null;

            IReadOnlyList<SignatureHtmlReference> references = html != null
                ? SignatureHtml.References(html.Text, signature.Name)
                : Array.Empty<SignatureHtmlReference>();
            HashSet<string> referenced = new HashSet<string>(
                references.Where(r => r.FileName != null).Select(r => r.FileName!), StringComparer.OrdinalIgnoreCase);

            List<SignaturePictureView> pictures = new List<SignaturePictureView>();
            List<string> others = new List<string>();
            if (signature.ResourceDirectory != null && Directory.Exists(signature.ResourceDirectory))
            {
                foreach (string file in Directory.GetFiles(signature.ResourceDirectory).OrderBy(f => f, StringComparer.OrdinalIgnoreCase))
                {
                    string fileName = Path.GetFileName(file);
                    if (!SignatureCatalog.IsPictureFile(fileName))
                    {
                        others.Add(fileName);
                        continue;
                    }

                    byte[] head = ReadHead(file, 64 * 1024);
                    bool sized = SignaturePictures.TryReadSize(head, out int width, out int height);
                    pictures.Add(new SignaturePictureView
                    {
                        File = fileName,
                        Path = file,
                        Bytes = new FileInfo(file).Length,
                        ContentType = SignaturePictures.ContentType(fileName),
                        Width = sized ? width : (int?)null,
                        Height = sized ? height : (int?)null,
                        ReferencedByHtml = referenced.Contains(fileName),
                    });
                }
            }

            List<string> missing = referenced
                .Where(r => !pictures.Any(p => string.Equals(p.File, r, StringComparison.OrdinalIgnoreCase))
                    && !others.Any(o => string.Equals(o, r, StringComparison.OrdinalIgnoreCase)))
                .OrderBy(r => r, StringComparer.OrdinalIgnoreCase)
                .ToList();

            List<SignatureRtfPictureView> rtfPictures = rtfSource == null
                ? new List<SignatureRtfPictureView>()
                : SignatureRtf.FindPictures(rtfSource).Select(p => new SignatureRtfPictureView
                {
                    Number = p.Number,
                    Kind = p.Kind,
                    Bytes = p.ByteCount,
                    Width = p.WidthTwips / 15,
                    Height = p.HeightTwips / 15,
                    Sha256 = p.Sha256,
                }).ToList();

            return new SignatureDetailView
            {
                Name = signature.Name,
                Formats = signature.Formats,
                Html = Cap(html?.Text, SignatureFormats.Html, truncated),
                HtmlEncoding = html?.EncodingName,
                Rtf = Cap(rtfSource == null ? null : SignatureRtf.Elide(rtfSource), SignatureFormats.Rtf, truncated),
                Text = Cap(text?.Text, SignatureFormats.Text, truncated),
                TextEncoding = text?.EncodingName,
                PictureReference = SignatureHtml.ReferenceFor(signature.Name, string.Empty),
                Pictures = pictures.Count > 0 ? pictures : null,
                OtherFiles = others.Count > 0 ? others : null,
                MissingPictures = missing.Count > 0 ? missing : null,
                RtfPictures = rtfPictures.Count > 0 ? rtfPictures : null,
                Account = signature.ScopeAccount,
                ManagedInMailbox = environment.IsManagedInMailbox(signature) ? true : (bool?)null,
                DefaultFor = uses,
                Truncated = truncated.Count > 0 ? truncated : null,
            };
        }

        private static string? Cap(string? content, string format, List<string> truncated)
        {
            if (content == null || content.Length <= MailService.BodyCharsCap)
            {
                return content;
            }

            truncated.Add(format);
            return content.Substring(0, MailService.BodyCharsCap);
        }

        private static byte[] ReadHead(string path, int max)
        {
            using FileStream stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            byte[] buffer = new byte[(int)Math.Min(max, stream.Length)];
            int read = 0;
            while (read < buffer.Length)
            {
                int n = stream.Read(buffer, read, buffer.Length - read);
                if (n == 0)
                {
                    break;
                }

                read += n;
            }

            return buffer;
        }
    }
}
