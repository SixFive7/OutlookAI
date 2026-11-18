using System;
using System.Collections.Generic;

namespace OutlookAI.Core.AddIn
{
    /// <summary>
    /// A mail item's recipients, read once, the one way Core and the add-in read them (the code-sharing
    /// refactor's stage S4, survey groups A08 and S04): the kind, the display name, and the SMTP address
    /// where Outlook knows one. Until then the task pane read <c>Recipient.Address</c>, which for an
    /// Exchange recipient is the X.500 address, and OutlookComSession read the same loop twice.
    /// </summary>
    public static class MailRecipients
    {
        /// <summary>
        /// PR_SMTP_ADDRESS, read through the recipient's PropertyAccessor: the SMTP address of a recipient
        /// whose <c>Address</c> is an Exchange (X.500) one.
        /// </summary>
        public const string SmtpAddressProperty = "http://schemas.microsoft.com/mapi/proptag/0x39FE001F";

        /// <summary>
        /// Every recipient of <paramref name="item"/>, in Outlook's order. A recipient whose type cannot be
        /// read counts as "to"; one whose SMTP address cannot be read keeps its <c>Address</c>; and a COM
        /// failure ends the list where it happened, with what was read so far. Every reference is released.
        /// </summary>
        public static List<ComRecipientInfo> Read(object item)
        {
            List<ComRecipientInfo> recipients = new List<ComRecipientInfo>();
            object? recipientsObject = null;
            try
            {
                recipientsObject = ((dynamic)item).Recipients;
                dynamic collection = (dynamic)recipientsObject!;
                int count = collection.Count;
                for (int i = 1; i <= count; i++)
                {
                    object? recipient = null;
                    try
                    {
                        recipient = collection[i];
                        dynamic r = (dynamic)recipient!;
                        int type = 1;
                        try
                        {
                            type = (int)r.Type;
                        }
                        catch (Exception ex) when (ComObjects.IsCallFailure(ex))
                        {
                        }

                        string kind = type == 2 ? "cc" : type == 3 ? "bcc" : "to";
                        string? name = TryGetString(() => (string?)r.Name);
                        string? address = TryGetSmtpAddress(recipient!) ?? TryGetString(() => (string?)r.Address);
                        recipients.Add(new ComRecipientInfo(kind, name, address));
                    }
                    finally
                    {
                        ComObjects.Release(recipient);
                    }
                }
            }
            catch (Exception ex) when (ComObjects.IsCallFailure(ex))
            {
            }
            finally
            {
                ComObjects.Release(recipientsObject);
            }

            return recipients;
        }

        private static string? TryGetSmtpAddress(object recipient)
        {
            object? accessor = null;
            try
            {
                accessor = ((dynamic)recipient).PropertyAccessor;
                object? value = ((dynamic)accessor!).GetProperty(SmtpAddressProperty);
                return value as string;
            }
            catch (Exception ex) when (ComObjects.IsCallFailure(ex))
            {
                return null;
            }
            finally
            {
                ComObjects.Release(accessor);
            }
        }

        private static string? TryGetString(Func<string?> getter)
        {
            try
            {
                return getter();
            }
            catch (Exception ex) when (ComObjects.IsCallFailure(ex))
            {
                return null;
            }
        }
    }

    /// <summary>One recipient of a mail item (COM-free data).</summary>
    public sealed class ComRecipientInfo
    {
        /// <summary>Creates a recipient snapshot.</summary>
        public ComRecipientInfo(string kind, string? name, string? address)
        {
            Kind = kind;
            Name = name;
            Address = address;
        }

        /// <summary>"to", "cc" or "bcc".</summary>
        public string Kind { get; }

        /// <summary>Display name.</summary>
        public string? Name { get; }

        /// <summary>SMTP address when resolvable, otherwise the raw provider address.</summary>
        public string? Address { get; }
    }
}
