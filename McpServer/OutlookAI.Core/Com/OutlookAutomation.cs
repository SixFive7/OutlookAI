using System;
using System.Runtime.InteropServices;
using OutlookAI.Core.AddIn;

namespace OutlookAI.Core.Com
{
    /// <summary>
    /// Outlook's automation object and a store of it found by name, in one place for everything that drives Outlook
    /// through COM - the product's session, the live tests' helpers and the corpus tool, which each wrote both out
    /// (the code-sharing refactor's S5b, the survey's T14). Not part of the add-in's share of Core: the add-in runs
    /// inside Outlook and never starts one.
    /// </summary>
    public static class OutlookAutomation
    {
        /// <summary>Outlook's <c>Outlook.Application</c> object, starting Outlook when it is not running.</summary>
        public static object CreateApplication()
        {
            Type progIdType = Type.GetTypeFromProgID("Outlook.Application")
                ?? throw new InvalidOperationException("Outlook.Application ProgID is not registered.");
            return Activator.CreateInstance(progIdType)
                ?? throw new InvalidOperationException("Failed to create Outlook.Application.");
        }

        /// <summary>
        /// The store of <paramref name="stores"/> whose display name is <paramref name="storeDisplayName"/>, compared
        /// ignoring case, or null. A store whose name will not read - a COM failure, or a value that is not a string -
        /// is passed over. Every store it read and did not return is released; the one returned is the caller's.
        /// </summary>
        public static object? FindStore(object stores, string storeDisplayName)
        {
            dynamic all = stores;
            int storeCount = all.Count;
            for (int i = 1; i <= storeCount; i++)
            {
                object candidate = all[i];
                string? name = null;
                try
                {
                    name = (string?)((dynamic)candidate).DisplayName;
                }
                catch (Exception ex) when (ex is COMException or Microsoft.CSharp.RuntimeBinder.RuntimeBinderException or InvalidCastException)
                {
                }

                if (string.Equals(name, storeDisplayName, StringComparison.OrdinalIgnoreCase))
                {
                    return candidate;
                }

                ComObjects.Release(candidate);
            }

            return null;
        }
    }
}
