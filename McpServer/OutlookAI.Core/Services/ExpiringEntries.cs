using System;
using System.Collections.Generic;

namespace OutlookAI.Core.Services
{
    /// <summary>
    /// What the server's three short-lived stores - send confirmations
    /// (<see cref="SendConfirmationTokens"/>), resumable exhaustive scans
    /// (<see cref="ExhaustiveScanCursors"/>) and update_draft intents
    /// (<see cref="DraftUpdateIntents"/>) - each did alike when it pruned: list the entries whose
    /// time to live has run out. Each removes them its own way; until the code-sharing refactor's
    /// stage S6 (Q154) each also listed them its own way.
    /// </summary>
    internal static class ExpiringEntries
    {
        /// <summary>
        /// The keys of the entries issued longer than <paramref name="timeToLive"/> before
        /// <paramref name="now"/>, in a list of their own, so that the caller can remove them from
        /// the dictionary it is walking. Call it under the store's lock.
        /// </summary>
        public static List<TKey> ExpiredKeys<TKey, TValue>(
            IEnumerable<KeyValuePair<TKey, TValue>> entries, Func<TValue, DateTime> issuedUtc, DateTime now, TimeSpan timeToLive)
        {
            List<TKey> expired = new List<TKey>();
            foreach (KeyValuePair<TKey, TValue> entry in entries)
            {
                if (now - issuedUtc(entry.Value) > timeToLive)
                {
                    expired.Add(entry.Key);
                }
            }

            return expired;
        }
    }
}
