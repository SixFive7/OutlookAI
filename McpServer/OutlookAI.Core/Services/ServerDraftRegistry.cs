using System;
using System.Collections.Generic;

namespace OutlookAI.Core.Services
{
    /// <summary>
    /// The guardrail that makes <c>discard_draft</c> safe (v3.MD D46 / S1 v3): a
    /// PER-PROCESS record of the draft EntryIDs THIS SERVER produced - every item
    /// created by new_draft / reply_draft / replyall_draft / forward_draft and every
    /// draft update_draft touched. <c>discard_draft</c> refuses anything that is not in
    /// here, so the one deletion-capable tool in the product can only ever reach mail
    /// the agent itself authored in this session; a mail the user wrote, an incoming
    /// mail, a sent item and a draft from a previous server process are all structurally
    /// out of reach, not merely policy-protected.
    /// <para>
    /// EntryIDs are NOT stable (v3.MD section 12: any move or relocate mints a new one),
    /// so <see cref="Replace"/> re-keys an entry when an operation changes the id, and
    /// registration is always by the id the item carries AFTER the operation settled.
    /// </para>
    /// Deliberately in-memory and unpersisted: a server restart empties it, exactly like
    /// hit ids and send tokens - a fresh process must not inherit deletion rights over
    /// items it cannot vouch for.
    /// </summary>
    public sealed class ServerDraftRegistry
    {
        /// <summary>Maximum remembered drafts; the oldest registration is evicted beyond it.</summary>
        public const int Capacity = 512;

        private readonly object _lock = new object();
        private readonly Dictionary<string, LinkedListNode<string>> _index =
            new Dictionary<string, LinkedListNode<string>>(StringComparer.OrdinalIgnoreCase);
        private readonly LinkedList<string> _order = new LinkedList<string>();

        /// <summary>Number of drafts currently remembered (test/diagnostic aid).</summary>
        public int Count
        {
            get
            {
                lock (_lock)
                {
                    return _index.Count;
                }
            }
        }

        /// <summary>
        /// Records a draft this server created or updated. Blank ids are ignored (a
        /// snapshot that could not read an EntryID must never widen the allowlist).
        /// </summary>
        public void Register(string? entryId)
        {
            if (string.IsNullOrWhiteSpace(entryId))
            {
                return;
            }

            string key = entryId!.Trim();
            lock (_lock)
            {
                if (_index.TryGetValue(key, out LinkedListNode<string>? existing))
                {
                    _order.Remove(existing);
                    _order.AddLast(existing);
                    return;
                }

                LinkedListNode<string> node = _order.AddLast(key);
                _index[key] = node;
                while (_index.Count > Capacity)
                {
                    LinkedListNode<string>? oldest = _order.First;
                    if (oldest == null)
                    {
                        break;
                    }

                    _order.Remove(oldest);
                    _index.Remove(oldest.Value);
                }
            }
        }

        /// <summary>
        /// Re-keys a remembered draft whose EntryID changed (a relocate/move mints a new
        /// id). The new id is registered whether or not the old one was known, so an
        /// update can never LOSE the right to discard what it just rewrote.
        /// </summary>
        public void Replace(string? oldEntryId, string? newEntryId)
        {
            if (!string.IsNullOrWhiteSpace(oldEntryId)
                && !string.Equals(oldEntryId, newEntryId, StringComparison.OrdinalIgnoreCase))
            {
                Forget(oldEntryId);
            }

            Register(newEntryId);
        }

        /// <summary>True when this server process created or last updated that draft.</summary>
        public bool Contains(string? entryId)
        {
            if (string.IsNullOrWhiteSpace(entryId))
            {
                return false;
            }

            lock (_lock)
            {
                return _index.ContainsKey(entryId!.Trim());
            }
        }

        /// <summary>Drops a registration (after a discard - the item no longer exists there).</summary>
        public void Forget(string? entryId)
        {
            if (string.IsNullOrWhiteSpace(entryId))
            {
                return;
            }

            lock (_lock)
            {
                if (_index.TryGetValue(entryId!.Trim(), out LinkedListNode<string>? node))
                {
                    _order.Remove(node);
                    _index.Remove(node.Value);
                }
            }
        }
    }
}
