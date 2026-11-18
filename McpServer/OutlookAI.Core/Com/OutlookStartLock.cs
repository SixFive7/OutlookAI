using System;
using System.IO;
using System.Threading;

namespace OutlookAI.Core.Com
{
    /// <summary>
    /// One OutlookAI connection to Outlook at a time, across every OutlookAI process of the user: only
    /// one may start Outlook, and the others wait (Q175 (e), decided by the maintainer 2026-11-18).
    /// <para>
    /// <b>Why.</b> Every Claude Code session runs its own server - eight were running on the maintainer's
    /// workstation on 2026-11-15 - and each opens its own connection. With Outlook closed, two of them
    /// asked at once would both start it and both make their first calls into it while its add-ins were
    /// still loading: the race behind that day's hang, twice over. So a connection holds this lock from
    /// before it looks whether Outlook runs until it is connected or has given up - start, settle wait
    /// and first calls included - and one that finds it held waits, then attaches to the Outlook the
    /// holder started. Connections to an Outlook that started long ago take it too; they hold it for
    /// well under a second.
    /// </para>
    /// <para>
    /// A named mutex in the session's namespace, which is where Outlook itself runs once per user. A
    /// holder that dies - the server ends its COM host at the connect deadline - leaves it abandoned, and
    /// Windows hands it to the next connection, which says so in the start log. Owned by a thread: the
    /// connection takes and releases it on its own STA thread, inside one work item.
    /// </para>
    /// </summary>
    public sealed class OutlookStartLock : IDisposable
    {
        /// <summary>The mutex's name: one per logon session, like Outlook.</summary>
        public const string DefaultName = @"Local\OutlookAI.OutlookStart";

        private readonly Mutex? _mutex;
        private bool _disposed;

        private OutlookStartLock(Mutex? mutex, bool previousHolderDied)
        {
            _mutex = mutex;
            PreviousHolderDied = previousHolderDied;
        }

        /// <summary>Whether the connection before this one died holding the lock.</summary>
        public bool PreviousHolderDied { get; }

        /// <summary>
        /// False for the stand-in handed out when the mutex itself could not be made - another object
        /// squatting on the name, say. The connection then goes ahead without it rather than fail: the lock
        /// narrows a race, and refusing every connection over a lock nobody can take would be an outage.
        /// </summary>
        public bool Held => _mutex != null;

        /// <summary>
        /// Takes the lock if it is free, without waiting; null when another connection holds it. Never
        /// throws: a lock that cannot be made at all comes back as a stand-in (<see cref="Held"/> false).
        /// Waiting is the caller's, so that it can keep its thread's messages flowing and keep looking at
        /// Outlook while it waits (<see cref="OutlookStartGuard"/>).
        /// </summary>
        public static OutlookStartLock? TryAcquire(string name = DefaultName)
        {
            Mutex? mutex = null;
            try
            {
                mutex = new Mutex(initiallyOwned: false, name);
                try
                {
                    if (mutex.WaitOne(0))
                    {
                        return new OutlookStartLock(mutex, previousHolderDied: false);
                    }
                }
                catch (AbandonedMutexException)
                {
                    // Ownership passed to this thread: the connection before it died holding the lock.
                    return new OutlookStartLock(mutex, previousHolderDied: true);
                }

                mutex.Dispose();
                return null;
            }
            catch (Exception ex) when (ex is UnauthorizedAccessException || ex is IOException
                || ex is WaitHandleCannotBeOpenedException || ex is ArgumentException)
            {
                mutex?.Dispose();
                return new OutlookStartLock(null, previousHolderDied: false);
            }
        }

        /// <summary>
        /// Whether another connection holds the lock right now - for the gate in front of each request,
        /// which takes no lock of its own. Asks by taking and at once releasing it, on the calling thread,
        /// so it can never be left holding it. A lock nobody ever made is not held.
        /// </summary>
        /// <param name="previousHolderDied">True when the look found the lock abandoned by a holder that died.</param>
        /// <param name="name">The mutex's name; tests use one of their own.</param>
        public static bool IsHeldElsewhere(out bool previousHolderDied, string name = DefaultName)
        {
            previousHolderDied = false;
            Mutex? mutex = null;
            try
            {
                if (!Mutex.TryOpenExisting(name, out mutex))
                {
                    return false;
                }

                bool acquired;
                try
                {
                    acquired = mutex.WaitOne(0);
                }
                catch (AbandonedMutexException)
                {
                    previousHolderDied = true;
                    acquired = true;
                }

                if (!acquired)
                {
                    return true;
                }

                mutex.ReleaseMutex();
                return false;
            }
            catch (Exception ex) when (ex is UnauthorizedAccessException || ex is IOException
                || ex is WaitHandleCannotBeOpenedException || ex is ArgumentException || ex is ApplicationException)
            {
                // Unknowable: say "not held", which leaves the decision to the connection's own lock.
                return false;
            }
            finally
            {
                mutex?.Dispose();
            }
        }

        /// <summary>Releases the lock. Must run on the thread that took it. Idempotent.</summary>
        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;
            if (_mutex == null)
            {
                return;
            }

            try
            {
                _mutex.ReleaseMutex();
            }
            catch (ApplicationException)
            {
                // Not owned by this thread - only possible through misuse; the handle is closed below,
                // and a mutex whose owner thread ends is released as abandoned.
            }

            _mutex.Dispose();
        }
    }
}
