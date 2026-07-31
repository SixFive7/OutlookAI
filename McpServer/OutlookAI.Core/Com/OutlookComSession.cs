using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

using OutlookAI.Core.IndexSearch;

namespace OutlookAI.Core.Com
{
    /// <summary>Data snapshot of an Outlook store (COM-free once returned).</summary>
    public sealed class ComStoreInfo
    {
        internal ComStoreInfo(string displayName, string storeId)
        {
            DisplayName = displayName;
            StoreId = storeId;
        }

        /// <summary>Store display name as shown in the folder pane.</summary>
        public string DisplayName { get; }

        /// <summary>StoreID hex string for Namespace.GetItemFromID's second argument.</summary>
        public string StoreId { get; }
    }

    /// <summary>Data snapshot of an opened item (COM-free once returned).</summary>
    public sealed class ComOpenResult
    {
        internal ComOpenResult(string entryId, string? subject, DateTime? receivedTime, int? itemClass)
        {
            EntryId = entryId;
            Subject = subject;
            ReceivedTime = receivedTime;
            ItemClass = itemClass;
        }

        /// <summary>The opened item's EntryID as reported by the object model.</summary>
        public string EntryId { get; }

        /// <summary>Item subject (null when the property is absent).</summary>
        public string? Subject { get; }

        /// <summary>ReceivedTime as reported by COM (local wall time, Kind=Unspecified).</summary>
        public DateTime? ReceivedTime { get; }

        /// <summary>OlObjectClass value (43 = olMail); null when unavailable.</summary>
        public int? ItemClass { get; }
    }

    /// <summary>One item captured by a store walk (COM-free once returned).</summary>
    public sealed class ComWalkedItem
    {
        internal ComWalkedItem(string entryId, string? subject, string? body, DateTime? receivedTime, string folderPath, int itemClass)
        {
            EntryId = entryId;
            Subject = subject;
            Body = body;
            ReceivedTime = receivedTime;
            FolderPath = folderPath;
            ItemClass = itemClass;
        }

        /// <summary>EntryID hex string.</summary>
        public string EntryId { get; }

        /// <summary>Subject (null when absent).</summary>
        public string? Subject { get; }

        /// <summary>Plain-text body. Only request walks with bodies on the designated test store (v3.MD S2/S4).</summary>
        public string? Body { get; }

        /// <summary>ReceivedTime (local wall time) or null (e.g. drafts).</summary>
        public DateTime? ReceivedTime { get; }

        /// <summary>Store-relative folder path, segments joined with '/'.</summary>
        public string FolderPath { get; }

        /// <summary>OlObjectClass value (43 = olMail).</summary>
        public int ItemClass { get; }
    }

    /// <summary>
    /// Outlook COM session: late-bound dynamic dispatch (no PIA - keeps dotnet build
    /// working), all COM confined to ONE dedicated STA thread that runs a REAL message
    /// pump (<see cref="PumpedStaRunner"/>, the v3.MD section-0.5.2 obligation). Started
    /// in Phase 1 for decode verification and the completeness-oracle walk; Phase 2 adds
    /// the read/attachment/accounts/folders/gap-sweep surface behind the MCP tools.
    /// Never quits or kills Outlook (S7); may start it when allowed (D17), except while
    /// the OutlookAISetup installer mutex is held.
    /// </summary>
    public sealed class OutlookComSession : IDisposable
    {
        private const int OlMailItemClass = 43;

        /// <summary>
        /// How many newest Deleted Items entries discard_draft inspects when re-locating a
        /// just-discarded draft (D46/C2). Bounded on purpose: the new EntryID is a
        /// convenience for undo, never a correctness requirement, and Deleted Items on a
        /// real mailbox holds tens of thousands of items.
        /// </summary>
        private const int DiscardRelocateScanCap = 40;

        /// <summary>
        /// PR_CONVERSATION_TOPIC. The object model exposes ConversationTopic read-only,
        /// so preserving a thread's grouping across a subject override goes through the
        /// PropertyAccessor with the MAPI proptag DASL (batch A - A3).
        /// </summary>
        private const string ConversationTopicDasl = "http://schemas.microsoft.com/mapi/proptag/0x0070001F";

        /// <summary>
        /// PR_CONVERSATION_INDEX (PT_BINARY). LIVE-PROVEN on this build (batch A - A3):
        /// assigning <c>MailItem.Subject</c> on a derived draft makes Outlook REGENERATE
        /// the conversation index header, which detaches the draft from its thread. The
        /// correct child index produced by Reply()/ReplyAll()/Forward() is captured before
        /// the subject write and restored through the PropertyAccessor afterwards.
        /// </summary>
        private const string ConversationIndexDasl = "http://schemas.microsoft.com/mapi/proptag/0x00710102";

        private readonly PumpedStaRunner _runner;
        private object? _application;
        private object? _namespace;
        private bool _disposed;
        private Process? _watchedProcess;
        private OutlookQuitSink? _quitSink;
        private OutlookQuitSinkRegistration? _quitSinkRegistration;
        private Action<OutlookComSession>? _onOutlookGone;
        private int _outlookGoneSignaled;

        /// <summary>
        /// D49: the non-displayed Explorer that keeps a window-less Outlook alive across an
        /// <c>Inspector.Close</c>. Released (NEVER Closed - closing the last Explorer is
        /// what makes Outlook exit) when the session is disposed.
        /// </summary>
        private object? _composeSurfacePin;

        private OutlookComSession(PumpedStaRunner runner, bool startedOutlook)
        {
            _runner = runner;
            StartedOutlook = startedOutlook;
        }

        /// <summary>True when connecting had to launch a new OUTLOOK.EXE process.</summary>
        public bool StartedOutlook { get; }

        /// <summary>PID of the OUTLOOK.EXE this session watches (null when the probe failed).</summary>
        public int? OutlookProcessId { get; private set; }

        /// <summary>True when the Application Quit event sink is advised (SF-2 proactive release path).</summary>
        public bool QuitSinkActive { get; private set; }

        /// <summary>
        /// D49: true when THIS session created the non-displayed Explorer that keeps a
        /// window-less Outlook alive. False when a window/Explorer already existed (nothing
        /// to pin) or the attempt failed - see <see cref="ComposeSurfacePinError"/>.
        /// </summary>
        public bool ComposeSurfacePinned { get; private set; }

        /// <summary>D49: content-free reason the process pin could not be created, else null.</summary>
        public string? ComposeSurfacePinError { get; private set; }

        /// <summary>
        /// D49: relinquishes the lifetime pin - closes every Explorer when NONE of them
        /// has a visible window, i.e. when the only thing keeping Outlook alive is the
        /// invisible surface this server holds. Returns how many were closed.
        /// <para>
        /// Refuses (returns 0) the moment any Outlook window is visible: closing a window
        /// the user can see is never this method's business, and the check is the same
        /// user-owned/not-user-owned rule the promotion uses.
        /// </para>
        /// <para>
        /// Note that closing the last Explorer is precisely what makes a window-less
        /// Outlook exit, which is why the normal Dispose path only RELEASES the pin. This
        /// exists so the disconnect-recovery suite can still stage a real Outlook exit -
        /// the scenario the pin otherwise (deliberately) prevents.
        /// </para>
        /// </summary>
        public int TryCloseInvisibleExplorers()
        {
            EnsureNotDisposed();
            return _runner.Run(() =>
            {
                if (ComposeSurface.CountUserVisibleWindows() > 0)
                {
                    return 0;
                }

                object? explorers = null;
                try
                {
                    explorers = ((dynamic)_application!).Explorers;
                    int count = (int)((dynamic)explorers!).Count;
                    int closed = 0;

                    // Reverse order: closing an Explorer removes it from the collection.
                    for (int i = count; i >= 1; i--)
                    {
                        object? explorer = null;
                        try
                        {
                            explorer = ((dynamic)explorers!).Item(i);
                            ((dynamic)explorer!).Close();
                            closed++;
                        }
                        catch (Exception ex) when (IsComCallFailure(ex))
                        {
                        }
                        finally
                        {
                            Release(explorer);
                        }
                    }

                    ComposeSurface.ForgetPin(_composeSurfacePin);
                    Release(_composeSurfacePin);
                    _composeSurfacePin = null;
                    ComposeSurfacePinned = false;
                    return closed;
                }
                catch (Exception ex) when (IsComCallFailure(ex))
                {
                    return 0;
                }
                finally
                {
                    Release(explorers);
                }
            });
        }

        /// <summary>True when an OUTLOOK.EXE process exists for this session's user.</summary>
        public static bool IsOutlookProcessRunning()
        {
            Process[] processes = Process.GetProcessesByName("OUTLOOK");
            try
            {
                return processes.Length > 0;
            }
            finally
            {
                foreach (Process process in processes)
                {
                    process.Dispose();
                }
            }
        }

        /// <summary>
        /// True when the add-in installer's SetupMutex is held (an add-in update is in
        /// progress) - starting Outlook is then forbidden (D17).
        /// </summary>
        public static bool IsInstallerMutexHeld()
        {
            return MutexExists("OutlookAISetup") || MutexExists(@"Global\OutlookAISetup");
        }

        private static bool MutexExists(string name)
        {
            Mutex? mutex = null;
            try
            {
                return Mutex.TryOpenExisting(name, out mutex);
            }
            catch (UnauthorizedAccessException)
            {
                // Exists but inaccessible - treat as held.
                return true;
            }
            finally
            {
                mutex?.Dispose();
            }
        }

        /// <summary>
        /// Connects to Outlook via the COM object model. Outlook is a single-instance COM
        /// server: creating Outlook.Application attaches to the running instance, or starts
        /// OUTLOOK.EXE when none is running (allowed per S7/D17;
        /// <paramref name="allowStartingOutlook"/>=false refuses instead). Never restarts
        /// or closes an existing instance.
        /// </summary>
        public static OutlookComSession Connect(bool allowStartingOutlook = true, Action<OutlookComSession>? onOutlookGone = null)
        {
            bool wasRunning = IsOutlookProcessRunning();
            if (!wasRunning)
            {
                if (IsInstallerMutexHeld())
                {
                    throw new InvalidOperationException(
                        "Outlook is not running and the OutlookAISetup installer mutex is held - retry after the add-in update completes (D17).");
                }

                if (!allowStartingOutlook)
                {
                    throw new InvalidOperationException("Outlook is not running and starting it was not allowed.");
                }
            }

            PumpedStaRunner runner = new PumpedStaRunner("OutlookAI.ComGateway.Sta");
            OutlookComSession session = new OutlookComSession(runner, startedOutlook: !wasRunning);
            session._onOutlookGone = onOutlookGone;
            try
            {
                runner.Run(() =>
                {
                    Type progIdType = Type.GetTypeFromProgID("Outlook.Application")
                        ?? throw new InvalidOperationException("Outlook.Application ProgID is not registered.");
                    session._application = Activator.CreateInstance(progIdType)
                        ?? throw new InvalidOperationException("Failed to create Outlook.Application.");
                    dynamic app = session._application;
                    dynamic ns = app.GetNamespace("MAPI");
                    session._namespace = (object)ns;
                    try
                    {
                        // Binds the default profile when Outlook was cold-started by us; a
                        // no-op/benign failure when a session is already active.
                        ns.Logon(Type.Missing, Type.Missing, false, false);
                    }
                    catch (COMException)
                    {
                    }

                    // SF-2, part 1 - DEFENSE-IN-DEPTH: advise the Application Quit sink on
                    // the pumped STA (section 0.5.2). Probe-measured 2026-07-23: on this
                    // build the event does NOT fire for a programmatic Quit from another
                    // client (Outlook parks instead) - the process-exit watcher below is
                    // the load-bearing signal; the sink covers any path that does raise
                    // the event, at near-zero cost. Best-effort either way.
                    try
                    {
                        session._quitSink = new OutlookQuitSink(session.SignalOutlookGone);
                        session._quitSinkRegistration = OutlookQuitSink.TryAdvise(session._application, session._quitSink);
                        session.QuitSinkActive = session._quitSinkRegistration != null;
                    }
                    catch (Exception ex) when (IsComCallFailure(ex))
                    {
                        session.QuitSinkActive = false;
                    }

                    // D49 THE PROCESS PIN. A window-less Outlook EXITS the moment the only
                    // Inspector closes (Phase-1 row 1) - which is the compose path's own
                    // last step, so every headless draft was killing the instance it had
                    // just used, taking update_draft (com_failure) and three live-suite
                    // collections (RPC_S_SERVER_UNAVAILABLE) with it. A NON-DISPLAYED
                    // Explorer owns an invisible window that keeps the process alive; it
                    // shows nothing, leaves Process.MainWindowHandle zero (so outlook_health
                    // still reports headless), and does not cost the user promotability -
                    // launching outlook.exe with the pin held still opens a normal window.
                    session._composeSurfacePin = ComposeSurface.TryPinProcess(
                        session._application!, session._namespace!, out string? pinError);
                    session.ComposeSurfacePinned = session._composeSurfacePin != null;
                    session.ComposeSurfacePinError = pinError;
                });

                // SF-2 fix, part 2: watch the OUTLOOK.EXE process itself (crash / hard
                // exit path - no Quit event fires then). Outlook is single-instance per
                // session, so the first process by name is THE instance we attached to.
                session.WireProcessExitWatch();
                return session;
            }
            catch
            {
                session.Dispose();
                throw;
            }
        }

        private void WireProcessExitWatch()
        {
            try
            {
                Process[] processes = Process.GetProcessesByName("OUTLOOK");
                for (int i = 1; i < processes.Length; i++)
                {
                    processes[i].Dispose();
                }

                if (processes.Length == 0)
                {
                    return;
                }

                _watchedProcess = processes[0];
                OutlookProcessId = _watchedProcess.Id;
                _watchedProcess.EnableRaisingEvents = true;
                _watchedProcess.Exited += OnWatchedProcessExited;
                if (_watchedProcess.HasExited)
                {
                    // Died between attach and wiring - Exited may already have fired
                    // before the handler was added; signal explicitly (idempotent).
                    SignalOutlookGone();
                }
            }
            catch (Exception ex) when (!(ex is OutOfMemoryException))
            {
                // Watching is an SF-2 hardening layer; the per-call liveness ping in
                // ComGateway.GetOrConnect still covers a silent death.
            }
        }

        private void OnWatchedProcessExited(object? sender, EventArgs e)
        {
            SignalOutlookGone();
        }

        /// <summary>
        /// Signals (once) that the attached Outlook is quitting or gone. Runs the
        /// gateway-provided callback on a worker thread; the callback disposes this
        /// session, which releases all COM refs on the STA.
        /// </summary>
        private void SignalOutlookGone()
        {
            if (System.Threading.Interlocked.Exchange(ref _outlookGoneSignaled, 1) != 0)
            {
                return;
            }

            Action<OutlookComSession>? callback = _onOutlookGone;
            if (callback == null)
            {
                return;
            }

            System.Threading.ThreadPool.QueueUserWorkItem(_ =>
            {
                try
                {
                    callback(this);
                }
                catch (Exception)
                {
                    // Detach-on-death must never take the host down.
                }
            });
        }

        /// <summary>Lists the stores of the active profile.</summary>
        public IReadOnlyList<ComStoreInfo> GetStores()
        {
            EnsureNotDisposed();
            return _runner.Run(() =>
            {
                dynamic ns = _namespace!;
                List<ComStoreInfo> result = new List<ComStoreInfo>();
                dynamic stores = ns.Stores;
                try
                {
                    int count = stores.Count;
                    for (int i = 1; i <= count; i++)
                    {
                        dynamic store = stores[i];
                        try
                        {
                            result.Add(new ComStoreInfo((string)store.DisplayName, (string)store.StoreID));
                        }
                        finally
                        {
                            Release(store);
                        }
                    }
                }
                finally
                {
                    Release(stores);
                }

                return (IReadOnlyList<ComStoreInfo>)result;
            });
        }

        /// <summary>
        /// Verify-on-open (v3.MD section 4): opens an EntryID via
        /// Namespace.GetItemFromID(entryIdHex, storeId) and snapshots Subject/ReceivedTime
        /// for comparison against the index row. Returns null on failure with a
        /// content-free error description (exception type + HRESULT only, S4).
        /// </summary>
        public ComOpenResult? TryOpenItem(string entryIdHex, string? storeId, out string? error)
        {
            EnsureNotDisposed();
            if (string.IsNullOrWhiteSpace(entryIdHex))
            {
                throw new ArgumentException("EntryID must not be blank.", nameof(entryIdHex));
            }

            string? capturedError = null;
            ComOpenResult? result = _runner.Run<ComOpenResult?>(() =>
            {
                dynamic ns = _namespace!;
                object? itemObject = null;
                try
                {
                    itemObject = storeId != null
                        ? ns.GetItemFromID(entryIdHex, storeId)
                        : ns.GetItemFromID(entryIdHex);
                    return Snapshot(itemObject);
                }
                catch (COMException ex)
                {
                    capturedError = string.Format(CultureInfo.InvariantCulture, "COMException 0x{0:X8}", ex.HResult);
                    return null;
                }
                catch (Exception ex) when (ex is InvalidOperationException || ex is ArgumentException
                    || ex is Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                {
                    capturedError = ex.GetType().Name;
                    return null;
                }
                finally
                {
                    Release(itemObject);
                }
            });

            error = capturedError;
            return result;
        }

        /// <summary>
        /// Narrow folder probe (v3.MD section 4 fallback mapping - on this machine's
        /// cached Exchange stores it is THE hit-mapping path, see <see cref="HitLocator"/>):
        /// walks the store's folder tree along <paramref name="folderPath"/>, then probes
        /// the folder with a DASL subject restriction and (when given) a ReceivedTime
        /// tolerance - an exact probe, not a scan. An empty subject falls back to a
        /// bounded time-only enumeration (small folders only). Returns the first matching
        /// item snapshot carrying the item's REAL EntryID.
        /// </summary>
        public ComOpenResult? TryResolveByPath(
            string storeDisplayName,
            IReadOnlyList<string> folderPath,
            string itemSubject,
            DateTime? indexReceivedUtc,
            out string? error,
            int toleranceSeconds = 120)
        {
            EnsureNotDisposed();
            if (string.IsNullOrWhiteSpace(storeDisplayName))
            {
                throw new ArgumentException("Store display name must not be blank.", nameof(storeDisplayName));
            }

            if (folderPath == null)
            {
                throw new ArgumentNullException(nameof(folderPath));
            }

            if (itemSubject == null)
            {
                throw new ArgumentNullException(nameof(itemSubject));
            }

            string? capturedError = null;
            ComOpenResult? result = _runner.Run<ComOpenResult?>(() =>
            {
                object? currentFolder = null;
                object? items = null;
                object? restricted = null;
                try
                {
                    currentFolder = WalkToFolder(storeDisplayName, folderPath, out string? walkError);
                    if (currentFolder == null)
                    {
                        capturedError = walkError;
                        return null;
                    }

                    dynamic folder = currentFolder!;

                    if (itemSubject.Length == 0)
                    {
                        // No subject to restrict on: bounded time-only enumeration,
                        // acceptable only for small folders (e.g. the tiny test-hub store).
                        if (!indexReceivedUtc.HasValue)
                        {
                            capturedError = "EmptySubjectAndNoReceivedTime";
                            return null;
                        }

                        items = folder.Items;
                        dynamic itemCollection = (dynamic)items!;
                        int total = itemCollection.Count;
                        if (total > 1000)
                        {
                            capturedError = "FolderTooLargeForTimeOnlyProbe";
                            return null;
                        }

                        for (int i = 1; i <= total; i++)
                        {
                            object? candidate = null;
                            try
                            {
                                candidate = itemCollection[i];
                                ComOpenResult snapshot = Snapshot(candidate);
                                if (ReceivedTimeMatches(snapshot.ReceivedTime, indexReceivedUtc.Value, toleranceSeconds)
                                    && (snapshot.Subject == null || snapshot.Subject.Length == 0))
                                {
                                    return snapshot;
                                }
                            }
                            finally
                            {
                                Release(candidate);
                            }
                        }

                        capturedError = "NoTimeOnlyMatch";
                        return null;
                    }

                    string subjectClause = "\"urn:schemas:httpmail:subject\" = '" + EscapeDaslValue(itemSubject) + "'";

                    // Tier A (Phase-2 narrowing, v3.MD section 0.8 Phase-1 guidance):
                    // Folder.GetTable with subject + received-time window - lightweight
                    // rows, no item RCWs until the final open. DASL date literals are UTC.
                    if (indexReceivedUtc.HasValue)
                    {
                        string windowed = "@SQL=(" + subjectClause + ") AND (\"urn:schemas:httpmail:datereceived\" >= '"
                            + FormatDaslUtc(indexReceivedUtc.Value.AddSeconds(-toleranceSeconds - 5)) + "' AND \"urn:schemas:httpmail:datereceived\" <= '"
                            + FormatDaslUtc(indexReceivedUtc.Value.AddSeconds(toleranceSeconds + 5)) + "')";
                        ComOpenResult? viaWindow = TryProbeViaGetTable(folder, windowed, indexReceivedUtc, toleranceSeconds);
                        if (viaWindow != null)
                        {
                            return viaWindow;
                        }
                    }

                    // Tier B: GetTable, subject-only.
                    ComOpenResult? viaSubject = TryProbeViaGetTable(folder, "@SQL=" + subjectClause, indexReceivedUtc, toleranceSeconds);
                    if (viaSubject != null)
                    {
                        return viaSubject;
                    }

                    // Tier C: legacy Items.Restrict (Phase-1 behavior) - correctness net
                    // in case GetTable misbehaves on a folder type.
                    items = folder.Items;
                    dynamic legacyItems = (dynamic)items!;
                    restricted = legacyItems.Restrict("@SQL=" + subjectClause);
                    dynamic restrictedItems = restricted!;
                    int count = restrictedItems.Count;
                    for (int i = 1; i <= count; i++)
                    {
                        object? candidate = null;
                        try
                        {
                            candidate = restrictedItems[i];
                            ComOpenResult snapshot = Snapshot(candidate);
                            if (!indexReceivedUtc.HasValue || ReceivedTimeMatches(snapshot.ReceivedTime, indexReceivedUtc.Value, toleranceSeconds))
                            {
                                return snapshot;
                            }
                        }
                        finally
                        {
                            Release(candidate);
                        }
                    }

                    capturedError = "NoSubjectTimeMatch";
                    return null;
                }
                catch (COMException ex)
                {
                    capturedError = string.Format(CultureInfo.InvariantCulture, "COMException 0x{0:X8}", ex.HResult);
                    return null;
                }
                finally
                {
                    Release(restricted);
                    Release(items);
                    Release(currentFolder);
                }
            });

            error = capturedError;
            return result;
        }

        /// <summary>
        /// STA-side folder resolution: walks a store's folder tree along the given
        /// store-relative path segments. Returns the folder RCW (CALLER must Release) or
        /// null with a content-free error. An empty path returns the store root folder.
        /// </summary>
        private object? WalkToFolder(string storeDisplayName, IReadOnlyList<string> folderPath, out string? error)
        {
            error = null;
            dynamic? store = FindStoreByDisplayName(storeDisplayName);
            if (store == null)
            {
                error = "StoreNotFound";
                return null;
            }

            object? currentFolder;
            try
            {
                currentFolder = store.GetRootFolder();
            }
            catch (Exception ex) when (IsComCallFailure(ex))
            {
                error = "RootFolderUnavailable";
                return null;
            }
            finally
            {
                Release(store);
            }

            foreach (string segment in folderPath)
            {
                object? next = null;
                dynamic folders = ((dynamic)currentFolder!).Folders;
                try
                {
                    next = folders[segment];
                }
                catch (Exception ex) when (IsComCallFailure(ex))
                {
                    error = "FolderNotFound";
                    Release(currentFolder);
                    return null;
                }
                finally
                {
                    Release(folders);
                }

                Release(currentFolder);
                currentFolder = next;
            }

            return currentFolder;
        }

        /// <summary>
        /// Opens an item and returns its attachment file names (used to verify
        /// attachment-hit parent mapping). Names stay in memory; callers must not log them
        /// for business stores (S4).
        /// </summary>
        public IReadOnlyList<string>? TryGetAttachmentFileNames(string entryIdHex, string? storeId, out string? error)
        {
            EnsureNotDisposed();
            string? capturedError = null;
            IReadOnlyList<string>? result = _runner.Run<IReadOnlyList<string>?>(() =>
            {
                dynamic ns = _namespace!;
                object? itemObject = null;
                object? attachments = null;
                try
                {
                    itemObject = storeId != null
                        ? ns.GetItemFromID(entryIdHex, storeId)
                        : ns.GetItemFromID(entryIdHex);
                    dynamic item = itemObject!;
                    attachments = item.Attachments;
                    dynamic attachmentCollection = attachments!;
                    int count = attachmentCollection.Count;
                    List<string> names = new List<string>(count);
                    for (int i = 1; i <= count; i++)
                    {
                        object? attachment = null;
                        try
                        {
                            attachment = attachmentCollection[i];
                            names.Add((string)((dynamic)attachment!).FileName);
                        }
                        catch (COMException)
                        {
                            // Some attachment types have no FileName; skip.
                        }
                        finally
                        {
                            Release(attachment);
                        }
                    }

                    return (IReadOnlyList<string>?)names;
                }
                catch (COMException ex)
                {
                    capturedError = string.Format(CultureInfo.InvariantCulture, "COMException 0x{0:X8}", ex.HResult);
                    return null;
                }
                finally
                {
                    Release(attachments);
                    Release(itemObject);
                }
            });

            error = capturedError;
            return result;
        }

        /// <summary>
        /// Total item count across every store's Outbox (the S7 quit-when-safe count -
        /// graceful tooling and the lifecycle tests refuse to close/quit Outlook while
        /// anything is pending). Stores without an Outbox (delegate caches) are skipped;
        /// returns -1 when the walk itself failed (callers treat unknown as unsafe).
        /// </summary>
        public int CountOutboxItems()
        {
            EnsureNotDisposed();
            return _runner.Run(() =>
            {
                dynamic ns = _namespace!;
                object? stores = null;
                try
                {
                    stores = (object)ns.Stores;
                    dynamic list = (dynamic)stores!;
                    int count = list.Count;
                    int total = 0;
                    for (int i = 1; i <= count; i++)
                    {
                        object? store = null;
                        object? outbox = null;
                        object? items = null;
                        try
                        {
                            store = list[i];
                            outbox = ((dynamic)store!).GetDefaultFolder(4); // olFolderOutbox
                            items = ((dynamic)outbox!).Items;
                            total += (int)((dynamic)items!).Count;
                        }
                        catch (Exception ex) when (IsComCallFailure(ex))
                        {
                            // Store without an Outbox (delegate cache) - fine.
                        }
                        finally
                        {
                            Release(items);
                            Release(outbox);
                            Release(store);
                        }
                    }

                    return total;
                }
                catch (Exception ex) when (IsComCallFailure(ex))
                {
                    return -1;
                }
                finally
                {
                    Release(stores);
                }
            });
        }

        /// <summary>Current MAPI profile name - doubles as the liveness ping for <see cref="ComGateway"/>.</summary>
        public string GetProfileName()
        {
            EnsureNotDisposed();
            return _runner.Run(() => (string)((dynamic)_namespace!).CurrentProfileName);
        }

        /// <summary>True when Outlook reports it is working offline.</summary>
        public bool IsNamespaceOffline()
        {
            EnsureNotDisposed();
            return _runner.Run(() =>
            {
                try
                {
                    return (bool)((dynamic)_namespace!).Offline;
                }
                catch (COMException)
                {
                    return false;
                }
            });
        }

        /// <summary>
        /// Full read of one item by its REAL EntryID (v3.MD section 8 L2): the COMPLETE
        /// plain-text body (HTML converted when Outlook has no text rendering; windowing
        /// happens in MailService against its body cache - soak fix D37), recipients
        /// with SMTP addresses, attachment list, and transport headers on request.
        /// <paramref name="includeBody"/>=false skips the body transfer entirely (the
        /// caller already holds a cached extraction). Returns null with a content-free
        /// error description on failure (S4).
        /// </summary>
        public ComItemDetail? TryReadItem(string entryIdHex, string? storeId, bool includeHeaders, bool includeBody, out string? error, bool includeHtml = false)
        {
            EnsureNotDisposed();
            if (string.IsNullOrWhiteSpace(entryIdHex))
            {
                throw new ArgumentException("EntryID must not be blank.", nameof(entryIdHex));
            }

            string? capturedError = null;
            ComItemDetail? result = _runner.Run<ComItemDetail?>(() =>
            {
                dynamic ns = _namespace!;
                object? itemObject = null;
                try
                {
                    itemObject = storeId != null
                        ? ns.GetItemFromID(entryIdHex, storeId)
                        : ns.GetItemFromID(entryIdHex);
                    return SnapshotDetail(itemObject!, includeHeaders, includeBody, includeHtml);
                }
                catch (COMException ex)
                {
                    capturedError = string.Format(CultureInfo.InvariantCulture, "COMException 0x{0:X8}", ex.HResult);
                    return null;
                }
                catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                {
                    capturedError = "RuntimeBinderException";
                    return null;
                }
                finally
                {
                    Release(itemObject);
                }
            });

            error = capturedError;
            return result;
        }

        /// <summary>
        /// Saves one attachment (1-based index) of an item to
        /// <paramref name="targetDirectory"/> (created if missing). Never overwrites: an
        /// existing name gets a numeric suffix. Returns the full saved path.
        /// </summary>
        public string? TrySaveAttachment(
            string entryIdHex,
            string? storeId,
            int attachmentIndex,
            string targetDirectory,
            out long sizeBytes,
            out string? error)
        {
            EnsureNotDisposed();
            if (string.IsNullOrWhiteSpace(entryIdHex))
            {
                throw new ArgumentException("EntryID must not be blank.", nameof(entryIdHex));
            }

            if (string.IsNullOrWhiteSpace(targetDirectory))
            {
                throw new ArgumentException("Target directory must not be blank.", nameof(targetDirectory));
            }

            string? capturedError = null;
            long capturedSize = 0;
            string? path = _runner.Run<string?>(() =>
            {
                dynamic ns = _namespace!;
                object? itemObject = null;
                object? attachments = null;
                object? attachment = null;
                try
                {
                    itemObject = storeId != null
                        ? ns.GetItemFromID(entryIdHex, storeId)
                        : ns.GetItemFromID(entryIdHex);
                    dynamic item = itemObject!;
                    attachments = item.Attachments;
                    dynamic attachmentCollection = (dynamic)attachments!;
                    int count = attachmentCollection.Count;
                    if (attachmentIndex < 1 || attachmentIndex > count)
                    {
                        capturedError = string.Format(
                            CultureInfo.InvariantCulture, "AttachmentIndexOutOfRange (count={0})", count);
                        return null;
                    }

                    attachment = attachmentCollection[attachmentIndex];
                    string fileName;
                    try
                    {
                        fileName = (string)((dynamic)attachment!).FileName;
                    }
                    catch (COMException)
                    {
                        fileName = "attachment";
                    }

                    Directory.CreateDirectory(targetDirectory);
                    string fullPath = MakeUniquePath(targetDirectory, SanitizeFileName(fileName));
                    ((dynamic)attachment!).SaveAsFile(fullPath);
                    capturedSize = new FileInfo(fullPath).Length;
                    return fullPath;
                }
                catch (COMException ex)
                {
                    capturedError = string.Format(CultureInfo.InvariantCulture, "COMException 0x{0:X8}", ex.HResult);
                    return null;
                }
                catch (IOException ex)
                {
                    capturedError = "IOException: " + ex.GetType().Name;
                    return null;
                }
                catch (UnauthorizedAccessException)
                {
                    capturedError = "UnauthorizedAccessException";
                    return null;
                }
                finally
                {
                    Release(attachment);
                    Release(attachments);
                    Release(itemObject);
                }
            });

            sizeBytes = capturedSize;
            error = capturedError;
            return path;
        }

        /// <summary>Lists the profile's mail accounts (list_accounts, D22/D25 flags built by the caller).</summary>
        public IReadOnlyList<ComAccountInfo> GetAccounts()
        {
            EnsureNotDisposed();
            return _runner.Run(() =>
            {
                dynamic ns = _namespace!;
                List<ComAccountInfo> result = new List<ComAccountInfo>();
                object? session = null;
                object? accounts = null;
                try
                {
                    session = ns.Session;
                    accounts = ((dynamic)session!).Accounts;
                    dynamic accountCollection = (dynamic)accounts!;
                    int count = accountCollection.Count;
                    for (int i = 1; i <= count; i++)
                    {
                        object? account = null;
                        object? deliveryStore = null;
                        try
                        {
                            account = accountCollection[i];
                            dynamic acc = (dynamic)account!;
                            string? smtp = TryGetString(() => (string?)acc.SmtpAddress);
                            string? display = TryGetString(() => (string?)acc.DisplayName);
                            string? deliveryName = null;
                            try
                            {
                                deliveryStore = acc.DeliveryStore;
                                if (deliveryStore != null)
                                {
                                    deliveryName = (string)((dynamic)deliveryStore).DisplayName;
                                }
                            }
                            catch (COMException)
                            {
                            }

                            result.Add(new ComAccountInfo(smtp, display, deliveryName));
                        }
                        finally
                        {
                            Release(deliveryStore);
                            Release(account);
                        }
                    }
                }
                finally
                {
                    Release(accounts);
                    Release(session);
                }

                return (IReadOnlyList<ComAccountInfo>)result;
            });
        }

        /// <summary>Detailed store list: display name, StoreID, Exchange type, cached flag.</summary>
        public IReadOnlyList<ComStoreDetail> GetStoreDetails()
        {
            EnsureNotDisposed();
            return _runner.Run(() =>
            {
                dynamic ns = _namespace!;
                List<ComStoreDetail> result = new List<ComStoreDetail>();
                dynamic stores = ns.Stores;
                try
                {
                    int count = stores.Count;
                    for (int i = 1; i <= count; i++)
                    {
                        dynamic store = stores[i];
                        try
                        {
                            string displayName = (string)store.DisplayName;
                            string storeId = (string)store.StoreID;
                            int? exchangeType = null;
                            bool? cached = null;
                            try
                            {
                                exchangeType = (int)store.ExchangeStoreType;
                            }
                            catch (COMException)
                            {
                            }

                            try
                            {
                                cached = (bool)store.IsCachedExchange;
                            }
                            catch (COMException)
                            {
                            }

                            result.Add(new ComStoreDetail(displayName, storeId, exchangeType, cached));
                        }
                        catch (COMException)
                        {
                            // Store with unreadable identity - skip.
                        }
                        finally
                        {
                            Release(store);
                        }
                    }
                }
                finally
                {
                    Release(stores);
                }

                return (IReadOnlyList<ComStoreDetail>)result;
            });
        }

        /// <summary>
        /// Folder tree listing (list_folders): the FULL tree of every requested store in
        /// a STABLE traversal order - stores sorted by display name, then depth-first
        /// with siblings sorted by folder name (case-insensitive ordinal) - so an offset
        /// into the flattened list pages deterministically across calls. Paths carry
        /// item/unread counts (PR_CONTENT_COUNT / PR_CONTENT_UNREAD). Bounded only by
        /// <paramref name="absoluteWalkCap"/> (a pathological-store guard, v3.MD
        /// section 12) - the per-call page bound lives in MailService.
        /// </summary>
        public IReadOnlyList<ComFolderInfo> ListFolders(string? storeDisplayName, int absoluteWalkCap)
        {
            EnsureNotDisposed();
            if (absoluteWalkCap < 1)
            {
                throw new ArgumentOutOfRangeException(nameof(absoluteWalkCap));
            }

            return _runner.Run(() =>
            {
                dynamic ns = _namespace!;
                List<ComFolderInfo> result = new List<ComFolderInfo>();
                dynamic stores = ns.Stores;
                List<(string Name, int Index)> storeOrder = new List<(string, int)>();
                try
                {
                    int count = stores.Count;
                    for (int i = 1; i <= count; i++)
                    {
                        object? probe = null;
                        try
                        {
                            probe = stores[i];
                            string name = (string)((dynamic)probe!).DisplayName;
                            if (storeDisplayName == null
                                || string.Equals(name, storeDisplayName, StringComparison.OrdinalIgnoreCase))
                            {
                                storeOrder.Add((name, i));
                            }
                        }
                        catch (Exception ex) when (IsComCallFailure(ex))
                        {
                        }
                        finally
                        {
                            Release(probe);
                        }
                    }

                    // Stable order leg 1: stores by display name (ties broken by
                    // profile position so equal names still page deterministically).
                    storeOrder.Sort((a, b) =>
                    {
                        int byName = string.Compare(a.Name, b.Name, StringComparison.OrdinalIgnoreCase);
                        return byName != 0 ? byName : a.Index.CompareTo(b.Index);
                    });

                    foreach ((string name, int index) in storeOrder)
                    {
                        if (result.Count >= absoluteWalkCap)
                        {
                            break;
                        }

                        object? store = null;
                        object? root = null;
                        try
                        {
                            store = stores[index];
                            root = ((dynamic)store!).GetRootFolder();
                            CollectFolders(root!, name, string.Empty, 1, absoluteWalkCap, result);
                        }
                        catch (Exception ex) when (IsComCallFailure(ex))
                        {
                        }
                        finally
                        {
                            Release(root);
                            Release(store);
                        }
                    }
                }
                finally
                {
                    Release(stores);
                }

                return (IReadOnlyList<ComFolderInfo>)result;
            });
        }

        /// <summary>
        /// Default folders a store-wide (or all-stores) sweep covers: the four folders
        /// mail can LAND in without any user action - Inbox (delivery), Sent Items
        /// (dispatch), Deleted Items (server-side delete/sweep rules - the live
        /// discovery case) and Junk Email (spam filter). Anything else a rule files
        /// into is reachable by scoping the search to that folder (soak fix 13).
        /// <para>
        /// Deliberately NOT the whole folder tree: measured on this machine a full walk
        /// costs ~10 ms per folder and these stores carry 41-46 folders each, i.e.
        /// seconds per search. These four cost 135 ms across all 5 stores (86 ms for
        /// the pre-fix Inbox+Sent pair).
        /// </para>
        /// <para>
        /// The set is identical for every store, which is what lets a cached all-stores
        /// sweep serve a store-scoped request (SweepCache): both cover the same folders
        /// per store.
        /// </para>
        /// </summary>
        public static readonly IReadOnlyList<string> DefaultSweepFolderKinds = new[]
        {
            "inbox", "sent", "deleted", "junk",
        };

        /// <summary>
        /// Folder cap for a folder-scoped sweep's subtree walk. Bounds the cost of
        /// scoping a search to a folder with a large subtree (~10 ms per folder).
        /// </summary>
        public const int MaxScopedSweepFolders = 40;

        // 6 = olFolderInbox, 5 = olFolderSentMail, 3 = olFolderDeletedItems, 23 = olFolderJunk.
        private static readonly (int FolderId, string Kind)[] DefaultSweepFolders =
        {
            (6, "inbox"), (5, "sent"), (3, "deleted"), (23, "junk"),
        };

        /// <summary>
        /// Fresh-mode gap sweep (v3.MD D19): enumerates the folders a search covers for
        /// items received/sent at or after <paramref name="sinceUtc"/>. Items are opened
        /// for authoritative properties and carry their REAL EntryIDs; bodies are fetched
        /// only when the caller needs term matching. Bounded by
        /// <paramref name="perFolderCap"/> per folder.
        /// <para>
        /// Scope follows the SEARCH scope (soak fix 13): with
        /// <paramref name="folderPath"/> set (which requires
        /// <paramref name="onlyStoreDisplayName"/>) the sweep covers exactly that folder
        /// AND its subfolders - the index tier's SCOPE= predicate is recursive, so the
        /// sweep must be too or a folder-scoped search would keep missing fresh mail.
        /// Without a folder path it covers <see cref="DefaultSweepFolderKinds"/> in every
        /// store (or in the named store).
        /// </para>
        /// <para>
        /// Never throws for a missing folder/store: an unresolvable scope is reported as
        /// a skipped folder so search degrades gracefully (D34) instead of failing.
        /// </para>
        /// </summary>
        public ComSweepResult SweepFoldersNewerThan(
            DateTime sinceUtc,
            int perFolderCap,
            bool includeBodies,
            string? onlyStoreDisplayName,
            IReadOnlyList<string>? folderPath = null,
            bool includeSubfolders = true)
        {
            EnsureNotDisposed();
            if (perFolderCap < 1)
            {
                throw new ArgumentOutOfRangeException(nameof(perFolderCap));
            }

            if (folderPath != null && folderPath.Count > 0 && onlyStoreDisplayName == null)
            {
                throw new ArgumentException(
                    "A folder-scoped sweep needs the store the folder lives in.",
                    nameof(onlyStoreDisplayName));
            }

            return _runner.Run(() =>
            {
                dynamic ns = _namespace!;
                SweepTally tally = new SweepTally();
                List<ComMailBrief> items = new List<ComMailBrief>();
                List<string> sweptFolders = new List<string>();
                int skipped = 0;

                if (folderPath != null && folderPath.Count > 0)
                {
                    SweepScopedFolder(
                        ns, onlyStoreDisplayName!, folderPath, sinceUtc, perFolderCap, includeBodies,
                        includeSubfolders, items, sweptFolders, ref skipped, tally);
                    return new ComSweepResult(
                        items, sweptFolders.Count, skipped, sweptFolders,
                        tally.Failed, tally.ItemCapped, tally.FolderCapReached);
                }

                dynamic stores = ns.Stores;
                try
                {
                    int count = stores.Count;
                    for (int i = 1; i <= count; i++)
                    {
                        dynamic store = stores[i];
                        try
                        {
                            string storeName;
                            string storeId;
                            try
                            {
                                storeName = (string)store.DisplayName;
                                storeId = (string)store.StoreID;
                            }
                            catch (COMException)
                            {
                                skipped += DefaultSweepFolders.Length;
                                continue;
                            }

                            if (onlyStoreDisplayName != null
                                && !string.Equals(storeName, onlyStoreDisplayName, StringComparison.OrdinalIgnoreCase))
                            {
                                continue;
                            }

                            foreach ((int folderId, string folderKind) in DefaultSweepFolders)
                            {
                                object? folder = null;
                                try
                                {
                                    folder = store.GetDefaultFolder(folderId);
                                    string label = DescribeSweptFolder(storeName, folder!, folderKind);
                                    SweepOutcome outcome = SweepFolder(
                                        ns, folder!, storeName, storeId, folderKind, sinceUtc, perFolderCap, includeBodies, items);
                                    if (outcome == SweepOutcome.Failed)
                                    {
                                        // A folder whose table could not be read has NO
                                        // freshness coverage - reporting it as swept was a
                                        // lie (section-12 no-silent-caps discipline).
                                        skipped++;
                                        tally.Failed++;
                                        continue;
                                    }

                                    sweptFolders.Add(label);
                                    if (outcome == SweepOutcome.ItemCapped)
                                    {
                                        tally.ItemCapped.Add(label);
                                    }
                                }
                                catch (Exception ex) when (IsComCallFailure(ex))
                                {
                                    // Store without that default folder (some delegate caches) - skip.
                                    skipped++;
                                }
                                finally
                                {
                                    Release(folder);
                                }
                            }
                        }
                        finally
                        {
                            Release(store);
                        }
                    }
                }
                finally
                {
                    Release(stores);
                }

                return new ComSweepResult(
                    items, sweptFolders.Count, skipped, sweptFolders,
                    tally.Failed, tally.ItemCapped, tally.FolderCapReached);
            });
        }

        /// <summary>
        /// Mutable counters a sweep collects so the caller can report partial coverage
        /// instead of implying complete coverage (section-12 no-silent-caps discipline).
        /// </summary>
        private sealed class SweepTally
        {
            internal int Failed { get; set; }

            internal List<string> ItemCapped { get; } = new List<string>();

            internal bool FolderCapReached { get; set; }
        }

        /// <summary>Why a single-folder sweep stopped.</summary>
        private enum SweepOutcome
        {
            /// <summary>Every item in the window was collected.</summary>
            Complete = 0,

            /// <summary>The per-folder item cap cut the (newest-first) list short.</summary>
            ItemCapped = 1,

            /// <summary>The folder's table could not be read - no coverage at all.</summary>
            Failed = 2,
        }

        /// <summary>
        /// STA-side folder-scoped sweep: walks to the requested folder and sweeps it,
        /// plus its subfolders when <paramref name="includeSubfolders"/> is set (bounded
        /// by <see cref="MaxScopedSweepFolders"/>). Mirrors the exhaustive scan's tree
        /// rule - only mail folders (DefaultItemType 0) are swept, but non-mail folders
        /// still get their subtrees visited.
        /// </summary>
        private void SweepScopedFolder(
            dynamic ns,
            string storeDisplayName,
            IReadOnlyList<string> folderPath,
            DateTime sinceUtc,
            int perFolderCap,
            bool includeBodies,
            bool includeSubfolders,
            List<ComMailBrief> items,
            List<string> sweptFolders,
            ref int skipped,
            SweepTally tally)
        {
            string? storeId = null;
            dynamic? store = FindStoreByDisplayName(storeDisplayName);
            if (store != null)
            {
                try
                {
                    storeId = (string)store.StoreID;
                }
                catch (Exception ex) when (IsComCallFailure(ex))
                {
                }
                finally
                {
                    Release(store);
                }
            }

            if (storeId == null)
            {
                skipped++;
                return;
            }

            object? root = WalkToFolder(storeDisplayName, folderPath, out string? walkError);
            if (root == null)
            {
                // Folder gone/renamed: the index tier still answers, so this degrades to
                // "one folder could not be swept" rather than failing the search.
                skipped++;
                return;
            }

            try
            {
                SweepFolderTree(
                    ns, root, storeDisplayName, storeId, string.Join("/", folderPath),
                    sinceUtc, perFolderCap, includeBodies, includeSubfolders,
                    items, sweptFolders, ref skipped, tally);
            }
            finally
            {
                Release(root);
            }
        }

        private void SweepFolderTree(
            dynamic ns,
            object folderObject,
            string storeName,
            string storeId,
            string relativePath,
            DateTime sinceUtc,
            int perFolderCap,
            bool includeBodies,
            bool includeSubfolders,
            List<ComMailBrief> items,
            List<string> sweptFolders,
            ref int skipped,
            SweepTally tally)
        {
            if (sweptFolders.Count >= MaxScopedSweepFolders)
            {
                skipped++;
                tally.FolderCapReached = true;
                return;
            }

            dynamic folder = folderObject;

            int defaultItemType = -1;
            try
            {
                defaultItemType = (int)folder.DefaultItemType;
            }
            catch (Exception ex) when (IsComCallFailure(ex))
            {
            }

            if (defaultItemType == 0)
            {
                string label = storeName + "/" + relativePath;
                SweepOutcome outcome = SweepFolder(
                    ns, folderObject, storeName, storeId, null, sinceUtc, perFolderCap, includeBodies, items);
                if (outcome == SweepOutcome.Failed)
                {
                    skipped++;
                    tally.Failed++;
                }
                else
                {
                    sweptFolders.Add(label);
                    if (outcome == SweepOutcome.ItemCapped)
                    {
                        tally.ItemCapped.Add(label);
                    }
                }
            }

            if (!includeSubfolders)
            {
                return;
            }

            object? subFolders = null;
            try
            {
                subFolders = folder.Folders;
                dynamic folderCollection = (dynamic)subFolders!;
                int count = folderCollection.Count;
                for (int i = 1; i <= count; i++)
                {
                    object? child = null;
                    try
                    {
                        child = folderCollection[i];
                        object childFolder = child!;
                        string childName = TryGetString(() => (string?)((dynamic)childFolder).Name) ?? "?";
                        SweepFolderTree(
                            ns, childFolder, storeName, storeId, relativePath + "/" + childName,
                            sinceUtc, perFolderCap, includeBodies, includeSubfolders,
                            items, sweptFolders, ref skipped, tally);
                    }
                    finally
                    {
                        Release(child);
                    }
                }
            }
            catch (Exception ex) when (IsComCallFailure(ex))
            {
                // A subtree we could not enumerate is a coverage hole, not a non-event:
                // count it so foldersSkipped stops under-reporting (soak fix 15).
                skipped++;
            }
            finally
            {
                Release(subFolders);
            }
        }

        private static string DescribeSweptFolder(string storeName, object folderObject, string fallbackKind)
        {
            string? name = TryGetString(() => (string?)((dynamic)folderObject).Name);
            return storeName + "/" + (string.IsNullOrEmpty(name) ? fallbackKind : name!);
        }

        /// <summary>
        /// COM conversation walk (thread tool fallback): opens the item, gets its
        /// Conversation, and snapshots up to <paramref name="maxItems"/> members with
        /// their real EntryIDs, ordered oldest-first.
        /// </summary>
        public IReadOnlyList<ComMailBrief>? TryGetConversationItems(string entryIdHex, string? storeId, int maxItems, out string? error)
        {
            EnsureNotDisposed();
            if (maxItems < 1)
            {
                throw new ArgumentOutOfRangeException(nameof(maxItems));
            }

            string? capturedError = null;
            IReadOnlyList<ComMailBrief>? result = _runner.Run<IReadOnlyList<ComMailBrief>?>(() =>
            {
                dynamic ns = _namespace!;
                object? itemObject = null;
                object? conversation = null;
                object? table = null;
                try
                {
                    itemObject = storeId != null
                        ? ns.GetItemFromID(entryIdHex, storeId)
                        : ns.GetItemFromID(entryIdHex);
                    dynamic item = itemObject!;
                    try
                    {
                        conversation = item.GetConversation();
                    }
                    catch (COMException)
                    {
                        conversation = null;
                    }

                    if (conversation == null)
                    {
                        // No conversation support/membership: the thread is the item itself.
                        ComMailBrief only = SnapshotBrief(ns, item, null, null, includeBody: false);
                        return (IReadOnlyList<ComMailBrief>)new List<ComMailBrief> { only };
                    }

                    table = ((dynamic)conversation!).GetTable();
                    dynamic t = (dynamic)table!;
                    int entryIdIndex = FindTableColumn(t, "EntryID");
                    if (entryIdIndex < 0)
                    {
                        capturedError = "ConversationTableWithoutEntryId";
                        return null;
                    }

                    List<ComMailBrief> briefs = new List<ComMailBrief>();
                    while (!(bool)t.EndOfTable && briefs.Count < maxItems)
                    {
                        object? row = null;
                        object? member = null;
                        try
                        {
                            row = t.GetNextRow();
                            object[] values = (object[])((dynamic)row!).GetValues();
                            if (entryIdIndex >= values.Length || values[entryIdIndex] is not string memberId || memberId.Length == 0)
                            {
                                continue;
                            }

                            try
                            {
                                member = ns.GetItemFromID(memberId);
                            }
                            catch (COMException)
                            {
                                continue;
                            }

                            briefs.Add(SnapshotBrief(ns, member!, null, null, includeBody: false));
                        }
                        finally
                        {
                            Release(member);
                            Release(row);
                        }
                    }

                    briefs.Sort((a, b) => DateTime.Compare(
                        a.ReceivedTime ?? DateTime.MinValue, b.ReceivedTime ?? DateTime.MinValue));
                    return (IReadOnlyList<ComMailBrief>)briefs;
                }
                catch (COMException ex)
                {
                    capturedError = string.Format(CultureInfo.InvariantCulture, "COMException 0x{0:X8}", ex.HResult);
                    return null;
                }
                finally
                {
                    Release(table);
                    Release(conversation);
                    Release(itemObject);
                }
            });

            error = capturedError;
            return result;
        }

        // ------------------------------------------------------------------ show-me (Phase 3, v3.MD L3)

        /// <summary>
        /// open_in_outlook backbone: opens the item by REAL EntryID and calls
        /// MailItem.Display() so it appears in an Inspector window on screen. Works with
        /// or without an Explorer window (a headless COM-started Outlook can still show
        /// Inspectors). Returns the displayed item's snapshot, or null with a
        /// content-free error (S4).
        /// </summary>
        public ComOpenResult? TryDisplayItem(string entryIdHex, string? storeId, out string? error)
        {
            EnsureNotDisposed();
            if (string.IsNullOrWhiteSpace(entryIdHex))
            {
                throw new ArgumentException("EntryID must not be blank.", nameof(entryIdHex));
            }

            string? capturedError = null;
            ComOpenResult? result = _runner.Run<ComOpenResult?>(() =>
            {
                dynamic ns = _namespace!;
                object? itemObject = null;
                try
                {
                    itemObject = storeId != null
                        ? ns.GetItemFromID(entryIdHex, storeId)
                        : ns.GetItemFromID(entryIdHex);
                    ComOpenResult snapshot = Snapshot(itemObject);
                    ((dynamic)itemObject!).Display();
                    return snapshot;
                }
                catch (Exception ex) when (IsComCallFailure(ex))
                {
                    capturedError = DescribeComFailure(ex);
                    return null;
                }
                finally
                {
                    Release(itemObject);
                }
            });

            error = capturedError;
            return result;
        }

        /// <summary>
        /// goto_folder backbone: resolves the store-relative folder path (empty path:
        /// the store's Inbox, falling back to its root folder), makes sure a VISIBLE
        /// Explorer window exists (created + shown when Outlook runs headless, D17/D30),
        /// sets ActiveExplorer().CurrentFolder to it and returns the resulting explorer
        /// state for verification.
        /// </summary>
        public ComExplorerState? TryGotoFolder(string storeDisplayName, IReadOnlyList<string>? folderPath, out string? error)
        {
            EnsureNotDisposed();
            if (string.IsNullOrWhiteSpace(storeDisplayName))
            {
                throw new ArgumentException("Store display name must not be blank.", nameof(storeDisplayName));
            }

            string? capturedError = null;
            ComExplorerState? result = _runner.Run<ComExplorerState?>(() =>
            {
                object? folder = null;
                object? explorer = null;
                try
                {
                    folder = ResolveNavigationFolder(storeDisplayName, folderPath, out string? folderError);
                    if (folder == null)
                    {
                        capturedError = folderError;
                        return null;
                    }

                    explorer = EnsureVisibleExplorer(folder, out string? explorerError);
                    if (explorer == null)
                    {
                        capturedError = explorerError;
                        return null;
                    }

                    ((dynamic)explorer!).CurrentFolder = folder;

                    // A just-created window (headless cold start) can report an empty
                    // CurrentFolder for a beat while it initializes - retry briefly.
                    ComExplorerState state = SnapshotExplorer(explorer!);
                    for (int attempt = 0; attempt < 6 && string.IsNullOrEmpty(state.CurrentFolderPath); attempt++)
                    {
                        Thread.Sleep(250);
                        state = SnapshotExplorer(explorer!);
                    }

                    return state;
                }
                catch (Exception ex) when (IsComCallFailure(ex))
                {
                    capturedError = DescribeComFailure(ex);
                    return null;
                }
                finally
                {
                    Release(explorer);
                    Release(folder);
                }
            });

            error = capturedError;
            return result;
        }

        /// <summary>
        /// show_search_results backbone: optionally navigates to a store/folder first
        /// (so current-folder scopes apply there), then drives Outlook's real search UI
        /// via Explorer.Search(query, olSearchScope). olSearchScope values are
        /// feature-tested live in Phase 3 (v3.MD risk register) - an unsupported value
        /// surfaces as a content-free error here. Returns the explorer state after the
        /// call (the search itself populates asynchronously in the UI).
        /// </summary>
        public ComExplorerState? TryShowSearchResults(
            string query,
            int olSearchScope,
            string? storeDisplayName,
            IReadOnlyList<string>? folderPath,
            out string? error)
        {
            EnsureNotDisposed();
            if (string.IsNullOrWhiteSpace(query))
            {
                throw new ArgumentException("Query must not be blank.", nameof(query));
            }

            string? capturedError = null;
            ComExplorerState? result = _runner.Run<ComExplorerState?>(() =>
            {
                object? folder = null;
                object? explorer = null;
                try
                {
                    if (storeDisplayName != null)
                    {
                        folder = ResolveNavigationFolder(storeDisplayName, folderPath, out string? folderError);
                        if (folder == null)
                        {
                            capturedError = folderError;
                            return null;
                        }
                    }

                    explorer = EnsureVisibleExplorer(folder, out string? explorerError);
                    if (explorer == null)
                    {
                        capturedError = explorerError;
                        return null;
                    }

                    dynamic e = explorer!;
                    if (folder != null)
                    {
                        e.CurrentFolder = folder;
                    }

                    e.Search(query, olSearchScope);
                    return SnapshotExplorer(explorer!);
                }
                catch (Exception ex) when (IsComCallFailure(ex))
                {
                    capturedError = DescribeComFailure(ex);
                    return null;
                }
                finally
                {
                    Release(explorer);
                    Release(folder);
                }
            });

            error = capturedError;
            return result;
        }

        /// <summary>Exits an active Explorer search (test cleanup / leaving the UI tidy).</summary>
        public bool TryClearSearch(out string? error)
        {
            EnsureNotDisposed();
            string? capturedError = null;
            bool result = _runner.Run(() =>
            {
                object? explorer = null;
                try
                {
                    explorer = ((dynamic)_application!).ActiveExplorer();
                    if (explorer == null)
                    {
                        capturedError = "NoActiveExplorer";
                        return false;
                    }

                    ((dynamic)explorer).ClearSearch();
                    return true;
                }
                catch (Exception ex) when (IsComCallFailure(ex))
                {
                    capturedError = DescribeComFailure(ex);
                    return false;
                }
                finally
                {
                    Release(explorer);
                }
            });

            error = capturedError;
            return result;
        }

        /// <summary>Snapshot of the active Explorer window, or null when none exists.</summary>
        public ComExplorerState? TryGetActiveExplorerState(out string? error)
        {
            EnsureNotDisposed();
            string? capturedError = null;
            ComExplorerState? result = _runner.Run<ComExplorerState?>(() =>
            {
                object? explorer = null;
                try
                {
                    explorer = ((dynamic)_application!).ActiveExplorer();
                    if (explorer == null)
                    {
                        capturedError = "NoActiveExplorer";
                        return null;
                    }

                    return SnapshotExplorer(explorer);
                }
                catch (Exception ex) when (IsComCallFailure(ex))
                {
                    capturedError = DescribeComFailure(ex);
                    return null;
                }
                finally
                {
                    Release(explorer);
                }
            });

            error = capturedError;
            return result;
        }

        /// <summary>Lists open Inspector windows with the shown item's EntryID (test verification).</summary>
        public IReadOnlyList<ComInspectorInfo> GetOpenInspectors()
        {
            EnsureNotDisposed();
            return _runner.Run(() =>
            {
                List<ComInspectorInfo> result = new List<ComInspectorInfo>();
                object? inspectors = null;
                try
                {
                    inspectors = ((dynamic)_application!).Inspectors;
                    dynamic collection = (dynamic)inspectors!;
                    int count = collection.Count;
                    for (int i = 1; i <= count; i++)
                    {
                        object? inspector = null;
                        object? item = null;
                        try
                        {
                            inspector = collection[i];
                            item = ((dynamic)inspector!).CurrentItem;
                            dynamic current = (dynamic)item!;
                            string? entryId = TryGetString(() => (string?)current.EntryID);
                            string? subject = TryGetString(() => (string?)current.Subject);
                            int? itemClass = null;
                            try
                            {
                                itemClass = (int)current.Class;
                            }
                            catch (Exception ex) when (IsComCallFailure(ex))
                            {
                            }

                            result.Add(new ComInspectorInfo(entryId, subject, itemClass));
                        }
                        catch (Exception ex) when (IsComCallFailure(ex))
                        {
                            // Inspector without a readable item - skip.
                        }
                        finally
                        {
                            Release(item);
                            Release(inspector);
                        }
                    }
                }
                catch (Exception ex) when (IsComCallFailure(ex))
                {
                    // No Inspectors collection - empty list.
                }
                finally
                {
                    Release(inspectors);
                }

                return (IReadOnlyList<ComInspectorInfo>)result;
            });
        }

        /// <summary>
        /// Closes the Inspector showing the given EntryID (tests close only windows they
        /// opened themselves - closing a window is NOT an Outlook restart, S7).
        /// olDiscard is used so nothing is saved/prompted.
        /// </summary>
        public bool TryCloseInspectorByEntryId(string entryIdHex, out string? error)
        {
            EnsureNotDisposed();
            if (string.IsNullOrWhiteSpace(entryIdHex))
            {
                throw new ArgumentException("EntryID must not be blank.", nameof(entryIdHex));
            }

            string? capturedError = null;
            bool result = _runner.Run(() =>
            {
                object? inspectors = null;
                try
                {
                    inspectors = ((dynamic)_application!).Inspectors;
                    dynamic collection = (dynamic)inspectors!;
                    int count = collection.Count;
                    for (int i = 1; i <= count; i++)
                    {
                        object? inspector = null;
                        object? item = null;
                        try
                        {
                            inspector = collection[i];
                            item = ((dynamic)inspector!).CurrentItem;
                            string? entryId = TryGetString(() => (string?)((dynamic)item!).EntryID);
                            if (entryId != null && string.Equals(entryId, entryIdHex, StringComparison.OrdinalIgnoreCase))
                            {
                                ((dynamic)inspector!).Close(1); // 1 = olDiscard
                                return true;
                            }
                        }
                        catch (Exception ex) when (IsComCallFailure(ex))
                        {
                            // Unreadable inspector - keep looking.
                        }
                        finally
                        {
                            Release(item);
                            Release(inspector);
                        }
                    }

                    capturedError = "InspectorNotFound";
                    return false;
                }
                catch (Exception ex) when (IsComCallFailure(ex))
                {
                    capturedError = DescribeComFailure(ex);
                    return false;
                }
                finally
                {
                    Release(inspectors);
                }
            });

            error = capturedError;
            return result;
        }

        /// <summary>Reads an Explorer pane's visibility (OlPane: 4 = navigation pane, 5 = to-do bar).</summary>
        public bool? TryGetExplorerPaneVisible(int pane, out string? error)
        {
            EnsureNotDisposed();
            string? capturedError = null;
            bool? result = _runner.Run<bool?>(() =>
            {
                object? explorer = null;
                try
                {
                    explorer = ((dynamic)_application!).ActiveExplorer();
                    if (explorer == null)
                    {
                        capturedError = "NoActiveExplorer";
                        return null;
                    }

                    return (bool)((dynamic)explorer).IsPaneVisible(pane);
                }
                catch (Exception ex) when (IsComCallFailure(ex))
                {
                    capturedError = DescribeComFailure(ex);
                    return null;
                }
                finally
                {
                    Release(explorer);
                }
            });

            error = capturedError;
            return result;
        }

        /// <summary>
        /// Shows/hides an Explorer pane (screenshot hygiene: the S5 screenshot hides the
        /// navigation pane so no other store's folder names are captured; restored after).
        /// </summary>
        public bool TrySetExplorerPaneVisible(int pane, bool visible, out string? error)
        {
            EnsureNotDisposed();
            string? capturedError = null;
            bool result = _runner.Run(() =>
            {
                object? explorer = null;
                try
                {
                    explorer = ((dynamic)_application!).ActiveExplorer();
                    if (explorer == null)
                    {
                        capturedError = "NoActiveExplorer";
                        return false;
                    }

                    ((dynamic)explorer).ShowPane(pane, visible);
                    return true;
                }
                catch (Exception ex) when (IsComCallFailure(ex))
                {
                    capturedError = DescribeComFailure(ex);
                    return false;
                }
                finally
                {
                    Release(explorer);
                }
            });

            error = capturedError;
            return result;
        }

        /// <summary>
        /// STA-side: resolves the navigation target for goto/show tools. Empty path =
        /// the store's Inbox when it has one (delegate caches may not), else the store
        /// root. Returns a folder RCW (CALLER must Release) or null + error.
        /// </summary>
        private object? ResolveNavigationFolder(string storeDisplayName, IReadOnlyList<string>? folderPath, out string? error)
        {
            if (folderPath != null && folderPath.Count > 0)
            {
                return WalkToFolder(storeDisplayName, folderPath, out error);
            }

            error = null;
            dynamic? store = FindStoreByDisplayName(storeDisplayName);
            if (store == null)
            {
                error = "StoreNotFound";
                return null;
            }

            try
            {
                try
                {
                    return store.GetDefaultFolder(6); // olFolderInbox
                }
                catch (Exception ex) when (IsComCallFailure(ex))
                {
                    // Store without an Inbox (some delegate caches) - fall back to root.
                }

                try
                {
                    return store.GetRootFolder();
                }
                catch (Exception ex) when (IsComCallFailure(ex))
                {
                    error = "RootFolderUnavailable";
                    return null;
                }
            }
            finally
            {
                Release(store);
            }
        }

        /// <summary>
        /// STA-side: returns the active Explorer, creating and displaying one (on
        /// <paramref name="preferredFolder"/>, else the default Inbox) when Outlook runs
        /// headless. The window is un-minimized and activated so show-me results are
        /// actually on screen. Returns an Explorer RCW (CALLER must Release) or null.
        /// </summary>
        private object? EnsureVisibleExplorer(object? preferredFolder, out string? error)
        {
            error = null;
            dynamic app = _application!;
            object? explorer = null;
            try
            {
                explorer = app.ActiveExplorer();
            }
            catch (Exception ex) when (IsComCallFailure(ex))
            {
            }

            // D49: ActiveExplorer() can hand back a LIFETIME PIN - a deliberately
            // never-displayed Explorer that stops a window-less Outlook from exiting.
            // Repurposing one as a show-me surface would make the pin the user's window,
            // and closing that window would then take the pin with it - exactly the
            // failure the pin exists to prevent. The test is process-wide on purpose: the
            // pin holding Outlook up may belong to ANOTHER server session in this
            // process, so ownership cannot be the criterion - COM identity against the
            // process-wide pin registry is.
            if (explorer != null && ComposeSurface.IsPin(explorer))
            {
                Release(explorer);
                explorer = null;
            }

            if (explorer == null)
            {
                object? defaultFolder = null;
                object? explorers = null;
                try
                {
                    object? folderToShow = preferredFolder;
                    if (folderToShow == null)
                    {
                        defaultFolder = ((dynamic)_namespace!).GetDefaultFolder(6);
                        folderToShow = defaultFolder;
                    }

                    explorers = app.Explorers;
                    explorer = ((dynamic)explorers!).Add(folderToShow, 0); // 0 = olFolderDisplayNormal
                    ((dynamic)explorer!).Display();
                }
                catch (Exception ex) when (IsComCallFailure(ex))
                {
                    error = DescribeComFailure(ex);
                    Release(explorer);
                    return null;
                }
                finally
                {
                    Release(explorers);
                    Release(defaultFolder);
                }
            }

            dynamic e = explorer!;
            try
            {
                if ((int)e.WindowState == 1) // olMinimized
                {
                    e.WindowState = 2; // olNormalWindow
                }
            }
            catch (Exception ex) when (IsComCallFailure(ex))
            {
            }

            try
            {
                e.Activate();
            }
            catch (Exception ex) when (IsComCallFailure(ex))
            {
            }

            return explorer;
        }

        /// <summary>STA-side: snapshots an Explorer's caption/current folder/window state.</summary>
        private static ComExplorerState SnapshotExplorer(object explorerObject)
        {
            dynamic explorer = explorerObject;
            string? caption = TryGetString(() => (string?)explorer.Caption);
            int? windowState = null;
            try
            {
                windowState = (int)explorer.WindowState;
            }
            catch (Exception ex) when (IsComCallFailure(ex))
            {
            }

            string? folderPath = null;
            string? folderName = null;
            object? current = null;
            try
            {
                current = explorer.CurrentFolder;
                if (current != null)
                {
                    folderPath = TryGetString(() => (string?)((dynamic)current).FolderPath);
                    folderName = TryGetString(() => (string?)((dynamic)current).Name);
                }
            }
            catch (Exception ex) when (IsComCallFailure(ex))
            {
            }
            finally
            {
                Release(current);
            }

            return new ComExplorerState(caption, folderPath, folderName, windowState);
        }

        internal static string DescribeComFailure(Exception ex)
        {
            return ex is COMException com
                ? string.Format(CultureInfo.InvariantCulture, "COMException 0x{0:X8}", com.HResult)
                : ex.GetType().Name;
        }

        // ------------------------------------------------------------------ drafts (Phase 4, v3.MD L4/D4)

        /// <summary>
        /// new_draft backbone (v3.MD sections 3/8 L4). Creates the mail DIRECTLY in the
        /// sending account's Drafts folder (Items.Add - a plain CreateItem would save
        /// into the DEFAULT store's Drafts), pins <c>SendUsingAccount</c> from the
        /// Account OBJECT first, then touches <c>GetInspector</c> so Outlook injects
        /// that account's NEW-MAIL signature natively, and writes the agent text ABOVE
        /// that signature region through the SAME held Inspector's WordEditor (batch A -
        /// A1). Saves to Drafts; <paramref name="display"/> additionally opens the draft
        /// in an Inspector for the user (D4 default behavior).
        /// </summary>
        public ComDraftCreateResult? TryCreateNewDraft(
            string accountSmtpAddress,
            IReadOnlyList<string> toRecipients,
            string subject,
            ComDraftBody body,
            bool display,
            ComSignatureOverride? signatureOverride,
            ComDraftOptions? options,
            out string? error)
        {
            EnsureNotDisposed();
            if (string.IsNullOrWhiteSpace(accountSmtpAddress))
            {
                throw new ArgumentException("Account SMTP address must not be blank.", nameof(accountSmtpAddress));
            }

            if (toRecipients == null)
            {
                throw new ArgumentNullException(nameof(toRecipients));
            }

            if (subject == null)
            {
                throw new ArgumentNullException(nameof(subject));
            }

            if (body == null)
            {
                throw new ArgumentNullException(nameof(body));
            }

            string? capturedError = null;
            ComDraftCreateResult? result = _runner.Run<ComDraftCreateResult?>(() =>
            {
                object? account = null;
                object? deliveryStore = null;
                object? draftsFolder = null;
                object? items = null;
                object? mail = null;
                try
                {
                    account = FindAccountBySmtp(accountSmtpAddress);
                    if (account == null)
                    {
                        capturedError = "AccountNotFound";
                        return null;
                    }

                    // Captured NOW: the pinned identity for the outcome snapshot when
                    // the post-save SendUsingAccount readback degrades (see SnapshotDraft).
                    string? pinnedAccountSmtp = TryGetString(() => (string?)((dynamic)account!).SmtpAddress);

                    try
                    {
                        deliveryStore = ((dynamic)account).DeliveryStore;
                    }
                    catch (Exception ex) when (IsComCallFailure(ex))
                    {
                    }

                    if (deliveryStore == null)
                    {
                        capturedError = "AccountHasNoDeliveryStore";
                        return null;
                    }

                    // Deterministic store identity for the outcome snapshot: on a FRESH
                    // (just-started) Outlook the saved item's Parent.Store probe can
                    // transiently fail (live-observed in the soak-fix batch) - the
                    // store we created the draft in is authoritative either way.
                    string? deliveryStoreName = TryGetString(() => (string?)((dynamic)deliveryStore!).DisplayName);
                    string? deliveryStoreId = TryGetString(() => (string?)((dynamic)deliveryStore!).StoreID);

                    draftsFolder = ((dynamic)deliveryStore).GetDefaultFolder(16); // olFolderDrafts

                    // Captured NOW (COM is demonstrably answering) as the deterministic
                    // folder identity for the outcome snapshot - see SnapshotDraft.
                    string? draftsFolderName = TryGetString(() => (string?)((dynamic)draftsFolder!).Name);
                    string? draftsFolderEntryId = TryGetString(() => (string?)((dynamic)draftsFolder!).EntryID);

                    items = ((dynamic)draftsFolder!).Items;
                    mail = ((dynamic)items!).Add(0); // olMailItem
                    dynamic draft = mail!;

                    // Identity FIRST: the Account OBJECT, before anything else touches
                    // the item (v3.MD section 3 - omitting it silently uses the default
                    // account; a string would not bind).
                    SetSendUsingAccount(mail!, account);

                    (bool signatureInjected, long textBefore, long textAfter, bool overrideApplied, string? overrideError, bool wordPlaced, bool surfacePromoted) =
                        ComposeDraft((object)draft, body, signatureOverride);
                    draft.Subject = subject;
                    List<string> unresolved = new List<string>();
                    AddRecipients(draft, toRecipients, 1, unresolved);
                    AddRecipients(draft, options?.CcRecipients ?? Array.Empty<string>(), 2, unresolved);
                    AddRecipients(draft, options?.BccRecipients ?? Array.Empty<string>(), 3, unresolved);
                    ApplyDraftOptions(draft, options);

                    // Attachments AFTER the composition closed the inspector (D46/C3):
                    // adding a file re-renders the item, and the Word edits must already
                    // be committed by then. A per-file COM refusal is reported through the
                    // saved-item snapshot below (requested count vs what is really there)
                    // rather than losing the draft.
                    _ = AddAttachmentsToDraft(draft, options?.AttachmentPaths);
                    draft.Save();

                    // The GetInspector touch leaves a HIDDEN Inspector alive inside
                    // Outlook (it shows up in Application.Inspectors - Phase-4 live
                    // finding). The Word compose path already closed it via
                    // Close(olSave); calling GetInspector again would only materialize a
                    // NEW one, so only the fallback path needs the cleanup. Display()
                    // below opens a fresh visible one for the final item when requested.
                    if (!wordPlaced)
                    {
                        CloseHiddenInspector(mail!);
                    }

                    mail = RelocateToFolderIfNeeded(mail!, draftsFolder!, out bool moved, out string? initialFolder, out bool inDraftsFolder);
                    if (display)
                    {
                        ((dynamic)mail!).Display();
                    }

                    string? folderFallbackName = inDraftsFolder ? draftsFolderName : null;
                    string? folderFallbackId = inDraftsFolder ? draftsFolderEntryId : null;
                    ComDraftInfo info = SnapshotDraft(
                        mail!,
                        deliveryStoreName,
                        deliveryStoreId,
                        folderFallbackName,
                        folderFallbackId,
                        pinnedAccountSmtp);
                    info = ResnapshotIfRecipientsEmpty(
                        info, deliveryStoreName, deliveryStoreId, folderFallbackName, folderFallbackId, pinnedAccountSmtp);
                    return new ComDraftCreateResult(
                        info,
                        accountResolved: true,
                        signatureInjected,
                        textBefore,
                        textAfter,
                        moved,
                        initialFolder,
                        display,
                        signatureOverride?.Name,
                        overrideApplied,
                        overrideError,
                        wordPlaced,
                        unresolved,
                        conversationTopicPreserved: null,
                        attachments: SnapshotAttachments(mail!),
                        composeSurfacePromoted: surfacePromoted,
                        composeSurfaceError: wordPlaced ? null : overrideError ?? "NoWordEditor");
                }
                catch (Exception ex) when (IsComCallFailure(ex))
                {
                    capturedError = DescribeComFailure(ex);
                    return null;
                }
                finally
                {
                    Release(mail);
                    Release(items);
                    Release(draftsFolder);
                    Release(deliveryStore);
                    Release(account);
                }
            });

            error = capturedError;
            return result;
        }

        /// <summary>
        /// reply_draft/replyall_draft/forward_draft backbone: derives the draft ONLY via
        /// COM <c>Reply()</c>/<c>ReplyAll()</c>/<c>Forward()</c> (threading + quoted
        /// history - v3.MD section 12: never rebuild replies with CreateItem), pins
        /// <c>SendUsingAccount</c> from the Account whose delivery store contains the
        /// source mail BEFORE the <c>GetInspector</c> signature touch, prepends the
        /// agent text ABOVE the quoted block, saves into that store's Drafts (moving the
        /// item there when Outlook saved it elsewhere) and optionally displays it (D4).
        /// </summary>
        public ComDraftCreateResult? TryCreateDerivedDraft(
            string sourceEntryIdHex,
            string? sourceStoreId,
            ComDerivedDraftKind kind,
            IReadOnlyList<string> toRecipients,
            ComDraftBody body,
            bool display,
            ComSignatureOverride? signatureOverride,
            ComDraftOptions? options,
            out string? error)
        {
            EnsureNotDisposed();
            if (string.IsNullOrWhiteSpace(sourceEntryIdHex))
            {
                throw new ArgumentException("Source EntryID must not be blank.", nameof(sourceEntryIdHex));
            }

            if (toRecipients == null)
            {
                throw new ArgumentNullException(nameof(toRecipients));
            }

            if (body == null)
            {
                throw new ArgumentNullException(nameof(body));
            }

            string? capturedError = null;
            ComDraftCreateResult? result = _runner.Run<ComDraftCreateResult?>(() =>
            {
                dynamic ns = _namespace!;
                object? source = null;
                object? sourceParent = null;
                object? sourceStore = null;
                object? account = null;
                object? draftsFolder = null;
                object? mail = null;
                try
                {
                    source = sourceStoreId != null
                        ? ns.GetItemFromID(sourceEntryIdHex, sourceStoreId)
                        : ns.GetItemFromID(sourceEntryIdHex);

                    // The store the source mail lives in drives both the sending
                    // account and the Drafts folder the draft must land in.
                    string? sourceStoreIdActual = null;
                    string? sourceStoreName = null;
                    try
                    {
                        sourceParent = ((dynamic)source!).Parent;
                        if (sourceParent != null)
                        {
                            sourceStore = ((dynamic)sourceParent).Store;
                            if (sourceStore != null)
                            {
                                sourceStoreIdActual = TryGetString(() => (string?)((dynamic)sourceStore!).StoreID);
                                sourceStoreName = TryGetString(() => (string?)((dynamic)sourceStore!).DisplayName);
                            }
                        }
                    }
                    catch (Exception ex) when (IsComCallFailure(ex))
                    {
                    }

                    dynamic sourceItem = source!;
                    mail = kind switch
                    {
                        ComDerivedDraftKind.Reply => sourceItem.Reply(),
                        ComDerivedDraftKind.ReplyAll => sourceItem.ReplyAll(),
                        ComDerivedDraftKind.Forward => sourceItem.Forward(),
                        _ => throw new ArgumentOutOfRangeException(nameof(kind)),
                    };
                    dynamic draft = mail!;

                    // Identity: the account delivering into the source store (what the
                    // UI would send from). Delegate-store mail has no matching account -
                    // recorded, SendUsingAccount left for Outlook to resolve.
                    bool accountResolved = false;
                    string? pinnedAccountSmtp = null;
                    if (sourceStoreIdActual != null || sourceStoreName != null)
                    {
                        account = FindAccountByDeliveryStore(sourceStoreIdActual, sourceStoreName);
                        if (account != null)
                        {
                            SetSendUsingAccount(mail!, account);
                            accountResolved = true;

                            // Captured NOW: the pinned identity for the outcome snapshot
                            // when the post-save readback degrades (see SnapshotDraft).
                            pinnedAccountSmtp = TryGetString(() => (string?)((dynamic)account!).SmtpAddress);
                        }
                    }

                    (bool signatureInjected, long textBefore, long textAfter, bool overrideApplied, string? overrideError, bool wordPlaced, bool surfacePromoted) =
                        ComposeDraft((object)draft, body, signatureOverride);
                    List<string> unresolved = new List<string>();
                    if (kind == ComDerivedDraftKind.Forward)
                    {
                        AddRecipients(draft, toRecipients, 1, unresolved);
                    }

                    // A2: Cc/Bcc are APPENDED - reply-all's own recipient list stands.
                    AddRecipients(draft, options?.CcRecipients ?? Array.Empty<string>(), 2, unresolved);
                    AddRecipients(draft, options?.BccRecipients ?? Array.Empty<string>(), 3, unresolved);
                    ApplyDraftOptions(draft, options);

                    // A3: a subject override replaces Outlook's RE:/FW: subject. Outlook
                    // recomputes PR_CONVERSATION_TOPIC from the new subject, which is the
                    // grouping key whenever a conversation cannot be resolved from the
                    // ConversationIndex GUID - so the SOURCE topic is written back after
                    // the subject write and before Save(). ConversationIndex itself is
                    // never touched: Reply()/Forward() already produced the correct child.
                    bool? topicPreserved = null;
                    if (!string.IsNullOrWhiteSpace(options?.SubjectOverride))
                    {
                        string? sourceTopic = TryGetString(() => (string?)sourceItem.ConversationTopic)
                            ?? TryGetPropertyString(sourceItem, ConversationTopicDasl);
                        string? childIndex = TryGetString(() => (string?)draft.ConversationIndex);

                        draft.Subject = options!.SubjectOverride;

                        // Order matters: index first (it carries the GUID the desktop
                        // groups by), topic second (the fallback grouping key).
                        byte[]? indexBytes = HexToBytes(childIndex);
                        bool indexRestored = indexBytes != null
                            && TrySetPropertyBinary(draft, ConversationIndexDasl, indexBytes);
                        bool topicRestored = !string.IsNullOrEmpty(sourceTopic)
                            && TrySetPropertyString(draft, ConversationTopicDasl, sourceTopic!);
                        topicPreserved = indexRestored && topicRestored;
                    }

                    // Attachments AFTER the composition closed the inspector (D46/C3).
                    _ = AddAttachmentsToDraft(draft, options?.AttachmentPaths);
                    draft.Save();

                    // Same hidden-Inspector cleanup as the new-draft path: only needed
                    // when the composition fell back to the HTML path, because the Word
                    // path already closed the held Inspector with Close(olSave).
                    if (!wordPlaced)
                    {
                        CloseHiddenInspector(mail!);
                    }

                    bool moved = false;
                    string? initialFolder = null;
                    bool inDraftsFolder = false;
                    string? draftsFolderName = null;
                    string? draftsFolderEntryId = null;
                    if (sourceStore != null)
                    {
                        try
                        {
                            draftsFolder = ((dynamic)sourceStore).GetDefaultFolder(16); // olFolderDrafts
                        }
                        catch (Exception ex) when (IsComCallFailure(ex))
                        {
                            // Store without a Drafts folder (some delegate caches) - the
                            // draft stays where Outlook saved it.
                        }
                    }

                    if (draftsFolder != null)
                    {
                        // Captured NOW (COM is demonstrably answering) as the
                        // deterministic folder identity for the outcome snapshot.
                        draftsFolderName = TryGetString(() => (string?)((dynamic)draftsFolder).Name);
                        draftsFolderEntryId = TryGetString(() => (string?)((dynamic)draftsFolder).EntryID);
                        mail = RelocateToFolderIfNeeded(mail!, draftsFolder, out moved, out initialFolder, out inDraftsFolder);
                    }

                    if (display)
                    {
                        ((dynamic)mail!).Display();
                    }

                    string? folderFallbackName = inDraftsFolder ? draftsFolderName : null;
                    string? folderFallbackId = inDraftsFolder ? draftsFolderEntryId : null;
                    string? smtpFallback = accountResolved ? pinnedAccountSmtp : null;
                    ComDraftInfo info = SnapshotDraft(
                        mail!,
                        sourceStoreName,
                        sourceStoreIdActual,
                        folderFallbackName,
                        folderFallbackId,
                        smtpFallback);
                    info = ResnapshotIfRecipientsEmpty(
                        info, sourceStoreName, sourceStoreIdActual, folderFallbackName, folderFallbackId, smtpFallback);
                    return new ComDraftCreateResult(
                        info,
                        accountResolved,
                        signatureInjected,
                        textBefore,
                        textAfter,
                        moved,
                        initialFolder,
                        display,
                        signatureOverride?.Name,
                        overrideApplied,
                        overrideError,
                        wordPlaced,
                        unresolved,
                        topicPreserved,
                        SnapshotAttachments(mail!),
                        composeSurfacePromoted: surfacePromoted,
                        composeSurfaceError: wordPlaced ? null : overrideError ?? "NoWordEditor");
                }
                catch (Exception ex) when (IsComCallFailure(ex))
                {
                    capturedError = DescribeComFailure(ex);
                    return null;
                }
                finally
                {
                    Release(mail);
                    Release(draftsFolder);
                    Release(account);
                    Release(sourceStore);
                    Release(sourceParent);
                    Release(source);
                }
            });

            error = capturedError;
            return result;
        }

        /// <summary>
        /// Re-opens a mail by EntryID and snapshots its identity/threading state
        /// (SendUsingAccount, parent folder, ConversationIndex) - the draft tests verify
        /// PERSISTED state through this instead of trusting the creation-time snapshot.
        /// </summary>
        public ComDraftInfo? TryGetMailInfo(string entryIdHex, string? storeId, out string? error)
        {
            EnsureNotDisposed();
            if (string.IsNullOrWhiteSpace(entryIdHex))
            {
                throw new ArgumentException("EntryID must not be blank.", nameof(entryIdHex));
            }

            string? capturedError = null;
            ComDraftInfo? result = _runner.Run<ComDraftInfo?>(() =>
            {
                dynamic ns = _namespace!;
                object? itemObject = null;
                try
                {
                    itemObject = storeId != null
                        ? ns.GetItemFromID(entryIdHex, storeId)
                        : ns.GetItemFromID(entryIdHex);
                    return SnapshotDraft(itemObject!);
                }
                catch (Exception ex) when (IsComCallFailure(ex))
                {
                    capturedError = DescribeComFailure(ex);
                    return null;
                }
                finally
                {
                    Release(itemObject);
                }
            });

            error = capturedError;
            return result;
        }

        /// <summary>
        /// Test-support surface (read-only): the RAW <c>HTMLBody</c> of a mail. The
        /// plain-text extraction the product returns collapses markup, so the
        /// signature-placement contract (agent text above an INTACT html signature) can
        /// only be asserted on the HTML itself.
        /// </summary>
        public string? TryGetHtmlBody(string entryIdHex, string? storeId, out string? error)
        {
            EnsureNotDisposed();
            if (string.IsNullOrWhiteSpace(entryIdHex))
            {
                throw new ArgumentException("EntryID must not be blank.", nameof(entryIdHex));
            }

            string? capturedError = null;
            string? result = _runner.Run<string?>(() =>
            {
                dynamic ns = _namespace!;
                object? itemObject = null;
                try
                {
                    itemObject = storeId != null
                        ? ns.GetItemFromID(entryIdHex, storeId)
                        : ns.GetItemFromID(entryIdHex);
                    return TryGetString(() => (string?)((dynamic)itemObject!).HTMLBody);
                }
                catch (Exception ex) when (IsComCallFailure(ex))
                {
                    capturedError = DescribeComFailure(ex);
                    return null;
                }
                finally
                {
                    Release(itemObject);
                }
            });

            error = capturedError;
            return result;
        }

        /// <summary>
        /// Identity of a store's default folder (6 = Inbox, 16 = Drafts, ...): EntryID +
        /// localized name. The draft tests compare a draft's parent folder EntryID
        /// against this instead of asserting locale-dependent folder names.
        /// </summary>
        public ComDefaultFolderInfo? TryGetDefaultFolderInfo(string storeDisplayName, int olDefaultFolderId, out string? error)
        {
            EnsureNotDisposed();
            if (string.IsNullOrWhiteSpace(storeDisplayName))
            {
                throw new ArgumentException("Store display name must not be blank.", nameof(storeDisplayName));
            }

            string? capturedError = null;
            ComDefaultFolderInfo? result = _runner.Run<ComDefaultFolderInfo?>(() =>
            {
                dynamic? store = FindStoreByDisplayName(storeDisplayName);
                if (store == null)
                {
                    capturedError = "StoreNotFound";
                    return null;
                }

                object? folder = null;
                try
                {
                    folder = store.GetDefaultFolder(olDefaultFolderId);
                    dynamic f = folder!;
                    return new ComDefaultFolderInfo((string)f.EntryID, (string)f.Name);
                }
                catch (Exception ex) when (IsComCallFailure(ex))
                {
                    capturedError = DescribeComFailure(ex);
                    return null;
                }
                finally
                {
                    Release(folder);
                    Release(store);
                }
            });

            error = capturedError;
            return result;
        }

        // ------------------------------------------------------------ move + archive (D39)

        /// <summary>
        /// Store-relative folder path from an OOM <c>Folder.FolderPath</c>
        /// (<c>\\Store Display Name\A\B</c> becomes <c>A/B</c>, the list_folders
        /// convention; the store root itself becomes an empty string). Pure logic,
        /// public for T1.
        /// </summary>
        public static string ToStoreRelativeFolderPath(string? folderPath, string? storeDisplayName)
        {
            if (string.IsNullOrEmpty(folderPath))
            {
                return string.Empty;
            }

            string path = folderPath!;
            if (path.StartsWith("\\\\", StringComparison.Ordinal))
            {
                path = path.Substring(2);
                bool stripped = false;
                if (!string.IsNullOrEmpty(storeDisplayName)
                    && path.StartsWith(storeDisplayName!, StringComparison.OrdinalIgnoreCase))
                {
                    string remainder = path.Substring(storeDisplayName!.Length);

                    // Exact-segment match only: the store name must be the WHOLE first
                    // segment (guards against one store name being a prefix of another).
                    if (remainder.Length == 0 || remainder[0] == '\\')
                    {
                        path = remainder;
                        stripped = true;
                    }
                }

                if (!stripped)
                {
                    // Unknown prefix: drop the first segment (the store) regardless.
                    int firstSeparator = path.IndexOf('\\');
                    path = firstSeparator < 0 ? string.Empty : path.Substring(firstSeparator);
                }
            }

            return path.TrimStart('\\').Replace('\\', '/');
        }

        /// <summary>
        /// Resolves a store's DESIGNATED Archive folder - the folder Outlook's own
        /// Archive action (Backspace), mobile swipe-archive and OWA use. Resolution is
        /// localization-proof and never guesses by name: primary =
        /// <c>Store.GetDefaultFolder(39)</c> (undocumented but live-proven value, see
        /// <see cref="ArchiveFolderResolution"/>), fallback = PR_IPM_ARCHIVE_ENTRYID on
        /// the store object. The resolved folder is VERIFIED (same store, mail folder,
        /// not one of the core default folders) before it is trusted - paranoia against
        /// the undocumented enum meaning something else on another build. Read-only;
        /// when a store has no designated archive folder the resolution FAILS
        /// (content-free error) and nothing is created.
        /// </summary>
        public ComArchiveFolderInfo? TryResolveArchiveFolder(string storeDisplayName, out string? error)
        {
            EnsureNotDisposed();
            if (string.IsNullOrWhiteSpace(storeDisplayName))
            {
                throw new ArgumentException("Store display name must not be blank.", nameof(storeDisplayName));
            }

            string? capturedError = null;
            ComArchiveFolderInfo? result = _runner.Run<ComArchiveFolderInfo?>(() =>
            {
                dynamic ns = _namespace!;
                dynamic? store = FindStoreByDisplayName(storeDisplayName);
                if (store == null)
                {
                    capturedError = "StoreNotFound";
                    return null;
                }

                object? folder = null;
                try
                {
                    string storeId = (string)store.StoreID;
                    string via = "outlookDefaultFolder";
                    try
                    {
                        folder = store.GetDefaultFolder(ArchiveFolderResolution.OlFolderArchive);
                    }
                    catch (Exception ex) when (IsComCallFailure(ex))
                    {
                        folder = null;
                    }

                    if (folder == null)
                    {
                        via = "storeArchiveProperty";
                        object? accessor = null;
                        try
                        {
                            accessor = store.PropertyAccessor;
                            object? value = ((dynamic)accessor!).GetProperty(ArchiveFolderResolution.ArchiveEntryIdPropertySchema);
                            string? hex = ArchiveFolderResolution.TryReadEntryIdHex(value);
                            if (hex != null)
                            {
                                folder = ns.GetFolderFromID(hex, storeId);
                            }
                        }
                        catch (Exception ex) when (IsComCallFailure(ex))
                        {
                            folder = null;
                        }
                        finally
                        {
                            Release(accessor);
                        }
                    }

                    if (folder == null)
                    {
                        capturedError = "NoDesignatedArchiveFolder";
                        return null;
                    }

                    dynamic f = folder;
                    string entryId = (string)f.EntryID;
                    string? verification = VerifyArchiveCandidate(f, store, storeId, entryId);
                    if (verification != null)
                    {
                        capturedError = verification;
                        return null;
                    }

                    return new ComArchiveFolderInfo(
                        storeDisplayName,
                        storeId,
                        entryId,
                        (string)f.Name,
                        ToStoreRelativeFolderPath((string?)f.FolderPath, storeDisplayName),
                        via);
                }
                catch (Exception ex) when (IsComCallFailure(ex))
                {
                    capturedError = DescribeComFailure(ex);
                    return null;
                }
                finally
                {
                    Release(folder);
                    Release(store);
                }
            });

            error = capturedError;
            return result;
        }

        /// <summary>
        /// STA-side verification of a resolved archive-folder candidate: it must live
        /// in the SAME store, be a mail folder, and not be one of the core default
        /// folders (Deleted Items/Outbox/Sent/Inbox/Drafts/Junk) - mis-designating any
        /// of those as "archive" would make archive_mail silently do something else.
        /// Returns a content-free error or null when the candidate is sound.
        /// </summary>
        private static string? VerifyArchiveCandidate(dynamic candidate, dynamic store, string storeId, string candidateEntryId)
        {
            object? candidateStore = null;
            try
            {
                candidateStore = candidate.Store;
                string? candidateStoreId = candidateStore != null ? (string?)((dynamic)candidateStore!).StoreID : null;
                if (!string.Equals(candidateStoreId, storeId, StringComparison.OrdinalIgnoreCase))
                {
                    return "ArchiveFolderVerificationFailed:store";
                }

                if ((int)candidate.DefaultItemType != 0)
                {
                    return "ArchiveFolderVerificationFailed:itemType";
                }

                // 3=Deleted Items 4=Outbox 5=Sent 6=Inbox 16=Drafts 23=Junk
                foreach (int coreDefault in new[] { 3, 4, 5, 6, 16, 23 })
                {
                    object? defaultFolder = null;
                    try
                    {
                        defaultFolder = store.GetDefaultFolder(coreDefault);
                        if (string.Equals((string)((dynamic)defaultFolder!).EntryID, candidateEntryId, StringComparison.OrdinalIgnoreCase))
                        {
                            return "ArchiveFolderVerificationFailed:coreDefault";
                        }
                    }
                    catch (Exception ex) when (IsComCallFailure(ex))
                    {
                    }
                    finally
                    {
                        Release(defaultFolder);
                    }
                }

                return null;
            }
            catch (Exception ex) when (IsComCallFailure(ex))
            {
                return "ArchiveFolderVerificationFailed:probe";
            }
            finally
            {
                Release(candidateStore);
            }
        }

        /// <summary>
        /// Moves one mail item to a folder path WITHIN ITS OWN STORE (D39 v1:
        /// same-store only - the target is resolved inside the store the item already
        /// lives in, so a cross-store move cannot happen by construction; when
        /// <paramref name="requireStoreDisplayName"/> is given and the item lives
        /// elsewhere the move is refused with <c>CrossStoreTarget:&lt;store&gt;</c>).
        /// Missing target segments are created only when <paramref name="createMissing"/>
        /// (mail folders, parents included). Refused targets: Deleted Items and its
        /// subtree (deletion semantics - the server has no delete surface), the Outbox,
        /// non-mail folders, and the item's current folder. The result carries old/new
        /// EntryIDs (EntryIDs CHANGE on any move) and the source path as undo address.
        /// </summary>
        public ComMoveItemResult? TryMoveItemToPath(
            string entryIdHex,
            string? storeId,
            IReadOnlyList<string> targetSegments,
            bool createMissing,
            string? requireStoreDisplayName,
            out string? error)
        {
            EnsureNotDisposed();
            if (string.IsNullOrWhiteSpace(entryIdHex))
            {
                throw new ArgumentException("EntryID must not be blank.", nameof(entryIdHex));
            }

            if (targetSegments == null || targetSegments.Count == 0)
            {
                throw new ArgumentException("Target folder path must have at least one segment.", nameof(targetSegments));
            }

            string? capturedError = null;
            ComMoveItemResult? result = _runner.Run<ComMoveItemResult?>(() =>
            {
                dynamic ns = _namespace!;
                object? item = null;
                object? parent = null;
                object? itemStore = null;
                object? targetFolder = null;
                try
                {
                    try
                    {
                        item = storeId != null
                            ? ns.GetItemFromID(entryIdHex, storeId)
                            : ns.GetItemFromID(entryIdHex);
                    }
                    catch (Exception ex) when (IsComCallFailure(ex))
                    {
                        capturedError = "ItemNotFound";
                        return null;
                    }

                    if (!IsMailItem(item!))
                    {
                        capturedError = "NotAMailItem";
                        return null;
                    }

                    dynamic mail = item!;
                    parent = mail.Parent;
                    dynamic parentFolder = parent!;
                    string fromFolderPath = (string)parentFolder.FolderPath;
                    string parentEntryId = (string)parentFolder.EntryID;
                    itemStore = parentFolder.Store;
                    dynamic ownStore = itemStore!;
                    string ownStoreName = (string)ownStore.DisplayName;

                    if (requireStoreDisplayName != null
                        && !string.Equals(ownStoreName, requireStoreDisplayName, StringComparison.OrdinalIgnoreCase))
                    {
                        capturedError = "CrossStoreTarget:" + ownStoreName;
                        return null;
                    }

                    List<string> createdPaths = new List<string>();
                    targetFolder = ResolveOrCreateFolder(ownStore, targetSegments, createMissing, createdPaths, out string? resolveError);
                    if (targetFolder == null)
                    {
                        capturedError = resolveError;
                        return null;
                    }

                    string? guardError = VerifyMoveTarget(ownStore, targetFolder, parentEntryId);
                    if (guardError != null)
                    {
                        capturedError = guardError;
                        return null;
                    }

                    return ExecuteMove(mail, targetFolder!, entryIdHex, ownStoreName, fromFolderPath, createdPaths);
                }
                catch (Exception ex) when (IsComCallFailure(ex))
                {
                    capturedError = DescribeComFailure(ex);
                    return null;
                }
                finally
                {
                    Release(targetFolder);
                    Release(itemStore);
                    Release(parent);
                    Release(item);
                }
            });

            error = capturedError;
            return result;
        }

        /// <summary>
        /// Moves one mail item to an already-resolved folder (archive_mail: the target
        /// is the store's designated Archive folder from
        /// <see cref="TryResolveArchiveFolder"/>). Same-store is enforced (the item's
        /// own store must match <paramref name="targetStoreId"/>); an item already in
        /// the target folder is refused with <c>AlreadyInTargetFolder</c>. Result
        /// semantics identical to <see cref="TryMoveItemToPath"/>.
        /// </summary>
        public ComMoveItemResult? TryMoveItemToFolderId(
            string entryIdHex,
            string? storeId,
            string targetFolderEntryId,
            string targetStoreId,
            out string? error)
        {
            EnsureNotDisposed();
            if (string.IsNullOrWhiteSpace(entryIdHex))
            {
                throw new ArgumentException("EntryID must not be blank.", nameof(entryIdHex));
            }

            if (string.IsNullOrWhiteSpace(targetFolderEntryId) || string.IsNullOrWhiteSpace(targetStoreId))
            {
                throw new ArgumentException("Target folder identity must not be blank.", nameof(targetFolderEntryId));
            }

            string? capturedError = null;
            ComMoveItemResult? result = _runner.Run<ComMoveItemResult?>(() =>
            {
                dynamic ns = _namespace!;
                object? item = null;
                object? parent = null;
                object? itemStore = null;
                object? targetFolder = null;
                try
                {
                    try
                    {
                        item = storeId != null
                            ? ns.GetItemFromID(entryIdHex, storeId)
                            : ns.GetItemFromID(entryIdHex);
                    }
                    catch (Exception ex) when (IsComCallFailure(ex))
                    {
                        capturedError = "ItemNotFound";
                        return null;
                    }

                    if (!IsMailItem(item!))
                    {
                        capturedError = "NotAMailItem";
                        return null;
                    }

                    dynamic mail = item!;
                    parent = mail.Parent;
                    dynamic parentFolder = parent!;
                    string fromFolderPath = (string)parentFolder.FolderPath;
                    string parentEntryId = (string)parentFolder.EntryID;
                    itemStore = parentFolder.Store;
                    dynamic ownStore = itemStore!;
                    string ownStoreName = (string)ownStore.DisplayName;
                    string ownStoreId = (string)ownStore.StoreID;

                    if (!string.Equals(ownStoreId, targetStoreId, StringComparison.OrdinalIgnoreCase))
                    {
                        capturedError = "CrossStoreTarget:" + ownStoreName;
                        return null;
                    }

                    if (string.Equals(parentEntryId, targetFolderEntryId, StringComparison.OrdinalIgnoreCase))
                    {
                        capturedError = "AlreadyInTargetFolder";
                        return null;
                    }

                    try
                    {
                        targetFolder = ns.GetFolderFromID(targetFolderEntryId, targetStoreId);
                    }
                    catch (Exception ex) when (IsComCallFailure(ex))
                    {
                        capturedError = "TargetFolderNotFound";
                        return null;
                    }

                    return ExecuteMove(mail, targetFolder!, entryIdHex, ownStoreName, fromFolderPath, Array.Empty<string>());
                }
                catch (Exception ex) when (IsComCallFailure(ex))
                {
                    capturedError = DescribeComFailure(ex);
                    return null;
                }
                finally
                {
                    Release(targetFolder);
                    Release(itemStore);
                    Release(parent);
                    Release(item);
                }
            });

            error = capturedError;
            return result;
        }

        /// <summary>
        /// STA-side: walks <paramref name="segments"/> from the store root, creating
        /// missing MAIL folders (Folders.Add type 6) when allowed. Returns the target
        /// folder RCW (caller releases) or null with a content-free error; created
        /// store-relative paths are appended to <paramref name="createdPaths"/>.
        /// </summary>
        private static object? ResolveOrCreateFolder(
            dynamic store,
            IReadOnlyList<string> segments,
            bool createMissing,
            List<string> createdPaths,
            out string? error)
        {
            error = null;
            object? current;
            try
            {
                current = store.GetRootFolder();
            }
            catch (Exception ex) when (IsComCallFailure(ex))
            {
                error = "RootFolderUnavailable";
                return null;
            }

            string pathSoFar = string.Empty;
            foreach (string segment in segments)
            {
                pathSoFar = pathSoFar.Length == 0 ? segment : pathSoFar + "/" + segment;
                object? next = null;
                object? folders = null;
                try
                {
                    folders = ((dynamic)current!).Folders;
                    try
                    {
                        next = ((dynamic)folders!)[segment];
                    }
                    catch (Exception ex) when (IsComCallFailure(ex))
                    {
                        next = null;
                    }

                    if (next == null)
                    {
                        if (!createMissing)
                        {
                            error = "TargetFolderNotFound";
                            Release(current);
                            return null;
                        }

                        try
                        {
                            // 6 = olFolderInbox: forces an IPF.Note (mail) folder
                            // regardless of the parent's type.
                            next = ((dynamic)folders!).Add(segment, 6);
                            createdPaths.Add(pathSoFar);
                        }
                        catch (Exception ex) when (IsComCallFailure(ex))
                        {
                            error = "TargetFolderCreateFailed";
                            Release(current);
                            return null;
                        }
                    }
                }
                finally
                {
                    Release(folders);
                }

                Release(current);
                current = next;
            }

            return current;
        }

        /// <summary>
        /// STA-side target guards shared by the move ops: the target must be a mail
        /// folder, must not be (or live under) Deleted Items - moving there is deletion
        /// semantics and the server has no delete surface (S1 v2) - must not be the
        /// Outbox, and must differ from the item's current folder. Content-free error
        /// or null.
        /// </summary>
        private static string? VerifyMoveTarget(dynamic store, object targetFolderObject, string sourceParentEntryId)
        {
            dynamic target = targetFolderObject;
            string targetEntryId = (string)target.EntryID;
            if (string.Equals(targetEntryId, sourceParentEntryId, StringComparison.OrdinalIgnoreCase))
            {
                return "AlreadyInTargetFolder";
            }

            try
            {
                if ((int)target.DefaultItemType != 0)
                {
                    return "TargetNotAMailFolder";
                }
            }
            catch (Exception ex) when (IsComCallFailure(ex))
            {
                return "TargetNotAMailFolder";
            }

            string? deletedItemsEntryId = TryGetDefaultFolderEntryId(store, 3);
            string? outboxEntryId = TryGetDefaultFolderEntryId(store, 4);
            if (outboxEntryId != null && string.Equals(targetEntryId, outboxEntryId, StringComparison.OrdinalIgnoreCase))
            {
                return "TargetIsOutbox";
            }

            if (deletedItemsEntryId != null)
            {
                // The target and every ancestor: a subfolder of Deleted Items is still
                // the trash subtree.
                object? cursor = null;
                try
                {
                    string cursorEntryId = targetEntryId;
                    dynamic current = target;
                    for (int depth = 0; depth < FolderWalkDepthGuard; depth++)
                    {
                        if (string.Equals(cursorEntryId, deletedItemsEntryId, StringComparison.OrdinalIgnoreCase))
                        {
                            return "TargetIsDeletedItems";
                        }

                        object? up;
                        try
                        {
                            up = current.Parent;
                        }
                        catch (Exception ex) when (IsComCallFailure(ex))
                        {
                            break;
                        }

                        Release(cursor);
                        cursor = up;
                        if (cursor == null)
                        {
                            break;
                        }

                        current = cursor;
                        try
                        {
                            cursorEntryId = (string)current.EntryID;
                        }
                        catch (Exception ex) when (IsComCallFailure(ex))
                        {
                            break; // reached the namespace/store level
                        }
                    }
                }
                finally
                {
                    Release(cursor);
                }
            }

            return null;
        }

        private static string? TryGetDefaultFolderEntryId(dynamic store, int olDefaultFolderId)
        {
            object? folder = null;
            try
            {
                folder = store.GetDefaultFolder(olDefaultFolderId);
                return (string)((dynamic)folder!).EntryID;
            }
            catch (Exception ex) when (IsComCallFailure(ex))
            {
                return null;
            }
            finally
            {
                Release(folder);
            }
        }

        /// <summary>
        /// STA-side: performs the actual <c>MailItem.Move</c> and snapshots the result.
        /// The returned moved item carries the NEW EntryID (EntryIDs change on any
        /// move); the original RCW is stale afterwards and released by the caller.
        /// </summary>
        private static ComMoveItemResult ExecuteMove(
            dynamic mail,
            object targetFolderObject,
            string oldEntryId,
            string storeDisplayName,
            string fromFolderPath,
            IReadOnlyList<string> createdPaths)
        {
            dynamic target = targetFolderObject;
            string toFolderPath = ToStoreRelativeFolderPath((string?)target.FolderPath, storeDisplayName);
            object? moved = null;
            try
            {
                moved = mail.Move(target);
                string newEntryId = (string)((dynamic)moved!).EntryID;
                return new ComMoveItemResult(
                    oldEntryId,
                    newEntryId,
                    storeDisplayName,
                    ToStoreRelativeFolderPath(fromFolderPath, storeDisplayName),
                    toFolderPath,
                    createdPaths);
            }
            finally
            {
                Release(moved);
            }
        }

        /// <summary>
        /// Sendable-state snapshot for the high-friction send flow (Phase 5, v3.MD D4):
        /// opens the item, requires it to be a mail item, and captures subject, Sent
        /// flag, recipients, plain-text body (content-hash input) and the account whose
        /// delivery store contains it. Read-only - nothing is modified.
        /// </summary>
        public ComSendableDraftState? TryGetSendableDraftState(string entryIdHex, string? storeId, out string? error)
        {
            EnsureNotDisposed();
            if (string.IsNullOrWhiteSpace(entryIdHex))
            {
                throw new ArgumentException("EntryID must not be blank.", nameof(entryIdHex));
            }

            string? capturedError = null;
            ComSendableDraftState? result = _runner.Run<ComSendableDraftState?>(() =>
            {
                dynamic ns = _namespace!;
                object? item = null;
                object? account = null;
                try
                {
                    item = storeId != null
                        ? ns.GetItemFromID(entryIdHex, storeId)
                        : ns.GetItemFromID(entryIdHex);

                    if (!IsMailItem(item!))
                    {
                        capturedError = "NotAMailItem";
                        return null;
                    }

                    ComDraftInfo info = SnapshotDraft(item!);
                    bool isSent = true; // fail CLOSED: unknown state is treated as not sendable
                    try
                    {
                        isSent = (bool)((dynamic)item!).Sent;
                    }
                    catch (Exception ex) when (IsComCallFailure(ex))
                    {
                    }

                    string? body = TryGetString(() => (string?)((dynamic)item!).Body);

                    string? accountSmtp = null;
                    if (info.StoreId != null || info.StoreDisplayName != null)
                    {
                        account = FindAccountByDeliveryStore(info.StoreId, info.StoreDisplayName);
                        if (account != null)
                        {
                            accountSmtp = TryGetString(() => (string?)((dynamic)account!).SmtpAddress);
                        }
                    }

                    return new ComSendableDraftState(
                        info.EntryId,
                        info.StoreId,
                        info.StoreDisplayName,
                        info.ParentFolderName,
                        info.Subject,
                        isSent,
                        body,
                        accountSmtp,
                        info.Recipients,
                        SnapshotAttachments(item!),
                        SendContentHash.DigestHtml(TryGetString(() => (string?)((dynamic)item!).HTMLBody)));
                }
                catch (Exception ex) when (IsComCallFailure(ex))
                {
                    capturedError = DescribeComFailure(ex);
                    return null;
                }
                finally
                {
                    Release(account);
                    Release(item);
                }
            });

            error = capturedError;
            return result;
        }

        /// <summary>
        /// Executes the CONFIRMED send of a saved draft (Phase 5, v3.MD D4/L5) as ONE
        /// STA operation so nothing can change between the checks and <c>Send()</c>:
        /// re-opens the draft, re-verifies it is unsent mail, RECOMPUTES the content
        /// hash against <paramref name="expectedContentHash"/> (token binding covers the
        /// validate-to-send gap), resolves the sending account FROM THE DRAFT'S OWN
        /// STORE (never from caller input - the send can never touch another account),
        /// pins it via the PROPERTYPUTREF accessor and HARD-VERIFIES the getter readback
        /// in-session (Phase-4 footgun - abort on mismatch), applies the optional
        /// SentOnBehalfOfName, and only then calls <c>Send()</c>. Every refusal happens
        /// BEFORE transport; a failure of the Send call itself is reported as
        /// "SendCallFailed:..." (state then unknown - the mail may sit in the Outbox).
        /// </summary>
        public ComSendResult? TrySendDraft(
            string entryIdHex,
            string? storeId,
            string expectedContentHash,
            string? sentOnBehalfOfName,
            out string? error)
        {
            EnsureNotDisposed();
            if (string.IsNullOrWhiteSpace(entryIdHex))
            {
                throw new ArgumentException("EntryID must not be blank.", nameof(entryIdHex));
            }

            if (string.IsNullOrWhiteSpace(expectedContentHash))
            {
                throw new ArgumentException("Expected content hash must not be blank.", nameof(expectedContentHash));
            }

            string? capturedError = null;
            ComSendResult? result = _runner.Run<ComSendResult?>(() =>
            {
                dynamic ns = _namespace!;
                object? item = null;
                object? account = null;
                try
                {
                    item = storeId != null
                        ? ns.GetItemFromID(entryIdHex, storeId)
                        : ns.GetItemFromID(entryIdHex);

                    if (!IsMailItem(item!))
                    {
                        capturedError = "NotAMailItem";
                        return null;
                    }

                    bool isSent = true;
                    try
                    {
                        isSent = (bool)((dynamic)item!).Sent;
                    }
                    catch (Exception ex) when (IsComCallFailure(ex))
                    {
                    }

                    if (isSent)
                    {
                        capturedError = "AlreadySent";
                        return null;
                    }

                    ComDraftInfo info = SnapshotDraft(item!);
                    string? body = TryGetString(() => (string?)((dynamic)item!).Body);

                    // D46: the attachment set and an HTML-body digest are hash inputs too,
                    // so a file added/removed - or a markup-only edit the plain text cannot
                    // show - after the token was issued invalidates it right here, inside
                    // the STA, immediately before Send().
                    string currentHash = SendContentHash.Compute(
                        info.Subject,
                        info.Recipients,
                        body,
                        sentOnBehalfOfName,
                        SnapshotAttachments(item!),
                        SendContentHash.DigestHtml(TryGetString(() => (string?)((dynamic)item!).HTMLBody)));
                    if (!string.Equals(currentHash, expectedContentHash, StringComparison.Ordinal))
                    {
                        capturedError = "ContentChangedSinceToken";
                        return null;
                    }

                    // Identity comes EXCLUSIVELY from the draft's own store - the
                    // account delivering into it (what the UI would send from).
                    account = FindAccountByDeliveryStore(info.StoreId, info.StoreDisplayName);
                    if (account == null)
                    {
                        capturedError = "NoSendingAccountForStore";
                        return null;
                    }

                    string? accountSmtp = TryGetString(() => (string?)((dynamic)account!).SmtpAddress);
                    if (string.IsNullOrWhiteSpace(accountSmtp))
                    {
                        capturedError = "NoSendingAccountForStore";
                        return null;
                    }

                    // ⚠ Phase-4 footgun: SendUsingAccount is PROPERTYPUTREF - use the
                    // explicit putref accessor, then HARD-VERIFY by in-session getter
                    // readback. A mismatch means the DEFAULT account would send: abort.
                    SetSendUsingAccount(item!, account);
                    string? readback = null;
                    object? pinned = null;
                    try
                    {
                        pinned = ((dynamic)item!).SendUsingAccount;
                        if (pinned != null)
                        {
                            readback = TryGetString(() => (string?)((dynamic)pinned!).SmtpAddress);
                        }
                    }
                    finally
                    {
                        Release(pinned);
                    }

                    if (!string.Equals(readback, accountSmtp, StringComparison.OrdinalIgnoreCase))
                    {
                        capturedError = "SendIdentityVerificationFailed";
                        return null;
                    }

                    if (!string.IsNullOrWhiteSpace(sentOnBehalfOfName))
                    {
                        ((dynamic)item!).SentOnBehalfOfName = sentOnBehalfOfName;
                    }

                    // Capture the outcome BEFORE Send() - the EntryID dies with it.
                    ComSendResult sendResult = new ComSendResult(
                        info.EntryId,
                        info.StoreDisplayName,
                        accountSmtp!,
                        string.IsNullOrWhiteSpace(sentOnBehalfOfName) ? null : sentOnBehalfOfName,
                        info.Subject,
                        info.Recipients);

                    try
                    {
                        ((dynamic)item!).Send();
                    }
                    catch (Exception ex) when (IsComCallFailure(ex))
                    {
                        capturedError = "SendCallFailed:" + DescribeComFailure(ex);
                        return null;
                    }

                    return sendResult;
                }
                catch (Exception ex) when (IsComCallFailure(ex))
                {
                    capturedError = DescribeComFailure(ex);
                    return null;
                }
                finally
                {
                    Release(account);
                    Release(item);
                }
            });

            error = capturedError;
            return result;
        }

        // ------------------------------------------------------------------ update / discard drafts (D46, soak fix 19)

        /// <summary>
        /// update_draft backbone (v3.MD D46/C1). Revises an EXISTING unsent draft in
        /// place: the draft region is REPLACED through the batch-A/B one-held-Inspector
        /// model (so the signature and any quoted original survive byte-identically),
        /// recipients follow REPLACE semantics per class, and a subject change carries the
        /// A3 conversation-index/topic restore so the draft stays in its thread.
        /// <para>
        /// DELIBERATELY NO HTMLBody FALLBACK, unlike the creators: the fallback SPLICES
        /// content in at the top of the body, which on an update would APPEND the new text
        /// above the old instead of replacing it. A failed Word step therefore discards the
        /// inspector (<c>Close(olDiscard)</c>) and refuses the whole update, leaving the
        /// draft exactly as it was - a refusal the agent can retry beats a silently
        /// duplicated body.
        /// </para>
        /// Preconditions are all fail-closed: mail item, UNSENT, and living in a Drafts
        /// folder.
        /// </summary>
        public ComDraftUpdateResult? TryUpdateDraft(
            string entryIdHex,
            string? storeId,
            ComDraftBody? body,
            string? subject,
            IReadOnlyList<string>? toRecipients,
            IReadOnlyList<string>? ccRecipients,
            IReadOnlyList<string>? bccRecipients,
            int? importance,
            bool? requestReadReceipt,
            ComSignatureOverride? signatureOverride,
            IReadOnlyList<string> attachmentsToAdd,
            IReadOnlyList<string> attachmentsToRemove,
            bool display,
            out string? error)
        {
            EnsureNotDisposed();
            if (string.IsNullOrWhiteSpace(entryIdHex))
            {
                throw new ArgumentException("EntryID must not be blank.", nameof(entryIdHex));
            }

            string? capturedError = null;
            ComDraftUpdateResult? result = _runner.Run<ComDraftUpdateResult?>(() =>
            {
                dynamic ns = _namespace!;
                object? item = null;
                object? fresh = null;
                try
                {
                    item = storeId != null
                        ? ns.GetItemFromID(entryIdHex, storeId)
                        : ns.GetItemFromID(entryIdHex);

                    capturedError = CheckEditableDraft(item!);
                    if (capturedError != null)
                    {
                        return null;
                    }

                    List<string> changed = new List<string>();
                    List<string> unresolved = new List<string>();
                    bool bodyReplaced = false;
                    bool wordPlaced = false;
                    bool overrideApplied = false;

                    // ⚠ THE ORDERING HERE COST TWO LIVE RUNS TO GET RIGHT, and neither of
                    // the two obvious shapes works on an ALREADY-SAVED item:
                    //
                    //  (a) compose in Word, then set properties on THIS reference and
                    //      Save() - the creators' order. On a NEW item that is correct,
                    //      because GetInspector binds to the very object we hold. On a
                    //      SAVED item Outlook's inspector edits its OWN MailItem instance;
                    //      Close(olSave) commits through that one, and our Save() then
                    //      writes the pre-edit content straight back over it. Observed
                    //      exactly: property changes stuck, every body rewrite vanished
                    //      while the call reported success.
                    //  (b) set properties and Save() FIRST, then compose. That loses the
                    //      body a different way: right after a Save() the freshly acquired
                    //      inspector has no Word editor yet, so the compose refuses with
                    //      NoWordEditor. (Re-acquiring the inspector to retry is worse
                    //      still - every GetInspector materializes ANOTHER hidden inspector
                    //      (section 12) and asking for a second one wedged Outlook solid.)
                    //
                    // What works is neither: compose FIRST on a freshly opened item (the
                    // proven shape - the inspector is the first thing touched after the
                    // open), let Close(olSave) commit it, then RE-OPEN the item and apply
                    // every property change to that fresh instance, which already contains
                    // the Word edits and can therefore be Save()d without clobbering them.
                    // D47: an inline image the draft ALREADY carries can only survive the
                    // re-render if it is embedded. One that is still a file:/// link (any
                    // draft composed before D47) is re-serialized by Word as a placeholder
                    // shape and vanishes. Counted on both sides of the compose so the
                    // outcome can REPORT the loss instead of letting the agent believe the
                    // signature came through whole. Only when Word actually runs.
                    int imagesBefore = 0;
                    bool countImages = body != null || signatureOverride != null;
                    if (countImages)
                    {
                        imagesBefore = OutlookAI.Core.Text.HtmlBodyComposer.CountInlineImages(
                            TryGetString(() => (string?)((dynamic)item!).HTMLBody));
                    }

                    if (body != null || signatureOverride != null)
                    {
                        (bool ok, string? composeError) = ReviseHeldDocument(item!, body, signatureOverride);
                        if (!ok)
                        {
                            capturedError = composeError ?? "BodyReplaceFailed";
                            return null;
                        }

                        wordPlaced = true;
                        bodyReplaced = body != null;
                        overrideApplied = signatureOverride != null;
                        if (body != null)
                        {
                            changed.Add("body");
                        }

                        if (signatureOverride != null)
                        {
                            changed.Add("signature");
                        }
                    }

                    // Re-open AFTER the inspector committed: this instance carries the Word
                    // edits, so saving it cannot undo them.
                    fresh = storeId != null
                        ? ns.GetItemFromID(entryIdHex, storeId)
                        : ns.GetItemFromID(entryIdHex);
                    dynamic draft = fresh!;

                    // Attachments: removals first, then additions, so removing and adding
                    // the same file name in one call means REPLACE.
                    List<string> removed = RemoveAttachmentsByName(draft, attachmentsToRemove);
                    (List<string> added, List<string> failedToAttach) = AddAttachmentsToDraft((object)draft, attachmentsToAdd);
                    if (removed.Count > 0)
                    {
                        changed.Add("attachmentsRemoved");
                    }

                    if (added.Count > 0)
                    {
                        changed.Add("attachmentsAdded");
                    }

                    // Recipients: REPLACE per class, and only for the classes the caller
                    // actually supplied (an omitted class is left alone).
                    if (toRecipients != null)
                    {
                        ReplaceRecipients(draft, 1, toRecipients, unresolved);
                        changed.Add("to");
                    }

                    if (ccRecipients != null)
                    {
                        ReplaceRecipients(draft, 2, ccRecipients, unresolved);
                        changed.Add("cc");
                    }

                    if (bccRecipients != null)
                    {
                        ReplaceRecipients(draft, 3, bccRecipients, unresolved);
                        changed.Add("bcc");
                    }

                    // Subject, with the A3 threading restore: assigning Subject makes
                    // Outlook REGENERATE PR_CONVERSATION_INDEX, detaching the draft from
                    // its thread - capture the draft's OWN index/topic and write them back
                    // afterwards (live-proven on this build in batch A).
                    bool? topicPreserved = null;
                    if (subject != null)
                    {
                        string? currentTopic = TryGetString(() => (string?)draft.ConversationTopic)
                            ?? TryGetPropertyString(draft, ConversationTopicDasl);
                        string? currentIndex = TryGetString(() => (string?)draft.ConversationIndex);

                        draft.Subject = subject;
                        changed.Add("subject");

                        byte[]? indexBytes = HexToBytes(currentIndex);
                        if (indexBytes != null || !string.IsNullOrEmpty(currentTopic))
                        {
                            bool indexRestored = indexBytes != null
                                && TrySetPropertyBinary(draft, ConversationIndexDasl, indexBytes!);
                            bool topicRestored = !string.IsNullOrEmpty(currentTopic)
                                && TrySetPropertyString(draft, ConversationTopicDasl, currentTopic!);
                            topicPreserved = indexRestored && topicRestored;
                        }
                    }

                    // Plain item properties.
                    if (importance != null)
                    {
                        try
                        {
                            draft.Importance = importance.Value;
                            changed.Add("importance");
                        }
                        catch (Exception ex) when (IsComCallFailure(ex))
                        {
                        }
                    }

                    if (requestReadReceipt != null)
                    {
                        try
                        {
                            draft.ReadReceiptRequested = requestReadReceipt.Value;
                            changed.Add("requestReadReceipt");
                        }
                        catch (Exception ex) when (IsComCallFailure(ex))
                        {
                        }
                    }

                    draft.Save();

                    bool displayed = false;
                    if (display)
                    {
                        try
                        {
                            draft.Display();
                            displayed = true;
                        }
                        catch (Exception ex) when (IsComCallFailure(ex))
                        {
                        }
                    }

                    int imagesDropped = 0;
                    if (countImages)
                    {
                        int imagesAfter = OutlookAI.Core.Text.HtmlBodyComposer.CountInlineImages(
                            TryGetString(() => (string?)((dynamic)fresh!).HTMLBody));
                        imagesDropped = Math.Max(0, imagesBefore - imagesAfter);
                    }

                    ComDraftInfo info = SnapshotDraft(fresh!);
                    return new ComDraftUpdateResult(
                        info,
                        changed,
                        unresolved,
                        SnapshotAttachments(fresh!),
                        added,
                        removed,
                        failedToAttach,
                        bodyReplaced,
                        wordPlaced,
                        displayed,
                        signatureOverride?.Name,
                        overrideApplied,
                        null,
                        topicPreserved,
                        imagesDropped);
                }
                catch (Exception ex) when (IsComCallFailure(ex))
                {
                    capturedError = DescribeComFailure(ex);
                    return null;
                }
                finally
                {
                    Release(fresh);
                    Release(item);
                }
            });

            error = capturedError;
            return result;
        }

        /// <summary>
        /// discard_draft backbone (v3.MD D46/C2, S1 v3 - the ONLY mail-deleting code path
        /// in the product). SOFT delete only: <c>MailItem.Delete()</c>, which moves the
        /// item to the store's Deleted Items exactly like pressing Delete in Outlook.
        /// <c>PermanentlyDelete</c> is never called, Deleted Items is never emptied and
        /// its contents are never touched. The caller has already proven the draft came
        /// from THIS server (<c>ServerDraftRegistry</c>); this layer re-proves the item is
        /// a mail item, is UNSENT and lives in a Drafts folder before deleting anything.
        /// A best-effort re-locate in Deleted Items returns the new EntryID so the discard
        /// stays reversible in the same way a move is (D39).
        /// </summary>
        public ComDraftDiscardResult? TryDiscardDraft(string entryIdHex, string? storeId, out string? error)
        {
            EnsureNotDisposed();
            if (string.IsNullOrWhiteSpace(entryIdHex))
            {
                throw new ArgumentException("EntryID must not be blank.", nameof(entryIdHex));
            }

            string? capturedError = null;
            ComDraftDiscardResult? result = _runner.Run<ComDraftDiscardResult?>(() =>
            {
                dynamic ns = _namespace!;
                object? item = null;
                object? parent = null;
                object? parentStore = null;
                try
                {
                    item = storeId != null
                        ? ns.GetItemFromID(entryIdHex, storeId)
                        : ns.GetItemFromID(entryIdHex);

                    capturedError = CheckEditableDraft(item!);
                    if (capturedError != null)
                    {
                        return null;
                    }

                    ComDraftInfo info = SnapshotDraft(item!);

                    // Deleted Items identity is resolved BEFORE the delete (afterwards the
                    // item's Parent is the target, not the source).
                    string? deletedItemsName = null;
                    string? deletedItemsEntryId = null;
                    parent = ((dynamic)item!).Parent;
                    if (parent != null)
                    {
                        parentStore = ((dynamic)parent!).Store;
                        if (parentStore != null)
                        {
                            object? deleted = null;
                            try
                            {
                                deleted = ((dynamic)parentStore!).GetDefaultFolder(3); // olFolderDeletedItems
                                deletedItemsName = TryGetString(() => (string?)((dynamic)deleted!).Name);
                                deletedItemsEntryId = TryGetString(() => (string?)((dynamic)deleted!).EntryID);
                            }
                            catch (Exception ex) when (IsComCallFailure(ex))
                            {
                            }
                            finally
                            {
                                Release(deleted);
                            }
                        }
                    }

                    // THE soft delete. Never PermanentlyDelete - S1 v3 allows a draft to be
                    // put in the bin, never to be destroyed.
                    ((dynamic)item!).Delete();

                    string? newEntryId = deletedItemsEntryId == null
                        ? null
                        : TryFindDiscardedCopy(deletedItemsEntryId, info.Subject, info.EntryId);

                    return new ComDraftDiscardResult(
                        info.EntryId,
                        newEntryId,
                        info.StoreDisplayName,
                        info.ParentFolderName,
                        deletedItemsName,
                        info.Subject);
                }
                catch (Exception ex) when (IsComCallFailure(ex))
                {
                    capturedError = DescribeComFailure(ex);
                    return null;
                }
                finally
                {
                    Release(parentStore);
                    Release(parent);
                    Release(item);
                }
            });

            error = capturedError;
            return result;
        }

        /// <summary>
        /// The shared fail-closed precondition gate for update_draft and discard_draft:
        /// the item must be a MAIL item, must be UNSENT, and must live in a Drafts folder.
        /// Returns a content-free error code, or null when the item may be edited.
        /// </summary>
        private static string? CheckEditableDraft(object itemObject)
        {
            if (!IsMailItem(itemObject))
            {
                return "NotAMailItem";
            }

            bool isSent = true; // fail CLOSED: an unreadable Sent flag is treated as sent
            try
            {
                isSent = (bool)((dynamic)itemObject).Sent;
            }
            catch (Exception ex) when (IsComCallFailure(ex))
            {
            }

            if (isSent)
            {
                return "AlreadySent";
            }

            return IsInDraftsFolder(itemObject) ? null : "NotInDraftsFolder";
        }

        /// <summary>
        /// True when the item's parent folder IS the store's Drafts folder or sits
        /// underneath it. Folder identity is compared by EntryID against
        /// <c>GetDefaultFolder(16)</c>, never by name - Drafts is localized (v3.MD D39's
        /// localization-proof rule).
        /// </summary>
        private static bool IsInDraftsFolder(object itemObject)
        {
            object? folder = null;
            object? store = null;
            try
            {
                folder = ((dynamic)itemObject).Parent;
                if (folder == null)
                {
                    return false;
                }

                store = ((dynamic)folder!).Store;
                if (store == null)
                {
                    return false;
                }

                string? draftsEntryId = null;
                object? drafts = null;
                try
                {
                    drafts = ((dynamic)store!).GetDefaultFolder(16); // olFolderDrafts
                    draftsEntryId = TryGetString(() => (string?)((dynamic)drafts!).EntryID);
                }
                catch (Exception ex) when (IsComCallFailure(ex))
                {
                }
                finally
                {
                    Release(drafts);
                }

                if (string.IsNullOrEmpty(draftsEntryId))
                {
                    return false;
                }

                // Walk up from the item's folder: Drafts itself, or any folder below it.
                object? current = folder;
                folder = null; // ownership moves to the walk loop
                for (int depth = 0; depth < FolderWalkDepthGuard && current != null; depth++)
                {
                    string? currentId = TryGetString(() => (string?)((dynamic)current!).EntryID);
                    if (string.Equals(currentId, draftsEntryId, StringComparison.OrdinalIgnoreCase))
                    {
                        Release(current);
                        return true;
                    }

                    object? next = null;
                    try
                    {
                        next = ((dynamic)current!).Parent;
                    }
                    catch (Exception ex) when (IsComCallFailure(ex))
                    {
                    }

                    Release(current);
                    current = next;

                    // The store root's Parent is the Namespace/store object, which has no
                    // EntryID - the walk ends there rather than looping.
                    if (current != null && TryGetString(() => (string?)((dynamic)current!).EntryID) == null)
                    {
                        Release(current);
                        current = null;
                    }
                }

                // Depth guard reached with a live reference still held.
                Release(current);
                return false;
            }
            catch (Exception ex) when (IsComCallFailure(ex))
            {
                return false;
            }
            finally
            {
                Release(store);
                Release(folder);
            }
        }

        /// <summary>
        /// The held-Inspector revision used by update_draft: acquire the inspector ONCE,
        /// optionally swap the signature region, optionally REPLACE the draft region with
        /// the new body, then flush with <c>Close(olSave)</c>. Any failure closes the
        /// inspector with <c>olDiscard</c> so a half-written document never reaches the
        /// item, and reports the error instead of falling back to an appending splice.
        /// </summary>
        private static (bool Ok, string? Error) ReviseHeldDocument(
            object draftObject,
            ComDraftBody? body,
            ComSignatureOverride? signatureOverride)
        {
            dynamic draft = draftObject;
            object? inspector = null;
            object? document = null;
            string? error = null;
            bool flushed = false;
            try
            {
                if (signatureOverride != null && !File.Exists(signatureOverride.FilePath))
                {
                    return (false, "SignatureFileMissing");
                }

                // EXACTLY ONE acquisition, deliberately - matching the creators.
                // A retry loop was tried here and REMOVED: every GetInspector call
                // materializes another hidden Inspector for the same item (v3.MD section
                // 12), and asking for a second one after releasing the first wedged
                // Outlook indefinitely on a headless instance. One inspector, held to the
                // close, is the only shape this codebase has ever proven.
                try
                {
                    inspector = draft.GetInspector;
                }
                catch (Exception ex) when (IsComCallFailure(ex))
                {
                }

                if (inspector == null)
                {
                    return (false, "NoInspector");
                }

                // ⚠ ACTIVATE BEFORE EDITING - live-measured, and the difference between
                // a working revision and a silent no-op. On a NEW item the WordEditor of a
                // hidden inspector is live and its edits commit. On an ALREADY-SAVED draft
                // the hidden inspector hands back a document whose edits go nowhere: the
                // call reports success and the stored HTMLBody changes by EXACTLY ZERO
                // bytes (measured: 38932 -> 38932, new text absent, old text intact).
                // Activating the inspector materializes the real editing surface.
                // D49: the Activate() is unchanged in PURPOSE and now correct in EFFECT.
                // It was always required (see above), but on a window-less Outlook a bare
                // Activate() PAINTS the compose window where the user can see it - a D33
                // violation that shipped unnoticed because it only happens headless, and
                // the headless path then died anyway. ComposeSurface parks the window
                // off-screen FIRST, activates, hides whatever became visible, and hands
                // back the WordEditor. Nothing is user-visible at any point.
                document = ComposeSurface.PromoteForWordEditor(inspector!, out string? promoteError);
                if (document == null)
                {
                    error = promoteError ?? "NoWordEditor";
                }

                if (error == null && signatureOverride != null)
                {
                    (bool sigOk, string? sigError) = ApplySignatureToDocument(document!, signatureOverride.FilePath);
                    error = sigOk ? null : sigError ?? "SignatureInsertFailed";
                }

                if (error == null && body != null)
                {
                    // replaceWholeDocumentWhenNoBoundary: an update must clear the PREVIOUS
                    // body even when the draft carries neither a signature nor a quoted
                    // original to bound the region.
                    (bool bodyOk, string? bodyError) = InsertBodyAboveSignature(
                        document!, body, replaceWholeDocumentWhenNoBoundary: true);
                    error = bodyOk ? null : bodyError ?? "BodyInsertFailed";
                }

                if (error == null)
                {
                    // D47, belt to the create-path braces: any picture still LINKED in this
                    // document (an older draft, or a signature just re-inserted by the
                    // override above) is embedded now, so this re-render is the last one
                    // that could drop it.
                    _ = EmbedLinkedPictures(document!);
                }

                if (error == null)
                {
                    // ⚠ THE COMMIT, and it is NOT the creators' Close(olSave).
                    // On a NEW item Close(olSave) is what writes the Word document into
                    // the item, because closing is what saves an item that was never
                    // saved. On an ALREADY-SAVED draft it does nothing - live-measured:
                    // the call reported success and the re-opened draft still held the
                    // OLD body, with the new text nowhere in the document.
                    // What does commit is saving the inspector's OWN item -
                    // Inspector.CurrentItem is the instance the WordEditor edits, so
                    // Save() on THAT writes the document through. The close is then
                    // olDiscard on purpose: the save already happened, and a second
                    // save-on-close would re-render the document from the item (the D37
                    // footgun) for no gain.
                    object? currentItem = null;
                    try
                    {
                        currentItem = ((dynamic)inspector!).CurrentItem;
                        if (currentItem != null)
                        {
                            ((dynamic)currentItem!).Save();
                            flushed = true;
                        }
                        else
                        {
                            error = "NoCurrentItem";
                        }
                    }
                    catch (Exception ex) when (IsComCallFailure(ex))
                    {
                        error = DescribeComFailure(ex);
                    }
                    finally
                    {
                        Release(currentItem);
                    }

                    if (flushed)
                    {
                        try
                        {
                            ((dynamic)inspector!).Close(1); // olDiscard - already saved
                        }
                        catch (Exception ex) when (IsComCallFailure(ex))
                        {
                        }
                    }
                }

                return (flushed, error);
            }
            catch (Exception ex) when (IsComCallFailure(ex))
            {
                return (false, DescribeComFailure(ex));
            }
            finally
            {
                if (!flushed && inspector != null)
                {
                    // Discard, never save: the draft must survive a failed revision
                    // untouched rather than half-rewritten. (On the success path the
                    // inspector was already closed with olDiscard after the explicit
                    // save.)
                    try
                    {
                        ((dynamic)inspector!).Close(1); // olDiscard
                    }
                    catch (Exception ex) when (IsComCallFailure(ex))
                    {
                    }
                }

                Release(document);
                Release(inspector);
            }
        }

        /// <summary>
        /// REPLACE semantics for one recipient class (D46/C1): every existing recipient of
        /// that type is removed - descending, because <c>Recipients.Remove</c> reindexes -
        /// and the supplied addresses are added in its place. Other classes are untouched.
        /// </summary>
        private static void ReplaceRecipients(dynamic mail, int type, IReadOnlyList<string> addresses, ICollection<string> unresolved)
        {
            object? recipients = null;
            try
            {
                recipients = mail.Recipients;
                dynamic collection = (dynamic)recipients!;
                int count = collection.Count;
                List<int> doomed = new List<int>();
                for (int i = 1; i <= count; i++)
                {
                    object? recipient = null;
                    try
                    {
                        recipient = collection[i];
                        int recipientType = (int)((dynamic)recipient!).Type;
                        if (recipientType == type)
                        {
                            doomed.Add(i);
                        }
                    }
                    catch (Exception ex) when (IsComCallFailure(ex))
                    {
                    }
                    finally
                    {
                        Release(recipient);
                    }
                }

                for (int i = doomed.Count - 1; i >= 0; i--)
                {
                    try
                    {
                        collection.Remove(doomed[i]);
                    }
                    catch (Exception ex) when (IsComCallFailure(ex))
                    {
                    }
                }
            }
            finally
            {
                Release(recipients);
            }

            AddRecipients(mail, addresses, type, unresolved);
        }

        /// <summary>
        /// Attaches already-validated absolute paths to a draft and returns the names that
        /// went on. The paths were existence/readability-checked PRE-COM
        /// (<c>DraftAttachments.Validate</c>); a COM refusal here is surfaced by throwing,
        /// because a draft that silently misses a file the agent believes it attached is
        /// exactly the failure mode the fail-closed validation exists to prevent.
        /// </summary>
        private static (List<string> Added, List<string> Failed) AddAttachmentsToDraft(dynamic mail, IReadOnlyList<string>? paths)
        {
            List<string> added = new List<string>();
            List<string> failed = new List<string>();
            if (paths == null || paths.Count == 0)
            {
                return (added, failed);
            }

            object? attachments = null;
            try
            {
                attachments = mail.Attachments;
                dynamic collection = (dynamic)attachments!;
                foreach (string path in paths)
                {
                    object? attachment = null;
                    try
                    {
                        attachment = collection.Add(path);
                        added.Add(Path.GetFileName(path));
                    }
                    catch (Exception ex) when (IsComCallFailure(ex))
                    {
                        // A COM refusal here is rare (every path was existence- and
                        // readability-checked pre-COM), but it must NOT abort a draft that
                        // already exists - that would leave the caller with a saved draft
                        // and an error claiming nothing happened. Report the file instead,
                        // loudly, in the outcome and the audit line.
                        failed.Add(Path.GetFileName(path));
                    }
                    finally
                    {
                        Release(attachment);
                    }
                }
            }
            finally
            {
                Release(attachments);
            }

            return (added, failed);
        }

        /// <summary>
        /// Removes attachments by FILE NAME (case-insensitive), descending because
        /// <c>Attachments.Delete</c> reindexes. Returns the names actually removed - a name
        /// that matched nothing simply does not appear, and the caller reports that.
        /// </summary>
        private static List<string> RemoveAttachmentsByName(dynamic mail, IReadOnlyList<string>? names)
        {
            List<string> removed = new List<string>();
            if (names == null || names.Count == 0)
            {
                return removed;
            }

            HashSet<string> wanted = new HashSet<string>(names, StringComparer.OrdinalIgnoreCase);
            object? attachments = null;
            try
            {
                attachments = mail.Attachments;
                dynamic collection = (dynamic)attachments!;
                int count = collection.Count;
                List<int> doomed = new List<int>();
                List<string> doomedNames = new List<string>();
                for (int i = 1; i <= count; i++)
                {
                    object? attachment = null;
                    try
                    {
                        attachment = collection[i];
                        string? fileName = TryGetString(() => (string?)((dynamic)attachment!).FileName);
                        if (fileName != null && wanted.Contains(fileName))
                        {
                            doomed.Add(i);
                            doomedNames.Add(fileName);
                        }
                    }
                    catch (Exception ex) when (IsComCallFailure(ex))
                    {
                    }
                    finally
                    {
                        Release(attachment);
                    }
                }

                for (int i = doomed.Count - 1; i >= 0; i--)
                {
                    object? attachment = null;
                    try
                    {
                        attachment = collection[doomed[i]];
                        ((dynamic)attachment!).Delete();
                        removed.Add(doomedNames[i]);
                    }
                    catch (Exception ex) when (IsComCallFailure(ex))
                    {
                    }
                    finally
                    {
                        Release(attachment);
                    }
                }
            }
            finally
            {
                Release(attachments);
            }

            removed.Reverse();
            return removed;
        }

        /// <summary>
        /// Attachment snapshot of an item (name + size, 1-based index) - the shape
        /// <c>read</c> already reports, reused by the draft tools' result and by the send
        /// content hash.
        /// </summary>
        private static IReadOnlyList<ComAttachmentInfo> SnapshotAttachments(object itemObject)
        {
            List<ComAttachmentInfo> infos = new List<ComAttachmentInfo>();
            object? attachments = null;
            try
            {
                attachments = ((dynamic)itemObject).Attachments;
                dynamic collection = (dynamic)attachments!;
                int count = collection.Count;
                for (int i = 1; i <= count; i++)
                {
                    object? attachment = null;
                    try
                    {
                        attachment = collection[i];
                        dynamic a = (dynamic)attachment!;
                        string? fileName = TryGetString(() => (string?)a.FileName);
                        long? size = null;
                        try
                        {
                            size = (long)(int)a.Size;
                        }
                        catch (Exception ex) when (IsComCallFailure(ex))
                        {
                        }

                        infos.Add(new ComAttachmentInfo(i, fileName, size));
                    }
                    finally
                    {
                        Release(attachment);
                    }
                }
            }
            catch (Exception ex) when (IsComCallFailure(ex))
            {
            }
            finally
            {
                Release(attachments);
            }

            return infos;
        }

        /// <summary>
        /// Best-effort re-locate of a just-discarded draft inside Deleted Items so the
        /// outcome can carry a usable newEntryId (EntryIDs change on ANY move). Read-only
        /// and failure-tolerant: nothing depends on finding it, and Deleted Items contents
        /// are never modified.
        /// </summary>
        private string? TryFindDiscardedCopy(string deletedItemsEntryId, string? subject, string oldEntryId)
        {
            object? folder = null;
            object? items = null;
            try
            {
                dynamic ns = _namespace!;
                folder = ns.GetFolderFromID(deletedItemsEntryId);
                if (folder == null)
                {
                    return null;
                }

                items = ((dynamic)folder!).Items;
                dynamic collection = (dynamic)items!;
                try
                {
                    collection.Sort("[LastModificationTime]", true);
                }
                catch (Exception ex) when (IsComCallFailure(ex))
                {
                }

                int count = collection.Count;
                int scanned = 0;
                for (int i = 1; i <= count && scanned < DiscardRelocateScanCap; i++, scanned++)
                {
                    object? candidate = null;
                    try
                    {
                        candidate = collection[i];
                        if (!IsMailItem(candidate!))
                        {
                            continue;
                        }

                        string? candidateSubject = TryGetString(() => (string?)((dynamic)candidate!).Subject);
                        if (!string.Equals(candidateSubject ?? string.Empty, subject ?? string.Empty, StringComparison.Ordinal))
                        {
                            continue;
                        }

                        string? candidateId = TryGetString(() => (string?)((dynamic)candidate!).EntryID);
                        if (candidateId != null
                            && !string.Equals(candidateId, oldEntryId, StringComparison.OrdinalIgnoreCase))
                        {
                            return candidateId;
                        }
                    }
                    catch (Exception ex) when (IsComCallFailure(ex))
                    {
                    }
                    finally
                    {
                        Release(candidate);
                    }
                }

                return null;
            }
            catch (Exception ex) when (IsComCallFailure(ex))
            {
                return null;
            }
            finally
            {
                Release(items);
                Release(folder);
            }
        }

        /// <summary>STA-side: true when the object reports OlObjectClass 43 (olMail).</summary>
        private static bool IsMailItem(object itemObject)
        {
            try
            {
                return (int)((dynamic)itemObject).Class == OlMailItemClass;
            }
            catch (Exception ex) when (IsComCallFailure(ex))
            {
                return false;
            }
        }

        /// <summary>
        /// ⚠ LIVE-VERIFIED FOOTGUN (Phase 4): <c>MailItem.SendUsingAccount</c> is a
        /// PROPERTYPUTREF property. A late-bound dynamic assignment
        /// (<c>mail.SendUsingAccount = account</c>) SILENTLY NO-OPS on this Outlook
        /// build - no exception, the getter stays null and the DEFAULT account would
        /// send. The putref accessor must be invoked explicitly. A failure here throws
        /// (identity is load-bearing); inner COM errors are unwrapped so the standard
        /// IsComCallFailure handling applies.
        /// </summary>
        private static void SetSendUsingAccount(object mailObject, object accountObject)
        {
            try
            {
                mailObject.GetType().InvokeMember(
                    "SendUsingAccount",
                    System.Reflection.BindingFlags.PutRefDispProperty
                        | System.Reflection.BindingFlags.Public
                        | System.Reflection.BindingFlags.Instance,
                    null,
                    mailObject,
                    new[] { accountObject },
                    CultureInfo.InvariantCulture);
            }
            catch (System.Reflection.TargetInvocationException ex) when (ex.InnerException != null)
            {
                throw ex.InnerException;
            }
        }

        /// <summary>
        /// STA-side body composition shared by both draft creators (rewritten in soak
        /// fix batch A - A1). BOTH paths now compose inside Word through ONE HELD
        /// Inspector, which is what Outlook's own compose window does:
        /// <c>GetInspector</c> makes Outlook inject the account's OWN signature natively
        /// (new-mail or reply/forward rendition, HTML and resources intact), an optional
        /// override swaps that region via the <c>_MailAutoSig</c> bookmark dance, the
        /// agent body is written ABOVE the signature region with the marker deleted and
        /// recreated around the untouched signature (the add-in's proven
        /// <c>AITaskPane.WriteDraftToDocument</c> technique), and
        /// <c>Inspector.Close(olSave)</c> flushes the document into the item.
        /// <para>
        /// The retired default path assigned <c>HTMLBody</c> once with the body spliced
        /// in after the &lt;body&gt; tag. That is string surgery on Outlook's own markup:
        /// it left the agent text OUTSIDE Word's WordSection1 container (so it did not
        /// inherit the message style), and when combined with an override on an account
        /// with no default signature it produced the A1 defect - the body ended up INSIDE
        /// the recreated <c>_MailAutoSig</c> bookmark (live-proven: the saved HTML opened
        /// with &lt;a name="_MailAutoSig"&gt; around the agent text), i.e. Outlook and the
        /// add-in both considered the whole message to be the signature.
        /// </para>
        /// <para>
        /// PROBED on this machine (D37, unchanged and load-bearing): Word-document edits
        /// NEVER reach the item via <c>item.Save()</c>; only Close(olSave) on the
        /// inspector that hosted the edits flushes them (an item.Save() BETWEEN the edits
        /// and the close re-renders the document from the item and silently wipes them,
        /// and a close via a re-acquired inspector reference loses them too).
        /// </para>
        /// If ANY step fails, the composition falls back to the previous wholesale
        /// HTMLBody assignment (whose input still carries the injected signature), so a
        /// draft is never lost or left body-less; the failure is reported content-free.
        /// </summary>
        private static (bool SignatureInjected, long TextBefore, long TextAfter, bool OverrideApplied, string? OverrideError, bool BodyPlacedViaWordEditor, bool SurfacePromoted) ComposeDraft(
            object draftObject,
            ComDraftBody body,
            ComSignatureOverride? signatureOverride)
        {
            dynamic draft = draftObject;
            string htmlBefore = TryGetString(() => (string?)draft.HTMLBody) ?? string.Empty;
            long textBefore = CountNonWhitespaceText(htmlBefore);
            string htmlAfter = htmlBefore;
            long textAfter = textBefore;
            bool injected = false;
            bool wordComposeDone = false;
            bool promoted = false;
            string? error = null;

            object? inspector = null;
            object? document = null;
            try
            {
                if (signatureOverride != null && !File.Exists(signatureOverride.FilePath))
                {
                    error = "SignatureFileMissing";
                }

                if (error == null)
                {
                    try
                    {
                        inspector = draft.GetInspector;
                    }
                    catch (Exception ex) when (IsComCallFailure(ex))
                    {
                    }

                    // Signature injection is measured across the GetInspector touch
                    // regardless of what happens next (text-based: HTML template
                    // expansion without a signature adds markup but no text).
                    htmlAfter = TryGetString(() => (string?)draft.HTMLBody) ?? string.Empty;
                    textAfter = CountNonWhitespaceText(htmlAfter);
                    injected = textAfter > textBefore;

                    if (inspector == null)
                    {
                        error = "NoInspector";
                    }
                }

                if (error == null)
                {
                    try
                    {
                        document = ((dynamic)inspector!).WordEditor;
                    }
                    catch (Exception ex) when (IsComCallFailure(ex))
                    {
                        // ⚠ Headless does NOT return null here, it THROWS
                        // COMException "The operation failed." (D49 Phase-1 finding 1).
                        // Note also that Inspector.IsWordMail() reports TRUE in exactly
                        // this state, so it is never a usable gate.
                    }

                    if (document == null)
                    {
                        // D49 THE EDITOR PROMOTION. Outlook is window-less, so the editor
                        // does not exist yet. Park the inspector's (already existing,
                        // invisible) window off-screen, Activate it, and hide whatever the
                        // activation put on screen: measured 53-79 ms to a live WordEditor
                        // with nothing user-visible at any point. Only reached when the
                        // editor was unobtainable, so a windowed Outlook never gets its
                        // windows touched and its behaviour is byte-identical to before.
                        document = ComposeSurface.PromoteForWordEditor(inspector!, out string? promoteError);
                        promoted = document != null;
                        if (document == null)
                        {
                            error = promoteError ?? "NoWordEditor";
                        }
                    }
                }

                if (error == null && signatureOverride != null)
                {
                    (bool sigOk, string? sigError) = ApplySignatureToDocument(document!, signatureOverride.FilePath);
                    error = sigOk ? null : sigError ?? "SignatureInsertFailed";
                }

                if (error == null)
                {
                    (bool bodyOk, string? bodyError) = InsertBodyAboveSignature(document!, body);
                    error = bodyOk ? null : bodyError ?? "BodyInsertFailed";
                }

                if (error == null)
                {
                    // D47: embed the signature's images instead of leaving them as
                    // file:/// links into the Signatures directory. Done HERE, on the
                    // create path, because that is where the link is born - an update
                    // then starts from an embedded cid: image and re-renders losslessly.
                    _ = EmbedLinkedPictures(document!);
                }

                if (error == null)
                {
                    // The load-bearing flush (probe-proven): olSave on the SAME held
                    // inspector commits the Word edits into the item.
                    try
                    {
                        ((dynamic)inspector!).Close(0); // olSave
                        wordComposeDone = true;
                    }
                    catch (Exception ex) when (IsComCallFailure(ex))
                    {
                        error = DescribeComFailure(ex);
                    }
                }
            }
            catch (Exception ex) when (IsComCallFailure(ex))
            {
                error = DescribeComFailure(ex);
            }
            finally
            {
                Release(document);
                Release(inspector);
            }

            if (!wordComposeDone)
            {
                // Fallback = the pre-batch-A composition. Its input HTML still contains
                // the injected signature, and the wholesale HTMLBody assignment
                // re-renders the Word document - discarding any partial Word edits.
                string html = TryGetString(() => (string?)draft.HTMLBody) ?? string.Empty;
                if (html.Length == 0)
                {
                    html = htmlAfter;
                }

                // An HTML body is ALREADY normalized markup - it must not be escaped here
                // (that would show the agent its own tags as text); it only gets the same
                // <div> wrapper the text path uses so the splice has one root element.
                string fragment = body.IsHtml
                    ? "<div>" + body.Html + "</div>"
                    : OutlookAI.Core.Text.HtmlBodyComposer.ToHtmlFragment(body.Text);
                draft.HTMLBody = OutlookAI.Core.Text.HtmlBodyComposer.InsertAtBodyTop(
                    html.Length > 0 ? html : null, fragment);
            }

            return (injected, textBefore, textAfter, signatureOverride != null && wordComposeDone, error, wordComposeDone, promoted && wordComposeDone);
        }

        /// <summary>
        /// The bookmark dance itself, on an already-acquired Word document: replace the
        /// _MailAutoSig region (or insert above _MailOriginal / at document end when no
        /// signature region exists) with the signature file's content and recreate the
        /// _MailAutoSig bookmark over it.
        /// </summary>
        private static (bool Applied, string? Error) ApplySignatureToDocument(object documentObject, string signatureFilePath)
        {
            dynamic doc = documentObject;
            object? bookmarks = null;
            try
            {
                bookmarks = doc.Bookmarks;
                dynamic bm = (dynamic)bookmarks!;

                // _MailAutoSig/_MailOriginal are HIDDEN bookmarks - invisible to
                // Exists() unless ShowHidden is on (add-in AITaskPane pattern).
                bm.ShowHidden = true;

                int insertAt;
                if ((bool)bm.Exists("_MailAutoSig"))
                {
                    // Replace: drop the marker, then the signature content itself.
                    object? sigBookmark = null;
                    object? sigRange = null;
                    try
                    {
                        sigBookmark = bm.Item("_MailAutoSig");
                        sigRange = ((dynamic)sigBookmark!).Range;
                        insertAt = (int)((dynamic)sigRange!).Start;
                        ((dynamic)sigBookmark).Delete();
                        ((dynamic)sigRange).Delete();
                    }
                    finally
                    {
                        Release(sigRange);
                        Release(sigBookmark);
                    }
                }
                else if ((bool)bm.Exists("_MailOriginal"))
                {
                    // No default signature (e.g. account without one): insert directly
                    // ABOVE the quoted original.
                    object? origBookmark = null;
                    object? origRange = null;
                    try
                    {
                        origBookmark = bm.Item("_MailOriginal");
                        origRange = ((dynamic)origBookmark!).Range;
                        insertAt = (int)((dynamic)origRange!).Start;
                    }
                    finally
                    {
                        Release(origRange);
                        Release(origBookmark);
                    }
                }
                else
                {
                    // Plain new draft: end of document (before the final paragraph mark).
                    object? content = null;
                    try
                    {
                        content = doc.Content;
                        insertAt = Math.Max(0, (int)((dynamic)content!).End - 1);
                    }
                    finally
                    {
                        Release(content);
                    }
                }

                int endBefore = GetDocumentEnd(doc);
                object? insertRange = null;
                try
                {
                    insertRange = doc.Range(insertAt, insertAt);
                    ((dynamic)insertRange!).InsertFile(signatureFilePath, Type.Missing, false, false, false);
                }
                finally
                {
                    Release(insertRange);
                }

                int endAfter = GetDocumentEnd(doc);
                int newEnd = insertAt + Math.Max(0, endAfter - endBefore);

                // Recreate the marker over the inserted content so Outlook (and the
                // add-in's draft/signature/quote split) keep working on this draft.
                object? newRange = null;
                try
                {
                    newRange = doc.Range(insertAt, newEnd);
                    bm.Add("_MailAutoSig", newRange);
                }
                finally
                {
                    Release(newRange);
                }

                return endAfter > endBefore ? (true, null) : (false, "InsertFileAddedNothing");
            }
            catch (Exception ex) when (IsComCallFailure(ex))
            {
                return (false, DescribeComFailure(ex));
            }
            finally
            {
                Release(bookmarks);
            }
        }

        /// <summary>
        /// Turns every LINKED picture in the compose document into an EMBEDDED one, and
        /// returns how many it converted. Best-effort by design: a failure here must never
        /// fail a composition, so every COM problem is swallowed and simply not counted.
        /// <para>
        /// ⚠ THE MECHANISM THIS EXISTS FOR (live-measured, D47). Word's
        /// <c>Range.InsertFile</c> of a signature .htm does NOT embed the signature's
        /// images: it inserts each one as a LINKED <c>InlineShape</c> pointing at the file
        /// on disk, and Outlook stores that link verbatim, so the saved draft carries
        /// <c>&lt;img src="file:///…\Signatures\&lt;name&gt;_files\logo.png"&gt;</c> and
        /// ZERO attachments. The link only renders because the file happens to sit on this
        /// machine. Re-rendering such a document - which is exactly what
        /// <c>update_draft</c> does when the held inspector re-materializes the saved HTML
        /// - makes Word re-serialize the unresolved linked picture as its VML placeholder
        /// AutoShape (<c>&lt;v:rect … alt="logo"&gt;</c>), and the <c>&lt;img&gt;</c>
        /// disappears. Embedding the picture at composition time removes the link, so
        /// Outlook writes real image bytes as an inline <c>cid:</c> attachment that
        /// survives any number of re-renders - and reaches the recipient, which a
        /// <c>file:///</c> link never could.
        /// </para>
        /// <c>SavePictureWithDocument</c> must be set BEFORE <c>BreakLink</c>: breaking a
        /// link whose picture is not stored with the document leaves nothing behind.
        /// </summary>
        private static int EmbedLinkedPictures(object documentObject)
        {
            dynamic doc = documentObject;
            int embedded = 0;
            object? shapes = null;
            try
            {
                shapes = doc.InlineShapes;
                dynamic collection = (dynamic)shapes!;
                int count = (int)collection.Count;
                for (int i = 1; i <= count; i++)
                {
                    object? shape = null;
                    object? link = null;
                    try
                    {
                        shape = collection[i];

                        // An already-embedded picture has no LinkFormat: late-bound COM
                        // answers either null or a binding failure, and both mean skip.
                        link = ((dynamic)shape!).LinkFormat;
                        if (link == null)
                        {
                            continue;
                        }

                        ((dynamic)link!).SavePictureWithDocument = true;
                        ((dynamic)link!).BreakLink();
                        embedded++;
                    }
                    catch (Exception ex) when (IsComCallFailure(ex))
                    {
                    }
                    finally
                    {
                        Release(link);
                        Release(shape);
                    }
                }
            }
            catch (Exception ex) when (IsComCallFailure(ex))
            {
            }
            finally
            {
                Release(shapes);
            }

            return embedded;
        }

        /// <summary>
        /// Writes the agent's plain-text body into the DRAFT region of an
        /// already-acquired Word document - the add-in's proven
        /// <c>AITaskPane.WriteDraftToDocument</c> technique ported to late-bound COM.
        /// <para>
        /// The draft region ends where the signature (<c>_MailAutoSig</c>) or, absent
        /// that, the quoted original (<c>_MailOriginal</c>) begins; with neither marker
        /// present the body simply goes to the document start. The boundary marker is
        /// DELETED before the write and RECREATED over the (shifted, byte-identical)
        /// region afterwards. That dance is load-bearing, not cosmetic: Word absorbs text
        /// inserted at a bookmark's start INTO that bookmark, so writing straight to
        /// Range(0,0) while <c>_MailAutoSig</c> starts at 0 made the agent body part of
        /// the signature region (the A1 defect - live-proven on an account without a
        /// default signature). Recreating the marker keeps Outlook's signature handling
        /// and the add-in's draft/signature/thread split working on the draft.
        /// </para>
        /// Line breaks become soft line breaks (vertical tab), matching the fallback
        /// path's &lt;br&gt; behavior; one paragraph mark separates the body from what
        /// follows. The signature and quote regions themselves are never modified.
        /// <para>
        /// An HTML body (batch B - B1) takes the same route with one substitution: the
        /// draft region is CLEARED and the markup is inserted with
        /// <c>Range.InsertFile</c> from a temporary .htm file - the same verb
        /// <see cref="ApplySignatureToDocument"/> already uses to place a signature, and
        /// the only rich-content mechanism proven in this codebase (the add-in inserts
        /// signatures the same way). Word's own HTML converter does the rendering, so
        /// headings, lists and tables arrive as real Word structures inside the message's
        /// WordSection1, and the boundary marker is re-anchored by the SAME
        /// <c>Content.End</c> delta arithmetic - a collapsed insert range does not span
        /// what was inserted, so the length must be measured, not read off the range.
        /// </para>
        /// </summary>
        private static (bool Inserted, string? Error) InsertBodyAboveSignature(
            object documentObject,
            ComDraftBody body,
            bool replaceWholeDocumentWhenNoBoundary = false)
        {
            dynamic doc = documentObject;
            object? bookmarks = null;
            string? boundary = null;
            int boundaryStart = -1;
            int boundaryEnd = -1;
            try
            {
                bookmarks = doc.Bookmarks;
                dynamic bm = (dynamic)bookmarks!;

                // _MailAutoSig/_MailOriginal are HIDDEN bookmarks - invisible to
                // Exists() unless ShowHidden is on (add-in AITaskPane pattern).
                bm.ShowHidden = true;

                if ((bool)bm.Exists("_MailAutoSig"))
                {
                    boundary = "_MailAutoSig";
                }
                else if ((bool)bm.Exists("_MailOriginal"))
                {
                    boundary = "_MailOriginal";
                }

                if (boundary != null)
                {
                    object? marker = null;
                    object? markerRange = null;
                    try
                    {
                        marker = bm.Item(boundary);
                        markerRange = ((dynamic)marker!).Range;
                        boundaryStart = (int)((dynamic)markerRange!).Start;
                        boundaryEnd = (int)((dynamic)markerRange).End;

                        // Marker only - the signature/quote CONTENT stays untouched.
                        ((dynamic)marker).Delete();
                    }
                    finally
                    {
                        Release(markerRange);
                        Release(marker);
                    }
                }

                // The DRAFT region [0, boundary) is Outlook's empty compose boilerplate on
                // a freshly created item - it is REPLACED, exactly like the add-in's
                // WriteDraftToDocument does, so the body starts at the top instead of
                // below the template's blank paragraphs. A trailing empty paragraph keeps
                // Outlook's own blank line between body and signature/quote.
                //
                // ⚠ WITH NO BOUNDARY MARKER the draft region has no upper bound, and the
                // two callers need OPPOSITE things (live-caught by the batch-C reply test):
                // on a CREATE the document holds only empty boilerplate, so writing at 0
                // is correct and safe; on an UPDATE the document holds the PREVIOUS body,
                // so writing at 0 would PREPEND - the new text above the old one, the exact
                // duplication update_draft exists to avoid. When the caller is revising, the
                // draft region therefore runs to the end of the document (minus Word's
                // final paragraph mark, which cannot be deleted).
                int insertAt = boundary != null ? boundaryStart : 0;
                if (boundary == null && replaceWholeDocumentWhenNoBoundary)
                {
                    insertAt = Math.Max(0, GetDocumentEnd(doc) - 1);
                }
                int writtenEnd;
                if (body.IsHtml)
                {
                    writtenEnd = InsertHtmlIntoDraftRegion(doc, body.Html, insertAt, boundary != null);
                }
                else
                {
                    string normalized = body.Text.Replace("\r\n", "\n").Replace('\n', '\v');
                    object? range = null;
                    try
                    {
                        range = doc.Range(0, insertAt);
                        ((dynamic)range!).Text = normalized + (boundary != null ? "\r\r" : "\r");
                        writtenEnd = (int)((dynamic)range).End;
                    }
                    finally
                    {
                        Release(range);
                    }
                }

                if (boundary != null)
                {
                    // The region moved by (new draft length - old draft length).
                    int shifted = boundaryEnd + (writtenEnd - insertAt);
                    if (shifted < writtenEnd)
                    {
                        shifted = writtenEnd;
                    }

                    // Word's HTML converter may merge the last inserted paragraph with the
                    // signature's first, leaving the document a character or two shorter
                    // than the arithmetic predicts. An out-of-range Range() would throw and
                    // silently demote the whole composition to the HTML fallback, so clamp.
                    int documentEnd = GetDocumentEnd(doc);
                    if (shifted > documentEnd)
                    {
                        shifted = documentEnd;
                    }

                    if (writtenEnd > shifted)
                    {
                        writtenEnd = shifted;
                    }

                    object? restoreRange = null;
                    try
                    {
                        restoreRange = doc.Range(writtenEnd, shifted);
                        bm.Add(boundary, restoreRange);
                    }
                    finally
                    {
                        Release(restoreRange);
                    }
                }

                return (true, null);
            }
            catch (Exception ex) when (IsComCallFailure(ex))
            {
                // The marker was deleted up front; restore it over the ORIGINAL range so
                // a failed write cannot silently strip the signature/quote anchor (the
                // add-in does the same). The caller falls back to the HTML path, which
                // re-renders the document from the item anyway.
                TryRestoreBookmark(bookmarks, boundary, boundaryStart, boundaryEnd);
                return (false, DescribeComFailure(ex));
            }
            finally
            {
                Release(bookmarks);
            }
        }

        /// <summary>
        /// Places an already-normalized HTML fragment into the draft region of the held
        /// inspector's Word document and returns the document offset just PAST it (which
        /// is where the signature/quote region now begins).
        /// <para>
        /// Mechanism: clear the compose boilerplate <c>[0, draftEnd)</c>, then
        /// <c>Range.InsertFile</c> a temporary .htm file at a COLLAPSED range at 0 - the
        /// same call <see cref="ApplySignatureToDocument"/> makes for signatures, i.e. the
        /// one HTML-into-Word route this codebase has proven. Word's HTML converter turns
        /// the markup into real Word structures inside the message body, so the result
        /// inherits the message style instead of sitting outside WordSection1 (the batch-A
        /// A1(ii) defect). The inserted LENGTH must be measured as a <c>Content.End</c>
        /// delta: a collapsed insert range does not grow to span what was inserted.
        /// </para>
        /// </summary>
        private static int InsertHtmlIntoDraftRegion(dynamic doc, string html, int draftEnd, bool hasBoundary)
        {
            string path = WriteTemporaryHtmlFile(html, hasBoundary);
            try
            {
                if (draftEnd > 0)
                {
                    // Outlook's empty compose boilerplate - replaced, not appended to,
                    // exactly like the plain-text path's Range(0, draftEnd).Text write.
                    object? clearRange = null;
                    try
                    {
                        clearRange = doc.Range(0, draftEnd);
                        ((dynamic)clearRange!).Delete();
                    }
                    finally
                    {
                        Release(clearRange);
                    }
                }

                int endBefore = GetDocumentEnd(doc);
                object? insertRange = null;
                try
                {
                    insertRange = doc.Range(0, 0);
                    ((dynamic)insertRange!).InsertFile(path, Type.Missing, false, false, false);
                }
                finally
                {
                    Release(insertRange);
                }

                int endAfter = GetDocumentEnd(doc);
                return Math.Max(0, endAfter - endBefore);
            }
            finally
            {
                try
                {
                    File.Delete(path);
                }
                catch (IOException)
                {
                }
                catch (UnauthorizedAccessException)
                {
                }
            }
        }

        /// <summary>
        /// Writes the fragment as a self-contained utf-8 .htm file for Word's converter.
        /// It lives under the shared state root (v3.MD section 0.5.2), never in a mailbox
        /// path, and is deleted immediately after the insert; a trailing empty paragraph
        /// is added when a signature/quote region follows, so Word cannot merge the last
        /// body paragraph into it.
        /// </summary>
        private static string WriteTemporaryHtmlFile(string html, bool hasBoundary)
        {
            string directory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "OutlookAI",
                "tmp");
            Directory.CreateDirectory(directory);
            PurgeStaleTemporaryFiles(directory);

            string path = Path.Combine(directory, "draft-body-" + Guid.NewGuid().ToString("N") + ".htm");
            string document = Services.SignatureManager.EnsureHtmlDocument(html + (hasBoundary ? "\r\n<p>&nbsp;</p>" : string.Empty));
            File.WriteAllText(path, document, new UTF8Encoding(false));
            return path;
        }

        /// <summary>Best-effort cleanup of temp bodies a crashed run could have left behind.</summary>
        private static void PurgeStaleTemporaryFiles(string directory)
        {
            try
            {
                DateTime cutoff = DateTime.UtcNow.AddHours(-1);
                foreach (string stale in Directory.GetFiles(directory, "draft-body-*.htm"))
                {
                    if (File.GetLastWriteTimeUtc(stale) < cutoff)
                    {
                        File.Delete(stale);
                    }
                }
            }
            catch (IOException)
            {
            }
            catch (UnauthorizedAccessException)
            {
            }
        }

        private static void TryRestoreBookmark(object? bookmarksObject, string? name, int start, int end)
        {
            if (bookmarksObject == null || name == null || start < 0 || end < start)
            {
                return;
            }

            object? range = null;
            try
            {
                dynamic bm = (dynamic)bookmarksObject;
                if ((bool)bm.Exists(name))
                {
                    return;
                }

                dynamic parentDoc = bm.Parent;
                range = parentDoc.Range(start, end);
                bm.Add(name, range);
            }
            catch (Exception ex) when (IsComCallFailure(ex))
            {
            }
            finally
            {
                Release(range);
            }
        }

        private static int GetDocumentEnd(dynamic doc)
        {
            object? content = null;
            try
            {
                content = doc.Content;
                return (int)((dynamic)content!).End;
            }
            finally
            {
                Release(content);
            }
        }

        /// <summary>
        /// Test-support surface (live signature tests): applies a signature override to
        /// an EXISTING saved draft - re-opens it, runs the same held-inspector bookmark
        /// dance the creators use (this time against the PREVIOUSLY applied signature's
        /// bookmark, i.e. the replace branch), flushes via Close(olSave) on that same
        /// inspector, then saves the item.
        /// </summary>
        public bool TryApplySignatureOverrideToDraft(string entryIdHex, string? storeId, string signatureFilePath, out string? error)
        {
            EnsureNotDisposed();
            if (string.IsNullOrWhiteSpace(entryIdHex))
            {
                throw new ArgumentException("EntryID must not be blank.", nameof(entryIdHex));
            }

            string? capturedError = null;
            bool applied = _runner.Run(() =>
            {
                dynamic ns = _namespace!;
                object? itemObject = null;
                object? inspector = null;
                object? document = null;
                try
                {
                    itemObject = storeId != null
                        ? ns.GetItemFromID(entryIdHex, storeId)
                        : ns.GetItemFromID(entryIdHex);

                    if (!File.Exists(signatureFilePath))
                    {
                        capturedError = "SignatureFileMissing";
                        return false;
                    }

                    inspector = ((dynamic)itemObject!).GetInspector;
                    if (inspector == null)
                    {
                        capturedError = "NoInspector";
                        return false;
                    }

                    try
                    {
                        document = ((dynamic)inspector).WordEditor;
                    }
                    catch (Exception ex) when (IsComCallFailure(ex))
                    {
                        // Headless throws rather than returning null (D49).
                    }

                    if (document == null)
                    {
                        // D49: same invisible promotion as the two compose paths.
                        document = ComposeSurface.PromoteForWordEditor(inspector, out string? promoteError);
                        if (document == null)
                        {
                            capturedError = promoteError ?? "NoWordEditor";
                            return false;
                        }
                    }

                    (bool ok, string? overrideError) = ApplySignatureToDocument(document, signatureFilePath);
                    capturedError = overrideError;
                    if (!ok)
                    {
                        return false;
                    }

                    // Probe-proven flush order: Close(olSave) on the held inspector
                    // FIRST, then item.Save().
                    ((dynamic)inspector!).Close(0);
                    ((dynamic)itemObject!).Save();
                    return true;
                }
                catch (Exception ex) when (IsComCallFailure(ex))
                {
                    capturedError = DescribeComFailure(ex);
                    return false;
                }
                finally
                {
                    Release(document);
                    Release(inspector);
                    Release(itemObject);
                }
            });

            error = capturedError;
            return applied;
        }

        /// <summary>
        /// STA-side: closes the hidden Inspector the GetInspector signature touch left
        /// behind. Without this, a display:false draft still surfaces in
        /// Application.Inspectors. <paramref name="saveWordEdits"/> decides the close
        /// mode - PROBED on this machine (D37): Word-document edits (the signature
        /// override path) do NOT reach the item via <c>item.Save()</c>; only
        /// <c>Inspector.Close(olSave)</c> flushes them, while <c>olDiscard</c> throws
        /// them away permanently. Default path (no Word edits): olDiscard - the item is
        /// already saved, nothing is lost.
        /// </summary>
        private static void CloseHiddenInspector(object mailObject, bool saveWordEdits = false)
        {
            object? inspector = null;
            try
            {
                inspector = ((dynamic)mailObject).GetInspector;
                if (inspector != null)
                {
                    ((dynamic)inspector).Close(saveWordEdits ? 0 : 1); // 0 = olSave, 1 = olDiscard
                }
            }
            catch (Exception ex) when (IsComCallFailure(ex))
            {
                // No inspector to close - fine.
            }
            finally
            {
                Release(inspector);
            }
        }

        private static long CountNonWhitespaceText(string html)
        {
            if (html.Length == 0)
            {
                return 0;
            }

            string text = OutlookAI.Core.Text.HtmlToText.Convert(html);
            long count = 0;
            foreach (char c in text)
            {
                if (!char.IsWhiteSpace(c))
                {
                    count++;
                }
            }

            return count;
        }

        /// <summary>STA-side: adds typed recipients (1 = To, 2 = Cc) and resolves them best-effort.</summary>
        /// <summary>
        /// APPENDS recipients of one type to whatever is already on the item (a
        /// reply-all's own recipient list is never replaced) and records every address
        /// Outlook could not resolve into <paramref name="unresolved"/>. Unresolved
        /// recipients stay ON the draft - they are legal there and the user can fix them
        /// - but they are reported instead of silently dropped (batch A, A2).
        /// </summary>
        private static void AddRecipients(dynamic mail, IReadOnlyList<string> addresses, int type, ICollection<string>? unresolved = null)
        {
            if (addresses.Count == 0)
            {
                return;
            }

            object? recipients = null;
            try
            {
                recipients = mail.Recipients;
                dynamic collection = (dynamic)recipients!;
                foreach (string address in addresses)
                {
                    object? recipient = null;
                    try
                    {
                        recipient = collection.Add(address);
                        ((dynamic)recipient!).Type = type;

                        bool resolved;
                        try
                        {
                            // Per-recipient (not ResolveAll) so the verdict belongs to
                            // THIS address and never to a pre-existing one.
                            resolved = (bool)((dynamic)recipient).Resolve();
                        }
                        catch (Exception ex) when (IsComCallFailure(ex))
                        {
                            resolved = false;
                        }

                        if (!resolved)
                        {
                            unresolved?.Add(address);
                        }
                    }
                    finally
                    {
                        Release(recipient);
                    }
                }
            }
            finally
            {
                Release(recipients);
            }
        }

        /// <summary>
        /// STA-side: applies the optional cross-tool draft properties (batch A, A4).
        /// Both are plain item properties; a failure is swallowed content-free because
        /// the draft itself is still valid and the outcome reports the item's ACTUAL
        /// values read back in <see cref="SnapshotDraft"/>.
        /// </summary>
        private static void ApplyDraftOptions(dynamic draft, ComDraftOptions? options)
        {
            if (options == null)
            {
                return;
            }

            if (options.Importance != null)
            {
                try
                {
                    draft.Importance = options.Importance.Value;
                }
                catch (Exception ex) when (IsComCallFailure(ex))
                {
                }
            }

            if (options.RequestReadReceipt != null)
            {
                try
                {
                    draft.ReadReceiptRequested = options.RequestReadReceipt.Value;
                }
                catch (Exception ex) when (IsComCallFailure(ex))
                {
                }
            }
        }

        /// <summary>
        /// PropertyAccessor WRITE (MAPI proptag DASL). Mirrors
        /// <see cref="TryGetPropertyString"/>'s failure handling: a store/provider that
        /// refuses the write is reported as false, never thrown - callers degrade
        /// gracefully and report the fact.
        /// </summary>
        private static bool TrySetPropertyString(dynamic comObject, string schemaName, string value)
        {
            return TrySetProperty(comObject, schemaName, value);
        }

        /// <summary>PropertyAccessor write of a PT_BINARY MAPI property.</summary>
        private static bool TrySetPropertyBinary(dynamic comObject, string schemaName, byte[] value)
        {
            return TrySetProperty(comObject, schemaName, value);
        }

        private static bool TrySetProperty(dynamic comObject, string schemaName, object value)
        {
            object? accessor = null;
            try
            {
                accessor = comObject.PropertyAccessor;
                ((dynamic)accessor!).SetProperty(schemaName, value);
                return true;
            }
            catch (COMException)
            {
                return false;
            }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
            {
                return false;
            }
            finally
            {
                Release(accessor);
            }
        }

        /// <summary>Hex string (as the object model reports ConversationIndex) to bytes; null when unusable.</summary>
        private static byte[]? HexToBytes(string? hex)
        {
            if (string.IsNullOrEmpty(hex) || (hex!.Length % 2) != 0)
            {
                return null;
            }

            byte[] bytes = new byte[hex.Length / 2];
            for (int i = 0; i < bytes.Length; i++)
            {
                // Substring, not AsSpan: the span overload of byte.TryParse does not exist
                // on net48, and Core's net48 target is a CI gate (R10 / D18 v2).
                if (!byte.TryParse(
                        hex.Substring(i * 2, 2),
                        System.Globalization.NumberStyles.HexNumber,
                        CultureInfo.InvariantCulture,
                        out byte parsed))
                {
                    return null;
                }

                bytes[i] = parsed;
            }

            return bytes;
        }

        /// <summary>
        /// STA-side: when the saved item's parent folder differs from
        /// <paramref name="targetFolder"/>, moves it there (EntryIDs CHANGE on move -
        /// v3.MD section 12; callers snapshot AFTER this). Returns the item to use from
        /// now on (the moved RCW when a move happened) and releases the stale one.
        /// </summary>
        private static object RelocateToFolderIfNeeded(object mailObject, object targetFolder, out bool moved, out string? initialFolderName, out bool inTargetFolder)
        {
            moved = false;
            initialFolderName = null;
            inTargetFolder = false;
            object? parent = null;
            try
            {
                parent = ((dynamic)mailObject).Parent;
                string? parentEntryId = null;
                if (parent != null)
                {
                    initialFolderName = TryGetString(() => (string?)((dynamic)parent).Name);
                    parentEntryId = TryGetString(() => (string?)((dynamic)parent).EntryID);
                }

                string? targetEntryId = TryGetString(() => (string?)((dynamic)targetFolder).EntryID);
                if (parentEntryId != null && targetEntryId != null
                    && string.Equals(parentEntryId, targetEntryId, StringComparison.OrdinalIgnoreCase))
                {
                    inTargetFolder = true;
                    return mailObject;
                }

                object movedItem = ((dynamic)mailObject).Move(targetFolder);
                Release(mailObject);
                moved = true;
                inTargetFolder = true;
                return movedItem;
            }
            catch (Exception ex) when (IsComCallFailure(ex))
            {
                // Move unavailable - keep the item where Outlook saved it (recorded via
                // the initial folder name); placement in the target stays unconfirmed.
                return mailObject;
            }
            finally
            {
                Release(parent);
            }
        }

        /// <summary>Delay before the one-shot recipient re-snapshot of a fresh draft (degraded-instance window, soak 2026-07-24).</summary>
        private const int RecipientResnapshotDelayMs = 1500;

        /// <summary>
        /// Creation-flow guard (soak 2026-07-24): every draft-creation flow carries at
        /// least one recipient by construction (reply/replyall derive them from the
        /// source, new/forward have a validated To) - a creation snapshot reading ZERO
        /// recipients is the degraded read shape observed while Outlook is booting or
        /// reconciling its stores (the item header answers while object-returning
        /// probes come back empty; the saved draft itself is correct). Remedy, bounded
        /// to one attempt: wait briefly, re-open the item FRESH by EntryID (a new COM
        /// proxy, not the possibly-degraded creation reference) and re-snapshot; the
        /// original snapshot is kept when the retry does not improve on it.
        /// </summary>
        private ComDraftInfo ResnapshotIfRecipientsEmpty(
            ComDraftInfo info,
            string? fallbackStoreName,
            string? fallbackStoreId,
            string? fallbackFolderName,
            string? fallbackFolderEntryId,
            string? fallbackSendUsingSmtp)
        {
            if (info.Recipients.Count > 0)
            {
                return info;
            }

            Thread.Sleep(RecipientResnapshotDelayMs);
            dynamic ns = _namespace!;
            object? reopened = null;
            try
            {
                reopened = info.StoreId != null
                    ? ns.GetItemFromID(info.EntryId, info.StoreId)
                    : ns.GetItemFromID(info.EntryId);
                ComDraftInfo retry = SnapshotDraft(
                    reopened!, fallbackStoreName, fallbackStoreId, fallbackFolderName, fallbackFolderEntryId, fallbackSendUsingSmtp);
                return retry.Recipients.Count > 0 ? retry : info;
            }
            catch (Exception ex) when (IsComCallFailure(ex))
            {
                return info;
            }
            finally
            {
                Release(reopened);
            }
        }

        /// <summary>Pumped wait between re-reads of an attachment whose size still reads zero (soak fix 21).</summary>
        private const int AttachmentSizeSettleDelayMs = 250;

        /// <summary>How many times a zero-sized attachment snapshot is re-read before it is reported as-is.</summary>
        private const int AttachmentSizeSettleAttempts = 4;

        /// <summary>
        /// Attachment snapshot for a draft RESULT, taken from a FRESH item reference.
        /// <para>
        /// Soak fix 21, and it was the defect a field report filed as "the draft picked up
        /// a ZERO-BYTE image001.png from the signature template". No bytes were ever lost:
        /// the compose reference these flows hold - the item whose hidden Inspector was
        /// just closed - answers <c>Attachment.Size</c> with ZERO for an attachment Outlook
        /// materialized during the composition, and in the HTMLBody-fallback shape answers
        /// <c>Attachments.Count</c> with zero as well. Measured 8/8 on a real account
        /// signature: <c>new_draft</c> reported <c>image001.png = 0</c> while <c>read</c>
        /// reported 3 035 on the SAME item milliseconds later, over 2 834 real PNG bytes.
        /// The tool was telling the calling agent its signature logo was empty.
        /// </para>
        /// <para>
        /// Remedy, and the ORDERING is the load-bearing part: re-opening the item by
        /// EntryID inside the composing call does NOT help - measured, including with
        /// bounded PUMPED waits, it still reads zero, because the compose flow's own item
        /// reference is still alive at that point. The size is committed by the time the
        /// NEXT call arrives, so this runs as its own COM call, after the creation call has
        /// returned and released everything (that is exactly why <c>read</c> always
        /// reported the truth while <c>new_draft</c> did not). The caller keeps whichever
        /// snapshot is better under <see cref="AttachmentSnapshotMerge"/>'s monotone rule,
        /// so a re-read can only ever improve a result.
        /// </para>
        /// </summary>
        public IReadOnlyList<ComAttachmentInfo> SnapshotAttachmentsById(string entryId, string? storeId)
        {
            IReadOnlyList<ComAttachmentInfo> best = Array.Empty<ComAttachmentInfo>();
            if (string.IsNullOrEmpty(entryId))
            {
                return best;
            }

            for (int attempt = 0; attempt < AttachmentSizeSettleAttempts; attempt++)
            {
                if (attempt > 0)
                {
                    // The retry exists ONLY for the not-yet-settled size; a snapshot that
                    // already knows every byte must not cost a wait. And the wait must be
                    // PUMPED: an unpumped Thread.Sleep on this STA blocks the very message
                    // queue Outlook needs to finish committing the attachment, so the size
                    // it waits for can never arrive (measured: 0 for the whole of a 400 ms
                    // unpumped sleep, correct on the next pumped call).
                    if (!AttachmentSnapshotMerge.HasUnsizedAttachment(best))
                    {
                        break;
                    }

                    PumpedStaRunner.PumpedWait(AttachmentSizeSettleDelayMs);
                }

                object? reopened = null;
                try
                {
                    dynamic ns = _namespace!;
                    reopened = string.IsNullOrEmpty(storeId)
                        ? ns.GetItemFromID(entryId)
                        : ns.GetItemFromID(entryId, storeId);
                    IReadOnlyList<ComAttachmentInfo> fresh = SnapshotAttachments(reopened!);
                    if (AttachmentSnapshotMerge.IsBetter(fresh, best))
                    {
                        best = fresh;
                    }
                }
                catch (Exception ex) when (IsComCallFailure(ex))
                {
                    break;
                }
                finally
                {
                    Release(reopened);
                }

                if (!AttachmentSnapshotMerge.HasUnsizedAttachment(best))
                {
                    break;
                }
            }

            return best;
        }

        /// <summary>STA-side identity/threading snapshot of a mail item.</summary>
        private ComDraftInfo SnapshotDraft(object itemObject, string? fallbackStoreName = null, string? fallbackStoreId = null, string? fallbackFolderName = null, string? fallbackFolderEntryId = null, string? fallbackSendUsingSmtp = null)
        {
            dynamic item = itemObject;
            string entryId = (string)item.EntryID;
            string? subject = TryGetString(() => (string?)item.Subject);
            string? conversationIndex = TryGetString(() => (string?)item.ConversationIndex);
            string? conversationId = TryGetString(() => (string?)item.ConversationID);
            string? conversationTopic = TryGetString(() => (string?)item.ConversationTopic)
                ?? TryGetPropertyString(item, ConversationTopicDasl);
            int? importance = null;
            bool readReceiptRequested = false;
            try
            {
                importance = (int)item.Importance;
            }
            catch (Exception ex) when (IsComCallFailure(ex))
            {
            }

            try
            {
                readReceiptRequested = (bool)item.ReadReceiptRequested;
            }
            catch (Exception ex) when (IsComCallFailure(ex))
            {
            }

            string? sendUsingSmtp = null;
            object? sendUsingAccount = null;
            try
            {
                sendUsingAccount = item.SendUsingAccount;
                if (sendUsingAccount != null)
                {
                    sendUsingSmtp = TryGetString(() => (string?)((dynamic)sendUsingAccount!).SmtpAddress);
                }
            }
            catch (Exception ex) when (IsComCallFailure(ex))
            {
            }
            finally
            {
                Release(sendUsingAccount);
            }

            string? folderName = null;
            string? folderEntryId = null;
            string? storeName = null;
            string? storeId = null;
            object? parent = null;
            object? parentStore = null;
            try
            {
                parent = item.Parent;
                if (parent != null)
                {
                    folderName = TryGetString(() => (string?)((dynamic)parent).Name);
                    folderEntryId = TryGetString(() => (string?)((dynamic)parent).EntryID);
                    parentStore = ((dynamic)parent).Store;
                    if (parentStore != null)
                    {
                        storeName = TryGetString(() => (string?)((dynamic)parentStore).DisplayName);
                        storeId = TryGetString(() => (string?)((dynamic)parentStore).StoreID);
                    }
                }
            }
            catch (Exception ex) when (IsComCallFailure(ex))
            {
            }
            finally
            {
                Release(parentStore);
                Release(parent);
            }

            // Fresh/busy-Outlook robustness (soak-fix batch): the Parent/Store probe
            // above is best-effort and can transiently fail right after a cold start or
            // while the UI is busy (live-observed: the relocate step confirmed the item
            // in Drafts and milliseconds later this Parent probe answered null) - fall
            // back to the caller-known store AND folder identity so DraftOutcome.Store/
            // .Folder stay deterministic. Folder fallbacks are passed only when the
            // caller CONFIRMED placement in that folder (RelocateToFolderIfNeeded).
            storeName ??= fallbackStoreName;
            storeId ??= fallbackStoreId;
            folderName ??= fallbackFolderName;
            folderEntryId ??= fallbackFolderEntryId;
            sendUsingSmtp ??= fallbackSendUsingSmtp; // Only passed when the caller PINNED the identity.

            List<ComRecipientInfo> recipients = new List<ComRecipientInfo>();
            object? recipientsObject = null;
            try
            {
                recipientsObject = item.Recipients;
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
                        catch (Exception ex) when (IsComCallFailure(ex))
                        {
                        }

                        string kind = type == 2 ? "cc" : type == 3 ? "bcc" : "to";
                        string? name = TryGetString(() => (string?)r.Name);
                        string? address = TryGetPropertyString(r, "http://schemas.microsoft.com/mapi/proptag/0x39FE001F")
                            ?? TryGetString(() => (string?)r.Address);
                        recipients.Add(new ComRecipientInfo(kind, name, address));
                    }
                    finally
                    {
                        Release(recipient);
                    }
                }
            }
            catch (Exception ex) when (IsComCallFailure(ex))
            {
            }
            finally
            {
                Release(recipientsObject);
            }

            return new ComDraftInfo(
                entryId,
                storeName,
                storeId,
                folderName,
                folderEntryId,
                subject,
                sendUsingSmtp,
                conversationIndex,
                conversationId,
                recipients,
                conversationTopic,
                importance,
                readReceiptRequested);
        }

        /// <summary>STA-side: the profile account with the given SmtpAddress (caller releases), or null.</summary>
        private object? FindAccountBySmtp(string smtpAddress)
        {
            dynamic ns = _namespace!;
            object? session = null;
            object? accounts = null;
            try
            {
                session = ns.Session;
                accounts = ((dynamic)session!).Accounts;
                dynamic collection = (dynamic)accounts!;
                int count = collection.Count;
                for (int i = 1; i <= count; i++)
                {
                    object? account = collection[i];
                    string? smtp = TryGetString(() => (string?)((dynamic)account!).SmtpAddress);
                    if (smtp != null && string.Equals(smtp, smtpAddress, StringComparison.OrdinalIgnoreCase))
                    {
                        return account;
                    }

                    Release(account);
                }

                return null;
            }
            finally
            {
                Release(accounts);
                Release(session);
            }
        }

        /// <summary>
        /// STA-side: the account whose DeliveryStore matches the given StoreID or (as a
        /// robustness fallback - store EntryID wrappings can differ between retrieval
        /// paths) the given store display name. Caller releases; null when no match.
        /// </summary>
        private object? FindAccountByDeliveryStore(string? storeId, string? storeDisplayName)
        {
            dynamic ns = _namespace!;
            object? session = null;
            object? accounts = null;
            try
            {
                session = ns.Session;
                accounts = ((dynamic)session!).Accounts;
                dynamic collection = (dynamic)accounts!;
                int count = collection.Count;
                for (int i = 1; i <= count; i++)
                {
                    object? account = collection[i];
                    object? deliveryStore = null;
                    try
                    {
                        deliveryStore = ((dynamic)account!).DeliveryStore;
                        if (deliveryStore != null)
                        {
                            string? deliveryStoreId = TryGetString(() => (string?)((dynamic)deliveryStore!).StoreID);
                            string? deliveryStoreName = TryGetString(() => (string?)((dynamic)deliveryStore!).DisplayName);
                            bool idMatch = storeId != null && deliveryStoreId != null
                                && string.Equals(deliveryStoreId, storeId, StringComparison.OrdinalIgnoreCase);
                            bool nameMatch = storeDisplayName != null && deliveryStoreName != null
                                && string.Equals(deliveryStoreName, storeDisplayName, StringComparison.OrdinalIgnoreCase);
                            if (idMatch || nameMatch)
                            {
                                return account;
                            }
                        }
                    }
                    catch (Exception ex) when (IsComCallFailure(ex))
                    {
                    }
                    finally
                    {
                        Release(deliveryStore);
                    }

                    Release(account);
                }

                return null;
            }
            finally
            {
                Release(accounts);
                Release(session);
            }
        }

        // ------------------------------------------------------------------ exhaustive scan (Phase 3, v3.MD D19)

        /// <summary>
        /// Exhaustive folder/date-bounded COM scan (search exhaustive:true): filters each
        /// mail folder in scope with a DASL restriction via Folder.GetTable -
        /// ci_phrasematch when Store.IsInstantSearchEnabled (feature-detected, per-folder
        /// LIKE downgrade on failure; v3.MD section 12: ci_* is valid in Restrict/GetTable
        /// only), plain LIKE otherwise. Matches are opened for authoritative snapshots
        /// carrying REAL EntryIDs. Scope = one folder (path given) or every folder of the
        /// store. Bounded by <paramref name="maxItems"/> and <paramref name="timeBudgetMs"/>.
        /// <paramref name="searchIn"/> selects which properties the terms must appear in
        /// (subject and/or body) - the same three scopes the index tier offers (D40).
        /// <paramref name="includeSubfolders"/> decides whether a folder-scoped scan walks
        /// the subtree; without a folder path the whole store tree is always walked.
        /// </summary>
        public ComExhaustiveResult ExhaustiveScan(
            string storeDisplayName,
            IReadOnlyList<string>? folderPath,
            IReadOnlyList<string>? terms,
            DateTime? sinceUtc,
            DateTime? beforeUtc,
            int maxItems,
            int timeBudgetMs,
            SearchIn searchIn = SearchInValues.Default,
            bool includeSubfolders = false)
        {
            EnsureNotDisposed();
            if (string.IsNullOrWhiteSpace(storeDisplayName))
            {
                throw new ArgumentException("Store display name must not be blank.", nameof(storeDisplayName));
            }

            if (maxItems < 1)
            {
                throw new ArgumentOutOfRangeException(nameof(maxItems));
            }

            if (timeBudgetMs < 1)
            {
                throw new ArgumentOutOfRangeException(nameof(timeBudgetMs));
            }

            return _runner.Run(() =>
            {
                dynamic? store = FindStoreByDisplayName(storeDisplayName);
                if (store == null)
                {
                    throw new InvalidOperationException(
                        "Store '" + storeDisplayName + "' was not found in Outlook. Use list_accounts for store names.");
                }

                string storeId;
                bool instantSearch = false;
                object? scanRoot = null;
                try
                {
                    storeId = (string)store.StoreID;
                    try
                    {
                        instantSearch = (bool)store.IsInstantSearchEnabled;
                    }
                    catch (Exception ex) when (IsComCallFailure(ex))
                    {
                        // Property unavailable - treat as no instant search (LIKE engine).
                    }

                    if (folderPath == null || folderPath.Count == 0)
                    {
                        scanRoot = store.GetRootFolder();
                    }
                }
                finally
                {
                    Release(store);
                }

                if (scanRoot == null)
                {
                    scanRoot = WalkToFolder(storeDisplayName, folderPath!, out string? walkError);
                    if (scanRoot == null)
                    {
                        throw new InvalidOperationException(
                            "Folder '" + string.Join("/", folderPath!) + "' was not found in store '" + storeDisplayName
                            + "' (" + (walkError ?? "unknown") + "). Use list_folders for paths.");
                    }
                }

                ExhaustiveScanState state = new ExhaustiveScanState(maxItems, TimeSpan.FromMilliseconds(timeBudgetMs))
                {
                    CiFilter = instantSearch
                        ? ExhaustiveDaslFilter.Build(terms, sinceUtc, beforeUtc, ExhaustiveEngine.CiPhraseMatch, searchIn)
                        : null,
                    LikeFilter = ExhaustiveDaslFilter.Build(terms, sinceUtc, beforeUtc, ExhaustiveEngine.Like, searchIn),
                };

                try
                {
                    // A whole-store scan always recurses; a folder-scoped one follows the
                    // caller's include_subfolders flag (soak fix 15 - before it, a
                    // folder-scoped exhaustive scan was unconditionally shallow while the
                    // index tier's folder scope was recursive).
                    bool recurse = folderPath == null || folderPath.Count == 0 || includeSubfolders;
                    ScanFolderTree(_namespace!, scanRoot, storeDisplayName, storeId, recurse, state);
                }
                finally
                {
                    Release(scanRoot);
                }

                string engine = state.UsedCi && state.UsedLike
                    ? "ci_phrasematch+like"
                    : state.UsedCi ? "ci_phrasematch" : "like";
                return new ComExhaustiveResult(
                    state.Items,
                    state.FoldersScanned,
                    state.FoldersSkipped,
                    engine,
                    instantSearch,
                    state.Truncated,
                    state.TimedOut);
            });
        }

        private void ScanFolderTree(
            dynamic ns,
            object folderObject,
            string storeName,
            string storeId,
            bool recurse,
            ExhaustiveScanState state)
        {
            // The deadline is evaluated PER FOLDER, not only inside the per-row drain:
            // a folder whose filter matches zero rows never entered the drain loop, so a
            // wide low-yield subtree used to overrun the budget without bound while
            // timedOut stayed false (soak fix 15).
            if (state.CheckDeadline() || state.ShouldStop)
            {
                return;
            }

            dynamic folder = folderObject;

            // Only mail folders are filtered (DefaultItemType 0 = olMailItem); other
            // folder types still get their subtrees visited (a calendar folder can hold
            // mail subfolders).
            int defaultItemType = -1;
            try
            {
                defaultItemType = (int)folder.DefaultItemType;
            }
            catch (Exception ex) when (IsComCallFailure(ex))
            {
            }

            if (defaultItemType == 0)
            {
                ScanSingleFolder(ns, folderObject, storeName, storeId, state);
            }

            if (!recurse || state.ShouldStop)
            {
                return;
            }

            object? subFolders = null;
            try
            {
                subFolders = folder.Folders;
                dynamic folderCollection = (dynamic)subFolders!;
                int count = folderCollection.Count;
                for (int i = 1; i <= count && !state.ShouldStop; i++)
                {
                    object? child = null;
                    try
                    {
                        child = folderCollection[i];
                        ScanFolderTree(ns, child!, storeName, storeId, true, state);
                    }
                    finally
                    {
                        Release(child);
                    }
                }
            }
            catch (Exception ex) when (IsComCallFailure(ex))
            {
                // A subtree that cannot be enumerated is a coverage hole: count it, so
                // foldersSkipped stops silently under-reporting (soak fix 15).
                state.FoldersSkipped++;
            }
            finally
            {
                Release(subFolders);
            }
        }

        private void ScanSingleFolder(dynamic ns, object folderObject, string storeName, string storeId, ExhaustiveScanState state)
        {
            dynamic folder = folderObject;
            string? folderName = TryGetString(() => (string?)folder.Name);

            string filter = state.CiFilter != null && !state.CiBroken ? state.CiFilter : state.LikeFilter;
            bool triedCi = ReferenceEquals(filter, state.CiFilter);

            object? table = null;
            try
            {
                try
                {
                    table = folder.GetTable(filter);
                }
                catch (Exception ex) when (triedCi && IsComCallFailure(ex))
                {
                    // ci_phrasematch rejected here - downgrade this and all later folders
                    // to LIKE (feature-detect rule, v3.MD section 12).
                    state.CiBroken = true;
                    triedCi = false;
                    filter = state.LikeFilter;
                    table = folder.GetTable(filter);
                }

                if (triedCi)
                {
                    state.UsedCi = true;
                }
                else
                {
                    state.UsedLike = true;
                }

                dynamic t = (dynamic)table!;
                int entryIdIndex = FindTableColumn(t, "EntryID");
                if (entryIdIndex < 0)
                {
                    state.FoldersSkipped++;
                    return;
                }

                state.FoldersScanned++;
                while (!(bool)t.EndOfTable && !state.ShouldStop)
                {
                    if (state.CheckDeadline())
                    {
                        return;
                    }

                    object? row = null;
                    object? member = null;
                    try
                    {
                        row = t.GetNextRow();
                        object[] values = (object[])((dynamic)row!).GetValues();
                        if (entryIdIndex >= values.Length || values[entryIdIndex] is not string entryId || entryId.Length == 0)
                        {
                            continue;
                        }

                        try
                        {
                            member = ns.GetItemFromID(entryId, storeId);
                        }
                        catch (Exception ex) when (IsComCallFailure(ex))
                        {
                            continue;
                        }

                        int itemClass;
                        try
                        {
                            itemClass = (int)((dynamic)member!).Class;
                        }
                        catch (Exception ex) when (IsComCallFailure(ex))
                        {
                            continue;
                        }

                        if (itemClass != OlMailItemClass)
                        {
                            continue;
                        }

                        state.Items.Add(SnapshotBrief(ns, member!, null, folderName, false, storeName, storeId));
                        if (state.Items.Count >= state.MaxItems)
                        {
                            state.Truncated = true;
                        }
                    }
                    finally
                    {
                        Release(member);
                        Release(row);
                    }
                }
            }
            catch (Exception ex) when (IsComCallFailure(ex))
            {
                state.FoldersSkipped++;
            }
            finally
            {
                Release(table);
            }
        }

        private sealed class ExhaustiveScanState
        {
            internal ExhaustiveScanState(int maxItems, TimeSpan budget)
            {
                MaxItems = maxItems;
                Budget = budget;
                Clock = Stopwatch.StartNew();
            }

            internal List<ComMailBrief> Items { get; } = new List<ComMailBrief>();

            internal int MaxItems { get; }

            internal TimeSpan Budget { get; }

            internal Stopwatch Clock { get; }

            internal string? CiFilter { get; set; }

            internal string LikeFilter { get; set; } = string.Empty;

            internal int FoldersScanned { get; set; }

            internal int FoldersSkipped { get; set; }

            internal bool UsedCi { get; set; }

            internal bool UsedLike { get; set; }

            internal bool CiBroken { get; set; }

            internal bool Truncated { get; set; }

            internal bool TimedOut { get; set; }

            internal bool ShouldStop => Truncated || TimedOut;

            /// <summary>
            /// Latches <see cref="TimedOut"/> once the budget is spent. Called at every
            /// FOLDER boundary as well as per row - a zero-match folder never reaches the
            /// row loop, so folder-level checking is what actually bounds a wide subtree.
            /// </summary>
            internal bool CheckDeadline()
            {
                if (TimedOut)
                {
                    return true;
                }

                if (Clock.Elapsed <= Budget)
                {
                    return false;
                }

                TimedOut = true;
                return true;
            }
        }

        /// <summary>
        /// Sweeps ONE folder's window. Returns how it ended so the caller can report
        /// partial coverage: before soak fix 15 a COM failure here was swallowed and the
        /// folder was still counted as successfully swept, and the per-folder item cap was
        /// wholly invisible (itemsSeen is post-cap, so "200" was indistinguishable from
        /// "exactly 200 existed" - and the table is sorted newest-first, so the OLDEST
        /// items in the freshness window were the ones silently dropped).
        /// </summary>
        private SweepOutcome SweepFolder(
            dynamic ns,
            object folderObject,
            string storeName,
            string storeId,
            string? folderKind,
            DateTime sinceUtc,
            int cap,
            bool includeBodies,
            List<ComMailBrief> results)
        {
            dynamic folder = folderObject;
            string? folderName = TryGetString(() => (string?)folder.Name);
            string filter = "@SQL=(\"urn:schemas:httpmail:datereceived\" >= '" + FormatDaslUtc(sinceUtc)
                + "') OR (\"urn:schemas:httpmail:date\" >= '" + FormatDaslUtc(sinceUtc) + "')";

            object? table = null;
            try
            {
                table = folder.GetTable(filter);
                dynamic t = (dynamic)table!;
                try
                {
                    // Sort needs the property present as a column; late-bound COM maps
                    // E_INVALIDARG to ArgumentException, hence the broad catch.
                    object? columns = null;
                    try
                    {
                        columns = t.Columns;
                        ((dynamic)columns!).Add("urn:schemas:httpmail:datereceived");
                    }
                    finally
                    {
                        Release(columns);
                    }

                    t.Sort("urn:schemas:httpmail:datereceived", true);
                }
                catch (Exception ex) when (IsComCallFailure(ex))
                {
                    // Unsorted sweep still works; the cap just cuts arbitrarily.
                }

                int entryIdIndex = FindTableColumn(t, "EntryID");
                if (entryIdIndex < 0)
                {
                    return SweepOutcome.Failed;
                }

                int taken = 0;
                while (!(bool)t.EndOfTable && taken < cap)
                {
                    object? row = null;
                    object? member = null;
                    try
                    {
                        row = t.GetNextRow();
                        object[] values = (object[])((dynamic)row!).GetValues();
                        if (entryIdIndex >= values.Length || values[entryIdIndex] is not string entryId || entryId.Length == 0)
                        {
                            continue;
                        }

                        try
                        {
                            member = ns.GetItemFromID(entryId, storeId);
                        }
                        catch (Exception ex) when (IsComCallFailure(ex))
                        {
                            continue;
                        }

                        results.Add(SnapshotBrief(ns, member!, folderKind, folderName, includeBodies, storeName, storeId));
                        taken++;
                    }
                    finally
                    {
                        Release(member);
                        Release(row);
                    }
                }

                // The cap is only a truncation when the table still had rows to give.
                return taken >= cap && !(bool)t.EndOfTable ? SweepOutcome.ItemCapped : SweepOutcome.Complete;
            }
            catch (Exception ex) when (IsComCallFailure(ex))
            {
                // GetTable/filter unsupported on this folder: NO freshness coverage here.
                return SweepOutcome.Failed;
            }
            finally
            {
                Release(table);
            }
        }

        private ComMailBrief SnapshotBrief(
            dynamic ns,
            object itemObject,
            string? folderKind,
            string? folderName,
            bool includeBody,
            string? storeNameHint = null,
            string? storeIdHint = null)
        {
            dynamic item = itemObject;
            string entryId = (string)item.EntryID;
            string? subject = TryGetString(() => (string?)item.Subject);
            DateTime? received = TryGetDateTime(() => (DateTime)item.ReceivedTime);
            string? senderName = TryGetString(() => (string?)item.SenderName);
            string? senderAddress = TryGetSenderSmtp(item);
            bool? isRead = null;
            try
            {
                isRead = !(bool)item.UnRead;
            }
            catch (COMException)
            {
            }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
            {
            }

            long? size = null;
            try
            {
                size = (long)(int)item.Size;
            }
            catch (COMException)
            {
            }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
            {
            }

            bool? hasAttachments = null;
            object? attachments = null;
            try
            {
                attachments = item.Attachments;
                hasAttachments = (int)((dynamic)attachments!).Count > 0;
            }
            catch (COMException)
            {
            }
            finally
            {
                Release(attachments);
            }

            string? storeName = storeNameHint;
            string? storeId = storeIdHint;
            string? resolvedFolderName = folderName;
            if (storeName == null || resolvedFolderName == null)
            {
                object? parent = null;
                object? parentStore = null;
                try
                {
                    parent = item.Parent;
                    if (parent != null)
                    {
                        resolvedFolderName ??= TryGetString(() => (string?)((dynamic)parent).Name);
                        if (storeName == null)
                        {
                            parentStore = ((dynamic)parent).Store;
                            if (parentStore != null)
                            {
                                storeName = TryGetString(() => (string?)((dynamic)parentStore).DisplayName);
                                storeId ??= TryGetString(() => (string?)((dynamic)parentStore).StoreID);
                            }
                        }
                    }
                }
                catch (COMException)
                {
                }
                finally
                {
                    Release(parentStore);
                    Release(parent);
                }
            }

            string? body = null;
            if (includeBody)
            {
                body = TryGetString(() => (string?)item.Body);
            }

            return new ComMailBrief(
                entryId,
                storeName ?? string.Empty,
                storeId,
                resolvedFolderName,
                folderKind,
                subject,
                senderName,
                senderAddress,
                received,
                isRead,
                hasAttachments,
                size,
                body);
        }

        private ComItemDetail SnapshotDetail(object itemObject, bool includeHeaders, bool includeBody, bool includeHtml = false)
        {
            dynamic item = itemObject;
            string entryId = (string)item.EntryID;
            int? itemClass = null;
            try
            {
                itemClass = (int)item.Class;
            }
            catch (COMException)
            {
            }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
            {
            }

            string? subject = TryGetString(() => (string?)item.Subject);
            DateTime? received = TryGetDateTime(() => (DateTime)item.ReceivedTime);
            DateTime? sent = TryGetDateTime(() => (DateTime)item.SentOn);
            string? senderName = TryGetString(() => (string?)item.SenderName);
            string? senderAddress = TryGetSenderSmtp(item);
            string? conversationId = TryGetString(() => (string?)item.ConversationID);

            bool? isRead = null;
            try
            {
                isRead = !(bool)item.UnRead;
            }
            catch (COMException)
            {
            }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
            {
            }

            long? size = null;
            try
            {
                size = (long)(int)item.Size;
            }
            catch (COMException)
            {
            }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
            {
            }

            // Body: Outlook maintains .Body as the plain-text rendering of HTML/RTF
            // mail; fall back to converting .HTMLBody ourselves when it is empty. The
            // FULL body is returned (windowing + caching live in MailService, D37);
            // includeBody=false skips the transfer for cache-served continuation reads.
            string body = string.Empty;
            string bodyOrigin = "none";
            if (includeBody)
            {
                string? nativeBody = TryGetString(() => (string?)item.Body);
                if (!string.IsNullOrEmpty(nativeBody))
                {
                    body = nativeBody!;
                    bodyOrigin = "text";
                }
                else
                {
                    string? html = TryGetString(() => (string?)item.HTMLBody);
                    if (!string.IsNullOrEmpty(html))
                    {
                        body = OutlookAI.Core.Text.HtmlToText.Convert(html);
                        bodyOrigin = "html-converted";
                    }
                }
            }

            long bodyTotal = body.Length;

            // The raw markup, only on request (read include_html): the plain-text
            // rendering above collapses exactly the structure an agent needs to verify -
            // formatting, where the signature region starts, where the quote begins.
            string? htmlBody = includeHtml ? (TryGetString(() => (string?)item.HTMLBody) ?? string.Empty) : null;

            // Recipients with SMTP resolution (PR_SMTP_ADDRESS via PropertyAccessor).
            List<ComRecipientInfo> recipients = new List<ComRecipientInfo>();
            object? recipientsObject = null;
            try
            {
                recipientsObject = item.Recipients;
                dynamic recipientCollection = (dynamic)recipientsObject!;
                int recipientCount = recipientCollection.Count;
                for (int i = 1; i <= recipientCount; i++)
                {
                    object? recipient = null;
                    try
                    {
                        recipient = recipientCollection[i];
                        dynamic r = (dynamic)recipient!;
                        int type = 1;
                        try
                        {
                            type = (int)r.Type;
                        }
                        catch (COMException)
                        {
                        }

                        string kind = type == 2 ? "cc" : type == 3 ? "bcc" : "to";
                        string? name = TryGetString(() => (string?)r.Name);
                        string? address = TryGetPropertyString(r, "http://schemas.microsoft.com/mapi/proptag/0x39FE001F")
                            ?? TryGetString(() => (string?)r.Address);
                        recipients.Add(new ComRecipientInfo(kind, name, address));
                    }
                    finally
                    {
                        Release(recipient);
                    }
                }
            }
            catch (COMException)
            {
            }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
            {
            }
            finally
            {
                Release(recipientsObject);
            }

            // Attachments.
            List<ComAttachmentInfo> attachmentInfos = new List<ComAttachmentInfo>();
            object? attachmentsObject = null;
            try
            {
                attachmentsObject = item.Attachments;
                dynamic attachmentCollection = (dynamic)attachmentsObject!;
                int attachmentCount = attachmentCollection.Count;
                for (int i = 1; i <= attachmentCount; i++)
                {
                    object? attachment = null;
                    try
                    {
                        attachment = attachmentCollection[i];
                        dynamic a = (dynamic)attachment!;
                        string? fileName = TryGetString(() => (string?)a.FileName);
                        long? attachmentSize = null;
                        try
                        {
                            attachmentSize = (long)(int)a.Size;
                        }
                        catch (COMException)
                        {
                        }

                        attachmentInfos.Add(new ComAttachmentInfo(i, fileName, attachmentSize));
                    }
                    finally
                    {
                        Release(attachment);
                    }
                }
            }
            catch (COMException)
            {
            }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
            {
            }
            finally
            {
                Release(attachmentsObject);
            }

            // Folder + store.
            string? folderPath = null;
            string? storeName = null;
            object? parent = null;
            object? parentStore = null;
            try
            {
                parent = item.Parent;
                if (parent != null)
                {
                    folderPath = TryGetString(() => (string?)((dynamic)parent).FolderPath);
                    parentStore = ((dynamic)parent).Store;
                    if (parentStore != null)
                    {
                        storeName = TryGetString(() => (string?)((dynamic)parentStore).DisplayName);
                    }
                }
            }
            catch (COMException)
            {
            }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
            {
            }
            finally
            {
                Release(parentStore);
                Release(parent);
            }

            string? internetMessageId = TryGetPropertyString(item, "http://schemas.microsoft.com/mapi/proptag/0x1035001F");
            string? headers = includeHeaders
                ? TryGetPropertyString(item, "http://schemas.microsoft.com/mapi/proptag/0x007D001F")
                : null;

            return new ComItemDetail(
                entryId,
                storeName,
                folderPath,
                itemClass,
                subject,
                senderName,
                senderAddress,
                received,
                sent,
                recipients,
                body,
                bodyTotal,
                bodyOrigin,
                attachmentInfos,
                size,
                isRead,
                conversationId,
                internetMessageId,
                headers,
                htmlBody);
        }

        /// <summary>
        /// Store-relative folder PATHS of one store, names only - no PR_CONTENT_COUNT /
        /// PR_CONTENT_UNREAD reads, no child counts, no sorting. Deliberately much cheaper
        /// than <see cref="ListFolders"/>, because the index tier calls it to learn a
        /// DELEGATE store's real nesting: the delegate index namespace is FLAT, so mapping
        /// a requested subtree onto the leaf names it contains (and spotting leaf-name
        /// collisions) needs the COM tree, and only Outlook has it.
        /// </summary>
        public IReadOnlyList<string> ListFolderPaths(string storeDisplayName, int absoluteWalkCap)
        {
            EnsureNotDisposed();
            if (string.IsNullOrWhiteSpace(storeDisplayName))
            {
                throw new ArgumentException("Store display name must not be blank.", nameof(storeDisplayName));
            }

            if (absoluteWalkCap < 1)
            {
                throw new ArgumentOutOfRangeException(nameof(absoluteWalkCap));
            }

            return _runner.Run(() =>
            {
                List<string> result = new List<string>();
                dynamic? store = FindStoreByDisplayName(storeDisplayName);
                if (store == null)
                {
                    return (IReadOnlyList<string>)result;
                }

                object? root = null;
                try
                {
                    root = store.GetRootFolder();
                    CollectFolderPaths(root!, string.Empty, 1, absoluteWalkCap, result);
                }
                catch (Exception ex) when (IsComCallFailure(ex))
                {
                }
                finally
                {
                    Release(root);
                    Release((object?)store);
                }

                return (IReadOnlyList<string>)result;
            });
        }

        /// <summary>
        /// Store-relative folder paths, as SEGMENTS, whose LEAF name equals
        /// <paramref name="leafName"/>.
        /// <para>
        /// Exists because the delegate index namespace is FLAT (D42): a delegate item in
        /// <c>Archive/AliExpress</c> is indexed as <c>&lt;host&gt;/1/&lt;delegate&gt;/AliExpress</c>,
        /// so walking that path from the delegate store root hits nothing and every such
        /// hit was unopenable. Resolving the leaf against the real COM tree restores them.
        /// Names-only walk, same cost as <see cref="ListFolderPaths"/>; several matches are
        /// possible (leaf collisions are a documented delegate reality) and all are
        /// returned, shallowest first, for the caller to probe in order.
        /// </para>
        /// </summary>
        public IReadOnlyList<IReadOnlyList<string>> FindFolderPathsByLeafName(
            string storeDisplayName, string leafName, int absoluteWalkCap)
        {
            EnsureNotDisposed();
            if (string.IsNullOrWhiteSpace(storeDisplayName))
            {
                throw new ArgumentException("Store display name must not be blank.", nameof(storeDisplayName));
            }

            if (string.IsNullOrWhiteSpace(leafName))
            {
                throw new ArgumentException("Leaf name must not be blank.", nameof(leafName));
            }

            return _runner.Run(() =>
            {
                List<IReadOnlyList<string>> matches = new List<IReadOnlyList<string>>();
                dynamic? store = FindStoreByDisplayName(storeDisplayName);
                if (store == null)
                {
                    return (IReadOnlyList<IReadOnlyList<string>>)matches;
                }

                object? root = null;
                try
                {
                    root = store.GetRootFolder();
                    CollectFolderSegmentsByLeaf(
                        root!, new List<string>(), leafName, 1, absoluteWalkCap, matches);
                }
                catch (Exception ex) when (IsComCallFailure(ex))
                {
                }
                finally
                {
                    Release(root);
                    Release((object?)store);
                }

                matches.Sort((a, b) => a.Count.CompareTo(b.Count));
                return (IReadOnlyList<IReadOnlyList<string>>)matches;
            });
        }

        private void CollectFolderSegmentsByLeaf(
            object folderObject,
            List<string> parentSegments,
            string leafName,
            int depth,
            int absoluteWalkCap,
            List<IReadOnlyList<string>> matches)
        {
            if (depth > FolderWalkDepthGuard || matches.Count >= absoluteWalkCap)
            {
                return;
            }

            dynamic folder = folderObject;
            object? subFolders = null;
            try
            {
                subFolders = folder.Folders;
                dynamic folderCollection = (dynamic)subFolders!;
                int count = folderCollection.Count;
                for (int i = 1; i <= count; i++)
                {
                    object? child = null;
                    try
                    {
                        child = folderCollection[i];
                        string name = TryGetString(() => (string?)((dynamic)child!).Name) ?? string.Empty;
                        if (name.Length == 0)
                        {
                            continue;
                        }

                        List<string> segments = new List<string>(parentSegments) { name };
                        if (string.Equals(name, leafName, StringComparison.OrdinalIgnoreCase))
                        {
                            matches.Add(segments);
                        }

                        CollectFolderSegmentsByLeaf(child!, segments, leafName, depth + 1, absoluteWalkCap, matches);
                    }
                    catch (Exception ex) when (IsComCallFailure(ex))
                    {
                    }
                    finally
                    {
                        Release(child);
                    }
                }
            }
            catch (Exception ex) when (IsComCallFailure(ex))
            {
                // Folder without enumerable children.
            }
            finally
            {
                Release(subFolders);
            }
        }

        private void CollectFolderPaths(
            object folderObject, string parentPath, int depth, int absoluteWalkCap, List<string> result)
        {
            if (depth > FolderWalkDepthGuard || result.Count >= absoluteWalkCap)
            {
                return;
            }

            dynamic folder = folderObject;
            object? subFolders = null;
            try
            {
                subFolders = folder.Folders;
                dynamic folderCollection = (dynamic)subFolders!;
                int count = folderCollection.Count;
                for (int i = 1; i <= count; i++)
                {
                    if (result.Count >= absoluteWalkCap)
                    {
                        return;
                    }

                    object? child = null;
                    try
                    {
                        child = folderCollection[i];
                        string name = TryGetString(() => (string?)((dynamic)child!).Name) ?? string.Empty;
                        if (name.Length == 0)
                        {
                            continue;
                        }

                        string path = parentPath.Length == 0 ? name : parentPath + "/" + name;
                        result.Add(path);
                        CollectFolderPaths(child!, path, depth + 1, absoluteWalkCap, result);
                    }
                    catch (Exception ex) when (IsComCallFailure(ex))
                    {
                    }
                    finally
                    {
                        Release(child);
                    }
                }
            }
            catch (Exception ex) when (IsComCallFailure(ex))
            {
                // Folder without enumerable children.
            }
            finally
            {
                Release(subFolders);
            }
        }

        /// <summary>Recursion guard for the full-tree folder walk (real trees are a handful of levels).</summary>
        private const int FolderWalkDepthGuard = 64;

        private void CollectFolders(
            object folderObject,
            string storeDisplayName,
            string parentPath,
            int depth,
            int absoluteWalkCap,
            List<ComFolderInfo> result)
        {
            if (depth > FolderWalkDepthGuard)
            {
                return;
            }

            dynamic folder = folderObject;
            object? subFolders = null;
            try
            {
                subFolders = folder.Folders;
                dynamic folderCollection = (dynamic)subFolders!;
                int count = folderCollection.Count;

                // Stable order leg 2: siblings sorted by name (case-insensitive
                // ordinal; collection position breaks ties) so the flattened
                // depth-first list - and with it any offset paging - is deterministic
                // regardless of the provider's enumeration order.
                List<(string Name, int Index)> order = new List<(string, int)>(count);
                for (int i = 1; i <= count; i++)
                {
                    object? probe = null;
                    try
                    {
                        probe = folderCollection[i];
                        order.Add(((string)((dynamic)probe!).Name, i));
                    }
                    catch (Exception ex) when (IsComCallFailure(ex))
                    {
                    }
                    finally
                    {
                        Release(probe);
                    }
                }

                order.Sort((a, b) =>
                {
                    int byName = string.Compare(a.Name, b.Name, StringComparison.OrdinalIgnoreCase);
                    return byName != 0 ? byName : a.Index.CompareTo(b.Index);
                });

                foreach ((string name, int index) in order)
                {
                    if (result.Count >= absoluteWalkCap)
                    {
                        return;
                    }

                    object? child = null;
                    try
                    {
                        child = folderCollection[index];
                        dynamic c = (dynamic)child!;
                        string path = parentPath.Length == 0 ? name : parentPath + "/" + name;
                        long? itemCount = TryGetPropertyLong(c, "http://schemas.microsoft.com/mapi/proptag/0x36020003");
                        long? unread = TryGetPropertyLong(c, "http://schemas.microsoft.com/mapi/proptag/0x36030003");
                        int childCount = 0;
                        object? grandChildren = null;
                        try
                        {
                            grandChildren = c.Folders;
                            childCount = (int)((dynamic)grandChildren!).Count;
                        }
                        catch (COMException)
                        {
                        }
                        finally
                        {
                            Release(grandChildren);
                        }

                        result.Add(new ComFolderInfo(storeDisplayName, path, name, itemCount, unread, childCount));
                        if (childCount > 0)
                        {
                            CollectFolders(child, storeDisplayName, path, depth + 1, absoluteWalkCap, result);
                        }
                    }
                    catch (Exception ex) when (IsComCallFailure(ex))
                    {
                    }
                    finally
                    {
                        Release(child);
                    }
                }
            }
            catch (COMException)
            {
                // Folder without enumerable children.
            }
            finally
            {
                Release(subFolders);
            }
        }

        private static string? TryGetSenderSmtp(dynamic item)
        {
            // PR_SENDER_SMTP_ADDRESS first (Exchange senders report an X.500 DN in
            // SenderEmailAddress), then the raw address as fallback.
            string? smtp = TryGetPropertyString(item, "http://schemas.microsoft.com/mapi/proptag/0x5D01001F");
            if (!string.IsNullOrEmpty(smtp))
            {
                return smtp;
            }

            return TryGetString(() => (string?)item.SenderEmailAddress);
        }

        private static string? TryGetPropertyString(dynamic comObject, string schemaName)
        {
            object? accessor = null;
            try
            {
                accessor = comObject.PropertyAccessor;
                object? value = ((dynamic)accessor!).GetProperty(schemaName);
                return value as string;
            }
            catch (COMException)
            {
                return null;
            }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
            {
                return null;
            }
            finally
            {
                Release(accessor);
            }
        }

        private static long? TryGetPropertyLong(dynamic comObject, string schemaName)
        {
            object? accessor = null;
            try
            {
                accessor = comObject.PropertyAccessor;
                object? value = ((dynamic)accessor!).GetProperty(schemaName);
                if (value == null)
                {
                    return null;
                }

                return Convert.ToInt64(value, CultureInfo.InvariantCulture);
            }
            catch (COMException)
            {
                return null;
            }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
            {
                return null;
            }
            catch (FormatException)
            {
                return null;
            }
            catch (InvalidCastException)
            {
                return null;
            }
            catch (OverflowException)
            {
                return null;
            }
            finally
            {
                Release(accessor);
            }
        }

        private ComOpenResult? TryProbeViaGetTable(dynamic folder, string daslFilter, DateTime? indexReceivedUtc, int toleranceSeconds)
        {
            object? table = null;
            object? storeObject = null;
            try
            {
                string? storeId = null;
                try
                {
                    storeObject = folder.Store;
                    if (storeObject != null)
                    {
                        storeId = (string)((dynamic)storeObject).StoreID;
                    }
                }
                catch (COMException)
                {
                }

                table = folder.GetTable(daslFilter);
                dynamic t = (dynamic)table!;
                int entryIdIndex = FindTableColumn(t, "EntryID");
                if (entryIdIndex < 0)
                {
                    return null;
                }

                dynamic ns = _namespace!;
                int scanned = 0;
                while (!(bool)t.EndOfTable && scanned < 500)
                {
                    scanned++;
                    object? row = null;
                    object? itemObject = null;
                    try
                    {
                        row = t.GetNextRow();
                        object[] values = (object[])((dynamic)row!).GetValues();
                        if (entryIdIndex >= values.Length || values[entryIdIndex] is not string entryId || entryId.Length == 0)
                        {
                            continue;
                        }

                        try
                        {
                            itemObject = storeId != null ? ns.GetItemFromID(entryId, storeId) : ns.GetItemFromID(entryId);
                        }
                        catch (COMException)
                        {
                            continue;
                        }

                        ComOpenResult snapshot = Snapshot(itemObject);
                        if (!indexReceivedUtc.HasValue
                            || ReceivedTimeMatches(snapshot.ReceivedTime, indexReceivedUtc.Value, toleranceSeconds))
                        {
                            return snapshot;
                        }
                    }
                    finally
                    {
                        Release(itemObject);
                        Release(row);
                    }
                }

                return null;
            }
            catch (Exception ex) when (IsComCallFailure(ex))
            {
                // GetTable unsupported / filter rejected here (late-bound COM maps
                // E_INVALIDARG to ArgumentException) - the caller falls back.
                return null;
            }
            finally
            {
                Release(table);
                Release(storeObject);
            }
        }

        /// <summary>
        /// Late-bound COM failures do not always surface as COMException: the dynamic
        /// binder maps E_INVALIDARG to ArgumentException, E_POINTER to
        /// ArgumentNullException, and binding problems to RuntimeBinderException. Optional
        /// COM paths must treat all of these as "that call did not work here". Public:
        /// every caller with an optional COM path (Phase 2 fact 2) uses this same test.
        /// </summary>
        public static bool IsComCallFailure(Exception ex)
        {
            // MissingMemberException: the dynamic COM binder's shape for a failed
            // GetIDsOfNames ("Could not get dispatch ID for X") - live-observed with
            // 0x800706BA (RPC_S_SERVER_UNAVAILABLE) when Outlook exited mid-call
            // (soak, 2026-07-24). InvalidComObjectException: a detached RCW - what a
            // released-by-the-exit-watcher reference throws when used mid-flight
            // (ComGateway already treats it as a disconnect shape).
            return ex is COMException
                || ex is ArgumentException
                || ex is InvalidCastException
                || ex is MissingMemberException
                || ex is InvalidComObjectException
                || ex is Microsoft.CSharp.RuntimeBinder.RuntimeBinderException;
        }

        private static int FindTableColumn(dynamic table, string columnName)
        {
            object? columns = null;
            try
            {
                columns = table.Columns;
                dynamic cols = (dynamic)columns!;
                int count = cols.Count;
                for (int i = 1; i <= count; i++)
                {
                    object? column = null;
                    try
                    {
                        column = cols[i];
                        string name = (string)((dynamic)column!).Name;
                        if (string.Equals(name, columnName, StringComparison.OrdinalIgnoreCase))
                        {
                            return i - 1;
                        }
                    }
                    finally
                    {
                        Release(column);
                    }
                }

                return -1;
            }
            catch (Exception ex) when (IsComCallFailure(ex))
            {
                return -1;
            }
            finally
            {
                Release(columns);
            }
        }

        private static string EscapeDaslValue(string value)
        {
            return value.Replace("'", "''");
        }

        /// <summary>DASL date literal: UTC, invariant US format (documented DASL semantics).</summary>
        private static string FormatDaslUtc(DateTime value)
        {
            DateTime utc = value.Kind == DateTimeKind.Local ? value.ToUniversalTime() : value;
            return utc.ToString("MM/dd/yyyy HH:mm:ss", CultureInfo.InvariantCulture);
        }

        private static string SanitizeFileName(string fileName)
        {
            char[] invalid = Path.GetInvalidFileNameChars();
            System.Text.StringBuilder sb = new System.Text.StringBuilder(fileName.Length);
            foreach (char c in fileName)
            {
                sb.Append(Array.IndexOf(invalid, c) >= 0 ? '_' : c);
            }

            string sanitized = sb.ToString().Trim();
            return sanitized.Length == 0 ? "attachment" : sanitized;
        }

        private static string MakeUniquePath(string directory, string fileName)
        {
            string candidate = Path.Combine(directory, fileName);
            if (!File.Exists(candidate))
            {
                return candidate;
            }

            string baseName = Path.GetFileNameWithoutExtension(fileName);
            string extension = Path.GetExtension(fileName);
            for (int i = 2; i < 10000; i++)
            {
                candidate = Path.Combine(
                    directory,
                    string.Format(CultureInfo.InvariantCulture, "{0} ({1}){2}", baseName, i, extension));
                if (!File.Exists(candidate))
                {
                    return candidate;
                }
            }

            throw new IOException("Could not find a free file name in the target directory.");
        }

        /// <summary>
        /// Recursively walks every folder of a store and snapshots all mail items
        /// (OlObjectClass 43), including plain-text bodies. Intended ONLY for the
        /// designated tiny test store (completeness oracle, v3.MD S2) - walking a
        /// multi-GB store would be the exact scan anti-pattern this project exists to
        /// avoid.
        /// </summary>
        public IReadOnlyList<ComWalkedItem> WalkStoreMailItems(string storeDisplayName)
        {
            EnsureNotDisposed();
            if (string.IsNullOrWhiteSpace(storeDisplayName))
            {
                throw new ArgumentException("Store display name must not be blank.", nameof(storeDisplayName));
            }

            return _runner.Run(() =>
            {
                List<ComWalkedItem> result = new List<ComWalkedItem>();
                dynamic? store = FindStoreByDisplayName(storeDisplayName);
                if (store == null)
                {
                    throw new InvalidOperationException("Store not found by display name.");
                }

                object? root = null;
                try
                {
                    root = store.GetRootFolder();
                    WalkFolder(root!, string.Empty, result);
                }
                finally
                {
                    Release(root);
                    Release(store);
                }

                return (IReadOnlyList<ComWalkedItem>)result;
            });
        }

        private void WalkFolder(object folderObject, string path, List<ComWalkedItem> result)
        {
            dynamic folder = folderObject;
            object? items = null;
            try
            {
                items = folder.Items;
                dynamic itemCollection = items!;
                int itemCount = itemCollection.Count;
                for (int i = 1; i <= itemCount; i++)
                {
                    object? item = null;
                    try
                    {
                        item = itemCollection[i];
                        dynamic mail = item!;
                        int itemClass;
                        try
                        {
                            itemClass = (int)mail.Class;
                        }
                        catch (COMException)
                        {
                            continue;
                        }

                        if (itemClass != OlMailItemClass)
                        {
                            continue;
                        }

                        string entryId = (string)mail.EntryID;
                        string? subject = TryGetString(() => (string?)mail.Subject);
                        string? body = TryGetString(() => (string?)mail.Body);
                        DateTime? received = TryGetDateTime(() => (DateTime)mail.ReceivedTime);
                        result.Add(new ComWalkedItem(entryId, subject, body, received, path, itemClass));
                    }
                    finally
                    {
                        Release(item);
                    }
                }
            }
            catch (COMException)
            {
                // Folders whose items cannot be enumerated (e.g. search folders) are skipped.
            }
            finally
            {
                Release(items);
            }

            object? subFolders = null;
            try
            {
                subFolders = folder.Folders;
                dynamic folderCollection = subFolders!;
                int folderCount = folderCollection.Count;
                for (int i = 1; i <= folderCount; i++)
                {
                    object? child = null;
                    try
                    {
                        child = folderCollection[i];
                        string name = (string)((dynamic)child!).Name;
                        string childPath = path.Length == 0 ? name : path + "/" + name;
                        WalkFolder(child, childPath, result);
                    }
                    finally
                    {
                        Release(child);
                    }
                }
            }
            catch (COMException)
            {
                // No subfolders / inaccessible.
            }
            finally
            {
                Release(subFolders);
            }
        }

        private dynamic? FindStoreByDisplayName(string displayName)
        {
            dynamic ns = _namespace!;
            dynamic stores = ns.Stores;
            try
            {
                int count = stores.Count;
                for (int i = 1; i <= count; i++)
                {
                    dynamic store = stores[i];
                    string name;
                    try
                    {
                        name = (string)store.DisplayName;
                    }
                    catch (COMException)
                    {
                        Release(store);
                        continue;
                    }

                    if (string.Equals(name, displayName, StringComparison.OrdinalIgnoreCase))
                    {
                        return store;
                    }

                    Release(store);
                }

                return null;
            }
            finally
            {
                Release(stores);
            }
        }

        private static ComOpenResult Snapshot(object? itemObject)
        {
            if (itemObject == null)
            {
                throw new InvalidOperationException("GetItemFromID returned null.");
            }

            dynamic item = itemObject;
            string entryId = (string)item.EntryID;
            string? subject = TryGetString(() => (string?)item.Subject);
            DateTime? received = TryGetDateTime(() => (DateTime)item.ReceivedTime);
            int? itemClass = null;
            try
            {
                itemClass = (int)item.Class;
            }
            catch (COMException)
            {
            }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
            {
            }

            return new ComOpenResult(entryId, subject, received, itemClass);
        }

        /// <summary>
        /// Compares a COM ReceivedTime (local wall time) against an index UTC timestamp,
        /// accepting both the UTC interpretation and the raw-local interpretation within
        /// the tolerance (the live tests record which interpretation actually holds).
        /// </summary>
        public static bool ReceivedTimeMatches(DateTime? comReceivedLocal, DateTime indexUtc, int toleranceSeconds)
        {
            if (!comReceivedLocal.HasValue)
            {
                return false;
            }

            DateTime local = comReceivedLocal.Value;
            DateTime asUtc = DateTime.SpecifyKind(local, DateTimeKind.Local).ToUniversalTime();
            double utcDelta = Math.Abs((asUtc - indexUtc).TotalSeconds);
            double rawDelta = Math.Abs((DateTime.SpecifyKind(local, DateTimeKind.Utc) - indexUtc).TotalSeconds);
            return utcDelta <= toleranceSeconds || rawDelta <= toleranceSeconds;
        }

        private static string? TryGetString(Func<string?> getter)
        {
            try
            {
                return getter();
            }
            catch (COMException)
            {
                return null;
            }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
            {
                return null;
            }
        }

        private static DateTime? TryGetDateTime(Func<DateTime> getter)
        {
            try
            {
                DateTime value = getter();
                // Outlook uses 4501-01-01 as "no value".
                return value.Year >= 4500 ? (DateTime?)null : value;
            }
            catch (COMException)
            {
                return null;
            }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
            {
                return null;
            }
        }

        internal static void Release(object? comObject)
        {
            if (comObject != null && Marshal.IsComObject(comObject))
            {
                Marshal.ReleaseComObject(comObject);
            }
        }

        private void EnsureNotDisposed()
        {
            if (_disposed)
            {
                throw new ObjectDisposedException(nameof(OutlookComSession));
            }
        }

        /// <summary>
        /// Releases COM references and stops the STA thread. Never quits Outlook - if
        /// this session started it headless, Outlook keeps running (index updates with
        /// it, D17) and, measured 2026-07-23, exits on its own ~11.5 minutes after the
        /// LAST client releases; a release also unsticks (~6 s) an Outlook parked by a
        /// quit-while-attached (SF-2).
        /// </summary>
        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;
            Process? watched = _watchedProcess;
            _watchedProcess = null;
            if (watched != null)
            {
                try
                {
                    watched.Exited -= OnWatchedProcessExited;
                    watched.EnableRaisingEvents = false;
                }
                catch (Exception)
                {
                    // Teardown of a possibly-dead process handle.
                }

                watched.Dispose();
            }

            try
            {
                _runner.Run(() =>
                {
                    _quitSinkRegistration?.Unadvise();
                    _quitSinkRegistration = null;
                    _quitSink = null;

                    // D49: CLOSE the pin, do not merely release it. Measured: an Explorer
                    // released but left in Outlook's collection outlives the session that
                    // made it and keeps Outlook up for good - a later Application.Quit
                    // (the user choosing Exit) then does NOT terminate the process, and
                    // repeated sessions accumulate invisible Explorers. Closing it here
                    // restores the pre-D49 lifecycle exactly: the pin exists for precisely
                    // as long as the session that needs it, and Outlook is left in the
                    // state it would have been in anyway. Any other window - the user's, or
                    // another live session's pin - still keeps Outlook running.
                    if (_composeSurfacePin != null)
                    {
                        try
                        {
                            ((dynamic)_composeSurfacePin).Close();
                        }
                        catch (Exception)
                        {
                            // Outlook may already be gone; teardown must not throw.
                        }
                    }

                    ComposeSurface.ForgetPin(_composeSurfacePin);
                    Release(_composeSurfacePin);
                    _composeSurfacePin = null;
                    Release(_namespace);
                    Release(_application);
                    _namespace = null;
                    _application = null;
                });
            }
            catch (Exception)
            {
                // Dispose must not throw; the RCWs are finalized by the GC below.
            }

            _runner.Dispose();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}
