using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;

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

        private readonly PumpedStaRunner _runner;
        private object? _application;
        private object? _namespace;
        private bool _disposed;
        private Process? _watchedProcess;
        private OutlookQuitSink? _quitSink;
        private OutlookQuitSinkRegistration? _quitSinkRegistration;
        private Action<OutlookComSession>? _onOutlookGone;
        private int _outlookGoneSignaled;

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
        /// Full read of one item by its REAL EntryID (v3.MD section 8 L2): plain-text
        /// body (HTML converted when Outlook has no text rendering), capped at
        /// <paramref name="maxBodyChars"/> with the true total reported, recipients with
        /// SMTP addresses, attachment list, and transport headers on request. Returns
        /// null with a content-free error description on failure (S4).
        /// </summary>
        public ComItemDetail? TryReadItem(string entryIdHex, string? storeId, bool includeHeaders, int maxBodyChars, out string? error)
        {
            EnsureNotDisposed();
            if (string.IsNullOrWhiteSpace(entryIdHex))
            {
                throw new ArgumentException("EntryID must not be blank.", nameof(entryIdHex));
            }

            if (maxBodyChars < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(maxBodyChars));
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
                    return SnapshotDetail(itemObject!, includeHeaders, maxBodyChars);
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
        /// Folder tree listing (list_folders): store-relative paths with item/unread
        /// counts (PR_CONTENT_COUNT / PR_CONTENT_UNREAD), depth- and count-capped for
        /// compact payloads (v3.MD section 12).
        /// </summary>
        public IReadOnlyList<ComFolderInfo> ListFolders(string? storeDisplayName, int maxDepth, int maxFolders)
        {
            EnsureNotDisposed();
            if (maxDepth < 1)
            {
                throw new ArgumentOutOfRangeException(nameof(maxDepth));
            }

            if (maxFolders < 1)
            {
                throw new ArgumentOutOfRangeException(nameof(maxFolders));
            }

            return _runner.Run(() =>
            {
                dynamic ns = _namespace!;
                List<ComFolderInfo> result = new List<ComFolderInfo>();
                dynamic stores = ns.Stores;
                try
                {
                    int count = stores.Count;
                    for (int i = 1; i <= count && result.Count < maxFolders; i++)
                    {
                        dynamic store = stores[i];
                        object? root = null;
                        try
                        {
                            string name;
                            try
                            {
                                name = (string)store.DisplayName;
                            }
                            catch (COMException)
                            {
                                continue;
                            }

                            if (storeDisplayName != null
                                && !string.Equals(name, storeDisplayName, StringComparison.OrdinalIgnoreCase))
                            {
                                continue;
                            }

                            try
                            {
                                root = store.GetRootFolder();
                            }
                            catch (COMException)
                            {
                                continue;
                            }

                            CollectFolders(root!, name, string.Empty, 1, maxDepth, maxFolders, result);
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
        /// Fresh-mode gap sweep (v3.MD D19): enumerates each store's Inbox and Sent
        /// Items for items received/sent at or after <paramref name="sinceUtc"/>. Items
        /// are opened for authoritative properties and carry their REAL EntryIDs; bodies
        /// are fetched only when the caller needs term matching. Bounded by
        /// <paramref name="perFolderCap"/> per folder.
        /// </summary>
        public ComSweepResult SweepDefaultFoldersNewerThan(
            DateTime sinceUtc,
            int perFolderCap,
            bool includeBodies,
            string? onlyStoreDisplayName)
        {
            EnsureNotDisposed();
            if (perFolderCap < 1)
            {
                throw new ArgumentOutOfRangeException(nameof(perFolderCap));
            }

            return _runner.Run(() =>
            {
                dynamic ns = _namespace!;
                List<ComMailBrief> items = new List<ComMailBrief>();
                int swept = 0;
                int skipped = 0;
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
                                skipped += 2;
                                continue;
                            }

                            if (onlyStoreDisplayName != null
                                && !string.Equals(storeName, onlyStoreDisplayName, StringComparison.OrdinalIgnoreCase))
                            {
                                continue;
                            }

                            // 6 = olFolderInbox, 5 = olFolderSentMail.
                            foreach ((int folderId, string folderKind) in new[] { (6, "inbox"), (5, "sent") })
                            {
                                object? folder = null;
                                try
                                {
                                    folder = store.GetDefaultFolder(folderId);
                                    SweepFolder(ns, folder!, storeName, storeId, folderKind, sinceUtc, perFolderCap, includeBodies, items);
                                    swept++;
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

                return new ComSweepResult(items, swept, skipped);
            });
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

        private static string DescribeComFailure(Exception ex)
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
        /// that account's signature, and sets HTMLBody exactly once with the agent text
        /// ABOVE the signature. Saves to Drafts; <paramref name="display"/> additionally
        /// opens the draft in an Inspector for the user (D4 default behavior).
        /// </summary>
        public ComDraftCreateResult? TryCreateNewDraft(
            string accountSmtpAddress,
            IReadOnlyList<string> toRecipients,
            IReadOnlyList<string> ccRecipients,
            string subject,
            string bodyText,
            bool display,
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

            if (ccRecipients == null)
            {
                throw new ArgumentNullException(nameof(ccRecipients));
            }

            if (subject == null)
            {
                throw new ArgumentNullException(nameof(subject));
            }

            if (bodyText == null)
            {
                throw new ArgumentNullException(nameof(bodyText));
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

                    (bool signatureInjected, long textBefore, long textAfter, string htmlAfter) =
                        TouchInspectorForSignature((object)draft);

                    string fragment = OutlookAI.Core.Text.HtmlBodyComposer.ToHtmlFragment(bodyText);
                    draft.HTMLBody = OutlookAI.Core.Text.HtmlBodyComposer.InsertAtBodyTop(
                        htmlAfter.Length > 0 ? htmlAfter : null, fragment);
                    draft.Subject = subject;
                    AddRecipients(draft, toRecipients, 1);
                    AddRecipients(draft, ccRecipients, 2);
                    draft.Save();

                    // The GetInspector touch left a HIDDEN Inspector alive inside
                    // Outlook (it shows up in Application.Inspectors - Phase-4 live
                    // finding). Close it now that the draft is saved; Display() below
                    // opens a fresh visible one for the final item when requested.
                    CloseHiddenInspector(mail!);
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
                        display);
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
            string bodyText,
            bool display,
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

            if (bodyText == null)
            {
                throw new ArgumentNullException(nameof(bodyText));
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

                    (bool signatureInjected, long textBefore, long textAfter, string htmlAfter) =
                        TouchInspectorForSignature((object)draft);

                    string fragment = OutlookAI.Core.Text.HtmlBodyComposer.ToHtmlFragment(bodyText);
                    draft.HTMLBody = OutlookAI.Core.Text.HtmlBodyComposer.InsertAtBodyTop(
                        htmlAfter.Length > 0 ? htmlAfter : null, fragment);
                    if (kind == ComDerivedDraftKind.Forward)
                    {
                        AddRecipients(draft, toRecipients, 1);
                    }

                    draft.Save();

                    // Same hidden-Inspector cleanup as the new-draft path (the
                    // GetInspector signature touch materializes one inside Outlook).
                    CloseHiddenInspector(mail!);

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
                        display);
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
                        info.Recipients);
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
                    string currentHash = SendContentHash.Compute(info.Subject, info.Recipients, body, sentOnBehalfOfName);
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
        /// STA-side signature injection (v3.MD section 3): with SendUsingAccount already
        /// pinned, touching GetInspector makes Outlook inject that account's signature
        /// exactly as if the user opened the compose window. Detection is TEXT-based
        /// (HTML template expansion without a signature adds markup but no text).
        /// Returns the post-touch HTML for composition.
        /// </summary>
        private static (bool SignatureInjected, long TextBefore, long TextAfter, string HtmlAfter) TouchInspectorForSignature(object draftObject)
        {
            dynamic draft = draftObject;
            string htmlBefore = TryGetString(() => (string?)draft.HTMLBody) ?? string.Empty;
            object? inspector = null;
            try
            {
                inspector = draft.GetInspector;
            }
            catch (Exception ex) when (IsComCallFailure(ex))
            {
                // No inspector available - signature stays uninjected; recorded via the
                // unchanged text length.
            }
            finally
            {
                Release(inspector);
            }

            string htmlAfter = TryGetString(() => (string?)draft.HTMLBody) ?? string.Empty;
            long textBefore = CountNonWhitespaceText(htmlBefore);
            long textAfter = CountNonWhitespaceText(htmlAfter);
            return (textAfter > textBefore, textBefore, textAfter, htmlAfter);
        }

        /// <summary>
        /// STA-side: closes the hidden Inspector the GetInspector signature touch left
        /// behind (olDiscard - the item is already saved, nothing is lost). Without
        /// this, a display:false draft still surfaces in Application.Inspectors.
        /// </summary>
        private static void CloseHiddenInspector(object mailObject)
        {
            object? inspector = null;
            try
            {
                inspector = ((dynamic)mailObject).GetInspector;
                if (inspector != null)
                {
                    ((dynamic)inspector).Close(1); // 1 = olDiscard
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
        private static void AddRecipients(dynamic mail, IReadOnlyList<string> addresses, int type)
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
                    }
                    finally
                    {
                        Release(recipient);
                    }
                }

                try
                {
                    collection.ResolveAll();
                }
                catch (Exception ex) when (IsComCallFailure(ex))
                {
                    // Unresolved recipients are legal on drafts; the user resolves on send.
                }
            }
            finally
            {
                Release(recipients);
            }
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

        /// <summary>STA-side identity/threading snapshot of a mail item.</summary>
        private ComDraftInfo SnapshotDraft(object itemObject, string? fallbackStoreName = null, string? fallbackStoreId = null, string? fallbackFolderName = null, string? fallbackFolderEntryId = null, string? fallbackSendUsingSmtp = null)
        {
            dynamic item = itemObject;
            string entryId = (string)item.EntryID;
            string? subject = TryGetString(() => (string?)item.Subject);
            string? conversationIndex = TryGetString(() => (string?)item.ConversationIndex);
            string? conversationId = TryGetString(() => (string?)item.ConversationID);

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
                recipients);
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
        /// </summary>
        public ComExhaustiveResult ExhaustiveScan(
            string storeDisplayName,
            IReadOnlyList<string>? folderPath,
            IReadOnlyList<string>? terms,
            DateTime? sinceUtc,
            DateTime? beforeUtc,
            int maxItems,
            int timeBudgetMs)
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
                        ? ExhaustiveDaslFilter.Build(terms, sinceUtc, beforeUtc, ExhaustiveEngine.CiPhraseMatch)
                        : null,
                    LikeFilter = ExhaustiveDaslFilter.Build(terms, sinceUtc, beforeUtc, ExhaustiveEngine.Like),
                };

                try
                {
                    bool recurse = folderPath == null || folderPath.Count == 0;
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
            if (state.ShouldStop)
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
                // No enumerable subfolders.
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
                    if (state.Clock.Elapsed > state.Budget)
                    {
                        state.TimedOut = true;
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
        }

        private void SweepFolder(
            dynamic ns,
            object folderObject,
            string storeName,
            string storeId,
            string folderKind,
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
                    return;
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
            }
            catch (Exception ex) when (IsComCallFailure(ex))
            {
                // GetTable/filter unsupported on this folder - counted as swept with zero rows.
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

        private ComItemDetail SnapshotDetail(object itemObject, bool includeHeaders, int maxBodyChars)
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
            // mail; fall back to converting .HTMLBody ourselves when it is empty.
            string body = string.Empty;
            string bodyOrigin = "none";
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

            long bodyTotal = body.Length;
            if (body.Length > maxBodyChars)
            {
                body = body.Substring(0, maxBodyChars);
            }

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
                headers);
        }

        private void CollectFolders(
            object folderObject,
            string storeDisplayName,
            string parentPath,
            int depth,
            int maxDepth,
            int maxFolders,
            List<ComFolderInfo> result)
        {
            dynamic folder = folderObject;
            object? subFolders = null;
            try
            {
                subFolders = folder.Folders;
                dynamic folderCollection = (dynamic)subFolders!;
                int count = folderCollection.Count;
                for (int i = 1; i <= count; i++)
                {
                    if (result.Count >= maxFolders)
                    {
                        return;
                    }

                    object? child = null;
                    try
                    {
                        child = folderCollection[i];
                        dynamic c = (dynamic)child!;
                        string name;
                        try
                        {
                            name = (string)c.Name;
                        }
                        catch (COMException)
                        {
                            continue;
                        }

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
                        if (depth < maxDepth && childCount > 0)
                        {
                            CollectFolders(child, storeDisplayName, path, depth + 1, maxDepth, maxFolders, result);
                        }
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

        private static void Release(object? comObject)
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
