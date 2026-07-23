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

        private OutlookComSession(PumpedStaRunner runner, bool startedOutlook)
        {
            _runner = runner;
            StartedOutlook = startedOutlook;
        }

        /// <summary>True when connecting had to launch a new OUTLOOK.EXE process.</summary>
        public bool StartedOutlook { get; }

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
        public static OutlookComSession Connect(bool allowStartingOutlook = true)
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
                });
                return session;
            }
            catch
            {
                session.Dispose();
                throw;
            }
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
                object? root = null;
                object? currentFolder = null;
                object? items = null;
                object? restricted = null;
                try
                {
                    dynamic? store = FindStoreByDisplayName(storeDisplayName);
                    if (store == null)
                    {
                        capturedError = "StoreNotFound";
                        return null;
                    }

                    try
                    {
                        root = store.GetRootFolder();
                    }
                    finally
                    {
                        Release(store);
                    }

                    currentFolder = root;
                    root = null;
                    foreach (string segment in folderPath)
                    {
                        dynamic folders = ((dynamic)currentFolder!).Folders;
                        object? next = null;
                        try
                        {
                            next = folders[segment];
                        }
                        catch (COMException)
                        {
                            capturedError = "FolderNotFound";
                            return null;
                        }
                        finally
                        {
                            Release(folders);
                        }

                        Release(currentFolder);
                        currentFolder = next;
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
                    Release(root);
                }
            });

            error = capturedError;
            return result;
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
                                catch (COMException)
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
                    t.Sort("urn:schemas:httpmail:datereceived", true);
                }
                catch (COMException)
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
                        catch (COMException)
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
            catch (COMException)
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
            catch (COMException)
            {
                // GetTable unsupported / filter rejected here - the caller falls back.
                return null;
            }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
            {
                return null;
            }
            finally
            {
                Release(table);
                Release(storeObject);
            }
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
            catch (COMException)
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
        /// Releases COM references and stops the STA thread. Never quits Outlook - if this
        /// session started it, it stays running (index updates resume with it, D17).
        /// </summary>
        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;
            try
            {
                _runner.Run(() =>
                {
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
