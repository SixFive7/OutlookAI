using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;

namespace OutlookAI.Core.Com
{
    /// <summary>
    /// Runs all work on one dedicated STA thread - Outlook COM objects must never be
    /// marshaled across threads (v3.MD section 12). Phase-1 minimal: a serialized work
    /// queue without a Win32 message pump, sufficient for synchronous outgoing calls
    /// (GetItemFromID, folder walks). Phase 2 replaces this with the pumped ComGateway
    /// required for v3.1 event sinks (v3.MD section 0.5.2).
    /// </summary>
    internal sealed class StaComRunner : IDisposable
    {
        private readonly BlockingCollection<KeyValuePair<Func<object?>, TaskCompletionSource<object?>>> _queue =
            new BlockingCollection<KeyValuePair<Func<object?>, TaskCompletionSource<object?>>>();

        private readonly Thread _thread;

        internal StaComRunner()
        {
            _thread = new Thread(Pump)
            {
                IsBackground = true,
                Name = "OutlookAI.Phase1.ComSta",
            };
            _thread.SetApartmentState(ApartmentState.STA);
            _thread.Start();
        }

        internal T Run<T>(Func<T> work)
        {
            if (work == null)
            {
                throw new ArgumentNullException(nameof(work));
            }

            TaskCompletionSource<object?> completion =
                new TaskCompletionSource<object?>(TaskCreationOptions.RunContinuationsAsynchronously);
            _queue.Add(new KeyValuePair<Func<object?>, TaskCompletionSource<object?>>(() => work(), completion));
            return (T)completion.Task.GetAwaiter().GetResult()!;
        }

        internal void Run(Action work)
        {
            Run<object?>(() =>
            {
                work();
                return null;
            });
        }

        private void Pump()
        {
            foreach (KeyValuePair<Func<object?>, TaskCompletionSource<object?>> item in _queue.GetConsumingEnumerable())
            {
                try
                {
                    item.Value.SetResult(item.Key());
                }
                catch (Exception ex)
                {
                    item.Value.SetException(ex);
                }
            }
        }

        public void Dispose()
        {
            _queue.CompleteAdding();
            _thread.Join(TimeSpan.FromSeconds(15));
            _queue.Dispose();
        }
    }

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
    /// Phase-1 minimal Outlook COM helper (v3.MD section 0.6 Phase 1): late-bound dynamic
    /// dispatch (no PIA - keeps dotnet build working), one dedicated STA thread, enough
    /// surface to verify EntryID decodes on open and to run the completeness-oracle store
    /// walk. Never quits or kills Outlook (S7); may start it when allowed (D17), except
    /// while the OutlookAISetup installer mutex is held.
    /// </summary>
    public sealed class OutlookComSession : IDisposable
    {
        private const int OlMailItemClass = 43;

        private readonly StaComRunner _runner;
        private object? _application;
        private object? _namespace;
        private bool _disposed;

        private OutlookComSession(StaComRunner runner, bool startedOutlook)
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

            StaComRunner runner = new StaComRunner();
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
        /// ItemPathDisplay fallback mapping (v3.MD section 4): walks the store's folder
        /// tree along <paramref name="folderPath"/>, then probes the folder with a narrow
        /// DASL subject restriction and (when given) a ReceivedTime tolerance - an exact
        /// probe, not a scan. Returns the first matching item snapshot.
        /// </summary>
        public ComOpenResult? TryResolveByPath(
            string storeDisplayName,
            IReadOnlyList<string> folderPath,
            string itemSubject,
            DateTime? indexReceivedUtc,
            out string? error)
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
                    items = folder.Items;
                    string filter = "@SQL=\"urn:schemas:httpmail:subject\" = '" + itemSubject.Replace("'", "''") + "'";
                    restricted = ((dynamic)items!).Restrict(filter);
                    dynamic restrictedItems = restricted!;
                    int count = restrictedItems.Count;
                    for (int i = 1; i <= count; i++)
                    {
                        object? candidate = null;
                        try
                        {
                            candidate = restrictedItems[i];
                            ComOpenResult snapshot = Snapshot(candidate);
                            if (!indexReceivedUtc.HasValue || ReceivedTimeMatches(snapshot.ReceivedTime, indexReceivedUtc.Value, toleranceSeconds: 120))
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
