using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Globalization;
using System.Runtime.InteropServices;

namespace OutlookAI.Core.IndexSearch
{
    /// <summary>Which query path reached the SystemIndex (v3.MD section 0.6 Phase 1: record either way).</summary>
    public enum IndexProviderKind
    {
        /// <summary>System.Data.OleDb over Search.CollatorDSO - the primary path.</summary>
        OleDb = 0,

        /// <summary>Late-bound ADODB COM (probe-proven on this machine) - the fallback path.</summary>
        AdodbCom = 1,
    }

    /// <summary>Executes Windows Search SQL against the SystemIndex catalog.</summary>
    public interface IIndexClient
    {
        /// <summary>Which provider this client uses.</summary>
        IndexProviderKind Provider { get; }

        /// <summary>
        /// Runs <paramref name="sql"/> and returns up to <paramref name="maxRows"/> rows as
        /// case-insensitive column-name -> value maps. WS-SQL recordsets are forward-only;
        /// rows are drained eagerly.
        /// </summary>
        /// <param name="commandTimeoutSeconds">
        /// Per-query timeout. Defaults to <see cref="OleDbIndexClient.DefaultCommandTimeoutSeconds"/>.
        /// outlook_health passes a much shorter one: it must answer while the machine is
        /// struggling, and a saturated indexer is exactly when it gets asked.
        /// </param>
        IReadOnlyList<IReadOnlyDictionary<string, object?>> ExecuteRows(string sql, int maxRows, int? commandTimeoutSeconds = null);
    }

    /// <summary>Primary client: System.Data.OleDb over the Search.CollatorDSO provider.</summary>
    public sealed class OleDbIndexClient : IIndexClient
    {
        /// <summary>Connection string from the validated probes (v3.MD section 4).</summary>
        public const string ConnectionString = "Provider=Search.CollatorDSO;Extended Properties='Application=Windows';";

        /// <inheritdoc />
        public IndexProviderKind Provider => IndexProviderKind.OleDb;

        /// <inheritdoc />
        /// <summary>Default per-query timeout for index queries.</summary>
        public const int DefaultCommandTimeoutSeconds = 30;

        public IReadOnlyList<IReadOnlyDictionary<string, object?>> ExecuteRows(string sql, int maxRows, int? commandTimeoutSeconds = null)
        {
            if (sql == null)
            {
                throw new ArgumentNullException(nameof(sql));
            }

            List<IReadOnlyDictionary<string, object?>> rows = new List<IReadOnlyDictionary<string, object?>>();
            using (OleDbConnection connection = new OleDbConnection(ConnectionString))
            {
                connection.Open();
                using (OleDbCommand command = new OleDbCommand(sql, connection))
                {
                    command.CommandTimeout = commandTimeoutSeconds ?? DefaultCommandTimeoutSeconds;
                    using (OleDbDataReader reader = command.ExecuteReader())
                    {
                        int fieldCount = reader.FieldCount;
                        string[] names = new string[fieldCount];
                        for (int i = 0; i < fieldCount; i++)
                        {
                            names[i] = reader.GetName(i);
                        }

                        while (rows.Count < maxRows && reader.Read())
                        {
                            Dictionary<string, object?> row = new Dictionary<string, object?>(fieldCount, StringComparer.OrdinalIgnoreCase);
                            for (int i = 0; i < fieldCount; i++)
                            {
                                object value = reader.GetValue(i);
                                row[names[i]] = value is DBNull ? null : value;
                            }

                            rows.Add(row);
                        }
                    }
                }
            }

            return rows;
        }
    }

    /// <summary>
    /// Fallback client: late-bound ADODB COM (the exact path the section-5 probe scripts
    /// proved on this machine). Only used when the OleDb provider path fails.
    /// </summary>
    public sealed class AdodbIndexClient : IIndexClient
    {
        /// <inheritdoc />
        public IndexProviderKind Provider => IndexProviderKind.AdodbCom;

        /// <inheritdoc />
        public IReadOnlyList<IReadOnlyDictionary<string, object?>> ExecuteRows(string sql, int maxRows, int? commandTimeoutSeconds = null)
        {
            if (sql == null)
            {
                throw new ArgumentNullException(nameof(sql));
            }

            Type connectionType = Type.GetTypeFromProgID("ADODB.Connection")
                ?? throw new InvalidOperationException("ADODB.Connection ProgID is not registered.");
            Type recordsetType = Type.GetTypeFromProgID("ADODB.Recordset")
                ?? throw new InvalidOperationException("ADODB.Recordset ProgID is not registered.");

            List<IReadOnlyDictionary<string, object?>> rows = new List<IReadOnlyDictionary<string, object?>>();
            dynamic? connection = null;
            dynamic? recordset = null;
            try
            {
                connection = Activator.CreateInstance(connectionType)
                    ?? throw new InvalidOperationException("Failed to create ADODB.Connection.");
                connection.Open(OleDbIndexClient.ConnectionString);

                recordset = Activator.CreateInstance(recordsetType)
                    ?? throw new InvalidOperationException("Failed to create ADODB.Recordset.");
                recordset.Open(sql, connection);

                dynamic fields = recordset.Fields;
                int fieldCount = fields.Count;
                string[] names = new string[fieldCount];
                for (int i = 0; i < fieldCount; i++)
                {
                    dynamic field = fields.Item(i);
                    names[i] = (string)field.Name;
                }

                while (rows.Count < maxRows && !(bool)recordset.EOF)
                {
                    Dictionary<string, object?> row = new Dictionary<string, object?>(fieldCount, StringComparer.OrdinalIgnoreCase);
                    for (int i = 0; i < fieldCount; i++)
                    {
                        object? value = fields.Item(i).Value;
                        row[names[i]] = value is DBNull ? null : value;
                    }

                    rows.Add(row);
                    recordset.MoveNext();
                }
            }
            finally
            {
                TryCloseAndRelease(recordset);
                TryCloseAndRelease(connection);
            }

            return rows;
        }

        private static void TryCloseAndRelease(dynamic? comObject)
        {
            if (comObject == null)
            {
                return;
            }

            try
            {
                if ((int)comObject.State != 0)
                {
                    comObject.Close();
                }
            }
            catch (COMException)
            {
                // Best-effort close; release below regardless.
            }

            object o = comObject;
            if (Marshal.IsComObject(o))
            {
                Marshal.ReleaseComObject(o);
            }
        }
    }

    /// <summary>Chooses the working index client (OleDb primary, ADODB COM fallback).</summary>
    public static class IndexClientFactory
    {
        /// <summary>
        /// Probes the OleDb x CollatorDSO path with a TOP 1 statement; on failure falls
        /// back to ADODB COM. <paramref name="report"/> describes the outcome for the
        /// decision log (record either way - v3.MD section 0.6 Phase 1).
        /// </summary>
        public static IIndexClient CreateAuto(out string report)
        {
            const string probeSql = "SELECT TOP 1 System.ItemUrl FROM SystemIndex";
            OleDbIndexClient oleDb = new OleDbIndexClient();
            try
            {
                int probeRows = oleDb.ExecuteRows(probeSql, 1).Count;
                report = string.Format(
                    CultureInfo.InvariantCulture,
                    "IndexClient=OleDb (Search.CollatorDSO via System.Data.OleDb; probe rows={0})",
                    probeRows);
                return oleDb;
            }
            catch (Exception oleDbError) when (oleDbError is OleDbException || oleDbError is InvalidOperationException || oleDbError is NotSupportedException || oleDbError is PlatformNotSupportedException)
            {
                AdodbIndexClient adodb = new AdodbIndexClient();
                int probeRows = adodb.ExecuteRows(probeSql, 1).Count;
                report = string.Format(
                    CultureInfo.InvariantCulture,
                    "IndexClient=AdodbCom fallback (OleDb failed: {0}: {1}; ADODB probe rows={2})",
                    oleDbError.GetType().Name,
                    oleDbError.Message,
                    probeRows);
                return adodb;
            }
        }
    }
}
