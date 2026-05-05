using System;
using System.IO;
using Newtonsoft.Json;

namespace Deriva.Excel.ETTJ
{
    internal static class EttjStatusStore
    {
        private static readonly object SyncRoot = new object();
        private static readonly string StatusFile =
            Path.Combine(EttjSettings.AppDataDir, "status.json");

        internal static StatusSnapshot Load()
        {
            lock (SyncRoot)
            {
                return LoadUnlocked();
            }
        }

        internal static void RecordSuccess(string item, string detail)
        {
            Record(item, "Success", detail);
        }

        internal static void RecordError(string item, string detail)
        {
            Record(item, "Error", detail);
        }

        private static void Record(string item, string status, string detail)
        {
            lock (SyncRoot)
            {
                var snapshot = LoadUnlocked();
                var entry = new StatusEntry
                {
                    LastUpdateLocal = DateTime.Now,
                    Status = status,
                    Detail = detail ?? string.Empty
                };

                if (string.Equals(item, "Holiday", StringComparison.OrdinalIgnoreCase))
                    snapshot.Holiday = entry;
                else if (string.Equals(item, "ETTJ", StringComparison.OrdinalIgnoreCase))
                    snapshot.ETTJ = entry;

                Directory.CreateDirectory(EttjSettings.AppDataDir);
                File.WriteAllText(StatusFile, JsonConvert.SerializeObject(snapshot, Formatting.Indented));
            }
        }

        private static StatusSnapshot LoadUnlocked()
        {
            try
            {
                if (File.Exists(StatusFile))
                {
                    var json = File.ReadAllText(StatusFile);
                    var snapshot = JsonConvert.DeserializeObject<StatusSnapshot>(json);
                    if (snapshot != null)
                    {
                        if (snapshot.Holiday == null)
                            snapshot.Holiday = StatusEntry.Empty();
                        if (snapshot.ETTJ == null)
                            snapshot.ETTJ = StatusEntry.Empty();
                        return snapshot;
                    }
                }
            }
            catch
            {
                // Ignore corrupt status and show an empty dashboard.
            }

            return new StatusSnapshot
            {
                Holiday = StatusEntry.Empty(),
                ETTJ = StatusEntry.Empty()
            };
        }
    }

    internal sealed class StatusSnapshot
    {
        public StatusEntry Holiday { get; set; }
        public StatusEntry ETTJ { get; set; }
    }

    internal sealed class StatusEntry
    {
        public DateTime? LastUpdateLocal { get; set; }
        public string Status { get; set; }
        public string Detail { get; set; }

        internal static StatusEntry Empty()
        {
            return new StatusEntry
            {
                LastUpdateLocal = null,
                Status = "Never",
                Detail = string.Empty
            };
        }
    }
}
