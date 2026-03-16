using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace KiCadSync
{
    public static class ManifestHelper
    {
        private static string ManifestPath(string syncDir) =>
            Path.Combine(syncDir, "manifest.json");

        public static JObject Load(string syncDir)
        {
            var path = ManifestPath(syncDir);
            if (!File.Exists(path))
                return new JObject(
                    new JProperty("schema_version", "1.0"),
                    new JProperty("history", new JArray()));
            return JObject.Parse(File.ReadAllText(path));
        }

        public static void RecordPush(string syncDir, string direction, string author,
            string comment, IEnumerable<ChangeRecord> changes)
        {
            var manifest = Load(syncDir);

            var record = new JObject(
                new JProperty("id", Guid.NewGuid().ToString()),
                new JProperty("timestamp", DateTimeOffset.UtcNow.ToString("o")),
                new JProperty("direction", direction),
                new JProperty("author", author),
                new JProperty("comment", comment),
                new JProperty("changes", JArray.FromObject(changes)));

            manifest["last_push"] = record;
            (manifest["history"] as JArray ?? new JArray()).Add(record);

            File.WriteAllText(ManifestPath(syncDir),
                manifest.ToString(Formatting.Indented));
        }

        public static JObject? GetLastPush(string syncDir) =>
            Load(syncDir)["last_push"] as JObject;
    }
}
