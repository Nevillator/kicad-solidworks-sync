using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace KiCadSync
{
    /// <summary>
    /// Handles reading/writing the sync directory and driving SolidWorks geometry operations.
    /// </summary>
    public class SyncManager
    {
        private readonly ISldWorks _swApp;
        private readonly string _syncDir;

        public SyncManager(ISldWorks swApp, string syncDir)
        {
            _swApp = swApp;
            _syncDir = syncDir;
        }

        // ── Pull (KiCad → SolidWorks) ─────────────────────────────────────────

        /// <summary>
        /// Import board.step from ecad_to_mcad/ into a new SolidWorks assembly.
        /// Returns the opened IModelDoc2 or null on failure.
        /// </summary>
        public IModelDoc2? PullFromKiCad(out List<ChangeRecord> changes)
        {
            changes = new List<ChangeRecord>();
            var stepPath = Path.Combine(_syncDir, "ecad_to_mcad", "board.step");
            if (!File.Exists(stepPath))
                throw new FileNotFoundException("No KiCad STEP file found.", stepPath);

            int errors = 0;
            var importData = _swApp.GetImportFileData(stepPath);
            var doc = _swApp.LoadFile4(stepPath, "", importData, ref errors);

            if (doc == null)
                throw new Exception($"SolidWorks failed to import STEP (error {errors}).");

            changes.Add(new ChangeRecord { Type = "3d_model_updated" });

            // TODO: apply layout.json component positions to the assembly
            var layoutPath = Path.Combine(_syncDir, "ecad_to_mcad", "layout.json");
            if (File.Exists(layoutPath))
                changes.AddRange(ReadLayoutChanges(layoutPath));

            return doc;
        }

        private IEnumerable<ChangeRecord> ReadLayoutChanges(string layoutPath)
        {
            var layout = JObject.Parse(File.ReadAllText(layoutPath));
            var components = layout["components"] as JArray ?? new JArray();
            foreach (var c in components)
                yield return new ChangeRecord
                {
                    Type = "component_added",
                    Ref = c["ref"]?.ToString(),
                };
        }

        // ── Push (SolidWorks → KiCad) ─────────────────────────────────────────

        /// <summary>
        /// Export board outline and component positions from the active SolidWorks assembly
        /// to mcad_to_ecad/. Returns the list of change records.
        /// </summary>
        public List<ChangeRecord> PushToKiCad(string comment)
        {
            var doc = _swApp.ActiveDoc as IModelDoc2
                ?? throw new InvalidOperationException("No active SolidWorks document.");

            var outDir = Path.Combine(_syncDir, "mcad_to_ecad");
            Directory.CreateDirectory(outDir);

            var changes = new List<ChangeRecord>();

            // Export board outline
            var outline = ExtractBoardOutline(doc);
            File.WriteAllText(Path.Combine(outDir, "board_outline.json"),
                JsonConvert.SerializeObject(outline, Formatting.Indented));
            changes.Add(new ChangeRecord { Type = "board_outline_updated" });

            // Export component moves
            var moves = ExtractComponentPositions(doc);
            File.WriteAllText(Path.Combine(outDir, "component_moves.json"),
                JsonConvert.SerializeObject(moves, Formatting.Indented));
            foreach (var m in moves.Components)
                changes.Add(new ChangeRecord { Type = "component_moved", Ref = m.Ref });

            // Update manifest
            ManifestHelper.RecordPush(_syncDir, "mcad_to_ecad", "SolidWorks", comment, changes);

            return changes;
        }

        // ── Geometry extraction (stubs — implement with SW API) ───────────────

        private BoardOutlineData ExtractBoardOutline(IModelDoc2 doc)
        {
            // TODO: traverse Edge_Cuts sketch in the SolidWorks part and extract
            // line/arc segments. For now return an empty outline as a placeholder.
            return new BoardOutlineData
            {
                SchemaVersion = "1.0",
                Segments = new List<object>(),
                Holes = new List<HoleData>(),
            };
        }

        private ComponentMovesData ExtractComponentPositions(IModelDoc2 doc)
        {
            // TODO: enumerate components in the assembly, read transform matrices,
            // and convert to KiCad coordinate space.
            return new ComponentMovesData
            {
                SchemaVersion = "1.0",
                Components = new List<ComponentMove>(),
            };
        }
    }

    // ── Data models ───────────────────────────────────────────────────────────

    public class ChangeRecord
    {
        [JsonProperty("type")]  public string Type { get; set; } = "";
        [JsonProperty("ref")]   public string? Ref { get; set; }
        [JsonProperty("from")]  public object? From { get; set; }
        [JsonProperty("to")]    public object? To { get; set; }
    }

    public class BoardOutlineData
    {
        [JsonProperty("schema_version")] public string SchemaVersion { get; set; } = "1.0";
        [JsonProperty("segments")]        public List<object> Segments { get; set; } = new();
        [JsonProperty("holes")]           public List<HoleData> Holes { get; set; } = new();
    }

    public class HoleData
    {
        [JsonProperty("type")]        public string Type { get; set; } = "round";
        [JsonProperty("center")]      public PointData Center { get; set; } = new();
        [JsonProperty("diameter_mm")] public double DiameterMm { get; set; }
    }

    public class PointData
    {
        [JsonProperty("x_mm")] public double X { get; set; }
        [JsonProperty("y_mm")] public double Y { get; set; }
    }

    public class ComponentMovesData
    {
        [JsonProperty("schema_version")] public string SchemaVersion { get; set; } = "1.0";
        [JsonProperty("components")]      public List<ComponentMove> Components { get; set; } = new();
    }

    public class ComponentMove
    {
        [JsonProperty("ref")]      public string Ref { get; set; } = "";
        [JsonProperty("position")] public ComponentPosition Position { get; set; } = new();
    }

    public class ComponentPosition
    {
        [JsonProperty("x_mm")]          public double X { get; set; }
        [JsonProperty("y_mm")]          public double Y { get; set; }
        [JsonProperty("rotation_deg")]  public double Rotation { get; set; }
    }
}
