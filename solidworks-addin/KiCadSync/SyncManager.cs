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
    /// Orchestrates sync operations between the sync directory and SolidWorks.
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
        /// Build a native SolidWorks assembly from KiCad sync data.
        /// Creates a parametric board part + imports component STEPs.
        /// Returns the opened assembly IModelDoc2.
        /// </summary>
        public IModelDoc2? PullFromKiCad(out List<ChangeRecord> changes)
        {
            changes = new List<ChangeRecord>();

            var outlinePath = Path.Combine(_syncDir, "ecad_to_mcad", "board_outline.json");
            var layoutPath = Path.Combine(_syncDir, "ecad_to_mcad", "layout.json");
            var stepPath = Path.Combine(_syncDir, "ecad_to_mcad", "board.step");

            if (!File.Exists(outlinePath))
                throw new FileNotFoundException("No board outline data from KiCad.", outlinePath);
            if (!File.Exists(layoutPath))
                throw new FileNotFoundException("No layout data from KiCad.", layoutPath);

            var layout = JObject.Parse(File.ReadAllText(layoutPath));
            var thicknessMm = (double)(layout["board"]?["thickness_mm"] ?? 1.6);

            // ── Step 1: Create native board part from outline ────────────────

            var saveDir = Path.Combine(_syncDir, "sw_working");
            Directory.CreateDirectory(saveDir);

            var builder = new BoardBuilder(_swApp);
            var boardPartPath = builder.CreateBoardPart(outlinePath, thicknessMm, saveDir);
            changes.Add(new ChangeRecord { Type = "board_outline_updated" });

            // ── Step 2: Import STEP to get component 3D models ──────────────

            IModelDoc2? stepDoc = null;
            if (File.Exists(stepPath))
            {
                int errors = 0;
                var importData = _swApp.GetImportFileData(stepPath);
                stepDoc = _swApp.LoadFile4(stepPath, "", importData, ref errors);
            }

            // ── Step 3: Build assembly ──────────────────────────────────────

            var assemblyPath = Path.Combine(saveDir, "Board_Assembly.sldasm");
            var assemblyDoc = _CreateAssembly(boardPartPath, stepDoc, layout, assemblyPath);

            // Close the temporary STEP import if we opened one
            if (stepDoc != null)
                _swApp.CloseDoc(((IModelDoc2)stepDoc).GetTitle());

            changes.Add(new ChangeRecord { Type = "3d_model_updated" });

            return assemblyDoc;
        }

        // ── Push (SolidWorks → KiCad) ─────────────────────────────────────────

        /// <summary>
        /// Export board outline and component positions from the active SolidWorks
        /// assembly to mcad_to_ecad/.
        /// </summary>
        public List<ChangeRecord> PushToKiCad(string comment)
        {
            var doc = _swApp.ActiveDoc as IModelDoc2
                ?? throw new InvalidOperationException("No active SolidWorks document.");

            var outDir = Path.Combine(_syncDir, "mcad_to_ecad");
            Directory.CreateDirectory(outDir);
            var changes = new List<ChangeRecord>();

            var layoutPath = Path.Combine(_syncDir, "ecad_to_mcad", "layout.json");

            // ── Extract board outline from the native part's sketch ─────────

            var boardDoc = _FindBoardPartDoc(doc);
            if (boardDoc != null)
            {
                var reader = new BoardReader();
                var outline = reader.ReadOutline(boardDoc);
                File.WriteAllText(
                    Path.Combine(outDir, "board_outline.json"),
                    JsonConvert.SerializeObject(outline, Formatting.Indented));
                changes.Add(new ChangeRecord { Type = "board_outline_updated" });
            }

            // ── Extract component positions ─────────────────────────────────

            if (doc is IAssemblyDoc assembly && File.Exists(layoutPath))
            {
                var map = ComponentMapper.BuildMap(assembly, layoutPath);
                var moves = ComponentMapper.ExtractPositions(map);
                File.WriteAllText(
                    Path.Combine(outDir, "component_moves.json"),
                    JsonConvert.SerializeObject(moves, Formatting.Indented));
                foreach (var m in moves.Components)
                    changes.Add(new ChangeRecord { Type = "component_moved", Ref = m.Ref });
            }

            ManifestHelper.RecordPush(_syncDir, "mcad_to_ecad", "SolidWorks", comment, changes);
            return changes;
        }

        // ── Assembly creation ───────────────────────────────────────────────────

        private IModelDoc2? _CreateAssembly(string boardPartPath, IModelDoc2? stepDoc,
            JObject layout, string assemblyPath)
        {
            // Create new assembly
            var asmDoc = _swApp.NewDocument(
                _swApp.GetUserPreferenceStringValue(
                    (int)swUserPreferenceStringValue_e.swDefaultTemplateAssembly),
                0, 0, 0) as IModelDoc2;
            if (asmDoc == null) return null;

            var assembly = asmDoc as IAssemblyDoc;

            // Add the board part as the first (fixed) component
            assembly.AddComponent5(boardPartPath, 0, "", false, "", 0, 0, 0);

            // If we have the STEP import open, we can extract individual component
            // bodies and save them as separate parts, then add to the assembly.
            // For now, add the entire STEP as a single component — the ComponentMapper
            // will handle identifying individual parts within it.
            if (stepDoc != null)
            {
                var stepTitle = stepDoc.GetTitle() as string;
                var stepPathOnDisk = stepDoc.GetPathName();
                if (!string.IsNullOrEmpty(stepPathOnDisk))
                    assembly.AddComponent5(stepPathOnDisk, 0, "", false, "", 0, 0, 0);
            }

            // Save
            int err = 0, warn = 0;
            asmDoc.Extension.SaveAs3(assemblyPath,
                (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                null, null, ref err, ref warn);

            return asmDoc;
        }

        // ── Helpers ─────────────────────────────────────────────────────────────

        private IModelDoc2? _FindBoardPartDoc(IModelDoc2 doc)
        {
            // If the active doc is a part, assume it's the board
            if (doc.GetType() == (int)swDocumentTypes_e.swDocPART)
                return doc;

            // If it's an assembly, find the board component
            if (doc is IAssemblyDoc assembly)
            {
                var layoutPath = Path.Combine(_syncDir, "ecad_to_mcad", "layout.json");
                if (File.Exists(layoutPath))
                {
                    var layout = JObject.Parse(File.ReadAllText(layoutPath));
                    var boardFileName = layout["board"]?["file_name"]?.ToString() ?? "";

                    var components = assembly.GetComponents(true) as object[];
                    if (components != null)
                    {
                        foreach (IComponent2 comp in components)
                        {
                            var name = comp.Name2 ?? "";
                            // Board component matches the KiCad PCB filename
                            if (!string.IsNullOrEmpty(boardFileName) &&
                                name.Contains(boardFileName, StringComparison.OrdinalIgnoreCase))
                            {
                                return comp.GetModelDoc2() as IModelDoc2;
                            }
                            // Fallback: look for "Board" in the name
                            if (name.Contains("Board", StringComparison.OrdinalIgnoreCase))
                            {
                                return comp.GetModelDoc2() as IModelDoc2;
                            }
                        }
                    }
                }
            }

            return null;
        }
    }

    // ── Shared data models ──────────────────────────────────────────────────

    public class ChangeRecord
    {
        [JsonProperty("type")]  public string Type { get; set; } = "";
        [JsonProperty("ref")]   public string? Ref { get; set; }
        [JsonProperty("from")]  public object? From { get; set; }
        [JsonProperty("to")]    public object? To { get; set; }
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
