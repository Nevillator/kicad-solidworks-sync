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

            SwAddin.Log("Pull: checking files...");
            if (!File.Exists(outlinePath))
                throw new FileNotFoundException("No board outline data from KiCad.", outlinePath);
            if (!File.Exists(layoutPath))
                throw new FileNotFoundException("No layout data from KiCad.", layoutPath);

            var layout = JObject.Parse(File.ReadAllText(layoutPath));
            var thicknessMm = (double)(layout["board"]?["thickness_mm"] ?? 1.6);
            SwAddin.Log($"Pull: thickness={thicknessMm}mm");

            // ── Step 1: Create native board part from outline ────────────────

            var saveDir = Path.Combine(_syncDir, "sw_working");
            Directory.CreateDirectory(saveDir);

            SwAddin.Log("Pull: creating board part...");
            var builder = new BoardBuilder(_swApp);
            var boardPartPath = builder.CreateBoardPart(outlinePath, thicknessMm, saveDir);
            SwAddin.Log($"Pull: board part saved to {boardPartPath}");
            changes.Add(new ChangeRecord { Type = "board_outline_updated" });

            // ── Step 2: Import STEP to get component 3D models ──────────────

            IModelDoc2? stepDoc = null;
            if (File.Exists(stepPath))
            {
                SwAddin.Log("Pull: importing STEP...");
                int errors = 0;
                var importData = _swApp.GetImportFileData(stepPath);
                SwAddin.Log($"Pull: importData={importData != null}");
                stepDoc = _swApp.LoadFile4(stepPath, "", importData, ref errors);
                SwAddin.Log($"Pull: STEP loaded, errors={errors}, doc={stepDoc != null}");
            }

            // ── Step 3: Build assembly ──────────────────────────────────────

            SwAddin.Log("Pull: creating assembly...");
            var assemblyPath = Path.Combine(saveDir, "Board_Assembly.sldasm");
            var assemblyDoc = _CreateAssembly(boardPartPath, stepDoc, layout, assemblyPath, saveDir, thicknessMm);
            SwAddin.Log($"Pull: assembly created={assemblyDoc != null}");

            // Close the temporary STEP import if we opened one
            if (stepDoc != null)
            {
                var stepTitle = stepDoc.GetTitle() as string;
                SwAddin.Log($"Pull: closing STEP doc '{stepTitle}'");
                if (!string.IsNullOrEmpty(stepTitle))
                    _swApp.CloseDoc(stepTitle);
            }

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
            JObject layout, string assemblyPath, string saveDir, double thicknessMm)
        {
            // Create new assembly
            var asmDoc = _swApp.NewDocument(
                _swApp.GetUserPreferenceStringValue(
                    (int)swUserPreferenceStringValue_e.swDefaultTemplateAssembly),
                0, 0, 0) as IModelDoc2;
            if (asmDoc == null) return null;

            var assembly = asmDoc as IAssemblyDoc;

            // Board part: top face at Z=0, bottom at Z=-thickness (matches KiCad STEP convention).
            assembly.AddComponent5(boardPartPath, 0, "", false, "", 0, 0, 0);

            // STEP sub-assembly: KiCad exports with IDF convention (B.Cu/bottom at Z=0,
            // F.Cu/top at Z=+thickness). Offset by -thickness so the bottom face aligns
            // with our .sldprt bottom at Z=-thickness.
            if (stepDoc != null)
            {
                var stepSavePath = Path.Combine(saveDir, "Components.sldasm");
                int stepErr = 0, stepWarn = 0;
                SwAddin.Log($"Pull: saving STEP import to {stepSavePath}");

                stepDoc.Extension.SaveAs3(stepSavePath,
                    (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                    (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                    null, null, ref stepErr, ref stepWarn);
                SwAddin.Log($"Pull: STEP saved, err={stepErr}, warn={stepWarn}");

                double zOffsetM = -thicknessMm / 1000.0;
                SwAddin.Log($"Pull: placing STEP at Z={zOffsetM * 1000:F3}mm");
                assembly.AddComponent5(stepSavePath, 0, "", false, "", 0, 0, zOffsetM);
                SwAddin.Log("Pull: STEP sub-assembly added to main assembly");
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
                                name.IndexOf(boardFileName, StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                return comp.GetModelDoc2() as IModelDoc2;
                            }
                            // Fallback: look for "Board" in the name
                            if (name.IndexOf("Board", StringComparison.OrdinalIgnoreCase) >= 0)
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
