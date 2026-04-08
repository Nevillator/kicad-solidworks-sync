using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
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
        private static readonly StringComparison OIC = StringComparison.OrdinalIgnoreCase;

        public SyncManager(ISldWorks swApp, string syncDir)
        {
            _swApp = swApp;
            _syncDir = syncDir;
        }

        // ── Pull (KiCad → SolidWorks) ─────────────────────────────────────────

        /// <summary>
        /// Build or update a native SolidWorks assembly from KiCad sync data.
        /// On first pull, prompts the user to choose where to save the assembly.
        /// On subsequent pulls, locates the existing assembly from the registry
        /// and updates it in place.
        /// </summary>
        public IModelDoc2? PullFromKiCad(out List<ChangeRecord> changes,
            IProgress<string>? progress = null, bool hideImport = false)
        {
            changes = new List<ChangeRecord>();

            var outlinePath = Path.Combine(_syncDir, "ecad_to_mcad", "board_outline.json");
            var layoutPath  = Path.Combine(_syncDir, "ecad_to_mcad", "layout.json");

            SwAddin.Log("Pull: checking files...");
            if (!File.Exists(outlinePath))
                throw new FileNotFoundException("No board outline data from KiCad.", outlinePath);
            if (!File.Exists(layoutPath))
                throw new FileNotFoundException("No layout data from KiCad.", layoutPath);

            var layout      = JObject.Parse(File.ReadAllText(layoutPath));
            var thicknessMm = (double)(layout["board"]?["thickness_mm"] ?? 1.6);
            var boardName   = layout["board"]?["file_name"]?.ToString() ?? "Board";

            SwAddin.Log($"Pull: boardName={boardName}, thickness={thicknessMm}mm");

            var saveDir = Path.Combine(_syncDir, "sw_working");
            Directory.CreateDirectory(saveDir);

            var registry = _LoadRegistry(saveDir);
            var entry    = _FindRegistryEntry(registry, boardName);

            if (entry != null)
            {
                // ── Update existing board ─────────────────────────────────────
                var assemblyPath  = entry["sldasm"]?.ToString() ?? "";
                var boardPartPath = entry["sldprt"]?.ToString() ?? "";
                SwAddin.Log($"Pull: update mode — assembly={assemblyPath}");

                // Open assembly BEFORE blanking so its window stays visible
                var asmDoc = _GetOpenDoc(assemblyPath);
                if (asmDoc == null)
                {
                    SwAddin.Log("Pull: opening existing assembly...");
                    int oe = 0, ow = 0;
                    asmDoc = _swApp.OpenDoc6(assemblyPath,
                        (int)swDocumentTypes_e.swDocASSEMBLY,
                        (int)swOpenDocOptions_e.swOpenDocOptions_Silent,
                        "", ref oe, ref ow) as IModelDoc2;
                }
                if (asmDoc == null)
                    throw new Exception($"Failed to open assembly: {assemblyPath}");

                if (hideImport)
                {
                    _swApp.DocumentVisible(false, (int)swDocumentTypes_e.swDocPART);
                    _swApp.UserControl = false;
                }
                try
                {
                    _UpdateExistingBoard(asmDoc, boardName, boardPartPath, assemblyPath,
                        outlinePath, layout, saveDir, thicknessMm, progress, hideImport);
                }
                finally
                {
                    if (hideImport)
                    {
                        _swApp.DocumentVisible(true, (int)swDocumentTypes_e.swDocPART);
                        _swApp.UserControl = true;
                    }
                }

                changes.Add(new ChangeRecord { Type = "board_outline_updated" });
                changes.Add(new ChangeRecord { Type = "3d_model_updated" });
                return asmDoc;
            }
            else
            {
                // ── First pull: prompt user for save location ─────────────────
                // Show save dialog BEFORE blanking the graphics area
                string? assemblyPath = null;
                using (var dlg = new SaveFileDialog
                {
                    Title      = $"Save Board Assembly — {boardName}",
                    Filter     = "SolidWorks Assembly (*.sldasm)|*.sldasm",
                    FileName   = $"{boardName}.sldasm",
                    DefaultExt = "sldasm",
                })
                {
                    if (dlg.ShowDialog() != DialogResult.OK)
                    {
                        SwAddin.Log("Pull: user cancelled save dialog.");
                        return null;
                    }
                    assemblyPath = dlg.FileName;
                }

                var boardPartPath = Path.Combine(saveDir, $"{boardName}_board.sldprt");
                SwAddin.Log($"Pull: first pull → {assemblyPath}");
                _CloseDocIfOpen(assemblyPath);
                _CloseDocIfOpen(boardPartPath);

                if (hideImport)
                {
                    _swApp.DocumentVisible(false, (int)swDocumentTypes_e.swDocPART);
                    _swApp.UserControl = false;
                }
                IModelDoc2? shown = null;
                bool success = false;
                try
                {
                    progress?.Report("Building board...");
                    SwAddin.Log("Pull: creating board part...");
                    var builder = new BoardBuilder(_swApp);
                    boardPartPath = builder.CreateBoardPart(outlinePath, thicknessMm, saveDir,
                        $"{boardName}_PCB");
                    SwAddin.Log($"Pull: board part → {boardPartPath}");
                    changes.Add(new ChangeRecord { Type = "board_outline_updated" });

                    progress?.Report("Creating assembly...");
                    SwAddin.Log("Pull: creating assembly...");
                    shown = _CreateAssembly(boardName, boardPartPath, layout,
                        assemblyPath, saveDir, thicknessMm, progress, hideImport);
                    SwAddin.Log("Pull: assembly created.");
                    changes.Add(new ChangeRecord { Type = "3d_model_updated" });

                    _RegisterBoard(registry, boardName, assemblyPath, boardPartPath);
                    _SaveRegistry(saveDir, registry);
                    success = true;
                }
                finally
                {
                    if (hideImport)
                    {
                        _swApp.DocumentVisible(true, (int)swDocumentTypes_e.swDocPART);
                        _swApp.UserControl = true;
                    }
                }

                if (!success) return null;
                SwAddin.Log($"Pull: assembly visible={shown != null}");
                return shown;
            }
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
                var map   = ComponentMapper.BuildMap(assembly, layoutPath);
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

        // ── Assembly creation (first pull) ──────────────────────────────────────

        private IModelDoc2 _CreateAssembly(string boardName, string boardPartPath,
            JObject layout, string assemblyPath, string saveDir, double thicknessMm,
            IProgress<string>? progress = null, bool hideImport = false)
        {
            // Step 1: Set import preferences, enable 3D Interconnect
            bool prev3DI = (bool)_swApp.GetUserPreferenceToggle(
                (int)swUserPreferenceToggle_e.swMultiCAD_Enable3DInterconnect);
            bool prevAnalytic = (bool)_swApp.GetUserPreferenceToggle(
                (int)swUserPreferenceToggle_e.swImportNeutralAnalyticalConversion);
            int prevAsmMap = _swApp.GetUserPreferenceIntegerValue(
                (int)swUserPreferenceIntegerValue_e.swImportNeutralAssemblyStructureMapping);
            _swApp.SetUserPreferenceToggle(
                (int)swUserPreferenceToggle_e.swMultiCAD_Enable3DInterconnect, true);
            _swApp.SetUserPreferenceToggle(
                (int)swUserPreferenceToggle_e.swImportNeutralAnalyticalConversion, false);
            _swApp.SetUserPreferenceIntegerValue(
                (int)swUserPreferenceIntegerValue_e.swImportNeutralAssemblyStructureMapping,
                2); // swImportNeutralAssemblyStructureMapping_e.swImportNeutralAssemblyStructureMapping_MultiBodyPart

            // Step 2: Build transform map from layout.json
            var transformMap = _BuildLayoutTransformMap(layout, thicknessMm);
            SwAddin.Log($"Pull: {transformMap.Count} components in layout");
            var total = transformMap.Count;

            // Step 3: Create fresh assembly
            var asmTemplate = _swApp.GetUserPreferenceStringValue(
                (int)swUserPreferenceStringValue_e.swDefaultTemplateAssembly);
            var asmDoc = _swApp.NewDocument(asmTemplate, 0, 0, 0) as IModelDoc2
                ?? throw new Exception("Pull: failed to create new assembly document.");
            var assembly = asmDoc as IAssemblyDoc
                ?? throw new Exception("Pull: new document is not an assembly.");

            int activateErr = 0;
            _swApp.ActivateDoc3(asmDoc.GetTitle(), true,
                (int)swRebuildOnActivation_e.swRebuildActiveDoc, ref activateErr);
            var activeAsm = _swApp.ActiveDoc as IAssemblyDoc ?? assembly;
            activeAsm.EditAssembly();

            var mathUtil    = _swApp.GetMathUtility() as IMathUtility;

            // Step 4: Insert board part, then pin its part origin to the assembly origin
            var boardResult = activeAsm.AddComponent5(boardPartPath, 0, "", false, "", 0, 0, 0);
            SwAddin.Log($"Pull: board AddComponent5={boardResult != null}");
            if (boardResult is IComponent2 boardComp && mathUtil != null)
            {
                // Translate board +thickness in Z so bottom face lands at assembly Z=0
                var boardData = new double[] { 1,0,0, 0,1,0, 0,0,1, 0, 0, thicknessMm/1000.0, 1 };
                var boardXform = mathUtil.CreateTransform(boardData) as MathTransform;
                if (boardXform != null) boardComp.Transform2 = boardXform;
            }
            _swApp.CloseDoc(Path.GetFileNameWithoutExtension(boardPartPath));

            // Step 5: Group placements by model, insert once per unique STEP, reuse for copies
            var compStepDir = Path.Combine(_syncDir, "ecad_to_mcad", "components");
            var refMap      = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            // Group: baseName → list of (kicadRef, xform)
            var groups = new Dictionary<string, List<(string kicadRef, double[]? xform)>>(
                StringComparer.OrdinalIgnoreCase);
            foreach (var kvp in transformMap)
            {
                if (!groups.TryGetValue(kvp.Value.baseName, out var list))
                    groups[kvp.Value.baseName] = list = new List<(string, double[]?)>();
                list.Add((kvp.Key, kvp.Value.xform));
            }

            int done = 0;
            foreach (var grp in groups)
            {
                var baseName  = grp.Key;
                var placements = grp.Value;
                var stepPath  = Path.Combine(compStepDir, $"{baseName}.step");

                if (!File.Exists(stepPath))
                {
                    SwAddin.Log($"  SKIP '{baseName}': no STEP");
                    done += placements.Count;
                    continue;
                }

                // First placement: InsertImportedComponent to load the STEP
                var (firstRef, firstXform) = placements[0];
                done++;
                progress?.Report($"Inserting {firstRef} ({done}/{total})...");
                activeAsm.InsertImportedComponent(stepPath, 0, 0, 0, out var firstObj);
                SwAddin.Log($"  InsertImportedComponent '{baseName}' ({firstRef}): {(firstObj != null ? "ok" : "null")}");

                string? loadedPath = null;
                if (firstObj is IComponent2 firstComp)
                {
                    refMap[firstRef] = firstComp.Name2 ?? "";
                    // Get the path of the now-loaded part for AddComponent5 reuse
                    loadedPath = firstComp.GetPathName();
                    if (firstXform != null && mathUtil != null)
                    {
                        var xf = mathUtil.CreateTransform(firstXform) as MathTransform;
                        if (xf != null) firstComp.Transform2 = xf;
                    }
                }

                // Subsequent placements: AddComponent5 reuses the in-memory part
                for (int i = 1; i < placements.Count; i++)
                {
                    var (kicadRef, xformData) = placements[i];
                    done++;
                    progress?.Report($"Inserting {kicadRef} ({done}/{total})...");

                    if (loadedPath == null)
                    {
                        SwAddin.Log($"  SKIP {kicadRef}: no loaded path for '{baseName}'");
                        continue;
                    }

                    var addedComp = activeAsm.AddComponent5(loadedPath, 0, "", false, "", 0, 0, 0)
                        as IComponent2;
                    SwAddin.Log($"  AddComponent5 copy '{kicadRef}': {(addedComp != null ? "ok" : "null")}");

                    if (addedComp != null)
                    {
                        refMap[kicadRef] = addedComp.Name2 ?? "";
                        if (xformData != null && mathUtil != null)
                        {
                            var xf = mathUtil.CreateTransform(xformData) as MathTransform;
                            if (xf != null) addedComp.Transform2 = xf;
                        }
                    }
                }
            }

            // Step 6: Restore import preferences
            _swApp.SetUserPreferenceToggle(
                (int)swUserPreferenceToggle_e.swMultiCAD_Enable3DInterconnect, prev3DI);
            _swApp.SetUserPreferenceToggle(
                (int)swUserPreferenceToggle_e.swImportNeutralAnalyticalConversion, prevAnalytic);
            _swApp.SetUserPreferenceIntegerValue(
                (int)swUserPreferenceIntegerValue_e.swImportNeutralAssemblyStructureMapping, prevAsmMap);

            // Step 7: Save
            progress?.Report("Saving...");
            int e = 0, w = 0;
            asmDoc.Extension.SaveAs3(assemblyPath,
                (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                null, null, ref e, ref w);
            SwAddin.Log($"Pull: assembly saved to '{Path.GetFileName(assemblyPath)}', err={e}");
            _SaveRefMap(saveDir, refMap);

            // Step 8: If hidden, close and reopen so it appears in the graphics area
            if (hideImport)
            {
                _swApp.CloseDoc(Path.GetFileNameWithoutExtension(assemblyPath));
                int e2 = 0, w2 = 0;
                var shown = _swApp.OpenDoc6(assemblyPath,
                    (int)swDocumentTypes_e.swDocASSEMBLY,
                    (int)swOpenDocOptions_e.swOpenDocOptions_Silent,
                    "", ref e2, ref w2) as IModelDoc2;
                return shown ?? throw new Exception("Pull: failed to reopen assembly after hide.");
            }
            return asmDoc;
        }

        // ── In-place update ─────────────────────────────────────────────────────

        private void _UpdateExistingBoard(IModelDoc2 asmDoc, string boardName,
            string boardPartPath, string assemblyPath, string outlinePath,
            JObject layout, string saveDir, double thicknessMm,
            IProgress<string>? progress = null, bool hideImport = false)
        {
            var assembly = asmDoc as IAssemblyDoc
                ?? throw new InvalidOperationException("Assembly doc is not an IAssemblyDoc");

            // Step 1 — Update board part in place (unchanged)
            var boardDoc = _GetOpenDoc(boardPartPath);
            if (boardDoc == null)
            {
                int oe = 0, ow = 0;
                boardDoc = _swApp.OpenDoc6(boardPartPath,
                    (int)swDocumentTypes_e.swDocPART,
                    (int)swOpenDocOptions_e.swOpenDocOptions_Silent,
                    "", ref oe, ref ow) as IModelDoc2;
            }
            if (boardDoc != null)
            {
                int activateErr = 0;
                _swApp.ActivateDoc3(boardDoc.GetTitle(), true,
                    (int)swRebuildOnActivation_e.swRebuildActiveDoc, ref activateErr);
                progress?.Report("Updating board...");
                SwAddin.Log("Pull: updating board part in place...");
                new BoardBuilder(_swApp).UpdateBoardPart(boardDoc, outlinePath, thicknessMm);
                _swApp.ActivateDoc3(asmDoc.GetTitle(), true,
                    (int)swRebuildOnActivation_e.swRebuildActiveDoc, ref activateErr);
            }
            else
            {
                SwAddin.Log("Pull: WARN — board part not found, recreating...");
                _CloseDocIfOpen(boardPartPath);
                var kicadDir = Path.GetDirectoryName(boardPartPath)!;
                new BoardBuilder(_swApp).CreateBoardPart(outlinePath, thicknessMm, kicadDir,
                    Path.GetFileNameWithoutExtension(boardPartPath));
            }

            // Step 2 — Build transform map from layout.json
            progress?.Report("Updating component positions...");
            var newTransforms = _BuildLayoutTransformMap(layout, thicknessMm);
            SwAddin.Log($"Pull: {newTransforms.Count} components in layout");

            // Step 4 — Build existing component map using saved ref→swName mapping
            var refMap      = _LoadRefMap(saveDir);
            var existingMap = new Dictionary<string, IComponent2>(StringComparer.OrdinalIgnoreCase);
            foreach (var kvp in refMap)
            {
                var comp = _FindComponentByName(assembly, kvp.Value);
                if (comp != null) existingMap[kvp.Key] = comp;
            }
            SwAddin.Log($"Pull: {existingMap.Count} existing components mapped by ref");

            var mathUtil = _swApp.GetMathUtility() as IMathUtility;

            // Step 5 — Update transforms for existing components; suppress deleted ones
            foreach (var kvp in existingMap)
            {
                if (newTransforms.TryGetValue(kvp.Key, out var entry) &&
                    entry.xform != null && mathUtil != null)
                {
                    var xform = mathUtil.CreateTransform(entry.xform) as MathTransform;
                    if (xform != null) kvp.Value.SetTransformAndSolve2(xform);
                    SwAddin.Log($"  transform updated: {kvp.Key}");
                }
                else
                {
                    kvp.Value.SetSuppression2((int)swComponentSuppressionState_e.swComponentSuppressed);
                    refMap.Remove(kvp.Key);
                    SwAddin.Log($"  suppressed (removed): {kvp.Key}");
                }
            }

            // Step 6 — Insert genuinely new components
            assembly.EditAssembly();
            var compRepoDir = Path.Combine(_syncDir, "sw_working", "components");
            Directory.CreateDirectory(compRepoDir);
            var newComponents = newTransforms.Where(k => !existingMap.ContainsKey(k.Key)).ToList();
            var newCount = newComponents.Count;
            var newIdx = 0;
            var compStepDir = Path.Combine(_syncDir, "ecad_to_mcad", "components");

            // Enable 3D Interconnect for InsertImportedComponent
            bool prev3DI = (bool)_swApp.GetUserPreferenceToggle(
                (int)swUserPreferenceToggle_e.swMultiCAD_Enable3DInterconnect);
            _swApp.SetUserPreferenceToggle(
                (int)swUserPreferenceToggle_e.swMultiCAD_Enable3DInterconnect, true);

            foreach (var kvp in newComponents)
            {
                newIdx++;
                var stepPath = Path.Combine(compStepDir, $"{_SanitizeName(kvp.Value.baseName)}.step");
                progress?.Report($"Inserting new component {newIdx}/{newCount}: {kvp.Key}...");

                if (!File.Exists(stepPath))
                {
                    SwAddin.Log($"  SKIP {kvp.Key}: no STEP at '{stepPath}'");
                    continue;
                }

                assembly.InsertImportedComponent(stepPath, 0, 0, 0, out var compObj);
                SwAddin.Log($"  InsertImportedComponent {kvp.Key}: {(compObj != null ? "ok" : "null")}");

                if (compObj is IComponent2 addedComp)
                {
                    refMap[kvp.Key] = addedComp.Name2 ?? "";
                    if (kvp.Value.xform != null && mathUtil != null)
                    {
                        var xform = mathUtil.CreateTransform(kvp.Value.xform) as MathTransform;
                        if (xform != null) addedComp.Transform2 = xform;
                    }
                }
            }

            _swApp.SetUserPreferenceToggle(
                (int)swUserPreferenceToggle_e.swMultiCAD_Enable3DInterconnect, prev3DI);

            _SaveRefMap(saveDir, refMap);
            progress?.Report("Saving...");
            _RebuildAndSave(asmDoc, assemblyPath);
        }

        private void _RebuildAndSave(IModelDoc2 doc, string path)
        {
            doc.ForceRebuild3(false);
            int err = 0, warn = 0;
            doc.Extension.SaveAs3(path,
                (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                null, null, ref err, ref warn);
            SwAddin.Log($"Pull: saved '{Path.GetFileName(path)}', err={err}, warn={warn}");
        }

        // ── Component repository helpers ─────────────────────────────────────────

        /// <summary>
        /// Snapshot current open document paths. STEP-imported docs have empty
        /// GetPathName(), so they are distinguishable from repo docs.
        /// </summary>
        private HashSet<string> _SnapshotPaths()
        {
            var set  = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var docs = _swApp.GetDocuments() as object[];
            if (docs != null)
                foreach (IModelDoc2 d in docs)
                    set.Add(d.GetPathName() ?? "");
            return set;
        }

        /// <summary>
        /// Close every document whose path was not in the snapshot (i.e. opened since
        /// the snapshot was taken). Iterates a local copy so CloseDoc is safe.
        /// </summary>
        private void _CloseDocsSince(HashSet<string> before)
        {
            var docs = _swApp.GetDocuments() as object[];
            if (docs == null) return;
            var toClose = docs.Cast<IModelDoc2>()
                .Where(d => !before.Contains(d.GetPathName() ?? ""))
                .Select(d => d.GetTitle())
                .ToList();
            foreach (var title in toClose)
                _swApp.CloseDoc(title);
        }

        /// <summary>
        /// Recursively save a STEP-imported doc tree (part or assembly, any depth) to
        /// the component repo directory. Saves sub-components before their parent so
        /// parent assembly references resolve correctly. Skips if dest file already exists.
        /// Returns the repo path for the root doc.
        /// </summary>
        private string _SaveDocTreeToRepo(IModelDoc2 doc, string baseName, string compRepoDir)
        {
            var isAsm    = doc is IAssemblyDoc;
            var ext      = isAsm ? ".sldasm" : ".sldprt";
            var destPath = Path.Combine(compRepoDir, $"{_SanitizeName(baseName)}{ext}");
            if (File.Exists(destPath)) return destPath;

            if (isAsm)
            {
                // Save sub-components first (bottom-up) so parent assembly can reference them
                var subComps = (doc as IAssemblyDoc)!.GetComponents(true) as object[];
                if (subComps != null)
                    foreach (IComponent2 sub in subComps)
                    {
                        var subDoc = sub.GetModelDoc2() as IModelDoc2;
                        if (subDoc == null) continue;
                        var subBase = _StripInstanceSuffix(sub.Name2 ?? "");
                        _SaveDocTreeToRepo(subDoc, subBase, compRepoDir);
                    }
            }

            int e = 0, w = 0;
            doc.Extension.SaveAs3(destPath,
                (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                null, null, ref e, ref w);
            SwAddin.Log($"  repo: saved '{baseName}'{ext}, err={e}");
            return destPath;
        }

        /// <summary>
        /// Import an individual component STEP from ecad_to_mcad/components/ and save its
        /// full doc tree to the repo. No-op if the repo file already exists (both .sldprt
        /// and .sldasm are checked). Returns the repo path, or empty string on failure.
        /// </summary>
        private string _ImportComponentToRepo(string baseName, string compRepoDir)
        {
            var prt = Path.Combine(compRepoDir, $"{_SanitizeName(baseName)}.sldprt");
            var asm = Path.Combine(compRepoDir, $"{_SanitizeName(baseName)}.sldasm");
            if (File.Exists(prt)) return prt;
            if (File.Exists(asm)) return asm;

            var stepPath = Path.Combine(_syncDir, "ecad_to_mcad", "components",
                                        $"{_SanitizeName(baseName)}.step");
            if (!File.Exists(stepPath))
            {
                SwAddin.Log($"  repo: no STEP for '{baseName}', skipping");
                return "";
            }

            var before = _SnapshotPaths();
            int errors = 0;
            var importData = _swApp.GetImportFileData(stepPath);
            var tempDoc    = _swApp.LoadFile4(stepPath, "", importData, ref errors) as IModelDoc2;
            if (tempDoc == null)
            {
                SwAddin.Log($"  repo: LoadFile4 failed for '{baseName}', err={errors}");
                _CloseDocsSince(before);
                return "";
            }

            var repoPath = _SaveDocTreeToRepo(tempDoc, baseName, compRepoDir);
            _CloseDocsSince(before);
            return repoPath;
        }

        /// <summary>
        /// Compute a 16-element SW IMathTransform ArrayData from layout.json component data.
        /// </summary>
        private static double[] _ComputeSwTransform(JToken comp, double thicknessMm)
        {
            const double DEG = Math.PI / 180.0;

            double fpX  = (double)(comp["position"]?["x_mm"]         ?? 0);
            double fpY  = (double)(comp["position"]?["y_mm"]         ?? 0);
            double fpRot= (double)(comp["position"]?["rotation_deg"]  ?? 0);
            double ox   = (double)(comp["model_offset"]?["x_mm"]     ?? 0);
            double oy   = (double)(comp["model_offset"]?["y_mm"]     ?? 0);
            double oz   = (double)(comp["model_offset"]?["z_mm"]     ?? 0);
            double mz   = (double)(comp["model_rotation"]?["z_deg"]  ?? 0);
            double mx   = (double)(comp["model_rotation"]?["x_deg"]  ?? 0);
            double my   = (double)(comp["model_rotation"]?["y_deg"]  ?? 0);

            bool isFront = string.Equals(comp["layer"]?.ToString(), "F.Cu",
                               StringComparison.OrdinalIgnoreCase);

            // ── Rotation ─────────────────────────────────────────────────────
            // KiCad model_rotation is intrinsic Rx→Ry→Rz applied in KiCad space
            // (X-right, Y-down screen, Z-up out-of-board).
            // Converting KiCad→SW by conjugating with F=diag(1,-1,1) (Y-flip):
            //   F·Rz(θ)·F = Rz(-θ),  F·Ry(θ)·F = Ry(θ),  F·Rx(θ)·F = Rx(-θ)
            // footprint rotation_deg is CCW in KiCad Y-down → CW in SW = negate.
            // Final SW rotation: Rz(-(fpRot+mz)) · Ry(my) · Rx(-mx)
            // KiCad applies: Rx(mx)·Ry(my)·Rz(mz) in model space, then Rz(fpRot) in board space.
            // Since board and SW share axes, fpRot is positive (CCW). Applied after model rotation
            // in model's local frame = right-multiply: Rx(mx)·Ry(my)·Rz(mz+fpRot)
            double[] R = _Mul3x3(_RotX3x3(mx * DEG),
                         _Mul3x3(_RotY3x3(my * DEG),
                                  _RotZ3x3((mz - fpRot) * DEG)));

            if (!isFront)
            {
                // B.Cu: post-multiply by Rx(180°) = diag(1,-1,-1) to flip onto bottom face
                R = _Mul3x3(R, new double[] { 1,0,0, 0,-1,0, 0,0,-1 });
            }

            // ── Translation ──────────────────────────────────────────────────
            // SW position = (fpX, -fpY) + Rz(fpRot) * (ox, oy)
            // Rz(fpRot) is a 2D rotation of the model_offset into board frame.
            double fpRotRad = fpRot * DEG;
            double cosF = Math.Cos(fpRotRad), sinF = Math.Sin(fpRotRad);
            double rotOx = cosF * ox - sinF * oy;
            double rotOy = sinF * ox + cosF * oy;

            double swX = ( fpX + rotOx) / 1000.0;
            double swY = (-fpY + rotOy) / 1000.0;
            double swZ = isFront
                ? ( thicknessMm + oz) / 1000.0
                : (-oz) / 1000.0;

            SwAddin.Log($"  xform '{comp["ref"]}': fp=({fpX:F2},{fpY:F2}) rot={fpRot:F1} " +
                        $"mx={mx:F1} my={my:F1} mz={mz:F1} " +
                        $"sw=({swX*1000:F2},{swY*1000:F2},{swZ*1000:F2})mm");

            // Row-major packing (SW IMathTransform ArrayData convention)
            double[] data = new double[16];
            data[0]=R[0]; data[1]=R[1]; data[2]=R[2];
            data[3]=R[3]; data[4]=R[4]; data[5]=R[5];
            data[6]=R[6]; data[7]=R[7]; data[8]=R[8];
            data[9]=swX; data[10]=swY; data[11]=swZ;
            data[12]=1;
            return data;
        }

        private static double[] _RotX3x3(double a)
        {
            double c = Math.Cos(a), s = Math.Sin(a);
            return new double[] { 1,0,0, 0,c,-s, 0,s,c };
        }
        private static double[] _RotY3x3(double a)
        {
            double c = Math.Cos(a), s = Math.Sin(a);
            return new double[] { c,0,s, 0,1,0, -s,0,c };
        }
        private static double[] _RotZ3x3(double a)
        {
            double c = Math.Cos(a), s = Math.Sin(a);
            return new double[] { c,-s,0, s,c,0, 0,0,1 };
        }
        private static double[] _Mul3x3(double[] A, double[] B)
        {
            var C = new double[9];
            for (int r = 0; r < 3; r++)
                for (int c = 0; c < 3; c++)
                    for (int k = 0; k < 3; k++)
                        C[r*3+c] += A[r*3+k] * B[k*3+c];
            return C;
        }

        /// <summary>
        /// Build a map from KiCad ref → (modelBaseName, sw transform) for every
        /// component in layout.json that has a 3D model.
        /// </summary>
        private Dictionary<string, (string baseName, double[]? xform)>
            _BuildLayoutTransformMap(JObject layout, double thicknessMm)
        {
            var result  = new Dictionary<string, (string, double[]?)>(StringComparer.OrdinalIgnoreCase);
            var comps   = layout["components"] as JArray;
            if (comps == null) return result;

            foreach (var comp in comps)
            {
                var kicadRef = comp["ref"]?.ToString();
                if (string.IsNullOrEmpty(kicadRef)) continue;
                if (!(bool)(comp["has_3d_model"] ?? false)) continue;

                var baseName = comp["model_name"]?.ToString() ?? "";
                if (string.IsNullOrEmpty(baseName)) continue;

                double[]? xform = null;
                try   { xform = _ComputeSwTransform(comp, thicknessMm); }
                catch (Exception ex) { SwAddin.Log($"  WARN: transform failed for '{kicadRef}': {ex.Message}"); }

                result[kicadRef!] = (_SanitizeName(baseName), xform);
            }
            return result;
        }

        private Dictionary<string, string> _LoadRefMap(string saveDir)
        {
            var path = Path.Combine(saveDir, "comp_ref_map.json");
            if (File.Exists(path))
                try
                {
                    return JsonConvert.DeserializeObject<Dictionary<string, string>>(
                               File.ReadAllText(path))
                           ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                }
                catch { }
            return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        }

        private void _SaveRefMap(string saveDir, Dictionary<string, string> map)
        {
            File.WriteAllText(
                Path.Combine(saveDir, "comp_ref_map.json"),
                JsonConvert.SerializeObject(map, Newtonsoft.Json.Formatting.Indented));
        }

        private static string _StripInstanceSuffix(string name)
        {
            var idx = name.LastIndexOf('-');
            return (idx > 0 && name.Substring(idx + 1).All(char.IsDigit))
                ? name.Substring(0, idx) : name;
        }

        private static string _SanitizeName(string name)
        {
            var invalid = Path.GetInvalidFileNameChars();
            return string.Concat(name.Select(c => invalid.Contains(c) ? '_' : c));
        }

        // ── Registry helpers ────────────────────────────────────────────────────

        private JObject _LoadRegistry(string saveDir)
        {
            var path = Path.Combine(saveDir, "board_registry.json");
            if (File.Exists(path))
                try { return JObject.Parse(File.ReadAllText(path)); }
                catch { }
            return new JObject { ["boards"] = new JArray() };
        }

        private void _SaveRegistry(string saveDir, JObject registry)
        {
            File.WriteAllText(
                Path.Combine(saveDir, "board_registry.json"),
                registry.ToString(Newtonsoft.Json.Formatting.Indented));
        }

        private JToken? _FindRegistryEntry(JObject registry, string boardName)
        {
            var boards = registry["boards"] as JArray;
            return boards?.FirstOrDefault(b =>
                string.Equals(b["kicad_name"]?.ToString(), boardName, OIC));
        }

        private void _RegisterBoard(JObject registry, string boardName,
            string assemblyPath, string boardPartPath)
        {
            var boards = registry["boards"] as JArray ?? new JArray();
            for (int i = boards.Count - 1; i >= 0; i--)
                if (string.Equals(boards[i]["kicad_name"]?.ToString(), boardName, OIC))
                    boards.RemoveAt(i);
            boards.Add(new JObject
            {
                ["kicad_name"] = boardName,
                ["sldasm"]     = assemblyPath,
                ["sldprt"]     = boardPartPath,
            });
            registry["boards"] = boards;
        }

        // ── Document / component helpers ────────────────────────────────────────

        private IModelDoc2? _GetOpenDoc(string filePath)
        {
            var docs = _swApp.GetDocuments() as object[];
            if (docs == null) return null;
            foreach (IModelDoc2 d in docs)
                if (string.Equals(d.GetPathName(), filePath, OIC))
                    return d;
            return null;
        }

        private void _CloseDocIfOpen(string filePath)
        {
            var docs = _swApp.GetDocuments() as object[];
            if (docs == null) return;
            foreach (IModelDoc2 d in docs)
            {
                if (string.Equals(d.GetPathName(), filePath, OIC))
                {
                    SwAddin.Log($"Pull: closing '{Path.GetFileName(filePath)}'");
                    _swApp.CloseDoc(d.GetTitle());
                    return;
                }
            }
        }

        private IComponent2? _FindComponentByPath(IAssemblyDoc asm, string path)
        {
            var comps = asm.GetComponents(false) as object[];
            if (comps == null) return null;
            foreach (IComponent2 c in comps)
                if (string.Equals(c.GetPathName(), path, OIC))
                    return c;
            return null;
        }

        private IComponent2? _FindComponentByName(IAssemblyDoc asm, string name)
        {
            var comps = asm.GetComponents(false) as object[];
            if (comps == null) return null;
            foreach (IComponent2 c in comps)
                if (string.Equals(c.Name2, name, OIC))
                    return c;
            return null;
        }

        private IModelDoc2? _FindBoardPartDoc(IModelDoc2 doc)
        {
            if (doc.GetType() == (int)swDocumentTypes_e.swDocPART)
                return doc;

            if (doc is IAssemblyDoc assembly)
            {
                var layoutPath = Path.Combine(_syncDir, "ecad_to_mcad", "layout.json");
                if (File.Exists(layoutPath))
                {
                    var layout = JObject.Parse(File.ReadAllText(layoutPath));
                    var boardFileName = layout["board"]?["file_name"]?.ToString() ?? "";
                    var saveDir = Path.Combine(_syncDir, "sw_working");
                    var entry   = _FindRegistryEntry(_LoadRegistry(saveDir), boardFileName);
                    var boardPartPath = entry?["sldprt"]?.ToString() ?? "";

                    var components = assembly.GetComponents(true) as object[];
                    if (components != null)
                    {
                        foreach (IComponent2 comp in components)
                        {
                            if (!string.IsNullOrEmpty(boardPartPath) &&
                                string.Equals(comp.GetPathName(), boardPartPath, OIC))
                                return comp.GetModelDoc2() as IModelDoc2;
                            var name = comp.Name2 ?? "";
                            if (!string.IsNullOrEmpty(boardFileName) &&
                                name.IndexOf(boardFileName, OIC) >= 0)
                                return comp.GetModelDoc2() as IModelDoc2;
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
