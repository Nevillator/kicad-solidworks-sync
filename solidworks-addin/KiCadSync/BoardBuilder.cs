using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json.Linq;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace KiCadSync
{
    /// <summary>
    /// Creates and updates a native parametric SolidWorks part from board outline data.
    /// The board is a real SW part with editable sketches — not an imported STEP body.
    /// </summary>
    public class BoardBuilder
    {
        private readonly ISldWorks _swApp;
        private readonly IMathUtility _mathUtil;

        // Sketch entity ID sidecar file name (stored alongside the SW part)
        private const string EntityMapFile = "_kicad_sketch_map.json";

        public BoardBuilder(ISldWorks swApp)
        {
            _swApp = swApp;
            _mathUtil = _swApp.IGetMathUtility();
        }

        /// <summary>
        /// Create a new native SolidWorks part from board outline JSON.
        /// Returns the path to the saved .sldprt file.
        /// </summary>
        public string CreateBoardPart(string outlineJsonPath, double thicknessMm, string saveDir)
        {
            var outline = JObject.Parse(File.ReadAllText(outlineJsonPath));
            var outerBoundary = outline["outer_boundary"] as JArray ?? new JArray();
            var holes = outline["holes"] as JArray ?? new JArray();
            var cutouts = outline["cutouts"] as JArray ?? new JArray();

            // Create new part
            var doc = _swApp.NewDocument(
                _swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart),
                0, 0, 0) as IModelDoc2;
            if (doc == null)
                throw new Exception("Failed to create new SolidWorks part.");

            var part = doc as IPartDoc;

            // Set units to mm
            var ext = doc.Extension;
            ext.SetUserPreferenceInteger(
                (int)swUserPreferenceIntegerValue_e.swUnitsLinear, 0,
                (int)swLengthUnit_e.swMM);

            // ── Step 1: Sketch the outer boundary on the Front Plane ────────
            // Front Plane = XY plane (normal = Z), matching KiCad's STEP export
            // where the board lies in XY and thickness goes along +Z.

            var frontPlane = (IFeature)doc.FeatureByName("Front Plane");
            doc.ClearSelection2(true);
            frontPlane.Select2(false, 0);
            doc.SketchManager.InsertSketch(true);

            var entityMap = new Dictionary<string, int>();
            _DrawSegments(doc.SketchManager, outerBoundary, entityMap, "outer");
            doc.SketchManager.InsertSketch(true); // close sketch

            // ── Step 2: Extrude the outline to board thickness ──────────────

            // Select the sketch we just created
            var lastSketch = _GetLastSketch(doc);
            if (lastSketch != null)
            {
                ((IFeature)lastSketch).Select2(false, 0);
                var thicknessMeters = thicknessMm / 1000.0;
                doc.FeatureManager.FeatureExtrusion2(
                    true,   // single direction
                    false,  // not reverse
                    false,  // not both directions
                    (int)swEndConditions_e.swEndCondBlind,
                    0,      // end condition 2 (unused)
                    thicknessMeters,
                    0,      // depth 2 (unused)
                    false, false, false, false,
                    0, 0,   // draft angles
                    false, false, false, false,
                    true,   // merge bodies
                    true, true,
                    0, 0,
                    false);
            }

            // ── Step 3: Add holes ───────────────────────────────────────────

            foreach (var hole in holes)
            {
                var holeType = hole["type"]?.ToString();
                if (holeType == "round")
                    _AddRoundHole(doc, hole, thicknessMm);
                else if (holeType == "slot")
                    _AddSlot(doc, hole, thicknessMm);
            }

            // ── Step 4: Add cutouts ─────────────────────────────────────────

            for (int i = 0; i < cutouts.Count; i++)
            {
                var cutoutSegs = cutouts[i] as JArray ?? new JArray();
                _AddCutout(doc, cutoutSegs, thicknessMm, entityMap, $"cutout_{i}");
            }

            // ── Save ────────────────────────────────────────────────────────

            var savePath = Path.Combine(saveDir, "Board.sldprt");
            int saveErr = 0, saveWarn = 0;
            doc.Extension.SaveAs3(savePath,
                (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                null, null, ref saveErr, ref saveWarn);

            // Save entity map sidecar
            var mapPath = Path.Combine(saveDir, EntityMapFile);
            File.WriteAllText(mapPath,
                Newtonsoft.Json.JsonConvert.SerializeObject(entityMap,
                    Newtonsoft.Json.Formatting.Indented));

            return savePath;
        }

        // ── Sketch drawing helpers ──────────────────────────────────────────

        private void _DrawSegments(ISketchManager skMgr, JArray segments,
            Dictionary<string, int> entityMap, string prefix)
        {
            for (int i = 0; i < segments.Count; i++)
            {
                var seg = segments[i] as JObject;
                if (seg == null) continue;

                var segType = seg["type"]?.ToString();
                ISketchSegment skSeg = null;

                if (segType == "line")
                {
                    var s = seg["start"];
                    var e = seg["end"];
                    double sx = (double)s["x_mm"] / 1000.0;
                    double sy = -(double)s["y_mm"] / 1000.0; // flip Y
                    double ex = (double)e["x_mm"] / 1000.0;
                    double ey = -(double)e["y_mm"] / 1000.0;

                    skSeg = skMgr.CreateLine(sx, sy, 0, ex, ey, 0) as ISketchSegment;
                }
                else if (segType == "arc")
                {
                    var s = seg["start"];
                    var m = seg["mid"];
                    var e = seg["end"];
                    double sx = (double)s["x_mm"] / 1000.0;
                    double sy = -(double)s["y_mm"] / 1000.0;
                    double mx = (double)m["x_mm"] / 1000.0;
                    double my = -(double)m["y_mm"] / 1000.0;
                    double ex = (double)e["x_mm"] / 1000.0;
                    double ey = -(double)e["y_mm"] / 1000.0;

                    // Create arc through three points
                    skSeg = skMgr.Create3PointArc(sx, sy, 0, ex, ey, 0, mx, my, 0)
                            as ISketchSegment;
                }

                if (skSeg != null)
                {
                    entityMap[$"{prefix}_{i}"] = skSeg.GetID();
                }
            }
        }

        private void _AddRoundHole(IModelDoc2 doc, JToken hole, double thicknessMm)
        {
            double cx = (double)hole["center"]["x_mm"] / 1000.0;
            double cy = -(double)hole["center"]["y_mm"] / 1000.0;
            double diameter = (double)hole["diameter_mm"] / 1000.0;
            double radius = diameter / 2.0;
            double thickness = thicknessMm / 1000.0;

            // Select the top face for the sketch
            _SelectTopFace(doc);
            doc.SketchManager.InsertSketch(true);

            doc.SketchManager.CreateCircle(cx, cy, 0, cx + radius, cy, 0);
            doc.SketchManager.InsertSketch(true);

            // Extruded cut through all
            var lastSketch = _GetLastSketch(doc);
            if (lastSketch != null)
            {
                ((IFeature)lastSketch).Select2(false, 0);
                doc.FeatureManager.FeatureCut4(
                    true,   // single direction
                    false,  // not reverse
                    false,  // not both directions
                    (int)swEndConditions_e.swEndCondThroughAll,
                    0,
                    thickness, thickness,
                    false, false, false, false,
                    0, 0,
                    false, false, false, false,
                    false, true, true,
                    true, true,
                    false, 0,
                    false, false);
            }
        }

        private void _AddSlot(IModelDoc2 doc, JToken slot, double thicknessMm)
        {
            double cx = (double)slot["center"]["x_mm"] / 1000.0;
            double cy = -(double)slot["center"]["y_mm"] / 1000.0;
            double width = (double)slot["width_mm"] / 1000.0;
            double length = (double)slot["length_mm"] / 1000.0;
            double angleDeg = (double)slot["angle_deg"];
            double angleRad = angleDeg * Math.PI / 180.0;
            double thickness = thicknessMm / 1000.0;

            double halfLen = (length - width) / 2.0;
            double r = width / 2.0;

            // Slot endpoints (centers of the end arcs)
            double dx = Math.Cos(angleRad);
            double dy = -Math.Sin(angleRad); // flip Y for SW
            double p1x = cx - halfLen * dx;
            double p1y = cy - halfLen * dy;
            double p2x = cx + halfLen * dx;
            double p2y = cy + halfLen * dy;

            _SelectTopFace(doc);
            doc.SketchManager.InsertSketch(true);

            // Use the sketch slot tool (straight slot)
            doc.SketchManager.CreateSketchSlot(
                (int)swSketchSlotCreationType_e.swSketchSlotCreationType_line,
                (int)swSketchSlotLengthType_e.swSketchSlotLengthType_CenterCenter,
                width,
                p1x, p1y, 0,
                p2x, p2y, 0,
                0, 0, 0,   // start direction (unused for line type)
                1, false);

            doc.SketchManager.InsertSketch(true);

            var lastSketch = _GetLastSketch(doc);
            if (lastSketch != null)
            {
                ((IFeature)lastSketch).Select2(false, 0);
                doc.FeatureManager.FeatureCut4(
                    true, false, false,
                    (int)swEndConditions_e.swEndCondThroughAll,
                    0, thickness, thickness,
                    false, false, false, false,
                    0, 0,
                    false, false, false, false,
                    false, true, true,
                    true, true,
                    false, 0,
                    false, false);
            }
        }

        private void _AddCutout(IModelDoc2 doc, JArray segments, double thicknessMm,
            Dictionary<string, int> entityMap, string prefix)
        {
            double thickness = thicknessMm / 1000.0;

            _SelectTopFace(doc);
            doc.SketchManager.InsertSketch(true);
            _DrawSegments(doc.SketchManager, segments, entityMap, prefix);
            doc.SketchManager.InsertSketch(true);

            var lastSketch = _GetLastSketch(doc);
            if (lastSketch != null)
            {
                ((IFeature)lastSketch).Select2(false, 0);
                doc.FeatureManager.FeatureCut4(
                    true, false, false,
                    (int)swEndConditions_e.swEndCondThroughAll,
                    0, thickness, thickness,
                    false, false, false, false,
                    0, 0,
                    false, false, false, false,
                    false, true, true,
                    true, true,
                    false, 0,
                    false, false);
            }
        }

        // ── Utilities ───────────────────────────────────────────────────────

        private void _SelectTopFace(IModelDoc2 doc)
        {
            // Select the largest planar face with normal pointing along +Z.
            // The board is sketched on the Front Plane (XY) and extruded along +Z,
            // so the top face normal = (0, 0, 1).
            var part = doc as IPartDoc;
            if (part == null) return;

            var bodies = part.GetBodies2((int)swBodyType_e.swSolidBody, false) as object[];
            if (bodies == null || bodies.Length == 0) return;

            IFace2 bestFace = null;
            double bestArea = 0;

            foreach (IBody2 body in bodies)
            {
                var faces = body.GetFaces() as object[];
                if (faces == null) continue;

                foreach (IFace2 face in faces)
                {
                    var surface = face.GetSurface() as ISurface;
                    if (surface == null || !surface.IsPlane()) continue;

                    var normal = face.Normal as double[];
                    if (normal == null || normal.Length < 3) continue;

                    if (normal[2] > 0.9)  // +Z = top face
                    {
                        double area = face.GetArea();
                        if (area > bestArea)
                        {
                            bestArea = area;
                            bestFace = face;
                        }
                    }
                }
            }

            if (bestFace != null)
            {
                doc.ClearSelection2(true);
                ((IEntity)bestFace).Select4(false, null);
            }
        }

        private ISketch _GetLastSketch(IModelDoc2 doc)
        {
            // Walk features in reverse to find the most recently added sketch
            IFeature feat = doc.FirstFeature();
            ISketch lastSketch = null;

            while (feat != null)
            {
                if (feat.GetTypeName2() == "ProfileFeature")
                    lastSketch = feat.GetSpecificFeature2() as ISketch;
                feat = feat.GetNextFeature() as IFeature;
            }

            return lastSketch;
        }
    }
}
