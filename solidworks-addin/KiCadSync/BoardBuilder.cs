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

        // Sketch entity ID sidecar file name (stored alongside the SW part)
        private const string EntityMapFile = "_kicad_sketch_map.json";

        public BoardBuilder(ISldWorks swApp)
        {
            _swApp = swApp;
        }

        private void _BeginDraw(IModelDoc2 doc)
        {
            doc.SketchManager.AddToDB = true;
            doc.SketchManager.AutoSolve = false;
            _swApp.SetUserPreferenceToggle(
                (int)swUserPreferenceToggle_e.swSketchAutomaticRelations, false);
        }

        private void _EndDraw(IModelDoc2 doc)
        {
            doc.SketchManager.AddToDB = false;
            doc.SketchManager.AutoSolve = true;
            _swApp.SetUserPreferenceToggle(
                (int)swUserPreferenceToggle_e.swSketchAutomaticRelations, true);
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
            SwAddin.Log($"BoardBuilder: outer={outerBoundary.Count} segs, holes={holes.Count}, cutouts={cutouts.Count}");

            // Create new part
            var template = _swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart);
            SwAddin.Log($"BoardBuilder: template='{template}'");
            var doc = _swApp.NewDocument(template, 0, 0, 0) as IModelDoc2;
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

            SwAddin.Log("BoardBuilder: selecting Front Plane...");
            doc.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
            doc.SketchManager.InsertSketch(true);
            SwAddin.Log("BoardBuilder: sketch opened, drawing segments...");
            _BeginDraw(doc);

            var entityMap = new Dictionary<string, int>();
            _DrawSegments(doc.SketchManager, outerBoundary, entityMap, "outer");

            var outlineSketch = doc.SketchManager.ActiveSketch; // capture BEFORE closing
            _EndDraw(doc);
            SwAddin.Log("BoardBuilder: closing sketch...");
            doc.SketchManager.InsertSketch(true); // close sketch

            // ── Step 2: Extrude the outline to board thickness ──────────────
            // Extrude in -Z direction so: top face (component side) = Z=0,
            // bottom face = Z=-thickness.  This matches KiCad's STEP export
            // convention where the component side is at Z=0.

            SwAddin.Log($"BoardBuilder: outlineSketch={outlineSketch != null}");
            if (outlineSketch != null)
            {
                ((IFeature)outlineSketch).Select2(false, 0);
                var thicknessMeters = thicknessMm / 1000.0;
                SwAddin.Log($"BoardBuilder: extruding {thicknessMeters}m in -Z...");
                doc.FeatureManager.FeatureExtrusion2(
                    true,   // single direction
                    true,   // REVERSE direction → extrude toward -Z
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

            // ── Step 5: Add pad drill holes ──────────────────────────────────

            var drills = outline["drills"] as JArray ?? new JArray();
            SwAddin.Log($"BoardBuilder: {drills.Count} drills to add");
            if (drills.Count > 0)
                _AddDrills(doc, drills, thicknessMm);

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
                SwAddin.Log($"  Drawing {prefix}_{i}: type={segType}");
                object? skSegObj = null;

                if (segType == "line")
                {
                    var s = seg["start"];
                    var e = seg["end"];
                    double sx = (double)s["x_mm"] / 1000.0;
                    double sy = -(double)s["y_mm"] / 1000.0; // flip Y
                    double ex = (double)e["x_mm"] / 1000.0;
                    double ey = -(double)e["y_mm"] / 1000.0;

                    SwAddin.Log($"    Line ({sx},{sy}) -> ({ex},{ey})");
                    skSegObj = skMgr.CreateLine(sx, sy, 0, ex, ey, 0);
                    SwAddin.Log($"    CreateLine returned: {skSegObj?.GetType().Name ?? "null"}");
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

                    skSegObj = skMgr.Create3PointArc(sx, sy, 0, ex, ey, 0, mx, my, 0);
                }

                if (skSegObj is ISketchSegment skSeg)
                {
                    try
                    {
                        var id = skSeg.GetID();
                        entityMap[$"{prefix}_{i}"] = Convert.ToInt32(id);
                    }
                    catch (Exception ex)
                    {
                        SwAddin.Log($"    GetID failed: {ex.Message}");
                    }
                }
            }
        }

        /// <summary>
        /// Draws a stadium (slot) profile into the currently open sketch and adds
        /// tangent/equal/parallel relations so it is fully constrained like a native slot.
        /// Returns true if all 4 entities were created successfully.
        /// Assumes sketch is already open on Front Plane.
        /// </summary>
        private bool _DrawStadium(IModelDoc2 doc, ISketchManager skMgr,
            double cx, double cy, double w, double h, double angleDeg)
        {
            double angleRad = angleDeg * Math.PI / 180.0;
            double axisX =  Math.Sin(angleRad);
            double axisY =  Math.Cos(angleRad);
            double perpX = -axisY;
            double perpY =  axisX;

            double halfLen   = (h - w) / 2.0;
            double arcRadius = w / 2.0;

            // Arc centers
            double ac1x = cx - halfLen * axisX,  ac1y = cy - halfLen * axisY;
            double ac2x = cx + halfLen * axisX,  ac2y = cy + halfLen * axisY;

            // Tangent points (corners of the bounding rectangle)
            double tlx = ac1x + arcRadius * perpX,  tly = ac1y + arcRadius * perpY;
            double trx = ac2x + arcRadius * perpX,  try_ = ac2y + arcRadius * perpY;
            double brx = ac2x - arcRadius * perpX,  bry = ac2y - arcRadius * perpY;
            double blx = ac1x - arcRadius * perpX,  bly = ac1y - arcRadius * perpY;

            // Arc midpoints (tips of each semicircle)
            double mid2x = ac2x + arcRadius * axisX,  mid2y = ac2y + arcRadius * axisY;
            double mid1x = ac1x - arcRadius * axisX,  mid1y = ac1y - arcRadius * axisY;

            SwAddin.Log($"      stadium: ac1=({ac1x*1000:F3},{ac1y*1000:F3}) ac2=({ac2x*1000:F3},{ac2y*1000:F3}) r={arcRadius*1000:F3}mm");

            var line1 = skMgr.CreateLine(tlx, tly, 0, trx, try_, 0) as ISketchSegment;
            var arc2  = skMgr.Create3PointArc(trx, try_, 0, brx, bry, 0, mid2x, mid2y, 0) as ISketchSegment;
            var line2 = skMgr.CreateLine(brx, bry, 0, blx, bly, 0) as ISketchSegment;
            var arc1  = skMgr.Create3PointArc(blx, bly, 0, tlx, tly, 0, mid1x, mid1y, 0) as ISketchSegment;

            SwAddin.Log($"      line1={line1!=null} arc2={arc2!=null} line2={line2!=null} arc1={arc1!=null}");

            if (line1 == null || arc2 == null || line2 == null || arc1 == null)
                return false;

            // Slot-like constraints: tangent at each junction, equal radii, parallel lines
            _SketchRelation(doc, line1, arc2,  "sgTANGENT");
            _SketchRelation(doc, arc2,  line2, "sgTANGENT");
            _SketchRelation(doc, line2, arc1,  "sgTANGENT");
            _SketchRelation(doc, arc1,  line1, "sgTANGENT");
            _SketchRelation(doc, arc1,  arc2,  "sgEQUAL");
            _SketchRelation(doc, line1, line2, "sgPARALLEL");

            return true;
        }

        private void _SketchRelation(IModelDoc2 doc, ISketchSegment a, ISketchSegment b, string relType)
        {
            try
            {
                a.Select4(false, null);
                b.Select4(true, null);
                doc.SketchAddConstraints(relType);
                doc.ClearSelection2(true);
            }
            catch (Exception ex)
            {
                SwAddin.Log($"      _SketchRelation({relType}) failed: {ex.Message}");
            }
        }

        private bool _DrawStadiumOrSlot(IModelDoc2 doc, ISketchManager skMgr,
            double cx, double cy, double w, double h, double angleDeg)
        {
            return _DrawStadium(doc, skMgr, cx, cy, w, h, angleDeg);
        }

        private void _AddRoundHole(IModelDoc2 doc, JToken hole, double thicknessMm)
        {
            double cx = (double)hole["center"]["x_mm"] / 1000.0;
            double cy = -(double)hole["center"]["y_mm"] / 1000.0;
            double diameter = (double)hole["diameter_mm"] / 1000.0;
            double radius = diameter / 2.0;

            doc.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
            doc.SketchManager.InsertSketch(true);
            _BeginDraw(doc);

            doc.SketchManager.CreateCircle(cx, cy, 0, cx + radius, cy, 0);
            var rhSketch = doc.SketchManager.ActiveSketch;
            _EndDraw(doc);
            doc.SketchManager.InsertSketch(true);

            if (rhSketch != null)
            {
                ((IFeature)rhSketch).Select2(false, 0);
                doc.FeatureManager.FeatureCut3(
                    false, false, true,
                    (int)swEndConditions_e.swEndCondThroughAll,
                    (int)swEndConditions_e.swEndCondThroughAll,
                    0, 0,
                    false, false, false, false,
                    0, 0,
                    false, false, false, false,
                    false, true, true,
                    true, true,
                    false, 0, 0.0, false);
            }
        }

        private void _AddSlot(IModelDoc2 doc, JToken slot, double thicknessMm)
        {
            double cx       = (double)slot["center"]["x_mm"] / 1000.0;
            double cy       = -(double)slot["center"]["y_mm"] / 1000.0;
            double width    = (double)slot["width_mm"]  / 1000.0;
            double length   = (double)slot["length_mm"] / 1000.0;
            double angleDeg = (double)slot["angle_deg"];

            SwAddin.Log($"  Slot ({cx*1000:F3},{cy*1000:F3}) {width*1000:F2}x{length*1000:F2}mm θ={angleDeg}°");

            if (length <= width)
            {
                SwAddin.Log($"  degenerate slot (length<=width), skipping");
                return;
            }

            doc.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
            doc.SketchManager.InsertSketch(true);
            _BeginDraw(doc);

            _DrawStadiumOrSlot(doc, doc.SketchManager, cx, cy, width, length, angleDeg);

            var slotSketch = doc.SketchManager.ActiveSketch;
            _EndDraw(doc);
            doc.SketchManager.InsertSketch(true);

            if (slotSketch == null) { SwAddin.Log("  WARN: slot sketch null"); return; }

            ((IFeature)slotSketch).Select2(false, 0);
            var cutFeat = (IFeature)doc.FeatureManager.FeatureCut3(
                false, false, true,
                (int)swEndConditions_e.swEndCondThroughAll,
                (int)swEndConditions_e.swEndCondThroughAll,
                0, 0,
                false, false, false, false,
                0, 0,
                false, false, false, false,
                false, true, true,
                true, true,
                false, 0, 0.0, false);
            SwAddin.Log($"  Slot cut {(cutFeat != null ? "OK" : "FAILED")}");
        }

        private void _AddCutout(IModelDoc2 doc, JArray segments, double thicknessMm,
            Dictionary<string, int> entityMap, string prefix)
        {
            doc.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
            doc.SketchManager.InsertSketch(true);
            _BeginDraw(doc);
            _DrawSegments(doc.SketchManager, segments, entityMap, prefix);
            var cutoutSketch = doc.SketchManager.ActiveSketch;
            _EndDraw(doc);
            doc.SketchManager.InsertSketch(true);

            if (cutoutSketch != null)
            {
                ((IFeature)cutoutSketch).Select2(false, 0);
                doc.FeatureManager.FeatureCut3(
                    false, false, true,
                    (int)swEndConditions_e.swEndCondThroughAll,
                    (int)swEndConditions_e.swEndCondThroughAll,
                    0, 0,
                    false, false, false, false,
                    0, 0,
                    false, false, false, false,
                    false, true, true,
                    true, true,
                    false, 0, 0.0, false);
            }
        }

        private void _AddDrills(IModelDoc2 doc, JArray drills, double thicknessMm)
        {
            // One sketch + one cut PER FOOTPRINT.
            // All pad shapes (round circles + oval stadiums) go into the same sketch,
            // then a single through-all-both FeatureCut3 cuts them all at once.
            //
            // Arc convention: dir=-1 (CW) sweeps OUTWARD on each end cap.
            // (CCW would sweep inward, self-intersecting the oval.)
            //
            // Angle convention: KiCad long axis at θ → SW (cos θ, -sin θ) after Y-flip.

            // One sketch+cut per footprint. All pad types go in the same sketch.
            // Formula confirmed: SW overall_length = 2*|P1→P2| + width,
            // so |P1→P2| = (overall_length - width) / 2 = halfLen.

            foreach (var fpDrills in drills)
            {
                var refDes = fpDrills["ref"]?.ToString() ?? "?";
                var pads = fpDrills["pads"] as JArray;
                if (pads == null || pads.Count == 0) continue;

                SwAddin.Log($"  Drilling {pads.Count} pads for {refDes}");

                doc.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
                doc.SketchManager.InsertSketch(true);
                _BeginDraw(doc);

                int drawn = 0;
                for (int pi = 0; pi < pads.Count; pi++)
                {
                    var pad = pads[pi];
                    var padType = pad["type"]?.ToString();
                    double cx = (double)(pad["center"]?["x_mm"] ?? 0) / 1000.0;
                    double cy = -(double)(pad["center"]?["y_mm"] ?? 0) / 1000.0;

                    if (padType == "round")
                    {
                        double r = (double)(pad["diameter_mm"] ?? 0) / 2.0 / 1000.0;
                        SwAddin.Log($"    [{pi}] round ({cx*1000:F3},{cy*1000:F3}) r={r*1000:F3}mm");
                        var circ = doc.SketchManager.CreateCircle(cx, cy, 0, cx + r, cy, 0);
                        if (circ != null) drawn++;
                        else SwAddin.Log($"    WARN: CreateCircle null");
                    }
                    else if (padType == "oval")
                    {
                        double w        = (double)(pad["width_mm"]  ?? 0) / 1000.0;
                        double h        = (double)(pad["height_mm"] ?? 0) / 1000.0;
                        double angleDeg = (double)(pad["angle_deg"] ?? 0);

                        SwAddin.Log($"    [{pi}] oval ({cx*1000:F3},{cy*1000:F3}) w={w*1000:F3} h={h*1000:F3} θ={angleDeg}°");

                        if (h <= w || w < 1e-6)
                        {
                            SwAddin.Log($"    degenerate oval (h<=w), falling back to circle r={w/2*1000:F3}mm");
                            var circ = doc.SketchManager.CreateCircle(cx, cy, 0, cx + w/2, cy, 0);
                            if (circ != null) drawn++;
                            continue;
                        }
                        if (_DrawStadiumOrSlot(doc, doc.SketchManager, cx, cy, w, h, angleDeg)) drawn++;
                    }
                }

                var sk = doc.SketchManager.ActiveSketch;
                _EndDraw(doc);
                doc.SketchManager.InsertSketch(true);

                SwAddin.Log($"    drawn={drawn}, sketch={sk != null}");
                if (sk == null || drawn == 0) continue;

                ((IFeature)sk).Select2(false, 0);
                var cutFeat = (IFeature)doc.FeatureManager.FeatureCut3(
                    false, false, true,
                    (int)swEndConditions_e.swEndCondThroughAll,
                    (int)swEndConditions_e.swEndCondThroughAll,
                    0, 0,
                    false, false, false, false,
                    0, 0,
                    false, false, false, false,
                    false, true, true,
                    true, true,
                    false, 0, 0.0, false);
                SwAddin.Log($"    cut: {(cutFeat != null ? "OK" : "FAILED")} for {refDes}");
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

    }
}
