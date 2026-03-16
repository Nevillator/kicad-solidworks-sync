using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace KiCadSync
{
    /// <summary>
    /// Reads the board outline back from the native SolidWorks part's sketch,
    /// converting to the sync JSON format for pushing to KiCad.
    /// </summary>
    public class BoardReader
    {
        /// <summary>
        /// Extract the board outline from the first sketch ("Boss-Extrude1" profile)
        /// in the board part, plus any cut features, and return the outline JSON structure.
        /// </summary>
        public BoardOutlineData ReadOutline(IModelDoc2 boardDoc)
        {
            var result = new BoardOutlineData();
            var part = boardDoc as IPartDoc;
            if (part == null)
                throw new InvalidOperationException("Board document is not a part.");

            // ── Read the outer boundary from the base extrude's sketch ──────

            IFeature feat = boardDoc.FirstFeature();
            ISketch outerSketch = null;
            var cutSketches = new List<ISketch>();

            while (feat != null)
            {
                var typeName = feat.GetTypeName2();

                // The base extrude's profile sketch contains the outer boundary
                if (typeName == "Extrusion" && outerSketch == null)
                {
                    // Get the sketch that defines this extrusion
                    var subFeat = feat.GetFirstSubFeature() as IFeature;
                    while (subFeat != null)
                    {
                        if (subFeat.GetTypeName2() == "ProfileFeature")
                        {
                            outerSketch = subFeat.GetSpecificFeature2() as ISketch;
                            break;
                        }
                        subFeat = subFeat.GetNextSubFeature() as IFeature;
                    }
                }
                // Cut features contain holes, slots, and cutouts
                else if (typeName == "Cut" || typeName == "ICE")  // ICE = Instant Cut Extrude
                {
                    var subFeat = feat.GetFirstSubFeature() as IFeature;
                    while (subFeat != null)
                    {
                        if (subFeat.GetTypeName2() == "ProfileFeature")
                        {
                            var sk = subFeat.GetSpecificFeature2() as ISketch;
                            if (sk != null) cutSketches.Add(sk);
                            break;
                        }
                        subFeat = subFeat.GetNextSubFeature() as IFeature;
                    }
                }

                feat = feat.GetNextFeature() as IFeature;
            }

            // ── Convert outer sketch to segments ────────────────────────────

            if (outerSketch != null)
            {
                result.OuterBoundary = _SketchToSegments(outerSketch);
            }

            // ── Classify cut sketches ───────────────────────────────────────

            foreach (var sk in cutSketches)
            {
                var segments = _SketchToSegments(sk);
                var classified = _ClassifyCutSketch(segments);

                if (classified.hole != null)
                    result.Holes.Add(classified.hole);
                else if (classified.cutoutSegments != null)
                    result.Cutouts.Add(classified.cutoutSegments);
            }

            return result;
        }

        // ── Sketch → segment conversion ─────────────────────────────────────

        private List<object> _SketchToSegments(ISketch sketch)
        {
            var segments = new List<object>();
            var sketchSegments = sketch.GetSketchSegments() as object[];
            if (sketchSegments == null) return segments;

            foreach (ISketchSegment skSeg in sketchSegments)
            {
                if (skSeg.ConstructionGeometry) continue;

                var segType = (swSketchSegments_e)skSeg.GetType();

                if (segType == swSketchSegments_e.swSketchLINE)
                {
                    var line = skSeg as ISketchLine;
                    if (line == null) continue;
                    var sp = line.GetStartPoint2() as ISketchPoint;
                    var ep = line.GetEndPoint2() as ISketchPoint;
                    if (sp == null || ep == null) continue;

                    segments.Add(new LineSegmentData
                    {
                        Start = _SwPointToKiCad(sp),
                        End = _SwPointToKiCad(ep),
                    });
                }
                else if (segType == swSketchSegments_e.swSketchARC)
                {
                    var arc = skSeg as ISketchArc;
                    if (arc == null) continue;
                    var sp = arc.GetStartPoint2() as ISketchPoint;
                    var ep = arc.GetEndPoint2() as ISketchPoint;
                    if (sp == null || ep == null) continue;

                    // Compute the arc midpoint from center, radius, start/end angles
                    var cp = arc.GetCenterPoint2() as ISketchPoint;
                    double radius = arc.GetRadius();
                    double startAngle = arc.GetStartAngle();
                    double endAngle = arc.GetEndAngle();
                    double midAngle = (startAngle + endAngle) / 2.0;

                    // If the arc wraps around (end < start), adjust
                    if (endAngle < startAngle)
                        midAngle = (startAngle + endAngle + 2 * Math.PI) / 2.0;

                    double mx = cp.X + radius * Math.Cos(midAngle);
                    double my = cp.Y + radius * Math.Sin(midAngle);

                    segments.Add(new ArcSegmentData
                    {
                        Start = _SwPointToKiCad(sp),
                        Mid = new PointData
                        {
                            X = mx * 1000.0,       // meters to mm
                            Y = -(my * 1000.0),     // flip Y
                        },
                        End = _SwPointToKiCad(ep),
                    });
                }
            }

            return segments;
        }

        private PointData _SwPointToKiCad(ISketchPoint pt)
        {
            return new PointData
            {
                X = pt.X * 1000.0,        // meters to mm
                Y = -(pt.Y * 1000.0),      // flip Y: SW Y-up → KiCad Y-down
            };
        }

        // ── Inner loop classification ───────────────────────────────────────

        private (HoleData? hole, List<object>? cutoutSegments) _ClassifyCutSketch(
            List<object> segments)
        {
            // Single circle → round hole
            // (A circle in SW sketch = one arc segment that is a full circle,
            //  or we detect it from the geometry)
            if (segments.Count == 1 && segments[0] is ArcSegmentData arc)
            {
                // Check if start ≈ end (full circle)
                if (_PointsClose(arc.Start, arc.End))
                {
                    // Compute center and diameter from start/mid/end
                    var center = _CircleCenter(arc);
                    var r = _Dist(center, arc.Start);
                    return (new HoleData
                    {
                        Type = "round",
                        Center = center,
                        DiameterMm = r * 2,
                    }, null);
                }
            }

            // TODO: detect slots (2 arcs + 2 lines with equal arc radii)
            // For now, everything else is a cutout
            return (null, segments);
        }

        private bool _PointsClose(PointData a, PointData b) =>
            Math.Abs(a.X - b.X) < 0.001 && Math.Abs(a.Y - b.Y) < 0.001;

        private PointData _CircleCenter(ArcSegmentData arc)
        {
            // Center from 3 points (start, mid, end)
            double ax = arc.Start.X, ay = arc.Start.Y;
            double bx = arc.Mid.X, by = arc.Mid.Y;
            double cx = arc.End.X, cy = arc.End.Y;

            double d = 2 * (ax * (by - cy) + bx * (cy - ay) + cx * (ay - by));
            if (Math.Abs(d) < 1e-12) return arc.Mid; // degenerate

            double ux = ((ax * ax + ay * ay) * (by - cy) +
                          (bx * bx + by * by) * (cy - ay) +
                          (cx * cx + cy * cy) * (ay - by)) / d;
            double uy = ((ax * ax + ay * ay) * (cx - bx) +
                          (bx * bx + by * by) * (ax - cx) +
                          (cx * cx + cy * cy) * (bx - ax)) / d;

            return new PointData { X = ux, Y = uy };
        }

        private double _Dist(PointData a, PointData b) =>
            Math.Sqrt((a.X - b.X) * (a.X - b.X) + (a.Y - b.Y) * (a.Y - b.Y));
    }

    // ── Typed segment data models for serialization ─────────────────────────

    public class LineSegmentData
    {
        [JsonProperty("type")] public string Type => "line";
        [JsonProperty("start")] public PointData Start { get; set; } = new();
        [JsonProperty("end")] public PointData End { get; set; } = new();
    }

    public class ArcSegmentData
    {
        [JsonProperty("type")] public string Type => "arc";
        [JsonProperty("start")] public PointData Start { get; set; } = new();
        [JsonProperty("mid")] public PointData Mid { get; set; } = new();
        [JsonProperty("end")] public PointData End { get; set; } = new();
    }

    /// <summary>
    /// Extended BoardOutlineData with typed outer_boundary and cutouts.
    /// Replaces the simpler version in SyncManager.cs.
    /// </summary>
    public class BoardOutlineData
    {
        [JsonProperty("schema_version")] public string SchemaVersion { get; set; } = "1.0";
        [JsonProperty("outer_boundary")] public List<object> OuterBoundary { get; set; } = new();
        [JsonProperty("holes")] public List<HoleData> Holes { get; set; } = new();
        [JsonProperty("cutouts")] public List<List<object>> Cutouts { get; set; } = new();
    }
}
