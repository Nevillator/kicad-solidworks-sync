using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Newtonsoft.Json.Linq;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace KiCadSync
{
    /// <summary>
    /// Maps SolidWorks assembly components to KiCad reference designators
    /// and extracts/applies component positions with coordinate conversion.
    /// </summary>
    public class ComponentMapper
    {
        /// <summary>
        /// Build a mapping from SW component names to KiCad refs.
        /// Uses the layout.json from the last KiCad push and the board filename
        /// to identify which assembly component is the board vs. electronic parts.
        /// </summary>
        public static ComponentMap BuildMap(IAssemblyDoc assembly, string layoutJsonPath)
        {
            var layout = JObject.Parse(File.ReadAllText(layoutJsonPath));
            var boardFileName = layout["board"]?["file_name"]?.ToString() ?? "";
            var components = layout["components"] as JArray ?? new JArray();

            // Build a set of known refs from KiCad
            var knownRefs = new HashSet<string>();
            var modelOffsets = new Dictionary<string, ModelOffset>();
            foreach (var comp in components)
            {
                var r = comp["ref"]?.ToString();
                if (r == null) continue;
                knownRefs.Add(r);

                var offset = comp["model_offset"];
                var rotation = comp["model_rotation"];
                if (offset != null)
                {
                    modelOffsets[r] = new ModelOffset
                    {
                        X = (double)(offset["x_mm"] ?? 0),
                        Y = (double)(offset["y_mm"] ?? 0),
                        Z = (double)(offset["z_mm"] ?? 0),
                        RotX = (double)(rotation?["x_deg"] ?? 0),
                        RotY = (double)(rotation?["y_deg"] ?? 0),
                        RotZ = (double)(rotation?["z_deg"] ?? 0),
                    };
                }
            }

            var map = new ComponentMap { BoardFileName = boardFileName };

            // Enumerate SW assembly components
            var swComponents = assembly.GetComponents(true) as object[];
            if (swComponents == null) return map;

            foreach (IComponent2 swComp in swComponents)
            {
                var name = swComp.Name2 ?? "";

                // Check if this is the board component
                // KiCad STEP export names the board after the PCB file
                if (!string.IsNullOrEmpty(boardFileName) &&
                    name.Contains(boardFileName, StringComparison.OrdinalIgnoreCase))
                {
                    map.BoardComponent = swComp;
                    continue;
                }

                // Try to match against known refs
                // SW component names from STEP import are typically "RefDes-1" or just "RefDes"
                var matchedRef = _MatchRef(name, knownRefs);
                if (matchedRef != null)
                {
                    map.ComponentRefs[matchedRef] = swComp;
                    if (modelOffsets.TryGetValue(matchedRef, out var mo))
                        map.ModelOffsets[matchedRef] = mo;
                }
                else
                {
                    map.UnmappedComponents.Add(swComp);
                }
            }

            return map;
        }

        /// <summary>
        /// Extract current positions of all mapped components, converting from
        /// SolidWorks coordinate space to KiCad coordinate space.
        /// </summary>
        public static ComponentMovesData ExtractPositions(ComponentMap map)
        {
            var moves = new ComponentMovesData();

            foreach (var kvp in map.ComponentRefs)
            {
                var refDes = kvp.Key;
                var swComp = kvp.Value;

                var transform = swComp.Transform2 as IMathTransform;
                if (transform == null) continue;

                var data = transform.ArrayData as double[];
                if (data == null || data.Length < 13) continue;

                // Translation: data[9]=X, data[10]=Y, data[11]=Z (in meters)
                // Rotation matrix: data[0..8] (row-major 3x3)
                double swX = data[9];
                double swY = data[10];

                // Convert SW meters → KiCad mm, flip Y
                double kicadX = swX * 1000.0;
                double kicadY = -swY * 1000.0;

                // Extract Z rotation from the 3x3 rotation matrix
                // R = | data[0] data[1] data[2] |
                //     | data[3] data[4] data[5] |
                //     | data[6] data[7] data[8] |
                // For rotation around Z: θ = atan2(R10, R00) = atan2(data[3], data[0])
                // Negate because Y is flipped
                double rotDeg = -Math.Atan2(data[3], data[0]) * 180.0 / Math.PI;

                // Subtract model offset to get footprint origin position
                if (map.ModelOffsets.TryGetValue(refDes, out var offset))
                {
                    // The model offset is in KiCad space (mm, Y-down),
                    // rotated by the component's rotation
                    double offRad = rotDeg * Math.PI / 180.0;
                    double cos = Math.Cos(offRad);
                    double sin = Math.Sin(offRad);
                    double rotatedOffX = offset.X * cos - offset.Y * sin;
                    double rotatedOffY = offset.X * sin + offset.Y * cos;
                    kicadX -= rotatedOffX;
                    kicadY -= rotatedOffY;
                }

                moves.Components.Add(new ComponentMove
                {
                    Ref = refDes,
                    Position = new ComponentPosition
                    {
                        X = Math.Round(kicadX, 4),
                        Y = Math.Round(kicadY, 4),
                        Rotation = Math.Round(rotDeg % 360.0, 4),
                    },
                });
            }

            return moves;
        }

        // ── Ref matching ────────────────────────────────────────────────────

        private static string? _MatchRef(string swName, HashSet<string> knownRefs)
        {
            // Try exact match first
            if (knownRefs.Contains(swName)) return swName;

            // SW often appends "-1", "-2" etc. to imported component names
            // Try stripping the suffix
            var dashIdx = swName.LastIndexOf('-');
            if (dashIdx > 0)
            {
                var prefix = swName.Substring(0, dashIdx);
                if (knownRefs.Contains(prefix)) return prefix;
            }

            // Try matching if the SW name starts with a known ref
            foreach (var r in knownRefs)
            {
                if (swName.StartsWith(r, StringComparison.OrdinalIgnoreCase) &&
                    (swName.Length == r.Length || !char.IsLetterOrDigit(swName[r.Length])))
                    return r;
            }

            return null;
        }
    }

    // ── Data types ──────────────────────────────────────────────────────────

    public class ComponentMap
    {
        public string BoardFileName { get; set; } = "";
        public IComponent2? BoardComponent { get; set; }
        public Dictionary<string, IComponent2> ComponentRefs { get; set; } = new();
        public Dictionary<string, ModelOffset> ModelOffsets { get; set; } = new();
        public List<IComponent2> UnmappedComponents { get; set; } = new();
    }

    public class ModelOffset
    {
        public double X { get; set; }
        public double Y { get; set; }
        public double Z { get; set; }
        public double RotX { get; set; }
        public double RotY { get; set; }
        public double RotZ { get; set; }
    }
}
