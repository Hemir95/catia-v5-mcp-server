"""Detail-design tools for CATIA V5 (community contribution).

Adds capabilities for detail engineering: sketch editing + driving dimensional
constraints, parameters & formulas (Knowledgeware), booleans / multi-body,
hole threading, feature rename/delete, face listing, circular measurement,
materials, and a CATScript escape hatch.

Follows the R62020 recipes validated against a live instance (sketch arcs via
CreateCircle with RADIAN angles, features materialize via part.UpdateObject /
part.Update, booleans take the body OBJECT, hole threading applied after
materialization, SPA measurables re-dispatched before reading).
"""

from __future__ import annotations

import json
import math
import os
import tempfile
from typing import Any

from catia_mcp.connection import CATIAConnection

try:
    import win32com.client

    _HAS_WIN32 = True
except ImportError:
    win32com = None
    _HAS_WIN32 = False


def _rebind(measurable: Any) -> Any:
    """Re-dispatch a SPA Measurable.

    Through pywin32 dynamic dispatch the measurable returned by
    ``GetMeasurable`` is inert (Volume/Area/COG read 0); re-dispatching binds
    the type information and yields real values.
    """
    if _HAS_WIN32:
        return win32com.client.Dispatch(measurable)
    return measurable


# ── Constants from the B30 type libraries (validated via tlb dump) ─────────
# CatConstraintType (MecModTypeLib) — REAL B30 values (dump 16-oct-2026)
CST_TYPE: dict[str, int] = {
    "reference": 0,
    "distance": 1,
    "on": 2,
    "concentricity": 3,
    "tangency": 4,
    "length": 5,
    "angle": 6,
    "planar_angle": 7,
    "parallelism": 8,
    "axis_parallelism": 9,
    "horizontality": 10,
    "perpendicularity": 11,
    "axis_perpendicularity": 12,
    "verticality": 13,
    "radius": 14,
    "symmetry": 15,
    "midpoint": 16,
    "equidistance": 17,
    "major_radius": 18,
    "minor_radius": 19,
    "surf_contact": 20,
    "lin_contact": 21,
    "ponc_contact": 22,
}
# Constraints that take ONE element (mono) vs TWO (bi) vs THREE (tri)
CST_ARITY = {
    "horizontality": 1, "verticality": 1, "radius": 1, "length": 1, "on": 1, "reference": 1,
    "distance": 2, "angle": 2, "concentricity": 2, "tangency": 2, "parallelism": 2,
    "perpendicularity": 2, "symmetry": 2, "midpoint": 2,
    "equidistance": 3,
}
DIMENSIONAL = {"distance", "length", "radius", "angle"}

        # R62020 threading constants
HOLE_THREADING_NORMAL = 0  # catThreadedHoleThreading
HOLE_THREADING_SMOOTH = 1  # catSmoothHoleThreading
# CatThreadSide (PartTypeLib): right=0, left=1
HOLE_THREAD_RIGHT = 0
HOLE_THREAD_LEFT = 1

# catCstMode (MecModTypeLib): driving=0, driven=1
CST_MODE_DRIVING = 0


class DetailTools:
    """Detail-design tools: constraints, parameters, booleans, threads, inspection."""

    def __init__(self, connection: CATIAConnection) -> None:
        self.conn = connection
        self._active_sketch: Any = None
        self._active_factory: Any = None
        self._last_geoms: list[Any] = []

    # ─────────────────────────────────────────────────────────── tool defs ──
    def get_tool_definitions(self) -> list[dict[str, Any]]:
        return [
            {
                "name": "catia_sketch_edit_open",
                "description": (
                    "Open an EXISTING sketch for edition (adds/modifies 2D geometry). "
                    "Use before catia_sketch_draw when the sketch already exists. "
                    "Without sketch_name, opens the last sketch of the active body."
                ),
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "sketch_name": {"type": "string", "description": "Name of the sketch to edit"},
                    },
                },
            },
            {
                "name": "catia_sketch_edit_close",
                "description": (
                    "Close the sketch edition currently open and update the part. "
                    "Call after drawing/constraining."
                ),
                "inputSchema": {"type": "object", "properties": {}},
            },
            {
                "name": "catia_sketch_draw",
                "description": (
                    "Draw 2D geometry into the ACTIVE sketch edition. "
                    "shape: 'point' (x,y), 'line' (x1,y1,x2,y2), 'circle' (cx,cy,r), "
                    "'arc' (cx,cy,r,start_deg,end_deg), 'rectangle_centered' (cx,cy,w,h). "
                    "Tracks created elements for subsequent constraints."
                ),
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "shape": {
                            "type": "string",
                            "enum": ["point", "line", "circle", "arc", "rectangle_centered"],
                        },
                        "x": {"type": "number"}, "y": {"type": "number"},
                        "x1": {"type": "number"}, "y1": {"type": "number"},
                        "x2": {"type": "number"}, "y2": {"type": "number"},
                        "cx": {"type": "number"}, "cy": {"type": "number"},
                        "r": {"type": "number"},
                        "start_deg": {"type": "number"}, "end_deg": {"type": "number"},
                        "w": {"type": "number"}, "h": {"type": "number"},
                    },
                    "required": ["shape"],
                },
            },
            {
                "name": "catia_sketch_constrain",
                "description": (
                    "Add a constraint to the ACTIVE sketch, by default on the last drawn elements. "
                    "type: distance/length (mm), radius (mm), angle (deg, between two lines), "
                    "horizontality, verticality, concentricity, tangency, parallelism, "
                    "perpendicularity, midpoint, symmetry. "
                    "Dimensional types also set the driving value. elem1/elem2 = indices into the "
                    "elements drawn this edition (1-based, 0 = axis/origin dependent element not tracked)."
                ),
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "type": {
                            "type": "string",
                            "enum": list(CST_ARITY.keys()),
                        },
                        "value": {"type": "number", "description": "Driving value (mm or deg for angle)"},
                        "elem1": {"type": "integer", "description": "1-based index of first element (default: last drawn)"},
                        "elem2": {"type": "integer", "description": "1-based index of second element"},
                        "elem3": {"type": "integer", "description": "1-based index of third element"},
                    },
                    "required": ["type"],
                },
            },
            {
                "name": "catia_sketch_list_constraints",
                "description": "List all constraints of a sketch (default: active/last sketch) with type, value and mode.",
                "inputSchema": {
                    "type": "object",
                    "properties": {"sketch_name": {"type": "string"}},
                },
            },
            {
                "name": "catia_parameter_create",
                "description": (
                    "Create a user parameter in the active part. kind: length/angle/mass/volume "
                    "(dimensional, value in mm/deg/kg/m3...), real, integer, boolean, string. "
                    "Parameters drive formulas and can constrain the design."
                ),
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "name": {"type": "string"},
                        "kind": {
                            "type": "string",
                            "enum": ["length", "angle", "mass", "volume", "real", "integer", "boolean", "string"],
                        },
                        "value": {"type": "number", "description": "Initial value (numeric kinds)"},
                        "text": {"type": "string", "description": "Value (string kind)"},
                    },
                    "required": ["name", "kind"],
                },
            },
            {
                "name": "catia_formula_create",
                "description": (
                    "Create a formula driving a parameter from an expression. "
                    "Example: param_name='Length', formula='2 * `PartBody\\Sketch.1\\Radius` + 5mm'. "
                    "Reference parameters inside the formula with backticks; units allowed."
                ),
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "name": {"type": "string", "description": "Formula name (optional)"},
                        "param_name": {"type": "string", "description": "Full name of the parameter to drive"},
                        "formula": {"type": "string", "description": "Expression body"},
                    },
                    "required": ["param_name", "formula"],
                },
            },
            {
                "name": "catia_relation_list",
                "description": "List all relations (formulas, design tables, rules) in the active part.",
                "inputSchema": {"type": "object", "properties": {}},
            },
            {
                "name": "catia_body_create",
                "description": "Create a new Body in the active part and make it the in-work object (for multi-body design and booleans).",
                "inputSchema": {
                    "type": "object",
                    "properties": {"name": {"type": "string", "description": "Name for the new body"}},
                },
            },
            {
                "name": "catia_bodies_list",
                "description": "List all bodies of the active part and which one is in-work.",
                "inputSchema": {"type": "object", "properties": {}},
            },
            {
                "name": "catia_set_inwork",
                "description": "Set the in-work object of the active part (new features will be added to that body).",
                "inputSchema": {
                    "type": "object",
                    "properties": {"body_name": {"type": "string"}},
                    "required": ["body_name"],
                },
            },
            {
                "name": "catia_boolean",
                "description": (
                    "Boolean operation of a body against the in-work/target body. "
                    "op: 'add' (union), 'subtract' (pocket-like), 'intersect'."
                ),
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "op": {"type": "string", "enum": ["add", "subtract", "intersect"]},
                        "tool_body": {"type": "string", "description": "Name of the body to fuse/cut/intersect with"},
                        "target_body": {
                            "type": "string",
                            "description": "Body receiving the boolean (default: PartBody/main body)",
                        },
                    },
                    "required": ["op", "tool_body"],
                },
            },
            {
                "name": "catia_hole_thread",
                "description": (
                    "Configure threading on an existing Hole feature: standard metric profile with "
                    "pitch, thread depth, major diameter and direction."
                ),
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "feature_name": {"type": "string", "description": "Name of the Hole feature"},
                        "pitch": {"type": "number", "description": "Thread pitch in mm (e.g., 1.0 for M6)"},
                        "depth": {"type": "number", "description": "Thread depth in mm"},
                        "diameter": {"type": "number", "description": "Major (nominal) diameter in mm (e.g., 6 for M6)"},
                        "direction": {"type": "string", "enum": ["right", "left"], "default": "right"},
                    },
                    "required": ["feature_name", "pitch", "depth", "diameter"],
                },
            },
            {
                "name": "catia_feature_rename",
                "description": "Rename a feature or sketch in the active part.",
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "feature_name": {"type": "string"},
                        "new_name": {"type": "string"},
                        "kind": {"type": "string", "enum": ["feature", "sketch"], "default": "feature"},
                    },
                    "required": ["feature_name", "new_name"],
                },
            },
            {
                "name": "catia_feature_delete",
                "description": "Delete a feature from the active body (via CATIA selection delete).",
                "inputSchema": {
                    "type": "object",
                    "properties": {"feature_name": {"type": "string"}},
                    "required": ["feature_name"],
                },
            },
            {
                "name": "catia_list_faces",
                "description": "List faces of a feature or of the whole body (default: last feature), for fillet/chamfer/shell targets.",
                "inputSchema": {
                    "type": "object",
                    "properties": {"feature_name": {"type": "string"}},
                },
            },
            {
                "name": "catia_measure_circles",
                "description": (
                    "List the CIRCULAR edges of a feature (or the whole solid) with radius, "
                    "circumference and center — the detail check for holes, fillets and bores. "
                    "Without feature_name, uses the last feature of the main body."
                ),
                "inputSchema": {
                    "type": "object",
                    "properties": {"feature_name": {"type": "string", "description": "Feature to inspect (e.g., 'Pad.1')"}},
                },
            },
            {
                "name": "catia_material_list",
                "description": "List materials available to the active part (from the material document/library).",
                "inputSchema": {"type": "object", "properties": {}},
            },
            {
                "name": "catia_material_apply",
                "description": "Apply a material (by name) to a body or the main body of the active part.",
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "material_name": {"type": "string"},
                        "body_name": {"type": "string", "description": "Target body (default: main body)"},
                    },
                    "required": ["material_name"],
                },
            },
            {
                "name": "catia_run_catscript",
                "description": (
                    "Escape hatch: execute a CATScript (VBScript dialect) macro against the live CATIA. "
                    "Use for APIs not yet exposed as tools. To return data, write it to "
                    "'%TEMP%\\catia_mcp_out.txt' from the script (CreateObject(\"Scripting.FileSystemObject\")); "
                    "A.R.I.A. reads it back automatically after execution."
                ),
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "script": {"type": "string", "description": "Full CATScript source (must define Sub CATMain())"},
                    },
                    "required": ["script"],
                },
            },
        ]

    # ───────────────────────────────────────────────────────────── execute ──
    def execute(self, tool_name: str, arguments: dict[str, Any]) -> str:
        match tool_name:
            case "catia_sketch_edit_open":
                return self._sketch_edit_open(arguments.get("sketch_name"))
            case "catia_sketch_edit_close":
                return self._sketch_edit_close()
            case "catia_sketch_draw":
                return self._sketch_draw(arguments)
            case "catia_sketch_constrain":
                return self._sketch_constrain(arguments)
            case "catia_sketch_list_constraints":
                return self._sketch_list_constraints(arguments.get("sketch_name"))
            case "catia_parameter_create":
                return self._parameter_create(arguments)
            case "catia_formula_create":
                return self._formula_create(arguments)
            case "catia_relation_list":
                return self._relation_list()
            case "catia_body_create":
                return self._body_create(arguments.get("name"))
            case "catia_bodies_list":
                return self._bodies_list()
            case "catia_set_inwork":
                return self._set_inwork(arguments["body_name"])
            case "catia_boolean":
                return self._boolean(arguments)
            case "catia_hole_thread":
                return self._hole_thread(arguments)
            case "catia_feature_rename":
                return self._feature_rename(arguments)
            case "catia_feature_delete":
                return self._feature_delete(arguments["feature_name"])
            case "catia_list_faces":
                return self._list_faces(arguments.get("feature_name"))
            case "catia_measure_circles":
                return self._measure_circles(arguments)
            case "catia_material_list":
                return self._material_list()
            case "catia_material_apply":
                return self._material_apply(arguments["material_name"], arguments.get("body_name"))
            case "catia_run_catscript":
                return self._run_catscript(arguments["script"])
            case _:
                raise ValueError(f"Unknown detail tool: {tool_name}")

    # ─────────────────────────────────────────────────────────── helpers ──
    @staticmethod
    def _set_prop(obj: Any, name: str, value: Any) -> bool:
        """Set a COM property with the R62020 fallback: dynamic dispatch often
        rejects direct property PUTs on method-returned objects; properties
        returning parameter objects are set through .Value (August recipe)."""
        try:
            setattr(obj, name, value)
            return True
        except Exception:
            try:
                getattr(obj, name).Value = value
                return True
            except Exception:
                return False

    def _get_sketch(self, sketch_name: str | None = None) -> Any:
        part = self.conn.get_active_part()
        # sketches may live in ANY body (multi-body parts): search all
        candidates = []
        try:
            candidates.append(part.MainBody.Sketches)
        except Exception:
            pass
        try:
            main_name = part.MainBody.Name
        except Exception:
            main_name = ""
        try:
            bodies = part.Bodies
            for i in range(1, bodies.Count + 1):
                try:
                    b = bodies.Item(i)
                    if b.Name != main_name:
                        candidates.append(b.Sketches)
                except Exception:
                    continue
        except Exception:
            pass

        if sketch_name:
            for sks in candidates:
                try:
                    return sks.Item(sketch_name)
                except Exception:
                    continue
            raise RuntimeError(f"Sketch '{sketch_name}' not found in any body.")
        for sks in candidates:
            try:
                if sks.Count > 0:
                    return sks.Item(sks.Count)
            except Exception:
                continue
        raise RuntimeError("No sketches found in any body.")

    def _find_shape(self, name: str) -> Any:
        """Find a feature by name across ALL bodies of the active part."""
        part = self.conn.get_active_part()
        scopes = []
        try:
            scopes.append(part.MainBody)
        except Exception:
            pass
        main_name = getattr(part.MainBody, "Name", "") if scopes else ""
        try:
            bodies = part.Bodies
            for i in range(1, bodies.Count + 1):
                try:
                    b = bodies.Item(i)
                    if b.Name != main_name:
                        scopes.append(b)
                except Exception:
                    continue
        except Exception:
            pass
        for b in scopes:
            try:
                return b.Shapes.Item(name)
            except Exception:
                continue
        raise RuntimeError(f"Feature '{name}' not found in any body.")

    def _sketch_edit_open(self, sketch_name: str | None) -> str:
        self.conn.ensure_connected()
        sketch = self._get_sketch(sketch_name)
        try:
            factory = sketch.OpenEdition()
        except Exception as e:
            raise RuntimeError(f"Could not open sketch edition: {e}") from e
        self._active_sketch = sketch
        self._active_factory = factory
        self._last_geoms = []
        try:
            name = sketch.Name
        except Exception:
            name = "?"
        return f"Sketch '{name}' open for edition. Ready to draw/constrain."

    def _sketch_edit_close(self) -> str:
        if self._active_sketch is None:
            return "No sketch edition is open."
        try:
            self._active_sketch.CloseEdition()
        except Exception:
            pass
        try:
            self.conn.get_active_part().Update()
        except Exception:
            pass
        self.conn.refresh_display()
        n = len(self._last_geoms)
        self._active_sketch = None
        self._active_factory = None
        self._last_geoms = []
        return f"Sketch edition closed ({n} elements drawn). Part updated."

    def _track(self, obj: Any) -> None:
        if obj is not None:
            self._last_geoms.append(obj)

    def _sketch_draw(self, args: dict[str, Any]) -> str:
        self.conn.ensure_connected()
        if self._active_factory is None:
            raise RuntimeError("No sketch edition open. Call catia_sketch_edit_open (or catia_create_sketch) first.")
        f = self._active_factory
        shape = args["shape"]
        obj = None
        if shape == "point":
            obj = f.CreatePoint(float(args["x"]), float(args["y"]))
            self._track(obj)
            return f"Point at ({args['x']}, {args['y']}) created."
        if shape == "line":
            obj = f.CreateLine(float(args["x1"]), float(args["y1"]), float(args["x2"]), float(args["y2"]))
            self._track(obj)
            return f"Line ({args['x1']},{args['y1']}) -> ({args['x2']},{args['y2']}) created."
        if shape == "circle":
            obj = f.CreateClosedCircle(float(args["cx"]), float(args["cy"]), float(args["r"]))
            self._track(obj)
            return f"Circle c=({args['cx']},{args['cy']}) r={args['r']} created."
        if shape == "arc":
            # R62020 recipe (validated Aug-2026): the ARC form is CreateCircle
            # with RADIAN angles (CreateArc does not resolve via IDispatch here)
            start_rad = math.radians(float(args.get("start_deg", 0.0)))
            end_rad = math.radians(float(args.get("end_deg", 360.0)))
            try:
                obj = f.CreateCircle(float(args["cx"]), float(args["cy"]), float(args["r"]), start_rad, end_rad)
            except Exception:
                steps = max(8, int(abs(args.get("end_deg", 360.0) - args.get("start_deg", 0.0)) / 5))
                pts = []
                for i in range(steps + 1):
                    ang = start_rad + (end_rad - start_rad) * i / steps
                    pts.append((float(args["cx"]) + float(args["r"]) * math.cos(ang),
                                float(args["cy"]) + float(args["r"]) * math.sin(ang)))
                for i in range(steps):
                    f.CreateLine(pts[i][0], pts[i][1], pts[i + 1][0], pts[i + 1][1])
                obj = None
            self._track(obj)
            return f"Arc c=({args['cx']},{args['cy']}) r={args['r']} {args.get('start_deg', 0)}-{args.get('end_deg', 360)} deg created."
        if shape == "rectangle_centered":
            obj = f.CreateCenteredRectangle(
                float(args["cx"]), float(args["cy"]), float(args["w"]), float(args["h"])
            )
            self._track(obj)
            return f"Centered rectangle c=({args['cx']},{args['cy']}) {args['w']}x{args['h']} created."
        raise ValueError(f"Unknown shape '{shape}'")

    def _ref_for_constraint(self, elem: Any) -> Any:
        """Get a Reference for a 2D element; try the most compatible routes."""
        # Route 1: pass the object itself (works on R62020, same style as shaft)
        return elem

    def _sketch_constrain(self, args: dict[str, Any]) -> str:
        self.conn.ensure_connected()
        if self._active_sketch is None:
            raise RuntimeError("No sketch edition open. Call catia_sketch_edit_open first.")
        ctype = args["type"].lower()
        if ctype not in CST_TYPE:
            raise ValueError(f"Unknown constraint type '{ctype}'")
        arity = CST_ARITY[ctype]
        cat_type = CST_TYPE[ctype]

        def pick(key: str) -> Any:
            idx = args.get(key)
            if idx is None:
                if not self._last_geoms:
                    raise RuntimeError("No elements drawn this edition; draw geometry first or pass indices.")
                return self._last_geoms[-1]
            if idx < 1 or idx > len(self._last_geoms):
                raise RuntimeError(f"{key}={idx} out of range (1..{len(self._last_geoms)})")
            return self._last_geoms[idx - 1]

        constraints = self._active_sketch.Constraints
        cst = None
        if arity == 1:
            cst = constraints.AddMonoEltCst(cat_type, self._ref_for_constraint(pick("elem1")))
        elif arity == 2:
            cst = constraints.AddBiEltCst(
                cat_type, self._ref_for_constraint(pick("elem1")), self._ref_for_constraint(pick("elem2"))
            )
        else:
            cst = constraints.AddTriEltCst(
                cat_type,
                self._ref_for_constraint(pick("elem1")),
                self._ref_for_constraint(pick("elem2")),
                self._ref_for_constraint(pick("elem3")),
            )

        value_txt = ""
        if ctype in DIMENSIONAL and args.get("value") is not None:
            v = float(args["value"])
            try:
                cst.Mode = CST_MODE_DRIVING
            except Exception:
                pass
            # the constraint's driving dimension is a parameter: set through
            # Dimension (Constraint has no writable Value property in B30)
            val = math.radians(v) if ctype == "angle" else v
            try:
                cst.Dimension.Value = val
                value_txt = f", driving value = {v}{' deg' if ctype == 'angle' else ' mm'}"
            except Exception as e:
                value_txt = f" (WARNING: value not set: {e})"

        try:
            self._active_sketch.CloseEdition()
            reopen_needed = True
        except Exception:
            reopen_needed = False
        if reopen_needed:
            self._active_factory = self._active_sketch.OpenEdition()

        name = "?"
        try:
            name = cst.Name
        except Exception:
            pass
        return f"Constraint '{name}' ({ctype}) created{value_txt}."

    def _sketch_list_constraints(self, sketch_name: str | None) -> str:
        self.conn.ensure_connected()
        sketch = self._active_sketch if (sketch_name is None and self._active_sketch is not None) else self._get_sketch(sketch_name)
        cons = sketch.Constraints
        inv = {v: k for k, v in CST_TYPE.items()}
        out = []
        for i in range(1, cons.Count + 1):
            c = cons.Item(i)
            info: dict[str, Any] = {"index": i}
            try:
                info["name"] = c.Name
            except Exception:
                pass
            try:
                t = c.Type
                info["type"] = inv.get(t, str(t))
            except Exception:
                pass
            try:
                info["value"] = round(float(c.Value), 6)
            except Exception:
                pass
            try:
                info["mode"] = "driving" if c.Mode == CST_MODE_DRIVING else "driven"
            except Exception:
                pass
            out.append(info)
        if not out:
            return "No constraints in this sketch."
        return json.dumps(out, indent=2)

    def _parameter_create(self, args: dict[str, Any]) -> str:
        self.conn.ensure_connected()
        part = self.conn.get_active_part()
        params = part.Parameters
        name = args["name"]
        kind = args["kind"].lower()
        if kind in ("length", "angle", "mass", "volume"):
            magnitude = {"length": "LENGTH", "angle": "ANGLE", "mass": "MASS", "volume": "VOLUME"}[kind]
            value = float(args.get("value", 0.0))
            param = params.CreateDimension(name, magnitude, value)
        elif kind == "real":
            param = params.CreateReal(name, float(args.get("value", 0.0)))
        elif kind == "integer":
            param = params.CreateInteger(name, int(args.get("value", 0)))
        elif kind == "boolean":
            param = params.CreateBoolean(name, bool(args.get("value", False)))
        elif kind == "string":
            param = params.CreateString(name, str(args.get("text", "")))
        else:
            raise ValueError(f"Unknown parameter kind '{kind}'")
        try:
            part.Update()
        except Exception:
            pass
        return f"Parameter '{name}' ({kind}) created."

    def _formula_create(self, args: dict[str, Any]) -> str:
        self.conn.ensure_connected()
        part = self.conn.get_active_part()
        params = part.Parameters
        relations = part.Relations
        param = params.Item(args["param_name"])
        name = args.get("name", "")
        comment = ""
        try:
            rel = relations.CreateFormula(name, comment, param, args["formula"])
        except Exception as e:
            raise RuntimeError(f"CreateFormula failed: {e}") from e
        try:
            part.Update()
        except Exception:
            pass
        rname = "?"
        try:
            rname = rel.Name
        except Exception:
            pass
        return f"Formula '{rname}' created, driving '{args['param_name']}'."

    def _relation_list(self) -> str:
        self.conn.ensure_connected()
        part = self.conn.get_active_part()
        rels = part.Relations
        out = []
        for i in range(1, rels.Count + 1):
            r = rels.Item(i)
            info: dict[str, Any] = {"index": i}
            try:
                info["name"] = r.Name
            except Exception:
                pass
            try:
                info["type"] = r.Type if hasattr(r, "Type") else "?"
            except Exception:
                pass
            out.append(info)
        if not out:
            return "No relations in this part."
        return json.dumps(out, indent=2)

    def _find_body(self, body_name: str | None = None) -> Any:
        part = self.conn.get_active_part()
        if body_name:
            return part.Bodies.Item(body_name)
        return part.MainBody

    def _body_create(self, name: str | None) -> str:
        self.conn.ensure_connected()
        part = self.conn.get_active_part()
        body = part.Bodies.Add()
        if name:
            try:
                body.Name = name
            except Exception:
                pass
        try:
            part.InWorkObject = body
        except Exception:
            pass
        try:
            bname = body.Name
        except Exception:
            bname = "?"
        return f"Body '{bname}' created and set as in-work object."

    def _bodies_list(self) -> str:
        self.conn.ensure_connected()
        part = self.conn.get_active_part()
        inwork = "?"
        try:
            inwork = part.InWorkObject.Name
        except Exception:
            pass
        out = []
        bodies = part.Bodies
        for i in range(1, bodies.Count + 1):
            b = bodies.Item(i)
            try:
                out.append({"index": i, "name": b.Name})
            except Exception:
                out.append({"index": i, "name": f"body{i}"})
        return json.dumps({"bodies": out, "in_work": inwork}, indent=2)

    def _set_inwork(self, body_name: str) -> str:
        self.conn.ensure_connected()
        part = self.conn.get_active_part()
        body = part.Bodies.Item(body_name)
        part.InWorkObject = body
        return f"In-work object set to '{body_name}'."

    def _boolean(self, args: dict[str, Any]) -> str:
        self.conn.ensure_connected()
        part = self.conn.get_active_part()
        sf = part.ShapeFactory
        op = args["op"].lower()
        tool_body = part.Bodies.Item(args["tool_body"])
        if args.get("target_body"):
            part.InWorkObject = part.Bodies.Item(args["target_body"])
        ref = part.CreateReferenceFromObject(tool_body)
        made = None
        if op == "add":
            made = sf.AddNewAdd(tool_body)
        elif op == "subtract":
            # B30 has NO AddNewSubtract: Part Design subtraction = AddNewRemove
            made = sf.AddNewRemove(tool_body)
        elif op == "intersect":
            made = sf.AddNewIntersect(tool_body)
        else:
            raise ValueError(f"Unknown boolean op '{op}'")
        try:
            part.UpdateObject(made)
        except Exception:
            part.Update()
        self.conn.refresh_display()
        fname = "?"
        try:
            fname = made.Name
        except Exception:
            pass
        return f"Boolean '{op}' with '{args['tool_body']}' done. Feature: '{fname}'"

    def _hole_thread(self, args: dict[str, Any]) -> str:
        self.conn.ensure_connected()
        part = self.conn.get_active_part()
        body = part.MainBody
        hole = None
        try:
            hole = self._find_shape(args["feature_name"])
        except Exception as e:
            raise RuntimeError(f"Hole '{args['feature_name']}' not found: {e}") from e
        # Threading is configured DIRECTLY on the Hole object in B30
        # (no ThreadSetting child object): ThreadingMode 0 = threaded.
        # ThreadDiameter (minor/core) is left to CATIA's standard tables unless
        # explicitly requested — forcing it = nominal can break the rebuild.
        self._set_prop(hole, "ThreadingMode", HOLE_THREADING_NORMAL)
        results = []
        results.append(("ThreadPitch", self._set_prop(hole, "ThreadPitch", float(args["pitch"]))))
        results.append(("ThreadDepth", self._set_prop(hole, "ThreadDepth", float(args["depth"]))))
        if args.get("diameter") is not None:
            results.append(("ThreadDiameter", self._set_prop(hole, "ThreadDiameter", float(args["diameter"]))))
        failed = [n for n, r in results if not r]
        if failed:
            return f"Thread config FAILED on: {', '.join(failed)}"
        try:
            self._set_prop(hole, "ThreadSide", HOLE_THREAD_RIGHT if args.get("direction", "right") == "right" else HOLE_THREAD_LEFT)
        except Exception:
            pass
        try:
            part.Update()
        except Exception:
            # rebuild failed -> revert to smooth hole and report honestly
            try:
                self._set_prop(hole, "ThreadingMode", HOLE_THREADING_SMOOTH)
                part.Update()
            except Exception:
                pass
            return "Thread config rejected by CATIA rebuild (reverted to smooth hole). Check pitch/depth/standard."
        return f"Thread configured on '{args['feature_name']}': pitch {args['pitch']} mm, depth {args['depth']} mm."

    def _feature_rename(self, args: dict[str, Any]) -> str:
        self.conn.ensure_connected()
        part = self.conn.get_active_part()
        body = part.MainBody
        kind = args.get("kind", "feature")
        try:
            obj = self._get_sketch(args["feature_name"]) if kind == "sketch" else self._find_shape(args["feature_name"])
        except Exception as e:
            raise RuntimeError(f"'{args['feature_name']}' not found: {e}") from e
        obj.Name = args["new_name"]
        try:
            part.Update()
        except Exception:
            pass
        return f"'{args['feature_name']}' renamed to '{args['new_name']}'."

    def _feature_delete(self, feature_name: str) -> str:
        self.conn.ensure_connected()
        part = self.conn.get_active_part()
        try:
            obj = self._find_shape(feature_name)
        except Exception as e:
            raise RuntimeError(f"'{feature_name}' not found: {e}") from e
        sel = self.conn.hso
        sel.Clear()
        sel.Add(obj)
        sel.Delete()
        try:
            part.Update()
        except Exception:
            pass
        self.conn.refresh_display()
        return f"Feature '{feature_name}' deleted."

    def _list_faces(self, feature_name: str | None) -> str:
        self.conn.ensure_connected()
        part = self.conn.get_active_part()
        scope = self._find_shape(feature_name) if feature_name else self._get_last_shape(part.MainBody)
        sel = self.conn.hso
        sel.Clear()
        sel.Add(scope)
        sel.Search("Topology.Face,sel")
        faces = []
        for i in range(1, sel.Count + 1):
            try:
                faces.append({"index": i, "name": sel.Item(i).Value.Name})
            except Exception:
                faces.append({"index": i, "name": f"Face.{i}"})
        sel.Clear()
        if not faces:
            return "No faces found."
        return json.dumps(faces, indent=2)

    def _get_last_shape(self, body: Any) -> Any:
        shapes = body.Shapes
        if shapes.Count == 0:
            raise RuntimeError("No features in the body.")
        return shapes.Item(shapes.Count)

    def _measure_circles(self, args: dict[str, Any]) -> str:
        self.conn.ensure_connected()
        part = self.conn.get_active_part()
        # B30: GetWorkbench lives on the DOCUMENT, not on Application
        spa = self.conn.active_document.GetWorkbench("SPAWorkbench")
        feature_name = args.get("feature_name")
        scope = self._find_shape(feature_name) if feature_name else self._get_last_shape(part.MainBody)
        sel = self.conn.hso
        sel.Clear()
        sel.Add(scope)
        out = []
        # route 1: circular EDGES
        sel.Search("Topology.Edge,sel")
        for i in range(1, sel.Count + 1):
            try:
                obj = sel.Item(i).Value
                ref = part.CreateReferenceFromObject(obj)
                m = _rebind(spa.GetMeasurable(ref))
                info: dict[str, Any] = {"index": i, "kind": "edge"}
                try:
                    info["radius_mm"] = round(float(m.Radius), 6)
                except Exception:
                    continue  # not a circular edge -> skip
                try:
                    info["circumference_mm"] = round(float(m.Length), 6)
                except Exception:
                    pass
                try:
                    c = [0.0, 0.0, 0.0]
                    m.GetCenter(c)
                    info["center_mm"] = [round(c[0], 4), round(c[1], 4), round(c[2], 4)]
                except Exception:
                    pass
                out.append(info)
            except Exception:
                continue
        sel.Clear()
        # route 2: cylindrical FACES (measurable.Radius works reliably there)
        sel.Add(scope)
        try:
            sel.Search("Topology.Face,sel")
        except Exception:
            sel.Clear()
            return json.dumps(out, indent=2) if out else "No circular edges found."
        for i in range(1, sel.Count + 1):
            try:
                obj = sel.Item(i).Value
                ref = part.CreateReferenceFromObject(obj)
                m = _rebind(spa.GetMeasurable(ref))
                r = float(m.Radius)  # raises on planar faces
                out.append({"index": i, "kind": "face", "radius_mm": round(r, 6)})
            except Exception:
                continue
        sel.Clear()
        if not out:
            return "No circular edges or cylindrical faces found."
        return json.dumps(out, indent=2)

    def _material_list(self) -> str:
        self.conn.ensure_connected()
        doc = self.conn.active_document
        mats = doc.Materials
        out = []
        for i in range(1, mats.Count + 1):
            try:
                out.append(mats.Item(i).Name)
            except Exception:
                pass
        if not out:
            return "No materials accessible from this document (load a material library in CATIA first)."
        return json.dumps(out, indent=2)

    def _material_apply(self, material_name: str, body_name: str | None) -> str:
        self.conn.ensure_connected()
        doc = self.conn.active_document
        part = self.conn.get_active_part()
        mat = doc.Materials.Item(material_name)
        target = part.Bodies.Item(body_name) if body_name else part.MainBody
        ref = part.CreateReferenceFromObject(target)
        target.ApplyMaterial(ref, mat)
        try:
            part.Update()
        except Exception:
            pass
        return f"Material '{material_name}' applied to '{body_name or 'main body'}'."

    def _run_catscript(self, script: str) -> str:
        self.conn.ensure_connected()
        temp_dir = tempfile.gettempdir()
        base = "catia_mcp_macro"
        script_path = os.path.join(temp_dir, base + ".CATScript")
        out_path = os.path.join(temp_dir, "catia_mcp_out.txt")
        with open(script_path, "w", encoding="latin-1") as f:
            f.write(script)
        if os.path.exists(out_path):
            try:
                os.remove(out_path)
            except Exception:
                pass
        # B30 signature (from tlb dump: 5 params):
        # ExecuteScript(libraryPath, catScriptLibraryTypeDirectory=1,
        #               programName, functionName, params)
        self.conn.app.SystemService.ExecuteScript(temp_dir, 1, base + ".CATScript", "CATMain", [])
        result = "CATScript executed."
        if os.path.exists(out_path):
            try:
                with open(out_path, "r", encoding="latin-1") as f:
                    content = f.read()
                result += "\n--- script output (%TEMP%\\catia_mcp_out.txt) ---\n" + content
            except Exception as e:
                result += f"\n(output file present but unreadable: {e})"
        return result
