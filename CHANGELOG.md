# Changelog

All notable changes to this project are documented here.

## [Unreleased]

### Added

- **New `detail` module (20 tools)** for detail engineering (addresses #2):
  sketch editing (`catia_sketch_edit_open/close`) with drawing and **driving
  dimensional constraints** (`catia_sketch_constrain`: distance/length/radius/
  angle with values — angles in radians at COM level, driving mode set),
  constraint listing, user **parameters** (dimension/real/integer/boolean/
  string) and **formulas** (Knowledgeware `Relations.CreateFormula`), relation
  listing, multi-body management (`catia_body_create`, `catia_bodies_list`,
  `catia_set_inwork`), **boolean operations** (`catia_boolean`: add/subtract/
  intersect — on R62020 they take the body OBJECT and subtraction is
  `AddNewRemove`), **hole threading** configuration
  (`catia_hole_thread`: pitch/depth/side applied to an existing Hole, nominal
  `ThreadDiameter` left to CATIA's standard tables), feature rename/delete,
  face listing, circular-edge/cylindrical-face measurement
  (`catia_measure_circles`), material listing/application, and a
  `catia_run_catscript` escape hatch that executes an ad-hoc CATScript and
  reads back its output file.

### Fixed (CATIA V5-6R2020 / B30, verified against a live instance)

- **Part Design**: Pad/Pocket/Hole raised `E_FAIL` on freshly created documents
  because the in-work object is a *Geometrical Set* — the sketch's body is now
  forced in-work before feature creation.
- **Part Design**: Shaft/Groove drive the revolution axis from the sketch's
  `AbsoluteAxis.VDirection` (a drawn line axis raises `E_FAIL`; wrapping the
  sketch itself with `CreateReferenceFromObject` raises `DISP_E_BADTYPE`) and
  set `FirstAngle` through the parameter object (dynamic-dispatch PUTs fail on
  factory-returned objects).
- **Part Design**: Hole uses `AddNewHoleFromSketch(sketch, depth)` — plain
  `AddNewHole(sketch, depth)` fails with a type mismatch on R62020 — and sets
  Diameter/BottomType/ThreadingMode via the dispatch fallback. Threading mode
  is `0` (`catThreadedHoleThreading`), not 1 (smooth).
- **Sketcher**: constraint type codes corrected against the B30 type libraries
  (`CatConstraintType`: distance=1, length=5, radius=14, angle=6, on=2,
  tangency=4, parallelism=8, horizontality=10, perpendicularity=11,
  verticality=13, reference=0). Angle values are converted to radians and
  dimensional constraints are set driving (`Mode=0`).
- **Sketcher**: arcs use `Factory2D.CreateCircle` with radian angles
  (`CreateArc` does not resolve through IDispatch on R62020), with a polyline
  fallback; `OpenEdition` returning `None` is retried once.
- **Measurement**: `GetWorkbench("SPAWorkbench")` lives on the **Document**, not
  on the Application (Application only has `GetWorkbenchId`) — the previous
  calls raised AttributeError on every measurement tool.
- **Measurement**: SPA `Measurable` objects are re-dispatched before reading;
  through pywin32 dynamic dispatch Volume/Area/COG otherwise read 0.
- **Measurement**: the COM API returns SI units (m3, m2, m) — outputs are now
  converted to mm-based values; the bounding box is derived from vertex points
  (`Measurable` has no bounding-box member in the V5 automation API).

## [0.2.0] — 2026-08-20

### Added
- **Generative Shape Design (GSD) module** — 24 new wireframe/surface tools
  (geometrical sets, 3D points/lines/planes/splines/circles, multi-sections
  surface, sweep, extrude, revolve, fill, blend, offset, join, split, trim,
  symmetry, ThickSurface/CloseSurface), bringing the server to **78 tools**.
  Contributed by @gaoflow in #4, runtime-tested against a live CATIA V5
  instance (30/30 smoke tests).
- GSD smoke test script (`scripts/gsd_smoke_test.py`).
- CI workflow running the offline test suite on Windows (Python 3.10/3.12).

### Fixed
- **All 14 Part Design tools raised `AttributeError`**: `Body` has no
  `ShapeFactory` in the CATIA V5 automation model — now uses
  `part.ShapeFactory`. Reported in #1 by @wuqing8577-netizen, fixed via #4
  (also proposed in #2 by @amirhossein199741).
- `Application.ActiveEditor` does not exist in CATIA V5 — view refresh,
  screenshot, set-view and fit-all now go through `ActiveWindow.ActiveViewer`
  (#4).
- Screenshots: CATIA V5 cannot capture PNG; the format is now chosen from the
  file extension (JPEG/BMP/TIFF) with a JPEG fallback for `.png` paths (#4).
- `pip install -e .` failed: invalid `build-backend` replaced with
  `setuptools.build_meta` (#4, also proposed in #2 and #5).
- Log file is written next to the package instead of the process working
  directory (which is `System32` when launched by Claude Desktop) (#4).
- `catia_new_part` / `catia_new_product` no longer crash when CATIA rejects
  the requested name (read-only `Part.Name` on some versions); the actual
  outcome is reported honestly. Contributed by @ewhenbula-svg in #5.
- **mcp SDK pinned to `<2`**: mcp 2.0.0 removed the low-level
  `Server.list_tools`/`call_tool` decorator API, breaking the server at
  startup on fresh installs.

## [0.1.0] — 2026-02-25

Initial release: 54 tools covering documents, 2D sketching, Part Design,
assembly, measurement, export and view control over CATIA V5 COM automation.
