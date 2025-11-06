# Architecture Overview

This document explains the internal architecture of the presentation generators, with emphasis on the smart schema-based flow in `build_flowchart_smart.py`.

## High-level pipeline

Smart generator (`build_flowchart_smart.py`) pipeline:

1. Input: schema `.txt` (explicit node IDs, optional titles/details, decision paths)
2. Parse → `nodes` dict and `edges` list + `start_ids`
3. Plan decision routes (prefer `Yes` to the right; otherwise down)
4. Assign columns; order nodes within columns (topological preference)
5. Place nodes with dynamic height estimates (content-driven)
6. Route connectors orthogonally with lane reservation and edge-hug avoidance; label decision outputs near origin
7. Render flow slide (title, optional key, logo top-right, footer) + notes slide

These stages live in `build_flowchart_smart.py`. The legacy bullet-based stages live in `build_slide_flexible.py`.

## Key modules and responsibilities (smart)

Parsing (schema → graph):
- `parse_schema_details_only(text)` returns `nodes` (id → {kind, title, details_lines}), `edges` (u,v,label), and `start_ids`.

Decision routing preferences:
- `plan_decision_routes(nodes, edges)` decides whether a decision’s outgoing edges prefer `right` or `down` (favoring a "Yes" label to the right when present).

Layout:
- `assign_columns(...)` computes a column index for each node using BFS with decision preferences.
- `order_within_columns(...)` topologically orders nodes within each column.
- Heights: `estimate_action_height` uses PIL font metrics to compute content-driven heights.

Drawing:
- `add_node_shape(...)` draws node shapes and text with styling.
- `route_orthogonal_detour(...)` adds connectors, avoiding running along box edges and reserving lanes to prevent overlaps.
- Decision labels are placed near the arrow origin with offsets based on direction.
- `_add_standard_footer(...)` adds left page count and right document info.
- `_add_key_box(...)` renders a titled Key box in the top-right area; height auto-sizes to `key.txt` contents.

## Data structures (smart)

Example schema excerpt and derived structures:

```json
{
  "type": "Start", "text": ["START"],
  "children": [
    { "type": "Information", "text": ["Information"], "children": [
      { "type": "Action", "text": ["Action/s"], "children": [
        { "type": "Decision", "text": ["Decision?"],
          "paths": [
            { "label": "Decision path 1", "steps": [ {"type":"Action", "text":["Heading","Action/s"]} ] },
            { "label": "Decision path 2", "steps": [ {"type":"Action", "text":["Heading","• Bullet Actions"]} ] },
            { "label": "Decision path 3", "steps": [ {"type":"Action", "text":["Action/s"]} ] }
          ]
        }
      ]}
    ]}
  ]
}
```

Nodes/edges (simplified):

```json
{
  "type": "Start", "text": ["START"],
  "paths": { "next": { "type": "Information", ... } }
}
```

`nodes` entry (computed layout example):

```json
{ "id": "action_7", "type": "Action", "lines": ["Heading","Action/s"],
  "left": 2.5, "top": 3.8, "width": 2.0, "height": 0.9 }
```

`connectors` entry (with decision label):

```json
{ "from_id": "decision_5", "to_id": "action_7",
  "from_anchor": "left", "to_anchor": "right",
  "label": "Decision path 1" }
```

## Visual rules

- Start/End are pill-shaped rounded rectangles; Decision is a diamond; Action/Information are rounded rectangles.
- Colors and text styling are defined by constants at the top of `build_flowchart_smart.py` (`PALETTE`).
- All text is black by default except the document heading which uses `#005077`.
- Shadows are disabled globally.
- Connectors:
  - Straight for aligned endpoints.
  - Otherwise, two-segment orthogonal (vertical first, then horizontal with arrowhead).
  - Rounded line caps for cleaner edges; labels are centered on the main segment.

## Extensibility points

- Add new node types:
  - Extend `BOX_STYLES` with fill color and shape type (`rounded` or `diamond`).
  - Update `calculate_box_size` if the new shape needs different padding.

- Change layout spacing or positions:
  - Tweak the `cfg` margins/gaps and the dynamic height estimation in `render(...)`.

- Alternate connector routing:
  - `route_orthogonal_detour` tries L-shapes first and then gutter routes; tune PADs and gutter offsets.

- The notes slide:
  - Adjust sizes and text in `_add_notes_slide`.

## Error handling & validation

Parsing is permissive for the smart schema; unrecognized lines are ignored unless they match the schema regexes. Errors in routing or layout fall back to simpler paths.

## Dependencies

- Python 3.8+
- `python-pptx`
- `lxml` (python-pptx)
- `Pillow` (font metrics for dynamic sizing)

## Build & run

See `README.md` for instructions. The entry point is `build_flowchart_smart.py` which auto-detects the newest schema `.txt`, optional `my_power_logo.png`, and optional `key.txt` (when `ShowKey: True` in the schema), then saves `<schema-stem>.pptx`.


