# MyPower Presentation Generator

This project generates PowerPoint slides for MyPower process documentation.

## Slides produced by build_flowchart_smart.py

1) Cover slide
- Title (left), logo (top-right)
- Version table (Version No, Author, Date Approved, Date Distributed)
- Distribution List section with a table of names/roles/departments
- Footer with page info and document details

2) Amendments slide
- Header/title and logo (top-right)
- “Amendments Since Initial Distribution” heading
- Table with columns: Revision No, Date, Amendments
- Footer with page info and document details

3) Flowchart slide
- Title (left), logo (top-right)
- Optional Key box (top-right area, below the logo) when the schema has `ShowKey: True` and `key.txt` exists
- Auto-laid-out nodes (Start, Information, Action, Decision, End) with dynamic heights
- Orthogonal connectors with decision labels and lane reservation (reduced overlaps)
- Footer with page info and document details

4) Notes slide
- Title (left), logo (top-right)
- Center heading: “Further notes & guidance here”
- Large bordered rectangle area for additional notes/guidance
- Footer with page info and document details


Generators:
- Recommended: `build_flowchart_smart.py` (auto-detects newest `.txt`, supports auto logo and optional Key box)
- Legacy: `build_slide_flexible.py` (embedded `BULLET_TEXT` format)

## Quick start

```bash
cd "/Users/shantamtaneja/Documents/Daryon-MyPower/Process_Diagrams"

# (optional) create a virtualenv
python3 -m venv .venv
source .venv/bin/activate

# install dependencies
pip install -r requirements.txt

# Place one or more schema .txt files in this folder
# (the newest .txt will be used by default)

# Optional: my_power_logo.png in this folder for auto logo (top-right)
# Optional: key.txt in this folder + `ShowKey: True` in the schema to show the Key box

# generate slides (auto-picks the newest .txt)
python3 build_flowchart_smart.py

# output filename is derived from the schema name, e.g. bullets_inverter_order.pptx
```

CLI options (smart generator):

```bash
python3 build_flowchart_smart.py \
  --schema bullets_inverter_order.txt \
  --logo path/to/logo.png \
  --out output.pptx \
  --showkey   # schema must also contain `ShowKey: True`; schema controls visibility
```



## How the smart generator works (high level)

1) Read the newest `.txt` schema (or the file passed via `--schema`).
2) Parse explicit-ID nodes and edges; compute decision route preferences (`right` vs `down`).
3) Assign columns and order within columns; compute dynamic box heights based on content.
4) Route connectors orthogonally with edge-hug avoidance and lane reservation; label decision outputs.
5) Render the flow slide (title, optional Key, logo top-right, footer) and the notes slide.

Arrows are orthogonal with lane reservation to reduce intersections. Start/End are pill-shaped rounded rectangles. Colors and fonts come from palette constants in the script to match MyPower branding.

Key box: shown only if the schema contains `ShowKey: True` (case-insensitive; accepts `yes`/`1`) and `key.txt` exists next to the script/executable. The key box height auto-sizes to its wrapped contents.

Logo: auto-detects `my_power_logo.png` in the folder if `--logo` is not passed; placed top-right on all slides, with title width adjusted.

## Smart schema format (excerpt)

The smart generator expects explicit IDs and optional titles/details; decisions specify `Path "Label" -> [ID]` lines. Example excerpt:

- Top-level items are single-line steps with the form: `Type: Text`.
  - Valid `Type` values: `Start`, `Information`, `Action`, `Decision`, `End`.
- A `Decision` step is special and must be followed by one or more indented Paths:
  - Each Path line is `Path "Label":` (quotes optional, colon required)
  - Under each Path, add one or more indented step blocks (usually an `Action:`). For a multi-line `Action:` block, place its lines indented beneath it.
  - Paths are rendered to the left, right and bottom (in the order they appear: path1, path2, path3).

```text
ShowKey: False

Start: [S1] START
 Leads to: [A1]

Action: [A1]
Title: Action Title
Details: First detail line
Details: • Second detail with bullet
Details: Third detail line
Leads to: [D1]

Decision: [D1] Should we proceed?
Path "Yes" -> [A2]
Path "No" -> [E1]

Action: [A2]
Title: Continue Process
Details: What happens when yes
Leads to: [E1]

End: [E1] END
```

## Layout and styling

- Start lozenge: `#92D050`
- Information box: `#2FC9FF`
- Action box: `#9DC3E6`
- Decision (diamond): `#EAB0FA`
- End lozenge: `#FFC000`
- Arrows: `#4472C4`
- Document heading: `#005077`
- All other text: `#000000`

Shadows are disabled across shapes. Box borders use a dark teal outline by default. If you want borders in a different color, tweak `DARK_TEAL` or the `add_box_shape` line style in the script.

## Adding your logo

Place `my_power_logo.png` in this folder and it will be used automatically (top-right). Override with `--logo`.

## Packaging an executable for distribution

The executable can be packaged for non-technical users. Build instructions are platform-specific (no cross-compilation).

### Prerequisites

```bash
pip install pyinstaller
```

### Build Steps

**For macOS/Linux:**
```bash
pyinstaller --onefile --name ProcessFlowBuilder build_flowchart_smart.py
```

**For Windows:**
```bash
pyinstaller --onefile --name ProcessFlowBuilder build_flowchart_smart.py
```

Output will be in the `dist/` folder:
- macOS/Linux: `dist/ProcessFlowBuilder`
- Windows: `dist/ProcessFlowBuilder.exe`

### Distribution Package

Create a distribution folder with:

1. **ProcessFlowBuilder** (or `.exe` on Windows)
2. **USER_GUIDE.md** (copy this file for end users)
3. **QUICK_START_WINDOWS_BUILD.md** (for maintainers who need a Windows build)
4. **Optional:** `my_power_logo.png` (template logo)
5. **Optional:** `key.txt` (template for Key box)

### Distribution Instructions

Provide to end users:
1. Extract the distribution folder
2. Place their `.txt` schema files in the same folder as the executable
3. Optionally add `my_power_logo.png` and `key.txt`
4. Run the executable (double-click on Windows; command line on Mac/Linux)
5. The PowerPoint file will be generated in the same folder

See `USER_GUIDE.md` for detailed instructions for non-technical users.

### Building a Windows executable from macOS

PyInstaller cannot cross-compile. This repo includes a GitHub Actions workflow that builds a Windows executable for you:

- Workflow file: `.github/workflows/build-windows.yml`
- Triggers: manual (workflow_dispatch), tag pushes matching `v*`, and changes to `build_flowchart_smart.py` on `main`
- Artifact: `ProcessFlowBuilder-Windows` (contains `ProcessFlowBuilder.exe`)

## Troubleshooting

- If Python can’t find `python-pptx`, ensure the virtual environment is activated and run `pip install python-pptx lxml`.
- If fonts or colors look off, check the constants near the top of `build_slide_flexible.py`.
- If PowerPoint warns that the file is in use, close the PPTX before regenerating.