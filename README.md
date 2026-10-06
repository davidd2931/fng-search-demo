# BOM Search Explorer — Portfolio Demo

A desktop portfolio demo for searching a small engineering bill-of-materials dataset.
The included records, names, identifiers, drawings, and project links are fictional
and generated locally. The demo has no company network configuration and does not
connect to a remote service.

## What it demonstrates

- Search across workbook attributes and spreadsheet cell contents.
- See which field matched, then inspect the matching BOM's component list.
- Search parts and find the BOMs that use them.
- Explore revision impact through assemblies and export a checklist.
- Open the bundled synthetic workbook, drawing, and CAD examples.

Company-specific templates and release integrations are intentionally omitted.
This is a portfolio demonstration, not an operational engineering system.

## Run it

Requires Python 3.10 or later on Windows with Tk installed.

```bash
python -m pip install -r requirements.txt
python tools/build_demo_data.py
python src/app.py
```

On Windows, `run-demo.bat` starts the app. The data generator can be run again at
any time to restore the six synthetic BOMs and their sample files.

## Example searches

- Search `gearbox` in BOM mode to find an assembly by name.
- Search `Drive Gear` in Components mode, then use **Where else is it used?**
- Select a part and open **What else would need revising?** to see the sample
  revision-impact tree.

The search tables are small Feather indexes built by `tools/build_demo_data.py`.
The source workbooks and files are also included so file-opening actions have
local examples to open.

## Project layout

```text
src/app.py                 desktop interface
src/demo_index.py          search and revision-impact data layer
src/demo_config.py         local-only paths and demo settings
src/demo_ecn.py            generic revision-label helper
demo_data/index/           synthetic Feather search tables
demo_data/files/           synthetic workbook, PDF, and CAD examples
tools/build_demo_data.py   regenerates the synthetic dataset
```

The dataset is deliberately small so the app can be explored without access to
any external files or services.
