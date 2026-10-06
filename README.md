# BOM Search Explorer — Portfolio Demo

A small desktop demo for searching a bill-of-materials (BOM) dataset. It contains
fictional sample records and runs locally on Windows. It does not connect to a
company network or remote service.

## Try it on Windows

### 1. Install Python (one time)

Install Python 3.10 or later from
[python.org](https://www.python.org/downloads/windows/). During installation,
select **Add python.exe to PATH**.

### 2. Download this demo

On this GitHub page, select **Code → Download ZIP**, then extract the ZIP file.

### 3. Start the app

Open the extracted folder and double-click **`run-demo.bat`**. On its first run,
it creates a local Python environment and installs the packages it needs. That
step requires an internet connection. Later starts open the app directly.

The six sample BOMs are already included; you do not need to generate or import
any data. If Windows SmartScreen appears, review the app source and only continue
if you are comfortable running it.

## What it demonstrates

- Search across workbook attributes and spreadsheet cell contents.
- See which field matched, then inspect the matching BOM's component list.
- Search parts and find the BOMs that use them.
- Explore revision impact through assemblies and export a checklist.
- Open bundled synthetic workbook, drawing, and CAD examples.

Company-specific templates and release integrations are intentionally omitted.
This is a portfolio demonstration, not an operational engineering system.

## Run manually (optional)

If you prefer a terminal, open PowerShell in the extracted folder and run:

```powershell
py -3 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe src\app.py
```

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
