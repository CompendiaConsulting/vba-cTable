# Installing cTable for Excel VBA

This guide explains how to add the cTable class library to an Excel VBA project. The library is implemented in VBA; it is not a COM add-in or an Excel worksheet function package.

> These instructions accompany the exported VBA modules in this repository. Import the files from `src/`, preserve their export attributes, and compile the VBA project after installation. Examples are illustrative unless explicitly identified as executed tests.

## 1. Prerequisites

- Desktop Microsoft Excel with VBA support and access to the Visual Basic Editor (VBE).
- A macro-enabled workbook (`.xlsm`) or other suitable VBA project.
- All of the library's exported source modules.
- **Microsoft Scripting Runtime**: the current cTable source contains early-bound `Scripting.Dictionary` declarations and explicitly requires this reference. On Excel for Windows, open **VBE → Tools → References** and enable **Microsoft Scripting Runtime**. On Mac, confirm that the reference is available in your Excel build; do not assume Windows COM/ActiveX features work on macOS. Some import and JSON paths may require additional platform review.

**Important:** Until the library is converted completely to late binding, importing the classes alone does not eliminate reference requirements. `CreateObject("Scripting.Dictionary")` at individual call sites does not replace the need for a reference while early-bound declarations remain.

## 2. Source files

Import these nine files from the repository (actual folder locations may vary):

| Module | Type | Purpose |
| --- | --- | --- |
| `cTable.cls` | Class | Main tabular data API |
| `cDataSource.cls` | Class | Input/source metadata |
| `cGroupBy.cls` | Class | Grouped operation handle |
| `cLineage.cls` | Class | Cell-source attribution engine |
| `cLineageRegistry.cls` | Class | Shared source and lineage-map registry |
| `cLog.cls` | Class | Status/error collection |
| `cLongVec.cls` | Class | Compact integer-vector helper |
| `cTableHistory.cls` | Class | Snapshot change/event analysis |
| `JsonConverter.bas` | Standard module | JSON parsing/serialization support |

**Preserve existing module attributes.** In particular `cLineageRegistry.cls` contains `Attribute VB_PredeclaredId = True` and must be **imported as an exported class** rather than pasted into a newly created blank class. Its default instance provides a shared source-ID namespace. Keep any third-party notices in `JsonConverter.bas` intact.

## 3. Import into Excel

1. Open the intended `.xlsm` workbook and open the **Visual Basic Editor** (typically **Alt+F11** on Windows; on Mac, use **Tools → Macro → Visual Basic Editor** or the Developer tab, depending on your Excel version).
2. Use **File → Import File…** to import each `.cls` class and the `.bas` standard module.
3. In the Project Explorer, confirm the eight classes are under **Class Modules** and `JsonConverter` is under **Modules**.
4. Enable the required reference(s) from **Tools → References** where supported.
5. Run **Debug → Compile VBAProject** to identify unresolved types, missing references, or source-version mismatches.
6. Save your workbook as `.xlsm`.

Do not rename the imported classes unless you are deliberately updating all dependent references.

### Source files distributed as `.txt`

If a source package contains exported VBA code under `.txt` filenames, make copies with their native exported extensions (`.cls` for class modules and `.bas` for the standard module), preserving the complete file contents and `Attribute` lines. Import those copies; do not paste text into blank classes.

## 4. Smoke test

Create a normal VBA standard module in your workbook and run:

```vb
Option Explicit

Public Sub SmokeTest_cTable()
    Dim t As New cTable
    t.read_excel_ws ThisWorkbook.Worksheets("Sheet1")
    Debug.Print "Rows: " & t.rowCount
    Debug.Print "Columns: " & t.ColumnCount
    Debug.Print t.head(5).ToString(5)
End Sub
```

**Before running:** put a header row and at least one data row on `Sheet1` (or change the sheet name). Open the Immediate Window in VBE (**Ctrl+G** on Windows). This test reads workbook data; it does not modify the source sheet.

## 5. Lineage is opt-in

To capture provenance at import, enable it **before** reading:

```vb
Dim t As New cTable
t.EnableLineage = True
t.read_excel_sheet "C:\Data\snapshot.xlsx", "AccountSnapshots"
Debug.Print t.LineageExplain(1, "AccountID")
```

Use a path appropriate to your machine. Enabling lineage *after* an import cannot retroactively recover cell origins. See [LINEAGE.md](LINEAGE.md).

## 6. Troubleshooting

| Symptom | Check |
| --- | --- |
| `User-defined type not defined` for `Scripting.Dictionary` | Missing/inaccessible Microsoft Scripting Runtime reference, or an unresolved early-bound dependency |
| `Can't find project or library` | A missing/broken reference in **Tools → References** |
| Errors mentioning `cLineageRegistry` members or source IDs | Import the original exported `.cls` preserving its `VB_PredeclaredId` attribute; do **not** use `New cLineageRegistry` |
| Errors opening workbook files | Verify file path, permissions, workbook format, and exact sheet name |
| Automation error while reading workbook metadata | Review `cDataSource.InitFromWorkbook` and the repository's Excel automation-error runbook; do not suppress unrelated errors globally |
| Compile failures on Mac | Identify Windows-only declarations/COM dependencies, and verify the specific Office for Mac version and supported references |

## 7. Updating the library

1. Back up your workbook/VBA project.
2. Review [CHANGELOG.md](../CHANGELOG.md) if present for API or behavior changes.
3. Replace the affected exported modules using the new **complete** exports, retaining class names and attributes. Avoid duplicate modules after import.
4. Compile the project and rerun regression tests, particularly filtering, joins, type conversion, lineage, and history functions.
5. Verify any workbooks or macros that depend on changed public signatures.

## See also

- [QUICKSTART.md](QUICKSTART.md) — example workflows
- [ARCHITECTURE.md](ARCHITECTURE.md) — module relationships
- [LINEAGE.md](LINEAGE.md) — provenance tracking
