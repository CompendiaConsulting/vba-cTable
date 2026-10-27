# Data Lineage in cTable

**Developer and user guide**
**Applies to:** `cTable`, `cDataSource`, `cLineage`, and `cLineageRegistry`
**Source reference:** The exported VBA modules in this repository. Consult `src/` for exact declarations and implementation details. Examples are illustrative unless explicitly identified as executed tests.

## Purpose

Lineage answers **"Where did this value come from?"** A `cTable` can attribute a value to its originating Excel worksheet, CSV/JSON input, another supported data source, or a synthetic source representing a manual write or computation. Lineage is useful for reconciling financial snapshots, explaining consolidated outputs, and tracing questionable values back to their inputs.

Lineage is **source attribution**, not a complete audit trail of every transformation or an immutable record of past cell values. A source label identifies provenance at the granularity the operation recorded. Do not interpret it as proof that every intermediate computation or row in the original workbook can be reconstructed.

## Quick start

Enable lineage **before** loading data:

```vb
Dim t As New cTable
Dim origin As cDataSource

t.EnableLineage = True
t.read_excel_sheet "C:\Data\AccountSnapshots_2026-10-01.xlsx", "AccountSnapshots"

Debug.Print t.LineageModeName
Debug.Print t.LineageExplain(1, "AccountID")

Set origin = t.LineageOfName(1, "AccountID")
Debug.Print origin.filePath
Debug.Print origin.sheetName
Debug.Print t.LineageReport()
```

Replace the sample path with one appropriate to your machine (Windows or macOS). Row numbers supplied to lineage methods are **1-based data-row indexes**, not Excel worksheet row numbers. In ordinary header-based imports, cTable row 1 represents the first data record.

> **Critical:** Enabling lineage after reading does not retroactively recover provenance. Enable it before each `read_*` operation. Disabling lineage with `t.EnableLineage = False` discards that table's lineage engine.

## Example: consolidate financial snapshots

The following example builds a single `cTable` from a preexisting `FileList` array of workbook paths. It retains per-workbook attribution where the import and append operations track it.

```vb
Dim t As cTable
Dim snapshot As cTable
Dim i As Long

Set t = Nothing
For i = LBound(FileList) To UBound(FileList)
    Set snapshot = New cTable
    snapshot.EnableLineage = True
    snapshot.read_excel_sheet CStr(FileList(i)), "AccountSnapshots"

    If t Is Nothing Then
        Set t = snapshot
    Else
        t.Append snapshot, strict:=True
    End If
Next i

If Not t Is Nothing Then
    Debug.Print "Rows: "; t.rowCount
    Debug.Print "Sources: "; t.LineageSourceCount
    Debug.Print t.LineageExplain(1, "AccountID")
    Debug.Print t.ColumnLineageLabel("AccountID")
    Debug.Print t.LineageReport()
End If
```

`Append` maps columns by name. In the bundled implementation, `strict:=True` rejects **incoming columns not present in the target**; it does not independently enforce bidirectional schema equality. For financial data, validate expected columns/types separately when necessary. Append is **UNION ALL-style**: it does not remove duplicate rows. A consolidated column can report mixed provenance because its cells came from several files; use `LineageOfName(row, column)` to inspect one value's origin.

`Append` rebuilds an array with the combined row count in the bundled implementation, so repeated appends can be expensive on large datasets. Consider batching and benchmarking if processing hundreds of large workbooks.

## Key public APIs

### cTable: configure and inspect

| Member | Use |
|---|---|
| `EnableLineage As Boolean` | Get/set whether this table tracks lineage; turn on before importing. |
| `LineageMode As ctLineageMode` | Machine-readable representation in use; `ctLinOff` when disabled. |
| `LineageModeName As String` | Human-readable lineage mode. |
| `LineageStoreName As String` | Description of current row-store representation. |
| `HasMixedRowLineage As Boolean` | Whether multiple row-source mappings are represented. |
| `HasMixedColumnLineage As Boolean` | Whether sources differ across columns. |
| `LineageSourceCount As Long` | Number of distinct sources recorded in the table; zero when off. |
| `LineageSources As Variant` | Array of labels for the distinct sources. |
| `LineageSourceObjects() As Variant` | Array of `cDataSource` objects representing recorded sources. |
| `LineageReport() As String` | Summary report of row/column shape, mode, and source labels. |
| `LineageColumnReport() As String` | Column-oriented lineage summary. |

### cTable: identify or modify attribution

| Member | Use |
|---|---|
| `LineageIdOf(rowNum, colNum) As Long` | Resolve the registered source ID for a cell. |
| `LineageOf(rowNum, colNum) As cDataSource` | Resolve source object using 1-based row and column numbers. |
| `LineageOfName(rowNum, colName) As cDataSource` | Resolve source object by column name. |
| `LineageExplain(rowNum, colName) As String` | One-line human-readable source explanation. |
| `ColumnLineage(colName) As cDataSource` | Return source only if the entire column has one source; otherwise `Nothing`. |
| `ColumnLineageLabel(colName) As String` | Label or diagnostic such as `(mixed: N sources)`. |
| `ColumnSourceCount(colName) As Long` | Count distinct source IDs for the column. |
| `WriteLineage As cDataSource` | Get/set the source attributed to manual `SetValue` writes. |
| `StampLineage(colName, ds)` | Attribute an entire column to a source. |
| `SetCellLineage(rowNum, colNum, ds)` | Attribute a specific cell to a source. |
| `DataSource As cDataSource` | Table-level source metadata; **not** necessarily a single definitive origin for a mixed-source table. |

Use `Set` when assigning returned `cDataSource` objects:

```vb
Dim ds As cDataSource
Set ds = t.LineageOfName(1, "AccountID")
Debug.Print ds.Describe()
```

Inspect a column after consolidation:

```vb
Debug.Print t.ColumnSourceCount("Expenditures")
Debug.Print t.ColumnLineageLabel("Expenditures")
Debug.Print t.LineageColumnReport()
```

### cDataSource: source identity and metadata

`cDataSource` is a metadata container. Depending on source type and available metadata, it exposes `filePath`, `sheetName`, `tableName`, `SourceType`, `Author`, `LastAuthor`, `lastSaveTime`, `ImportDate`, `ImportUser`, `Fingerprint`, `label`, and `Describe()`.

Supported initialization routines include `InitFromPath`, `InitFromWorkbook`, `InitFromPathDeclared`, `InitFromAccessPath`, `InitFromCSV`, `InitFromJSON`, `InitFromRange`, `InitCalculated`, and `InitManual`. Not all properties have a meaningful value for every source kind. Treat metadata as evidence to examine rather than guaranteed complete information.

Once the registry interns a `cDataSource`, the instance is frozen. Do not mutate registered source descriptors. Create a new source descriptor when the origin changes.

## Design and implementation

### Responsibilities

```text
cTable
  |  EnableLineage, LineageOfName, LineageReport, ...
  v
cLineage                 Per-table source IDs and compact row/column mapping
  |  resolves IDs and interned maps via
  v
cLineageRegistry         Shared source and map registries (default instance)
  |
  +-- cDataSource         Source metadata / identity / labels
  +-- integer source maps Reused by tables and rows
```

- **`cTable`** is the supported application-facing API. It owns or references a `cLineage` engine while tracking is enabled.
- **`cLineage`** stores compact lineage mappings, not a `cDataSource` object in every cell.
- **`cLineageRegistry`** interns source descriptions and source-ID vectors for efficient sharing.
- **`cDataSource`** describes each physical or synthetic origin.

### Compact representation

`cLineage` resolves a cell's source ID using the most specific applicable metadata. The implementation can share one source across the table, share source maps across rows, represent changing row bands, and add column or cell overlays. Modes exposed by `ctLineageMode` include `ctLinOff`, `ctLinUniform`, `ctLinByColumn`, and `ctLinByRow` (consult the enum in the source for additional modes). The representation may evolve after an append, join, column calculation, or targeted write; do not infer the provenance of every cell from a table's mode alone.

`cLineage.IdAt(row, column)` performs direct source lookup; cTable presents this using `LineageIdOf`, `LineageOf`, and `LineageOfName`.

### Shared registry: important import constraint

`cLineageRegistry` is deliberately implemented as a **predeclared VBA class instance**, not an ordinary class that callers should instantiate:

```vb
Dim id As Long
id = cLineageRegistry.SourceId(ds)  ' correct
```

**Do not use** `New cLineageRegistry`: that creates a private ID space incompatible with the shared registry. Preserve `Attribute VB_PredeclaredId = True` when exporting/importing the class. Import the actual exported `.cls` file rather than copying only its visible code into a blank class module.

Do not call `cLineageRegistry.ResetRegistries` while any active, lineage-enabled `cTable` relies on existing registry IDs. Resetting the registry invalidates the lookup space for those tables.

### Source identity and unknown origins

The registry represents sources by integer IDs. Reserved IDs include `ctSrcUnknown`, representing a source that was not recorded or was lost, and synthetic categories for manually assigned and calculated values. Interning uses `cDataSource.Fingerprint` so equivalent descriptions can share a registered ID. The precise fields involved in identity should be confirmed in `cDataSource.Fingerprint` before using ID equality as a compliance or deduplication criterion.

**Design principle:** if the engine cannot determine an exact source, it should mark provenance as *Unknown*, rather than silently claim a specific input.

## Transformation semantics and cautions

The bundled implementation has lineage-aware code paths for several transformations, including row gathers, column mappings, appends, and joins. The relevant internal methods are:

| Internal method | Responsibility |
|---|---|
| `x_ApplyRowGather` | Carry source mappings through row reordering/filtering. |
| `x_ApplyColumnMapping` | Remap attribution after column projection/reordering. |
| `x_AppendFrom` | Append incoming row attribution using the target-column mapping. |
| `x_InitJoin` | Build combined lineage for joined output rows and columns. |
| `StampColumn` / `SetCell` | Override column- or cell-level attribution. |
| `CloneMe` | Clone the lineage state. |

A calculation may be attributed to a synthetic *Calculated* source rather than preserve a full expression-level dependency graph. Do not assume a derived result can identify every input cell contributing to the calculation. Validate the specific operation's implementation and test it before making auditability claims.

### Suggested regression tests

When modifying operations involving rows, columns, joins, appends, or calculations, verify at least:

1. A single workbook import resolves its actual path and worksheet on representative cells.
2. Appending two files preserves the correct source for rows from each file.
3. Sorting/filtering does not detach source attribution from the corresponding data row.
4. Column reordering preserves attribution by logical column, not old positional index.
5. Joins correctly attribute columns originating on either side, including unmatched rows.
6. Manual cell writes and computed columns produce the documented source category.
7. A mixed-origin column returns `Nothing` from `ColumnLineage`, and a useful result from `ColumnLineageLabel`.
8. Disabled lineage reports its disabled state, rather than inventing source details.

Test the above with **actual values**, not only counts or a non-empty report. Record platform/Excel version and expected results in the repository's test documentation.

## Troubleshooting

| Symptom | Likely explanation / action |
|---|---|
| `LineageExplain` says tracking is disabled | Set `EnableLineage = True` **before** reading each file. |
| Existing values show `Unknown` | Tracking may have been enabled only after import, or a transformation could not preserve exact source attribution. Re-import with lineage enabled and inspect the operation. |
| `ColumnLineage` returns `Nothing` | A column may legitimately have multiple source IDs; use `ColumnLineageLabel`, `ColumnSourceCount`, or individual row lookups. |
| Source information is incomplete | Some workbook properties or file metadata may be unavailable; examine `cDataSource.Describe` and input metadata. |
| Registry/source identifiers are inconsistent | Check that `cLineageRegistry` was imported with its predeclared attribute and that registries were not reset while live tables existed. |
| Workbook import throws an Automation error | Diagnose workbook metadata access separately from worksheet reading. Consult the repository's Excel automation error runbook, if present. |

## Guidance for maintainers and AI coding agents

- Preserve public signatures and existing lineage semantics unless a breaking change is deliberately approved.
- Update lineage mappings whenever data rows or columns are gathered, reordered, appended, replaced, or projected.
- Do not store raw `cDataSource` objects in every cell; reuse the shared registry and mappings.
- Preserve `VB_PredeclaredId` on `cLineageRegistry` and never instantiate a private registry.
- Prefer *Unknown* over misattribution when exact lineage cannot be determined.
- Distinguish source attribution from transformation history and row-level data provenance.
- Add regression tests for both values and sources, especially after sorting, joining, and appending.
- Update this document and the changelog whenever public lineage behavior changes.

## Related source files

- `cTable.cls` — user-facing lineage operations and data transformations
- `cDataSource.cls` — source descriptors and metadata
- `cLineage.cls` — compact cell-lineage engine
- `cLineageRegistry.cls` — shared interned source/map registry
- `cLongVec.cls` — integer-vector support used by lineage internals
- `cTableHistory.cls` — separate time-oriented history capability; **not** synonymous with lineage

---

*Validate lineage behavior against the current repository source and representative test data before relying on it in an application.*
