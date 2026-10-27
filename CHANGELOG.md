# Changelog

Notable changes to the **cTable VBA class library** are documented here. This file follows the general structure of [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

> **Historical baseline:** Historical entries were reconstructed from dated source comments and describe development milestones, not verified release tags. Entries explicitly identified as releases should correspond to repository release tags.

## [Unreleased]

This section is reserved for noteworthy changes made **after** adoption of this changelog. Add entries here in the same commit as the corresponding code changes, under `Added`, `Changed`, `Deprecated`, `Removed`, `Fixed`, or `Security` as applicable.

## [0.1.0] - 2026-10-26

### Initial Release

First release of the cTable VBA class library.

### Features

- Pandas-inspired tabular data manipulation in Excel VBA.
- Excel, CSV, Access, and JSON data import capabilities.
- Filtering, sorting, joins, and table concatenation.
- Grouping, aggregation, and statistical calculations.
- Column data types and missing-value handling.
- Optional cell-level data lineage and source attribution.
- Snapshot history and change detection.
- Excel worksheet and JSON export capabilities.

### Deployment

- Runs within Microsoft Excel VBA without requiring Python or pandas.
- Supports distribution as VBA modules embedded in macro-enabled Excel workbooks.
- Requires the documented VBA modules and platform-specific dependencies.

### Known Limitations

- Not all pandas functionality is implemented.
- Excel for Mac compatibility has not been fully validated and has known limitations
for Scripting.Dictionary objects
- Some operations rely on Windows-specific components.

### Added

### Changed

### Fixed
- **Implemented; pending runtime validation in Excel:** Make `WhereNot` and `WhereNotIn` honor `nullBehavior:="exclude"` (default) and `"include"` independently of predicate inversion. Previously, missing rows were incorrectly included by default in inverse filters.

### Removed
- **Implemented; pending runtime validation in Excel:** Remove the redundant `nullBehavior:="omit"` option. Replace existing uses with `"exclude"`.

## Pre-release development history

### 2026-10-07 — Filtering interface and missing-value behavior

**Changed**
- Extended `Where` and `WhereIn` to accept column names or 1-based column indexes through `Variant` column references.
- Added optional `nullBehavior` handling to filtering, defaulting to `"exclude"`.

**Added**
- `WhereNot` and `WhereNotIn` operations. Their inverse behavior applies to non-missing values; missing-value treatment follows `nullBehavior`.

**Compatibility note:** Code using the former named argument `ColumnName:=` may need adjustment if the parameter was renamed to `col`. Positional calls should be checked against the currently shipped signatures.

### 2026-08-27 — Extended source reading

**Added / Changed**
- Added `Unmerge` support to Excel read methods for working with merged cells.
- Extended `read_json` input handling to support additional file/remote-source scenarios, including SharePoint-related paths described in the original development notes. Availability depends on the environment and implementation.

### 2026-08-24 — Limited source reads

**Added**
- Optional `nrows` limit on `read_excel_sheet`, `read_excel_table`, `read_access_table`, `read_excel_range`, `read_excel_ws`, `read_csv`, and `read_json`.
- `nrows` counts data records rather than header rows; handling of read limits differs by source format.

### 2026-08-21 — Snapshot change-history helper

**Added**
- `cTableHistory` for analyzing time-stamped snapshots and producing change events for entities.
- `History` and `HistoryFor` integration in `cTable`, including lazy timeline/context caches and stale-state detection after table mutation.

**Behavioral note:** The first snapshot is treated as a baseline, not as a wave of newly added entities. Subsequent additions, removals, and attribute changes produce event records.

### 2026-08-03 — Grouped operations

**Added / Changed**
- `cGroupBy` handle to support grouped difference, shift, rank, and cumulative operations through `t.GroupBy(...).Col(...)`.
- Version checking to prevent reuse of stale grouping handles following table modifications.

**Behavioral note:** Difference and cumulative calculations use the table's current row order; group creation does not implicitly sort records.

### 2026-07-27 — Lineage storage refinements

**Changed**
- Refined `cLineage` tracking with compact source maps, row bands, and column/cell overlays for mixed-source tables.
- Established the principle that uncertain provenance is reported as `Unknown` rather than attributed incorrectly.

### 2026-07-26 — Cell-level provenance

**Added**
- Optional cell lineage tracking through `EnableLineage` and the `cLineage`, `cLineageRegistry`, and `cDataSource` components.
- Cell, column, and table reporting APIs such as `LineageOf`, `LineageOfName`, `LineageExplain`, `ColumnLineage`, and `LineageReport`.
- Support for calculated, manual, literal, missing, and unknown sources.

**Changed**
- Updated sorting and join paths to propagate lineage across row reordering and two-sided joins.
- Adjusted the `DataSource` compatibility property for tables with multiple origins; it can return `Nothing` when lineage has more than one source.

**Migration note:** `LineageSources` returns source labels; use `LineageSourceObjects` for actual `cDataSource` objects.

### 2026-07-23 — Column type system

**Added**
- `ctDataType` and `ctConvertErrors` enumerations and column dtype inspection/conversion APIs.
- Type inference on reads (`InferTypes`), numeric/date conversion helpers, and missing-value operations (`NA`, `NaT`, `isna`, `notna`, `dropna`, `fillna`).

**Changed**
- Excel export supports dtype-aware formatting where enabled.
- Derived and structurally modified tables propagate their column type information.

### Earlier evolution — pandas-shaped API

**Changed**
- Introduced pandas-inspired naming for several import/export, indexing, and uniqueness methods, including `read_excel_sheet`, `read_excel_range`, `to_excel_range`, `set_index`, `index_lookup`, and `unique`.

The source comments do not establish a verified release date for this set of naming changes, and this is **not** a complete inventory of early development.

## Maintenance rules

- Record externally observable additions, changes, fixes, and compatibility breaks; do not copy every internal code comment.
- Use formal version headings (`[x.y.z] - YYYY-MM-DD`) only when an actual version has been tagged or released.
- Preserve exact API spelling and defaults; verify claims against the implementation and regression tests.
- Link to relevant Git issues or commits when useful.
- Retain historical implementation rationale in source comments or architecture documentation rather than duplicating it here.

## Related documentation

- [Architecture](docs/ARCHITECTURE.md)
- [Lineage](docs/LINEAGE.md)
- [Quick start](docs/QUICKSTART.md)
- [Contributing](CONTRIBUTING.md)
