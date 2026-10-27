# cTable VBA API Reference

> This reference accompanies the exported VBA modules in this repository. Exact declarations and implementation details are defined in `src/`. Examples are illustrative unless explicitly identified as executed tests.

`cTable` is a pandas-inspired, in-memory table abstraction for Excel VBA. Its primary object owns a rectangular `Variant` data array, column headers, optional type metadata, indexes, diagnostics, history state, and optional cell-level source lineage. Related classes implement grouped calculations, provenance, and snapshot change analysis.

## Start here

- [Installation](INSTALLATION.md) — import classes, configure dependencies, and compile.
- [Quick Start](QUICKSTART.md) — end-to-end examples.
- [Architecture](ARCHITECTURE.md) — storage, object responsibilities, and mutation contracts.
- [Lineage](LINEAGE.md) — source tracking and attribution.
- **Detailed method guides:** [Filtering](api/FILTERING.md) · [Sorting](api/SORTING.md) · [Joins](api/JOINS.md).
- [Changelog](../CHANGELOG.md) — development history and release notes.

## Key conventions

| Topic | Contract in the bundled implementation |
|---|---|
| Indexing | Public row and column positions are **1-based**. |
| Object results | Use `Set` when a method returns a `cTable` or another VBA object. |
| Mutability | Read the individual method contract; operations are **not** uniformly in-place. |
| Missing values | `Null` is the principal missing marker; `p_IsNA` also recognizes `Empty`, errors, and blank/whitespace strings. |
| Data types | `ctVariant`, `ctString`, `ctLong`, `ctDouble`, `ctCurrency`, `ctDate`, and `ctBoolean`. |
| Column reference | Many APIs accept a column **name or index**; others require a **name**. Use each declaration. |
| Comparisons | Some are text-insensitive, numeric-aware, or dtype-aware; behavior varies by operation. |
| Lineage | Enable **before** data import to preserve the originating source. |
| Source | `src/cTable.cls` and its companion classes are authoritative in your checked-out repository. |

## Typical pipeline

```vb
Dim t As cTable
Dim filtered As cTable
Dim sorted As cTable

Set t = New cTable
t.EnableLineage = True
t.read_excel_sheet filePath:="", sheetName:="AccountSnapshots"

Set filtered = t.Where("Expenditures", 0, "gt")
Set sorted = filtered.sort_values( _
    by:=Array("RunDate", "AccountID"), _
    ascending:=Array(False, True))

sorted.to_excel_range ThisWorkbook.Worksheets("Results").Range("A1")
```

The example assumes `AccountSnapshots` exists in the workbook containing the code, `Results` already exists, and the named columns are present. A blank `filePath` selects `ThisWorkbook` in the bundled reader.

## API by task

This is a **navigation map**, not a substitute for the exact signatures in the source. Names below were checked against the bundled public declarations. Methods beginning `x_` are integration hooks, not normal end-user entry points.

### Import, export, and construction

| Task | Methods | Notes |
|---|---|---|
| Excel input | `read_excel_sheet`, `read_excel_table`, `read_excel_ws`, `read_excel_range` | Headers, type inference, row limits, optional source-row capture. |
| Other input | `read_access_table`, `read_csv`, `read_json`, `read_text` | Environment-specific dependencies; see installation guide. |
| In-memory construction | `from_array` | Load supplied headers and data array. |
| Excel output | `to_excel_range`, `ReplaceExcelTableFrom_cTable` | Write into existing sheet/range or Excel table. |
| JSON output | `to_json` | Requires relevant JSON/file I/O support. |
| Copy, display | `Clone`, `Describe`, `ToString`, `info` | Different output types and levels of detail. |

### Inspect, validate, and address rows/columns

`rowCount`, `ColumnCount`, `headers`, `HeadersArray`, `data`, `DataArray`, `shape`, `columnNumber`, `columnName`, `GetHeaderIndex`, `ColumnExists`, `RequireColumns`, `MissingColumns`, `UnexpectedColumns`, `GetValue`, `GetValueByName`, `SetValue`, `SetValueByName`, `GetColumnArrayByName`, `Log`.

### Filtering, selection, and duplicates

| Method | Effect | Detailed reference |
|---|---|---|
| `Where`, `WhereNot` | Single-column scalar predicate and inversion. Returns a new table. | [Filtering](api/FILTERING.md#scalar-filtering-where-and-wherenot) |
| `WhereIn`, `WhereNotIn` | Membership predicates. Returns a new table. | [Filtering](api/FILTERING.md#membership-filtering-wherein-and-wherenotin) |
| `SelectColumns` | Project a subset of columns. | [Filtering](api/FILTERING.md#related-selection-operations) |
| `head`, `tail`, `dropna`, `drop_duplicates` | Row selection and duplicate handling. | [Filtering](api/FILTERING.md#related-selection-operations) |
| `CountTrue` | Count `True` values in a Boolean mask. | Source declaration |

### Sorting and ranking

| Method | Effect | Detailed reference |
|---|---|---|
| `sort_values` | Per-column sort direction; copy by default, optional in-place. | [Sorting](api/SORTING.md#sort_values) |
| `SortBy` | Legacy in-place multi-key sort using one direction. | [Sorting](api/SORTING.md#sortby-legacy) |
| `nlargest`, `nsmallest` | Top/bottom rows, optional tie handling. | [Sorting](api/SORTING.md#nlargest-and-nsmallest) |
| `rank`, `AddRankColumn`, `SetRankColumn` | Rank values rather than rearranging rows. | Source declaration |

### Joins, concatenation, and aggregation

| Method | Effect | Detailed reference |
|---|---|---|
| `merge` | Inner, left, right, outer, cross joins; composite keys; validation. | [Joins](api/JOINS.md#merge--recommended-for-new-code) |
| `MakeJoin` | Legacy inner/left join, one pair of keys. | [Joins](api/JOINS.md#makejoin--legacy) |
| `Append`, `concat` | Stack records by column name. | [Joins](api/JOINS.md#stacking-rows-is-not-a-join) |
| `append_many`, `from_tables` | Combine collections or arrays of cTables. | [Joins](api/JOINS.md#multiple-workbook-example) |
| `GroupBy` | Return a grouped-operation handle (`cGroupBy`). | `cGroupBy.cls` |
| `AggregateDict`, `sum`, `mean`, `median`, `min`, `max`, `count`, `nunique`, `var`, `std`, `quantile`, `unique`, `GetUniqueValues` | Aggregate or inspect values. | Source declarations |

### Calculation and transformation

- **Differences/shifts:** `diff`, `shift`, `AddDiffColumn`, `SetDiffColumn`, `WithDiffColumn`, `AddShiftColumn`, `SetShiftColumn`.
- **Cumulative:** `cumsum`, `cumprod`, `cummax`, `cummin` and corresponding `Add...Column` / `Set...Column` variants.
- **Calculated columns:** `AddComputedColumn`, `SetComputedColumn`, `WithComputedColumn`, `SetColumnValue`.
- **Windows and expanding series:** `rolling`, `expanding`, `AddRollingColumn`, `SetRollingColumn`, `WithRollingColumn`, `AddExpandingColumn`, `SetExpandingColumn`.
- **Shape:** `pivot`, `melt`, `AddRow`, `DropRowInPlace`, `InsertColumn`, `DropColNameInPlace`, `DropColNoInPlace`, `DropColumnIfExists`, `RenameColumn`, `rename_columns`, `RenameColsByArr`, `ReorderColumnInPlace`, `ReorderColumnRelativeInPlace`.
- **Strings:** `str_trim`, `str_upper`, `str_lower`, `str_title_case`, `str_replace`, `str_contains`, `str_count`, `AddStrCountColumn`, `SetStrCountColumn`.
- **Dates:** `dt`, `dt_part`, `dt_name`, `dt_fy`, `dt_fq`, `dt_format`, `dt_value`, `dt_format_value`, `AddDtColumn`, `SetDtColumn`, `WithDtColumn`, `AddDtFormatColumn`, `SetDtFormatColumn`, `WithDtFormatColumn`.

### Types, nulls, and indexing

`dtype`, `dtype_at`, `dtype_string`, `dtypes`, `dtypes_string`, `dtypes_array`, `set_dtypes_from`, `dtype_name`, `dtype_from_name`, `convert_dtypes`, `astype`, `astype_many`, `with_astype`, `to_numeric`, `to_datetime`, `NA`, `NaT`, `IsNAValue`, `isna`, `notna`, `hasnans`, `dropna`, `fillna`, `FillNull`, `set_index`, `index_lookup`.

### Provenance, history, and diagnostics

- **Source and lineage:** `DataSource`, `EnableLineage`, `WriteLineage`, `LineageIdOf`, `LineageOf`, `LineageOfName`, `LineageExplain`, `ColumnLineage`, `LineageSourceObjects`, `LineageSources`, `LineageMode`, `LineageModeName`, `LineageStoreName`, `LineageSourceCount`, `ColumnLineageLabel`, `ColumnSourceCount`, `LineageColumnReport`, `LineageReport`, `HasMixedRowLineage`, `HasMixedColumnLineage`, `StampLineage`, `SetCellLineage`.
- **History:** `history`, `HistoryFor` (implemented with `cTableHistory`).
- **Diagnostics:** `Log`, `Describe`, `ToString`, `info`.

For the operational model and ownership rules, consult [Lineage](LINEAGE.md) and [Architecture](ARCHITECTURE.md).

## Method source pointers

Search src/cTable.cls for the method names discussed in this guide. Source line numbers change as the implementation evolves.

| Method | Bundled source line |
|---|---:|
| `Where`, `WhereNot`, `WhereIn`, `WhereNotIn` | 6137, 6143, 6149, 6154 |
| `SelectColumns` | 6165 |
| `MakeJoin`, `merge` | 6231, 6813 |
| `SortBy`, `sort_values` | 6969, 7051 |
| `nlargest`, `nsmallest` | 8598, 8613 |
| `concat`, `Append` | 8964, 9016 |
| `read_excel_sheet` | 4496 |
| `to_excel_range` | 9107 |
| `append_many`, `from_tables` | 14250, 14327 |

## Documentation maintenance

1. Treat exported VBA `.cls` / `.bas` source as authoritative; do not infer behavior solely from pandas.
2. When changing a signature, update the method guide, examples, and `CHANGELOG.md` in the same pull request.
3. When changing a branch of behavior (especially missing values, joins, lineage, and in-place mutation), add executable regression tests; documentation is **not** proof of execution.
4. Regenerate any machine-oriented API index from source; do not hand-maintain line numbers through refactors.
5. Record known discrepancies between source comments and actual branches explicitly; see [Filtering: implementation caveat](api/FILTERING.md#missing-cell-policy).
