# Joins and row concatenation

[← API Reference](../API-REFERENCE.md) · [Filtering](FILTERING.md) · [Sorting](SORTING.md) · [Lineage](../LINEAGE.md)

> This guide describes joins and table combination in `src/cTable.cls`, including `merge`, `MakeJoin`, `Append`, `concat`, `append_many`, and `from_tables`. Consult the source for exact declarations and implementation details. Examples are illustrative unless explicitly identified as executed tests.

## Joins versus stacking

| Goal | Method | Result |
|---|---|---|
| Match records based on key fields | `merge` (preferred), `MakeJoin` (legacy) | New table with columns from both inputs |
| Stack two tables with equivalent headers | `concat` | New table with additional rows |
| Add rows to an existing table | `Append` | Mutates existing table |
| Stack many imported snapshots | `from_tables` or `append_many` | New table or existing-table mutation, respectively |

These are not interchangeable: a **join adds related columns**, while a **concat/append adds rows**, akin to SQL `UNION ALL`.

## `merge` — recommended for new code

### Signature

```vb
Public Function merge(ByVal right As cTable, _
    Optional ByVal how As String = "inner", _
    Optional ByVal on_arg As Variant, _
    Optional ByVal left_on As Variant, _
    Optional ByVal right_on As Variant, _
    Optional ByVal suffixes As Variant, _
    Optional ByVal indicator As Boolean = False, _
    Optional ByVal validate As String = "", _
    Optional ByVal sort As Boolean = False, _
    Optional ByVal indicator_name As String = "_merge") As cTable
```

`merge` **returns a new `cTable`**, leaving both input tables unchanged.

### Join types

| `how` | Matched rows | Unmatched rows |
|---|---|---|
| `"inner"` (default) | All matching left/right combinations | Omitted |
| `"left"` | All matching combinations | Left rows retained, right fields `Null` |
| `"right"` | All matching combinations | Right rows retained, left fields `Null` |
| `"outer"` or `"full"` | All matching combinations | Unmatched rows from both sides retained |
| `"cross"` | Cartesian product | No key conditions; row count can be `left.rowCount × right.rowCount` |

### Keys and column layout

| Parameter | Meaning |
|---|---|
| `on_arg` | Single shared column name or array of names. Matching key columns appear **once** in the result. |
| `left_on`, `right_on` | Names or arrays paired positionally; both key sets are **retained** as distinct columns. |
| *Neither specified* | The intersection of **all shared column names** is used as the composite join key. |
| `suffixes` | Two strings for overlapping non-key names; defaults to `Array("_x", "_y")`. |
| `indicator` | Adds a provenance-of-match-status column with `"left_only"`, `"right_only"`, or `"both"`. |
| `indicator_name` | Custom name for that column, default `"_merge"`. |
| `validate` | Declared cardinality; see below. |
| `sort` | If `True`, the result is sorted by key columns. |

**Important:** Leaving `on_arg` unspecified is *not* always the same as picking a primary key. If two financial tables share both `AccountID` and `RunDate`, the default joins on **both** names. Specify `on_arg:="AccountID"` when you intend to join only by AccountID.

### Examples

```vb
Dim actuals As cTable
Dim lookup As cTable
Dim enriched As cTable

' Assume actuals and lookup are populated cTables.
Set enriched = actuals.merge( _
    right:=lookup, _
    how:="left", _
    on_arg:="AccountID", _
    indicator:=True, _
    validate:="many_to_one")
```

The cardinality setting enforces that the **right-side AccountID** is unique; multiple actuals rows may refer to it.

**Composite key:**

```vb
Set enriched = actuals.merge( _
    right:=lookup, _
    how:="left", _
    on_arg:=Array("AccountID", "CategoryCode"), _
    suffixes:=Array("_snapshot", "_reference"))
```

**Different key names:**

```vb
Set enriched = actuals.merge( _
    right:=lookup, _
    how:="inner", _
    left_on:="AccountID", _
    right_on:="AccountCode")
```

In the last form, both `AccountID` and `AccountCode` are retained.

**Cross join:**

```vb
Set enriched = actuals.merge(right:=lookup, how:="cross")
```

Cross joins can grow very large. For instance, 10,000 left rows and 1,000 right rows can produce 10 million output rows; consider the memory footprint of VBA `Variant` arrays before using them.

### Cardinality validation

The `validate` string is normalized by lowercasing, trimming, and replacing colons with underscores. Valid choices include:

| Value | Constraint |
|---|---|
| `""` or `"many_to_many"` | No uniqueness check |
| `"one_to_one"` / `"1_1"` | Both key sets unique |
| `"one_to_many"` / `"1_m"` | Left key unique |
| `"many_to_one"` / `"m_1"` | Right key unique |

A failed validation raises a VBA error before generating the result. Validation is skipped for a cross join.

### How duplicate and missing keys behave

- The join engine indexes all row occurrences per key and emits **every matching pair**. Duplicate keys on both sides can multiply output rows.
- Composite keys are constructed using **length-prefixed components**, avoiding ambiguity between combinations such as `("AB", "C")` and `("A", "BC")`.
- Matching is **case-insensitive**, using `vbTextCompare` on dictionary keys.
- Missing keys are represented with a shared sentinel. As implemented, **missing keys can match missing keys**, rather than automatically being treated as nonmatches. This matters when source data contains blank AccountIDs or account IDs.
- Date key components are normalized to their serial numbers, avoiding locale-specific display-string differences.
- Key conversion otherwise relies largely on `CStr`, so validate key types and formatting upstream, particularly leading-zero identifiers.

### Column collisions and provenance

With the pandas-style `merge` naming convention, overlapping non-key columns receive suffixes on **both** sides (for example, `Amount_x` and `Amount_y`). When `on_arg` merges identical key columns, a right-only row in an outer/right join uses the right key value in the retained key field.

If lineage was enabled on the left table, the merge engine combines left and right provenance, assigns a Missing source to absent-side values, and specially handles coalesced keys from right-only rows. See [Lineage](../LINEAGE.md) for the source model. Do not treat the indicator column as a comprehensive lineage graph; it indicates only join match status.

## `MakeJoin` — legacy

### Signature

```vb
Public Function MakeJoin(ByVal JoinType As String, right As cTable, _
    ByVal leftKey As String, ByVal rightKey As String, _
    Optional LeftAsNames As Object = Nothing, _
    Optional RightAsNames As Object = Nothing) As cTable
```

`MakeJoin` supports **only `"inner"` and `"left"`**, with one pair of key columns. Both facades use the same internal `p_MergeCore`, but this older facade has different column-naming behavior: overlapping columns on the **right** are renamed with `_1`, `_2`, etc., rather than applying left and right suffixes symmetrically.

```vb
Dim joined As cTable
Set joined = actuals.MakeJoin("left", lookup, "AccountID", "AccountCode")
```

It retains the left key and drops the corresponding right key from the output. `LeftAsNames` and `RightAsNames` allow caller-supplied renaming dictionaries. Prefer `merge` for new development, composite keys, right/outer/cross joins, indicators, and uniqueness validation.

## Stacking rows is not a join

### `Append` — modify this table

```vb
Public Sub Append(other As cTable, Optional ByVal strict As Boolean = True)
```

`Append` adds `other`'s rows to the **existing** `cTable`, matching incoming columns by name rather than position. It preserves duplicates and can propagate each input's lineage.

**Caution:** With `strict:=True`, `Append` checks that every **incoming** column exists in the destination; it does **not** require the destination and incoming schema to have exactly the same set of columns. An incoming column not found in the destination raises an error. Missing destination fields in incoming rows are initialized to `Null`.

With `strict:=False`, unmatched incoming columns are **not added** to the destination; their values are effectively ignored. Existing destination columns not present in `other` get `Null` in appended rows.

### `concat` — return a combined copy

```vb
Public Function concat(ByVal other As cTable, _
    Optional ByVal strict As Boolean = True) As cTable
```

`concat` returns a **new** table. Unlike `Append`, its `strict:=True` path checks **both directions** for column-set equality (apart from its special handling of a left table with no columns). It then clones the left table and appends the right one. This is the better choice when an exact schema match matters.

```vb
Dim combined As cTable
Set combined = january.concat(february, strict:=True)
```

### `append_many` and `from_tables`

```vb
Public Function append_many(ByVal tables As Variant, _
    Optional ByVal strict As Boolean = True, _
    Optional ByVal SkipEmpty As Boolean = True) As Long

Public Function from_tables(ByVal tables As Variant, _
    Optional ByVal strict As Boolean = True, _
    Optional ByVal SkipEmpty As Boolean = True) As cTable
```

Both accept a `Collection` or array of `cTable` objects.

- `append_many` appends to the **current table** and returns the number of input tables appended.
- `from_tables` constructs and returns a **new table**, cloning the first eligible table and appending subsequent tables. Inputs are not modified.
- `SkipEmpty:=True` skips tables with zero rows or zero columns.

`from_tables` and `append_many` delegate later appends to `Append`; therefore, **`strict:=True` there does not provide the same bidirectional schema check as `concat`**. Use explicit `RequireColumns` / schema comparison when complete column-set equality is required.

## Multiple-workbook example

For identically structured `AccountSnapshots` worksheets in `FileList`:

```vb
Dim tables As Collection
Dim builder As cTable
Dim snapshot As cTable
Dim combined As cTable
Dim i As Long

Set tables = New Collection

For i = LBound(FileList) To UBound(FileList)
    Set snapshot = New cTable
    snapshot.EnableLineage = True
    snapshot.read_excel_sheet CStr(FileList(i)), "AccountSnapshots"
    tables.Add snapshot
Next i

Set builder = New cTable
Set combined = builder.from_tables(tables, strict:=True)

Debug.Print "Snapshot records: " & combined.rowCount
Debug.Print combined.LineageReport
```

The example assumes `FileList` is a populated array of workbook paths and all workbooks have the sheet and expected headers. The collection approach retains every input table until the collection is released; for very large datasets, consider the cost of retaining multiple table arrays.

**Why initialize from the first actual table?** Calling `Append` directly on a newly constructed cTable with zero columns cannot establish a schema. `from_tables` handles initialization by cloning the first eligible table. For incremental processing, set your master table to the first successfully read table (or its clone), then call `Append` for remaining tables.

## Suggested regression tests

1. Inner, left, right, outer, and cross joins with unique keys.
2. One-to-many and many-to-many multiplication; each cardinality validation mode.
3. Composite keys, differently named keys, and `on_arg` default behavior when multiple columns are shared.
4. Missing key matches, typed dates, case-only differences, and leading-zero IDs.
5. Overlapping column names, custom suffixes, optional match indicator, and right-only coalesced keys.
6. Provenance for present and absent sides, especially `outer`/`right` joins.
7. `Append` versus `concat` strictness when one table has extra or missing columns.
8. Empty tables, `SkipEmpty`, `from_tables` initialization, and source-table nonmutation.

## Source pointers

Search src/cTable.cls for the method names discussed in this guide. Source line numbers change as the implementation evolves.
