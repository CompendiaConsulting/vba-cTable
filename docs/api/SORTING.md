# Sorting and row ranking

[← API Reference](../API-REFERENCE.md) · [Filtering](FILTERING.md) · [Joins](JOINS.md)

> This guide describes sorting in `src/cTable.cls`. Consult the source for exact declarations and implementation details. Examples are illustrative unless explicitly identified as executed tests.

## Choose the right operation

| Requirement | Recommended method | Source table changed? |
|---|---|---|
| Several sort keys with **different directions** | `sort_values` | Not by default |
| In-place sort with the **same direction** for every key | `SortBy` or `sort_values(..., inplace:=True)` | Yes |
| First `n` rows after a sort | `sort_values(...).head(n)` | No, unless sorting in place |
| Top or bottom `n` ranked rows | `nlargest` / `nsmallest` | No |

## `sort_values`

### Signature

```vb
Public Function sort_values( _
    ByVal by As Variant, _
    Optional ByVal ascending As Variant = True, _
    Optional ByVal inplace As Boolean = False, _
    Optional ByVal na_position As String = "last") As cTable
```

| Argument | Meaning | Default |
|---|---|---|
| `by` | Column name or array of column names; leftmost key is primary | Required |
| `ascending` | Single Boolean applies to all keys; Boolean array supplies one direction for each key | `True` |
| `inplace` | `False`: sort a clone and return it; `True`: sort current table and return `Nothing` | `False` |
| `na_position` | `"first"` or `"last"` for missing values, independent of sort direction | `"last"` |

**Example: newest snapshot first, ascending AccountID within each snapshot**

```vb
Dim sorted As cTable
Set sorted = t.sort_values( _
    by:=Array("RunDate", "AccountID"), _
    ascending:=Array(False, True))
```

This leaves `t` unchanged and assigns a newly sorted table to `sorted`.

**Example: sort `t` in place**

```vb
Call t.sort_values( _
    by:=Array("RunDate", "AccountID"), _
    ascending:=Array(False, True), _
    inplace:=True)
```

Do **not** use `Set t = t.sort_values(..., inplace:=True)`: the method explicitly returns `Nothing` in this mode.

**Example: common direction and missing values first**

```vb
Set sorted = t.sort_values( _
    by:=Array("AccountID", "RunDate"), _
    ascending:=True, _
    na_position:="first")
```

### Missing-value behavior

`p_CompareSortRowToBuffer` calls `p_IsNA`, treating `Null`, `Empty`, error values, and empty/whitespace text as missing. When either comparison value is missing, `na_position` determines placement; the sorting code **does not reverse** that placement for a descending column.

For example, `ascending:=False, na_position:="last"` still places missing values last rather than first.

### Types and comparison semantics

The method compares according to the declared `ctDataType`:

- `ctString` — case-insensitive text comparison.
- `ctLong`, `ctDouble`, `ctCurrency` — numeric comparison.
- `ctDate` — date serial-number comparison.
- `ctBoolean` — `False = 0`, `True = 1`.
- `ctVariant` — compatible Date/Boolean/numeric pairs are compared in their respective domains; otherwise values fall back to case-insensitive text comparison.

When data is unexpectedly sorted lexically (for example, numeric-looking strings), inspect `t.dtypes_string` and the import's type inference settings. `sort_values` expects **column names**; unlike some other cTable methods, its implementation converts entries to strings and resolves names, not numeric positions.

### Validation and row order

- An empty `by` array or unknown column name causes an error.
- If `ascending` is an array, its length must match the number of sort columns.
- Invalid `na_position` causes an error.
- The method uses a Shell-sort-like implementation; **do not assume stable ordering of rows whose sort keys compare equal**. Add a deterministic tie-break column if reproducible tie order matters.
- Sorting clears indexes because existing row-number mappings become invalid.
- Row-level lineage is permuted along with the data where needed, preserving provenance.

## `SortBy` (legacy)

### Signature

```vb
Public Sub SortBy( _
    ByVal COLS As Variant, _
    Optional ByVal NonAlphaAtEnd As Boolean = True, _
    Optional ByVal SortAscending As Boolean = True)
```

`SortBy` **always sorts the existing table** and has a **single `SortAscending` flag** for all supplied keys. It accepts one column name or an array of names. The `NonAlphaAtEnd` flag belongs to the legacy comparison behavior and is not the same parameter as `sort_values`'s `na_position`.

```vb
' Both columns ascending; modifies t.
t.SortBy Array("RunDate", "AccountID"), SortAscending:=True

' Both columns descending; modifies t.
t.SortBy Array("RunDate", "AccountID"), SortAscending:=False
```

For descending `RunDate` and ascending `AccountID`, **use `sort_values`**, not `SortBy`.

## `nlargest` and `nsmallest`

### Signatures

```vb
Public Function nlargest(ByVal n As Long, _
    ByVal columns As Variant, _
    Optional ByVal keep As String = "first") As cTable

Public Function nsmallest(ByVal n As Long, _
    ByVal columns As Variant, _
    Optional ByVal keep As String = "first") As cTable
```

These methods select and return rows with the largest or smallest **ranking values**. They do not modify the original table.

| Argument | Description |
|---|---|
| `n` | Desired row count. Zero or negative returns an empty table with the same schema. |
| `columns` | A ranking column name or array of column names / column references supported by the implementation. |
| `keep` | `"first"` or `"last"` selects which occurrences win a tie; `"all"` includes every row tied at the cutoff, potentially exceeding `n`. |

The implementation excludes candidate rows that have missing values in **any ranking column**. It also validates that the declared column type is rankable. Do not assume arbitrary text columns can be ranked as if this were a general sort.

```vb
Dim largest As cTable
Set largest = t.nlargest(10, "Expenditures")

' Include all records tied with the 10th largest value.
Set largest = t.nlargest(10, "Expenditures", keep:="all")
```

If your aim is *all rows in a specified display order*, use `sort_values`; if your aim is the *highest/lowest subset*, use these methods.

## Financial snapshot example

For a table where `Expenditures` is a **running total per AccountID**, a difference calculation needs an **ascending chronological order** to produce forward deltas. Only after calculating should the result be reordered for display:

```vb
Dim work As cTable

Set work = t.sort_values( _
    by:=Array("AccountID", "RunDate"), _
    ascending:=Array(True, True))

work.AddDiffColumn colName:="Expenditures", _
    by:="AccountID", _
    newColName:="Expenditures_Net", _
    baseline:=0

Call work.sort_values( _
    by:=Array("RunDate", "AccountID"), _
    ascending:=Array(False, True), _
    inplace:=True)
```

This assumes each AccountID has a meaningful sequence of snapshots. If an AccountID has duplicate `RunDate` values, supply another deterministic key or reconcile duplicates before differencing.

## Suggested regression tests

1. Multi-key mixed directions produce expected order.
2. `inplace:=False` does not alter the source; `inplace:=True` returns `Nothing` and changes row order.
3. `na_position` remains independent of ascending/descending directions.
4. Missing sort keys, mismatched Boolean arrays, and invalid policy strings raise errors.
5. Sorting with mixed source lineage retains the correct per-cell attribution.
6. Index lookups remain correct after in-place sorting and index rebuild.
7. `nlargest` / `nsmallest` exclude missing ranking entries and handle `keep:="all"` ties as documented.

## Source pointers

Search src/cTable.cls for the method names discussed in this guide. Source line numbers change as the implementation evolves.
