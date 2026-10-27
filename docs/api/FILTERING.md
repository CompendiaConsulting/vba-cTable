# Filtering — `Where`, `WhereNot`, `WhereIn`, `WhereNotIn`

[← API Reference](../API-REFERENCE.md) · [Sorting](SORTING.md) · [Joins](JOINS.md)

> This guide describes filtering in `src/cTable.cls`, including the public filter methods and the `p_NullPolicy`, `p_WherePredicate`, and `p_FilterRows` helpers. Examples are illustrative unless explicitly identified as executed tests.

## Overview

These four methods return a **new `cTable` containing selected rows**, leaving the source table unchanged. Column names or 1-based integer positions identify the column; the result preserves the source row order, available data type metadata, and enabled cell lineage.

| Method | Predicate | Result |
|---|---|---|
| `Where` | Matches a scalar comparison | `cTable` |
| `WhereNot` | Inverts a scalar comparison | `cTable` |
| `WhereIn` | Matches an element of an array | `cTable` |
| `WhereNotIn` | Inverts array membership | `cTable` |

## Scalar filtering: `Where` and `WhereNot`

### Signatures

```vb
Public Function Where(ByVal col As Variant, ByVal value As Variant, _
    Optional ByVal operatorType As String = "eq", _
    Optional ByVal nullBehavior As String = "exclude") As cTable

Public Function WhereNot(ByVal col As Variant, ByVal value As Variant, _
    Optional ByVal operatorType As String = "eq", _
    Optional ByVal nullBehavior As String = "exclude") As cTable
```

| Argument | Accepted values | Default / meaning |
|---|---|---|
| `col` | Column name or **1-based** numeric column index | Required; invalid or missing column raises an error. |
| `value` | Scalar value, including a missing value for special `eq` / `noteq` checks | Required. |
| `operatorType` | `eq`, `noteq`, `gt`, `gte`, `lt`, `lte`, `startswith`, `endswith`, `contains` | `eq`; case-insensitive operator spelling. |
| `nullBehavior` | `exclude`, `include` | `exclude`; missing-cell selection is independent of inverse predicate direction. |

### Examples

```vb
Dim t As cTable
Dim result As cTable
Set t = New cTable
t.read_excel_sheet filePath:="", sheetName:="AccountSnapshots"

' Expenditures greater than $100,000.
Set result = t.Where("Expenditures", 100000, "gt")

' Match a specific AccountID; the first column is #1, not #0.
Set result = t.Where("AccountID", "A100")

' Negative predicate: AccountID does not equal A100.
Set result = t.WhereNot("AccountID", "A100")

' Text containing an identifying fragment (case-insensitive).
Set result = t.Where("AccountID", "ops", "contains")
```

### Supported operators

| Operator | Result for non-missing values |
|---|---|
| `eq` | Equality (numeric equality if both operands are numeric; otherwise case-insensitive string comparison). |
| `noteq` | Inequality with the same numeric/text comparison choices. |
| `gt` / `gte` | Greater than / greater than or equal to. |
| `lt` / `lte` | Less than / less than or equal to. |
| `startswith` / `endswith` | Case-insensitive prefix/suffix comparison. |
| `contains` | Case-insensitive substring search. |

**Type nuance:** Relational comparisons use numerical comparison when *both* operands satisfy `IsNumeric`; otherwise they use VBA relational string comparison. Do not assume numeric ordering if financial amounts were imported as strings containing nonnumeric formatting. Equality may treat numeric-looking strings and numeric values as equal. String relational behavior can depend on VBA comparison settings; don't assume it follows the same rules as `eq`.

## Membership filtering: `WhereIn` and `WhereNotIn`

### Signatures

```vb
Public Function WhereIn(ByVal col As Variant, ByVal values As Variant, _
    Optional ByVal nullBehavior As String = "exclude") As cTable

Public Function WhereNotIn(ByVal col As Variant, ByVal values As Variant, _
    Optional ByVal nullBehavior As String = "exclude") As cTable
```

`values` must be a **non-empty array**, typically `Array(...)`. The implementation builds a case-insensitive `Scripting.Dictionary` keyed by `CStr` of each non-missing value; comparisons therefore differ from the numeric-aware scalar `Where` comparisons.

```vb
Dim selected As cTable
Dim excluded As cTable

Set selected = t.WhereIn("AccountID", Array("A100", "B200", "C300"))
Set excluded = t.WhereNotIn("AccountID", Array("A100", "B200"))
```

A missing value in the membership array explicitly asks the operation to test for missing cells:

```vb
Set selected = t.WhereIn("AccountID", Array("A100", Null))
```

## Missing values and `nullBehavior`

The underlying missing-value predicate treats `Null`, `Empty`, VBA error values, and blank/whitespace-only strings as missing. A missing **search value** is a distinct case from a missing **table cell**.

### Missing search value

Scalar missing-value queries are allowed only with `eq` and `noteq`:

```vb
Set result = t.Where("AccountID", Null)               ' Missing AccountIDs
Set result = t.WhereNot("AccountID", Null)            ' Non-missing AccountIDs
Set result = t.Where("AccountID", Null, "noteq")      ' Non-missing AccountIDs
```

`Where("AccountID", Null, "gt")` raises an error. In a membership query, `Array(Null)` explicitly includes missing values; the inverted membership filter excludes them.

### Missing-cell policy

For non-missing search values and membership arrays **without** an explicit missing-value element, the revised policy is independent of whether the filter is positive or inverse:

| `nullBehavior` | `Where` / `WhereIn`: select missing rows? | `WhereNot` / `WhereNotIn`: select missing rows? |
|---|---|---|
| `"exclude"` (default) | No | No |
| `"include"` | Yes | Yes |

`"omit"` is removed as a redundant option; callers using it must change to `"exclude"`. A missing query value (for example `Where("AccountID", Null)`) or an explicit missing entry in a membership array (for example `Array("A100", Null)`) continues to express a direct missing-value predicate and is not overridden by `nullBehavior`.

```vb
Set result = t.WhereNot("AccountID", "A100")
Set result = t.WhereNotIn("AccountID", Array("A100"))
Set result = t.WhereNot("AccountID", "A100", nullBehavior:="include")
```

**Version note:** The current source applies nullBehavior independently of predicate inversion when the search does not explicitly request missing values. Explicit missing-value searches are handled as predicates. This description is based on static source inspection; runtime verification in Excel remains necessary.

## Combining filters

Because each filter returns another `cTable`, apply successive calls to a temporary result:

```vb
Dim current As cTable

Set current = t.Where("RunDate", #10/1/2026#, "gte")
Set current = current.Where("Expenditures", 0, "gt")
Set current = current.WhereIn("AccountID", Array("A100", "B200"))
```

Each additional call selects from the previously selected rows (**AND** semantics). Be careful not to confuse `WhereNot` with `operatorType:="noteq"` around missing values; one inverts the entire evaluated predicate, whereas `noteq` is an operator.

## Related selection operations

| Operation | Purpose | Mutates source? |
|---|---|---|
| `SelectColumns(COLS)` | Return a table containing selected columns. | No |
| `head(n)` / `tail(n)` | Return the first / last `n` rows. | No |
| `dropna(subsetCols, how)` | Build a table excluding rows according to missing-value rules. | No |
| `drop_duplicates(subset, keep, inplace)` | Remove duplicate rows using selected key columns. | Only with `inplace:=True` |

### Example: project columns after filtering

```vb
Dim output As cTable
Set output = t.Where("Expenditures", 0, "gt")
Set output = output.SelectColumns(Array("RunDate", "AccountID", "Expenditures"))
```

## Validation and performance

- Use `t.RequireColumns Array("AccountID", "Expenditures")` to fail early if headers differ from your expectation.
- `WhereIn` / `WhereNotIn` prepare a dictionary for lookup and are useful for a list of accepted categories.
- Each successive filtering call creates a new `cTable` and copies the selected rows. With large snapshots, avoid many unnecessary intermediate copies.
- Filters preserve row order; if you need newest-to-oldest output, follow with [sorting](SORTING.md).
- If source lineage is enabled, derived rows retain corresponding provenance through the implementation's row-gather path; see [Lineage](../LINEAGE.md).

## Suggested regression cases

These are **tests to implement and execute**, not tests claimed as passed:

1. Name and integer-index column selection produce the same rows.
2. All nine scalar operators produce expected results on numeric and string data.
3. `WhereIn` is case-insensitive for strings; invalid or empty arrays raise errors.
4. All four methods are tested with missing and non-missing cells, including `Null`, `Empty`, and whitespace strings.
5. `nullBehavior` values `exclude`, `include`, and invalid input (including rejected `omit`) are covered independently.
6. The source table remains unchanged, source row order persists, and lineage/dtypes propagate correctly.

## Source pointers

Search src/cTable.cls for the method names discussed in this guide. Source line numbers change as the implementation evolves.
