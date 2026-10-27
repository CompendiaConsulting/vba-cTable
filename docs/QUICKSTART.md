# cTable Quickstart

cTable is a pandas-inspired table manipulation library for **Excel VBA**. It reads worksheet/workbook data into a VBA table, performs filtering, sorting, joining, and grouped calculations, and writes results back to Excel.

> These examples accompany the exported VBA modules in this repository. Consult `src/` for exact method declarations and implementation details. Examples are illustrative unless explicitly identified as executed tests. See [INSTALLATION.md](INSTALLATION.md) before running them.

## 1. Read data, filter, sort, export

Assume `Sheet1` in the current workbook has headers `AccountID`, `RunDate`, and `Expenditures`, with at least one data row. Add this to a standard VBA module:

```vb
Option Explicit

Public Sub Example_FilterAndSort()
    Dim t As New cTable
    Dim result As cTable

    t.read_excel_ws ThisWorkbook.Worksheets("Sheet1")
    Set result = t.Where("Expenditures", 1000, "gt")

    ' Newest snapshots first, then AccountID ascending.
    Call result.sort_values( _
        by:=Array("RunDate", "AccountID"), _
        ascending:=Array(False, True), _
        inplace:=True)

    result.to_excel_range ThisWorkbook.Worksheets("Results").Range("A1")
    Debug.Print "Selected records: "; result.rowCount
End Sub
```

Create a blank `Results` worksheet before running. `Where` produces a new cTable; the `sort_values` call sorts `result` in place. `to_excel_range` writes to the specified location; it does not create the destination sheet.

## 2. Combine many financial snapshot workbooks

Suppose `FileList` is a populated 1-D `Variant` array of full paths to Excel workbooks. Each workbook contains a `AccountSnapshots` sheet with compatible headers.

```vb
Public Function LoadAccountSnapshotsFiles(ByVal FileList As Variant) As cTable
    Dim t As New cTable
    Dim tFile As cTable
    Dim i As Long
    Dim filesDone As Long
    Dim started As Boolean

    For i = LBound(FileList) To UBound(FileList)
        Set tFile = New cTable
        tFile.EnableLineage = True  ' Before importing this file
        tFile.read_excel_sheet CStr(FileList(i)), "AccountSnapshots"

        If Not started Then
            Set t = tFile
            started = True
        Else
            t.Append tFile, strict:=True
        End If

        filesDone = filesDone + 1
        If filesDone Mod 10 = 0 Then
            Debug.Print "Files: "; filesDone; _
                        " | Records: "; t.rowCount
        End If
    Next i

    If started Then
        Call t.sort_values( _
            by:=Array("RunDate", "AccountID"), _
            ascending:=Array(False, True), _
            inplace:=True)
    End If

    Set LoadAccountSnapshotsFiles = t
End Function
```

`Append` performs **row stacking**, comparable to SQL `UNION ALL`: it does not remove duplicates. With `strict:=True`, an incoming column that does not exist in the target raises an error; this is not a complete, bidirectional schema-equality check. The first source establishes the destination columns. Empty inputs and missing files should be validated by the calling application; add error handling for production ingestion.

Example call (replace paths with your own):

```vb
Dim paths As Variant
Dim combined As cTable
paths = Array("C:\Data\status-01.xlsx", "C:\Data\status-02.xlsx")
Set combined = LoadAccountSnapshotsFiles(paths)
Debug.Print combined.rowCount
Debug.Print combined.LineageExplain(1, "AccountID")
```

`LineageExplain` can identify the contributing workbook for an imported cell when lineage was captured. Refer to [LINEAGE.md](LINEAGE.md) for details and caveats.

## 3. Explore a table

```vb
Debug.Print t.rowCount
Debug.Print t.ColumnCount
Debug.Print t.head(5).ToString(5)
Debug.Print t.info()
Debug.Print t.dtypes_string
```

Here `t` denotes an already populated `cTable` variable. `head` returns a new table containing the first rows; `info` and `dtypes_string` summarize the structure and inferred types.

## 4. Work with dates and types

Read methods infer types by default (`InferTypes:=True`). The library supports `ctVariant`, `ctString`, `ctLong`, `ctDouble`, `ctCurrency`, `ctDate`, and `ctBoolean`. VBA `Null` is the missing-value marker. Check or convert a column as needed:

```vb
Debug.Print t.dtype_string("Expenditures")
Call t.to_numeric("Expenditures")
Debug.Print t.dtype_string("Expenditures")
```

`to_numeric` uses a conversion error policy; verify unexpected source values before relying on totals.

## 5. Grouped time-series calculations

For cumulative financial data, **sort into the order required by your calculation before computing changes**. The `cGroupBy` handle delegates grouped operations to `cTable` and detects staleness after table mutations.

Example grouped first differences:

```vb
Dim g As cGroupBy
Dim changes As Variant

Set g = t.GroupBy("AccountID").Col("Expenditures")
changes = g.diff()
```

The `changes` array is a calculated result; this example does **not** automatically add a column. Consult the API reference for `AssignDiff` or cTable's column operations when attaching results. If you sort, append, or modify `t` after constructing `g`, create a fresh GroupBy handle.

## 6. Provenance and troubleshooting

```vb
Debug.Print t.LineageModeName
Debug.Print t.LineageSourceCount
Debug.Print t.LineageReport()
Debug.Print t.LineageExplain(1, "AccountID")
```

Cell lineage is disabled by default unless configured otherwise at the registry level. Enable it before importing to get useful source attribution; an imported multi-file table can have mixed sources, so `t.DataSource` may be `Nothing` even though individual cells have lineage.

## API conventions to remember

- Row and column **indexes are 1-based**. File arrays may be 0-based or 1-based; use `LBound` / `UBound`.
- Methods like `Where` and `head` return new `cTable` objects: assign with `Set`.
- Mutating methods change the existing table; check signatures and `inplace` arguments.
- Sorting affects the direction/sign of ordered differences and cumulative calculations.
- `Append` retains duplicate records.
- `Null` indicates missing data; filtering includes optional `nullBehavior` controls.
- Preserve workbook paths/headers carefully: imported data is only as consistent as the input snapshots.

## Further reading

- [INSTALLATION.md](INSTALLATION.md)
- [ARCHITECTURE.md](ARCHITECTURE.md)
- [LINEAGE.md](LINEAGE.md)
