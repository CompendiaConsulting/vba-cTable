# cTable VBA Library

A pandas-inspired library for working with tabular data in Excel VBA. Load data into memory, filter and join records, calculate grouped values, reshape tables, and write results back to Excel. Supporting classes provide source metadata, optional cell lineage, snapshot change history, and diagnostics.

The public API is defined by the VBA source in [`src/`](src/). Similarity to pandas is a naming and design aid; verify signatures and behavior against this library.

**Current version:** 0.1.0 (initial release candidate)

See [CHANGELOG.md](CHANGELOG.md) for release history.

## Why cTable?

Python's pandas library provides powerful capabilities for data manipulation, analysis, and ETL (Extract, Transform, Load). However, many organizations operate in controlled IT environments where Python installation, execution, or deployment is restricted by security policies, software approval processes, or application management requirements.

These restrictions can prevent Excel VBA developers from taking advantage of pandas-style data processing, even when their organizations already use Microsoft Excel extensively.

**cTable was developed to bridge this gap.**

cTable brings pandas-inspired data manipulation and ETL capabilities into the Excel VBA environment, allowing developers to build sophisticated data-processing applications directly within macro-enabled Excel workbooks (`.xlsm`).

This approach offers several advantages:

- **No Python runtime required.** Data processing executes through Excel VBA rather than Python.
- **Simplified deployment.** Developers can distribute macro-enabled workbooks containing the required VBA components, without requiring end users to install Python, pandas, or a separate application.
- **Familiar development environment.** Excel VBA programmers can use concepts such as filtering, joins, grouping, aggregation, sorting, and data transformation.
- **Works within established Excel workflows.** Organizations can build on existing spreadsheet processes instead of introducing an entirely new technology stack.
- **Data traceability.** Optional lineage capabilities help identify the sources of imported and transformed data.

cTable is not intended to reproduce every capability of pandas or match its performance. Rather, it makes a useful subset of pandas-inspired data-processing functionality available to Excel VBA developers, particularly in environments where deploying Python-based solutions is impractical or prohibited.

**Deployment note:** cTable does not require Python, but its VBA modules and external dependencies must be included in the workbook. Microsoft Excel must be available, macros must be permitted by organizational policy, and certain features may require additional Windows components or approved VBA references.

## Project status and support

This repository is the authoritative upstream for cTable, maintained by Brian A. Crawford. Contributions, bug reports, and documentation improvements are welcome. Maintenance and support are provided on a best-effort basis; response times, fixes, and release dates are not guaranteed. See CONTRIBUTING.md for contribution guidance.

## Folder Contents

| File | Purpose |
|---|---|
| [`cTable.cls`](src/cTable.cls) | Main table class: import/export, row and column operations, filtering, joins, sorting, statistics, types, and calculations. |
| [`cDataSource.cls`](src/cDataSource.cls) | Source identity and metadata, including file, sheet/table, workbook authorship, save time, and import information. |
| [`cGroupBy.cls`](src/cGroupBy.cls) | Grouped calculation interface for differences, shifts, ranks, cumulative values, rolling/expanding calculations, and date parts. |
| [`cLineage.cls`](src/cLineage.cls) | Stores and transforms source attribution for table cells. Normally accessed through cTable. |
| [`cLineageRegistry.cls`](src/cLineageRegistry.cls) | Shared source and lineage-map registry, reserved source identifiers, and the project lineage default. Requires its predeclared class instance. |
| [`cLongVec.cls`](src/cLongVec.cls) | Compact Long-vector helper used by the lineage engine. |
| [`cLog.cls`](src/cLog.cls) | Collects status and error messages. Available through t.Log. |
| [`cTableHistory.cls`](src/cTableHistory.cls) | Builds cached change-event tables from repeated snapshots. Access through t.History and t.HistoryFor. |
| [`JsonConverter.bas`](src/JsonConverter.bas) | Bundled VBA-JSON parser/serializer and date/time helpers. Retain its third-party notices. |

Sources distributed as *_cls.txt or *_bas.txt are text copies of exported VBA modules. The filenames above are the recommended import filenames; the class/module names remain those declared in the source attributes.

## Setup

- Use a macro-enabled Excel workbook (.xlsm) or add-in (.xlam).
- Import all eight class modules and JsonConverter.bas through the VBA editor’s File > Import File command. If starting from .txt attachments, create .cls copies of the classes and a .bas copy of JsonConverter. Preserve the original text copies.
- Preserve all exported headers and Attribute lines. In particular, cLineageRegistry requires Attribute VB_PredeclaredId = True. Import it rather than pasting it into a blank class, and use cLineageRegistry directly rather than creating it with New.
- In Tools > References, enable Microsoft Scripting Runtime. The source includes early-bound Scripting.Dictionary declarations; late-bound calls elsewhere do not remove this requirement.
- Run Debug > Compile VBAProject, then run the quick start against a small workbook.

## Host and Platform Dependencies

The library uses Excel objects such as Workbook, Worksheet, Range, and ListObject. Use Excel VBA as the host for these examples.

The current source also uses Windows components on particular code paths:

| Feature | Additional dependency |
|---|---|
| Access-table import | ADO and an appropriate Access database provider. |
| CSV text I/O and JSON file export | ADODB.Stream. |
| HTTP reads | MSXML or WinHTTP, according to the selected authentication/transport mode. |
| Regular-expression matching | VBScript.RegExp. |

Treat this baseline as Windows-oriented. JsonConverter contains Mac conditional code, but that does not make the complete library portable to macOS. Dictionary availability, COM dependencies, and Excel-specific paths must be addressed and tested before using the complete library on a Mac.

## Quick Start

This example reads the AccountSnapshots sheet in the workbook containing the code. Headers must include RunDate, AccountID, and numeric Expenditures; an output sheet named NetExpenditures must already exist. Each AccountID should have one record per snapshot date.

```vb
Option Explicit

Public Sub BuildNetExpenditures()
    Dim t As cTable
    Dim result As cTable
    Dim output As Worksheet

    Set t = New cTable
    t.EnableLineage = True
    t.read_excel_sheet filePath:="", sheetName:="AccountSnapshots"
    t.RequireColumns Array("RunDate", "AccountID", "Expenditures")

    ' Order each AccountID chronologically before calculating differences.
    Set result = t.sort_values( _
        by:=Array("AccountID", "RunDate"), ascending:=Array(True, True))
    result.AddDiffColumn colName:="Expenditures", by:="AccountID", _
        newColName:="Expenditures_Net", baseline:=0

    ' Display the newest snapshots first.
    Set result = result.sort_values( _
        by:=Array("RunDate", "AccountID"), ascending:=Array(False, True))
    Set output = ThisWorkbook.Worksheets("NetExpenditures")
    result.to_excel_range output.Range("A1")

    Debug.Print "Rows: " & result.rowCount
    Debug.Print result.LineageReport
End Sub
```

An empty filePath selects ThisWorkbook for read_excel_sheet; supply a full workbook path to read another file. The reader defaults to headers, type inference, and all data rows. Use nrows:=100 for a small trial import. IncludeSourceRow:=True adds the original worksheet row number.

With baseline:=0, the earliest expenditure for each AccountID is compared with zero. Later values are compared with the preceding row for that AccountID. Calculations follow the table’s current row order; sort before using differences, shifts, or cumulative calculations.

to_excel_range writes values and, by default, formats based on column types. It does not clear old output outside the new result’s footprint. Clear the intended output area separately when replacing a larger prior result.

## Common Operations

These fragments assume t is a populated cTable; declare result tables as cTable and returned arrays as Variant.

```vb
' Filter and select columns; both return new tables.
Set result = t.Where("Expenditures", 100000, "gte")
Set result = result.SelectColumns(Array("RunDate", "AccountID", "Expenditures"))

' Membership filtering.
Set result = t.WhereIn("AccountID", Array("AccountID-A", "AccountID-B"))

' Distinct values and a scalar total.
fileList = t.GetUniqueValues("fPath")
Debug.Print t.sum("Expenditures")

' Create or replace a column with a constant value.
t.SetColumnValue "Source", "Snapshot import", ctString

' Grouped transforms return values aligned to the current rows.
netValues = t.GroupBy("AccountID").Col("Expenditures").diff(baseline:=0)
```

## API Guide

| Task | Entry points |
|---|---|
| Read Excel | `read_excel_sheet`, `read_excel_table`, `read_excel_range`, `read_excel_ws` |
| Read other data | `from_array`, `read_access_table`, `read_csv`, `read_json`, `read_text` |
| Inspect | `rowCount`, `ColumnCount`, `headers`, `shape`, `GetValue`, `columnNumber`, `ColumnExists`, `RequireColumns`, `info`, `ToString` |
| Filter and select | `Where`, `WhereNot`, `WhereIn`, `WhereNotIn`, `SelectColumns`, `head`, `tail`, `dropna`, `drop_duplicates` |
| Combine | `MakeJoin`, `merge`, `Append`, `concat`, `append_many`, `from_tables` |
| Sort | `SortBy`, `sort_values`, `nlargest`, `nsmallest` |
| Edit columns | `InsertColumn`, `RenameColumn`, `SetColumnValue`, `DropColumnIfExists`, `ReorderColumnInPlace`, `ReorderColumnRelativeInPlace` |
| Calculate | `AddComputedColumn`, `SetComputedColumn`, `diff`, `shift`, `rank`, `cumsum`, `rolling`, `expanding`, `GroupBy` |
| Summarize | `sum`, `mean`, `median`, `min`, `max`, `count`, `nunique`, `quantile`, `AggregateDict` |
| Reshape | `pivot`, `melt` |
| Types and missing data | `dtype`, `dtypes`, `convert_dtypes`, `astype`, `to_numeric`, `to_datetime`, `isna`, `notna`, `fillna` |
| Strings and dates | `str_trim`, `str_replace`, `str_contains`, `dt`, `dt_fy`, `dt_fq`, `dt_format` |
| Export | `to_excel_range`, `ReplaceExcelTableFrom_cTable`, `to_json` |

Use to_excel_range, rather than to_excel, for range export. Grouped transforms preserve row alignment; AggregateDict provides grouped aggregation as a dictionary.

## Return Values And Mutation

| Pattern | Behavior |
|---|---|
| `Set result = t.Where(...)` | Returns a new table. |
| `Set result = t.sort_values(...)` | Returns a sorted copy by default. |
| `Call t.sort_values(..., inplace:=True)` | Sorts t and returns Nothing; do not assign the return value back to t. |
| `t.SortBy ...` | Sorts t in place, with a common direction for its sort keys. |
| `t.Append other` | Adds rows to t, aligning columns by name. |
| `Set result = t.concat(other)` | Returns a combined table; default strict mode requires matching column sets. |
| `t.AddDiffColumn ...` | Adds a calculated column to t. |
| `values = t.diff(...)` | Returns an array rather than a table. |

Use Set for object results. Public row and column positions are 1-based. Many column arguments accept names or numeric positions, but this varies by method; check the declaration before supplying a number.

## Types And Missing Values

Supported column types are ctVariant, ctString, ctLong, ctDouble, ctCurrency, ctDate, and ctBoolean. Read methods normally infer types; review t.dtypes_string when identifiers, dates, or currency need particular treatment. Disable inference where necessary to preserve source representations such as leading-zero identifiers.

VBA Null is the universal missing-value marker. NA and NaT also return Null; blank or whitespace-only strings count as missing in the library’s default missing-value tests. Use IsNull or IsNAValue for scalar checks rather than value = Null.

The four Where* filters default to nullBehavior:="exclude". WhereNot and WhereNotIn therefore invert their positive counterparts over non-missing values; missing rows are excluded from both by default. Use nullBehavior:="include" when missing rows should be retained.

## Source Metadata And Cell Lineage

cDataSource describes an input or a declared source. Lineage records source attribution at cell level, allowing a combined table to retain sources from different files or join inputs.

- Enable lineage before reading: t.EnableLineage = True. It is disabled by default unless the project default has been changed.
- For future table instances, set cLineageRegistry.LineageDefaultEnabled = True before constructing them.
- Inspect a cell with t.LineageExplain(1, "Expenditures") or obtain a source object with Set ds = t.LineageOfName(1, "Expenditures").
- Inspect summaries with LineageReport, LineageColumnReport, and LineageSourceCount.
- LineageSources returns source-label strings; LineageSourceObjects returns source objects.
- With lineage active, t.DataSource returns a source only when the table resolves to a single source; otherwise it can return Nothing. Check object results before accessing their properties. Enabling lineage after loading cannot reconstruct provenance that was not recorded during the read.

Source attribution is not a full dependency graph of every input used in a calculation. Use explicit source stamps and write attribution where the application needs them. Avoid resetting the shared registries while tables using their identifiers remain alive.

## Snapshot Change History

cTableHistory compares repeated snapshots and returns change events. It is separate from cell lineage and is not a general log of every method call.

```vb
' t contains RunDate, AccountID, and Expenditures across multiple snapshots.
Dim changes As cTable
t.History.Init "RunDate"
Set changes = t.HistoryFor("AccountID", columns:=Array("Expenditures"))
Debug.Print changes.ToString
```

Use at most one non-missing entity key per snapshot. The first snapshot establishes the baseline. Subsequent snapshots can produce Changed, Added, and Removed events, including before/after values and snapshot dates. History queries build cached indexes without sorting or modifying the source table.

After mutating the table, reinitialize with t.History.Init "RunDate" before querying again. Likewise, obtain fresh cGroupBy handles after table mutations rather than retaining stale handles.

## Using cTable with AI Assistants

cTable's complete VBA source code and documentation are available in this repository, making it possible to use AI coding assistants such as ChatGPT, Claude, Gemini, and others to help develop, troubleshoot, and extend cTable-based Excel applications.

By providing an AI assistant with the relevant cTable class modules and documentation, developers can obtain assistance with:

- Writing VBA procedures that use cTable's data manipulation and ETL capabilities.
- Understanding method signatures, parameters, and return values.
- Debugging errors and unexpected results.
- Building more complex workflows involving filtering, joins, aggregation, and data lineage.
- Developing and reviewing enhancements to the library itself.

**For best results, provide the assistant with the relevant source files and API documentation**, and ask it to verify method names and signatures against the actual cTable implementation rather than assuming that equivalent pandas methods exist.

AI-generated code should always be reviewed, compiled, and tested in Excel VBA before use, particularly in applications involving financial or other sensitive data.

When using cloud-based AI services, follow your organization's policies regarding source code, proprietary information, and sensitive data. Do not share confidential workbooks or data with an external service unless authorized.

## Diagnostics And Maintenance

- Check t.Log.HasErrors and inspect t.Log.ErrorMessages; status messages are available through StatusMessages. Public methods can also raise VBA errors, so use normal error handling.
- For workbook metadata automation failures, consult excel-automation-error-runbook.md if supplied with the source distribution. Optional document-property failures should be distinguished from failures to read worksheet data.
- Keep the related classes together when updating. Preserve export attributes, public API compatibility, and third-party notices.
- Compile after import and exercise small fixtures covering filtering with missing values, joins, append alignment, types, lineage, and snapshot history before using a changed version in production.
The exported VBA modules in src/ define the public API and implementation. Examples are illustrative and have not been executed in an Excel VBA host unless explicitly stated otherwise.

## Documentation

- [Installation](docs/INSTALLATION.md)
- [Quick Start](docs/QUICKSTART.md)
- [Architecture](docs/ARCHITECTURE.md)
- [Data Lineage](docs/LINEAGE.md)
- [Changelog](CHANGELOG.md)
- [Contributing](CONTRIBUTING.md)
- [API Reference](docs/API-REFERENCE.md) — cTable methods,
  parameters, examples, and detailed reference guides.

## License

Copyright (c) 2026 Brian A. Crawford.

cTable is open-source software distributed under the
[MIT License](LICENSE).

You may use, modify, and distribute cTable in personal,
academic, government, and commercial applications,
including proprietary Excel VBA projects, subject to
the terms of the MIT License.

This repository includes the third-party VBA-JSON
component (`JsonConverter.bas`), which retains its
original copyright and licensing terms.

See [Third-Party Notices](THIRD-PARTY-NOTICES.md)
for details.
