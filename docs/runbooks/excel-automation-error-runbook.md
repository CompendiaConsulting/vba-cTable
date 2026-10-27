# Runbook: Excel Automation Error During cTable Source Metadata Initialization

## Purpose

This runbook describes how to diagnose and resolve an Excel Automation error that occurs when `cTable.read_excel_sheet` initializes a `cDataSource` object for an otherwise readable workbook.

The issue commonly appears as:

```text
Run-time error '-2147467259 (80004005)':
Automation error
Unspecified error
```

---

## Scope

This runbook applies to workbook reads performed through:

```vba
t.read_excel_sheet filePath, sheetName, hasHeaders:=False
```

It specifically addresses failures during `cDataSource.InitFromWorkbook`, not failures opening a workbook, locating a worksheet, or reading worksheet values.

---

## Symptoms

Typical symptoms include:

* `read_excel_sheet` raises Automation error `80004005`.
* The workbook opens successfully in Excel.
* The target worksheet exists.
* `UsedRange` can be inspected successfully.
* Worksheet cells can be read manually in the Immediate window or a test procedure.
* Ignoring the error allows the remainder of the table-load procedure to complete.

### Example Validation

```vba
Dim wb As Workbook
Dim ws As Worksheet
Dim rng As Range
Dim samplePath As String
samplePath = "C:\Data\sample.xlsx"

Set wb = Workbooks.Open(samplePath, ReadOnly:=True, UpdateLinks:=False)
Set ws = wb.Worksheets("Sheet1")
Set rng = ws.UsedRange

Debug.Print rng.Address
Debug.Print rng.Rows.Count
Debug.Print rng.Columns.Count
Debug.Print rng.Cells(1, 1).Value
```

### Example Successful Output

```text
$A$1:$K$108
108
11
Sample Report
```

---

## Root Cause

`cDataSource.InitFromWorkbook` retrieves optional Excel built-in document properties:

```vba
m_Author = wb.BuiltinDocumentProperties("Author")
m_LastSaveTime = wb.BuiltinDocumentProperties("Last Save Time")
m_LastAuthor = wb.BuiltinDocumentProperties("Last Author")
```

For some valid workbooks, Excel raises an Automation error when one or more built-in document properties are accessed. This may occur with externally generated, minimally populated, protected, cloud-synchronized, or otherwise nonstandard `.xlsx` files.

The workbook can still be valid and fully readable. The error is related to optional workbook metadata, not worksheet content.

---

## Impact

### Before the Fix

* A failure retrieving optional metadata prevented `read_excel_sheet` from loading valid worksheet data.
* The failure occurred before `p_LoadFromRange` was reached.
* Valid files could not be imported despite Excel being able to read the target worksheet and range.

### After the Fix

* `cDataSource` remains initialized with required source information:
  * File path
  * Sheet name
  * Table name, if applicable
  * Source type
  * Source kind
  * Import date
  * Import user
* Any unavailable optional metadata remains at its default value:
  * `Author = ""`
  * `LastAuthor = ""`
  * `LastSaveTime = 0`
* `cTable.read_excel_sheet` proceeds normally.

---

## Resolution

Treat Excel built-in document properties as optional metadata. Use safe helper functions when retrieving document properties, rather than assigning document-property values directly.

### 1. Safe Text-Property Helper

Add this private helper to `cDataSource`:

```vba
Private Function p_BuiltInDocumentPropertyText( _
        ByVal wb As Workbook, _
        ByVal propertyName As String) As String

    Dim propertyValue As Variant

    On Error Resume Next
    propertyValue = wb.BuiltinDocumentProperties(propertyName).Value

    If Err.Number = 0 Then
        If Not IsNull(propertyValue) _
           And Not IsEmpty(propertyValue) _
           And Not IsError(propertyValue) Then

            p_BuiltInDocumentPropertyText = CStr(propertyValue)
        End If
    End If

    Err.Clear
    On Error GoTo 0
End Function
```

### 2. Safe Date-Property Helper

Add this private helper to `cDataSource`:

```vba
Private Function p_BuiltInDocumentPropertyDate( _
        ByVal wb As Workbook, _
        ByVal propertyName As String) As Date

    Dim propertyValue As Variant

    On Error Resume Next
    propertyValue = wb.BuiltinDocumentProperties(propertyName).Value

    If Err.Number = 0 Then
        If Not IsNull(propertyValue) _
           And Not IsEmpty(propertyValue) _
           And Not IsError(propertyValue) Then

            If IsDate(propertyValue) Then
                p_BuiltInDocumentPropertyDate = CDate(propertyValue)
            End If
        End If
    End If

    Err.Clear
    On Error GoTo 0
End Function
```

### 3. Updated `InitFromWorkbook`

Replace direct calls to `BuiltinDocumentProperties` in `InitFromWorkbook` with the safe helpers:

```vba
Public Sub InitFromWorkbook(ByVal wb As Workbook, _
                            ByVal filePath As String, _
                            Optional ByVal sheetName As String = "", _
                            Optional ByVal tableName As String = "")

    p_GuardFrozen

    m_FilePath = filePath
    m_SheetName = sheetName
    m_TableName = tableName
    m_SourceType = "ExcelSheet"

    If Len(tableName) > 0 Then
        m_Kind = ctKindExcelTable
    Else
        m_Kind = ctKindExcelSheet
    End If

    If Not wb Is Nothing Then
        m_Author = p_BuiltInDocumentPropertyText(wb, "Author")

        m_LastSaveTime = p_BuiltInDocumentPropertyDate( _
                            wb, "Last Save Time")

        m_LastAuthor = p_BuiltInDocumentPropertyText( _
                           wb, "Last Author")
    End If

    p_Stamp
End Sub
```

### 4. Updated `InitFromPath`

Apply the same safe helpers in `InitFromPath`, which accesses the same built-in document properties.

**Replace this vulnerable block:**

```vba
If Not wb Is Nothing Then
    m_Author = wb.BuiltinDocumentProperties("Author")
    m_LastSaveTime = wb.BuiltinDocumentProperties("Last Save Time")
    m_LastAuthor = wb.BuiltinDocumentProperties("Last Author")

    If didOpen Then wb.Close False
End If
```

**With:**

```vba
If Not wb Is Nothing Then
    m_Author = p_BuiltInDocumentPropertyText(wb, "Author")

    m_LastSaveTime = p_BuiltInDocumentPropertyDate( _
                        wb, "Last Save Time")

    m_LastAuthor = p_BuiltInDocumentPropertyText( _
                       wb, "Last Author")

    If didOpen Then wb.Close SaveChanges:=False
End If
```

---

## Validation Steps

After applying the fix, test the affected workbook:

```vba
Dim t As New cTable
Dim samplePath As String
samplePath = "C:\Data\sample.xlsx"

t.read_excel_sheet samplePath, "Sheet1", hasHeaders:=False

Debug.Print t.rowCount
Debug.Print t.ColumnCount
Debug.Print t.GetValue(1, 1)
```

### Expected Outcome

* No Automation error is raised.
* The row and column counts match the selected worksheet range.
* The first table value matches the worksheet.
* `DataSource` is initialized even if optional document properties are unavailable.

### Illustrative Expected Output”

```text
108
11
Sample Report
```

### Metadata Inspection (Optional)

```vba
Debug.Print t.DataSource.Author
Debug.Print t.DataSource.LastAuthor
Debug.Print t.DataSource.LastSaveTime
Debug.Print t.DataSource.Describe
```

*Note: Blank author, last-author, or save-time fields are acceptable when Excel cannot provide that optional metadata.*

---

## Troubleshooting

### 1. The Error Still Occurs

* Use VBA Debug mode to identify the exact failing line.
* Test workbook and worksheet access separately:

```vba
Dim wb As Workbook
Dim ws As Worksheet
Dim rng As Range

Set wb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
Set ws = wb.Worksheets(sheetName)
Set rng = ws.UsedRange

Debug.Print rng.Address
Debug.Print rng.Cells(1, 1).Value
```

If this test fails, the issue is not limited to `cDataSource` metadata retrieval.

### 2. The Error Occurs While Opening the Workbook

Investigate:

* Invalid or inaccessible path
* File permissions
* File locking
* OneDrive or SharePoint synchronization state
* Corrupted workbook
* Excel Protected View or trust settings
* External links or workbook prompts

*Caution: Do not broadly suppress this error in `read_excel_sheet`.*

### 3. The Error Occurs During `p_LoadFromRange`

Investigate:

* Merged cells
* Invalid range references
* Worksheet protection
* Cell error values
* Large or unusual `UsedRange`
* Excel COM instability

*Note: The document-property fix does not address failures in range reading or table loading.*

### 4. The Error Occurs in Lineage Processing

Test the source object and lineage registry separately:

```vba
Dim ds As cDataSource
Dim sourceId As Long

Set ds = New cDataSource
ds.InitFromWorkbook wb, filePath, sheetName

sourceId = cLineageRegistry.SourceId(ds)

Debug.Print sourceId
Debug.Print ds.Fingerprint
Debug.Print ds.Label
```

If `InitFromWorkbook` completes but `SourceId` fails, investigate `cLineageRegistry`, source interning, and lineage initialization rather than workbook document properties.

---

## Design Guidance

* Workbook document properties are optional enrichment metadata.
* A valid worksheet import should not fail because author or save-time metadata cannot be retrieved.
* Use narrowly scoped error handling around individual optional metadata reads.
* Do not place a broad `On Error Resume Next` around the entire source initialization or workbook-read procedure.
* Preserve normal error behavior for required operations:
  * Opening the workbook
  * Locating the worksheet
  * Identifying the range
  * Reading worksheet values
  * Loading values into `cTable`
* Continue to use `LastSaveTime` in source fingerprints when it is available. An unavailable value is represented by the existing zero-date behavior.

---

## Related Components

* `cTable.read_excel_sheet`
* `cDataSource.InitFromWorkbook`
* `cDataSource.InitFromPath`
* `cLineageRegistry`
* `cLineage`
* `p_LoadFromRange`

---

## Change History

| Date | Change |
| :--- | :--- |
| 2026-10-16 | Added safe retrieval of Excel built-in document properties to prevent Automation error `80004005` from blocking valid workbook imports. |
