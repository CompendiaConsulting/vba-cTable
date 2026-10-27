Attribute VB_Name = "Test_cTableRegression"
Option Explicit
' cTable regression suite
' Copyright (c) 2026 Brian A. Crawford
' SPDX-License-Identifier: MIT
'
' Import as a standard VBA module, then run RunAllTests.
' Fixtures are in memory; no worksheets or external files are needed.
' Failure raises an error after reporting results in the Immediate Window.
' Intended for the corrected exclude/include Null policy.

Private mPassed As Long
Private mFailed As Long
Private mChecks As Long

Public Sub RunAllTests()
    mPassed = 0: mFailed = 0: mChecks = 0
    Debug.Print String$(64, "-")
    Debug.Print "cTable regression tests: " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Debug.Print String$(64, "-")

    RunCase 1, "Import / accessors"
    RunCase 2, "Where / WhereIn"
    RunCase 3, "Inverse filters and Null policy"
    RunCase 4, "Sort copy and in-place"
    RunCase 5, "Append aligns headers"
    RunCase 6, "concat preserves inputs"
    RunCase 7, "Left merge and indicator"
    RunCase 8, "SelectColumns"
    RunCase 9, "Clone independent copy"
    RunCase 10, "drop_duplicates"
    RunCase 11, "Grouped difference baseline"
    RunCase 12, "Lineage basic smoke test"

    Debug.Print String$(64, "-")
    Debug.Print "SUMMARY: " & mPassed & " passed, " & mFailed & _
                " failed, " & mChecks & " assertions"
    If mFailed > 0 Then Err.Raise vbObjectError + 2800, _
        "Test_cTableRegression.RunAllTests", CStr(mFailed) & " regression tests failed."
End Sub

Private Sub RunCase(ByVal caseNum As Long, ByVal description As String)
    On Error GoTo CaseFailed
    Select Case caseNum
        Case 1: TestImport
        Case 2: TestFilters
        Case 3: TestMissingPolicy
        Case 4: TestSort
        Case 5: TestAppend
        Case 6: TestConcat
        Case 7: TestMerge
        Case 8: TestProjection
        Case 9: TestClone
        Case 10: TestDuplicates
        Case 11: TestDiff
        Case 12: TestLineage
    End Select
    mPassed = mPassed + 1
    Debug.Print "PASS | " & description
    Exit Sub
CaseFailed:
    mFailed = mFailed + 1
    Debug.Print "FAIL | " & description & " | " & Err.Number & " | " & Err.Description
    Err.Clear
End Sub

Private Sub Check(ByVal condition As Boolean, ByVal message As String)
    mChecks = mChecks + 1
    If Not condition Then Err.Raise vbObjectError + 2801, _
        "cTable regression", message
End Sub

Private Sub Eq(ByVal expected As Variant, ByVal actual As Variant, _
               ByVal message As String)
    mChecks = mChecks + 1
    If IsNull(expected) Then
        If Not IsNull(actual) Then Err.Raise vbObjectError + 2802, _
            "cTable regression", message & ": expected Null"
        Exit Sub
    End If
    If IsNull(actual) Then Err.Raise vbObjectError + 2802, _
        "cTable regression", message & ": unexpected Null"
    If CStr(expected) <> CStr(actual) Then Err.Raise vbObjectError + 2802, _
        "cTable regression", message & ": expected " & CStr(expected) & _
        ", got " & CStr(actual)
End Sub

Private Function Fixture(Optional ByVal withLineage As Boolean = False) As cTable
    Dim h(1 To 3) As String
    Dim a(1 To 4, 1 To 3) As Variant
    Dim t As New cTable
    h(1) = "RunDate": h(2) = "LOA": h(3) = "Expenditures"
    a(1, 1) = DateSerial(2026, 1, 1): a(1, 2) = "A": a(1, 3) = 100
    a(2, 1) = DateSerial(2026, 1, 1): a(2, 2) = "B": a(2, 3) = 50
    a(3, 1) = DateSerial(2026, 2, 1): a(3, 2) = "A": a(3, 3) = 170
    a(4, 1) = DateSerial(2026, 2, 1): a(4, 2) = "B": a(4, 3) = 90
    t.EnableLineage = withLineage
    t.from_array h, a, 4, 3
    Set Fixture = t
End Function

Private Function MissingFixture() As cTable
    Dim h(1 To 1) As String
    Dim a(1 To 4, 1 To 1) As Variant
    Dim t As New cTable
    h(1) = "LOA"
    a(1, 1) = "A": a(2, 1) = "B": a(3, 1) = Null: a(4, 1) = "A"
    t.from_array h, a, 4, 1
    Set MissingFixture = t
End Function

Private Sub TestImport()
    Dim t As cTable: Set t = Fixture()
    Eq 4, t.rowCount, "rowCount"
    Eq 3, t.ColumnCount, "ColumnCount"
    Eq "LOA", t.columnName(2), "header"
    Eq "A", t.GetValue(1, "LOA"), "GetValue by name"
    Eq 170, t.GetValue(3, 3), "GetValue by number"
End Sub

Private Sub TestFilters()
    Dim t As cTable, r As cTable
    Set t = Fixture()
    Set r = t.Where("LOA", "A"): Eq 2, r.rowCount, "Where equality"
    Set r = t.Where(3, 100, "gte"): Eq 2, r.rowCount, "Where >= by index"
    Set r = t.WhereIn("LOA", Array("B")): Eq 2, r.rowCount, "WhereIn"
    Eq 4, t.rowCount, "filters leave source unchanged"
End Sub

Private Sub TestMissingPolicy()
    Dim t As cTable, r As cTable
    Set t = MissingFixture()
    Set r = t.WhereNot("LOA", "A")
    Eq 1, r.rowCount, "WhereNot excludes missing by default"
    Set r = t.WhereNotIn("LOA", Array("A"))
    Eq 1, r.rowCount, "WhereNotIn excludes missing by default"
    Set r = t.Where("LOA", "A", nullBehavior:="include")
    Eq 3, r.rowCount, "Where include missing"
    Set r = t.WhereNot("LOA", "A", nullBehavior:="include")
    Eq 2, r.rowCount, "WhereNot include missing"
    Set r = t.Where("LOA", Null): Eq 1, r.rowCount, "explicit Null search"
    Set r = t.WhereNot("LOA", Null): Eq 3, r.rowCount, "explicit inverse Null"
End Sub

Private Sub TestSort()
    Dim t As cTable, r As cTable, ignored As cTable
    Set t = Fixture()
    Set r = t.sort_values(by:=Array("RunDate", "LOA"), _
                          ascending:=Array(False, True))
    Eq 170, r.GetValue(1, "Expenditures"), "RunDate descending; LOA ascending"
    Eq 100, t.GetValue(1, "Expenditures"), "copy sort preserves original"
    Set ignored = t.sort_values(by:="Expenditures", _
                                 ascending:=False, inplace:=True)
    Check (ignored Is Nothing), "inplace sort returns Nothing"
    Eq 170, t.GetValue(1, "Expenditures"), "inplace sort modifies original"
End Sub

Private Sub TestAppend()
    Dim t As cTable, extra As New cTable
    Dim h(1 To 3) As String, a(1 To 1, 1 To 3) As Variant
    Set t = Fixture()
    h(1) = "Expenditures": h(2) = "LOA": h(3) = "RunDate"
    a(1, 1) = 250: a(1, 2) = "A": a(1, 3) = DateSerial(2026, 3, 1)
    extra.from_array h, a, 1, 3
    t.Append extra, strict:=True
    Eq 5, t.rowCount, "append adds one row"
    Eq 250, t.GetValue(5, "Expenditures"), "append maps by header"
    Eq "A", t.GetValue(5, "LOA"), "append LOA mapping"
End Sub

Private Sub TestConcat()
    Dim t As cTable, other As cTable, result As cTable
    Set t = Fixture(): Set other = Fixture()
    Set result = t.concat(other)
    Eq 8, result.rowCount, "concat result rows"
    Eq 4, t.rowCount, "left source unchanged"
    Eq 4, other.rowCount, "right source unchanged"
End Sub

Private Sub TestMerge()
    Dim t As cTable, other As New cTable, result As cTable
    Dim h(1 To 2) As String, a(1 To 2, 1 To 2) As Variant
    Set t = Fixture()
    h(1) = "LOA": h(2) = "Directorate"
    a(1, 1) = "A": a(1, 2) = "East"
    a(2, 1) = "C": a(2, 2) = "West"
    other.from_array h, a, 2, 2
    Set result = t.merge(other, how:="left", on_arg:="LOA", indicator:=True)
    Eq 4, result.rowCount, "left merge row count"
    Eq "East", result.GetValue(1, "Directorate"), "matched right data"
    Check IsNull(result.GetValue(2, "Directorate")), "unmatched right is Null"
    Eq "left_only", result.GetValue(2, "_merge"), "left-only indicator"
    Eq "both", result.GetValue(1, "_merge"), "matched indicator"
End Sub

Private Sub TestProjection()
    Dim t As cTable, r As cTable
    Set t = Fixture()
    Set r = t.SelectColumns(Array("LOA", "Expenditures"))
    Eq 2, r.ColumnCount, "projected column count"
    Eq "LOA", r.columnName(1), "projected order"
    Eq 3, t.ColumnCount, "projection leaves source unchanged"
End Sub

Private Sub TestClone()
    Dim t As cTable, c As cTable
    Set t = Fixture(): Set c = t.Clone()
    c.SetValue 1, "LOA", "Z"
    Eq "Z", c.GetValue(1, "LOA"), "cloned table changed"
    Eq "A", t.GetValue(1, "LOA"), "original table unaffected"
End Sub

Private Sub TestDuplicates()
    Dim t As cTable, r As cTable
    Set t = Fixture()
    Set r = t.drop_duplicates(subset:="LOA")
    Eq 2, r.rowCount, "unique LOA rows"
    Eq "A", r.GetValue(1, "LOA"), "first A kept"
    Eq "B", r.GetValue(2, "LOA"), "first B kept"
End Sub

Private Sub TestDiff()
    Dim t As cTable: Set t = Fixture()
    t.AddDiffColumn colName:="Expenditures", by:="LOA", _
                    newColName:="Net", baseline:=0
    Eq 100, t.GetValue(1, "Net"), "A baseline zero"
    Eq 50, t.GetValue(2, "Net"), "B baseline zero"
    Eq 70, t.GetValue(3, "Net"), "A change"
    Eq 40, t.GetValue(4, "Net"), "B change"
End Sub

Private Sub TestLineage()
    Dim t As cTable, src As cDataSource
    Set t = Fixture(True)
    Check t.EnableLineage, "lineage enabled before import"
    Set src = t.LineageOfName(1, "Expenditures")
    Check Not (src Is Nothing), "source object available"
    Check Len(t.LineageExplain(1, "Expenditures")) > 0, "lineage description"
End Sub
