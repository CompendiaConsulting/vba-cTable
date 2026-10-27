VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cGroupBy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' cGroupBy - Grouped data operations and transformations
'
' Copyright (c) 2026 Brian A. Crawford
' SPDX-License-Identifier: MIT
'
' Licensed under the MIT License.
' See LICENSE in the repository root.
' ============================================================
' Version 20260803b
' cGroupBy  -  a lightweight grouping handle over a cTable
'
' REQUIRES: cTable with Patch 1 (grouped differencing) AND Patch 2 (rank and
'           cumulative operations) installed.
'
' This REPLACES the earlier cGroupBy in its entirety. Select all in the
' existing class module and paste over it.
'
' PURPOSE
'   Chained, pandas-shaped syntax over the transforms that live in cTable:
'
'       df.groupby("AccountID")["Expenditure"].diff()
'       t.GroupBy("AccountID").Col("Expenditure").diff()
'
'   This class holds NO group membership of its own. It stores the key names
'   and delegates every call back to cTable, so there is exactly one
'   implementation of the grouping logic and nothing to keep in sync.
'
' STALENESS
'   The one real hazard in a grouping object is caching row numbers that a
'   later sort, filter, or append invalidates. cTable stamps a version counter
'   in p_ClearIndexes - the chokepoint every structural mutation already passes
'   through - and this class captures it at construction and re-checks on every
'   call. A grouping used after the table moved raises rather than returning
'   plausible wrong numbers.
'
'   The guard is deliberately over-eager: writes that only change values, such
'   as SetValue or FillNull, also bump the counter and will invalidate a
'   grouping that would technically still have been valid. Failing closed is
'   the right trade; call GroupBy again.
'
' ORDER
'   Nothing here sorts. For diff, shift and the cumulative operations the sort
'   order determines the RESULT - and for diff, its SIGN. rank is the one
'   exception: it reads values rather than positions, so only method:="first"
'   and tie order depend on where rows sit.
'
' OPTIONAL DEFAULT MEMBER
'   To write t.GroupBy("AccountID")("Expenditure").diff() instead of .Col(...),
'   export this class, add the attribute line immediately after the Col
'   signature, and re-import:
'
'       Public Function Col(ByVal colName As String) As cGroupBy
'       Attribute Col.VB_UserMemId = 0
'
'   The VBE will not let you type that line directly; it survives only through
'   an export/import cycle.

Private m_Tbl As cTable
Private m_By As Variant             ' key column name, or array of names
Private m_Version As Long           ' cTable structural version at construction
Private m_Col As String             ' selected value column; "" until Col()

'==============================================================================
' CONSTRUCTION
'==============================================================================

' Called by cTable.GroupBy. Not intended for direct use.
Public Sub x_Init(ByVal t As cTable, ByVal by As Variant, _
                  Optional ByVal colName As String = "")
    Set m_Tbl = t
    m_By = by
    m_Version = t.x_Version
    m_Col = colName
End Sub

'==============================================================================
' SELECTION
'==============================================================================

' Select the value column, returning a new handle. The original is unchanged,
' so one grouping can feed several columns:
'
'   Dim g As cGroupBy
'   Set g = t.GroupBy("AccountID")
'   obligated = g.Col("Obligation").cumsum()
'   expended  = g.Col("Expenditure").cumsum()
Public Function Col(ByVal colName As String) As cGroupBy
    p_Check

    If m_Tbl.columnNumber(colName) = 0 Then
        Err.Raise vbObjectError + 1323, "cGroupBy.Col", _
                  "Column not found: """ & colName & """"
    End If

    Dim g As cGroupBy
    Set g = New cGroupBy
    g.x_Init m_Tbl, m_By, colName

    Set Col = g
End Function

'==============================================================================
' TRANSFORMS  (one value per row, in current row order)
'
' Each returns a 1-based Variant array of RowCount elements. Rows whose group
' key is missing return Null.
'==============================================================================

' pandas groupby(K)[V].diff(periods)
' pandas groupby(K)[V].diff(periods), plus a baseline for unpaired rows.
'
'   t.GroupBy("AccountID").Col("Obligations").diff(baseline:=0)
Public Function diff(Optional ByVal periods As Long = 1, _
                     Optional ByVal baseline As Variant) As Variant
    p_Check
    diff = m_Tbl.diff(p_ValueCol("diff"), m_By, periods, True, baseline)
End Function

' pandas groupby(K)[V].shift(periods, fill_value=...)
Public Function shift(Optional ByVal periods As Long = 1, _
                      Optional ByVal fill_value As Variant) As Variant
    p_Check
    shift = m_Tbl.shift(p_ValueCol("shift"), m_By, periods, True, fill_value)
End Function


' pandas groupby(K)[V].rank(method, ascending, na_option, pct)
Public Function rank(Optional ByVal method As String = "average", _
                     Optional ByVal ascending As Boolean = True, _
                     Optional ByVal na_option As String = "keep", _
                     Optional ByVal pct As Boolean = False) As Variant
    p_Check
    rank = m_Tbl.rank(p_ValueCol("rank"), m_By, method, ascending, _
                      na_option, pct)
End Function


' pandas groupby(K)[V].cumsum(skipna)
Public Function cumsum(Optional ByVal skipna As Boolean = True) As Variant
    p_Check
    cumsum = m_Tbl.cumsum(p_ValueCol("cumsum"), m_By, skipna)
End Function


' pandas groupby(K)[V].cumprod(skipna)
Public Function cumprod(Optional ByVal skipna As Boolean = True) As Variant
    p_Check
    cumprod = m_Tbl.cumprod(p_ValueCol("cumprod"), m_By, skipna)
End Function


' pandas groupby(K)[V].cummax(skipna)
Public Function cummax(Optional ByVal skipna As Boolean = True) As Variant
    p_Check
    cummax = m_Tbl.cummax(p_ValueCol("cummax"), m_By, skipna)
End Function


' pandas groupby(K)[V].cummin(skipna)
Public Function cummin(Optional ByVal skipna As Boolean = True) As Variant
    p_Check
    cummin = m_Tbl.cummin(p_ValueCol("cummin"), m_By, skipna)
End Function

'==============================================================================
' IN-PLACE ASSIGNMENT
'
' The closest available equivalent to the pandas assignment idiom:
'
'   df["Expenditure_Diff"] = df.groupby("AccountID")["Expenditure"].diff()
'   t.GroupBy("AccountID").Col("Expenditure").AssignDiff "Expenditure_Diff"
'
' Each creates or replaces the target column, and each INVALIDATES the handle,
' because adding a column bumps the table's version. That is deliberate: a
' handle must not survive a change to the table it points at. Re-select from
' the table for the next assignment.
'
' Passing "" for newColName uses the default suffix for that operation.
'==============================================================================

Public Sub AssignDiff(Optional ByVal newColName As String = "", _
                      Optional ByVal periods As Long = 1)
    p_Check
    m_Tbl.SetDiffColumn p_ValueCol("AssignDiff"), m_By, periods, newColName
End Sub


Public Sub AssignShift(Optional ByVal newColName As String = "", _
                       Optional ByVal periods As Long = 1, _
                       Optional ByVal fill_value As Variant)
    p_Check
    m_Tbl.SetShiftColumn p_ValueCol("AssignShift"), m_By, periods, _
                         newColName, True, fill_value
End Sub


Public Sub AssignRank(Optional ByVal newColName As String = "", _
                      Optional ByVal method As String = "average", _
                      Optional ByVal ascending As Boolean = True, _
                      Optional ByVal na_option As String = "keep", _
                      Optional ByVal pct As Boolean = False)
    p_Check
    m_Tbl.SetRankColumn p_ValueCol("AssignRank"), m_By, method, ascending, _
                        na_option, pct, newColName
End Sub


Public Sub AssignCumSum(Optional ByVal newColName As String = "", _
                        Optional ByVal skipna As Boolean = True)
    p_Check
    m_Tbl.SetCumSumColumn p_ValueCol("AssignCumSum"), m_By, skipna, newColName
End Sub


Public Sub AssignCumProd(Optional ByVal newColName As String = "", _
                         Optional ByVal skipna As Boolean = True)
    p_Check
    m_Tbl.SetCumProdColumn p_ValueCol("AssignCumProd"), m_By, skipna, newColName
End Sub


Public Sub AssignCumMax(Optional ByVal newColName As String = "", _
                        Optional ByVal skipna As Boolean = True)
    p_Check
    m_Tbl.SetCumMaxColumn p_ValueCol("AssignCumMax"), m_By, skipna, newColName
End Sub


Public Sub AssignCumMin(Optional ByVal newColName As String = "", _
                        Optional ByVal skipna As Boolean = True)
    p_Check
    m_Tbl.SetCumMinColumn p_ValueCol("AssignCumMin"), m_By, skipna, newColName
End Sub

'==============================================================================
' INTROSPECTION
'==============================================================================

' The key columns backing this grouping, as a 0-based array of names.
Public Property Get keys() As Variant
    If IsArray(m_By) Then
        keys = m_By
    Else
        keys = Array(CStr(m_By))
    End If
End Property


' The selected value column, or "" if none has been chosen yet.
Public Property Get valueColumn() As String
    valueColumn = m_Col
End Property


' True while the underlying table still matches the state this handle was
' built against. Lets callers test without provoking an error.
Public Property Get IsValid() As Boolean
    If m_Tbl Is Nothing Then Exit Property
    IsValid = (m_Tbl.x_Version = m_Version)
End Property

'==============================================================================
' PRIVATE
'==============================================================================

Private Sub p_Check()
    If m_Tbl Is Nothing Then
        Err.Raise vbObjectError + 1320, "cGroupBy", _
                  "This grouping was never initialised. Use cTable.GroupBy."
    End If

    If m_Tbl.x_Version <> m_Version Then
        Err.Raise vbObjectError + 1321, "cGroupBy", _
                  "The underlying cTable changed after this grouping was " & _
                  "created, so its row positions are no longer valid. " & _
                  "Call GroupBy again."
    End If
End Sub


Private Function p_ValueCol(ByVal callerName As String) As String
    If Len(m_Col) = 0 Then
        Err.Raise vbObjectError + 1322, "cGroupBy." & callerName, _
                  "Select a value column first, e.g. " & _
                  "GroupBy(""AccountID"").Col(""Expenditure"")."
    End If

    p_ValueCol = m_Col
End Function

'========================


' pandas groupby(K)[V].rolling(window, ...).agg
Public Function rolling(ByVal window As Variant, _
                        Optional ByVal operation As String = "sum", _
                        Optional ByVal on_arg As String = "", _
                        Optional ByVal min_periods As Long = -1, _
                        Optional ByVal closed As String = "right") As Variant
    p_Check
    rolling = m_Tbl.rolling(p_ValueCol("rolling"), window, m_By, operation, _
                            on_arg, min_periods, closed)
End Function


' pandas groupby(K)[V].expanding(...).agg
Public Function expanding(Optional ByVal operation As String = "sum", _
                          Optional ByVal min_periods As Long = 1) As Variant
    p_Check
    expanding = m_Tbl.expanding(p_ValueCol("expanding"), m_By, operation, _
                                min_periods)
End Function


' pandas Series.dt, reached through a grouping for symmetry with the other
' accessors. The grouping itself is irrelevant to the result - .dt is a
' per-cell transform - but routing it here keeps one chained idiom.
Public Function dt(ByVal part As String, _
                   Optional ByVal fiscalStartMonth As Long = 10) As Variant
    p_Check
    dt = m_Tbl.dt(p_ValueCol("dt"), part, fiscalStartMonth)
End Function
