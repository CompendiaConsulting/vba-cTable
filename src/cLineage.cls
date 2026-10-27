VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cLineage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' cLineage - Cell-level data provenance and source tracking
'
' Copyright (c) 2026 Brian A. Crawford
' SPDX-License-Identifier: MIT
'
' Licensed under the MIT License.
' See LICENSE in the repository root.
' ============================================================
'
' Version 20260727
'==============================================================================
' cLineage  -  Per-table cell-lineage engine for cTable
'
' REQUIRES: cLineageRegistry (predeclared singleton class), cDataSource,
'           cLongVec, Microsoft Scripting Runtime
'
' MODEL
'   Lineage of a cell is resolved most-specific-first through four tiers:
'
'     1. Cell override   sparse Dictionary or dense cLongVec, per column
'     2. Column stamp    one source ID for an entire column
'     3. Row block map   each row points at an interned column map
'     4. Table default   a single interned column map for all rows
'
'   Tier 3 collapses to tier 4 whenever every row agrees, which is the case
'   for every table that has not been through a concat/union. Tiers 1 and 2
'   are not allocated at all until something is overwritten.
'
' COST
'   Uniform table          O(ColCount) longs, total
'   After a join           O(ColCount) longs, total (columns differ, rows do not)
'   After a concat         O(bands) while row order is contiguous,
'                          4 bytes/row once sorting or filtering interleaves it
'   Overwritten cells      ~40 bytes/cell sparse, promoting to 4 bytes/row
'                          for that one column past the density threshold
'
'   A 2,000,000 x 30 table holds roughly 960 MB of Variants. A dense row map
'   for it is 8 MB - under 1%. A per-cell map would be 240 MB. That asymmetry
'   is the whole reason this class exists.
'
' INVARIANTS
'   * Interned maps are never mutated. Structural changes intern new maps.
'   * Any path that cannot compute lineage exactly assigns ctSrcUnknown.
'     Lineage degrades; it never lies.
'   * Row and column transforms funnel through exactly two routines,
'     x_ApplyRowGather and x_ApplyColumnMapping.
'==============================================================================

Private Const OVERLAY_DENSE_MIN As Long = 4096   ' promote sparse->dense above
Private Const OVERLAY_DENSE_DIV As Long = 16     ' ...and above RowCount/16
Private Const MAX_BANDS As Long = 64             ' beyond this, store dense

Private Enum ctRowStore
    ctRowUniform = 0
    ctRowBanded = 1
    ctRowDense = 2
End Enum

Private m_RowCount As Long
Private m_ColCount As Long

' --- tiers 3 and 4: row -> column-map assignment -----------------------------
Private m_RowStore  As Long          ' ctRowStore; Long so it can pass ByRef
Private m_UniMap    As Long          ' ctRowUniform
Private m_BandStart() As Long        ' ctRowBanded, 1..m_BandCount, ascending
Private m_BandMap()   As Long
Private m_BandCount   As Long
Private m_RowMap()    As Long        ' ctRowDense, 1..m_RowCount

' --- tiers 1 and 2: per-column overlays --------------------------------------
Private m_ColMode()   As Byte        ' 0 none, 1 stamp, 2 sparse, 3 dense
Private m_ColStamp()  As Long
Private m_ColSparse() As Object      ' Dictionary(rowNum -> srcId)
Private m_ColDense()  As cLongVec
Private m_AnyOverlay  As Boolean

'==============================================================================
' CONSTRUCTION
'==============================================================================

Public Sub InitUniform(ByVal rowCount As Long, ByVal colCount As Long, _
                       ByVal srcId As Long)
    m_RowCount = rowCount
    m_ColCount = colCount
    m_RowStore = ctRowUniform
    m_UniMap = cLineageRegistry.UniformMapId(colCount, srcId)
    p_ClearOverlays
End Sub

' Internal: prepare an empty engine with a known shape.
Public Sub x_InitShell(ByVal colCount As Long, ByVal rowCount As Long)
    m_ColCount = colCount
    m_RowCount = rowCount
    m_RowStore = ctRowUniform
    m_UniMap = cLineageRegistry.UniformMapId(colCount, ctSrcUnknown)
    p_ClearOverlays
End Sub

Private Sub p_ClearOverlays()
    m_AnyOverlay = False
    Erase m_ColMode
    Erase m_ColStamp
    Erase m_ColSparse
    Erase m_ColDense
End Sub

Private Sub p_EnsureOverlayArrays()
    If m_AnyOverlay Then Exit Sub
    If m_ColCount <= 0 Then Exit Sub
    ReDim m_ColMode(1 To m_ColCount)
    ReDim m_ColStamp(1 To m_ColCount)
    ReDim m_ColSparse(1 To m_ColCount)
    ReDim m_ColDense(1 To m_ColCount)
    m_AnyOverlay = True
End Sub

Public Function CloneMe() As cLineage
    Dim out As cLineage
    Set out = New cLineage
    out.x_InitShell m_ColCount, m_RowCount
    out.x_CopyStateFrom Me
    Set CloneMe = out
End Function

' Internal: deep-copy row structure and overlays from another engine.
Public Sub x_CopyStateFrom(ByVal src As cLineage)
    m_RowCount = src.rowCount
    m_ColCount = src.ColumnCount
    src.x_ExportRowStore m_RowStore, m_UniMap, m_BandStart, m_BandMap, _
                         m_BandCount, m_RowMap
    p_ClearOverlays
    If Not src.HasOverlays Then Exit Sub

    p_EnsureOverlayArrays
    Dim c As Long
    For c = 1 To m_ColCount
        Select Case src.x_ColMode(c)
            Case 1
                m_ColMode(c) = 1
                m_ColStamp(c) = src.x_ColStamp(c)
            Case 2
                m_ColMode(c) = 2
                Set m_ColSparse(c) = p_CloneDict(src.x_ColSparse(c))
            Case 3
                m_ColMode(c) = 3
                Set m_ColDense(c) = src.x_ColDense(c).CloneMe
        End Select
    Next c
End Sub

Private Function p_CloneDict(ByVal d As Object) As Object
    Dim out As Object
    Set out = CreateObject("Scripting.Dictionary")
    Dim k As Variant
    For Each k In d.keys
        out.Add k, d(k)
    Next k
    Set p_CloneDict = out
End Function

'==============================================================================
' INTERNAL ACCESSORS  (project-scope; not part of the cTable public API)
'==============================================================================

Public Sub x_ExportRowStore(ByRef store As Long, ByRef uniMap As Long, _
                            ByRef bandStart() As Long, ByRef bandMap() As Long, _
                            ByRef bandCount As Long, ByRef rowMap() As Long)
    store = m_RowStore
    uniMap = m_UniMap
    bandCount = m_BandCount
    If m_BandCount > 0 Then
        bandStart = m_BandStart
        bandMap = m_BandMap
    Else
        Erase bandStart
        Erase bandMap
    End If
    If m_RowStore = ctRowDense Then rowMap = m_RowMap Else Erase rowMap
End Sub

Public Property Get HasOverlays() As Boolean
    HasOverlays = m_AnyOverlay
End Property

Public Function x_ColMode(ByVal c As Long) As Byte
    If m_AnyOverlay Then x_ColMode = m_ColMode(c)
End Function

Public Function x_ColStamp(ByVal c As Long) As Long
    If m_AnyOverlay Then x_ColStamp = m_ColStamp(c)
End Function

Public Function x_ColSparse(ByVal c As Long) As Object
    If m_AnyOverlay Then Set x_ColSparse = m_ColSparse(c)
End Function

Public Function x_ColDense(ByVal c As Long) As cLongVec
    If m_AnyOverlay Then Set x_ColDense = m_ColDense(c)
End Function

'==============================================================================
' SHAPE
'==============================================================================

Public Property Get rowCount() As Long
    rowCount = m_RowCount
End Property

Public Property Get ColumnCount() As Long
    ColumnCount = m_ColCount
End Property

Public Property Get Mode() As ctLineageMode
    If m_AnyOverlay Then
        Dim c As Long
        For c = 1 To m_ColCount
            If m_ColMode(c) >= 2 Then
                Mode = ctLinFull
                Exit Property
            End If
        Next c
    End If

    If m_RowStore <> ctRowUniform Then
        Mode = ctLinByRow
        Exit Property
    End If

    Dim uniSrc As Long
    uniSrc = cLineageRegistry.MapUniformSource(m_UniMap)
    If uniSrc = 0 Then
        Mode = ctLinByColumn
    ElseIf m_AnyOverlay Then
        Mode = ctLinByColumn
    Else
        Mode = ctLinUniform
    End If
End Property

Public Property Get HasMixedRowSources() As Boolean
    HasMixedRowSources = (m_RowStore <> ctRowUniform)
End Property

Public Property Get HasMixedColumnSources() As Boolean
    HasMixedColumnSources = (Mode = ctLinByColumn Or Mode = ctLinByRow _
                             Or Mode = ctLinFull)
End Property

'==============================================================================
' RESOLUTION  (the hot path)
'==============================================================================

Public Function RowMapAt(ByVal r As Long) As Long
    Select Case m_RowStore
        Case ctRowUniform
            RowMapAt = m_UniMap
        Case ctRowDense
            If r >= 1 And r <= m_RowCount Then RowMapAt = m_RowMap(r)
        Case Else
            RowMapAt = p_BandMapAt(r)
    End Select
End Function

Private Function p_BandMapAt(ByVal r As Long) As Long
    Dim lo As Long, hi As Long, mid As Long
    lo = 1: hi = m_BandCount
    Do While lo <= hi
        mid = (lo + hi) \ 2
        If m_BandStart(mid) > r Then
            hi = mid - 1
        Else
            lo = mid + 1
        End If
    Loop
    If hi >= 1 Then p_BandMapAt = m_BandMap(hi) Else p_BandMapAt = m_UniMap
End Function

' Source ID for one cell. O(1).
Public Function IdAt(ByVal r As Long, ByVal c As Long) As Long
    If r < 1 Or r > m_RowCount Or c < 1 Or c > m_ColCount Then
        IdAt = ctSrcUnknown
        Exit Function
    End If

    If m_AnyOverlay Then
        Select Case m_ColMode(c)
            Case 1
                IdAt = m_ColStamp(c)
                Exit Function
            Case 2
                If m_ColSparse(c).Exists(r) Then
                    IdAt = m_ColSparse(c)(r)
                    Exit Function
                End If
            Case 3
                Dim v As Long
                v = m_ColDense(c).item(r)
                If v <> 0 Then
                    IdAt = v
                    Exit Function
                End If
        End Select
    End If

    IdAt = cLineageRegistry.MapEntry(RowMapAt(r), c)
End Function

' Bulk read: every row's source ID for one column, in one pass.
Public Function ColumnIds(ByVal c As Long) As Long()
    Dim out() As Long
    If m_RowCount < 1 Then
        ReDim out(1 To 1)
        ColumnIds = out
        Exit Function
    End If

    ReDim out(1 To m_RowCount)
    Dim r As Long
    For r = 1 To m_RowCount
        out(r) = IdAt(r, c)
    Next r
    ColumnIds = out
End Function

' Single source ID if the whole column agrees, else 0.
Public Function ColumnUniformSource(ByVal c As Long) As Long
    If m_AnyOverlay Then
        If m_ColMode(c) = 1 Then
            ColumnUniformSource = m_ColStamp(c)
            Exit Function
        End If
        If m_ColMode(c) >= 2 Then Exit Function      ' mixed by construction
    End If

    If m_RowStore = ctRowUniform Then
        ColumnUniformSource = cLineageRegistry.MapEntry(m_UniMap, c)
        Exit Function
    End If

    Dim first As Long, r As Long, v As Long
    first = cLineageRegistry.MapEntry(RowMapAt(1), c)
    For r = 2 To m_RowCount
        v = cLineageRegistry.MapEntry(RowMapAt(r), c)
        If v <> first Then Exit Function
    Next r
    ColumnUniformSource = first
End Function

' Distinct source IDs in ONE column. One pass over the rows: fine for
' reporting, not for a hot loop. Returns an empty array when there are no rows.
Public Function ColumnDistinctSourceIds(ByVal c As Long) As Variant
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    If c >= 1 And c <= m_ColCount Then
        Dim r As Long
        For r = 1 To m_RowCount
            d(IdAt(r, c)) = 1
        Next r
    End If
    ColumnDistinctSourceIds = d.keys
End Function

' Distinct source IDs present anywhere in the table.
Public Function DistinctSourceIds() As Variant
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")

    Dim maps As Object
    Set maps = p_DistinctMapIds
    Dim k As Variant, c As Long
    For Each k In maps.keys
        For c = 1 To m_ColCount
            d(cLineageRegistry.MapEntry(CLng(k), c)) = 1
        Next c
    Next k

    If m_AnyOverlay Then
        For c = 1 To m_ColCount
            Select Case m_ColMode(c)
                Case 1
                    d(m_ColStamp(c)) = 1
                Case 2
                    For Each k In m_ColSparse(c).Items
                        d(CLng(k)) = 1
                    Next k
                Case 3
                    Dim r As Long, v As Long
                    For r = 1 To m_RowCount
                        v = m_ColDense(c).item(r)
                        If v <> 0 Then d(v) = 1
                    Next r
            End Select
        Next c
    End If

    DistinctSourceIds = d.keys
End Function

'==============================================================================
' WRITES  (tiers 1 and 2)
'==============================================================================

' Replace lineage for an entire column - AddComputedColumn, astype-with-loss,
' bulk column replacement.
Public Sub StampColumn(ByVal c As Long, ByVal srcId As Long)
    If c < 1 Or c > m_ColCount Then Exit Sub
    p_EnsureOverlayArrays
    m_ColMode(c) = 1
    m_ColStamp(c) = srcId
    Set m_ColSparse(c) = Nothing
    Set m_ColDense(c) = Nothing
End Sub

' Replace lineage for one cell - SetValue, FillNull, targeted repair.
Public Sub SetCell(ByVal r As Long, ByVal c As Long, ByVal srcId As Long)
    If r < 1 Or r > m_RowCount Or c < 1 Or c > m_ColCount Then Exit Sub
    p_EnsureOverlayArrays

    Select Case m_ColMode(c)
        Case 0
            Set m_ColSparse(c) = CreateObject("Scripting.Dictionary")
            m_ColMode(c) = 2
            m_ColSparse(c)(r) = srcId

        Case 1
            If m_ColStamp(c) = srcId Then Exit Sub
            ' A stamped column covers every row, so demote to dense.
            Dim dv As cLongVec
            Set dv = New cLongVec
            dv.Init m_RowCount
            dv.Fill m_ColStamp(c)
            dv.SetItem r, srcId
            Set m_ColDense(c) = dv
            m_ColMode(c) = 3
            m_ColStamp(c) = 0

        Case 2
            m_ColSparse(c)(r) = srcId
            If p_ShouldDensify(m_ColSparse(c).count) Then p_Densify c

        Case 3
            m_ColDense(c).SetItem r, srcId
    End Select
End Sub

Private Function p_ShouldDensify(ByVal n As Long, _
                                 Optional ByVal rowsTotal As Long = -1) As Boolean
    If rowsTotal < 0 Then rowsTotal = m_RowCount
    If n < OVERLAY_DENSE_MIN Then Exit Function
    p_ShouldDensify = (n > rowsTotal \ OVERLAY_DENSE_DIV)
End Function

Private Sub p_Densify(ByVal c As Long)
    Dim dv As cLongVec
    Set dv = New cLongVec
    dv.Init m_RowCount
    Dim k As Variant
    For Each k In m_ColSparse(c).keys
        dv.SetItem CLng(k), m_ColSparse(c)(k)
    Next k
    Set m_ColDense(c) = dv
    Set m_ColSparse(c) = Nothing
    m_ColMode(c) = 3
End Sub

' Bulk overlay load. Zero means "no override". Picks sparse or dense itself.
Public Sub x_SetColumnOverlay(ByVal c As Long, ByRef vals() As Long, _
                              ByVal n As Long)
    Dim Nz As Long, i As Long
    For i = 1 To n
        If vals(i) <> 0 Then Nz = Nz + 1
    Next i

    If Nz = 0 Then
        If m_AnyOverlay Then
            m_ColMode(c) = 0
            Set m_ColSparse(c) = Nothing
            Set m_ColDense(c) = Nothing
        End If
        Exit Sub
    End If

    p_EnsureOverlayArrays

    If p_ShouldDensify(Nz, n) Then
        Dim dv As cLongVec
        Set dv = New cLongVec
        dv.FromArray vals, n
        Set m_ColDense(c) = dv
        Set m_ColSparse(c) = Nothing
        m_ColMode(c) = 3
    Else
        Dim d As Object
        Set d = CreateObject("Scripting.Dictionary")
        For i = 1 To n
            If vals(i) <> 0 Then d(i) = vals(i)
        Next i
        Set m_ColSparse(c) = d
        Set m_ColDense(c) = Nothing
        m_ColMode(c) = 2
    End If
End Sub

' Raw overlay value for a cell, 0 if none. Used by the transform routines.
Private Function p_OverlayAt(ByVal r As Long, ByVal c As Long) As Long
    If Not m_AnyOverlay Then Exit Function
    Select Case m_ColMode(c)
        Case 1
            p_OverlayAt = m_ColStamp(c)
        Case 2
            If m_ColSparse(c).Exists(r) Then p_OverlayAt = m_ColSparse(c)(r)
        Case 3
            p_OverlayAt = m_ColDense(c).item(r)
    End Select
End Function

'==============================================================================
' ROW TRANSFORM  (Where, head, tail, dropna, drop_duplicates, sort, ...)
'
' Every row-order or row-subset change in cTable routes through here.
' rows(i) is the source row that becomes output row i.
'==============================================================================

Public Sub x_ApplyRowGather(ByRef rows() As Long, ByVal n As Long)
    If n <= 0 Then
        m_RowCount = 0
        m_RowStore = ctRowUniform
        p_ClearOverlays
        Exit Sub
    End If

    ' Overlays first: they are read with the OLD row numbering, so
    ' m_RowCount must not change until every read is done.
    If m_AnyOverlay Then
        Dim c As Long, i As Long
        Dim buf() As Long
        ReDim buf(1 To n)
        Dim oldMode() As Byte
        oldMode = m_ColMode

        For c = 1 To m_ColCount
            If oldMode(c) = 1 Then
                ' A stamp covers all rows; it survives any gather unchanged.
            ElseIf oldMode(c) >= 2 Then
                For i = 1 To n
                    buf(i) = p_OverlayAt(rows(i), c)
                Next i
                p_ResetColumnOverlay c
                x_SetColumnOverlay c, buf, n
            End If
        Next c
    End If

    ' Row store.
    If m_RowStore <> ctRowUniform Then
        Dim newRM() As Long
        ReDim newRM(1 To n)
        Dim j As Long
        For j = 1 To n
            newRM(j) = RowMapAt(rows(j))
        Next j
        m_RowCount = n
        p_AdoptDenseRowMap newRM, n
    Else
        m_RowCount = n
    End If
End Sub

Private Sub p_ResetColumnOverlay(ByVal c As Long)
    If Not m_AnyOverlay Then Exit Sub
    m_ColMode(c) = 0
    m_ColStamp(c) = 0
    Set m_ColSparse(c) = Nothing
    Set m_ColDense(c) = Nothing
End Sub

' Adopt a dense row map, collapsing it to uniform or banded when possible.
Private Sub p_AdoptDenseRowMap(ByRef rm() As Long, ByVal n As Long)
    Dim runs As Long, i As Long
    runs = 1
    For i = 2 To n
        If rm(i) <> rm(i - 1) Then runs = runs + 1
    Next i

    If runs = 1 Then
        m_RowStore = ctRowUniform
        m_UniMap = rm(1)
        Erase m_BandStart
        Erase m_BandMap
        m_BandCount = 0
        Erase m_RowMap
        Exit Sub
    End If

    If runs <= MAX_BANDS Then
        ReDim m_BandStart(1 To runs)
        ReDim m_BandMap(1 To runs)
        Dim b As Long
        b = 1
        m_BandStart(1) = 1
        m_BandMap(1) = rm(1)
        For i = 2 To n
            If rm(i) <> rm(i - 1) Then
                b = b + 1
                m_BandStart(b) = i
                m_BandMap(b) = rm(i)
            End If
        Next i
        m_BandCount = runs
        m_RowStore = ctRowBanded
        Erase m_RowMap
        Exit Sub
    End If

    m_RowMap = rm
    m_RowStore = ctRowDense
    Erase m_BandStart
    Erase m_BandMap
    m_BandCount = 0
End Sub

' Row appended with no known origin (AddRow with literal data).
Public Sub x_AppendRow(ByVal srcId As Long)
    Dim newMap As Long
    newMap = cLineageRegistry.UniformMapId(m_ColCount, srcId)

    If m_RowStore = ctRowUniform And (m_RowCount = 0 Or newMap = m_UniMap) Then
        m_RowCount = m_RowCount + 1
        m_UniMap = newMap
        Exit Sub
    End If

    Dim rm() As Long
    ReDim rm(1 To m_RowCount + 1)
    Dim r As Long
    For r = 1 To m_RowCount
        rm(r) = RowMapAt(r)
    Next r
    rm(m_RowCount + 1) = newMap
    m_RowCount = m_RowCount + 1
    p_AdoptDenseRowMap rm, m_RowCount
End Sub

Public Sub x_RemoveRow(ByVal rowNum As Long)
    If rowNum < 1 Or rowNum > m_RowCount Then Exit Sub

    ' Uniform with no overlays: nothing to renumber.
    If m_RowStore = ctRowUniform And Not m_AnyOverlay Then
        m_RowCount = m_RowCount - 1
        Exit Sub
    End If

    Dim rows() As Long
    ReDim rows(1 To m_RowCount)
    Dim i As Long, k As Long
    For i = 1 To m_RowCount
        If i <> rowNum Then
            k = k + 1
            rows(k) = i
        End If
    Next i
    x_ApplyRowGather rows, k
End Sub

'==============================================================================
' COLUMN TRANSFORM  (SelectColumns, InsertColumn, DropCol, ReorderColumn)
'
' oldOf(newPos) = source column position, or 0 for a brand-new column.
'==============================================================================

Public Sub x_ApplyColumnMapping(ByRef oldOf() As Long, ByVal newCount As Long, _
                                ByVal newColSrcId As Long)
    ' Re-intern every live map into the new column space.
    Dim maps As Object
    Set maps = p_DistinctMapIds

    Dim remap As Object
    Set remap = CreateObject("Scripting.Dictionary")

    Dim vec() As Long
    ReDim vec(1 To IIf(newCount > 0, newCount, 1))

    Dim k As Variant, p As Long, oldId As Long
    For Each k In maps.keys
        oldId = CLng(k)
        For p = 1 To newCount
            If oldOf(p) = 0 Then
                vec(p) = newColSrcId
            Else
                vec(p) = cLineageRegistry.MapEntry(oldId, oldOf(p))
            End If
        Next p
        remap(oldId) = cLineageRegistry.mapId(vec, newCount)
    Next k

    p_RemapRowStore remap

    ' Permute the overlays.
    If m_AnyOverlay Then
        Dim oldMode() As Byte, oldStamp() As Long
        Dim oldSparse() As Object, oldDense() As cLongVec
        oldMode = m_ColMode
        oldStamp = m_ColStamp
        oldSparse = m_ColSparse
        oldDense = m_ColDense

        ReDim m_ColMode(1 To newCount)
        ReDim m_ColStamp(1 To newCount)
        ReDim m_ColSparse(1 To newCount)
        ReDim m_ColDense(1 To newCount)

        For p = 1 To newCount
            If oldOf(p) >= 1 Then
                m_ColMode(p) = oldMode(oldOf(p))
                m_ColStamp(p) = oldStamp(oldOf(p))
                Set m_ColSparse(p) = oldSparse(oldOf(p))
                Set m_ColDense(p) = oldDense(oldOf(p))
            End If
        Next p
    End If

    m_ColCount = newCount
End Sub

Private Sub p_RemapRowStore(ByVal remap As Object)
    Select Case m_RowStore
        Case ctRowUniform
            m_UniMap = remap(m_UniMap)
        Case ctRowBanded
            Dim b As Long
            For b = 1 To m_BandCount
                m_BandMap(b) = remap(m_BandMap(b))
            Next b
        Case ctRowDense
            Dim r As Long
            For r = 1 To m_RowCount
                m_RowMap(r) = remap(m_RowMap(r))
            Next r
    End Select
End Sub

Private Function p_DistinctMapIds() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Select Case m_RowStore
        Case ctRowUniform
            d(m_UniMap) = 1
        Case ctRowBanded
            Dim b As Long
            For b = 1 To m_BandCount
                d(m_BandMap(b)) = 1
            Next b
        Case ctRowDense
            Dim r As Long
            For r = 1 To m_RowCount
                d(m_RowMap(r)) = 1
            Next r
    End Select
    Set p_DistinctMapIds = d
End Function

' --- convenience wrappers ----------------------------------------------------

Public Sub x_InsertColumnAt(ByVal pos As Long, ByVal srcId As Long)
    Dim n As Long
    n = m_ColCount + 1
    Dim oldOf() As Long
    ReDim oldOf(1 To n)
    Dim p As Long
    For p = 1 To n
        If p < pos Then
            oldOf(p) = p
        ElseIf p = pos Then
            oldOf(p) = 0
        Else
            oldOf(p) = p - 1
        End If
    Next p
    x_ApplyColumnMapping oldOf, n, srcId
End Sub

Public Sub x_RemoveColumnAt(ByVal pos As Long)
    If m_ColCount <= 1 Then
        x_InitShell 0, m_RowCount
        Exit Sub
    End If
    Dim oldOf() As Long
    ReDim oldOf(1 To m_ColCount - 1)
    Dim src As Long, dst As Long
    For src = 1 To m_ColCount
        If src <> pos Then
            dst = dst + 1
            oldOf(dst) = src
        End If
    Next src
    x_ApplyColumnMapping oldOf, m_ColCount - 1, ctSrcUnknown
End Sub

' colIdx(newPos) = old column position. Handles both SelectColumns and any
' reordering, since both are just a column mapping.
Public Function ProjectColumns(ByRef colIdx() As Long, ByVal n As Long) As cLineage
    Dim out As cLineage
    Set out = CloneMe
    out.x_ApplyColumnMapping colIdx, n, ctSrcUnknown
    Set ProjectColumns = out
End Function

'==============================================================================
' CONCAT / APPEND
'
' targetColOfOther(k) = this table's column position for other's column k,
'                       or 0 if that column is dropped.
'==============================================================================

Public Sub x_AppendFrom(ByVal other As cLineage, _
                        ByRef targetColOfOther() As Long, _
                        ByVal otherColCount As Long)

    Dim oldRC As Long
    oldRC = m_RowCount
    Dim addRC As Long
    addRC = other.rowCount
    If addRC = 0 Then Exit Sub

    ' Translate other's column maps into this table's column space, cached.
    Dim xlate As Object
    Set xlate = CreateObject("Scripting.Dictionary")

    Dim rm() As Long
    ReDim rm(1 To oldRC + addRC)
    Dim r As Long
    For r = 1 To oldRC
        rm(r) = RowMapAt(r)
    Next r

    Dim vec() As Long
    ReDim vec(1 To m_ColCount)
    Dim k As Long, oid As Long

    For r = 1 To addRC
        oid = other.RowMapAt(r)
        If Not xlate.Exists(oid) Then
            Dim c As Long
            For c = 1 To m_ColCount
                vec(c) = ctSrcMissing
            Next c
            For k = 1 To otherColCount
                If targetColOfOther(k) > 0 Then
                    vec(targetColOfOther(k)) = cLineageRegistry.MapEntry(oid, k)
                End If
            Next k
            xlate(oid) = cLineageRegistry.mapId(vec, m_ColCount)
        End If
        rm(oldRC + r) = xlate(oid)
    Next r

    ' Carry over other's overlays, offset into the appended row range.
    If other.HasOverlays Then
        Dim buf() As Long
        ReDim buf(1 To oldRC + addRC)
        For k = 1 To otherColCount
            Dim tc As Long
            tc = targetColOfOther(k)
            If tc > 0 Then
                If other.x_ColMode(k) <> 0 Then
                    For r = 1 To oldRC
                        buf(r) = p_OverlayAt(r, tc)
                    Next r
                    For r = 1 To addRC
                        buf(oldRC + r) = other.x_RawOverlayAt(r, k)
                    Next r
                    p_ResetColumnOverlay tc
                    x_SetColumnOverlay tc, buf, oldRC + addRC
                End If
            End If
        Next k
    End If

    m_RowCount = oldRC + addRC
    p_AdoptDenseRowMap rm, m_RowCount
End Sub

Public Function x_RawOverlayAt(ByVal r As Long, ByVal c As Long) As Long
    x_RawOverlayAt = p_OverlayAt(r, c)
End Function

'==============================================================================
' JOIN
'
' outLeft(i)  = left row that produced output row i
' outRight(i) = right row, or 0 for a left-join non-match
' rSrcCol(j)  = right table's original column index for output column
'               leftColCount + j
'==============================================================================

Public Sub x_InitJoin(ByVal leftLin As cLineage, ByVal rightLin As cLineage, _
                      ByRef outLeft() As Long, ByRef outRight() As Long, _
                      ByVal n As Long, _
                      ByRef rSrcCol() As Long, ByVal rUsed As Long)

    Dim lcc As Long
    lcc = leftLin.ColumnCount
    m_ColCount = lcc + rUsed
    m_RowCount = n
    p_ClearOverlays

    If n = 0 Then
        m_RowStore = ctRowUniform
        m_UniMap = cLineageRegistry.UniformMapId(m_ColCount, ctSrcUnknown)
        Exit Sub
    End If

    Dim combo As Object
    Set combo = CreateObject("Scripting.Dictionary")

    Dim rm() As Long
    ReDim rm(1 To n)
    Dim vec() As Long
    ReDim vec(1 To m_ColCount)

    Dim i As Long, c As Long, lm As Long, rmp As Long
    Dim key As String

    For i = 1 To n
        ' CHANGED: a row with no left match is legal now (right / outer join).
        If outLeft(i) = 0 Then lm = 0 Else lm = leftLin.RowMapAt(outLeft(i))
        If outRight(i) = 0 Then rmp = 0 Else rmp = rightLin.RowMapAt(outRight(i))
        key = lm & ":" & rmp

        If Not combo.Exists(key) Then
            For c = 1 To lcc
                ' CHANGED: mirror the right-side treatment. A left column on a
                ' right-only row holds a Null this join manufactured.
                If lm = 0 Then
                    vec(c) = ctSrcMissing
                Else
                    vec(c) = cLineageRegistry.MapEntry(lm, c)
                End If
            Next c
            For c = 1 To rUsed
                If rmp = 0 Then
                    vec(lcc + c) = ctSrcMissing
                Else
                    vec(lcc + c) = cLineageRegistry.MapEntry(rmp, rSrcCol(c))
                End If
            Next c
            combo(key) = cLineageRegistry.mapId(vec, m_ColCount)
        End If
        rm(i) = combo(key)
    Next i

    p_AdoptDenseRowMap rm, n

    ' Gather both sides' overlays through the output row order.
    Dim buf() As Long
    ReDim buf(1 To n)

    If leftLin.HasOverlays Then
        For c = 1 To lcc
            If leftLin.x_ColMode(c) <> 0 Then
                For i = 1 To n
                    ' CHANGED: symmetric with the right-side gather below.
                    If outLeft(i) = 0 Then
                        buf(i) = ctSrcMissing
                    Else
                        buf(i) = leftLin.x_RawOverlayAt(outLeft(i), c)
                    End If
                Next i
                x_SetColumnOverlay c, buf, n
            End If
        Next c
    End If

    If rightLin.HasOverlays Then
        For c = 1 To rUsed
            If rightLin.x_ColMode(rSrcCol(c)) <> 0 Then
                For i = 1 To n
                    If outRight(i) = 0 Then
                        buf(i) = ctSrcMissing
                    Else
                        buf(i) = rightLin.x_RawOverlayAt(outRight(i), rSrcCol(c))
                    End If
                Next i
                x_SetColumnOverlay lcc + c, buf, n
            End If
        Next c
    End If
End Sub


'==============================================================================
' DIAGNOSTICS
'==============================================================================

' Human-readable current mode, e.g. "by row". Never raises.
Public Function ModeName() As String
    ModeName = p_ModeName(Mode)
End Function

Public Function Report() As String
    Dim sb As String
    sb = "cLineage  [" & m_RowCount & " rows x " & m_ColCount & " cols]" & vbCrLf
    sb = sb & "  mode      : " & p_ModeName(Mode) & vbCrLf
    sb = sb & "  row store : " & StoreName & vbCrLf

    Dim ids As Variant
    ids = DistinctSourceIds
    sb = sb & "  sources   : " & (UBound(ids) - LBound(ids) + 1) & vbCrLf

    Dim i As Long
    For i = LBound(ids) To UBound(ids)
        sb = sb & "     " & cLineageRegistry.SourceLabel(CLng(ids(i))) & vbCrLf
    Next i
    Report = sb
End Function

' Human-readable row-storage form, e.g. "banded (3 bands)". Never raises.
Public Function StoreName() As String
    Select Case m_RowStore
        Case ctRowUniform: StoreName = "uniform (1 map)"
        Case ctRowBanded:  StoreName = "banded (" & m_BandCount & " bands)"
        Case Else:         StoreName = "dense (" & _
                                       Format$(m_RowCount * 4 / 1048576#, "0.0") & " MB)"
    End Select
End Function

Private Function p_ModeName(ByVal m As ctLineageMode) As String
    Select Case m
        Case ctLinOff:      p_ModeName = "off"
        Case ctLinUniform:  p_ModeName = "uniform"
        Case ctLinByColumn: p_ModeName = "by column"
        Case ctLinByRow:    p_ModeName = "by row"
        Case Else:          p_ModeName = "full"
    End Select
End Function
