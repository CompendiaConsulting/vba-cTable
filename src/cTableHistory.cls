VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTableHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

' ============================================================
' cTableHistory - Snapshot/change-history helper for cTable
'
' Copyright (c) 2026 Brian A. Crawford
' SPDX-License-Identifier: MIT
'
' Licensed under the MIT License.
' See LICENSE in the repository root.
' ============================================================
' Version 20260821-v1
'==============================================================================
' cTableHistory - Snapshot/change-history helper for cTable
'
' PURPOSE
'   Builds compact, reusable change-event caches from a cTable containing
'   repeated snapshots identified by a timestamp/date column.
'
' PUBLIC SHAPE
'
'       t.History.Init "RunDate"
'
'   Set h = t.HistoryFor( _
'               "AccountID", _
'               "A100", _
'               Array("Status", "AccountName"))
'
'   Set h = t.HistoryFor( _
'               "AccountID", _
'               Array("A100", "B200"), _
'               Array("CategoryCode", "Status", "AccountName"))
'
' DESIGN
'   * Init configures the time axis and captures the cTable version.
'   * The chronological row/snapshot index is built lazily on the first query.
'   * Each entity context is built once for its selected identity column and
'     cached as a compact event cTable containing only actual changes.
'   * The source cTable is never sorted or modified.
'   * cTable owns this object. This object DOES NOT retain a cTable reference,
'     avoiding a permanent circular reference in VBA's reference-counted
'     object model.
'   * cTable passes itself only during x_HistoryFor, while a context is built.
'
' EVENT TABLE
'   <TimeColumn>
'   Previous<TimeColumn>
'   <EntityColumn>
'   ColumnName
'   BeforeValue
'   AfterValue
'   ChangeType        "Changed", "Added", or "Removed"
'
' ENTITY LIFECYCLE
'   The first snapshot is the baseline: existing entities do NOT produce
'   "Added" events merely because that is the first time the data set sees
'   them. An entity appearing after the first snapshot produces one Added
'   event; an entity disappearing produces one Removed event.
'
'   Added/Removed produce an entity lifecycle row with
'   ColumnName = <EntityColumn>, plus one state row for each non-missing
'   attribute (Null -> value on Added; value -> Null on Removed). The lifecycle
'   row is retained even when HistoryFor supplies a columns filter.
'
' UNIQUENESS
'   A non-missing entity key may occur at most once per snapshot.
'   Rows with missing entity keys are skipped because they cannot be
'   associated with a specific entity.
'
' STALENESS
'   Init records the version armed by cTable.History. cTable must call
'   x_MarkStale whenever its mutation version changes. Queries fail closed
'   while stale; call t.History.Init again.
'
' REQUIREMENTS
'   cTable with the integration edits supplied with this class.
'==============================================================================

Private m_Initialized As Boolean
Private m_Stale As Boolean
Private m_ArmedVersion As Long
Private m_SourceVersion As Long

Private m_TimeColumn As String
Private m_TimeCol As Long

Private m_TimelineBuilt As Boolean
Private m_RowOrder() As Long
Private m_SnapshotStart() As Long
Private m_SnapshotEnd() As Long
Private m_SnapshotDate() As Date
Private m_SnapshotCount As Long

' entity-column name -> compact event cTable
Private m_Caches As Object

'==============================================================================
' LIFECYCLE / OWNER HOOKS
'==============================================================================

Private Sub Class_Initialize()
    p_ResetCaches
End Sub

' Called by cTable.History before returning this helper.
' No parent reference is retained.
Public Sub x_Arm(ByVal currentVersion As Long)
    m_ArmedVersion = currentVersion
End Sub

' Called from cTable's mutation/version chokepoint.
Public Sub x_MarkStale(ByVal currentVersion As Long)
    m_ArmedVersion = currentVersion
    If m_Initialized Then m_Stale = True
End Sub

'==============================================================================
' INITIALIZATION
'==============================================================================

' Configure/rebuild history around one timestamp/date column.
'
' Init deliberately does not copy or retain the source cTable. The timeline is
' materialized on the next HistoryFor call, when cTable passes itself into
' x_HistoryFor. This keeps initialization cheap and avoids an owner/helper
' circular reference.
Public Sub Init(ByVal TimeColumn As String)

    Dim nm As String
    nm = Trim$(TimeColumn)

    If Len(nm) = 0 Then
        Err.Raise vbObjectError + 1600, "cTableHistory.Init", _
                  "timeColumn cannot be blank."
    End If

    m_TimeColumn = nm
    m_TimeCol = 0

    m_SourceVersion = m_ArmedVersion
    m_Initialized = True
    m_Stale = False

    p_ClearTimeline
    p_ResetCaches
End Sub

'==============================================================================
' STATE / DIAGNOSTICS
'==============================================================================

Public Property Get IsInitialized() As Boolean
    IsInitialized = m_Initialized
End Property

Public Property Get IsStale() As Boolean
    IsStale = m_Stale
End Property

Public Property Get TimeColumn() As String
    TimeColumn = m_TimeColumn
End Property

Public Property Get SourceVersion() As Long
    SourceVersion = m_SourceVersion
End Property

Public Property Get TimelineBuilt() As Boolean
    TimelineBuilt = m_TimelineBuilt
End Property

Public Property Get SnapshotCount() As Long
    SnapshotCount = m_SnapshotCount
End Property

Public Property Get ContextCount() As Long
    If Not m_Caches Is Nothing Then ContextCount = m_Caches.count
End Property

Public Function HasContext(ByVal entityColumn As String) As Boolean
    If m_Caches Is Nothing Then Exit Function
    HasContext = m_Caches.Exists(Trim$(entityColumn))
End Function

'==============================================================================
' QUERY ENTRY POINT - called by cTable.HistoryFor
'==============================================================================

' Project-internal public entry point. cTable remains the intended user-facing
' query facade.
Public Function x_HistoryFor( _
        ByVal parent As cTable, _
        ByVal entityColumn As String, _
        Optional ByVal entityValues As Variant, _
        Optional ByVal columns As Variant) As cTable

    Const CALLER As String = "cTable.HistoryFor"

    p_CheckReady parent, CALLER
    p_EnsureTimeline parent, CALLER

    Dim entityName As String
    entityName = Trim$(entityColumn)

    If Len(entityName) = 0 Then
        Err.Raise vbObjectError + 1601, CALLER, _
                  "entityColumn cannot be blank."
    End If

    Dim entityCol As Long
    entityCol = parent.columnNumber(entityName)

    If entityCol = 0 Then
        Err.Raise vbObjectError + 1602, CALLER, _
                  "Entity column not found: """ & entityName & """"
    End If

    If entityCol = m_TimeCol Then
        Err.Raise vbObjectError + 1603, CALLER, _
                  "The time column cannot also be the entity column."
    End If

    p_ValidateOutputNames entityName, CALLER

    Dim cache As cTable

    If m_Caches.Exists(entityName) Then
        Set cache = m_Caches(entityName)
    Else
        Set cache = p_BuildContext(parent, entityName, entityCol)
        m_Caches.Add entityName, cache
    End If

    Set x_HistoryFor = p_FilterCache( _
                            cache, _
                            parent, _
                            entityName, _
                            entityValues, _
                            columns)
End Function

'==============================================================================
' READINESS / STALENESS
'==============================================================================

Private Sub p_CheckReady(ByVal parent As cTable, ByVal callerName As String)

    If Not m_Initialized Then
        Err.Raise vbObjectError + 1604, callerName, _
                  "History has not been initialized. " & _
                  "Call t.History.Init ""RunDate"" first."
    End If

    If parent Is Nothing Then
        Err.Raise vbObjectError + 1605, callerName, _
                  "The source cTable is Nothing."
    End If

    If parent.x_Version <> m_SourceVersion Then
        m_Stale = True
    End If

    If m_Stale Then
        Err.Raise vbObjectError + 1606, callerName, _
                  "History is stale because the source cTable changed after " & _
                  "History.Init. Call t.History.Init again."
    End If
End Sub

'==============================================================================
' TIMELINE
'==============================================================================

Private Sub p_ClearTimeline()
    m_TimelineBuilt = False
    m_TimeCol = 0
    m_SnapshotCount = 0

    Erase m_RowOrder
    Erase m_SnapshotStart
    Erase m_SnapshotEnd
    Erase m_SnapshotDate
End Sub

Private Sub p_EnsureTimeline(ByVal parent As cTable, ByVal callerName As String)

    If m_TimelineBuilt Then Exit Sub

    m_TimeCol = parent.columnNumber(m_TimeColumn)

    If m_TimeCol = 0 Then
        Err.Raise vbObjectError + 1607, callerName, _
                  "History time column not found: """ & m_TimeColumn & """"
    End If

    Dim n As Long
    n = parent.rowCount

    If n = 0 Then
        m_TimelineBuilt = True
        Exit Sub
    End If

    Dim timeKey() As Double
    ReDim timeKey(1 To n)
    ReDim m_RowOrder(1 To n)

    Dim r As Long
    Dim ok As Boolean
    Dim v As Variant

    For r = 1 To n
        v = parent.GetValue(r, m_TimeCol)

        If p_IsNA(v) Then
            Err.Raise vbObjectError + 1608, callerName, _
                      "History time column """ & m_TimeColumn & _
                      """ contains a missing value at row " & r & "."
        End If

        timeKey(r) = p_TimeNumber(v, ok)

        If Not ok Then
            Err.Raise vbObjectError + 1609, callerName, _
                      "Value in history time column """ & m_TimeColumn & _
                      """ at row " & r & " is not date/time-like."
        End If

        m_RowOrder(r) = r
    Next r

    p_SortTimeline m_RowOrder, timeKey, n

    ReDim m_SnapshotStart(1 To n)
    ReDim m_SnapshotEnd(1 To n)
    ReDim m_SnapshotDate(1 To n)

    Dim i As Long

    For i = 1 To n

        Dim isNewSnapshot As Boolean

        If i = 1 Then
            isNewSnapshot = True
        Else
            isNewSnapshot = (timeKey(i) <> timeKey(i - 1))
        End If

        If isNewSnapshot Then

            If m_SnapshotCount > 0 Then
                m_SnapshotEnd(m_SnapshotCount) = i - 1
            End If

            m_SnapshotCount = m_SnapshotCount + 1
            m_SnapshotStart(m_SnapshotCount) = i
            m_SnapshotDate(m_SnapshotCount) = CDate(timeKey(i))
        End If
    Next i

    m_SnapshotEnd(m_SnapshotCount) = n

    ReDim Preserve m_SnapshotStart(1 To m_SnapshotCount)
    ReDim Preserve m_SnapshotEnd(1 To m_SnapshotCount)
    ReDim Preserve m_SnapshotDate(1 To m_SnapshotCount)

    m_TimelineBuilt = True
End Sub

Private Function p_TimeNumber(ByVal v As Variant, ByRef ok As Boolean) As Double

    ok = False
    On Error GoTo Fail

    If VarType(v) = vbDate Then
        p_TimeNumber = CDbl(CDate(v))
    ElseIf IsNumeric(v) Then
        ' Excel/VBA date serial.
        p_TimeNumber = CDbl(CDate(CDbl(v)))
    ElseIf IsDate(v) Then
        p_TimeNumber = CDbl(CDate(v))
    Else
        Exit Function
    End If

    ok = True
    Exit Function

Fail:
    Err.Clear
    p_TimeNumber = 0#
End Function

' Sort source row numbers chronologically without touching the parent cTable.
' Ties preserve original source-row order.
Private Sub p_SortTimeline( _
        ByRef rows() As Long, _
        ByRef timeKey() As Double, _
        ByVal n As Long)

    If n < 2 Then Exit Sub

    Dim gap As Long
    gap = 1

    Do While gap < n \ 3
        gap = gap * 3 + 1
    Loop

    Dim i As Long
    Dim j As Long
    Dim holdRow As Long
    Dim holdTime As Double

    Do While gap >= 1

        For i = gap + 1 To n

            holdRow = rows(i)
            holdTime = timeKey(i)
            j = i

            Do While j > gap

                If Not p_TimeComesAfter( _
                            timeKey(j - gap), rows(j - gap), _
                            holdTime, holdRow) Then Exit Do

                rows(j) = rows(j - gap)
                timeKey(j) = timeKey(j - gap)
                j = j - gap
            Loop

            rows(j) = holdRow
            timeKey(j) = holdTime
        Next i

        gap = (gap - 1) \ 3
    Loop
End Sub

Private Function p_TimeComesAfter( _
        ByVal timeA As Double, _
        ByVal rowA As Long, _
        ByVal timeB As Double, _
        ByVal rowB As Long) As Boolean

    If timeA > timeB Then
        p_TimeComesAfter = True
    ElseIf timeA = timeB Then
        p_TimeComesAfter = (rowA > rowB)
    End If
End Function

'==============================================================================
' CONTEXT CACHE BUILD
'==============================================================================

Private Function p_BuildContext( _
        ByVal parent As cTable, _
        ByVal entityName As String, _
        ByVal entityCol As Long) As cTable

    Dim eventDate() As Date
    Dim previousDate() As Date
    Dim eventEntity() As Variant
    Dim eventColumn() As String
    Dim beforeValue() As Variant
    Dim afterValue() As Variant
    Dim changeType() As String

    Dim eventCount As Long
    Dim capacity As Long

    If m_SnapshotCount = 0 Then
        Set p_BuildContext = p_MakeEventTable( _
                                parent, entityName, entityCol, _
                                eventDate, previousDate, eventEntity, _
                                eventColumn, beforeValue, afterValue, _
                                changeType, 0)
        Exit Function
    End If

    Dim previousEntities As Object
    Set previousEntities = p_SnapshotEntities( _
                                parent, 1, entityName, entityCol)

    Dim currentEntities As Object
    Dim s As Long
    Dim k As Variant
    Dim previousRow As Long
    Dim currentRow As Long
    Dim c As Long
    Dim entityValue As Variant
    Dim before As Variant
    Dim after As Variant

    For s = 2 To m_SnapshotCount

        Set currentEntities = p_SnapshotEntities( _
                                    parent, s, entityName, entityCol)

        ' Existing entities: compare every eligible attribute.
        ' New entities: one lifecycle event.
        For Each k In currentEntities.keys

            currentRow = CLng(currentEntities(k))
            entityValue = parent.GetValue(currentRow, entityCol)

            If previousEntities.Exists(k) Then

                previousRow = CLng(previousEntities(k))

                For c = 1 To parent.ColumnCount

                    If c <> m_TimeCol And c <> entityCol Then

                        before = parent.GetValue(previousRow, c)
                        after = parent.GetValue(currentRow, c)

                        If Not p_ValuesEqual( _
                                    before, after, parent.dtype_at(c)) Then

                            p_AddEvent _
                                eventDate, previousDate, eventEntity, _
                                eventColumn, beforeValue, afterValue, _
                                changeType, eventCount, capacity, _
                                m_SnapshotDate(s), _
                                m_SnapshotDate(s - 1), _
                                entityValue, _
                                parent.columnName(c), _
                                before, after, "Changed"
                        End If
                    End If
                Next c

            Else
                ' Entity lifecycle marker.
                p_AddEvent _
                    eventDate, previousDate, eventEntity, _
                    eventColumn, beforeValue, afterValue, _
                    changeType, eventCount, capacity, _
                    m_SnapshotDate(s), _
                    m_SnapshotDate(s - 1), _
                    entityValue, _
                    entityName, _
                    Null, entityValue, "Added"

                ' Also expose the entity's initial attributes as Null -> value events
                ' so filtered history includes its observed state when it first appears.
                For c = 1 To parent.ColumnCount

                    If c <> m_TimeCol And c <> entityCol Then

                        after = parent.GetValue(currentRow, c)

                        If Not p_IsNA(after) Then
                            p_AddEvent _
                                eventDate, previousDate, eventEntity, _
                                eventColumn, beforeValue, afterValue, _
                                changeType, eventCount, capacity, _
                                m_SnapshotDate(s), _
                                m_SnapshotDate(s - 1), _
                                entityValue, _
                                parent.columnName(c), _
                                Null, after, "Added"
                        End If
                    End If
                Next c
            End If
        Next k

        ' Entities that existed in the prior snapshot but not this one.
        For Each k In previousEntities.keys

            If Not currentEntities.Exists(k) Then

                previousRow = CLng(previousEntities(k))
                entityValue = parent.GetValue(previousRow, entityCol)

                ' Entity lifecycle marker.
                p_AddEvent _
                    eventDate, previousDate, eventEntity, _
                    eventColumn, beforeValue, afterValue, _
                    changeType, eventCount, capacity, _
                    m_SnapshotDate(s), _
                    m_SnapshotDate(s - 1), _
                    entityValue, _
                    entityName, _
                    entityValue, Null, "Removed"

                ' Mirror the added-state behavior as value -> Null events.
                For c = 1 To parent.ColumnCount

                    If c <> m_TimeCol And c <> entityCol Then

                        before = parent.GetValue(previousRow, c)

                        If Not p_IsNA(before) Then
                            p_AddEvent _
                                eventDate, previousDate, eventEntity, _
                                eventColumn, beforeValue, afterValue, _
                                changeType, eventCount, capacity, _
                                m_SnapshotDate(s), _
                                m_SnapshotDate(s - 1), _
                                entityValue, _
                                parent.columnName(c), _
                                before, Null, "Removed"
                        End If
                    End If
                Next c
            End If
        Next k

        Set previousEntities = currentEntities
    Next s

    Set p_BuildContext = p_MakeEventTable( _
                                parent, entityName, entityCol, _
                                eventDate, previousDate, eventEntity, _
                                eventColumn, beforeValue, afterValue, _
                                changeType, eventCount)
End Function

' Build entityKey -> source row for one snapshot.
' Missing entity keys are intentionally skipped.
Private Function p_SnapshotEntities( _
        ByVal parent As cTable, _
        ByVal snapshotNo As Long, _
        ByVal entityName As String, _
        ByVal entityCol As Long) As Object

    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare

    Dim pos As Long
    Dim sourceRow As Long
    Dim v As Variant
    Dim key As String

    For pos = m_SnapshotStart(snapshotNo) To m_SnapshotEnd(snapshotNo)

        sourceRow = m_RowOrder(pos)
        v = parent.GetValue(sourceRow, entityCol)

        If Not p_IsNA(v) Then

            key = p_ValueKey(v, parent.dtype_at(entityCol))

            If d.Exists(key) Then
                Err.Raise vbObjectError + 1610, _
                          "cTable.HistoryFor", _
                          "Entity column """ & entityName & _
                          """ is not unique within snapshot " & _
                          Format$(m_SnapshotDate(snapshotNo), _
                                  "yyyy-mm-dd hh:nn:ss") & _
                          ". Duplicate rows include " & _
                          d(key) & " and " & sourceRow & "."
            End If

            d.Add key, sourceRow
        End If
    Next pos

    Set p_SnapshotEntities = d
End Function

'==============================================================================
' EVENT ARRAYS / EVENT TABLE
'==============================================================================

Private Sub p_AddEvent( _
        ByRef eventDate() As Date, _
        ByRef previousDate() As Date, _
        ByRef eventEntity() As Variant, _
        ByRef eventColumn() As String, _
        ByRef beforeValue() As Variant, _
        ByRef afterValue() As Variant, _
        ByRef changeType() As String, _
        ByRef eventCount As Long, _
        ByRef capacity As Long, _
        ByVal atDate As Date, _
        ByVal priorDate As Date, _
        ByVal entityValue As Variant, _
        ByVal columnName As String, _
        ByVal before As Variant, _
        ByVal after As Variant, _
        ByVal kind As String)

    eventCount = eventCount + 1

    p_EnsureEventCapacity _
        eventCount, capacity, _
        eventDate, previousDate, eventEntity, eventColumn, _
        beforeValue, afterValue, changeType

    eventDate(eventCount) = atDate
    previousDate(eventCount) = priorDate
    eventEntity(eventCount) = entityValue
    eventColumn(eventCount) = columnName
    beforeValue(eventCount) = before
    afterValue(eventCount) = after
    changeType(eventCount) = kind
End Sub

Private Sub p_EnsureEventCapacity( _
        ByVal needed As Long, _
        ByRef capacity As Long, _
        ByRef eventDate() As Date, _
        ByRef previousDate() As Date, _
        ByRef eventEntity() As Variant, _
        ByRef eventColumn() As String, _
        ByRef beforeValue() As Variant, _
        ByRef afterValue() As Variant, _
        ByRef changeType() As String)

    If needed <= capacity Then Exit Sub

    Dim newCapacity As Long

    If capacity = 0 Then
        newCapacity = 128
    Else
        newCapacity = capacity * 2
    End If

    Do While newCapacity < needed
        newCapacity = newCapacity * 2
    Loop

    If capacity = 0 Then
        ReDim eventDate(1 To newCapacity)
        ReDim previousDate(1 To newCapacity)
        ReDim eventEntity(1 To newCapacity)
        ReDim eventColumn(1 To newCapacity)
        ReDim beforeValue(1 To newCapacity)
        ReDim afterValue(1 To newCapacity)
        ReDim changeType(1 To newCapacity)
    Else
        ReDim Preserve eventDate(1 To newCapacity)
        ReDim Preserve previousDate(1 To newCapacity)
        ReDim Preserve eventEntity(1 To newCapacity)
        ReDim Preserve eventColumn(1 To newCapacity)
        ReDim Preserve beforeValue(1 To newCapacity)
        ReDim Preserve afterValue(1 To newCapacity)
        ReDim Preserve changeType(1 To newCapacity)
    End If

    capacity = newCapacity
End Sub

Private Function p_MakeEventTable( _
        ByVal parent As cTable, _
        ByVal entityName As String, _
        ByVal entityCol As Long, _
        ByRef eventDate() As Date, _
        ByRef previousDate() As Date, _
        ByRef eventEntity() As Variant, _
        ByRef eventColumn() As String, _
        ByRef beforeValue() As Variant, _
        ByRef afterValue() As Variant, _
        ByRef changeType() As String, _
        ByVal eventCount As Long) As cTable

    Const COLS As Long = 7

    Dim hdrs(1 To COLS) As String

    hdrs(1) = m_TimeColumn
    hdrs(2) = "Previous" & m_TimeColumn
    hdrs(3) = entityName
    hdrs(4) = "ColumnName"
    hdrs(5) = "BeforeValue"
    hdrs(6) = "AfterValue"
    hdrs(7) = "ChangeType"

    Dim outData() As Variant
    ReDim outData(1 To IIf(eventCount > 0, eventCount, 1), 1 To COLS)

    Dim r As Long

    For r = 1 To eventCount
        outData(r, 1) = eventDate(r)
        outData(r, 2) = previousDate(r)
        outData(r, 3) = eventEntity(r)
        outData(r, 4) = eventColumn(r)
        outData(r, 5) = beforeValue(r)
        outData(r, 6) = afterValue(r)
        outData(r, 7) = changeType(r)
    Next r

    Dim out As cTable
    Set out = New cTable

    ' History output is derived metadata. V1 deliberately does not claim
    ' source-cell lineage.
    out.EnableLineage = False
    out.from_array hdrs, outData, eventCount, COLS

    Dim types(1 To COLS) As Variant
    types(1) = CLng(ctDate)
    types(2) = CLng(ctDate)
    types(3) = CLng(parent.dtype_at(entityCol))
    types(4) = CLng(ctString)
    types(5) = CLng(ctVariant)
    types(6) = CLng(ctVariant)
    types(7) = CLng(ctString)

    out.dtypes_array = types

    ' Cache once in the requested stable order. Any filtered result preserves
    ' this order.
    If eventCount > 1 Then
        Dim ignored As cTable
        Set ignored = out.sort_values( _
                            Array(m_TimeColumn, entityName, "ColumnName"), _
                            ascending:=True, _
                            inplace:=True)
    End If

    Set p_MakeEventTable = out
End Function

'==============================================================================
' QUERY FILTER
'==============================================================================

Private Function p_FilterCache( _
        ByVal cache As cTable, _
        ByVal parent As cTable, _
        ByVal entityName As String, _
        ByVal entityValues As Variant, _
        ByVal columns As Variant) As cTable

    Dim entityOutCol As Long
    Dim columnNameCol As Long

    entityOutCol = cache.columnNumber(entityName)
    columnNameCol = cache.columnNumber("ColumnName")

    Dim filterEntities As Boolean
    Dim filterColumns As Boolean

    filterEntities = Not p_ArgIsMissing(entityValues)
    filterColumns = Not p_ArgIsMissing(columns)

    Dim entitySet As Object
    Dim columnSet As Object

    If filterEntities Then
        Set entitySet = p_BuildEntityFilter( _
                                entityValues, _
                                cache.dtype_at(entityOutCol))
    End If

    If filterColumns Then
        Set columnSet = p_BuildColumnFilter( _
                                parent, entityName, columns)
    End If

    Dim selected() As Long
    ReDim selected(1 To IIf(cache.rowCount > 0, cache.rowCount, 1))

    Dim selectedCount As Long
    Dim r As Long
    Dim keep As Boolean
    Dim entityValue As Variant
    Dim changedColumn As String
    Dim key As String

    For r = 1 To cache.rowCount

        keep = True

        If filterEntities Then

            entityValue = cache.GetValue(r, entityOutCol)

            If p_IsNA(entityValue) Then
                keep = False
            Else
                key = p_ValueKey( _
                            entityValue, _
                            cache.dtype_at(entityOutCol))

                keep = entitySet.Exists(key)
            End If
        End If

        If keep And filterColumns Then

            changedColumn = CStr(cache.GetValue(r, columnNameCol))

            ' Lifecycle events are always retained.
            If StrComp(changedColumn, entityName, vbTextCompare) <> 0 Then
                keep = columnSet.Exists(changedColumn)
            End If
        End If

        If keep Then
            selectedCount = selectedCount + 1
            selected(selectedCount) = r
        End If
    Next r

    Set p_FilterCache = p_CopyRows(cache, selected, selectedCount)
End Function

Private Function p_BuildEntityFilter( _
        ByVal values As Variant, _
        ByVal entityType As ctDataType) As Object

    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare

    Dim i As Long
    Dim v As Variant

    If IsArray(values) Then

        If p_ArrayHasElements(values) Then
            For i = LBound(values) To UBound(values)
                v = values(i)
                If Not p_IsNA(v) Then
                    d(p_ValueKey(v, entityType)) = True
                End If
            Next i
        End If

    ElseIf Not p_IsNA(values) Then
        d(p_ValueKey(values, entityType)) = True
    End If

    Set p_BuildEntityFilter = d
End Function

Private Function p_BuildColumnFilter( _
        ByVal parent As cTable, _
        ByVal entityName As String, _
        ByVal columns As Variant) As Object

    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare

    Dim names As Variant

    If IsArray(columns) Then

        If Not p_ArrayHasElements(columns) Then
            Set p_BuildColumnFilter = d
            Exit Function
        End If

        names = columns
    Else
        names = Array(CStr(columns))
    End If

    Dim i As Long
    Dim nm As String

    For i = LBound(names) To UBound(names)

        nm = Trim$(CStr(names(i)))

        If Len(nm) = 0 Then
            Err.Raise vbObjectError + 1611, _
                      "cTable.HistoryFor", _
                      "columns cannot contain a blank column name."
        End If

        If parent.columnNumber(nm) = 0 Then
            Err.Raise vbObjectError + 1612, _
                      "cTable.HistoryFor", _
                      "History column not found: """ & nm & """"
        End If

        If StrComp(nm, m_TimeColumn, vbTextCompare) = 0 Then
            Err.Raise vbObjectError + 1613, _
                      "cTable.HistoryFor", _
                      "The time column is already included in every history " & _
                      "result and cannot be tracked as a changed attribute."
        End If

        d(nm) = True
    Next i

    Set p_BuildColumnFilter = d
End Function

Private Function p_CopyRows( _
        ByVal src As cTable, _
        ByRef selected() As Long, _
        ByVal selectedCount As Long) As cTable

    Dim hdrs() As String
    hdrs = src.HeadersArray

    Dim cc As Long
    cc = src.ColumnCount

    Dim outData() As Variant
    ReDim outData(1 To IIf(selectedCount > 0, selectedCount, 1), _
                  1 To IIf(cc > 0, cc, 1))

    Dim r As Long
    Dim c As Long
    Dim sourceRow As Long

    For r = 1 To selectedCount

        sourceRow = selected(r)

        For c = 1 To cc
            outData(r, c) = src.GetValue(sourceRow, c)
        Next c
    Next r

    Dim out As cTable
    Set out = New cTable
    out.EnableLineage = False
    out.from_array hdrs, outData, selectedCount, cc
    out.dtypes_array = src.dtypes_array

    Set p_CopyRows = out
End Function

'==============================================================================
' VALUE SEMANTICS
'==============================================================================

' Same default NA convention used by cTable: Null, Empty, Error, and
' whitespace-only strings are missing.
Private Function p_IsNA(ByVal v As Variant) As Boolean

    If IsNull(v) Then
        p_IsNA = True
        Exit Function
    End If

    If IsEmpty(v) Then
        p_IsNA = True
        Exit Function
    End If

    If IsError(v) Then
        p_IsNA = True
        Exit Function
    End If

    If VarType(v) = vbString Then
        p_IsNA = (Len(Trim$(CStr(v))) = 0)
    End If
End Function

Private Function p_ValuesEqual( _
        ByVal a As Variant, _
        ByVal b As Variant, _
        ByVal declaredType As ctDataType) As Boolean

    Dim aNA As Boolean
    Dim bNA As Boolean

    aNA = p_IsNA(a)
    bNA = p_IsNA(b)

    If aNA Or bNA Then
        p_ValuesEqual = (aNA And bNA)
        Exit Function
    End If

    On Error GoTo Different

    Select Case declaredType

        Case ctString
            p_ValuesEqual = _
                (StrComp(CStr(a), CStr(b), vbTextCompare) = 0)

        Case ctLong
            p_ValuesEqual = (CLng(a) = CLng(b))

        Case ctDouble
            p_ValuesEqual = (CDbl(a) = CDbl(b))

        Case ctCurrency
            p_ValuesEqual = (CCur(a) = CCur(b))

        Case ctDate
            p_ValuesEqual = (CDbl(CDate(a)) = CDbl(CDate(b)))

        Case ctBoolean
            p_ValuesEqual = (CBool(a) = CBool(b))

        Case Else
            p_ValuesEqual = p_VariantValuesEqual(a, b)
    End Select

    Exit Function

Different:
    Err.Clear
    p_ValuesEqual = False
End Function

Private Function p_VariantValuesEqual( _
        ByVal a As Variant, _
        ByVal b As Variant) As Boolean

    On Error GoTo Different

    If VarType(a) = vbString And VarType(b) = vbString Then
        p_VariantValuesEqual = _
            (StrComp(CStr(a), CStr(b), vbTextCompare) = 0)

    ElseIf VarType(a) = vbDate And VarType(b) = vbDate Then
        p_VariantValuesEqual = _
            (CDbl(CDate(a)) = CDbl(CDate(b)))

    ElseIf VarType(a) = vbBoolean And VarType(b) = vbBoolean Then
        p_VariantValuesEqual = (CBool(a) = CBool(b))

    ElseIf IsNumeric(a) And IsNumeric(b) Then
        p_VariantValuesEqual = (CDbl(a) = CDbl(b))

    ElseIf VarType(a) = VarType(b) Then
        p_VariantValuesEqual = _
            (StrComp(CStr(a), CStr(b), vbTextCompare) = 0)

    Else
        p_VariantValuesEqual = False
    End If

    Exit Function

Different:
    Err.Clear
    p_VariantValuesEqual = False
End Function

' Collision-resistant enough for entity identity within one VBA session.
' Type prefixes prevent String "1", numeric 1, Boolean True, etc. from sharing
' an entity identity in an untyped column.
Private Function p_ValueKey( _
        ByVal v As Variant, _
        ByVal declaredType As ctDataType) As String

    On Error GoTo Fallback

    Select Case declaredType

        Case ctString
            p_ValueKey = "S|" & CStr(v)

        Case ctLong, ctDouble, ctCurrency
            p_ValueKey = "N|" & Trim$(Str$(CDbl(v)))

        Case ctDate
            p_ValueKey = "D|" & Trim$(Str$(CDbl(CDate(v))))

        Case ctBoolean
            If CBool(v) Then
                p_ValueKey = "B|1"
            Else
                p_ValueKey = "B|0"
            End If

        Case Else
            Select Case VarType(v)

                Case vbString
                    p_ValueKey = "S|" & CStr(v)

                Case vbDate
                    p_ValueKey = "D|" & _
                                 Trim$(Str$(CDbl(CDate(v))))

                Case vbBoolean
                    If CBool(v) Then
                        p_ValueKey = "B|1"
                    Else
                        p_ValueKey = "B|0"
                    End If

                Case vbByte, vbInteger, vbLong, vbSingle, _
                     vbDouble, vbCurrency, vbDecimal
                    p_ValueKey = "N|" & Trim$(Str$(CDbl(v)))

                Case Else
                    p_ValueKey = "V|" & typeName(v) & "|" & CStr(v)
            End Select
    End Select

    Exit Function

Fallback:
    Err.Clear
    p_ValueKey = "V|" & typeName(v) & "|" & CStr(v)
End Function

'==============================================================================
' ARGUMENT / CACHE HELPERS
'==============================================================================

Private Function p_ArgIsMissing(ByVal v As Variant) As Boolean

    If IsMissing(v) Then
        p_ArgIsMissing = True
        Exit Function
    End If

    ' An omitted Optional forwarded through another procedure can arrive as
    ' an Error variant instead of retaining IsMissing=True.
    On Error Resume Next
    p_ArgIsMissing = IsError(v)
    On Error GoTo 0
End Function

Private Function p_ArrayHasElements(ByVal values As Variant) As Boolean

    If Not IsArray(values) Then Exit Function

    Dim lo As Long
    Dim hi As Long

    On Error GoTo EmptyArray

    lo = LBound(values)
    hi = UBound(values)

    p_ArrayHasElements = (hi >= lo)
    Exit Function

EmptyArray:
    p_ArrayHasElements = False
End Function

Private Sub p_ResetCaches()

    Set m_Caches = CreateObject("Scripting.Dictionary")
    m_Caches.CompareMode = vbTextCompare
End Sub

Private Sub p_ValidateOutputNames( _
        ByVal entityName As String, _
        ByVal callerName As String)

    Dim previousTimeName As String
    previousTimeName = "Previous" & m_TimeColumn

    If StrComp(entityName, previousTimeName, vbTextCompare) = 0 _
       Or StrComp(entityName, "ColumnName", vbTextCompare) = 0 _
       Or StrComp(entityName, "BeforeValue", vbTextCompare) = 0 _
       Or StrComp(entityName, "AfterValue", vbTextCompare) = 0 _
       Or StrComp(entityName, "ChangeType", vbTextCompare) = 0 Then

        Err.Raise vbObjectError + 1614, callerName, _
                  "Entity column """ & entityName & """ conflicts with a " & _
                  "reserved history-output column name."
    End If

    If StrComp(m_TimeColumn, "ColumnName", vbTextCompare) = 0 _
       Or StrComp(m_TimeColumn, "BeforeValue", vbTextCompare) = 0 _
       Or StrComp(m_TimeColumn, "AfterValue", vbTextCompare) = 0 _
       Or StrComp(m_TimeColumn, "ChangeType", vbTextCompare) = 0 Then

        Err.Raise vbObjectError + 1615, callerName, _
                  "Time column """ & m_TimeColumn & """ conflicts with a " & _
                  "reserved history-output column name."
    End If
End Sub
