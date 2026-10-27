VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ============================================================
' cTable - A pandas-inspired data manipulation library for VBA
'
' Copyright (c) 2026 Brian A. Crawford
' SPDX-License-Identifier: MIT
'
' Licensed under the MIT License.
' See LICENSE in the repository root.
' ============================================================
'
' Version 20261007 - name-or-index column references, configurable Null filters
' cTable  -  A pandas-inspired tabular data class for VBA / Excel
' REQUIRES: Microsoft Scripting Runtime reference module
' COMPATIBILITY CHANGES from V1 match pandas names for similar methods:
'       RENAMED LoadFromAccessTable to read_access_table
'       RENAMED LoadFromArrays to from_array
'       RENAMED LoadFromExcelRange to read_excel_range
'       RENAMED read_excel_sheet to read_excel_sheet
'       RENAMED LoadFromExcelTable to read_excel_table
'       RENAMED read_excel_ws to read_excel_ws
'       RENAMED ToExcelRange to to_excel_range
'       RENAMED BuildIndex to set_index
'       RENAMED IndexLookup to index_lookup
'       RENAMED Distinct to unique
'
' NEW IN 20260723b - COLUMN DATA TYPE SYSTEM (pandas-style, self-contained):
'       ADDED Public Enum ctDataType (ctVariant/ctString/ctLong/ctDouble/
'             ctCurrency/ctDate/ctBoolean) and ctConvertErrors
'             (ctErrorsPreserve/ctErrorsCoerce/ctErrorsRaise)
'       ADDED dtype / dtype_at (Get/Let), dtypes, dtypes_string, dtypes_array,
'             set_dtypes_from, dtype_name, dtype_from_name
'       ADDED convert_dtypes (pandas convert_dtypes; auto-called by all read
'             methods via new optional InferTypes parameter, default True)
'       ADDED astype / with_astype / to_numeric / to_datetime
'       ADDED NA / NaT / IsNAValue / isna / notna / count / hasnans /
'             dropna / fillna / info
'       CHANGED to_excel_range and ReplaceExcelTableFrom_cTable: new optional
'             ApplyFormats (default True) applies Excel NumberFormats per
'             column dtype; ctVariant columns are untouched, so untyped
'             tables behave exactly as V1
'       CHANGED derived-table methods (Where, WhereIn, SelectColumns, Clone,
'             MakeJoin, dropna) propagate dtypes to their results
'       CHANGED structural column ops (InsertColumn, DropColNoInPlace,
'             ReorderColumn*) preserve dtypes through the change
'
' NEW IN 20260726 - CELL LINEAGE
'       ADDED EnableLineage (default False; per-project default via
'             cLineageRegistry.LineageDefaultEnabled)
'       ADDED LineageOf / LineageOfName / LineageIdOf / LineageExplain /
'             ColumnLineage / LineageSources / LineageReport / LineageMode /
'             HasMixedRowLineage / HasMixedColumnLineage
'       ADDED StampLineage / SetCellLineage / WriteLineage
'       ADDED LineageModeName / LineageStoreName / LineageSourceCount /
'             ColumnLineageLabel / ColumnSourceCount / LineageColumnReport -
'             all total functions, safe when lineage is off
'       RENAMED LineageSources now returns an array of label STRINGS; the
'             cDataSource objects moved to LineageSourceObjects
'       CHANGED DataSource Get is now a compatibility shim: with lineage on it
'             returns the single source only when the table has one, else
'             Nothing. With lineage off it behaves exactly as V1.
'       CHANGED MakeJoin now carries BOTH sides' lineage (V1 discarded the
'             right table's DataSource entirely)
'       CHANGED p_DoJoin signature: emits the left/right row pairing
'       CHANGED p_ShellSort tracks the row permutation when lineage is active
'       REQUIRES cLineageRegistry (import, do not paste), cLineage, cLongVec,
'             cDataSource v20260726 or later
'
' NEW IN 20260824 - LIMITED SOURCE READS
'       ADDED Optional nrows As Long = 0 to read_excel_sheet,
'             read_excel_table, read_access_table, read_excel_range,
'             read_excel_ws, read_csv, and read_json.
'       nrows counts DATA rows; headers do not count.
'       Excel/Access/limited CSV cap work at the source-read stage.
'       JSON caps materialized records after ParseJson.
'
' NEW IN 20260827
'       CHANGED read_json now allows file path OR SharePoint as input
'       ADDED read_excel_* now include Unmerge option to fill in merged cells
'
' MISSING-VALUE MODEL
'   VBA Null is the universal missing marker, contextual to the column dtype
'   exactly as in pandas (Null in a ctDate column ~ NaT; Null in a ctDouble
'   column ~ NaN). The NA and NaT properties return Null and exist for
'   readable intent. IsNAValue() is the scalar-level test (pandas
'   pd.isna(scalar)); the isna property is the column/table mask (pandas
'   df.isna()). Blank / whitespace-only strings count as NA by default,
'   matching V1 conventions in FillNull and NonAlphaAtEnd sorting.
'==============================================================================
'
' REQUIREMENTS
'   Microsoft Scripting Runtime must be referenced in the VBA project:
'   Tools > References > "Microsoft Scripting Runtime"
'
' INTERNAL STORAGE
'   m_Data()     Variant(1..RowCount, 1..ColCount)  - all cell values
'   m_Headers()  String (1..ColCount)               - column names
'   m_ColTypes() ctDataType (1..ColCount)           - column data types
'                (lazily allocated; unallocated = all ctVariant = V1 behavior)
'   m_Indexes    Dictionary(indexName ->            - named column indexes
'                  Dictionary(keyValue -> Collection of Long row numbers))
'
' CONVENTIONS
'   * All public row/column numbers are 1-based.
'   * Methods that return a new cTable never modify the source table.
'   * "Null" is used for missing values (e.g. left-join non-matches, new columns).
'   * String comparisons are case-insensitive (vbTextCompare) throughout.
'   * Where() supports both equality and numeric range operators.
'==============================================================================

'------------------------------------------------------------------------------
' Valid cTable column data types.
' (Public Enum is legal in a VBA class module and is visible project-wide.)
'------------------------------------------------------------------------------
Public Enum ctDataType
    ctVariant = 0      ' untyped / mixed (pandas "object"). Default; a column
                       ' of ctVariant behaves exactly like cTable V1.
    ctString = 1       ' VBA String
    ctLong = 2         ' VBA Long          (pandas int64)
    ctDouble = 3       ' VBA Double        (pandas float64)
    ctCurrency = 4     ' VBA Currency      (fixed-point, 4 dp)
    ctDate = 5         ' VBA Date          (pandas datetime64)
    ctBoolean = 6      ' VBA Boolean
End Enum

'------------------------------------------------------------------------------
' Conversion error handling (mirrors pandas errors= argument).
' (An enum because Public Const is not allowed in class modules.)
'------------------------------------------------------------------------------
Public Enum ctConvertErrors
    ctErrorsPreserve = 0   ' failure -> keep original value (default; safe)
    ctErrorsCoerce = 1     ' failure -> Null / NA   (pandas errors='coerce')
    ctErrorsRaise = 2      ' failure -> Err.Raise   (pandas errors='raise')
End Enum
' Join types.
Public Enum ctJoinHow
    ctJoinInner = 1
    ctJoinLeft = 2
    ctJoinRight = 3
    ctJoinOuter = 4
    ctJoinCross = 5
End Enum

' Column-naming style for overlapping names.
Public Enum ctJoinNaming
    ctNamePandas = 0        ' overlapping names get suffixes on BOTH sides
    ctNameLegacy = 1        ' right side only, _1 / _2 ... (MakeJoin)
End Enum

Public Enum ctJoinValidate
    ctValidateNone = 0
    ctValidateOneToOne = 1
    ctValidateOneToMany = 2
    ctValidateManyToOne = 3
    ctValidateManyToMany = 4
End Enum
'      Private state
Private m_Data()    As Variant      ' (1..m_RowCount, 1..m_ColCount)
Private m_Headers() As String       ' (1..m_ColCount)
Private m_RowCount  As Long
Private m_ColCount  As Long
Private m_Indexes   As Object       ' Scripting.Dictionary

Private m_ColTypes() As ctDataType  ' (1..m_ColCount); unallocated = ctVariant
Private m_TypesAllocated As Boolean

Private m_DataSource As cDataSource
Private m_Log As cLog
Private m_History As cTableHistory  ' History object

Private m_Lin As cLineage           ' Nothing when lineage tracking is off
Private m_WriteSrcId As Long        ' source attributed to manual cell writes
Private m_Version As Long           ' bumped by every structural mutation


'==============================================================================
' CLASS LIFECYCLE
'==============================================================================

Private Sub Class_Initialize()
    m_RowCount = 0
    m_ColCount = 0
    m_TypesAllocated = False
    Set m_Indexes = CreateObject("Scripting.Dictionary")
    m_Indexes.CompareMode = vbTextCompare

    m_WriteSrcId = ctSrcManual
    If cLineageRegistry.LineageDefaultEnabled Then
        Set m_Lin = New cLineage
        m_Lin.InitUniform 0, 0, ctSrcUnknown
    End If
End Sub

Private Sub Class_Terminate()
    Set m_Indexes = Nothing
    Set m_Lin = Nothing
    Set m_History = Nothing
End Sub

'==============================================================================
'
' LINEAGE - PRIVATE HELPERS
'
' Every hook in this class is guarded by "If Not m_Lin Is Nothing". With
' lineage disabled the class behaves exactly as it did before.
'
'==============================================================================

' Guarantee the engine's shape matches the table's. Any structural path that
' bypassed the lineage hooks is caught here and degraded to Unknown rather
' than left reporting stale positions. Lineage degrades; it never lies.
Private Sub p_LinSync()
    If m_Lin Is Nothing Then Exit Sub
    If m_Lin.ColumnCount <> m_ColCount Or m_Lin.rowCount <> m_RowCount Then
        m_Lin.InitUniform m_RowCount, m_ColCount, ctSrcUnknown
        If m_Log Is Nothing Then Set m_Log = New cLog
        m_Log.AddStatus "cTable lineage: shape drift detected; reset to Unknown."
    End If
End Sub

' Record the current m_DataSource as the origin of every cell. Called at the
' end of each read_* method, once RowCount and ColumnCount are final.
Private Sub p_LinAdoptSource()
    If m_Lin Is Nothing Then Exit Sub
    m_Lin.InitUniform m_RowCount, m_ColCount, p_LinSrcId(m_DataSource)
End Sub

' True when a Dictionary.Keys result actually holds something. An empty
' Dictionary yields LBound 0 / UBound -1, which trips naive count arithmetic.
Private Function p_LinHasIds(ByRef ids As Variant) As Boolean
    If Not IsArray(ids) Then Exit Function
    p_LinHasIds = (UBound(ids) >= LBound(ids))
End Function

Private Function p_LinSrcId(ByVal ds As cDataSource) As Long
    If ds Is Nothing Then
        p_LinSrcId = ctSrcUnknown
    Else
        p_LinSrcId = cLineageRegistry.SourceId(ds)
    End If
End Function

' Attach a row-gathered copy of this table's lineage to a derived table.
Private Sub p_LinAttachGather(ByVal target As cTable, _
                              ByRef rows() As Long, ByVal n As Long)
    If m_Lin Is Nothing Then Exit Sub
    Dim lg As cLineage
    Set lg = m_Lin.CloneMe
    lg.x_ApplyRowGather rows, n
    target.x_SetLineage lg
End Sub

' Another cTable's engine, or a fresh Unknown engine if it has none. Used by
' MakeJoin and Append so a tracked table can absorb an untracked one.
Private Function p_LinOther(ByVal other As cTable) As cLineage
    Dim lg As cLineage
    Set lg = other.x_Lineage
    If lg Is Nothing Then
        Set lg = New cLineage
        lg.InitUniform other.rowCount, other.ColumnCount, ctSrcUnknown
    End If
    Set p_LinOther = lg
End Function

'==============================================================================
'
' PRIVATE HELPERS
'
'==============================================================================
' ---------------------------------------------------------------------------
' Convert the nunique axis argument to 0 or 1.
'
' Supported values follow pandas:
'   0 or "index"   = calculate one unique count per column
'   1 or "columns" = calculate one unique count per row
' ---------------------------------------------------------------------------

'==============================================================================
'
' GROUPED RANK AND CUMULATIVE OPERATIONS
'
' All of these share the ordering contract established by diff: nothing sorts,
' every operation walks rows in their CURRENT order, and grouping is by the
' class-wide vbTextCompare convention. Rows whose group key is NA are excluded
' by default and receive Null.
'
' Rank is the exception to "order does not matter" being safe: rank READS the
' values rather than the positions, so row order changes nothing about the
' ranks themselves - only method:="first" and tie order depend on position.
'
'==============================================================================

' ---------------------------------------------------------------------------
' SHARED ARGUMENT AND LABEL HELPERS
' ---------------------------------------------------------------------------

' True when an Optional Variant was omitted, including when it was forwarded
' from another procedure's omitted Optional (which arrives as an Error
' variant rather than as Missing).
Private Function p_ArgIsMissing(ByVal v As Variant) As Boolean
    If IsMissing(v) Then
        p_ArgIsMissing = True
        Exit Function
    End If

    On Error Resume Next
    p_ArgIsMissing = IsError(v)
    On Error GoTo 0
End Function


' True when this row is excluded from grouping because a key value is NA.
Private Function p_RowKeyIsNA(ByVal rowNum As Long, _
                              ByRef keyCols() As Long, _
                              ByVal keyCount As Long, _
                              ByVal dropNAKeys As Boolean) As Boolean

    If Not dropNAKeys Then Exit Function
    If keyCount = 0 Then Exit Function

    Dim i As Long
    For i = 1 To keyCount
        If p_IsNA(m_Data(rowNum, keyCols(i))) Then
            p_RowKeyIsNA = True
            Exit Function
        End If
    Next i
End Function


' Append one "name=value" fragment to a comma-separated argument list.
Private Sub p_AddExtra(ByRef extras As String, ByVal fragment As String)
    If Len(extras) > 0 Then extras = extras & ", "
    extras = extras & fragment
End Sub


' ---------------------------------------------------------------------------
' The lineage label for a grouped operation, e.g.
'
'   diff(Expenditure by AccountID)
'   rank(Expenditure by AccountID, method=min, pct=True)
'   cumsum(Expenditure by AccountID, CategoryCode)
'
' Only NON-DEFAULT arguments appear. That keeps the common label short and,
' more importantly, keeps it STABLE: adding a new argument with a default
' does not silently re-intern every existing calculated source under a new
' fingerprint.
' ---------------------------------------------------------------------------
Private Function p_OpLabel(ByVal opName As String, _
                           ByVal colName As String, _
                           ByRef keyCols() As Long, _
                           ByVal keyCount As Long, _
                           ByVal extras As String) As String

    Dim s As String
    s = opName & "(" & colName

    If keyCount > 0 Then
        Dim i As Long
        Dim k As String

        For i = 1 To keyCount
            If i > 1 Then k = k & ", "
            k = k & m_Headers(keyCols(i))
        Next i

        s = s & " by " & k
    End If

    If Len(extras) > 0 Then s = s & ", " & extras

    p_OpLabel = s & ")"
End Function


' ---------------------------------------------------------------------------
' GROUP MEMBERSHIP
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' Build explicit group membership in compressed form.
'
' Returns the group count, and fills:
'   groupOf(1..RowCount)      group ordinal per row; 0 = excluded (NA key)
'   memberStart(1..count+1)   offsets into member()
'   member(1..totalMembers)   row numbers, ascending within each group
'
' diff and shift use the cheaper p_NeighborInGroup chain, which is all they
' need. rank needs to sort each group's members, and the cumulative operations
' need per-group running state, so both want membership rather than links.
' Both routines derive their keys from p_GroupRowKey, so grouping semantics
' cannot drift between them.
' ---------------------------------------------------------------------------
Private Function p_BuildGroups(ByRef keyCols() As Long, _
                               ByVal keyCount As Long, _
                               ByVal dropNAKeys As Boolean, _
                               ByRef groupOf() As Long, _
                               ByRef memberStart() As Long, _
                               ByRef member() As Long) As Long

    Dim n As Long
    n = m_RowCount

    ReDim groupOf(1 To IIf(n > 0, n, 1))
    ReDim memberStart(1 To 1)
    ReDim member(1 To 1)
    memberStart(1) = 1

    If n = 0 Then Exit Function

    Dim gc As Long
    Dim r As Long

    If keyCount = 0 Then
        ' Ungrouped: one implicit group holding every row.
        gc = 1
        For r = 1 To n
            groupOf(r) = 1
        Next r
    Else
        Dim ids As Object
        Set ids = CreateObject("Scripting.Dictionary")
        ids.CompareMode = vbTextCompare

        Dim key As String

        For r = 1 To n
            If p_RowKeyIsNA(r, keyCols, keyCount, dropNAKeys) Then
                groupOf(r) = 0
            Else
                key = p_GroupRowKey(r, keyCols, keyCount)

                If ids.Exists(key) Then
                    groupOf(r) = ids(key)
                Else
                    gc = gc + 1
                    ids(key) = gc
                    groupOf(r) = gc
                End If
            End If
        Next r
    End If

    If gc = 0 Then Exit Function

    ' Count, then lay out offsets, then fill - a counting sort, so member()
    ' comes out in ascending row order within each group without sorting.
    Dim sizes() As Long
    ReDim sizes(1 To gc)

    Dim total As Long

    For r = 1 To n
        If groupOf(r) > 0 Then
            sizes(groupOf(r)) = sizes(groupOf(r)) + 1
            total = total + 1
        End If
    Next r

    ReDim memberStart(1 To gc + 1)
    memberStart(1) = 1

    Dim g As Long
    For g = 1 To gc
        memberStart(g + 1) = memberStart(g) + sizes(g)
    Next g

    ReDim member(1 To IIf(total > 0, total, 1))

    Dim cursor() As Long
    ReDim cursor(1 To gc)

    For g = 1 To gc
        cursor(g) = memberStart(g)
    Next g

    For r = 1 To n
        g = groupOf(r)
        If g > 0 Then
            member(cursor(g)) = r
            cursor(g) = cursor(g) + 1
        End If
    Next r

    p_BuildGroups = gc
End Function


'==============================================================================
' RANK
'==============================================================================

' method -> 1 average, 2 min, 3 max, 4 first, 5 dense
Private Function p_RankMethodCode(ByVal method As String) As Long
    Select Case LCase$(Trim$(method))
        Case "average": p_RankMethodCode = 1
        Case "min":     p_RankMethodCode = 2
        Case "max":     p_RankMethodCode = 3
        Case "first":   p_RankMethodCode = 4
        Case "dense":   p_RankMethodCode = 5
        Case Else
            Err.Raise vbObjectError + 1331, "cTable.rank", _
                      "method must be ""average"", ""min"", ""max"", " & _
                      """first"", or ""dense""."
    End Select
End Function


' na_option -> 1 keep, 2 top, 3 bottom
Private Function p_RankNaCode(ByVal na_option As String) As Long
    Select Case LCase$(Trim$(na_option))
        Case "keep":   p_RankNaCode = 1
        Case "top":    p_RankNaCode = 2
        Case "bottom": p_RankNaCode = 3
        Case Else
            Err.Raise vbObjectError + 1332, "cTable.rank", _
                      "na_option must be ""keep"", ""top"", or ""bottom""."
    End Select
End Function


' Two rows share a rank position. Missing values tie with each other and with
' nothing else, which is what puts them in a single block under top / bottom.
Private Function p_RankTied(ByVal rowA As Long, ByVal rowB As Long, _
                            ByVal ci As Long) As Boolean

    Dim aNA As Boolean
    Dim bNA As Boolean

    aNA = p_IsNA(m_Data(rowA, ci))
    bNA = p_IsNA(m_Data(rowB, ci))

    If aNA <> bNA Then Exit Function

    If aNA Then
        p_RankTied = True
    Else
        p_RankTied = (p_RankValueAsDouble(rowA, ci) = _
                      p_RankValueAsDouble(rowB, ci))
    End If
End Function


' ---------------------------------------------------------------------------
' Rank one group's members into out(), which is indexed by ROW number.
' Ranks are written as Doubles; the caller narrows them if the method is
' integral.
' ---------------------------------------------------------------------------
Private Sub p_RankGroup(ByVal ci As Long, _
                        ByRef member() As Long, _
                        ByVal firstIdx As Long, _
                        ByVal lastIdx As Long, _
                        ByVal methodCode As Long, _
                        ByVal ascending As Boolean, _
                        ByVal naCode As Long, _
                        ByVal pct As Boolean, _
                        ByRef Out() As Variant)

    Dim m As Long
    m = lastIdx - firstIdx + 1
    If m <= 0 Then Exit Sub

    ' Split into valued and missing, preserving original row order so that
    ' method:="first" and the stable sort both see appearance order.
    Dim valued() As Long
    Dim missing() As Long
    ReDim valued(1 To m)
    ReDim missing(1 To m)

    Dim nV As Long
    Dim nm As Long
    Dim i As Long
    Dim r As Long

    For i = firstIdx To lastIdx
        r = member(i)

        If p_IsNA(m_Data(r, ci)) Then
            nm = nm + 1
            missing(nm) = r
        Else
            nV = nV + 1
            valued(nV) = r
        End If
    Next i

    ' Reuse the class's own stable ranking sort, so tie ordering here matches
    ' what nlargest and nsmallest already produce.
    If nV > 1 Then
        Dim colIdx(1 To 1) As Long
        colIdx(1) = ci

        Dim work() As Long
        ReDim work(1 To nV)

        p_MergeSortRankRows valued, work, 1, nV, colIdx, 1, Not ascending
    End If

    ' Assemble the ranking order according to na_option.
    Dim ordered() As Long
    Dim total As Long

    Select Case naCode

        Case 1      ' keep: missing values are not ranked and consume no slot
            total = nV
            For i = 1 To nm
                Out(missing(i)) = Null
            Next i

            If total = 0 Then Exit Sub
            ReDim ordered(1 To total)
            For i = 1 To nV
                ordered(i) = valued(i)
            Next i

        Case 2      ' top: missing ranks first
            total = m
            ReDim ordered(1 To total)
            For i = 1 To nm
                ordered(i) = missing(i)
            Next i
            For i = 1 To nV
                ordered(nm + i) = valued(i)
            Next i

        Case Else   ' bottom: missing ranks last
            total = m
            ReDim ordered(1 To total)
            For i = 1 To nV
                ordered(i) = valued(i)
            Next i
            For i = 1 To nm
                ordered(nV + i) = missing(i)
            Next i
    End Select

    ' Walk tie runs. Positions i..j in `ordered` are the slots the run
    ' occupies, which is exactly what every method is defined in terms of.
    Dim j As Long
    Dim k As Long
    Dim denseRank As Long
    Dim rk As Double

    i = 1

    Do While i <= total

        j = i
        Do While j < total
            If Not p_RankTied(ordered(j + 1), ordered(i), ci) Then Exit Do
            j = j + 1
        Loop

        denseRank = denseRank + 1

        For k = i To j
            Select Case methodCode
                Case 1:    rk = (CDbl(i) + CDbl(j)) / 2#   ' average
                Case 2:    rk = i                          ' min
                Case 3:    rk = j                          ' max
                Case 4:    rk = k                          ' first
                Case Else: rk = denseRank                  ' dense
            End Select

            Out(ordered(k)) = rk
        Next k

        i = j + 1
    Loop

    If pct Then
        ' pandas divides by the number of ranked observations, except under
        ' method:="dense" where it divides by the highest dense rank - the
        ' count of DISTINCT values - so that the result still reaches 1.0.
        Dim denom As Double

        If methodCode = 5 Then
            denom = denseRank
        Else
            denom = total
        End If

        If denom > 0 Then
            For i = 1 To total
                Out(ordered(i)) = CDbl(Out(ordered(i))) / denom
            Next i
        End If
    End If
End Sub


' ---------------------------------------------------------------------------
' rank
'
' Returns a 1-based Variant array of RowCount elements holding each row's rank
' within its group. Nothing is modified. pandas GroupBy.rank.
'
' method     "average" (default) mean of the slots a tie run occupies
'            "min"     lowest slot in the run
'            "max"     highest slot in the run
'            "first"   ties broken by order of appearance
'            "dense"   like min, but with no gaps after a tie
' ascending  True (default) ranks smallest as 1
' na_option  "keep" (default) missing values rank as Null and consume no slot
'            "top"     missing values rank first
'            "bottom"  missing values rank last
' pct        True returns the rank as a proportion rather than a position
'
' The column must have a rankable dtype - ctLong, ctDouble, ctCurrency,
' ctDate, or ctBoolean - because ranking reuses p_RankValueAsDouble. An
' untyped column raises; call convert_dtypes first.
'
'   v = t.rank("Expenditure", "AccountID")
'   v = t.rank("Expenditure", "AccountID", "min", ascending:=False)
'   v = t.rank("Expenditure", "AccountID", pct:=True)
' ---------------------------------------------------------------------------
Public Function rank(ByVal colName As Variant, _
                     Optional ByVal by As Variant, _
                     Optional ByVal method As String = "average", _
                     Optional ByVal ascending As Boolean = True, _
                     Optional ByVal na_option As String = "keep", _
                     Optional ByVal pct As Boolean = False, _
                     Optional ByVal dropna_keys As Boolean = True) As Variant
    colName = p_LabelColumnRef(colName)

    Dim ci As Long
    ci = p_ColIdx(colName)

    If ci = 0 Then
        Err.Raise vbObjectError + 1303, "cTable.rank", _
                  "Column not found: """ & colName & """"
    End If

    If m_RowCount = 0 Then
        rank = Array()
        Exit Function
    End If

    p_EnsureTypes

    If Not p_IsRankableDType(m_ColTypes(ci)) Then
        Err.Raise vbObjectError + 1333, "cTable.rank", _
                  "Column """ & colName & """ has dtype " & _
                  Me.dtype_name(m_ColTypes(ci)) & " and cannot be ranked. " & _
                  "Call convert_dtypes or astype first."
    End If

    Dim methodCode As Long
    Dim naCode As Long
    methodCode = p_RankMethodCode(method)
    naCode = p_RankNaCode(na_option)

    Dim keyCols() As Long
    Dim keyCount As Long
    keyCount = p_ResolveGroupKeys(keyCols, "cTable.rank", by)

    Dim groupOf() As Long
    Dim memberStart() As Long
    Dim member() As Long
    Dim gc As Long
    gc = p_BuildGroups(keyCols, keyCount, dropna_keys, _
                       groupOf, memberStart, member)

    Dim Out() As Variant
    ReDim Out(1 To m_RowCount)

    Dim r As Long
    For r = 1 To m_RowCount
        Out(r) = Null           ' Variant default is Empty, not Null
    Next r

    Dim g As Long
    For g = 1 To gc
        p_RankGroup ci, member, memberStart(g), memberStart(g + 1) - 1, _
                    methodCode, ascending, naCode, pct, Out
    Next g

    ' Every method except average yields whole positions.
    If methodCode <> 1 And Not pct Then
        For r = 1 To m_RowCount
            If Not IsNull(Out(r)) Then Out(r) = CLng(Out(r))
        Next r
    End If

    rank = Out
End Function


'==============================================================================
' CUMULATIVE OPERATIONS
'==============================================================================

' opCode: 1 cumsum, 2 cumprod, 3 cummax, 4 cummin

' Numeric extraction for the cumulative fold. Tolerant of ctVariant in the
' same way p_DiffPair is, so untyped columns still work.
Private Function p_CumNumeric(ByVal v As Variant, _
                              ByVal colType As ctDataType, _
                              ByRef ok As Boolean) As Double
    ok = True
    On Error GoTo Fail

    Select Case colType
        Case ctBoolean
            p_CumNumeric = p_BoolAsLong(v)

        Case ctDate
            p_CumNumeric = CDbl(CDate(v))

        Case ctLong, ctDouble, ctCurrency
            p_CumNumeric = CDbl(v)

        Case ctString
            GoTo Fail

        Case Else
            If VarType(v) = vbBoolean Then
                p_CumNumeric = p_BoolAsLong(v)
            ElseIf VarType(v) = vbDate Then
                p_CumNumeric = CDbl(CDate(v))
            ElseIf IsNumeric(v) Then
                p_CumNumeric = CDbl(v)
            Else
                GoTo Fail
            End If
    End Select

    Exit Function

Fail:
    Err.Clear
    ok = False
    p_CumNumeric = 0#
End Function


' ---------------------------------------------------------------------------
' The running fold itself. One pass, one slot of state per group.
'
' skipna:=True  (default) an NA yields NA at that position but leaves the
'               running state untouched, so the series continues afterwards.
' skipna:=False the first NA poisons the rest of that group, matching pandas.
' ---------------------------------------------------------------------------
Private Function p_CumulativeValues(ByVal ci As Long, _
                                    ByRef groupOf() As Long, _
                                    ByVal gc As Long, _
                                    ByVal opCode As Long, _
                                    ByVal skipna As Boolean) As Variant

    Dim colType As ctDataType
    colType = m_ColTypes(ci)

    Dim Out() As Variant
    ReDim Out(1 To m_RowCount)

    Dim r As Long
    For r = 1 To m_RowCount
        Out(r) = Null
    Next r

    If gc = 0 Then
        p_CumulativeValues = Out
        Exit Function
    End If

    ' Currency sums stay in Currency so money keeps exact 4-dp arithmetic.
    Dim useCurrency As Boolean
    useCurrency = (opCode = 1 And colType = ctCurrency)

    Dim hasState() As Boolean
    Dim blocked() As Boolean
    Dim dblState() As Double
    Dim curState() As Currency
    Dim rowState() As Long

    ReDim hasState(1 To gc)
    ReDim blocked(1 To gc)
    ReDim dblState(1 To gc)
    ReDim curState(1 To gc)
    ReDim rowState(1 To gc)

    Dim g As Long
    Dim v As Variant
    Dim x As Double
    Dim ok As Boolean

    For r = 1 To m_RowCount

        g = groupOf(r)

        If g > 0 Then
            If Not blocked(g) Then

                v = m_Data(r, ci)

                If p_IsNA(v) Then
                    If Not skipna Then blocked(g) = True
                    ' out(r) is already Null either way.
                Else
                    x = p_CumNumeric(v, colType, ok)

                    If Not ok Then
                        Err.Raise vbObjectError + 1334, "cTable cumulative", _
                                  "Value in column """ & m_Headers(ci) & _
                                  """ at row " & r & " is not numeric."
                    End If

                    Select Case opCode

                        Case 1      ' cumsum
                            If useCurrency Then
                                If hasState(g) Then
                                    curState(g) = curState(g) + CCur(v)
                                Else
                                    curState(g) = CCur(v)
                                End If
                                Out(r) = curState(g)
                            Else
                                If hasState(g) Then
                                    dblState(g) = dblState(g) + x
                                Else
                                    dblState(g) = x
                                End If
                                Out(r) = dblState(g)
                            End If

                        Case 2      ' cumprod
                            If hasState(g) Then
                                dblState(g) = dblState(g) * x
                            Else
                                dblState(g) = x
                            End If
                            Out(r) = dblState(g)

                        Case 3      ' cummax
                            If Not hasState(g) Then
                                dblState(g) = x
                                rowState(g) = r
                            ElseIf x > dblState(g) Then
                                dblState(g) = x
                                rowState(g) = r
                            End If
                            ' Emit the ORIGINAL value, so a Date stays a Date
                            ' and a Currency keeps its scale.
                            Out(r) = m_Data(rowState(g), ci)

                        Case Else   ' cummin
                            If Not hasState(g) Then
                                dblState(g) = x
                                rowState(g) = r
                            ElseIf x < dblState(g) Then
                                dblState(g) = x
                                rowState(g) = r
                            End If
                            Out(r) = m_Data(rowState(g), ci)
                    End Select

                    hasState(g) = True
                End If
            End If
        End If
    Next r

    p_CumulativeValues = Out
End Function


' The dtype a cumulative column should carry.
'
' cumsum WIDENS integers to ctDouble on purpose. A running total overflows
' ctLong at about 2.1 billion, which a column of expenditures reaches easily,
' and ctDouble represents integers exactly to 2^53 so nothing is lost.
Private Function p_CumResultType(ByVal srcType As ctDataType, _
                                 ByVal opCode As Long) As ctDataType
    Select Case opCode

        Case 1      ' cumsum
            Select Case srcType
                Case ctCurrency
                    p_CumResultType = ctCurrency
                Case ctLong, ctDouble, ctBoolean
                    p_CumResultType = ctDouble
                Case Else
                    p_CumResultType = ctVariant
            End Select

        Case 2      ' cumprod - always widened; products escalate fast
            Select Case srcType
                Case ctLong, ctDouble, ctCurrency, ctBoolean
                    p_CumResultType = ctDouble
                Case Else
                    p_CumResultType = ctVariant
            End Select

        Case Else   ' cummax / cummin return values taken from the column
            p_CumResultType = srcType
    End Select
End Function


' Shared front end for the four cumulative functions.
Private Function p_Cumulative(ByVal opName As String, _
                              ByVal opCode As Long, _
                              ByVal colName As String, _
                              ByVal skipna As Boolean, _
                              ByVal dropNAKeys As Boolean, _
                              Optional ByVal by As Variant) As Variant

    Dim callerName As String
    callerName = "cTable." & opName

    Dim ci As Long
    ci = p_ColIdx(colName)

    If ci = 0 Then
        Err.Raise vbObjectError + 1303, callerName, _
                  "Column not found: """ & colName & """"
    End If

    If m_RowCount = 0 Then
        p_Cumulative = Array()
        Exit Function
    End If

    p_EnsureTypes

    Dim colType As ctDataType
    colType = m_ColTypes(ci)

    If colType = ctString Then
        Err.Raise vbObjectError + 1335, callerName, _
                  opName & " is not defined for ctString column """ & _
                  colName & """."
    End If

    If colType = ctDate And opCode <= 2 Then
        Err.Raise vbObjectError + 1336, callerName, _
                  "Adding or multiplying dates is not meaningful. " & _
                  "cummax and cummin do work on ctDate columns."
    End If

    Dim keyCols() As Long
    Dim keyCount As Long
    keyCount = p_ResolveGroupKeys(keyCols, callerName, by)

    Dim groupOf() As Long
    Dim memberStart() As Long
    Dim member() As Long
    Dim gc As Long
    gc = p_BuildGroups(keyCols, keyCount, dropNAKeys, _
                       groupOf, memberStart, member)

    p_Cumulative = p_CumulativeValues(ci, groupOf, gc, opCode, skipna)
End Function


' ---------------------------------------------------------------------------
' cumsum / cumprod / cummax / cummin
'
' Each returns a 1-based Variant array of RowCount elements. Nothing is
' modified. pandas GroupBy.cumsum and friends.
'
' skipna  True (default) NA yields NA at that position without breaking the
'         running series; False poisons the remainder of the group.
'
'   v = t.cumsum("Expenditure", "AccountID")
'   v = t.cummax("RunDate", "AccountID")
' ---------------------------------------------------------------------------
Public Function cumsum(ByVal colName As Variant, _
                       Optional ByVal by As Variant, _
                       Optional ByVal skipna As Boolean = True, _
                       Optional ByVal dropna_keys As Boolean = True) As Variant
    colName = p_LabelColumnRef(colName)
    cumsum = p_Cumulative("cumsum", 1, colName, skipna, dropna_keys, by)
End Function


Public Function cumprod(ByVal colName As Variant, _
                        Optional ByVal by As Variant, _
                        Optional ByVal skipna As Boolean = True, _
                        Optional ByVal dropna_keys As Boolean = True) As Variant
    colName = p_LabelColumnRef(colName)
    cumprod = p_Cumulative("cumprod", 2, colName, skipna, dropna_keys, by)
End Function


Public Function cummax(ByVal colName As Variant, _
                       Optional ByVal by As Variant, _
                       Optional ByVal skipna As Boolean = True, _
                       Optional ByVal dropna_keys As Boolean = True) As Variant
    colName = p_LabelColumnRef(colName)
    cummax = p_Cumulative("cummax", 3, colName, skipna, dropna_keys, by)
End Function


Public Function cummin(ByVal colName As Variant, _
                       Optional ByVal by As Variant, _
                       Optional ByVal skipna As Boolean = True, _
                       Optional ByVal dropna_keys As Boolean = True) As Variant
    colName = p_LabelColumnRef(colName)
    cummin = p_Cumulative("cummin", 4, colName, skipna, dropna_keys, by)
End Function


'==============================================================================
' COLUMN WRITERS
'==============================================================================

' Resolve and validate the destination column name before any work is done,
' so an Add* that is going to be refused fails before computing anything.
Private Function p_ResolveTargetName(ByVal newColName As String, _
                                     ByVal colName As String, _
                                     ByVal suffix As String, _
                                     ByVal allowReplace As Boolean, _
                                     ByVal callerName As String) As String
    Dim tn As String
    tn = Trim$(newColName)
    If Len(tn) = 0 Then tn = colName & suffix

    If Not allowReplace Then
        If p_ColIdx(tn) > 0 Then
            Err.Raise vbObjectError + 1304, callerName, _
                      "Column already exists: """ & tn & """"
        End If
    End If

    p_ResolveTargetName = tn
End Function


' ---------------------------------------------------------------------------
' Write a computed 1-based value array into a new or existing column, apply
' the dtype, and stamp the column's lineage. Shared by every Add*/Set* writer
' in this section.
' ---------------------------------------------------------------------------
Private Sub p_WriteTransformColumn(ByVal targetName As String, _
                                   ByVal values As Variant, _
                                   ByVal finalType As ctDataType, _
                                   ByVal label As String)

    Dim targetCol As Long
    targetCol = p_ColIdx(targetName)

    If targetCol = 0 Then
        Me.InsertColumn targetName
        targetCol = m_ColCount
    End If

    p_EnsureTypes

    Dim r As Long
    For r = 1 To m_RowCount
        m_Data(r, targetCol) = values(r)
    Next r

    If finalType <> ctVariant Then
        ' Group-boundary and excluded-row Nulls convert cleanly: p_TryConvert
        ' passes NA through as Null for every dtype.
        p_CoerceColumn targetCol, finalType, ctErrorsRaise
    End If

    m_ColTypes(targetCol) = finalType

    If Not m_Lin Is Nothing Then
        p_LinSync
        m_Lin.StampColumn targetCol, cLineageRegistry.CalculatedSourceId(label)
    End If

    p_ClearIndexes
End Sub


' ---------------------------------------------------------------------------
' SHIFT WRITER
' ---------------------------------------------------------------------------
Private Sub p_ApplyShift(ByVal colName As String, _
                         ByVal periods As Long, _
                         ByVal newColName As String, _
                         ByVal dropNAKeys As Boolean, _
                         ByVal fill_value As Variant, _
                         ByVal allowReplace As Boolean, _
                         ByVal callerName As String, _
                         Optional ByVal by As Variant)

    Dim srcCol As Long
    srcCol = p_ColIdx(colName)

    If srcCol = 0 Then
        Err.Raise vbObjectError + 1303, callerName, _
                  "Column not found: """ & colName & """"
    End If

    Dim targetName As String
    targetName = p_ResolveTargetName(newColName, colName, "_Lag", _
                                     allowReplace, callerName)

    p_EnsureTypes

    Dim srcType As ctDataType
    srcType = m_ColTypes(srcCol)

    Dim values As Variant
    values = Me.shift(colName, by, periods, dropNAKeys, fill_value)

    Dim keyCols() As Long
    Dim keyCount As Long
    keyCount = p_ResolveGroupKeys(keyCols, callerName, by)

    Dim extras As String
    If periods <> 1 Then p_AddExtra extras, "periods=" & CStr(periods)
    If Not p_ArgIsMissing(fill_value) Then
        p_AddExtra extras, "fill_value=" & p_LabelScalar(fill_value)
    End If

    p_WriteTransformColumn targetName, values, srcType, _
        p_OpLabel("shift", colName, keyCols, keyCount, extras)
End Sub


' A short, stable rendering of a scalar for use inside a lineage label.
Private Function p_LabelScalar(ByVal v As Variant) As String
    If IsNull(v) Or IsEmpty(v) Then
        p_LabelScalar = "Null"
    ElseIf VarType(v) = vbString Then
        p_LabelScalar = """" & CStr(v) & """"
    ElseIf VarType(v) = vbDate Then
        p_LabelScalar = Format$(v, "yyyy-mm-dd")
    Else
        p_LabelScalar = CStr(v)
    End If
End Function


' ---------------------------------------------------------------------------
' RANK WRITER
' ---------------------------------------------------------------------------
Private Sub p_ApplyRank(ByVal colName As String, _
                        ByVal method As String, _
                        ByVal ascending As Boolean, _
                        ByVal na_option As String, _
                        ByVal pct As Boolean, _
                        ByVal dropNAKeys As Boolean, _
                        ByVal newColName As String, _
                        ByVal allowReplace As Boolean, _
                        ByVal callerName As String, _
                        Optional ByVal by As Variant)

    Dim targetName As String
    targetName = p_ResolveTargetName(newColName, colName, "_Rank", _
                                     allowReplace, callerName)

    Dim values As Variant
    values = Me.rank(colName, by, method, ascending, na_option, pct, dropNAKeys)

    Dim keyCols() As Long
    Dim keyCount As Long
    keyCount = p_ResolveGroupKeys(keyCols, callerName, by)

    Dim extras As String
    If p_RankMethodCode(method) <> 1 Then
        p_AddExtra extras, "method=" & LCase$(Trim$(method))
    End If
    If Not ascending Then p_AddExtra extras, "ascending=False"
    If p_RankNaCode(na_option) <> 1 Then
        p_AddExtra extras, "na_option=" & LCase$(Trim$(na_option))
    End If
    If pct Then p_AddExtra extras, "pct=True"

    Dim finalType As ctDataType
    If p_RankMethodCode(method) = 1 Or pct Then
        finalType = ctDouble
    Else
        finalType = ctLong
    End If

    p_WriteTransformColumn targetName, values, finalType, _
        p_OpLabel("rank", colName, keyCols, keyCount, extras)
End Sub


' ---------------------------------------------------------------------------
' CUMULATIVE WRITER
' ---------------------------------------------------------------------------
Private Sub p_ApplyCumulative(ByVal opName As String, _
                              ByVal opCode As Long, _
                              ByVal colName As String, _
                              ByVal skipna As Boolean, _
                              ByVal dropNAKeys As Boolean, _
                              ByVal newColName As String, _
                              ByVal allowReplace As Boolean, _
                              ByVal callerName As String, _
                              Optional ByVal by As Variant)

    Dim srcCol As Long
    srcCol = p_ColIdx(colName)

    If srcCol = 0 Then
        Err.Raise vbObjectError + 1303, callerName, _
                  "Column not found: """ & colName & """"
    End If

    Dim targetName As String
    targetName = p_ResolveTargetName(newColName, colName, _
                                     "_" & opName, allowReplace, callerName)

    p_EnsureTypes

    Dim srcType As ctDataType
    srcType = m_ColTypes(srcCol)

    Dim values As Variant
    values = p_Cumulative(opName, opCode, colName, skipna, dropNAKeys, by)

    Dim keyCols() As Long
    Dim keyCount As Long
    keyCount = p_ResolveGroupKeys(keyCols, callerName, by)

    Dim extras As String
    If Not skipna Then p_AddExtra extras, "skipna=False"

    p_WriteTransformColumn targetName, values, _
        p_CumResultType(srcType, opCode), _
        p_OpLabel(opName, colName, keyCols, keyCount, extras)
End Sub


'==============================================================================
' PUBLIC COLUMN WRITERS
'
' Add* raises if the destination already exists, matching AddComputedColumn.
' Set* creates or replaces, matching SetComputedColumn and the pandas
' assignment idiom df["X"] = ... which overwrites silently.
'
' Default destination names: colName & "_Rank", "_cumsum", "_cumprod",
' "_cummax", "_cummin", "_Lag".
'==============================================================================

Public Sub AddRankColumn(ByVal colName As Variant, _
                         Optional ByVal by As Variant, _
                         Optional ByVal method As String = "average", _
                         Optional ByVal ascending As Boolean = True, _
                         Optional ByVal na_option As String = "keep", _
                         Optional ByVal pct As Boolean = False, _
                         Optional ByVal newColName As String = "", _
                         Optional ByVal dropna_keys As Boolean = True)
    colName = p_LabelColumnRef(colName)

    p_ApplyRank colName, method, ascending, na_option, pct, dropna_keys, _
                newColName, False, "cTable.AddRankColumn", by
End Sub


Public Sub SetRankColumn(ByVal colName As Variant, _
                         Optional ByVal by As Variant, _
                         Optional ByVal method As String = "average", _
                         Optional ByVal ascending As Boolean = True, _
                         Optional ByVal na_option As String = "keep", _
                         Optional ByVal pct As Boolean = False, _
                         Optional ByVal newColName As String = "", _
                         Optional ByVal dropna_keys As Boolean = True)
    colName = p_LabelColumnRef(colName)

    p_ApplyRank colName, method, ascending, na_option, pct, dropna_keys, _
                newColName, True, "cTable.SetRankColumn", by
End Sub


Public Sub AddCumSumColumn(ByVal colName As Variant, _
                           Optional ByVal by As Variant, _
                           Optional ByVal skipna As Boolean = True, _
                           Optional ByVal newColName As String = "", _
                           Optional ByVal dropna_keys As Boolean = True)
    colName = p_LabelColumnRef(colName)

    p_ApplyCumulative "cumsum", 1, colName, skipna, dropna_keys, _
                      newColName, False, "cTable.AddCumSumColumn", by
End Sub


Public Sub SetCumSumColumn(ByVal colName As Variant, _
                           Optional ByVal by As Variant, _
                           Optional ByVal skipna As Boolean = True, _
                           Optional ByVal newColName As String = "", _
                           Optional ByVal dropna_keys As Boolean = True)
    colName = p_LabelColumnRef(colName)

    p_ApplyCumulative "cumsum", 1, colName, skipna, dropna_keys, _
                      newColName, True, "cTable.SetCumSumColumn", by
End Sub


Public Sub AddCumProdColumn(ByVal colName As Variant, _
                            Optional ByVal by As Variant, _
                            Optional ByVal skipna As Boolean = True, _
                            Optional ByVal newColName As String = "", _
                            Optional ByVal dropna_keys As Boolean = True)
    colName = p_LabelColumnRef(colName)

    p_ApplyCumulative "cumprod", 2, colName, skipna, dropna_keys, _
                      newColName, False, "cTable.AddCumProdColumn", by
End Sub


Public Sub SetCumProdColumn(ByVal colName As Variant, _
                            Optional ByVal by As Variant, _
                            Optional ByVal skipna As Boolean = True, _
                            Optional ByVal newColName As String = "", _
                            Optional ByVal dropna_keys As Boolean = True)
    colName = p_LabelColumnRef(colName)

    p_ApplyCumulative "cumprod", 2, colName, skipna, dropna_keys, _
                      newColName, True, "cTable.SetCumProdColumn", by
End Sub


Public Sub AddCumMaxColumn(ByVal colName As Variant, _
                           Optional ByVal by As Variant, _
                           Optional ByVal skipna As Boolean = True, _
                           Optional ByVal newColName As String = "", _
                           Optional ByVal dropna_keys As Boolean = True)
    colName = p_LabelColumnRef(colName)

    p_ApplyCumulative "cummax", 3, colName, skipna, dropna_keys, _
                      newColName, False, "cTable.AddCumMaxColumn", by
End Sub


Public Sub SetCumMaxColumn(ByVal colName As Variant, _
                           Optional ByVal by As Variant, _
                           Optional ByVal skipna As Boolean = True, _
                           Optional ByVal newColName As String = "", _
                           Optional ByVal dropna_keys As Boolean = True)
    colName = p_LabelColumnRef(colName)

    p_ApplyCumulative "cummax", 3, colName, skipna, dropna_keys, _
                      newColName, True, "cTable.SetCumMaxColumn", by
End Sub


Public Sub AddCumMinColumn(ByVal colName As Variant, _
                           Optional ByVal by As Variant, _
                           Optional ByVal skipna As Boolean = True, _
                           Optional ByVal newColName As String = "", _
                           Optional ByVal dropna_keys As Boolean = True)
    colName = p_LabelColumnRef(colName)

    p_ApplyCumulative "cummin", 4, colName, skipna, dropna_keys, _
                      newColName, False, "cTable.AddCumMinColumn", by
End Sub


Public Sub SetCumMinColumn(ByVal colName As Variant, _
                           Optional ByVal by As Variant, _
                           Optional ByVal skipna As Boolean = True, _
                           Optional ByVal newColName As String = "", _
                           Optional ByVal dropna_keys As Boolean = True)
    colName = p_LabelColumnRef(colName)

    p_ApplyCumulative "cummin", 4, colName, skipna, dropna_keys, _
                      newColName, True, "cTable.SetCumMinColumn", by
End Sub

Private Function p_NormalizeNuniqueAxis(ByVal axis As Variant) As Long
    If IsNull(axis) Then
        Err.Raise vbObjectError + 1230, _
                  "cTable.nunique", _
                  "axis cannot be Null."
    End If

    If IsEmpty(axis) Then
        p_NormalizeNuniqueAxis = 0
        Exit Function
    End If

    If VarType(axis) = vbString Then
        Select Case LCase$(Trim$(CStr(axis)))
            Case "0", "index"
                p_NormalizeNuniqueAxis = 0

            Case "1", "columns"
                p_NormalizeNuniqueAxis = 1

            Case Else
                Err.Raise vbObjectError + 1231, _
                          "cTable.nunique", _
                          "axis must be 0, 1, ""index"", or ""columns""."
        End Select

        Exit Function
    End If

    If IsNumeric(axis) Then
        Dim numericAxis As Double
        numericAxis = CDbl(axis)

        If numericAxis = 0# Then
            p_NormalizeNuniqueAxis = 0
        ElseIf numericAxis = 1# Then
            p_NormalizeNuniqueAxis = 1
        Else
            Err.Raise vbObjectError + 1231, _
                      "cTable.nunique", _
                      "axis must be 0 or 1."
        End If

        Exit Function
    End If

    Err.Raise vbObjectError + 1231, _
              "cTable.nunique", _
              "axis must be 0, 1, ""index"", or ""columns""."
End Function


' Produce a consistent string representation for a numeric value.
'
' CDec preserves Currency and normal decimal values more accurately than ' immediately converting everything to Double. Extremely large Double ' values fall back to Str$(CDbl(...)).
Private Function p_CanonicalNumericKey(ByVal v As Variant) As String
    On Error GoTo DoubleFallback

    p_CanonicalNumericKey = CStr(CDec(v))
    Exit Function

DoubleFallback:
    Err.Clear
    p_CanonicalNumericKey = Trim$(Str$(CDbl(v)))
End Function


' ---------------------------------------------------------------------------
' Create a type-aware dictionary key for one non-NA scalar value.
'
' Prefixes prevent values of different conceptual types from colliding:
'   "1" as String is different from 1 as a number.
'   True is different from numeric 1.
'
' String comparison remains case-insensitive because the distinct-value ' dictionaries use vbTextCompare, matching the rest of cTable.
' ---------------------------------------------------------------------------
Private Function p_NuniqueValueKey( _
        ByVal v As Variant, _
        ByVal declaredType As ctDataType) As String

    If p_IsNA(v) Then
        p_NuniqueValueKey = "NA:"
        Exit Function
    End If

    Select Case declaredType
        Case ctString
            p_NuniqueValueKey = "S:" & CStr(v)

        Case ctLong, ctDouble, ctCurrency
            p_NuniqueValueKey = _
                "N:" & p_CanonicalNumericKey(v)

        Case ctDate
            p_NuniqueValueKey = _
                "D:" & Trim$(Str$(CDbl(CDate(v))))

        Case ctBoolean
            If CBool(v) Then
                p_NuniqueValueKey = "B:1"
            Else
                p_NuniqueValueKey = "B:0"
            End If

        Case ctVariant
            ' Determine the key from the actual Variant subtype.
            If IsObject(v) Then
                Err.Raise vbObjectError + 1232, _
                          "cTable.nunique", _
                          "Object values cannot be counted by nunique."
            End If

            If (VarType(v) And vbArray) = vbArray Then
                Err.Raise vbObjectError + 1233, _
                          "cTable.nunique", _
                          "Array values cannot be counted by nunique."
            End If

            Select Case VarType(v)
                Case vbString
                    p_NuniqueValueKey = "S:" & CStr(v)

                Case vbBoolean
                    If CBool(v) Then
                        p_NuniqueValueKey = "B:1"
                    Else
                        p_NuniqueValueKey = "B:0"
                    End If

                Case vbDate
                    p_NuniqueValueKey = _
                        "D:" & Trim$(Str$(CDbl(CDate(v))))

                Case vbByte, vbInteger, vbLong, vbSingle, _
                     vbDouble, vbCurrency, vbDecimal

                    p_NuniqueValueKey = _
                        "N:" & p_CanonicalNumericKey(v)

                Case Else
                    On Error GoTo UnsupportedValue
                    p_NuniqueValueKey = _
                        "V:" & typeName(v) & ":" & CStr(v)
            End Select

        Case Else
            Err.Raise vbObjectError + 1234, _
                      "cTable.nunique", _
                      "Unsupported dtype: " & _
                      Me.dtype_name(declaredType)
    End Select

    Exit Function

UnsupportedValue:
    Err.Raise vbObjectError + 1235, _
              "cTable.nunique", _
              "Value of type """ & typeName(v) & _
              """ cannot be counted by nunique."
End Function
' ---------------------------------------------------------------------------
' Normalize the pandas-style drop_duplicates keep argument.
'
' Returns:
'    1 = keep first occurrence
'   -1 = keep last occurrence
'    0 = remove every member of a duplicate group
'
' Supported values:
'   "first"
'   "last"
'   False
' ---------------------------------------------------------------------------
Private Function p_NormalizeDropDuplicatesKeep( _
        ByVal keep As Variant) As Long

    If IsNull(keep) Then
        Err.Raise vbObjectError + 1240, _
                  "cTable.drop_duplicates", _
                  "keep cannot be Null."
    End If

    If IsEmpty(keep) Then
        p_NormalizeDropDuplicatesKeep = 1
        Exit Function
    End If

    If VarType(keep) = vbBoolean Then
        If CBool(keep) Then
            Err.Raise vbObjectError + 1241, _
                      "cTable.drop_duplicates", _
                      "keep=True is not valid. Use ""first"", ""last"", or False."
        End If

        p_NormalizeDropDuplicatesKeep = 0
        Exit Function
    End If

    If VarType(keep) = vbString Then
        Select Case LCase$(Trim$(CStr(keep)))
            Case "first"
                p_NormalizeDropDuplicatesKeep = 1

            Case "last"
                p_NormalizeDropDuplicatesKeep = -1

            Case Else
                Err.Raise vbObjectError + 1241, _
                          "cTable.drop_duplicates", _
                          "keep must be ""first"", ""last"", or False."
        End Select

        Exit Function
    End If

    Err.Raise vbObjectError + 1241, _
              "cTable.drop_duplicates", _
              "keep must be ""first"", ""last"", or False."
End Function


' ---------------------------------------------------------------------------
' Create a collision-resistant, dtype-aware key for selected columns in one ' row. Each scalar component is generated by p_NuniqueValueKey.
'
' The length prefix prevents combinations such as:
'   ("AB", "C")
' and
'   ("A", "BC")
'
' from producing the same composite key.
' ---------------------------------------------------------------------------
Private Function p_DuplicateRowKey( _
        ByVal rowNum As Long, _
        ByRef colIndexes() As Long, _
        ByVal selectedColCount As Long) As String

    Dim result As String
    Dim component As String
    Dim i As Long
    Dim ci As Long

    For i = 1 To selectedColCount
        ci = colIndexes(i)

        component = p_NuniqueValueKey( _
                        m_Data(rowNum, ci), _
                        m_ColTypes(ci))

        result = result & _
                 CStr(Len(component)) & _
                 ":" & component & "|"
    Next i

    p_DuplicateRowKey = result
End Function


' Count unique values in one column.
Private Function p_NuniqueColumn(ByVal colNum As Long, _
                                 ByVal dropna As Boolean) As Long
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare

    Dim r As Long
    Dim v As Variant
    Dim key As String

    For r = 1 To m_RowCount
        v = m_Data(r, colNum)

        If p_IsNA(v) Then
            If Not dropna Then
                ' All missing-value representations count as one category.
                If Not seen.Exists("NA:") Then
                    seen.Add "NA:", True
                End If
            End If
        Else
            key = p_NuniqueValueKey( _
                        v, _
                        m_ColTypes(colNum))

            If Not seen.Exists(key) Then
                seen.Add key, True
            End If
        End If
    Next r

    p_NuniqueColumn = seen.count
End Function


' Count unique values across one row.
Private Function p_NuniqueRow(ByVal rowNum As Long, _
                              ByVal dropna As Boolean) As Long
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare

    Dim c As Long
    Dim v As Variant
    Dim key As String

    For c = 1 To m_ColCount
        v = m_Data(rowNum, c)

        If p_IsNA(v) Then
            If Not dropna Then
                If Not seen.Exists("NA:") Then
                    seen.Add "NA:", True
                End If
            End If
        Else
            key = p_NuniqueValueKey( _
                        v, _
                        m_ColTypes(c))

            If Not seen.Exists(key) Then
                seen.Add key, True
            End If
        End If
    Next c

    p_NuniqueRow = seen.count
End Function

' Returns True when a Variant contains a nonempty array.
Private Function p_ArrayHasElements(ByVal values As Variant) As Boolean
    If Not IsArray(values) Then Exit Function

    Dim lowerBound As Long
    Dim upperBound As Long

    On Error GoTo EmptyArray

    lowerBound = LBound(values)
    upperBound = UBound(values)

    p_ArrayHasElements = (upperBound >= lowerBound)
    Exit Function

EmptyArray:
    p_ArrayHasElements = False
End Function


' Joins the contents of a one-dimensional Variant array.
Private Function p_JoinVariantArray( _
        ByVal values As Variant, _
        Optional ByVal delimiter As String = ", ") As String

    If Not p_ArrayHasElements(values) Then Exit Function

    Dim result As String
    Dim i As Long

    For i = LBound(values) To UBound(values)
        If Len(result) > 0 Then result = result & delimiter
        result = result & CStr(values(i))
    Next i

    p_JoinVariantArray = result
End Function

' Returns True when a dtype can be used by nlargest/nsmallest.
Private Function p_IsRankableDType(ByVal t As ctDataType) As Boolean
    Select Case t
        Case ctLong, ctDouble, ctCurrency, ctDate, ctBoolean
            p_IsRankableDType = True

        Case Else
            p_IsRankableDType = False
    End Select
End Function


' Convert a typed ranking value to Double for comparison.
Private Function p_RankValueAsDouble(ByVal rowNum As Long, _
                                     ByVal colNum As Long) As Double
    Dim v As Variant
    v = m_Data(rowNum, colNum)

    Select Case m_ColTypes(colNum)
        Case ctBoolean
            If CBool(v) Then
                p_RankValueAsDouble = 1#
            Else
                p_RankValueAsDouble = 0#
            End If

        Case ctDate
            p_RankValueAsDouble = CDbl(CDate(v))

        Case ctLong, ctDouble, ctCurrency
            p_RankValueAsDouble = CDbl(v)

        Case Else
            Err.Raise vbObjectError + 1200, _
                      "cTable ranking", _
                      "Column """ & m_Headers(colNum) & _
                      """ cannot be ranked."
    End Select
End Function


' Lexicographically compare two rows using the ranking columns.
'
' Returns:
'   -1 when rowA < rowB
'    0 when the ranking keys are equal
'    1 when rowA > rowB
Private Function p_CompareRankRows(ByVal rowA As Long, _
                                   ByVal rowB As Long, _
                                   ByRef colIndexes() As Long, _
                                   ByVal rankingColCount As Long) As Long
    Dim i As Long
    Dim colNum As Long
    Dim valueA As Double
    Dim valueB As Double

    For i = 1 To rankingColCount
        colNum = colIndexes(i)

        valueA = p_RankValueAsDouble(rowA, colNum)
        valueB = p_RankValueAsDouble(rowB, colNum)

        If valueA < valueB Then
            p_CompareRankRows = -1
            Exit Function
        ElseIf valueA > valueB Then
            p_CompareRankRows = 1
            Exit Function
        End If
    Next i

    p_CompareRankRows = 0
End Function


' Merge two sorted sections of a row-number array.
Private Sub p_MergeRankRows(ByRef rowNumbers() As Long, _
                            ByRef work() As Long, _
                            ByVal first As Long, _
                            ByVal middle As Long, _
                            ByVal last As Long, _
                            ByRef colIndexes() As Long, _
                            ByVal rankingColCount As Long, _
                            ByVal descending As Boolean)
    Dim leftPos As Long
    Dim rightPos As Long
    Dim outputPos As Long
    Dim comparison As Long
    Dim chooseLeft As Boolean

    leftPos = first
    rightPos = middle + 1
    outputPos = first

    Do While leftPos <= middle And rightPos <= last
        comparison = p_CompareRankRows( _
                        rowNumbers(leftPos), _
                        rowNumbers(rightPos), _
                        colIndexes, _
                        rankingColCount)

        If descending Then
            chooseLeft = (comparison >= 0)
        Else
            chooseLeft = (comparison <= 0)
        End If

        ' Selecting the left value when keys are equal makes the sort stable.
        If chooseLeft Then
            work(outputPos) = rowNumbers(leftPos)
            leftPos = leftPos + 1
        Else
            work(outputPos) = rowNumbers(rightPos)
            rightPos = rightPos + 1
        End If

        outputPos = outputPos + 1
    Loop

    Do While leftPos <= middle
        work(outputPos) = rowNumbers(leftPos)
        leftPos = leftPos + 1
        outputPos = outputPos + 1
    Loop

    Do While rightPos <= last
        work(outputPos) = rowNumbers(rightPos)
        rightPos = rightPos + 1
        outputPos = outputPos + 1
    Loop

    For outputPos = first To last
        rowNumbers(outputPos) = work(outputPos)
    Next outputPos
End Sub


' Stable merge sort for an array of source row numbers.
Private Sub p_MergeSortRankRows(ByRef rowNumbers() As Long, _
                                ByRef work() As Long, _
                                ByVal first As Long, _
                                ByVal last As Long, _
                                ByRef colIndexes() As Long, _
                                ByVal rankingColCount As Long, _
                                ByVal descending As Boolean)
    If first >= last Then Exit Sub

    Dim middle As Long
    middle = (first + last) \ 2

    p_MergeSortRankRows _
        rowNumbers, work, _
        first, middle, _
        colIndexes, rankingColCount, descending

    p_MergeSortRankRows _
        rowNumbers, work, _
        middle + 1, last, _
        colIndexes, rankingColCount, descending

    p_MergeRankRows _
        rowNumbers, work, _
        first, middle, last, _
        colIndexes, rankingColCount, descending
End Sub


' Shared implementation for nlargest and nsmallest.
Private Function p_RankedRows(ByVal n As Long, _
                              ByVal columns As Variant, _
                              ByVal largest As Boolean, _
                              ByVal keep As String) As cTable
    Dim result As cTable

    If m_ColCount = 0 Then
        Set result = New cTable
        Set p_RankedRows = result
        Exit Function
    End If

    If n <= 0 Or m_RowCount = 0 Then
        Dim emptyRows(1 To 1) As Long

        Set result = p_BuildFromRows(emptyRows, 0)
        Set result.DataSource = Me.DataSource
        result.set_dtypes_from Me

        Set p_RankedRows = result
        Exit Function
    End If

    p_EnsureTypes

    ' Accept either one column name or an array of column names.
    Dim columnList As Variant

    If IsArray(columns) Then
        columnList = columns
    Else
        columnList = Array(columns)
    End If

    Dim rankingColCount As Long
    rankingColCount = UBound(columnList) - LBound(columnList) + 1

    If rankingColCount <= 0 Then
        Err.Raise vbObjectError + 1201, _
                  "cTable ranking", _
                  "At least one ranking column is required."
    End If

    Dim colIndexes() As Long
    ReDim colIndexes(1 To rankingColCount)

    Dim i As Long
    Dim colName As String
    Dim colNum As Long

    For i = 1 To rankingColCount
        colName = p_LabelColumnRef(columnList(LBound(columnList) + i - 1))
        colNum = p_ColIdx(colName)

        If colNum = 0 Then
            Err.Raise vbObjectError + 1202, _
                      "cTable ranking", _
                      "Column not found: """ & colName & """"
        End If

        If Not p_IsRankableDType(m_ColTypes(colNum)) Then
            Err.Raise vbObjectError + 1203, _
                      "cTable ranking", _
                      "Column """ & colName & """ has dtype " & _
                      Me.dtype_name(m_ColTypes(colNum)) & _
                      " and cannot be used for ranking."
        End If

        colIndexes(i) = colNum
    Next i

    Dim keepMode As String
    keepMode = LCase$(Trim$(keep))

    Select Case keepMode
        Case "first", "last", "all"
            ' Valid.

        Case Else
            Err.Raise vbObjectError + 1204, _
                      "cTable ranking", _
                      "keep must be ""first"", ""last"", or ""all""."
    End Select

    ' Collect rows having non-NA values in every ranking column.
    Dim candidateRows() As Long
    ReDim candidateRows(1 To m_RowCount)

    Dim candidateCount As Long
    Dim r As Long
    Dim includeRow As Boolean

    If keepMode = "last" Then
        ' Reverse initial order. Because the merge sort is stable, rows with
        ' identical ranking keys then retain last-occurrence-first order.
        For r = m_RowCount To 1 Step -1
            includeRow = True

            For i = 1 To rankingColCount
                If p_IsNA(m_Data(r, colIndexes(i))) Then
                    includeRow = False
                    Exit For
                End If
            Next i

            If includeRow Then
                candidateCount = candidateCount + 1
                candidateRows(candidateCount) = r
            End If
        Next r
    Else
        For r = 1 To m_RowCount
            includeRow = True

            For i = 1 To rankingColCount
                If p_IsNA(m_Data(r, colIndexes(i))) Then
                    includeRow = False
                    Exit For
                End If
            Next i

            If includeRow Then
                candidateCount = candidateCount + 1
                candidateRows(candidateCount) = r
            End If
        Next r
    End If

    If candidateCount = 0 Then
        Set result = p_BuildFromRows(candidateRows, 0)
        Set result.DataSource = Me.DataSource
        result.set_dtypes_from Me

        Set p_RankedRows = result
        Exit Function
    End If

    Dim work() As Long
    ReDim work(1 To candidateCount)

    p_MergeSortRankRows _
        candidateRows, _
        work, _
        1, _
        candidateCount, _
        colIndexes, _
        rankingColCount, _
        largest

    Dim resultCount As Long

    If n < candidateCount Then
        resultCount = n
    Else
        resultCount = candidateCount
    End If

    ' keep="all" includes every row tied with the nth row.
    If keepMode = "all" _
       And resultCount > 0 _
       And resultCount < candidateCount Then

        Do While resultCount < candidateCount
            If p_CompareRankRows( _
                    candidateRows(resultCount), _
                    candidateRows(resultCount + 1), _
                    colIndexes, _
                    rankingColCount) <> 0 Then

                Exit Do
            End If

            resultCount = resultCount + 1
        Loop
    End If

    Set result = p_BuildFromRows(candidateRows, resultCount)
    Set result.DataSource = Me.DataSource
    result.set_dtypes_from Me

    Set p_RankedRows = result
End Function

' ---------------------------------------------------------------------------
' Returns True for dtypes treated as numeric by aggregation methods.
'
' Boolean values are treated pandas-style:
'   True  = 1
'   False = 0
' ---------------------------------------------------------------------------
Private Function p_IsNumericDType(ByVal t As ctDataType) As Boolean
    Select Case t
        Case ctLong, ctDouble, ctCurrency, ctBoolean
            p_IsNumericDType = True

        Case Else
            p_IsNumericDType = False
    End Select
End Function


' True when a supplied Variant means "no particular column".
Private Function p_IsBlankColumnArgument(ByVal v As Variant) As Boolean
    If IsNull(v) Or IsEmpty(v) Then
        p_IsBlankColumnArgument = True
    Else
        p_IsBlankColumnArgument = (Len(Trim$(CStr(v))) = 0)
    End If
End Function


' Returns whether an aggregation supports a particular dtype when ' numeric_only=False.
Private Function p_AggregationSupportsType(ByVal operation As String, _
                                           ByVal t As ctDataType) As Boolean
    Select Case LCase$(operation)
        Case "count"
            ' count operates on every dtype, including ctVariant.
            p_AggregationSupportsType = True

        Case "min", "max"
            ' Ordered scalar dtypes are supported.
            p_AggregationSupportsType = _
                (t = ctString Or _
                 t = ctLong Or _
                 t = ctDouble Or _
                 t = ctCurrency Or _
                 t = ctDate Or _
                 t = ctBoolean)

        Case "mean", "median", "quantile"
            ' Dates can be averaged or interpolated using their VBA
            ' serial-number representation.
            p_AggregationSupportsType = _
                p_IsNumericDType(t) Or t = ctDate

        Case "var", "std"
            ' A date standard deviation would require a timedelta dtype,
            ' which cTable does not currently have.
            p_AggregationSupportsType = p_IsNumericDType(t)

        Case Else
            p_AggregationSupportsType = False
    End Select
End Function


' ---------------------------------------------------------------------------
' Collect a column into a Double array for numerical calculations.
'
' Returns:
'   True  = collection completed
'   False = skipna=False and an NA value was encountered
'
' validCount receives the number of non-NA values collected.
' ---------------------------------------------------------------------------
Private Function p_CollectNumericValues(ByVal ci As Long, _
                                        ByVal skipna As Boolean, _
                                        ByVal allowDate As Boolean, _
                                        ByRef values() As Double, _
                                        ByRef validCount As Long) As Boolean
    If ci < 1 Or ci > m_ColCount Then
        Err.Raise vbObjectError + 930, _
                  "cTable.p_CollectNumericValues", _
                  "Column number out of bounds: " & ci
    End If

    p_EnsureTypes

    Dim colType As ctDataType
    colType = m_ColTypes(ci)

    If Not p_IsNumericDType(colType) Then
        If Not (allowDate And colType = ctDate) Then
            Err.Raise vbObjectError + 931, _
                      "cTable aggregation", _
                      "Column """ & m_Headers(ci) & """ has dtype " & _
                      Me.dtype_name(colType) & _
                      " and cannot be used for this aggregation."
        End If
    End If

    ReDim values(1 To IIf(m_RowCount > 0, m_RowCount, 1))
    validCount = 0

    Dim r As Long
    Dim v As Variant
    Dim converted As Variant
    Dim ok As Boolean
    Dim numericValue As Double

    For r = 1 To m_RowCount
        v = m_Data(r, ci)

        If p_IsNA(v) Then
            If Not skipna Then
                p_CollectNumericValues = False
                Exit Function
            End If
        Else
            Select Case colType
                Case ctBoolean
                    converted = p_TryConvert(v, ctBoolean, ok)

                    If Not ok Then
                        Err.Raise vbObjectError + 932, _
                                  "cTable aggregation", _
                                  "Invalid Boolean value in column """ & _
                                  m_Headers(ci) & """ at row " & r & "."
                    End If

                    If CBool(converted) Then
                        numericValue = 1#
                    Else
                        numericValue = 0#
                    End If

                Case ctLong, ctDouble, ctCurrency
                    ' VBA converts True to -1. Handle Boolean source values
                    ' explicitly so aggregate behavior uses True=1.
                    If VarType(v) = vbBoolean Then
                        If CBool(v) Then
                            numericValue = 1#
                        Else
                            numericValue = 0#
                        End If
                    Else
                        converted = p_TryConvert(v, colType, ok)

                        If Not ok Then
                            Err.Raise vbObjectError + 933, _
                                      "cTable aggregation", _
                                      "Invalid numeric value in column """ & _
                                      m_Headers(ci) & """ at row " & r & "."
                        End If

                        numericValue = CDbl(converted)
                    End If

                Case ctDate
                    converted = p_TryConvert(v, ctDate, ok)

                    If Not ok Then
                        Err.Raise vbObjectError + 934, _
                                  "cTable aggregation", _
                                  "Invalid date value in column """ & _
                                  m_Headers(ci) & """ at row " & r & "."
                    End If

                    numericValue = CDbl(CDate(converted))

                Case Else
                    Err.Raise vbObjectError + 935, _
                              "cTable aggregation", _
                              "Unsupported dtype for column """ & _
                              m_Headers(ci) & """."
            End Select

            validCount = validCount + 1
            values(validCount) = numericValue
        End If
    Next r

    p_CollectNumericValues = True
End Function


' In-place QuickSort for a 1-based Double array.
Private Sub p_QuickSortDoubles(ByRef values() As Double, _
                               ByVal first As Long, _
                               ByVal last As Long)
    If first >= last Then Exit Sub

    Dim low As Long
    Dim high As Long
    Dim pivot As Double
    Dim temp As Double

    low = first
    high = last
    pivot = values((first + last) \ 2)

    Do While low <= high
        Do While values(low) < pivot
            low = low + 1
        Loop

        Do While values(high) > pivot
            high = high - 1
        Loop

        If low <= high Then
            temp = values(low)
            values(low) = values(high)
            values(high) = temp

            low = low + 1
            high = high - 1
        End If
    Loop

    If first < high Then p_QuickSortDoubles values, first, high
    If low < last Then p_QuickSortDoubles values, low, last
End Sub


' Calculate a quantile from an already-sorted 1-based Double array.
Private Function p_QuantileFromSorted(ByRef values() As Double, _
                                      ByVal valueCount As Long, _
                                      ByVal q As Double, _
                                      ByVal interpolation As String) As Double
    If q < 0# Or q > 1# Then
        Err.Raise vbObjectError + 936, _
                  "cTable.quantile", _
                  "q must be between 0 and 1."
    End If

    If valueCount <= 0 Then
        Err.Raise vbObjectError + 937, _
                  "cTable.quantile", _
                  "Cannot calculate a quantile from an empty array."
    End If

    If valueCount = 1 Then
        p_QuantileFromSorted = values(1)
        Exit Function
    End If

    ' Zero-based fractional location.
    Dim position As Double
    position = (valueCount - 1) * q

    Dim lowerBase As Long
    Dim lowerIndex As Long
    Dim upperIndex As Long
    Dim fraction As Double

    lowerBase = Fix(position)
    lowerIndex = lowerBase + 1
    upperIndex = lowerIndex
    fraction = position - lowerBase

    If fraction > 0# Then upperIndex = lowerIndex + 1

    Select Case LCase$(Trim$(interpolation))
        Case "linear", ""
            p_QuantileFromSorted = _
                values(lowerIndex) + _
                fraction * (values(upperIndex) - values(lowerIndex))

        Case "lower"
            p_QuantileFromSorted = values(lowerIndex)

        Case "higher"
            p_QuantileFromSorted = values(upperIndex)

        Case "midpoint"
            p_QuantileFromSorted = _
                (values(lowerIndex) + values(upperIndex)) / 2#

        Case "nearest"
            Dim nearestIndex As Long

            ' VBA Round uses banker's rounding, which is suitable for choosing
            ' between equally near positions.
            nearestIndex = CLng(Round(position, 0)) + 1

            If nearestIndex < 1 Then nearestIndex = 1
            If nearestIndex > valueCount Then nearestIndex = valueCount

            p_QuantileFromSorted = values(nearestIndex)

        Case Else
            Err.Raise vbObjectError + 938, _
                      "cTable.quantile", _
                      "Unknown interpolation method: """ & _
                      interpolation & """."
    End Select
End Function


' Quantile or median for one column.
Private Function p_QuantileColumn(ByVal ci As Long, _
                                  ByVal q As Double, _
                                  ByVal skipna As Boolean, _
                                  ByVal interpolation As String) As Variant
    Dim values() As Double
    Dim validCount As Long

    If Not p_CollectNumericValues( _
            ci, skipna, True, values, validCount) Then

        p_QuantileColumn = Null
        Exit Function
    End If

    If validCount = 0 Then
        p_QuantileColumn = Null
        Exit Function
    End If

    p_QuickSortDoubles values, 1, validCount

    Dim result As Double
    result = p_QuantileFromSorted( _
                values, validCount, q, interpolation)

    If m_ColTypes(ci) = ctDate Then
        p_QuantileColumn = CDate(result)
    Else
        p_QuantileColumn = result
    End If
End Function


' Arithmetic mean for one column.
Private Function p_MeanColumn(ByVal ci As Long, _
                              ByVal skipna As Boolean) As Variant
    Dim values() As Double
    Dim validCount As Long

    If Not p_CollectNumericValues( _
            ci, skipna, True, values, validCount) Then

        p_MeanColumn = Null
        Exit Function
    End If

    If validCount = 0 Then
        p_MeanColumn = Null
        Exit Function
    End If

    Dim total As Double
    Dim i As Long

    For i = 1 To validCount
        total = total + values(i)
    Next i

    If m_ColTypes(ci) = ctDate Then
        p_MeanColumn = CDate(total / validCount)
    Else
        p_MeanColumn = total / validCount
    End If
End Function


' Variance for one column using Welford's numerically stable algorithm.
Private Function p_VarianceColumn(ByVal ci As Long, _
                                  ByVal skipna As Boolean, _
                                  ByVal ddof As Long) As Variant
    If ddof < 0 Then
        Err.Raise vbObjectError + 939, _
                  "cTable.var", _
                  "ddof cannot be negative."
    End If

    Dim values() As Double
    Dim validCount As Long

    If Not p_CollectNumericValues( _
            ci, skipna, False, values, validCount) Then

        p_VarianceColumn = Null
        Exit Function
    End If

    If validCount <= ddof Then
        p_VarianceColumn = Null
        Exit Function
    End If

    Dim runningMean As Double
    Dim m2 As Double
    Dim delta As Double
    Dim i As Long

    For i = 1 To validCount
        delta = values(i) - runningMean
        runningMean = runningMean + delta / i
        m2 = m2 + delta * (values(i) - runningMean)
    Next i

    Dim result As Double
    result = m2 / (validCount - ddof)

    ' Protect against tiny negative floating-point artifacts.
    If result < 0# And result > -0.000000000001 Then result = 0#

    p_VarianceColumn = result
End Function


' Minimum or maximum for one ordered column.
Private Function p_ExtremeColumn(ByVal ci As Long, _
                                 ByVal findMaximum As Boolean, _
                                 ByVal skipna As Boolean) As Variant
    p_EnsureTypes

    Dim colType As ctDataType
    colType = m_ColTypes(ci)

    If Not p_AggregationSupportsType( _
            IIf(findMaximum, "max", "min"), colType) Then

        Err.Raise vbObjectError + 940, _
                  "cTable aggregation", _
                  "Column """ & m_Headers(ci) & """ has dtype " & _
                  Me.dtype_name(colType) & _
                  " and cannot be ordered."
    End If

    Dim hasValue As Boolean
    Dim bestValue As Variant
    Dim currentValue As Variant
    Dim converted As Variant
    Dim ok As Boolean
    Dim comparison As Long

    Dim r As Long
    For r = 1 To m_RowCount
        currentValue = m_Data(r, ci)

        If p_IsNA(currentValue) Then
            If Not skipna Then
                p_ExtremeColumn = Null
                Exit Function
            End If
        Else
            converted = p_TryConvert(currentValue, colType, ok)

            If Not ok Then
                Err.Raise vbObjectError + 941, _
                          "cTable aggregation", _
                          "Invalid value in column """ & _
                          m_Headers(ci) & """ at row " & r & "."
            End If

            If Not hasValue Then
                bestValue = converted
                hasValue = True
            Else
                Select Case colType
                    Case ctString
                        comparison = StrComp( _
                            CStr(converted), _
                            CStr(bestValue), _
                            vbTextCompare)

                    Case ctBoolean
                        Dim currentBoolNumber As Long
                        Dim bestBoolNumber As Long

                        If CBool(converted) Then currentBoolNumber = 1
                        If CBool(bestValue) Then bestBoolNumber = 1

                        If currentBoolNumber < bestBoolNumber Then
                            comparison = -1
                        ElseIf currentBoolNumber > bestBoolNumber Then
                            comparison = 1
                        Else
                            comparison = 0
                        End If

                    Case Else
                        If CDbl(converted) < CDbl(bestValue) Then
                            comparison = -1
                        ElseIf CDbl(converted) > CDbl(bestValue) Then
                            comparison = 1
                        Else
                            comparison = 0
                        End If
                End Select

                If findMaximum Then
                    If comparison > 0 Then bestValue = converted
                Else
                    If comparison < 0 Then bestValue = converted
                End If
            End If
        End If
    Next r

    If hasValue Then
        p_ExtremeColumn = bestValue
    Else
        p_ExtremeColumn = Null
    End If
End Function


' Dispatch one column to the requested aggregation.
Private Function p_AggregateColumnIndex(ByVal operation As String, _
                                        ByVal ci As Long, _
                                        ByVal skipna As Boolean, _
                                        ByVal q As Double, _
                                        ByVal interpolation As String, _
                                        ByVal ddof As Long) As Variant
    Select Case LCase$(operation)
        Case "count"
            p_AggregateColumnIndex = _
                m_RowCount - p_ColNACount(ci)

        Case "mean"
            p_AggregateColumnIndex = _
                p_MeanColumn(ci, skipna)

        Case "median"
            p_AggregateColumnIndex = _
                p_QuantileColumn(ci, 0.5, skipna, "linear")

        Case "min"
            p_AggregateColumnIndex = _
                p_ExtremeColumn(ci, False, skipna)

        Case "max"
            p_AggregateColumnIndex = _
                p_ExtremeColumn(ci, True, skipna)

        Case "var"
            p_AggregateColumnIndex = _
                p_VarianceColumn(ci, skipna, ddof)

        Case "std"
            Dim varianceResult As Variant
            varianceResult = p_VarianceColumn(ci, skipna, ddof)

            If IsNull(varianceResult) Then
                p_AggregateColumnIndex = Null
            Else
                p_AggregateColumnIndex = Sqr(CDbl(varianceResult))
            End If

        Case "quantile"
            ' pandas-style quantile always excludes NA values.
            p_AggregateColumnIndex = _
                p_QuantileColumn(ci, q, True, interpolation)

        Case Else
            Err.Raise vbObjectError + 942, _
                      "cTable aggregation", _
                      "Unknown aggregation: """ & operation & """"
    End Select
End Function


' Aggregate one explicitly named column.
Private Function p_AggregateNamedColumn(ByVal operation As String, _
                                        ByVal colName As String, _
                                        ByVal skipna As Boolean, _
                                        ByVal numeric_only As Boolean, _
                                        ByVal q As Double, _
                                        ByVal interpolation As String, _
                                        ByVal ddof As Long) As Variant
    Dim ci As Long
    ci = p_ColIdx(colName)

    If ci = 0 Then
        Err.Raise vbObjectError + 943, _
                  "cTable." & operation, _
                  "Column not found: """ & colName & """"
    End If

    p_EnsureTypes

    If numeric_only And Not p_IsNumericDType(m_ColTypes(ci)) Then
        Err.Raise vbObjectError + 944, _
                  "cTable." & operation, _
                  "Column """ & colName & """ has dtype " & _
                  Me.dtype_name(m_ColTypes(ci)) & _
                  " and is excluded by numeric_only=True."
    End If

    If Not p_AggregationSupportsType(operation, m_ColTypes(ci)) Then
        Err.Raise vbObjectError + 945, _
                  "cTable." & operation, _
                  "Column """ & colName & """ has dtype " & _
                  Me.dtype_name(m_ColTypes(ci)) & _
                  " and does not support " & operation & "."
    End If

    p_AggregateNamedColumn = _
        p_AggregateColumnIndex( _
            operation, ci, skipna, q, interpolation, ddof)
End Function


' Aggregate all applicable columns and return column name -> result.
Private Function p_AggregateAllColumns(ByVal operation As String, _
                                       ByVal skipna As Boolean, _
                                       ByVal numeric_only As Boolean, _
                                       ByVal q As Double, _
                                       ByVal interpolation As String, _
                                       ByVal ddof As Long) As Object
    Dim results As Object
    Set results = CreateObject("Scripting.Dictionary")
    results.CompareMode = vbTextCompare

    p_EnsureTypes

    Dim c As Long
    Dim includeColumn As Boolean

    For c = 1 To m_ColCount
        If numeric_only Then
            includeColumn = p_IsNumericDType(m_ColTypes(c))
        Else
            includeColumn = _
                p_AggregationSupportsType(operation, m_ColTypes(c))

            If Not includeColumn Then
                Err.Raise vbObjectError + 946, _
                          "cTable." & operation, _
                          "Column """ & m_Headers(c) & """ has dtype " & _
                          Me.dtype_name(m_ColTypes(c)) & _
                          " and does not support " & operation & _
                          ". Use numeric_only:=True to exclude it."
            End If
        End If

        If includeColumn Then
            results.Add m_Headers(c), _
                p_AggregateColumnIndex( _
                    operation, _
                    c, _
                    skipna, _
                    q, _
                    interpolation, _
                    ddof)
        End If
    Next c

    Set p_AggregateAllColumns = results
End Function



' Sum one numeric column.
'
' pandas-style behavior:
'   skipna=True
'       Ignore NA values.
'
'   skipna=False
'       Return Null if any NA value is encountered.
'
'   minCount=0
'       An empty or all-NA column returns zero when skipna=True.
'
'   minCount>0
'       Return Null if fewer than minCount valid values exist.
'
' Return type:
'   ctLong       -> Long when possible; Double if the total exceeds Long
'   ctDouble     -> Double
'   ctCurrency   -> Currency
'   ctBoolean    -> Long count of True values
Private Function p_SumNumericColumn(ByVal ci As Long, _
                                    ByVal skipna As Boolean, _
                                    ByVal minCount As Long) As Variant

    If ci < 1 Or ci > m_ColCount Then
        Err.Raise vbObjectError + 921, _
                  "cTable.p_SumNumericColumn", _
                  "Column number out of bounds: " & ci
    End If

    If minCount < 0 Then
        Err.Raise vbObjectError + 922, _
                  "cTable.sum", _
                  "min_count cannot be negative."
    End If

    p_EnsureTypes

    Dim colType As ctDataType
    colType = m_ColTypes(ci)

    If Not p_IsNumericDType(colType) Then
        Err.Raise vbObjectError + 923, _
                  "cTable.sum", _
                  "Column """ & m_Headers(ci) & """ has dtype " & _
                  Me.dtype_name(colType) & " and cannot be summed."
    End If

    Dim totalDouble As Double
    Dim totalCurrency As Currency
    Dim validCount As Long

    Dim r As Long
    Dim v As Variant
    Dim converted As Variant
    Dim ok As Boolean

    For r = 1 To m_RowCount
        v = m_Data(r, ci)

        If p_IsNA(v) Then
            If Not skipna Then
                p_SumNumericColumn = Null
                Exit Function
            End If
        Else
            validCount = validCount + 1

            Select Case colType
                Case ctBoolean
                    converted = p_TryConvert(v, ctBoolean, ok)

                    If Not ok Then
                        Err.Raise vbObjectError + 924, _
                                  "cTable.sum", _
                                  "Invalid Boolean value in column """ & _
                                  m_Headers(ci) & """ at row " & r & "."
                    End If

                    If CBool(converted) Then
                        totalDouble = totalDouble + 1#
                    End If

                Case ctCurrency
                    ' VBA normally treats True numerically as -1.
                    ' For pandas-style aggregation, explicitly use 1 or 0.
                    If VarType(v) = vbBoolean Then
                        If CBool(v) Then
                            totalCurrency = totalCurrency + CCur(1)
                        End If
                    Else
                        converted = p_TryConvert(v, ctCurrency, ok)

                        If Not ok Then
                            Err.Raise vbObjectError + 925, _
                                      "cTable.sum", _
                                      "Non-numeric value in column """ & _
                                      m_Headers(ci) & """ at row " & r & "."
                        End If

                        totalCurrency = totalCurrency + CCur(converted)
                    End If

                Case ctLong, ctDouble
                    If VarType(v) = vbBoolean Then
                        If CBool(v) Then
                            totalDouble = totalDouble + 1#
                        End If
                    Else
                        converted = p_TryConvert(v, colType, ok)

                        If Not ok Then
                            Err.Raise vbObjectError + 925, _
                                      "cTable.sum", _
                                      "Non-numeric value in column """ & _
                                      m_Headers(ci) & """ at row " & r & "."
                        End If

                        totalDouble = totalDouble + CDbl(converted)
                    End If
            End Select
        End If
    Next r

    ' min_count is based on valid, non-NA values.
    If validCount < minCount Then
        p_SumNumericColumn = Null
        Exit Function
    End If

    Select Case colType
        Case ctCurrency
            p_SumNumericColumn = totalCurrency

        Case ctBoolean
            p_SumNumericColumn = CLng(totalDouble)

        Case ctLong
            ' Preserve Long where possible, but avoid overflow when a
            ' column of Long values has a total beyond the Long range.
            If totalDouble >= -2147483648# _
               And totalDouble <= 2147483647# _
               And totalDouble = Fix(totalDouble) Then

                p_SumNumericColumn = CLng(totalDouble)
            Else
                p_SumNumericColumn = totalDouble
            End If

        Case ctDouble
            p_SumNumericColumn = totalDouble
    End Select
End Function

' Returns 1-based column index for colName; 0 if not found.
Private Function p_ColIdx(ByVal colName As String) As Long
    Dim i As Long
    For i = 1 To m_ColCount
        If StrComp(m_Headers(i), colName, vbTextCompare) = 0 Then
            p_ColIdx = i
            Exit Function
        End If
    Next i
    p_ColIdx = 0
End Function

 ' ---------------------------------------------------------------------------
' COLUMN REFERENCES (2026-10-07)
' Names remain strings (including "2026"); integral numeric Variants are
' 1-based positions. Lookup mode returns 0 for missing names/out-of-range.
' Reject unsupported argument types in either mode.
' ---------------------------------------------------------------------------
Private Function p_ResolveCol(ByVal col As Variant, _
                             Optional ByVal strict As Boolean = True) As Long
    Dim ci As Long
    Dim kind As VbVarType
    Dim n As Double
    If IsObject(col) Or IsArray(col) Then GoTo BadType
    If IsNull(col) Or IsEmpty(col) Or IsError(col) Then GoTo BadType
    kind = VarType(col)
    Select Case kind
        Case vbString
            ci = p_ColIdx(CStr(col))
        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
            n = CDbl(col)
            If n <> Fix(n) Then GoTo BadIndex
            If n >= 1# And n <= CDbl(m_ColCount) Then ci = CLng(n)
        Case Else
            GoTo BadType
    End Select
    If ci = 0 And strict Then
        Err.Raise vbObjectError + 1701, "cTable.p_ResolveCol", _
                  "Column not found or index out of range: " & CStr(col)
    End If
    p_ResolveCol = ci
    Exit Function
BadIndex:
    Err.Raise vbObjectError + 1702, "cTable.p_ResolveCol", _
              "Column index must be a whole number: " & CStr(col)
BadType:
    Err.Raise vbObjectError + 1703, "cTable.p_ResolveCol", _
              "Column reference must be a name or integer index."
End Function

' Preserve original spelling for existing string callers (including errors).
' For numeric callers supply the real header for names, generated destinations,
' diagnostic output and calculated-lineage fingerprints.
Private Function p_LabelColumnRef(ByVal col As Variant) As String
    If Not IsObject(col) And Not IsArray(col) Then
        If VarType(col) = vbString Then
            p_LabelColumnRef = CStr(col)
            Exit Function
        End If
    End If
    p_LabelColumnRef = m_Headers(p_ResolveCol(col))
End Function

Private Function p_HasOptionalCol(ByVal col As Variant) As Boolean
    If p_ArgIsMissing(col) Then Exit Function
    If IsNull(col) Or IsEmpty(col) Then Exit Function
    If VarType(col) = vbString Then
        If Len(Trim$(CStr(col))) = 0 Then Exit Function
    End If
    p_HasOptionalCol = True
End Function

' Sanitise header array: replace blank or duplicate entries with "Column#".
Private Sub p_SanitizeHeaders(ByRef hdrs() As String, ByVal colCount As Long)
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    Dim i As Long
    For i = 1 To colCount
        Dim h As String
        h = Trim(hdrs(i))
        If h = "" Or seen.Exists(h) Then
            h = "Column" & i
        End If
        hdrs(i) = h
        seen(h) = 1
    Next i
    Set seen = Nothing
End Sub

' Reset the index dictionary.
Private Sub p_ClearIndexes()
    m_Version = m_Version + 1

    Set m_Indexes = CreateObject("Scripting.Dictionary")
    m_Indexes.CompareMode = vbTextCompare

    If Not m_History Is Nothing Then
        m_History.x_MarkStale m_Version
    End If
End Sub

' Open or reuse an already-open workbook; returns Nothing on failure.
Private Function p_GetWorkbook(ByVal filePath As String, _
                                ByRef didOpen As Boolean) As Workbook
    didOpen = False
    Dim fName As String
    fName = Dir(filePath)
    Dim wb As Workbook
    On Error Resume Next
    Set wb = Workbooks(fName)
    On Error GoTo 0
    If wb Is Nothing Then
        On Error Resume Next
        Set wb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
        On Error GoTo 0
        didOpen = True
    End If
    Set p_GetWorkbook = wb
End Function
Private Sub p_ValidateNRows(ByVal nrows As Long, _
                            ByVal callerName As String)

    If nrows < 0 Then
        Err.Raise vbObjectError + 1650, callerName, _
                  "nrows cannot be negative."
    End If

End Sub


' ---------------------------------------------------------------------------
' p_RangeHasMerge
'   True when any cell in rng participates in a merged block.
'
'   Range.MergeCells returns False when nothing is merged, True when every
'   cell is merged, and Null when the range is mixed. Null is the common case
'   for a real sheet, so anything that is not an explicit False is treated as
'   "scan this range".
' ---------------------------------------------------------------------------
Private Function p_RangeHasMerge(ByVal Rng As Excel.Range) As Boolean

    If Rng Is Nothing Then Exit Function

    Dim v As Variant

    On Error Resume Next
    v = Rng.MergeCells
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        ' Unknown. Scanning is the safe answer.
        p_RangeHasMerge = True
        Exit Function
    End If
    On Error GoTo 0

    If VarType(v) = vbBoolean Then
        p_RangeHasMerge = CBool(v)
    Else
        p_RangeHasMerge = True
    End If

End Function

' ---------------------------------------------------------------------------
' p_FillMergedValues
'   Propagates each merged block's value across every cell of that block, in
'   the already-materialized value array.
'
'   Excel stores a merged block's value only in its top-left anchor cell.
'   Every other cell in the block reads back as Empty. A sheet that merges
'   repeated labels for appearance therefore imports as one populated row
'   followed by a run of blanks. This restores the anchor value to every row
'   and column of the block so the cTable column is fully populated.
'
'   Handles vertical merges, horizontal merges, and rectangular blocks.
'
'   raw
'       A 1-based 2-D array whose (1,1) element corresponds to rng.Cells(1,1),
'       which is what Range.Value and Range.Value2 return. Modified in place.
'
'   useValue2
'       False reads the anchor with .Value, True with .Value2. Pass whichever
'       property produced raw, so dates and currency stay consistent with the
'       unmerged cells around them.
'
'   If a block starts above or to the left of rng, the anchor value is still
'   read from the true anchor outside rng and applied to the portion of the
'   block that falls inside rng.
'
'   Cost: one MergeCells test per column, plus one probe per merged block in
'   the columns that have any. A range with no merged cells costs a single
'   test on the whole range.
' ---------------------------------------------------------------------------
Private Sub p_FillMergedValues(ByVal Rng As Excel.Range, _
                               ByRef raw As Variant, _
                               Optional ByVal useValue2 As Boolean = False)

    If Rng Is Nothing Then Exit Sub
    If Not IsArray(raw) Then Exit Sub

    ' Whole-range short circuit. Most sheets have no merged cells at all and
    ' pay only this one test.
    If Not p_RangeHasMerge(Rng) Then Exit Sub

    Dim rawRows As Long, rawCols As Long

    On Error Resume Next
    rawRows = UBound(raw, 1)
    rawCols = UBound(raw, 2)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    If rawRows < 1 Or rawCols < 1 Then Exit Sub

    Dim rngTop As Long, rngLeft As Long
    rngTop = Rng.Row
    rngLeft = Rng.Column

    Dim blocksFilled As Long

    Dim r As Long, c As Long
    Dim rr As Long, cc As Long
    Dim r1 As Long, r2 As Long, c1 As Long, c2 As Long
    Dim cel As Excel.Range
    Dim ma As Excel.Range
    Dim v As Variant

    For c = 1 To rawCols

        ' Skip whole columns that contain no merged cells. This keeps the
        ' per-cell walk off large blank regions of a UsedRange.
        If p_RangeHasMerge(Rng.columns(c)) Then

            r = 1

            Do While r <= rawRows

                ' Only a non-anchor cell of a merged block reads back Empty,
                ' so an already-populated cell never needs probing.
                If IsEmpty(raw(r, c)) Then

                    Set cel = Rng.Cells(r, c)

                    If cel.MergeCells Then

                        Set ma = cel.MergeArea

                        If useValue2 Then
                            v = ma.Cells(1, 1).Value2
                        Else
                            v = ma.Cells(1, 1).value
                        End If

                        ' Block bounds expressed as raw-array subscripts.
                        r1 = ma.Row - rngTop + 1
                        c1 = ma.Column - rngLeft + 1
                        r2 = r1 + ma.rows.count - 1
                        c2 = c1 + ma.columns.count - 1

                        ' Clip to the part of the block inside rng.
                        If r1 < 1 Then r1 = 1
                        If c1 < 1 Then c1 = 1
                        If r2 > rawRows Then r2 = rawRows
                        If c2 > rawCols Then c2 = rawCols

                        For rr = r1 To r2
                            For cc = c1 To c2
                                raw(rr, cc) = v
                            Next cc
                        Next rr

                        blocksFilled = blocksFilled + 1

                        ' r is inside the block, so r2 >= r and this always
                        ' advances.
                        r = r2 + 1

                    Else
                        r = r + 1
                    End If

                Else
                    r = r + 1
                End If

            Loop

        End If

    Next c

    Set ma = Nothing
    Set cel = Nothing

    If blocksFilled > 0 Then
        If m_Log Is Nothing Then Set m_Log = New cLog
        m_Log.AddStatus "Unmerge: filled " & blocksFilled & _
                        " merged block(s) from their anchor values."
    End If

End Sub

Private Function p_LimitReadRange(ByVal Rng As Excel.Range, _
                                  ByVal hasHeaders As Boolean, _
                                  ByVal nrows As Long) As Excel.Range

    If Rng Is Nothing Then Exit Function

    If nrows = 0 Then
        Set p_LimitReadRange = Rng
        Exit Function
    End If

    ' If the requested data-row count already exceeds the physical source
    ' row count, the full source range is necessarily sufficient. Testing
    ' this first also avoids Long overflow when hasHeaders=True.
    If nrows >= Rng.rows.count Then
        Set p_LimitReadRange = Rng
        Exit Function
    End If

    Dim wantedRows As Long
    wantedRows = nrows
    If hasHeaders Then wantedRows = wantedRows + 1

    Set p_LimitReadRange = _
        Rng.Resize(wantedRows, Rng.columns.count)

End Function

' Core range loader - reads rng into m_Data / m_Headers.
' Unmerge=True back-fills merged blocks from their anchor cell before the
' header/data split, so vertically merged label columns arrive fully populated.
Private Sub p_LoadFromRange(ByVal Rng As Range, ByVal hasHeaders As Boolean, _
                            Optional ByVal Unmerge As Boolean = False)
    If Rng Is Nothing Then Exit Sub

    Dim rawRows As Long, rawCols As Long
    rawRows = Rng.rows.count
    rawCols = Rng.columns.count

    ' Pull values into a 1-based 2-D array (handle single-cell edge case).
    Dim raw As Variant
    If rawRows = 1 And rawCols = 1 Then
        ReDim raw(1 To 1, 1 To 1)
        raw(1, 1) = Rng.value
    Else
        raw = Rng.value
    End If

    ' Applied to the whole block, headers included, so a merged header cell
    ' spans correctly. Duplicate names are then resolved by p_SanitizeHeaders.
    If Unmerge Then p_FillMergedValues Rng, raw, False

    m_ColCount = rawCols
    ReDim m_Headers(1 To m_ColCount)

    Dim dataStart As Long           ' first raw row that contains data
    Dim c As Long
    If hasHeaders Then
        For c = 1 To m_ColCount
            m_Headers(c) = CStr(raw(1, c))
        Next c
        p_SanitizeHeaders m_Headers, m_ColCount
        dataStart = 2
        m_RowCount = rawRows - 1
    Else
        For c = 1 To m_ColCount
            m_Headers(c) = "Column" & c
        Next c
        dataStart = 1
        m_RowCount = rawRows
    End If

    ' Fresh load - reset dtypes (read methods re-infer via convert_dtypes).
    m_TypesAllocated = False

    If m_RowCount <= 0 Then
        m_RowCount = 0
        p_ClearIndexes
        Exit Sub
    End If

    ReDim m_Data(1 To m_RowCount, 1 To m_ColCount)
    Dim r As Long
    For r = 1 To m_RowCount
        For c = 1 To m_ColCount
            m_Data(r, c) = raw(dataStart + r - 1, c)
        Next c
    Next r
    p_ClearIndexes
End Sub

' Shell sort m_Data in place by an ordered list of column indices.
' colIndices() is 0-based, values are 1-based column numbers.
Private Sub p_ShellSort(colIndices() As Long, ByVal numCols As Long, _
                        ByVal ascending As Boolean, ByVal NonAlphaAtEnd As Boolean)
    If m_RowCount <= 1 Then Exit Sub

    ' Compute initial gap (Knuth sequence).
    Dim gap As Long
    gap = 1
    Do While gap < m_RowCount \ 3
        gap = gap * 3 + 1
    Loop

    Dim i As Long, j As Long, c As Long
    Dim swapRow() As Variant
    ReDim swapRow(1 To m_ColCount)

    ' When lineage varies by row or carries overrides, track the permutation
    ' alongside the data and replay it through the engine once at the end.
    ' perm(i) is the ORIGINAL row index now sitting at position i. Comparison
    ' semantics are untouched, so sort results are identical either way.
    Dim trackLin As Boolean
    Dim perm() As Long
    Dim swapTag As Long
    trackLin = False
    If Not m_Lin Is Nothing Then
        trackLin = (m_Lin.HasMixedRowSources Or m_Lin.HasOverlays)
    End If
    If trackLin Then
        ReDim perm(1 To m_RowCount)
        For i = 1 To m_RowCount
            perm(i) = i
        Next i
    End If

    Do While gap >= 1
        For i = gap + 1 To m_RowCount
            ' Stash row i.
            For c = 1 To m_ColCount
                swapRow(c) = m_Data(i, c)
            Next c
            If trackLin Then swapTag = perm(i)
            j = i
            Do While j > gap
                Dim cmpRes As Long
                cmpRes = p_CompareRows(j - gap, swapRow, colIndices, numCols, NonAlphaAtEnd)
                Dim shift As Boolean
                shift = IIf(ascending, cmpRes > 0, cmpRes < 0)
                If Not shift Then Exit Do
                ' Move row j-gap one position down.
                For c = 1 To m_ColCount
                    m_Data(j, c) = m_Data(j - gap, c)
                Next c
                If trackLin Then perm(j) = perm(j - gap)
                j = j - gap
            Loop
            For c = 1 To m_ColCount
                m_Data(j, c) = swapRow(c)
            Next c
            If trackLin Then perm(j) = swapTag
        Next i
        gap = (gap - 1) \ 3
    Loop

    If trackLin Then m_Lin.x_ApplyRowGather perm, m_RowCount
End Sub
' ---------------------------------------------------------------------------
' Compare two non-NA scalar values using the declared cTable dtype.
'
' Returns:
'   -1 = left value sorts before right value
'    0 = values are equal
'    1 = left value sorts after right value
'
' Direction is not applied here. This always performs the natural ascending ' comparison. The caller reverses the result for descending columns.
' ---------------------------------------------------------------------------
Private Function p_CompareSortScalars( _
        ByVal leftValue As Variant, _
        ByVal rightValue As Variant, _
        ByVal declaredType As ctDataType) As Long

    Dim leftNumber As Double
    Dim rightNumber As Double

    Select Case declaredType
        Case ctString
            p_CompareSortScalars = StrComp( _
                                        CStr(leftValue), _
                                        CStr(rightValue), _
                                        vbTextCompare)
            Exit Function

        Case ctLong, ctDouble, ctCurrency
            leftNumber = CDbl(leftValue)
            rightNumber = CDbl(rightValue)

        Case ctDate
            ' VBA dates are stored internally as serial numbers.
            leftNumber = CDbl(CDate(leftValue))
            rightNumber = CDbl(CDate(rightValue))

        Case ctBoolean
            If CBool(leftValue) Then
                leftNumber = 1#
            Else
                leftNumber = 0#
            End If

            If CBool(rightValue) Then
                rightNumber = 1#
            Else
                rightNumber = 0#
            End If

        Case ctVariant
            ' For a mixed/untyped column, compare compatible Variant
            ' subtypes naturally and fall back to text comparison.
            If VarType(leftValue) = vbDate _
               And VarType(rightValue) = vbDate Then

                leftNumber = CDbl(CDate(leftValue))
                rightNumber = CDbl(CDate(rightValue))

            ElseIf VarType(leftValue) = vbBoolean _
               And VarType(rightValue) = vbBoolean Then

                If CBool(leftValue) Then
                    leftNumber = 1#
                Else
                    leftNumber = 0#
                End If

                If CBool(rightValue) Then
                    rightNumber = 1#
                Else
                    rightNumber = 0#
                End If

            ElseIf IsNumeric(leftValue) _
               And IsNumeric(rightValue) Then

                leftNumber = CDbl(leftValue)
                rightNumber = CDbl(rightValue)

            Else
                p_CompareSortScalars = StrComp( _
                                            CStr(leftValue), _
                                            CStr(rightValue), _
                                            vbTextCompare)
                Exit Function
            End If

        Case Else
            p_CompareSortScalars = StrComp( _
                                        CStr(leftValue), _
                                        CStr(rightValue), _
                                        vbTextCompare)
            Exit Function
    End Select

    If leftNumber < rightNumber Then
        p_CompareSortScalars = -1
    ElseIf leftNumber > rightNumber Then
        p_CompareSortScalars = 1
    Else
        p_CompareSortScalars = 0
    End If
End Function
' ---------------------------------------------------------------------------
' Compare one table row against a buffered row using multiple sort columns.
'
' Each column has its own ascending/descending flag.
'
' Missing-value placement is independent of direction:
'   naFirst=True  -> NA values sort first
'   naFirst=False -> NA values sort last
' ---------------------------------------------------------------------------
Private Function p_CompareSortRowToBuffer( _
        ByVal rowNum As Long, _
        ByRef compareRow() As Variant, _
        ByRef colIndexes() As Long, _
        ByRef ascendingFlags() As Boolean, _
        ByVal sortColumnCount As Long, _
        ByVal naFirst As Boolean) As Long

    Dim sortPosition As Long
    Dim colNum As Long

    Dim leftValue As Variant
    Dim rightValue As Variant

    Dim leftIsNA As Boolean
    Dim rightIsNA As Boolean

    Dim comparison As Long

    For sortPosition = 1 To sortColumnCount
        colNum = colIndexes(sortPosition)

        leftValue = m_Data(rowNum, colNum)
        rightValue = compareRow(colNum)

        leftIsNA = p_IsNA(leftValue)
        rightIsNA = p_IsNA(rightValue)

        If leftIsNA Or rightIsNA Then
            If leftIsNA And rightIsNA Then
                comparison = 0

            ElseIf leftIsNA Then
                If naFirst Then
                    comparison = -1
                Else
                    comparison = 1
                End If

            Else
                If naFirst Then
                    comparison = 1
                Else
                    comparison = -1
                End If
            End If

            ' Do not reverse NA placement for descending columns.
        Else
            comparison = p_CompareSortScalars( _
                            leftValue, _
                            rightValue, _
                            m_ColTypes(colNum))

            If comparison <> 0 Then
                If Not ascendingFlags(sortPosition) Then
                    comparison = -comparison
                End If
            End If
        End If

        If comparison <> 0 Then
            p_CompareSortRowToBuffer = comparison
            Exit Function
        End If
    Next sortPosition

    p_CompareSortRowToBuffer = 0
End Function


' ---------------------------------------------------------------------------
' Sort m_Data in place using independent directions for each sort column.
'
' colIndexes() and ascendingFlags() are both 1-based:
'
'   colIndexes(1)      = primary sort column
'   ascendingFlags(1)  = direction for primary sort column
'
'   colIndexes(2)      = secondary sort column
'   ascendingFlags(2)  = direction for secondary sort column
' ---------------------------------------------------------------------------
Private Sub p_ShellSortValues( _
        ByRef colIndexes() As Long, _
        ByRef ascendingFlags() As Boolean, _
        ByVal sortColumnCount As Long, _
        ByVal naFirst As Boolean)

    If m_RowCount <= 1 Then Exit Sub

    p_EnsureTypes

    Dim gap As Long
    gap = 1

    Do While gap < m_RowCount \ 3
        gap = gap * 3 + 1
    Loop

    Dim swapRow() As Variant
    ReDim swapRow(1 To m_ColCount)

    ' Track the original row positions when lineage requires it.
    Dim trackLineage As Boolean
    Dim permutation() As Long
    Dim swapLineageRow As Long

    trackLineage = False

    If Not m_Lin Is Nothing Then
        trackLineage = _
            (m_Lin.HasMixedRowSources Or m_Lin.HasOverlays)
    End If

    Dim i As Long
    Dim j As Long
    Dim c As Long
    Dim comparison As Long

    If trackLineage Then
        ReDim permutation(1 To m_RowCount)

        For i = 1 To m_RowCount
            permutation(i) = i
        Next i
    End If

    Do While gap >= 1
        For i = gap + 1 To m_RowCount
            ' Buffer the row being inserted.
            For c = 1 To m_ColCount
                swapRow(c) = m_Data(i, c)
            Next c

            If trackLineage Then
                swapLineageRow = permutation(i)
            End If

            j = i

            Do While j > gap
                comparison = p_CompareSortRowToBuffer( _
                                rowNum:=j - gap, _
                                compareRow:=swapRow, _
                                colIndexes:=colIndexes, _
                                ascendingFlags:=ascendingFlags, _
                                sortColumnCount:=sortColumnCount, _
                                naFirst:=naFirst)

                ' Positive means the existing row belongs after the
                ' buffered row under the requested sort specification.
                If comparison <= 0 Then Exit Do

                For c = 1 To m_ColCount
                    m_Data(j, c) = m_Data(j - gap, c)
                Next c

                If trackLineage Then
                    permutation(j) = permutation(j - gap)
                End If

                j = j - gap
            Loop

            For c = 1 To m_ColCount
                m_Data(j, c) = swapRow(c)
            Next c

            If trackLineage Then
                permutation(j) = swapLineageRow
            End If
        Next i

        gap = (gap - 1) \ 3
    Loop

    If trackLineage Then
        m_Lin.x_ApplyRowGather permutation, m_RowCount
    End If

    ' Existing indexes contain obsolete row numbers after a sort.
    p_ClearIndexes
End Sub



' Compare m_Data row rowIdx against a 1-based Variant row buffer.
' Returns -1, 0, or 1 using multi-column tie-breaking.
Private Function p_CompareRows( _
        ByVal rowIdx As Long, _
        ByRef compareRow() As Variant, _
        ByRef colIndices() As Long, _
        ByVal numCols As Long, _
        ByVal NonAlphaAtEnd As Boolean) As Long

    Dim k As Long
    Dim ci As Long

    Dim v1 As Variant
    Dim v2 As Variant

    Dim isNA1 As Boolean
    Dim isNA2 As Boolean

    Dim cmp As Long
    Dim d1 As Double
    Dim d2 As Double

    For k = 0 To numCols - 1
        ci = colIndices(k)

        v1 = m_Data(rowIdx, ci)
        v2 = compareRow(ci)

        isNA1 = p_IsNA(v1)
        isNA2 = p_IsNA(v2)

        ' --------------------------------------------------------
        ' Missing-value comparison.
        ' --------------------------------------------------------
        If isNA1 Or isNA2 Then
            If isNA1 And isNA2 Then
                cmp = 0

            ElseIf isNA1 Then
                If NonAlphaAtEnd Then
                    cmp = 1
                Else
                    cmp = -1
                End If

            Else
                If NonAlphaAtEnd Then
                    cmp = -1
                Else
                    cmp = 1
                End If
            End If

        Else
            ' ----------------------------------------------------
            ' Compare according to the declared cTable dtype.
            ' ----------------------------------------------------
            Select Case m_ColTypes(ci)
                Case ctDate
                    d1 = CDbl(CDate(v1))
                    d2 = CDbl(CDate(v2))

                    If d1 < d2 Then
                        cmp = -1
                    ElseIf d1 > d2 Then
                        cmp = 1
                    Else
                        cmp = 0
                    End If

                Case ctLong, ctDouble, ctCurrency
                    d1 = CDbl(v1)
                    d2 = CDbl(v2)

                    If d1 < d2 Then
                        cmp = -1
                    ElseIf d1 > d2 Then
                        cmp = 1
                    Else
                        cmp = 0
                    End If

                Case ctBoolean
                    If CBool(v1) Then
                        d1 = 1#
                    Else
                        d1 = 0#
                    End If

                    If CBool(v2) Then
                        d2 = 1#
                    Else
                        d2 = 0#
                    End If

                    If d1 < d2 Then
                        cmp = -1
                    ElseIf d1 > d2 Then
                        cmp = 1
                    Else
                        cmp = 0
                    End If

                Case ctString
                    cmp = StrComp( _
                              CStr(v1), _
                              CStr(v2), _
                              vbTextCompare)

                Case ctVariant
                    ' Use the actual Variant subtypes where possible.
                    If VarType(v1) = vbDate _
                       And VarType(v2) = vbDate Then

                        d1 = CDbl(CDate(v1))
                        d2 = CDbl(CDate(v2))

                        If d1 < d2 Then
                            cmp = -1
                        ElseIf d1 > d2 Then
                            cmp = 1
                        Else
                            cmp = 0
                        End If

                    ElseIf IsNumeric(v1) And IsNumeric(v2) Then
                        d1 = CDbl(v1)
                        d2 = CDbl(v2)

                        If d1 < d2 Then
                            cmp = -1
                        ElseIf d1 > d2 Then
                            cmp = 1
                        Else
                            cmp = 0
                        End If

                    Else
                        cmp = StrComp( _
                                  CStr(v1), _
                                  CStr(v2), _
                                  vbTextCompare)
                    End If

                Case Else
                    cmp = StrComp( _
                              CStr(v1), _
                              CStr(v2), _
                              vbTextCompare)
            End Select
        End If

        If cmp <> 0 Then
            p_CompareRows = cmp
            Exit Function
        End If
    Next k

    p_CompareRows = 0
End Function


' Build a result cTable from a filtered subset of current rows.
Private Function p_BuildFromRows(matchRows() As Long, _
                                   ByVal matchCount As Long) As cTable
    Set p_BuildFromRows = New cTable
    Dim empty2D() As Variant
    ReDim empty2D(1 To IIf(matchCount > 0, matchCount, 1), 1 To IIf(m_ColCount > 0, m_ColCount, 1))

    If matchCount = 0 Or m_ColCount = 0 Then
        p_BuildFromRows.from_array m_Headers, empty2D, 0, m_ColCount
        p_LinAttachGather p_BuildFromRows, matchRows, 0
        Exit Function
    End If

    Dim resData() As Variant
    ReDim resData(1 To matchCount, 1 To m_ColCount)
    Dim r As Long, c As Long
    For r = 1 To matchCount
        For c = 1 To m_ColCount
            resData(r, c) = m_Data(matchRows(r), c)
        Next c
    Next r
    p_BuildFromRows.from_array m_Headers, resData, matchCount, m_ColCount
    p_LinAttachGather p_BuildFromRows, matchRows, matchCount
End Function
' ---------------------------------------------------------------------------
' Replace the current rows with the selected source rows.
'
' Headers and dtypes are unchanged because this operation only removes rows.
' Existing indexes are cleared because their stored row numbers are no ' longer valid.
' ---------------------------------------------------------------------------
Private Sub p_KeepRowsInPlace( _
        ByRef selectedRows() As Long, _
        ByVal selectedCount As Long)

    If selectedCount < 0 Or selectedCount > m_RowCount Then
        Err.Raise vbObjectError + 1242, _
                  "cTable.p_KeepRowsInPlace", _
                  "Invalid selected row count: " & selectedCount
    End If

    ' No rows were removed.
    If selectedCount = m_RowCount Then Exit Sub

    ' Preserve the columns and dtypes, but remove every data row.
    If selectedCount = 0 Then
        Erase m_Data
        m_RowCount = 0
        If Not m_Lin Is Nothing Then m_Lin.x_ApplyRowGather selectedRows, 0
        p_ClearIndexes
        Exit Sub
    End If

    Dim newData() As Variant
    ReDim newData(1 To selectedCount, 1 To m_ColCount)

    Dim outputRow As Long
    Dim sourceRow As Long
    Dim c As Long

    For outputRow = 1 To selectedCount
        sourceRow = selectedRows(outputRow)

        For c = 1 To m_ColCount
            newData(outputRow, c) = m_Data(sourceRow, c)
        Next c
    Next outputRow

    m_Data = newData
    m_RowCount = selectedCount

    If Not m_Lin Is Nothing Then
        m_Lin.x_ApplyRowGather selectedRows, selectedCount
    End If

    ' Existing indexes contain obsolete row numbers.
    p_ClearIndexes
End Sub

' Inner implementation for Join, supports custom headers via newHdrs()
Private Function p_DoJoin(right As cTable, ByVal leftKey As String, _
                            ByVal rightKey As String, _
                            ByVal isLeft As Boolean, _
                            ByRef newHdrs() As String, _
                            ByRef outLeft() As Long, _
                            ByRef outRight() As Long, _
                            ByRef outCount As Long) As cTable
    ' Validate key columns.
    Dim lci As Long
    lci = p_ColIdx(leftKey)
    If lci = 0 Then
        Err.Raise vbObjectError + 51, "cTable.Join", _
                  "Left key column not found: """ & leftKey & """"
    End If
    Dim rci As Long
    rci = right.columnNumber(rightKey)
    If rci = 0 Then
        Err.Raise vbObjectError + 52, "cTable.Join", _
                  "Right key column not found: """ & rightKey & """"
    End If

    ' Hash the right table on its key column for O(n) lookup.
    Dim rIdx As Object
    Set rIdx = CreateObject("Scripting.Dictionary")
    rIdx.CompareMode = vbTextCompare
    Dim i As Long
    For i = 1 To right.rowCount
        Dim rKeyStr As String
        rKeyStr = CStr(right.GetValue(i, rci))
        If Not rIdx.Exists(rKeyStr) Then
            Set rIdx(rKeyStr) = New Collection
        End If
        rIdx(rKeyStr).Add i
    Next i

    ' Collect result rows into a Collection of 1-based Variant arrays.
    Dim rowBuf As Collection
    Set rowBuf = New Collection

    ' Record which left row and which right row produced each output row.
    ' This is all the lineage engine needs to split the result by column.
    Dim capLin As Boolean
    Dim capSize As Long
    capLin = (Not m_Lin Is Nothing)
    outCount = 0
    If capLin Then
        capSize = 1024
        ReDim outLeft(1 To capSize)
        ReDim outRight(1 To capSize)
    End If

    Dim totalCols As Long
    totalCols = UBound(newHdrs)

    Dim c As Long
    For i = 1 To m_RowCount
        Dim lKeyStr As String
        lKeyStr = CStr(m_Data(i, lci))

        If rIdx.Exists(lKeyStr) Then
            Dim matches As Collection
            Set matches = rIdx(lKeyStr)
            Dim m As Long
            For m = 1 To matches.count
                Dim ri As Long
                ri = matches(m)
                Dim nr() As Variant
                ReDim nr(1 To totalCols)
                ' Left columns
                For c = 1 To m_ColCount
                    nr(c) = m_Data(i, c)
                Next c
                ' Right columns
                For c = 1 To (totalCols - m_ColCount)
                    nr(m_ColCount + c) = right.GetValue(ri, c + IIf(c >= rci, 1, 0))
                Next c
                rowBuf.Add nr

                If capLin Then
                    outCount = outCount + 1
                    If outCount > capSize Then
                        capSize = capSize * 2
                        ReDim Preserve outLeft(1 To capSize)
                        ReDim Preserve outRight(1 To capSize)
                    End If
                    outLeft(outCount) = i
                    outRight(outCount) = ri
                End If
            Next m
        ElseIf isLeft Then
            ' Left join: emit left row with Null on right side.
            Dim nrL() As Variant
            ReDim nrL(1 To totalCols)
            For c = 1 To m_ColCount
                nrL(c) = m_Data(i, c)
            Next c
            For c = 1 To (totalCols - m_ColCount)
                nrL(m_ColCount + c) = Null
            Next c
            rowBuf.Add nrL

            If capLin Then
                outCount = outCount + 1
                If outCount > capSize Then
                    capSize = capSize * 2
                    ReDim Preserve outLeft(1 To capSize)
                    ReDim Preserve outRight(1 To capSize)
                End If
                outLeft(outCount) = i
                outRight(outCount) = 0        ' no right-side match
            End If
        End If
    Next i

    ' Convert Collection to 2-D array and package as cTable.
    Set p_DoJoin = New cTable
    Dim emptyPlaceholder() As Variant
    ReDim emptyPlaceholder(1 To 1, 1 To totalCols)

    If rowBuf.count = 0 Then
        p_DoJoin.from_array newHdrs, emptyPlaceholder, 0, totalCols
        Exit Function
    End If

    Dim resData() As Variant
    ReDim resData(1 To rowBuf.count, 1 To totalCols)
    Dim r As Long, k As Long
    For r = 1 To rowBuf.count
        Dim rowArr As Variant
        rowArr = rowBuf(r)
        For k = 1 To totalCols
            resData(r, k) = rowArr(k)
        Next k
    Next r
    p_DoJoin.from_array newHdrs, resData, rowBuf.count, totalCols
End Function

Private Function p_HeaderExists(ByVal colName As String) As Boolean
    p_HeaderExists = (p_ColIdx(colName) > 0)
End Function

'==============================================================================
'
' PRIVATE TYPE-SYSTEM HELPERS  (new in 20260723b)
'
'==============================================================================

' Guarantee m_ColTypes exists and matches m_ColCount. If a structural change
' bypassed the type system, rebuild as all-ctVariant rather than erroring.
Private Sub p_EnsureTypes()
    If m_ColCount = 0 Then
        m_TypesAllocated = False
        Exit Sub
    End If
    Dim needRebuild As Boolean
    needRebuild = Not m_TypesAllocated
    If m_TypesAllocated Then
        Dim ub As Long: ub = -1
        On Error Resume Next
        ub = UBound(m_ColTypes)
        On Error GoTo 0
        If ub <> m_ColCount Then needRebuild = True
    End If
    If needRebuild Then
        ReDim m_ColTypes(1 To m_ColCount)
        Dim i As Long
        For i = 1 To m_ColCount
            m_ColTypes(i) = ctVariant
        Next i
        m_TypesAllocated = True
    End If
End Sub

' Scalar missing-value test (core of IsNAValue / isna / dropna / fillna).
Private Function p_IsNA(ByVal v As Variant, _
                        Optional ByVal TreatEmptyStringAsNA As Boolean = True) As Boolean
    p_IsNA = False
    If IsNull(v) Then p_IsNA = True: Exit Function
    If IsEmpty(v) Then p_IsNA = True: Exit Function
    If IsError(v) Then p_IsNA = True: Exit Function
    If TreatEmptyStringAsNA Then
        If VarType(v) = vbString Then
            If Len(Trim$(v)) = 0 Then p_IsNA = True
        End If
    End If
End Function

' Strip currency/grouping decoration so "$1,234.50" or "(123.45)" parses.
Private Function p_CleanNumericString(ByVal s As String) As String
    Dim t As String
    t = Trim$(s)
    t = Replace(t, "$", "")
    t = Replace(t, ",", "")
    If Len(t) >= 2 Then
        If Left$(t, 1) = "(" And right$(t, 1) = ")" Then
            t = "-" & mid$(t, 2, Len(t) - 2)     ' accounting negative
        End If
    End If
    p_CleanNumericString = t
End Function

' Attempt to convert a single value to the target ctDataType.
' NA passes through as Null with ok=True (missing is valid in every dtype).
' On failure: ok=False and the ORIGINAL value is returned.
Private Function p_TryConvert(ByVal v As Variant, _
                              ByVal t As ctDataType, _
                              ByRef ok As Boolean) As Variant
    ok = True
    If p_IsNA(v) Then
        p_TryConvert = Null
        Exit Function
    End If

    On Error GoTo ConvFail
    Select Case t
        Case ctVariant
            p_TryConvert = v

        Case ctString
            p_TryConvert = CStr(v)

        Case ctLong
            If VarType(v) = vbString Then
                Dim sL As String: sL = p_CleanNumericString(CStr(v))
                If Not IsNumeric(sL) Then GoTo ConvFail
                p_TryConvert = CLng(CDbl(sL))
            ElseIf IsNumeric(v) Or VarType(v) = vbBoolean Then
                p_TryConvert = CLng(v)
            ElseIf VarType(v) = vbDate Then
                p_TryConvert = CLng(CDbl(v))
            Else
                GoTo ConvFail
            End If

        Case ctDouble
            If VarType(v) = vbString Then
                Dim sD As String: sD = p_CleanNumericString(CStr(v))
                If Not IsNumeric(sD) Then GoTo ConvFail
                p_TryConvert = CDbl(sD)
            ElseIf IsNumeric(v) Or VarType(v) = vbBoolean Or VarType(v) = vbDate Then
                p_TryConvert = CDbl(v)
            Else
                GoTo ConvFail
            End If

        Case ctCurrency
            If VarType(v) = vbString Then
                Dim sc As String: sc = p_CleanNumericString(CStr(v))
                If Not IsNumeric(sc) Then GoTo ConvFail
                p_TryConvert = CCur(sc)
            ElseIf IsNumeric(v) Or VarType(v) = vbBoolean Then
                p_TryConvert = CCur(v)
            Else
                GoTo ConvFail
            End If

        Case ctDate
            If VarType(v) = vbDate Then
                p_TryConvert = v
            ElseIf VarType(v) = vbString Then
                If Not IsDate(v) Then GoTo ConvFail
                p_TryConvert = CDate(v)
            ElseIf IsNumeric(v) Then
                p_TryConvert = CDate(CDbl(v))
            Else
                GoTo ConvFail
            End If

        Case ctBoolean
            If VarType(v) = vbBoolean Then
                p_TryConvert = v
            ElseIf IsNumeric(v) Then
                p_TryConvert = CBool(v)
            ElseIf VarType(v) = vbString Then
                Select Case LCase$(Trim$(CStr(v)))
                    Case "true", "yes", "y", "1":  p_TryConvert = True
                    Case "false", "no", "n", "0": p_TryConvert = False
                    Case Else: GoTo ConvFail
                End Select
            Else
                GoTo ConvFail
            End If

        Case Else
            GoTo ConvFail
    End Select
    Exit Function

ConvFail:
    On Error GoTo 0
    ok = False
    p_TryConvert = v
End Function

' Infer the best ctDataType for one column by scanning its non-NA values.
Private Function p_InferColumnType(ByVal ci As Long) As ctDataType
    Dim nBool As Long, nDate As Long, nLong As Long, nDbl As Long
    Dim nCur As Long, nStr As Long, nOther As Long, nTotal As Long
    Dim r As Long, v As Variant

    For r = 1 To m_RowCount
        v = m_Data(r, ci)
        If Not p_IsNA(v) Then
            nTotal = nTotal + 1
            Select Case VarType(v)
                Case vbBoolean
                    nBool = nBool + 1
                Case vbDate
                    nDate = nDate + 1
                Case vbCurrency
                    nCur = nCur + 1
                Case vbByte, vbInteger, vbLong
                    nLong = nLong + 1
                Case vbSingle, vbDouble, vbDecimal
                    Dim d As Double: d = CDbl(v)
                    If d = Fix(d) And Abs(d) <= 2147483647# Then
                        nLong = nLong + 1
                    Else
                        nDbl = nDbl + 1
                    End If
                Case vbString
                    Dim s As String: s = Trim$(CStr(v))
                    Dim sc As String: sc = p_CleanNumericString(s)
                    Select Case True
                        Case LCase$(s) = "true" Or LCase$(s) = "false"
                            nBool = nBool + 1
                        Case InStr(1, s, "$") > 0 And IsNumeric(sc)
                            nCur = nCur + 1
                        Case IsNumeric(sc) And sc <> s
                            nCur = nCur + 1
                        Case IsNumeric(s)
                            Dim ds As Double: ds = CDbl(s)
                            If ds = Fix(ds) And Abs(ds) <= 2147483647# _
                               And InStr(1, s, ".") = 0 _
                               And InStr(1, LCase$(s), "e") = 0 Then
                                nLong = nLong + 1
                            Else
                                nDbl = nDbl + 1
                            End If
                        Case IsDate(s)
                            nDate = nDate + 1
                        Case Else
                            nStr = nStr + 1
                    End Select
                Case Else
                    nOther = nOther + 1
            End Select
        End If
    Next r

    If nTotal = 0 Then
        p_InferColumnType = ctVariant
    ElseIf nStr > 0 Or nOther > 0 Then
        If nStr = nTotal Then
            p_InferColumnType = ctString
        Else
            p_InferColumnType = ctVariant
        End If
    ElseIf nDate = nTotal Then
        p_InferColumnType = ctDate
    ElseIf nBool = nTotal Then
        p_InferColumnType = ctBoolean
    ElseIf nDate > 0 Or nBool > 0 Then
        p_InferColumnType = ctVariant
    ElseIf nCur > 0 Then
        p_InferColumnType = ctCurrency
    ElseIf nDbl > 0 Then
        p_InferColumnType = ctDouble
    ElseIf nLong = nTotal Then
        p_InferColumnType = ctLong
    Else
        p_InferColumnType = ctVariant
    End If
End Function

' Convert every value in a column; returns the number of failed conversions.
Private Function p_CoerceColumn(ByVal ci As Long, _
                                ByVal targetType As ctDataType, _
                                ByVal errors As ctConvertErrors) As Long
    Dim failCount As Long, r As Long, ok As Boolean, nV As Variant
    For r = 1 To m_RowCount
        nV = p_TryConvert(m_Data(r, ci), targetType, ok)
        If ok Then
            m_Data(r, ci) = nV
        Else
            failCount = failCount + 1
            Select Case errors
                Case ctErrorsRaise
                    Err.Raise vbObjectError + 910, "cTable.astype", _
                        "Cannot convert value """ & CStr(m_Data(r, ci)) & _
                        """ (row " & r & ", column """ & m_Headers(ci) & _
                        """) to " & Me.dtype_name(targetType)
                Case ctErrorsCoerce
                    m_Data(r, ci) = Null
                Case Else
                    ' ctErrorsPreserve: leave original value
            End Select
        End If
    Next r
    p_CoerceColumn = failCount
End Function

' NA count for one column (used by info / count / hasnans).
Private Function p_ColNACount(ByVal ci As Long) As Long
    Dim r As Long, n As Long
    For r = 1 To m_RowCount
        If p_IsNA(m_Data(r, ci)) Then n = n + 1
    Next r
    p_ColNACount = n
End Function

' Excel NumberFormat string for a dtype (used by typed export).
Private Function p_NumberFormatFor(ByVal t As ctDataType, _
                                   Optional ByVal dateFormat As String = "mm/dd/yyyy") As String
    Select Case t
        Case ctString:   p_NumberFormatFor = "@"
        Case ctLong:     p_NumberFormatFor = "#,##0"
        Case ctDouble:   p_NumberFormatFor = "General"
        Case ctCurrency: p_NumberFormatFor = "$#,##0.00"
        Case ctDate:     p_NumberFormatFor = dateFormat
        Case ctBoolean:  p_NumberFormatFor = "General"
        Case Else:       p_NumberFormatFor = "General"
    End Select
End Function

' V1 helper retained for AggregateDict compatibility.
Private Function CleanCurrency(val As Variant) As Double
    Dim s As String
    s = CStr(val)
    s = Replace(s, "$", "")
    s = Replace(s, ",", "")
    s = Trim(s)
    If IsNumeric(s) Then
        CleanCurrency = CDbl(s)
    Else
        CleanCurrency = 0
    End If
End Function

'==============================================================================
'
' PUBLIC INITIALISATION METHODS
'
'==============================================================================

' ---------------------------------------------------------------------------
' from_array
'   Core public initialiser - accepts 1-based header and data arrays directly.
'   Useful for unit-testing or building a cTable programmatically.
'   hdrs(1..colCount), data(1..rowCount, 1..colCount)
' ---------------------------------------------------------------------------
Public Sub from_array(hdrs() As String, data() As Variant, _
                           ByVal rowCount As Long, ByVal colCount As Long)
    m_ColCount = colCount
    m_RowCount = rowCount
    m_TypesAllocated = False

    If colCount <= 0 Then
        m_ColCount = 0: m_RowCount = 0
        If Not m_Lin Is Nothing Then m_Lin.InitUniform 0, 0, ctSrcUnknown
        p_ClearIndexes
        Exit Sub
    End If

    ReDim m_Headers(1 To m_ColCount)
    Dim i As Long
    For i = 1 To m_ColCount
        m_Headers(i) = hdrs(i)
    Next i
    p_SanitizeHeaders m_Headers, m_ColCount

    If rowCount > 0 Then
        ReDim m_Data(1 To rowCount, 1 To m_ColCount)
        Dim j As Long
        For i = 1 To rowCount
            For j = 1 To m_ColCount
                m_Data(i, j) = data(i, j)
            Next j
        Next i
    End If

    ' from_array replaces shape wholesale and is public, so it is the lineage
    ' reset point. Derived-table methods overwrite this immediately with the
    ' correct engine; anything that forgets reports Unknown.
    If Not m_Lin Is Nothing Then
        m_Lin.InitUniform m_RowCount, m_ColCount, ctSrcUnknown
    End If

    p_ClearIndexes
End Sub

'------------------------------------------------------------------------------
' read_excel_sheet
'
' Loads data from a worksheet range, starting at Offset within the sheet's
' UsedRange. Offset=1 starts at the first UsedRange row.
'
' IncludeSourceRow:
'   True  = appends SourceRowColumnName containing the physical worksheet row
'           number for each loaded data row.
'   False = no source-row column is added.
'
' When hasHeaders=False, all source rows are data rows. Therefore, if the
' read range begins on worksheet row 5, the first _SourceRow value is 5.
'------------------------------------------------------------------------------
Public Sub read_excel_sheet(ByVal filePath As String, _
                            Optional ByVal sheetName As String = "", _
                            Optional ByVal hasHeaders As Boolean = True, _
                            Optional ByVal offset As Long = 1, _
                            Optional ByVal InferTypes As Boolean = True, _
                            Optional ByVal nrows As Long = 0, _
                            Optional ByVal Unmerge As Boolean = False, _
                            Optional ByVal IncludeSourceRow As Boolean = False, _
                            Optional ByVal SourceRowColumnName As String = "_SourceRow")

    Const CALLER As String = "cTable.read_excel_sheet"

    p_ValidateNRows nrows, CALLER

    Dim priorScreenUpdating As Boolean
    Dim priorCalculation As XlCalculation

    Dim didOpen As Boolean
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim ur As Range
    Dim Rng As Range

    Dim actualFilePath As String
    Dim skip As Long
    Dim firstDataWorksheetRow As Long

    Dim errNumber As Long
    Dim errDescription As String

    On Error GoTo ErrHandler

    priorScreenUpdating = Application.ScreenUpdating
    priorCalculation = Application.Calculation

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    '-------------------------------------------------------------------------
    ' Open or identify the workbook.
    '-------------------------------------------------------------------------
    If Len(filePath) = 0 Then
        Set wb = ThisWorkbook
        didOpen = False
        actualFilePath = wb.fullName
    Else
        Set wb = p_GetWorkbook(filePath, didOpen)

        If wb Is Nothing Then
            Err.Raise vbObjectError + 1, CALLER, _
                      "Cannot open workbook: " & filePath
        End If

        actualFilePath = filePath
    End If

    '-------------------------------------------------------------------------
    ' Establish source metadata and initialize lineage, when enabled.
    '-------------------------------------------------------------------------
    Set m_DataSource = New cDataSource
    m_DataSource.InitFromWorkbook wb, actualFilePath, sheetName

    If Not m_Lin Is Nothing Then
        m_Lin.InitUniform 0, 0, _
            cLineageRegistry.SourceId(m_DataSource)
    End If

    '-------------------------------------------------------------------------
    ' Locate the requested worksheet.
    '-------------------------------------------------------------------------
    If Len(sheetName) = 0 Then
        If wb.Worksheets.count > 0 Then
            Set ws = wb.Worksheets(1)
        End If
    Else
        On Error Resume Next
        Set ws = wb.Worksheets(sheetName)
        On Error GoTo ErrHandler
    End If

    If ws Is Nothing Then
        Err.Raise vbObjectError + 2, CALLER, _
                  "Sheet not found: " & sheetName
    End If

    Set ur = ws.UsedRange

    If ur Is Nothing Then GoTo CleanExit

    ' offset=1 means the first row of UsedRange.
    skip = IIf(offset < 1, 0, offset - 1)

    If skip >= ur.rows.count Then GoTo CleanExit

    Set Rng = ur.offset(skip, 0) _
                .Resize(ur.rows.count - skip, ur.columns.count)

    ' nrows counts data rows, not the header row.
    Set Rng = p_LimitReadRange(Rng, hasHeaders, nrows)

    '-------------------------------------------------------------------------
    ' Read the selected range into cTable storage.
    '-------------------------------------------------------------------------
    p_LoadFromRange Rng, hasHeaders, Unmerge

    If InferTypes Then Me.convert_dtypes

    ' Attribute the loaded cells to this workbook/sheet source.
    p_LinAdoptSource

    '-------------------------------------------------------------------------
    ' Add physical Excel source-row numbers after loading.
    '
    ' Example:
    '   rng.Row = 10 and hasHeaders=True  -> first data row is Excel row 11
    '   rng.Row = 10 and hasHeaders=False -> first data row is Excel row 10
    '-------------------------------------------------------------------------
    If IncludeSourceRow And m_RowCount > 0 Then

        firstDataWorksheetRow = Rng.Row

        If hasHeaders Then
            firstDataWorksheetRow = firstDataWorksheetRow + 1
        End If

        p_AddExcelSourceRowColumn _
            firstDataWorksheetRow:=firstDataWorksheetRow, _
            SourceRowColumnName:=SourceRowColumnName
    End If

CleanExit:
    On Error Resume Next

    If didOpen Then
        If Not wb Is Nothing Then wb.Close SaveChanges:=False
    End If

    Application.Calculation = priorCalculation
    Application.ScreenUpdating = priorScreenUpdating

    On Error GoTo 0

    If errNumber <> 0 Then
        Err.Raise errNumber, CALLER, errDescription
    End If

    Exit Sub

ErrHandler:
    errNumber = Err.Number
    errDescription = Err.Description
    Resume CleanExit

End Sub


'------------------------------------------------------------------------------
' read_excel_table
'
' Loads data from a named Excel Table (ListObject).
'
' If sheetName is omitted, the workbook is searched for tableName.
'
' IncludeSourceRow:
'   True  = appends SourceRowColumnName containing the physical worksheet row
'           number for each loaded ListObject data row.
'   False = no source-row column is added.
'
' For a ListObject, the source row is always based on DataBodyRange. The
' ListObject header row is not part of DataBodyRange and is never included in
' cTable data rows.
'------------------------------------------------------------------------------
Public Sub read_excel_table(ByVal filePath As String, _
                            ByVal tableName As String, _
                            Optional ByVal sheetName As String = "", _
                            Optional ByVal InferTypes As Boolean = True, _
                            Optional ByVal nrows As Long = 0, _
                            Optional ByVal Unmerge As Boolean = False, _
                            Optional ByVal IncludeSourceRow As Boolean = False, _
                            Optional ByVal SourceRowColumnName As String = "_SourceRow")

    Const CALLER As String = "cTable.read_excel_table"

    p_ValidateNRows nrows, CALLER

    Dim didOpen As Boolean
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim wsx As Worksheet
    Dim lo As ListObject

    Dim firstDataWorksheetRow As Long

    On Error GoTo ErrHandler

    '-------------------------------------------------------------------------
    ' Open or reuse the workbook.
    '-------------------------------------------------------------------------
    Set wb = p_GetWorkbook(filePath, didOpen)

    If wb Is Nothing Then
        Err.Raise vbObjectError + 1, CALLER, _
                  "Cannot open workbook: " & filePath
    End If

    '-------------------------------------------------------------------------
    ' Locate the Excel Table.
    '-------------------------------------------------------------------------
    If Len(sheetName) > 0 Then

        On Error Resume Next
        Set ws = wb.Worksheets(sheetName)
        On Error GoTo ErrHandler

        If ws Is Nothing Then
            Err.Raise vbObjectError + 2, CALLER, _
                      "Sheet not found: " & sheetName
        End If

        On Error Resume Next
        Set lo = ws.ListObjects(tableName)
        On Error GoTo ErrHandler

    Else

        For Each wsx In wb.Worksheets
            On Error Resume Next
            Set lo = wsx.ListObjects(tableName)
            On Error GoTo ErrHandler

            If Not lo Is Nothing Then
                Set ws = wsx
                Exit For
            End If
        Next wsx

    End If

    If lo Is Nothing Then
        Err.Raise vbObjectError + 3, CALLER, _
                  "Table not found: """ & tableName & """"
    End If

    ' If the caller omitted sheetName, record the actual sheet that contains
    ' the Excel Table.
    If ws Is Nothing Then
        Set ws = lo.parent
    End If

    '-------------------------------------------------------------------------
    ' Establish source metadata and initialize lineage when enabled.
    '-------------------------------------------------------------------------
    Set m_DataSource = New cDataSource
    m_DataSource.InitFromPath filePath, ws.name, tableName

    If Not m_Lin Is Nothing Then
        m_Lin.InitUniform 0, 0, _
            cLineageRegistry.SourceId(m_DataSource)
    End If

    '-------------------------------------------------------------------------
    ' Load and sanitize ListObject headers.
    '-------------------------------------------------------------------------
    m_ColCount = lo.ListColumns.count
    ReDim m_Headers(1 To m_ColCount)

    Dim c As Long

    For c = 1 To m_ColCount
        m_Headers(c) = lo.ListColumns(c).name
    Next c

    p_SanitizeHeaders m_Headers, m_ColCount
    m_TypesAllocated = False

    '-------------------------------------------------------------------------
    ' Load ListObject data body.
    '-------------------------------------------------------------------------
    If lo.DataBodyRange Is Nothing Then

        ' The table has headers but no data rows.
        m_RowCount = 0
        Erase m_Data

    Else

        m_RowCount = lo.DataBodyRange.rows.count

        If nrows > 0 Then
            If nrows < m_RowCount Then
                m_RowCount = nrows
            End If
        End If

        Dim dataRng As Range
        Set dataRng = lo.DataBodyRange.Resize(m_RowCount, m_ColCount)

        Dim raw As Variant

        If m_RowCount = 1 And m_ColCount = 1 Then
            ReDim raw(1 To 1, 1 To 1)
            raw(1, 1) = dataRng.value
        Else
            raw = dataRng.value
        End If

        ' ListObjects normally cannot contain merged cells. This remains
        ' available for API symmetry with the other Excel readers.
        If Unmerge Then
            p_FillMergedValues dataRng, raw, False
        End If

        ReDim m_Data(1 To m_RowCount, 1 To m_ColCount)

        Dim r As Long

        For r = 1 To m_RowCount
            For c = 1 To m_ColCount
                m_Data(r, c) = raw(r, c)
            Next c
        Next r

        ' DataBodyRange begins at the physical worksheet row for cTable row 1.
        firstDataWorksheetRow = dataRng.Row
    End If

    p_ClearIndexes

    If InferTypes Then Me.convert_dtypes

    p_LinAdoptSource

    '-------------------------------------------------------------------------
    ' Add a physical worksheet source-row column.
    '
    ' The first row of DataBodyRange is the first cTable data row, regardless
    ' of whether nrows limits the read. The ListObject's header row is not
    ' counted or included.
    '-------------------------------------------------------------------------
    If IncludeSourceRow And m_RowCount > 0 Then
        p_AddExcelSourceRowColumn _
            firstDataWorksheetRow:=firstDataWorksheetRow, _
            SourceRowColumnName:=SourceRowColumnName
    End If

CleanExit:
    On Error Resume Next

    If didOpen Then
        If Not wb Is Nothing Then wb.Close SaveChanges:=False
    End If

    On Error GoTo 0
    Exit Sub

ErrHandler:
    Dim errNumber As Long
    Dim errDescription As String

    errNumber = Err.Number
    errDescription = Err.Description

    Resume CleanFail

CleanFail:
    On Error Resume Next

    If didOpen Then
        If Not wb Is Nothing Then wb.Close SaveChanges:=False
    End If

    On Error GoTo 0

    Err.Raise errNumber, CALLER, errDescription

End Sub



' ---------------------------------------------------------------------------
' read_access_table
'   Loads a table from an Access database via ADO.
'   InferTypes=True (default) maps ADO field types directly to ctDataType.
'   if nrows = 0 then return all rows otherwise return nrows of data (excluding header)
' ---------------------------------------------------------------------------
Public Sub read_access_table(ByVal filePath As String, ByVal tableName As String, _
                             Optional ByVal InferTypes As Boolean = True, _
                             Optional ByVal nrows As Long = 0)

    Const CALLER As String = "cTable.read_access_table"

    p_ValidateNRows nrows, CALLER

    Dim conn As Object
    Dim rs As Object
    Dim connStr As String
    Dim extPos As Long
    Dim ext As String
    Dim r As Long, c As Long
    Dim errNum As Long
    Dim errDesc As String

    On Error GoTo CleanUp

    ' Extract file extension robustly
    extPos = InStrRev(filePath, ".")
    If extPos > 0 Then
        ext = LCase$(mid$(filePath, extPos + 1))
    Else
        ext = ""
    End If

    If ext = "accdb" Then
        connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & filePath
    ElseIf ext = "mdb" Then
        connStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & filePath
    Else
        Err.Raise vbObjectError + 200, CALLER, _
                  "Unsupported Access file type: " & filePath
    End If

    ' Initialize DataSource metadata first
    Set m_DataSource = New cDataSource
    m_DataSource.InitFromAccessPath filePath, tableName

    ' Create ADO connection and recordset
    Set conn = CreateObject("ADODB.Connection")
    Set rs = CreateObject("ADODB.Recordset")

    ' Open connection (read-only mode)
    conn.Open connStr

    ' Limit the query itself when nrows is supplied.
    Dim sql As String
    If nrows > 0 Then
        sql = "SELECT TOP " & CStr(nrows) & " * FROM [" & tableName & "]"
    Else
        sql = "SELECT * FROM [" & tableName & "]"
    End If

    ' Open table with static cursor (supports RecordCount)
    rs.Open sql, conn, 3, 1

    ' Build headers from the recordset fields
    m_ColCount = rs.fields.count
    ReDim m_Headers(1 To m_ColCount)
    For c = 1 To m_ColCount
        m_Headers(c) = rs.fields(c - 1).name
    Next c
    p_SanitizeHeaders m_Headers, m_ColCount
    m_TypesAllocated = False

    ' Map ADO field types to ctDataType (exact - no inference needed)
    If InferTypes Then
        Dim tArr() As Variant
        ReDim tArr(1 To m_ColCount)
        For c = 1 To m_ColCount
            Select Case rs.fields(c - 1).Type
                Case 2, 3, 16, 17, 18, 19, 20, 21
                    tArr(c) = CLng(ctLong)
                Case 4, 5, 131, 14
                    tArr(c) = CLng(ctDouble)
                Case 6
                    tArr(c) = CLng(ctCurrency)
                Case 7, 133, 134, 135
                    tArr(c) = CLng(ctDate)
                Case 11
                    tArr(c) = CLng(ctBoolean)
                Case 8, 129, 130, 200, 201, 202, 203
                    tArr(c) = CLng(ctString)
                Case Else
                    tArr(c) = CLng(ctVariant)
            End Select
        Next c
        Me.dtypes_array = tArr
    End If

    ' Load data rows
    m_RowCount = rs.RecordCount
    If m_RowCount > 0 And m_ColCount > 0 Then
        ReDim m_Data(1 To m_RowCount, 1 To m_ColCount)
        r = 1
        rs.MoveFirst
        Do While Not rs.EOF
            For c = 1 To m_ColCount
                m_Data(r, c) = rs.fields(c - 1).value
            Next c
            r = r + 1
            rs.MoveNext
        Loop
    Else
        ReDim m_Data(1 To 1, 1 To m_ColCount)
    End If

    p_ClearIndexes
    p_LinAdoptSource

CleanUp:
    errNum = Err.Number
    errDesc = Err.Description

    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
        Set conn = Nothing
    End If
    On Error GoTo 0

    If errNum <> 0 Then
        If m_Log Is Nothing Then Set m_Log = New cLog
        m_Log.AddError CALLER & ": " & errDesc
        Err.Raise errNum, CALLER, errDesc
    End If

End Sub


'------------------------------------------------------------------------------
' read_excel_range
'
' Loads data from an Excel Range object directly.
'
' IncludeSourceRow:
'   True  = appends SourceRowColumnName containing the physical worksheet row
'           number from which each cTable data row was read.
'   False = no source-row column is added.
'
' The source row values are based on the actual worksheet rows represented by
' the supplied range, not merely 1..RowCount.
'
' Examples:
'
'   ' Range A1:D100, with a header in row 1:
'   t.read_excel_range Range("A1:D100"), hasHeaders:=True, _
'                      IncludeSourceRow:=True
'
'   ' _SourceRow starts at 2.
'
'   ' Range A10:D100, without headers:
'   t.read_excel_range Range("A10:D100"), hasHeaders:=False, _
'                      IncludeSourceRow:=True
'
'   ' _SourceRow starts at 10.
'------------------------------------------------------------------------------
Public Sub read_excel_range(ByVal Rng As Excel.Range, _
                            Optional hasHeaders As Boolean = True, _
                            Optional ws As Excel.Worksheet = Nothing, _
                            Optional wb As Excel.Workbook = Nothing, _
                            Optional ByVal InferTypes As Boolean = True, _
                            Optional ByVal nrows As Long = 0, _
                            Optional ByVal Unmerge As Boolean = False, _
                            Optional ByVal IncludeSourceRow As Boolean = False, _
                            Optional ByVal SourceRowColumnName As String = "_SourceRow")

    Const CALLER As String = "cTable.read_excel_range"

    p_ValidateNRows nrows, CALLER

    Dim readRng As Excel.Range

    Dim filePath As String
    Dim rawRows As Long
    Dim rawCols As Long

    Dim arr As Variant

    Dim dataStart As Long
    Dim firstDataWorksheetRow As Long

    Dim r As Long
    Dim c As Long

    On Error GoTo ErrHandler

    If Rng Is Nothing Then Exit Sub

    ' Derive worksheet and workbook from the range when the caller does not
    ' explicitly supply them.
    If ws Is Nothing Then
        Set ws = Rng.Worksheet
    End If

    If wb Is Nothing Then
        Set wb = ws.parent
    End If

    If wb Is Nothing Then
        Err.Raise vbObjectError + 1, CALLER, _
                  "Could not determine the workbook for the supplied range."
    End If

    ' nrows counts data rows, excluding the header row when present.
    Set readRng = p_LimitReadRange(Rng, hasHeaders, nrows)

    If readRng Is Nothing Then Exit Sub

    rawRows = readRng.rows.count
    rawCols = readRng.columns.count

    If rawRows <= 0 Or rawCols <= 0 Then Exit Sub

    If Len(wb.path) > 0 Then
        filePath = wb.fullName
    Else
        filePath = "(Unsaved Workbook)"
    End If

    '-------------------------------------------------------------------------
    ' Read values into a 1-based two-dimensional array.
    '-------------------------------------------------------------------------
    If rawRows = 1 And rawCols = 1 Then
        ReDim arr(1 To 1, 1 To 1)
        arr(1, 1) = readRng.Value2
    Else
        arr = readRng.Value2
    End If

    ' Restore values in merged blocks before splitting headers and data.
    If Unmerge Then
        p_FillMergedValues readRng, arr, True
    End If

    '-------------------------------------------------------------------------
    ' Build headers.
    '-------------------------------------------------------------------------
    m_ColCount = rawCols
    ReDim m_Headers(1 To m_ColCount)
    m_TypesAllocated = False

    If hasHeaders Then

        For c = 1 To m_ColCount
            m_Headers(c) = CStr(arr(1, c))
        Next c

        p_SanitizeHeaders m_Headers, m_ColCount

        dataStart = 2
        m_RowCount = rawRows - 1

        ' The first cTable data row is directly beneath the Excel header row.
        firstDataWorksheetRow = readRng.Row + 1

    Else

        For c = 1 To m_ColCount
            m_Headers(c) = "Column" & c
        Next c

        dataStart = 1
        m_RowCount = rawRows

        ' Every supplied Excel row is a data row.
        firstDataWorksheetRow = readRng.Row

    End If

    '-------------------------------------------------------------------------
    ' Handle header-only ranges.
    '-------------------------------------------------------------------------
    If m_RowCount <= 0 Then
        m_RowCount = 0
        Erase m_Data

        p_ClearIndexes

        If InferTypes Then Me.convert_dtypes

        GoTo SetMetadata
    End If

    '-------------------------------------------------------------------------
    ' Copy data values into cTable storage.
    '-------------------------------------------------------------------------
    ReDim m_Data(1 To m_RowCount, 1 To m_ColCount)

    For r = 1 To m_RowCount
        For c = 1 To m_ColCount
            m_Data(r, c) = arr(dataStart + r - 1, c)
        Next c
    Next r

    p_ClearIndexes

    If InferTypes Then Me.convert_dtypes

SetMetadata:
    '-------------------------------------------------------------------------
    ' Record provenance metadata.
    '-------------------------------------------------------------------------
    If m_DataSource Is Nothing Then
        Set m_DataSource = New cDataSource
    End If

    m_DataSource.InitFromWorkbook wb, filePath, ws.name
    m_DataSource.SourceType = "ExcelRange"

    If m_Log Is Nothing Then
        Set m_Log = New cLog
    End If

    m_Log.AddStatus _
        "Loaded data from range " & _
        readRng.Address(External:=True) & _
        " in worksheet '" & ws.name & _
        "' of workbook '" & wb.name & _
        "'. Rows: " & m_RowCount & _
        ", Columns: " & m_ColCount & _
        ", Headers: " & hasHeaders

    ' Adopt the source after the data shape is final.
    p_LinAdoptSource

    '-------------------------------------------------------------------------
    ' Add physical worksheet row numbers after source lineage is established.
    '-------------------------------------------------------------------------
    If IncludeSourceRow And m_RowCount > 0 Then
        p_AddExcelSourceRowColumn _
            firstDataWorksheetRow:=firstDataWorksheetRow, _
            SourceRowColumnName:=SourceRowColumnName
    End If

    Exit Sub

ErrHandler:
    If m_Log Is Nothing Then
        Set m_Log = New cLog
    End If

    m_Log.AddError _
        "Error loading data from range: " & Err.Description

    Err.Raise Err.Number, CALLER, Err.Description

End Sub



'------------------------------------------------------------------------------
' read_excel_ws
'
' Loads data from a worksheet object, starting at Offset within the sheet's
' UsedRange. Offset=1 starts at the first UsedRange row.
'
' IncludeSourceRow:
'   True  = appends SourceRowColumnName containing the physical worksheet row
'           number for each loaded data row.
'   False = no source-row column is added.
'
' Examples:
'
'   ' Header is on the first loaded row; data starts on the next worksheet row.
'   t.read_excel_ws ws, hasHeaders:=True, IncludeSourceRow:=True
'
'   ' No header; every loaded worksheet row is data.
'   t.read_excel_ws ws, hasHeaders:=False, IncludeSourceRow:=True
'
'   ' Start at the fifth row of UsedRange.
'   t.read_excel_ws ws, hasHeaders:=False, offset:=5, _
'                   IncludeSourceRow:=True
'------------------------------------------------------------------------------
Public Sub read_excel_ws(ByVal ws As Excel.Worksheet, _
                         Optional ByVal hasHeaders As Boolean = True, _
                         Optional ByVal offset As Long = 1, _
                         Optional ByVal InferTypes As Boolean = True, _
                         Optional ByVal nrows As Long = 0, _
                         Optional ByVal Unmerge As Boolean = False, _
                         Optional ByVal IncludeSourceRow As Boolean = False, _
                         Optional ByVal SourceRowColumnName As String = "_SourceRow")

    Const CALLER As String = "cTable.read_excel_ws"

    p_ValidateNRows nrows, CALLER

    If ws Is Nothing Then
        Err.Raise vbObjectError + 2, CALLER, _
                  "Worksheet is Nothing."
    End If

    ' Preserve caller-owned Excel application state.
    Dim priorScreenUpdating As Boolean
    Dim priorCalculation As XlCalculation

    priorScreenUpdating = Application.ScreenUpdating
    priorCalculation = Application.Calculation

    Dim errNumber As Long
    Dim errDescription As String

    Dim wb As Workbook
    Dim filePath As String
    Dim ur As Range
    Dim Rng As Range

    Dim skip As Long
    Dim firstDataWorksheetRow As Long

    On Error GoTo ErrHandler

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    '-------------------------------------------------------------------------
    ' Establish source metadata.
    '-------------------------------------------------------------------------
    Set m_DataSource = New cDataSource

    Set wb = ws.parent

    If Len(wb.path) > 0 Then
        filePath = wb.fullName
    Else
        filePath = "(Unsaved Workbook)"
    End If

    m_DataSource.InitFromWorkbook wb, filePath, ws.name

    If Not m_Lin Is Nothing Then
        m_Lin.InitUniform 0, 0, _
            cLineageRegistry.SourceId(m_DataSource)
    End If

    '-------------------------------------------------------------------------
    ' Identify the source range.
    '-------------------------------------------------------------------------
    Set ur = ws.UsedRange

    If ur Is Nothing Then GoTo CleanExit

    ' offset=1 starts at the first UsedRange row.
    skip = IIf(offset < 1, 0, offset - 1)

    If skip >= ur.rows.count Then GoTo CleanExit

    Set Rng = ur.offset(skip, 0) _
                .Resize(ur.rows.count - skip, ur.columns.count)

    ' nrows limits data rows, excluding a header row when present.
    Set Rng = p_LimitReadRange(Rng, hasHeaders, nrows)

    '-------------------------------------------------------------------------
    ' Load values, infer types, and adopt source lineage.
    '-------------------------------------------------------------------------
    p_LoadFromRange Rng, hasHeaders, Unmerge

    If InferTypes Then Me.convert_dtypes

    p_LinAdoptSource

    '-------------------------------------------------------------------------
    ' Append the original physical worksheet row number for every data row.
    '
    ' If rng begins at worksheet row 10:
    '   hasHeaders=True  -> first cTable data row came from worksheet row 11
    '   hasHeaders=False -> first cTable data row came from worksheet row 10
    '-------------------------------------------------------------------------
    If IncludeSourceRow And m_RowCount > 0 Then

        firstDataWorksheetRow = Rng.Row

        If hasHeaders Then
            firstDataWorksheetRow = firstDataWorksheetRow + 1
        End If

        p_AddExcelSourceRowColumn _
            firstDataWorksheetRow:=firstDataWorksheetRow, _
            SourceRowColumnName:=SourceRowColumnName
    End If

CleanExit:
    On Error Resume Next
    Application.Calculation = priorCalculation
    Application.ScreenUpdating = priorScreenUpdating
    On Error GoTo 0

    If errNumber <> 0 Then
        Err.Raise errNumber, CALLER, errDescription
    End If

    Exit Sub

ErrHandler:
    errNumber = Err.Number
    errDescription = Err.Description
    Resume CleanExit

End Sub


'==============================================================================
'
' PUBLIC PROPERTIES
'
'==============================================================================

' Number of data rows (excluding headers).
Public Property Get rowCount() As Long
    rowCount = m_RowCount
End Property

' Number of columns.
Public Property Get ColumnCount() As Long
    ColumnCount = m_ColCount
End Property

' 1-based 1-D String array of column names.
Public Property Get headers() As Variant
    If m_ColCount = 0 Then headers = Array(): Exit Property
    Dim h() As String
    ReDim h(1 To m_ColCount)
    Dim i As Long
    For i = 1 To m_ColCount: h(i) = m_Headers(i): Next i
    headers = h
End Property

' 1-based 2-D Variant array of data values (copy).
Public Property Get data() As Variant
    If m_RowCount = 0 Or m_ColCount = 0 Then data = Array(): Exit Property
    Dim d() As Variant
    ReDim d(1 To m_RowCount, 1 To m_ColCount)
    Dim r As Long, c As Long
    For r = 1 To m_RowCount
        For c = 1 To m_ColCount
            d(r, c) = m_Data(r, c)
        Next c
    Next r
    data = d
End Property

' 1-based column number for colName; 0 if not found.
Public Property Get columnNumber(ByVal colName As Variant) As Long
    columnNumber = p_ResolveCol(colName, False)
End Property

' Column name at 1-based column number; "" if out of range.
Public Property Get columnName(ByVal colNumber As Long) As String
    If colNumber < 1 Or colNumber > m_ColCount Then
        columnName = ""
    Else
        columnName = m_Headers(colNumber)
    End If
End Property

'      Internal accessors for Clone / from_array round-trips
Public Property Get HeadersArray() As String()
    HeadersArray = m_Headers
End Property

Public Property Get DataArray() As Variant
    If m_RowCount = 0 Then
        DataArray = Array()
    Else
        DataArray = m_Data
    End If
End Property

' Compatibility shim. With lineage OFF this behaves exactly as V1. With
' lineage ON it returns the table's single source when there is one, and
' Nothing when the table draws on several - a V1 caller then gets an honest
' Nothing instead of a confidently wrong answer.
Public Property Get DataSource() As cDataSource
    If m_Lin Is Nothing Then
        Set DataSource = m_DataSource
        Exit Property
    End If

    Dim ids As Variant
    ids = m_Lin.DistinctSourceIds
    If UBound(ids) - LBound(ids) + 1 <> 1 Then Exit Property
    Set DataSource = cLineageRegistry.SourceById(CLng(ids(LBound(ids))))
End Property

Public Property Set DataSource(ByVal ds As cDataSource)
    Set m_DataSource = ds
End Property

' Internal: lets one cTable hand a lineage engine to another. VBA cannot
' reach another instance's private members, so these must be Public.
Public Property Get x_Lineage() As cLineage
    Set x_Lineage = m_Lin
End Property

Public Sub x_SetLineage(ByVal lin As cLineage)
    Set m_Lin = lin
End Sub

'==============================================================================
'
' LINEAGE - PUBLIC API
'
' Answers "where did the value in this cell come from?"
' See cTable_Lineage_Primer.md.
'
'==============================================================================

' Set this BEFORE a read_* call. Enabling it afterwards seeds every existing
' cell as Unknown, because the read had nothing to record.
Public Property Get EnableLineage() As Boolean
    EnableLineage = Not (m_Lin Is Nothing)
End Property

Public Property Let EnableLineage(ByVal v As Boolean)
    If Not v Then
        Set m_Lin = Nothing
        Exit Property
    End If
    If Not m_Lin Is Nothing Then Exit Property

    Set m_Lin = New cLineage
    m_Lin.InitUniform m_RowCount, m_ColCount, p_LinSrcId(m_DataSource)

    If m_RowCount > 0 And m_DataSource Is Nothing Then
        If m_Log Is Nothing Then Set m_Log = New cLog
        m_Log.AddStatus "cTable lineage enabled after data was loaded; " & _
                        "existing cells report Unknown."
    End If
End Property

' Read-only. Reports the representation actually in use, so it cannot drift
' out of step with reality the way a hand-maintained flag would.
Public Property Get LineageMode() As ctLineageMode
    If m_Lin Is Nothing Then
        LineageMode = ctLinOff
    Else
        LineageMode = m_Lin.mode
    End If
End Property

Public Property Get HasMixedRowLineage() As Boolean
    If Not m_Lin Is Nothing Then HasMixedRowLineage = m_Lin.HasMixedRowSources
End Property

Public Property Get HasMixedColumnLineage() As Boolean
    If Not m_Lin Is Nothing Then HasMixedColumnLineage = m_Lin.HasMixedColumnSources
End Property

' Source attributed to SetValue / SetValueByName. Defaults to Manual.
Public Property Get WriteLineage() As cDataSource
    Set WriteLineage = cLineageRegistry.SourceById(m_WriteSrcId)
End Property

Public Property Set WriteLineage(ByVal ds As cDataSource)
    If ds Is Nothing Then
        m_WriteSrcId = ctSrcManual
    Else
        m_WriteSrcId = cLineageRegistry.SourceId(ds)
    End If
End Property

' O(1). Use this, not LineageOf, inside loops over many cells - it allocates
' nothing and can be compared directly against ctSrcCalculated and friends.
Public Function LineageIdOf(ByVal rowNum As Long, ByVal colNum As Long) As Long
    If m_Lin Is Nothing Then
        LineageIdOf = ctSrcUnknown
    Else
        LineageIdOf = m_Lin.IdAt(rowNum, colNum)
    End If
End Function

Public Function LineageOf(ByVal rowNum As Long, ByVal colNum As Long) As cDataSource
    Set LineageOf = cLineageRegistry.SourceById(LineageIdOf(rowNum, colNum))
End Function

Public Function LineageOfName(ByVal rowNum As Long, _
                              ByVal colName As Variant) As cDataSource
    colName = p_LabelColumnRef(colName)
    Dim ci As Long: ci = p_ColIdx(colName)
    If ci = 0 Then
        Err.Raise vbObjectError + 1320, "cTable.LineageOfName", _
                  "Column not found: """ & colName & """"
    End If
    Set LineageOfName = LineageOf(rowNum, ci)
End Function

' One-line answer suitable for showing a user.
Public Function LineageExplain(ByVal rowNum As Long, _
                               ByVal colName As Variant) As String
    colName = p_LabelColumnRef(colName)
    Dim ci As Long: ci = p_ColIdx(colName)
    If ci = 0 Then
        LineageExplain = "Column not found: """ & colName & """"
        Exit Function
    End If
    If m_Lin Is Nothing Then
        LineageExplain = "Row " & rowNum & ", column """ & colName & _
                         """ -> lineage tracking is disabled"
        Exit Function
    End If
    LineageExplain = "Row " & rowNum & ", column """ & colName & """ -> " & _
                     cLineageRegistry.SourceLabel(m_Lin.IdAt(rowNum, ci))
End Function

' Returns Nothing when the column's rows do not share a single source, which
' after a concat is normal.
Public Function ColumnLineage(ByVal colName As Variant) As cDataSource
    colName = p_LabelColumnRef(colName)
    If m_Lin Is Nothing Then Exit Function
    Dim ci As Long: ci = p_ColIdx(colName)
    If ci = 0 Then Exit Function

    Dim id As Long
    id = m_Lin.ColumnUniformSource(ci)
    If id = 0 Then Exit Function
    Set ColumnLineage = cLineageRegistry.SourceById(id)
End Function

' 1-based array of the distinct cDataSource OBJECTS feeding this table.
' Empty array when lineage is off. Was named LineageSources before 20260727.
Public Function LineageSourceObjects() As Variant
    LineageSourceObjects = Array()
    If m_Lin Is Nothing Then Exit Function

    Dim ids As Variant
    ids = m_Lin.DistinctSourceIds
    If Not p_LinHasIds(ids) Then Exit Function

    Dim n As Long
    n = UBound(ids) - LBound(ids) + 1
    Dim Out() As cDataSource
    ReDim Out(1 To n)
    Dim i As Long
    For i = 1 To n
        Set Out(i) = cLineageRegistry.SourceById(CLng(ids(LBound(ids) + i - 1)))
    Next i
    LineageSourceObjects = Out
End Function

'------------------------------------------------------------------------------
' TOTAL DIAGNOSTICS
'
' Every member below is a total function. With lineage disabled, with no rows,
' or with a column name that does not exist, they return "", 0, or an empty
' array. None of them returns Nothing and none of them raises, so they are all
' safe to type straight into the Immediate window.
'
' ColumnLineage is the ONE member of the lineage API that can hand back
' Nothing. Use ColumnLineageLabel or ColumnSourceCount unless you are
' deliberately branching on cardinality.
'------------------------------------------------------------------------------

' "off", "uniform", "by column", "by row", or "full".
Public Property Get LineageModeName() As String
    If m_Lin Is Nothing Then
        LineageModeName = "off"
    Else
        LineageModeName = m_Lin.ModeName
    End If
End Property

' Row-storage form, e.g. "uniform (1 map)", "banded (3 bands)", "dense (8.0 MB)".
' Empty string when lineage is off.
Public Property Get LineageStoreName() As String
    If Not m_Lin Is Nothing Then LineageStoreName = m_Lin.StoreName
End Property

' Number of distinct sources anywhere in the table. 0 when lineage is off.
Public Property Get LineageSourceCount() As Long
    If m_Lin Is Nothing Then Exit Property
    Dim ids As Variant
    ids = m_Lin.DistinctSourceIds
    If Not p_LinHasIds(ids) Then Exit Property
    LineageSourceCount = UBound(ids) - LBound(ids) + 1
End Property

' 1-based array of source LABEL strings. Empty array when lineage is off.
Public Property Get LineageSources() As Variant
    LineageSources = Array()
    If m_Lin Is Nothing Then Exit Property

    Dim ids As Variant
    ids = m_Lin.DistinctSourceIds
    If Not p_LinHasIds(ids) Then Exit Property

    Dim n As Long
    n = UBound(ids) - LBound(ids) + 1
    Dim Out() As String
    ReDim Out(1 To n)
    Dim i As Long
    For i = 1 To n
        Out(i) = cLineageRegistry.SourceLabel(CLng(ids(LBound(ids) + i - 1)))
    Next i
    LineageSources = Out
End Property

' Source label for a column, or a parenthesised explanation. Never raises:
'   "(lineage disabled)"      tracking is off
'   "(no such column: X)"     name not found
'   "(no rows)"               nothing to attribute
'   "(mixed: N sources)"      the column draws on more than one origin
Public Function ColumnLineageLabel(ByVal colName As Variant) As String
    colName = p_LabelColumnRef(colName)
    If m_Lin Is Nothing Then
        ColumnLineageLabel = "(lineage disabled)"
        Exit Function
    End If

    Dim ci As Long: ci = p_ColIdx(colName)
    If ci = 0 Then
        ColumnLineageLabel = "(no such column: " & colName & ")"
        Exit Function
    End If
    If m_RowCount = 0 Then
        ColumnLineageLabel = "(no rows)"
        Exit Function
    End If

    Dim id As Long
    id = m_Lin.ColumnUniformSource(ci)
    If id <> 0 Then
        ColumnLineageLabel = cLineageRegistry.SourceLabel(id)
    Else
        ColumnLineageLabel = "(mixed: " & ColumnSourceCount(colName) & " sources)"
    End If
End Function

' 0 = lineage off, no such column, or no rows. 1 = single source. 2+ = mixed.
Public Function ColumnSourceCount(ByVal colName As Variant) As Long
    colName = p_LabelColumnRef(colName)
    If m_Lin Is Nothing Or m_RowCount = 0 Then Exit Function
    Dim ci As Long: ci = p_ColIdx(colName)
    If ci = 0 Then Exit Function

    If m_Lin.ColumnUniformSource(ci) <> 0 Then
        ColumnSourceCount = 1
        Exit Function
    End If

    Dim ids As Variant
    ids = m_Lin.ColumnDistinctSourceIds(ci)
    If Not p_LinHasIds(ids) Then Exit Function
    ColumnSourceCount = UBound(ids) - LBound(ids) + 1
End Function

' Per-column breakdown, expanding any column that draws on several sources.
Public Function LineageColumnReport() As String
    If m_Lin Is Nothing Then
        LineageColumnReport = "cTable lineage: disabled"
        Exit Function
    End If

    Dim sb As String, c As Long, i As Long
    sb = "Lineage by column  [" & m_RowCount & " rows, " & _
         LineageModeName & ", " & LineageStoreName & "]" & vbCrLf

    For c = 1 To m_ColCount
        sb = sb & "  " & m_Headers(c) & ": " & _
             ColumnLineageLabel(m_Headers(c)) & vbCrLf

        If ColumnSourceCount(m_Headers(c)) > 1 Then
            Dim ids As Variant
            ids = m_Lin.ColumnDistinctSourceIds(c)
            For i = LBound(ids) To UBound(ids)
                sb = sb & "      - " & _
                     cLineageRegistry.SourceLabel(CLng(ids(i))) & vbCrLf
            Next i
        End If
    Next c
    LineageColumnReport = sb
End Function

Public Function LineageReport() As String
    If m_Lin Is Nothing Then
        LineageReport = "cTable lineage: disabled"
    Else
        LineageReport = m_Lin.Report
    End If
End Function

' Assert lineage yourself when you know something the class does not.
Public Sub StampLineage(ByVal colName As Variant, ByVal ds As cDataSource)
    colName = p_LabelColumnRef(colName)
    If m_Lin Is Nothing Then Exit Sub
    Dim ci As Long: ci = p_ColIdx(colName)
    If ci = 0 Then
        Err.Raise vbObjectError + 1321, "cTable.StampLineage", _
                  "Column not found: """ & colName & """"
    End If
    p_LinSync
    m_Lin.StampColumn ci, p_LinSrcId(ds)
End Sub

Public Sub SetCellLineage(ByVal rowNum As Long, ByVal colNum As Long, _
                          ByVal ds As cDataSource)
    If m_Lin Is Nothing Then Exit Sub
    p_LinSync
    m_Lin.SetCell rowNum, colNum, p_LinSrcId(ds)
End Sub

Public Property Get Log() As cLog
    If m_Log Is Nothing Then Set m_Log = New cLog
    Set Log = m_Log
End Property

' Returns a 1-based 2-element array: (RowCount, ColumnCount)
Public Property Get shape() As Variant
    Dim shp(1 To 2) As Long
    shp(1) = m_RowCount
    shp(2) = m_ColCount
    shape = shp
End Property

'==============================================================================
'
' CELL-LEVEL ACCESS
'
'==============================================================================

' Get / Set by 1-based row and column number.
Public Function GetValue(ByVal rowNum As Long, ByVal col As Variant) As Variant
    Dim ci As Long
    ci = p_ResolveCol(col, False)
    If rowNum < 1 Or rowNum > m_RowCount Or ci = 0 Then
        GetValue = Null
        Exit Function
    End If
    GetValue = m_Data(rowNum, ci)
End Function

Public Sub SetValue(ByVal rowNum As Long, _
                    ByVal col As Variant, _
                    ByVal val As Variant)
    Dim ci As Long
    ci = p_ResolveCol(col, False)
    If rowNum < 1 Or rowNum > m_RowCount Or ci = 0 Then
        Err.Raise vbObjectError + 10, "cTable.SetValue", "Index out of bounds"
    End If
    m_Data(rowNum, ci) = val
    If Not m_Lin Is Nothing Then m_Lin.SetCell rowNum, ci, m_WriteSrcId
    p_ClearIndexes
End Sub

' Compatibility facades preserve original name-only and missing-column rules.
Public Function GetValueByName(ByVal rowNum As Long, ByVal colName As String) As Variant
    Dim ci As Long: ci = p_ColIdx(colName)
    If ci = 0 Then GetValueByName = Null: Exit Function
    GetValueByName = m_Data(rowNum, ci)
End Function

Public Sub SetValueByName(ByVal rowNum As Long, ByVal colName As String, ByVal val As Variant)
    Dim ci As Long: ci = p_ColIdx(colName)
    If ci = 0 Then
        Err.Raise vbObjectError + 11, "cTable.SetValueByName", _
                  "Column not found: """ & colName & """"
    End If
    m_Data(rowNum, ci) = val
    If Not m_Lin Is Nothing Then m_Lin.SetCell rowNum, ci, m_WriteSrcId
    p_ClearIndexes
End Sub

'==============================================================================
'
' BUILD INDEX
'
'==============================================================================

' Creates a hash index on colName for fast lookups via index_lookup().
Public Sub set_index(ByVal colName As Variant, ByVal indexName As String)
    colName = p_LabelColumnRef(colName)
    Dim ci As Long: ci = p_ColIdx(colName)
    If ci = 0 Then
        Err.Raise vbObjectError + 20, "cTable.set_index", _
                  "Column not found: """ & colName & """"
    End If

    Dim idx As Object
    Set idx = CreateObject("Scripting.Dictionary")
    idx.CompareMode = vbTextCompare

    Dim i As Long
    For i = 1 To m_RowCount
        Dim key As String
        key = CStr(m_Data(i, ci))
        If Not idx.Exists(key) Then Set idx(key) = New Collection
        idx(key).Add i
    Next i
    m_Indexes(indexName) = idx
End Sub

' Returns a 0-based Long array of row numbers matching key in the named index.
' Returns Empty if indexName or key is not found.
Public Function index_lookup(ByVal indexName As String, _
                             ByVal key As String) As Variant
    If Not m_Indexes.Exists(indexName) Then index_lookup = Empty: Exit Function
    Dim idx As Object: Set idx = m_Indexes(indexName)
    If Not idx.Exists(key) Then index_lookup = Empty: Exit Function

    Dim col As Collection: Set col = idx(key)
    Dim result() As Long
    ReDim result(0 To col.count - 1)
    Dim j As Long
    For j = 0 To col.count - 1
        result(j) = col(j + 1)
    Next j
    index_lookup = result
End Function

'==============================================================================
'
' WHERE  (row filter)
'
'==============================================================================

Private Function p_NullPolicy(ByVal nullBehavior As String) As Long
    Select Case LCase$(Trim$(nullBehavior))
        Case "exclude": p_NullPolicy = 0
        Case "include": p_NullPolicy = 1
        Case Else
            Err.Raise vbObjectError + 1704, "cTable.Where", _
                "nullBehavior must be exclude or include."
    End Select
End Function

' One scalar predicate, used by both positive and inverse filters.
' Nonmissing comparisons retain the original Where numeric/text behavior.
Private Function p_WherePredicate(ByVal cell As Variant, _
                                  ByVal value As Variant, _
                                  ByVal op As String) As Boolean
    Dim cellStr As String, sVal As String
    Dim numericCell As Boolean, numericVal As Boolean
    cellStr = CStr(cell)
    sVal = CStr(value)
    numericCell = IsNumeric(cell)
    numericVal = IsNumeric(value)
    Select Case op
        Case "eq"
            If numericCell And numericVal Then
                p_WherePredicate = (CDbl(cell) = CDbl(value))
            Else
                p_WherePredicate = (StrComp(cellStr, sVal, vbTextCompare) = 0)
            End If
        Case "noteq"
            If numericCell And numericVal Then
                p_WherePredicate = (CDbl(cell) <> CDbl(value))
            Else
                p_WherePredicate = (StrComp(cellStr, sVal, vbTextCompare) <> 0)
            End If
        Case "startswith"
            p_WherePredicate = (StrComp(Left$(cellStr, Len(sVal)), sVal, vbTextCompare) = 0)
        Case "endswith"
            p_WherePredicate = (StrComp(right$(cellStr, Len(sVal)), sVal, vbTextCompare) = 0)
        Case "contains"
            p_WherePredicate = (InStr(1, cellStr, sVal, vbTextCompare) > 0)
        Case "gt", "gte", "lt", "lte"
            Dim cmp As Long
            If numericCell And numericVal Then
                If CDbl(cell) < CDbl(value) Then
                    cmp = -1
                ElseIf CDbl(cell) > CDbl(value) Then
                    cmp = 1
                End If
            Else
                ' Original Where used VBA's relational string comparison.
                If cellStr < sVal Then
                    cmp = -1
                ElseIf cellStr > sVal Then
                    cmp = 1
                End If
            End If
            Select Case op
                Case "gt": p_WherePredicate = cmp > 0
                Case "gte": p_WherePredicate = cmp >= 0
                Case "lt": p_WherePredicate = cmp < 0
                Case "lte": p_WherePredicate = cmp <= 0
            End Select
    End Select
End Function

Private Function p_FilterRows(ByVal columnRef As Variant, _
                              ByVal query As Variant, _
                              ByVal op As String, _
                              ByVal invert As Boolean, _
                              ByVal membership As Boolean, _
                              ByVal nullBehavior As String) As cTable

    Dim ci As Long
    ci = p_ResolveCol(columnRef)

    Dim nullMode As Long
    nullMode = p_NullPolicy(nullBehavior)

    Dim opCode As String
    opCode = LCase$(Trim$(op))

    Dim lookup As Object
    Dim queryMissing As Boolean
    Dim membershipIncludesNA As Boolean

    '-------------------------------------------------------------------------
    ' Validate and prepare either a membership lookup or a scalar predicate.
    '-------------------------------------------------------------------------
    If membership Then

        If Not IsArray(query) Then
            Err.Raise vbObjectError + 35, _
                      "cTable.WhereIn", _
                      "values must be an array."
        End If

        If Not p_ArrayHasElements(query) Then
            Err.Raise vbObjectError + 36, _
                      "cTable.WhereIn", _
                      "values cannot be empty."
        End If

        Set lookup = CreateObject("Scripting.Dictionary")
        lookup.CompareMode = vbTextCompare

        Dim k As Long

        For k = LBound(query) To UBound(query)
            If p_IsNA(query(k)) Then
                lookup(Chr$(1) & "NA") = True
            Else
                lookup(CStr(query(k))) = True
            End If
        Next k

        ' Important: only evaluate lookup.Exists after lookup is initialized.
        membershipIncludesNA = lookup.Exists(Chr$(1) & "NA")

    Else

        Select Case opCode
            Case "eq", "noteq", "startswith", "endswith", "contains", _
                 "gt", "gte", "lt", "lte"
                ' Valid operator.

            Case Else
                Err.Raise vbObjectError + 31, _
                          "cTable.Where", _
                          "Unknown operator: " & op
        End Select

        queryMissing = p_IsNA(query)

        If queryMissing _
           And opCode <> "eq" _
           And opCode <> "noteq" Then

            Err.Raise vbObjectError + 1705, _
                      "cTable.Where", _
                      "A missing search value requires eq or noteq."
        End If
    End If

    '-------------------------------------------------------------------------
    ' Select source row numbers.
    '-------------------------------------------------------------------------
    Dim selected() As Long
    ReDim selected(1 To IIf(m_RowCount > 0, m_RowCount, 1))

    Dim r As Long
    Dim n As Long
    Dim hit As Boolean
    Dim missing As Boolean

    For r = 1 To m_RowCount

        missing = p_IsNA(m_Data(r, ci))
        hit = False

        '---------------------------------------------------------------------
        ' Apply nullBehavior only when:
        '   * the current cell is missing;
        '   * the scalar query is not missing; and
        '   * a membership query did not explicitly request missing values.
        '
        ' membershipIncludesNA is a Boolean computed above. It is False for
        ' ordinary Where / WhereNot calls, so no lookup object is accessed here.
        '---------------------------------------------------------------------
        If missing _
           And Not queryMissing _
           And Not membershipIncludesNA Then

            ' Missing-cell policy is independent of predicate inversion.
            ' An explicit missing search is handled in the Else branch.
            Select Case nullMode
                Case 0      ' exclude
                    hit = False

                Case 1      ' include
                    hit = True
            End Select

        Else

            If membership Then

                If missing Then
                    hit = lookup.Exists(Chr$(1) & "NA")
                Else
                    hit = lookup.Exists(CStr(m_Data(r, ci)))
                End If

            ElseIf queryMissing Then

                hit = missing

                If opCode = "noteq" Then
                    hit = Not hit
                End If

            Else

                hit = p_WherePredicate( _
                            m_Data(r, ci), _
                            query, _
                            opCode)
            End If

            If invert Then
                hit = Not hit
            End If
        End If

        If hit Then
            n = n + 1
            selected(n) = r
        End If
    Next r

    '-------------------------------------------------------------------------
    ' Build and return the filtered result.
    '-------------------------------------------------------------------------
    Dim result As cTable
    Set result = p_BuildFromRows(selected, n)

    Set result.DataSource = Me.DataSource
    result.set_dtypes_from Me

    Set p_FilterRows = result

End Function


Public Function Where(ByVal col As Variant, ByVal value As Variant, _
                      Optional ByVal operatorType As String = "eq", _
                      Optional ByVal nullBehavior As String = "exclude") As cTable
    Set Where = p_FilterRows(col, value, operatorType, False, False, nullBehavior)
End Function

Public Function WhereNot(ByVal col As Variant, ByVal value As Variant, _
                         Optional ByVal operatorType As String = "eq", _
                         Optional ByVal nullBehavior As String = "exclude") As cTable
    Set WhereNot = p_FilterRows(col, value, operatorType, True, False, nullBehavior)
End Function

Public Function WhereIn(ByVal col As Variant, ByVal values As Variant, _
                        Optional ByVal nullBehavior As String = "exclude") As cTable
    Set WhereIn = p_FilterRows(col, values, "", False, True, nullBehavior)
End Function

Public Function WhereNotIn(ByVal col As Variant, ByVal values As Variant, _
                           Optional ByVal nullBehavior As String = "exclude") As cTable
    Set WhereNotIn = p_FilterRows(col, values, "", True, True, nullBehavior)
End Function

'==============================================================================
'
' SELECT COLUMNS
'
'==============================================================================

Public Function SelectColumns(ByVal COLS As Variant) As cTable
    Dim numCols As Long
    numCols = UBound(COLS) - LBound(COLS) + 1

    Dim colIdx() As Long
    Dim newHdrs() As String
    ReDim colIdx(1 To numCols)
    ReDim newHdrs(1 To numCols)

    Dim k As Long
    For k = 1 To numCols
        Dim cName As String
        cName = p_LabelColumnRef(COLS(LBound(COLS) + k - 1))
        Dim ci As Long: ci = p_ColIdx(cName)
        If ci = 0 Then
            Err.Raise vbObjectError + 40, "cTable.SelectColumns", _
                      "Column not found: """ & cName & """"
        End If
        colIdx(k) = ci
        newHdrs(k) = m_Headers(ci)
    Next k

    Set SelectColumns = New cTable
    Dim emptyPH() As Variant
    ReDim emptyPH(1 To 1, 1 To numCols)

    If m_RowCount = 0 Then
        SelectColumns.from_array newHdrs, emptyPH, 0, numCols
        Set SelectColumns.DataSource = Me.DataSource
        SelectColumns.set_dtypes_from Me
        If Not m_Lin Is Nothing Then
            SelectColumns.x_SetLineage m_Lin.ProjectColumns(colIdx, numCols)
        End If
        Exit Function
    End If

    Dim resData() As Variant
    ReDim resData(1 To m_RowCount, 1 To numCols)
    Dim r As Long, c As Long
    For r = 1 To m_RowCount
        For c = 1 To numCols
            resData(r, c) = m_Data(r, colIdx(c))
        Next c
    Next r
    SelectColumns.from_array newHdrs, resData, m_RowCount, numCols
    Set SelectColumns.DataSource = Me.DataSource
    SelectColumns.set_dtypes_from Me
    If Not m_Lin Is Nothing Then
        SelectColumns.x_SetLineage m_Lin.ProjectColumns(colIdx, numCols)
    End If
End Function

'==============================================================================
'
' JOINS
'
'==============================================================================

' ---------------------------------------------------------------------------
' MakeJoin  (legacy)
'
' Retained unchanged for existing callers. New code should prefer merge, which
' supports multi-column keys, right / outer / cross joins, and pandas naming.
'
'   Set j = t.MakeJoin("left", other, "AccountID", "AccountID_Ref")
' ---------------------------------------------------------------------------
Public Function MakeJoin(ByVal JoinType As String, right As cTable, _
                         ByVal leftKey As String, ByVal rightKey As String, _
                         Optional LeftAsNames As Object = Nothing, _
                         Optional RightAsNames As Object = Nothing) As cTable

    Dim howCode As Long

    Select Case LCase(JoinType)
        Case "inner": howCode = ctJoinInner
        Case "left":  howCode = ctJoinLeft
        Case Else
            Err.Raise vbObjectError + 120, "cTable.MakeJoin", _
                      "Unknown join type: """ & JoinType & """"
    End Select

    Dim lKeys() As Long
    Dim rKeys() As Long
    ReDim lKeys(1 To 1)
    ReDim rKeys(1 To 1)

    lKeys(1) = p_ColIdx(leftKey)
    If lKeys(1) = 0 Then
        Err.Raise vbObjectError + 51, "cTable.Join", _
                  "Left key column not found: """ & leftKey & """"
    End If

    rKeys(1) = right.columnNumber(rightKey)
    If rKeys(1) = 0 Then
        Err.Raise vbObjectError + 52, "cTable.Join", _
                  "Right key column not found: """ & rightKey & """"
    End If

    Set MakeJoin = p_MergeCore(right, lKeys, rKeys, 1, howCode, _
                               True, ctNameLegacy, "", "", _
                               False, "", ctValidateNone, False, _
                               LeftAsNames, RightAsNames, "cTable.MakeJoin")
End Function

'==============================================================================
'
' MERGE
'
' One engine, two facades: merge is the pandas-shaped public entry point and
' MakeJoin is the legacy one. Both call p_MergeCore, so there is a single
' implementation of matching, column layout and lineage.
'
' LINEAGE
'   Left columns carry left provenance, right columns carry right provenance,
'   and an unmatched side reports ctSrcMissing rather than guessing. That is
'   cLineage.x_InitJoin's job; Patch 4a made its treatment of a missing LEFT
'   side symmetric with its existing treatment of a missing right side.
'
'   One subtlety the engine cannot see: when the key columns are MERGED (the
'   pandas `on` form, and every MakeJoin), a right-only row's key column is
'   populated from the RIGHT table even though it sits in a left column
'   position. x_InitJoin would report ctSrcMissing there, which would be a
'   lie - the value is real. p_MergeCore patches those cells afterwards with
'   the right table's own ids.
'
'==============================================================================




' ---------------------------------------------------------------------------
' One key component, rendered as text.
'
' Deliberately the same permissive rule p_DoJoin used - plain CStr - with two
' repairs: Null becomes an explicit sentinel instead of raising error 94, and
' a Date becomes its serial number instead of a locale-dependent string.
' ---------------------------------------------------------------------------
Private Function p_JoinKeyPart(ByVal v As Variant) As String
    If p_IsNA(v) Then
        p_JoinKeyPart = Chr$(1) & "NA"
    ElseIf VarType(v) = vbDate Then
        p_JoinKeyPart = Chr$(1) & "D" & Trim$(Str$(CDbl(CDate(v))))
    Else
        p_JoinKeyPart = CStr(v)
    End If
End Function


' Composite key for one row of any cTable. Length-prefixed so ("AB","C") and
' ("A","BC") cannot collide.
Private Function p_JoinRowKey(ByVal tbl As cTable, ByVal rowNum As Long, _
                              ByRef COLS() As Long, ByVal n As Long) As String
    Dim s As String
    Dim part As String
    Dim i As Long

    For i = 1 To n
        part = p_JoinKeyPart(tbl.GetValue(rowNum, COLS(i)))
        s = s & CStr(Len(part)) & ":" & part & "|"
    Next i

    p_JoinRowKey = s
End Function


' key -> Collection of row numbers.
Private Function p_JoinIndex(ByVal tbl As cTable, ByRef COLS() As Long, _
                             ByVal n As Long) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare

    Dim i As Long
    Dim k As String

    For i = 1 To tbl.rowCount
        k = p_JoinRowKey(tbl, i, COLS, n)
        If Not d.Exists(k) Then Set d(k) = New Collection
        d(k).Add i
    Next i

    Set p_JoinIndex = d
End Function


' True when no key combination repeats.
Private Function p_KeysAreUnique(ByVal tbl As cTable, ByRef COLS() As Long, _
                                 ByVal n As Long) As Boolean
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare

    Dim i As Long
    Dim k As String

    For i = 1 To tbl.rowCount
        k = p_JoinRowKey(tbl, i, COLS, n)
        If d.Exists(k) Then Exit Function
        d(k) = 1
    Next i

    p_KeysAreUnique = True
End Function


Private Sub p_AppendPair(ByRef jLeft() As Long, ByRef jRight() As Long, _
                         ByRef n As Long, ByRef cap As Long, _
                         ByVal li As Long, ByVal ri As Long)
    n = n + 1
    If n > cap Then
        cap = cap * 2
        ReDim Preserve jLeft(1 To cap)
        ReDim Preserve jRight(1 To cap)
    End If
    jLeft(n) = li
    jRight(n) = ri
End Sub


' ---------------------------------------------------------------------------
' The engine. Everything public routes through here.
'
' lKeys / rKeys   1-based key column numbers on each side, pairwise
' mergeKeys       True drops the right key columns and coalesces a right-only
'                 row's key values up into the left key columns
' naming          ctNamePandas or ctNameLegacy
' ---------------------------------------------------------------------------
Private Function p_MergeCore(ByVal right As cTable, _
                             ByRef lKeys() As Long, ByRef rKeys() As Long, _
                             ByVal keyCount As Long, _
                             ByVal howCode As Long, _
                             ByVal mergeKeys As Boolean, _
                             ByVal naming As Long, _
                             ByVal suffixLeft As String, _
                             ByVal suffixRight As String, _
                             ByVal indicator As Boolean, _
                             ByVal indicatorName As String, _
                             ByVal validateCode As Long, _
                             ByVal sortResult As Boolean, _
                             ByVal leftRename As Object, _
                             ByVal rightRename As Object, _
                             ByVal callerName As String) As cTable

    p_EnsureTypes

    Dim i As Long, j As Long, c As Long, li As Long, ri As Long

    ' --- validate ---------------------------------------------------------
    If validateCode <> ctValidateNone And howCode <> ctJoinCross Then
        Select Case validateCode
            Case ctValidateOneToOne, ctValidateOneToMany
                If Not p_KeysAreUnique(Me, lKeys, keyCount) Then
                    Err.Raise vbObjectError + 1360, callerName, _
                              "validate failed: the left key is not unique."
                End If
        End Select

        Select Case validateCode
            Case ctValidateOneToOne, ctValidateManyToOne
                If Not p_KeysAreUnique(right, rKeys, keyCount) Then
                    Err.Raise vbObjectError + 1361, callerName, _
                              "validate failed: the right key is not unique."
                End If
        End Select
    End If

    ' --- which right columns are emitted ----------------------------------
    Dim rColCount As Long
    rColCount = right.ColumnCount

    Dim dropRight() As Boolean
    ReDim dropRight(1 To IIf(rColCount > 0, rColCount, 1))

    If mergeKeys Then
        For j = 1 To keyCount
            dropRight(rKeys(j)) = True
        Next j
    End If

    Dim rSrcCol() As Long
    ReDim rSrcCol(1 To IIf(rColCount > 0, rColCount, 1))

    Dim rUsed As Long
    For c = 1 To rColCount
        If Not dropRight(c) Then
            rUsed = rUsed + 1
            rSrcCol(rUsed) = c
        End If
    Next c

    Dim totalCols As Long
    totalCols = m_ColCount + rUsed
    If indicator Then totalCols = totalCols + 1

    ' --- headers ----------------------------------------------------------
    Dim hdrs() As String
    ReDim hdrs(1 To totalCols)

    For i = 1 To m_ColCount
        hdrs(i) = m_Headers(i)
    Next i

    If naming = ctNamePandas Then
        ' Overlapping non-key names are suffixed on BOTH sides, as pandas does.
        Dim clash As Boolean
        For c = 1 To rUsed
            clash = False
            For i = 1 To m_ColCount
                If StrComp(m_Headers(i), right.columnName(rSrcCol(c)), _
                           vbTextCompare) = 0 Then
                    clash = True
                    hdrs(i) = m_Headers(i) & suffixLeft
                    Exit For
                End If
            Next i

            If clash Then
                hdrs(m_ColCount + c) = right.columnName(rSrcCol(c)) & suffixRight
            Else
                hdrs(m_ColCount + c) = right.columnName(rSrcCol(c))
            End If
        Next c
    Else
        ' Legacy: right side only, _1 / _2 ..., checked against LEFT headers.
        Dim baseName As String
        Dim newName As String
        Dim suffixNo As Long
        Dim collision As Boolean

        For c = 1 To rUsed
            baseName = right.columnName(rSrcCol(c))
            newName = baseName

            collision = False
            For i = 1 To m_ColCount
                If StrComp(m_Headers(i), newName, vbTextCompare) = 0 Then
                    collision = True
                    Exit For
                End If
            Next i

            If collision Then
                suffixNo = 1
                Do
                    newName = baseName & "_" & suffixNo
                    suffixNo = suffixNo + 1
                    collision = False
                    For i = 1 To m_ColCount
                        If StrComp(m_Headers(i), newName, vbTextCompare) = 0 Then
                            collision = True
                            Exit For
                        End If
                    Next i
                Loop While collision
            End If

            hdrs(m_ColCount + c) = newName
        Next c
    End If

    ' Optional caller-supplied renaming, applied last so it wins.
    If Not leftRename Is Nothing Then
        For i = 1 To m_ColCount
            If leftRename.Exists(hdrs(i)) Then hdrs(i) = CStr(leftRename(hdrs(i)))
        Next i
    End If

    If Not rightRename Is Nothing Then
        For c = 1 To rUsed
            If rightRename.Exists(hdrs(m_ColCount + c)) Then
                hdrs(m_ColCount + c) = CStr(rightRename(hdrs(m_ColCount + c)))
            End If
        Next c
    End If

    If indicator Then hdrs(totalCols) = indicatorName

    ' --- pair up the rows -------------------------------------------------
    Dim jLeft() As Long, jRight() As Long
    Dim n As Long, cap As Long
    cap = 1024
    ReDim jLeft(1 To cap)
    ReDim jRight(1 To cap)

    Dim idx As Object
    Dim matches As Collection
    Dim k As String
    Dim m As Long

    Select Case howCode

        Case ctJoinCross
            For i = 1 To m_RowCount
                For j = 1 To right.rowCount
                    p_AppendPair jLeft, jRight, n, cap, i, j
                Next j
            Next i

        Case ctJoinRight
            ' Index the LEFT and walk the RIGHT, so the result keeps the
            ' right table's row order, as pandas how="right" does.
            Set idx = p_JoinIndex(Me, lKeys, keyCount)

            For j = 1 To right.rowCount
                k = p_JoinRowKey(right, j, rKeys, keyCount)

                If idx.Exists(k) Then
                    Set matches = idx(k)
                    For m = 1 To matches.count
                        p_AppendPair jLeft, jRight, n, cap, matches(m), j
                    Next m
                Else
                    p_AppendPair jLeft, jRight, n, cap, 0, j
                End If
            Next j

        Case Else
            Set idx = p_JoinIndex(right, rKeys, keyCount)

            Dim rMatched() As Boolean
            ReDim rMatched(1 To IIf(right.rowCount > 0, right.rowCount, 1))

            For i = 1 To m_RowCount
                k = p_JoinRowKey(Me, i, lKeys, keyCount)

                If idx.Exists(k) Then
                    Set matches = idx(k)
                    For m = 1 To matches.count
                        p_AppendPair jLeft, jRight, n, cap, i, matches(m)
                        rMatched(matches(m)) = True
                    Next m
                ElseIf howCode = ctJoinLeft Or howCode = ctJoinOuter Then
                    p_AppendPair jLeft, jRight, n, cap, i, 0
                End If
            Next i

            If howCode = ctJoinOuter Then
                For j = 1 To right.rowCount
                    If Not rMatched(j) Then
                        p_AppendPair jLeft, jRight, n, cap, 0, j
                    End If
                Next j
            End If
    End Select

    ' --- data -------------------------------------------------------------
    Dim outData() As Variant
    ReDim outData(1 To IIf(n > 0, n, 1), 1 To totalCols)

    For i = 1 To n
        li = jLeft(i)
        ri = jRight(i)

        For c = 1 To m_ColCount
            If li = 0 Then
                outData(i, c) = Null
            Else
                outData(i, c) = m_Data(li, c)
            End If
        Next c

        ' A merged key column on a right-only row takes the right's value.
        If mergeKeys And li = 0 And ri > 0 Then
            For j = 1 To keyCount
                outData(i, lKeys(j)) = right.GetValue(ri, rKeys(j))
            Next j
        End If

        For c = 1 To rUsed
            If ri = 0 Then
                outData(i, m_ColCount + c) = Null
            Else
                outData(i, m_ColCount + c) = right.GetValue(ri, rSrcCol(c))
            End If
        Next c

        If indicator Then
            If li = 0 Then
                outData(i, totalCols) = "right_only"
            ElseIf ri = 0 Then
                outData(i, totalCols) = "left_only"
            Else
                outData(i, totalCols) = "both"
            End If
        End If
    Next i

    ' --- assemble ---------------------------------------------------------
    Dim result As cTable
    Set result = New cTable
    result.from_array hdrs, outData, n, totalCols
    Set result.DataSource = Me.DataSource

    ' Types are carried POSITIONALLY. set_dtypes_from matches by name, which
    ' silently dropped the dtype of every renamed column.
    Dim outTypes() As ctDataType
    ReDim outTypes(1 To totalCols)

    For i = 1 To m_ColCount
        outTypes(i) = m_ColTypes(i)
    Next i
    For c = 1 To rUsed
        outTypes(m_ColCount + c) = right.dtype_at(rSrcCol(c))
    Next c
    If indicator Then outTypes(totalCols) = ctString

    result.dtypes_array = outTypes

    ' --- lineage ----------------------------------------------------------
    If Not m_Lin Is Nothing Then
        Dim rightLin As cLineage
        Set rightLin = p_LinOther(right)

        Dim jl As cLineage
        Set jl = New cLineage
        jl.x_InitJoin m_Lin, rightLin, jLeft, jRight, n, rSrcCol, rUsed

        ' x_InitJoin marks every left column of a right-only row as Missing.
        ' For a MERGED key column that is wrong: the value is real and came
        ' from the right table. Patch those cells with the right's own ids.
        If mergeKeys And n > 0 Then
            Dim anyRightOnly As Boolean
            For i = 1 To n
                If jLeft(i) = 0 Then
                    anyRightOnly = True
                    Exit For
                End If
            Next i

            If anyRightOnly Then
                Dim ids() As Long
                ReDim ids(1 To n)

                For j = 1 To keyCount
                    For i = 1 To n
                        If jLeft(i) = 0 And jRight(i) > 0 Then
                            ids(i) = rightLin.IdAt(jRight(i), rKeys(j))
                            If ids(i) = 0 Then ids(i) = ctSrcUnknown
                        Else
                            ids(i) = jl.IdAt(i, lKeys(j))
                        End If
                    Next i

                    p_LinSetColumnIds jl, lKeys(j), ids, n
                Next j
            End If
        End If

        If indicator Then
            jl.x_InsertColumnAt totalCols, _
                cLineageRegistry.CalculatedSourceId("merge indicator")
        End If

        result.x_SetLineage jl
    End If

    ' --- optional sort by the join keys ------------------------------------
    If sortResult And n > 0 And keyCount > 0 Then
        Dim sortNames() As String
        ReDim sortNames(0 To keyCount - 1)
        For j = 1 To keyCount
            sortNames(j - 1) = hdrs(lKeys(j))
        Next j
        result.sort_values sortNames, inplace:=True
    End If

    Set p_MergeCore = result
End Function


' Map a pandas how string to a code.
Private Function p_JoinHowCode(ByVal how As String) As Long
    Select Case LCase$(Trim$(how))
        Case "inner": p_JoinHowCode = ctJoinInner
        Case "left":  p_JoinHowCode = ctJoinLeft
        Case "right": p_JoinHowCode = ctJoinRight
        Case "outer", "full": p_JoinHowCode = ctJoinOuter
        Case "cross": p_JoinHowCode = ctJoinCross
        Case Else
            Err.Raise vbObjectError + 1362, "cTable.merge", _
                      "how must be ""inner"", ""left"", ""right"", " & _
                      """outer"", or ""cross""."
    End Select
End Function


Private Function p_JoinValidateCode(ByVal validate As String) As Long
    Select Case LCase$(Trim$(Replace(validate, ":", "_")))
        Case "":                             p_JoinValidateCode = ctValidateNone
        Case "one_to_one", "1_1":            p_JoinValidateCode = ctValidateOneToOne
        Case "one_to_many", "1_m":           p_JoinValidateCode = ctValidateOneToMany
        Case "many_to_one", "m_1":           p_JoinValidateCode = ctValidateManyToOne
        Case "many_to_many", "m_m":          p_JoinValidateCode = ctValidateManyToMany
        Case Else
            Err.Raise vbObjectError + 1363, "cTable.merge", _
                      "validate must be ""one_to_one"", ""one_to_many"", " & _
                      """many_to_one"", or ""many_to_many""."
    End Select
End Function


' Column names present in BOTH tables, in left order. pandas' default key.
Private Function p_CommonColumns(ByVal right As cTable, _
                                 ByRef lKeys() As Long, _
                                 ByRef rKeys() As Long) As Long
    Dim n As Long
    Dim c As Long
    Dim rc As Long

    ReDim lKeys(1 To IIf(m_ColCount > 0, m_ColCount, 1))
    ReDim rKeys(1 To IIf(m_ColCount > 0, m_ColCount, 1))

    For c = 1 To m_ColCount
        rc = right.columnNumber(m_Headers(c))
        If rc > 0 Then
            n = n + 1
            lKeys(n) = c
            rKeys(n) = rc
        End If
    Next c

    p_CommonColumns = n
End Function


' ---------------------------------------------------------------------------
' merge
'
' pandas DataFrame.merge. Returns a new cTable; neither side is modified.
'
' right        the other cTable
' how          "inner" (default), "left", "right", "outer", "cross"
' on_arg       key column(s) present in BOTH tables under the same name. The
'              key appears ONCE in the result. Omitted means the intersection
'              of the two tables' column names, as in pandas.
' left_on      key column(s) on the left, paired positionally with right_on.
' right_on     Both key sets are KEPT in the result, as in pandas.
' suffixes     two-element array applied to overlapping non-key names,
'              default Array("_x", "_y")
' indicator    True adds a column reporting "both", "left_only", "right_only"
' validate     "one_to_one", "one_to_many", "many_to_one", "many_to_many"
' sort         True orders the result by the join keys
'
'   Set m = t.merge(other, "left", "AccountID")
'   Set m = t.merge(other, "outer", Array("AccountID", "CategoryCode"), indicator:=True)
'   Set m = t.merge(other, "inner", left_on:="AccountID", right_on:="AccountID_Ref")
'   Set m = t.merge(other, "cross")
' ---------------------------------------------------------------------------
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

    Const CALLER As String = "cTable.merge"

    If right Is Nothing Then
        Err.Raise vbObjectError + 1364, CALLER, "right table is Nothing."
    End If

    Dim howCode As Long
    howCode = p_JoinHowCode(how)

    p_EnsureTypes

    Dim lKeys() As Long
    Dim rKeys() As Long
    Dim keyCount As Long
    Dim mergeKeys As Boolean

    Dim haveOn As Boolean
    Dim haveLeftOn As Boolean
    Dim haveRightOn As Boolean
    haveOn = Not p_ArgIsMissing(on_arg)
    haveLeftOn = Not p_ArgIsMissing(left_on)
    haveRightOn = Not p_ArgIsMissing(right_on)

    If howCode = ctJoinCross Then
        If haveOn Or haveLeftOn Or haveRightOn Then
            Err.Raise vbObjectError + 1365, CALLER, _
                      "A cross join takes no key columns."
        End If
        ReDim lKeys(1 To 1)
        ReDim rKeys(1 To 1)
        keyCount = 0
        mergeKeys = False

    ElseIf haveLeftOn Or haveRightOn Then
        If Not (haveLeftOn And haveRightOn) Then
            Err.Raise vbObjectError + 1366, CALLER, _
                      "left_on and right_on must be supplied together."
        End If
        If haveOn Then
            Err.Raise vbObjectError + 1367, CALLER, _
                      "Use on_arg, or left_on with right_on, but not both."
        End If

        keyCount = p_ResolveColumnList(lKeys, CALLER, "left_on", left_on)

        Dim rNames As Variant
        If IsArray(right_on) Then
            rNames = right_on
        Else
            rNames = Array(right_on)
        End If

        If UBound(rNames) - LBound(rNames) + 1 <> keyCount Then
            Err.Raise vbObjectError + 1368, CALLER, _
                      "left_on and right_on must name the same number of " & _
                      "columns."
        End If

        ReDim rKeys(1 To keyCount)
        Dim j As Long
        Dim nm As String

        For j = 1 To keyCount
            nm = p_LabelColumnRef(rNames(LBound(rNames) + j - 1))
            rKeys(j) = right.columnNumber(nm)
            If rKeys(j) = 0 Then
                Err.Raise vbObjectError + 1369, CALLER, _
                          "right_on column not found: """ & nm & """"
            End If
        Next j

        ' Both key sets are kept, as pandas does when the names differ.
        mergeKeys = False

    Else
        If haveOn Then
            keyCount = p_ResolveColumnList(lKeys, CALLER, "on_arg", on_arg)

            ReDim rKeys(1 To keyCount)
            For j = 1 To keyCount
                rKeys(j) = right.columnNumber(m_Headers(lKeys(j)))
                If rKeys(j) = 0 Then
                    Err.Raise vbObjectError + 1370, CALLER, _
                              "on_arg column """ & m_Headers(lKeys(j)) & _
                              """ is not present in the right table."
                End If
            Next j
        Else
            keyCount = p_CommonColumns(right, lKeys, rKeys)

            If keyCount = 0 Then
                Err.Raise vbObjectError + 1371, CALLER, _
                          "The two tables share no column names, so there " & _
                          "is no key to merge on. Supply on_arg, or left_on " & _
                          "with right_on."
            End If
        End If

        mergeKeys = True
    End If

    ' --- suffixes ---------------------------------------------------------
    Dim sufL As String
    Dim sufR As String
    sufL = "_x"
    sufR = "_y"

    If Not p_ArgIsMissing(suffixes) Then
        If Not IsArray(suffixes) Then
            Err.Raise vbObjectError + 1372, CALLER, _
                      "suffixes must be a two-element array, e.g. " & _
                      "Array(""_left"", ""_right"")."
        End If
        If UBound(suffixes) - LBound(suffixes) + 1 <> 2 Then
            Err.Raise vbObjectError + 1372, CALLER, _
                      "suffixes must be a two-element array."
        End If
        sufL = CStr(suffixes(LBound(suffixes)))
        sufR = CStr(suffixes(LBound(suffixes) + 1))
    End If

    If sufL = sufR Then
        Err.Raise vbObjectError + 1373, CALLER, _
                  "The two suffixes must differ, or overlapping columns " & _
                  "would end up with the same name."
    End If

    Dim indName As String
    indName = Trim$(indicator_name)
    If Len(indName) = 0 Then indName = "_merge"

    Set merge = p_MergeCore(right, lKeys, rKeys, keyCount, howCode, _
                            mergeKeys, ctNamePandas, sufL, sufR, _
                            indicator, indName, _
                            p_JoinValidateCode(validate), sort, _
                            Nothing, Nothing, CALLER)
End Function


'==============================================================================
'
' SORT
'
'==============================================================================

Public Sub SortBy( _
        ByVal COLS As Variant, _
        Optional ByVal NonAlphaAtEnd As Boolean = True, _
        Optional ByVal SortAscending As Boolean = True)

    If Not IsArray(COLS) Then
        COLS = Array(COLS)
    End If

    If m_RowCount <= 1 Then Exit Sub

    p_EnsureTypes

    Dim n As Long
    n = UBound(COLS) - LBound(COLS) + 1

    Dim ci() As Long
    ReDim ci(0 To n - 1)

    Dim k As Long
    Dim cName As String
    Dim idx As Long

    For k = 0 To n - 1
        cName = CStr(COLS(LBound(COLS) + k))
        idx = p_ColIdx(cName)

        If idx = 0 Then
            Err.Raise vbObjectError + 60, _
                      "cTable.SortBy", _
                      "Column not found: """ & cName & """"
        End If

        ci(k) = idx
    Next k

    p_ShellSort ci, n, SortAscending, NonAlphaAtEnd
    p_ClearIndexes
End Sub

' ---------------------------------------------------------------------------
' sort_values
'
' Sort a cTable by one or more columns, with an independent direction for ' each column. Mirrors the principal behavior of pandas DataFrame.sort_values.
'
' by:
'   One column:
'       Set result = t.sort_values("RunDate")
'
'   Multiple columns:
'       Set result = t.sort_values( _
'           Array("RunDate", "AccountID"))
'
' ascending:
'   One Boolean applies to every sort column:
'       ascending:=True
'
'   An array specifies one direction per sort column:
'       ascending:=Array(False, True)
'
'       False = descending
'       True  = ascending
'
' inplace:
'   False = return a sorted cTable; source remains unchanged
'   True  = sort this cTable and return Nothing
'
' na_position:
'   "first" = missing values appear first
'   "last"  = missing values appear last
'
' Examples:
'
'   Set result = t.sort_values( _
'       by:=Array("RunDate", "AccountID"), _
'       ascending:=Array(False, True))
'
'   Call t.sort_values( _
'       by:=Array("RunDate", "AccountID"), _
'       ascending:=Array(False, True), _
'       inplace:=True)
' ---------------------------------------------------------------------------
Public Function sort_values( _
        ByVal by As Variant, _
        Optional ByVal ascending As Variant = True, _
        Optional ByVal inplace As Boolean = False, _
        Optional ByVal na_position As String = "last") As cTable

    ' ------------------------------------------------------------
    ' Non-inplace mode: sort a clone and return it.
    ' ------------------------------------------------------------
    If Not inplace Then
        Dim result As cTable
        Dim ignoredResult As cTable

        Set result = Me.Clone

        Set ignoredResult = result.sort_values( _
                                by:=by, _
                                ascending:=ascending, _
                                inplace:=True, _
                                na_position:=na_position)

        Set sort_values = result
        Exit Function
    End If

    ' ------------------------------------------------------------
    ' Normalize the by argument into a column-name array.
    ' ------------------------------------------------------------
    Dim sortColumns As Variant

    If IsArray(by) Then
        If Not p_ArrayHasElements(by) Then
            Err.Raise vbObjectError + 1250, _
                      "cTable.sort_values", _
                      "by cannot be an empty array."
        End If

        sortColumns = by
    Else
        If IsObject(by) Or IsNull(by) Or IsEmpty(by) Then
            Err.Raise vbObjectError + 1251, _
                      "cTable.sort_values", _
                      "by must contain a column name or an array of names."
        End If

        sortColumns = Array(CStr(by))
    End If

    Dim sortColumnCount As Long
    sortColumnCount = _
        UBound(sortColumns) - LBound(sortColumns) + 1

    ' ------------------------------------------------------------
    ' Resolve column names into 1-based column numbers.
    ' ------------------------------------------------------------
    Dim colIndexes() As Long
    ReDim colIndexes(1 To sortColumnCount)

    Dim i As Long
    Dim colName As String
    Dim colNum As Long

    For i = 1 To sortColumnCount
        colName = CStr( _
                    sortColumns( _
                        LBound(sortColumns) + i - 1))

        colNum = p_ColIdx(colName)

        If colNum = 0 Then
            Err.Raise vbObjectError + 1252, _
                      "cTable.sort_values", _
                      "Column not found: """ & colName & """"
        End If

        colIndexes(i) = colNum
    Next i

    ' ------------------------------------------------------------
    ' Normalize ascending into one Boolean per sort column.
    ' ------------------------------------------------------------
    Dim ascendingFlags() As Boolean
    ReDim ascendingFlags(1 To sortColumnCount)

    If IsArray(ascending) Then
        If Not p_ArrayHasElements(ascending) Then
            Err.Raise vbObjectError + 1253, _
                      "cTable.sort_values", _
                      "ascending cannot be an empty array."
        End If

        Dim ascendingCount As Long
        ascendingCount = _
            UBound(ascending) - LBound(ascending) + 1

        If ascendingCount <> sortColumnCount Then
            Err.Raise vbObjectError + 1254, _
                      "cTable.sort_values", _
                      "The ascending array must contain one Boolean for " & _
                      "each column in by. Expected " & sortColumnCount & _
                      ", received " & ascendingCount & "."
        End If

        For i = 1 To sortColumnCount
            ascendingFlags(i) = CBool( _
                ascending(LBound(ascending) + i - 1))
        Next i
    Else
        For i = 1 To sortColumnCount
            ascendingFlags(i) = CBool(ascending)
        Next i
    End If

    ' ------------------------------------------------------------
    ' Validate missing-value placement.
    ' ------------------------------------------------------------
    Dim naFirst As Boolean

    Select Case LCase$(Trim$(na_position))
        Case "first"
            naFirst = True

        Case "last"
            naFirst = False

        Case Else
            Err.Raise vbObjectError + 1255, _
                      "cTable.sort_values", _
                      "na_position must be ""first"" or ""last""."
    End Select

    p_EnsureTypes

    p_ShellSortValues _
        colIndexes:=colIndexes, _
        ascendingFlags:=ascendingFlags, _
        sortColumnCount:=sortColumnCount, _
        naFirst:=naFirst

    ' pandas returns None when inplace=True.
    Set sort_values = Nothing
End Function

'==============================================================================
'
' COLUMN MANIPULATION
'
'==============================================================================

' Drop a column by name.
' If IgnoreMissing=True and the column is not found, do nothing.
Public Sub DropColNameInPlace(ByVal colName As Variant, _
                              Optional ByVal IgnoreMissing As Boolean = False)
    colName = p_LabelColumnRef(colName)

    Dim ci As Long
    ci = p_ColIdx(colName)

    If ci = 0 Then
        If IgnoreMissing Then Exit Sub

        Err.Raise vbObjectError + 70, "cTable.DropColNameInPlace", _
                  "Column not found: """ & colName & """"
    End If

    DropColNoInPlace ci
End Sub

' Drop a column by 1-based column number.
Public Sub DropColNoInPlace(ByVal colNumber As Long)
    If colNumber < 1 Or colNumber > m_ColCount Then
        Err.Raise vbObjectError + 71, "cTable.DropColNoInPlace", _
                  "Column number out of bounds: " & colNumber
    End If

    p_EnsureTypes

    Dim newCC As Long: newCC = m_ColCount - 1
    If newCC = 0 Then
        m_ColCount = 0
        m_RowCount = 0
        m_TypesAllocated = False
        If Not m_Lin Is Nothing Then m_Lin.InitUniform 0, 0, ctSrcUnknown
        p_ClearIndexes
        Exit Sub
    End If

    ' Rebuild headers and dtypes in the same pass.
    Dim newH() As String: ReDim newH(1 To newCC)
    Dim newT() As ctDataType: ReDim newT(1 To newCC)
    Dim dst As Long: dst = 0
    Dim src As Long
    For src = 1 To m_ColCount
        If src <> colNumber Then
            dst = dst + 1
            newH(dst) = m_Headers(src)
            newT(dst) = m_ColTypes(src)
        End If
    Next src

    ' Rebuild data.
    If m_RowCount > 0 Then
        Dim newD() As Variant
        ReDim newD(1 To m_RowCount, 1 To newCC)
        Dim r As Long
        For r = 1 To m_RowCount
            dst = 0
            For src = 1 To m_ColCount
                If src <> colNumber Then
                    dst = dst + 1
                    newD(r, dst) = m_Data(r, src)
                End If
            Next src
        Next r
        m_Data = newD
    End If

    m_Headers = newH
    m_ColTypes = newT
    m_ColCount = newCC
    If Not m_Lin Is Nothing Then m_Lin.x_RemoveColumnAt colNumber
    p_ClearIndexes
End Sub
Public Function DropColumnIfExists(ByVal colName As Variant) As Boolean
    colName = p_LabelColumnRef(colName)
    If Me.columnNumber(colName) > 0 Then
        Me.DropColNameInPlace colName
        DropColumnIfExists = True
    Else
        DropColumnIfExists = False
    End If
End Function

' Insert a blank column (filled with Null) at the given 1-based position.
' colNumber=0 (default) appends to the right.
'
' dataType defaults to ctVariant. Null is valid for every cTable dtype.
Public Sub InsertColumn(ByVal colName As String, _
                        Optional ByVal colNumber As Long = 0, _
                        Optional ByVal dataType As ctDataType = ctVariant)

    If Len(Trim$(colName)) = 0 Then
        Err.Raise vbObjectError + 72, _
                  "cTable.InsertColumn", _
                  "Column name cannot be blank."
    End If

    If p_ColIdx(colName) > 0 Then
        Err.Raise vbObjectError + 73, _
                  "cTable.InsertColumn", _
                  "Column already exists: """ & colName & """"
    End If

    p_EnsureTypes

    Dim insertAt As Long

    If colNumber <= 0 Or colNumber > m_ColCount Then
        insertAt = m_ColCount + 1
    Else
        insertAt = colNumber
    End If

    Dim newCC As Long
    newCC = m_ColCount + 1

    ' Rebuild headers and dtypes.
    Dim newH() As String
    Dim newT() As ctDataType
    Dim pos As Long

    ReDim newH(1 To newCC)
    ReDim newT(1 To newCC)

    For pos = 1 To newCC
        If pos < insertAt Then
            newH(pos) = m_Headers(pos)
            newT(pos) = m_ColTypes(pos)

        ElseIf pos = insertAt Then
            newH(pos) = Trim$(colName)
            newT(pos) = dataType

        Else
            newH(pos) = m_Headers(pos - 1)
            newT(pos) = m_ColTypes(pos - 1)
        End If
    Next pos

    ' Rebuild data. New cells are Null.
    If m_RowCount > 0 Then
        Dim newD() As Variant
        Dim r As Long

        ReDim newD(1 To m_RowCount, 1 To newCC)

        For r = 1 To m_RowCount
            For pos = 1 To newCC
                If pos < insertAt Then
                    newD(r, pos) = m_Data(r, pos)

                ElseIf pos = insertAt Then
                    newD(r, pos) = Null

                Else
                    newD(r, pos) = m_Data(r, pos - 1)
                End If
            Next pos
        Next r

        m_Data = newD
    End If

    m_Headers = newH
    m_ColTypes = newT
    m_ColCount = newCC
    m_TypesAllocated = True

    If Not m_Lin Is Nothing Then
        m_Lin.x_InsertColumnAt insertAt, ctSrcUnknown
    End If

    p_ClearIndexes
End Sub


' Rename an existing column.
Public Sub RenameColumn(ByVal oldName As String, ByVal newName As String)
    Dim ci As Long: ci = p_ColIdx(oldName)
    If ci = 0 Then
        Err.Raise vbObjectError + 80, "cTable.RenameColumn", _
                  "Column not found: """ & oldName & """"
    End If
    m_Headers(ci) = newName
    p_ClearIndexes
End Sub
' -----------------------------------------------------------------------------
' rename_columns (multiple)
'
' Renames columns by physical position, starting with column 1.
'
' Rules:
'   - newNames may contain 1 through ColumnCount entries.
'   - More entries than ColumnCount raises an error.
'   - Null, Empty, or blank text leaves that column's current name unchanged.
'   - Duplicate resulting names raise an error.
'   - A supplied name is trimmed before assignment.
'
' Examples:
'
'   ' Rename columns 1 and 3 only.
'   t.rename_columns Array("AccountID", Null, "Organization")
'
'   ' Rename the first four columns of a six-column table.
'   t.rename_columns Array("AccountID", "Fund", "Organization", "Amount")
' -----------------------------------------------------------------------------
Public Sub rename_columns(ByVal newNames As Variant)

    Const CALLER As String = "cTable.rename_columns"

    If Not IsArray(newNames) Then
        Err.Raise vbObjectError + 1730, CALLER, _
                  "newNames must be an array of column names."
    End If

    If Not p_ArrayHasElements(newNames) Then Exit Sub

    Dim suppliedCount As Long
    suppliedCount = UBound(newNames) - LBound(newNames) + 1

    If suppliedCount > m_ColCount Then
        Err.Raise vbObjectError + 1731, CALLER, _
                  "newNames contains " & suppliedCount & _
                  " entries, but the table has only " & m_ColCount & _
                  " columns."
    End If

    ' Build and validate the entire resulting header list before assigning it.
    ' This ensures the table is unchanged if validation fails.
    Dim proposedHeaders() As String
    ReDim proposedHeaders(1 To m_ColCount)

    Dim c As Long
    For c = 1 To m_ColCount
        proposedHeaders(c) = m_Headers(c)
    Next c

    Dim sourceIndex As Long
    Dim value As Variant
    Dim proposedName As String

    For c = 1 To suppliedCount
        sourceIndex = LBound(newNames) + c - 1
        value = newNames(sourceIndex)

        ' Null, Empty, or blank text means "do not rename this column."
        If IsNull(value) Or IsEmpty(value) Then
            ' Leave proposedHeaders(c) unchanged.

        ElseIf IsObject(value) Or IsError(value) Or IsArray(value) Then
            Err.Raise vbObjectError + 1732, CALLER, _
                      "Name at position " & c & _
                      " must be text, Null, Empty, or blank."

        Else
            proposedName = Trim$(CStr(value))

            If Len(proposedName) > 0 Then
                proposedHeaders(c) = proposedName
            End If
        End If
    Next c

    ' Validate all final names, including names that were retained.
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare

    For c = 1 To m_ColCount
        If Len(Trim$(proposedHeaders(c))) = 0 Then
            Err.Raise vbObjectError + 1733, CALLER, _
                      "Column " & c & " would have a blank name."
        End If

        If seen.Exists(proposedHeaders(c)) Then
            Err.Raise vbObjectError + 1734, CALLER, _
                      "Duplicate resulting column name: """ & _
                      proposedHeaders(c) & """"
        End If

        seen.Add proposedHeaders(c), True
    Next c

    m_Headers = proposedHeaders
    p_ClearIndexes
End Sub


'==============================================================================
'
' COMPUTED COLUMNS
'
'==============================================================================

' ---------------------------------------------------------------------------
' AddComputedColumn
'
' Adds a new column and populates it by calling a public VBA function once ' for each row. The calculation function receives a Dictionary containing ' the row's values, keyed by column name.
'
' resultType:
'   ctVariant  = leave the calculated column untyped (existing behavior)
'   Other type = convert and validate every calculated result
'
' Example:
'   t.AddComputedColumn _
'       "AvailableSpending", _
'       "CalcAvailableSpending", _
'       ctCurrency
' ---------------------------------------------------------------------------
Public Sub AddComputedColumn( _
        ByVal colName As String, _
        ByVal computeFuncName As String, _
        Optional ByVal resultType As ctDataType = ctVariant)

    If p_HeaderExists(colName) Then
        Err.Raise vbObjectError + 1100, _
                  "cTable.AddComputedColumn", _
                  "Column already exists: """ & colName & """"
    End If

    ' InsertColumn appends the new column and initializes it with Null.
    Me.InsertColumn colName

    Dim targetCol As Long
    targetCol = m_ColCount

    Dim rowDict As Object
    Set rowDict = CreateObject("Scripting.Dictionary")
    rowDict.CompareMode = vbTextCompare

    Dim r As Long
    Dim c As Long
    Dim result As Variant

    For r = 1 To m_RowCount
        rowDict.RemoveAll

        ' Build a dictionary containing the source columns for this row.
        ' Exclude the newly added destination column.
        For c = 1 To m_ColCount
            If c <> targetCol Then
                rowDict(m_Headers(c)) = m_Data(r, c)
            End If
        Next c

        result = Application.Run(computeFuncName, rowDict)

        ' Assign directly so indexes are not rebuilt once per row.
        m_Data(r, targetCol) = result
    Next r

    p_EnsureTypes

    If resultType <> ctVariant Then
        ' Convert and validate all calculated values.
        ' Null values remain valid for every dtype.
        p_CoerceColumn _
            targetCol, _
            resultType, _
            ctErrorsRaise

        m_ColTypes(targetCol) = resultType
     Else
        m_ColTypes(targetCol) = p_InferColumnType(targetCol)

        If m_ColTypes(targetCol) <> ctVariant Then
            p_CoerceColumn _
                targetCol, _
                m_ColTypes(targetCol), _
                ctErrorsPreserve
        End If
     End If

    ' The whole column now originates from the calculation, by name.
    If Not m_Lin Is Nothing Then
        p_LinSync
        m_Lin.StampColumn targetCol, _
                          cLineageRegistry.CalculatedSourceId(computeFuncName)
    End If

    ' Calculated values may invalidate existing indexes.
    p_ClearIndexes
End Sub

' ---------------------------------------------------------------------------
' SetComputedColumn
'
' Creates or replaces a column using a public function in a standard module.
' The calculation function receives a dictionary containing the current row.
'
' Example:
'   t.SetComputedColumn _
'       "AvailableSpending", _
'       "CalcAvailableSpending", _
'       ctCurrency
' ---------------------------------------------------------------------------
Public Sub SetComputedColumn( _
        ByVal colName As String, _
        ByVal computeFuncName As String, _
        Optional ByVal resultType As ctDataType = ctVariant)

    Dim targetCol As Long
    targetCol = p_ColIdx(colName)

    ' Add the column if it does not already exist.
    If targetCol = 0 Then
        Me.InsertColumn colName
        targetCol = m_ColCount
    End If

    Dim rowValues As Object
    Set rowValues = CreateObject("Scripting.Dictionary")
    rowValues.CompareMode = vbTextCompare

    Dim r As Long
    Dim c As Long
    Dim calculatedValue As Variant

    For r = 1 To m_RowCount
        rowValues.RemoveAll

        ' Build the row dictionary, excluding the destination column.
        For c = 1 To m_ColCount
            If c <> targetCol Then
                rowValues(m_Headers(c)) = m_Data(r, c)
            End If
        Next c

        calculatedValue = Application.Run( _
                                computeFuncName, _
                                rowValues)

        ' Direct assignment avoids clearing indexes on every row.
        m_Data(r, targetCol) = calculatedValue
    Next r

    p_EnsureTypes

    If resultType = ctVariant Then
        ' Infer the resulting dtype from the calculated values.
        m_ColTypes(targetCol) = p_InferColumnType(targetCol)

        If m_ColTypes(targetCol) <> ctVariant Then
            p_CoerceColumn _
                targetCol, _
                m_ColTypes(targetCol), _
                ctErrorsPreserve
        End If
    Else
        ' Assign and validate the requested result dtype.
        p_CoerceColumn _
            targetCol, _
            resultType, _
            ctErrorsRaise

        m_ColTypes(targetCol) = resultType
    End If

    If Not m_Lin Is Nothing Then
        p_LinSync
        m_Lin.StampColumn targetCol, _
                          cLineageRegistry.CalculatedSourceId(computeFuncName)
    End If

    p_ClearIndexes
End Sub

' ---------------------------------------------------------------------------
' WithComputedColumn
'
' Returns a new cTable containing an additional calculated column.
' The source cTable is not modified.
'
' resultType:
'   ctVariant  = leave the calculated column untyped
'   Other type = convert and validate calculated values
'
' Example:
'   Set result = t.WithComputedColumn( _
'       "AvailableSpending", _
'       "CalcAvailableSpending", _
'       ctCurrency)
' ---------------------------------------------------------------------------
Public Function WithComputedColumn( _
        ByVal colName As String, _
        ByVal computeFuncName As String, _
        Optional ByVal resultType As ctDataType = ctVariant) As cTable

    Dim newTable As cTable
    Set newTable = Me.Clone

    newTable.AddComputedColumn _
        colName:=colName, _
        computeFuncName:=computeFuncName, _
        resultType:=resultType

    Set WithComputedColumn = newTable
End Function
'==============================================================================
'
' GROUPED DIFFERENCING AND LAGGING
'
' pandas equivalence:
'     df = df.sort_values(["AccountID", "RunDate"])
'     df["Expenditure_Diff"] = df.groupby("AccountID")["Expenditure"].diff()
'
'     t.sort_values Array("AccountID", "RunDate"), inplace:=True
'     t.AddDiffColumn "Expenditure", "AccountID"
'
' ORDER IS THE CALLER'S RESPONSIBILITY. Neither diff nor shift sorts. Each
' walks rows in their CURRENT order and pairs every row with the nearest
' earlier row that shares its group key, exactly as pandas does. Sorting
' RunDate DESCENDING therefore yields older-minus-newer, the negative of the
' snippet above. This is a sign error that produces plausible numbers, so
' verify the sort direction before trusting any output.
'
' GROUP KEYS are compared with the class-wide vbTextCompare convention, so
' "abc" and "ABC" fall in the same group. Keys are built by the existing
' p_NuniqueValueKey / p_DuplicateRowKey helpers, so they are dtype-aware and
' collision-resistant across mixed types.
'
' NA KEYS are excluded by default (dropna_keys:=True), matching pandas: a row
' whose key is Null gets a Null result and never links to another row. Pass
' dropna_keys:=False to treat missing as its own group.
'
'==============================================================================

' ---------------------------------------------------------------------------
' Normalize a by argument into 1-based column numbers.
'
' Returns the key count. A missing, blank, or empty-array by returns 0,
' meaning "one implicit group containing every row" - i.e. ungrouped diff.
' keyCols is left unallocated in that case and must not be indexed.
' ---------------------------------------------------------------------------
Private Function p_ResolveGroupKeys(ByRef keyCols() As Long, _
                                    ByVal callerName As String, _
                                    Optional ByVal by As Variant) As Long

    If IsMissing(by) Then Exit Function

    ' A missing optional forwarded from another procedure arrives as an
    ' Error variant rather than as Missing. Treat both as "no grouping".
    Dim isErrorMarker As Boolean
    On Error Resume Next
    isErrorMarker = IsError(by)
    On Error GoTo 0
    If isErrorMarker Then Exit Function

    Dim names As Variant

    If IsArray(by) Then
        If Not p_ArrayHasElements(by) Then Exit Function
        names = by
    Else
        If IsObject(by) Then
            Err.Raise vbObjectError + 1300, callerName, _
                      "by must be a column name or an array of column names."
        End If
        If p_IsBlankColumnArgument(by) Then Exit Function
        names = Array(by)
    End If

    Dim n As Long
    n = UBound(names) - LBound(names) + 1
    ReDim keyCols(1 To n)

    Dim i As Long
    Dim nm As String
    Dim ci As Long

    For i = 1 To n
        nm = p_LabelColumnRef(names(LBound(names) + i - 1))
        ci = p_ColIdx(nm)

        If ci = 0 Then
            Err.Raise vbObjectError + 1301, callerName, _
                      "Group key column not found: """ & nm & """"
        End If

        keyCols(i) = ci
    Next i

    p_ResolveGroupKeys = n
End Function


' ---------------------------------------------------------------------------
' One row's group key. Single-key groupings skip the length-prefixing that
' only matters for disambiguating composites, which is the common case and
' roughly halves the string work on large tables.
'
' Requires p_EnsureTypes to have been called.
' ---------------------------------------------------------------------------
Private Function p_GroupRowKey(ByVal rowNum As Long, _
                               ByRef keyCols() As Long, _
                               ByVal keyCount As Long) As String

    If keyCount = 1 Then
        p_GroupRowKey = p_NuniqueValueKey( _
                            m_Data(rowNum, keyCols(1)), _
                            m_ColTypes(keyCols(1)))
    Else
        p_GroupRowKey = p_DuplicateRowKey(rowNum, keyCols, keyCount)
    End If
End Function


' ---------------------------------------------------------------------------
' Build a linked list of same-group neighbours, one pass, in current row order.
'
' lookBack = True   link(r) = nearest EARLIER row sharing r's key   (lag)
' lookBack = False  link(r) = nearest LATER row sharing r's key     (lead)
'
' link(r) = 0 means r has no such neighbour: the first row of its group, or a
' row excluded by an NA key.
'
' Chaining rather than materializing member lists keeps this O(rows) in time
' and one Long array plus one key dictionary in space, and gives periods > 1
' for free by walking the chain repeatedly.
' ---------------------------------------------------------------------------
Private Function p_NeighborInGroup(ByRef keyCols() As Long, _
                                   ByVal keyCount As Long, _
                                   ByVal dropNAKeys As Boolean, _
                                   ByVal lookBack As Boolean) As Long()

    Dim link() As Long
    ReDim link(1 To IIf(m_RowCount > 0, m_RowCount, 1))

    If m_RowCount = 0 Then
        p_NeighborInGroup = link
        Exit Function
    End If

    Dim firstRow As Long
    Dim lastRow As Long
    Dim stepBy As Long

    If lookBack Then
        firstRow = 1
        lastRow = m_RowCount
        stepBy = 1
    Else
        firstRow = m_RowCount
        lastRow = 1
        stepBy = -1
    End If

    ' Ungrouped: the neighbour is simply the adjacent row.
    If keyCount = 0 Then
        Dim r0 As Long
        Dim adj As Long

        For r0 = firstRow To lastRow Step stepBy
            adj = r0 - stepBy
            If adj >= 1 And adj <= m_RowCount Then link(r0) = adj
        Next r0

        p_NeighborInGroup = link
        Exit Function
    End If

    Dim lastSeen As Object
    Set lastSeen = CreateObject("Scripting.Dictionary")
    lastSeen.CompareMode = vbTextCompare

    Dim r As Long
    Dim i As Long
    Dim key As String
    Dim skipRow As Boolean

    For r = firstRow To lastRow Step stepBy

        skipRow = False

        If dropNAKeys Then
            For i = 1 To keyCount
                If p_IsNA(m_Data(r, keyCols(i))) Then
                    skipRow = True
                    Exit For
                End If
            Next i
        End If

        If Not skipRow Then
            key = p_GroupRowKey(r, keyCols, keyCount)
            If lastSeen.Exists(key) Then link(r) = lastSeen(key)
            lastSeen(key) = r
        End If
    Next r

    p_NeighborInGroup = link
End Function


' True as 1 rather than VBA's -1, matching p_CollectNumericValues.
Private Function p_BoolAsLong(ByVal v As Variant) As Long
    If CBool(v) Then p_BoolAsLong = 1 Else p_BoolAsLong = 0
End Function


' ---------------------------------------------------------------------------
' a - b, dtype-aware. Null whenever either side is missing.
'
' ctCurrency stays in Currency rather than routing through Double, so money
' keeps exact 4-decimal arithmetic. A ctDate difference is a span in DAYS
' returned as a Double, not a Date; cTable has no timedelta dtype.
' ---------------------------------------------------------------------------
Private Function p_DiffPair(ByVal a As Variant, ByVal b As Variant, _
                            ByVal colType As ctDataType) As Variant

    If p_IsNA(a) Or p_IsNA(b) Then
        p_DiffPair = Null
        Exit Function
    End If

    On Error GoTo NotSubtractable

    Select Case colType
        Case ctCurrency
            p_DiffPair = CCur(a) - CCur(b)

        Case ctLong
            p_DiffPair = CLng(a) - CLng(b)

        Case ctDouble
            p_DiffPair = CDbl(a) - CDbl(b)

        Case ctBoolean
            p_DiffPair = p_BoolAsLong(a) - p_BoolAsLong(b)

        Case ctDate
            p_DiffPair = CDbl(CDate(a)) - CDbl(CDate(b))

        Case ctString
            Err.Raise vbObjectError + 1302, "cTable.diff", _
                      "diff is not defined for a ctString column."

        Case Else
            ' ctVariant: decide from the actual Variant subtypes.
            If VarType(a) = vbDate And VarType(b) = vbDate Then
                p_DiffPair = CDbl(CDate(a)) - CDbl(CDate(b))
            ElseIf VarType(a) = vbCurrency Or VarType(b) = vbCurrency Then
                p_DiffPair = CCur(a) - CCur(b)
            ElseIf IsNumeric(a) And IsNumeric(b) Then
                p_DiffPair = CDbl(a) - CDbl(b)
            Else
                p_DiffPair = Null
            End If
    End Select

    Exit Function

NotSubtractable:
    Err.Clear
    p_DiffPair = Null
End Function


' ---------------------------------------------------------------------------
' The dtype a diff column should carry, given its source column's dtype.
'
' This is deliberately derived rather than inferred from the resulting values.
' p_InferColumnType would classify a ctDouble column whose differences all
' happen to land on whole numbers as ctLong, silently narrowing the column
' and truncating any fractional value written to it later.
' ---------------------------------------------------------------------------
Private Function p_DiffResultType(ByVal srcType As ctDataType) As ctDataType
    Select Case srcType
        Case ctCurrency
            p_DiffResultType = ctCurrency

        Case ctLong
            p_DiffResultType = ctLong

        Case ctDouble
            p_DiffResultType = ctDouble

        Case ctBoolean
            p_DiffResultType = ctLong

        Case ctDate
            p_DiffResultType = ctDouble      ' a span in days

        Case Else
            p_DiffResultType = ctVariant
    End Select
End Function


' A stable, human-readable lineage label, e.g. "diff(Expenditure by AccountID)".
Private Function p_DiffLabel(ByVal opName As String, _
                             ByVal colName As String, _
                             ByRef keyCols() As Long, _
                             ByVal keyCount As Long, _
                             ByVal periods As Long, _
                             ByVal baseline As Variant) As String

    Dim extras As String
    If periods <> 1 Then p_AddExtra extras, "periods=" & CStr(periods)

    ' Only a supplied baseline appears, so labels for existing calls do not
    ' change and their interned sources are not disturbed.
    If Not p_ArgIsMissing(baseline) Then
        p_AddExtra extras, "baseline=" & p_LabelScalar(baseline)
    End If

    p_DiffLabel = p_OpLabel(opName, colName, keyCols, keyCount, extras)
End Function


' ---------------------------------------------------------------------------
' diff
'
' Returns a 1-based Variant array of RowCount elements: each row's value minus
' the value `periods` positions earlier within its group. Nothing is modified.
'
' periods > 0   difference against an EARLIER row  (pandas default)
' periods < 0   difference against a LATER row
' periods = 0   zero for every non-NA row, as in pandas
'
' Rows with no such neighbour receive Null.
'
'   v = t.diff("Expenditure", "AccountID")
'   v = t.diff("Expenditure", Array("AccountID", "CategoryCode"), 2)
'   v = t.diff("Balance")                       ' ungrouped, row over row
' ---------------------------------------------------------------------------
Public Function diff(ByVal colName As Variant, _
                     Optional ByVal by As Variant, _
                     Optional ByVal periods As Long = 1, _
                     Optional ByVal dropna_keys As Boolean = True, _
                     Optional ByVal baseline As Variant) As Variant
    colName = p_LabelColumnRef(colName)

    Dim ci As Long
    ci = p_ColIdx(colName)

    If ci = 0 Then
        Err.Raise vbObjectError + 1303, "cTable.diff", _
                  "Column not found: """ & colName & """"
    End If

    If m_RowCount = 0 Then
        diff = Array()
        Exit Function
    End If

    p_EnsureTypes

    Dim colType As ctDataType
    colType = m_ColTypes(ci)

    If colType = ctString Then
        Err.Raise vbObjectError + 1302, "cTable.diff", _
                  "diff is not defined for ctString column """ & colName & """."
    End If

    Dim keyCols() As Long
    Dim keyCount As Long
    keyCount = p_ResolveGroupKeys(keyCols, "cTable.diff", by)

    Dim link() As Long
    link = p_NeighborInGroup(keyCols, keyCount, dropna_keys, (periods >= 0))

    Dim steps As Long
    steps = Abs(periods)

    Dim Out() As Variant
    ReDim Out(1 To m_RowCount)

    Dim r As Long
    Dim p As Long
    Dim k As Long

    For r = 1 To m_RowCount

        If steps = 0 Then
            Out(r) = p_DiffPair(m_Data(r, ci), m_Data(r, ci), colType)
        Else
            p = r

            For k = 1 To steps
                p = link(p)
                If p = 0 Then Exit For
            Next k

            If p = 0 Then
                ' No partner. Measure against the baseline if one was given,
                ' otherwise report the absence honestly as Null.
                If p_ArgIsMissing(baseline) Then
                    Out(r) = Null
                Else
                    Out(r) = p_DiffPair(m_Data(r, ci), baseline, colType)
                End If
            Else
                Out(r) = p_DiffPair(m_Data(r, ci), m_Data(p, ci), colType)
            End If

        End If
    Next r

    diff = Out
End Function


' ---------------------------------------------------------------------------
' shift
'
' The lag itself rather than the difference: each row receives the VALUE from
' `periods` positions earlier within its group, Null where there is none.
' Same ordering and grouping rules as diff. pandas groupby(K)[V].shift(n).
'
'   v = t.shift("Expenditure", "AccountID")           ' prior period's amount
' ---------------------------------------------------------------------------
Public Function shift(ByVal colName As Variant, _
                      Optional ByVal by As Variant, _
                      Optional ByVal periods As Long = 1, _
                      Optional ByVal dropna_keys As Boolean = True, _
                      Optional ByVal fill_value As Variant) As Variant
    colName = p_LabelColumnRef(colName)

    Dim ci As Long
    ci = p_ColIdx(colName)

    If ci = 0 Then
        Err.Raise vbObjectError + 1303, "cTable.shift", _
                  "Column not found: """ & colName & """"
    End If

    If m_RowCount = 0 Then
        shift = Array()
        Exit Function
    End If

    p_EnsureTypes

    Dim keyCols() As Long
    Dim keyCount As Long
    keyCount = p_ResolveGroupKeys(keyCols, "cTable.shift", by)

    Dim link() As Long
    link = p_NeighborInGroup(keyCols, keyCount, dropna_keys, (periods >= 0))

    Dim useFill As Boolean
    useFill = Not p_ArgIsMissing(fill_value)

    Dim steps As Long
    steps = Abs(periods)

    Dim Out() As Variant
    ReDim Out(1 To m_RowCount)

    Dim r As Long
    Dim p As Long
    Dim k As Long

    For r = 1 To m_RowCount

        If p_RowKeyIsNA(r, keyCols, keyCount, dropna_keys) Then
            ' Not part of any group, so there is nothing to fill.
            Out(r) = Null

        ElseIf steps = 0 Then
            Out(r) = m_Data(r, ci)

        Else
            p = r

            For k = 1 To steps
                p = link(p)
                If p = 0 Then Exit For
            Next k

            If p = 0 Then
                If useFill Then Out(r) = fill_value Else Out(r) = Null
            Else
                Out(r) = m_Data(p, ci)
            End If
        End If
    Next r

    shift = Out
End Function



' ---------------------------------------------------------------------------
' Shared writer for AddDiffColumn / SetDiffColumn.
'
' Values are computed against the CURRENT column layout before any column is
' added, so a target name that overwrites the source column or a group key
' still produces correct results.
' ---------------------------------------------------------------------------
Private Sub p_ApplyDiff(ByVal colName As String, _
                        ByVal periods As Long, _
                        ByVal newColName As String, _
                        ByVal resultType As ctDataType, _
                        ByVal dropNAKeys As Boolean, _
                        ByVal allowReplace As Boolean, _
                        ByVal callerName As String, _
                        Optional ByVal by As Variant, _
                        Optional ByVal baseline As Variant)

    Dim srcCol As Long
    srcCol = p_ColIdx(colName)

    If srcCol = 0 Then
        Err.Raise vbObjectError + 1303, callerName, _
                  "Column not found: """ & colName & """"
    End If

    Dim targetName As String
    targetName = Trim$(newColName)
    If Len(targetName) = 0 Then targetName = colName & "_Diff"

    Dim targetCol As Long
    targetCol = p_ColIdx(targetName)

    If targetCol > 0 And Not allowReplace Then
        Err.Raise vbObjectError + 1304, callerName, _
                  "Column already exists: """ & targetName & """"
    End If

    p_EnsureTypes

    Dim srcType As ctDataType
    srcType = m_ColTypes(srcCol)

    Dim values As Variant
    values = Me.diff(colName, by, periods, dropNAKeys, baseline)

    Dim keyCols() As Long
    Dim keyCount As Long
    keyCount = p_ResolveGroupKeys(keyCols, callerName, by)

    If targetCol = 0 Then
        Me.InsertColumn targetName
        targetCol = m_ColCount
        p_EnsureTypes
    End If

    Dim r As Long
    For r = 1 To m_RowCount
        m_Data(r, targetCol) = values(r)
    Next r

    Dim finalType As ctDataType
    If resultType = ctVariant Then
        finalType = p_DiffResultType(srcType)
    Else
        finalType = resultType
    End If

    If finalType <> ctVariant Then
        ' Group-boundary Nulls convert cleanly: p_TryConvert passes NA
        ' through as Null for every dtype.
        p_CoerceColumn targetCol, finalType, ctErrorsRaise
    End If

    m_ColTypes(targetCol) = finalType

    ' The whole column originates from the calculation, by name.
    If Not m_Lin Is Nothing Then
        p_LinSync
        m_Lin.StampColumn targetCol, _
            cLineageRegistry.CalculatedSourceId( _
                p_DiffLabel("diff", colName, keyCols, keyCount, _
                periods, baseline))
    End If

    p_ClearIndexes
End Sub


' ---------------------------------------------------------------------------
' AddDiffColumn
'
' Adds a grouped difference column in place. Raises if the target name is
' already taken, matching AddComputedColumn.
'
' newColName  blank (default) -> colName & "_Diff"
' resultType  ctVariant (default) -> derived from the source column's dtype:
'             ctCurrency stays ctCurrency, ctDate becomes ctDouble days,
'             an untyped source produces an untyped result. Any other value
'             is applied and validated with ctErrorsRaise.
'
'   t.sort_values Array("AccountID", "RunDate"), inplace:=True
'   t.AddDiffColumn "Expenditure", "AccountID", newColName:="Expenditure_Diff"
' ---------------------------------------------------------------------------
Public Sub AddDiffColumn(ByVal colName As Variant, _
                         Optional ByVal by As Variant, _
                         Optional ByVal periods As Long = 1, _
                         Optional ByVal newColName As String = "", _
                         Optional ByVal resultType As ctDataType = ctVariant, _
                         Optional ByVal dropna_keys As Boolean = True, _
                         Optional ByVal baseline As Variant)
    colName = p_LabelColumnRef(colName)

    p_ApplyDiff colName, periods, newColName, resultType, dropna_keys, _
                False, "cTable.AddDiffColumn", by, baseline
End Sub


' ---------------------------------------------------------------------------
' SetDiffColumn
'
' Creates or replaces the target column, matching SetComputedColumn and the
' pandas assignment idiom df["X"] = ... which overwrites without complaint.
' ---------------------------------------------------------------------------
Public Sub SetDiffColumn(ByVal colName As Variant, _
                         Optional ByVal by As Variant, _
                         Optional ByVal periods As Long = 1, _
                         Optional ByVal newColName As String = "", _
                         Optional ByVal resultType As ctDataType = ctVariant, _
                         Optional ByVal dropna_keys As Boolean = True, _
                         Optional ByVal baseline As Variant)
    colName = p_LabelColumnRef(colName)

    p_ApplyDiff colName, periods, newColName, resultType, dropna_keys, _
                True, "cTable.SetDiffColumn", by, baseline
End Sub


' ---------------------------------------------------------------------------
' WithDiffColumn
'
' Returns a new cTable carrying the diff column. The source is unchanged.
' ---------------------------------------------------------------------------
Public Function WithDiffColumn(ByVal colName As Variant, _
                               Optional ByVal by As Variant, _
                               Optional ByVal periods As Long = 1, _
                               Optional ByVal newColName As String = "", _
                               Optional ByVal resultType As ctDataType = ctVariant, _
                               Optional ByVal dropna_keys As Boolean = True, _
                               Optional ByVal baseline As Variant) As cTable
    colName = p_LabelColumnRef(colName)

    Dim t As cTable
    Set t = Me.Clone

    t.AddDiffColumn colName, by, periods, newColName, resultType, _
                    dropna_keys, baseline

    Set WithDiffColumn = t
End Function


' ---------------------------------------------------------------------------
' AddShiftColumn / SetShiftColumn
'
' The lag written to a column. Add raises if the target exists; Set replaces.
' The result carries the SOURCE column's dtype, because a shifted value is the
' original value moved rather than a derived quantity.
' ---------------------------------------------------------------------------
Public Sub AddShiftColumn(ByVal colName As Variant, _
                          Optional ByVal by As Variant, _
                          Optional ByVal periods As Long = 1, _
                          Optional ByVal newColName As String = "", _
                          Optional ByVal dropna_keys As Boolean = True, _
                          Optional ByVal fill_value As Variant)
    colName = p_LabelColumnRef(colName)

    p_ApplyShift colName, periods, newColName, dropna_keys, fill_value, _
                 False, "cTable.AddShiftColumn", by
End Sub
Public Sub SetShiftColumn(ByVal colName As Variant, _
                          Optional ByVal by As Variant, _
                          Optional ByVal periods As Long = 1, _
                          Optional ByVal newColName As String = "", _
                          Optional ByVal dropna_keys As Boolean = True, _
                          Optional ByVal fill_value As Variant)
    colName = p_LabelColumnRef(colName)

    p_ApplyShift colName, periods, newColName, dropna_keys, fill_value, _
                 True, "cTable.SetShiftColumn", by
End Sub

' ---------------------------------------------------------------------------
' GroupBy
'
' Returns a cGroupBy handle for chained, pandas-shaped calls. Requires the
' cGroupBy class plus EDIT 1 and EDIT 2 above.
'
'   t.GroupBy("AccountID").Col("Expenditure").diff()
' ---------------------------------------------------------------------------
Public Function GroupBy(ByVal by As Variant) As cGroupBy

    Dim keyCols() As Long
    Dim keyCount As Long
    keyCount = p_ResolveGroupKeys(keyCols, "cTable.GroupBy", by)

    If keyCount = 0 Then
        Err.Raise vbObjectError + 1310, "cTable.GroupBy", _
                  "GroupBy requires at least one key column."
    End If

    Dim gb As cGroupBy
    Set gb = New cGroupBy
    gb.x_Init Me, by

    Set GroupBy = gb
End Function

' ---------------------------------------------------------------------------
' History
'
' Lazily creates the table's cTableHistory helper.
'
'   t.History.Init "RunDate"
' ---------------------------------------------------------------------------
Public Property Get history() As cTableHistory

    If m_History Is Nothing Then
        Set m_History = New cTableHistory
    End If

    ' Arm Init with the current table version without creating a permanent
    ' cTable <-> cTableHistory back-reference.
    m_History.x_Arm m_Version

    Set history = m_History
End Property

' ---------------------------------------------------------------------------
' HistoryFor
'
' Query the initialized history using one entity-identity column.
'
' entityValues omitted -> all entities
' columns omitted      -> all changed attributes
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
' ---------------------------------------------------------------------------
Public Function HistoryFor( _
        ByVal entityColumn As String, _
        Optional ByVal entityValues As Variant, _
        Optional ByVal columns As Variant) As cTable

    Dim hist As cTableHistory
    Set hist = Me.history

    Set HistoryFor = hist.x_HistoryFor( _
                            Me, _
                            entityColumn, _
                            entityValues, _
                            columns)
End Function

' Mutation version counter. Cached helper objects (cGroupBy, cTableHistory)
' capture or observe this value so stale row/value-dependent state fails closed.
Public Property Get x_Version() As Long
    x_Version = m_Version
End Property

'==============================================================================
'
' ROW SELECTION
'
'==============================================================================
' ---------------------------------------------------------------------------
' head
'
' Returns the first n rows as a new cTable.
'
' n >= 0:
'   Return the first n rows.
'
' n < 0:
'   Return all rows except the final Abs(n) rows.
'
' The source table is unchanged.
' ---------------------------------------------------------------------------
Public Function head(Optional ByVal n As Long = 5) As cTable
    Dim result As cTable

    If m_ColCount = 0 Then
        Set result = New cTable
        Set head = result
        Exit Function
    End If

    Dim resultCount As Long

    If n >= 0 Then
        If n < m_RowCount Then
            resultCount = n
        Else
            resultCount = m_RowCount
        End If
    Else
        resultCount = m_RowCount + n

        If resultCount < 0 Then resultCount = 0
    End If

    Dim selectedRows() As Long
    ReDim selectedRows(1 To IIf(resultCount > 0, resultCount, 1))

    Dim r As Long
    For r = 1 To resultCount
        selectedRows(r) = r
    Next r

    Set result = p_BuildFromRows(selectedRows, resultCount)
    Set result.DataSource = Me.DataSource
    result.set_dtypes_from Me

    Set head = result
End Function
' ---------------------------------------------------------------------------
' tail
'
' Returns the final n rows as a new cTable.
'
' n >= 0:
'   Return the final n rows.
'
' n < 0:
'   Return all rows except the first Abs(n) rows.
'
' ---------------------------------------------------------------------------
Public Function tail(Optional ByVal n As Long = 5) As cTable
    Dim result As cTable

    If m_ColCount = 0 Then
        Set result = New cTable
        Set tail = result
        Exit Function
    End If

    Dim resultCount As Long
    Dim firstSourceRow As Long

    If n >= 0 Then
        If n < m_RowCount Then
            resultCount = n
        Else
            resultCount = m_RowCount
        End If

        firstSourceRow = m_RowCount - resultCount + 1
    Else
        firstSourceRow = Abs(n) + 1

        If firstSourceRow <= m_RowCount Then
            resultCount = m_RowCount - firstSourceRow + 1
        Else
            resultCount = 0
        End If
    End If

    Dim selectedRows() As Long
    ReDim selectedRows(1 To IIf(resultCount > 0, resultCount, 1))

    Dim outputRow As Long
    Dim sourceRow As Long

    sourceRow = firstSourceRow

    For outputRow = 1 To resultCount
        selectedRows(outputRow) = sourceRow
        sourceRow = sourceRow + 1
    Next outputRow

    Set result = p_BuildFromRows(selectedRows, resultCount)
    Set result.DataSource = Me.DataSource
    result.set_dtypes_from Me

    Set tail = result
End Function
' ---------------------------------------------------------------------------
' nlargest
'
' Returns the n rows having the largest values in the specified columns.
' Multiple columns are evaluated from left to right.
'
' columns:
'   One column name:
'       t.nlargest(5, "Spending")
'
'   Multiple column names:
'       t.nlargest(5, Array("Spending", "AuthorizedSpending"))
'
' keep:
'   "first" = retain earlier source rows when ranking keys are tied
'   "last"  = retain later source rows when ranking keys are tied
'   "all"   = include every row tied with the nth row
'
' Rows containing NA in a ranking column are excluded.
' ---------------------------------------------------------------------------
Public Function nlargest(ByVal n As Long, _
                         ByVal columns As Variant, _
                         Optional ByVal keep As String = "first") As cTable
    Set nlargest = p_RankedRows( _
                        n:=n, _
                        columns:=columns, _
                        largest:=True, _
                        keep:=keep)
End Function
' ---------------------------------------------------------------------------
' nsmallest
'
' Returns the n rows having the smallest values in the specified columns.
' Parameters and behavior match nlargest.
' ---------------------------------------------------------------------------
Public Function nsmallest(ByVal n As Long, _
                          ByVal columns As Variant, _
                          Optional ByVal keep As String = "first") As cTable
    Set nsmallest = p_RankedRows( _
                         n:=n, _
                         columns:=columns, _
                         largest:=False, _
                         keep:=keep)
End Function


'==============================================================================
'
' ROW MANIPULATION
'
'==============================================================================
' ---------------------------------------------------------------------------
' drop_duplicates
'
' Removes duplicate rows using pandas-style subset, keep, and inplace ' arguments.
'
' subset:
'   Omitted:
'       Compare all columns.
'
'   One column:
'       t.drop_duplicates(subset:="CustomerID")
'
'   Multiple columns:
'       t.drop_duplicates( _
'           subset:=Array("CustomerID", "OrderDate"))
'
' keep:
'   "first" = retain the first row in each duplicate group
'   "last"  = retain the last row in each duplicate group
'   False   = remove every row belonging to a duplicate group
'
' inplace:
'   False = return a new cTable; source is unchanged
'   True  = modify this cTable and return Nothing, matching pandas
'
' The original row order is preserved.
' ---------------------------------------------------------------------------
Public Function drop_duplicates( _
        Optional ByVal subset As Variant, _
        Optional ByVal keep As Variant = "first", _
        Optional ByVal inplace As Boolean = False) As cTable

    Dim result As cTable

    ' A table without columns cannot contain meaningful duplicate rows.
    If m_ColCount = 0 Then
        If inplace Then
            Set drop_duplicates = Nothing
        Else
            Set drop_duplicates = Me.Clone
        End If

        Exit Function
    End If

    p_EnsureTypes

    Dim keepMode As Long
    keepMode = p_NormalizeDropDuplicatesKeep(keep)

    ' ------------------------------------------------------------
    ' Resolve subset into 1-based column numbers.
    ' ------------------------------------------------------------
    Dim useCols() As Long
    Dim useColCount As Long
    Dim i As Long
    Dim colName As String
    Dim ci As Long

    If IsMissing(subset) Or IsEmpty(subset) Or IsNull(subset) Then
        useColCount = m_ColCount
        ReDim useCols(1 To useColCount)

        For i = 1 To useColCount
            useCols(i) = i
        Next i

    ElseIf IsArray(subset) Then
        If Not p_ArrayHasElements(subset) Then
            Err.Raise vbObjectError + 1243, _
                      "cTable.drop_duplicates", _
                      "subset cannot be an empty array."
        End If

        useColCount = UBound(subset) - LBound(subset) + 1
        ReDim useCols(1 To useColCount)

        For i = 1 To useColCount
            colName = p_LabelColumnRef(subset(LBound(subset) + i - 1))
            ci = p_ColIdx(colName)

            If ci = 0 Then
                Err.Raise vbObjectError + 1244, _
                          "cTable.drop_duplicates", _
                          "Column not found: """ & colName & """"
            End If

            useCols(i) = ci
        Next i

    Else
        If IsObject(subset) Then
            Err.Raise vbObjectError + 1245, _
                      "cTable.drop_duplicates", _
                      "subset must be a column name or a one-dimensional array."
        End If

        useColCount = 1
        ReDim useCols(1 To 1)

        colName = CStr(subset)
        ci = p_ColIdx(colName)

        If ci = 0 Then
            Err.Raise vbObjectError + 1244, _
                      "cTable.drop_duplicates", _
                      "Column not found: """ & colName & """"
        End If

        useCols(1) = ci
    End If

    ' ------------------------------------------------------------
    ' Identify the rows that should remain.
    ' ------------------------------------------------------------
    Dim selectedRows() As Long
    ReDim selectedRows(1 To IIf(m_RowCount > 0, m_RowCount, 1))

    Dim selectedCount As Long
    Dim rowNum As Long
    Dim rowKey As String

    Select Case keepMode
        ' --------------------------------------------------------
        ' Keep first occurrence.
        ' --------------------------------------------------------
        Case 1
            Dim seen As Object
            Set seen = CreateObject("Scripting.Dictionary")
            seen.CompareMode = vbTextCompare

            For rowNum = 1 To m_RowCount
                rowKey = p_DuplicateRowKey( _
                            rowNum, _
                            useCols, _
                            useColCount)

                If Not seen.Exists(rowKey) Then
                    seen.Add rowKey, True

                    selectedCount = selectedCount + 1
                    selectedRows(selectedCount) = rowNum
                End If
            Next rowNum

        ' --------------------------------------------------------
        ' Keep last occurrence.
        ' --------------------------------------------------------
        Case -1
            Dim lastRows As Object
            Set lastRows = CreateObject("Scripting.Dictionary")
            lastRows.CompareMode = vbTextCompare

            ' Record the final row number for each duplicate key.
            For rowNum = 1 To m_RowCount
                rowKey = p_DuplicateRowKey( _
                            rowNum, _
                            useCols, _
                            useColCount)

                If lastRows.Exists(rowKey) Then
                    lastRows(rowKey) = rowNum
                Else
                    lastRows.Add rowKey, rowNum
                End If
            Next rowNum

            ' Second pass preserves the original row ordering.
            For rowNum = 1 To m_RowCount
                rowKey = p_DuplicateRowKey( _
                            rowNum, _
                            useCols, _
                            useColCount)

                If CLng(lastRows(rowKey)) = rowNum Then
                    selectedCount = selectedCount + 1
                    selectedRows(selectedCount) = rowNum
                End If
            Next rowNum

        ' --------------------------------------------------------
        ' keep=False: retain only keys appearing exactly once.
        ' --------------------------------------------------------
        Case 0
            Dim rowCounts As Object
            Set rowCounts = CreateObject("Scripting.Dictionary")
            rowCounts.CompareMode = vbTextCompare

            For rowNum = 1 To m_RowCount
                rowKey = p_DuplicateRowKey( _
                            rowNum, _
                            useCols, _
                            useColCount)

                If rowCounts.Exists(rowKey) Then
                    rowCounts(rowKey) = CLng(rowCounts(rowKey)) + 1
                Else
                    rowCounts.Add rowKey, 1&
                End If
            Next rowNum

            For rowNum = 1 To m_RowCount
                rowKey = p_DuplicateRowKey( _
                            rowNum, _
                            useCols, _
                            useColCount)

                If CLng(rowCounts(rowKey)) = 1 Then
                    selectedCount = selectedCount + 1
                    selectedRows(selectedCount) = rowNum
                End If
            Next rowNum
    End Select

    ' ------------------------------------------------------------
    ' Return a new table or modify this table.
    ' ------------------------------------------------------------
    If inplace Then
        p_KeepRowsInPlace selectedRows, selectedCount

        ' pandas returns None when inplace=True.
        Set drop_duplicates = Nothing
    Else
        Set result = p_BuildFromRows( _
                        selectedRows, _
                        selectedCount)

        Set result.DataSource = Me.DataSource
        result.set_dtypes_from Me

        Set drop_duplicates = result
    End If
End Function

' Delete one row in place using its 1-based row number.
Public Sub DropRowInPlace(ByVal rowNum As Long)
    If rowNum < 1 Or rowNum > m_RowCount Then
        Err.Raise vbObjectError + 1010, _
                  "cTable.DropRowInPlace", _
                  "Row number out of bounds: " & rowNum
    End If

    If m_RowCount = 1 Then
        Erase m_Data
        m_RowCount = 0
        If Not m_Lin Is Nothing Then m_Lin.x_RemoveRow rowNum
        p_ClearIndexes
        Exit Sub
    End If

    Dim newData() As Variant
    ReDim newData(1 To m_RowCount - 1, 1 To m_ColCount)

    Dim sourceRow As Long
    Dim targetRow As Long
    Dim c As Long

    targetRow = 0

    For sourceRow = 1 To m_RowCount
        If sourceRow <> rowNum Then
            targetRow = targetRow + 1

            For c = 1 To m_ColCount
                newData(targetRow, c) = m_Data(sourceRow, c)
            Next c
        End If
    Next sourceRow

    m_Data = newData
    m_RowCount = m_RowCount - 1

    If Not m_Lin Is Nothing Then m_Lin.x_RemoveRow rowNum

    ' Existing indexes contain obsolete row numbers.
    p_ClearIndexes
End Sub

Public Sub AddRow(Optional ByVal rowData As Variant)
    If m_ColCount = 0 Then Exit Sub

    Dim newRC As Long: newRC = m_RowCount + 1
    Dim newD() As Variant
    ReDim newD(1 To newRC, 1 To m_ColCount)

    Dim r As Long, c As Long
    For r = 1 To m_RowCount
        For c = 1 To m_ColCount
            newD(r, c) = m_Data(r, c)
        Next c
    Next r

    If IsMissing(rowData) Or IsEmpty(rowData) Then
        For c = 1 To m_ColCount: newD(newRC, c) = Null: Next c
    ElseIf IsArray(rowData) Then
        Dim lo As Long: lo = LBound(rowData)
        Dim rLen As Long: rLen = UBound(rowData) - lo + 1
        For c = 1 To m_ColCount
            If c <= rLen Then
                newD(newRC, c) = rowData(lo + c - 1)
            Else
                newD(newRC, c) = Null
            End If
        Next c
    Else
        newD(newRC, 1) = rowData
        For c = 2 To m_ColCount: newD(newRC, c) = Null: Next c
    End If

    m_Data = newD
    m_RowCount = newRC
    If Not m_Lin Is Nothing Then m_Lin.x_AppendRow ctSrcManual
    p_ClearIndexes
End Sub
' ---------------------------------------------------------------------------
' concat
'
' Returns a new cTable containing the rows of this table followed by the ' rows of another cTable. Neither source table is modified.
'
' This is the non-mutating counterpart to Append.
'
' strict=True:
'   Both tables must contain exactly the same column names, although their
'   physical column order may differ.
'
' strict=False:
'   Columns from the right table that do not exist in the left table are
'   ignored. Left-table columns missing from the right table are filled
'   with Null.
'
' Row positions in the returned table are automatically numbered ' 1 To RowCount. cTable does not currently maintain pandas-style row labels, ' so an ignore_index parameter is unnecessary.
'
' Example:
'   Set combined = January.concat(February)
' ---------------------------------------------------------------------------
Public Function concat( _
        ByVal other As cTable, _
        Optional ByVal strict As Boolean = True) As cTable

    If other Is Nothing Then
        Err.Raise vbObjectError + 1210, _
                  "cTable.concat", _
                  "The table to concatenate cannot be Nothing."
    End If

    Dim result As cTable

    ' Handle an empty left table.
    If m_ColCount = 0 Then
        Set result = other.Clone
        Set concat = result
        Exit Function
    End If

    ' In strict mode, require identical column sets.
    If strict Then
        Dim missingFromRight As Variant
        Dim missingFromLeft As Variant

        missingFromRight = other.MissingColumns(Me.headers)
        missingFromLeft = Me.MissingColumns(other.headers)

        If p_ArrayHasElements(missingFromRight) Then
            Err.Raise vbObjectError + 1211, _
                      "cTable.concat", _
                      "The right table is missing columns required by " & _
                      "the left table: " & _
                      p_JoinVariantArray(missingFromRight, ", ")
        End If

        If p_ArrayHasElements(missingFromLeft) Then
            Err.Raise vbObjectError + 1212, _
                      "cTable.concat", _
                      "The right table contains columns not present in " & _
                      "the left table: " & _
                      p_JoinVariantArray(missingFromLeft, ", ")
        End If
    End If

    Set result = Me.Clone

    ' Append aligns matching columns by name.
    result.Append other, strict

    Set concat = result
End Function

Public Sub Append(other As cTable, Optional ByVal strict As Boolean = True)
    If other.rowCount = 0 Then Exit Sub

    Dim colMap() As Long
    ReDim colMap(1 To other.ColumnCount)
    Dim k As Long
    For k = 1 To other.ColumnCount
        colMap(k) = p_ColIdx(other.columnName(k))
        If strict And colMap(k) = 0 Then
            Err.Raise vbObjectError + 90, "cTable.Append", _
                      "Column """ & other.columnName(k) & """ not found in target table"
        End If
    Next k

    Dim newRC As Long: newRC = m_RowCount + other.rowCount
    Dim newD() As Variant: ReDim newD(1 To newRC, 1 To m_ColCount)

    Dim r As Long, c As Long
    For r = 1 To m_RowCount
        For c = 1 To m_ColCount: newD(r, c) = m_Data(r, c): Next c
    Next r

    Dim destRow As Long: destRow = m_RowCount
    For r = 1 To other.rowCount
        destRow = destRow + 1
        For c = 1 To m_ColCount: newD(destRow, c) = Null: Next c
        For k = 1 To other.ColumnCount
            Dim ci As Long: ci = colMap(k)
            If ci > 0 Then newD(destRow, ci) = other.GetValue(r, k)
        Next k
    Next r

    m_Data = newD
    m_RowCount = newRC

    ' colMap(k) is this table's column for other's column k, which is exactly
    ' what x_AppendFrom needs. Left columns with no counterpart get Missing.
    If Not m_Lin Is Nothing Then
        m_Lin.x_AppendFrom p_LinOther(other), colMap, other.ColumnCount
    End If

    p_ClearIndexes
End Sub

' Replace all Null / empty values in a column with fillValue.
Public Sub FillNull(ByVal colName As Variant, ByVal fillValue As Variant)
    colName = p_LabelColumnRef(colName)
    Dim ci As Long: ci = p_ColIdx(colName)
    If ci = 0 Then
        Err.Raise vbObjectError + 100, "cTable.FillNull", _
                  "Column not found: """ & colName & """"
    End If
    Dim r As Long
    For r = 1 To m_RowCount
        If IsNull(m_Data(r, ci)) Or m_Data(r, ci) = "" Then
            m_Data(r, ci) = fillValue
            If Not m_Lin Is Nothing Then m_Lin.SetCell r, ci, ctSrcLiteral
        End If
    Next r
    p_ClearIndexes
End Sub

Public Function unique(ByVal colName As Variant) As Variant
    colName = p_LabelColumnRef(colName)
    Dim ci As Long: ci = p_ColIdx(colName)
    If ci = 0 Then
        Err.Raise vbObjectError + 110, "cTable.unique", _
                  "Column not found: """ & colName & """"
    End If
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    Dim r As Long
    For r = 1 To m_RowCount
        Dim k As String: k = CStr(m_Data(r, ci))
        If Not d.Exists(k) Then d(k) = m_Data(r, ci)
    Next r
    unique = d.Items
End Function

'==============================================================================
'
' OUTPUT
'
'==============================================================================

' Write the table to an Excel range.
' target is the top-left cell of the destination.
' includeHeaders=True (default) writes the header row first.
' ApplyFormats=True (default) applies an Excel NumberFormat to each column
' based on its dtype. ctVariant columns are untouched, so untyped tables
' get exact V1 behavior. Pass ApplyFormats:=False to skip.
Public Sub to_excel_range(ByVal target As Range, _
                         Optional ByVal includeHeaders As Boolean = True, _
                         Optional ByVal ApplyFormats As Boolean = True, _
                         Optional ByVal dateFormat As String = "mm/dd/yyyy")
    If target Is Nothing Or m_ColCount = 0 Then Exit Sub

    Dim startRow As Long: startRow = 0
    Dim c As Long
    If includeHeaders Then
        For c = 1 To m_ColCount
            target.Cells(1, c).value = m_Headers(c)
        Next c
        startRow = 1
    End If

    If m_RowCount > 0 Then
        Dim Out() As Variant
        ReDim Out(1 To m_RowCount, 1 To m_ColCount)
        Dim r As Long
        For r = 1 To m_RowCount
            For c = 1 To m_ColCount
                Out(r, c) = IIf(IsNull(m_Data(r, c)), "", m_Data(r, c))
            Next c
        Next r
        target.Cells(startRow + 1, 1).Resize(m_RowCount, m_ColCount).value = Out
    End If

    If ApplyFormats Then
        p_EnsureTypes
        For c = 1 To m_ColCount
            If m_ColTypes(c) <> ctVariant Then
                target.Cells(startRow + 1, c) _
                      .Resize(IIf(m_RowCount > 0, m_RowCount, 1), 1) _
                      .NumberFormat = p_NumberFormatFor(m_ColTypes(c), dateFormat)
            End If
        Next c
    End If
End Sub
' ---------------------------------------------------------------------------
' to_json
'
' Write this cTable to a UTF-8 JSON file.
'
' orient:
'   "records"  -> [{"Col1":v1,"Col2":v2}, ...]
'   "columns"  -> {"Col1":[v1,v2,...],"Col2":[v1,v2,...]}
'
' indent:
'   0 = compact JSON
'   >0 = pretty-print using that many spaces per nesting level
'
' date_format:
'   "iso"    -> "2026-08-28T14:30:00"
'   "serial" -> VBA / Excel date serial number
'
' force_ascii:
'   False = write Unicode characters directly
'   True  = escape non-ASCII characters as \uXXXX
'
' Examples:
'
'   t.to_json "C:\Temp\data.json"
'
'   t.to_json _
'       "C:\Temp\data.json", _
'       orient:="records", _
'       indent:=4
'
' ---------------------------------------------------------------------------
Public Sub to_json( _
        ByVal filePath As String, _
        Optional ByVal orient As String = "records", _
        Optional ByVal indent As Long = 2, _
        Optional ByVal date_format As String = "iso", _
        Optional ByVal force_ascii As Boolean = False)

    Const CALLER As String = "cTable.to_json"
    Const adTypeText As Long = 2
    Const adSaveCreateOverWrite As Long = 2

    If Len(Trim$(filePath)) = 0 Then
        Err.Raise vbObjectError + 1660, CALLER, _
                  "filePath cannot be blank."
    End If

    If indent < 0 Then
        Err.Raise vbObjectError + 1661, CALLER, _
                  "indent cannot be negative."
    End If

    Dim orientName As String
    orientName = LCase$(Trim$(orient))

    Select Case orientName
        Case "records", "columns"
            ' valid
        Case Else
            Err.Raise vbObjectError + 1662, CALLER, _
                      "orient must be ""records"" or ""columns""."
    End Select

    Dim dateFormatName As String
    dateFormatName = LCase$(Trim$(date_format))

    Select Case dateFormatName
        Case "iso", "serial"
            ' valid
        Case Else
            Err.Raise vbObjectError + 1663, CALLER, _
                      "date_format must be ""iso"" or ""serial""."
    End Select

    p_EnsureTypes

    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")

    On Error GoTo ErrHandler

    stm.Type = adTypeText
    stm.Charset = "utf-8"
    stm.Open

    Select Case orientName
        Case "records"
            p_WriteJsonRecords _
                stm, _
                indent, _
                dateFormatName, _
                force_ascii

        Case "columns"
            p_WriteJsonColumns _
                stm, _
                indent, _
                dateFormatName, _
                force_ascii
    End Select

    stm.SaveToFile filePath, adSaveCreateOverWrite

CleanExit:
    On Error Resume Next
    If Not stm Is Nothing Then
        If stm.State <> 0 Then stm.Close
    End If
    Set stm = Nothing
    On Error GoTo 0
    Exit Sub

ErrHandler:
    Dim errNumber As Long
    Dim errDescription As String

    errNumber = Err.Number
    errDescription = Err.Description

    Resume CleanFail

CleanFail:
    On Error Resume Next
    If Not stm Is Nothing Then
        If stm.State <> 0 Then stm.Close
    End If
    Set stm = Nothing
    On Error GoTo 0

    Err.Raise errNumber, CALLER, _
              "Could not write JSON file """ & filePath & """: " & _
              errDescription
End Sub
'==============================================================================
' JSON OUTPUT HELPERS
'==============================================================================
Private Sub p_WriteJsonRecords( _
        ByVal stm As Object, _
        ByVal indent As Long, _
        ByVal dateFormat As String, _
        ByVal forceAscii As Boolean)

    Dim pretty As Boolean
    pretty = (indent > 0)

    Dim r As Long
    Dim c As Long

    stm.WriteText "["

    For r = 1 To m_RowCount

        If r > 1 Then
            stm.WriteText ","
        End If

        If pretty Then
            stm.WriteText vbCrLf & Space$(indent)
        End If

        stm.WriteText "{"

        For c = 1 To m_ColCount

            If c > 1 Then
                stm.WriteText ","
            End If

            If pretty Then
                stm.WriteText vbCrLf & Space$(indent * 2)
            End If

            stm.WriteText _
                p_JsonQuote(m_Headers(c), forceAscii) & _
                ":" & IIf(pretty, " ", "") & _
                p_JsonValue( _
                    m_Data(r, c), _
                    m_ColTypes(c), _
                    dateFormat, _
                    forceAscii)

        Next c

        If pretty And m_ColCount > 0 Then
            stm.WriteText vbCrLf & Space$(indent)
        End If

        stm.WriteText "}"

    Next r

    If pretty And m_RowCount > 0 Then
        stm.WriteText vbCrLf
    End If

    stm.WriteText "]"
End Sub


Private Sub p_WriteJsonColumns( _
        ByVal stm As Object, _
        ByVal indent As Long, _
        ByVal dateFormat As String, _
        ByVal forceAscii As Boolean)

    Dim pretty As Boolean
    pretty = (indent > 0)

    Dim c As Long
    Dim r As Long

    stm.WriteText "{"

    For c = 1 To m_ColCount

        If c > 1 Then
            stm.WriteText ","
        End If

        If pretty Then
            stm.WriteText vbCrLf & Space$(indent)
        End If

        stm.WriteText _
            p_JsonQuote(m_Headers(c), forceAscii) & _
            ":" & IIf(pretty, " ", "") & "["

        For r = 1 To m_RowCount

            If r > 1 Then
                stm.WriteText ","
            End If

            If pretty Then
                stm.WriteText vbCrLf & Space$(indent * 2)
            End If

            stm.WriteText _
                p_JsonValue( _
                    m_Data(r, c), _
                    m_ColTypes(c), _
                    dateFormat, _
                    forceAscii)

        Next r

        If pretty And m_RowCount > 0 Then
            stm.WriteText vbCrLf & Space$(indent)
        End If

        stm.WriteText "]"

    Next c

    If pretty And m_ColCount > 0 Then
        stm.WriteText vbCrLf
    End If

    stm.WriteText "}"
End Sub


Private Function p_JsonValue( _
        ByVal v As Variant, _
        ByVal declaredType As ctDataType, _
        ByVal dateFormat As String, _
        ByVal forceAscii As Boolean) As String

    ' cTable NA -> JSON null
    If p_IsNA(v) Then
        p_JsonValue = "null"
        Exit Function
    End If

    Select Case declaredType

        Case ctString
            p_JsonValue = p_JsonQuote(CStr(v), forceAscii)

        Case ctLong
            p_JsonValue = CStr(CLng(v))

        Case ctDouble
            p_JsonValue = p_JsonNumber(CDbl(v))

        Case ctCurrency
            p_JsonValue = p_JsonNumber(CDbl(CCur(v)))

        Case ctDate
            p_JsonValue = _
                p_JsonDate(CDate(v), dateFormat, forceAscii)

        Case ctBoolean
            If CBool(v) Then
                p_JsonValue = "true"
            Else
                p_JsonValue = "false"
            End If

        Case Else
            ' ctVariant: infer from the actual Variant subtype.
            Select Case VarType(v)

                Case vbBoolean
                    If CBool(v) Then
                        p_JsonValue = "true"
                    Else
                        p_JsonValue = "false"
                    End If

                Case vbDate
                    p_JsonValue = _
                        p_JsonDate(CDate(v), dateFormat, forceAscii)

                Case vbByte, vbInteger, vbLong, _
                     vbSingle, vbDouble, vbCurrency, vbDecimal

                    p_JsonValue = p_JsonNumber(CDbl(v))

                Case Else
                    p_JsonValue = _
                        p_JsonQuote(CStr(v), forceAscii)
            End Select
    End Select
End Function

Private Function p_JsonDate( _
        ByVal d As Date, _
        ByVal dateFormat As String, _
        ByVal forceAscii As Boolean) As String

    If dateFormat = "serial" Then
        p_JsonDate = p_JsonNumber(CDbl(d))
    Else
        p_JsonDate = _
            p_JsonQuote( _
                Format$(d, "yyyy-mm-dd\THH:nn:ss"), _
                forceAscii)
    End If
End Function
Private Function p_JsonNumber(ByVal value As Double) As String

    Dim s As String
    s = Trim$(Str$(value))

    ' JSON requires a digit before the decimal point.
    If Left$(s, 1) = "." Then
        s = "0" & s
    ElseIf Left$(s, 2) = "-." Then
        s = "-0" & mid$(s, 2)
    End If

    p_JsonNumber = s
End Function

Private Function p_JsonQuote( _
        ByVal value As String, _
        ByVal forceAscii As Boolean) As String

    Dim sb As String
    Dim i As Long
    Dim ch As String
    Dim code As Long

    sb = """"

    For i = 1 To Len(value)

        ch = mid$(value, i, 1)
        code = AscW(ch)

        If code < 0 Then code = code + 65536

        Select Case code

            Case 34
                sb = sb & "\"""

            Case 92
                sb = sb & "\\"

            Case 8
                sb = sb & "\b"

            Case 9
                sb = sb & "\t"

            Case 10
                sb = sb & "\n"

            Case 12
                sb = sb & "\f"

            Case 13
                sb = sb & "\r"

            Case 0 To 31
                sb = sb & "\u" & right$("0000" & Hex$(code), 4)

            Case Else
                If forceAscii And code > 127 Then
                    sb = sb & _
                         "\u" & right$("0000" & Hex$(code), 4)
                Else
                    sb = sb & ch
                End If

        End Select
    Next i

    p_JsonQuote = sb & """"
End Function



'==============================================================================
'
' CLONE
'
'==============================================================================

Public Function Clone() As cTable
    Set Clone = New cTable
    If m_ColCount = 0 Then Exit Function
    Dim ph() As Variant: ReDim ph(1 To 1, 1 To m_ColCount)
    If m_RowCount = 0 Then
        Clone.from_array m_Headers, ph, 0, m_ColCount
    Else
        Clone.from_array m_Headers, m_Data, m_RowCount, m_ColCount
    End If
    Set Clone.DataSource = Me.DataSource
    Clone.set_dtypes_from Me
    If Not m_Lin Is Nothing Then Clone.x_SetLineage m_Lin.CloneMe
End Function

'==============================================================================
'
' DIAGNOSTICS
'
'==============================================================================

Public Function Describe() As String

    Dim sb As String
    Dim i As Long

    sb = "cTable  [" & _
         m_RowCount & _
         " rows x " & _
         m_ColCount & _
         " columns]" & vbCrLf

    If m_ColCount > 0 Then
        sb = sb & "Columns : "

        For i = 1 To m_ColCount
            sb = sb & m_Headers(i)

            If i < m_ColCount Then
                sb = sb & ", "
            End If
        Next i

        sb = sb & vbCrLf
    End If

    If m_RowCount > 0 Then

        sb = sb & "First row: "

        For i = 1 To m_ColCount
            sb = sb & _
                 "[" & _
                 p_DisplayValue(m_Data(1, i)) & _
                 "]"

            If i < m_ColCount Then
                sb = sb & "  "
            End If
        Next i

        sb = sb & vbCrLf

        If m_RowCount > 1 Then

            sb = sb & "Last row : "

            For i = 1 To m_ColCount
                sb = sb & _
                     "[" & _
                     p_DisplayValue( _
                         m_Data(m_RowCount, i)) & _
                     "]"

                If i < m_ColCount Then
                    sb = sb & "  "
                End If
            Next i

        End If
    End If

    Describe = sb

End Function

Public Function ToString( _
        Optional ByVal maxRows As Long = 50) As String

    Dim sb As String
    Dim i As Long
    Dim c As Long
    Dim limit As Long

    ' ------------------------------------------------------------
    ' Headers
    ' ------------------------------------------------------------
    For c = 1 To m_ColCount
        sb = sb & m_Headers(c)

        If c < m_ColCount Then
            sb = sb & vbTab
        End If
    Next c

    If m_ColCount > 0 Then
        sb = sb & vbCrLf
    End If

    ' ------------------------------------------------------------
    ' Determine how many rows to display.
    '
    ' maxRows <= 0 means display all rows.
    ' ------------------------------------------------------------
    If maxRows <= 0 Then
        limit = m_RowCount

    ElseIf maxRows > m_RowCount Then
        limit = m_RowCount

    Else
        limit = maxRows
    End If

    ' ------------------------------------------------------------
    ' Data
    ' ------------------------------------------------------------
    For i = 1 To limit

        For c = 1 To m_ColCount

            sb = sb & _
                 p_DisplayValue(m_Data(i, c))

            If c < m_ColCount Then
                sb = sb & vbTab
            End If

        Next c

        sb = sb & vbCrLf

    Next i

    ' ------------------------------------------------------------
    ' Truncation notice
    ' ------------------------------------------------------------
    If limit < m_RowCount Then
        sb = sb & _
             "... (" & _
             (m_RowCount - limit) & _
             " more rows)"
    End If

    ToString = sb

End Function



Private Function p_DisplayValue(ByVal v As Variant) As String
    If IsNull(v) Then
        p_DisplayValue = ""
    ElseIf IsEmpty(v) Then
        p_DisplayValue = ""
    ElseIf IsError(v) Then
        p_DisplayValue = "#ERROR"
    Else
        p_DisplayValue = CStr(v)
    End If
End Function

' ---------------------------------------------------------------------------
' ReplaceExcelTableFrom_cTable
'   Replaces all data in the given Excel ListObject (table) with the contents
'   of this cTable, matching columns by name (case-insensitive).
'   ApplyFormats=True (default) applies dtype number formats to the
'   ListObject data body, matched by column name.
' ---------------------------------------------------------------------------
Public Sub ReplaceExcelTableFrom_cTable(lo As ListObject, _
                                        Optional RedIfEmptyCols As Variant = Empty, _
                                        Optional ByVal ApplyFormats As Boolean = True, _
                                        Optional ByVal dateFormat As String = "mm/dd/yyyy")
    Dim excelColCount As Long
    Dim excelHeaders() As String
    Dim colMap() As Long
    Dim r As Long, c As Long

    excelColCount = lo.ListColumns.count

    ReDim excelHeaders(1 To excelColCount)
    ReDim colMap(1 To excelColCount)
    For c = 1 To excelColCount
        excelHeaders(c) = lo.ListColumns(c).name
        colMap(c) = Me.columnNumber(excelHeaders(c))
    Next c

    If Not lo.DataBodyRange Is Nothing Then
        lo.DataBodyRange.ClearFormats
        lo.DataBodyRange.ClearContents
    End If

    Dim newRows As Long
    newRows = Me.rowCount
    lo.Resize lo.Range.Resize(newRows + 1, excelColCount)

    If newRows = 0 Then Exit Sub

    Dim outArr() As Variant
    ReDim outArr(1 To newRows, 1 To excelColCount)
    For r = 1 To newRows
        For c = 1 To excelColCount
            If colMap(c) > 0 Then
                outArr(r, c) = Me.GetValue(r, colMap(c))
            Else
                outArr(r, c) = ""
            End If
        Next c
    Next r

    ' NA -> blank cells
    For r = 1 To newRows
        For c = 1 To excelColCount
            If IsNull(outArr(r, c)) Then outArr(r, c) = ""
        Next c
    Next r

    lo.DataBodyRange.value = outArr

    ' Dtype-driven number formats
    If ApplyFormats Then
        p_EnsureTypes
        For c = 1 To excelColCount
            If colMap(c) > 0 Then
                If m_ColTypes(colMap(c)) <> ctVariant Then
                    lo.ListColumns(c).DataBodyRange.NumberFormat = _
                        p_NumberFormatFor(m_ColTypes(colMap(c)), dateFormat)
                End If
            End If
        Next c
    End If

    ' Optional: Highlight specified columns in red if empty/null/""
    If Not IsEmpty(RedIfEmptyCols) Then
        Dim highlightCols As Object
        Set highlightCols = CreateObject("Scripting.Dictionary")
        highlightCols.CompareMode = vbTextCompare
        Dim arrIdx As Long
        For arrIdx = LBound(RedIfEmptyCols) To UBound(RedIfEmptyCols)
            highlightCols(CStr(RedIfEmptyCols(arrIdx))) = True
        Next arrIdx

        For c = 1 To excelColCount
            If highlightCols.Exists(excelHeaders(c)) Then
                For r = 1 To newRows
                    Dim cellVal As Variant
                    cellVal = lo.DataBodyRange.Cells(r, c).value
                    If IsNull(cellVal) Or cellVal = "" Then
                        lo.DataBodyRange.Cells(r, c).Interior.color = vbRed
                    End If
                Next r
            End If
        Next c
    End If
End Sub

'==============================================================================
'
' AGGREGATION
'
'==============================================================================
' ---------------------------------------------------------------------------
' nunique
'
' Counts distinct values in a column, in every column, or across every row.
'
' Named-column mode:
'   t.nunique("Spending")
'       Returns a scalar Long.
'
' Whole-table axis=0 mode:
'   Set results = t.nunique()
'   Set results = t.nunique(axis:=0, dropna:=True)
'
'       Returns a Scripting.Dictionary:
'           column name -> unique-value count
'
' Whole-table axis=1 mode:
'   results = t.nunique(axis:=1)
'
'       Returns a 1-based Long array:
'           results(row number) -> unique values across that row
'
' axis:
'   0 or "index"   = one count per column
'   1 or "columns" = one count per row
'
' dropna:
'   True  = exclude missing values from the count
'   False = count missing values as one distinct value
'
' Note:
'   colName is a cTable convenience extension. When colName is supplied,
'   axis must be 0.
' ---------------------------------------------------------------------------
Public Function nunique( _
        Optional ByVal colName As Variant = "", _
        Optional ByVal axis As Variant = 0, _
        Optional ByVal dropna As Boolean = True) As Variant
    If p_HasOptionalCol(colName) Then colName = p_LabelColumnRef(colName)

    p_EnsureTypes

    Dim normalizedAxis As Long
    normalizedAxis = p_NormalizeNuniqueAxis(axis)

    ' ------------------------------------------------------------
    ' Named-column mode: return one Long.
    ' ------------------------------------------------------------
    If Len(Trim$(colName)) > 0 Then
        If normalizedAxis <> 0 Then
            Err.Raise vbObjectError + 1236, _
                      "cTable.nunique", _
                      "A named column can only be used with axis=0."
        End If

        Dim colNum As Long
        colNum = p_ColIdx(colName)

        If colNum = 0 Then
            Err.Raise vbObjectError + 1237, _
                      "cTable.nunique", _
                      "Column not found: """ & colName & """"
        End If

        nunique = p_NuniqueColumn(colNum, dropna)
        Exit Function
    End If

    ' ------------------------------------------------------------
    ' axis=0: return column name -> unique count.
    ' ------------------------------------------------------------
    If normalizedAxis = 0 Then
        Dim columnResults As Object
        Set columnResults = CreateObject("Scripting.Dictionary")
        columnResults.CompareMode = vbTextCompare

        Dim c As Long
        For c = 1 To m_ColCount
            columnResults.Add _
                m_Headers(c), _
                p_NuniqueColumn(c, dropna)
        Next c

        Set nunique = columnResults
        Exit Function
    End If

    ' ------------------------------------------------------------
    ' axis=1: return one count for every row.
    ' ------------------------------------------------------------
    If m_RowCount = 0 Then
        nunique = Array()
        Exit Function
    End If

    Dim rowResults() As Long
    ReDim rowResults(1 To m_RowCount)

    Dim r As Long
    For r = 1 To m_RowCount
        rowResults(r) = p_NuniqueRow(r, dropna)
    Next r

    nunique = rowResults
End Function

' ---------------------------------------------------------------------------
' sum
'   pandas-style column aggregation using the existing cTable dtype and
'   missing-value systems.
'
' Usage:
'   tbl.sum("Amount")
'       Returns a scalar Variant for one column.
'
'   tbl.sum(numeric_only:=True)
'       Returns a Scripting.Dictionary containing one result per numeric
'       column. This is the VBA equivalent of the Series returned by
'       pandas DataFrame.sum(axis=0).
'
' Parameters:
'   colName
'       Optional column name. If omitted, sums columns independently.
'
'   skipna
'       True  = ignore NA values.
'       False = return Null when any NA is encountered.
'
'   numeric_only
'       True  = whole-table mode includes only numeric/Boolean columns.
'       False = whole-table mode raises an error if a nonnumeric column
'               is encountered.
'
'       This parameter has no filtering effect when a specific column is
'       requested. An explicitly selected nonnumeric column raises an error.
'
'   min_count
'       Minimum number of valid values required. If fewer valid values exist,
'       the result is Null. Default 0 means an empty/all-NA column returns 0
'       when skipna=True.
'
' Returns:
'   Scalar Variant when colName is supplied.
'   Scripting.Dictionary when colName is omitted.
' ---------------------------------------------------------------------------
Public Function sum(Optional ByVal colName As Variant, _
                    Optional ByVal skipna As Boolean = True, _
                    Optional ByVal numeric_only As Boolean = False, _
                    Optional ByVal min_count As Long = 0) As Variant
    If p_HasOptionalCol(colName) Then colName = p_LabelColumnRef(colName)

    If min_count < 0 Then
        Err.Raise vbObjectError + 920, _
                  "cTable.sum", _
                  "min_count cannot be negative."
    End If

    p_EnsureTypes

    Dim hasColumn As Boolean
    hasColumn = Not IsMissing(colName)

    If hasColumn Then
        If IsEmpty(colName) Or IsNull(colName) Then
            hasColumn = False
        Else
            hasColumn = (Len(Trim$(CStr(colName))) > 0)
        End If
    End If

    ' ------------------------------------------------------------
    ' One-column mode: return a scalar.
    ' ------------------------------------------------------------
    If hasColumn Then
        Dim ci As Long
        ci = p_ColIdx(CStr(colName))

        If ci = 0 Then
            Err.Raise vbObjectError + 926, _
                      "cTable.sum", _
                      "Column not found: """ & CStr(colName) & """"
        End If

        If Not p_IsNumericDType(m_ColTypes(ci)) Then
            Err.Raise vbObjectError + 927, _
                      "cTable.sum", _
                      "Column """ & m_Headers(ci) & """ has dtype " & _
                      Me.dtype_name(m_ColTypes(ci)) & _
                      " and cannot be summed."
        End If

        sum = p_SumNumericColumn(ci, skipna, min_count)
        Exit Function
    End If

    ' ------------------------------------------------------------
    ' Whole-table mode: return column name -> sum dictionary.
    ' ------------------------------------------------------------
    Dim results As Scripting.Dictionary
    Set results = New Scripting.Dictionary
    results.CompareMode = vbTextCompare

    Dim c As Long
    For c = 1 To m_ColCount
        If p_IsNumericDType(m_ColTypes(c)) Then
            results.Add m_Headers(c), _
                        p_SumNumericColumn(c, skipna, min_count)

        ElseIf Not numeric_only Then
            Err.Raise vbObjectError + 928, _
                      "cTable.sum", _
                      "Column """ & m_Headers(c) & """ has dtype " & _
                      Me.dtype_name(m_ColTypes(c)) & _
                      " and cannot be summed. Use numeric_only:=True " & _
                      "to exclude nonnumeric columns."
        End If
    Next c

    Set sum = results
End Function

' ---------------------------------------------------------------------------
' median
' ---------------------------------------------------------------------------
Public Function median(Optional ByVal colName As Variant, _
                       Optional ByVal skipna As Boolean = False, _
                       Optional ByVal numeric_only As Boolean = False) As Variant
    If p_HasOptionalCol(colName) Then colName = p_LabelColumnRef(colName)
    If IsMissing(colName) Then
        Set median = p_AggregateAllColumns( _
                        "median", skipna, numeric_only, _
                        0.5, "linear", 1)

    ElseIf p_IsBlankColumnArgument(colName) Then
        Set median = p_AggregateAllColumns( _
                        "median", skipna, numeric_only, _
                        0.5, "linear", 1)

    Else
        median = p_AggregateNamedColumn( _
                    "median", CStr(colName), skipna, numeric_only, _
                    0.5, "linear", 1)
    End If
End Function


' ---------------------------------------------------------------------------
' mean
' ---------------------------------------------------------------------------
Public Function mean(Optional ByVal colName As Variant, _
                     Optional ByVal skipna As Boolean = False, _
                     Optional ByVal numeric_only As Boolean = False) As Variant
    If p_HasOptionalCol(colName) Then colName = p_LabelColumnRef(colName)
    If IsMissing(colName) Then
        Set mean = p_AggregateAllColumns( _
                        "mean", skipna, numeric_only, _
                        0.5, "linear", 1)

    ElseIf p_IsBlankColumnArgument(colName) Then
        Set mean = p_AggregateAllColumns( _
                        "mean", skipna, numeric_only, _
                        0.5, "linear", 1)

    Else
        mean = p_AggregateNamedColumn( _
                    "mean", CStr(colName), skipna, numeric_only, _
                    0.5, "linear", 1)
    End If
End Function


' ---------------------------------------------------------------------------
' min
' ---------------------------------------------------------------------------
Public Function min(Optional ByVal colName As Variant, _
                    Optional ByVal skipna As Boolean = False, _
                    Optional ByVal numeric_only As Boolean = False) As Variant
    If p_HasOptionalCol(colName) Then colName = p_LabelColumnRef(colName)
    If IsMissing(colName) Then
        Set min = p_AggregateAllColumns( _
                        "min", skipna, numeric_only, _
                        0.5, "linear", 1)

    ElseIf p_IsBlankColumnArgument(colName) Then
        Set min = p_AggregateAllColumns( _
                        "min", skipna, numeric_only, _
                        0.5, "linear", 1)

    Else
        min = p_AggregateNamedColumn( _
                    "min", CStr(colName), skipna, numeric_only, _
                    0.5, "linear", 1)
    End If
End Function


' ---------------------------------------------------------------------------
' max
' ---------------------------------------------------------------------------
Public Function max(Optional ByVal colName As Variant, _
                    Optional ByVal skipna As Boolean = False, _
                    Optional ByVal numeric_only As Boolean = False) As Variant
    If p_HasOptionalCol(colName) Then colName = p_LabelColumnRef(colName)
    If IsMissing(colName) Then
        Set max = p_AggregateAllColumns( _
                        "max", skipna, numeric_only, _
                        0.5, "linear", 1)

    ElseIf p_IsBlankColumnArgument(colName) Then
        Set max = p_AggregateAllColumns( _
                        "max", skipna, numeric_only, _
                        0.5, "linear", 1)

    Else
        max = p_AggregateNamedColumn( _
                    "max", CStr(colName), skipna, numeric_only, _
                    0.5, "linear", 1)
    End If
End Function


' ---------------------------------------------------------------------------
' var
'
' ddof=1 returns sample variance, matching pandas.
' ddof=0 returns population variance.
' ---------------------------------------------------------------------------
Public Function var(Optional ByVal colName As Variant, _
                    Optional ByVal skipna As Boolean = False, _
                    Optional ByVal numeric_only As Boolean = False, _
                    Optional ByVal ddof As Long = 1) As Variant
    If p_HasOptionalCol(colName) Then colName = p_LabelColumnRef(colName)
    If IsMissing(colName) Then
        Set var = p_AggregateAllColumns( _
                        "var", skipna, numeric_only, _
                        0.5, "linear", ddof)

    ElseIf p_IsBlankColumnArgument(colName) Then
        Set var = p_AggregateAllColumns( _
                        "var", skipna, numeric_only, _
                        0.5, "linear", ddof)

    Else
        var = p_AggregateNamedColumn( _
                    "var", CStr(colName), skipna, numeric_only, _
                    0.5, "linear", ddof)
    End If
End Function


' ---------------------------------------------------------------------------
' std
'
' ddof=1 returns sample standard deviation, matching pandas.
' ddof=0 returns population standard deviation.
' ---------------------------------------------------------------------------
Public Function std(Optional ByVal colName As Variant, _
                    Optional ByVal skipna As Boolean = False, _
                    Optional ByVal numeric_only As Boolean = False, _
                    Optional ByVal ddof As Long = 1) As Variant
    If p_HasOptionalCol(colName) Then colName = p_LabelColumnRef(colName)
    If IsMissing(colName) Then
        Set std = p_AggregateAllColumns( _
                        "std", skipna, numeric_only, _
                        0.5, "linear", ddof)

    ElseIf p_IsBlankColumnArgument(colName) Then
        Set std = p_AggregateAllColumns( _
                        "std", skipna, numeric_only, _
                        0.5, "linear", ddof)

    Else
        std = p_AggregateNamedColumn( _
                    "std", CStr(colName), skipna, numeric_only, _
                    0.5, "linear", ddof)
    End If
End Function


' ---------------------------------------------------------------------------
' quantile
'
' q:
'   0.00 = minimum
'   0.25 = first quartile
'   0.50 = median
'   0.75 = third quartile
'   1.00 = maximum
'
' interpolation:
'   linear, lower, higher, midpoint, nearest
'
' quantile always excludes NA values, following pandas; there is therefore ' no skipna argument.
' ---------------------------------------------------------------------------
Public Function quantile(Optional ByVal colName As Variant, _
                         Optional ByVal q As Double = 0.5, _
                         Optional ByVal numeric_only As Boolean = False, _
                         Optional ByVal interpolation As String = "linear") As Variant
    If p_HasOptionalCol(colName) Then colName = p_LabelColumnRef(colName)
    If q < 0# Or q > 1# Then
        Err.Raise vbObjectError + 947, _
                  "cTable.quantile", _
                  "q must be between 0 and 1."
    End If

    If IsMissing(colName) Then
        Set quantile = p_AggregateAllColumns( _
                            "quantile", True, numeric_only, _
                            q, interpolation, 1)

    ElseIf p_IsBlankColumnArgument(colName) Then
        Set quantile = p_AggregateAllColumns( _
                            "quantile", True, numeric_only, _
                            q, interpolation, 1)

    Else
        quantile = p_AggregateNamedColumn( _
                        "quantile", CStr(colName), True, numeric_only, _
                        q, interpolation, 1)
    End If
End Function

Public Function AggregateDict(GroupBy As String, aggValue As String, Optional aggType As String = "sum") As Scripting.Dictionary
    Dim dict As Scripting.Dictionary
    Set dict = New Scripting.Dictionary
    dict.CompareMode = vbTextCompare

    Dim tempDict As Scripting.Dictionary
    Set tempDict = New Scripting.Dictionary
    tempDict.CompareMode = vbTextCompare

    Dim groupCol As Long, aggCol As Long
    groupCol = p_ColIdx(GroupBy)
    aggCol = p_ColIdx(aggValue)
    If groupCol = 0 Or aggCol = 0 Then
        Err.Raise vbObjectError + 513, "cTable.AggregateDict", "Column not found"
    End If

    Dim r As Long
    Dim key As String, val As Variant

    For r = 1 To m_RowCount
        key = CStr(m_Data(r, groupCol))
        val = m_Data(r, aggCol)
        If Not tempDict.Exists(key) Then
            Set tempDict(key) = New Collection
        End If
        tempDict(key).Add val
    Next r

    Dim values As Collection
    Dim minVal As Variant, maxVal As Variant, sumVal As Double, avgVal As Double
    Dim groupKey As Variant

    For Each groupKey In tempDict.keys
        Set values = tempDict(groupKey)
        Select Case LCase(aggType)
            Case "sum", "avg"
                sumVal = 0
                Dim cnt As Long: cnt = 0
                For Each val In values
                    sumVal = sumVal + CleanCurrency(val)
                    cnt = cnt + 1
                Next
                If LCase(aggType) = "sum" Then
                    dict.Add groupKey, sumVal
                Else
                    If cnt > 0 Then
                        avgVal = sumVal / cnt
                        dict.Add groupKey, avgVal
                    Else
                        dict.Add groupKey, Null
                    End If
                End If
            Case "min"
                minVal = CleanCurrency(values(1))
                For Each val In values
                    If CleanCurrency(val) < minVal Then minVal = CleanCurrency(val)
                Next
                dict.Add groupKey, minVal
            Case "max"
                maxVal = CleanCurrency(values(1))
                For Each val In values
                    If CleanCurrency(val) > maxVal Then maxVal = CleanCurrency(val)
                Next
                dict.Add groupKey, maxVal
            Case Else
                dict.Add groupKey, values
        End Select
    Next groupKey

    Set AggregateDict = dict
End Function

Public Function GetUniqueValues(ByVal colName As Variant, Optional SortValues As Boolean = True) As Variant
    colName = p_LabelColumnRef(colName)
    Dim ci As Long: ci = p_ColIdx(colName)
    If ci = 0 Then
        Err.Raise vbObjectError + 110, "cTable.GetUniqueValues", _
                  "Column not found: """ & colName & """"
    End If

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Dim r As Long
    For r = 1 To m_RowCount
        Dim v As Variant
        v = m_Data(r, ci)
        If Not dict.Exists(CStr(v)) Then dict.Add CStr(v), v
    Next r

    Dim arr As Variant
    arr = dict.Items

    If SortValues Then
        Dim i As Long, j As Long
        Dim temp As Variant
        For i = LBound(arr) To UBound(arr) - 1
            For j = i + 1 To UBound(arr)
                If arr(i) > arr(j) Then
                    temp = arr(i)
                    arr(i) = arr(j)
                    arr(j) = temp
                End If
            Next j
        Next i
    End If
    GetUniqueValues = arr
End Function

'==============================================================================
'
' COLUMN REORDERING
'
'==============================================================================

Public Sub ReorderColumnRelativeInPlace(ByVal colName As Variant, posOffset As Integer, Optional RelColName As String = "")
    colName = p_LabelColumnRef(colName)
    Dim srcIdx As Long, relIdx As Long, newIdx As Long
    Dim colCount As Long: colCount = m_ColCount
    Dim i As Long

    srcIdx = p_ColIdx(colName)
    If srcIdx = 0 Then
        Err.Raise vbObjectError + 701, "cTable.ReorderColumnRelativeInPlace", _
                  "Column not found: """ & colName & """"
    End If

    If Len(RelColName) > 0 Then
        relIdx = p_ColIdx(RelColName)
        If relIdx = 0 Then
            Err.Raise vbObjectError + 702, "cTable.ReorderColumnRelativeInPlace", _
                      "Relative column not found: """ & RelColName & """"
        End If
    Else
        relIdx = srcIdx
    End If

    Dim tempOrder() As Long
    ReDim tempOrder(1 To colCount - 1)
    Dim k As Long: k = 0
    For i = 1 To colCount
        If i <> srcIdx Then
            k = k + 1
            tempOrder(k) = i
        End If
    Next i

    Dim relPos As Long
    relPos = relIdx
    If srcIdx < relIdx Then relPos = relPos - 1

    If posOffset < 0 Then
        newIdx = relPos + posOffset + 1
    Else
        newIdx = relPos + posOffset
    End If

    If newIdx < 1 Then newIdx = 1
    If newIdx > colCount - 1 Then newIdx = colCount - 1

    If srcIdx < relIdx Then
        If newIdx = srcIdx Then Exit Sub
    ElseIf srcIdx > relIdx Then
        If newIdx = srcIdx - 1 Then Exit Sub
    Else
        If newIdx = srcIdx Then Exit Sub
    End If

    Dim finalOrder() As Long
    ReDim finalOrder(1 To colCount)
    Dim outPos As Long: outPos = 0

    For i = 1 To newIdx - 1
        outPos = outPos + 1
        finalOrder(outPos) = tempOrder(i)
    Next i

    outPos = outPos + 1
    finalOrder(outPos) = srcIdx

    For i = newIdx To colCount - 1
        outPos = outPos + 1
        finalOrder(outPos) = tempOrder(i)
    Next i

    ' Rebuild headers, dtypes, and data using the order map.
    p_EnsureTypes
    Dim newHdrs() As String: ReDim newHdrs(1 To colCount)
    Dim newTypes() As ctDataType: ReDim newTypes(1 To colCount)
    Dim newData() As Variant
    If m_RowCount > 0 Then ReDim newData(1 To m_RowCount, 1 To colCount)

    For i = 1 To colCount
        newHdrs(i) = m_Headers(finalOrder(i))
        newTypes(i) = m_ColTypes(finalOrder(i))
    Next i

    Dim r As Long, c As Long
    If m_RowCount > 0 Then
        For r = 1 To m_RowCount
            For c = 1 To colCount
                newData(r, c) = m_Data(r, finalOrder(c))
            Next c
        Next r
    End If

    m_Headers = newHdrs
    m_ColTypes = newTypes
    If m_RowCount > 0 Then m_Data = newData

    ' finalOrder(newPos) = old column position - exactly the mapping the
    ' lineage engine wants.
    If Not m_Lin Is Nothing Then
        m_Lin.x_ApplyColumnMapping finalOrder, colCount, ctSrcUnknown
    End If

    p_ClearIndexes
End Sub

Public Sub ReorderColumnInPlace(ByVal colName As Variant, posOffset As Integer)
    colName = p_LabelColumnRef(colName)
    Dim srcIdx As Long, tgtIdx As Long
    Dim colCount As Long: colCount = m_ColCount
    Dim i As Long

    srcIdx = p_ColIdx(colName)
    If srcIdx = 0 Then
        Err.Raise vbObjectError + 711, "cTable.ReorderColumnInPlace", _
                  "Column not found: """ & colName & """"
    End If

    If posOffset < 1 Or posOffset > colCount Then
        tgtIdx = colCount
    Else
        tgtIdx = posOffset
    End If

    If srcIdx = tgtIdx Then Exit Sub

    Dim tempOrder() As Long
    ReDim tempOrder(1 To colCount - 1)
    Dim k As Long: k = 0
    For i = 1 To colCount
        If i <> srcIdx Then
            k = k + 1
            tempOrder(k) = i
        End If
    Next i

    Dim finalOrder() As Long
    ReDim finalOrder(1 To colCount)
    Dim outPos As Long: outPos = 0

    If tgtIdx = 1 Then
        outPos = outPos + 1
        finalOrder(outPos) = srcIdx
        For i = 1 To colCount - 1
            outPos = outPos + 1
            finalOrder(outPos) = tempOrder(i)
        Next i
    ElseIf tgtIdx >= colCount Then
        For i = 1 To colCount - 1
            outPos = outPos + 1
            finalOrder(outPos) = tempOrder(i)
        Next i
        outPos = outPos + 1
        finalOrder(outPos) = srcIdx
    Else
        For i = 1 To tgtIdx - 1
            outPos = outPos + 1
            finalOrder(outPos) = tempOrder(i)
        Next i
        outPos = outPos + 1
        finalOrder(outPos) = srcIdx
        For i = tgtIdx To colCount - 1
            outPos = outPos + 1
            finalOrder(outPos) = tempOrder(i)
        Next i
    End If

    p_EnsureTypes
    Dim newHdrs() As String: ReDim newHdrs(1 To colCount)
    Dim newTypes() As ctDataType: ReDim newTypes(1 To colCount)
    Dim newData() As Variant
    If m_RowCount > 0 Then ReDim newData(1 To m_RowCount, 1 To colCount)

    For i = 1 To colCount
        newHdrs(i) = m_Headers(finalOrder(i))
        newTypes(i) = m_ColTypes(finalOrder(i))
    Next i

    Dim r As Long, c As Long
    If m_RowCount > 0 Then
        For r = 1 To m_RowCount
            For c = 1 To colCount
                newData(r, c) = m_Data(r, finalOrder(c))
            Next c
        Next r
    End If

    m_Headers = newHdrs
    m_ColTypes = newTypes
    If m_RowCount > 0 Then m_Data = newData

    ' finalOrder(newPos) = old column position - exactly the mapping the
    ' lineage engine wants.
    If Not m_Lin Is Nothing Then
        m_Lin.x_ApplyColumnMapping finalOrder, colCount, ctSrcUnknown
    End If

    p_ClearIndexes
End Sub

'==============================================================================
'
' COLUMN VALIDATION / UTILITY
'
'==============================================================================
' Returns True if a column exists, using the class-wide case-insensitive
' column-name comparison.
Public Function ColumnExists(ByVal colName As Variant) As Boolean
    ColumnExists = (p_ResolveCol(colName, False) > 0)
End Function

Public Function GetColumnArrayByName(colNames As Variant) As Variant
    Dim n As Long
    n = UBound(colNames) - LBound(colNames) + 1
    Dim idxs() As Long
    ReDim idxs(1 To n)

    Dim i As Long
    For i = 1 To n
        Dim cn As String
        cn = CStr(colNames(LBound(colNames) + i - 1))
        Dim ci As Long
        ci = p_ColIdx(cn)
        If ci = 0 Then
            Err.Raise vbObjectError + 1000, "cTable.GetColumnArrayByName", _
                      "Column not found: """ & cn & """"
        End If
        idxs(i) = ci
    Next i
    GetColumnArrayByName = idxs
End Function

Public Function GetHeaderIndex(ByVal hdrName As String) As Long
    GetHeaderIndex = p_ColIdx(hdrName)
End Function

Public Function MissingColumns(expectedColumns As Variant, Optional ShowMatches As Boolean = False) As Variant
    Dim found As Object, missing As Object
    Set found = CreateObject("Scripting.Dictionary")
    Set missing = CreateObject("Scripting.Dictionary")
    found.CompareMode = vbTextCompare
    missing.CompareMode = vbTextCompare

    Dim actualHeaders As Variant
    actualHeaders = Me.headers

    Dim actualDict As Object
    Set actualDict = CreateObject("Scripting.Dictionary")
    actualDict.CompareMode = vbTextCompare

    Dim i As Long
    For i = LBound(actualHeaders) To UBound(actualHeaders)
        actualDict(actualHeaders(i)) = True
    Next i

    For i = LBound(expectedColumns) To UBound(expectedColumns)
        If actualDict.Exists(expectedColumns(i)) Then
            found(expectedColumns(i)) = True
        Else
            missing(expectedColumns(i)) = True
        End If
    Next i

    If ShowMatches Then
        If found.count = 0 Then MissingColumns = Array() Else MissingColumns = found.keys
    Else
        If missing.count = 0 Then MissingColumns = Array() Else MissingColumns = missing.keys
    End If
End Function

Public Function UnexpectedColumns(expectedColumns As Variant, Optional ShowMatches As Boolean = False) As Variant
    Dim found As Object, unexpected As Object
    Set found = CreateObject("Scripting.Dictionary")
    Set unexpected = CreateObject("Scripting.Dictionary")
    found.CompareMode = vbTextCompare
    unexpected.CompareMode = vbTextCompare

    Dim expectedDict As Object
    Set expectedDict = CreateObject("Scripting.Dictionary")
    expectedDict.CompareMode = vbTextCompare

    Dim i As Long
    For i = LBound(expectedColumns) To UBound(expectedColumns)
        expectedDict(expectedColumns(i)) = True
    Next i

    Dim actualHeaders As Variant
    actualHeaders = Me.headers

    For i = LBound(actualHeaders) To UBound(actualHeaders)
        If expectedDict.Exists(actualHeaders(i)) Then
            found(actualHeaders(i)) = True
        Else
            unexpected(actualHeaders(i)) = True
        End If
    Next i

    If ShowMatches Then
        If found.count = 0 Then UnexpectedColumns = Array() Else UnexpectedColumns = found.keys
    Else
        If unexpected.count = 0 Then UnexpectedColumns = Array() Else UnexpectedColumns = unexpected.keys
    End If
End Function

'==============================================================================
'
' DATA TYPES - NA VALUES  (pandas NaN / NaT / pd.isna)   [new in 20260723b]
'
'==============================================================================

' Universal missing value (pandas NaN / None analogue).
Public Property Get NA() As Variant
    NA = Null
End Property

' Missing datetime (pandas NaT). Identical to NA; readable intent for dates.
Public Property Get NaT() As Variant
    NaT = Null
End Property

' Scalar-level NA test (pandas pd.isna(value)).
Public Function IsNAValue(ByVal v As Variant, _
                          Optional ByVal TreatEmptyStringAsNA As Boolean = True) As Boolean
    IsNAValue = p_IsNA(v, TreatEmptyStringAsNA)
End Function

'==============================================================================
'
' DATA TYPES - DTYPE PROPERTIES  [new in 20260723b]
'
'==============================================================================

' Single column dtype by NAME (pandas s.dtype).
Public Property Get dtype(ByVal colName As Variant) As ctDataType
    colName = p_LabelColumnRef(colName)
    Dim ci As Long: ci = p_ColIdx(colName)
    If ci = 0 Then
        Err.Raise vbObjectError + 901, "cTable.dtype", _
                  "Column not found: """ & colName & """"
    End If
    p_EnsureTypes
    dtype = m_ColTypes(ci)
End Property

Public Property Let dtype(ByVal colName As Variant, ByVal t As ctDataType)
    colName = p_LabelColumnRef(colName)
    Dim ci As Long: ci = p_ColIdx(colName)
    If ci = 0 Then
        Err.Raise vbObjectError + 901, "cTable.dtype", _
                  "Column not found: """ & colName & """"
    End If
    p_EnsureTypes
    m_ColTypes(ci) = t
    p_ClearIndexes
End Property

' Single column dtype by 1-based column NUMBER.
Public Property Get dtype_at(ByVal colNumber As Long) As ctDataType
    If colNumber < 1 Or colNumber > m_ColCount Then
        Err.Raise vbObjectError + 902, "cTable.dtype_at", _
                  "Column number out of bounds: " & colNumber
    End If
    p_EnsureTypes
    dtype_at = m_ColTypes(colNumber)
End Property

Public Property Let dtype_at(ByVal colNumber As Long, ByVal t As ctDataType)
    If colNumber < 1 Or colNumber > m_ColCount Then
        Err.Raise vbObjectError + 902, "cTable.dtype_at", _
                  "Column number out of bounds: " & colNumber
    End If
    p_EnsureTypes
    m_ColTypes(colNumber) = t
    p_ClearIndexes
End Property

' Returns the dtype name string for a column by name
Public Property Get dtype_string(ByVal colName As Variant) As String
    colName = p_LabelColumnRef(colName)
    Dim ci As Long: ci = p_ColIdx(colName)
    If ci = 0 Then
        Err.Raise vbObjectError + 950, "cTable.dtype_string", _
                  "Column not found: """ & colName & """"
    End If
    p_EnsureTypes
    dtype_string = Me.dtype_name(m_ColTypes(ci))
End Property

' All dtypes (pandas df.dtypes): Scripting.Dictionary name -> ctDataType.
Public Property Get dtypes() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    p_EnsureTypes
    Dim i As Long
    For i = 1 To m_ColCount
        d(m_Headers(i)) = m_ColTypes(i)
    Next i
    Set dtypes = d
End Property

' Human-readable dtype listing for Debug.Print (pandas print(df.dtypes)).
Public Property Get dtypes_string() As String
    p_EnsureTypes
    Dim sb As String, i As Long
    For i = 1 To m_ColCount
        sb = sb & m_Headers(i) & vbTab & Me.dtype_name(m_ColTypes(i)) & vbCrLf
    Next i
    dtypes_string = sb
End Property

' Raw dtypes as a 1-based Variant array of Longs (for cross-instance copy).
Public Property Get dtypes_array() As Variant
    p_EnsureTypes
    If m_ColCount = 0 Then
        dtypes_array = Array()
        Exit Property
    End If
    Dim v() As Variant
    ReDim v(1 To m_ColCount)
    Dim i As Long
    For i = 1 To m_ColCount
        v(i) = CLng(m_ColTypes(i))
    Next i
    dtypes_array = v
End Property

Public Property Let dtypes_array(ByVal v As Variant)
    p_EnsureTypes
    If m_ColCount = 0 Then Exit Property
    If Not IsArray(v) Then Exit Property
    Dim i As Long, lo As Long
    lo = LBound(v)
    For i = 1 To m_ColCount
        If lo + i - 1 <= UBound(v) Then
            m_ColTypes(i) = CLng(v(lo + i - 1))
        End If
    Next i
End Property

' Copy dtypes from another cTable, matching columns BY NAME.
Public Sub set_dtypes_from(ByVal src As cTable)
    If src Is Nothing Then Exit Sub
    p_EnsureTypes
    Dim i As Long
    For i = 1 To m_ColCount
        Dim srcCol As Long
        srcCol = src.columnNumber(m_Headers(i))
        If srcCol > 0 Then m_ColTypes(i) = src.dtype_at(srcCol)
    Next i
End Sub

' Display name for a ctDataType value.
Public Function dtype_name(ByVal t As ctDataType) As String
    Select Case t
        Case ctVariant:  dtype_name = "variant"
        Case ctString:   dtype_name = "string"
        Case ctLong:     dtype_name = "long"
        Case ctDouble:   dtype_name = "double"
        Case ctCurrency: dtype_name = "currency"
        Case ctDate:     dtype_name = "date"
        Case ctBoolean:  dtype_name = "boolean"
        Case Else:       dtype_name = "variant"
    End Select
End Function

' Parse a type name to the enum; accepts pandas-flavoured aliases.
Public Function dtype_from_name(ByVal typeName As String) As ctDataType
    Select Case LCase$(Trim$(typeName))
        Case "variant", "object", "":            dtype_from_name = ctVariant
        Case "string", "str", "text":            dtype_from_name = ctString
        Case "long", "int", "int64", "integer":  dtype_from_name = ctLong
        Case "double", "float", "float64":       dtype_from_name = ctDouble
        Case "currency", "money":                dtype_from_name = ctCurrency
        Case "date", "datetime", "datetime64":   dtype_from_name = ctDate
        Case "boolean", "bool":                  dtype_from_name = ctBoolean
        Case Else
            Err.Raise vbObjectError + 900, "cTable.dtype_from_name", _
                      "Unknown data type name: """ & typeName & """"
    End Select
End Function

'==============================================================================
'
' DATA TYPES - INFERENCE + CONVERSION  [new in 20260723b]
'
'==============================================================================

' Infer the best dtype for every column, or one column if colName given
' (pandas df.convert_dtypes()). Called automatically by read methods.
Public Sub convert_dtypes(Optional ByVal colName As Variant = "", _
                          Optional ByVal coerce As Boolean = True)
    If p_HasOptionalCol(colName) Then colName = p_LabelColumnRef(colName)
    p_EnsureTypes
    Dim c As Long, cStart As Long, cEnd As Long
    If Len(colName) > 0 Then
        cStart = p_ColIdx(colName)
        If cStart = 0 Then
            Err.Raise vbObjectError + 903, "cTable.convert_dtypes", _
                      "Column not found: """ & colName & """"
        End If
        cEnd = cStart
    Else
        cStart = 1: cEnd = m_ColCount
    End If

    For c = cStart To cEnd
        Dim t As ctDataType
        t = p_InferColumnType(c)
        m_ColTypes(c) = t
        If coerce And t <> ctVariant Then
            p_CoerceColumn c, t, ctErrorsPreserve
        End If
    Next c
    p_ClearIndexes
End Sub

' Convert a column to a new dtype IN PLACE (pandas astype).
' Returns the number of failed conversions.
Public Function astype(ByVal colName As Variant, _
                       ByVal newType As ctDataType, _
                       Optional ByVal errors As ctConvertErrors = ctErrorsPreserve) As Long
    colName = p_LabelColumnRef(colName)
    Dim ci As Long: ci = p_ColIdx(colName)
    If ci = 0 Then
        Err.Raise vbObjectError + 904, "cTable.astype", _
                  "Column not found: """ & colName & """"
    End If
    p_EnsureTypes
    astype = p_CoerceColumn(ci, newType, errors)
    m_ColTypes(ci) = newType
    p_ClearIndexes
End Function

' Non-mutating astype: returns a converted copy, source untouched.
Public Function with_astype(ByVal colName As Variant, _
                            ByVal newType As ctDataType, _
                            Optional ByVal errors As ctConvertErrors = ctErrorsPreserve) As cTable
    colName = p_LabelColumnRef(colName)
    Dim t As cTable
    Set t = Me.Clone
    t.astype colName, newType, errors
    Set with_astype = t
End Function

' pandas pd.to_numeric(s, errors='coerce')
Public Function to_numeric(ByVal colName As Variant, _
                           Optional ByVal errors As ctConvertErrors = ctErrorsCoerce) As Long
    colName = p_LabelColumnRef(colName)
    to_numeric = Me.astype(colName, ctDouble, errors)
End Function

' pandas pd.to_datetime(s, errors='coerce')
Public Function to_datetime(ByVal colName As Variant, _
                            Optional ByVal errors As ctConvertErrors = ctErrorsCoerce) As Long
    colName = p_LabelColumnRef(colName)
    to_datetime = Me.astype(colName, ctDate, errors)
End Function

'==============================================================================
'
' DATA TYPES - NA OPERATIONS  [new in 20260723b]
'
'==============================================================================

' NA mask (pandas df.isna() / s.isna()).
'   t.isna("Col") -> 1-based 1-D Boolean array (1..RowCount)
'   t.isna        -> 1-based 2-D Boolean array (1..RowCount, 1..ColCount)
Public Property Get isna(Optional ByVal colName As Variant = "") As Variant
    If p_HasOptionalCol(colName) Then colName = p_LabelColumnRef(colName)
    Dim r As Long, c As Long
    If Len(colName) > 0 Then
        Dim ci As Long: ci = p_ColIdx(colName)
        If ci = 0 Then
            Err.Raise vbObjectError + 905, "cTable.isna", _
                      "Column not found: """ & colName & """"
        End If
        If m_RowCount = 0 Then isna = Array(): Exit Property
        Dim mask1() As Boolean
        ReDim mask1(1 To m_RowCount)
        For r = 1 To m_RowCount
            mask1(r) = p_IsNA(m_Data(r, ci))
        Next r
        isna = mask1
    Else
        If m_RowCount = 0 Or m_ColCount = 0 Then isna = Array(): Exit Property
        Dim mask2() As Boolean
        ReDim mask2(1 To m_RowCount, 1 To m_ColCount)
        For r = 1 To m_RowCount
            For c = 1 To m_ColCount
                mask2(r, c) = p_IsNA(m_Data(r, c))
            Next c
        Next r
        isna = mask2
    End If
End Property

' Inverse mask (pandas df.notna() / s.notna()).
Public Property Get notna(Optional ByVal colName As Variant = "") As Variant
    If p_HasOptionalCol(colName) Then colName = p_LabelColumnRef(colName)
    Dim r As Long, c As Long
    If Len(colName) > 0 Then
        Dim ci As Long: ci = p_ColIdx(colName)
        If ci = 0 Then
            Err.Raise vbObjectError + 905, "cTable.notna", _
                      "Column not found: """ & colName & """"
        End If
        If m_RowCount = 0 Then notna = Array(): Exit Property
        Dim mask1() As Boolean
        ReDim mask1(1 To m_RowCount)
        For r = 1 To m_RowCount
            mask1(r) = Not p_IsNA(m_Data(r, ci))
        Next r
        notna = mask1
    Else
        If m_RowCount = 0 Or m_ColCount = 0 Then notna = Array(): Exit Property
        Dim mask2() As Boolean
        ReDim mask2(1 To m_RowCount, 1 To m_ColCount)
        For r = 1 To m_RowCount
            For c = 1 To m_ColCount
                mask2(r, c) = Not p_IsNA(m_Data(r, c))
            Next c
        Next r
        notna = mask2
    End If
End Property

' Non-NA counts.
'
'   t.count("Amount")              -> Long
'   Set d = t.count                -> Dictionary for all columns
'   Set d = t.count(numeric_only:=True)
'
' count always excludes NA values; it therefore has no skipna parameter.
Public Function count(Optional ByVal colName As Variant, _
                      Optional ByVal numeric_only As Boolean = False) As Variant
    If p_HasOptionalCol(colName) Then colName = p_LabelColumnRef(colName)
    If IsMissing(colName) Then
        Set count = p_AggregateAllColumns( _
                        "count", True, numeric_only, _
                        0.5, "linear", 1)

    ElseIf p_IsBlankColumnArgument(colName) Then
        Set count = p_AggregateAllColumns( _
                        "count", True, numeric_only, _
                        0.5, "linear", 1)

    Else
        count = p_AggregateNamedColumn( _
                    "count", CStr(colName), True, numeric_only, _
                    0.5, "linear", 1)
    End If
End Function


' True if any NA exists in the column / whole table (pandas s.hasnans).
Public Property Get hasnans(Optional ByVal colName As Variant = "") As Boolean
    If p_HasOptionalCol(colName) Then colName = p_LabelColumnRef(colName)
    Dim c As Long, r As Long
    If Len(colName) > 0 Then
        Dim ci As Long: ci = p_ColIdx(colName)
        If ci = 0 Then
            Err.Raise vbObjectError + 906, "cTable.hasnans", _
                      "Column not found: """ & colName & """"
        End If
        hasnans = (p_ColNACount(ci) > 0)
    Else
        For c = 1 To m_ColCount
            For r = 1 To m_RowCount
                If p_IsNA(m_Data(r, c)) Then hasnans = True: Exit Property
            Next r
        Next c
        hasnans = False
    End If
End Property

' Return a NEW cTable with NA rows removed; source unchanged (pandas dropna).
Public Function dropna(Optional ByVal subsetCols As Variant, _
                       Optional ByVal how As String = "any") As cTable
    Dim useCols() As Long, nCols As Long, k As Long
    If IsMissing(subsetCols) Or IsEmpty(subsetCols) Then
        nCols = m_ColCount
        ReDim useCols(1 To IIf(nCols > 0, nCols, 1))
        For k = 1 To nCols: useCols(k) = k: Next k
    Else
        nCols = UBound(subsetCols) - LBound(subsetCols) + 1
        ReDim useCols(1 To nCols)
        For k = 1 To nCols
            Dim cn As String: cn = CStr(subsetCols(LBound(subsetCols) + k - 1))
            useCols(k) = p_ColIdx(cn)
            If useCols(k) = 0 Then
                Err.Raise vbObjectError + 907, "cTable.dropna", _
                          "Column not found: """ & cn & """"
            End If
        Next k
    End If

    Dim dropIfAny As Boolean
    dropIfAny = (LCase$(Trim$(how)) <> "all")

    Dim matchRows() As Long
    ReDim matchRows(1 To IIf(m_RowCount > 0, m_RowCount, 1))
    Dim matchCount As Long, r As Long

    For r = 1 To m_RowCount
        Dim naCount As Long: naCount = 0
        For k = 1 To nCols
            If p_IsNA(m_Data(r, useCols(k))) Then naCount = naCount + 1
        Next k
        Dim keep As Boolean
        If dropIfAny Then
            keep = (naCount = 0)
        Else
            keep = (naCount < nCols)
        End If
        If keep Then
            matchCount = matchCount + 1
            matchRows(matchCount) = r
        End If
    Next r

    Set dropna = p_BuildFromRows(matchRows, matchCount)
    Set dropna.DataSource = Me.DataSource
    dropna.set_dtypes_from Me
End Function

' Replace NA values IN PLACE (pandas fillna).
Public Sub fillna(ByVal fillValue As Variant, _
                  Optional ByVal colName As Variant = "")
    If p_HasOptionalCol(colName) Then colName = p_LabelColumnRef(colName)
    p_EnsureTypes
    Dim c As Long, cStart As Long, cEnd As Long, r As Long
    If Len(colName) > 0 Then
        cStart = p_ColIdx(colName)
        If cStart = 0 Then
            Err.Raise vbObjectError + 908, "cTable.fillna", _
                      "Column not found: """ & colName & """"
        End If
        cEnd = cStart
    Else
        cStart = 1: cEnd = m_ColCount
    End If

    For c = cStart To cEnd
        Dim fv As Variant, ok As Boolean
        fv = p_TryConvert(fillValue, m_ColTypes(c), ok)
        If Not ok Then fv = fillValue
        For r = 1 To m_RowCount
            If p_IsNA(m_Data(r, c)) Then
                m_Data(r, c) = fv
                If Not m_Lin Is Nothing Then m_Lin.SetCell r, c, ctSrcLiteral
            End If
        Next r
    Next c
    p_ClearIndexes
End Sub

' dtypes + non-null counts per column (pandas df.info()).
Public Function info() As String
    p_EnsureTypes
    Dim sb As String
    sb = "cTable  [" & m_RowCount & " rows x " & m_ColCount & " columns]" & vbCrLf
    Dim i As Long
    For i = 1 To m_ColCount
        sb = sb & " " & i & "  " & m_Headers(i) & vbTab & _
             Me.dtype_name(m_ColTypes(i)) & vbTab & _
             "non-null: " & (m_RowCount - p_ColNACount(i)) & vbCrLf
    Next i
    info = sb
End Function

'----------------- NEWLY ADDED - Move to proper section after testing
'------------------------------------------------------------------------------
' RenameColsByArr
'
' Renames columns using either:
'   1. An array of Scripting.Dictionary-compatible objects
'        Each dictionary key   = existing column name
'        Each dictionary value = new column name
'
'   2. A string containing dictionary-style key:value pairs:
'        "'AvailableSpending': 'Available Spending', 'UnliquidatedObligations':'Obligations'"
'
' If renameDictArr is an empty array or an empty string, this method does nothing.
'
' Raises errors similar to RenameColumn when a source column is not found.
'------------------------------------------------------------------------------
Public Sub RenameColsByArr(ByVal renameDictArr As Variant)

    Dim renameArr As Variant

    ' Allow a string shorthand such as:
    '   "'OldName': 'New Name', 'OldName2': 'New Name2'"
    If VarType(renameDictArr) = vbString Then
        renameArr = p_CvtDictStrToDictArray(CStr(renameDictArr))
    Else
        renameArr = renameDictArr
    End If

    If IsEmpty(renameArr) Then Exit Sub

    If Not IsArray(renameArr) Then
        Err.Raise vbObjectError + 81, "cTable.RenameColsByArr", _
                  "Input must be an array of dictionaries or a dictionary-style string."
    End If

    ' Empty array does nothing.
    Dim lb As Long
    Dim ub As Long

    On Error GoTo EmptyArray
    lb = LBound(renameArr)
    ub = UBound(renameArr)
    On Error GoTo 0

    If ub < lb Then Exit Sub

    Dim i As Long
    Dim d As Object
    Dim oldName As Variant
    Dim newName As String
    Dim ci As Long

    For i = lb To ub

        If IsObject(renameArr(i)) Then
            Set d = renameArr(i)
        Else
            Err.Raise vbObjectError + 82, "cTable.RenameColsByArr", _
                      "Array element " & CStr(i) & " is not a dictionary-compatible object."
        End If

        For Each oldName In d.keys

            ci = p_ColIdx(CStr(oldName))

            If ci = 0 Then
                Err.Raise vbObjectError + 80, "cTable.RenameColsByArr", _
                          "Column not found: """ & CStr(oldName) & """"
            End If

            newName = CStr(d(oldName))
            m_Headers(ci) = newName

        Next oldName

    Next i

    p_ClearIndexes
    Exit Sub

EmptyArray:
    Err.Clear
    Exit Sub

End Sub


'------------------------------------------------------------------------------
' p_CvtDictStrToDictArray
'
' Converts a string containing delimited key:value pairs into a Variant holding
' a 0-based array with one Scripting.Dictionary-compatible object.
'
' Default accepted form:
'   "'Key1': 'Value1', 'Key2':'Value2'"
'
' Also accepts optional surrounding braces:
'   "{'Key1': 'Value1', 'Key2':'Value2'}"
'
' Also accepts double-quoted keys and values:
'   """Key1"": ""Value1"", ""Key2"": ""Value2"""
'
' Returns:
'   Empty input  -> Array()
'   Valid input  -> Variant containing a 0-based array of dictionaries
'
' Notes:
'   This helper is intentionally generic so it can be reused by other methods.
'   The simple parser expects comma-separated key:value pairs.
'   Commas and colons inside keys or values are not supported.
'------------------------------------------------------------------------------
Private Function p_CvtDictStrToDictArray(ByVal dictStr As String) As Variant

    Dim s As String
    s = Trim$(dictStr)

    If Len(s) = 0 Then
        p_CvtDictStrToDictArray = Array()
        Exit Function
    End If

    ' Allow optional surrounding braces.
    If Left$(s, 1) = "{" And right$(s, 1) = "}" Then
        s = mid$(s, 2, Len(s) - 2)
        s = Trim$(s)
    End If

    If Len(s) = 0 Then
        p_CvtDictStrToDictArray = Array()
        Exit Function
    End If

    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare

    Dim pairs As Variant
    pairs = Split(s, ",")

    Dim i As Long
    Dim pairText As String
    Dim delimiterPos As Long
    Dim keyText As String
    Dim valueText As String

    For i = LBound(pairs) To UBound(pairs)

        pairText = Trim$(CStr(pairs(i)))

        If Len(pairText) > 0 Then

            delimiterPos = InStr(1, pairText, ":", vbBinaryCompare)

            If delimiterPos = 0 Then
                Err.Raise vbObjectError + 83, "cTable.p_CvtDictStrToDictArray", _
                          "Invalid key:value pair: " & pairText
            End If

            keyText = Trim$(Left$(pairText, delimiterPos - 1))
            valueText = Trim$(mid$(pairText, delimiterPos + 1))

            keyText = p_UnquoteTextToken(keyText)
            valueText = p_UnquoteTextToken(valueText)

            If Len(keyText) = 0 Then
                Err.Raise vbObjectError + 84, "cTable.p_CvtDictStrToDictArray", _
                          "Dictionary key cannot be blank."
            End If

            If Len(valueText) = 0 Then
                Err.Raise vbObjectError + 85, "cTable.p_CvtDictStrToDictArray", _
                          "Dictionary value cannot be blank for key """ & keyText & """."
            End If

            d(keyText) = valueText

        End If

    Next i

    If d.count = 0 Then
        p_CvtDictStrToDictArray = Array()
        Exit Function
    End If

    Dim outArr(0 To 0) As Object
    Set outArr(0) = d

    p_CvtDictStrToDictArray = outArr

End Function


'------------------------------------------------------------------------------
' p_UnquoteTextToken
'
' Removes one matching pair of surrounding single quotes or double quotes from
' a text token. Whitespace outside the quote pair is ignored.
'
' Examples:
'   "'ABC'"     -> "ABC"
'   """ABC"""   -> "ABC"
'   " ABC "     -> "ABC"
'------------------------------------------------------------------------------
Private Function p_UnquoteTextToken(ByVal token As String) As String

    Dim t As String
    t = Trim$(token)

    If Len(t) >= 2 Then

        If Left$(t, 1) = "'" And right$(t, 1) = "'" Then
            t = mid$(t, 2, Len(t) - 2)

        ElseIf Left$(t, 1) = """" And right$(t, 1) = """" Then
            t = mid$(t, 2, Len(t) - 2)

        End If

    End If

    p_UnquoteTextToken = Trim$(t)

End Function

' ---------------------------------------------------------------------------
' Resolve a column-name argument into 1-based column numbers.
'
' Like p_ResolveGroupKeys, but with error text naming the ARGUMENT rather than
' calling everything a group key. Returns 0 for a missing or empty argument.
' ---------------------------------------------------------------------------
Private Function p_ResolveColumnList(ByRef outCols() As Long, _
                                     ByVal callerName As String, _
                                     ByVal argName As String, _
                                     Optional ByVal COLS As Variant) As Long

    If p_ArgIsMissing(COLS) Then Exit Function

    Dim names As Variant

    If IsArray(COLS) Then
        If Not p_ArrayHasElements(COLS) Then Exit Function
        names = COLS
    Else
        If IsObject(COLS) Then
            Err.Raise vbObjectError + 1340, callerName, _
                      argName & " must be a column name or an array of names."
        End If
        If p_IsBlankColumnArgument(COLS) Then Exit Function
        names = Array(COLS)
    End If

    Dim n As Long
    n = UBound(names) - LBound(names) + 1
    ReDim outCols(1 To n)

    Dim i As Long
    Dim j As Long
    Dim nm As String
    Dim ci As Long

    For i = 1 To n
        nm = p_LabelColumnRef(names(LBound(names) + i - 1))
        ci = p_ColIdx(nm)

        If ci = 0 Then
            Err.Raise vbObjectError + 1341, callerName, _
                      argName & " column not found: """ & nm & """"
        End If

        For j = 1 To i - 1
            If outCols(j) = ci Then
                Err.Raise vbObjectError + 1342, callerName, _
                          argName & " lists column """ & nm & """ twice."
            End If
        Next j

        outCols(i) = ci
    Next i

    p_ResolveColumnList = n
End Function


' Mark columns as claimed, refusing any column claimed twice across arguments.
Private Sub p_ClaimColumns(ByRef claimed() As Boolean, _
                           ByRef COLS() As Long, ByVal n As Long, _
                           ByVal callerName As String, ByVal argName As String)
    Dim i As Long
    For i = 1 To n
        If claimed(COLS(i)) Then
            Err.Raise vbObjectError + 1343, callerName, _
                      "Column """ & m_Headers(COLS(i)) & """ cannot be used " & _
                      "as " & argName & " and elsewhere at the same time."
        End If
        claimed(COLS(i)) = True
    Next i
End Sub


' Every column not yet claimed, in original order.
Private Function p_UnclaimedColumns(ByRef claimed() As Boolean, _
                                    ByRef outCols() As Long) As Long
    Dim n As Long
    Dim c As Long

    ReDim outCols(1 To IIf(m_ColCount > 0, m_ColCount, 1))

    For c = 1 To m_ColCount
        If Not claimed(c) Then
            n = n + 1
            outCols(n) = c
        End If
    Next c

    p_UnclaimedColumns = n
End Function


' ---------------------------------------------------------------------------
' Distinct key combinations over keyCols, in first-appearance order.
'
' Returns the distinct count and fills:
'   keyOf(1..RowCount)   ordinal of each row's key
'   repRow(1..count)     the FIRST source row exhibiting that key
'
' Unlike p_BuildGroups this does not exclude NA keys: pandas keeps a missing
' index or column label as its own bucket in a reshape.
' ---------------------------------------------------------------------------
Private Function p_DistinctKeys(ByRef keyCols() As Long, _
                                ByVal keyCount As Long, _
                                ByRef keyOf() As Long, _
                                ByRef repRow() As Long) As Long

    ReDim keyOf(1 To IIf(m_RowCount > 0, m_RowCount, 1))
    ReDim repRow(1 To IIf(m_RowCount > 0, m_RowCount, 1))

    If m_RowCount = 0 Or keyCount = 0 Then Exit Function

    Dim ids As Object
    Set ids = CreateObject("Scripting.Dictionary")
    ids.CompareMode = vbTextCompare

    Dim n As Long
    Dim r As Long
    Dim key As String

    For r = 1 To m_RowCount
        key = p_GroupRowKey(r, keyCols, keyCount)

        If ids.Exists(key) Then
            keyOf(r) = ids(key)
        Else
            n = n + 1
            ids(key) = n
            keyOf(r) = n
            repRow(n) = r
        End If
    Next r

    p_DistinctKeys = n
End Function


' Order two rows by the key columns. NA sorts last, as in sort_values with
' the default na_position.
Private Function p_CompareKeyRows(ByVal rowA As Long, ByVal rowB As Long, _
                                  ByRef keyCols() As Long, _
                                  ByVal keyCount As Long) As Long
    Dim i As Long
    Dim ci As Long
    Dim aNA As Boolean
    Dim bNA As Boolean
    Dim c As Long

    For i = 1 To keyCount
        ci = keyCols(i)
        aNA = p_IsNA(m_Data(rowA, ci))
        bNA = p_IsNA(m_Data(rowB, ci))

        If aNA And bNA Then
            ' equal on this column; fall through to the next
        ElseIf aNA Then
            p_CompareKeyRows = 1
            Exit Function
        ElseIf bNA Then
            p_CompareKeyRows = -1
            Exit Function
        Else
            c = p_CompareSortScalars(m_Data(rowA, ci), _
                                     m_Data(rowB, ci), _
                                     m_ColTypes(ci))
            If c <> 0 Then
                p_CompareKeyRows = c
                Exit Function
            End If
        End If
    Next i
End Function


' Sort key ordinals by their representative rows. Shell sort, because the
' distinct-key count can reach the thousands and an insertion sort would not.
Private Sub p_SortKeyOrder(ByRef order() As Long, ByVal n As Long, _
                           ByRef repRow() As Long, _
                           ByRef keyCols() As Long, _
                           ByVal keyCount As Long)
    If n < 2 Then Exit Sub

    Dim gap As Long
    Dim i As Long
    Dim j As Long
    Dim tmp As Long

    gap = 1
    Do While gap < n \ 3
        gap = gap * 3 + 1
    Loop

    Do While gap >= 1
        For i = gap + 1 To n
            tmp = order(i)
            j = i

            Do While j > gap
                If p_CompareKeyRows(repRow(order(j - gap)), repRow(tmp), _
                                    keyCols, keyCount) <= 0 Then Exit Do
                order(j) = order(j - gap)
                j = j - gap
            Loop

            order(j) = tmp
        Next i

        gap = gap \ 3
    Loop
End Sub


' ---------------------------------------------------------------------------
' Render one key value for use inside a generated column header.
'
' Dates get an ISO form rather than CStr, whose output depends on the machine's
' regional settings - a pivot whose headers change when the workbook moves
' between machines is a genuine hazard for anything downstream that matches on
' those names.
' ---------------------------------------------------------------------------
Private Function p_FormatKeyPart(ByVal v As Variant, _
                                 ByVal colType As ctDataType, _
                                 ByVal fmt As String) As String
    If p_IsNA(v) Then
        p_FormatKeyPart = "NA"
        Exit Function
    End If

    On Error GoTo Fallback

    If Len(fmt) > 0 Then
        p_FormatKeyPart = Format$(v, fmt)
    ElseIf colType = ctDate Or VarType(v) = vbDate Then
        p_FormatKeyPart = Format$(v, "yyyy-mm-dd")
    Else
        p_FormatKeyPart = CStr(v)
    End If

    Exit Function

Fallback:
    Err.Clear
    p_FormatKeyPart = CStr(v)
End Function


Private Function p_KeyHeader(ByVal rowNum As Long, _
                             ByRef keyCols() As Long, _
                             ByVal keyCount As Long, _
                             ByVal sep As String, _
                             ByVal fmt As String) As String
    Dim s As String
    Dim i As Long

    For i = 1 To keyCount
        If i > 1 Then s = s & sep
        s = s & p_FormatKeyPart(m_Data(rowNum, keyCols(i)), _
                                m_ColTypes(keyCols(i)), fmt)
    Next i

    p_KeyHeader = s
End Function


' A cell's source id, never 0. x_SetColumnOverlay reads 0 as "no overlay",
' which would silently fall back to the row map.
Private Function p_LinIdSafe(ByVal r As Long, ByVal c As Long) As Long
    If m_Lin Is Nothing Then
        p_LinIdSafe = ctSrcUnknown
        Exit Function
    End If

    p_LinIdSafe = m_Lin.IdAt(r, c)
    If p_LinIdSafe = 0 Then p_LinIdSafe = ctSrcUnknown
End Function


' ---------------------------------------------------------------------------
' Write one column's per-cell source ids into a lineage engine.
'
' Collapses to a single stamp when every cell agrees, which is the common case
' for a table loaded from one file. Without this a wide pivot would allocate a
' cLongVec of RowCount Longs per column - the same size again as the data.
' ---------------------------------------------------------------------------
Private Sub p_LinSetColumnIds(ByVal lg As cLineage, ByVal c As Long, _
                              ByRef vals() As Long, ByVal n As Long)
    If n < 1 Then Exit Sub

    Dim uniform As Boolean
    Dim i As Long

    uniform = True
    For i = 2 To n
        If vals(i) <> vals(1) Then
            uniform = False
            Exit For
        End If
    Next i

    If uniform Then
        lg.StampColumn c, vals(1)
    Else
        lg.x_SetColumnOverlay c, vals, n
    End If
End Sub


'==============================================================================
' PIVOT
'==============================================================================

' ---------------------------------------------------------------------------
' pivot
'
' Long to wide. pandas DataFrame.pivot.
'
'   df = df.pivot(index="AccountID", columns="RunDate", values="Expenditure")
'   Set w = t.pivot("RunDate", "AccountID", "Expenditure")
'
' columns      column(s) whose distinct values become new column headers
' index        column(s) identifying each output row. Omitted means "every
'              column not named in columns or values" - pandas' index=None
'              means "use the existing index", which cTable does not have.
' values       column(s) supplying the grid values. Omitted means "every
'              column not named in index or columns".
' sort         True (default) orders both output rows and generated columns by
'              their key values, as pandas does. False keeps first-appearance
'              order, which is often what a report wants.
' col_sep      separator joining parts of a generated header
' col_format   VBA Format$ pattern applied to key values in headers. Empty
'              (default) renders dates as yyyy-mm-dd and everything else with
'              CStr.
' max_columns  guard against an accidental pivot on a near-unique column.
'
' RAISES on a duplicate index/columns pair, exactly as pandas does, naming the
' pair rather than silently keeping one of them.
'
' Generated headers must not collide with each other or with an index column
' name; a collision raises rather than being renamed to "Column7".
' ---------------------------------------------------------------------------
Public Function pivot(ByVal columns As Variant, _
                      Optional ByVal index As Variant, _
                      Optional ByVal values As Variant, _
                      Optional ByVal sort As Boolean = True, _
                      Optional ByVal col_sep As String = "_", _
                      Optional ByVal col_format As String = "", _
                      Optional ByVal max_columns As Long = 4096) As cTable

    Const CALLER As String = "cTable.pivot"

    If m_ColCount = 0 Then
        Set pivot = New cTable
        Exit Function
    End If

    p_EnsureTypes

    ' --- resolve the three column roles ---------------------------------
    Dim colCols() As Long
    Dim colCount As Long
    colCount = p_ResolveColumnList(colCols, CALLER, "columns", columns)

    If colCount = 0 Then
        Err.Raise vbObjectError + 1344, CALLER, _
                  "pivot requires at least one columns column."
    End If

    Dim claimed() As Boolean
    ReDim claimed(1 To m_ColCount)
    p_ClaimColumns claimed, colCols, colCount, CALLER, "columns"

    Dim idxCols() As Long
    Dim idxCount As Long
    Dim valCols() As Long
    Dim valCount As Long

    Dim haveIndex As Boolean
    Dim haveValues As Boolean
    haveIndex = Not p_ArgIsMissing(index)
    haveValues = Not p_ArgIsMissing(values)

    If Not haveIndex And Not haveValues Then
        Err.Raise vbObjectError + 1345, CALLER, _
                  "Supply index, values, or both. With neither, there is " & _
                  "nothing to distinguish row identity from grid content."
    End If

    If haveIndex Then
        idxCount = p_ResolveColumnList(idxCols, CALLER, "index", index)
        p_ClaimColumns claimed, idxCols, idxCount, CALLER, "index"
    End If

    If haveValues Then
        valCount = p_ResolveColumnList(valCols, CALLER, "values", values)
        p_ClaimColumns claimed, valCols, valCount, CALLER, "values"
    End If

    If Not haveIndex Then idxCount = p_UnclaimedColumns(claimed, idxCols)
    If Not haveValues Then valCount = p_UnclaimedColumns(claimed, valCols)

    If idxCount = 0 Then
        Err.Raise vbObjectError + 1346, CALLER, _
                  "No columns are left to identify each output row."
    End If

    If valCount = 0 Then
        Err.Raise vbObjectError + 1347, CALLER, _
                  "No columns are left to fill the grid."
    End If

    ' --- distinct keys ---------------------------------------------------
    Dim idxKeyOf() As Long, idxRep() As Long
    Dim colKeyOf() As Long, colRep() As Long
    Dim nrows As Long, nKeys As Long

    nrows = p_DistinctKeys(idxCols, idxCount, idxKeyOf, idxRep)
    nKeys = p_DistinctKeys(colCols, colCount, colKeyOf, colRep)

    If nKeys > max_columns Then
        Err.Raise vbObjectError + 1348, CALLER, _
                  "Pivoting on """ & m_Headers(colCols(1)) & """ would " & _
                  "produce " & nKeys & " column groups, over the " & _
                  max_columns & " limit. A near-unique column is usually " & _
                  "the cause. Raise max_columns if this is intended."
    End If

    Dim outCC As Long
    outCC = idxCount + nKeys * valCount

    ' --- ordering --------------------------------------------------------
    Dim rowOrder() As Long, colOrder() As Long
    ReDim rowOrder(1 To IIf(nrows > 0, nrows, 1))
    ReDim colOrder(1 To IIf(nKeys > 0, nKeys, 1))

    Dim i As Long, j As Long, k As Long, r As Long, c As Long

    For i = 1 To nrows: rowOrder(i) = i: Next i
    For i = 1 To nKeys: colOrder(i) = i: Next i

    If sort Then
        p_SortKeyOrder rowOrder, nrows, idxRep, idxCols, idxCount
        p_SortKeyOrder colOrder, nKeys, colRep, colCols, colCount
    End If

    ' Position of each key ordinal in the output.
    Dim rowPos() As Long, colPos() As Long
    ReDim rowPos(1 To IIf(nrows > 0, nrows, 1))
    ReDim colPos(1 To IIf(nKeys > 0, nKeys, 1))

    For i = 1 To nrows: rowPos(rowOrder(i)) = i: Next i
    For i = 1 To nKeys: colPos(colOrder(i)) = i: Next i

    ' --- source row for each grid cell, and duplicate detection ----------
    Dim cellRow() As Long
    ReDim cellRow(1 To IIf(nrows > 0, nrows, 1), 1 To IIf(nKeys > 0, nKeys, 1))

    Dim rp As Long, cp As Long

    For r = 1 To m_RowCount
        rp = rowPos(idxKeyOf(r))
        cp = colPos(colKeyOf(r))

        If cellRow(rp, cp) <> 0 Then
            Err.Raise vbObjectError + 1349, CALLER, _
                      "Duplicate entry for index " & _
                      p_KeyHeader(r, idxCols, idxCount, col_sep, col_format) & _
                      " and columns " & _
                      p_KeyHeader(r, colCols, colCount, col_sep, col_format) & _
                      " (rows " & cellRow(rp, cp) & " and " & r & "). " & _
                      "pivot cannot choose between them; aggregate first."
        End If

        cellRow(rp, cp) = r
    Next r

    ' --- headers ---------------------------------------------------------
    Dim hdrs() As String
    ReDim hdrs(1 To outCC)

    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare

    For i = 1 To idxCount
        hdrs(i) = m_Headers(idxCols(i))
        seen(hdrs(i)) = 1
    Next i

    Dim pos As Long
    Dim hdr As String
    pos = idxCount

    For j = 1 To valCount
        For k = 1 To nKeys
            pos = pos + 1

            hdr = p_KeyHeader(colRep(colOrder(k)), colCols, colCount, _
                              col_sep, col_format)

            If valCount > 1 Then
                hdr = m_Headers(valCols(j)) & col_sep & hdr
            End If

            If seen.Exists(hdr) Then
                Err.Raise vbObjectError + 1350, CALLER, _
                          "Generated column name """ & hdr & """ collides " & _
                          "with an existing name. Use col_sep or col_format " & _
                          "to separate them."
            End If

            seen(hdr) = 1
            hdrs(pos) = hdr
        Next k
    Next j

    ' --- data ------------------------------------------------------------
    Dim outData() As Variant
    ReDim outData(1 To IIf(nrows > 0, nrows, 1), 1 To outCC)

    Dim srcRow As Long

    For i = 1 To nrows
        srcRow = idxRep(rowOrder(i))
        For j = 1 To idxCount
            outData(i, j) = m_Data(srcRow, idxCols(j))
        Next j
    Next i

    pos = idxCount

    For j = 1 To valCount
        For k = 1 To nKeys
            pos = pos + 1

            For i = 1 To nrows
                srcRow = cellRow(i, k)
                If srcRow = 0 Then
                    outData(i, pos) = Null
                Else
                    outData(i, pos) = m_Data(srcRow, valCols(j))
                End If
            Next i
        Next k
    Next j

    ' --- assemble --------------------------------------------------------
    Dim result As cTable
    Set result = New cTable
    result.from_array hdrs, outData, nrows, outCC
    Set result.DataSource = Me.DataSource

    Dim outTypes() As ctDataType
    ReDim outTypes(1 To outCC)

    For j = 1 To idxCount
        outTypes(j) = m_ColTypes(idxCols(j))
    Next j

    pos = idxCount
    For j = 1 To valCount
        For k = 1 To nKeys
            pos = pos + 1
            outTypes(pos) = m_ColTypes(valCols(j))
        Next k
    Next j

    result.dtypes_array = outTypes

    ' --- lineage: a scatter, so build it fresh ---------------------------
    If Not m_Lin Is Nothing And nrows > 0 Then
        Dim lg As cLineage
        Set lg = New cLineage
        lg.InitUniform nrows, outCC, ctSrcUnknown

        Dim ids() As Long
        ReDim ids(1 To nrows)

        For j = 1 To idxCount
            For i = 1 To nrows
                ids(i) = p_LinIdSafe(idxRep(rowOrder(i)), idxCols(j))
            Next i
            p_LinSetColumnIds lg, j, ids, nrows
        Next j

        pos = idxCount
        For j = 1 To valCount
            For k = 1 To nKeys
                pos = pos + 1

                For i = 1 To nrows
                    srcRow = cellRow(i, k)
                    If srcRow = 0 Then
                        ' A hole in the grid. No file produced this Null.
                        ids(i) = ctSrcMissing
                    Else
                        ids(i) = p_LinIdSafe(srcRow, valCols(j))
                    End If
                Next i

                p_LinSetColumnIds lg, pos, ids, nrows
            Next k
        Next j

        result.x_SetLineage lg
    End If

    Set pivot = result
End Function


'==============================================================================
' MELT
'==============================================================================

' ---------------------------------------------------------------------------
' melt
'
' Wide to long. pandas DataFrame.melt.
'
'   df = df.melt(id_vars=["AccountID"], value_vars=["Q1","Q2"], var_name="Quarter")
'   Set l = t.melt("AccountID", Array("Q1","Q2"), "Quarter", "Amount")
'
' id_vars     column(s) repeated on every output row. Omitted means none.
' value_vars  column(s) to unpivot. Omitted means every column not in id_vars.
' var_name    name of the column holding the source column NAME
' value_name  name of the column holding the value
'
' ROW ORDER matches pandas: all rows for the first value_var, then all rows for
' the second, and so on. To keep each source row's melted values adjacent
' instead, sort the result by the id_vars afterwards.
'
' The value column's dtype is the shared dtype of every value_var when they
' agree, and ctVariant when they do not - stacking a Currency column onto a
' Date column has no single honest type.
' ---------------------------------------------------------------------------
Public Function melt(Optional ByVal id_vars As Variant, _
                     Optional ByVal value_vars As Variant, _
                     Optional ByVal var_name As String = "variable", _
                     Optional ByVal value_name As String = "value") As cTable

    Const CALLER As String = "cTable.melt"

    If m_ColCount = 0 Then
        Set melt = New cTable
        Exit Function
    End If

    p_EnsureTypes

    Dim idCols() As Long
    Dim idCount As Long
    idCount = p_ResolveColumnList(idCols, CALLER, "id_vars", id_vars)

    Dim claimed() As Boolean
    ReDim claimed(1 To m_ColCount)
    p_ClaimColumns claimed, idCols, idCount, CALLER, "id_vars"

    Dim valCols() As Long
    Dim valCount As Long

    If p_ArgIsMissing(value_vars) Then
        valCount = p_UnclaimedColumns(claimed, valCols)
    Else
        valCount = p_ResolveColumnList(valCols, CALLER, "value_vars", value_vars)
        p_ClaimColumns claimed, valCols, valCount, CALLER, "value_vars"
    End If

    If valCount = 0 Then
        Err.Raise vbObjectError + 1351, CALLER, _
                  "There are no columns left to unpivot."
    End If

    Dim vn As String
    Dim nn As String
    vn = Trim$(var_name)
    nn = Trim$(value_name)
    If Len(vn) = 0 Then vn = "variable"
    If Len(nn) = 0 Then nn = "value"

    If StrComp(vn, nn, vbTextCompare) = 0 Then
        Err.Raise vbObjectError + 1352, CALLER, _
                  "var_name and value_name must differ."
    End If

    Dim i As Long
    Dim j As Long
    Dim r As Long

    For i = 1 To idCount
        If StrComp(m_Headers(idCols(i)), vn, vbTextCompare) = 0 _
           Or StrComp(m_Headers(idCols(i)), nn, vbTextCompare) = 0 Then
            Err.Raise vbObjectError + 1353, CALLER, _
                      "var_name or value_name collides with the id_vars " & _
                      "column """ & m_Headers(idCols(i)) & """."
        End If
    Next i

    Dim outCC As Long
    Dim outN As Long
    outCC = idCount + 2
    outN = m_RowCount * valCount

    Dim hdrs() As String
    ReDim hdrs(1 To outCC)
    For i = 1 To idCount
        hdrs(i) = m_Headers(idCols(i))
    Next i
    hdrs(idCount + 1) = vn
    hdrs(idCount + 2) = nn

    ' --- the mapping from output row back to source ----------------------
    Dim RowOf() As Long
    Dim varOf() As Long
    ReDim RowOf(1 To IIf(outN > 0, outN, 1))
    ReDim varOf(1 To IIf(outN > 0, outN, 1))

    Dim p As Long
    For j = 1 To valCount
        For r = 1 To m_RowCount
            p = p + 1
            RowOf(p) = r
            varOf(p) = j
        Next r
    Next j

    ' --- data ------------------------------------------------------------
    Dim outData() As Variant
    ReDim outData(1 To IIf(outN > 0, outN, 1), 1 To outCC)

    For p = 1 To outN
        r = RowOf(p)
        For i = 1 To idCount
            outData(p, i) = m_Data(r, idCols(i))
        Next i
        outData(p, idCount + 1) = m_Headers(valCols(varOf(p)))
        outData(p, idCount + 2) = m_Data(r, valCols(varOf(p)))
    Next p

    Dim result As cTable
    Set result = New cTable
    result.from_array hdrs, outData, outN, outCC
    Set result.DataSource = Me.DataSource

    ' --- dtypes ----------------------------------------------------------
    Dim outTypes() As ctDataType
    ReDim outTypes(1 To outCC)

    For i = 1 To idCount
        outTypes(i) = m_ColTypes(idCols(i))
    Next i

    outTypes(idCount + 1) = ctString

    Dim sharedType As ctDataType
    sharedType = m_ColTypes(valCols(1))
    For j = 2 To valCount
        If m_ColTypes(valCols(j)) <> sharedType Then
            sharedType = ctVariant
            Exit For
        End If
    Next j

    outTypes(idCount + 2) = sharedType
    result.dtypes_array = outTypes

    ' --- lineage: a repeating row gather ---------------------------------
    If Not m_Lin Is Nothing And outN > 0 Then
        Dim lg As cLineage
        Set lg = m_Lin.CloneMe

        ' x_ApplyRowGather reads rows() in the OLD numbering and does not
        ' require n to match the current row count, so repeating each source
        ' row once per value_var is a legitimate gather. Every id_vars
        ' column's provenance survives exactly.
        lg.x_ApplyRowGather RowOf, outN

        Dim oldOf() As Long
        ReDim oldOf(1 To outCC)
        For i = 1 To idCount
            oldOf(i) = idCols(i)
        Next i
        oldOf(idCount + 1) = 0
        oldOf(idCount + 2) = 0

        ' Both new columns start as Literal. That is correct for the variable
        ' column, which holds schema metadata rather than data; the value
        ' column is overwritten with exact per-cell ids below.
        lg.x_ApplyColumnMapping oldOf, outCC, ctSrcLiteral

        Dim ids() As Long
        ReDim ids(1 To outN)

        For p = 1 To outN
            ids(p) = p_LinIdSafe(RowOf(p), valCols(varOf(p)))
        Next p

        p_LinSetColumnIds lg, idCount + 2, ids, outN

        result.x_SetLineage lg
    End If

    Set melt = result
End Function
' CHANGE CLASSIFICATION
'   Public API addition in cTable ......... dt, AddDtColumn, SetDtColumn,
'                                           WithDtColumn, rolling, expanding,
'                                           Add/Set/WithRollingColumn,
'                                           Add/SetExpandingColumn
'   Internal lineage behaviour ............ none. Every new column routes
'                                           through the existing
'                                           p_WriteTransformColumn chokepoint,
'                                           which stamps the whole column with
'                                           a CalculatedSourceId built from a
'                                           stable p_OpLabel string.
'   Registry / source model ............... none
'   Storage / performance ................. none for existing paths
'
' BACKWARD-COMPATIBILITY RISK
'   None identified. No existing signature, default, or dtype rule changes.
'   The only namespace claims are the new member names above; verify none of
'   them collide with a member added since 20260727.
'
' LINEAGE-CORRECTNESS RISK
'   A window value is derived from a RANGE of rows, so exact per-cell
'   provenance is not representable in the current column-stamp model. These
'   columns are therefore stamped calculated with a label naming the operation,
'   source column, group keys, window and ordering column. That degrades
'   precision, which is permitted; it does not invent a source, which is not.
'
' ORDERING CONTRACT  (differs from diff / shift - read this)
'   diff and shift walk rows in their CURRENT order and make ordering the
'   caller's problem. Windows split the difference:
'
'     * OFFSET windows ("7D", "90D", ...) sort each group's rows internally by
'       the on_arg column before walking. Order of the table is irrelevant and
'       a mis-sorted table cannot produce plausible-but-wrong numbers. Results
'       are written back to the rows they belong to.
'     * INTEGER windows (7, 30, ...) and expanding use the CURRENT row order,
'       exactly like cumsum. Sorting is the caller's responsibility there.
'
'==============================================================================


'==============================================================================
'
' DATETIME ACCESSORS   (pandas .dt)
'
'   v = t.dt("RunDATE", "date")            -> 1-based Variant array
'   t.AddDtColumn "RunDATE", "date", "AsOfDate"
'
' Cells that are missing, or that cannot be read as a date, yield Null. That
' is pandas errors="coerce" behaviour and it keeps a single bad cell from
' aborting a whole ETL run.
'
' PARTS (case-insensitive)
'   date | normalize .................. midnight of that day        ctDate
'   time .............................. time of day, date part 0    ctDate
'   year month day hour minute second . component                   ctLong
'   quarter ........................... 1-4                         ctLong
'   dayofweek | weekday ............... Monday=0 .. Sunday=6        ctLong
'   dayofyear ......................... 1-366                       ctLong
'   week | weekofyear | isoweek ....... ISO-8601 week number        ctLong
'   isoyear ........................... ISO-8601 week-numbering yr  ctLong
'   daysinmonth ....................... 28-31                       ctLong
'   month_name day_name ............... locale month / day name     ctString
'   is_month_start is_month_end ....... boundary tests              ctBoolean
'   is_quarter_start is_quarter_end ... boundary tests              ctBoolean
'   is_year_start is_year_end ......... boundary tests              ctBoolean
'   month_start month_end ............. snap to period boundary     ctDate
'   quarter_start quarter_end ......... snap to period boundary     ctDate
'   year_start year_end ............... snap to period boundary     ctDate
'   week_start week_end ............... Monday / Sunday of the week ctDate
'   fiscal_year ....................... see fiscalStartMonth        ctLong
'   fiscal_quarter .................... 1-4 from fiscalStartMonth   ctLong
'
' fiscalStartMonth defaults to 10 (US federal FY). A fiscal year that starts
' in any month after January is labelled by the calendar year it ENDS in, so
' 2026-10-01 falls in FY2027. A fiscalStartMonth of 1 means FY = calendar year.
'==============================================================================

Public Function dt(ByVal colName As Variant, _
                   ByVal part As String, _
                   Optional ByVal fiscalStartMonth As Long = 10) As Variant
    colName = p_LabelColumnRef(colName)

    Const CALLER As String = "cTable.dt"

    Dim ci As Long
    ci = p_ColIdx(colName)

    If ci = 0 Then
        Err.Raise vbObjectError + 1440, CALLER, _
                  "Column not found: """ & colName & """"
    End If

    If fiscalStartMonth < 1 Or fiscalStartMonth > 12 Then
        Err.Raise vbObjectError + 1441, CALLER, _
                  "fiscalStartMonth must be 1 to 12."
    End If

    Dim partCode As Long
    partCode = p_DtPartCode(part, CALLER)

    If m_RowCount = 0 Then
        dt = Array()
        Exit Function
    End If

    p_EnsureTypes

    Dim Out() As Variant
    ReDim Out(1 To m_RowCount)

    Dim r As Long
    Dim v As Variant
    Dim ok As Boolean
    Dim d As Date

    For r = 1 To m_RowCount
        v = p_TryConvert(m_Data(r, ci), ctDate, ok)

        If Not ok Then
            Out(r) = Null
        ElseIf IsNull(v) Then
            Out(r) = Null
        Else
            d = CDate(v)
            Out(r) = p_DtPart(d, partCode, fiscalStartMonth)
        End If
    Next r

    dt = Out
End Function


' Add a derived datetime column. Add raises if the name is taken; Set replaces.
Public Sub AddDtColumn(ByVal colName As Variant, _
                       ByVal part As String, _
                       Optional ByVal newColName As String = "", _
                       Optional ByVal fiscalStartMonth As Long = 10, _
                       Optional ByVal resultType As ctDataType = ctVariant)
    colName = p_LabelColumnRef(colName)

    p_ApplyDt colName, part, newColName, fiscalStartMonth, resultType, _
              False, "cTable.AddDtColumn"
End Sub


Public Sub SetDtColumn(ByVal colName As Variant, _
                       ByVal part As String, _
                       Optional ByVal newColName As String = "", _
                       Optional ByVal fiscalStartMonth As Long = 10, _
                       Optional ByVal resultType As ctDataType = ctVariant)
    colName = p_LabelColumnRef(colName)

    p_ApplyDt colName, part, newColName, fiscalStartMonth, resultType, _
              True, "cTable.SetDtColumn"
End Sub


' Non-mutating form. The source table is unchanged.
Public Function WithDtColumn(ByVal colName As Variant, _
                             ByVal part As String, _
                             Optional ByVal newColName As String = "", _
                             Optional ByVal fiscalStartMonth As Long = 10, _
                             Optional ByVal resultType As ctDataType = ctVariant) As cTable
    colName = p_LabelColumnRef(colName)

    Dim t As cTable
    Set t = Me.Clone

    t.AddDtColumn colName, part, newColName, fiscalStartMonth, resultType

    Set WithDtColumn = t
End Function


'------------------------------------------------------------------------------
' DT INTERNALS
'------------------------------------------------------------------------------

Private Sub p_ApplyDt(ByVal colName As String, _
                      ByVal part As String, _
                      ByVal newColName As String, _
                      ByVal fiscalStartMonth As Long, _
                      ByVal resultType As ctDataType, _
                      ByVal allowReplace As Boolean, _
                      ByVal callerName As String)

    Dim srcCol As Long
    srcCol = p_ColIdx(colName)

    If srcCol = 0 Then
        Err.Raise vbObjectError + 1440, callerName, _
                  "Column not found: """ & colName & """"
    End If

    Dim partCode As Long
    partCode = p_DtPartCode(part, callerName)

    Dim targetName As String
    targetName = p_ResolveTargetName(newColName, colName, _
                                     "_" & p_DtCanonicalName(partCode), _
                                     allowReplace, callerName)

    Dim values As Variant
    values = Me.dt(colName, part, fiscalStartMonth)

    Dim finalType As ctDataType
    If resultType = ctVariant Then
        finalType = p_DtResultType(partCode)
    Else
        finalType = resultType
    End If

    Dim extras As String
    If partCode = 28 Or partCode = 29 Then
        If fiscalStartMonth <> 10 Then
            p_AddExtra extras, "fiscalStartMonth=" & CStr(fiscalStartMonth)
        End If
    End If

    Dim noKeys() As Long

    p_WriteTransformColumn targetName, values, finalType, _
        p_OpLabel("dt." & p_DtCanonicalName(partCode), colName, _
                  noKeys, 0, extras)
End Sub


' part name -> stable internal code. One string comparison per CALL, not per
' row, which matters on tables with hundreds of thousands of rows.
Private Function p_DtPartCode(ByVal part As String, _
                              ByVal callerName As String) As Long

    Select Case LCase$(Trim$(Replace(part, "-", "_")))

        Case "date", "normalize":                    p_DtPartCode = 1
        Case "time":                                 p_DtPartCode = 2
        Case "year":                                 p_DtPartCode = 3
        Case "month":                                p_DtPartCode = 4
        Case "day":                                  p_DtPartCode = 5
        Case "hour":                                 p_DtPartCode = 6
        Case "minute":                               p_DtPartCode = 7
        Case "second":                               p_DtPartCode = 8
        Case "quarter":                              p_DtPartCode = 9
        Case "dayofweek", "day_of_week", "weekday":  p_DtPartCode = 10
        Case "dayofyear", "day_of_year":             p_DtPartCode = 11
        Case "week", "weekofyear", "isoweek", "iso_week": p_DtPartCode = 12
        Case "daysinmonth", "days_in_month":         p_DtPartCode = 13
        Case "month_name", "monthname":              p_DtPartCode = 14
        Case "day_name", "dayname":                  p_DtPartCode = 15
        Case "is_month_start":                       p_DtPartCode = 16
        Case "is_month_end":                         p_DtPartCode = 17
        Case "is_quarter_start":                     p_DtPartCode = 18
        Case "is_quarter_end":                       p_DtPartCode = 19
        Case "is_year_start":                        p_DtPartCode = 20
        Case "is_year_end":                          p_DtPartCode = 21
        Case "month_start":                          p_DtPartCode = 22
        Case "month_end":                            p_DtPartCode = 23
        Case "quarter_start":                        p_DtPartCode = 24
        Case "quarter_end":                          p_DtPartCode = 25
        Case "year_start":                           p_DtPartCode = 26
        Case "year_end":                             p_DtPartCode = 27
        Case "fiscal_year", "fiscalyear", "fy":      p_DtPartCode = 28
        Case "fiscal_quarter", "fiscalquarter", "fq": p_DtPartCode = 29
        Case "isoyear", "iso_year":                  p_DtPartCode = 30
        Case "week_start":                           p_DtPartCode = 31
        Case "week_end":                             p_DtPartCode = 32

        Case Else
            Err.Raise vbObjectError + 1442, callerName, _
                      "Unknown datetime part: """ & part & """. See the " & _
                      "PARTS table above cTable.dt for the supported list."
    End Select
End Function


' The canonical spelling, used for default column names and lineage labels so
' that "weekday" and "dayofweek" intern to the same calculated source.
Private Function p_DtCanonicalName(ByVal partCode As Long) As String
    Select Case partCode
        Case 1:  p_DtCanonicalName = "date"
        Case 2:  p_DtCanonicalName = "time"
        Case 3:  p_DtCanonicalName = "year"
        Case 4:  p_DtCanonicalName = "month"
        Case 5:  p_DtCanonicalName = "day"
        Case 6:  p_DtCanonicalName = "hour"
        Case 7:  p_DtCanonicalName = "minute"
        Case 8:  p_DtCanonicalName = "second"
        Case 9:  p_DtCanonicalName = "quarter"
        Case 10: p_DtCanonicalName = "dayofweek"
        Case 11: p_DtCanonicalName = "dayofyear"
        Case 12: p_DtCanonicalName = "week"
        Case 13: p_DtCanonicalName = "daysinmonth"
        Case 14: p_DtCanonicalName = "month_name"
        Case 15: p_DtCanonicalName = "day_name"
        Case 16: p_DtCanonicalName = "is_month_start"
        Case 17: p_DtCanonicalName = "is_month_end"
        Case 18: p_DtCanonicalName = "is_quarter_start"
        Case 19: p_DtCanonicalName = "is_quarter_end"
        Case 20: p_DtCanonicalName = "is_year_start"
        Case 21: p_DtCanonicalName = "is_year_end"
        Case 22: p_DtCanonicalName = "month_start"
        Case 23: p_DtCanonicalName = "month_end"
        Case 24: p_DtCanonicalName = "quarter_start"
        Case 25: p_DtCanonicalName = "quarter_end"
        Case 26: p_DtCanonicalName = "year_start"
        Case 27: p_DtCanonicalName = "year_end"
        Case 28: p_DtCanonicalName = "fiscal_year"
        Case 29: p_DtCanonicalName = "fiscal_quarter"
        Case 30: p_DtCanonicalName = "isoyear"
        Case 31: p_DtCanonicalName = "week_start"
        Case 32: p_DtCanonicalName = "week_end"
        Case Else: p_DtCanonicalName = "dt"
    End Select
End Function


Private Function p_DtResultType(ByVal partCode As Long) As ctDataType
    Select Case partCode
        Case 1, 2, 22, 23, 24, 25, 26, 27, 31, 32
            p_DtResultType = ctDate

        Case 14, 15
            p_DtResultType = ctString

        Case 16, 17, 18, 19, 20, 21
            p_DtResultType = ctBoolean

        Case Else
            p_DtResultType = ctLong
    End Select
End Function


Private Function p_DtPart(ByVal d As Date, _
                          ByVal partCode As Long, _
                          ByVal fiscalStartMonth As Long) As Variant

    Dim qStartMonth As Long
    Dim offset As Long
    Dim isoYr As Long

    Select Case partCode

        Case 1:  p_DtPart = CDate(Int(CDbl(d)))
        Case 2:  p_DtPart = CDate(CDbl(d) - Int(CDbl(d)))
        Case 3:  p_DtPart = CLng(Year(d))
        Case 4:  p_DtPart = CLng(Month(d))
        Case 5:  p_DtPart = CLng(Day(d))
        Case 6:  p_DtPart = CLng(Hour(d))
        Case 7:  p_DtPart = CLng(Minute(d))
        Case 8:  p_DtPart = CLng(Second(d))
        Case 9:  p_DtPart = CLng((Month(d) - 1) \ 3 + 1)

        ' pandas convention: Monday = 0.
        Case 10: p_DtPart = CLng(Weekday(d, vbMonday) - 1)

        Case 11: p_DtPart = CLng(DatePart("y", d))
        Case 12: p_DtPart = CLng(p_IsoWeek(d, isoYr))
        Case 13: p_DtPart = CLng(Day(DateSerial(Year(d), Month(d) + 1, 0)))
        Case 14: p_DtPart = Format$(d, "mmmm")
        Case 15: p_DtPart = Format$(d, "dddd")

        Case 16: p_DtPart = (Day(d) = 1)
        Case 17: p_DtPart = (Day(d) = Day(DateSerial(Year(d), Month(d) + 1, 0)))

        Case 18
            qStartMonth = ((Month(d) - 1) \ 3) * 3 + 1
            p_DtPart = (Day(d) = 1 And Month(d) = qStartMonth)

        Case 19
            qStartMonth = ((Month(d) - 1) \ 3) * 3 + 1
            p_DtPart = (Int(CDbl(d)) = _
                        CDbl(DateSerial(Year(d), qStartMonth + 3, 0)))

        Case 20: p_DtPart = (Month(d) = 1 And Day(d) = 1)
        Case 21: p_DtPart = (Month(d) = 12 And Day(d) = 31)

        Case 22: p_DtPart = DateSerial(Year(d), Month(d), 1)
        Case 23: p_DtPart = DateSerial(Year(d), Month(d) + 1, 0)

        Case 24
            qStartMonth = ((Month(d) - 1) \ 3) * 3 + 1
            p_DtPart = DateSerial(Year(d), qStartMonth, 1)

        Case 25
            qStartMonth = ((Month(d) - 1) \ 3) * 3 + 1
            p_DtPart = DateSerial(Year(d), qStartMonth + 3, 0)

        Case 26: p_DtPart = DateSerial(Year(d), 1, 1)
        Case 27: p_DtPart = DateSerial(Year(d), 12, 31)

        Case 28
            If fiscalStartMonth <= 1 Then
                p_DtPart = CLng(Year(d))
            ElseIf Month(d) >= fiscalStartMonth Then
                p_DtPart = CLng(Year(d) + 1)
            Else
                p_DtPart = CLng(Year(d))
            End If

        Case 29
            offset = Month(d) - fiscalStartMonth
            If offset < 0 Then offset = offset + 12
            p_DtPart = CLng(offset \ 3 + 1)

        Case 30
            p_IsoWeek d, isoYr
            p_DtPart = CLng(isoYr)

        Case 31
            p_DtPart = DateAdd("d", -(Weekday(d, vbMonday) - 1), _
                               CDate(Int(CDbl(d))))

        Case 32
            p_DtPart = DateAdd("d", 7 - Weekday(d, vbMonday), _
                               CDate(Int(CDbl(d))))

        Case Else
            p_DtPart = Null
    End Select
End Function


' ISO-8601 week number, with the week-numbering year returned by reference.
' The ISO week of a date is the week containing that week's Thursday, which
' makes both values fall out of one DateAdd.
Private Function p_IsoWeek(ByVal d As Date, ByRef isoYear As Long) As Long
    Dim dayOnly As Date
    dayOnly = CDate(Int(CDbl(d)))

    Dim thu As Date
    thu = DateAdd("d", 4 - Weekday(dayOnly, vbMonday), dayOnly)

    isoYear = Year(thu)
    p_IsoWeek = (DatePart("y", thu) - 1) \ 7 + 1
End Function


'==============================================================================
'
' ROLLING AND EXPANDING WINDOWS   (pandas .rolling / .expanding)
'
'   v = t.rolling("Obligations_Net", "90D", "AccountID", "sum", "AsOfDate")
'   t.AddRollingColumn "Obligations_Net", "7D", "AccountID", "sum", "AsOfDate", _
'                      newColName:="Obligations_last7"
'
' window
'   Long        trailing N ROWS, current row included. Uses current row order.
'   String      trailing TIME span. Requires on_arg. Suffixes:
'                 D / day / days        W / week / weeks
'                 H / hour / hours      min / T / minute / minutes
'                 S / sec / second / seconds
'               "M" is rejected as ambiguous - use "30D" or "31D".
'
' by          group key name, or array of names. Windows never cross a group.
' operation   sum mean count min max std var median first last
' on_arg      ordering column for offset windows. Must be date-like.
' min_periods -1 (default) = pandas default: the window size for integer
'             windows, 1 for offset windows and expanding. Fewer non-missing
'             values in the window than this yields Null.
' closed      right (default) | both | left | neither. Offset windows only.
'             right means (t - window, t]: the trailing span EXCLUDING the
'             point exactly one window back.
'
' Missing values inside the window are skipped, not propagated, matching
' pandas. A window whose values are all missing therefore has count 0 and
' returns Null under the default min_periods.
'
' RESULT DTYPE
'   count ................. ctLong
'   sum ................... ctCurrency from ctCurrency, else ctDouble
'   mean std var median ... ctDouble
'   min max first last .... the source column's dtype
'
'   std and var use ddof = 1, as pandas does, so they need at least two
'   non-missing values.
'==============================================================================

Public Function rolling(ByVal colName As Variant, _
                        ByVal window As Variant, _
                        Optional ByVal by As Variant, _
                        Optional ByVal operation As String = "sum", _
                        Optional ByVal on_arg As String = "", _
                        Optional ByVal min_periods As Long = -1, _
                        Optional ByVal closed As String = "right", _
                        Optional ByVal dropna_keys As Boolean = True) As Variant
    colName = p_LabelColumnRef(colName)

    Dim label As String
    Dim finalType As ctDataType

    rolling = p_WindowCore("cTable.rolling", colName, window, by, operation, _
                           on_arg, min_periods, closed, dropna_keys, False, _
                           label, finalType)
End Function


Public Function expanding(ByVal colName As Variant, _
                          Optional ByVal by As Variant, _
                          Optional ByVal operation As String = "sum", _
                          Optional ByVal min_periods As Long = 1, _
                          Optional ByVal dropna_keys As Boolean = True) As Variant
    colName = p_LabelColumnRef(colName)

    Dim label As String
    Dim finalType As ctDataType

    expanding = p_WindowCore("cTable.expanding", colName, 0&, by, operation, _
                             "", min_periods, "right", dropna_keys, True, _
                             label, finalType)
End Function


'------------------------------------------------------------------------------
' WINDOW COLUMN WRITERS
'------------------------------------------------------------------------------

Public Sub AddRollingColumn(ByVal colName As Variant, _
                            ByVal window As Variant, _
                            Optional ByVal by As Variant, _
                            Optional ByVal operation As String = "sum", _
                            Optional ByVal on_arg As String = "", _
                            Optional ByVal newColName As String = "", _
                            Optional ByVal min_periods As Long = -1, _
                            Optional ByVal closed As String = "right", _
                            Optional ByVal resultType As ctDataType = ctVariant, _
                            Optional ByVal dropna_keys As Boolean = True)
    colName = p_LabelColumnRef(colName)

    p_ApplyWindow "cTable.AddRollingColumn", colName, window, by, operation, _
                  on_arg, newColName, min_periods, closed, resultType, _
                  dropna_keys, False, False
End Sub


Public Sub SetRollingColumn(ByVal colName As Variant, _
                            ByVal window As Variant, _
                            Optional ByVal by As Variant, _
                            Optional ByVal operation As String = "sum", _
                            Optional ByVal on_arg As String = "", _
                            Optional ByVal newColName As String = "", _
                            Optional ByVal min_periods As Long = -1, _
                            Optional ByVal closed As String = "right", _
                            Optional ByVal resultType As ctDataType = ctVariant, _
                            Optional ByVal dropna_keys As Boolean = True)
    colName = p_LabelColumnRef(colName)

    p_ApplyWindow "cTable.SetRollingColumn", colName, window, by, operation, _
                  on_arg, newColName, min_periods, closed, resultType, _
                  dropna_keys, True, False
End Sub


Public Function WithRollingColumn(ByVal colName As Variant, _
                                  ByVal window As Variant, _
                                  Optional ByVal by As Variant, _
                                  Optional ByVal operation As String = "sum", _
                                  Optional ByVal on_arg As String = "", _
                                  Optional ByVal newColName As String = "", _
                                  Optional ByVal min_periods As Long = -1, _
                                  Optional ByVal closed As String = "right", _
                                  Optional ByVal resultType As ctDataType = ctVariant, _
                                  Optional ByVal dropna_keys As Boolean = True) As cTable
    colName = p_LabelColumnRef(colName)

    Dim t As cTable
    Set t = Me.Clone

    t.AddRollingColumn colName, window, by, operation, on_arg, newColName, _
                       min_periods, closed, resultType, dropna_keys

    Set WithRollingColumn = t
End Function


Public Sub AddExpandingColumn(ByVal colName As Variant, _
                              Optional ByVal by As Variant, _
                              Optional ByVal operation As String = "sum", _
                              Optional ByVal newColName As String = "", _
                              Optional ByVal min_periods As Long = 1, _
                              Optional ByVal resultType As ctDataType = ctVariant, _
                              Optional ByVal dropna_keys As Boolean = True)
    colName = p_LabelColumnRef(colName)

    p_ApplyWindow "cTable.AddExpandingColumn", colName, 0&, by, operation, _
                  "", newColName, min_periods, "right", resultType, _
                  dropna_keys, False, True
End Sub


Public Sub SetExpandingColumn(ByVal colName As Variant, _
                              Optional ByVal by As Variant, _
                              Optional ByVal operation As String = "sum", _
                              Optional ByVal newColName As String = "", _
                              Optional ByVal min_periods As Long = 1, _
                              Optional ByVal resultType As ctDataType = ctVariant, _
                              Optional ByVal dropna_keys As Boolean = True)
    colName = p_LabelColumnRef(colName)

    p_ApplyWindow "cTable.SetExpandingColumn", colName, 0&, by, operation, _
                  "", newColName, min_periods, "right", resultType, _
                  dropna_keys, True, True
End Sub


Private Sub p_ApplyWindow(ByVal callerName As String, _
                          ByVal colName As String, _
                          ByVal window As Variant, _
                          ByVal by As Variant, _
                          ByVal operation As String, _
                          ByVal on_arg As String, _
                          ByVal newColName As String, _
                          ByVal min_periods As Long, _
                          ByVal closed As String, _
                          ByVal resultType As ctDataType, _
                          ByVal dropNAKeys As Boolean, _
                          ByVal allowReplace As Boolean, _
                          ByVal expandingMode As Boolean)

    Dim suffix As String
    If expandingMode Then
        suffix = "_exp" & LCase$(Trim$(operation))
    Else
        suffix = "_roll" & p_WindowText(window)
    End If

    Dim targetName As String
    targetName = p_ResolveTargetName(newColName, colName, suffix, _
                                     allowReplace, callerName)

    Dim label As String
    Dim finalType As ctDataType
    Dim values As Variant

    values = p_WindowCore(callerName, colName, window, by, operation, _
                          on_arg, min_periods, closed, dropNAKeys, _
                          expandingMode, label, finalType)

    If resultType <> ctVariant Then finalType = resultType

    p_WriteTransformColumn targetName, values, finalType, label
End Sub


'------------------------------------------------------------------------------
' WINDOW ENGINE
'
' One implementation behind rolling, expanding and all their column writers,
' so window semantics cannot drift between entry points.
'------------------------------------------------------------------------------

Private Function p_WindowCore(ByVal callerName As String, _
                              ByVal colName As String, _
                              ByVal window As Variant, _
                              ByVal by As Variant, _
                              ByVal operation As String, _
                              ByVal on_arg As String, _
                              ByVal min_periods As Long, _
                              ByVal closed As String, _
                              ByVal dropNAKeys As Boolean, _
                              ByVal expandingMode As Boolean, _
                              ByRef outLabel As String, _
                              ByRef outType As ctDataType) As Variant

    Dim ci As Long
    ci = p_ColIdx(colName)

    If ci = 0 Then
        Err.Raise vbObjectError + 1450, callerName, _
                  "Column not found: """ & colName & """"
    End If

    Dim opCode As Long
    opCode = p_WindowOpCode(operation, callerName)

    ' --- window shape ---------------------------------------------------
    Dim winRows As Long
    Dim winDays As Double
    Dim isOffset As Boolean

    If Not expandingMode Then
        p_ParseWindow window, winRows, winDays, isOffset, callerName
    End If

    Dim onCol As Long
    If isOffset Then
        If Len(Trim$(on_arg)) = 0 Then
            Err.Raise vbObjectError + 1451, callerName, _
                      "A time-based window such as """ & CStr(window) & _
                      """ needs on_arg to name the date column that " & _
                      "orders the rows."
        End If

        onCol = p_ColIdx(on_arg)

        If onCol = 0 Then
            Err.Raise vbObjectError + 1452, callerName, _
                      "on_arg column not found: """ & on_arg & """"
        End If

    ElseIf Len(Trim$(on_arg)) > 0 Then
        Err.Raise vbObjectError + 1453, callerName, _
                  "on_arg applies only to time-based windows. A numeric " & _
                  "window counts rows in their current order."
    End If

    Dim closedCode As Long
    closedCode = p_WindowClosedCode(closed, callerName)

    p_EnsureTypes

    Dim srcType As ctDataType
    srcType = m_ColTypes(ci)

    ' Reject combinations that would return arithmetic on things that are not
    ' quantities. min / max / first / last / count remain legal everywhere.
    Select Case opCode
        Case 1, 2, 6, 7, 8
            If srcType = ctString Or srcType = ctDate Then
                Err.Raise vbObjectError + 1454, callerName, _
                          "Operation """ & operation & """ is not " & _
                          "meaningful on a " & Me.dtype_name(srcType) & _
                          " column. Use min, max, first, last or count."
            End If
    End Select

    outType = p_WindowResultType(srcType, opCode)

    ' --- min_periods defaults -------------------------------------------
    Dim minP As Long
    minP = min_periods

    If minP < 0 Then
        If expandingMode Or isOffset Then
            minP = 1
        Else
            minP = winRows
        End If
    End If

    ' --- lineage label ---------------------------------------------------
    Dim keyCols() As Long
    Dim keyCount As Long
    keyCount = p_ResolveGroupKeys(keyCols, callerName, by)

    Dim extras As String
    p_AddExtra extras, "op=" & LCase$(Trim$(operation))

    If Not expandingMode Then
        p_AddExtra extras, "window=" & p_WindowText(window)
        If isOffset Then p_AddExtra extras, "on=" & m_Headers(onCol)
        If closedCode <> 1 Then
            p_AddExtra extras, "closed=" & LCase$(Trim$(closed))
        End If
    End If

    If minP <> IIf(expandingMode Or isOffset, 1, winRows) Then
        p_AddExtra extras, "min_periods=" & CStr(minP)
    End If

    Dim opName As String
    If expandingMode Then opName = "expanding" Else opName = "rolling"

    outLabel = p_OpLabel(opName, colName, keyCols, keyCount, extras)

    ' --- degenerate table -------------------------------------------------
    If m_RowCount = 0 Then
        p_WindowCore = Array()
        Exit Function
    End If

    ' --- grouping ---------------------------------------------------------
    Dim groupOf() As Long
    Dim memberStart() As Long
    Dim member() As Long
    Dim gc As Long

    gc = p_BuildGroups(keyCols, keyCount, dropNAKeys, _
                       groupOf, memberStart, member)

    p_WindowCore = p_WindowValues(ci, onCol, winDays, winRows, opCode, _
                                  minP, closedCode, expandingMode, isOffset, _
                                  memberStart, member, gc)
End Function


Private Function p_WindowValues(ByVal ci As Long, _
                                ByVal onCol As Long, _
                                ByVal winDays As Double, _
                                ByVal winRows As Long, _
                                ByVal opCode As Long, _
                                ByVal minP As Long, _
                                ByVal closedCode As Long, _
                                ByVal expandingMode As Boolean, _
                                ByVal isOffset As Boolean, _
                                ByRef memberStart() As Long, _
                                ByRef member() As Long, _
                                ByVal gc As Long) As Variant

    Const TOL As Double = 0.0000001         ' ~8.6 ms, guards date arithmetic

    Dim n As Long
    n = m_RowCount

    Dim Out() As Variant
    ReDim Out(1 To n)

    Dim r As Long
    For r = 1 To n
        Out(r) = Null
    Next r

    If gc = 0 Then
        p_WindowValues = Out
        Exit Function
    End If

    ' --- ordering key, once for the whole table --------------------------
    Dim onVal() As Double
    Dim onNA() As Boolean
    ReDim onVal(1 To n)
    ReDim onNA(1 To n)

    Dim v As Variant
    Dim ok As Boolean

    If isOffset Then
        For r = 1 To n
            v = p_TryConvert(m_Data(r, onCol), ctDate, ok)

            If (Not ok) Or IsNull(v) Then
                onNA(r) = True
            Else
                onVal(r) = CDbl(CDate(v))
            End If
        Next r
    End If

    Dim seqArr() As Long
    Dim work() As Long
    ReDim seqArr(1 To n)
    ReDim work(1 To n)

    ' --- walk each group --------------------------------------------------
    Dim g As Long
    Dim sLo As Long
    Dim sHi As Long
    Dim p As Long
    Dim k As Long
    Dim i As Long
    Dim L As Long
    Dim rr As Long
    Dim curL As Long
    Dim curR As Long
    Dim cnt As Long
    Dim s As Double
    Dim ss As Double
    Dim lo As Double
    Dim t As Double
    Dim d As Double

    For g = 1 To gc
        sLo = memberStart(g)
        sHi = memberStart(g + 1) - 1

        If sHi >= sLo Then

            ' Offset windows order the group by the on_arg column, so the
            ' table's own row order cannot corrupt the result.
            If isOffset Then
                p_MergeSortByKey member, work, sLo, sHi, onVal
            End If

            k = 0
            For p = sLo To sHi
                r = member(p)

                If isOffset Then
                    If Not onNA(r) Then
                        k = k + 1
                        seqArr(k) = r
                    End If
                Else
                    k = k + 1
                    seqArr(k) = r
                End If
            Next p

            curL = 1
            curR = 0
            cnt = 0
            s = 0
            ss = 0

            For i = 1 To k

                ' --- right edge --------------------------------------
                If isOffset Then
                    t = onVal(seqArr(i))
                    rr = curR

                    Do While rr < k
                        If closedCode <= 2 Then
                            If onVal(seqArr(rr + 1)) < t + TOL Then
                                rr = rr + 1
                            Else
                                Exit Do
                            End If
                        Else
                            If onVal(seqArr(rr + 1)) < t - TOL Then
                                rr = rr + 1
                            Else
                                Exit Do
                            End If
                        End If
                    Loop
                Else
                    rr = i
                End If

                ' --- left edge ---------------------------------------
                If expandingMode Then
                    L = 1
                ElseIf isOffset Then
                    lo = t - winDays
                    L = curL

                    Do While L <= rr
                        If closedCode = 2 Or closedCode = 3 Then
                            If onVal(seqArr(L)) > lo - TOL Then Exit Do
                        Else
                            If onVal(seqArr(L)) > lo + TOL Then Exit Do
                        End If
                        L = L + 1
                    Loop
                Else
                    L = i - winRows + 1
                    If L < 1 Then L = 1
                End If

                ' --- slide the accumulators --------------------------
                ' [L..rr] is contiguous, and both edges only ever move
                ' forward, so add-then-remove keeps curL <= curR + 1.
                Do While curR < rr
                    curR = curR + 1
                    v = m_Data(seqArr(curR), ci)

                    If Not p_IsNA(v) Then
                        d = p_WindowNum(v, ok)
                        If ok Then
                            cnt = cnt + 1
                            s = s + d
                            ss = ss + d * d
                        End If
                    End If
                Loop

                Do While curL < L
                    v = m_Data(seqArr(curL), ci)

                    If Not p_IsNA(v) Then
                        d = p_WindowNum(v, ok)
                        If ok Then
                            cnt = cnt - 1
                            s = s - d
                            ss = ss - d * d
                        End If
                    End If

                    curL = curL + 1
                Loop

                Out(seqArr(i)) = p_WindowResult(ci, seqArr, L, rr, opCode, _
                                                cnt, s, ss, minP)
            Next i
        End If
    Next g

    p_WindowValues = Out
End Function


' Turn the current window state into one cell value.
'
' sum / mean / count / std / var come from the running accumulators, so they
' cost O(1) per row. min / max / median / first / last scan the window, which
' is why they are the ones to avoid on very wide windows over very long series.
Private Function p_WindowResult(ByVal ci As Long, _
                                ByRef seqArr() As Long, _
                                ByVal L As Long, _
                                ByVal rr As Long, _
                                ByVal opCode As Long, _
                                ByVal cnt As Long, _
                                ByVal s As Double, _
                                ByVal ss As Double, _
                                ByVal minP As Long) As Variant

    If cnt < minP Then
        p_WindowResult = Null
        Exit Function
    End If

    Dim p As Long
    Dim v As Variant
    Dim best As Variant
    Dim d As Double
    Dim ok As Boolean
    Dim variance As Double

    Select Case opCode

        Case 1                                  ' sum
            p_WindowResult = s

        Case 2                                  ' mean
            If cnt < 1 Then
                p_WindowResult = Null
            Else
                p_WindowResult = s / cnt
            End If

        Case 3                                  ' count
            p_WindowResult = cnt

        Case 4, 5                               ' min / max
            best = Null

            Dim cmp As Long

            For p = L To rr
                v = m_Data(seqArr(p), ci)

                If Not p_IsNA(v) Then
                    If IsNull(best) Then
                        best = v
                    Else
                        cmp = p_WindowCompare(v, best)

                        If opCode = 4 Then
                            If cmp < 0 Then best = v
                        Else
                            If cmp > 0 Then best = v
                        End If
                    End If
                End If
            Next p

            p_WindowResult = best

        Case 6, 7                               ' std / var, ddof = 1
            If cnt < 2 Then
                p_WindowResult = Null
            Else
                variance = (ss - (s * s) / cnt) / (cnt - 1)
                If variance < 0 Then variance = 0      ' cancellation guard

                If opCode = 7 Then
                    p_WindowResult = variance
                Else
                    p_WindowResult = Sqr(variance)
                End If
            End If

        Case 8                                  ' median
            Dim vals() As Double
            ReDim vals(1 To cnt + 1)

            Dim m As Long
            m = 0

            For p = L To rr
                v = m_Data(seqArr(p), ci)

                If Not p_IsNA(v) Then
                    d = p_WindowNum(v, ok)
                    If ok Then
                        m = m + 1
                        vals(m) = d
                    End If
                End If
            Next p

            If m = 0 Then
                p_WindowResult = Null
            Else
                p_SortDoubles vals, 1, m

                If (m Mod 2) = 1 Then
                    p_WindowResult = vals((m + 1) \ 2)
                Else
                    p_WindowResult = (vals(m \ 2) + vals(m \ 2 + 1)) / 2
                End If
            End If

        Case 9                                  ' first non-missing
            p_WindowResult = Null

            For p = L To rr
                v = m_Data(seqArr(p), ci)
                If Not p_IsNA(v) Then
                    p_WindowResult = v
                    Exit For
                End If
            Next p

        Case 10                                 ' last non-missing
            p_WindowResult = Null

            For p = rr To L Step -1
                v = m_Data(seqArr(p), ci)
                If Not p_IsNA(v) Then
                    p_WindowResult = v
                    Exit For
                End If
            Next p

        Case Else
            p_WindowResult = Null
    End Select
End Function


'------------------------------------------------------------------------------
' WINDOW HELPERS
'------------------------------------------------------------------------------

' Parse a window argument into either a row count or a span in days.
Private Sub p_ParseWindow(ByVal window As Variant, _
                          ByRef winRows As Long, _
                          ByRef winDays As Double, _
                          ByRef isOffset As Boolean, _
                          ByVal callerName As String)

    winRows = 0
    winDays = 0
    isOffset = False

    If IsNull(window) Or IsEmpty(window) Then
        Err.Raise vbObjectError + 1455, callerName, _
                  "window is required."
    End If

    If VarType(window) <> vbString Then
        winRows = CLng(window)

        If winRows < 1 Then
            Err.Raise vbObjectError + 1456, callerName, _
                      "A numeric window must be 1 or more."
        End If

        Exit Sub
    End If

    Dim raw As String
    raw = Trim$(CStr(window))

    If Len(raw) = 0 Then
        Err.Raise vbObjectError + 1455, callerName, "window is required."
    End If

    ' A bare numeric string is a row count, matching rolling("5").
    If IsNumeric(raw) Then
        winRows = CLng(raw)

        If winRows < 1 Then
            Err.Raise vbObjectError + 1456, callerName, _
                      "A numeric window must be 1 or more."
        End If

        Exit Sub
    End If

    ' Split leading quantity from trailing unit.
    Dim i As Long
    Dim ch As String
    Dim numPart As String

    For i = 1 To Len(raw)
        ch = mid$(raw, i, 1)

        If (ch >= "0" And ch <= "9") Or ch = "." Then
            numPart = numPart & ch
        Else
            Exit For
        End If
    Next i

    Dim unitPart As String
    unitPart = LCase$(Trim$(mid$(raw, i)))

    Dim qty As Double
    If Len(numPart) = 0 Then
        qty = 1
    ElseIf IsNumeric(numPart) Then
        qty = CDbl(numPart)
    Else
        Err.Raise vbObjectError + 1457, callerName, _
                  "Could not read a quantity from window """ & raw & """."
    End If

    If qty <= 0 Then
        Err.Raise vbObjectError + 1456, callerName, _
                  "A window span must be greater than zero."
    End If

    Dim factor As Double

    Select Case unitPart
        Case "d", "day", "days":                   factor = 1
        Case "w", "week", "weeks":                 factor = 7
        Case "h", "hr", "hour", "hours":           factor = 1# / 24#
        Case "t", "min", "mins", "minute", "minutes": factor = 1# / 1440#
        Case "s", "sec", "secs", "second", "seconds": factor = 1# / 86400#

        Case "m", "month", "months"
            Err.Raise vbObjectError + 1458, callerName, _
                      "Month windows are ambiguous because months differ " & _
                      "in length. Use an explicit day count such as " & _
                      """30D"" or ""31D""."

        Case Else
            Err.Raise vbObjectError + 1458, callerName, _
                      "Unknown window unit """ & unitPart & """ in """ & _
                      raw & """. Use D, W, H, min or S."
    End Select

    winDays = qty * factor
    isOffset = True
End Sub


Private Function p_WindowText(ByVal window As Variant) As String
    If VarType(window) = vbString Then
        p_WindowText = UCase$(Trim$(CStr(window)))
    Else
        p_WindowText = CStr(window)
    End If
End Function


Private Function p_WindowOpCode(ByVal operation As String, _
                                ByVal callerName As String) As Long

    Select Case LCase$(Trim$(operation))
        Case "sum":                       p_WindowOpCode = 1
        Case "mean", "avg", "average":    p_WindowOpCode = 2
        Case "count":                     p_WindowOpCode = 3
        Case "min":                       p_WindowOpCode = 4
        Case "max":                       p_WindowOpCode = 5
        Case "std":                       p_WindowOpCode = 6
        Case "var":                       p_WindowOpCode = 7
        Case "median":                    p_WindowOpCode = 8
        Case "first":                     p_WindowOpCode = 9
        Case "last":                      p_WindowOpCode = 10

        Case Else
            Err.Raise vbObjectError + 1459, callerName, _
                      "Unknown window operation: """ & operation & """. " & _
                      "Use sum, mean, count, min, max, std, var, median, " & _
                      "first or last."
    End Select
End Function


' 1 right (default), 2 both, 3 left, 4 neither
Private Function p_WindowClosedCode(ByVal closed As String, _
                                    ByVal callerName As String) As Long

    Select Case LCase$(Trim$(closed))
        Case "", "right":   p_WindowClosedCode = 1
        Case "both":        p_WindowClosedCode = 2
        Case "left":        p_WindowClosedCode = 3
        Case "neither":     p_WindowClosedCode = 4

        Case Else
            Err.Raise vbObjectError + 1460, callerName, _
                      "closed must be right, both, left or neither."
    End Select
End Function


Private Function p_WindowResultType(ByVal srcType As ctDataType, _
                                    ByVal opCode As Long) As ctDataType
    Select Case opCode

        Case 3                                  ' count
            p_WindowResultType = ctLong

        Case 1                                  ' sum
            Select Case srcType
                Case ctCurrency
                    p_WindowResultType = ctCurrency
                Case ctLong, ctDouble, ctBoolean
                    p_WindowResultType = ctDouble
                Case Else
                    p_WindowResultType = ctVariant
            End Select

        Case 2, 6, 7, 8                         ' mean std var median
            p_WindowResultType = ctDouble

        Case Else                               ' min max first last
            p_WindowResultType = srcType
    End Select
End Function


' Numeric reading of a cell for accumulation. Anything that is not a quantity
' returns ok = False and is treated exactly like a missing value, so one stray
' text cell degrades a single observation rather than failing the run.
Private Function p_WindowNum(ByVal v As Variant, ByRef ok As Boolean) As Double
    ok = False

    If p_IsNA(v) Then Exit Function

    Select Case VarType(v)

        Case vbBoolean
            ' pandas counts True as 1, where VBA CDbl(True) is -1.
            If CBool(v) Then p_WindowNum = 1 Else p_WindowNum = 0
            ok = True

        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, _
             vbCurrency, vbDecimal, vbDate

            p_WindowNum = CDbl(v)
            ok = True

        Case vbString
            Dim s As String
            s = p_CleanNumericString(CStr(v))

            If IsNumeric(s) Then
                p_WindowNum = CDbl(s)
                ok = True
            End If
    End Select
End Function


' Ordering comparison for min / max. Quantities compare numerically; anything
' else falls back to the class-wide case-insensitive text convention.
Private Function p_WindowCompare(ByVal a As Variant, _
                                 ByVal b As Variant) As Long
    Dim da As Double
    Dim db As Double
    Dim oka As Boolean
    Dim okb As Boolean

    da = p_WindowNum(a, oka)
    db = p_WindowNum(b, okb)

    If oka And okb Then
        If da < db Then
            p_WindowCompare = -1
        ElseIf da > db Then
            p_WindowCompare = 1
        End If

        Exit Function
    End If

    p_WindowCompare = StrComp(CStr(a), CStr(b), vbTextCompare)
End Function


' Stable merge sort of a slice of idx() by keyVal(), keyed by ROW NUMBER.
' Stability matters: rows sharing a timestamp keep their table order, so
' first / last inside a window stay reproducible.
Private Sub p_MergeSortByKey(ByRef idx() As Long, _
                             ByRef work() As Long, _
                             ByVal lo As Long, _
                             ByVal hi As Long, _
                             ByRef keyVal() As Double)

    If hi - lo < 1 Then Exit Sub

    Dim mid As Long
    mid = (lo + hi) \ 2

    p_MergeSortByKey idx, work, lo, mid, keyVal
    p_MergeSortByKey idx, work, mid + 1, hi, keyVal

    Dim i As Long
    Dim j As Long
    Dim k As Long

    i = lo
    j = mid + 1
    k = lo

    Do While i <= mid And j <= hi
        If keyVal(idx(j)) < keyVal(idx(i)) Then
            work(k) = idx(j)
            j = j + 1
        Else
            work(k) = idx(i)
            i = i + 1
        End If

        k = k + 1
    Loop

    Do While i <= mid
        work(k) = idx(i)
        i = i + 1
        k = k + 1
    Loop

    Do While j <= hi
        work(k) = idx(j)
        j = j + 1
        k = k + 1
    Loop

    For k = lo To hi
        idx(k) = work(k)
    Next k
End Sub


' Insertion sort for the small slices the median path produces.
Private Sub p_SortDoubles(ByRef a() As Double, _
                          ByVal lo As Long, _
                          ByVal hi As Long)
    Dim i As Long
    Dim j As Long
    Dim key As Double

    For i = lo + 1 To hi
        key = a(i)
        j = i - 1

        Do While j >= lo
            If a(j) <= key Then Exit Do
            a(j + 1) = a(j)
            j = j - 1
        Loop

        a(j + 1) = key
    Next i
End Sub
'==============================================================================
' READABILITY DELEGATES
'==============================================================================

' Straight alias of dt. Same parts, same rules, same return shape.
'
'   v = t.dt_part("RunDATE", "month")     ' 8 for August, Null if not a date
Public Function dt_part(ByVal colName As Variant, _
                        ByVal part As String, _
                        Optional ByVal fiscalStartMonth As Long = 10) As Variant
    colName = p_LabelColumnRef(colName)

    dt_part = Me.dt(colName, part, fiscalStartMonth)
End Function


' Month and weekday names.
'
' which:  "month"  -> August       (default)
'         "day"    -> Friday
'         "mon"    -> Aug
'         "dy"     -> Fri
'
' CAUTION: these come from VBA Format$, which is LOCALE-dependent. On a French
' Windows install "month" returns "ao  t". If a report has to read the same on
' every machine, format the number instead and map it yourself.
Public Function dt_name(ByVal colName As Variant, _
                        Optional ByVal which As String = "month") As Variant
    colName = p_LabelColumnRef(colName)

    Select Case LCase$(Trim$(which))
        Case "", "month", "monthname", "month_name"
            dt_name = Me.dt(colName, "month_name")

        Case "day", "dayname", "day_name", "weekday"
            dt_name = Me.dt(colName, "day_name")

        Case "mon", "monthabbr", "month_abbr"
            dt_name = Me.dt_format(colName, "mmm")

        Case "dy", "dayabbr", "day_abbr"
            dt_name = Me.dt_format(colName, "ddd")

        Case Else
            Err.Raise vbObjectError + 1470, "cTable.dt_name", _
                      "which must be month, day, mon or dy."
    End Select
End Function


' Fiscal year. startMonth defaults to 10 for the US federal year, and a fiscal
' year beginning after January is labelled by the calendar year it ENDS in, so
' 2026-10-01 is FY2027. startMonth:=1 makes FY the calendar year.
'
'   v = t.dt_fy("RunDATE")        ' federal
'   v = t.dt_fy("RunDATE", 7)     ' July start
Public Function dt_fy(ByVal colName As Variant, _
                      Optional ByVal startMonth As Long = 10) As Variant
    colName = p_LabelColumnRef(colName)

    dt_fy = Me.dt(colName, "fiscal_year", startMonth)
End Function


' Fiscal quarter, 1 to 4, counted from startMonth.
Public Function dt_fq(ByVal colName As Variant, _
                      Optional ByVal startMonth As Long = 10) As Variant
    colName = p_LabelColumnRef(colName)

    dt_fq = Me.dt(colName, "fiscal_quarter", startMonth)
End Function


'==============================================================================
'
' dt_format   -   a date column rendered as text
'
'   v = t.dt_format("RunDATE", "us")            ' 08/04/2026
'   v = t.dt_format("RunDATE", "yyyy-mm-dd")    ' 2026-08-04
'   v = t.dt_format("RunDATE", "%Y-%m-%d")      ' 2026-08-04
'   t.AddDtFormatColumn "RunDATE", "us", "RunDateText"
'
' Cells that are missing or unreadable as a date yield Null, exactly as dt
' does. The result dtype is ctString.
'
' THREE WAYS TO SPELL A FORMAT, tried in this order:
'
' 1. A NAMED PRESET. Safest, because it sidesteps both traps below.
'
'      iso            2026-08-04
'      iso_datetime   2026-08-04 14:30:00
'      us             08/04/2026
'      us_short       08/04/26
'      us_datetime    08/04/2026 02:30 PM
'      eu             04/08/2026
'      eu_dot         04.08.2026
'      compact        20260804
'      month_year     August 2026
'      mon_year       Aug 2026
'      time24         14:30:00
'      time12         02:30 PM
'
' 2. A strftime PATTERN, recognised by the "%" sign. Use this when you want
'    pandas-shaped code or when minutes are involved, because %M is
'    unambiguous where VBA's "mm" is not.
'
'      %Y %y %m %d %H %I %M %S %B %b %A %a %p %j %%
'
' 3. A RAW VBA Format$ PATTERN, for anything the first two do not cover.
'
' TWO VBA TRAPS THIS FUNCTION DEFUSES
'
'   * "/" and ":" are not literals to VBA Format$. They are placeholders for
'     the SYSTEM date and time separators, so "mm/dd/yyyy" renders as
'     08.04.2026 on a machine configured with dots. Every "/" and ":" in a
'     raw or strftime pattern is escaped here, so what you write is what you
'     get on every machine.
'
'   * In a VBA pattern "mm" means MONTH and "nn" means MINUTE. "hh:mm"
'     silently produces the hour and the month. This function cannot guess
'     your intent, so it leaves raw patterns alone - use "hh:nn", the time24
'     preset, or "%H:%M" instead.
'==============================================================================

Public Function dt_format(ByVal colName As Variant, _
                          Optional ByVal fmt As String = "iso") As Variant
    colName = p_LabelColumnRef(colName)

    Const CALLER As String = "cTable.dt_format"

    Dim ci As Long
    ci = p_ColIdx(colName)

    If ci = 0 Then
        Err.Raise vbObjectError + 1471, CALLER, _
                  "Column not found: """ & colName & """"
    End If

    Dim pattern As String
    pattern = p_ResolveDateFormat(fmt)

    If m_RowCount = 0 Then
        dt_format = Array()
        Exit Function
    End If

    p_EnsureTypes

    Dim Out() As Variant
    ReDim Out(1 To m_RowCount)

    Dim r As Long
    Dim v As Variant
    Dim ok As Boolean

    For r = 1 To m_RowCount
        v = p_TryConvert(m_Data(r, ci), ctDate, ok)

        If (Not ok) Or IsNull(v) Then
            Out(r) = Null
        Else
            Out(r) = Format$(CDate(v), pattern)
        End If
    Next r

    dt_format = Out
End Function


Public Sub AddDtFormatColumn(ByVal colName As Variant, _
                             Optional ByVal fmt As String = "iso", _
                             Optional ByVal newColName As String = "")
    colName = p_LabelColumnRef(colName)

    p_ApplyDtFormat colName, fmt, newColName, False, _
                    "cTable.AddDtFormatColumn"
End Sub


Public Sub SetDtFormatColumn(ByVal colName As Variant, _
                             Optional ByVal fmt As String = "iso", _
                             Optional ByVal newColName As String = "")
    colName = p_LabelColumnRef(colName)

    p_ApplyDtFormat colName, fmt, newColName, True, _
                    "cTable.SetDtFormatColumn"
End Sub


Public Function WithDtFormatColumn(ByVal colName As Variant, _
                                   Optional ByVal fmt As String = "iso", _
                                   Optional ByVal newColName As String = "") As cTable
    colName = p_LabelColumnRef(colName)
    Dim t As cTable
    Set t = Me.Clone

    t.AddDtFormatColumn colName, fmt, newColName

    Set WithDtFormatColumn = t
End Function


Private Sub p_ApplyDtFormat(ByVal colName As String, _
                            ByVal fmt As String, _
                            ByVal newColName As String, _
                            ByVal allowReplace As Boolean, _
                            ByVal callerName As String)

    If p_ColIdx(colName) = 0 Then
        Err.Raise vbObjectError + 1471, callerName, _
                  "Column not found: """ & colName & """"
    End If

    Dim targetName As String
    targetName = p_ResolveTargetName(newColName, colName, "_Text", _
                                     allowReplace, callerName)

    Dim values As Variant
    values = Me.dt_format(colName, fmt)

    ' The RESOLVED pattern goes in the label, so "us" and "mm/dd/yyyy" intern
    ' to one calculated source and genuinely different formats intern to two.
    Dim extras As String
    p_AddExtra extras, "fmt=" & p_ResolveDateFormat(fmt)

    Dim noKeys() As Long

    p_WriteTransformColumn targetName, values, ctString, _
        p_OpLabel("dt.format", colName, noKeys, 0, extras)
End Sub


'------------------------------------------------------------------------------
' FORMAT RESOLUTION
'------------------------------------------------------------------------------

Private Function p_ResolveDateFormat(ByVal fmt As String) As String
    Dim f As String
    f = Trim$(fmt)

    If Len(f) = 0 Then f = "iso"

    ' Preset lookup ignores case, spaces and underscores, so "us_short",
    ' "usshort" and "US Short" are the same preset.
    Dim keyName As String
    keyName = LCase$(Replace(Replace(f, " ", ""), "_", ""))

    Select Case keyName
        Case "iso", "isodate":  p_ResolveDateFormat = "yyyy-mm-dd"
        Case "isodatetime":     p_ResolveDateFormat = "yyyy-mm-dd hh\:nn\:ss"
        Case "us", "usdate":    p_ResolveDateFormat = "mm\/dd\/yyyy"
        Case "usshort":         p_ResolveDateFormat = "mm\/dd\/yy"
        Case "usdatetime":      p_ResolveDateFormat = "mm\/dd\/yyyy hh\:nn AM/PM"
        Case "eu", "eudate":    p_ResolveDateFormat = "dd\/mm\/yyyy"
        Case "eudot":           p_ResolveDateFormat = "dd\.mm\.yyyy"
        Case "compact":         p_ResolveDateFormat = "yyyymmdd"
        Case "monthyear":       p_ResolveDateFormat = "mmmm yyyy"
        Case "monyear":         p_ResolveDateFormat = "mmm yyyy"
        Case "time24":          p_ResolveDateFormat = "hh\:nn\:ss"
        Case "time12":          p_ResolveDateFormat = "hh\:nn AM/PM"

        Case Else
            ' Escape first, so a "/" the caller typed becomes a literal while
            ' the "/" that %p introduces inside AM/PM does not.
            Dim escaped As String
            escaped = p_EscapeDateSeparators(f)

            If InStr(f, "%") > 0 Then
                p_ResolveDateFormat = p_StrftimeToVba(escaped)
            Else
                p_ResolveDateFormat = escaped
            End If
    End Select
End Function


' VBA Format$ reads "/" as "the system date separator" and ":" as "the system
' time separator". A leading backslash forces the literal character instead.
' Characters already escaped are left alone.
Private Function p_EscapeDateSeparators(ByVal f As String) As String
    Dim outBuf As String
    Dim i As Long
    Dim ch As String
    Dim prev As String

    For i = 1 To Len(f)
        ch = mid$(f, i, 1)

        If i > 1 Then prev = mid$(f, i - 1, 1) Else prev = ""

        If (ch = "/" Or ch = ":" Or ch = ".") And prev <> "\" Then
            outBuf = outBuf & "\" & ch
        Else
            outBuf = outBuf & ch
        End If
    Next i

    p_EscapeDateSeparators = outBuf
End Function


' strftime tokens to VBA Format$ tokens. Longest-lived gotcha handled here:
' %M is MINUTES and becomes "nn", never "mm".
Private Function p_StrftimeToVba(ByVal f As String) As String
    Dim s As String
    s = f

    ' A doubled percent is a literal percent. Park it first so the token
    ' replacements below cannot see it. This has to be a Dim, not a Const:
    ' VBA will not accept a function call in a constant expression.
    Dim pctMark As String
    pctMark = Chr$(1)

    s = Replace(s, "%%", pctMark)

    s = Replace(s, "%Y", "yyyy")
    s = Replace(s, "%y", "yy")
    s = Replace(s, "%m", "mm")
    s = Replace(s, "%d", "dd")
    s = Replace(s, "%B", "mmmm")
    s = Replace(s, "%b", "mmm")
    s = Replace(s, "%A", "dddd")
    s = Replace(s, "%a", "ddd")
    s = Replace(s, "%j", "y")
    s = Replace(s, "%H", "hh")
    s = Replace(s, "%I", "hh")
    s = Replace(s, "%M", "nn")
    s = Replace(s, "%S", "ss")
    s = Replace(s, "%p", "AM/PM")

    s = Replace(s, pctMark, "%")

    p_StrftimeToVba = s
End Function


'==============================================================================
' SCALAR HELPERS
'
' The column accessors above all return arrays, which is the right shape for
' a column but the wrong shape inside an AddComputedColumn callback, where
' you hold one value from a row Dictionary. These take that value directly.
'
'   Public Function TagQuarter(ByVal rowValues As Object) As Variant
'       TagQuarter = gTbl.dt_value(rowValues("RunDATE"), "fiscal_quarter")
'   End Function
'==============================================================================

Public Function dt_value(ByVal v As Variant, _
                         ByVal part As String, _
                         Optional ByVal fiscalStartMonth As Long = 10) As Variant

    If fiscalStartMonth < 1 Or fiscalStartMonth > 12 Then
        Err.Raise vbObjectError + 1472, "cTable.dt_value", _
                  "fiscalStartMonth must be 1 to 12."
    End If

    Dim partCode As Long
    partCode = p_DtPartCode(part, "cTable.dt_value")

    Dim converted As Variant
    Dim ok As Boolean
    converted = p_TryConvert(v, ctDate, ok)

    If (Not ok) Or IsNull(converted) Then
        dt_value = Null
    Else
        dt_value = p_DtPart(CDate(converted), partCode, fiscalStartMonth)
    End If
End Function


Public Function dt_format_value(ByVal v As Variant, _
                                Optional ByVal fmt As String = "iso") As Variant
    Dim converted As Variant
    Dim ok As Boolean
    converted = p_TryConvert(v, ctDate, ok)

    If (Not ok) Or IsNull(converted) Then
        dt_format_value = Null
    Else
        dt_format_value = Format$(CDate(converted), p_ResolveDateFormat(fmt))
    End If
End Function

' ===================================================================================
' added 20260810-1823
' ---------------------------------------------------------------------------
' SetColumnValue
'
' Creates or replaces a column with one constant value.
'
' This is the vectorized equivalent of:
'   For r = 1 To t.rowCount
'       t.SetValueByName r, colName, value
'   Next r
'
' resultType:
'   ctVariant  = leave the column untyped unless InferIfVariant=True
'   Other type = convert the constant value and assign that dtype
'
' sourceId:
'   Optional lineage source for the whole column. By default this is Literal.
'
' Examples:
'   t.SetColumnValue "Source", filePath, ctString
'   t.SetColumnValue "IsCurrent", True, ctBoolean
' ---------------------------------------------------------------------------
Public Sub SetColumnValue(ByVal colName As String, _
                          ByVal value As Variant, _
                          Optional ByVal resultType As ctDataType = ctVariant, _
                          Optional ByVal errors As ctConvertErrors = ctErrorsRaise, _
                          Optional ByVal InferIfVariant As Boolean = False)

    Dim targetCol As Long
    targetCol = p_ColIdx(colName)

    If targetCol = 0 Then
        Me.InsertColumn colName
        targetCol = m_ColCount
    End If

    p_EnsureTypes

    Dim finalValue As Variant
    Dim ok As Boolean
    Dim r As Long

    If resultType <> ctVariant Then
        finalValue = p_TryConvert(value, resultType, ok)

        If Not ok Then
            Select Case errors
                Case ctErrorsRaise
                    Err.Raise vbObjectError + 1480, _
                              "cTable.SetColumnValue", _
                              "Cannot convert value """ & CStr(value) & _
                              """ to " & Me.dtype_name(resultType) & "."

                Case ctErrorsCoerce
                    finalValue = Null

                Case Else
                    finalValue = value
            End Select
        End If
    Else
        finalValue = value
    End If

    For r = 1 To m_RowCount
        m_Data(r, targetCol) = finalValue
    Next r

    If resultType <> ctVariant Then
        m_ColTypes(targetCol) = resultType

    ElseIf InferIfVariant Then
        m_ColTypes(targetCol) = p_InferColumnType(targetCol)

        If m_ColTypes(targetCol) <> ctVariant Then
            p_CoerceColumn targetCol, m_ColTypes(targetCol), ctErrorsPreserve
        End If

    Else
        m_ColTypes(targetCol) = ctVariant
    End If

    If Not m_Lin Is Nothing Then
        p_LinSync
        m_Lin.StampColumn targetCol, ctSrcLiteral
    End If

    p_ClearIndexes
End Sub

' ---------------------------------------------------------------------------
' RequireColumns
'
' Raises one clear error if any required columns are missing.
'
' requiredColumns may be:
'   - a single column name string
'   - an array of column names
'
' Example:
'   t.RequireColumns Array("AccountID", "RunDATE", "Amount")
' ---------------------------------------------------------------------------
Public Sub RequireColumns(ByVal requiredColumns As Variant, _
                          Optional ByVal callerName As String = "cTable.RequireColumns")

    Dim names As Variant

    If IsArray(requiredColumns) Then
        If Not p_ArrayHasElements(requiredColumns) Then Exit Sub
        names = requiredColumns
    Else
        If IsObject(requiredColumns) Or IsNull(requiredColumns) Or IsEmpty(requiredColumns) Then
            Err.Raise vbObjectError + 1490, _
                      callerName, _
                      "requiredColumns must be a column name or an array of column names."
        End If

        names = Array(CStr(requiredColumns))
    End If

    Dim missing As Object
    Set missing = CreateObject("Scripting.Dictionary")
    missing.CompareMode = vbTextCompare

    Dim i As Long
    Dim nm As String

    For i = LBound(names) To UBound(names)
        nm = Trim$(CStr(names(i)))

        If Len(nm) = 0 Then
            Err.Raise vbObjectError + 1491, _
                      callerName, _
                      "requiredColumns contains a blank column name."
        End If

        If p_ColIdx(nm) = 0 Then
            missing(nm) = True
        End If
    Next i

    If missing.count > 0 Then
        Err.Raise vbObjectError + 1492, _
                  callerName, _
                  "Missing required column(s): " & p_JoinVariantArray(missing.keys, ", ")
    End If
End Sub
' ---------------------------------------------------------------------------
' astype_many
'
' Converts multiple columns in one call.
'
' specs must be an array whose elements are two-item arrays:
'   Array(columnName, ctDataType)
'
' Example:
'   t.astype_many Array( _
'       Array("RunDATE", ctDate), _
'       Array("AccountID", ctString), _
'       Array("Amount", ctDouble)), _
'       ctErrorsCoerce
'
' Returns:
'   Scripting.Dictionary column name -> failed conversion count
' ---------------------------------------------------------------------------
Public Function astype_many(ByVal specs As Variant, _
                            Optional ByVal errors As ctConvertErrors = ctErrorsPreserve, _
                            Optional ByVal IgnoreMissing As Boolean = False) As Object

    If Not IsArray(specs) Then
        Err.Raise vbObjectError + 1493, _
                  "cTable.astype_many", _
                  "specs must be an array of Array(columnName, ctDataType) entries."
    End If

    If Not p_ArrayHasElements(specs) Then
        Dim emptyResults As Object
        Set emptyResults = CreateObject("Scripting.Dictionary")
        emptyResults.CompareMode = vbTextCompare
        Set astype_many = emptyResults
        Exit Function
    End If

    Dim results As Object
    Set results = CreateObject("Scripting.Dictionary")
    results.CompareMode = vbTextCompare

    Dim i As Long
    Dim spec As Variant
    Dim colName As String
    Dim targetType As ctDataType
    Dim ci As Long
    Dim failCount As Long

    p_EnsureTypes

    For i = LBound(specs) To UBound(specs)

        spec = specs(i)

        If Not IsArray(spec) Then
            Err.Raise vbObjectError + 1494, _
                      "cTable.astype_many", _
                      "Each specs element must be Array(columnName, ctDataType)."
        End If

        If UBound(spec) - LBound(spec) + 1 < 2 Then
            Err.Raise vbObjectError + 1495, _
                      "cTable.astype_many", _
                      "Each specs element must contain columnName and ctDataType."
        End If

        colName = CStr(spec(LBound(spec)))
        targetType = CLng(spec(LBound(spec) + 1))

        ci = p_ColIdx(colName)

        If ci = 0 Then
            If IgnoreMissing Then
                results(colName) = Null
            Else
                Err.Raise vbObjectError + 1496, _
                          "cTable.astype_many", _
                          "Column not found: """ & colName & """"
            End If
        Else
            failCount = p_CoerceColumn(ci, targetType, errors)
            m_ColTypes(ci) = targetType
            results(colName) = failCount
        End If
    Next i

    p_ClearIndexes

    Set astype_many = results
End Function
' ---------------------------------------------------------------------------
' append_many
'
' Appends multiple cTable objects to this table in place.
'
' tables may be:
'   - Collection of cTable objects
'   - Array of cTable objects
'
' strict is passed through to Append.
'
' Returns the number of non-empty tables appended.
' ---------------------------------------------------------------------------
Public Function append_many(ByVal tables As Variant, _
                            Optional ByVal strict As Boolean = True, _
                            Optional ByVal SkipEmpty As Boolean = True) As Long

    Dim appended As Long
    Dim item As Variant
    Dim tbl As cTable
    Dim i As Long

    If IsObject(tables) Then

        If typeName(tables) <> "Collection" Then
            Err.Raise vbObjectError + 1497, _
                      "cTable.append_many", _
                      "tables must be a Collection or an array of cTable objects."
        End If

        For Each item In tables
            If Not item Is Nothing Then
                Set tbl = item

                If Not (SkipEmpty And (tbl.ColumnCount = 0 Or tbl.rowCount = 0)) Then
                    Me.Append tbl, strict
                    appended = appended + 1
                End If
            End If
        Next item

    ElseIf IsArray(tables) Then

        If Not p_ArrayHasElements(tables) Then
            append_many = 0
            Exit Function
        End If

        For i = LBound(tables) To UBound(tables)
            If IsObject(tables(i)) Then
                If Not tables(i) Is Nothing Then
                    Set tbl = tables(i)

                    If Not (SkipEmpty And (tbl.ColumnCount = 0 Or tbl.rowCount = 0)) Then
                        Me.Append tbl, strict
                        appended = appended + 1
                    End If
                End If
            Else
                Err.Raise vbObjectError + 1498, _
                          "cTable.append_many", _
                          "Array element " & CStr(i) & " is not a cTable object."
            End If
        Next i

    Else
        Err.Raise vbObjectError + 1497, _
                  "cTable.append_many", _
                  "tables must be a Collection or an array of cTable objects."
    End If

    append_many = appended
End Function
'---------------------------------------------------
' These wre called out as optional improvements
'---------------------------------------------------
' ---------------------------------------------------------------------------
' from_tables
'
' Builds a new cTable by cloning the first non-empty table and appending the
' rest. The source tables are not modified.
'
' tables may be:
'   - Collection of cTable objects
'   - Array of cTable objects
'
' Example:
'   Dim builder As New cTable
'   Set combined = builder.from_tables(loadedTables, strict:=False)
' ---------------------------------------------------------------------------
Public Function from_tables(ByVal tables As Variant, _
                            Optional ByVal strict As Boolean = True, _
                            Optional ByVal SkipEmpty As Boolean = True) As cTable

    Dim result As cTable
    Dim tbl As cTable
    Dim item As Variant
    Dim i As Long
    Dim started As Boolean

    Set result = New cTable

    If IsObject(tables) Then

        If typeName(tables) <> "Collection" Then
            Err.Raise vbObjectError + 1499, _
                      "cTable.from_tables", _
                      "tables must be a Collection or an array of cTable objects."
        End If

        For Each item In tables
            If Not item Is Nothing Then
                Set tbl = item

                If Not (SkipEmpty And (tbl.ColumnCount = 0 Or tbl.rowCount = 0)) Then
                    If Not started Then
                        Set result = tbl.Clone
                        started = True
                    Else
                        result.Append tbl, strict
                    End If
                End If
            End If
        Next item

    ElseIf IsArray(tables) Then

        If p_ArrayHasElements(tables) Then
            For i = LBound(tables) To UBound(tables)
                If IsObject(tables(i)) Then
                    If Not tables(i) Is Nothing Then
                        Set tbl = tables(i)

                        If Not (SkipEmpty And (tbl.ColumnCount = 0 Or tbl.rowCount = 0)) Then
                            If Not started Then
                                Set result = tbl.Clone
                                started = True
                            Else
                                result.Append tbl, strict
                            End If
                        End If
                    End If
                Else
                    Err.Raise vbObjectError + 1500, _
                              "cTable.from_tables", _
                              "Array element " & CStr(i) & " is not a cTable object."
                End If
            Next i
        End If

    Else
        Err.Raise vbObjectError + 1499, _
                  "cTable.from_tables", _
                  "tables must be a Collection or an array of cTable objects."
    End If

    Set from_tables = result
End Function

'------------------------------------------------------------------------------
'       STRING METHODS
'------------------------------------------------------------------------------
Public Sub str_trim(ByVal colName As Variant, _
                    Optional ByVal side As String = "both")
    colName = p_LabelColumnRef(colName)

    Dim ci As Long
    ci = p_ColIdx(colName)

    If ci = 0 Then
        Err.Raise vbObjectError + 1500, "cTable.str_trim", _
                  "Column not found: """ & colName & """"
    End If

    Dim r As Long
    Dim v As Variant
    Dim s As String
    Dim mode As String

    mode = LCase$(Trim$(side))

    For r = 1 To m_RowCount
        v = m_Data(r, ci)

        If Not p_IsNA(v) Then
            s = CStr(v)

            Select Case mode
                Case "both", "trim"
                    m_Data(r, ci) = Trim$(s)

                Case "left", "ltrim"
                    m_Data(r, ci) = LTrim$(s)

                Case "right", "rtrim"
                    m_Data(r, ci) = RTrim$(s)

                Case Else
                    Err.Raise vbObjectError + 1501, "cTable.str_trim", _
                              "side must be ""both"", ""left"", or ""right""."
            End Select
        End If
    Next r

    p_EnsureTypes
    m_ColTypes(ci) = ctString

    p_ClearIndexes
End Sub
'==============================================================================
'
' STRING HELPERS
'
' Pandas-inspired string operations for one column.
'
' In-place mutators:
'   t.str_upper "CustomerName"
'   t.str_lower "CustomerName"
'   t.str_title_case "CustomerName"
'   t.str_replace "CustomerName", "old", "new"
'
' Mask:
'   v = t.str_contains("CustomerName", "abc")
'
' Missing values remain Null and are not converted to strings.
'==============================================================================

'------------------------------------------------------------------------------
' Convert each non-NA value in a column to uppercase, in place.
'------------------------------------------------------------------------------
Public Sub str_upper(ByVal colName As Variant)
    colName = p_LabelColumnRef(colName)
    p_StrApplySimple colName, "upper", "cTable.str_upper"
End Sub


'------------------------------------------------------------------------------
' Convert each non-NA value in a column to lowercase, in place.
'------------------------------------------------------------------------------
Public Sub str_lower(ByVal colName As Variant)
    colName = p_LabelColumnRef(colName)
    p_StrApplySimple colName, "lower", "cTable.str_lower"
End Sub


'------------------------------------------------------------------------------
' Convert each non-NA value in a column to title case, in place.
'
' Uses VBA StrConv(..., vbProperCase). This is simple title casing and may not
' perfectly match pandas or specialized name-casing rules for apostrophes,
' acronyms, or particles.
'------------------------------------------------------------------------------
Public Sub str_title_case(ByVal colName As Variant)
    colName = p_LabelColumnRef(colName)
    p_StrApplySimple colName, "title", "cTable.str_title_case"
End Sub


'------------------------------------------------------------------------------
' Replace text in each non-NA value in a column, in place.
'
' findText:
'   Text to search for.
'
' replaceText:
'   Replacement text.
'
' compare:
'   vbTextCompare   = case-insensitive replacement, default
'   vbBinaryCompare = case-sensitive replacement
'
' Example:
'   t.str_replace "CustomerName", "inc.", "Inc"
'   t.str_replace "CustomerName", "abc", "xyz", vbBinaryCompare
'------------------------------------------------------------------------------
Public Sub str_replace(ByVal colName As Variant, _
                       ByVal findText As String, _
                       ByVal replaceText As String, _
                       Optional ByVal compare As VbCompareMethod = vbTextCompare)
    colName = p_LabelColumnRef(colName)

    Dim ci As Long
    ci = p_ColIdx(colName)

    If ci = 0 Then
        Err.Raise vbObjectError + 1510, "cTable.str_replace", _
                  "Column not found: """ & colName & """"
    End If

    If Len(findText) = 0 Then
        Err.Raise vbObjectError + 1511, "cTable.str_replace", _
                  "findText cannot be blank."
    End If

    Dim r As Long
    Dim v As Variant

    For r = 1 To m_RowCount
        v = m_Data(r, ci)

        If Not p_IsNA(v) Then
            m_Data(r, ci) = Replace(CStr(v), findText, replaceText, , , compare)

            If Not m_Lin Is Nothing Then
                m_Lin.SetCell r, ci, ctSrcCalculated
            End If
        End If
    Next r

    p_EnsureTypes
    m_ColTypes(ci) = ctString

    p_ClearIndexes
End Sub


'------------------------------------------------------------------------------
' Return a 1-based Boolean array indicating whether each row's column value
' contains searchText.
'
' Missing values return False by default.
'
' compare:
'   vbTextCompare   = case-insensitive search, default
'   vbBinaryCompare = case-sensitive search
'
' Example:
'   Dim mask As Variant
'   mask = t.str_contains("CustomerName", "abc")
'------------------------------------------------------------------------------
Public Function str_contains(ByVal colName As Variant, _
                             ByVal searchText As String, _
                             Optional ByVal compare As VbCompareMethod = vbTextCompare) As Variant
    colName = p_LabelColumnRef(colName)

    Dim ci As Long
    ci = p_ColIdx(colName)

    If ci = 0 Then
        Err.Raise vbObjectError + 1520, "cTable.str_contains", _
                  "Column not found: """ & colName & """"
    End If

    If Len(searchText) = 0 Then
        Err.Raise vbObjectError + 1521, "cTable.str_contains", _
                  "searchText cannot be blank."
    End If

    If m_RowCount = 0 Then
        str_contains = Array()
        Exit Function
    End If

    Dim Out() As Boolean
    ReDim Out(1 To m_RowCount)

    Dim r As Long
    Dim v As Variant

    For r = 1 To m_RowCount
        v = m_Data(r, ci)

        If p_IsNA(v) Then
            Out(r) = False
        Else
            Out(r) = (InStr(1, CStr(v), searchText, compare) > 0)
        End If
    Next r

    str_contains = Out
End Function


'------------------------------------------------------------------------------
' Shared implementation for simple one-column string transformations.
'
' operation:
'   upper
'   lower
'   title
'------------------------------------------------------------------------------
Private Sub p_StrApplySimple(ByVal colName As String, _
                             ByVal operation As String, _
                             ByVal callerName As String)

    Dim ci As Long
    ci = p_ColIdx(colName)

    If ci = 0 Then
        Err.Raise vbObjectError + 1530, callerName, _
                  "Column not found: """ & colName & """"
    End If

    Dim r As Long
    Dim v As Variant
    Dim s As String

    For r = 1 To m_RowCount
        v = m_Data(r, ci)

        If Not p_IsNA(v) Then
            s = CStr(v)

            Select Case LCase$(operation)
                Case "upper"
                    m_Data(r, ci) = UCase$(s)

                Case "lower"
                    m_Data(r, ci) = LCase$(s)

                Case "title"
                    m_Data(r, ci) = StrConv(s, vbProperCase)

                Case Else
                    Err.Raise vbObjectError + 1531, callerName, _
                              "Unknown string operation: """ & operation & """"
            End Select

            If Not m_Lin Is Nothing Then
                m_Lin.SetCell r, ci, ctSrcCalculated
            End If
        End If
    Next r

    p_EnsureTypes
    m_ColTypes(ci) = ctString

    p_ClearIndexes
End Sub

Public Sub AddStrCountColumn(ByVal colName As Variant, _
                             ByVal pattern As String, _
                             Optional ByVal newColName As String = "", _
                             Optional ByVal match_mode As String = "literal", _
                             Optional ByVal compare As VbCompareMethod = vbTextCompare, _
                             Optional ByVal regex_ignore_case As Boolean = True, _
                             Optional ByVal case_sensitive As Boolean = False, _
                             Optional ByVal na_value As Long = 0)
    colName = p_LabelColumnRef(colName)

    p_ApplyStrCount colName, pattern, newColName, match_mode, compare, _
                    regex_ignore_case, case_sensitive, na_value, False, _
                    "cTable.AddStrCountColumn"
End Sub


Public Sub SetStrCountColumn(ByVal colName As Variant, _
                             ByVal pattern As String, _
                             Optional ByVal newColName As String = "", _
                             Optional ByVal match_mode As String = "literal", _
                             Optional ByVal compare As VbCompareMethod = vbTextCompare, _
                             Optional ByVal regex_ignore_case As Boolean = True, _
                             Optional ByVal case_sensitive As Boolean = False, _
                             Optional ByVal na_value As Long = 0)
    colName = p_LabelColumnRef(colName)

    p_ApplyStrCount colName, pattern, newColName, match_mode, compare, _
                    regex_ignore_case, case_sensitive, na_value, True, _
                    "cTable.SetStrCountColumn"
End Sub
Private Sub p_ApplyStrCount(ByVal colName As String, _
                            ByVal pattern As String, _
                            ByVal newColName As String, _
                            ByVal match_mode As String, _
                            ByVal compare As VbCompareMethod, _
                            ByVal regex_ignore_case As Boolean, _
                            ByVal case_sensitive As Boolean, _
                            ByVal na_value As Long, _
                            ByVal allowReplace As Boolean, _
                            ByVal callerName As String)

    If p_ColIdx(colName) = 0 Then
        Err.Raise vbObjectError + 1543, callerName, _
                  "Column not found: """ & colName & """"
    End If

    Dim targetName As String
    targetName = p_ResolveTargetName(newColName, colName, "_count", _
                                     allowReplace, callerName)

    Dim values As Variant
    values = Me.str_count(colName, pattern, match_mode, compare, _
                          regex_ignore_case, case_sensitive, na_value)

    Dim extras As String
    p_AddExtra extras, "pattern=" & p_LabelScalar(pattern)
    p_AddExtra extras, "mode=" & LCase$(Trim$(match_mode))

    If LCase$(Trim$(match_mode)) = "regex" Then
        If Not regex_ignore_case Then
            p_AddExtra extras, "ignore_case=False"
        End If
    End If

    If case_sensitive Then
        p_AddExtra extras, "case_sensitive=True"
    End If

    Dim noKeys() As Long

    p_WriteTransformColumn targetName, values, ctLong, _
        p_OpLabel("str.count", colName, noKeys, 0, extras)
End Sub
'==============================================================================
'
' STRING HELPER: str_count
'
' Counts matches in each non-NA value of a column. Returns a 1-based Long array.
'
' match_mode:
'   "literal"       Counts non-overlapping literal occurrences. Default.
'   "regex"         Counts regex matches using VBScript.RegExp.
'   "like"          Returns 1 if the entire string matches a VBA Like pattern.
'   "like_contains" Returns 1 if any substring matches a VBA Like pattern.
'
' Examples:
'   v = t.str_count("CustomerName", "LLC")
'   v = t.str_count("CustomerName", "\d+", match_mode:="regex")
'   v = t.str_count("CustomerName", "*LLC*", match_mode:="like")
'   v = t.str_count("CustomerName", "LLC", match_mode:="like_contains")
'
' Notes:
'   VBA Like is controlled by Option Compare for case sensitivity. To avoid
'   relying on the module setting, this implementation normalizes text and
'   pattern to lowercase when case_sensitive:=False.
'==============================================================================

Public Function str_count(ByVal colName As Variant, _
                          ByVal pattern As String, _
                          Optional ByVal match_mode As String = "literal", _
                          Optional ByVal compare As VbCompareMethod = vbTextCompare, _
                          Optional ByVal regex_ignore_case As Boolean = True, _
                          Optional ByVal case_sensitive As Boolean = False, _
                          Optional ByVal na_value As Long = 0) As Variant
    colName = p_LabelColumnRef(colName)

    Dim ci As Long
    ci = p_ColIdx(colName)

    If ci = 0 Then
        Err.Raise vbObjectError + 1540, "cTable.str_count", _
                  "Column not found: """ & colName & """"
    End If

    If Len(pattern) = 0 Then
        Err.Raise vbObjectError + 1541, "cTable.str_count", _
                  "pattern cannot be blank."
    End If

    If m_RowCount = 0 Then
        str_count = Array()
        Exit Function
    End If

    Dim mode As String
    mode = LCase$(Trim$(match_mode))

    Select Case mode
        Case "literal", "regex", "like", "like_contains", "likecontains"
            ' valid
        Case Else
            Err.Raise vbObjectError + 1542, "cTable.str_count", _
                      "match_mode must be ""literal"", ""regex"", ""like"", or ""like_contains""."
    End Select

    Dim Out() As Long
    ReDim Out(1 To m_RowCount)

    Dim re As Object

    If mode = "regex" Then
        Set re = CreateObject("VBScript.RegExp")
        re.pattern = pattern
        re.Global = True
        re.IgnoreCase = regex_ignore_case
        re.Multiline = False
    End If

    Dim r As Long
    Dim v As Variant
    Dim s As String

    For r = 1 To m_RowCount
        v = m_Data(r, ci)

        If p_IsNA(v) Then
            Out(r) = na_value
        Else
            s = CStr(v)

            Select Case mode
                Case "literal"
                    Out(r) = p_StrLiteralCount(s, pattern, compare)

                Case "regex"
                    Out(r) = p_StrRegexCount(s, re)

                Case "like"
                    If p_StrLikeMatch(s, pattern, case_sensitive) Then
                        Out(r) = 1
                    Else
                        Out(r) = 0
                    End If

                Case "like_contains", "likecontains"
                    If p_StrLikeContains(s, pattern, case_sensitive) Then
                        Out(r) = 1
                    Else
                        Out(r) = 0
                    End If
            End Select
        End If
    Next r

    str_count = Out
End Function
'------------------------------------------------------------------------------
' Count non-overlapping literal occurrences of findText in textValue.
'------------------------------------------------------------------------------
Private Function p_StrLiteralCount(ByVal textValue As String, _
                                   ByVal findText As String, _
                                   ByVal compare As VbCompareMethod) As Long

    Dim pos As Long
    Dim n As Long
    Dim stepLen As Long

    stepLen = Len(findText)
    pos = 1

    Do
        pos = InStr(pos, textValue, findText, compare)

        If pos = 0 Then Exit Do

        n = n + 1
        pos = pos + stepLen
    Loop

    p_StrLiteralCount = n
End Function


'------------------------------------------------------------------------------
' Count regex matches using a precompiled VBScript.RegExp object.
'------------------------------------------------------------------------------
Private Function p_StrRegexCount(ByVal textValue As String, _
                                 ByVal re As Object) As Long

    Dim matches As Object
    Set matches = re.Execute(textValue)

    p_StrRegexCount = matches.count
End Function


'------------------------------------------------------------------------------
' Full-string VBA Like match.
'
' If case_sensitive=False, both value and pattern are lowercased before Like.
' This avoids depending on Option Compare Binary/Text at the top of the module.
'------------------------------------------------------------------------------
Private Function p_StrLikeMatch(ByVal textValue As String, _
                                ByVal likePattern As String, _
                                ByVal case_sensitive As Boolean) As Boolean

    If case_sensitive Then
        p_StrLikeMatch = (textValue Like likePattern)
    Else
        p_StrLikeMatch = (LCase$(textValue) Like LCase$(likePattern))
    End If
End Function


'------------------------------------------------------------------------------
' Contains-style VBA Like match.
'
' This returns True when the text contains something matching likePattern.
'
' Because VBA Like is a full-string matcher, this wraps the pattern in leading
' and trailing "*" wildcards unless the caller already supplied them.
'
' Examples:
'   pattern "LLC"      matches text containing LLC
'   pattern "L?C"      matches LAC, LBC, LLC, etc. anywhere in the string
'   pattern "*LLC*"    works as supplied
'------------------------------------------------------------------------------
Private Function p_StrLikeContains(ByVal textValue As String, _
                                   ByVal likePattern As String, _
                                   ByVal case_sensitive As Boolean) As Boolean

    Dim pat As String
    pat = likePattern

    If Left$(pat, 1) <> "*" Then pat = "*" & pat
    If right$(pat, 1) <> "*" Then pat = pat & "*"

    p_StrLikeContains = p_StrLikeMatch(textValue, pat, case_sensitive)
End Function

Public Function CountTrue(ByVal mask As Variant) As Long
    Dim i As Long

    If Not IsArray(mask) Then Exit Function

    For i = LBound(mask) To UBound(mask)
        If CBool(mask(i)) Then CountTrue = CountTrue + 1
    Next i
End Function
'------------------------------------------------------------------------------
' read_csv
'
' Loads a CSV / delimited text file into cTable.
'
' filePath:
'   Full path to the CSV or delimited text file.
'
' delimiter:
'   Field delimiter. Default ",".
'
' hasHeaders:
'   True  = first parsed row supplies column names.
'   False = generated headers are used: Column1, Column2, etc.
'
' InferTypes:
'   True = auto-detect column dtypes through convert_dtypes.
'
' nrows:
'   0 = read all data records.
'   >0 = read only this number of DATA records. The header does not count.
'
' IncludeSourceRecord:
'   True  = append SourceRecordColumnName, containing the 1-based ordinal of
'           each parsed CSV data record.
'   False = do not append a source-record column.
'
' SourceRecordColumnName:
'   Defaults to "_SourceRecord".
'
' IMPORTANT:
'   _SourceRecord is a parsed-record ordinal, not the physical text-file line
'   number. Blank lines are skipped. Use a separate _SourceLine capability if
'   physical line numbers are ever required.
'
' Supports:
'   - UTF-8 files with or without BOM
'   - quoted fields
'   - escaped quotes inside quoted fields
'   - ragged rows padded with Null
'
' Does not support embedded line breaks inside quoted fields.
'------------------------------------------------------------------------------
Public Sub read_csv(ByVal filePath As String, _
                    Optional ByVal delimiter As String = ",", _
                    Optional ByVal hasHeaders As Boolean = True, _
                    Optional ByVal InferTypes As Boolean = True, _
                    Optional ByVal nrows As Long = 0, _
                    Optional ByVal IncludeSourceRecord As Boolean = False, _
                    Optional ByVal SourceRecordColumnName As String = "_SourceRecord")

    Const CALLER As String = "cTable.read_csv"

    p_ValidateNRows nrows, CALLER

    If Len(Trim$(filePath)) = 0 Then
        Err.Raise vbObjectError + 1500, CALLER, _
                  "filePath is required."
    End If

    If Len(delimiter) = 0 Then
        Err.Raise vbObjectError + 1501, CALLER, _
                  "delimiter cannot be blank."
    End If

    If Dir$(filePath) = "" Then
        Err.Raise vbObjectError + 1502, CALLER, _
                  "File not found: " & filePath
    End If

    If m_Log Is Nothing Then Set m_Log = New cLog
    m_Log.AddStatus "read_csv starting: " & filePath

    Dim rows As Collection
    Set rows = New Collection

    Dim maxCols As Long

    '-------------------------------------------------------------------------
    ' Read and parse the source.
    '
    ' For limited reads, stop after the requested number of parsed rows.
    ' When headers are present, one additional parsed row is needed.
    '-------------------------------------------------------------------------
    If nrows > 0 Then

        Dim maxParsedRows As Long
        maxParsedRows = nrows

        If hasHeaders Then
            maxParsedRows = maxParsedRows + 1
        End If

        Set rows = p_ReadDelimitedRowsLimited( _
                        filePath:=filePath, _
                        delimiter:=delimiter, _
                        maxParsedRows:=maxParsedRows, _
                        maxCols:=maxCols)

    Else

        Dim fileText As String
        fileText = p_ReadTextFile(filePath, "utf-8")
        fileText = p_StripTextBOM(fileText)

        ' Normalize all recognized line endings to LF.
        fileText = Replace(fileText, vbCrLf, vbLf)
        fileText = Replace(fileText, vbCr, vbLf)

        If Len(fileText) = 0 Then
            p_ResetEmptyTableForRead
            m_Log.AddStatus "read_csv found an empty file."
            Exit Sub
        End If

        Dim lines As Variant
        lines = Split(fileText, vbLf)

        Dim lineIndex As Long
        Dim lineText As String
        Dim fields As Variant
        Dim fieldCount As Long

        For lineIndex = LBound(lines) To UBound(lines)

            lineText = CStr(lines(lineIndex))

            If lineIndex = LBound(lines) Then
                lineText = p_StripTextBOM(lineText)
            End If

            ' Match existing behavior: ignore completely blank physical lines.
            If Len(lineText) > 0 Then

                fields = p_ParseDelimitedLine(lineText, delimiter)

                ' Defensive BOM cleanup on the first field of the first
                ' accepted parsed row.
                If rows.count = 0 Then
                    If p_ArrayHasElements(fields) Then
                        fields(LBound(fields)) = _
                            p_StripTextBOM(CStr(fields(LBound(fields))))
                    End If
                End If

                rows.Add fields

                If p_ArrayHasElements(fields) Then
                    fieldCount = UBound(fields) - LBound(fields) + 1

                    If fieldCount > maxCols Then
                        maxCols = fieldCount
                    End If
                End If
            End If
        Next lineIndex
    End If

    If rows.count = 0 Or maxCols = 0 Then
        p_ResetEmptyTableForRead
        m_Log.AddStatus _
            "read_csv parsed no rows. Check delimiter and file content."
        Exit Sub
    End If

    '-------------------------------------------------------------------------
    ' Establish headers and determine how many accepted data records exist.
    '-------------------------------------------------------------------------
    Dim c As Long
    Dim dataStart As Long
    Dim headerRow As Variant

    m_ColCount = maxCols
    ReDim m_Headers(1 To m_ColCount)

    If hasHeaders Then

        headerRow = rows(1)

        For c = 1 To m_ColCount
            If c - 1 <= UBound(headerRow) Then
                m_Headers(c) = CStr(headerRow(c - 1))
            Else
                m_Headers(c) = "Column" & c
            End If
        Next c

        m_Headers(1) = p_StripTextBOM(m_Headers(1))
        p_SanitizeHeaders m_Headers, m_ColCount

        dataStart = 2
        m_RowCount = rows.count - 1

    Else

        For c = 1 To m_ColCount
            m_Headers(c) = "Column" & c
        Next c

        dataStart = 1
        m_RowCount = rows.count
    End If

    m_TypesAllocated = False

    '-------------------------------------------------------------------------
    ' Header-only file.
    '-------------------------------------------------------------------------
    If m_RowCount <= 0 Then
        m_RowCount = 0
        Erase m_Data

        p_ClearIndexes

        If InferTypes Then Me.convert_dtypes

        p_SetCsvDataSource filePath
        p_LinAdoptSource

        m_Log.AddStatus _
            "read_csv loaded headers only. Rows: 0, Columns: " & m_ColCount

        Exit Sub
    End If

    '-------------------------------------------------------------------------
    ' Materialize data values.
    '-------------------------------------------------------------------------
    ReDim m_Data(1 To m_RowCount, 1 To m_ColCount)

    Dim r As Long
    Dim srcRow As Long
    Dim rowFields As Variant
    Dim rawValue As String

    For r = 1 To m_RowCount

        srcRow = dataStart + r - 1
        rowFields = rows(srcRow)

        For c = 1 To m_ColCount

            If c - 1 <= UBound(rowFields) Then

                rawValue = CStr(rowFields(c - 1))

                If Len(rawValue) = 0 Then
                    m_Data(r, c) = Null
                Else
                    m_Data(r, c) = rawValue
                End If

            Else
                m_Data(r, c) = Null
            End If
        Next c
    Next r

    p_ClearIndexes

    If InferTypes Then Me.convert_dtypes

    '-------------------------------------------------------------------------
    ' Source metadata and lineage must be established before adding the
    ' generated source-record field.
    '-------------------------------------------------------------------------
    p_SetCsvDataSource filePath
    p_LinAdoptSource

    '-------------------------------------------------------------------------
    ' The first loaded data record is always record 1, whether the source had
    ' a header row or not. The header is schema, not a data record.
    '-------------------------------------------------------------------------
    If IncludeSourceRecord Then
        p_AddCsvSourceRecordColumn SourceRecordColumnName
    End If

    m_Log.AddStatus _
        "read_csv loaded file '" & filePath & "'. Rows: " & _
        m_RowCount & ", Columns: " & m_ColCount & _
        ", Headers: " & hasHeaders & _
        ", SourceRecord: " & IncludeSourceRecord

End Sub


' ---------------------------------------------------------------------------
' p_ReadTextFile
'   Reads a text file with ADODB.Stream using the requested charset.
'   UTF-8 decoding prevents BOM bytes from appearing as "      ".
' ---------------------------------------------------------------------------
Private Function p_ReadTextFile(ByVal filePath As String, _
                                Optional ByVal Charset As String = "utf-8") As String

    Const CALLER As String = "cTable.p_ReadTextFile"

    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")

    On Error GoTo ReadFail

    With stm
        .Type = 2                 ' adTypeText
        .Charset = Charset
        .Open
        .LoadFromFile filePath
        p_ReadTextFile = .ReadText(-1)
        .Close
    End With

    Set stm = Nothing
    Exit Function

ReadFail:
    On Error Resume Next
    If Not stm Is Nothing Then
        If stm.State = 1 Then stm.Close
    End If
    Set stm = Nothing
    On Error GoTo 0

    Err.Raise vbObjectError + 1504, CALLER, _
              "Error reading text file '" & filePath & "': " & Err.Description
End Function
Private Function p_ReadDelimitedRowsLimited( _
        ByVal filePath As String, _
        ByVal delimiter As String, _
        ByVal maxParsedRows As Long, _
        ByRef maxCols As Long) As Collection

    Const CALLER As String = "cTable.p_ReadDelimitedRowsLimited"

    Dim Out As Collection
    Set Out = New Collection

    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")

    On Error GoTo ReadFail

    With stm
        .Type = 2                 ' adTypeText
        .Charset = "utf-8"
        .Open
        .LoadFromFile filePath

        ' Read a small prefix only to determine whether the file is LF/CRLF
        ' or old-style CR delimited. Then rewind before the real read.
        Dim probe As String
        probe = .ReadText(65536)

        If InStr(1, probe, vbLf, vbBinaryCompare) > 0 Then
            .LineSeparator = 10   ' adLF; CRLF leaves a trailing CR to trim
        ElseIf InStr(1, probe, vbCr, vbBinaryCompare) > 0 Then
            .LineSeparator = 13   ' adCR
        Else
            .LineSeparator = 10
        End If

        .position = 0

        Dim firstPhysicalLine As Boolean
        firstPhysicalLine = True

        Dim lineText As String
        Dim fields As Variant
        Dim fieldCount As Long

        Do While Not .EOS And Out.count < maxParsedRows

            lineText = .ReadText(-2)    ' adReadLine

            If firstPhysicalLine Then
                lineText = p_StripTextBOM(lineText)
                firstPhysicalLine = False
            End If

            ' When LF is the selected separator, a CRLF file leaves CR at
            ' the end of the returned text. Remove it before parsing.
            If Len(lineText) > 0 Then
                If right$(lineText, 1) = vbCr Then
                    lineText = Left$(lineText, Len(lineText) - 1)
                End If
            End If

            ' Match the full-file reader: completely blank lines are skipped.
            If Len(lineText) > 0 Then

                fields = p_ParseDelimitedLine(lineText, delimiter)

                If Out.count = 0 Then
                    If p_ArrayHasElements(fields) Then
                        fields(LBound(fields)) = _
                            p_StripTextBOM(CStr(fields(LBound(fields))))
                    End If
                End If

                Out.Add fields

                If p_ArrayHasElements(fields) Then
                    fieldCount = UBound(fields) - LBound(fields) + 1
                    If fieldCount > maxCols Then maxCols = fieldCount
                End If

            End If

        Loop

        .Close
    End With

    Set stm = Nothing
    Set p_ReadDelimitedRowsLimited = Out
    Exit Function

ReadFail:
    Dim errDesc As String
    errDesc = Err.Description

    On Error Resume Next
    If Not stm Is Nothing Then
        If stm.State = 1 Then stm.Close
    End If
    Set stm = Nothing
    On Error GoTo 0

    Err.Raise vbObjectError + 1505, CALLER, _
              "Error reading limited text rows from '" & _
              filePath & "': " & errDesc

End Function

' ---------------------------------------------------------------------------
' p_StripTextBOM
'   Removes common BOM / marker characters from the start of a string.
'
' Handles:
'   UTF-8 BOM decoded correctly      ChrW$(&HFEFF)
'   UTF-8 BOM misread as ANSI        "      "
' ---------------------------------------------------------------------------
Private Function p_StripTextBOM(ByVal s As String) As String

    If Len(s) = 0 Then
        p_StripTextBOM = s
        Exit Function
    End If

    If Left$(s, 1) = ChrW$(&HFEFF) Then
        s = mid$(s, 2)
    End If

    If Len(s) >= 3 Then
        If Left$(s, 3) = "      " Then
            s = mid$(s, 4)
        End If
    End If

    p_StripTextBOM = s

End Function
' ---------------------------------------------------------------------------
' p_ParseDelimitedLine
'   Parses one delimited line into a 0-based Variant array.
'
' Supports:
'   abc,def
'   "abc","def"
'   "a,b","def"
'   "He said ""yes""",def
' ---------------------------------------------------------------------------
Private Function p_ParseDelimitedLine(ByVal lineText As String, _
                                      ByVal delimiter As String) As Variant

    Dim fields As Collection
    Set fields = New Collection

    Dim i As Long
    Dim ch As String
    Dim token As String
    Dim inQuotes As Boolean
    Dim delimLen As Long

    delimLen = Len(delimiter)
    i = 1

    Do While i <= Len(lineText)

        ch = mid$(lineText, i, 1)

        If ch = """" Then

            If inQuotes Then

                If i < Len(lineText) Then
                    If mid$(lineText, i + 1, 1) = """" Then
                        token = token & """"
                        i = i + 2
                    Else
                        inQuotes = False
                        i = i + 1
                    End If
                Else
                    inQuotes = False
                    i = i + 1
                End If

            Else
                inQuotes = True
                i = i + 1
            End If

        ElseIf Not inQuotes _
           And delimLen > 0 _
           And mid$(lineText, i, delimLen) = delimiter Then

            fields.Add token
            token = vbNullString
            i = i + delimLen

        Else

            token = token & ch
            i = i + 1

        End If

    Loop

    fields.Add token

    Dim Out() As Variant
    ReDim Out(0 To fields.count - 1)

    Dim n As Long
    For n = 1 To fields.count
        Out(n - 1) = fields(n)
    Next n

    p_ParseDelimitedLine = Out

End Function

' ---------------------------------------------------------------------------
' p_ResetEmptyTableForRead
'   Clears this cTable after a read method determines there is no data.
' ---------------------------------------------------------------------------
Private Sub p_ResetEmptyTableForRead()

    m_RowCount = 0
    m_ColCount = 0

    Erase m_Data
    Erase m_Headers

    m_TypesAllocated = False

    If Not m_Lin Is Nothing Then
        m_Lin.InitUniform 0, 0, ctSrcUnknown
    End If

    p_ClearIndexes

End Sub

' ---------------------------------------------------------------------------
' p_SetCsvDataSource
'   CSV provenance using file system metadata only.
'
'   Does NOT call cDataSource.InitFromPath because that opens the path as an
'   Excel workbook. For .csv files, that visibly opens a separate Excel window.
' ---------------------------------------------------------------------------
Private Sub p_SetCsvDataSource(ByVal filePath As String)

    Set m_DataSource = New cDataSource
    m_DataSource.InitFromCSV filePath

End Sub

'/////////////////// Add read_json
' ---------------------------------------------------------------------------
' read_json
'   Loads JSON data into cTable.
'
'   filePath       Full path to the JSON file, OR an http/https URL such as a
'                  SharePoint document URL. The two are distinguished by
'                  p_IsHttpUrl; URLs are fetched with p_ReadTextFromUrl.
'
'   authType       Only used when filePath is a URL. One of:
'                    "xmlhttp"        MSXML2.XMLHTTP (default)
'                    "serverxmlhttp"  MSXML2.ServerXMLHTTP.6.0
'                    "winhttp"        WinHttp.WinHttpRequest.5.1, with
'                                     SetAutoLogonPolicy 0 for integrated auth
'                    "clientcert"     WinHttp with SetClientCertificate
'                    "bearer"         XMLHTTP with an Authorization header
'
'   authValue      Certificate location for "clientcert", or the token text
'                  for "bearer". Ignored otherwise.
'
'   recordPath     Optional path to the array of records inside a top-level
'                  object. Examples:
'                    ""              top-level JSON is the records array
'                    "records"       root("records")
'                    "data.items"    root("data")("items")
'
'   InferTypes     True auto-detects dtypes via convert_dtypes.
'
'   flattenNested  False:
'                    Nested objects/arrays are serialized back to JSON text.
'
'                  True:
'                    Nested dictionaries are flattened using dot names.
'                    Example:
'                      {"person":{"name":"Ann"}}
'                    becomes:
'                      person.name = Ann
'
'                    Nested arrays are still serialized as JSON text.
'
'   missingAsNull  True:
'                    Missing properties in a record become Null.
'   if nrows = 0 then return all rows otherwise return nrows of data (excluding header)
'
' Requirements:
'   Requires JsonConverter.bas, with JsonConverter.ParseJson available.
'
' Supported JSON shape:
'   [
'       {"A": 1, "B": "x"},
'       {"A": 2, "B": "y"}
'   ]
'
'   or, with recordPath:="records":
'
'   {
'       "records": [
'           {"A": 1, "B": "x"},
'           {"A": 2, "B": "y"}
'       ]
'   }
' ---------------------------------------------------------------------------
Public Sub read_json(ByVal filePath As String, _
                     Optional ByVal recordPath As String = "", _
                     Optional ByVal InferTypes As Boolean = True, _
                     Optional ByVal flattenNested As Boolean = False, _
                     Optional ByVal missingAsNull As Boolean = True, _
                     Optional ByVal nrows As Long = 0, _
                     Optional ByVal authType As String = "xmlhttp", _
                     Optional ByVal authValue As String = "")

    Const CALLER As String = "cTable.read_json"

    p_ValidateNRows nrows, CALLER

    ' filePath is ByVal, so trimming here is local to this call.
    filePath = Trim$(filePath)

    If Len(filePath) = 0 Then
        Err.Raise vbObjectError + 1600, CALLER, "filePath is required."
    End If

    Dim isRemote As Boolean
    isRemote = p_IsHttpUrl(filePath)

    ' Dir$ raises run-time error 52 ("Bad file name or number") when handed an
    ' http/https URL, so the existence test is only meaningful for local and
    ' UNC paths.
    If Not isRemote Then
        If Dir$(filePath) = "" Then
            Err.Raise vbObjectError + 1601, CALLER, _
                      "File not found: " & filePath
        End If
    End If

    If m_Log Is Nothing Then Set m_Log = New cLog
    m_Log.AddStatus "read_json starting: " & filePath

    Dim jsonText As String
    jsonText = p_ReadTextFromPathOrUrl(filePath, "utf-8", authType, authValue)
    jsonText = p_StripTextBOM(jsonText)

    ' A SharePoint URL that needs interactive sign-in, or a sharing link that
    ' points at the web viewer rather than the file, returns an HTML page with
    ' HTTP 200. Detect that here so the failure is reported as a retrieval
    ' problem instead of a confusing JSON parse error.
    If isRemote Then
        p_JsonGuardHtmlResponse jsonText, filePath, CALLER
    End If

    If Len(Trim$(jsonText)) = 0 Then
        p_ResetEmptyTableForRead
        m_Log.AddStatus "read_json found an empty file."
        Exit Sub
    End If

    Dim root As Variant

    On Error GoTo ParseErr
    Set root = JsonConverter.ParseJson(jsonText)
    On Error GoTo 0

    Dim records As Collection
    Set records = p_JsonResolveRecords(root, recordPath, CALLER)

    If records Is Nothing Then
        p_ResetEmptyTableForRead
        m_Log.AddStatus "read_json parsed no records."
        Exit Sub
    End If

    If records.count = 0 Then
        p_ResetEmptyTableForRead
        m_Log.AddStatus "read_json parsed an empty records array."
        Exit Sub
    End If

    ' JSON is still read and ParseJson is still run against the entire file.
    ' nrows limits only the records normalized/materialized into this cTable.
    Dim recordLimit As Long
    recordLimit = records.count

    If nrows > 0 Then
        If nrows < recordLimit Then recordLimit = nrows
    End If

    ' --------------------------------------------------------------------
    ' First pass:
    '   Normalize each JSON record to a Scripting.Dictionary and collect
    '   the union of all field names, preserving first appearance order.
    ' --------------------------------------------------------------------
    Dim normalized As Collection
    Set normalized = New Collection

    Dim headerSeen As Object
    Set headerSeen = CreateObject("Scripting.Dictionary")
    headerSeen.CompareMode = vbTextCompare

    Dim headerOrder As Collection
    Set headerOrder = New Collection

    Dim i As Long
    Dim rec As Variant
    Dim rowDict As Object
    Dim k As Variant

    For i = 1 To recordLimit

        p_AssignJsonVariant rec, records(i)

        Set rowDict = p_JsonRecordToDictionary( _
                        rec, _
                        flattenNested, _
                        CALLER)

        normalized.Add rowDict

        For Each k In rowDict.keys
            If Not headerSeen.Exists(CStr(k)) Then
                headerSeen.Add CStr(k), True
                headerOrder.Add CStr(k)
            End If
        Next k

    Next i

    If headerOrder.count = 0 Then
        p_ResetEmptyTableForRead
        m_Log.AddStatus "read_json records had no fields."
        Exit Sub
    End If

    ' --------------------------------------------------------------------
    ' Headers
    ' --------------------------------------------------------------------
    m_ColCount = headerOrder.count
    m_RowCount = normalized.count
    m_TypesAllocated = False

    ReDim m_Headers(1 To m_ColCount)

    Dim c As Long
    For c = 1 To m_ColCount
        m_Headers(c) = CStr(headerOrder(c))
    Next c

    p_SanitizeHeaders m_Headers, m_ColCount

    ' --------------------------------------------------------------------
    ' Data
    ' --------------------------------------------------------------------
    ReDim m_Data(1 To m_RowCount, 1 To m_ColCount)

    Dim r As Long
    Dim fieldName As String

    For r = 1 To m_RowCount

        Set rowDict = normalized(r)

        For c = 1 To m_ColCount

            fieldName = CStr(headerOrder(c))

            If rowDict.Exists(fieldName) Then
                m_Data(r, c) = rowDict(fieldName)
            Else
                If missingAsNull Then
                    m_Data(r, c) = Null
                Else
                    m_Data(r, c) = vbNullString
                End If
            End If

        Next c

    Next r

    p_ClearIndexes

    If InferTypes Then Me.convert_dtypes

    p_SetJsonDataSource filePath
    p_LinAdoptSource

    m_Log.AddStatus "read_json loaded " & _
                    IIf(isRemote, "URL '", "file '") & filePath & "'. Rows: " & _
                    m_RowCount & ", Columns: " & m_ColCount & _
                    ", recordPath: " & IIf(Len(recordPath) = 0, "(root)", recordPath)

    Exit Sub

ParseErr:
    Err.Raise vbObjectError + 1602, CALLER, _
              "Error parsing JSON from '" & filePath & "': " & Err.Description

End Sub


' ---------------------------------------------------------------------------
' p_SetJsonDataSource
'   JSON provenance using file system metadata only.
'   Does not open the file in Excel.
' ---------------------------------------------------------------------------
Private Sub p_SetJsonDataSource(ByVal filePath As String)

    Set m_DataSource = New cDataSource

    If p_IsHttpUrl(filePath) Then

        ' cDataSource.InitFromJSON reads file system metadata, which does not
        ' exist for an http/https source. Provenance is best effort here: if it
        ' cannot be recorded, the read still succeeds and lineage falls back to
        ' ctSrcUnknown via p_LinSrcId.
        On Error Resume Next
        m_DataSource.InitFromJSON filePath
        If Err.Number <> 0 Then
            Err.Clear
            Set m_DataSource = Nothing
            If m_Log Is Nothing Then Set m_Log = New cLog
            m_Log.AddStatus "read_json could not record file provenance for a " & _
                            "URL source. Lineage source is Unknown."
        End If
        On Error GoTo 0

    Else
        m_DataSource.InitFromJSON filePath
    End If

End Sub

' ---------------------------------------------------------------------------
' p_JsonGuardHtmlResponse
'   Raises a descriptive error when a URL fetch returned an HTML document
'   instead of JSON.
'
'   SharePoint does this in two common situations:
'     1. The request was not authenticated, so an ADFS / Microsoft sign-in
'        page came back with HTTP status 200.
'     2. The URL is a sharing link or a web-viewer link rather than a direct
'        link to the file contents.
' ---------------------------------------------------------------------------
Private Sub p_JsonGuardHtmlResponse(ByVal jsonText As String, _
                                    ByVal url As String, _
                                    ByVal callerName As String)

    Dim probe As String
    probe = LTrim$(jsonText)

    If Len(probe) > 512 Then probe = Left$(probe, 512)
    probe = LCase$(probe)

    If Left$(probe, 1) = "{" Or Left$(probe, 1) = "[" Then Exit Sub

    If Left$(probe, 9) = "<!doctype" Or _
       Left$(probe, 5) = "<html" Or _
       InStr(1, probe, "<html", vbTextCompare) > 0 Then

        Err.Raise vbObjectError + 1614, callerName, _
                  "The URL returned an HTML page rather than JSON. This is " & _
                  "usually a SharePoint sign-in page, or a sharing/viewer " & _
                  "link instead of a direct file link. Try the direct " & _
                  "document URL, or append ""?download=1"", or use " & _
                  "authType:=""winhttp"". URL: " & url
    End If

End Sub
' ---------------------------------------------------------------------------
' p_JsonResolveRecords
'   Returns a Collection of record objects to load.
'
'   Supports:
'       ""                              top-level JSON must be an array
'       "company.departments"           object -> object -> array
'       "company.products"              object -> object -> array
'       "company.departments.employees" object -> object -> array -> array
'
'   When the path passes through a JSON array, this walks every item in that
'   array and continues resolving the remaining path. This makes nested record
'   paths useful for common API-shaped JSON.
' ---------------------------------------------------------------------------
Private Function p_JsonResolveRecords(ByVal root As Variant, _
                                      ByVal recordPath As String, _
                                      ByVal callerName As String) As Collection

    Dim Out As Collection
    Set Out = New Collection

    If Len(Trim$(recordPath)) = 0 Then

        If IsObject(root) Then
            If typeName(root) = "Collection" Then
                Set p_JsonResolveRecords = root
                Exit Function
            End If
        End If

        Err.Raise vbObjectError + 1603, callerName, _
                  "Top-level JSON must be an array of records when " & _
                  "recordPath is omitted. For this file, try " & _
                  """company.departments"", ""company.products"", or " & _
                  """company.departments.employees""."
    End If

    Dim parts As Variant
    parts = Split(recordPath, ".")

    Dim startNode As Variant
    p_AssignJsonVariant startNode, root

    p_JsonCollectPathMatches startNode, parts, LBound(parts), Out, callerName, recordPath

    If Out.count = 0 Then
        Err.Raise vbObjectError + 1608, callerName, _
                  "recordPath """ & recordPath & _
                  """ resolved to no records."
    End If

    Set p_JsonResolveRecords = Out

End Function

' ---------------------------------------------------------------------------
' p_JsonCollectPathMatches
'   Recursive path walker for p_JsonResolveRecords.
'
'   If node is a Dictionary:
'       resolve the current path segment as a property name.
'
'   If node is a Collection:
'       apply the same path segment to every item in the array.
'
'   When all path segments have been consumed:
'       if node is a Collection, add every item as one record;
'       otherwise add node itself as one record.
' ---------------------------------------------------------------------------
Private Sub p_JsonCollectPathMatches(ByVal node As Variant, _
                                     ByVal parts As Variant, _
                                     ByVal partIndex As Long, _
                                     ByVal Out As Collection, _
                                     ByVal callerName As String, _
                                     ByVal fullpath As String)

    Dim item As Variant
    Dim d As Object
    Dim keyName As String

    ' Path fully consumed. The current node is the record set.
    If partIndex > UBound(parts) Then

        If IsObject(node) Then
            If typeName(node) = "Collection" Then
                For Each item In node
                    Out.Add item
                Next item
            Else
                Out.Add node
            End If
        Else
            Out.Add node
        End If

        Exit Sub
    End If

    keyName = Trim$(CStr(parts(partIndex)))

    If Len(keyName) = 0 Then
        Err.Raise vbObjectError + 1604, callerName, _
                  "recordPath contains a blank segment: """ & fullpath & """"
    End If

    If Not IsObject(node) Then
        ' A scalar cannot be walked further.
        Exit Sub
    End If

    Select Case typeName(node)

        Case "Dictionary"
            Set d = node

            If d.Exists(keyName) Then
                Dim nextNode As Variant
                p_AssignJsonVariant nextNode, d(keyName)

                p_JsonCollectPathMatches nextNode, parts, partIndex + 1, _
                                         Out, callerName, fullpath
            End If


        Case "Collection"
            Dim nextItem As Variant

            For Each item In node
                p_AssignJsonVariant nextItem, item

                p_JsonCollectPathMatches nextItem, parts, partIndex, _
                                         Out, callerName, fullpath
            Next item


        Case Else
            ' Unknown object type. Cannot walk further.

    End Select

End Sub
' ---------------------------------------------------------------------------
' p_AssignJsonVariant
'   Assigns either an object JSON node or scalar JSON value into a Variant.
'
'   VBA requires Set when the value is an object. Without this helper, lines
'   like:
'       rec = records(i)
'   can raise run-time error 450 when records(i) is a Dictionary.
' ---------------------------------------------------------------------------
Private Sub p_AssignJsonVariant(ByRef target As Variant, ByVal source As Variant)

    If IsObject(source) Then
        Set target = source
    Else
        target = source
    End If

End Sub

' ---------------------------------------------------------------------------
' p_JsonRecordToDictionary
'   Converts one JSON record into a Dictionary.
'
'   Accepted record shapes:
'       JSON object      -> fields from the object
'       scalar value     -> one column named "value"
'
'   If flattenNested=True, nested objects become dotted columns.
'   Nested arrays are serialized as JSON text.
' ---------------------------------------------------------------------------
Private Function p_JsonRecordToDictionary(ByVal rec As Variant, _
                                          ByVal flattenNested As Boolean, _
                                          ByVal callerName As String) As Object

    Dim Out As Object
    Set Out = CreateObject("Scripting.Dictionary")
    Out.CompareMode = vbTextCompare

    If IsObject(rec) Then

        If typeName(rec) = "Dictionary" Then

            If flattenNested Then
                p_JsonFlattenDictionary rec, Out, vbNullString
            Else
                Dim k As Variant
                Dim v As Variant

                For Each k In rec.keys
                    p_AssignJsonVariant v, rec(k)
                    Out(CStr(k)) = p_JsonValueToCell(v)
                Next k
            End If

        Else
            ' Unexpected object type. Store as text.
            Out("value") = p_JsonValueToCell(rec)
        End If

    Else
        Out("value") = p_JsonValueToCell(rec)
    End If

    Set p_JsonRecordToDictionary = Out

End Function

' ---------------------------------------------------------------------------
' p_JsonFlattenDictionary
'   Recursively flattens JSON dictionaries into dotted column names.
'
'   Example:
'       {"person":{"name":"Ann","office":"HQ"}}
'
'   Output:
'       person.name
'       person.office
'
'   Arrays remain JSON text because exploding arrays creates multiple possible
'   table shapes and should be a separate method.
' ---------------------------------------------------------------------------
Private Sub p_JsonFlattenDictionary(ByVal src As Object, _
                                    ByVal Out As Object, _
                                    ByVal prefix As String)

    Dim k As Variant
    Dim fullName As String
    Dim v As Variant

    For Each k In src.keys

        If Len(prefix) = 0 Then
            fullName = CStr(k)
        Else
            fullName = prefix & "." & CStr(k)
        End If

        p_AssignJsonVariant v, src(k)

        If IsObject(v) Then

            If typeName(v) = "Dictionary" Then
                p_JsonFlattenDictionary v, Out, fullName
            Else
                Out(fullName) = p_JsonValueToCell(v)
            End If

        Else
            Out(fullName) = p_JsonValueToCell(v)
        End If

    Next k

End Sub


' ---------------------------------------------------------------------------
' p_JsonValueToCell
'   Converts a parsed JSON value into a cTable cell value.
'
'   Scalars become scalars.
'   Null / Empty become Null.
'   JSON objects and arrays become compact JSON text.
' ---------------------------------------------------------------------------
Private Function p_JsonValueToCell(ByVal v As Variant) As Variant

    If IsObject(v) Then
        p_JsonValueToCell = JsonConverter.ConvertToJson(v, Whitespace:=0)
        Exit Function
    End If

    If IsNull(v) Or IsEmpty(v) Then
        p_JsonValueToCell = Null
    Else
        p_JsonValueToCell = v
    End If

End Function
'///////////////////Add for SharePoint file access

' ---------------------------------------------------------------------------
' p_IsHttpUrl
'   True when the supplied path is an HTTP/HTTPS URL.
' ---------------------------------------------------------------------------
Private Function p_IsHttpUrl(ByVal pathOrUrl As String) As Boolean

    Dim s As String
    s = LCase$(Trim$(pathOrUrl))

    p_IsHttpUrl = _
        (Left$(s, 7) = "http://") Or _
        (Left$(s, 8) = "https://")

End Function
' ---------------------------------------------------------------------------
' p_IsSharePointUrl
'   True when the supplied path appears to be a SharePoint URL.
'
' Note:
'   This intentionally detects common SharePoint Online URLs only.
'   If you also use other internal URL patterns, extend this test.
' ---------------------------------------------------------------------------
Private Function p_IsSharePointUrl(ByVal pathOrUrl As String) As Boolean

    Dim s As String
    s = LCase$(Trim$(pathOrUrl))

    p_IsSharePointUrl = _
        p_IsHttpUrl(s) And _
        (InStr(1, s, "sharepoint.com", vbTextCompare) > 0)

End Function
' ---------------------------------------------------------------------------
' p_ReadTextFromPathOrUrl
'   Reads text from either:
'       - local / network file path
'       - SharePoint / HTTP URL
'
' For local files, uses existing p_ReadTextFile.
' For HTTP URLs, uses p_ReadTextFromUrl.
' ---------------------------------------------------------------------------
Private Function p_ReadTextFromPathOrUrl(ByVal pathOrUrl As String, _
                                         Optional ByVal Charset As String = "utf-8", _
                                         Optional ByVal authType As String = "xmlhttp", _
                                         Optional ByVal authValue As String = "") As String

    Const CALLER As String = "cTable.p_ReadTextFromPathOrUrl"

    Dim s As String
    s = Trim$(pathOrUrl)

    If Len(s) = 0 Then
        Err.Raise vbObjectError + 1610, CALLER, _
                  "pathOrUrl is required."
    End If

    If p_IsHttpUrl(s) Then
        p_ReadTextFromPathOrUrl = p_ReadTextFromUrl(s, authType, authValue)
    Else
        If Dir$(s) = "" Then
            Err.Raise vbObjectError + 1611, CALLER, _
                      "File not found: " & s
        End If

        p_ReadTextFromPathOrUrl = p_ReadTextFile(s, Charset)
    End If

End Function
' ---------------------------------------------------------------------------
' p_ReadTextFromUrl
'   Downloads text from an HTTP/HTTPS URL.
'
' Intended for SharePoint-hosted JSON/CSV files where the user already has
' access and the URL can be fetched by MSXML2.XMLHTTP.
'
' Caveats:
'   - Does not perform explicit Microsoft 365 OAuth.
'   - May fail with 401/403 if the current Windows/Office session is not
'     usable by XMLHTTP.
'   - May receive an HTML sign-in page instead of JSON if SharePoint redirects.
' ---------------------------------------------------------------------------
Private Function p_ReadTextFromUrl(ByVal url As String, _
                                   Optional ByVal authType As String = "xmlhttp", _
                                   Optional ByVal authValue As String = "") As String

    Const CALLER As String = "cTable.p_ReadTextFromUrl"

    Dim mode As String
    mode = LCase$(Trim$(authType))

    Dim http As Object
    Dim usedCert As Boolean

    ' Argument validation and object creation happen before the error handler
    ' is armed, so a bad authType reports its own message rather than being
    ' rewrapped as a download failure.
    Select Case mode

        Case "", "none", "xmlhttp", "sharepoint"
            Set http = CreateObject("MSXML2.XMLHTTP")

        Case "serverxmlhttp"
            Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")

        Case "winhttp", "default"
            Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

        Case "clientcert", "piv", "pki"
            If Len(Trim$(authValue)) = 0 Then
                Err.Raise vbObjectError + 1615, CALLER, _
                          "authValue is required when authType=""clientcert""."
            End If
            Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
            usedCert = True

        Case "bearer"
            If Len(Trim$(authValue)) = 0 Then
                Err.Raise vbObjectError + 1616, CALLER, _
                          "authValue is required when authType=""bearer""."
            End If
            Set http = CreateObject("MSXML2.XMLHTTP")

        Case Else
            Err.Raise vbObjectError + 1617, CALLER, _
                      "Unsupported authType: """ & authType & """. " & _
                      "Use ""xmlhttp"", ""serverxmlhttp"", ""winhttp"", " & _
                      """clientcert"", ""piv"", ""pki"", or ""bearer""."

    End Select

    On Error GoTo HttpFail

    http.Open "GET", url, False
    http.setRequestHeader "Accept", "application/json, text/plain, */*"

    If mode = "bearer" Then
        http.setRequestHeader "Authorization", "Bearer " & authValue
    End If

    If usedCert Then
        http.SetClientCertificate authValue
    End If

    ' WinHttp only. Allows the current Windows credentials to be sent
    ' automatically, which is what makes integrated SharePoint auth work.
    If mode = "winhttp" Or mode = "default" Then
        On Error Resume Next
        http.SetAutoLogonPolicy 0
        On Error GoTo HttpFail
    End If

    http.send

    If http.Status >= 200 And http.Status < 300 Then
        p_ReadTextFromUrl = CStr(http.responseText)
        Exit Function
    End If

    Err.Raise vbObjectError + 1612, CALLER, _
              "Failed to download URL. HTTP status: " & http.Status & _
              IIf(Len(http.statusText) > 0, " " & http.statusText, "") & _
              ". authType: " & mode

    Exit Function

HttpFail:
    Err.Raise vbObjectError + 1613, CALLER, _
              "Error downloading URL '" & url & "': " & Err.Description

End Function

'////////////////// read_text test

' ---------------------------------------------------------------------------
' read_text
'   TEST METHOD.
'
'   Reads text from either a local path or an HTTP/SharePoint URL and loads it
'   into this cTable as a one-row, one-column table.
'
'   Intended to test SharePoint / URL retrieval before wiring the same helper
'   into read_json or read_csv.
'
' Result:
'   Column: text
'   Row 1 : full downloaded/read text
'
' Usage:
'   t.read_text "C:\Temp\sample.txt"
'
'   t.read_text "https://example.com/data/sample.txt"
'
'   Debug.Print t.GetValueByName(1, "text")
' ---------------------------------------------------------------------------
Public Sub read_text(ByVal filePath As String, _
                     Optional ByVal authType As String = "xmlhttp", _
                     Optional ByVal authValue As String = "")

    Const CALLER As String = "cTable.read_text"

    If Len(Trim$(filePath)) = 0 Then
        Err.Raise vbObjectError + 1700, CALLER, "filePath is required."
    End If

    If m_Log Is Nothing Then Set m_Log = New cLog
    m_Log.AddStatus "read_text starting: " & filePath

    Dim textValue As String
    textValue = p_testReadTextFromPathOrUrl(filePath, authType, authValue)
    textValue = p_StripTextBOM(textValue)

    Dim hdrs(1 To 1) As String
    hdrs(1) = "text"

    Dim data(1 To 1, 1 To 1) As Variant
    data(1, 1) = textValue

    Me.from_array hdrs, data, 1, 1

    p_EnsureTypes
    m_ColTypes(1) = ctString

    If m_Log Is Nothing Then Set m_Log = New cLog
    m_Log.AddStatus "read_text loaded source. Characters: " & Len(textValue)

End Sub
' ---------------------------------------------------------------------------
' p_testIsHttpUrl
'   TEST HELPER.
' ---------------------------------------------------------------------------
Private Function p_testIsHttpUrl(ByVal pathOrUrl As String) As Boolean

    Dim s As String
    s = LCase$(Trim$(pathOrUrl))

    p_testIsHttpUrl = _
        (Left$(s, 7) = "http://") Or _
        (Left$(s, 8) = "https://")

End Function
' ---------------------------------------------------------------------------
' p_testReadTextFromPathOrUrl
'   TEST HELPER.
'
'   Local file:
'       uses existing p_ReadTextFile.
'
'   HTTP/SharePoint URL:
'       uses p_testReadTextFromUrl.
' ---------------------------------------------------------------------------
Private Function p_testReadTextFromPathOrUrl(ByVal pathOrUrl As String, _
                                             Optional ByVal authType As String = "xmlhttp", _
                                             Optional ByVal authValue As String = "") As String

    Const CALLER As String = "cTable.p_testReadTextFromPathOrUrl"

    Dim s As String
    s = Trim$(pathOrUrl)

    If Len(s) = 0 Then
        Err.Raise vbObjectError + 1710, CALLER, _
                  "pathOrUrl is required."
    End If

    If p_testIsHttpUrl(s) Then
        p_testReadTextFromPathOrUrl = p_testReadTextFromUrl(s, authType, authValue)
    Else
        If Dir$(s) = "" Then
            Err.Raise vbObjectError + 1711, CALLER, _
                      "File not found: " & s
        End If

        p_testReadTextFromPathOrUrl = p_ReadTextFile(s, "utf-8")
    End If

End Function
' ---------------------------------------------------------------------------
' p_testReadTextFromPathOrUrl
'   TEST HELPER.
'
'   Local file:
'       uses existing p_ReadTextFile.
'
'   HTTP/SharePoint URL:
'       uses p_testReadTextFromUrl.
' ---------------------------------------------------------------------------


' ---------------------------------------------------------------------------
' p_testReadTextFromUrl
'   TEST HELPER.
'
'   authType:
'       "xmlhttp" or "none"
'           Uses MSXML2.XMLHTTP.
'
'       "serverxmlhttp"
'           Uses MSXML2.ServerXMLHTTP.6.0.
'
'       "winhttp"
'           Uses WinHttp.WinHttpRequest.5.1 without an explicit
'           client certificate.
'
'       "clientcert", "piv", "pki"
'           Uses WinHttp.WinHttpRequest.5.1 with a client certificate
'           specified by authValue.
'
'       "bearer"
'           Uses MSXML2.XMLHTTP with an Authorization header
'           containing the bearer token supplied in authValue.
'
'   The default transport is "xmlhttp". Select the transport and
'   authentication mode required by the destination service.
' ---------------------------------------------------------------------------
Private Function p_testReadTextFromUrl(ByVal url As String, _
                                       Optional ByVal authType As String = "xmlhttp", _
                                       Optional ByVal authValue As String = "") As String

    Const CALLER As String = "cTable.p_testReadTextFromUrl"

    Dim mode As String
    mode = LCase$(Trim$(authType))

    On Error GoTo HttpFail

    Select Case mode

        Case "", "none", "xmlhttp", "sharepoint"

            Dim xmlhttp As Object
            Set xmlhttp = CreateObject("MSXML2.XMLHTTP")

            xmlhttp.Open "GET", url, False
            xmlhttp.setRequestHeader "Accept", "application/json, text/plain, */*"
            xmlhttp.send

            If xmlhttp.Status >= 200 And xmlhttp.Status < 300 Then
                p_testReadTextFromUrl = CStr(xmlhttp.responseText)
                Exit Function
            End If

            Err.Raise vbObjectError + 1712, CALLER, _
                      "XMLHTTP failed. HTTP status: " & xmlhttp.Status & _
                      IIf(Len(xmlhttp.statusText) > 0, " " & xmlhttp.statusText, "")

        Case "serverxmlhttp"

            Dim sxh As Object
            Set sxh = CreateObject("MSXML2.ServerXMLHTTP.6.0")

            sxh.Open "GET", url, False
            sxh.setRequestHeader "Accept", "application/json, text/plain, */*"
            sxh.send

            If sxh.Status >= 200 And sxh.Status < 300 Then
                p_testReadTextFromUrl = CStr(sxh.responseText)
                Exit Function
            End If

            Err.Raise vbObjectError + 1713, CALLER, _
                      "ServerXMLHTTP failed. HTTP status: " & sxh.Status & _
                      IIf(Len(sxh.statusText) > 0, " " & sxh.statusText, "")

        Case "winhttp", "default"

            Dim whr As Object
            Set whr = CreateObject("WinHttp.WinHttpRequest.5.1")

            whr.Open "GET", url, False
            whr.setRequestHeader "Accept", "application/json, text/plain, */*"

            On Error Resume Next
            whr.SetAutoLogonPolicy 0
            On Error GoTo HttpFail

            whr.send

            If whr.Status >= 200 And whr.Status < 300 Then
                p_testReadTextFromUrl = CStr(whr.responseText)
                Exit Function
            End If

            Err.Raise vbObjectError + 1714, CALLER, _
                      "WinHttp failed. HTTP status: " & whr.Status & _
                      IIf(Len(whr.statusText) > 0, " " & whr.statusText, "")

        Case "clientcert", "piv", "pki"

            If Len(Trim$(authValue)) = 0 Then
                Err.Raise vbObjectError + 1715, CALLER, _
                          "authValue is required when authType=""clientcert""."
            End If

            Dim whrCert As Object
            Set whrCert = CreateObject("WinHttp.WinHttpRequest.5.1")

            whrCert.Open "GET", url, False
            whrCert.setRequestHeader "Accept", "application/json, text/plain, */*"
            whrCert.SetClientCertificate authValue
            whrCert.send

            If whrCert.Status >= 200 And whrCert.Status < 300 Then
                p_testReadTextFromUrl = CStr(whrCert.responseText)
                Exit Function
            End If

            Err.Raise vbObjectError + 1716, CALLER, _
                      "WinHttp clientcert failed. HTTP status: " & whrCert.Status & _
                      IIf(Len(whrCert.statusText) > 0, " " & whrCert.statusText, "")

        Case "bearer"

            If Len(Trim$(authValue)) = 0 Then
                Err.Raise vbObjectError + 1717, CALLER, _
                          "authValue is required when authType=""bearer""."
            End If

            Dim xhrBearer As Object
            Set xhrBearer = CreateObject("MSXML2.XMLHTTP")

            xhrBearer.Open "GET", url, False
            xhrBearer.setRequestHeader "Accept", "application/json, text/plain, */*"
            xhrBearer.setRequestHeader "Authorization", "Bearer " & authValue
            xhrBearer.send

            If xhrBearer.Status >= 200 And xhrBearer.Status < 300 Then
                p_testReadTextFromUrl = CStr(xhrBearer.responseText)
                Exit Function
            End If

            Err.Raise vbObjectError + 1718, CALLER, _
                      "Bearer XMLHTTP failed. HTTP status: " & xhrBearer.Status & _
                      IIf(Len(xhrBearer.statusText) > 0, " " & xhrBearer.statusText, "")

        Case Else

            Err.Raise vbObjectError + 1719, CALLER, _
                      "Unsupported authType: """ & authType & """. " & _
                      "Use ""xmlhttp"", ""serverxmlhttp"", ""winhttp"", " & _
                      """clientcert"", ""piv"", ""pki"", or ""bearer""."

    End Select

    Exit Function

HttpFail:
    Err.Raise vbObjectError + 1720, CALLER, _
              "Error downloading URL '" & url & "': " & Err.Description

End Function

'------------------------------------------------------------------------------
' p_AddExcelSourceRowColumn
'
' Appends a ctLong column containing the physical worksheet row number from
' which each cTable data row was read.
'
' firstDataWorksheetRow is the Excel worksheet row corresponding to cTable
' row 1. For example:
'
'   hasHeaders=True  -> rng.Row + 1
'   hasHeaders=False -> rng.Row
'
' The column is intentionally added immediately after the source read, before
' callers can sort, filter, append, or otherwise change row order.
'------------------------------------------------------------------------------
Private Sub p_AddExcelSourceRowColumn( _
        ByVal firstDataWorksheetRow As Long, _
        Optional ByVal SourceRowColumnName As String = "_SourceRow")

    Const CALLER As String = "cTable.p_AddExcelSourceRowColumn"

    Dim targetName As String
    targetName = Trim$(SourceRowColumnName)

    If Len(targetName) = 0 Then
        targetName = "_SourceRow"
    End If

    If p_ColIdx(targetName) > 0 Then
        Err.Raise vbObjectError + 1730, CALLER, _
                  "Column already exists: """ & targetName & """"
    End If

    If m_RowCount <= 0 Then Exit Sub

    Dim values() As Variant
    ReDim values(1 To m_RowCount)

    Dim r As Long

    For r = 1 To m_RowCount
        values(r) = CLng(firstDataWorksheetRow + r - 1)
    Next r

    ' Uses the common writer so dtype, indexes, and lineage are handled
    ' consistently with the other derived-column operations.
    p_WriteTransformColumn _
        targetName:=targetName, _
        values:=values, _
        finalType:=ctLong, _
        label:="source_excel_row"
End Sub

'------------------------------------------------------------------------------
' p_AddCsvSourceRecordColumn
'
' Appends a ctLong column containing the 1-based ordinal of each parsed CSV
' data record.
'
' This is intentionally a record number, not a physical text-file line number:
'   - blank physical lines are skipped by read_csv;
'   - headers do not count as records;
'   - a future CSV parser could permit a quoted record to span physical lines.
'
' Thus, _SourceRecord = 1 means the first accepted DATA record from the CSV
' source, regardless of whether the file contains a header row.
'------------------------------------------------------------------------------
Private Sub p_AddCsvSourceRecordColumn( _
        Optional ByVal SourceRecordColumnName As String = "_SourceRecord")

    Const CALLER As String = "cTable.p_AddCsvSourceRecordColumn"

    Dim targetName As String
    targetName = Trim$(SourceRecordColumnName)

    If Len(targetName) = 0 Then
        targetName = "_SourceRecord"
    End If

    If p_ColIdx(targetName) > 0 Then
        Err.Raise vbObjectError + 1732, CALLER, _
                  "Column already exists: """ & targetName & """"
    End If

    If m_RowCount <= 0 Then Exit Sub

    Dim values() As Variant
    ReDim values(1 To m_RowCount)

    Dim r As Long

    For r = 1 To m_RowCount
        values(r) = CLng(r)
    Next r

    ' Use the shared column writer so dtype, index clearing, and lineage are
    ' handled consistently with other generated cTable columns.
    p_WriteTransformColumn _
        targetName:=targetName, _
        values:=values, _
        finalType:=ctLong, _
        label:="source_csv_record"
End Sub
