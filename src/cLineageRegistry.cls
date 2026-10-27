VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cLineageRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' ============================================================
' cLineageRegistry  -  Global interning registries for cTable lineage
'
' Copyright (c) 2026 Brian A. Crawford
' SPDX-License-Identifier: MIT
'
' Licensed under the MIT License.
' See LICENSE in the repository root.
' ============================================================
' Version 20260726
'==============================================================================
' cLineageRegistry  -  Global interning registries for cTable lineage
'
' PURPOSE
'   Lineage is stored as small integer IDs, never as object references. This
'   module owns the two registries that turn heavyweight objects into IDs:
'
'     Source registry : cDataSource      -> Long   (a "source ID")
'     Map registry    : Long() of srcIDs -> Long   (a "column map ID")
'
'   Both registries intern by value: identical content always returns the same
'   ID, so a 2,000,000-row table that came from three files holds exactly
'   three source IDs and (usually) three map IDs, not 2,000,000 objects.
'
' IMMUTABILITY CONTRACT
'   Everything handed to this module is treated as frozen. A column map is
'   never edited in place because it may be shared by several cTables. All
'   structural changes intern a NEW map (copy-on-write).
'
' SINGLETON
'   This class carries VB_PredeclaredId = True, so VBA creates one global
'   default instance automatically. Reference it by name, never with New:
'
'       id = cLineageRegistry.SourceId(ds)        ' correct
'       Set r = New cLineageRegistry              ' wrong - private ID space
'
'   IMPORT THIS FILE, DO NOT PASTE IT. The PredeclaredId attribute lives in
'   the .cls header and is lost if the code is copied into a blank class
'   module. Use File > Import File... in the VBE.
'
'   If the attribute is missing, every internal call site (cLineageRegistry.X)
'   fails at compile time rather than silently splitting the registry, and
'   p_Ensure raises on any instance created with New.
'
' REQUIRES: Microsoft Scripting Runtime
'==============================================================================

'------------------------------------------------------------------------------
' Reserved source IDs. IDs below ctSrcFirstUser are pre-seeded and stable
' across the life of the project, so they can be compared against directly.
'------------------------------------------------------------------------------
Public Enum ctSrcReserved
    ctSrcNone = 0           ' sentinel only: "no entry" inside an array
    ctSrcUnknown = 1        ' lineage was never recorded / was lost
    ctSrcCalculated = 2     ' generic computed value
    ctSrcManual = 3         ' written by SetValue / user code
    ctSrcLiteral = 4        ' fill value, default, constant
    ctSrcMissing = 5        ' Null introduced by outer join or concat widening
    ctSrcFirstUser = 16
End Enum

'------------------------------------------------------------------------------
' Cheapest representation currently in use by a table (read-only diagnostic).
'------------------------------------------------------------------------------
Public Enum ctLineageMode
    ctLinOff = 0            ' tracking disabled; zero cost
    ctLinUniform = 1        ' every cell shares one source
    ctLinByColumn = 2       ' one map for all rows, sources differ by column
    ctLinByRow = 3          ' sources differ by row band (concat/union)
    ctLinFull = 4           ' row-level plus cell/column overrides
End Enum

Private m_DefaultEnabled As Boolean

Private m_Src()      As cDataSource
Private m_SrcCount   As Long
Private m_SrcKey     As Object          ' fingerprint -> Long

Private m_Map()      As Variant         ' each element is a Long() 1..n
Private m_MapUni()   As Long            ' srcID if map is uniform, else 0
Private m_MapLen()   As Long
Private m_MapCount   As Long
Private m_MapKey     As Object          ' map key -> Long

'==============================================================================
' LIFECYCLE
'==============================================================================

' Project-wide default for cTable.EnableLineage on newly created tables.
Public Property Get LineageDefaultEnabled() As Boolean
    LineageDefaultEnabled = m_DefaultEnabled
End Property

Public Property Let LineageDefaultEnabled(ByVal v As Boolean)
    m_DefaultEnabled = v
End Property

Private Sub p_Ensure()
    If Not m_SrcKey Is Nothing Then Exit Sub

    ' First touch of this instance: verify it is the default one. Costs a
    ' single object comparison per session, never on the hot path.
    If Not Me Is cLineageRegistry Then
        Err.Raise vbObjectError + 1310, "cLineageRegistry", _
                  "Use the default instance (cLineageRegistry.X). " & _
                  "New cLineageRegistry creates a private, incompatible ID space."
    End If

    Set m_SrcKey = CreateObject("Scripting.Dictionary")
    m_SrcKey.CompareMode = vbTextCompare
    Set m_MapKey = CreateObject("Scripting.Dictionary")
    m_MapKey.CompareMode = vbBinaryCompare

    ReDim m_Src(0 To 63)
    ReDim m_Map(0 To 63)
    ReDim m_MapUni(0 To 63)
    ReDim m_MapLen(0 To 63)

    m_SrcCount = ctSrcFirstUser - 1
    m_MapCount = 0

    p_SeedReserved
End Sub

Private Sub p_SeedReserved()
    p_SeedOne ctSrcUnknown, "Unknown", "Lineage not recorded"
    p_SeedOne ctSrcCalculated, "Calculated", "Derived by cTable computation"
    p_SeedOne ctSrcManual, "Manual", "Written directly by code or user"
    p_SeedOne ctSrcLiteral, "Literal", "Constant or fill value"
    p_SeedOne ctSrcMissing, "Missing", "Introduced as Null by a join or concat"
End Sub

Private Sub p_SeedOne(ByVal id As Long, ByVal kindName As String, _
                      ByVal note As String)
    Dim ds As cDataSource
    Set ds = New cDataSource
    ds.InitSynthetic kindName, note
    Set m_Src(id) = ds
    m_SrcKey.Add ds.Fingerprint, id
End Sub

' Clears every registry. Only safe when no cTable with lineage enabled is
' still alive - existing IDs become meaningless.
Public Sub ResetRegistries()
    Set m_SrcKey = Nothing
    Set m_MapKey = Nothing
    Erase m_Src
    Erase m_Map
    Erase m_MapUni
    Erase m_MapLen
    m_SrcCount = 0
    m_MapCount = 0
End Sub

'==============================================================================
' SOURCE REGISTRY
'==============================================================================

' Intern a cDataSource and return its ID. Nothing maps to ctSrcUnknown.
Public Function SourceId(ByVal ds As cDataSource) As Long
    p_Ensure
    If ds Is Nothing Then
        SourceId = ctSrcUnknown
        Exit Function
    End If

    Dim k As String
    k = ds.Fingerprint

    If m_SrcKey.Exists(k) Then
        SourceId = m_SrcKey(k)
        Exit Function
    End If

    m_SrcCount = m_SrcCount + 1
    If m_SrcCount > UBound(m_Src) Then
        ReDim Preserve m_Src(0 To UBound(m_Src) * 2 + 1)
    End If

    Set m_Src(m_SrcCount) = ds
    m_SrcKey.Add k, m_SrcCount
    SourceId = m_SrcCount
End Function

Public Function SourceById(ByVal id As Long) As cDataSource
    p_Ensure
    If id < 1 Or id > m_SrcCount Then
        Set SourceById = m_Src(ctSrcUnknown)
    Else
        Set SourceById = m_Src(id)
    End If
End Function

' Intern a pseudo-source describing a calculation, so lineage can answer
' "calculated by CalcAvailableSpending" rather than just "calculated".
Public Function CalculatedSourceId(ByVal exprName As String) As Long
    If Len(Trim$(exprName)) = 0 Then
        CalculatedSourceId = ctSrcCalculated
        Exit Function
    End If
    Dim ds As cDataSource
    Set ds = New cDataSource
    ds.InitCalculated exprName
    CalculatedSourceId = SourceId(ds)
End Function

Public Function SourceLabel(ByVal id As Long) As String
    SourceLabel = SourceById(id).Label
End Function

Public Function SourceCount() As Long
    p_Ensure
    SourceCount = m_SrcCount
End Function

'==============================================================================
' COLUMN MAP REGISTRY
'
' A column map is a frozen Long() of length ColCount holding one source ID per
' column. Rows point at a map ID, so N rows sharing a shape cost one map.
'==============================================================================

Private Function p_MapKeyOf(ByRef ids() As Long, ByVal n As Long) As String
    Dim i As Long
    Dim allSame As Boolean
    allSame = True
    For i = 2 To n
        If ids(i) <> ids(1) Then
            allSame = False
            Exit For
        End If
    Next i

    If allSame Then
        p_MapKeyOf = "U|" & n & "|" & ids(1)
        Exit Function
    End If

    ' Fixed-width encoding avoids delimiter scanning and keeps keys compact.
    Dim sb As String
    sb = "V|" & n & "|"
    For i = 1 To n
        sb = sb & Hex$(ids(i)) & ","
    Next i
    p_MapKeyOf = sb
End Function

' Intern a column map. ids() must be 1-based with n elements.
Public Function mapId(ByRef ids() As Long, ByVal n As Long) As Long
    p_Ensure
    If n <= 0 Then
        mapId = 0
        Exit Function
    End If

    Dim k As String
    k = p_MapKeyOf(ids, n)

    If m_MapKey.Exists(k) Then
        mapId = m_MapKey(k)
        Exit Function
    End If

    m_MapCount = m_MapCount + 1
    If m_MapCount > UBound(m_Map) Then
        Dim newUB As Long
        newUB = UBound(m_Map) * 2 + 1
        ReDim Preserve m_Map(0 To newUB)
        ReDim Preserve m_MapUni(0 To newUB)
        ReDim Preserve m_MapLen(0 To newUB)
    End If

    Dim frozen() As Long
    ReDim frozen(1 To n)
    Dim i As Long
    For i = 1 To n
        frozen(i) = ids(i)
    Next i

    m_Map(m_MapCount) = frozen
    m_MapLen(m_MapCount) = n
    If Left$(k, 1) = "U" Then
        m_MapUni(m_MapCount) = ids(1)
    Else
        m_MapUni(m_MapCount) = 0
    End If

    m_MapKey.Add k, m_MapCount
    mapId = m_MapCount
End Function

' Convenience: a map of n columns all pointing at one source.
Public Function UniformMapId(ByVal n As Long, ByVal srcId As Long) As Long
    If n <= 0 Then
        UniformMapId = 0
        Exit Function
    End If
    Dim ids() As Long
    ReDim ids(1 To n)
    Dim i As Long
    For i = 1 To n
        ids(i) = srcId
    Next i
    UniformMapId = mapId(ids, n)
End Function

' O(1) lookup of one column's source within a map. This is the hot path.
Public Function MapEntry(ByVal mapId As Long, ByVal colIdx As Long) As Long
    p_Ensure
    If mapId < 1 Or mapId > m_MapCount Then
        MapEntry = ctSrcUnknown
        Exit Function
    End If
    If colIdx < 1 Or colIdx > m_MapLen(mapId) Then
        MapEntry = ctSrcUnknown
        Exit Function
    End If
    MapEntry = m_Map(mapId)(colIdx)
End Function

' Returns a working copy of the map. Callers may modify the copy freely; the
' registry entry is untouched.
Public Function MapVector(ByVal mapId As Long) As Long()
    p_Ensure
    Dim out() As Long
    If mapId < 1 Or mapId > m_MapCount Then
        ReDim out(1 To 1)
        out(1) = ctSrcUnknown
        MapVector = out
        Exit Function
    End If

    Dim n As Long
    n = m_MapLen(mapId)
    ReDim out(1 To n)
    Dim i As Long
    For i = 1 To n
        out(i) = m_Map(mapId)(i)
    Next i
    MapVector = out
End Function

Public Function MapLength(ByVal mapId As Long) As Long
    p_Ensure
    If mapId < 1 Or mapId > m_MapCount Then Exit Function
    MapLength = m_MapLen(mapId)
End Function

' Returns the shared source ID if every column in the map agrees, else 0.
Public Function MapUniformSource(ByVal mapId As Long) As Long
    p_Ensure
    If mapId < 1 Or mapId > m_MapCount Then Exit Function
    MapUniformSource = m_MapUni(mapId)
End Function

Public Function MapCount() As Long
    p_Ensure
    MapCount = m_MapCount
End Function

'==============================================================================
' DIAGNOSTICS
'==============================================================================

Public Function RegistryReport() As String
    p_Ensure
    Dim sb As String
    sb = "cLineageRegistry" & vbCrLf & _
         "  sources : " & m_SrcCount & vbCrLf & _
         "  maps    : " & m_MapCount & vbCrLf

    Dim i As Long
    For i = 1 To m_SrcCount
        If Not m_Src(i) Is Nothing Then
            sb = sb & "  [" & i & "] " & m_Src(i).Label & vbCrLf
        End If
    Next i
    RegistryReport = sb
End Function
