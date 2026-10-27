VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cDataSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ============================================================
' cDataSource  -  Immutable metadata container for a cTable lineage source
'
' Copyright (c) 2026 Brian A. Crawford
' SPDX-License-Identifier: MIT
'
' Licensed under the MIT License.
' See LICENSE in the repository root.
' ============================================================
'
' Version 20260726
'==============================================================================
' cDataSource  -  Immutable metadata container for a cTable lineage source
'
' CHANGES FROM V1
'   ADDED  Kind (ctSourceKind) so synthetic sources - calculated, manual,
'          literal, missing, unknown - are first-class rather than sentinels.
'   ADDED  Fingerprint: a stable value-identity key. Two cDataSource objects
'          describing the same origin produce the same fingerprint, which is
'          what lets mLineage intern them to a single ID.
'   ADDED  Label: one-line human summary, used by lineage reporting.
'   ADDED  InitCalculated, InitManual, InitSynthetic, InitFromPathDeclared.
'   ADDED  IsInitialized and a soft freeze. Once interned, a cDataSource must
'          not change, or its registry ID would silently point at new content.
'   KEPT   Every V1 member and Init* signature.
'
' All existing V1 code continues to work unchanged.
'==============================================================================

Public Enum ctSourceKind
    ctKindUnknown = 0
    ctKindExcelSheet = 1
    ctKindExcelTable = 2
    ctKindExcelRange = 3
    ctKindAccessTable = 4
    ctKindCSV = 5
    ctKindArray = 6
    ctKindJSON = 7
    ctKindCalculated = 10
    ctKindManual = 11
    ctKindLiteral = 12
    ctKindMissing = 13
End Enum


Private m_FilePath As String
Private m_SheetName As String
Private m_TableName As String
Private m_Author As String
Private m_LastSaveTime As Date
Private m_LastAuthor As String
Private m_ImportDate As Date
Private m_ImportUser As String
Private m_SourceType As String
Private m_Kind As ctSourceKind
Private m_Expression As String       ' calculation name, for ctKindCalculated
Private m_Note As String
Private m_Init As Boolean
Private m_Frozen As Boolean

'==============================================================================
' INITIALIZATION  (V1 signatures preserved)
'==============================================================================

Public Sub InitFromPath(ByVal filePath As String, _
                        Optional ByVal sheetName As String = "", _
                        Optional ByVal tableName As String = "")
    p_GuardFrozen
    m_FilePath = filePath
    m_SheetName = sheetName
    m_TableName = tableName
    m_SourceType = "ExcelSheet"
    m_Kind = ctKindExcelSheet

    Dim wb As Workbook
    Dim didOpen As Boolean
    On Error Resume Next
    Set wb = Workbooks(Dir(filePath))
    On Error GoTo 0
    If wb Is Nothing Then
        Set wb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
        didOpen = True
    End If

    If Not wb Is Nothing Then
        m_Author = wb.BuiltinDocumentProperties("Author")
        m_LastSaveTime = wb.BuiltinDocumentProperties("Last Save Time")
        m_LastAuthor = wb.BuiltinDocumentProperties("Last Author")
        If didOpen Then wb.Close False
    End If
    p_Stamp
End Sub

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

    ' Workbook document properties are optional metadata. A workbook may be
    ' readable even if Excel cannot retrieve one or more of these properties.
    If Not wb Is Nothing Then
        m_Author = p_BuiltInDocumentPropertyText(wb, "Author")
        m_LastSaveTime = p_BuiltInDocumentPropertyDate(wb, "Last Save Time")
        m_LastAuthor = p_BuiltInDocumentPropertyText(wb, "Last Author")
    End If

    p_Stamp
End Sub
'------------------------------------------------------------------------------
' Returns a built-in document property as text.
'
' Missing, inaccessible, or unsupported properties are treated as unavailable
' optional metadata and return an empty string.
'------------------------------------------------------------------------------
Private Function p_BuiltInDocumentPropertyText( _
        ByVal wb As Workbook, _
        ByVal propertyName As String) As String

    Dim propertyValue As Variant

    On Error Resume Next
    propertyValue = wb.BuiltinDocumentProperties(propertyName).value

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


'------------------------------------------------------------------------------
' Returns a built-in document property as a Date.
'
' Missing, inaccessible, unsupported, or non-date properties return VBA's
' zero Date value, which is already how cDataSource represents unavailable
' LastSaveTime metadata.
'------------------------------------------------------------------------------
Private Function p_BuiltInDocumentPropertyDate( _
        ByVal wb As Workbook, _
        ByVal propertyName As String) As Date

    Dim propertyValue As Variant

    On Error Resume Next
    propertyValue = wb.BuiltinDocumentProperties(propertyName).value

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


' Declare an Excel source WITHOUT opening the workbook. Use when the file is
' unavailable, when running outside Excel (Access host), or in tests. Pass
' lastSave if you know it - it participates in the fingerprint.
Public Sub InitFromPathDeclared(ByVal filePath As String, _
                                Optional ByVal sheetName As String = "", _
                                Optional ByVal tableName As String = "", _
                                Optional ByVal lastSave As Date)
    p_GuardFrozen
    m_FilePath = filePath
    m_SheetName = sheetName
    m_TableName = tableName
    m_SourceType = "ExcelSheet"
    m_Kind = IIf(Len(tableName) > 0, ctKindExcelTable, ctKindExcelSheet)
    m_LastSaveTime = lastSave
    p_Stamp
End Sub

Public Sub InitFromAccessPath(ByVal filePath As String, _
                              Optional ByVal tableName As String = "")
    p_GuardFrozen
    m_FilePath = filePath
    m_TableName = tableName
    m_SourceType = "AccessTable"
    m_Kind = ctKindAccessTable
    m_LastSaveTime = FileDateTime(filePath)
    p_Stamp
End Sub

Public Sub InitFromCSV(ByVal filePath As String)
    p_GuardFrozen

    m_FilePath = filePath
    m_SourceType = "CSV"
    m_Kind = ctKindCSV

    ' File system metadata only. This does not open the file in Excel.
    On Error Resume Next
    If Len(Dir$(filePath)) > 0 Then
        m_LastSaveTime = FileDateTime(filePath)
    End If
    On Error GoTo 0

    p_Stamp
End Sub
Public Sub InitFromJSON(ByVal filePath As String)
    p_GuardFrozen

    m_FilePath = filePath
    m_SourceType = "JSON"
    m_Kind = ctKindJSON

    ' File system metadata only. This does not open the file in Excel.
    On Error Resume Next
    If Len(Dir$(filePath)) > 0 Then
        m_LastSaveTime = FileDateTime(filePath)
    End If
    On Error GoTo 0

    p_Stamp
End Sub


Public Sub InitFromRange(ByVal Rng As Object)
    p_GuardFrozen
    m_SourceType = "ExcelRange"
    m_Kind = ctKindExcelRange
    On Error Resume Next
    m_FilePath = Rng.Worksheet.parent.fullName
    m_SheetName = Rng.Worksheet.name
    m_TableName = Rng.Address(False, False)
    On Error GoTo 0
    p_Stamp
End Sub

' A value produced by a calculation rather than read from anywhere.
Public Sub InitCalculated(ByVal expressionName As String)
    p_GuardFrozen
    m_SourceType = "Calculated"
    m_Kind = ctKindCalculated
    m_Expression = expressionName
    p_Stamp
End Sub

Public Sub InitManual(Optional ByVal note As String = "")
    p_GuardFrozen
    m_SourceType = "Manual"
    m_Kind = ctKindManual
    m_Note = note
    p_Stamp
End Sub

' Used by mLineage to seed the reserved source IDs.
Public Sub InitSynthetic(ByVal kindName As String, ByVal note As String)
    p_GuardFrozen
    m_SourceType = kindName
    m_Note = note
    Select Case LCase$(kindName)
        Case "calculated": m_Kind = ctKindCalculated
        Case "manual":     m_Kind = ctKindManual
        Case "literal":    m_Kind = ctKindLiteral
        Case "missing":    m_Kind = ctKindMissing
        Case Else:         m_Kind = ctKindUnknown
    End Select
    p_Stamp
End Sub

Private Sub p_Stamp()
    m_ImportDate = Now()
    m_ImportUser = Environ$("USERNAME")
    m_Init = True
End Sub

Private Sub p_GuardFrozen()
    If m_Frozen Then
        Err.Raise vbObjectError + 1300, "cDataSource", _
                  "This data source has been interned into the lineage " & _
                  "registry and can no longer be modified. Create a new " & _
                  "cDataSource instead."
    End If
End Sub

' Called once by mLineage on intern. Idempotent.
Public Sub Freeze()
    m_Frozen = True
End Sub

Public Property Get IsInitialized() As Boolean
    IsInitialized = m_Init
End Property

Public Property Get IsFrozen() As Boolean
    IsFrozen = m_Frozen
End Property

'==============================================================================
' VALUE IDENTITY
'
' Fingerprint deliberately excludes ImportDate and ImportUser: reading the
' same file twice in one session must produce one lineage source, not two.
' LastSaveTime IS included, so re-reading a file that has since been edited
' correctly produces a distinct source.
'==============================================================================

Public Property Get Fingerprint() As String
    Dim sb As String
    sb = CStr(CLng(m_Kind)) & "|" & _
         LCase$(m_FilePath) & "|" & _
         LCase$(m_SheetName) & "|" & _
         LCase$(m_TableName) & "|" & _
         LCase$(m_Expression)

    If m_LastSaveTime <> 0 Then
        sb = sb & "|" & Format$(m_LastSaveTime, "yyyy-mm-dd hh:nn:ss")
    Else
        sb = sb & "|"
    End If

    If m_Kind = ctKindUnknown Or m_Kind = ctKindManual _
       Or m_Kind = ctKindLiteral Or m_Kind = ctKindMissing Then
        sb = sb & "|" & LCase$(m_SourceType) & "|" & LCase$(m_Note)
    End If

    Fingerprint = sb
End Property

' One-line human summary, e.g.
'   ExcelSheet  Q3Actuals.xlsx!Detail  (saved 2026-07-01 14:22 by jsmith)
'   Calculated  CalcAvailableSpending
Public Property Get label() As String
    Dim sb As String

    Select Case m_Kind
        Case ctKindCalculated
            label = "Calculated  " & _
                    IIf(Len(m_Expression) > 0, m_Expression, "(unnamed)")
            Exit Property
        Case ctKindManual, ctKindLiteral, ctKindMissing, ctKindUnknown
            If Len(m_FilePath) = 0 Then
                label = m_SourceType & IIf(Len(m_Note) > 0, "  (" & m_Note & ")", "")
                Exit Property
            End If
    End Select

    sb = m_SourceType & "  " & p_BaseName(m_FilePath)
    If Len(m_SheetName) > 0 Then sb = sb & "!" & m_SheetName
    If Len(m_TableName) > 0 Then sb = sb & "[" & m_TableName & "]"
    If m_LastSaveTime <> 0 Then
        sb = sb & "  (saved " & Format$(m_LastSaveTime, "yyyy-mm-dd hh:nn")
        If Len(m_LastAuthor) > 0 Then sb = sb & " by " & m_LastAuthor
        sb = sb & ")"
    End If
    label = sb
End Property

Private Function p_BaseName(ByVal p As String) As String
    Dim i As Long
    i = InStrRev(p, "\")
    If i = 0 Then i = InStrRev(p, "/")
    If i = 0 Then p_BaseName = p Else p_BaseName = mid$(p, i + 1)
End Function

'==============================================================================
' PROPERTIES  (V1)
'==============================================================================

Public Property Get filePath() As String
    filePath = m_FilePath
End Property

Public Property Get sheetName() As String
    sheetName = m_SheetName
End Property

Public Property Get tableName() As String
    tableName = m_TableName
End Property

Public Property Get Author() As String
    Author = m_Author
End Property

Public Property Get lastSaveTime() As Date
    lastSaveTime = m_LastSaveTime
End Property

Public Property Get LastAuthor() As String
    LastAuthor = m_LastAuthor
End Property

Public Property Get ImportDate() As Date
    ImportDate = m_ImportDate
End Property

Public Property Get ImportUser() As String
    ImportUser = m_ImportUser
End Property

Public Property Get SourceType() As String
    SourceType = m_SourceType
End Property

Public Property Let SourceType(ByVal v As String)
    p_GuardFrozen
    m_SourceType = v
End Property

Public Property Get kind() As ctSourceKind
    kind = m_Kind
End Property

Public Property Get Expression() As String
    Expression = m_Expression
End Property

Public Property Get note() As String
    note = m_Note
End Property

'==============================================================================
' DIAGNOSTIC
'==============================================================================

Public Function Describe() As String
    Dim sb As String
    sb = "cDataSource [" & m_SourceType & "]" & vbCrLf
    If Len(m_FilePath) > 0 Then sb = sb & "FilePath: " & m_FilePath & vbCrLf
    If m_SheetName <> "" Then sb = sb & "SheetName: " & m_SheetName & vbCrLf
    If m_TableName <> "" Then sb = sb & "TableName: " & m_TableName & vbCrLf
    If m_Expression <> "" Then sb = sb & "Expression: " & m_Expression & vbCrLf
    If m_Author <> "" Then sb = sb & "Author: " & m_Author & vbCrLf
    If m_LastSaveTime <> 0 Then sb = sb & "LastSaveTime: " & m_LastSaveTime & vbCrLf
    If m_LastAuthor <> "" Then sb = sb & "LastAuthor: " & m_LastAuthor & vbCrLf
    sb = sb & "ImportDate: " & m_ImportDate & vbCrLf
    sb = sb & "ImportUser: " & m_ImportUser & vbCrLf
    sb = sb & "Fingerprint: " & Fingerprint & vbCrLf
    Describe = sb
End Function
