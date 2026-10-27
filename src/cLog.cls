VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Module: cLog

Option Explicit

' ============================================================
' cLog - Status and error message management
'
' Copyright (c) 2026 Brian A. Crawford
' SPDX-License-Identifier: MIT
'
' Licensed under the MIT License.
' See LICENSE in the repository root.
' ============================================================

Private mStatus As Collection
Private mErrors As Collection

Private Sub Class_Initialize()
    Set mStatus = New Collection
    Set mErrors = New Collection
End Sub

Public Sub AddStatus(ByVal msg As String)
    mStatus.Add msg
End Sub

Public Sub AddError(ByVal msg As String)
    mErrors.Add msg
End Sub

Public Property Get StatusMessages() As Collection
    Set StatusMessages = mStatus
End Property

Public Property Get ErrorMessages() As Collection
    Set ErrorMessages = mErrors
End Property

' --- New Properties ---

Public Property Get HasMessages() As Boolean
    HasMessages = (mStatus.count > 0 Or mErrors.count > 0)
End Property

Public Property Get HasErrors() As Boolean
    HasErrors = (mErrors.count > 0)
End Property

Public Property Get HasStatus() As Boolean
    HasStatus = (mStatus.count > 0)
End Property

Public Property Get ErrorCount() As Long
    ErrorCount = mErrors.count
End Property

Public Property Get StatusCount() As Long
    StatusCount = mStatus.count
End Property
