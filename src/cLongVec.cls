VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cLongVec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' cLongVec - A minimal fixed-length Long vector with O(1) get/set.
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
' cLongVec  -  Minimal fixed-length Long vector with O(1) get/set.
'
' VBA forbids Public arrays in class modules, and assigning through a Variant
' that holds an array (v(i) = x) does not work. This wrapper gives cLineage a
' dense per-column overlay it can write to one element at a time without
' copying the whole array.
'
' Value 0 always means "no override - fall through to the next lineage tier".
'==============================================================================

Private m_V() As Long
Private m_N As Long

Public Sub Init(ByVal n As Long)
    m_N = n
    If n > 0 Then ReDim m_V(1 To n) Else Erase m_V
End Sub

Public Property Get count() As Long
    count = m_N
End Property

Public Property Get item(ByVal i As Long) As Long
    If i < 1 Or i > m_N Then Exit Property
    item = m_V(i)
End Property

Public Sub SetItem(ByVal i As Long, ByVal v As Long)
    If i < 1 Or i > m_N Then Exit Sub
    m_V(i) = v
End Sub

Public Sub Fill(ByVal v As Long)
    Dim i As Long
    For i = 1 To m_N
        m_V(i) = v
    Next i
End Sub

' Number of elements that actually carry an override.
Public Function NonZeroCount() As Long
    Dim i As Long, k As Long
    For i = 1 To m_N
        If m_V(i) <> 0 Then k = k + 1
    Next i
    NonZeroCount = k
End Function

' Load from a 1-based Long array.
Public Sub FromArray(ByRef src() As Long, ByVal n As Long)
    Init n
    Dim i As Long
    For i = 1 To n
        m_V(i) = src(i)
    Next i
End Sub

Public Function CloneMe() As cLongVec
    Set CloneMe = New cLongVec
    CloneMe.FromArray m_V, m_N
End Function
