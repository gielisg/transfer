VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "sssNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Base 1

Private m_col As Collection
Private m_NodeItem As sssNodeItem

Public Property Set NodeItem(ByVal vData As sssNodeItem)
    Set m_NodeItem = vData
End Property

Public Property Get NodeItem() As sssNodeItem
    If m_NodeItem Is Nothing Then Set m_NodeItem = New sssNodeItem
    Set NodeItem = m_NodeItem
End Property
Public Function Count() As Integer
    If m_col Is Nothing Then
        Count = 0
    Else
        Count = m_col.Count
    End If
End Function
Public Sub AddNode(objNode As Object)
    If m_col Is Nothing Then Set m_col = New Collection
    m_col.Add objNode
End Sub
Public Function item(Index As Integer) As Object
    Set item = m_col.item(Index)
End Function
'Fix for count integer, cannot break binary compatibility.
Public Function LongCount() As Long
    If m_col Is Nothing Then
        LongCount = 0
    Else
        LongCount = m_col.Count
    End If
End Function
Public Function Longitem(Index As Long) As Object
    Set Longitem = m_col.item(Index)
End Function

