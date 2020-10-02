VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "sssNodeItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Base 1

Private m_dic As Dictionary

Public Sub AddAttribute(ByVal strKey As String, ByVal varValue As Variant)
    m_dic.Add strKey, varValue
End Sub
Public Function GetAttribute(ByVal strKey As String) As Variant
    If Not m_dic.Exists(strKey) Then
        GetAttribute = ""
    Else
        GetAttribute = m_dic.Item(strKey)
    End If
End Function
Public Function Exists(ByVal strKey As String) As Boolean
    Exists = m_dic.Exists(strKey)
End Function
Private Sub Class_Initialize()
    Set m_dic = New Dictionary
End Sub

Private Sub Class_Terminate()
    Set m_dic = Nothing
End Sub
