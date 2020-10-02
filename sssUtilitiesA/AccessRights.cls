VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AccessRights"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Base 1

' ----------------------------------------------------------
' This class is a custom collection of AccessRight objects
' ----------------------------------------------------------
Private m_dic As Dictionary

Public Function Item(ByVal Index As Variant) As AccessRight
Attribute Item.VB_UserMemId = 0
    If Not m_dic.Exists(Index) Then
        Err.Raise vbObjectError + 100, "AccessRights.Item", "Invalid index."
    Else
        Set Item = m_dic.Item(Index)
    End If
End Function

Public Function Count(Optional ByVal strBusinessUnitCode As String) As Long
    Count = m_dic.Count
End Function

Friend Sub Add(objItem As AccessRight, strIndex As String)
    m_dic.Add strIndex, objItem
End Sub
Public Function Exists(ByVal lngObjectId As Long) As Boolean
   
    Exists = m_dic.Exists(CStr(lngObjectId))
    
End Function
Public Function GetAccessRight(ByVal lngObjectId As Long) As enAccessRight
    
    If m_dic.Exists(CStr(lngObjectId)) Then
        GetAccessRight = m_dic.Item(CStr(lngObjectId)).AccessRight
    Else
        Err.Raise vbObjectError + 100, "AccessRights.GetAccessRight", "No Access Right determined for object."
    End If
End Function
Public Sub Remove(ByVal Index As Variant)
    If m_dic.Exists(Index) Then m_dic.Remove (Index)
End Sub

Public Sub Clear()
    m_dic.RemoveAll
End Sub

Private Sub Class_Initialize()
    Set m_dic = New Dictionary
End Sub

Private Sub Class_Terminate()
    Set m_dic = Nothing
End Sub
