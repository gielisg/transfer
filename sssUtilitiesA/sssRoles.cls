VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "sssRoles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Base 1

Public Type typRoles
    RoleLevel As Integer
    RoleId As Long
End Type

Private ma_Roles() As Variant
Private m_iNumberOfRoles As Integer

Public Property Get Count() As Integer
    Count = m_iNumberOfRoles
End Property
Public Function InClause() As String
    Dim aRole As typRoles
    Dim idx As Integer
    Dim sInClause As String
    
    For idx = 1 To Me.Count
        aRole = ma_Roles(idx)
        
        sInClause = sInClause & CStr(aRole.RoleId)
        '***    add the comma
        If idx < Me.Count Then
            sInClause = sInClause & ","
        End If
    Next
    InClause = sInClause
End Function
Public Sub Item(iIndex As Integer, ByRef lRoleId As Long, ByRef iLevel As Integer)
    Dim aRole As typRoles
    
    If iIndex <= m_iNumberOfRoles Then
        aRole = ma_Roles(iIndex)
        lRoleId = aRole.RoleId
        iLevel = aRole.RoleLevel
    End If
    
End Sub
Public Function RoleExists(lRoleId As Long)
    Dim idx As Integer
    Dim aRole As typRoles
    Dim bFound As Boolean
    
    bFound = False
    
    For idx = 1 To m_iNumberOfRoles
        aRole = ma_Roles(idx)
        If lRoleId = aRole.RoleId Then
            bFound = True
            Exit For
        End If
    Next
    RoleExists = bFound
End Function
Public Sub AddRole(iLevel As Integer, lRoleId As Long)
    '***********************************************************************************************
    'Purpose    :
    'Created on :
    'Created by: Gordon
    '***********************************************************************************************

    '-------------------------------
    '*** Generic Procedure Code
    '-------------------------------
    On Error GoTo ERROR_HANDLER
    
    Dim strModuleAndProcName As String
    strModuleAndProcName = "sssRoles - AddRole Procedure"

    Dim idx As Integer
    
    If RoleExists(lRoleId) Then GoTo EXIT_PROC
    
    '***    do we need to redim the array
    If (m_iNumberOfRoles + 1) > UBound(ma_Roles) Then
        '***    add another five
        idx = m_iNumberOfRoles + 5
        ReDim Preserve ma_Roles(idx)
    End If
    
    '***    add to the next position in the array
    m_iNumberOfRoles = m_iNumberOfRoles + 1
    Dim aRole As typRoles
    aRole.RoleId = lRoleId
    aRole.RoleLevel = iLevel
    ma_Roles(m_iNumberOfRoles) = aRole
    
EXIT_PROC:
    
    '------------------------------------------------------------
    '*** Generic Error Handling Code (ensure Exit Proc is above)
    '------------------------------------------------------------
    Exit Sub
    
ERROR_HANDLER:

    Dim lErrNum As Long
    Dim strErrDesc As String
    Dim strErrSource As String
    
    lErrNum = Err.Number
    strErrDesc = Err.Description
    strErrSource = "[" & App.Title & " - " & strModuleAndProcName & " - Build " & App.Major & "." & App.Minor & "." & App.Revision
    If Erl > 0 Then strErrSource = strErrSource & " - Line No " & str(Erl)
    If InStr(1, Err.Source, "-->") > 0 Then
        '*** This error has already been handled by our code
        strErrSource = strErrSource & "]" & vbNewLine & "  --> " & Err.Source
    Else
        '*** Newly generated error, log it here.
        strErrSource = strErrSource & "]" & vbNewLine & "  --> [Source: " & Err.Source & "]"
        On Error Resume Next
        LogError lErrNum, strErrSource, strErrDesc
    End If
    '(ALWAYS comment out two of the three options below)
    '*** EITHER display the error here
    'DisplayError lErrNum, strErrSource, strErrDesc
    '*** OR raise the error to the calling procedure
    Err.Raise lErrNum, strErrSource, strErrDesc
    '*** OR ignore the error and continue
    'Resume Next
End Sub

Private Sub Class_Initialize()
    m_iNumberOfRoles = 0
    ReDim ma_Roles(5)
End Sub
