VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Autonumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Base 1

Private m_objUser As CUser
Property Let User(User As CUser)
    Set m_objUser = User
End Property
Public Function GetNextNumber(strKey As String) As Long
    '***********************************************************************************************
    'Purpose    : Returns a number from the stack (autonum table)
    'Created on :
    'Created by: Gordon
    '***********************************************************************************************

    '-------------------------------
    '*** Generic Procedure Code
    '-------------------------------
    On Error GoTo ERROR_HANDLER
    
    Dim strModuleAndProcName As String
    strModuleAndProcName = "Autonumber - GetNextNumber Function"
    Dim dbAccess As dbAccess
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    
    '***    some validations
    If m_objUser Is Nothing Then
        Err.Raise vbObjectError + 100, strModuleAndProcName, "User object has not been set."
    End If
    
    '***    obtain a dbAccess object
    Set dbAccess = m_objUser.GetDbAccessObject(m_objUser.DefaultBusinessUnitCode)
    
    strSQL = "execute procedure sp_autonum(p_autonum=?)"
    Set rs = dbAccess.LoadRS(strSQL, strKey)
    
    GetNextNumber = rs(0)
    
EXIT_PROC:
    
    '------------------------------------------------------------
    '*** Generic Error Handling Code (ensure Exit Proc is above)
    '------------------------------------------------------------
    Exit Function
    
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

End Function
