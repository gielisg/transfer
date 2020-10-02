VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "sssLauncher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Base 1

Private m_objParent As Object
Public AddInfo As Collection
Private m_objUser As sssUtilitiesA.CUser

'Public Sub Start(ArgumentsString As String)
''    Set objForm = New frmLauncher
'    Set AddInfo = New Collection
'
''    If Not m_objParent Is Nothing Then        Set objForm.Parent = objParent
''    End If
''
''    If Not objForm.Start(ArgumentsString) Then
''        'populate error, out here
''        Exit Sub
''    End If
'
''    On Error GoTo 0
''
''    'Show Modal, otherwise application ends after the form is displayed.
''    objForm.Show vbModal
'
'
'
'
'End Sub
Public Property Set Parent(objParent As Object)
    Set m_objParent = objParent
End Property
Public Property Set User(objUser As sssUtilitiesA.CUser)
    Set m_objUser = objUser
End Property
Public Sub Start(ArgumentString As String)

    Const ksMethod = "Start"
    
    Dim objDataContainer As CommonDataComponent.CommonData
    Dim objApplication As ILaunchedApplication
    
    Dim sAppTitle As String
    Dim strErrMsg As String
    Dim sAppName As String
    
    On Error GoTo ErrorHandler
    
    Set AddInfo = New Collection
    
    '***    start the data object
    Set objDataContainer = CreateObject("CommondataComponent.CommonData")
    objDataContainer.CmdLine = ArgumentString
    
    With objDataContainer
    
        '***    Application name is a required argument
        If Trim(.Value(gksAppName)) = "" Then
            strErrMsg = "Application Name Missing In parameters passed." & Chr(13) & Chr(10) & "Launched program will not run."
            MsgBox strErrMsg, vbOKOnly Or vbCritical, "Application Launch Control"
        End If
               
        '***    Database name is a required argument
        If Trim(.Value(gksDBName)) = "" Then
            strErrMsg = "Database Name Missing In parameters passed." & Chr(13) & Chr(10) & "Launched program will not run."
            MsgBox strErrMsg, vbOKOnly Or vbCritical, "Application Launch Control"
            Exit Sub
        End If
        
        ' Get the db instance from the registry and add to common data.
        ' this is also a required value
        If .Value(gksDBInstance) = "" Then
            '***    get the DBInstance from the user
            .Value(gksDBInstance) = m_objUser.DBInstance
'          '// Added by Abha on 22/05/2001 as The latest version of Informix Client write to HKEY_LOCAL_MACHINE and the old version write to HKEY_CURRENT_USER
'          '// So first check in HKEY_LOCAL_MACHINE if latest machine and then HKEY_CURRENT_USER
'          .Value(gksDBInstance) = QueryValue(HKEY_LOCAL_MACHINE, strInformixServerSubkey, "INFORMIXSERVER")
'            If .Value(gksDBInstance) = "" Then
'                .Value(gksDBInstance) = QueryValue(HKEY_CURRENT_USER, strInformixServerSubkey, "INFORMIXSERVER")
'              If Trim(.Value(gksDBInstance)) = "" Then
'                  strErrMsg = "Informix DatabaseServer Name Missing in system registry." & Chr(13) & Chr(10) & "Launched program will not run."
'                  MsgBox strErrMsg, vbOKOnly Or vbCritical, "Application Launch Control"
'
'                Exit Sub
'              End If
'            End If
        End If
        
        ' Get the host name from the registry and add to common data
        ' this is also a required value
        If .Value(gksDBHost) = "" Then
            .Value(gksDBHost) = m_objUser.DBHost
'          .Value(gksDBHost) = QueryValue(HKEY_LOCAL_MACHINE, strInformixHostSubkey & .Value(gksDBInstance), "HOST")
'          If Trim(.Value(gksDBHost)) = "" Then
'              strErrMsg = "Informix SqlHost Name Missing in system registry" & Chr(13) & Chr(10) & "Launched program will not run."
'              MsgBox strErrMsg, vbOKOnly Or vbCritical, "Application Launch Control"
'
'                Exit Sub
'          End If
        End If
        
  
        
        ' Create an instance of the application
        strErrMsg = "Attempting to create object " & .AppName

          sAppName = .AppName
          sAppName = Trim(Replace(sAppName, "'", " "))
          Set objApplication = CreateObject(sAppName)
        
        strErrMsg = "Object created."

        If Not m_objParent Is Nothing Then Set objApplication.Parent = m_objParent
        
        strErrMsg = "Attempting to start application."
        
        objApplication.Start .Stream
        
    End With
    
    '***    this is a bit of a kludge. Attempt to extract the results from
    '***    the old style dialog applications and preserve in the the AddInfo
    '***    collection
    On Error Resume Next
    
    If Not objApplication.AddInfo Is Nothing Then
        Set AddInfo = objApplication.AddInfo
    End If
    
    
    On Error GoTo ErrorHandler
    
    Exit Sub
            
            
ErrorHandler:


    MsgBox "Error " & Err.Number & vbNewLine & Err.Description & vbNewLine & "Proc: " & TypeName$(Me) & ksMethod & vbNewLine & "Line Number: " & Erl & vbNewLine & "Status: " & strErrMsg & vbNewLine & "Arguments:" & ReplacePassword(Command)

End Sub

Private Function ReplacePassword(ByVal str As String) As String
    Dim pwdStart As Long
    Dim pwdEnd As Long
    Dim strPassword  As String
    Dim strReplace As String
    
    Const Passwordstring As String = "pwd="
    'search for pwd position
    pwdStart = InStr(1, LCase(str), Passwordstring, vbTextCompare)
    If pwdStart = 0 Then
        Exit Function
    Else
        'get end of pwd
        pwdEnd = InStr(pwdStart, str, "/", vbTextCompare)
        strPassword = Mid$(str, pwdStart + Len(Passwordstring), pwdEnd - pwdStart - Len(Passwordstring))
        'check if passwdor come with " around it
        If Right(strPassword, 1) = Chr(34) And Left(strPassword, 1) = Chr(34) Then
            strReplace = Chr(34) & String(Len(strPassword) - 2, "*") & Chr(34)
        Else
            strReplace = String(Len(strPassword), "*")
        End If
        ReplacePassword = Left(str, pwdStart + Len(Passwordstring) - 1) & strReplace & Mid(str, pwdEnd, Len(str) - pwdEnd)
    End If
        
End Function

