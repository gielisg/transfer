Attribute VB_Name = "modSecurity"
Option Explicit
Option Base 1

Const sPre As String = "selcomm8903uneu6jie9h4rjrk2jrbgv38y329ruhrbb32r37ry-329r32rb7i76bjhjhytef:f"
Const sPost As String = "46821j9nuh87h8nbb600=---selcomm909898087841n3b4y32g43291ybvbvtfdk.,ijh87u"

'***    windows api declarations
Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Function IsLicenceKeyFilePresent(ByVal sAppName As String) As Boolean

    Dim fs As New Scripting.FileSystemObject
    Dim sSystemDir As String
    
    '***    determine the systems directory
    sSystemDir = SystemDirectory
    
    '***    the file is written to the systems directory
    sAppName = sAppName & ".txt"
    sAppName = sSystemDir & "\" & sAppName
    
    IsLicenceKeyFilePresent = fs.FileExists(sAppName)
        
End Function

Public Function WriteLicenceKeyFile(ByVal sAppName As String) As Boolean

    Dim fs As New Scripting.FileSystemObject
    Dim f As Scripting.File
    Dim ts As Scripting.TextStream
    Dim sTmp As String
    Dim sSystemDir As String
    
    '***    determine the systems directory
    sSystemDir = SystemDirectory
    '***    determine the disk serial number
    sTmp = fs.GetDrive(fs.GetDriveName(fs.GetAbsolutePathName(App.Path))).SerialNumber
    sTmp = sAppName & sTmp
    '***    encrypt the string
    Dim CMD5 As New CMD5
    sTmp = CMD5.MD5(sTmp)
    '***    a basic spoof if hacker realises it is MD5
    sTmp = sPre & sTmp & sPost
    '***    encrypt again
    sTmp = CMD5.MD5(sTmp)
    
    '***    the file is written to the systems directory
    sAppName = sAppName & ".txt"
    sAppName = sSystemDir & "\" & sAppName
    Set ts = fs.CreateTextFile(sAppName)
    ts.WriteLine sTmp
    
    Set ts = Nothing
    Set fs = Nothing
    WriteLicenceKeyFile = True
End Function
Public Function CheckLicenceKey(ByVal sAppName As String) As Boolean

    Dim fs As New Scripting.FileSystemObject
    Dim f As Scripting.File
    Dim ts As Scripting.TextStream
    Dim sTmp As String
    Dim sLine As String
    Dim sFileName As String
    Dim sSystemDir As String
    
    On Error GoTo ERROR_HANDLER
    
    '***    determine the systems directory
    sSystemDir = SystemDirectory
    '***    the file is written to the systems directory
    sFileName = sAppName & ".txt"
    sFileName = sSystemDir & "\" & sFileName
    Set ts = fs.OpenTextFile(sFileName)
    sLine = ts.ReadLine
        
    '***    determine the disk serial number
    sTmp = fs.GetDrive(fs.GetDriveName(fs.GetAbsolutePathName(App.Path))).SerialNumber
    sTmp = sAppName & sTmp
    '***    encrypt the string
    Dim CMD5 As New CMD5
    sTmp = CMD5.MD5(sTmp)
    '***    a basic spoof if hacker realises it is MD5
    sTmp = sPre & sTmp & sPost
    '***    encrypt again
    sTmp = CMD5.MD5(sTmp)
    
    '***    check against the licence key from the file
    If sTmp <> sLine Then
        CheckLicenceKey = False
    Else
        CheckLicenceKey = True
    End If
    
ERROR_HANDLER:
    
    Set ts = Nothing
    Set fs = Nothing
    
End Function
Private Function SystemDirectory() As String
    Dim WinPath As String
    WinPath = String(145, Chr(0))
    SystemDirectory = Left(WinPath, GetSystemDirectory(WinPath, 145))
End Function
