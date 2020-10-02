VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmChangePassword 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Password"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3990
   Icon            =   "frmChangePassword.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   3990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Height          =   2235
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   3975
      Begin VB.TextBox txtOldPassword 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2145
         MaxLength       =   18
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   315
         Width           =   1575
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   2040
         TabIndex        =   10
         Top             =   1605
         Width           =   1215
      End
      Begin VB.CommandButton cmdChangePassword 
         Caption         =   "&Change"
         Default         =   -1  'True
         Height          =   375
         Left            =   600
         TabIndex        =   9
         Top             =   1605
         Width           =   1215
      End
      Begin VB.TextBox txtPassword 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2145
         MaxLength       =   18
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   697
         Width           =   1575
      End
      Begin VB.TextBox txtReenterPassword 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2145
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label lblPassword 
         Alignment       =   1  'Right Justify
         Caption         =   "Enter Old Password :"
         Height          =   255
         Index           =   2
         Left            =   225
         TabIndex        =   8
         Top             =   315
         Width           =   1815
      End
      Begin VB.Label lblPassword 
         Alignment       =   1  'Right Justify
         Caption         =   "Enter New Password :"
         Height          =   255
         Index           =   1
         Left            =   225
         TabIndex        =   5
         Top             =   697
         Width           =   1815
      End
      Begin VB.Label lblPassword 
         Caption         =   "Re-enter New Password :"
         Height          =   255
         Index           =   0
         Left            =   225
         TabIndex        =   4
         Top             =   1080
         Width           =   1815
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   2220
      Width           =   3990
      _ExtentX        =   7038
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6985
            Text            =   "Enter New Password."
            TextSave        =   "Enter New Password."
         EndProperty
      EndProperty
   End
   Begin VB.Label lblStatus 
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   1680
      Width           =   3975
   End
End
Attribute VB_Name = "frmChangePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Public blnCancel As Boolean
Public WithEvents m_objUser As CUser
Attribute m_objUser.VB_VarHelpID = -1

Private Sub cmdCancel_Click()
    '-------------------------------
    '*** Generic Procedure Code
    '-------------------------------
1     On Error GoTo errHandler
2     Dim sModuleAndProcName As String
3     sModuleAndProcName = "frmChangePassword - cmdCancel_Click Function"

    Unload Me
        
'------------------------------------------------------------
'*** Generic Error Handling Code (ensure Exit Proc is above)
'------------------------------------------------------------
errHandler:
  If Err.Number = 0 Then Exit Sub
  Dim lErrNum As Long, sErrDesc As String, sErrSource As String
  lErrNum = Err.Number
  sErrDesc = Err.Description
  sErrSource = "[" & App.Title & " - " & sModuleAndProcName & " - Build " & App.Major & "." & App.Minor & "." & App.Revision
  If Erl > 0 Then sErrSource = sErrSource & " - Line No " & str(Erl)
  If InStr(1, Err.Source, "-->") > 0 Then
    '*** This error has already been handled by our code
    sErrSource = sErrSource & "]" & vbNewLine & "  --> " & Err.Source
  Else
    '*** Newly generated error, log it here.
    sErrSource = sErrSource & "]" & vbNewLine & "  --> [Source: " & Err.Source & "]"
    On Error Resume Next
    LogError lErrNum, sErrSource, sErrDesc
  End If
  '(ALWAYS comment out two of the three options below)
  '*** EITHER display the error here
    DisplayError lErrNum, sErrSource, sErrDesc
  '*** OR raise the error to the calling procedure
    'Err.Raise lErrNum, sErrSource, sErrDesc
  '*** OR ignore the error and continue
    'Resume Next

End Sub
Private Sub cmdChangePassword_Click()
'-------------------------------
'*** Generic Procedure Code
'-------------------------------
1     On Error GoTo errHandler
2     Dim sModuleAndProcName As String
3     sModuleAndProcName = "frmChangePassword - cmdChangePassword_Click Function"
  
    Dim strSQL As String
    Dim rs As ADODB.Recordset
    Dim dbAccess As dbAccess
    Dim colParams As Collection
    Dim colParamTypes As Collection
    Dim sMessage As String
    Dim sPassword As String
    Dim idx As Integer
    Dim colDatabase As Collection
    Dim typDatasource As typDatasourceNew2
    
    '***    check the user has entered the correct existing password
    If Trim(txtOldPassword.Text) <> Trim(m_objUser.Password) Then
        sMessage = "Invalid existing password."
        txtOldPassword.SetFocus
        MsgBox sMessage, vbCritical, "Invalid Password"
        Exit Sub
    End If
    
    '***    get rid of the list of invalid passwords
    If LCase(txtPassword.Text) = "password" Or LCase(txtPassword.Text) = "passwd" Or _
        LCase(txtPassword.Text) = "test" Or LCase(txtPassword.Text) = "123" Or _
        LCase(txtPassword.Text) = "1234" Or LCase(txtPassword.Text) = "selcomm" Or _
        LCase(txtPassword.Text) = "select" Or _
        (LCase(txtPassword.Text) = LCase(m_objUser.LogonCode)) _
    Then
        sMessage = txtPassword.Text & " is not a permitted password. Choose another."
        MsgBox sMessage, vbCritical, "Invalid Password"
        Exit Sub
    End If
    
    '***    invoke the string validator
    
    ' Check user inputs
    If txtPassword = "" Then
        MsgBox "Please enter Password !!!", vbCritical, "Change Password"
        txtPassword.SetFocus
        Exit Sub
    End If
  
    If txtReenterPassword = "" Then
        MsgBox "Please re-enter Password !!!", vbCritical, "Change Password"
        txtReenterPassword.SetFocus
        Exit Sub
    End If
    
    If txtReenterPassword <> txtPassword Then
        MsgBox "Re-entered password must be the same !!!", vbCritical, "Change Password"
        txtReenterPassword.Text = ""
        txtReenterPassword.SetFocus
        Exit Sub
    End If
    
    '***    validate the password
    sPassword = Trim(txtPassword.Text)
    
    Set dbAccess = m_objUser.GetDbAccessObject(m_objUser.CurrentBusinessUnitCode)
                
    strSQL = "execute procedure ss_validate_str(p_strtypecd=?, p_teststring=?)"
    
    Set rs = dbAccess.LoadRS(strSQL, "PASSWD", sPassword)
    
    While Not rs.EOF
        If FieldToString(rs, 1) <> "O" Then
            MsgBox FieldToString(rs, 2), vbExclamation, "Invalid Password"
            Exit Sub
        End If
        rs.MoveNext
    Wend
        
    '***    get the dbAccess object
    Set dbAccess = m_objUser.GetDbAccessObject(m_objUser.DefaultBusinessUnitCode)
    
    '***    Loop through all databases if multiple
    strSQL = "execute procedure sp_chg_oper_pass (p_login_code = ?, p_old_passwd = ?, p_new_passwd = ?, resume_flg = 'Y', tran_flg = 'Y')"
    
    Set rs = dbAccess.LoadRS(strSQL, m_objUser.LogonCode, m_objUser.Password, sPassword)
    
    If rs("StatusCode") = 0 Then
        blnCancel = False
        m_objUser.Password = Trim(txtPassword.Text)
        
        Unload Me
    Else
        MsgBox "Unable to change password." & vbNewLine & _
            Trim(rs(rs.Fields.Count - 1)), vbCritical + vbOKOnly, "Password Change Failed"
    End If
    
'------------------------------------------------------------
'*** Generic Error Handling Code (ensure Exit Proc is above)
'------------------------------------------------------------
errHandler:
  If Err.Number = 0 Then Exit Sub
  Dim lErrNum As Long, sErrDesc As String, sErrSource As String
  lErrNum = Err.Number
  sErrDesc = Err.Description
  sErrSource = "[" & App.Title & " - " & sModuleAndProcName & " - Build " & App.Major & "." & App.Minor & "." & App.Revision
  If Erl > 0 Then sErrSource = sErrSource & " - Line No " & str(Erl)
  If InStr(1, Err.Source, "-->") > 0 Then
    '*** This error has already been handled by our code
    sErrSource = sErrSource & "]" & vbNewLine & "  --> " & Err.Source
  Else
    '*** Newly generated error, log it here.
    sErrSource = sErrSource & "]" & vbNewLine & "  --> [Source: " & Err.Source & "]"
    On Error Resume Next
    LogError lErrNum, sErrSource, sErrDesc
  End If
  '(ALWAYS comment out two of the three options below)
  '*** EITHER display the error here
    DisplayError lErrNum, sErrSource, sErrDesc
  '*** OR raise the error to the calling procedure
    'Err.Raise lErrNum, sErrSource, sErrDesc
  '*** OR ignore the error and continue
    'Resume Next

    
End Sub

Private Sub Form_Load()
    blnCancel = True
End Sub

