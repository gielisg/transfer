VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmLogon 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Logon"
   ClientHeight    =   1995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3990
   Icon            =   "frmLogon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   3990
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   1740
      Width           =   3990
      _ExtentX        =   7038
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4409
            MinWidth        =   176
            Text            =   "Enter User Id and Password."
            TextSave        =   "Enter User Id and Password."
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame4 
      Height          =   1695
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   3975
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   240
         Top             =   960
      End
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   3480
         Top             =   1080
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   2040
         TabIndex        =   3
         Top             =   1185
         Width           =   1215
      End
      Begin VB.CommandButton cmdLogon 
         Caption         =   "&Logon"
         Default         =   -1  'True
         Height          =   375
         Left            =   600
         TabIndex        =   2
         Top             =   1185
         Width           =   1215
      End
      Begin VB.TextBox txtUserName 
         Height          =   315
         Left            =   1440
         MaxLength       =   18
         TabIndex        =   0
         Top             =   345
         Width           =   1575
      End
      Begin VB.TextBox txtPassword 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1440
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lblUserId 
         Alignment       =   1  'Right Justify
         Caption         =   "User Id :"
         Height          =   255
         Left            =   480
         TabIndex        =   6
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblPassword 
         Caption         =   "Password :"
         Height          =   255
         Left            =   600
         TabIndex        =   5
         Top             =   720
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Choose Role"
      Height          =   1695
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   3975
      Begin VB.CommandButton cmdChoose 
         Caption         =   "&Choose"
         Height          =   375
         Left            =   600
         TabIndex        =   11
         Top             =   1185
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel2 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   2040
         TabIndex        =   10
         Top             =   1185
         Width           =   1215
      End
      Begin VB.ListBox lstRoles 
         Height          =   840
         Left            =   240
         TabIndex        =   9
         Top             =   255
         Width           =   3495
      End
   End
End
Attribute VB_Name = "frmLogon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Public m_bLogonStatus As Boolean
Public WithEvents m_objUser As CUser
Attribute m_objUser.VB_VarHelpID = -1
Public m_retCode As enLogonReturnCode
Public m_blnRoleSelectionRequired As Boolean
Private m_blnChoose As Boolean
Private m_bStart As Boolean

Private Sub cmdCancel_Click()
    '-------------------------------
    '*** Generic Procedure Code
    '-------------------------------
1     On Error GoTo errHandler
2     Dim sModuleAndProcName As String
3     sModuleAndProcName = "frmLogon - cmdCancel_Click Function"

    
    If Timer1.Enabled Then
        Timer1.Enabled = False
    Else
        If cmdCancel.Caption = "Cancel" Then
            m_retCode = icFailure
            Unload Me
        ElseIf cmdCancel.Caption = "Log Off" Then
    '        m_bLogonStatus = False
            ClearTxt
            EnableButtons
            frmLogon.Show
            cmdCancel.Caption = "Cancel"
            StatusBar1.Panels(1).Text = "Enter UserId & Password ....."
            txtUserName.SetFocus
        End If
    End If
        
    m_bLogonStatus = False
        
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

Private Sub cmdCancel2_Click()
    m_objUser.LogonStatus = icFailure
    Unload Me
End Sub

Private Sub cmdChoose_Click()
    m_blnChoose = True
    lstRoles_DblClick
End Sub

Private Sub cmdLogon_Click()
'-------------------------------
'*** Generic Procedure Code
'-------------------------------
1     On Error GoTo errHandler
2     Dim sModuleAndProcName As String
3     sModuleAndProcName = "frmLogon - cmdLogon_Click Function"
  
    Dim sMessage As String
    Dim objfrm As frmChangePassword
    
    m_blnChoose = False
    
    ' Check user inputs (blank passwords are not allowed, if not then replace this code
    ' with the commented out code below)
    If txtUserName = "" Then
        StatusBar1.Panels(1).Text = "Please enter User Id !!!"
        txtPassword.Text = ""
        txtUserName.SetFocus
        Exit Sub
    End If
    If txtPassword = "" Then
        StatusBar1.Panels(1).Text = "Please enter Password !!!"
        txtPassword.Text = ""
        txtPassword.SetFocus
        Exit Sub
    End If
    
    '***    Attempt to Logon the User
    m_objUser.LogonCode = txtUserName
    m_objUser.Password = txtPassword
    
    MousePointer = vbHourglass
    
    m_retCode = m_objUser.Logon
  
    MousePointer = vbDefault
    
    Select Case m_retCode
    
    Case enLogonReturnCode.icFailure:
        m_bLogonStatus = False
        'StatusBar1.Panels(1).Text = "Logon Failed!!! Invalid logon/password combination ..."
        txtPassword = ""
        txtPassword.SetFocus
        Exit Sub
    Case enLogonReturnCode.icInvalidUserName:
        m_bLogonStatus = False
        txtPassword = ""
        'StatusBar1.Panels(1).Text = "Logon Failed!!! Invalid logon/password combination ..."
        txtPassword.SetFocus
        Exit Sub
    Case enLogonReturnCode.icSuccess:
        SaveSetting "Selcomm", "LogonCode", "LogonCode", txtUserName
        m_bLogonStatus = True
        If m_objUser.NumberOfRoles < 2 Or GetSetting("Selcomm", "Role", "Role", "") <> "" Then
            m_blnChoose = True
            Me.Hide
        End If
    Case enLogonReturnCode.icResetPassword:
        MsgBox m_objUser.LoginErrorMessage, vbInformation + vbOKOnly, "Reset Password"
        Set objfrm = New frmChangePassword
        Set objfrm.m_objUser = m_objUser
        objfrm.Show vbModal, Me
        
        If Not objfrm.blnCancel Then
            m_retCode = m_objUser.Logon
            If m_retCode = icSuccess Then
                m_bLogonStatus = True
                If m_objUser.NumberOfRoles < 2 Or GetSetting("Selcomm", "Role", "Role", "") <> "" Then
                    m_blnChoose = True
                    Me.Hide
                Else
                    Exit Sub
                End If
            End If
        Else
            m_bLogonStatus = False
            'StatusBar1.Panels(1).Text = m_objUser.LoginErrorMessage
            Exit Sub
        End If
    Case enLogonReturnCode.icBusinessUnitLoadFailure:
        m_bLogonStatus = False
        StatusBar1.Panels(1).Text = "Failed to load business units."
        txtPassword = ""
        txtPassword.SetFocus
        Exit Sub
    Case enLogonReturnCode.icConfigurationLoadFailure:
        m_bLogonStatus = False
        StatusBar1.Panels(1).Text = "Failed to load configurations."
        txtPassword = ""
        txtPassword.SetFocus
        Exit Sub
    Case enLogonReturnCode.icConnectionConfigLoadFailure:
        m_bLogonStatus = False
        StatusBar1.Panels(1).Text = "Failed to load connection configurations."
        txtPassword = ""
        txtPassword.SetFocus
        Exit Sub
    Case enLogonReturnCode.icRoleLoadFailure:
        m_bLogonStatus = False
        StatusBar1.Panels(1).Text = "Failed to load roles."
        txtPassword = ""
        txtPassword.SetFocus
        Exit Sub
    Case enLogonReturnCode.icDatabaseFailure:
        m_bLogonStatus = False
        StatusBar1.Panels(1).Text = "Database failure."
        txtPassword = ""
        txtPassword.SetFocus
        Exit Sub
    Case enLogonReturnCode.icLicenseConfigFailure:
        m_bLogonStatus = False
        StatusBar1.Panels(1).Text = "Licence error. Please contact support."
        Exit Sub
    Case enLogonReturnCode.icLicenseFailure:
        m_bLogonStatus = False
        StatusBar1.Panels(1).Text = "Max User Licence Limit. Please contact support."
        Exit Sub
    Case enLogonReturnCode.icLockOut
        m_bLogonStatus = False
        StatusBar1.Panels(1).Text = m_objUser.LoginErrorMessage
    End Select
Exit Sub

'------------------------------------------------------------
'*** Generic Error Handling Code (ensure Exit Proc is above)
'------------------------------------------------------------
errHandler:
  If Err.Number = 0 Then Exit Sub
    Me.MousePointer = vbDefault
    '***  custom error handler
    If Err.Number = 999 Then
        MsgBox "No Roles defined for this logon.", vbCritical, "No Roles Defined"
        Exit Sub
    End If

    If Err.Number = 998 Then
        MsgBox "Invalid Business Unit Code passed or no Business Units defined.", vbCritical, "Invalid Business Units"
        Exit Sub
    End If
    
    If Err.Number = 1003 Then
        MsgBox "No license defined.", vbCritical, "No license"
        Exit Sub
    End If
    
    If Err.Number = 1004 Then
        MsgBox "No more license available.", vbExclamation, "No license"
        Exit Sub
    End If
    
    If Err.Number = 1006 Then
        MsgBox Err.Description, vbCritical + vbOKOnly, "Invalid IP address"
        Exit Sub
    End If
    
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
    If Err.Number = (vbObjectError + 1300) Then       '***    a command line argument error
        m_bLogonStatus = False
        StatusBar1.Panels(1).Text = "Invalid command line parameters. "
        txtPassword = ""
        MsgBox Err.Description, vbCritical, "Invalid Command Line"
    Else
        DisplayError lErrNum, sErrSource, sErrDesc
    End If
    '*** OR raise the error to the calling procedure
    'Err.Raise lErrNum, sErrSource, sErrDesc
    '*** OR ignore the error and continue
    'Resume Next

    
End Sub



Private Sub ClearTxt()
    
    txtUserName.Text = ""
    txtPassword.Text = ""
    
End Sub

Private Sub EnableButtons()
'-------------------------------
'*** Generic Procedure Code
'-------------------------------
1     On Error GoTo errHandler
2     Dim sModuleAndProcName As String
3     sModuleAndProcName = "frmLogon - cmdCancel_Click Function"

    cmdLogon.Enabled = True
    cmdCancel.Enabled = True
    txtUserName.Enabled = True
    txtPassword.Enabled = True
    
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
    'DisplayError lErrNum, sErrSource, sErrDesc
  '*** OR raise the error to the calling procedure
    Err.Raise lErrNum, sErrSource, sErrDesc
  '*** OR ignore the error and continue
    'Resume Next
    
    
End Sub

Private Sub Form_Load()
    On Error Resume Next

    frmLogon.WindowState = vbNormal
    frmLogon.Move (Screen.Width - frmLogon.Width) / 2, (Screen.Height - frmLogon.Height) / 2
    StatusBar1.Panels(2).Text = m_objUser.ComputerName
    StatusBar1.Panels(2).Width = Me.TextWidth(StatusBar1.Panels(2).Text) + 150
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not m_blnChoose Then cmdCancel2_Click
End Sub

Private Sub lstRoles_DblClick()
    m_blnChoose = True
    m_objUser.SetRole (lstRoles.List(lstRoles.ListIndex))
End Sub

Private Sub lstRoles_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        m_blnChoose = True
        m_objUser.SetRole (lstRoles.List(lstRoles.ListIndex))
    End If
End Sub

Private Sub m_objUser_LoadConfigurationFinish()
    StatusBar1.Panels(1).Text = "Finished loading connection configurations."
End Sub

Private Sub m_objUser_LoadConfigurationStart()
    StatusBar1.Panels(1).Text = "Loading connection configurations."
End Sub

Private Sub m_objUser_LoadLicenseFinsish()
    StatusBar1.Panels(1).Text = "Finished loading license."
'    Unload Me
End Sub

Private Sub m_objUser_LoadLicenseStart()
    StatusBar1.Panels(1).Text = "Loading license."
End Sub

Private Sub m_objUser_LoadParentRoleFinish()
    StatusBar1.Panels(1).Text = "Finished loading parent roles."
    Unload Me
End Sub

Private Sub m_objUser_LoadParentRoleStart()
    StatusBar1.Panels(1).Text = "Loading parent roles."
End Sub

Private Sub m_objUser_LogonFinish()
    StatusBar1.Panels(1).Text = "Logon successful!"
End Sub

Private Sub m_objUser_LogonInvalidUserName()
    StatusBar1.Panels(1).Text = "Logon failed. Invalid User Id..."
End Sub

Private Sub m_objUser_LogonStart()
    StatusBar1.Panels(1).Text = "Attempting to Logon."
End Sub

Private Sub m_objUser_LoadRolesFinish()
    StatusBar1.Panels(1).Text = "Finished Loading Roles."
End Sub

Private Sub m_objUser_LoadRolesStart()
    StatusBar1.Panels(1).Text = "Loading Roles."
End Sub

Private Sub m_objUser_RoleSelectionFinish()
    StatusBar1.Panels(1).Text = "Role Selected."
    Frame1.Enabled = False
    Unload Me
End Sub

Private Sub m_objUser_RoleSelectionRequired()
    Dim rs As ADODB.Recordset
    
    StatusBar1.Panels(1).Text = "Choose Role."
    Me.Frame4.Visible = False
    Me.Frame1.Visible = True
    m_blnRoleSelectionRequired = True
    '***    load the list box
    Set rs = m_objUser.Roles
    While Not rs.EOF
        lstRoles.AddItem rs("role_narr")
        rs.MoveNext
    Wend
    lstRoles.ListIndex = 0
End Sub

Private Sub m_objUser_StatusBarText(strText As String)
    StatusBar1.Panels(1).Text = strText
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    
    StatusBar1.Panels(1) = "Automatic login in now."
    
    cmdLogon_Click
End Sub

Private Sub txtPassword_Click()
    If Timer1.Enabled Then Timer1.Enabled = False
End Sub

Private Sub txtUserName_Click()
    If Timer1.Enabled Then Timer1.Enabled = False
End Sub

Private Sub txtUserName_GotFocus()
    If Not m_bStart Then
        m_bStart = True
        If Len(Trim(txtUserName.Text)) > 0 Then
            txtPassword.SetFocus
        End If
    Else
        If Len(txtUserName) > 0 Then
            txtUserName.SelStart = 0
            txtUserName.SelLength = Len(txtUserName)
        End If
    End If
End Sub
