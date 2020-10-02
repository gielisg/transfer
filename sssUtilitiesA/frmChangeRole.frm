VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmChangeRole 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Role"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3990
   Icon            =   "frmChangeRole.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   3990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkDefault 
      Caption         =   "Set as default role."
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton cmdChangeRole 
      Caption         =   "&Select"
      Default         =   -1  'True
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Choose Role"
      Height          =   1695
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   3975
      Begin VB.ListBox lstRoles 
         Height          =   1035
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   3495
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   2775
      Width           =   3990
      _ExtentX        =   7038
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4313
            Text            =   "Choose new role."
            TextSave        =   "Choose new role."
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2646
            MinWidth        =   2646
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmChangeRole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Public WithEvents m_objUser As CUser
Attribute m_objUser.VB_VarHelpID = -1
Dim ClickSelect As Boolean
'
Private Sub chkDefault_Click()
    chkDefault_GotFocus
End Sub

Private Sub chkDefault_GotFocus()
    If chkDefault.Value = Checked Then
        StatusBar1.Panels(1) = "Set selected role as default."
    Else
        StatusBar1.Panels(1) = "Not set selected role as default."
    End If
End Sub

Private Sub cmdCancel_Click()
    m_objUser.RoleChanged = False
    Me.Hide
End Sub

Private Sub cmdCancel_GotFocus()
    StatusBar1.Panels(1) = "Cancel and exit."
End Sub

Private Sub cmdChangeRole_Click()
    m_objUser.SetRole (lstRoles.List(lstRoles.ListIndex))
    
    ClickSelect = True
    m_objUser.SetDefaultRole = (chkDefault.Value = Checked)
    Me.Hide
End Sub

Private Sub cmdChangeRole_GotFocus()
    StatusBar1.Panels(1) = "Select selected role."
End Sub

Private Sub Form_Load()
    On Error Resume Next

    frmChangeRole.WindowState = vbNormal
    frmChangeRole.Move (Screen.Width - frmChangeRole.Width) / 2, (Screen.Height - frmChangeRole.Height) / 2

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    m_objUser.RoleChanged = ClickSelect
End Sub

Private Sub lstRoles_DblClick()
    cmdChangeRole_Click
End Sub

Private Sub lstRoles_GotFocus()
    StatusBar1.Panels(1) = "Choose new role."
End Sub

Private Sub lstRoles_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then m_objUser.SetRole (lstRoles.List(lstRoles.ListIndex))
'    Me.Hide
End Sub
Public Sub Populate()
    Dim strCurrentRole As String
'    Dim rs As ADODB.Recordset
    Dim idx As Integer
    Dim kdx As Integer
    Dim colRoleList As Collection
    Dim RoleList As sssRoleList
    
    strCurrentRole = m_objUser.CurrentRole
    StatusBar1.Panels(2).Text = strCurrentRole
    '***    load the list box
'    Set rs = m_objUser.Roles
'    While Not rs.EOF
'        idx = idx + 1
'        lstRoles.AddItem rs("role_narr")
'        If strCurrentRole = rs("role_narr") Then kdx = idx
'        rs.MoveNext
'    Wend
    Set colRoleList = m_objUser.RoleLists
    For idx = 1 To colRoleList.Count
        Set RoleList = colRoleList(idx)
            lstRoles.AddItem RoleList.RoleNarr
        If strCurrentRole = RoleList.RoleNarr Then
            kdx = idx
        End If
    Next
    lstRoles.ListIndex = kdx - 1
End Sub
