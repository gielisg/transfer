VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmChangeBusinessUnit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Business Unit"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3990
   Icon            =   "frmChangeBusinessUnit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   3990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Choose Business Unit"
      Height          =   1695
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   3975
      Begin VB.ListBox lstBusinessUnits 
         Height          =   1035
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   3495
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   1770
      Width           =   3990
      _ExtentX        =   7038
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4313
            Text            =   "Choose new business unit."
            TextSave        =   "Choose new business unit."
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2646
            MinWidth        =   2646
         EndProperty
      EndProperty
   End
   Begin VB.Label lblStatus 
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   1680
      Width           =   3975
   End
End
Attribute VB_Name = "frmChangeBusinessUnit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Public WithEvents m_objUser As CUser
Attribute m_objUser.VB_VarHelpID = -1
Public m_varKeys As Variant

Private Sub Form_Load()
    On Error Resume Next

    frmChangeBusinessUnit.WindowState = vbNormal
    frmChangeBusinessUnit.Move (Screen.Width - frmChangeBusinessUnit.Width) / 2, (Screen.Height - frmChangeBusinessUnit.Height) / 2

End Sub

Private Sub lstBusinessUnits_DblClick()
    m_objUser.SetBusinessUnit (m_varKeys(lstBusinessUnits.ListIndex))
    Me.Hide
End Sub

Private Sub lstBusinessUnits_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then m_objUser.SetBusinessUnit (m_varKeys(lstBusinessUnits.ListIndex))
    Me.Hide
End Sub
Public Sub Populate()

    Dim arr As Variant
    Dim idx As Integer
    Dim kdx As Integer
    
    StatusBar1.Panels(2).Text = m_objUser.CurrentBusinessUnit
    
    '***    load the list box
    arr = m_objUser.BusinessUnits.Items
    m_varKeys = m_objUser.BusinessUnits.Keys
    
    For idx = 0 To m_objUser.BusinessUnits.Count - 1
        If m_objUser.CurrentBusinessUnitCode = arr(idx) Then
            kdx = idx
        End If
        lstBusinessUnits.AddItem arr(idx)
    Next
    
    lstBusinessUnits.ListIndex = kdx
    
End Sub
