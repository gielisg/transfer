VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmOverAmount 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5790
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   5175
   Icon            =   "frmOverAmount.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   5175
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraService 
      Caption         =   "Service"
      Height          =   855
      Left            =   120
      TabIndex        =   29
      Top             =   120
      Width           =   4935
      Begin VB.CommandButton cmdService 
         Caption         =   "..."
         Height          =   285
         Left            =   4560
         TabIndex        =   1
         Top             =   360
         Width           =   285
      End
      Begin VB.ComboBox cmbService 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Number"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   645
      Left            =   0
      ScaleHeight     =   645
      ScaleWidth      =   5175
      TabIndex        =   28
      Top             =   4770
      Width           =   5175
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Default         =   -1  'True
         Height          =   375
         Left            =   1920
         TabIndex        =   17
         Top             =   120
         Width           =   1440
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   375
         Left            =   3600
         TabIndex        =   18
         Top             =   120
         Width           =   1440
      End
   End
   Begin VB.Frame fraDetail 
      Caption         =   "Details"
      Height          =   1455
      Left            =   120
      TabIndex        =   23
      Top             =   3240
      Width           =   4935
      Begin VB.TextBox txtMarkup 
         Height          =   285
         Left            =   3600
         MaxLength       =   7
         TabIndex        =   12
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton optEndDate 
         Caption         =   "End Date"
         Height          =   255
         Left            =   2520
         TabIndex        =   14
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton optOngoing 
         Caption         =   "On-going"
         Height          =   255
         Left            =   2520
         TabIndex        =   15
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox txtAmount 
         Height          =   285
         Left            =   1080
         MaxLength       =   17
         TabIndex        =   11
         Top             =   360
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker dtpStart 
         Height          =   285
         Left            =   1080
         TabIndex        =   13
         Top             =   720
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Format          =   57606145
         CurrentDate     =   38272
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   285
         Left            =   3600
         TabIndex        =   16
         Top             =   720
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Format          =   57606145
         CurrentDate     =   38272
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Markup Ratio"
         Height          =   255
         Left            =   2520
         TabIndex        =   26
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Start Date"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Amount"
         Height          =   255
         Left            =   360
         TabIndex        =   24
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame fraPackageCharge 
      Caption         =   "Package/Charge"
      Height          =   2055
      Left            =   120
      TabIndex        =   20
      Top             =   1080
      Width           =   4935
      Begin VB.TextBox txtBillNarrative 
         Height          =   285
         Left            =   1080
         MaxLength       =   32
         TabIndex        =   10
         Top             =   1560
         Width           =   3735
      End
      Begin VB.ComboBox cmbCharge 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1200
         Width           =   3375
      End
      Begin VB.ComboBox cmbOption 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   720
         Width           =   3735
      End
      Begin VB.ComboBox cmbPackage 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   3375
      End
      Begin VB.CommandButton cmdPackage 
         Caption         =   "..."
         Height          =   285
         Left            =   4560
         TabIndex        =   4
         Top             =   360
         Width           =   285
      End
      Begin VB.TextBox txtPackage 
         Height          =   285
         Left            =   1080
         TabIndex        =   3
         Top             =   360
         Width           =   3735
      End
      Begin VB.CommandButton cmdCharge 
         Caption         =   "..."
         Height          =   285
         Left            =   4560
         TabIndex        =   9
         Top             =   1200
         Width           =   285
      End
      Begin VB.TextBox txtCharge 
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1200
         Width           =   3735
      End
      Begin VB.TextBox txtOption 
         Height          =   285
         Left            =   1080
         TabIndex        =   6
         Top             =   720
         Width           =   3735
      End
      Begin VB.Label Label8 
         Caption         =   "Bill Narrative"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Option"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Charge"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Package"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   360
         Width           =   735
      End
   End
   Begin MSComctlLib.StatusBar ssb 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   19
      Top             =   5415
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5521
            MinWidth        =   176
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   176
            MinWidth        =   176
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   176
            MinWidth        =   176
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            Object.Width           =   1693
            MinWidth        =   176
            TextSave        =   "14/06/2013"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            Object.Width           =   1376
            MinWidth        =   176
            TextSave        =   "11:08 AM"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
   End
End
Attribute VB_Name = "frmOverAmount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public objDatabase As clsDatabase
Public objCommonDB As Object
Dim objRS As ADODB.Recordset
Dim colPackageCode As Collection
Dim colOption As Collection
Dim colChargeCode As Collection
Dim colChangeAmount As Collection
Dim colMarkup As Collection

Dim strContentsBeforePaste As String
Dim blnLoading As Boolean
Dim LastPackageIndex As Integer
Dim LastOptionIndex As Integer
Dim LastChargeIndex As Integer

Dim DefaultPackage As Integer
Dim DefaultOption As Integer

Public Enum enuMode
    typNew = 0
    typUpdate = 1
End Enum

Dim objfrm As frmAvailCharge

Public Mode As enuMode
Public ServiceNumber As String

Dim colSpCnRef As Collection

Public OverrideID As Long
Public PackageNarr As String
Public OptionNarr As String
Public ChargeNarr As String
Public BillBarrative As String
Public Amount As Double
Public Markup As Double
Public StartDate As Date
Public EndDate As Date

Public blnCancel As Boolean

Private Const PackageChangeObject As String = "sssChangePackageUI.Application"
Private Const FixRound As Double = 0.000000000001

Private Const UpdateCaption As String = "&Update"
Private Const CloseCaption As String = "&Close"
Private Const CancelCaption As String = "&Cancel"

Dim WithEvents objfrmService As frmOtherService
Attribute objfrmService.VB_VarHelpID = -1

Private Function DistTitle() As String
    If Mode = typNew Then
        DistTitle = "New Override"
    Else
        DistTitle = "Update Override"
    End If
End Function

Public Function Populate() As Boolean
'***********************************************************************************************
'Purpose    : populate form data
'Inputs     :
'Outputs    : None.
'Version: $SSSVersion:
'Created on : 13/03/2001
'Created by : Patrick
'***********************************************************************************************

'-------------------------------
'*** Generic Procedure Code
'-------------------------------
    On Error GoTo errHandler
    Dim sModuleAndProcName As String
    sModuleAndProcName = "frmOverAmount - Populate Procedure"
    
'-------------------------------
'*** Specific Procedure Code
'-------------------------------
    
    blnLoading = True
    If Not SetData Then Exit Function
    
    If Mode = typNew Then
        LoadPackage
       
    End If
    
    blnLoading = False
    Populate = True
    
Exit Function
'------------------------------------------------------------
'*** Generic Error Handling Code (ensure Exit Proc is above)
'------------------------------------------------------------
errHandler:
    If Err.Number = 0 Then Exit Function
    
    Dim lErrNum As Long, sErrDesc As String, sErrSource As String
    lErrNum = Err.Number
    sErrDesc = Err.Description
    sErrSource = "[" & App.Title & " - " & sModuleAndProcName & " - Build " & App.Major & "." & App.Minor & "." & App.Revision
    If Erl > 0 Then sErrSource = sErrSource & " - Line No " & CStr(Erl)
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
'    Err.Raise lErrNum, sErrSource, sErrDesc
    '*** OR ignore the error and continue
    'Resume Next
End Function

Private Function SetData() As Boolean
'***********************************************************************************************
'Purpose    : Set data up for invoice option form
'Inputs     :
'Outputs    : None.
'Version: $SSSVersion:
'Created on : 13/03/2001
'Created by : Patrick
'***********************************************************************************************

'-------------------------------
'*** Generic Procedure Code
'-------------------------------
    On Error GoTo errHandler
    Dim sModuleAndProcName As String
    sModuleAndProcName = "frmOverAmount - SetData Procedure"
    
'-------------------------------
'*** Specific Procedure Code
'-------------------------------
    
    If ServiceNumber <> "" Then
        Me.Caption = DistTitle & " [Service: " & ServiceNumber & "]"
    Else
        Me.Caption = DistTitle & " [Account: " & objDatabase.objApplication.ContactCode & "]"
    End If
    
    If Mode = typNew Then
        txtPackage.Visible = False
        txtOption.Visible = False
        txtCharge.Visible = False
        txtMarkup = "0"
        dtpStart.Value = Date
        dtpEnd.Value = DateAdd("m", 3, Date)
        optOngoing.Value = True
        If objDatabase.objApplication.Mode = typService Then
            fraService.Visible = False
            fraPackageCharge.Top = fraPackageCharge.Top - 960
            fraDetail.Top = fraDetail.Top - 960
            Height = Height - 960
        Else
            cmbService.AddItem "<All Services>"
            cmbService.ListIndex = 0
        End If
    Else
        cmdSave.Caption = UpdateCaption
        cmbPackage.Visible = False
        cmbOption.Visible = False
        cmbCharge.Visible = False
        cmdPackage.Visible = False
        cmdCharge.Visible = False
        txtPackage.Enabled = False
        txtOption.Enabled = False
        txtCharge.Enabled = False
        txtPackage = PackageNarr
        txtOption = OptionNarr
        txtCharge = ChargeNarr
        txtBillNarrative = BillBarrative
        
        txtAmount = Amount
        txtMarkup = Markup
        
        dtpStart.Value = StartDate
        If EndDate = #1/1/9999# Then
            optOngoing.Value = True
            dtpEnd.Value = DateAdd("m", 3, StartDate)
        Else
            optEndDate.Value = True
            dtpEnd.Value = EndDate
        End If
        
        fraService.Visible = False
        fraPackageCharge.Top = fraPackageCharge.Top - 960
        fraDetail.Top = fraDetail.Top - 960
        Height = Height - 960
    End If
    
    LastPackageIndex = -1
    cmdSave.Enabled = False
    
    With ssb
        .Panels(2).Text = objDatabase.objApplication.User.logoncode
        .Panels(3).Text = objDatabase.objApplication.User.CurrentBusinessUnit
        .Panels(2).Width = Me.TextWidth(.Panels(2).Text) + 150
        .Panels(3).Width = Me.TextWidth(.Panels(3).Text) + 150
    End With
        
    SetData = True
Exit Function
'------------------------------------------------------------
'*** Generic Error Handling Code (ensure Exit Proc is above)
'------------------------------------------------------------
errHandler:
    If Err.Number = 0 Then Exit Function
    
    Dim lErrNum As Long, sErrDesc As String, sErrSource As String
    lErrNum = Err.Number
    sErrDesc = Err.Description
    sErrSource = "[" & App.Title & " - " & sModuleAndProcName & " - Build " & App.Major & "." & App.Minor & "." & App.Revision
    If Erl > 0 Then sErrSource = sErrSource & " - Line No " & CStr(Erl)
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
'    Err.Raise lErrNum, sErrSource, sErrDesc
    '*** OR ignore the error and continue
    'Resume Next

End Function

Private Function LoadPackage() As Boolean
'***********************************************************************************************
'Purpose    :
'Outputs    : None.
'Version: $SSSVersion:
'Created on : 07/02/2001
'Created by : Patrick
'Modification:
'***********************************************************************************************

'-------------------------------
'*** Generic Procedure Code
'-------------------------------
    On Error GoTo errHandler
    Dim sModuleAndProcName As String
    sModuleAndProcName = "frmOverAmount - LoadPackage"

'-------------------------------
'*** Specific Procedure Code
'-------------------------------
    Dim idxPackage As Integer
    Dim SpCnRef As Long
    
    If objDatabase.objApplication.Mode = typContact Then
        If cmbService.ListIndex > 0 Then
            SpCnRef = colSpCnRef(cmbService.ListIndex)
        End If
    End If
    
    Set colPackageCode = New Collection
    
    Set objRS = objDatabase.GetAccountPackage(SpCnRef)
                
    DefaultPackage = 0
    DefaultOption = 0
                    
    While Not objRS.EOF
        If objRS(0) = 0 Then
            cmbPackage.AddItem "[" & NullText(objRS(1)) & "] " & NullText(objRS(2))
            colPackageCode.Add NullText(objRS(1))
            If NullText(objRS(3)) = "Y" Then
                DefaultPackage = cmbPackage.ListCount - 1
                DefaultOption = NullText(objRS(4))
            End If
        End If
        
        objRS.MoveNext
    Wend
        
    If cmbPackage.ListCount > 0 Then
        cmbPackage.ListIndex = DefaultPackage
    Else
        cmdPackage_Click
    End If
    
    Set objRS = Nothing
    
    LoadPackage = True
    
Exit Function
'------------------------------------------------------------
'*** Generic Error Handling Code (ensure Exit Proc is above)
'------------------------------------------------------------
errHandler:
    Me.MousePointer = vbDefault
    If Err.Number = 0 Then Exit Function
    
    LoadPackage = False
    Dim lErrNum As Long, sErrDesc As String, sErrSource As String
    lErrNum = Err.Number
    sErrDesc = Err.Description
    sErrSource = "[" & App.Title & " - " & sModuleAndProcName & " - Build " & App.Major & "." & App.Minor & "." & App.Revision
    If Erl > 0 Then sErrSource = sErrSource & " - Line No " & CStr(Erl)
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
    
End Function

Public Sub DisplayOptions()
'***********************************************************************************************
'Purpose    : name
'Inputs     :
'Outputs    : None.
'Version: $SSSVersion:
'Created on : 07/02/2001
'Created by : Patrick
'***********************************************************************************************

'-------------------------------
'*** Generic Procedure Code
'-------------------------------
    On Error GoTo errHandler
    Dim sModuleAndProcName As String
    sModuleAndProcName = "frmOverAmount - DisplayOptions"
    
'-------------------------------
'*** Specific Procedure Code
'-------------------------------
    Dim idx  As Integer
    Dim idxOption As Integer
    
    cmbOption.Clear
    Set colOption = New Collection
    
    cmbOption.AddItem "<All Options>"
    
    Set objRS = objCommonDB.GetOptions(colPackageCode(cmbPackage.ListIndex + 1))
    
    While Not objRS.EOF
        If objRS(0) = 0 Then
            cmbOption.AddItem Trim$(objRS(2))
            colOption.Add Trim$(objRS(1))
        End If
        objRS.MoveNext
    Wend
    
    LastOptionIndex = -1
    
    If cmbPackage.ListIndex = DefaultPackage Then
        If DefaultOption <> 0 Then
            For idx = 1 To colOption.Count
                If colOption(idx) = DefaultOption Then
                    idxOption = idx
                    Exit For
                End If
            Next
        End If
    End If
    cmbOption.ListIndex = idxOption
    
    Set objRS = Nothing
  
Exit Sub
'------------------------------------------------------------
'*** Generic Error Handling Code (ensure Exit Proc is above)
'------------------------------------------------------------
errHandler:
    If Err.Number = 0 Then Exit Sub
    
    Dim lErrNum As Long, sErrDesc As String, sErrSource As String
    lErrNum = Err.Number
    sErrDesc = Err.Description
    sErrSource = "[" & App.Title & " - " & sModuleAndProcName & " - Build " & App.Major & "." & App.Minor & "." & App.Revision
    If Erl > 0 Then sErrSource = sErrSource & " - Line No " & CStr(Erl)
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

Public Sub DisplayCharge()
'***********************************************************************************************
'Purpose    : name
'Inputs     :
'Outputs    : None.
'Version: $SSSVersion:
'Created on : 07/02/2001
'Created by : Patrick
'***********************************************************************************************

'-------------------------------
'*** Generic Procedure Code
'-------------------------------
    On Error GoTo errHandler
    Dim sModuleAndProcName As String
    sModuleAndProcName = "frmOverAmount - DisplayCharge"
    
'-------------------------------
'*** Specific Procedure Code
'-------------------------------
    Dim idx  As Long
    Dim PackageOption As Long
    
    If cmbOption.ListIndex > 0 Then
        PackageOption = colOption(cmbOption.ListIndex)
    End If
    
    cmbCharge.Clear
    
    Set colChargeCode = New Collection
    Set colChangeAmount = New Collection
    Set colMarkup = New Collection
    
    Set objRS = objDatabase.GetPackageCharge(colPackageCode(cmbPackage.ListIndex + 1), PackageOption)
    
    While Not objRS.EOF
        cmbCharge.AddItem "[" & Trim$(objRS(0)) & "] " & Trim$(objRS(1))
        colChargeCode.Add Trim$(objRS(0))
        colChangeAmount.Add ""
        colMarkup.Add ""
        objRS.MoveNext
    Wend
    
    If cmbCharge.ListCount > 0 Then
        LastChargeIndex = -1
        cmbCharge.ListIndex = 0
'    Else
'        If MsgBox("There is no package charge. Would you like to load network charge?", vbQuestion + vbYesNo, DistTitle) = vbYes Then
'            cmdCharge_Click
'        End If
    End If
    
    Set objRS = Nothing
  
Exit Sub
'------------------------------------------------------------
'*** Generic Error Handling Code (ensure Exit Proc is above)
'------------------------------------------------------------
errHandler:
    If Err.Number = 0 Then Exit Sub
    
    Dim lErrNum As Long, sErrDesc As String, sErrSource As String
    lErrNum = Err.Number
    sErrDesc = Err.Description
    sErrSource = "[" & App.Title & " - " & sModuleAndProcName & " - Build " & App.Major & "." & App.Minor & "." & App.Revision
    If Erl > 0 Then sErrSource = sErrSource & " - Line No " & CStr(Erl)
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
        
Private Function SaveChargeOverride() As Boolean
'***********************************************************************************************
'Purpose    :
'Outputs    : None.
'Version: $SSSVersion:
'Created on : 07/02/2001
'Created by : Patrick
'Modification:
'***********************************************************************************************

'-------------------------------
'*** Generic Procedure Code
'-------------------------------
    On Error GoTo errHandler
    Dim sModuleAndProcName As String
    sModuleAndProcName = "frmOverAmount - SaveChargeOverride"

'-------------------------------
'*** Specific Procedure Code
'-------------------------------
    Dim PackageOption As Long
    Dim EndDate As Date
    Dim SpCnRef As Long
    
    If objDatabase.objApplication.Mode = typContact Then
        If cmbService.ListIndex > 0 Then
            SpCnRef = colSpCnRef(cmbService.ListIndex)
        End If
    End If
    
    If cmbOption.ListIndex > 0 Then
        PackageOption = colOption(cmbOption.ListIndex)
    End If
    
    If optEndDate Then
        EndDate = dtpEnd.Value
    Else
        EndDate = #1/1/9999#
    End If
    
    Set objRS = objDatabase.GetSaveChargeOverride(colPackageCode(cmbPackage.ListIndex + 1), _
        PackageOption, colChargeCode(cmbCharge.ListIndex + 1), dtpStart.Value, EndDate, _
        Val(txtAmount), Val(txtMarkup), SpCnRef, Trim(txtBillNarrative))
                
    If objRS(0) = 0 Then
        SaveChargeOverride = True
        OverrideID = CLng(objRS(1))
    Else
        MsgBox "Unable to save charge override." & vbNewLine & _
            Trim(objRS(objRS.Fields.Count - 1)), vbCritical + vbOKOnly, DistTitle
    End If
        
    Set objRS = Nothing
        
Exit Function
'------------------------------------------------------------
'*** Generic Error Handling Code (ensure Exit Proc is above)
'------------------------------------------------------------
errHandler:
    Me.MousePointer = vbDefault
    If Err.Number = 0 Then Exit Function
    
    Dim lErrNum As Long, sErrDesc As String, sErrSource As String
    lErrNum = Err.Number
    sErrDesc = Err.Description
    sErrSource = "[" & App.Title & " - " & sModuleAndProcName & " - Build " & App.Major & "." & App.Minor & "." & App.Revision
    If Erl > 0 Then sErrSource = sErrSource & " - Line No " & CStr(Erl)
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
    
End Function

Private Function UpdateChargeOverride() As Boolean
'***********************************************************************************************
'Purpose    :
'Outputs    : None.
'Version: $SSSVersion:
'Created on : 07/02/2001
'Created by : Patrick
'Modification:
'***********************************************************************************************

'-------------------------------
'*** Generic Procedure Code
'-------------------------------
    On Error GoTo errHandler
    Dim sModuleAndProcName As String
    sModuleAndProcName = "frmOverAmount - UpdateChargeOverride"

'-------------------------------
'*** Specific Procedure Code
'-------------------------------
    Dim EndDate As Date
        
    If optEndDate Then
        EndDate = dtpEnd.Value
    Else
        EndDate = #1/1/9999#
    End If
    
    Set objRS = objDatabase.GetUpdateChargeOverride(OverrideID, dtpStart.Value, _
        EndDate, Val(txtAmount), Val(txtMarkup), Trim(txtBillNarrative))
                
    If objRS(0) = 0 Then
        UpdateChargeOverride = True
    Else
        MsgBox "Unable to update charge override." & vbNewLine & _
            Trim(objRS(objRS.Fields.Count - 1)), vbCritical + vbOKOnly, DistTitle
    End If
        
    Set objRS = Nothing
        
Exit Function
'------------------------------------------------------------
'*** Generic Error Handling Code (ensure Exit Proc is above)
'------------------------------------------------------------
errHandler:
    Me.MousePointer = vbDefault
    If Err.Number = 0 Then Exit Function
    
    Dim lErrNum As Long, sErrDesc As String, sErrSource As String
    lErrNum = Err.Number
    sErrDesc = Err.Description
    sErrSource = "[" & App.Title & " - " & sModuleAndProcName & " - Build " & App.Major & "." & App.Minor & "." & App.Revision
    If Erl > 0 Then sErrSource = sErrSource & " - Line No " & CStr(Erl)
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
    
End Function

Private Function Validate() As Boolean
    If Mode = typNew Then
        If cmbPackage.ListCount = 0 Then
            MsgBox "You must select a package.", vbExclamation + vbOKOnly, DistTitle
            cmdPackage.SetFocus
            Exit Function
        End If
        
        If cmbCharge.ListCount = 0 Then
            MsgBox "You must select a charge.", vbExclamation + vbOKOnly, DistTitle
            cmdCharge.SetFocus
            Exit Function
        End If
    End If
    
    If Not IsNumeric(txtAmount) Then
        MsgBox "Invalid charge amount.", vbExclamation + vbOKOnly, DistTitle
        txtAmount.SetFocus
        Exit Function
    Else
        If PropRound(txtAmount, 6) <> Val(txtAmount) Then
            MsgBox "Invalid charge amount.", vbExclamation + vbOKOnly, DistTitle
            txtAmount.SetFocus
            Exit Function
        End If
        
        If txtAmount < -999999.999999 Or txtAmount > 999999.999999 Then
            MsgBox "Invalid charge amount.", vbExclamation + vbOKOnly, DistTitle
            txtAmount.SetFocus
            Exit Function
        End If
    End If
    
    If Trim(txtMarkup) <> "" Then
        If Not IsNumeric(txtMarkup) Then
            MsgBox "Invalid markup ratio.", vbExclamation + vbOKOnly, DistTitle
            txtMarkup.SetFocus
            Exit Function
        Else
            If PropRound(txtMarkup, 4) <> Val(txtMarkup) Then
                MsgBox "Invalid markup ratio.", vbExclamation + vbOKOnly, DistTitle
                txtMarkup.SetFocus
                Exit Function
            End If
            
            If txtMarkup > 99.9999 Then
                MsgBox "Invalid markup ratio.", vbExclamation + vbOKOnly, DistTitle
                txtMarkup.SetFocus
                Exit Function
            End If
        End If
    End If
    Validate = True
End Function

Private Function NullText(str As Variant) As String
    If IsNull(str) Then
        NullText = ""
    Else
        NullText = Trim$(str)
    End If
End Function

Private Function PropRound(Number As Double, NumDigitsAfterDecimal As Long) As Double
    ' this function is not bullet proof function, but enough for handle 2 digits after decimal point
    PropRound = Round(Number + FixRound, NumDigitsAfterDecimal)
End Function

Private Sub cmbCharge_Click()
    If cmbCharge.ListIndex = LastChargeIndex Then Exit Sub
    
    LastChargeIndex = cmbCharge.ListIndex
    txtAmount = colChangeAmount(cmbCharge.ListIndex + 1)
    txtMarkup = colMarkup(cmbCharge.ListIndex + 1)
    
    cmdSave.Enabled = False
    cmdClose.Caption = CloseCaption
End Sub

Private Sub cmbCharge_GotFocus()
    ssb.Panels(1) = "Charge."
End Sub

Private Sub cmbOption_Click()
    If cmbOption.ListIndex = LastOptionIndex Then Exit Sub
    LastOptionIndex = cmbOption.ListIndex
    
    DisplayCharge
End Sub

Private Sub cmbOption_GotFocus()
    ssb.Panels(1) = "Option."
End Sub

Private Sub cmbPackage_Click()
    If cmbPackage.ListIndex = LastPackageIndex Then Exit Sub
    
    LastPackageIndex = cmbPackage.ListIndex
    
    DisplayOptions
End Sub

Private Sub cmbPackage_GotFocus()
    ssb.Panels(1) = "Package."
End Sub

Private Sub cmbService_GotFocus()
    ssb.Panels(1) = "Service."
End Sub

Private Sub cmdCharge_GotFocus()
    ssb.Panels(1) = "Select charge."
End Sub

Private Sub cmdClose_Click()
    UnloadForm
End Sub

Private Sub cmdCharge_Click()
    
    Dim PackageOption As Long
    

    If objfrm Is Nothing Then
        Set objfrm = New frmAvailCharge

        Set objfrm.objCommonDB = objCommonDB
        objfrm.PackageCode = colPackageCode(cmbPackage.ListIndex + 1)

        If cmbOption.ListIndex > 0 Then
            PackageOption = cmbOption.ListIndex
        End If

         objfrm.PackageOption = PackageOption

        If objfrm.Populate(objDatabase) Then
            objfrm.Show vbModal, Me
        End If
    Else
        objfrm.blnCancelled = True
        objfrm.Show vbModal, Me
    End If

    If Not objfrm.blnCancelled Then
        SelectCharge objfrm.SelectedChargeCode, objfrm.SelectedChargeNarr, objfrm.SelectedAmount, objfrm.SelectedMarkup
    End If
    
End Sub

'Private Sub SetDate()
'    Set objRS = objDatabase.GetChargeLastUsed(ChargeCode)
'
'    If objRS(0) = 0 Then
'        If Not IsNull(objRS(1)) Then
'            dtpStart.MinDate = DateAdd("d", 1, CDate(objRS(1)))
'        Else
'            dtpStart.MinDate = VB_NULL_DATE
'        End If
'    End If
'End Sub

Private Sub cmdClose_GotFocus()
    ssb.Panels(1) = "Exit."
End Sub

Private Sub cmdPackage_Click()
    Dim objPackage As Object
    
    Set objPackage = CreateObject(PackageChangeObject)
        
    With objPackage
        .ParentForm = Me
        .User = objDatabase.objApplication.User
        .Arguments = "/Mode=3"
        .SpCnRef = 0
        .ServiceType = "ALL"
        .ContactCode = objDatabase.objApplication.ContactCode
        
        If Not .Start Then
            MsgBox "Could not load the package change window", vbInformation + vbOKOnly, DistTitle
        ElseIf .Canceled = False Then
            SelectPackage .PackageCode, .PackageNarr
        End If
    End With
End Sub

Private Sub SelectPackage(PackageCode As String, PackageNarr As String)
    Dim idx As Integer
    Dim blnFound As Boolean
    
    For idx = 1 To colPackageCode.Count
        If PackageCode = colPackageCode(idx) Then
            blnFound = True
            Exit For
        End If
    Next
    
    If blnFound Then
        cmbPackage.ListIndex = idx - 1
    Else
        cmbPackage.AddItem "[" & PackageCode & "] " & PackageNarr
        colPackageCode.Add PackageCode
        cmbPackage.ListIndex = cmbPackage.ListCount - 1
    End If
End Sub

Private Sub SelectCharge(ChargeCode As String, ChargeNarr As String, ChargeAmount As Variant, Markup As Variant)
    Dim idx As Integer
    Dim blnFound As Boolean
    
    For idx = 1 To colChargeCode.Count
        If ChargeCode = colChargeCode(idx) Then
            blnFound = True
            Exit For
        End If
    Next
    
    If blnFound Then
        cmbCharge.ListIndex = idx - 1
    Else
        cmbCharge.AddItem "[" & ChargeCode & "] " & ChargeNarr
        colChargeCode.Add ChargeCode
        colChangeAmount.Add ChargeAmount
        colMarkup.Add Markup
        cmbCharge.ListIndex = cmbCharge.ListCount - 1
    End If
End Sub

Private Sub UnloadForm()
    Set objRS = Nothing
    Set objDatabase = Nothing
    Set objCommonDB = Nothing
    Unload Me
End Sub

Private Sub cmdPackage_GotFocus()
    ssb.Panels(1) = "Select package."
End Sub

Private Sub cmdSave_Click()
    If Not Validate Then Exit Sub
            
    If Mode = typNew Then
        If SaveChargeOverride Then
            UnloadForm
        End If
    Else
        If UpdateChargeOverride Then
            UnloadForm
        End If
    End If
End Sub

Private Sub cmdSave_GotFocus()
    If Mode = typNew Then
        ssb.Panels(1) = "Save."
    Else
        ssb.Panels(1) = "Update."
    End If
End Sub

Private Sub cmdService_Click()
    SelectOtherService
End Sub

Public Sub SelectOtherService()
    If objfrmService Is Nothing Then
    
        Set objfrmService = New frmOtherService
                
        If objfrmService.Load(objDatabase) Then
            objfrmService.Show vbModal, Me
        Else
            Set objfrmService = Nothing
        End If
    Else
        objfrmService.Show vbModal, Me
    End If
End Sub

Private Sub cmdService_GotFocus()
    ssb.Panels(1) = "Select service."
End Sub

Private Sub dtpEnd_Change()
    If Mode = typUpdate Then CanSave
End Sub

Private Sub dtpEnd_GotFocus()
    ssb.Panels(1) = "End date."
End Sub

Private Sub dtpStart_Change()
    dtpEnd.MinDate = dtpStart.Value
    If Mode = typUpdate Then CanSave
End Sub

Private Sub dtpStart_GotFocus()
    ssb.Panels(1) = "Start date."
End Sub

Private Sub mnuClose_Click()
    cmdClose_Click
End Sub

Private Sub mnuFile_Click()
    mnuSave.Caption = cmdSave.Caption
    mnuSave.Enabled = cmdSave.Enabled
    
    mnuClose.Caption = cmdClose.Caption
End Sub

Private Sub mnuSave_Click()
    cmdSave_Click
End Sub

Private Sub objfrmService_SelectClick(SpCnRef As Long, PhoneNumber As String)
    Dim idx As Integer
    Dim blnExist As Boolean
    
    If colSpCnRef Is Nothing Then
        Set colSpCnRef = New Collection
    End If
    
    For idx = 1 To colSpCnRef.Count
        If colSpCnRef(idx) = SpCnRef Then
            blnExist = True
            Exit For
        End If
    Next
    
    If Not blnExist Then
        colSpCnRef.Add SpCnRef
        cmbService.AddItem PhoneNumber
        cmbService.ListIndex = cmbService.ListCount - 1
    Else
        cmbService.ListIndex = idx
    End If
End Sub

Private Sub optEndDate_Click()
    dtpEnd.Enabled = optEndDate.Value
    If Mode = typUpdate Then CanSave
    optEndDate_GotFocus
End Sub

Private Sub optEndDate_GotFocus()
    ssb.Panels(1) = "Has end date."
End Sub

Private Sub optOngoing_Click()
    dtpEnd.Enabled = Not optOngoing.Value
    If Mode = typUpdate Then CanSave
    optOngoing_GotFocus
End Sub

Private Sub optOngoing_GotFocus()
    ssb.Panels(1) = "On-going charge override."
End Sub

Private Sub txtAmount_Change()
    If Trim(txtAmount) <> "" Then CanSave
End Sub

Private Sub txtAmount_GotFocus()
    ssb.Panels(1) = "Amount."
End Sub

Private Sub txtAmount_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 3          'Copy - Ctrl+C
        Case 8          'Backspace
        Case 22         'Paste - Ctrl+V
            strContentsBeforePaste = txtAmount
        Case 24         'Cut - Ctrl+X
        Case 45         '-
            If InStr(Me.ActiveControl.Text, "-") > 0 Then KeyAscii = 0
        Case 46         '.
            If InStr(Me.ActiveControl.Text, ".") > 0 Then KeyAscii = 0
        Case 48 To 57   '0-9
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtAmount_KeyUp(KeyCode As Integer, Shift As Integer)
    HandleTextBoxKeyUp KeyCode
End Sub

Private Sub txtAmount_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    HandleTextBoxMouseDown Button
End Sub

Private Sub txtAmount_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    HandleTextBoxMouseUp Button
End Sub
Private Sub HandleTextBoxKeyUp(KeyCode As Integer)
    On Error Resume Next
    If KeyCode = 86 Then  'paste - Ctrl+V
        'for "$1", "9e9", "1-" isnumeric returns true
        'and cdbl("123e123123") gets an ovefloe error
        If Not IsNumeric(Me.ActiveControl.Text) Then
            Me.ActiveControl.Text = strContentsBeforePaste
        ElseIf CStr(CDbl(Me.ActiveControl.Text)) <> Me.ActiveControl.Text Then
            Me.ActiveControl.Text = strContentsBeforePaste
        End If
    End If
End Sub
Private Sub HandleTextBoxMouseDown(Button As Integer)
    If Button = vbRightButton Then strContentsBeforePaste = Me.ActiveControl.Text
End Sub
Private Sub HandleTextBoxMouseUp(Button As Integer)
    On Error Resume Next
    If Button = vbRightButton Then
        If Not IsNumeric(Me.ActiveControl.Text) Then
            Me.ActiveControl.Text = strContentsBeforePaste
        ElseIf CStr(CDbl(Me.ActiveControl.Text)) <> Me.ActiveControl.Text Then
            Me.ActiveControl.Text = strContentsBeforePaste
        End If
    End If
End Sub

Private Sub CanSave()
    If Not blnLoading Then
        cmdSave.Enabled = True
        cmdClose.Caption = CancelCaption
    End If
End Sub

Private Sub txtBillNarrative_Change()
    If Mode = typUpdate Then CanSave
End Sub

Private Sub txtBillNarrative_GotFocus()
    ssb.Panels(1) = "Bill narrative."
End Sub

Private Sub txtCharge_GotFocus()
    ssb.Panels(1) = "Charge."
End Sub

Private Sub txtMarkup_Change()
    CanSave
End Sub

Private Sub txtMarkup_GotFocus()
    ssb.Panels(1) = "Markup ratio."
End Sub

Private Sub txtMarkup_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 3          'Copy - Ctrl+C
        Case 8          'Backspace
        Case 22         'Paste - Ctrl+V
            strContentsBeforePaste = txtAmount
        Case 24         'Cut - Ctrl+X
        Case 46         '.
            If InStr(Me.ActiveControl.Text, ".") > 0 Then KeyAscii = 0
        Case 48 To 57   '0-9
        Case Else
            KeyAscii = 0
    End Select

End Sub

Private Sub txtMarkup_KeyUp(KeyCode As Integer, Shift As Integer)
    HandleTextBoxKeyUp KeyCode
End Sub

Private Sub txtMarkup_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    HandleTextBoxMouseDown Button
End Sub

Private Sub txtMarkup_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    HandleTextBoxMouseUp Button
End Sub
