VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmChargeOverride 
   Caption         =   "Form1"
   ClientHeight    =   5685
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11490
   Icon            =   "frmChargeOverride.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5685
   ScaleWidth      =   11490
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkUnused 
      Caption         =   "Unused Records"
      Height          =   255
      Left            =   4200
      TabIndex        =   3
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CheckBox chkOther 
      Caption         =   "Service Level Records"
      Height          =   255
      Left            =   2040
      TabIndex        =   2
      Top             =   2640
      Width           =   2055
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   3360
      Width           =   1440
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "&End"
      Height          =   375
      Left            =   5160
      TabIndex        =   7
      Top             =   3360
      Width           =   1440
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   3360
      Width           =   1440
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   6840
      TabIndex        =   8
      Top             =   3360
      Width           =   1440
   End
   Begin VB.CheckBox chkOld 
      Caption         =   "Old Records"
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4320
      Top             =   0
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   4680
      TabIndex        =   6
      Top             =   3000
      Width           =   1440
   End
   Begin VSFlex7LCtl.VSFlexGrid vsf 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   7695
      _cx             =   13573
      _cy             =   4048
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483624
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   3
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   20
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmChargeOverride.frx":000C
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   1
      ExplorerBar     =   5
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin MSComctlLib.StatusBar ssb 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   5310
      Width           =   11490
      _ExtentX        =   20267
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16210
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
   Begin VB.Label lblDisplayed 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   6210
      TabIndex        =   11
      Top             =   2640
      Width           =   1485
   End
   Begin VB.Label lblShow 
      Caption         =   "Show:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2640
      Width           =   495
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNew 
         Caption         =   "New"
      End
      Begin VB.Menu mnuUpdate 
         Caption         =   "Update"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnuEnd 
         Caption         =   "End"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintAll 
         Caption         =   "Print All"
      End
      Begin VB.Menu mnuPrintSelection 
         Caption         =   "Print Selection"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "View"
      Visible         =   0   'False
      Begin VB.Menu mnuViewService 
         Caption         =   "Service Charge Overrides"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuCopyAll 
         Caption         =   "Copy All"
      End
      Begin VB.Menu mnuCopySelection 
         Caption         =   "Copy Selection"
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuNew1 
         Caption         =   "New"
      End
      Begin VB.Menu mnuUpdate1 
         Caption         =   "Update"
      End
      Begin VB.Menu mnuDelete1 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnuEnd1 
         Caption         =   "End"
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintAll2 
         Caption         =   "Print All"
      End
      Begin VB.Menu mnuPrintSelection2 
         Caption         =   "Print Selection"
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCopyAll2 
         Caption         =   "Copy All"
      End
      Begin VB.Menu mnuCopySelection2 
         Caption         =   "Copy Selection"
      End
      Begin VB.Menu mnuSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose1 
         Caption         =   "Close"
      End
   End
End
Attribute VB_Name = "frmChargeOverride"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim objDatabase As clsDatabase

Dim objApplication As Application
Dim objCommonDB As Object
Dim objRS As ADODB.Recordset
Dim strSQL As String

Dim PhoneNumber As String
Dim strProcessStep As String

Private Const DisTitle As String = "Charge Override"
Private Const MinFormWidth As Long = 9000
Private Const MinFormHeight As Long = 4800

Private Const ObjectName As String = "Object Name: "
Private Const PropertyName As String = "Property Name: "
Private Const MethodName As String = "Method Name: "
Private Const CommonDBObject As String = "sssCommonDB.CommonDB"
Private Const cstCurrent As String = "Current"
Private Const cstOld As String = "Old"
Private Const cstFuture As String = "Future"

Private Const EndCaption As String = "&End"
Private Const UnEndCaption As String = "R&e-Open"
Private Const cstYES As String = "Yes"
Private Const cstNO As String = "No"

Private Const OverrideIDCol As Integer = 0
Private Const TypeCol As Integer = 2
Private Const PackageNarrCol As Integer = 3
Private Const OPtionNarrCol As Integer = 4
Private Const ChargeNarrCol As Integer = 6
Private Const BillNarrativeCol As Integer = 7
Private Const AmountCol As Integer = 8
Private Const MarkupCol As Integer = 9
Private Const StartDateCol As Integer = 10
Private Const EndDateCol As Integer = 11
Private Const UsedCol As Integer = 15

Dim LastChargeOverrideID As Long
Dim blnChargeOverrideChanges As Boolean
Dim idx As Long

Public Function Populate(obj As Application) As Boolean
'***********************************************************************************************
'Purpose    : Populate form and set data
'Inputs     :
'Outputs    : None.
'Version: $SSSVersion:
'Created on : 20/03/2001
'Created by : Patrick
'***********************************************************************************************

'-------------------------------
'*** Generic Procedure Code
'-------------------------------
    On Error GoTo errHandler
    Dim sModuleAndProcName As String
    sModuleAndProcName = "frmChargeOverride - Populate Procedure"
    
'-------------------------------
'*** Specific Procedure Code
'-------------------------------

    Timer1.Enabled = False
    Set objApplication = obj
    
    Populate = False
    
    On Error GoTo ObjErrorHandler
    strProcessStep = ObjectName & CommonDBObject
    
    Set objCommonDB = CreateObject(CommonDBObject)
    
    strProcessStep = strProcessStep & vbNewLine & _
        PropertyName & "objApplication"

    Set objCommonDB.objApplication = objApplication
    
    On Error GoTo 0
    
    Set objDatabase = New clsDatabase
    Set objDatabase.objApplication = objApplication
    
    Set objApplication = obj
    
    If Not SetData Then
        Exit Function
    End If
    
    Populate = True
    Timer1.Enabled = True
    
Exit Function
'------------------------------------------------------------
'*** Generic object error handler
'------------------------------------------------------------
ObjErrorHandler:
    Err.Description = Err.Description & vbNewLine & strProcessStep

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
'    Err.Raise lErrNum, sErrSource, sErrDesc
    '*** OR ignore the error and continue
    'Resume Next
    
End Function

Private Function SetData() As Boolean
'***********************************************************************************************
'Purpose    : Set data up for
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
    sModuleAndProcName = "frmChargeOverride - SetData Procedure"
    
'-------------------------------
'*** Specific Procedure Code
'-------------------------------
    If objApplication.Mode = typContact Then
        Me.Caption = DisTitle & " [Account: " & objApplication.ContactCode & "]"
        vsf.ColHidden(1) = True
    Else
        If GetPhoneNumber Then
            Me.Caption = DisTitle & " [Service: " & PhoneNumber & "]"
        Else
            Exit Function
        End If
        
        mnuViewService.Caption = "Account Charge Override"
        chkOther.Caption = "Account Level Records"
    End If
    
    vsf.Rows = vsf.FixedRows
    
    chkUnused.Value = Checked
    
    With ssb
        .Panels(2).Text = objApplication.User.logoncode
        .Panels(3).Text = objApplication.User.CurrentBusinessUnit
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

Private Sub UnloadForm()
    Set objRS = Nothing
    Set objDatabase = Nothing
    Set objCommonDB = Nothing
    Set objApplication = Nothing
    Unload Me
End Sub

Private Sub chkOld_Click()
    ShowHideRow
    chkOld_GotFocus
End Sub

Private Sub ShowHideRow()
    Dim FirstShowRow As Integer
    Dim TotalRecord As Integer
    Dim DisplayedRecord As Integer
    
    With vsf
        TotalRecord = .Rows - .FixedRows
        
        If chkOld.Value = Checked And chkUnused.Value = Checked Then
            .RowHidden(-1) = False
            If .Rows > .FixedRows Then
                FirstShowRow = .FixedRows
            End If
            DisplayedRecord = TotalRecord
        Else
            For idx = .FixedRows To .Rows - 1
                            
                If chkOld.Value = Checked Then
                    .RowHidden(idx) = (.TextMatrix(idx, 11) = "Y")
                ElseIf chkUnused.Value = Checked Then
                    .RowHidden(idx) = (.TextMatrix(idx, 2) = cstOld)
                Else
                    .RowHidden(idx) = (.TextMatrix(idx, 11) = "Y") Or (.TextMatrix(idx, 2) = cstOld)
                End If
                
                If Not .RowHidden(idx) Then
                    DisplayedRecord = DisplayedRecord + 1
                End If
                
                If FirstShowRow = 0 Then
                    If Not .RowHidden(idx) Then
                        FirstShowRow = idx
                    End If
                End If
            Next
        End If
        
        If FirstShowRow > 0 Then
            .Select FirstShowRow, 0, FirstShowRow, .Cols - 1
        End If
    
        vsf_SelChange
        
        lblDisplayed = DisplayedRecord & " of " & TotalRecord & " Records Displayed"
        lblDisplayed.Left = Me.ScaleWidth - 240 - lblDisplayed.Width
    End With
End Sub

Private Sub chkOld_GotFocus()
    ssb.Panels(1) = "Show " & LCase(chkOld.Caption) & "."
End Sub

Private Sub chkOther_Click()
    LoadChargeOverride
    vsf.ColHidden(1) = Not (chkOther.Value = Checked)
    chkOther_GotFocus
End Sub

Private Sub chkOther_GotFocus()
    ssb.Panels(1) = "Show " & LCase(chkOther.Caption) & "."
End Sub

Private Sub chkunused_Click()
    ShowHideRow
    chkunused_GotFocus
End Sub

Private Sub chkunused_GotFocus()
    ssb.Panels(1) = "Show " & LCase(chkUnused.Caption) & "."
End Sub

Private Sub cmdClose_Click()
    UnloadForm
End Sub

Private Sub cmdClose_GotFocus()
    ssb.Panels(1) = "Exit."
End Sub

Private Sub cmdDelete_Click()
    If MsgBox("Are you sure you want to delete this charge override?", _
        vbQuestion + vbYesNo, DisTitle) = vbNo Then Exit Sub
    
    If DeleteChargeOverride Then LoadChargeOverride
End Sub

Private Sub cmdDelete_GotFocus()
    ssb.Panels(1) = "Delete charge override."
End Sub

Private Sub cmdEnd_Click()
    If cmdEnd.Caption = UnEndCaption Then
        If UnEndChargeOverride Then LoadChargeOverride
    ElseIf EndChargeOverride Then
        If blnChargeOverrideChanges Then LoadChargeOverride
    End If
End Sub

Private Sub cmdEnd_GotFocus()
    ssb.Panels(1) = Replace(cmdEnd.Caption, "&", "") & " charge override."
End Sub

Private Sub cmdNew_Click()
    If NewChargeOverride Then
        If blnChargeOverrideChanges Then LoadChargeOverride
    End If
End Sub

Public Function NewChargeOverride() As Boolean
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
    sModuleAndProcName = "frmChargeOverride - NewChargeOverride"

'-------------------------------
'*** Specific Procedure Code
'-------------------------------
    Dim objfrm As frmOverAmount
    
    Set objfrm = New frmOverAmount
               
    With objfrm
        Set .objCommonDB = objCommonDB
        Set .objDatabase = objDatabase
        .Mode = typNew
        .ServiceNumber = PhoneNumber
        
        If .Populate Then
            .Show vbModal, Me
        End If
        If Not .blnCancel Then blnChargeOverrideChanges = True
        
        If blnChargeOverrideChanges Then LastChargeOverrideID = .OverrideID
    End With
    
    Set objfrm = Nothing
    
    NewChargeOverride = True
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
    'Err.Raise lErrNum, sErrSource, sErrDesc
    '*** OR ignore the error and continue
    'Resume Next
    
    
End Function

Public Function EndChargeOverride() As Boolean
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
    sModuleAndProcName = "frmChargeOverride - EndChargeOverride"

'-------------------------------
'*** Specific Procedure Code
'-------------------------------
    Dim objfrmEndChargeOverride As frmEndChargeOverride
    
    Set objfrmEndChargeOverride = New frmEndChargeOverride
    
    With objfrmEndChargeOverride
        .ChargeOverrideID = vsf.TextMatrix(vsf.RowSel, 0)
        If .Populate(objDatabase) Then
            .Show vbModal, Me
        End If
        If Not .blnCancel Then blnChargeOverrideChanges = True
        
        If blnChargeOverrideChanges Then LastChargeOverrideID = .ChargeOverrideID
    End With
    
    Set objfrmEndChargeOverride = Nothing
    
    EndChargeOverride = True
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
    'Err.Raise lErrNum, sErrSource, sErrDesc
    '*** OR ignore the error and continue
    'Resume Next
    
    
End Function

Public Function UnEndChargeOverride() As Boolean
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
    sModuleAndProcName = "frmChargeOverride - UnEndChargeOverride"

'-------------------------------
'*** Specific Procedure Code
'-------------------------------
    UnEndChargeOverride = objDatabase.EndChargeOverride(vsf.TextMatrix(vsf.RowSel, 0), #1/1/9999#)
    LastChargeOverrideID = vsf.TextMatrix(vsf.RowSel, 0)
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
    'Err.Raise lErrNum, sErrSource, sErrDesc
    '*** OR ignore the error and continue
    'Resume Next
    
    
End Function

Private Sub cmdNew_GotFocus()
    ssb.Panels(1) = "New charge override."
End Sub

Private Sub cmdUpdate_Click()
    If UpdateChargeOverride Then
        LoadChargeOverride
    End If
End Sub
Private Function UpdateChargeOverride() As Boolean

    Dim objfrm As frmOverAmount
    Set objfrm = New frmOverAmount
    
    With objfrm
        .Mode = typUpdate
        .OverrideID = vsf.TextMatrix(vsf.RowSel, OverrideIDCol)
        .PackageNarr = vsf.TextMatrix(vsf.RowSel, PackageNarrCol)
        .OptionNarr = vsf.TextMatrix(vsf.RowSel, OPtionNarrCol)
        .ChargeNarr = vsf.TextMatrix(vsf.RowSel, ChargeNarrCol)
        .BillBarrative = vsf.TextMatrix(vsf.RowSel, BillNarrativeCol)
        .StartDate = vsf.TextMatrix(vsf.RowSel, StartDateCol)
        .EndDate = vsf.TextMatrix(vsf.RowSel, EndDateCol)
        .Amount = vsf.TextMatrix(vsf.RowSel, AmountCol)
        .Markup = vsf.TextMatrix(vsf.RowSel, MarkupCol)
        
        If objApplication.Mode = typService Then
            .ServiceNumber = PhoneNumber
        Else
            If vsf.TextMatrix(vsf.RowSel, 1) <> "<All Services>" Then
                .ServiceNumber = vsf.TextMatrix(vsf.RowSel, 1)
            End If
        End If
        
        Set .objCommonDB = objCommonDB
        Set .objDatabase = objDatabase
        
        If .Populate Then
            .Show vbModal, Me
        End If
        
        If Not .blnCancel Then
            blnChargeOverrideChanges = True
            LastChargeOverrideID = .OverrideID
            UpdateChargeOverride = True
        End If
    End With
    Set objfrm = Nothing
    
End Function

Private Sub cmdUpdate_GotFocus()
    ssb.Panels(1) = "Update charge override."
End Sub

Private Sub Form_Resize()
    If Me.WindowState <> 1 Then
        'set minimeum size
        Me.ScaleMode = 1 'twip
        If Height < MinFormHeight Then
            Height = MinFormHeight
        End If
        If Width < MinFormWidth Then
            Width = MinFormWidth
        End If
        
        vsf.Move 120, 240, Me.ScaleWidth - 240, Me.ScaleHeight - 1800
                
        cmdClose.Left = Me.ScaleWidth - 1800
        cmdEnd.Left = cmdClose.Left - 1680
        cmdDelete.Left = cmdClose.Left - 1680 * 2
        cmdUpdate.Left = cmdClose.Left - 1680 * 3
        cmdNew.Left = cmdClose.Left - 1680 * 4
        
        cmdClose.Top = Me.ScaleHeight - 960
        cmdUpdate.Top = cmdClose.Top
        cmdEnd.Top = cmdClose.Top
        cmdDelete.Top = cmdClose.Top
        cmdNew.Top = cmdClose.Top
        lblShow.Top = vsf.Height + 300
        chkOld.Top = lblShow.Top
        chkOther.Top = lblShow.Top
        chkUnused.Top = lblShow.Top
        lblDisplayed.Top = lblShow.Top
        lblDisplayed.Left = Me.ScaleWidth - 120 - lblDisplayed.Width

    End If
End Sub

Private Function LoadChargeOverride() As Boolean
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
    sModuleAndProcName = "frmChargeOverride - LoadChargeOverride"

'-------------------------------
'*** Specific Procedure Code
'-------------------------------
    Dim strType As String
    Dim SelectRow As Long
    Dim idxRow As Long
    Dim strServiceNumber As String
    Dim strUsed As String
    Dim blnOld As Boolean
    Dim blnNotUsed As Boolean
    
    Me.MousePointer = vbHourglass
    ssb.Panels(1).Text = "Loading account charge override..."
       
    Set objRS = objDatabase.GetChargeOverride(chkOther.Value = Checked)
    
    With vsf
        .Rows = .FixedRows
                
        While Not objRS.EOF
            If objRS(0) = 0 Then
                
                If IsNull(objRS("phone_num")) Then
                    strServiceNumber = "<All Services>"
                Else
                    strServiceNumber = NullText(objRS("phone_num"))
                End If
                
                If DateDiff("d", objRS("ChgOverrideWefr"), Date) < 0 Then
                    strType = cstFuture
                ElseIf DateDiff("d", Date, objRS("ChgOverrideWeto")) < 0 Then
                    strType = cstOld
                    blnOld = True
                Else
                    strType = cstCurrent
                End If
                
                If NullText(objRS("editable")) = "Y" Then
                    strUsed = cstNO
                    blnNotUsed = True
                Else
                    strUsed = cstYES
                End If
                
                .AddItem NullText(objRS("ChgOverrideId")) & vbTab & _
                    strServiceNumber & vbTab & _
                    strType & vbTab & _
                    NullText(objRS("pkg_narr")) & vbTab & _
                    NullText(objRS("pkg_opt_narr")) & vbTab & _
                    NullText(objRS("chg_code")) & vbTab & _
                    NullText(objRS("chg_narr")) & vbTab & _
                    NullText(objRS("ChgOverrideBNarr")) & vbTab & _
                    Format(objRS("chg_amount"), "Currency") & vbTab & _
                    NullText(objRS("nt_cost_comp")) & vbTab & _
                    UserDateTime(objRS("ChgOverrideWefr")) & vbTab & _
                    UserDateTime(objRS("ChgOverrideWeto")) & vbTab & _
                    NullText(objRS("pkg_code")) & vbTab & _
                    NullText(objRS("option")) & vbTab & _
                    NullText(objRS("chg_code")) & vbTab & _
                    strUsed & vbTab & _
                    UserDateTime(objRS("created_tm")) & vbTab & _
                    NullText(objRS("created_by")) & vbTab & _
                    UserDateTime(objRS("last_updated")) & vbTab & _
                    NullText(objRS("updated_by")), .Rows
                                
            ElseIf objRS(0) <> 100 Then
                MsgBox "Unable to load account charge override details." & vbNewLine & _
                    Trim$(objRS(objRS.Fields.Count - 1)), vbCritical + vbOKOnly, DisTitle
                GoTo ExitFunction
            End If
            
            objRS.MoveNext
        Wend
               
        .AutoSize 0, .Cols - 1
        If Not blnOld Then
            chkOld.Value = Unchecked
            chkOld.Enabled = False
        Else
            chkOld.Enabled = True
        End If
        
        If Not blnNotUsed Then
            chkUnused.Value = Unchecked
            chkUnused.Enabled = False
        Else
            chkUnused.Enabled = True
        End If
        
        ShowHideRow
        SelectRow = .FixedRows
        
        If .Rows > .FixedRows Then
            If LastChargeOverrideID <> 0 Then
                For idxRow = .FixedRows To .Rows - 1
                    If (.TextMatrix(idxRow, 0) = LastChargeOverrideID) _
                        And Not .RowHidden(idxRow) Then
                        
                        SelectRow = idxRow
                        Exit For
                    End If
                Next
                .Select SelectRow, 0, SelectRow, .Cols - 1
                .ShowCell SelectRow, 0
            End If
            
        End If
    End With
    
    Set objRS = Nothing
    
    LoadChargeOverride = True

ExitFunction:
    Me.MousePointer = vbDefault
    ssb.Panels(1) = ""
    
Exit Function
'------------------------------------------------------------
'*** Generic Error Handling Code (ensure Exit Proc is above)
'------------------------------------------------------------
errHandler:
    Me.MousePointer = vbDefault
    If Err.Number = 0 Then Exit Function
    
    LoadChargeOverride = False
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

Private Function DeleteChargeOverride() As Boolean
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
    sModuleAndProcName = "frmChargeOverride - DeleteChargeOverride"

'-------------------------------
'*** Specific Procedure Code
'-------------------------------

    Me.MousePointer = vbHourglass
    ssb.Panels(1).Text = "Deleting account charge override..."
       
    Set objRS = objDatabase.GetDeleteChargeOverride(vsf.TextMatrix(vsf.RowSel, 0))
    If objRS(0) = 0 Then
        DeleteChargeOverride = True
    Else
        MsgBox "Unable to delete charge override.", vbCritical + vbOKOnly, DisTitle
    End If
    
    Set objRS = Nothing
    
    Me.MousePointer = vbDefault
    ssb.Panels(1) = ""
Exit Function
'------------------------------------------------------------
'*** Generic Error Handling Code (ensure Exit Proc is above)
'------------------------------------------------------------
errHandler:
    Me.MousePointer = vbDefault
    If Err.Number = 0 Then Exit Function
    
    DeleteChargeOverride = False
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

Private Sub mnuClose_Click()
    cmdClose_Click
End Sub

Private Sub mnuClose1_Click()
    cmdClose_Click
End Sub

Private Sub mnuCopyAll_Click()
    vsfCopyAll vsf, DisTitle, False, False, ssb, Me
End Sub

Private Sub mnuCopyAll2_Click()
    mnuCopyAll_Click
End Sub

Private Sub mnuCopySelection_Click()
    vsfCopySelection vsf, DisTitle, False, ssb, Me
End Sub

Private Sub mnuCopySelection2_Click()
    mnuCopySelection_Click
End Sub

Private Sub mnuDelete_Click()
    cmdDelete_Click
End Sub

Private Sub mnuDelete1_Click()
    cmdDelete_Click
End Sub

Private Sub mnuEnd_Click()
    cmdEnd_Click
End Sub

Private Sub mnuEnd1_Click()
    cmdEnd_Click
End Sub

Private Sub mnuFile_Click()
    mnuUpdate.Enabled = cmdUpdate.Enabled
    mnuDelete.Enabled = cmdDelete.Enabled
    mnuEnd.Caption = Replace(cmdEnd.Caption, "&", "")
    mnuEnd.Enabled = cmdEnd.Enabled
End Sub

Private Sub mnuNew_Click()
    cmdNew_Click
End Sub

Private Sub mnuNew1_Click()
    cmdNew_Click
End Sub

Private Sub mnuPopUp_Click()
    mnuUpdate1.Enabled = cmdUpdate.Enabled
    mnuDelete1.Enabled = cmdDelete.Enabled
    mnuEnd1.Caption = Replace(cmdEnd.Caption, "&", "")
    mnuEnd1.Enabled = cmdEnd.Enabled
End Sub

Private Sub mnuPrintAll_Click()
    vsfPrintAll vsf, DisTitle, , , 500, 500, ssb, Me
End Sub

Private Sub mnuPrintAll2_Click()
    mnuPrintAll_Click
End Sub

Private Sub mnuPrintSelection_Click()
    vsfPrintSelection vsf, DisTitle, , , 500, 500, ssb, Me
End Sub

Private Sub mnuUpdate_Click()
    cmdUpdate_Click
End Sub

Private Sub mnuUpdate1_Click()
    cmdUpdate_Click
End Sub

Private Sub mnuViewService_Click()
    mnuViewService.Checked = Not mnuViewService.Checked
    
    vsf.ColHidden(1) = Not mnuViewService.Checked
    
    LoadChargeOverride
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    If Not LoadChargeOverride Then
        UnloadForm
    End If
End Sub

Private Function NullText(str As Variant) As String
    If IsNull(str) Then
        NullText = ""
    Else
        NullText = Trim$(str)
    End If
End Function

Private Function UserDateTime(Dt As Variant) As String
    'return format, using system short date format if time is 00:00:00
    'or system short date format+ " " + system long time dormat
    Dim strTemp As String
    If Not IsNull(Dt) Then
        If Format(Dt, "hh:mm:ss") = Format("00:00:00", "hh:mm:ss") Then
            strTemp = Format(Dt, "Short Date")
        Else
            strTemp = Format(Dt, "Short Date") & " " & Format(Dt, "Long Time")
        End If
    Else
        strTemp = ""
    End If
    UserDateTime = strTemp
End Function

Private Sub vsf_GotFocus()
    ssb.Panels(1) = "Charge overrides."
End Sub

Private Sub vsf_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuPopup
    End If
End Sub

Private Sub vsf_SelChange()
    On Error Resume Next
    With vsf
        If .Rows >= .FixedRows And .RowSel >= .FixedRows Then
            cmdUpdate.Enabled = (.TextMatrix(.RowSel, UsedCol) = cstNO) And Not .RowHidden(.RowSel)
            If DateDiff("d", .TextMatrix(.RowSel, EndDateCol), Date) = 0 And Not .RowHidden(.RowSel) Then
                cmdEnd.Caption = UnEndCaption
                cmdEnd.Enabled = True
            Else
                cmdEnd.Caption = EndCaption
                cmdEnd.Enabled = (.TextMatrix(.RowSel, TypeCol) <> cstOld) _
                    And (DateDiff("d", .TextMatrix(.RowSel, EndDateCol), Date) < 0) And Not .RowHidden(.RowSel)
            End If
        Else
            cmdUpdate.Enabled = False
            cmdEnd.Caption = EndCaption
            cmdEnd.Enabled = False
        End If
    End With
    cmdDelete.Enabled = cmdUpdate.Enabled
End Sub

Public Function GetPhoneNumber() As Boolean
'***********************************************************************************************
'Purpose    : Get phone number for display on form
'Inputs     :
'Outputs    :
'Version: $SSSVersion:
'Created on : 13/03/2001
'Created by : Patrick
'***********************************************************************************************

'-------------------------------
'*** Generic Procedure Code
'-------------------------------
    On Error GoTo errHandler
    Dim sModuleAndProcName As String
    sModuleAndProcName = "frmChargeOverride - GetPhoneNumber Procedure"
    
'-------------------------------
'*** Specific Procedure Code
'-------------------------------
    
    On Error GoTo ObjErrorHandler
    
    strProcessStep = ObjectName & CommonDBObject
    strProcessStep = strProcessStep & vbNewLine & _
        MethodName & "GetPhoneNumber"
        
    Set objRS = objCommonDB.GetPhoneNumber(objApplication.SpCnRef)
    
    On Error GoTo 0
    
    If Not objRS.EOF Then
            GetPhoneNumber = True
            PhoneNumber = Trim(objRS!phone_num)
    Else
        GetPhoneNumber = False
        
        MsgBox "Unable to get service number for service " & objApplication.SpCnRef, _
            vbCritical + vbOKOnly, DisTitle
    End If
    Set objRS = Nothing
Exit Function
'------------------------------------------------------------
'*** Generic object error handler
'------------------------------------------------------------
ObjErrorHandler:
    Err.Description = Err.Description & vbNewLine & strProcessStep
'------------------------------------------------------------
'*** Generic Error Handling Code (ensure Exit Proc is above)
'------------------------------------------------------------
errHandler:
    If Err.Number = 0 Then Exit Function
    GetPhoneNumber = False
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


