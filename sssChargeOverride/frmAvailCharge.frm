VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmAvailCharge 
   Caption         =   "  "
   ClientHeight    =   5070
   ClientLeft      =   165
   ClientTop       =   630
   ClientWidth     =   9420
   Icon            =   "frmAvailCharge.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5070
   ScaleWidth      =   9420
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   6240
      TabIndex        =   1
      Top             =   3720
      Width           =   1440
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select"
      Default         =   -1  'True
      Height          =   375
      Left            =   4200
      TabIndex        =   0
      Top             =   3720
      Width           =   1440
   End
   Begin MSComctlLib.StatusBar ssb 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   4695
      Width           =   9420
      _ExtentX        =   16616
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12717
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
            TextSave        =   "26/08/2008"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            Object.Width           =   1217
            MinWidth        =   176
            TextSave        =   "2:05 PM"
         EndProperty
      EndProperty
   End
   Begin VSFlex7LCtl.VSFlexGrid vsf 
      Height          =   2295
      Left            =   360
      TabIndex        =   3
      Top             =   720
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
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmAvailCharge.frx":000C
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
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuSelect 
         Caption         =   "Select"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
   End
End
Attribute VB_Name = "frmAvailCharge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim objRS As ADODB.Recordset
Dim objDatabase As clsDatabase
Dim NodeSelected As Integer

Private Const DisTitle As String = "Available Charge"
Private Const cstYES As String = "Yes"
Private Const cstNO As String = "No"

Private Const MinFormHeight As Long = 4200
Private Const MinFormWidth As Long = 4500

Public objCommonDB As Object
Public blnCancelled As Boolean
Public PackageCode As Long
Public PackageOption As Long

Public SelectedChargeCode As String
Public SelectedChargeNarr As String
Public SelectedAmount As Variant
Public SelectedMarkup As Variant

Public Function Populate(xobjDatabase As clsDatabase) As Boolean
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
    sModuleAndProcName = "frmAvailCharge - Populate Procedure"
    
'-------------------------------
'*** Specific Procedure Code
'-------------------------------
    
    blnCancelled = True
    Set objDatabase = xobjDatabase
            
    SetData
    
    Populate = True
    
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
'    Err.Raise lErrNum, sErrSource, sErrDesc
    '*** OR ignore the error and continue
    'Resume Next
    
End Function

Private Sub SetData()
'***********************************************************************************************
'Purpose    : Set data up for miscelaneous charges form
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
    sModuleAndProcName = "frmAvailCharge - SetData Procedure"
    
'-------------------------------
'*** Specific Procedure Code
'-------------------------------
    Dim idx As Integer

    Me.Caption = DisTitle
    vsf.Rows = vsf.FixedRows
    cmdSelect.Enabled = False
    
    With ssb
        .Panels(2).Text = objDatabase.objApplication.User.logoncode
        .Panels(3).Text = objDatabase.objApplication.User.CurrentBusinessUnit
        .Panels(2).Width = Me.TextWidth(.Panels(2).Text) + 150
        .Panels(3).Width = Me.TextWidth(.Panels(3).Text) + 150
    End With
    
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
'    Err.Raise lErrNum, sErrSource, sErrDesc
    '*** OR ignore the error and continue
    'Resume Next

End Sub

Private Sub HideForm()
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    HideForm
End Sub

Private Sub cmdCancel_GotFocus()
    ssb.Panels(1) = "Exit."
End Sub

Private Sub cmdSelect_Click()
    SelectedChargeCode = vsf.TextMatrix(vsf.RowSel, 0)
    SelectedChargeNarr = vsf.TextMatrix(vsf.RowSel, 1)
    SelectedAmount = vsf.TextMatrix(vsf.RowSel, 3)
    SelectedMarkup = vsf.TextMatrix(vsf.RowSel, 4)
    
    
    blnCancelled = False
    HideForm
End Sub

Private Sub cmdSelect_GotFocus()
    ssb.Panels(1) = "Select charge."
End Sub

Private Sub Form_Resize()
    If Me.WindowState <> 1 Then
        If Height < MinFormHeight Then
            Height = MinFormHeight
        End If
        If Width < MinFormWidth Then
            Width = MinFormWidth
        End If
        
        vsf.Move 120, 240, Me.ScaleWidth - 300, Me.ScaleHeight - 1260
                
        cmdCancel.Top = Me.ScaleHeight - 900
        cmdSelect.Top = cmdCancel.Top
        cmdCancel.Left = Width - 1800
        cmdSelect.Left = cmdCancel.Left - 1680
    End If

End Sub

Private Sub mnuClose_Click()
    cmdCancel.Value = True
End Sub

Private Sub mnuFile_Click()
    mnuSelect.Enabled = cmdSelect.Enabled
End Sub

Private Sub mnuSelect_Click()
    cmdSelect.Value = True
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    If Not Load Then
        HideForm
    End If
End Sub

Private Function Load() As Boolean
    Me.MousePointer = vbHourglass
    Me.Refresh
    
    If Not DisplayAvailCharges Then
        Exit Function
    End If
    
    Me.MousePointer = vbDefault
        
    Load = True
End Function

Private Function DisplayAvailCharges() As Boolean
'***********************************************************************************************
'Purpose    : display charges on form
'Inputs     :
'Outputs    : None.
'Version: $SSSVersion:
'Created on : 23/03/2001
'Created by : Patrick
'***********************************************************************************************

'-------------------------------
'*** Generic Procedure Code
'-------------------------------
    On Error GoTo errHandler
    Dim sModuleAndProcName As String
    sModuleAndProcName = "frmAvailCharge - DisplayAvailCharges Procedure"
    
'-------------------------------
'*** Specific Procedure Code
'-------------------------------
    
    Me.MousePointer = vbHourglass
    
    ssb.Panels(1).Text = "Loading available charges..."
    
    Set objRS = objDatabase.GetCharge(PackageCode, PackageOption)
    
    vsf.Rows = vsf.FixedRows
    
    While Not objRS.EOF
        If objRS(0) = 0 Then
            vsf.AddItem NullText(objRS(1)) & vbTab & _
                NullText(objRS(2)) & vbTab & _
                NullText(objRS(4)) & vbTab & _
                NullText(objRS(5)) & vbTab & _
                NullText(objRS(6)), vsf.Rows
        Else
            MsgBox Trim$(objRS(objRS.Fields.Count - 1)), vbCritical + vbOKOnly, DisTitle
            GoTo ExitFunction
        End If
        objRS.MoveNext
    Wend
    
    vsf.Select vsf.FixedRows, 0, vsf.FixedRows, vsf.Cols - 1
    
    vsf.AutoSize 0, vsf.Cols - 1
    
    cmdSelect.Enabled = True
    DisplayAvailCharges = True
    
ExitFunction:
    
    Me.MousePointer = vbDefault
    ssb.Panels(1) = ""
Exit Function
'------------------------------------------------------------
'*** Generic Error Handling Code (ensure Exit Proc is above)
'------------------------------------------------------------
errHandler:
    Me.MousePointer = vbDefault
    DisplayAvailCharges = False
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

Private Function NullText(str As Variant) As String
    If IsNull(str) Then
        NullText = ""
    Else
        NullText = Trim$(str)
    End If
    
End Function

Private Sub vsf_DblClick()
    If vsf.MouseRow < vsf.FixedRows Then Exit Sub
    If cmdSelect.Enabled Then cmdSelect_Click
End Sub

Private Sub vsf_GotFocus()
    ssb.Panels(1) = "Charge list."
End Sub

