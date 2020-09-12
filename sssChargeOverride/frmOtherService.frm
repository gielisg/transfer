VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmOtherService 
   Caption         =   "Form1"
   ClientHeight    =   4605
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9000
   Icon            =   "frmOtherService.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4605
   ScaleWidth      =   9000
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select"
      Default         =   -1  'True
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   6000
      TabIndex        =   2
      Top             =   2880
      Width           =   1335
   End
   Begin VSFlex7LCtl.VSFlexGrid vsfService 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   7335
      _cx             =   12938
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
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmOtherService.frx":000C
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
   Begin VB.Label Label1 
      Caption         =   "Service List"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuSelect 
         Caption         =   "Select"
      End
      Begin VB.Menu mnuSpe1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
   End
End
Attribute VB_Name = "frmOtherService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Event SelectClick(SpCnRef As Long, PhoneNumber As String)

Private Const DisTitlt As String = "Service List"
Private Const MinFormHeight As Integer = 4500
Private Const MinFormWidth As Integer = 5070

Dim objRS As ADODB.Recordset
Dim objDatabase As Object

Public Function Load(vobjDatabase) As Boolean
'Purpose    : name
'Inputs     :
'Outputs    : None.
'Version: $SSSVersion:
'Created on : 13/09/2001
'Created by : Patrick
'***********************************************************************************************

'-------------------------------
'*** Generic Procedure Code
'-------------------------------
    On Error GoTo ErrorHandler
    Dim sModuleAndProcName As String
    sModuleAndProcName = "ctlsssCallsHistory frmOtherService - Load Procedure"
    
'-------------------------------
'*** Specific Procedure Code
'-------------------------------
    Set objDatabase = vobjDatabase
    
    vsfService.Rows = vsfService.FixedRows
    Me.Caption = DisTitlt
    Load = LoadService
Exit Function
'------------------------------------------------------------
'*** Generic Error Handling Code (ensure Exit Proc is above)
'------------------------------------------------------------
ErrorHandler:
    
    Load = False
        
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

Public Function LoadService() As Boolean
'***********************************************************************************************
'Purpose    : name
'Inputs     :
'Outputs    : None.
'Version: $SSSVersion:
'Created on : 13/09/2001
'Created by : Patrick
'***********************************************************************************************

'-------------------------------
'*** Generic Procedure Code
'-------------------------------
    On Error GoTo ErrorHandler
    Dim sModuleAndProcName As String
    sModuleAndProcName = "ctlsssCallsHistory frmOtherService LoadService"
    
'-------------------------------
'*** Specific Procedure Code
'-------------------------------
    Dim idx As Integer
    Dim strNarrative As String
    Dim ServiceIndex As Integer
    
    Set objRS = objDatabase.GetServiceList
                
    With vsfService
        While Not objRS.EOF
            If objRS(0) = 0 Then
                                        
                If NullText(objRS(3)) <> "" Then
                    strNarrative = NullText(objRS(3))
                Else
                    strNarrative = NullText(objRS(2))
                End If
                
                .AddItem NullText(objRS(1)) & vbTab & _
                    NullText(objRS(2)) & vbTab & _
                    strNarrative & vbTab & _
                    Format(objRS(4), "Short Date") & vbTab & _
                    NullText(objRS(6)) & vbTab & _
                    IIf(NullText(objRS(10)) = "Y", "Yes", "No"), .Rows

            ElseIf objRS(0) <> 100 Then
                MsgBox "Unalbe to get service list." & vbNewLine & _
                    Trim$(objRS(objRS.Fields.Count - 1)), vbCritical + vbOKOnly, DisTitlt
                GoTo ExitFunction
            End If
            objRS.MoveNext
        Wend
        
        If .Rows > .FixedRows Then
            ServiceIndex = .FixedRows
            
            For idx = .FixedRows To .Rows - 1
                If .TextMatrix(idx, 0) = objDatabase.objApplication.SpCnRef Then
                    ServiceIndex = idx
                    .Select idx, 0, idx, .Cols - 1
                    Exit For
                End If
            Next
            
            .Select ServiceIndex, 0, ServiceIndex, .Cols - 1
            .AutoSize 0, .Cols - 1
            LoadService = True
        Else
            MsgBox "There is no service for this account.", vbExclamation + vbOKOnly, DisTitlt
        End If
    End With
  
ExitFunction:

Exit Function
'------------------------------------------------------------
'*** Generic Error Handling Code (ensure Exit Proc is above)
'------------------------------------------------------------
ErrorHandler:
    
    LoadService = False
        
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

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdSelect_Click()
    Dim PhoneNumber As String
    PhoneNumber = vsfService.TextMatrix(vsfService.RowSel, 1)
    Me.Hide
    RaiseEvent SelectClick(vsfService.TextMatrix(vsfService.RowSel, 0), PhoneNumber)
End Sub

Private Sub Form_Resize()
    If Me.WindowState <> 1 Then
        If Height < MinFormHeight Then
            Height = MinFormHeight
        End If
        If Width < MinFormWidth Then
            Width = MinFormWidth
        End If
        
        vsfService.Height = Height - 1800
        vsfService.Width = Width - 450
        cmdCancel.Top = Height - 1200
        cmdSelect.Top = cmdCancel.Top
        cmdCancel.Left = Width - 1800
        cmdSelect.Left = cmdCancel.Left - 1680
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

Private Sub mnuClose_Click()
    cmdCancel_Click
End Sub

Private Sub mnuSelect_Click()
    cmdSelect_Click
End Sub


Private Sub vsfService_DblClick()
    cmdSelect_Click
End Sub
