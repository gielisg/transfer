VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmEndChargeOverride 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   1500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3465
   Icon            =   "frmEndChargeOverride.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   3465
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   840
      Width           =   1440
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "&End"
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1440
   End
   Begin MSComCtl2.DTPicker dtpEnd 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      _Version        =   393216
      Format          =   66781185
      CurrentDate     =   38273
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "End Date"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   360
      Width           =   735
   End
End
Attribute VB_Name = "frmEndChargeOverride"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim objDatabase As clsDatabase
Dim objRS As ADODB.Recordset

Private Const DisTitle As String = "End Charge Override"

Public blnCancel As Boolean
Public ChargeOverrideID As Long

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
    sModuleAndProcName = "frmEndChargeOverride - Populate Procedure"
    
'-------------------------------
'*** Specific Procedure Code
'-------------------------------
    Set objDatabase = xobjDatabase
    Me.Caption = DisTitle
    blnCancel = True
    
    dtpEnd = Date
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

Private Function EndChargeOverride() As Boolean
'***********************************************************************************************
'Purpose    : Name
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
    sModuleAndProcName = "frmEndChargeOverride - EndChargeOverride Procedure"

'-------------------------------
'*** Specific Procedure Code
'-------------------------------
    
    EndChargeOverride = objDatabase.EndChargeOverride(ChargeOverrideID, dtpEnd.Value)
    
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

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdEnd_Click()
    If EndChargeOverride Then
        blnCancel = False
        Unload Me
    End If
End Sub
