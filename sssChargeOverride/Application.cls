VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Application"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Const mksMode As String = "Mode"

Dim objfrm As frmChargeOverride
Dim objCDC As CommonDataComponent.CommonData

Public Enum enuType
    typContact = 0
    typService = 1
End Enum

Public dbAccess As Object
'local variable(s) to hold property value(s)
Private mvarUser As Object 'local copy
Private mvarParentForm As Object 'local copy
Private mvarContactCode As String 'local copy
Private mvarSpCnRef As Long 'local copy
Private mvarArguments As String 'local copy
Private mvarCanceled As Boolean 'local copy
Private mvarMode As enuMode 'local copy

Public Property Let Mode(ByVal vData As enuType)
    mvarMode = vData
End Property

Public Property Get Mode() As enuType
    Mode = mvarMode
End Property

Public Property Let Canceled(ByVal vData As Boolean)
    mvarCanceled = vData
End Property

Public Property Get Canceled() As Boolean
    Canceled = mvarCanceled
End Property

Public Property Let Arguments(ByVal vData As String)
    mvarArguments = vData
End Property

Public Property Get Arguments() As String
    Arguments = mvarArguments
End Property

Public Property Let SpCnRef(ByVal vData As Long)
    mvarSpCnRef = vData
End Property

Public Property Get SpCnRef() As Long
    SpCnRef = mvarSpCnRef
End Property

Public Property Let ContactCode(ByVal vData As String)
    mvarContactCode = vData
End Property

Public Property Get ContactCode() As String
    ContactCode = mvarContactCode
End Property

Public Property Set ParentForm(ByVal vData As Object)
    Set mvarParentForm = vData
End Property
Public Property Let ParentForm(ByVal vData As Object)
    Set mvarParentForm = vData
End Property

Public Property Get ParentForm() As Object
    Set ParentForm = mvarParentForm
End Property

Public Property Set User(ByVal vData As Object)
    Set mvarUser = vData
End Property
Public Property Let User(ByVal vData As Object)
    Set mvarUser = vData
End Property

Public Property Get User() As Object
    Set User = mvarUser
End Property

Public Function Start(Optional StartMode As Integer = 0) As Boolean
'***********************************************************************************************
'Purpose    : start the Quick Note application
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
    sModuleAndProcName = "Application - Start Procedure"
    
'-------------------------------
'*** Specific Procedure Code
'-------------------------------

    Start = False
1    Set objCDC = New CommonDataComponent.CommonData
2    Set dbAccess = CreateObject("sssUtilitiesA.dbAccess")
3    Set dbAccess = User.GetDbAccessObject(User.CurrentBusinessUnitCode)
    
4    objCDC.Stream = Arguments
    
    If Val(objCDC.Value(mksMode)) = 1 Then
        Mode = typService
    Else
        Mode = typContact
    End If
    
    Set objfrm = New frmChargeOverride
    
    If Not objfrm.Populate(Me) Then
        Exit Function
    End If
    
    If StartMode = 0 Then
        If Not ParentForm Is Nothing Then
            objfrm.Show vbModal, ParentForm
        Else
            objfrm.Show vbModal
        End If
    Else
        objfrm.Show vbModeless
    End If
    
    DoEvents
    Start = True
    
StartError:
65    Set objfrm = Nothing
    
Exit Function
'------------------------------------------------------------
'*** Generic Error Handling Code (ensure Exit Proc is above)
'------------------------------------------------------------
errHandler:
    If Err.Number = 0 Then Exit Function
    Start = False
    
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
