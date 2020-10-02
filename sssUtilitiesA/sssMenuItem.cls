VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "sssMenuItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Base 1

Private m_lngObjectId As Long
Private m_strCaption As String
Private m_blnChecked As Boolean
Private m_intGroup As Integer
Private m_intOrder As Integer
Private m_strDll As String
Private m_blnEnabled As Boolean
Private m_blnDefault As Boolean
Private m_strCmd As String
Private m_strTooltip As String
Private m_strArgs As String
Private m_blnModality As Boolean
Private m_strActionFlag As String

Public Property Get ActionFlag() As String
    ActionFlag = m_strActionFlag
End Property
Public Property Let ActionFlag(strActionFlag As String)
    m_strActionFlag = strActionFlag
End Property
Public Property Get ObjectId() As Long
    ObjectId = m_lngObjectId
End Property
Public Property Let ObjectId(lngObjectId As Long)
    m_lngObjectId = lngObjectId
End Property
Public Property Get Caption() As String
    Caption = m_strCaption
End Property
Public Property Let Caption(strCaption As String)
    m_strCaption = strCaption
End Property
Public Property Get Default() As Boolean
    Default = m_blnDefault
End Property
Public Property Let Default(blnDefault As Boolean)
    m_blnDefault = blnDefault
End Property
Public Property Get Checked() As Boolean
    Checked = m_blnChecked
End Property
Public Property Let Checked(blnChecked As Boolean)
    m_blnChecked = blnChecked
End Property
Public Property Let Group(intGroup As Integer)
    m_intGroup = intGroup
End Property
Public Property Get Group() As Integer
    Group = m_intGroup
End Property
Public Property Let Order(intOrder As Integer)
    m_intOrder = intOrder
End Property
Public Property Get Order() As Integer
    Order = m_intOrder
End Property
Public Property Let Dll(strDll As String)
    m_strDll = strDll
End Property
Public Property Get Dll() As String
    Dll = m_strDll
End Property
Public Property Let Enabled(blnEnabled As Boolean)
    m_blnEnabled = blnEnabled
End Property
Public Property Get Enabled() As Boolean
    Enabled = m_blnEnabled
End Property
Public Property Let cmd(strCmd As String)
    m_strCmd = strCmd
End Property
Public Property Get cmd() As String
    cmd = m_strCmd
End Property

Public Property Let Tooltip(strTooltip As String)
    m_strTooltip = strTooltip
End Property
Public Property Get Tooltip() As String
    Tooltip = m_strTooltip
End Property
Public Property Let Args(strArgs As String)
    m_strArgs = strArgs
End Property
Public Property Get Args() As String
    Args = m_strArgs
End Property
Public Property Let Modality(blnModality As Boolean)
    m_blnModality = blnModality
End Property
Public Property Get Modality() As Boolean
    Modality = m_blnModality
End Property

Private Sub Class_Initialize()
    m_blnChecked = False
    m_blnModality = True
End Sub
