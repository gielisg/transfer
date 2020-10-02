VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AccessRight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Base 1

Public Enum enAccessRight
    FullAccess = 5
    ReadOnly = 3
    NoAccess = 0
End Enum

Private m_AccessRight As enAccessRight
Private m_strObjectNarr As String
Private m_lngObjectId As Long

Public Property Get AccessRight() As enAccessRight
    AccessRight = m_AccessRight
End Property
Public Property Let AccessRight(vData As enAccessRight)
    m_AccessRight = vData
End Property
Public Property Get ObjectNarr() As String
  ObjectNarr = m_strObjectNarr
End Property
Public Property Let ObjectNarr(ByVal Value As String)
  m_strObjectNarr = Value
End Property
Public Property Get ObjectId() As Long
  ObjectId = m_lngObjectId
End Property
Public Property Let ObjectId(ByVal Value As Long)
  m_lngObjectId = Value
End Property

Private Sub Class_Initialize()
    m_AccessRight = NoAccess
End Sub
