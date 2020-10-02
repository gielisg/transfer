VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "sssRoleList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarRoleID As Integer 'local copy
Private mvarRoleNarr As String 'local copy

Public Property Let RoleNarr(ByVal vData As String)
    mvarRoleNarr = vData
End Property

Public Property Get RoleNarr() As String
    RoleNarr = mvarRoleNarr
End Property

Public Property Let RoleID(ByVal vData As Integer)
    mvarRoleID = vData
End Property

Public Property Get RoleID() As Integer
    RoleID = mvarRoleID
End Property



