VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "sssAttributeRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Base 1

Public Enum RuleType
    WarningRule = 0
    ErrorRule = 1
End Enum

Public Enum ScriptType
    VBScript = 0
    RegularExpression = 1
End Enum

'***    unique identifier of the rule
Private m_lngId As Long
'***    rule type
Private m_RuleType As RuleType
'***    rule script
Private m_strRuleScript As String
'***    script type
Private m_ScriptType As ScriptType
'***    message
Private m_strMessage As String
'***    comment(brief description of rule for user)
Private m_strComment As String
'***    dll to load for special processing
Private m_strDll As String

Public Property Let Id(ByVal vData As Long)
    m_lngId = vData
End Property
Public Property Get Id() As Long
    Id = m_lngId
End Property
Public Property Let RuleType(ByVal vData As RuleType)
    m_RuleType = vData
End Property
Public Property Get RuleType() As RuleType
    RuleType = m_RuleType
End Property
Public Property Let RuleScript(ByVal vData As String)
    m_strRuleScript = vData
End Property
Public Property Get RuleScript() As String
    RuleScript = m_strRuleScript
End Property
Public Property Let ScriptType(ByVal vData As ScriptType)
    m_ScriptType = vData
End Property
Public Property Get ScriptType() As ScriptType
    ScriptType = m_ScriptType
End Property
Public Property Let Message(ByVal vData As String)
    m_strMessage = vData
End Property
Public Property Get Message() As String
    Message = m_strMessage
End Property
Public Property Let Comment(ByVal vData As String)
    m_strComment = vData
End Property
Public Property Get Comment() As String
    Comment = m_strComment
End Property
Public Property Let Dll(ByVal vData As String)
    m_strDll = vData
End Property
Public Property Get Dll() As String
    Dll = m_strDll
End Property
