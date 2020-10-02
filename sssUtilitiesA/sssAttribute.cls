VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "sssAttribute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit
Option Base 1

'***    the Id of the Entity Attribute
Private m_strId As String
'***    the attribute configuration
Private m_sssAttributeConfig As sssAttributeConfig
'***    the attribute value
Private m_varValue As Variant
'***    the attribute label (we sometimes need to override the config label)
Private m_strLabel As String
Public Property Let Id(ByVal vData As String)
    m_strId = vData
End Property
Public Property Get Id() As String
    Id = m_strId
End Property
Public Property Set AttributeConfig(vData As sssAttributeConfig)
    Set m_sssAttributeConfig = vData
    '***    default the label but we can override
    m_strLabel = vData.Label
End Property
Public Property Get AttributeConfig() As sssAttributeConfig
    Set AttributeConfig = m_sssAttributeConfig
End Property
Public Property Let Value(ByVal vData As Variant)
    m_varValue = vData
End Property
Public Property Get Value() As Variant
    Value = m_varValue
End Property
Public Property Let Label(ByVal vData As String)
    m_strLabel = vData
End Property
Public Property Get Label() As String
    Label = m_strLabel
End Property

