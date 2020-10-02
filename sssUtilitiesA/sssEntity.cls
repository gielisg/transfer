VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "sssEntity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit
Option Base 1

'***    entity id (unique identifier for this entity)
Private m_lngId As Long
'***    entity name
Private m_strName As String
'***    associated icon
Private m_strIcon As String
'***    associated menu id
Private m_lngPopupMenuId As Long
'***    tooltip
Private m_strTooltip As String

'***    attribute collection
Private m_colAttributes As Collection
'***    attribute groups
'Private m_colAttributeGroups As Collection
'***    collection of rules to apply to attribute
Private m_colRules As Collection

Public Property Let Id(ByVal vData As Long)
    m_lngId = vData
End Property
Public Property Get Id() As Long
    Id = m_lngId
End Property
Public Property Let Name(ByVal vData As String)
    m_strName = vData
End Property
Public Property Get Name() As String
    Name = m_strName
End Property
Public Property Let Icon(ByVal vData As String)
    m_strIcon = vData
End Property
Public Property Get Icon() As String
    Icon = m_strIcon
End Property
Public Property Let PopupMenuId(ByVal vData As Long)
    m_lngPopupMenuId = vData
End Property
Public Property Get PopupMenuId() As Long
    PopupMenuId = m_lngPopupMenuId
End Property
Public Property Let Tooltip(ByVal vData As String)
    m_strTooltip = vData
End Property
Public Property Get Tooltip() As String
    Tooltip = m_strTooltip
End Property

Public Function AttributeItem(ByVal Index As Variant) As sssAttribute
    If m_colAttributes Is Nothing Then Set m_colAttributes = New Collection
    Set AttributeItem = m_colAttributes(Index)
End Function
Public Function AttributeCount() As Long
    If m_colAttributes Is Nothing Then Set m_colAttributes = New Collection
    AttributeCount = m_colAttributes.Count
End Function

Friend Sub AddAttribute(objItem As sssAttribute, Optional strIndex As String)
    
    If m_colAttributes Is Nothing Then Set m_colAttributes = New Collection
    If Len(strIndex) = 0 Then
        m_colAttributes.Add objItem
    Else
        m_colAttributes.Add objItem, strIndex
    End If

End Sub
Public Function RuleItem(ByVal Index As Variant) As sssAttributeRule
    If m_colRules Is Nothing Then Set m_colRules = New Collection
    Set RuleItem = m_colRules(Index)
End Function
Public Function RuleCount() As Long
    If m_colRules Is Nothing Then Set m_colRules = New Collection
    RuleCount = m_colRules.Count
End Function

Friend Sub AddRule(objItem As sssAttributeRule, Optional strIndex As String)
    
    If m_colRules Is Nothing Then Set m_colRules = New Collection
    If Len(strIndex) = 0 Then
        m_colRules.Add objItem
    Else
        m_colRules.Add objItem, strIndex
    End If
End Sub

'Public Function Groups() As Collection
'    If m_colAttributeGroups Is Nothing Then Set m_colAttributeGroups = New Collection
'    Set Groups = m_colAttributeGroups
'End Function
'Public Function GroupCount() As Long
'    If m_colAttributeGroups Is Nothing Then Set m_colAttributeGroups = New Collection
'    GroupCount = m_colAttributeGroups.Count
'End Function


