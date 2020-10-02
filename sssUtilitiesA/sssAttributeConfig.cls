VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "sssAttributeConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit
Option Base 1

Public Enum AttributeType
    RealAndList = 0
    RealOnly = 1
    ListOnly = 2
End Enum

Public Enum EditType
    NoEdit = 0
    InPlaceEdit = 1
    Dialog = 2
End Enum

Public Enum AttributeDataType
    StringType = 0
    DateType = 1
    BooleanType = 2
    ListType = 3
    CurrencyType = 4
    DecimalType = 5
    IntegerType = 6
    FileNameType = 7
    LongStringType = 8
End Enum

'***    attribute id (unique identifier for this attribute)
Private m_lngId As Long
'***    group id (allows attributes to be associated with each other)
Private m_lngGroupId As Long
'***    attribute name
Private m_strName As String
'***    label of the attribute
Private m_strLabel As String
Private m_varDefaultValue As Variant
Private m_AttributeType As AttributeType
Private m_EditType As EditType
Private m_AttributeDataType As AttributeDataType
'***    edit mask
Private m_strMask As String
Private m_strIcon As String
'***    associated menu id
Private m_lngPopupMenuId As Long
Private m_strTooltip As String
'***    validate data value from this list
Private m_dicListValues As Dictionary
Private m_strListSQL As String
'***    validate using this SQL
Private m_strFKSQL As String
'***    autonumber key
Private m_strAutonumberKey As String
'***    collection of rules to apply to attribute
Private m_colRules As Collection
'***    dispbold
Private m_blnDisplayBold As Boolean
'***    dispcolour
Private m_strDisplayColour As String

Public Property Let Id(ByVal vData As Long)
    m_lngId = vData
End Property
Public Property Get Id() As Long
    Id = m_lngId
End Property
Public Property Let GroupId(ByVal vData As Long)
    m_lngGroupId = vData
End Property
Public Property Get GroupId() As Long
    Id = m_lngGroupId
End Property
Public Property Let Name(ByVal vData As String)
    m_strName = vData
End Property
Public Property Get Name() As String
    Name = m_strName
End Property
Public Property Let Label(ByVal vData As String)
    m_strLabel = vData
End Property
Public Property Get Label() As String
    Label = m_strLabel
End Property

Public Property Let DefaultValue(ByVal vData As Variant)
    m_varDefaultValue = vData
End Property
Public Property Get DefaultValue() As Variant
    DefaultValue = m_varDefaultValue
End Property
Public Property Let EditType(ByVal vData As EditType)
    m_EditType = vData
End Property
Public Property Get EditType() As EditType
    EditType = m_EditType
End Property
Public Property Let AttributeType(ByVal vData As AttributeType)
    m_AttributeType = vData
End Property
Public Property Get AttributeType() As AttributeType
    AttributeType = m_AttributeType
End Property

Public Property Let AttributeDataType(ByVal vData As AttributeDataType)
    m_AttributeDataType = vData
End Property
Public Property Get AttributeDataType() As AttributeDataType
    AttributeDataType = m_AttributeDataType
End Property
Public Property Let Mask(ByVal vData As String)
    m_strMask = vData
End Property
Public Property Get Mask() As String
    Mask = m_strMask
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
Public Property Let ListValues(ByVal vData As Dictionary)
    Set m_dicListValues = vData
End Property
Public Property Get ListValues() As Dictionary
    If m_dicListValues Is Nothing Then Set m_dicListValues = New Dictionary
    Set ListValues = m_dicListValues
End Property
Public Property Let FKSQL(ByVal vData As String)
    m_strFKSQL = vData
End Property
Public Property Get FKSQL() As String
    FKSQL = m_strFKSQL
End Property
Public Property Let ListSQL(ByVal vData As String)
    m_strListSQL = vData
End Property
Public Property Get ListSQL() As String
    ListSQL = m_strListSQL
End Property
Public Property Let AutonumberKey(ByVal vData As String)
    m_strAutonumberKey = vData
End Property
Public Property Get AutonumberKey() As String
    AutonumberKey = m_strAutonumberKey
End Property
Public Function RuleItem(ByVal Index As Variant) As sssAttributeRule
    If m_colRules Is Nothing Then Set m_colRules = New Collection
    Set RuleItem = m_colRules(Index)
End Function
Public Function RuleCount() As Long
    If m_colRules Is Nothing Then Set m_colRules = New Collection
    RuleCount = m_colRules.Count
End Function

Public Property Let DisplayBold(ByVal vData As Boolean)
    m_blnDisplayBold = vData
End Property

Public Property Get DisplayBold() As Boolean
    DisplayBold = m_blnDisplayBold
End Property

Public Property Let DisplayColour(ByVal vData As String)
    m_strDisplayColour = vData
End Property

Public Property Get DisplayColour() As String
    DisplayColour = m_strDisplayColour
End Property
'Friend Function Clone(objNewAttribute As Object) As Object
'    '***    this clones the attribute object
'    objNewAttribute.Id = Me.Id
'    objnewattribute.
'End Function
Friend Sub AddRule(objItem As sssAttributeRule, Optional strIndex As String)
    
    If m_colRules Is Nothing Then Set m_colRules = New Collection
    If Len(strIndex) = 0 Then
        m_colRules.Add objItem
    Else
        m_colRules.Add objItem, strIndex
    End If
    
End Sub
