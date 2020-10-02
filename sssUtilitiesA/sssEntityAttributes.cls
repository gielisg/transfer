VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "sssEntityAttributes"
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

'***    attribute configuration collection
Private m_colAttributeConfigs As Collection

'***    attribute values collection
Private m_colAttributes As Collection

'***    collection of rules to apply to the entity (entity level)
Private m_colRules As Collection

Private m_objUser As CUser

Public Property Let User(User As CUser)
    Set m_objUser = User
End Property
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
'Public Function GetAttributeByName(strName As String, Optional blnNew As Boolean = False) As sssAttribute
'    Dim objAttribute As sssAttribute
'    Dim objNewAttribute As sssAttribute
'    Dim idx As Integer
'    Dim blnFokjund As Boolean
'
'    If m_colAttributes Is Nothing Then Set m_colAttributes = New Collection
'
'    For idx = 1 To m_colAttributes.Count
'        Set objAttribute = m_colAttributes.Item(idx)
'        If UCase(objAttribute.Name) = UCase(strName) Then
'            blnFound = True
'            Exit For
'        End If
'    Next
'
'    If blnNew Then
'        Set objNewAttribute = New sssAttribute
'        objNewAttribute.AttributeDataType = objAttribute.AttributeDataType
'        objNewAttribute.AttributeType = objAttribute.AttributeType
'        objNewAttribute.AutonumberKey = objAttribute.AutonumberKey
'        objNewAttribute.DefaultValue = objAttribute.DefaultValue
'        objNewAttribute.EditType = objAttribute.EditType
'        objNewAttribute.AttributeDataType = objAttribute.AttributeDataType
'        objNewAttribute.AttributeDataType = objAttribute.AttributeDataType
'        objNewAttribute.AttributeDataType = objAttribute.AttributeDataType
'        objNewAttribute.AttributeDataType = objAttribute.AttributeDataType
'        objNewAttribute.AttributeDataType = objAttribute.AttributeDataType
'        objNewAttribute.AttributeDataType = objAttribute.AttributeDataType
'        objNewAttribute.AttributeDataType = objAttribute.AttributeDataType
'        objNewAttribute.AttributeDataType = objAttribute.AttributeDataType
'        objNewAttribute.AttributeDataType = objAttribute.AttributeDataType
'        objNewAttribute.AttributeDataType = objAttribute.AttributeDataType
'
'    If blnFound Then
'        Set GetAttributeByName = objAttribute
'    Else
'        Set GetAttributeByName = Nothing
'    End If
'
'End Function
Public Function AttributeConfigItem(ByVal Index As Variant) As sssAttributeConfig
    If m_colAttributeConfigs Is Nothing Then Set m_colAttributeConfigs = New Collection
    Set AttributeConfigItem = m_colAttributeConfigs(Index)
End Function
Public Function AttributeConfigCount() As Long
    If m_colAttributeConfigs Is Nothing Then Set m_colAttributeConfigs = New Collection
    AttributeConfigCount = m_colAttributeConfigs.Count
End Function

Friend Sub AddAttributeConfig(objItem As sssAttributeConfig, Optional strIndex As String)
    
    If m_colAttributeConfigs Is Nothing Then Set m_colAttributeConfigs = New Collection
    If Len(strIndex) = 0 Then
        m_colAttributeConfigs.Add objItem
    Else
        m_colAttributeConfigs.Add objItem, strIndex
    End If

End Sub
Public Function GetAttributeConfigByName(strName As String) As sssAttributeConfig
    Dim objAttributeConfig As sssAttributeConfig
    Dim idx As Integer
    Dim blnFound As Boolean
    
    If m_colAttributeConfigs Is Nothing Then Set m_colAttributeConfigs = New Collection
    
    For idx = 1 To m_colAttributeConfigs.Count
        Set objAttributeConfig = m_colAttributeConfigs.Item(idx)
        If UCase(objAttributeConfig.Name) = UCase(strName) Then
            blnFound = True
            Exit For
        End If
    Next
        
    If blnFound Then
        Set GetAttributeConfigByName = objAttributeConfig
    Else
        Set GetAttributeConfigByName = Nothing
    End If

End Function
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
Public Function Load(lngId As Long) As Boolean
    '***********************************************************************************************
    'Purpose    : Load the attribute configurations
    'Created on :
    'Created by: Gordon
    '***********************************************************************************************

    '-------------------------------
    '*** Generic Procedure Code
    '-------------------------------
    On Error GoTo ERROR_HANDLER
    Dim strModuleAndProcName As String
    strModuleAndProcName = "sssEntityAttributes - Load Sub"
    
    Dim dbAccess As dbAccess
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    Dim objRule As sssAttributeRule
    Dim objAttributeConfig As sssAttributeConfig
    Dim idx As Integer
    Dim blnFound As Boolean
    
    strSQL = "select * from entities_cnfg where entity_id = " & CStr(m_lngId)
    Set dbAccess = m_objUser.GetDbAccessObject(m_objUser.CurrentBusinessUnitCode)
    Set rs = dbAccess.LoadRS(strSQL)
    
    While Not rs.EOF
        blnFound = True
        m_strName = FieldToString(rs, "entity_narr")
        m_strIcon = FieldToString(rs, "icon")
        If Not IsNull(rs("menu_id")) Then m_lngPopupMenuId = rs("menu_id")
        m_strTooltip = FieldToString(rs, "tooltip")
        rs.MoveNext
    Wend
    
    If Not blnFound Then Err.Raise vbObjectError + 100, strModuleAndProcName, "Invalid Entity Id"
    
    '***    get the entity rules
    strSQL = "select * from rules a, ent_rule_rel b " & _
                "where a.rule_id = b.rule_id " & _
                "and b.entity_id = " & CStr(m_lngId)
    
    Set rs = dbAccess.LoadRS(strSQL)
    
    Set m_colRules = New Collection
    
    While Not rs.EOF
        
        Set objRule = New sssAttributeRule
        objRule.Id = rs("rule_id")
        objRule.RuleType = rs("rule_type")
        objRule.ScriptType = rs("script_type")
        objRule.RuleScript = FieldToString(rs, "rule_script")
        objRule.Message = FieldToString(rs, "rule_message")
        objRule.Comment = FieldToString(rs, "rule_comment")
        objRule.Dll = FieldToString(rs, "rule_dll")
        m_colRules.Add objRule
        
SKIP_RECORD:
        
        rs.MoveNext
    Wend
    
    '***    get the entity attribute configs
    strSQL = "select * from entity_attrib a, ent_attrib_rel b " & _
                "where a.attribute_id = b.attribute_id " & _
                "and b.entity_id = " & CStr(m_lngId) & " " & _
                "order by attribute_ord"
    
    Set m_colAttributeConfigs = New Collection
    
    Set rs = dbAccess.LoadRS(strSQL)
    
    While Not rs.EOF
        '***    check that the user has access rights to this attribute
        'TODO
        Set objAttributeConfig = New sssAttributeConfig
        objAttributeConfig.Id = rs("attribute_id")
        objAttributeConfig.Name = FieldToString(rs, "attribute_name")
        objAttributeConfig.Label = FieldToString(rs, "attribute_label")
        objAttributeConfig.DefaultValue = FieldToString(rs, "attrib_dflt_val")
        objAttributeConfig.Icon = FieldToString(rs, "attribute_icon")
        objAttributeConfig.AttributeType = rs("attrib_type")
        objAttributeConfig.EditType = rs("attrib_edit_typ")
        objAttributeConfig.AttributeDataType = rs("attrib_data_typ")
        objAttributeConfig.Mask = FieldToString(rs, "attribute_mask")
        objAttributeConfig.PopupMenuId = rs("attrib_menu_id")
        objAttributeConfig.Tooltip = FieldToString(rs, "attrib_tooltip")
        objAttributeConfig.FKSQL = FieldToString(rs, "attrib_fk_sql")
        objAttributeConfig.AutonumberKey = FieldToString(rs, "attrib_auto_key")
        objAttributeConfig.ListSQL = FieldToString(rs, "attrib_list_sql")
        
        m_colAttributeConfigs.Add objAttributeConfig
        
'SKIP_RECORD:

        rs.MoveNext
    Wend
    
    If m_colAttributeConfigs.Count = 0 Then GoTo EXIT_PROC
    
    '***    get the attri bute rules
    strSQL = "select * from rules a, ent_attrib_rule b " & _
                "where a.rule_id and b.rule_id " & _
                "and b.attribute_id in (select attribute_id " & _
                "from ent_attrib_rel " & _
                "where entity_id = " & CStr(m_lngId) & " ) " & _
                "order by attribute_id "
    
    Set rs = dbAccess.LoadRS(strSQL)
    
    While Not rs.EOF
        Set objRule = New sssAttributeRule
        objRule.Id = rs("rule_id")
        objRule.RuleType = rs("rule_type")
        objRule.ScriptType = rs("script_type")
        objRule.RuleScript = FieldToString(rs, "rule_script")
        objRule.Message = FieldToString(rs, "rule_message")
        objRule.Comment = FieldToString(rs, "rule_comment")
        objRule.Dll = FieldToString(rs, "rule_dll")
        
        '***    locate the appropriate attribute to add the rule to
        For idx = 1 To m_colAttributeConfigs.Count
            Set objAttributeConfig = m_colAttributeConfigs.Item(idx)
            If objAttributeConfig.Id = rs("attribute_id") Then
                objAttributeConfig.AddRule objRule
                Exit For
            End If
        Next
        
        rs.MoveNext
    Wend
    
EXIT_PROC:

    Load = blnFound
    
    Exit Function
    '------------------------------------------------------------
    '*** Generic Error Handling Code (ensure Exit Proc is above)
    '------------------------------------------------------------
ERROR_HANDLER:

    Dim lErrNum As Long
    Dim strErrDesc As String
    Dim strErrSource As String
    
    lErrNum = Err.Number
    strErrDesc = Err.Description
    strErrSource = "[" & App.Title & " - " & strModuleAndProcName & " - Build " & App.Major & "." & App.Minor & "." & App.Revision
    If Erl > 0 Then strErrSource = strErrSource & " - Line No " & str(Erl)
    If InStr(1, Err.Source, "-->") > 0 Then
        '*** This error has already been handled by our code
        strErrSource = strErrSource & "]" & vbNewLine & "  --> " & Err.Source
    Else
        '*** Newly generated error, log it here.
        strErrSource = strErrSource & "]" & vbNewLine & "  --> [Source: " & Err.Source & "]"
        On Error Resume Next
        LogError lErrNum, strErrSource, strErrDesc
    End If
    '(ALWAYS comment out two of the three options below)
    '*** EITHER display the error here
    'DisplayError lErrNum, strErrSource, strErrDesc
    '*** OR raise the error to the calling procedure
    Err.Raise lErrNum, strErrSource, strErrDesc
    '*** OR ignore the error and continue
    'Resume Next
End Function
'Public Function Groups() As Collection
'    If m_colAttributeGroups Is Nothing Then Set m_colAttributeGroups = New Collection
'    Set Groups = m_colAttributeGroups
'End Function
'Public Function GroupCount() As Long
'    If m_colAttributeGroups Is Nothing Then Set m_colAttributeGroups = New Collection
'    GroupCount = m_colAttributeGroups.Count
'End Function


