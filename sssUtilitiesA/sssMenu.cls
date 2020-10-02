VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "sssMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Base 1

' ----------------------------------------------------------
' This class is a custom collection of MenuItem objects
' ----------------------------------------------------------
Private m_lngMenuId As Long
Private m_colItems As Collection
Private m_dbAccess As dbAccess
Private m_objUser As CUser
Private m_AccessRight As AccessRight

Public Property Get MenuId() As Long
    MenuId = m_lngMenuId
End Property
Public Property Let MenuId(lngMenuId As Long)
    m_lngMenuId = lngMenuId
End Property
Public Property Let dbAccess(dbAccess As dbAccess)
    Set m_dbAccess = dbAccess
End Property
Public Property Let User(User As CUser)
    Set m_objUser = User
End Property
Public Property Get AccessRight() As AccessRight
    AccessRight = m_AccessRight
End Property
Public Property Let AccessRight(vData As AccessRight)
    m_AccessRight = vData
End Property
Public Function item(ByVal Index As Variant) As sssMenuItem
Attribute item.VB_UserMemId = 0
  Set item = m_colItems(Index)
End Function
Public Function Count() As Long
  Count = m_colItems.Count
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
  Set NewEnum = m_colItems.[_NewEnum]
End Function

Friend Sub Add(objItem As sssMenuItem, Optional strIndex As String)

  If Len(strIndex) = 0 Then
    m_colItems.Add objItem
  Else
     m_colItems.Add objItem, strIndex
  End If
  
End Sub
'Public Function Exists(lngObjectId As Long) As Boolean
'    Dim intIdx As Integer
'    Dim blnExists As Boolean
'    For intIdx = 1 To m_colItems.Count
'        If m_colItems(intIdx).ObjectId = lngObjectId Then
'            blnExists = True
'            Exit For
'        End If
'    Next
'End Function
'Public Function GetAccessRight(lngObjectId As Long) As enAccessRight
'    Dim intIdx As Integer
'    Dim blnExists As Boolean
'    For intIdx = 1 To m_colItems.Count
'        If m_colItems.ObjectId = lngObjectId Then
'            blnExists = True
'            Exit For
'        End If
'    Next
'    If blnExists Then
'        GetAccessRight = m_colItems.AccessRight
'    Else
'        Err.Raise vbObjectError + 100, "AccessRights.GetAccessRight", "No Access Right determined for object."
'    End If
'End Function
Public Sub Remove(ByVal Index As Variant)
  m_colItems.Remove Index
End Sub

'this LoadMenuItems sub is replaced by LoadMenuItems2, keep for binary compatibility
Public Sub LoadMenuItems(Optional lngMenuId As Long)
    '***********************************************************************************************
    'Purpose    : Loads up the menu items for the menu
    'Created on :
    'Created by: Gordon
    '***********************************************************************************************

    '-------------------------------
    '*** Generic Procedure Code
    '-------------------------------
    On Error GoTo ERROR_HANDLER
    
    Dim strModuleAndProcName As String
    strModuleAndProcName = "Menu - LoadMenuItems Function"
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    Dim objMenuItem As sssMenuItem
    Dim lngGroupId As Long
    Dim intOrder As Integer
    Dim lngObjectId As Long
    Dim lngKey As Long
    Dim intAccess As Integer
    Dim idx As Integer
    
    '***    some validations
    If m_objUser Is Nothing Then
        Err.Raise vbObjectError + 100, "Menu.LoadMenuItems", "No user set."
    End If
    If lngMenuId = 0 And m_lngMenuId = 0 Then

        Err.Raise vbObjectError + 100, "Menu.LoadMenuItems", "No menu id set."

    End If

    If m_dbAccess Is Nothing Then
        Err.Raise vbObjectError + 100, "Menu.LoadMenuItems", "No dbAccess object set."
    End If
    
    '***    the SQL
    strSQL = "select menu_items.* " & _
                "from menu_items " & _
                "where menu_items.menu_id = " & CStr(m_lngMenuId) & " " & _
                "order by mnuitm_grp, mnuitm_order "
                
    Set rs = m_dbAccess.LoadRS(strSQL)
    
    While Not rs.EOF
        '***    groups 2,5,8 should only have one item
        lngGroupId = rs("mnuitm_grp")
        intOrder = rs("mnuitm_order")
        
'        If (lngGroupId = 2 Or lngGroupId = 5 Or lngGroupId = 8) _
'        And intOrder > 1 Then
'            'GoTo SKIP_RECORD
'        End If
        '***    if the object_id is null then we do not do access checking
        If IsNull(rs("object_id")) Then
            intAccess = enAccessRight.FullAccess
        Else
            lngObjectId = rs("object_id")
        
            '***    check the access for this menuitem object
            intAccess = m_objUser.GetAccessRights(lngObjectId)
            If intAccess = enAccessRight.NoAccess Then GoTo SKIP_RECORD
        End If
        '***    use lngKey as the collection key
        lngKey = rs("mnuitm_id")
        
        Set objMenuItem = New sssMenuItem
        objMenuItem.Caption = FieldToString(rs, "mnuitm_caption")
        objMenuItem.cmd = FieldToString(rs, "mnuitm_cmd")
        objMenuItem.Dll = FieldToString(rs, "mnuitm_dll")
        objMenuItem.Group = lngGroupId
        objMenuItem.ObjectId = lngObjectId
        objMenuItem.Order = intOrder
        objMenuItem.Tooltip = FieldToString(rs, "mnuitm_tooltip")
        If FieldToString(rs, "mnuitm_check") = "Y" Then
            objMenuItem.Checked = True
        Else
            objMenuItem.Checked = False
        End If
        '***    decide if the menu item is enabled
        If intAccess = enAccessRight.FullAccess Then
            objMenuItem.Enabled = True
        Else
            objMenuItem.Enabled = False
        End If
        '***    check the default
        If FieldToString(rs, "mnuitm_default") = "Y" Then
            objMenuItem.Default = True
        Else
            objMenuItem.Default = False
        End If
        
        Me.Add objMenuItem, CStr(lngKey)
        
SKIP_RECORD:
        
        rs.MoveNext
    Wend
    
    '***    if last menu item caption = "-" (a seperator) then delete
    For idx = Me.Count To 1 Step -1
        Set objMenuItem = Me.item(idx)
        
        '***    is it a seperator
        If objMenuItem.Caption = "-" Then
            Me.Remove (idx)
        Else
            GoTo EXIT_SUB
        End If
    Next
    
EXIT_SUB:

    Exit Sub
    
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
End Sub
Public Sub LoadMenuItems2(Optional lngMenuId As Long, Optional strContactCode As String, _
    Optional lngSpCnRef As Long, Optional strServiceTypeCode As String, _
    Optional strServicestatusCode As String)
    '***********************************************************************************************
    'Purpose    : Loads up the menu items for the menu
    'Created on :
    'Created by: Gordon
    '***********************************************************************************************

    '-------------------------------
    '*** Generic Procedure Code
    '-------------------------------
    On Error GoTo ERROR_HANDLER
    
    Dim strModuleAndProcName As String
    strModuleAndProcName = "Menu - LoadMenuItems2 Function"
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    Dim objMenuItem As sssMenuItem
    Dim lngGroupId As Long
    Dim intOrder As Integer
    Dim lngObjectId As Long
    Dim lngKey As Long
    Dim intAccess As Integer
    Dim idx As Integer
    
    '***    some validations
    If m_objUser Is Nothing Then
        Err.Raise vbObjectError + 100, "Menu.LoadMenuItems", "No user set."
    End If
    If lngMenuId = 0 Then
        If m_lngMenuId = 0 Then
            Err.Raise vbObjectError + 100, "Menu.LoadMenuItems", "No menu id set."
        Else
            lngMenuId = m_lngMenuId
        End If
    End If
    If m_dbAccess Is Nothing Then
        Err.Raise vbObjectError + 100, "Menu.LoadMenuItems", "No dbAccess object set."
    End If
    
    '***    the SQL
'    strSQL = "select menu_items.* " & _
'                "from menu_items " & _
'                "where menu_items.menu_id = " & CStr(m_lngMenuId) & " " & _
'                "order by mnuitm_grp, mnuitm_order "
                
    strSQL = "execute procedure ss_menu_items(p_menu_id = " & lngMenuId & _
        ", p_role_id = " & m_objUser.CurrentRoleID
        
    If strContactCode <> "" Then
        strSQL = strSQL & ", p_contact_code = '" & strContactCode & "'"
    ElseIf lngSpCnRef <> 0 Then
        strSQL = strSQL & ", p_sp_cn_ref = " & lngSpCnRef
    Else
        If strServiceTypeCode <> "" Then
            strSQL = strSQL & ", p_srvctypecode = " & strServiceTypeCode
        End If
        If strServicestatusCode <> "" Then
            strSQL = strSQL & ", p_sys_status = " & strServicestatusCode
        End If
    End If
    
    strSQL = strSQL & ", resume_flg = 'Y', isolation_flg = 'D')"
    
    Set rs = m_dbAccess.LoadRS(strSQL)
    
    While Not rs.EOF
        If rs(0) = 0 Then
            '***    use lngKey as the collection key
            lngKey = rs(1)
            
            Set objMenuItem = New sssMenuItem
            objMenuItem.Caption = FieldToString(rs, "mnuitm_caption")
            objMenuItem.cmd = FieldToString(rs, "mnuitm_cmd")
            objMenuItem.Dll = FieldToString(rs, "mnuitm_dll")
            objMenuItem.Group = FieldToString(rs, "mnuitm_grp")
            objMenuItem.ObjectId = Val(FieldToString(rs, "object_id"))
            objMenuItem.Order = FieldToString(rs, "mnuitm_order")
            objMenuItem.Tooltip = FieldToString(rs, "mnuitm_tooltip")
            If FieldToString(rs, "mnutim_check") = "Y" Then
                objMenuItem.Checked = True
            Else
                objMenuItem.Checked = False
            End If
            objMenuItem.Enabled = True
            '***    check the default
            If FieldToString(rs, "mnuitm_default") = "Y" Then
                objMenuItem.Default = True
            Else
                objMenuItem.Default = False
            End If
            
            If FieldToString(rs, "modal_flg") = "N" Then
                objMenuItem.Modality = False
            Else
                objMenuItem.Modality = True
            End If
            
            objMenuItem.ActionFlag = FieldToString(rs, "action_flg")
            
            Me.Add objMenuItem, CStr(lngKey)
        End If
SKIP_RECORD:
        
        rs.MoveNext
    Wend
    
    '***    if last menu item caption = "-" (a seperator) then delete
    For idx = Me.Count To 1 Step -1
        Set objMenuItem = Me.item(idx)
        
        '***    is it a seperator
        If objMenuItem.Caption = "-" Then
            Me.Remove (idx)
        Else
            GoTo EXIT_SUB
        End If
    Next
    
EXIT_SUB:

    Exit Sub
    
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
End Sub
Public Sub Clear()
  Do While m_colItems.Count > 0
    m_colItems.Remove (1)
  Loop
End Sub

Private Sub Class_Initialize()
  Set m_colItems = New Collection
End Sub

Private Sub Class_Terminate()
  Set m_colItems = Nothing
End Sub

