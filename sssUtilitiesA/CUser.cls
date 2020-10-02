VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"ConnectionConfiguration"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Option Base 1

Const ODBCKey As String = "Software\ODBC\ODBC.INI"
Const ODBCSubKey As String = "ODBC Data Sources"

Private Const ConnectionTimeOutKey As String = "DB_CONNECTION_TIMEOUT"

Public Enum enLogonReturnCode
  icSuccess = 0
  icInvalidUserName = 1
  icFailure = 2
  icNotLoggedOn = 3
  icConfigurationLoadFailure = 4
  icRoleLoadFailure = 5
  icConnectionConfigLoadFailure = 6
  icBusinessUnitLoadFailure = 7
  icResetPassword = 8
  icDatabaseFailure = 9
  icLicenseConfigFailure = 10
  icLicenseFailure = 11
  icInvalidIP = 12
  icLockOut = 13
End Enum

Public Type typAccessRight
    ObjectId As Long
    ObjectNarr As String
    AccessRight As AccessRights
End Type

Public Type typDatasource
    BusinessUnitCode As String
    DataSourceCode As String
    DataSourceLocal As Boolean
    DataSourceName As String
    DataSourceNarr As String
End Type
'keep binary compatibility
'the above type is not used
Public Type typDatasourceNew
    BusinessUnitCode As String
    DataSourceCode As String
    DataSourceLocal As Boolean
    DataSourceName As String
    DataSourceNarr As String
    BusinessUnitNarr As String
End Type

Public Type typDatasourceNew2
    BusinessUnitCode As String
    DataSourceCode As String
    DataSourceLocal As Boolean
    DataSourceName As String
    DataSourceNarr As String
    BusinessUnitNarr As String
    BusUnitColour As String
End Type

Public Enum ArgumentType
    Primitive = 0
    object = 1
End Enum

'***    properties
Private m_strLogonCode As String
Private m_strFirstName As String
Public m_strAppName As String
Private m_strName As String
Private m_strContactCode As String
Private m_strEmail As String
Private m_intLevel As Integer
Private m_strPassword As String
Private m_strUnixPassword As String
Private m_strUserType As String
Private m_strDefaultLocationCode As String
Private m_strDefaultBusinessUnitCode As String
Private m_blnLicense As Boolean
Private m_blnSetDefaultRole As Boolean
Private m_blnRoleChanged As Boolean
Private m_lngConnectionTimeout As Long
Private m_strLoginErrorMessage As String
'***    Current Role
Private m_lngRoleId As Long
Private m_blnAdminRole As Boolean
'***    flag to indicate that we use the one set of menues across all
'***    business units
Private m_blnGlobalMenues As Boolean
'***    flag to indicate that we use the one set of access rights across
'***    all business units
Private m_blnGlobalAccess As Boolean
'***    flag to indicate that we use the one set of entity configs
Private m_blnGlobalEntity As Boolean
'***    Master object id (used for loading access rights upon logon)
Private m_lngMasterObjectId As Long
'***    Connection Configuration object
Private m_ConnectionConfiguration As ConnectionConfiguration
'***    command line (seem to need to pass this around)
Private m_strCommandLine As String
'***    Initial configuration
Private m_enmInitialAccessMode As AccessMode
Private m_strInitialDSN As String
Private m_strInitialURL As String
Private m_strDatabaseName As String
Private m_strDBInstance As String
Private m_strDBHost As String
Private m_strDBService As String
Private m_strConnectMode As String
Private m_intNumberOfRoles As Integer
Private m_strDealerID As String
Private m_strIPAddress As String
Private m_strComputerName As String
Private m_blnNoPassword As Boolean

'***    current business unit
Private m_strCurrentBusinessUnitCode As String
Private m_strCurrentBusinessUnit As String

'***    Logon status
Private m_enLogonStatus As enLogonReturnCode
Private m_blnAutoLogon As Boolean
Private m_iEncryptMethod As Integer
'***    Parent Roles for the current role
Private m_Roles As sssRoles
'***    Access Rights
Private m_AccessRights As AccessRights

'***    Database access objects
Private m_dicDbAccessObjects As Dictionary
'***    Roles
Private m_rsRoles As ADODB.Recordset
Private m_RoleList As Collection
'***    Menues
Private m_dicMenues As Dictionary
'***    Arguments (used to transfer arguments to external modules)
Private m_dicArguments As Dictionary
'***    Business Units
Private m_dicBusinessUnits As Dictionary
'***    Configurations
Private m_Configurations As Configurations
'***    Entity Configurations
Private m_dicEntities As Dictionary

Private m_colDataSource As Collection
Private m_blnSwitchDatabase As Boolean

'***    Events
Event LoadConfigurationStart()
Event LoadConfigurationFinish()

Event LoadRolesStart()
Event LoadRolesFinish()
Event RoleSelectionRequired()
Event RoleSelectionFinish()
Event LoadParentRoleStart()
Event LoadParentRoleFinish()

Event LoadMenuDataStart()
Event LoadMenuDataFinish()

Event LoadEntityDataStart()
Event LoadEntityDataFinish()

Event LoadLicenseStart()
Event LoadLicenseFinsish()

Event LoadIPaddressStart()
Event LoadIPAddressFinish()

Event ChangePasswordStart()
Event ChangePasswordFinish()
Event LogonStart()
Event LogonFailed()
Event LogonInvalidUserName()
Event LogonFinish()
Event StatusBarText(strText As String)

Event MenuCreated(lMenuId As Long, sMenuName As String)
Event MenuDestroyed(lMenuId As Long, sMenuName As String)
Event MRUReload(lMRU As Long)

Event SendMessage(sMessage As String)
Event ReceiveMessage(sMessage As String)

Event SwitchDataSource()

Public Sub RaiseMRUReload(lMRU As Long)
    RaiseEvent MRUReload(lMRU)
End Sub
Public Sub AddArgument(ByVal strKey As String, varValue As Variant)
    If m_dicArguments Is Nothing Then Set m_dicArguments = New Dictionary
    If m_dicArguments.Exists(strKey) Then m_dicArguments.Remove (strKey)
    m_dicArguments.Add strKey, varValue
End Sub
Public Sub RemoveAllArguments()
    If Not m_dicArguments Is Nothing Then
        m_dicArguments.RemoveAll
    End If
End Sub
Public Function ArgumentExists(ByVal strKey As String) As Boolean
    If m_dicArguments Is Nothing Then Set m_dicArguments = New Dictionary
    If m_dicArguments.Exists(strKey) Then
        ArgumentExists = True
    Else
        ArgumentExists = False
    End If
End Function
Public Function ArgumentItem(ByVal strKey As String) As Variant
    If m_dicArguments Is Nothing Then Set m_dicArguments = New Dictionary
    If Not m_dicArguments.Exists(strKey) Then
        Err.Raise vbObjectError + 100, "CUser.ArgumentItem", "Invalid argument."
    End If
    If IsObject(m_dicArguments.item(strKey)) Then
        Set ArgumentItem = m_dicArguments.item(strKey)
    Else
        ArgumentItem = m_dicArguments.item(strKey)
    End If
End Function
Public Function GetAccessRights(lngObjectId As Long) As enAccessRight
    '***********************************************************************************************
    'Purpose    : Returns the access rights token for this user for a particular
    '               object.
    'Created on :
    'Created by: Gordon
    '***********************************************************************************************

    '-------------------------------
    '*** Generic Procedure Code
    '-------------------------------
    On Error GoTo ERROR_HANDLER
    
    Dim strModuleAndProcName As String
    strModuleAndProcName = "CUser - AccessRights Function"
    Dim blnFound As Boolean
    Dim intCounter As Integer
    Dim strSQL As String
    Dim rs As ADODB.Recordset
    Dim dbAccess As dbAccess
    Dim AccessRight As AccessRight
    
    '***    administrators have complete rights
    If m_lngRoleId = 1 Then
        GetAccessRights = FullAccess
        GoTo EXIT_PROC
    End If
    '***    see if we have already fetched the access rights for this
    '***    object from the database. If not load the access rights.
'    If Not m_AccessRights.Exists(CStr(lngObjectId)) Then
'        ***    TODO write individual access rights fetch routine
'        ***    We need to see if there are any permissions for either ourselves ("A" part
'        ***    of the SQL) or for other roles ("B" part of the select).
'        ***    The order by clause will ensure that we process the access permissions
'        ***    set for this role before processing those for other roles.
'        strSQL = "select unique 'A', objects.object_id, token_id, object_narr " & _
'                "from role_obj_perm, objects " & _
'                "where role_id in (" & m_Roles.InClause & ") " & _
'                "and role_obj_perm.object_id = objects.object_id " & _
'                "and objects.object_id = ? " & _
'                "union " & _
'                "select 'A', objects.object_id, token_id, object_narr " & _
'                "from role_obj_perm, objects, obj_rel " & _
'                "where role_id in (" & m_Roles.InClause & ") " & _
'                "and role_obj_perm.object_id = objects.object_id " & _
'                "and objects.object_id = obj_rel.child_object_id " & _
'                "and obj_rel.par_object_id = ? "
'        strSQL = strSQL & _
'                " union " & _
'                "select unique 'B', objects.object_id, token_id, object_narr " & _
'                "from role_obj_perm, objects " & _
'                "where role_id not in (" & m_Roles.InClause & ") " & _
'                "and role_obj_perm.object_id = objects.object_id " & _
'                "and objects.object_id = ? " & _
'                "union " & _
'                "select 'B', objects.object_id, token_id, object_narr " & _
'                "from role_obj_perm, objects, obj_rel " & _
'                "where role_id not in (" & m_Roles.InClause & ") " & _
'                "and role_obj_perm.object_id = objects.object_id " & _
'                "and objects.object_id = obj_rel.child_object_id " & _
'                "and obj_rel.par_object_id = ? " & _
'                "order by 1 "
'        Set dbAccess = GetDbAccessObject(m_strDefaultBusinessUnitCode)
'        Set rs = dbAccess.LoadRS(strSQL, lngObjectId, lngObjectId, lngObjectId, lngObjectId)
'
'        blnFound = False
'
'        While Not rs.EOF
'            Set AccessRight = New AccessRight
'            AccessRight.ObjectId = CStr(rs(1))
'            ***    Add the access objects that are defined for this role
'            If rs(0) = "A" Then
'                If rs(2) = 5 Then
'                    AccessRight.AccessRight = FullAccess
'                ElseIf rs(2) = 3 Then
'                    AccessRight.AccessRight = ReadOnly
'                Else
'                    AccessRight.AccessRight = NoAccess
'                End If
'            Else
'                AccessRight.AccessRight = NoAccess
'            End If
'            AccessRight.ObjectNarr = FieldToString(rs, "object_narr")
'
'            If AccessRight.ObjectId = lngObjectId Then blnFound = True
'
'            ***    if the access object exists then we compare the level
'            ***    of access we already have to the new level. If the new
'            ***    level is greater then we replace the existing access
'            ***    with the new
'            If m_AccessRights.Exists(CStr(AccessRight.ObjectId)) Then
'                If m_AccessRights(CStr(rs(1))).AccessRight < AccessRight.AccessRight Then
'                    m_AccessRights.Remove CStr(rs(1))
'                    ***    add it
'                    m_AccessRights.Add AccessRight, CStr(AccessRight.ObjectId)
'                End If
'            Else
'                ***    add it
'                m_AccessRights.Add AccessRight, CStr(AccessRight.ObjectId)
'            End If
'
'            rs.MoveNext
'        Wend
'
'        ***    give full access if nothing in database
'        If Not blnFound Then
'            GetAccessRights = FullAccess
'            Set AccessRight = New AccessRight
'            AccessRight.ObjectId = AccessRight.ObjectId
'            AccessRight.AccessRight = FullAccess
'            AccessRight.ObjectNarr = ""
'            m_AccessRights.Add AccessRight, CStr(lngObjectId)
'        End If
'    End If
    
    If Not m_AccessRights.Exists(CStr(lngObjectId)) Then
        strSQL = "execute procedure ss_token_id (p_role_id = ?, p_object_id = ?)"
        Set dbAccess = GetDbAccessObject(m_strDefaultBusinessUnitCode)
        Set rs = dbAccess.LoadRS(strSQL, CurrentRoleID, lngObjectId)
        
        Set AccessRight = New AccessRight
        AccessRight.ObjectId = lngObjectId
        AccessRight.AccessRight = rs(1)
        
        m_AccessRights.Add AccessRight, CStr(AccessRight.ObjectId)
    End If
        
    GetAccessRights = m_AccessRights.GetAccessRight(lngObjectId)
    
EXIT_PROC:
    
    '------------------------------------------------------------
    '*** Generic Error Handling Code (ensure Exit Proc is above)
    '------------------------------------------------------------
    Exit Function
    
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

Public Property Get LoginErrorMessage() As String
    LoginErrorMessage = m_strLoginErrorMessage
End Property
Public Property Get MessageServerIP() As String
    MessageServerIP = "10.1.13.208"
End Property
Public Property Get MessageServerPort() As Long
    MessageServerPort = 9215
End Property
Public Property Let GlobalAccess(blnGlobal As Boolean)
    m_blnGlobalAccess = blnGlobal
End Property
Public Property Get GlobalAccess() As Boolean
    GlobalAccess = m_blnGlobalAccess
End Property
Public Property Let GlobalMenu(blnGlobal As Boolean)
    m_blnGlobalMenues = blnGlobal
End Property
Public Property Get GlobalMenu() As Boolean
    GlobalMenu = m_blnGlobalMenues
End Property
Public Property Let GlobalEntity(blnGlobal As Boolean)
    m_blnGlobalEntity = blnGlobal
End Property
Public Property Get GlobalEntity() As Boolean
    GlobalEntity = m_blnGlobalEntity
End Property
Public Property Get CurrentBusinessUnitCode() As String
    CurrentBusinessUnitCode = m_strCurrentBusinessUnitCode
End Property
Public Property Let CurrentBusinessUnitCode(ByVal strCurrentBusinessUnitCode As String)
    m_strCurrentBusinessUnitCode = strCurrentBusinessUnitCode
End Property
Public Property Get CurrentBusinessUnit() As String
    CurrentBusinessUnit = m_strCurrentBusinessUnit
End Property
Public Property Let CurrentBusinessUnit(ByVal strCurrentBusinessUnit As String)
    m_strCurrentBusinessUnit = strCurrentBusinessUnit
End Property
Public Property Get DBDatabase() As String
    DBDatabase = m_strDatabaseName
End Property
Public Property Get DBInstance() As String
    Dim strInformixServerSubkey As String
    
    strInformixServerSubkey = "Software\Informix\Environment"

    If Trim(m_strDBInstance) = "" Then
        m_strDBInstance = QueryValue(HKEY_LOCAL_MACHINE, strInformixServerSubkey, "INFORMIXSERVER")
        If Trim(m_strDBInstance) = "" Then
            m_strDBInstance = QueryValue(HKEY_CURRENT_USER, strInformixServerSubkey, "INFORMIXSERVER")
        End If
    End If
    DBInstance = m_strDBInstance
End Property
Public Property Get DBHost() As String
    Dim strInformixHostSubkey As String
    
    strInformixHostSubkey = "SOFTWARE\Informix\SqlHosts\" & Trim(DBInstance)
    
    If Trim(m_strDBHost) = "" Then
        m_strDBHost = QueryValue(HKEY_LOCAL_MACHINE, strInformixHostSubkey, "HOST")
    End If
    DBHost = m_strDBHost
End Property
Public Property Get DBService() As String
    Dim strInformixHostSubkey As String
    strInformixHostSubkey = "Software\Informix\SqlHosts\" & Trim(DBInstance)
    
    If Trim(m_strDBService) = "" Then
        m_strDBService = QueryValue(HKEY_LOCAL_MACHINE, strInformixHostSubkey, "SERVICE")
    End If
    DBService = m_strDBService
End Property
Public Property Get InitialDSN() As String
  InitialDSN = m_strInitialDSN
End Property
Public Property Get InitialURL() As String
  InitialURL = m_strInitialURL
End Property
Public Property Get InitialAccessMode() As AccessMode
  InitialAccessMode = m_enmInitialAccessMode
End Property
Public Property Get LogonStatus() As enLogonReturnCode
  LogonStatus = m_enLogonStatus
End Property
Public Property Let LogonStatus(enStatus As enLogonReturnCode)
  m_enLogonStatus = enStatus
End Property
Public Property Get LogonCode() As String
  LogonCode = m_strLogonCode
End Property
Public Property Let LogonCode(ByVal strValue As String)
  m_strLogonCode = strValue
End Property
Public Property Get FirstName() As String
  FirstName = m_strFirstName
End Property
Public Property Get Name() As String
  Name = m_strName
End Property
Public Property Get ContactCode() As String
  ContactCode = m_strContactCode
End Property
Public Property Get Email() As String
  Email = m_strEmail
End Property
Public Property Let Email(ByVal strValue As String)
  m_strEmail = strValue
End Property
Public Property Get ConnectionTimeout() As Long
  ConnectionTimeout = m_lngConnectionTimeout
End Property
Public Property Let ConnectionTimeout(ByVal strValue As Long)
  m_lngConnectionTimeout = strValue
End Property

Public Property Get Level() As String
  Level = m_intLevel
End Property

Public Property Get Password() As String
  Password = m_strPassword
End Property

Public Property Let Password(ByVal strValue As String)
  m_strPassword = strValue
End Property

Public Property Get UnixPassword() As String
  UnixPassword = m_strUnixPassword
End Property

Public Property Get UserType() As String
  UserType = m_strUserType
End Property
Public Property Get DefaultLocationCode() As String
  DefaultLocationCode = m_strDefaultLocationCode
End Property
Public Property Get DefaultBusinessUnitCode() As String
    DefaultBusinessUnitCode = m_strDefaultBusinessUnitCode
End Property
Public Property Get ConnectionConfiguration() As ConnectionConfiguration
    If m_ConnectionConfiguration Is Nothing Then LoadConnectionConfiguration
    Set ConnectionConfiguration = m_ConnectionConfiguration
End Property
Public Property Get Component(ByVal strEntity As String, Optional ByVal strBusinessUnitCode As String) As String
    If m_Configurations Is Nothing Then LoadConfigurations
    Component = m_Configurations.GetComponent(strEntity, strBusinessUnitCode)
End Property
Public Function CheckLicence() As Boolean
    '***********************************************************************************************
    'Purpose    : Basic Licence Checking
    'Created on :
    'Created by: Gordon
    '***********************************************************************************************

    '-------------------------------
    '*** Generic Procedure Code
    '-------------------------------
    On Error GoTo ERROR_HANDLER
    Dim strModuleAndProcName As String
    strModuleAndProcName = "CUser - CheckLicence Function   "
    
    '***    see if the licencing key file is
    
'    '***    translate the form
'    TranslateForm frm
'    Dim sLogon As String
'
'    '***    fetch the last successful logon
'    sLogon = GetSetting("Selcomm", "LogonCode", "LogonCode", "")
'    frm.txtUserName.Text = sLogon
'
'    '***    open the form
'    frm.Show vbModal
'
'    '***    null the form
'    Unload frm
'
EXIT_PROC:

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
Public Property Get CommandLine() As String
    CommandLine = m_strCommandLine
End Property
Public Property Let CommandLine(ByVal strValue As String)
    m_strCommandLine = strValue
End Property
Public Property Get NumberOfRoles() As Integer
    NumberOfRoles = m_intNumberOfRoles
End Property
Public Property Get Roles() As ADODB.Recordset
    m_rsRoles.MoveFirst
    Set Roles = m_rsRoles
End Property
Public Property Get RoleLists() As Collection
    Set RoleLists = m_RoleList
End Property
Public Property Get BusinessUnits() As Dictionary
    Set BusinessUnits = m_dicBusinessUnits
End Property
Public Property Get CurrentRole() As String
    Dim RoleList As sssRoleList
    Dim idx As Integer
    If m_RoleList Is Nothing Then
        CurrentRole = ""
        Exit Property
    End If
    
    For idx = 1 To m_RoleList.Count
        Set RoleList = m_RoleList(idx)
        If RoleList.RoleID = m_lngRoleId Then
            CurrentRole = RoleList.RoleNarr
            Exit For
        End If
    Next
End Property
Public Property Get CurrentRoleID() As Integer
    CurrentRoleID = m_lngRoleId
End Property
Public Property Get IsAdministrator() As Boolean
    IsAdministrator = m_blnAdminRole
End Property
Public Property Get AutoLogon() As Boolean
    AutoLogon = m_blnAutoLogon
End Property
Public Property Let AutoLogon(blnAutoLogon As Boolean)
    m_blnAutoLogon = blnAutoLogon
End Property
Public Property Get SetDefaultRole() As Boolean
    SetDefaultRole = m_blnSetDefaultRole
End Property
Public Property Let SetDefaultRole(blnSetDefaultRole As Boolean)
    m_blnSetDefaultRole = blnSetDefaultRole
End Property
Public Property Get RoleChanged() As Boolean
    RoleChanged = m_blnRoleChanged
End Property
Public Property Let RoleChanged(blnRoleChanged As Boolean)
    m_blnRoleChanged = blnRoleChanged
End Property
Public Property Get DealerID() As String
    DealerID = m_strDealerID
End Property
Public Property Let IPAddress(strIPAddress As String)
    m_strIPAddress = strIPAddress
End Property
Public Property Get DataSource() As Collection
    If m_colDataSource Is Nothing Then
        Set m_colDataSource = New Collection
        LoadDataSource
    End If
    Set DataSource = m_colDataSource
End Property
Public Property Get ComputerName() As String
    ComputerName = m_strComputerName
End Property
Public Property Get NoPassword() As Boolean
    NoPassword = m_blnNoPassword
End Property

Public Function IsBusinessUnitAvailable(strBusinessUnitCode As String) As Boolean
    If m_dicBusinessUnits.Exists(strBusinessUnitCode) Then
        IsBusinessUnitAvailable = True
    Else
        IsBusinessUnitAvailable = False
    End If
End Function

'Public Function Logon() As enLogonReturnCode
'    '***********************************************************************************************
'    'Purpose    : Logon the user
'    'Created on :
'    'Created by: Gordon
'    '***********************************************************************************************
'
'    '-------------------------------
'    '*** Generic Procedure Code
'    '-------------------------------
'    On Error GoTo ERROR_HANDLER
'    Dim strModuleAndProcName As String
'    strModuleAndProcName = "CUser - Logon Function"
'
'    Static idx As Integer
'
'    Dim strSQL As String
'    Dim rs As ADODB.Recordset
'    Dim strErrorDescription As String
'    Dim dbAccess As dbAccess
'    Dim blnFound As Boolean
'    Dim strPassword As String
'    Dim strTmp As String
'
'    '***    raise the event
'    RaiseEvent LogonStart
'
'    '***    fetch the initial parameters from the command line
'    On Error GoTo COMMAND_LINE_FAILURE
'    'only read connamd line when first log on
'    If idx = 0 Then
'        ReadCommand
'        idx = idx + 1
'    End If
'
'    On Error GoTo ERROR_HANDLER
'
'    m_enLogonStatus = icFailure
'
'    strSQL = "select a.contact_code, a.firstname, a.name, " & _
'            "b.operator_lvl, b.passwd, b.user_type, b.email, b.dflt_location, " & _
'            "b.dflt_bus_unit, encrypt_method " & _
'            "from contact a, operator b " & _
'            "where a.contact_code = b.contact_code " & _
'            "and login_code = ? "
'
'    '***    create the database access object
'    m_enLogonStatus = icDatabaseFailure
'    Set dbAccess = New dbAccess
'    dbAccess.AccessMode = m_enmInitialAccessMode
'    dbAccess.UserName = m_strLogonCode
'    dbAccess.Password = m_strPassword
'    dbAccess.DSN = m_strInitialDSN
'    dbAccess.URL = m_strInitialURL
'    Set rs = dbAccess.LoadRS(strSQL, m_strLogonCode)
'
'    While Not rs.EOF
'        blnFound = True
'        '***    set the members
'        m_strContactCode = FieldToString(rs, "contact_code")
'        m_strFirstName = FieldToString(rs, "firstname")
'        m_strName = FieldToString(rs, "name")
'        If Not IsNull(rs("operator_lvl")) Then m_intLevel = rs("operator_lvl")
'        strPassword = FieldToString(rs, "passwd")
'        m_strUserType = FieldToString(rs, "user_type")
'        m_strEmail = FieldToString(rs, "email")
'        m_strDefaultLocationCode = FieldToString(rs, "dflt_location")
'        If m_strConnectMode <> "C" Then m_strDefaultBusinessUnitCode = FieldToString(rs, "dflt_bus_unit")
'        m_iEncryptMethod = FieldToString(rs, "encrypt_method")
'
'        rs.MoveNext
'    Wend
'
'    Set rs = Nothing
'    Set dbAccess = Nothing
'
'    '***    if no record then invalid logon code
'    If Not blnFound Then
'        m_enLogonStatus = icInvalidUserName
'        RaiseEvent LogonInvalidUserName
'        GoTo EXIT_PROC
'    End If
'
'    '***    compare the password from the database to that entered by the user
'    '***    if the password is encrypted in the database then decrypt
'    If m_iEncryptMethod = 1 Then
'        Dim objMD5 As New CMD5
'
'        strTmp = objMD5.MD5(m_strPassword)
'    Else
'        strTmp = m_strPassword
'    End If
'
'    If strPassword = strTmp Then
'        RaiseEvent LogonFinish
'
'        '***    load connection configuration
'        m_enLogonStatus = icConnectionConfigLoadFailure
'        LoadConnectionConfiguration
'        '***    load IP
'        m_enLogonStatus = icInvalidIP
'        LoadIPAddress
'        '***    load business units
'        m_enLogonStatus = icBusinessUnitLoadFailure
'        LoadBusinessUnits
'        '***    load user license
'        m_enLogonStatus = icLicenseConfigFailure
'        LoadLicense
'        '***    load roles
'        m_enLogonStatus = icRoleLoadFailure
'        LoadRoles
'        '***    load configurations
'        m_enLogonStatus = icConfigurationLoadFailure
'        LoadConfigurations
'
'        '***    Get connection timeout
'        LoadConnectionTimeOut
'
'        '***    report success
'        m_enLogonStatus = icSuccess
'    Else
'        m_enLogonStatus = icFailure
'    End If
'
'EXIT_PROC:
'
'  Logon = m_enLogonStatus
'
'  Exit Function
'
'COMMAND_LINE_FAILURE:
'  Logon = m_enLogonStatus
'  Err.Raise Err.Number, Err.Source, Err.Description
'  Exit Function
'
'ERROR_HANDLER:
'
'  Logon = m_enLogonStatus
'  RaiseEvent LogonFailed
'  Err.Raise Err.Number, Err.Source, Err.Description
'  Exit Function
'
'End Function
Public Function Logon() As enLogonReturnCode
    '***********************************************************************************************
    'Purpose    : Logon the user
    'Created on :
    'Created by: Gordon
    '***********************************************************************************************

    '-------------------------------
    '*** Generic Procedure Code
    '-------------------------------
    On Error GoTo ERROR_HANDLER
    Dim strModuleAndProcName As String
    strModuleAndProcName = "CUser - Logon Function"
    
    Static idx As Integer
    
    Dim strSQL As String
    Dim rs As ADODB.Recordset
    Dim strErrorDescription As String
    Dim dbAccess As dbAccess
    Dim blnFound As Boolean
    Dim strPassword As String
    Dim strIPAddress As String
    Dim strComputerName As String
    Dim strTmp As String
    Dim colParameters As Collection

    '***    raise the event
    RaiseEvent LogonStart
    
    '***    fetch the initial parameters from the command line
    On Error GoTo COMMAND_LINE_FAILURE
    'only read connamd line when first log on
    If idx = 0 Then
'        ReadCommand
        idx = idx + 1
    End If
    
    On Error GoTo ERROR_HANDLER
    
    If IsTerminalSession Then
        strIPAddress = GetClientIPAddress
        strComputerName = GetClientComputerName
    Else
        strIPAddress = "0.0.0.0"
        strComputerName = GetLocalComputerName
    End If
    
    strSQL = "execute procedure sp_operator_login (p_login_code = ?, p_ip_address = ?, p_machine_name = ?"
    
    If m_blnNoPassword Then
        strSQL = strSQL & ", p_no_password = 'Y'"
    Else
        strSQL = strSQL & ", p_passwd = ?"
    End If
    
    If AutoLogon Then
        strSQL = strSQL & ", p_auto_logon = 'Y'"
    End If
    strSQL = strSQL & ", resume_flg = 'Y', tran_flg = 'Y')"
    
    '***    create the database access object
    m_enLogonStatus = icDatabaseFailure
    
    Set dbAccess = New dbAccess
    dbAccess.AccessMode = m_enmInitialAccessMode
    dbAccess.UserName = m_strLogonCode
    dbAccess.Password = m_strPassword
    dbAccess.DSN = m_strInitialDSN
    dbAccess.URL = m_strInitialURL
     
    Set colParameters = New Collection
    colParameters.Add m_strLogonCode
    colParameters.Add strIPAddress
    colParameters.Add strComputerName
    
    If Not m_blnNoPassword Then
        colParameters.Add m_strPassword
    End If
    
    dbAccess.Parameters = colParameters
    Set rs = dbAccess.LoadRS(strSQL)
    
    If rs("StatusCode") = 0 Then
        m_strContactCode = FieldToString(rs, "contact_code")
        m_strName = FieldToString(rs, "name")
        m_strFirstName = FieldToString(rs, "firstname")
        
        RaiseEvent LogonFinish
        
        'Get business unit code if missing in command line
        If m_strCurrentBusinessUnitCode = "" Then
            Set colParameters = New Collection
            
            dbAccess.Parameters = colParameters
            strSQL = "execute procedure ss_dflt_bus_unit (resume_flg = 'Y')"
            
            Set rs = dbAccess.LoadRS(strSQL)
            
            If rs(0) = 0 Then
                m_strCurrentBusinessUnitCode = rs(1)
                m_strDefaultBusinessUnitCode = m_strCurrentBusinessUnitCode
            Else
                m_strLoginErrorMessage = FieldToString(rs, rs.Fields.Count - 1)
                GoTo EXIT_PROC
            End If
        End If
        
        'dealer id could be processed before user
        If m_strDealerID <> "" Then
            If InvalidDealerID(m_strDealerID) Then
                m_strLoginErrorMessage = "Dealer ID supplied '" & m_strDealerID & "' is invalid. Please contact your system administrator."
                GoTo EXIT_PROC
            End If
        End If
        '***    load connection configuration
        m_enLogonStatus = icConnectionConfigLoadFailure
        LoadConnectionConfiguration
        '***    load IP
'        m_enLogonStatus = icInvalidIP
'        LoadIPAddress
        '***    load business units
        
        m_enLogonStatus = icBusinessUnitLoadFailure
        LoadBusinessUnits
        '***    load user license
        m_enLogonStatus = icLicenseConfigFailure
        LoadLicense
        '***    load roles
        m_enLogonStatus = icRoleLoadFailure
        LoadRoles
        '***    load configurations
        m_enLogonStatus = icConfigurationLoadFailure
        LoadConfigurations
        
        '***    Get connection timeout
        LoadConnectionTimeOut
        
        LoadUnixPassword
        
        '***    report success
        m_enLogonStatus = icSuccess
    Else
        If rs("StatusCode") = 100 Then
            m_enLogonStatus = icResetPassword
        ElseIf rs("StatusCode") = 400 Then
            m_enLogonStatus = icLockOut
        Else
            m_enLogonStatus = icFailure
        End If
        m_strLoginErrorMessage = FieldToString(rs, "ErrorMsg")
        RaiseEvent StatusBarText(m_strLoginErrorMessage)
        GoTo EXIT_PROC
    End If
    
    Set rs = Nothing
    Set dbAccess = Nothing

EXIT_PROC:
  
  Logon = m_enLogonStatus
  
  Exit Function
  
COMMAND_LINE_FAILURE:
  Logon = m_enLogonStatus
  Err.Raise Err.Number, Err.Source, Err.Description
  Exit Function
  
ERROR_HANDLER:
    
  Logon = m_enLogonStatus
  RaiseEvent LogonFailed
  Err.Raise Err.Number, Err.Source, Err.Description
  Exit Function
  
End Function

Public Sub DisplayLogonForm(Optional ByVal lngMasterObjectId As Long = 0, _
                            Optional sAppName As String = "")
    '***********************************************************************************************
    'Purpose    : Displays the logon form
    'Created on :
    'Created by: Gordon
    '***********************************************************************************************

    '-------------------------------
    '*** Generic Procedure Code
    '-------------------------------
    On Error GoTo ERROR_HANDLER
    Dim strModuleAndProcName As String
    strModuleAndProcName = "CUser - DisplayLogonForm Sub"
    
LOGON_AGAIN:
    
    Dim frm As New frmLogon
    Dim sMsg As String
    
    'read command here, was in logon
    ReadCommand
    
    '***    set the master object id
    m_lngMasterObjectId = lngMasterObjectId
       
    Set frm.m_objUser = Me
    Load frm
    
    '***    set the user object
'    m_strAppName = sAppName
    
'    '***    translate the form
'    TranslateForm frm
    Dim sLogon As String
    
    '***    fetch the last successful logon
    'Logon Code might in commandline
    If m_strLogonCode = "" Then
        sLogon = GetSetting("Selcomm", "LogonCode", "LogonCode", "")
        If sLogon = "" Then
            sLogon = Environ("USERNAME")
        End If
    Else
        sLogon = m_strLogonCode
    End If
    
    If sLogon = "" Then
        sLogon = GetSetting("Selcomm", "LogonCode", "LogonCode", "")
    End If
    
    frm.txtUserName.Text = sLogon
    
    If m_blnNoPassword Or m_strPassword <> "" Then
        frm.txtPassword = m_strPassword
        frm.Timer1.Enabled = True
    End If
    
    '***    open the form
    frm.Show vbModal
    
    '***    record the logon code
    SaveSetting "Selcomm", "LogonCode", "LogonCode", Me.LogonCode
    
    '***    check the licence key, if the user is in the INSTALLER role
    '***    user then write the disk drive serial number to the file
    '***    otherwise check the disk drive serial number
'    If frm.m_bLogonStatus = True Then
'        If Me.CurrentRole = "INSTALLER" Then
'            If Not WriteLicenceKeyFile(m_strAppName) Then
'                MsgBox "Failed to generate licence key file for this application. Please contact Select Software Solutions.", vbApplicationModal, "Licence Key File"
'                Me.LogonStatus = icFailure
'                Exit Sub
'            Else
'                MsgBox "Licence Key File generated. You must now logon under your normal logon code.", vbApplicationModal, "Licence Key Generated"
'                m_lngRoleId = 0
'                Me.LogonStatus = icFailure
'                GoTo LOGON_AGAIN
'            End If
'        Else
'            If Not CheckLicenceKey(m_strAppName) Then
'                MsgBox "Invalid licence key file for this application. Please contact Select Software Solutions.", vbApplicationModal, "Invalid Licence Key"
'                Me.LogonStatus = icFailure
'                Exit Sub
'            End If
'        End If
'    End If
    '***    null the form
    Unload frm
    
EXIT_PROC:

    Exit Sub
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

End Sub
Public Sub DisplayChangePasswordForm()
    '***********************************************************************************************
    'Purpose    : Displays the password change form
    'Created on :
    'Created by: Gordon
    '***********************************************************************************************

    '-------------------------------
    '*** Generic Procedure Code
    '-------------------------------
    On Error GoTo ERROR_HANDLER
    Dim strModuleAndProcName As String
    strModuleAndProcName = "CUser - DisplayChangePasswordForm Sub"
    
    Dim frm As New frmChangePassword
    
    '***    set the user object
    Set frm.m_objUser = Me
    
    '***    open the form
    frm.Show vbModal
    
    '***    null the form
    Unload frm
    
EXIT_PROC:

    Exit Sub
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

End Sub
Public Sub DisplayChangeRoleForm()
    '***********************************************************************************************
    'Purpose    : Displays the role change form
    'Created on :
    'Created by: Gordon
    '***********************************************************************************************

    '-------------------------------
    '*** Generic Procedure Code
    '-------------------------------
    On Error GoTo ERROR_HANDLER
    Dim strModuleAndProcName As String
    strModuleAndProcName = "CUser - DisplayChangeRoleForm Sub"
    
    Dim frm As New frmChangeRole
    
    '***    set the user object
    Set frm.m_objUser = Me
    
    '***    populate the form
    frm.Populate
    
    '***    open the form
    frm.Show vbModal
    
    '***    null the form
    Unload frm
    
EXIT_PROC:

    Exit Sub
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

End Sub
Public Sub DisplayChangeBusinessUnitForm()
    '***********************************************************************************************
    'Purpose    : Displays the business unit change form
    'Created on :
    'Created by: Gordon
    '***********************************************************************************************

    '-------------------------------
    '*** Generic Procedure Code
    '-------------------------------
    On Error GoTo ERROR_HANDLER
    Dim strModuleAndProcName As String
    strModuleAndProcName = "CUser - DisplayChangeBusinessUnitForm Sub"
    
    Dim frm As New frmChangeBusinessUnit
    
    '***    set the user object
    Set frm.m_objUser = Me
    
    '***    populate the form
    frm.Populate
    
    '***    open the form
    frm.Show vbModal
    
    '***    null the form
    Unload frm
    
EXIT_PROC:

    Exit Sub
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

End Sub
Private Sub LoadConnectionConfiguration()
    '***********************************************************************************************
    'Purpose    : Gets the connection configuration object.
    'Created on :
    'Created by: Gordon
    '***********************************************************************************************

    '-------------------------------
    '*** Generic Procedure Code
    '-------------------------------
    On Error GoTo ERROR_HANDLER
    Dim strModuleAndProcName As String
    strModuleAndProcName = "CUser - LoadConnectionConfiguration Sub"
    
    '***    raise the event
    RaiseEvent LoadConfigurationStart
    
    '***    create the object.
    Set m_ConnectionConfiguration = New ConnectionConfiguration

    '***    set the required values
    m_ConnectionConfiguration.User = Me
    
    '***    if we are operating in command line mode then directly set the
    '***    connection configuration object. In this case there will be
    '***    only one connection configured.
    If m_strConnectMode = "C" Then
        m_ConnectionConfiguration.SetProperties m_strCurrentBusinessUnitCode, m_strInitialDSN, _
                            m_strDatabaseName, m_enmInitialAccessMode, _
                            m_strInitialURL
        
    Else
        '***    load up all the connections as usual
        m_ConnectionConfiguration.Populate
    End If
    
EXIT_PROC:

    '***    raise the event
    RaiseEvent LoadConfigurationFinish
    
    Exit Sub
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

End Sub
Private Sub LoadConfigurations()
    '***********************************************************************************************
    'Purpose    : Gets the configuration object.
    'Created on :
    'Created by: Gordon
    '***********************************************************************************************

    '-------------------------------
    '*** Generic Procedure Code
    '-------------------------------
    On Error GoTo ERROR_HANDLER
    Dim strModuleAndProcName As String
    strModuleAndProcName = "CUser - LoadConfigurations Sub"
    
    '***    raise the event
    RaiseEvent LoadConfigurationStart
    
    '***    create the object.
    Set m_Configurations = New Configurations

    '***    set the required values
    m_Configurations.User = Me
    m_Configurations.LoadComponents
    
EXIT_PROC:

    '***    raise the event
    RaiseEvent LoadConfigurationFinish
    
    Exit Sub
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

End Sub
Public Sub LoadRoles()
    '***********************************************************************************************
    'Purpose    : Load user roles
    'Created on :
    'Created by: Gordon
    '***********************************************************************************************

    '-------------------------------
    '*** Generic Procedure Code
    '-------------------------------
    On Error GoTo ERROR_HANDLER
    Dim strModuleAndProcName As String
    strModuleAndProcName = "CUser - LoadRoles Sub"
    
    Dim dbAccess As dbAccess
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    Dim intCounter As Integer
    Dim sRole As String
    Dim sDefaultRole As String
    Dim bDefaultRoleCurrent As Boolean
'    Dim bLicenseFileExit As Boolean
    Dim bInstallerRole As Boolean
    Dim RoleList As sssRoleList
    Dim bRoleExists As Boolean
    Dim cOldRoleList As Collection
    Dim lOldRoleID As Long
    
    '***    raise the event
    RaiseEvent LoadRolesStart
    '***    the SQL
    strSQL = "select usr_roles.role_id , role_narr " & _
            "from usr_roles, roles " & _
            "where usr_roles.role_id = roles.role_id " & _
            "and login_code = '" & m_strLogonCode & "' " & _
            "and to_dt >= current " & _
            "and from_dt <= current " & _
            "order by usr_roles.role_id "
            
                        '"and from_dt <= today and to_dt >= today " & _
    '***    fetch the default role
    '***    use the current role when switch database
    If CurrentRole = "" Or CurrentRole = "INSTALLER" Then
        sDefaultRole = GetSetting("Selcomm", "Role", "Role", "")
    Else
        sDefaultRole = CurrentRole
    End If
    bDefaultRoleCurrent = False
    
    '***    get the dbAccess object
    Set dbAccess = GetDbAccessObject(m_strDefaultBusinessUnitCode)
    Set rs = dbAccess.LoadRS(strSQL)
        
'    bLicenseFileExit = IsLicenceKeyFilePresent(m_strAppName)
    
    Set cOldRoleList = m_RoleList
    lOldRoleID = m_lngRoleId
    
    Set m_rsRoles = rs
    Set m_RoleList = New Collection
    
    While Not m_rsRoles.EOF
        '***    if no license file exists, check if has 'installer' role
'        If bLicenseFileExit Then
            '***    see if the default role is still current
            If sDefaultRole = FieldToString(m_rsRoles, "role_narr") Then
                bDefaultRoleCurrent = True
            End If
            intCounter = intCounter + 1
            If intCounter = 1 Then
                m_lngRoleId = CInt(rs("role_id"))
                '***    is the user an administrator
                '***    role_id = 1 is ADMINISTRATOR
                If rs(0) = 1 Then m_blnAdminRole = True
                sRole = FieldToString(m_rsRoles, "role_narr")
            End If
            
            'check if role exists in new database
            If m_blnSwitchDatabase Then
                If sDefaultRole = FieldToString(m_rsRoles, "role_narr") Then
                    bRoleExists = True
                End If
            End If
'        Else
'            '*** 'installer' role
'            If rs(0) = 100 Then
'                bInstallerRole = True
'                intCounter = intCounter + 1
'                sRole = FieldToString(m_rsRoles, "role_narr")
'            End If
'        End If
        
        Set RoleList = New sssRoleList
        RoleList.RoleID = CInt(m_rsRoles("role_id"))
        RoleList.RoleNarr = FieldToString(m_rsRoles, "role_narr")
        m_RoleList.Add RoleList
        
        m_rsRoles.MoveNext
    Wend
    m_intNumberOfRoles = intCounter
        
    'when switch database, error if role does not exist in the new database
    If m_blnSwitchDatabase Then
        If Not bRoleExists Then
            GoTo ROLE_NOT_EXISTS
        End If
    End If
        
    '***    clear the default role
    If Not bDefaultRoleCurrent Then
        sDefaultRole = ""
    End If
    
    If intCounter = 0 Then
'        If bLicenseFileExit Then
            GoTo NO_ROLES
'        Else
'            GoTo NO_LICENSEFILE_NO_INSTALLERROLE
'        End If
    End If
    
'    m_rsRoles.MoveFirst
    Set rs = Nothing
    Set dbAccess = Nothing
    
EXIT_PROC:

    '***    raise the event
    RaiseEvent LoadRolesFinish
    
    If Trim(sDefaultRole) <> "" Then
        SetRole sDefaultRole
    ElseIf m_intNumberOfRoles > 1 Then
        RaiseEvent RoleSelectionRequired
    Else
        SetRole sRole
    End If
    
    Exit Sub
    
NO_ROLES:
    Err.Raise 999, "LoadRoles", "No Roles defined"
    Exit Sub
'NO_LICENSEFILE_NO_INSTALLERROLE:
'    Err.Raise 1001, "LoadRoles", "You must log on as an installer to complete this installation of Selcomm."
'    Exit Sub
ROLE_NOT_EXISTS:
    Set m_RoleList = cOldRoleList
    m_lngRoleId = lOldRoleID
    Err.Raise 1005, "LoadRoles", sDefaultRole & " does not exist in the new database."
    '------------------------------------------------------------
    '*** Generic Error Handling Code (ensure Exit Proc is above)
    '------------------------------------------------------------
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
Public Sub LoadLicense()
    '***********************************************************************************************
    'Purpose    : Load user lincese
    'Created on :
    'Created by: Patrick
    '***********************************************************************************************

    '-------------------------------
    '*** Generic Procedure Code
    '-------------------------------
    On Error GoTo ERROR_HANDLER
    Dim strModuleAndProcName As String
    strModuleAndProcName = "CUser - LoadLicense Sub"
    
    Dim dbAccess As dbAccess
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    
    '***    raise the event
    RaiseEvent LoadLicenseStart
    '***    the SQL
    strSQL = "execute procedure ss_license()"
                
    '***    get the dbAccess object
    Set dbAccess = GetDbAccessObject(m_strDefaultBusinessUnitCode)
    Set rs = dbAccess.LoadRS(strSQL)
    
    If Not rs.EOF Then
        If rs(0) = 0 Then
            m_blnLicense = (Trim$(rs(3)) = "Y")
            
            If Not m_blnLicense Then
                m_enLogonStatus = icLicenseFailure
                GoTo LICENSE_LIMIT
            End If
        Else
            GoTo NO_LICENSE
        End If
    End If
    
    Set rs = Nothing
    Set dbAccess = Nothing
    
EXIT_PROC:

    '***    raise the event
    RaiseEvent LoadLicenseFinsish
    
    Exit Sub
    
LICENSE_LIMIT:
    Err.Raise 1004, "LoadLicense", "No more license available."
    Exit Sub
NO_LICENSE:
    Err.Raise 1003, "LoadLicense", "No license defined."
    '------------------------------------------------------------
    '*** Generic Error Handling Code (ensure Exit Proc is above)
    '------------------------------------------------------------
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

Private Sub LoadParentRoles(Optional iLevel As Integer = 1, Optional lRoleId As Long = -1)
    '***********************************************************************************************
    'Purpose    : Loads the parent roles
    'Created on :
    'Created by: Gordon
    '***********************************************************************************************

    '-------------------------------
    '*** Generic Procedure Code
    '-------------------------------
    On Error GoTo ERROR_HANDLER
    Dim strModuleAndProcName As String
    strModuleAndProcName = "CUser - LoadParentRoles Sub"
    
    Dim dbAccess As dbAccess
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    
    Dim lParentRoleId As Long
    Dim idx As Integer
    
    '***    clear the roles object if this is the first level
    If iLevel = 1 Then
        '***    raise the event
        RaiseEvent LoadParentRoleStart
    
        Set m_Roles = New sssRoles
        '***    add the first role (the current role) obviously level 1
        m_Roles.AddRole iLevel, m_lngRoleId
        '***    set the role id to the current role
        lRoleId = m_lngRoleId
    End If
    
    '***    fetch the parent roles for this role
    strSQL = "select par_role_id  " & _
            "from roles_rel  " & _
            "where child_role_id = ? "

    '***    get the dbAccess object
    Set dbAccess = GetDbAccessObject(m_strDefaultBusinessUnitCode)
    Set rs = dbAccess.LoadRS(strSQL, lRoleId)

    '***    we cycle the parent roles. First we check if the parent role
    '***    is already included in our role collection. If it is we discard
    '***    it, this is to avoid circular references. Otherwise we add it to
    '***    the roles collection then recursively call this procedure
    While Not rs.EOF
        lParentRoleId = rs(0)
        '***    see if we already have this role
        If m_Roles.RoleExists(lParentRoleId) Then GoTo SKIP_RECORD
        
        '***    add this role to the collection
        idx = iLevel + 1
        m_Roles.AddRole idx, lParentRoleId
        '***    recursively call this procedure for the next level down for the parent role
        LoadParentRoles idx, lParentRoleId
        
SKIP_RECORD:

        rs.MoveNext
    Wend

EXIT_PROC:

    Set rs = Nothing
    Set dbAccess = Nothing

    '***    raise the event
    If iLevel = 1 Then RaiseEvent LoadParentRoleStart
    
    Exit Sub
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

End Sub
Public Sub LoadBusinessUnits()
    '***********************************************************************************************
    'Purpose    : Load business units for this user
    'Created on :
    'Created by: Gordon
    '***********************************************************************************************

    '-------------------------------
    '*** Generic Procedure Code
    '-------------------------------
    On Error GoTo ERROR_HANDLER
    Dim strModuleAndProcName As String
    strModuleAndProcName = "CUser - LoadBusinessUnits Sub"
    
    Dim dbAccess As dbAccess
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    Dim intCounter As Integer
    
    Dim sBusinessUnitCode As String
    

    '***    some validations
    If m_strLogonCode = "" Then Err.Raise vbObjectError + 100, strModuleAndProcName, "No logon code set."
    
'    '***    the SQL
'    strSQL = "select bus_usr.bus_unit_code , bus_unit_narr " & _
'            "from bus_usr, bus_unit " & _
'            "where bus_usr.bus_unit_code = bus_unit.bus_unit_code " & _
'            "and login_code = ? " & _
'            "order by 1 "
'
'    '***    get the dbAccess object
    Set dbAccess = GetDbAccessObject(m_strDefaultBusinessUnitCode)
'    Set rs = dbAccess.LoadRS(strSQL, m_strLogonCode)
'
'    intCounter = 0
'    While Not rs.EOF
'        intCounter = intCounter + 1
'        '***    add the business unit to the dictionary
'        m_dicBusinessUnits.Add FieldToString(rs, "bus_unit_code"), FieldToString(rs, "bus_unit_narr")
'        rs.MoveNext
'    Wend
'
'    If intCounter > 0 Then GoTo EXIT_PROC
    
    '***    the SQL, select all business units
    strSQL = "select bus_unit_code , bus_unit_narr " & _
            "from bus_unit " & _
            "where bus_unit_display = 'Y'"
        
    Set rs = dbAccess.LoadRS(strSQL)
    
    intCounter = 0
    While Not rs.EOF
        intCounter = intCounter + 1
        sBusinessUnitCode = FieldToString(rs, "bus_unit_code")
        
        '***    add the business unit to the dictionary
        If Not m_dicBusinessUnits.Exists(sBusinessUnitCode) Then
            m_dicBusinessUnits.Add sBusinessUnitCode, FieldToString(rs, "bus_unit_narr")
        End If
        rs.MoveNext
    Wend
    
    If intCounter = 0 Then Err.Raise 998, strModuleAndProcName, "No business units defined."
    
EXIT_PROC:
    Set rs = Nothing
    Set dbAccess = Nothing
        
    Exit Sub
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

End Sub

Public Sub LoadIPAddress()
    '***********************************************************************************************
    'Purpose    : Load IP address range for access
    'Created on :
    'Created by: Gordon
    '***********************************************************************************************

    '-------------------------------
    '*** Generic Procedure Code
    '-------------------------------
    On Error GoTo ERROR_HANDLER
    Dim strModuleAndProcName As String
    strModuleAndProcName = "CUser - LoadIPAddress Sub"
    
    Dim idx As Integer
    Dim blnValidIP As Boolean
    Dim IPAddress() As String
    Dim IPRange As String
    Dim ClientIPAddrsss As String
    Dim dbAccess As dbAccess
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    
    '***    raise the event
    RaiseEvent LoadIPaddressStart
    
    'if not terminal server, all valid
    If Not IsThisTerminalServer Then
        blnValidIP = True
    Else
        'if console, all valid
        If Not IsTerminalSession Then
            blnValidIP = True
        Else
            '***    the SQL
            strSQL = "select cnfg_val from sys_cnfg where cnfg_key = 'VALID_LOGIN_IP'"
            
        '    '***    get the dbAccess object
            Set dbAccess = GetDbAccessObject(m_strDefaultBusinessUnitCode)
            Set rs = dbAccess.LoadRS(strSQL)
            
            If Not rs.EOF Then
                IPRange = FieldToString(rs, 0)
                
                If IPRange = "" Or InStr(UCase(IPRange), "ALL") > 0 Then
                    blnValidIP = True
                Else
                    While InStr(IPRange, "  ") > 0
                        IPRange = Replace(IPRange, "  ", " ")
                    Wend
                    
                    IPAddress = Split(IPRange, " ")
                        
                    'running on console, not IP address returned
                    ClientIPAddrsss = GetClientIPAddress
                    'hard code for testing only
        '            ClientIPAddrsss = "10.1.14.114"
                    If ClientIPAddrsss <> "" Then
                        
                        For idx = 0 To UBound(IPAddress)
                            If IPAddressInRange(ClientIPAddrsss, IPAddress(idx)) Then
                                blnValidIP = True
                                Exit For
                            End If
                        Next
                    Else
                        blnValidIP = True
                    End If
                End If
            Else
                blnValidIP = True
            End If
            
            Set rs = Nothing
            Set dbAccess = Nothing
        End If
    End If
    
    RaiseEvent LoadIPAddressFinish
        
    If Not blnValidIP Then Err.Raise 1006, strModuleAndProcName, "You are not permitted to logon from address " & ClientIPAddrsss & " ."
    
EXIT_PROC:

        
    Exit Sub
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

End Sub

Public Sub LoadDataSource(Optional blnRefresh As Boolean = False)
    '***********************************************************************************************
    'Purpose    : Load Data source list
    'Created on :
    'Created by: Patrick
    '***********************************************************************************************

    '-------------------------------
    '*** Generic Procedure Code
    '-------------------------------
    On Error GoTo ERROR_HANDLER
    Dim strModuleAndProcName As String
    strModuleAndProcName = "CUser - LoadDataSource Sub"
    
    Dim dbAccess As dbAccess
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    Dim typDatasource As typDatasourceNew2
    
    If blnRefresh Then Set m_colDataSource = New Collection
    
'    '***    the SQL
    strSQL = "execute procedure ss_datasrce(p_login_code = ?, resume_flg = 'Y');"

'    '***    get the dbAccess object
    Set dbAccess = GetDbAccessObject(m_strDefaultBusinessUnitCode)
    Set rs = dbAccess.LoadRS(strSQL, m_strLogonCode)
    
    While Not rs.EOF
        With typDatasource
            .BusinessUnitCode = FieldToString(rs, "bus_unit_code")
            .BusinessUnitNarr = FieldToString(rs, "bus_unit_narr")
            .DataSourceCode = FieldToString(rs, "datasourcecode")
            .DataSourceLocal = (FieldToString(rs, "datasrcestat") = "L")
            .DataSourceName = FieldToString(rs, "datasrcename")
            .DataSourceNarr = FieldToString(rs, "datasrcenarr")
            .BusUnitColour = FieldToString(rs, "bus_unit_colour")
            m_colDataSource.Add typDatasource
        End With
        rs.MoveNext
    Wend
        
EXIT_PROC:

    Set rs = Nothing
    Set dbAccess = Nothing
        
    Exit Sub
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

End Sub

Private Sub LoadEntityConfiguration(ByVal lngEntityId As Long)
    '***********************************************************************************************
    'Purpose    : Load entity configuration data
    'Created on :
    'Created by: Gordon
    '***********************************************************************************************

    '-------------------------------
    '*** Generic Procedure Code
    '-------------------------------
    On Error GoTo ERROR_HANDLER
    Dim strModuleAndProcName As String
    strModuleAndProcName = "CUser - LoadEntityConfiguration Sub"
    
    Dim objEntity As sssEntityAttributes
    Dim strBusinessUnitCode As String
    
    '***    raise the event
    RaiseEvent LoadEntityDataStart
    
    '***    determine which business unit to load the access information from
    If m_blnGlobalEntity Then
        strBusinessUnitCode = m_strDefaultBusinessUnitCode
    Else
        strBusinessUnitCode = m_strCurrentBusinessUnitCode
    End If
    
    '***    see if we already have this menu loaded
    If EntityExists(strBusinessUnitCode, lngEntityId) Then
        GoTo EXIT_PROC
    End If
    
    '***    create the menu data
    Set objEntity = New sssEntityAttributes
    
    '***    initialise required information
    objEntity.Id = lngEntityId
    objEntity.User = Me
    
    '***    call the load
    objEntity.Load lngEntityId
    
    '***    add the menu item to the collection
    AddEntity strBusinessUnitCode, objEntity, lngEntityId
        
EXIT_PROC:
    
    '***    raise the event
    RaiseEvent LoadEntityDataFinish
    
    Exit Sub
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

End Sub
Private Sub LoadMenuData(ByVal lngMenuId As Long)
    '***********************************************************************************************
    'Purpose    : Load menu
    'Created on :
    'Created by: Gordon
    '***********************************************************************************************

    '-------------------------------
    '*** Generic Procedure Code
    '-------------------------------
    On Error GoTo ERROR_HANDLER
    Dim strModuleAndProcName As String
    strModuleAndProcName = "CUser - LoadMenuData Sub"
    
    Dim objMenu As sssMenu
    Dim strBusinessUnitCode As String
    
    '***    raise the event
    RaiseEvent LoadMenuDataStart
    
    '***    determine which business unit to load the access information from
    If m_blnGlobalMenues Then
        strBusinessUnitCode = m_strDefaultBusinessUnitCode
    Else
        strBusinessUnitCode = m_strCurrentBusinessUnitCode
    End If
    
    '***    see if we already have this menu loaded
    If MenuExists(strBusinessUnitCode, lngMenuId) Then
        GoTo EXIT_PROC
    End If
    
    '***    create the menu data
    Set objMenu = New sssMenu
    
    '***    initialise required information
    objMenu.MenuId = lngMenuId
    objMenu.User = Me
    objMenu.dbAccess = Me.GetDbAccessObject(strBusinessUnitCode)
    
    '***    check the access permissions
    If GetAccessRights(lngMenuId) = FullAccess Then
        '***    call the load
        objMenu.LoadMenuItems2 lngMenuId
    End If
    
    '***    add the menu item to the collection
    AddMenu strBusinessUnitCode, objMenu, lngMenuId
        
EXIT_PROC:
    
    '***    raise the event
    RaiseEvent LoadMenuDataFinish
    
    Exit Sub
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

End Sub
Private Sub LoadMenuData2(ByVal lngMenuId As Long, Optional strContactCode As String, _
    Optional lngSpCnRef As Long, Optional strServiceTypeCode As String, _
    Optional strServicestatusCode As String)
    '***********************************************************************************************
    'Purpose    : Load menu
    'Created on :
    'Created by: Gordon
    '***********************************************************************************************

    '-------------------------------
    '*** Generic Procedure Code
    '-------------------------------
    On Error GoTo ERROR_HANDLER
    Dim strModuleAndProcName As String
    strModuleAndProcName = "CUser - LoadMenuData Sub"
    
    Dim objMenu As sssMenu
    Dim strBusinessUnitCode As String
    
    '***    raise the event
    RaiseEvent LoadMenuDataStart
    
    '***    determine which business unit to load the access information from
    If m_blnGlobalMenues Then
        strBusinessUnitCode = m_strDefaultBusinessUnitCode
    Else
        strBusinessUnitCode = m_strCurrentBusinessUnitCode
    End If
    
    '***    see if we already have this menu loaded
    If MenuExists(strBusinessUnitCode, lngMenuId) Then
        GoTo EXIT_PROC
    End If
    
    '***    create the menu data
    Set objMenu = New sssMenu
    
    '***    initialise required information
    objMenu.MenuId = lngMenuId
    objMenu.User = Me
    objMenu.dbAccess = Me.GetDbAccessObject(strBusinessUnitCode)
    
    '***    check the access permissions
    If GetAccessRights(lngMenuId) = FullAccess Then
        '***    call the load
        objMenu.LoadMenuItems2 lngMenuId, strContactCode, lngSpCnRef, strServiceTypeCode, strServicestatusCode
    End If
    
    '***    add the menu item to the collection
    AddMenu strBusinessUnitCode, objMenu, lngMenuId
        
EXIT_PROC:
    
    '***    raise the event
    RaiseEvent LoadMenuDataFinish
    
    Exit Sub
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

End Sub

Public Sub LoadConnectionTimeOut()
    '***********************************************************************************************
    'Purpose    : Load database connection timeout
    'Created on :
    'Created by: Gordon
    '***********************************************************************************************

    '-------------------------------
    '*** Generic Procedure Code
    '-------------------------------
    On Error Resume Next
    Dim strModuleAndProcName As String
    strModuleAndProcName = "CUser - LoadConnectionTimeOut Sub"
    
    Dim dbAccess As dbAccess
    Dim rs As ADODB.Recordset
    Dim strSQL As String

    '***    get the dbAccess object
    Set dbAccess = GetDbAccessObject(m_strDefaultBusinessUnitCode)
    
    '***    the SQL, select all business units
    strSQL = "select cnfg_val from sys_cnfg where cnfg_key = '" & ConnectionTimeOutKey & "'"
            
    Set rs = dbAccess.LoadRS(strSQL)
    
    If Not rs.EOF Then
        If Val(rs(0)) >= 1 And Val(rs(0)) <= 9999 Then
            ConnectionTimeout = Val(rs(0))
        End If
    End If

    Set rs = Nothing
    Set dbAccess = Nothing
        
End Sub

Public Sub LoadUnixPassword()
    '***********************************************************************************************
    'Purpose    : Load Unit passwrod from registry and decrypt
    'Created on :
    'Created by: Gordon
    '***********************************************************************************************

    '-------------------------------
    '*** Generic Procedure Code
    '-------------------------------
    On Error Resume Next
    Dim strModuleAndProcName As String
    strModuleAndProcName = "CUser - LoadUnixPassword Sub"
    
    Dim UnixPassword As String

    UnixPassword = QueryValue(HKEY_CURRENT_USER, "software\ODBC\ODBC.INI\" & InitialDSN, "PWD")
    UnixPassword = DecryptInformix(LogonCode, UnixPassword)
    
    m_strUnixPassword = UnixPassword
        
End Sub


Public Function GetEntity(ByVal lngEntityId As Long) As sssEntityAttributes
    '***********************************************************************************************
    'Purpose    : Return the Entity
    'Created on :
    'Created by: Gordon
    '***********************************************************************************************

    '-------------------------------
    '*** Generic Procedure Code
    '-------------------------------
    On Error GoTo ERROR_HANDLER
    Dim strModuleAndProcName As String
    strModuleAndProcName = "CUser - GetEntity Function"
    
    Dim dic As Dictionary
    Dim strBusinessUnitCode As String
    
    '***    Load entity data
    LoadEntityConfiguration lngEntityId
       
    '***    determine which business unit to load the access information from
    If m_blnGlobalEntity Then
        strBusinessUnitCode = m_strDefaultBusinessUnitCode
    Else
        strBusinessUnitCode = m_strCurrentBusinessUnitCode
    End If
    
    '***    does the dictionary have a collection of entities for this business unit?
    If Not EntityExists(strBusinessUnitCode, lngEntityId) Then
'        Err.Raise vbObjectError + 100, "CUser.GetEntity", "Invalid Entity Id."
        Set GetEntity = Nothing
    Else
        Set dic = m_dicEntities.item(strBusinessUnitCode)
    End If
    If Not dic.Exists(CStr(lngEntityId)) Then
'        Err.Raise vbObjectError + 100, "CUser.GetEntity", "Invalid Entity Id."
        Set GetEntity = Nothing
    Else
        Set GetEntity = dic.item(CStr(lngEntityId))
    End If
    
    
EXIT_PROC:
    
    
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

Public Function GetMenu(ByVal lngMenuId As Long) As sssMenu
    '***********************************************************************************************
    'Purpose    : Return the menu
    'Created on :
    'Created by: Gordon
    '***********************************************************************************************

    '-------------------------------
    '*** Generic Procedure Code
    '-------------------------------
    On Error GoTo ERROR_HANDLER
    Dim strModuleAndProcName As String
    strModuleAndProcName = "CUser - GetMenu Function"
    
    Dim dic As Dictionary
    Dim strBusinessUnitCode As String
    
    If lngMenuId <= 0 Then Exit Function
    
    
    '***    Load menu data
    LoadMenuData lngMenuId
       
    '***    determine which business unit to load the access information from
    If m_blnGlobalMenues Then
        strBusinessUnitCode = m_strDefaultBusinessUnitCode
    Else
        strBusinessUnitCode = m_strCurrentBusinessUnitCode
    End If
    
    '***    does the dictionary have a collection of menues for this business unit?
    If Not MenuExists(strBusinessUnitCode, lngMenuId) Then
        Err.Raise vbObjectError + 100, "CUser.GetMenu", "Invalid Menu Id."
    Else
        Set dic = m_dicMenues.item(strBusinessUnitCode)
    End If
    
    If Not dic.Exists(CStr(lngMenuId)) Then
        Err.Raise vbObjectError + 100, "CUser.GetMenu", "Invalid Menu Id."
    Else
        Set GetMenu = dic.item(CStr(lngMenuId))
    End If
    
    
EXIT_PROC:
    
    
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
Public Function GetMenu2(ByVal lngMenuId As Long, Optional strContactCode As String, _
    Optional lngSpCnRef As Long, Optional strServiceTypeCode As String, _
    Optional strServicestatusCode As String) As sssMenu
    '***********************************************************************************************
    'Purpose    : Return the menu
    'Created on :
    'Created by: Gordon
    '***********************************************************************************************

    '-------------------------------
    '*** Generic Procedure Code
    '-------------------------------
    On Error GoTo ERROR_HANDLER
    Dim strModuleAndProcName As String
    strModuleAndProcName = "CUser - GetMenu Function"
    
    Dim dic As Dictionary
    Dim strBusinessUnitCode As String
    
    If lngMenuId <= 0 Then Exit Function
    
    'refresh menu
    RemoveMenu lngMenuId
    
    '***    Load menu data
    LoadMenuData2 lngMenuId, strContactCode, lngSpCnRef, strServiceTypeCode, strServicestatusCode
       
    '***    determine which business unit to load the access information from
    If m_blnGlobalMenues Then
        strBusinessUnitCode = m_strDefaultBusinessUnitCode
    Else
        strBusinessUnitCode = m_strCurrentBusinessUnitCode
    End If
    
    '***    does the dictionary have a collection of menues for this business unit?
    If Not MenuExists(strBusinessUnitCode, lngMenuId) Then
        Err.Raise vbObjectError + 100, "CUser.GetMenu", "Invalid Menu Id."
    Else
        Set dic = m_dicMenues.item(strBusinessUnitCode)
    End If
    
    If Not dic.Exists(CStr(lngMenuId)) Then
        Err.Raise vbObjectError + 100, "CUser.GetMenu", "Invalid Menu Id."
    Else
        Set GetMenu2 = dic.item(CStr(lngMenuId))
    End If
    
EXIT_PROC:
    
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


Public Function GetDbAccessObject(Optional ByVal strBusinessUnitCode As String, _
                                Optional ByVal CreateNew As Boolean = False) _
    As dbAccess
    '***********************************************************************************************
    'Purpose    : Returns a dbAccess object
    'Created on :
    'Created by: Gordon
    '***********************************************************************************************

    '-------------------------------
    '*** Generic Procedure Code
    '-------------------------------
    On Error GoTo ERROR_HANDLER
    Dim strModuleAndProcName As String
    strModuleAndProcName = "CUser - CreateDbAccessObject Sub"
    
    Dim dbAccess As dbAccess
    Dim blnFound As Boolean
    Dim intCounter As Integer
    
    '***    if strBusinessUnitCode is not passed then use the default from the
    '***    connection configuration object
    If IsMissing(strBusinessUnitCode) Or strBusinessUnitCode = "" Then
        '***    use the default business unit if none requested
        strBusinessUnitCode = m_strDefaultBusinessUnitCode
    End If
    
    '***    if a new object has not been requested try and find an existing one
    If Not CreateNew Then
        For intCounter = 1 To m_dicDbAccessObjects.Count
            Set dbAccess = m_dicDbAccessObjects(CStr(intCounter))
            If dbAccess.BusinessUnitCode = strBusinessUnitCode Then
                blnFound = True
                Exit For
            End If
        Next
    End If
    
    '***    create the dbAccess object if required
    If Not blnFound Then
        Set dbAccess = New dbAccess
        dbAccess.UserName = m_strLogonCode
        dbAccess.Password = m_strPassword
        dbAccess.AccessMode = Me.ConnectionConfiguration.AccessMethod(strBusinessUnitCode)
        dbAccess.DatabaseName = Me.ConnectionConfiguration.Database(strBusinessUnitCode)
        dbAccess.DSN = Me.ConnectionConfiguration.DSN(strBusinessUnitCode)
        dbAccess.URL = Me.ConnectionConfiguration.URL(strBusinessUnitCode)
        dbAccess.BusinessUnitCode = strBusinessUnitCode
        
        If Me.ConnectionConfiguration.BusinessUnitCodes(strBusinessUnitCode) = "" Then
            Err.Raise -105, strModuleAndProcName, "Invalid business unit code."
        Else
            dbAccess.BusinessUnitCode = Me.ConnectionConfiguration.BusinessUnitCodes(strBusinessUnitCode)
        End If
        '***    set the Id
        dbAccess.Id = CStr(m_dicDbAccessObjects.Count + 1)
        m_dicDbAccessObjects.Add dbAccess.Id, dbAccess
    End If
    
    Set GetDbAccessObject = dbAccess
    
EXIT_PROC:
    
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
Public Sub SetBusinessUnit(ByVal strBusinessUnitCode As String)

    '***********************************************************************************************
    'Purpose    : Sets the business unit
    'Created on :
    'Created by: Gordon
    '***********************************************************************************************

    '-------------------------------
    '*** Generic Procedure Code
    '-------------------------------
    On Error GoTo ERROR_HANDLER
    Dim strModuleAndProcName As String
    strModuleAndProcName = "CUser - SetBusinessUnit Sub"
    Dim lngRoleId As Long
    
    '***    validate the business unit
    If Not m_dicBusinessUnits.Exists(strBusinessUnitCode) Then _
        Err.Raise vbObjectError + 100, strModuleAndProcName, "Invalid business unit."
    
    '***    do we need to do anything?
    If strBusinessUnitCode = m_strCurrentBusinessUnitCode Then GoTo EXIT_PROC

    '***    role is different
    m_strCurrentBusinessUnitCode = strBusinessUnitCode

'    Set m_dicDbAccessObjects = New Dictionary
'    Set m_colMenues = New Collection
    Set m_dicArguments = New Dictionary
  
EXIT_PROC:
    
    Exit Sub
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

End Sub
Public Sub SetRole(ByVal strRole As String)

    '***********************************************************************************************
    'Purpose    : Sets the role id
    'Created on :
    'Created by: Gordon
    '***********************************************************************************************

    '-------------------------------
    '*** Generic Procedure Code
    '-------------------------------
    On Error GoTo ERROR_HANDLER
    Dim strModuleAndProcName As String
    strModuleAndProcName = "CUser - SetRole Sub"
    Dim lngRoleId As Long
    Dim idx As Integer
    Dim RoleList As sssRoleList
    
'    m_rsRoles.MoveFirst
'    While Not m_rsRoles.EOF
'        If m_rsRoles("role_narr") = strRole Then
'            lngRoleId = m_rsRoles("role_id")
'        End If
'        m_rsRoles.MoveNext
'    Wend
    For idx = 1 To m_RoleList.Count
        Set RoleList = m_RoleList(idx)
        If RoleList.RoleNarr = strRole Then
            lngRoleId = RoleList.RoleID
            Exit For
        End If
        
    Next
'    '***    do we need to do anything?
'    If lngRoleId = m_lngRoleId Then GoTo EXIT_PROC

    '***    role is different
    m_lngRoleId = lngRoleId
    
    '***    clear out the old configurations
    Set m_dicMenues = New Dictionary
    Set m_AccessRights = New AccessRights
    Set m_dicArguments = New Dictionary
    
    '***    load the parent roles
    LoadParentRoles
    
EXIT_PROC:
    
    RaiseEvent RoleSelectionFinish
    
    Exit Sub
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

End Sub
Public Sub ReadCommand()

    '***********************************************************************************************
    'Purpose    : Populates the sssUtilities object
    'Created on :
    'Created by: Gordon
    '***********************************************************************************************

    '-------------------------------
    '*** Generic Procedure Code
    '-------------------------------
    On Error GoTo ERROR_HANDLER
    Dim strModuleAndProcName As String
    strModuleAndProcName = "ConnectionConfiguration - ReadCommand Sub"
    
    Dim strCommandLine As String
    Dim argv() As String
    Dim parameterv() As String
    Dim databasev() As String
    Dim strTemp As String
    Dim strMessage As String
    Dim strTmp As String
    Dim idx As Integer
    
    Dim blnNewCommand As Boolean
    
    Dim ODBCDatabase As String
    Dim ODBCInstance As String
    Dim ODBCHost As String
    
    Dim colODBC As Collection
    Dim valueInfo As Variant
    
    Const cstODBCE As String = "odbc"
    Const cstUSER As String = "user"
    Const cstPASSWORD As String = "password"
    Const cstBUSINESSUNITCODE As String = "business_unit"
    Const cstDEALER As String = "dealer"
    
    '***    Reads commandline
    strCommandLine = m_strCommandLine
    
    If strCommandLine = "" Then
        Err.Raise 100, "ConnectionConfiguration.ReadCommand", "Default Webserver URL must be passed as first argument."
    End If
    
    'check command format
    blnNewCommand = (InStr(LCase(strCommandLine), cstODBCE & "=") > 0)
    
    If blnNewCommand Then
        argv = Split(strCommandLine)
        
        m_enmInitialAccessMode = ODBC
        m_strConnectMode = "C"
        
        For idx = 0 To UBound(argv)
            If InStr(argv(idx), "=") = 0 Then
                Err.Raise 100, "ConnectionConfiguration.ReadCommand", "Invalid command field - " & argv(idx) & "."
            End If
            
            parameterv = Split(argv(idx), "=")
            
            If UBound(parameterv) <> 1 Then
                Err.Raise 100, "ConnectionConfiguration.ReadCommand", "Invalid command field - " & argv(idx) & "."
            End If
            
            If LCase(parameterv(0)) = cstODBCE Then
                m_strInitialDSN = parameterv(1)
                databasev = Split(m_strInitialDSN, ":")
                
                If UBound(databasev) <> 2 Then
                    Err.Raise 100, "ConnectionConfiguration.ReadCommand", "Invalid odbc field - " & parameterv(1) & "."
                End If
                
                m_strDatabaseName = databasev(0)
                m_strDBInstance = databasev(1)
                m_strDBHost = databasev(2)
                
                If m_strDatabaseName = "" Or m_strDBInstance = "" Or m_strDBHost = "" Then
                    Err.Raise 100, "ConnectionConfiguration.ReadCommand", "Invalid odbc field - " & parameterv(1) & "."
                End If
                
                'validate ODBC detail, in case no password
                Set colODBC = EnumRegistryValues(HKEY_CURRENT_USER, ODBCKey & "\" & m_strInitialDSN)
                
                If colODBC.Count = 0 Then
                    Err.Raise 100, "ConnectionConfiguration.ReadCommand", "ODBC datasource " & m_strInitialDSN & " is not configured."
                Else
                    For Each valueInfo In colODBC
                        If valueInfo(0) = "DATABASE" Then
                            ODBCDatabase = valueInfo(1)
                        ElseIf valueInfo(0) = "SERVER" Then
                            ODBCInstance = valueInfo(1)
                            
                            ODBCHost = QueryValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Informix\SqlHosts\" & ODBCInstance, "HOST")
                        End If
                    Next
                    
                    If m_strDatabaseName <> ODBCDatabase Then
                        strMessage = "Check the Database Name."
                    End If
                    
                    If m_strDBInstance <> ODBCInstance Then
                        If strMessage = "" Then
                            strMessage = "Check the Server Name."
                        Else
                            strMessage = strMessage & vbNewLine & "Check the Server Name."
                        End If
                    End If
                    
                    If m_strDBHost <> ODBCHost Then
                        If strMessage = "" Then
                            strMessage = "Check the Host Name."
                        Else
                            strMessage = strMessage & vbNewLine & "Check the Host Name."
                        End If
                    End If
                    
                    If strMessage <> "" Then
                        Err.Raise 100, "ConnectionConfiguration.ReadCommand", "ODBC Details mismatch:" & vbNewLine & strMessage
                    End If
                End If
            ElseIf LCase(parameterv(0)) = cstDEALER Then
                m_strDealerID = UCase(parameterv(1))
                
            ElseIf LCase(parameterv(0)) = cstUSER Then
                m_strLogonCode = parameterv(1)
            ElseIf LCase(parameterv(0)) = cstPASSWORD Then
                If parameterv(1) = "{no-password}" Then
                    m_blnNoPassword = True
                End If
                m_strPassword = parameterv(1)
            ElseIf LCase(parameterv(0)) = cstBUSINESSUNITCODE Then
                m_strCurrentBusinessUnitCode = parameterv(1)
                m_strDefaultBusinessUnitCode = m_strCurrentBusinessUnit
            End If
        Next
    Else
        '***    Gets all commandline arguments in argv.
        SeperateArguments strCommandLine, argv()
    
        '***    see that we got the minimum number of arguments
        If UBound(argv()) < 5 Then
            Err.Raise vbObjectError + 1300, strModuleAndProcName, "Must have a minimum of five arguments. ODBC/HTTP DSN URL DATABASE C/D"
        End If
        '***    checks and assigns values from commandline
        strTemp = argv(1)
    
        '***    must be passed the initial access method
        If UCase(strTemp) <> "ODBC" Then
            Err.Raise vbObjectError + 1300, strModuleAndProcName, "Argument 1 must be 'ODBC'."
        Else
            m_enmInitialAccessMode = ODBC
        End If
        
        '***    get the DSN
        m_strInitialDSN = argv(2)
        m_strInitialURL = argv(3)
        If m_strInitialURL <> "chost01/selcomm" Then
            Err.Raise vbObjectError + 1300, strModuleAndProcName, "Argument 3 must be 'chost01/selcomm'"
        End If
        m_strDatabaseName = argv(4)
        m_strConnectMode = argv(5)
            
        '***    if the connect mode is C (command line) then we must also pass the business unit code
        If UCase(m_strConnectMode) <> "C" Then
            Err.Raise vbObjectError + 1300, strModuleAndProcName, "Argument 3 must be 'C'"
        ElseIf UBound(argv()) < 8 Then
            Err.Raise vbObjectError + 1300, strModuleAndProcName, "Must have a minimum of eight arguments for command line connection. ODBC DSN chost01/selcomm DATABASE C BUSINESS_UNIT_CODE DB_HOST DB_INSTANCE "
        End If
        
        If UCase(m_strConnectMode) = "C" Then
            m_strCurrentBusinessUnitCode = UCase(argv(6))
            m_strDefaultBusinessUnitCode = m_strCurrentBusinessUnitCode
            m_strDBHost = argv(7)
            m_strDBInstance = argv(8)
        End If
        
        'check dsn against database, server and host
        If m_strInitialDSN <> m_strDatabaseName & ":" & m_strDBInstance & ":" & m_strDBHost Then
            Err.Raise vbObjectError + 1300, strModuleAndProcName, "Command line is configured incorrectly: hostname not consistent with dsn name."
        End If
        
        If (UBound(argv()) >= 9) Then
            m_strDealerID = UCase$(Trim(argv(9)))
        End If
    End If
    
    Exit Sub
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

End Sub
Private Function InvalidDealerID(DealerID As String) As Boolean
    '***********************************************************************************************
    'Purpose    : Check Dealer ID passed in via the command line is valid
    'Created on :
    'Created by: Kazim
    '***********************************************************************************************

    '-------------------------------
    '*** Generic Procedure Code
    '-------------------------------
    On Error GoTo ERROR_HANDLER
    Dim strModuleAndProcName As String
    strModuleAndProcName = "Read Command - Validate Dealer ID"
    
    Dim rs As ADODB.Recordset
    Dim dbAccess As dbAccess
    Dim strSQL As String
    
    InvalidDealerID = False

    '***    get the connection object
    Set dbAccess = GetDbAccessObject(m_strDefaultBusinessUnitCode)
    
    '***    see whether dealer ID exists
    strSQL = "select * from dealers where dealer_id = ?"
    Set rs = dbAccess.LoadRS(strSQL, DealerID)
    
    If rs.BOF And rs.EOF Then InvalidDealerID = True
    
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
Private Sub AddEntity(ByVal strBusinessUnitCode As String, objEntity As sssEntityAttributes, lngEntityId As Long)
    Dim dic As Dictionary
    Dim idx As Integer
    
    '***    create the menues dictionary if required
    If m_dicEntities Is Nothing Then Set m_dicEntities = New Dictionary
    '***    does the dictionary have a collection of menues for this business unit?
    If Not m_dicEntities.Exists(strBusinessUnitCode) Then
        Set dic = New Dictionary
        m_dicEntities.Add strBusinessUnitCode, dic
    Else
        Set dic = m_dicEntities.item(strBusinessUnitCode)
    End If
    
    '***    see if this menu is already loaded
    If Not dic.Exists(CStr(lngEntityId)) Then
        dic.Add CStr(lngEntityId), objEntity
    End If
End Sub
Private Sub AddMenu(ByVal strBusinessUnitCode As String, objMenu As sssMenu, lngMenuId As Long)
    Dim dic As Dictionary
    Dim idx As Integer
    
    '***    create the menues dictionary if required
    If m_dicMenues Is Nothing Then Set m_dicMenues = New Dictionary
    '***    does the dictionary have a collection of menues for this business unit?
    If Not m_dicMenues.Exists(strBusinessUnitCode) Then
        Set dic = New Dictionary
        m_dicMenues.Add strBusinessUnitCode, dic
    Else
        Set dic = m_dicMenues.item(strBusinessUnitCode)
    End If
    
    '***    see if this menu is already loaded
    If Not dic.Exists(CStr(lngMenuId)) Then
        dic.Add CStr(lngMenuId), objMenu
    End If
End Sub

Public Sub RemoveMenu(lngMenuId As Long)
    Dim dic As Dictionary
    Dim idx As Integer
    Dim strBusinessUnitCode As String
    
    '***    create the menues dictionary if required
    If m_dicMenues Is Nothing Then Exit Sub
    
    If m_blnGlobalMenues Then
        strBusinessUnitCode = m_strDefaultBusinessUnitCode
    Else
        strBusinessUnitCode = m_strCurrentBusinessUnitCode
    End If
    
    '***    does the dictionary have a collection of menues for this business unit?
    If m_dicMenues.Exists(strBusinessUnitCode) Then
        Set dic = m_dicMenues.item(strBusinessUnitCode)
                
        
        '***    see if this menu is already loaded
        If dic.Exists(CStr(lngMenuId)) Then
            dic.Remove CStr(lngMenuId)
        End If
    End If
    
End Sub

Private Function EntityExists(ByVal strBusinessUnitCode As String, lngEntityId As Long) As Boolean
    Dim dic As Dictionary
    Dim idx As Integer
    
    '***    create the menues dictionary if required
    If m_dicEntities Is Nothing Then Set m_dicEntities = New Dictionary
    '***    does the dictionary have a collection of menues for this business unit?
    If Not m_dicEntities.Exists(strBusinessUnitCode) Then
        Set dic = New Dictionary
        m_dicEntities.Add strBusinessUnitCode, dic
    Else
        Set dic = m_dicEntities.item(strBusinessUnitCode)
    End If
    EntityExists = dic.Exists(CStr(lngEntityId))
End Function
Private Function MenuExists(ByVal strBusinessUnitCode As String, lngMenuId As Long) As Boolean
    Dim dic As Dictionary
    Dim idx As Integer
    
    '***    create the menues dictionary if required
    If m_dicMenues Is Nothing Then Set m_dicMenues = New Dictionary
    '***    does the dictionary have a collection of menues for this business unit?
    If Not m_dicMenues.Exists(strBusinessUnitCode) Then
        Set dic = New Dictionary
        m_dicMenues.Add strBusinessUnitCode, dic
    Else
        Set dic = m_dicMenues.item(strBusinessUnitCode)
    End If
    MenuExists = dic.Exists(CStr(lngMenuId))
End Function
'Private Sub AddAccess(ByVal strBusinessUnitCode As String, objAccess As AccessRight, lngObjectId As Long)
'    Dim dic As Dictionary
'    Dim idx As Integer
'
'    '***    create the access rights dictionary if required
'    If m_dicAccessRights Is Nothing Then Set m_dicAccessRights = New Dictionary
'    '***    does the dictionary have a collection of menues for this business unit?
'    If Not m_dicAccessRights.Exists(strBusinessUnitCode) Then
'        Set dic = New Dictionary
'        m_dicAccessRights.Add strBusinessUnitCode, dic
'    Else
'        Set dic = m_dicAccessRights.Item(strBusinessUnitCode)
'    End If
'
'    '***    see if this menu is already loaded
'    If Not dic.Exists(CStr(lngObjectId)) Then
'        dic.Add CStr(lngObjectId), objAccess
'    End If
'End Sub
Public Function CreateMessageString(sMessage As String, _
                        Optional sCategory As String = "GENERAL", _
                        Optional lSessionId As Long = 0, _
                        Optional iPriority As Integer = 2) As String
    Dim objMessage As sssMessage
    
    Set objMessage = CreateMessageObject(sMessage, sCategory, _
                        lSessionId, iPriority)
    
    CreateMessageString = objMessage.Serialize
    
    Set objMessage = Nothing
    
End Function
Public Function CreateMessageObject(sMessage As String, _
                        Optional sCategory As String = "GENERAL", _
                        Optional lSessionId As Long = 0, _
                        Optional iPriority As Integer = 2) As sssMessage
    Dim objMessage As New sssMessage
    
    objMessage.Category = sCategory
    objMessage.ContactCode = Me.ContactCode
    objMessage.CurrentBusinessUnit = Me.CurrentBusinessUnit
    objMessage.CurrentBusinessUnitCode = Me.CurrentBusinessUnitCode
    objMessage.DSN = Me.GetDbAccessObject(Me.CurrentBusinessUnitCode).DSN
    objMessage.Email = Me.Email
    objMessage.FirstName = Me.FirstName
    objMessage.LogonCode = Me.LogonCode
    objMessage.Message = sMessage
    objMessage.Name = Me.Name
    objMessage.Password = Me.Password
    objMessage.Priority = iPriority
    
    Set CreateMessageObject = objMessage
    
    Set objMessage = Nothing
    
End Function
Private Sub SeperateArguments(szCommandLine As String, argv() As String)
    Dim i As Integer
    Dim nPos As Integer
    Dim str As String
    
    nPos = InStr(1, szCommandLine, " ")
    ReDim argv(1) As String

    i = 1
    
    While nPos > 0
        ReDim Preserve argv(i) As String
        str = Trim(Mid(szCommandLine, 1, nPos))
        argv(i) = str
        szCommandLine = Trim(Mid(szCommandLine, nPos + 1, Len(szCommandLine)))
        nPos = InStr(1, szCommandLine, " ")
        i = i + 1
    Wend
    
    If szCommandLine <> "" Then
        ReDim Preserve argv(i) As String
        str = Trim(szCommandLine)
        argv(i) = str
    End If
End Sub
Public Sub RaiseMenuEvent(sMode As String, lMenuId As Long, sMenuName As String)
    If sMode = "D" Then     '***    destroyed
        RaiseEvent MenuDestroyed(lMenuId, sMenuName)
    Else
        RaiseEvent MenuCreated(lMenuId, sMenuName)
    End If
End Sub
Public Sub RaiseSendMessageEvent(sMessage As String, _
                            Optional sCategory As String = "GENERAL", _
                            Optional iPriority As Integer = 3)
    Dim sMessageString As String
    sMessageString = CreateMessageString(sMessage, sCategory, 0, iPriority)
    RaiseEvent SendMessage(sMessageString)
    
End Sub
Public Sub RaiseReceiveMessageEvent(sMessage As String)
    RaiseEvent ReceiveMessage(sMessage)
End Sub

Private Sub Class_Initialize()
    m_enLogonStatus = icNotLoggedOn
    Set m_dicBusinessUnits = New Dictionary
    Set m_dicDbAccessObjects = New Dictionary
    Set m_AccessRights = New AccessRights
    m_blnGlobalMenues = True
    m_blnGlobalAccess = True
    m_strComputerName = StrConv(GetLocalComputerName, vbProperCase)
End Sub

Private Sub Class_Terminate()
    Set m_dicBusinessUnits = Nothing
    Set m_dicDbAccessObjects = Nothing
    Set m_AccessRights = Nothing
    Set m_Configurations = Nothing
End Sub

Public Function SwitchDataSource(strNewDatabaseName As String, strNewBusinessUnitCode As String)

    '***********************************************************************************************
    'Purpose    : name
    'Created on :
    'Created by: Gordon
    '***********************************************************************************************

    '-------------------------------
    '*** Generic Procedure Code
    '-------------------------------
    On Error GoTo ERROR_HANDLER
    Dim strModuleAndProcName As String
    strModuleAndProcName = "CUser - SwitchDataSource Sub"
    
    Dim dbAccess As dbAccess
    Dim OriginalDSN As String
    Dim OrigianlDatabaseName As String
    Dim OrigianlDefaultBusinessUnitCode As String
    Dim OriginalCurrentBusinessUnitCode As String
    
    m_blnSwitchDatabase = True
    
    'try close old connection to database
    Set dbAccess = GetDbAccessObject(CurrentBusinessUnitCode)
    dbAccess.CloseDatabase
    
    Set dbAccess = Nothing
    
    'record original config
    OriginalDSN = m_strInitialDSN
    OrigianlDatabaseName = m_strDatabaseName
    OrigianlDefaultBusinessUnitCode = m_strDefaultBusinessUnitCode
    OriginalCurrentBusinessUnitCode = m_strCurrentBusinessUnitCode
    
    Me.CurrentBusinessUnitCode = strNewBusinessUnitCode
        
    If Not ODBCExist(strNewDatabaseName) Then
        If Not CreateNewODBC(m_strInitialDSN, strNewDatabaseName) Then GoTo EXITFUNCTION
        If Not CopyODBC(m_strInitialDSN, strNewDatabaseName) Then GoTo EXITFUNCTION
    End If
    
    m_strInitialDSN = strNewDatabaseName & ":" & m_strDBInstance & ":" & m_strDBHost
    m_strDatabaseName = strNewDatabaseName
    m_strDefaultBusinessUnitCode = strNewBusinessUnitCode
    
    If Logon <> icSuccess Then
        m_strInitialDSN = OriginalDSN
        m_strDatabaseName = OrigianlDatabaseName
        m_strDefaultBusinessUnitCode = OrigianlDefaultBusinessUnitCode
        m_strCurrentBusinessUnitCode = OriginalCurrentBusinessUnitCode
        GoTo EXITFUNCTION
    End If
    
    LoadDataSource True
    
    RaiseEvent SwitchDataSource
    SwitchDataSource = True
    
EXITFUNCTION:
    m_blnSwitchDatabase = False
    
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
'    Err.Raise lErrNum, strErrSource, strErrDesc
    '*** OR ignore the error and continue
    Resume Next

End Function

Private Function ODBCExist(strNewDatabaseName As String) As Boolean
    Dim ODBCName As String
    Dim DriverName As String
        
    ODBCName = strNewDatabaseName & ":" & m_strDBInstance & ":" & m_strDBHost
    
    If Not KeyExists(HKEY_CURRENT_USER, ODBCKey & "\" & ODBCName) Then Exit Function
    
    DriverName = GetRegistryValue(HKEY_CURRENT_USER, ODBCKey & "\" & ODBCSubKey, ODBCName, DriverName)
    
    If DriverName = "" Then Exit Function
    
    ODBCExist = True
End Function

Private Function CopyODBC(OldODBCName As String, NewDatabaseName As String) As Boolean
    Dim NewODBCName As String
    Dim colNew As Collection
    Dim colOld As Collection
    Dim valueInfo As Variant
    
    NewODBCName = NewDatabaseName & ":" & m_strDBInstance & ":" & m_strDBHost
    
    Set colOld = EnumRegistryValues(HKEY_CURRENT_USER, ODBCKey & "\" & OldODBCName)
    Set colNew = New Collection
    
    For Each valueInfo In colOld
        If valueInfo(0) = "DATABASE" Then
            valueInfo(1) = NewDatabaseName
        End If
        colNew.Add valueInfo
    Next
    
    If CreateKey(HKEY_CURRENT_USER, ODBCKey & "\" & NewODBCName) <> 0 Then
        Exit Function
    End If
    
    For Each valueInfo In colNew
        If Not SetRegistryValue(HKEY_CURRENT_USER, ODBCKey & "\" & NewODBCName, valueInfo(0), valueInfo(1), REG_SZ) Then
            Exit Function
        End If
    Next
    
    CopyODBC = True
End Function

Private Function CreateNewODBC(OldODBCName As String, NewDatabaseName As String) As Boolean
    Dim NewODBCName As String
    Dim ODBCValue As String
    
    NewODBCName = NewDatabaseName & ":" & m_strDBInstance & ":" & m_strDBHost
    
    ODBCValue = GetRegistryValue(HKEY_CURRENT_USER, ODBCKey & "\" & ODBCSubKey, OldODBCName)
    
    If ODBCValue = "" Then Exit Function
    
    If SetRegistryValue(HKEY_CURRENT_USER, ODBCKey & "\" & ODBCSubKey, NewODBCName, ODBCValue, REG_SZ) Then
        CreateNewODBC = True
    End If
End Function

Private Function GetClientIPAddress() As String
    Dim objClientInfo As WTSClientInfo.clsTSClientInfo
    
    Set objClientInfo = New WTSClientInfo.clsTSClientInfo
    
    GetClientIPAddress = Trim(objClientInfo.ClientAddress)
    
    Set objClientInfo = Nothing
End Function

Private Function GetClientComputerName() As String
    Dim objClientInfo As WTSClientInfo.clsTSClientInfo
    
    Set objClientInfo = New WTSClientInfo.clsTSClientInfo
    
    GetClientComputerName = Trim(objClientInfo.ClientName)
    
    Set objClientInfo = Nothing
End Function

