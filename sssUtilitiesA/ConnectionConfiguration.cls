VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConnectionConfiguration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Base 1

'***    the User
Private WithEvents m_User As CUser
Attribute m_User.VB_VarHelpID = -1

Dim m_blnPopulated As Boolean

'***    dictionarys
Dim m_dicBusinessUnitCodes As Dictionary
Dim m_dicBusinessUnits As Dictionary
Dim m_dicDSNs As Dictionary
Dim m_dicDatabases As Dictionary
Dim m_dicMSMQServers As Dictionary
Dim m_dicMTSServers As Dictionary
Dim m_dicAccessMethods As Dictionary
Dim m_dicURLs As Dictionary

'***    default indicator
Dim m_intDefaultIndex As Integer

'***    Events
Event PopulateStart()
Event PopulateFinish()

Public Property Let User(objUser As CUser)
    Set m_User = objUser
End Property
Public Sub SetProperties(sBusinessUnitCode As String, sDSN As String, _
                            sDatabase As String, enmAccessmode As AccessMode, _
                            sURL As String)
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
    strModuleAndProcName = "ConnectionConfiguration - SetProperties Sub"
    
    Dim lngResult As Long
    Dim strSQL As String
    Dim rs As ADODB.Recordset
    Dim intRecordCounter As Integer
    Dim strErrorDescription As String
    Dim sBusinessUnit As String
    Dim dbAccess As dbAccess
    
    '***    raise the event
    RaiseEvent PopulateStart
    
    '***    reset
    Set m_dicBusinessUnitCodes = New Dictionary
    Set m_dicBusinessUnits = New Dictionary
    Set m_dicDSNs = New Dictionary
    Set m_dicDatabases = New Dictionary
    Set m_dicMSMQServers = New Dictionary
    Set m_dicMTSServers = New Dictionary
    Set m_dicAccessMethods = New Dictionary
    Set m_dicURLs = New Dictionary
    
    '***    First see if there is a user specific configuration
    strSQL = "select bus_unit_narr " & _
            "from  bus_unit " & _
            "where bus_unit_code = ? and bus_unit_display = 'Y'"
    
    '***    create the database access object
    Set dbAccess = New dbAccess
    
    dbAccess.UserName = m_User.LogonCode
    dbAccess.Password = m_User.Password
    dbAccess.AccessMode = m_User.InitialAccessMode
    dbAccess.DSN = m_User.InitialDSN
    dbAccess.URL = m_User.InitialURL
    Set rs = dbAccess.LoadRS(strSQL, sBusinessUnitCode)
    '***    construct an error if required
    If dbAccess.ErrorCode <> "" Then
        'strErrorDescription = GetError(rs)
        GoTo DATABASE_ERROR
    End If
    intRecordCounter = 0
    
    While Not rs.EOF
        intRecordCounter = intRecordCounter + 1
        sBusinessUnit = FieldToString(rs, "bus_unit_narr")
        rs.MoveNext
    Wend
    
    If intRecordCounter > 0 Then
    
        '***    assemble the dictionarys
        m_dicBusinessUnitCodes.Add sBusinessUnitCode, sBusinessUnitCode
        m_dicBusinessUnits.Add sBusinessUnitCode, sBusinessUnit
        m_dicDSNs.Add sBusinessUnitCode, sDSN
        m_dicDatabases.Add sBusinessUnitCode, sDatabase
        If enmAccessmode = ODBC Then
            m_dicAccessMethods.Add sBusinessUnitCode, "O"
        Else
            m_dicAccessMethods.Add sBusinessUnitCode, "H"
        End If
        m_dicURLs.Add sBusinessUnitCode, sURL
        
        m_intDefaultIndex = intRecordCounter
        m_User.CurrentBusinessUnit = sBusinessUnit
        m_User.CurrentBusinessUnitCode = sBusinessUnitCode
    End If
    
    Set rs = Nothing
    Set dbAccess = Nothing
    
EXIT_PROC:
    m_blnPopulated = True
    '***    raise the event
    RaiseEvent PopulateFinish
    
    Exit Sub
    
DATABASE_ERROR:

    Err.Raise vbObjectError + 200, "ConnectionConfiguration.SetProperties", strErrorDescription
    
    GoTo EXIT_PROC

INVALID_BUSINESS_UNIT:

    Err.Raise vbObjectError + 200, "ConnectionConfiguration.SetProperties", "Invalid Business Unit"
    
    GoTo EXIT_PROC
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
Public Property Get AccessMethod(ByVal strBusinessUnitCode As String) As AccessMode
    If Not m_blnPopulated Then Populate
    If m_dicAccessMethods.item(strBusinessUnitCode) = "O" Then
        AccessMethod = ODBC
    Else
        AccessMethod = http
    End If
End Property
Public Property Get BusinessUnitCodes() As Dictionary
    If Not m_blnPopulated Then Populate
    Set BusinessUnitCodes = m_dicBusinessUnitCodes
End Property
Public Property Get BusinessUnits() As Dictionary
    If Not m_blnPopulated Then Populate
    Set BusinessUnits = m_dicBusinessUnits
End Property
Public Property Get BusinessUnit(ByVal strBusinessUnitCode As String) As String
    If Not m_blnPopulated Then Populate
    BusinessUnit = m_dicBusinessUnits.item(strBusinessUnitCode)
End Property
Public Property Get DSNs() As Dictionary
    If Not m_blnPopulated Then Populate
    Set DSNs = m_dicDSNs
End Property
Public Property Get DSN(ByVal strBusinessUnitCode As String) As String
    If Not m_blnPopulated Then Populate
    DSN = m_dicDSNs.item(strBusinessUnitCode)
End Property
Public Property Get Databases() As Dictionary
    If Not m_blnPopulated Then Populate
    Set Databases = m_dicDatabases
End Property
Public Property Get Database(ByVal strBusinessUnitCode As String) As String
    If Not m_blnPopulated Then Populate
    Database = m_dicDatabases.item(strBusinessUnitCode)
End Property
Public Property Get MSMQServers() As Dictionary
    If Not m_blnPopulated Then Populate
    Set MSMQServers = m_dicMSMQServers
End Property
Public Property Get MSMQServer(ByVal strBusinessUnitCode As String) As String
    If Not m_blnPopulated Then Populate
    MSMQServer = m_dicMSMQServers.item(strBusinessUnitCode)
End Property
Public Property Get MTSServers() As Dictionary
    If Not m_blnPopulated Then Populate
    Set MTSServers = m_dicMTSServers
End Property
Public Property Get MTSServer(ByVal strBusinessUnitCode As String) As String
    If Not m_blnPopulated Then Populate
    MTSServer = m_dicMTSServers.item(strBusinessUnitCode)
End Property
Public Property Get URLs() As Dictionary
    If Not m_blnPopulated Then Populate
    Set URLs = m_dicURLs
End Property
Public Property Get URL(ByVal strBusinessUnitCode As String) As String
    If Not m_blnPopulated Then Populate
    URL = m_dicURLs.item(strBusinessUnitCode)
End Property
Public Property Get DefaultBusinessUnitCode() As String
    If Not m_blnPopulated Then Populate
    If m_intDefaultIndex = 0 Then
        Err.Raise 100, "ConnectionConfiguration.DefaultBusinessUnitCode", "No default index set."
    Else
        DefaultBusinessUnitCode = m_dicBusinessUnitCodes.item(m_intDefaultIndex)
    End If
End Property
Public Property Get DefaultBusinessUnit() As String
    If Not m_blnPopulated Then Populate
    If m_intDefaultIndex = 0 Then
        Err.Raise 100, "ConnectionConfiguration.DefaultBusinessUnit", "No default index set."
    Else
        DefaultBusinessUnit = m_dicBusinessUnits.item(m_intDefaultIndex)
    End If
End Property
Public Property Get DefaultDSN() As String
    If Not m_blnPopulated Then Populate
    If m_intDefaultIndex = 0 Then
        Err.Raise 100, "ConnectionConfiguration.DefaultDSN", "No default index set."
    Else
        DefaultDSN = m_dicDSNs.item(m_intDefaultIndex)
    End If
End Property
Public Property Get DefaultDatabase() As String
    If Not m_blnPopulated Then Populate
    If m_intDefaultIndex = 0 Then
        Err.Raise 100, "ConnectionConfiguration.DefaultDatabase", "No default index set."
    Else
        DefaultDatabase = m_dicDatabases.item(m_intDefaultIndex)
    End If
End Property
Public Property Get DefaultMSMQServer() As String
    If Not m_blnPopulated Then Populate
    If m_intDefaultIndex = 0 Then
        Err.Raise 100, "ConnectionConfiguration.DefaultMSMQServer", "No default index set."
    Else
        DefaultMSMQServer = m_dicMSMQServers.item(m_intDefaultIndex)
    End If
End Property
Public Property Get DefaultMTSServer() As String
    If Not m_blnPopulated Then Populate
    If m_intDefaultIndex = 0 Then
        Err.Raise 100, "ConnectionConfiguration.DefaultMTSServer", "No default index set."
    Else
        DefaultMTSServer = m_dicMTSServers.item(m_intDefaultIndex)
    End If
End Property
Public Property Get DefaultURL() As String
    If Not m_blnPopulated Then Populate
    If m_intDefaultIndex = 0 Then
        Err.Raise 100, "ConnectionConfiguration.DefaultMTSServer", "No default index set."
    Else
        DefaultURL = m_dicURLs.item(m_intDefaultIndex)
    End If
End Property

Public Sub Populate()
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
    strModuleAndProcName = "ConnectionConfiguration - Populate Sub"
    
    Dim lngResult As Long
    Dim strSQL As String
    Dim rs As ADODB.Recordset
    Dim intRecordCounter As Integer
    Dim strErrorDescription As String
    Dim dbAccess As dbAccess
    Dim strBusinessUnitCode As String
    
    '***    raise the event
    RaiseEvent PopulateStart
    
    '***    If the logon code has not been set raise an error
    If m_User.LogonCode = "" Then
        Err.Raise vbObjectError + 300, "ConnectionConfiguration.Populate", "User logon code must be set."
    End If
    
    '***    Fetch the configuration data from the Selcomm DSN
    '***
    '***    Fetch the configuration data from the Selcomm DSN
    strSQL = "select conn_cnfg.*, bus_unit_narr " & _
            "from conn_cnfg, bus_unit " & _
            "where conn_cnfg.bus_unit_code = bus_unit.bus_unit_code "
    Set rs = dbAccess.LoadRS(strSQL)
    '***    construct an error if required
    If dbAccess.ErrorCode <> "" Then
        'strErrorDescription = GetError(rs)
        GoTo DATABASE_ERROR
    End If
    
    intRecordCounter = 0
    While Not rs.EOF
        intRecordCounter = intRecordCounter + 1
        strBusinessUnitCode = FieldToString(rs, "bus_unit_code")
        
        '***    if this business unit is not available for this user then ignore.
        'If Not m_User.IsBusinessUnitAvailable(strBusinessUnitCode) Then GoTo SKIP_RECORD
        
        '***    assemble the dictionarys
        m_dicBusinessUnitCodes.Add strBusinessUnitCode, FieldToString(rs, "bus_unit_code")
        m_dicBusinessUnits.Add strBusinessUnitCode, FieldToString(rs, "bus_unit_narr")
        m_dicDSNs.Add strBusinessUnitCode, FieldToString(rs, "dsn")
        m_dicDatabases.Add strBusinessUnitCode, FieldToString(rs, "dflt_database")
        m_dicMTSServers.Add strBusinessUnitCode, FieldToString(rs, "mts_server")
        m_dicMSMQServers.Add strBusinessUnitCode, FieldToString(rs, "msmq_server")
        m_dicAccessMethods.Add strBusinessUnitCode, FieldToString(rs, "access_method")
        m_dicURLs.Add strBusinessUnitCode, FieldToString(rs, "URL")
        
        '***    record the default index
        If FieldToString(rs, "is_dflt") = "Y" Then
            m_intDefaultIndex = intRecordCounter
            m_User.CurrentBusinessUnit = FieldToString(rs, "bus_unit_narr")
            m_User.CurrentBusinessUnitCode = strBusinessUnitCode
        End If
        
SKIP_RECORD:
        
        rs.MoveNext
    Wend
    Set rs = Nothing
    Set dbAccess = Nothing
    
    '***    set the populated flag
    m_blnPopulated = True
    
EXIT_PROC:

    '***    raise the event
    RaiseEvent PopulateFinish
    
    Exit Sub
    
DATABASE_ERROR:

    Err.Raise vbObjectError + 200, "ConnectionConfiguration.Populate", strErrorDescription
    
    GoTo EXIT_PROC
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
'Public Function CreateDbAccessObject(Optional ByVal strBusinessUnitCode As String) _
'    As dbAccess
'    '***********************************************************************************************
'    'Purpose    : Returns a dbAccess object configured for use.
'    'Created on :
'    'Created by: Gordon
'    '***********************************************************************************************
'
'    '-------------------------------
'    '*** Generic Procedure Code
'    '-------------------------------
'    On Error GoTo ERROR_HANDLER
'    Dim strModuleAndProcName As String
'    strModuleAndProcName = "CUser - CreateDbAccessObject Sub"
'
'    '***    have we configured the configuration object.
'    Dim dbAccess As New dbAccess
'
'    '***    if strBusinessUnitCode is not passed then use the default from the
'    '***    connection configuration object
'    If IsMissing(strBusinessUnitCode) Then
'        dbAccess.AccessMode = Me.ConnectionConfiguration.URLs
'        dbAccess.DatabaseName = Me.ConnectionConfiguration.DefaultDatabase
'        dbAccess.DSN = Me.ConnectionConfiguration.DefaultDSN
'        dbAccess.URL = Me.ConnectionConfiguration.DefaultURL
'    Else
'    End If
'
'    '***    set the Id
'    dbAccess.Id = CStr(m_dicDbAccessObjects.Count + 1)
'    m_dicDbAccessObjects.Add dbAccess.Id, dbAccess
'
'    Set CreateDbAccessObject = dbAccess
'
'EXIT_PROC:
'
'    Exit Function
'    '------------------------------------------------------------
'    '*** Generic Error Handling Code (ensure Exit Proc is above)
'    '------------------------------------------------------------
'ERROR_HANDLER:
'
'    Dim lErrNum As Long
'    Dim strErrDesc As String
'    Dim strErrSource As String
'
'    lErrNum = Err.Number
'    strErrDesc = Err.Description
'    strErrSource = "[" & App.Title & " - " & strModuleAndProcName & " - Build " & App.Major & "." & App.Minor & "." & App.Revision
'    If Erl > 0 Then strErrSource = strErrSource & " - Line No " & str(Erl)
'    If InStr(1, Err.Source, "-->") > 0 Then
'        '*** This error has already been handled by our code
'        strErrSource = strErrSource & "]" & vbNewLine & "  --> " & Err.Source
'    Else
'        '*** Newly generated error, log it here.
'        strErrSource = strErrSource & "]" & vbNewLine & "  --> [Source: " & Err.Source & "]"
'        On Error Resume Next
'        LogError lErrNum, strErrSource, strErrDesc
'    End If
'    '(ALWAYS comment out two of the three options below)
'    '*** EITHER display the error here
'    'DisplayError lErrNum, strErrSource, strErrDesc
'    '*** OR raise the error to the calling procedure
'    Err.Raise lErrNum, strErrSource, strErrDesc
'    '*** OR ignore the error and continue
'    'Resume Next
'
'End Function


Private Sub Class_Initialize()
    Set m_dicBusinessUnitCodes = New Dictionary
    Set m_dicBusinessUnits = New Dictionary
    Set m_dicDSNs = New Dictionary
    Set m_dicDatabases = New Dictionary
    Set m_dicMSMQServers = New Dictionary
    Set m_dicMTSServers = New Dictionary
    Set m_dicAccessMethods = New Dictionary
    Set m_dicURLs = New Dictionary
    m_blnPopulated = False
End Sub

Private Sub Class_Terminate()
    Set m_dicBusinessUnitCodes = Nothing
    Set m_dicBusinessUnits = Nothing
    Set m_dicDSNs = Nothing
    Set m_dicDatabases = Nothing
    Set m_dicMSMQServers = Nothing
    Set m_dicMTSServers = Nothing
    Set m_dicAccessMethods = Nothing
    Set m_dicURLs = Nothing
    Set m_User = Nothing
End Sub
