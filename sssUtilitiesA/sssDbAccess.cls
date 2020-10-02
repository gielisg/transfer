VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "dbAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Option Explicit
Option Base 1

'***    the current transaction mode
Public Enum CurrTransMode
    InTransaction = 0
    NotInTransaction = 1
End Enum
'***    the transaction mode
Public Enum TransAction
    beginwork = 0
    CommitWork = 1
    RollbackWork = 2
End Enum
'***    the access mode
Public Enum AccessMode
    ODBC = 0
    http = 1
End Enum
'***    the Lock Mode
Public Enum LockMode
    wait = 0
    NotWait = 1
End Enum
'***    the isolation mode
Public Enum IsolationMode
    CommittedRead = 0
    DirtyRead = 1
End Enum

'***    state variables
'***    the business unit
Private m_strBusinessUnitCode As String
'***    Id (we may wish to identify different objects)
Private m_strId As String
'***    the current transaction mode
Private m_intTransMode As CurrTransMode
'***    the access mode
Private m_intAccessMode As AccessMode
'***    the current lock mode
Private m_intLockMode As LockMode
'***    the current isolation mode
Private m_intIsolationMode As IsolationMode
'***    the user name
Private m_strUserName As String
'***    the user password
Private m_strPassword As String
'***    the DSN
Private m_strDSN As String
'***    the URL
Private m_strURL As String
'***    the database
Private m_strDatabaseName As String
'***    the error code
Private m_strErrorCode As String
'***    the error description
Private m_strErrorDesc As String
'***    the error message (error description plus more info)
Private m_strErrorMessage As String
'***    the error source
Private m_strErrorSource As String
'***    the last serial number created
Private m_lngSerialNumber As Long
'***    the name of the serial column and table
Private m_strSerialCol As String
Private m_strSerialTable As String
'***    maximum number of records to fetch
Private m_intMaxRows As Integer
'***    last database access datetime
Private m_dtLastDatabaseAccess As Date
Private m_blnRunning As Boolean
Private m_blnStopAutoDisconnect As Boolean
'***    the database connection (only used for ODBC style connections)
Private m_Cn As ADODB.Connection
'***    parameter collection (used when the calling program needs to pass a
'***    dynamic number of parameters so cannot use the ParamArray). Will
'***    use the ParamArray in preference to the parameter collection.
Private m_colParameters As Collection
Private m_colParameterTypes As Collection

'***    properties
Public Property Get Id() As String
    Id = m_strId
End Property
Public Property Let Id(vData As String)
    m_strId = vData
End Property
Public Property Get CurrTransMode() As CurrTransMode
    CurrTransMode = m_intTransMode
End Property
Public Property Get AccessMode() As AccessMode
    AccessMode = m_intAccessMode
End Property
Public Property Let AccessMode(vData As AccessMode)
    m_intAccessMode = vData
End Property
Public Property Get LockMode() As LockMode
    LockMode = m_intLockMode
End Property
Public Property Get IsolationMode() As IsolationMode
    IsolationMode = m_intIsolationMode
End Property
Public Property Get UserName() As String
    UserName = m_strUserName
End Property
Public Property Let UserName(vData As String)
    m_strUserName = vData
End Property
Public Property Get BusinessUnitCode() As String
    BusinessUnitCode = m_strBusinessUnitCode
End Property
Public Property Let BusinessUnitCode(vData As String)
    m_strBusinessUnitCode = vData
End Property
Public Property Get Password() As String
    Password = m_strPassword
End Property
Public Property Let Password(vData As String)
    m_strPassword = vData
End Property
Public Property Get DSN() As String
    DSN = m_strDSN
End Property
Public Property Let DSN(vData As String)
    m_strDSN = vData
End Property
Public Property Get URL() As String
    URL = m_strURL
End Property
Public Property Let URL(vData As String)
    m_strURL = vData
End Property
Public Property Get DatabaseName() As String
    DatabaseName = m_strDatabaseName
End Property
Public Property Let DatabaseName(vData As String)
    m_strDatabaseName = vData
End Property
Public Property Get ErrorCode() As String
    ErrorCode = m_strErrorCode
End Property
Public Property Get errorDesc() As String
    errorDesc = m_strErrorDesc
End Property
Public Property Get ErrorMessage() As String
    ErrorMessage = m_strErrorMessage
End Property
Public Property Get errorSource() As String
    errorSource = m_strErrorSource
End Property
Public Property Get SerialNumber() As Long
    SerialNumber = m_lngSerialNumber
End Property
Public Property Get SerialCol() As String
    SerialCol = m_strSerialCol
End Property
Public Property Let SerialCol(vData As String)
    m_strSerialCol = vData
    m_lngSerialNumber = 0
End Property
Public Property Get SerialTable() As String
    SerialTable = m_strSerialTable
End Property
Public Property Let SerialTable(vData As String)
    m_strSerialTable = vData
End Property
Public Property Get MaxRows() As Integer
    MaxRows = m_intMaxRows
End Property
Public Property Let MaxRows(vData As Integer)
    m_intMaxRows = vData
End Property
Public Property Get Running() As Boolean
    Running = m_blnRunning
End Property
Public Property Let Running(vData As Boolean)
    m_blnRunning = vData
End Property
Public Property Get StopAutoDisconnect() As Boolean
    StopAutoDisconnect = m_blnStopAutoDisconnect
End Property
Public Property Let StopAutoDisconnect(vData As Boolean)
    m_blnStopAutoDisconnect = vData
End Property

Public Property Get LastDatabaseAccess() As Date
    LastDatabaseAccess = m_dtLastDatabaseAccess
End Property
Public Property Let LastDatabaseAccess(vData As Date)
    m_dtLastDatabaseAccess = vData
End Property

Public Property Let Parameters(colParams As Collection)
    Set m_colParameters = colParams
End Property
Public Property Let ParameterTypes(colParams As Collection)
    Set m_colParameterTypes = colParams
End Property
'***    methods
Public Function LoadRS( _
                ByVal strSQL As String, _
                ParamArray Parameters() _
                ) _
        As ADODB.Recordset
    '***********************************************************************************************
    '   Purpose: Load a record set from the database
    '   Created on : 1/9/2001
    '   Created by : Gordon
    '   Modified on :
    '***********************************************************************************************
    
    '-------------------------------
    '*** Generic Procedure Code
    '-------------------------------
    On Error GoTo errHandler
    Dim sModuleAndProcName As String
    sModuleAndProcName = "dbAccess - LoadRS Function"
    
    Dim strURL As String
    Dim rs As ADODB.Recordset
    Dim cmd As ADODB.Command
    Dim param As ADODB.Parameter
    Dim idx As Integer
    Dim kdx As Integer
    Dim var As Variant
    
    Running = True
    
    '***    clear the errors
    m_strErrorCode = ""
    m_strErrorDesc = ""
    
    '***    if the database connection does not exist then create it.
    If m_Cn Is Nothing Then
        Connect
    End If
    
    '***    check the validity of the operation
    ValidateOperation
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
        
    '***    if the mode is ODBC
    If m_intAccessMode = ODBC Then
        
        '***    see if SQL is parameterised
        If UBound(Parameters) >= 0 Then
            
            '***    substitute the placeholders if any
            If InStr(1, strSQL, "?") <> 0 Then
                For idx = 0 To UBound(Parameters)
                    SubstitutePlaceholders strSQL, Parameters(idx), VarTypeString(Parameters(idx))
                Next
            End If
            '***    set the max records
'''            Set rs = New ADODB.Recordset
            rs.MaxRecords = m_intMaxRows
            rs.ActiveConnection = m_Cn
            
            '***    fetch the data
            rs.Open (strSQL)
'            Set cmd = New ADODB.Command
'            Set cmd.ActiveConnection = m_Cn
'
'            cmd.CommandText = strSQL
'            Set rs = cmd.Execute(, Parameters)
            
        ElseIf Not m_colParameters Is Nothing Then
            'Dim objParam As ADODB.Parameter
        
            Set cmd = New ADODB.Command
            Set cmd.ActiveConnection = m_Cn
            cmd.CommandText = strSQL
            idx = m_colParameters.Count
            '***    if there is nothing in the parameter collection
            If idx = 0 Then GoTo NO_PARAM_COL
            
            Dim varParams() As Variant
            ReDim varParams(idx)
            For idx = 1 To m_colParameters.Count
'                Set objParam = cmd.CreateParameter(CStr(idx), adVariant, adParamInput)
'                objParam.Value = CVar(m_colParameters.item(idx))
'                cmd.Parameters.Append objParam
                varParams(idx) = CVar(m_colParameters.item(idx))
'                Set param = cmd.CreateParameter(CStr(idx), adChar, adParamInput, Len(CVar(m_colParameters.item(idx))))
'                param.Value = CVar(m_colParameters.item(idx))
'
'                cmd.Parameters.Append param
            Next
        
            '***    set the parameter collection to nothing so it is not
            '***    used again
            Set m_colParameters = Nothing
            cmd.ActiveConnection.CursorLocation = adUseClient
            Set rs = cmd.Execute(, varParams)
        Else
NO_PARAM_COL:
            '***    set the max records
'''            Set rs = New ADODB.Recordset
            rs.MaxRecords = m_intMaxRows
            rs.ActiveConnection = m_Cn
            
            '***    fetch the data
        
            rs.Open (strSQL)
        
        End If
    Else    '***    http
        '***    fetch the data
        strURL = "http://" & m_strURL & "/Common/Server/ExecuteSQL.asp" & _
               "?DSN=" & URLEncode(m_strDSN) & _
               "&SQL=" & URLEncode(strSQL) & _
               "&Mode=" & URLEncode("S") & _
               "&MaxRows=" & URLEncode(m_intMaxRows)
'               "&LockMode=" & URLEncode(m_intLockMode)
              ' "&IsolationLevel=" & URLEncode(m_intIsolationMode)
    
        ' Determine if there are any parameters to pass
        Dim varValue As Variant
        Dim nCount As Integer
        If UBound(Parameters) >= 0 Then
            kdx = UBound(Parameters)
            nCount = 0
            For idx = 0 To kdx
                nCount = nCount + 1
                varValue = Parameters(idx)
                strURL = strURL & "&Param" & CStr(nCount) & "=" & URLEncode(varValue)
            Next
            strURL = strURL & "&ParamCount=" & CStr(nCount)
        End If
        
        '***    check the collection object
        If Not m_colParameters Is Nothing Then
            If m_colParameters.Count > 0 Then
                For idx = 1 To m_colParameters.Count
                    nCount = nCount + 1
                    varValue = m_colParameters.item(idx)
                    strURL = strURL & "&Param" & CStr(nCount) & "=" & URLEncode(varValue)
                Next
                strURL = strURL & "&ParamCount=" & CStr(nCount)
            End If
            Set m_colParameters = Nothing
        End If
        
'''       Set rs = New ADODB.Recordset
       rs.Open strURL
           
    End If
    
    'offline recordset
    If Not rs.ActiveConnection Is Nothing Then
        Set rs.ActiveConnection = Nothing
    End If
        
    '***    error handling ????
    m_strErrorCode = ""
    m_strErrorDesc = ""
    For idx = 0 To rs.Fields.Count - 1
        If rs.Fields.item(idx).Name = "err_no" Then
            If Not IsNull(rs(idx)) Then m_strErrorCode = CStr(rs.Fields.item(idx).Value)
            'If Not IsNull(rs(1)) Then m_strErrorDesc = CStr(rs.Fields.Item(idx).Value)
        End If
        If rs.Fields.item(idx).Name = "err_desc" Then
            If Not IsNull(rs(1)) Then m_strErrorDesc = CStr(rs.Fields.item(idx).Value)
        End If
    Next
    
    Set LoadRS = rs
    Set rs = Nothing
    
    LastDatabaseAccess = Now
    Running = False
    '***    error handling
    If m_strErrorCode <> "" Then
        Err.Raise CLng(m_strErrorCode), sModuleAndProcName, m_strErrorDesc
    End If
    
    Exit Function
    '------------------------------------------------------------
    '*** Generic Error Handling Code (ensure Exit Proc is above)
    '------------------------------------------------------------
errHandler:
    Running = False

    If Err.Number = 0 Then Exit Function
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim sErrSource As String
    
    lErrNum = Err.Number
    sErrDesc = Err.Description
    sErrSource = "[" & App.Title & " - " & sModuleAndProcName & " - Build " & App.Major & "." & App.Minor & "." & App.Revision
    If Erl > 0 Then sErrSource = sErrSource & " - Line No " & str(Erl)
    If InStr(1, Err.Source, "-->") > 0 Then
        '*** This error has already been handled by our code
        sErrSource = sErrSource & "]" & vbNewLine & "  --> " & Err.Source
    Else
        '*** Newly generated error, log it here.
        sErrSource = sErrSource & "]" & vbNewLine & "  --> [Source: " & Err.Source & "]"
        On Error Resume Next
        LogError lErrNum, sErrSource, sErrDesc
    End If
    '***    CUSTOM error handler
    m_strErrorCode = CStr(lErrNum)
    m_strErrorDesc = sErrDesc
    m_strErrorSource = sErrSource
    '***    set the error code and description in the object
    m_strErrorCode = m_strErrorCode
    m_strErrorDesc = m_strErrorDesc
    
    '(ALWAYS comment out two of the three options below)
    '*** EITHER display the error here
    'DisplayError lErrNum, sErrSource, sErrDesc
    '*** OR raise the error to the calling procedure
    Err.Raise lErrNum, sErrSource, sErrDesc
    '*** OR ignore the error and continue
    'Resume Next
End Function
Public Sub ExecuteSQL(ByVal strSQL As String, ParamArray Parameters())

    '***********************************************************************************************
    '   Purpose: Executes an SQL statement
    '   Created on : 1/9/2001
    '   Created by : Gordon
    '   Modified on :
    '***********************************************************************************************
    
    '-------------------------------
    '*** Generic Procedure Code
    '-------------------------------
    On Error GoTo errHandler
    Dim sModuleAndProcName As String
    sModuleAndProcName = "dbAccess - ExecuteSQL Sub"
    Dim cmd As ADODB.Command
    Dim idx As Integer
    Dim kdx As Integer
    Dim rs As ADODB.Recordset
    Dim strURL As String
    Dim strSQL2 As String
    
    Running = True
    '***    clear the errors
    m_strErrorCode = ""
    m_strErrorDesc = ""
    
    '***    check the validity of the operation
    ValidateOperation
    
    '***    if the mode is ODBC
    If m_intAccessMode = ODBC Then
        '***    if there is no connection
        If m_Cn Is Nothing Then Connect
        
        '***    see if SQL is parameterised
        If UBound(Parameters) >= 0 Then
            Set cmd = New ADODB.Command
            Set cmd.ActiveConnection = m_Cn
            cmd.CommandText = strSQL
            Set rs = cmd.Execute(, Parameters)
        ElseIf Not m_colParameters Is Nothing Then
            'Dim objParam As ADODB.Parameter
            
            Set cmd = New ADODB.Command
            Set cmd.ActiveConnection = m_Cn
            cmd.CommandText = strSQL
            'cmd.CommandType = adCmdFile
            idx = m_colParameters.Count
            Dim varParams() As Variant
            ReDim varParams(idx)
            For idx = 1 To m_colParameters.Count
'                Set objParam = cmd.CreateParameter(CStr(idx), adVariant, adParamInput)
'                objParam.Value = CVar(m_colParameters.Item(idx))
'                cmd.Parameters.Append objParam
                varParams(idx) = CVar(m_colParameters.item(idx))
            Next
            '***    set the parameter collection to nothing so it is not
            '***    used again
            Set m_colParameters = Nothing
            Set rs = cmd.Execute(, varParams)
        Else
            '***    execute the SQL
            m_Cn.Execute strSQL
        End If
   
        '***    if this statement contains a serial field then we fetch the
        '***    serial number from the field
        If m_strSerialCol <> "" Then
            strSQL2 = "select max(" & m_strSerialCol & ") from " & _
                    m_strSerialTable
            Set rs = New ADODB.Recordset
            rs.ActiveConnection = m_Cn
            rs.Open strSQL2
            While Not rs.EOF
                m_lngSerialNumber = CLng(rs(0))
                rs.MoveNext
            Wend
            Set rs = Nothing
        End If
    Else    '***    http
            
        '***    fetch the data
        strURL = "http://" & m_strURL & "/Common/Server/ExecuteSQL.asp" & _
               "?DSN=" & URLEncode(m_strDSN) & _
               "&SQL=" & URLEncode(strSQL) & _
               "&Mode=" & URLEncode("S") & _
               "&MaxRows=" & URLEncode(m_intMaxRows)
'               "&LockMode=" & URLEncode(m_intLockMode)
              ' "&IsolationLevel=" & URLEncode(m_intIsolationMode)
              
        '***    Determine if there are any parameters to pass
        Dim varValue As Variant
        Dim nCount As Integer
        If UBound(Parameters) >= 0 Then
            kdx = UBound(Parameters)
            nCount = 0
            For idx = 0 To kdx
                nCount = nCount + 1
                varValue = Parameters(idx)
                strURL = strURL & "&ParamDT" & CStr(nCount) & "=" & URLEncode(varValue)
            Next
            strURL = strURL & "&ParamDTCount=" & CStr(nCount)
        End If
        
        '***    check the parameter collection object
        If Not m_colParameters Is Nothing Then
            If m_colParameters.Count > 0 Then
                For idx = 1 To m_colParameters.Count
                    nCount = nCount + 1
                    varValue = m_colParameters.item(idx)
                    strURL = strURL & "&Param" & CStr(nCount) & "=" & URLEncode(varValue)
                Next
                strURL = strURL & "&ParamCount=" & CStr(nCount)
            End If
            Set m_colParameters = Nothing
        End If
        '***    Determine if there are any parameters to pass
        '***    check the parameter type collection object
        If Not m_colParameterTypes Is Nothing Then
            If m_colParameterTypes.Count > 0 Then
                For idx = 1 To m_colParameterTypes.Count
                    nCount = nCount + 1
                    varValue = m_colParameterTypes.item(idx)
                    strURL = strURL & "&ParamDT" & CStr(nCount) & "=" & URLEncode(varValue)
                Next
                strURL = strURL & "&ParamDTCount=" & CStr(nCount)
            End If
            Set m_colParameterTypes = Nothing
        End If
        
        Set rs = New ADODB.Recordset
        rs.Open strURL
       
    End If
    
    LastDatabaseAccess = Now
    Running = False
    Exit Sub
    '------------------------------------------------------------
    '*** Generic Error Handling Code (ensure Exit Proc is above)
    '------------------------------------------------------------
errHandler:
    Running = False
    If Err.Number = 0 Then Exit Sub
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim sErrSource As String
    
    lErrNum = Err.Number
    sErrDesc = Err.Description
    sErrSource = "[" & App.Title & " - " & sModuleAndProcName & " - Build " & App.Major & "." & App.Minor & "." & App.Revision
    If Erl > 0 Then sErrSource = sErrSource & " - Line No " & str(Erl)
    If InStr(1, Err.Source, "-->") > 0 Then
        '*** This error has already been handled by our code
        sErrSource = sErrSource & "]" & vbNewLine & "  --> " & Err.Source
    Else
        '*** Newly generated error, log it here.
        sErrSource = sErrSource & "]" & vbNewLine & "  --> [Source: " & Err.Source & "]"
        On Error Resume Next
        LogError lErrNum, sErrSource, sErrDesc
    End If
    '***    CUSTOM error handler
    m_strErrorCode = CStr(lErrNum)
    m_strErrorDesc = sErrDesc
    m_strErrorSource = sErrSource
    '(ALWAYS comment out two of the three options below)
    '*** EITHER display the error here
    'DisplayError lErrNum, sErrSource, sErrDesc
    '*** OR raise the error to the calling procedure
    Err.Raise lErrNum, sErrSource, sErrDesc
    '*** OR ignore the error and continue
    'Resume Next

End Sub
Public Sub SetLockMode(enmLockMode As LockMode)

    '***********************************************************************************************
    '   Purpose: Sets the lock mode of the database
    '   Created on : 1/9/2001
    '   Created by : Gordon
    '   Modified on :
    '***********************************************************************************************
    
    '-------------------------------
    '*** Generic Procedure Code
    '-------------------------------
    On Error GoTo errHandler
    Dim sModuleAndProcName As String
    sModuleAndProcName = "dbAccess - SetLockMode Sub"
    
    '***    if the current lock mode is the one requested then do nothing
    If m_intLockMode = enmLockMode Then Exit Sub
    
    If m_intAccessMode = http Then
        SetLockModeHTTP enmLockMode
    Else
        SetLockModeODBC enmLockMode
    End If
    '***    set the value of the lock mode
    m_intLockMode = enmLockMode
    
    Exit Sub
    '------------------------------------------------------------
    '*** Generic Error Handling Code (ensure Exit Proc is above)
    '------------------------------------------------------------
errHandler:
    If Err.Number = 0 Then Exit Sub
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim sErrSource As String
    
    lErrNum = Err.Number
    sErrDesc = Err.Description
    sErrSource = "[" & App.Title & " - " & sModuleAndProcName & " - Build " & App.Major & "." & App.Minor & "." & App.Revision
    If Erl > 0 Then sErrSource = sErrSource & " - Line No " & str(Erl)
    If InStr(1, Err.Source, "-->") > 0 Then
        '*** This error has already been handled by our code
        sErrSource = sErrSource & "]" & vbNewLine & "  --> " & Err.Source
    Else
        '*** Newly generated error, log it here.
        sErrSource = sErrSource & "]" & vbNewLine & "  --> [Source: " & Err.Source & "]"
        On Error Resume Next
        LogError lErrNum, sErrSource, sErrDesc
    End If
    '***    CUSTOM error handler
    m_strErrorCode = CStr(lErrNum)
    m_strErrorDesc = sErrDesc
    m_strErrorSource = sErrSource
    '(ALWAYS comment out two of the three options below)
    '*** EITHER display the error here
    'DisplayError lErrNum, sErrSource, sErrDesc
    '*** OR raise the error to the calling procedure
    Err.Raise lErrNum, sErrSource, sErrDesc
    '*** OR ignore the error and continue
    'Resume Next
End Sub
Private Sub SetLockModeODBC(enmLockMode As LockMode)
    '***********************************************************************************************
    '   Purpose: Sets the lock mode of the database for ODBC
    '   Created on : 1/9/2001
    '   Created by : Gordon
    '   Modified on :
    '***********************************************************************************************
    
    '-------------------------------
    '*** Generic Procedure Code
    '-------------------------------
    On Error GoTo errHandler
    Dim sModuleAndProcName As String
    sModuleAndProcName = "dbAccess - SetLockModeODBC Sub"
    
    If m_Cn Is Nothing Then Connect
    
    If enmLockMode = wait Then
        m_Cn.Execute "SET LOCK MODE TO WAIT"
    Else
        m_Cn.Execute "SET LOCK MODE TO NOT WAIT "
    End If
    
    Exit Sub
    '------------------------------------------------------------
    '*** Generic Error Handling Code (ensure Exit Proc is above)
    '------------------------------------------------------------
errHandler:
    If Err.Number = 0 Then Exit Sub
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim sErrSource As String
    
    lErrNum = Err.Number
    sErrDesc = Err.Description
    sErrSource = "[" & App.Title & " - " & sModuleAndProcName & " - Build " & App.Major & "." & App.Minor & "." & App.Revision
    If Erl > 0 Then sErrSource = sErrSource & " - Line No " & str(Erl)
    If InStr(1, Err.Source, "-->") > 0 Then
        '*** This error has already been handled by our code
        sErrSource = sErrSource & "]" & vbNewLine & "  --> " & Err.Source
    Else
        '*** Newly generated error, log it here.
        sErrSource = sErrSource & "]" & vbNewLine & "  --> [Source: " & Err.Source & "]"
        On Error Resume Next
        LogError lErrNum, sErrSource, sErrDesc
    End If
    '***    CUSTOM error handler
    m_strErrorCode = CStr(lErrNum)
    m_strErrorDesc = sErrDesc
    m_strErrorSource = sErrSource
    '(ALWAYS comment out two of the three options below)
    '*** EITHER display the error here
    'DisplayError lErrNum, sErrSource, sErrDesc
    '*** OR raise the error to the calling procedure
    Err.Raise lErrNum, sErrSource, sErrDesc
    '*** OR ignore the error and continue
    'Resume Next
End Sub
Private Sub SetLockModeHTTP(enmLockMode As LockMode)
    '***********************************************************************************************
    '   Purpose: Sets the lock mode of the database for HTTP
    '   Created on : 1/9/2001
    '   Created by : Gordon
    '   Modified on :
    '***********************************************************************************************
    
    '-------------------------------
    '*** Generic Procedure Code
    '-------------------------------
    On Error GoTo errHandler
    Dim sModuleAndProcName As String
    sModuleAndProcName = "dbAccess - SetLockModeHTTP Sub"
    '***    TODO what about HTPP nothing for now
    Dim rs As ADODB.Recordset
    Dim strURL As String
    
    '***    validate the properties
    ValidateOperation
    
    '***    fetch the data
    strURL = "http://" & m_strURL & "/Common/Server/ExecuteSQL.asp" & _
            "?DSN=" & URLEncode(m_strDSN) & _
            "&SQL=" & URLEncode("SET LOCK MODE")
'            "&LockMode=" & URLEncode(m_intLockMode) & _
'            "&IsolationLevel=" & URLEncode(m_intIsolationMode)
 
    Set rs = New ADODB.Recordset
    rs.Open strURL
    '***    TODO check the SQL worked!!
    
    Set rs = Nothing
    Exit Sub
    '------------------------------------------------------------
    '*** Generic Error Handling Code (ensure Exit Proc is above)
    '------------------------------------------------------------
errHandler:
    If Err.Number = 0 Then Exit Sub
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim sErrSource As String
    
    lErrNum = Err.Number
    sErrDesc = Err.Description
    sErrSource = "[" & App.Title & " - " & sModuleAndProcName & " - Build " & App.Major & "." & App.Minor & "." & App.Revision
    If Erl > 0 Then sErrSource = sErrSource & " - Line No " & str(Erl)
    If InStr(1, Err.Source, "-->") > 0 Then
        '*** This error has already been handled by our code
        sErrSource = sErrSource & "]" & vbNewLine & "  --> " & Err.Source
    Else
        '*** Newly generated error, log it here.
        sErrSource = sErrSource & "]" & vbNewLine & "  --> [Source: " & Err.Source & "]"
        On Error Resume Next
        LogError lErrNum, sErrSource, sErrDesc
    End If
    '***    CUSTOM error handler
    m_strErrorCode = CStr(lErrNum)
    m_strErrorDesc = sErrDesc
    m_strErrorSource = sErrSource
    '(ALWAYS comment out two of the three options below)
    '*** EITHER display the error here
    'DisplayError lErrNum, sErrSource, sErrDesc
    '*** OR raise the error to the calling procedure
    Err.Raise lErrNum, sErrSource, sErrDesc
    '*** OR ignore the error and continue
    'Resume Next
End Sub

Public Sub SetIsolationMode(enmIsolationMode As IsolationMode)
    '***********************************************************************************************
    '   Purpose: Sets the isolation mode of the database
    '   Created on : 1/9/2001
    '   Created by : Gordon
    '   Modified on :
    '***********************************************************************************************
    
    '-------------------------------
    '*** Generic Procedure Code
    '-------------------------------
    On Error GoTo errHandler
    Dim sModuleAndProcName As String
    sModuleAndProcName = "dbAccess - SetIsolationMode Sub"
    
    '***    if the current isolation mode is the one requested then do nothing
    If m_intIsolationMode = enmIsolationMode Then Exit Sub
    
    If m_intAccessMode = http Then
        SetIsolationModeHTTP enmIsolationMode
    Else
        SetIsolationModeODBC enmIsolationMode
    End If
    
    Exit Sub
    '------------------------------------------------------------
    '*** Generic Error Handling Code (ensure Exit Proc is above)
    '------------------------------------------------------------
errHandler:
    If Err.Number = 0 Then Exit Sub
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim sErrSource As String
    
    lErrNum = Err.Number
    sErrDesc = Err.Description
    sErrSource = "[" & App.Title & " - " & sModuleAndProcName & " - Build " & App.Major & "." & App.Minor & "." & App.Revision
    If Erl > 0 Then sErrSource = sErrSource & " - Line No " & str(Erl)
    If InStr(1, Err.Source, "-->") > 0 Then
        '*** This error has already been handled by our code
        sErrSource = sErrSource & "]" & vbNewLine & "  --> " & Err.Source
    Else
        '*** Newly generated error, log it here.
        sErrSource = sErrSource & "]" & vbNewLine & "  --> [Source: " & Err.Source & "]"
        On Error Resume Next
        LogError lErrNum, sErrSource, sErrDesc
    End If
    '***    CUSTOM error handler
    m_strErrorCode = CStr(lErrNum)
    m_strErrorDesc = sErrDesc
    m_strErrorSource = sErrSource
    '(ALWAYS comment out two of the three options below)
    '*** EITHER display the error here
    'DisplayError lErrNum, sErrSource, sErrDesc
    '*** OR raise the error to the calling procedure
    Err.Raise lErrNum, sErrSource, sErrDesc
    '*** OR ignore the error and continue
    'Resume Next
End Sub
Private Sub SetIsolationModeODBC(enmIsolationMode As IsolationMode)
    '***********************************************************************************************
    '   Purpose: Sets the isolation mode of the database for ODBC
    '   Created on : 1/9/2001
    '   Created by : Gordon
    '   Modified on :
    '***********************************************************************************************
    
    '-------------------------------
    '*** Generic Procedure Code
    '-------------------------------
    On Error GoTo errHandler
    Dim sModuleAndProcName As String
    sModuleAndProcName = "dbAccess - SetIsolationModeODBC Sub"
    
    If enmIsolationMode = CommittedRead Then
        m_Cn.Execute "SET ISOLATION TO COMMITTED READ"
    Else
        m_Cn.Execute "SET ISOLATION TO DIRTY READ"
    End If
    
    Exit Sub
    '------------------------------------------------------------
    '*** Generic Error Handling Code (ensure Exit Proc is above)
    '------------------------------------------------------------
errHandler:
    If Err.Number = 0 Then Exit Sub
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim sErrSource As String
    
    lErrNum = Err.Number
    sErrDesc = Err.Description
    sErrSource = "[" & App.Title & " - " & sModuleAndProcName & " - Build " & App.Major & "." & App.Minor & "." & App.Revision
    If Erl > 0 Then sErrSource = sErrSource & " - Line No " & str(Erl)
    If InStr(1, Err.Source, "-->") > 0 Then
        '*** This error has already been handled by our code
        sErrSource = sErrSource & "]" & vbNewLine & "  --> " & Err.Source
    Else
        '*** Newly generated error, log it here.
        sErrSource = sErrSource & "]" & vbNewLine & "  --> [Source: " & Err.Source & "]"
        On Error Resume Next
        LogError lErrNum, sErrSource, sErrDesc
    End If
    '***    CUSTOM error handler
    m_strErrorCode = CStr(lErrNum)
    m_strErrorDesc = sErrDesc
    m_strErrorSource = sErrSource
    '(ALWAYS comment out two of the three options below)
    '*** EITHER display the error here
    'DisplayError lErrNum, sErrSource, sErrDesc
    '*** OR raise the error to the calling procedure
    Err.Raise lErrNum, sErrSource, sErrDesc
    '*** OR ignore the error and continue
    'Resume Next
End Sub
Private Sub SetIsolationModeHTTP(enmIsolationMode As IsolationMode)
    '***********************************************************************************************
    '   Purpose: Sets the isolation mode of the database for HTTP
    '   Created on : 1/9/2001
    '   Created by : Gordon
    '   Modified on :
    '***********************************************************************************************
    
    '-------------------------------
    '*** Generic Procedure Code
    '-------------------------------
    On Error GoTo errHandler
    Dim sModuleAndProcName As String
    sModuleAndProcName = "dbAccess - SetIsolationModeHTTP Sub"
    Dim rs As ADODB.Recordset
    Dim strURL As String
    
    '***    validate the properties
    ValidateOperation
    
    '***    fetch the data
    strURL = "http://" & m_strURL & "/Common/Server/ExecuteSQL.asp" & _
            "?DSN=" & URLEncode(m_strDSN) & _
            "&SQL=" & URLEncode("SET ISOLATION LEVEL")
'            "&LockMode=" & URLEncode(m_intLockMode) & _
'            "&IsolationLevel=" & URLEncode(m_intIsolationMode)
'
    Set rs = New ADODB.Recordset
    rs.Open strURL
    '***    TODO check the SQL worked!!
    
    Set rs = Nothing
    '***    TODO
    Exit Sub
    '------------------------------------------------------------
    '*** Generic Error Handling Code (ensure Exit Proc is above)
    '------------------------------------------------------------
errHandler:
    If Err.Number = 0 Then Exit Sub
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim sErrSource As String
    
    lErrNum = Err.Number
    sErrDesc = Err.Description
    sErrSource = "[" & App.Title & " - " & sModuleAndProcName & " - Build " & App.Major & "." & App.Minor & "." & App.Revision
    If Erl > 0 Then sErrSource = sErrSource & " - Line No " & str(Erl)
    If InStr(1, Err.Source, "-->") > 0 Then
        '*** This error has already been handled by our code
        sErrSource = sErrSource & "]" & vbNewLine & "  --> " & Err.Source
    Else
        '*** Newly generated error, log it here.
        sErrSource = sErrSource & "]" & vbNewLine & "  --> [Source: " & Err.Source & "]"
        On Error Resume Next
        LogError lErrNum, sErrSource, sErrDesc
    End If
    '***    CUSTOM error handler
    m_strErrorCode = CStr(lErrNum)
    m_strErrorDesc = sErrDesc
    m_strErrorSource = sErrSource
    '(ALWAYS comment out two of the three options below)
    '*** EITHER display the error here
    'DisplayError lErrNum, sErrSource, sErrDesc
    '*** OR raise the error to the calling procedure
    Err.Raise lErrNum, sErrSource, sErrDesc
    '*** OR ignore the error and continue
    'Resume Next
End Sub
Public Sub BeginTrans()
    '***********************************************************************************************
    '   Purpose: Starts a transaction
    '   Created on : 1/9/2001
    '   Created by : Gordon
    '   Modified on :
    '***********************************************************************************************
    
    '-------------------------------
    '*** Generic Procedure Code
    '-------------------------------
    On Error GoTo errHandler
    Dim sModuleAndProcName As String
    sModuleAndProcName = "dbAccess - BeginTrans Sub"
    
    Running = True
    
    If m_intAccessMode = http Then
        BeginTransHTTP
    Else
        BeginTransODBC
    End If
    
    m_intTransMode = InTransaction
    LastDatabaseAccess = Now
    Running = False
    Exit Sub
    '------------------------------------------------------------
    '*** Generic Error Handling Code (ensure Exit Proc is above)
    '------------------------------------------------------------
errHandler:
    If Err.Number = 0 Then Exit Sub
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim sErrSource As String
    
    lErrNum = Err.Number
    sErrDesc = Err.Description
    sErrSource = "[" & App.Title & " - " & sModuleAndProcName & " - Build " & App.Major & "." & App.Minor & "." & App.Revision
    If Erl > 0 Then sErrSource = sErrSource & " - Line No " & str(Erl)
    If InStr(1, Err.Source, "-->") > 0 Then
        '*** This error has already been handled by our code
        sErrSource = sErrSource & "]" & vbNewLine & "  --> " & Err.Source
    Else
        '*** Newly generated error, log it here.
        sErrSource = sErrSource & "]" & vbNewLine & "  --> [Source: " & Err.Source & "]"
        On Error Resume Next
        LogError lErrNum, sErrSource, sErrDesc
    End If
    '***    CUSTOM error handler
    m_strErrorCode = CStr(lErrNum)
    m_strErrorDesc = sErrDesc
    m_strErrorSource = sErrSource
    '(ALWAYS comment out two of the three options below)
    '*** EITHER display the error here
    'DisplayError lErrNum, sErrSource, sErrDesc
    '*** OR raise the error to the calling procedure
    Err.Raise lErrNum, sErrSource, sErrDesc
    '*** OR ignore the error and continue
    'Resume Next
End Sub
Private Sub BeginTransODBC()
    '***********************************************************************************************
    '   Purpose: Starts a transaction for an ODBC connection
    '   Created on : 1/9/2001
    '   Created by : Gordon
    '   Modified on :
    '***********************************************************************************************
    
    '-------------------------------
    '*** Generic Procedure Code
    '-------------------------------
    On Error GoTo errHandler
    Dim sModuleAndProcName As String
    sModuleAndProcName = "dbAccess - BeginTransODBC Sub"
    
    If m_Cn Is Nothing Then Connect
    
    m_Cn.BeginTrans
    
    Exit Sub
    '------------------------------------------------------------
    '*** Generic Error Handling Code (ensure Exit Proc is above)
    '------------------------------------------------------------
errHandler:
    If Err.Number = 0 Then Exit Sub
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim sErrSource As String
    
    lErrNum = Err.Number
    sErrDesc = Err.Description
    sErrSource = "[" & App.Title & " - " & sModuleAndProcName & " - Build " & App.Major & "." & App.Minor & "." & App.Revision
    If Erl > 0 Then sErrSource = sErrSource & " - Line No " & str(Erl)
    If InStr(1, Err.Source, "-->") > 0 Then
        '*** This error has already been handled by our code
        sErrSource = sErrSource & "]" & vbNewLine & "  --> " & Err.Source
    Else
        '*** Newly generated error, log it here.
        sErrSource = sErrSource & "]" & vbNewLine & "  --> [Source: " & Err.Source & "]"
        On Error Resume Next
        LogError lErrNum, sErrSource, sErrDesc
    End If
    '***    CUSTOM error handler
    m_strErrorCode = CStr(lErrNum)
    m_strErrorDesc = sErrDesc
    m_strErrorSource = sErrSource
    '(ALWAYS comment out two of the three options below)
    '*** EITHER display the error here
    'DisplayError lErrNum, sErrSource, sErrDesc
    '*** OR raise the error to the calling procedure
    Err.Raise lErrNum, sErrSource, sErrDesc
    '*** OR ignore the error and continue
    'Resume Next
End Sub
Private Sub BeginTransHTTP()
    '***********************************************************************************************
    '   Purpose: Starts a transaction for an HTTP connection
    '   Created on : 1/9/2001
    '   Created by : Gordon
    '   Modified on :
    '***********************************************************************************************
    
    '-------------------------------
    '*** Generic Procedure Code
    '-------------------------------
    On Error GoTo errHandler
    Dim sModuleAndProcName As String
    sModuleAndProcName = "dbAccess - BeginTransHTTP Sub"
    
    Dim rs As ADODB.Recordset
    Dim strURL As String
    
    '***    validate the properties
    ValidateOperation
    
    '***    fetch the data
    strURL = "http://" & m_strURL & "/Common/Server/ExecuteSQL.asp" & _
            "?DSN=" & URLEncode(m_strDSN) & _
            "&SQL=" & URLEncode("BEGIN WORK")
'            "&LockMode=" & URLEncode(m_intLockMode) & _
'            "&IsolationLevel=" & URLEncode(m_intIsolationMode)
 
    Set rs = New ADODB.Recordset
    rs.Open strURL
    '***    TODO check the SQL worked!!
    
    Set rs = Nothing
    Exit Sub
    '------------------------------------------------------------
    '*** Generic Error Handling Code (ensure Exit Proc is above)
    '------------------------------------------------------------
errHandler:
    If Err.Number = 0 Then Exit Sub
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim sErrSource As String
    
    lErrNum = Err.Number
    sErrDesc = Err.Description
    sErrSource = "[" & App.Title & " - " & sModuleAndProcName & " - Build " & App.Major & "." & App.Minor & "." & App.Revision
    If Erl > 0 Then sErrSource = sErrSource & " - Line No " & str(Erl)
    If InStr(1, Err.Source, "-->") > 0 Then
        '*** This error has already been handled by our code
        sErrSource = sErrSource & "]" & vbNewLine & "  --> " & Err.Source
    Else
        '*** Newly generated error, log it here.
        sErrSource = sErrSource & "]" & vbNewLine & "  --> [Source: " & Err.Source & "]"
        On Error Resume Next
        LogError lErrNum, sErrSource, sErrDesc
    End If
    '***    CUSTOM error handler
    m_strErrorCode = CStr(lErrNum)
    m_strErrorDesc = sErrDesc
    m_strErrorSource = sErrSource
    '(ALWAYS comment out two of the three options below)
    '*** EITHER display the error here
    'DisplayError lErrNum, sErrSource, sErrDesc
    '*** OR raise the error to the calling procedure
    Err.Raise lErrNum, sErrSource, sErrDesc
    '*** OR ignore the error and continue
    'Resume Next
End Sub
Public Sub CommitTrans()
    '***********************************************************************************************
    '   Purpose: Commits a transaction
    '   Created on : 1/9/2001
    '   Created by : Gordon
    '   Modified on :
    '***********************************************************************************************
    
    '-------------------------------
    '*** Generic Procedure Code
    '-------------------------------
    On Error GoTo errHandler
    Dim sModuleAndProcName As String
    sModuleAndProcName = "dbAccess - CommitTrans Sub"
    
    If m_intAccessMode = http Then
        CommitTransHTTP
    Else
        CommitTransODBC
    End If
    
    LastDatabaseAccess = Now
    
    m_intTransMode = NotInTransaction
    Exit Sub
    '------------------------------------------------------------
    '*** Generic Error Handling Code (ensure Exit Proc is above)
    '------------------------------------------------------------
errHandler:
    If Err.Number = 0 Then Exit Sub
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim sErrSource As String
    
    lErrNum = Err.Number
    sErrDesc = Err.Description
    sErrSource = "[" & App.Title & " - " & sModuleAndProcName & " - Build " & App.Major & "." & App.Minor & "." & App.Revision
    If Erl > 0 Then sErrSource = sErrSource & " - Line No " & str(Erl)
    If InStr(1, Err.Source, "-->") > 0 Then
        '*** This error has already been handled by our code
        sErrSource = sErrSource & "]" & vbNewLine & "  --> " & Err.Source
    Else
        '*** Newly generated error, log it here.
        sErrSource = sErrSource & "]" & vbNewLine & "  --> [Source: " & Err.Source & "]"
        On Error Resume Next
        LogError lErrNum, sErrSource, sErrDesc
    End If
    '***    CUSTOM error handler
    m_strErrorCode = CStr(lErrNum)
    m_strErrorDesc = sErrDesc
    m_strErrorSource = sErrSource
    '(ALWAYS comment out two of the three options below)
    '*** EITHER display the error here
    'DisplayError lErrNum, sErrSource, sErrDesc
    '*** OR raise the error to the calling procedure
    Err.Raise lErrNum, sErrSource, sErrDesc
    '*** OR ignore the error and continue
    'Resume Next
End Sub
Private Sub CommitTransODBC()
    '***********************************************************************************************
    '   Purpose: Commits a transaction for an ODBC connection
    '   Created on : 1/9/2001
    '   Created by : Gordon
    '   Modified on :
    '***********************************************************************************************
    
    '-------------------------------
    '*** Generic Procedure Code
    '-------------------------------
    On Error GoTo errHandler
    Dim sModuleAndProcName As String
    sModuleAndProcName = "dbAccess - CommitTransODBC Sub"
    
    m_Cn.CommitTrans
    
    Exit Sub
    '------------------------------------------------------------
    '*** Generic Error Handling Code (ensure Exit Proc is above)
    '------------------------------------------------------------
errHandler:
    If Err.Number = 0 Then Exit Sub
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim sErrSource As String
    
    lErrNum = Err.Number
    sErrDesc = Err.Description
    sErrSource = "[" & App.Title & " - " & sModuleAndProcName & " - Build " & App.Major & "." & App.Minor & "." & App.Revision
    If Erl > 0 Then sErrSource = sErrSource & " - Line No " & str(Erl)
    If InStr(1, Err.Source, "-->") > 0 Then
        '*** This error has already been handled by our code
        sErrSource = sErrSource & "]" & vbNewLine & "  --> " & Err.Source
    Else
        '*** Newly generated error, log it here.
        sErrSource = sErrSource & "]" & vbNewLine & "  --> [Source: " & Err.Source & "]"
        On Error Resume Next
        LogError lErrNum, sErrSource, sErrDesc
    End If
    '***    CUSTOM error handler
    m_strErrorCode = CStr(lErrNum)
    m_strErrorDesc = sErrDesc
    m_strErrorSource = sErrSource
    '(ALWAYS comment out two of the three options below)
    '*** EITHER display the error here
    'DisplayError lErrNum, sErrSource, sErrDesc
    '*** OR raise the error to the calling procedure
    Err.Raise lErrNum, sErrSource, sErrDesc
    '*** OR ignore the error and continue
    'Resume Next
End Sub
Private Sub CommitTransHTTP()
    '***********************************************************************************************
    '   Purpose: Commits a transaction for HTTP
    '   Created on : 1/9/2001
    '   Created by : Gordon
    '   Modified on :
    '***********************************************************************************************
    
    '-------------------------------
    '*** Generic Procedure Code
    '-------------------------------
    On Error GoTo errHandler
    Dim sModuleAndProcName As String
    sModuleAndProcName = "dbAccess - CommitTransHTTP Sub"
    Dim rs As ADODB.Recordset
    Dim strURL As String
    
    '***    validate the properties
    ValidateOperation
    
    '***    fetch the data
    strURL = "http://" & m_strURL & "/Common/Server/ExecuteSQL.asp" & _
            "?DSN=" & URLEncode(m_strDSN) & _
            "&SQL=" & URLEncode("COMMIT WORK")
'            "&LockMode=" & URLEncode(m_intLockMode) & _
'            "&IsolationLevel=" & URLEncode(m_intIsolationMode)
 
    Set rs = New ADODB.Recordset
    rs.Open strURL
    '***    TODO check the SQL worked!!
    
    Set rs = Nothing
    Exit Sub
    '------------------------------------------------------------
    '*** Generic Error Handling Code (ensure Exit Proc is above)
    '------------------------------------------------------------
errHandler:
    If Err.Number = 0 Then Exit Sub
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim sErrSource As String
    
    lErrNum = Err.Number
    sErrDesc = Err.Description
    sErrSource = "[" & App.Title & " - " & sModuleAndProcName & " - Build " & App.Major & "." & App.Minor & "." & App.Revision
    If Erl > 0 Then sErrSource = sErrSource & " - Line No " & str(Erl)
    If InStr(1, Err.Source, "-->") > 0 Then
        '*** This error has already been handled by our code
        sErrSource = sErrSource & "]" & vbNewLine & "  --> " & Err.Source
    Else
        '*** Newly generated error, log it here.
        sErrSource = sErrSource & "]" & vbNewLine & "  --> [Source: " & Err.Source & "]"
        On Error Resume Next
        LogError lErrNum, sErrSource, sErrDesc
    End If
    '***    CUSTOM error handler
    m_strErrorCode = CStr(lErrNum)
    m_strErrorDesc = sErrDesc
    m_strErrorSource = sErrSource
    '(ALWAYS comment out two of the three options below)
    '*** EITHER display the error here
    'DisplayError lErrNum, sErrSource, sErrDesc
    '*** OR raise the error to the calling procedure
    Err.Raise lErrNum, sErrSource, sErrDesc
    '*** OR ignore the error and continue
    'Resume Next
End Sub
Public Sub RollbackTrans()
    '***********************************************************************************************
    '   Purpose: Rollback a transaction
    '   Created on : 1/9/2001
    '   Created by : Gordon
    '   Modified on :
    '***********************************************************************************************
    
    '-------------------------------
    '*** Generic Procedure Code
    '-------------------------------
    On Error GoTo errHandler
    Dim sModuleAndProcName As String
    sModuleAndProcName = "dbAccess - RollbackTrans Sub"
    
    If m_intAccessMode = http Then
        RollbackTransHTTP
    Else
        RollbackTransODBC
    End If
    
    LastDatabaseAccess = Now
    
    m_intTransMode = NotInTransaction
    Exit Sub
    '------------------------------------------------------------
    '*** Generic Error Handling Code (ensure Exit Proc is above)
    '------------------------------------------------------------
errHandler:
    If Err.Number = 0 Then Exit Sub
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim sErrSource As String
    
    lErrNum = Err.Number
    sErrDesc = Err.Description
    sErrSource = "[" & App.Title & " - " & sModuleAndProcName & " - Build " & App.Major & "." & App.Minor & "." & App.Revision
    If Erl > 0 Then sErrSource = sErrSource & " - Line No " & str(Erl)
    If InStr(1, Err.Source, "-->") > 0 Then
        '*** This error has already been handled by our code
        sErrSource = sErrSource & "]" & vbNewLine & "  --> " & Err.Source
    Else
        '*** Newly generated error, log it here.
        sErrSource = sErrSource & "]" & vbNewLine & "  --> [Source: " & Err.Source & "]"
        On Error Resume Next
        LogError lErrNum, sErrSource, sErrDesc
    End If
    '***    CUSTOM error handler
    m_strErrorCode = CStr(lErrNum)
    m_strErrorDesc = sErrDesc
    m_strErrorSource = sErrSource
    '(ALWAYS comment out two of the three options below)
    '*** EITHER display the error here
    'DisplayError lErrNum, sErrSource, sErrDesc
    '*** OR raise the error to the calling procedure
    Err.Raise lErrNum, sErrSource, sErrDesc
    '*** OR ignore the error and continue
    'Resume Next
End Sub
Private Sub RollbackTransODBC()
    '***********************************************************************************************
    '   Purpose: Rollback a transaction for an ODBC connection
    '   Created on : 1/9/2001
    '   Created by : Gordon
    '   Modified on :
    '***********************************************************************************************
    
    '-------------------------------
    '*** Generic Procedure Code
    '-------------------------------
    On Error GoTo errHandler
    Dim sModuleAndProcName As String
    sModuleAndProcName = "dbAccess - RollbackTransODBC Sub"
    
    m_Cn.RollbackTrans
    
    Exit Sub
    '------------------------------------------------------------
    '*** Generic Error Handling Code (ensure Exit Proc is above)
    '------------------------------------------------------------
errHandler:
    If Err.Number = 0 Then Exit Sub
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim sErrSource As String
    
    lErrNum = Err.Number
    sErrDesc = Err.Description
    sErrSource = "[" & App.Title & " - " & sModuleAndProcName & " - Build " & App.Major & "." & App.Minor & "." & App.Revision
    If Erl > 0 Then sErrSource = sErrSource & " - Line No " & str(Erl)
    If InStr(1, Err.Source, "-->") > 0 Then
        '*** This error has already been handled by our code
        sErrSource = sErrSource & "]" & vbNewLine & "  --> " & Err.Source
    Else
        '*** Newly generated error, log it here.
        sErrSource = sErrSource & "]" & vbNewLine & "  --> [Source: " & Err.Source & "]"
        On Error Resume Next
        LogError lErrNum, sErrSource, sErrDesc
    End If
    '***    CUSTOM error handler
    m_strErrorCode = CStr(lErrNum)
    m_strErrorDesc = sErrDesc
    m_strErrorSource = sErrSource
    '(ALWAYS comment out two of the three options below)
    '*** EITHER display the error here
    'DisplayError lErrNum, sErrSource, sErrDesc
    '*** OR raise the error to the calling procedure
    Err.Raise lErrNum, sErrSource, sErrDesc
    '*** OR ignore the error and continue
    'Resume Next
End Sub
Private Sub RollbackTransHTTP()
    '***********************************************************************************************
    '   Purpose: Rollback a transaction for HTTP
    '   Created on : 1/9/2001
    '   Created by : Gordon
    '   Modified on :
    '***********************************************************************************************
    
    '-------------------------------
    '*** Generic Procedure Code
    '-------------------------------
    On Error GoTo errHandler
    Dim sModuleAndProcName As String
    sModuleAndProcName = "dbAccess - RollbackTransHTTP Sub"
    Dim rs As ADODB.Recordset
    Dim strURL As String
    
    '***    validate the properties
    ValidateOperation
    
    '***    fetch the data
    strURL = "http://" & m_strURL & "/Common/Server/ExecuteSQL.asp" & _
            "?DSN=" & URLEncode(m_strDSN) & _
            "&SQL=" & URLEncode("ROLLBACK WORK")
'            "&LockMode=" & URLEncode(m_intLockMode) & _
'            "&IsolationLevel=" & URLEncode(m_intIsolationMode)
 
    Set rs = New ADODB.Recordset
    rs.Open strURL
    '***    TODO check the SQL worked!!
    
    Set rs = Nothing
    Exit Sub
    '------------------------------------------------------------
    '*** Generic Error Handling Code (ensure Exit Proc is above)
    '------------------------------------------------------------
errHandler:
    If Err.Number = 0 Then Exit Sub
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim sErrSource As String
    
    lErrNum = Err.Number
    sErrDesc = Err.Description
    sErrSource = "[" & App.Title & " - " & sModuleAndProcName & " - Build " & App.Major & "." & App.Minor & "." & App.Revision
    If Erl > 0 Then sErrSource = sErrSource & " - Line No " & str(Erl)
    If InStr(1, Err.Source, "-->") > 0 Then
        '*** This error has already been handled by our code
        sErrSource = sErrSource & "]" & vbNewLine & "  --> " & Err.Source
    Else
        '*** Newly generated error, log it here.
        sErrSource = sErrSource & "]" & vbNewLine & "  --> [Source: " & Err.Source & "]"
        On Error Resume Next
        LogError lErrNum, sErrSource, sErrDesc
    End If
    '***    CUSTOM error handler
    m_strErrorCode = CStr(lErrNum)
    m_strErrorDesc = sErrDesc
    m_strErrorSource = sErrSource
    '(ALWAYS comment out two of the three options below)
    '*** EITHER display the error here
    'DisplayError lErrNum, sErrSource, sErrDesc
    '*** OR raise the error to the calling procedure
    Err.Raise lErrNum, sErrSource, sErrDesc
    '*** OR ignore the error and continue
    'Resume Next
End Sub
Public Sub Connect()
    '***********************************************************************************************
    '   Purpose: Connects to database
    '   Created on : 1/9/2001
    '   Created by : Gordon
    '   Modified on :
    '***********************************************************************************************
    
    '-------------------------------
    '*** Generic Procedure Code
    '-------------------------------
    On Error GoTo errHandler
    Dim sModuleAndProcName As String
    sModuleAndProcName = "dbAccess - Connect Sub"
    
    If m_intAccessMode = http Then
        'ConnectHTTP
    Else
        ConnectODBC
    End If
    
    Exit Sub
    '------------------------------------------------------------
    '*** Generic Error Handling Code (ensure Exit Proc is above)
    '------------------------------------------------------------
errHandler:
    If Err.Number = 0 Then Exit Sub
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim sErrSource As String
    
    lErrNum = Err.Number
    sErrDesc = Err.Description
    sErrSource = "[" & App.Title & " - " & sModuleAndProcName & " - Build " & App.Major & "." & App.Minor & "." & App.Revision
    If Erl > 0 Then sErrSource = sErrSource & " - Line No " & str(Erl)
    If InStr(1, Err.Source, "-->") > 0 Then
        '*** This error has already been handled by our code
        sErrSource = sErrSource & "]" & vbNewLine & "  --> " & Err.Source
    Else
        '*** Newly generated error, log it here.
        sErrSource = sErrSource & "]" & vbNewLine & "  --> [Source: " & Err.Source & "]"
        On Error Resume Next
        LogError lErrNum, sErrSource, sErrDesc
    End If
    '***    CUSTOM error handler
    m_strErrorCode = CStr(lErrNum)
    m_strErrorDesc = sErrDesc
    m_strErrorSource = sErrSource
    '(ALWAYS comment out two of the three options below)
    '*** EITHER display the error here
    'DisplayError lErrNum, sErrSource, sErrDesc
    '*** OR raise the error to the calling procedure
    Err.Raise lErrNum, sErrSource, sErrDesc
    '*** OR ignore the error and continue
    'Resume Next
End Sub

Public Sub CloseDatabase()
    '***********************************************************************************************
    '   Purpose: Closes the database
    '   Created on : 1/9/2001
    '   Created by : Gordon
    '   Modified on :
    '***********************************************************************************************
    
    '-------------------------------
    '*** Generic Procedure Code
    '-------------------------------
    On Error GoTo errHandler
    Dim sModuleAndProcName As String
    sModuleAndProcName = "dbAccess - CloseDatabase Sub"
    
    If m_intAccessMode = http Then
        CloseDatabaseHTTP
    Else
        CloseDatabaseODBC
    End If
    
    Exit Sub
    '------------------------------------------------------------
    '*** Generic Error Handling Code (ensure Exit Proc is above)
    '------------------------------------------------------------
errHandler:
    If Err.Number = 0 Then Exit Sub
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim sErrSource As String
    
    lErrNum = Err.Number
    sErrDesc = Err.Description
    sErrSource = "[" & App.Title & " - " & sModuleAndProcName & " - Build " & App.Major & "." & App.Minor & "." & App.Revision
    If Erl > 0 Then sErrSource = sErrSource & " - Line No " & str(Erl)
    If InStr(1, Err.Source, "-->") > 0 Then
        '*** This error has already been handled by our code
        sErrSource = sErrSource & "]" & vbNewLine & "  --> " & Err.Source
    Else
        '*** Newly generated error, log it here.
        sErrSource = sErrSource & "]" & vbNewLine & "  --> [Source: " & Err.Source & "]"
        On Error Resume Next
        LogError lErrNum, sErrSource, sErrDesc
    End If
    '***    CUSTOM error handler
    m_strErrorCode = CStr(lErrNum)
    m_strErrorDesc = sErrDesc
    m_strErrorSource = sErrSource
    '(ALWAYS comment out two of the three options below)
    '*** EITHER display the error here
    'DisplayError lErrNum, sErrSource, sErrDesc
    '*** OR raise the error to the calling procedure
    Err.Raise lErrNum, sErrSource, sErrDesc
    '*** OR ignore the error and continue
    'Resume Next
End Sub
Private Sub CloseDatabaseODBC()
    '***********************************************************************************************
    '   Purpose: Closes the database for ODBC connection
    '   Created on : 1/9/2001
    '   Created by : Gordon
    '   Modified on :
    '***********************************************************************************************
    
    '-------------------------------
    '*** Generic Procedure Code
    '-------------------------------
    On Error GoTo errHandler
    Dim sModuleAndProcName As String
    sModuleAndProcName = "dbAccess - CloseDatabaseODBC Sub"
    
    If Not m_Cn Is Nothing Then
        If m_Cn.State = adStateOpen Then m_Cn.Close
        Set m_Cn = Nothing
    End If

    Exit Sub
    '------------------------------------------------------------
    '*** Generic Error Handling Code (ensure Exit Proc is above)
    '------------------------------------------------------------
errHandler:
    If Err.Number = 0 Then Exit Sub
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim sErrSource As String
    
    lErrNum = Err.Number
    sErrDesc = Err.Description
    sErrSource = "[" & App.Title & " - " & sModuleAndProcName & " - Build " & App.Major & "." & App.Minor & "." & App.Revision
    If Erl > 0 Then sErrSource = sErrSource & " - Line No " & str(Erl)
    If InStr(1, Err.Source, "-->") > 0 Then
        '*** This error has already been handled by our code
        sErrSource = sErrSource & "]" & vbNewLine & "  --> " & Err.Source
    Else
        '*** Newly generated error, log it here.
        sErrSource = sErrSource & "]" & vbNewLine & "  --> [Source: " & Err.Source & "]"
        On Error Resume Next
        LogError lErrNum, sErrSource, sErrDesc
    End If
    '***    CUSTOM error handler
    m_strErrorCode = CStr(lErrNum)
    m_strErrorDesc = sErrDesc
    m_strErrorSource = sErrSource
    '(ALWAYS comment out two of the three options below)
    '*** EITHER display the error here
    'DisplayError lErrNum, sErrSource, sErrDesc
    '*** OR raise the error to the calling procedure
    Err.Raise lErrNum, sErrSource, sErrDesc
    '*** OR ignore the error and continue
    'Resume Next
End Sub
Private Sub CloseDatabaseHTTP()
    '***********************************************************************************************
    '   Purpose: Closes the database for HTTP connection
    '   Created on : 1/9/2001
    '   Created by : Gordon
    '   Modified on :
    '***********************************************************************************************
    
    '-------------------------------
    '*** Generic Procedure Code
    '-------------------------------
    On Error GoTo errHandler
    Dim sModuleAndProcName As String
    sModuleAndProcName = "dbAccess - CloseDatabaseHTTP Sub"
    Dim rs As ADODB.Recordset
    Dim strURL As String
    
    '***    validate the properties
    ValidateOperation
    
    '***    fetch the data
    strURL = "http://" & m_strURL & "/Common/Server/ExecuteSQL.asp" & _
            "?DSN=" & URLEncode(m_strDSN) & _
            "&SQL=" & URLEncode("CLOSE DATABASE")
'            "&LockMode=" & URLEncode(m_intLockMode) & _
'            "&IsolationLevel=" & URLEncode(m_intIsolationMode)
 
    Set rs = New ADODB.Recordset
    rs.Open strURL
    '***    TODO check the SQL worked!!
    
    Set rs = Nothing
    

    Exit Sub
    '------------------------------------------------------------
    '*** Generic Error Handling Code (ensure Exit Proc is above)
    '------------------------------------------------------------
errHandler:
    If Err.Number = 0 Then Exit Sub
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim sErrSource As String
    
    lErrNum = Err.Number
    sErrDesc = Err.Description
    sErrSource = "[" & App.Title & " - " & sModuleAndProcName & " - Build " & App.Major & "." & App.Minor & "." & App.Revision
    If Erl > 0 Then sErrSource = sErrSource & " - Line No " & str(Erl)
    If InStr(1, Err.Source, "-->") > 0 Then
        '*** This error has already been handled by our code
        sErrSource = sErrSource & "]" & vbNewLine & "  --> " & Err.Source
    Else
        '*** Newly generated error, log it here.
        sErrSource = sErrSource & "]" & vbNewLine & "  --> [Source: " & Err.Source & "]"
        On Error Resume Next
        LogError lErrNum, sErrSource, sErrDesc
    End If
    '***    CUSTOM error handler
    m_strErrorCode = CStr(lErrNum)
    m_strErrorDesc = sErrDesc
    m_strErrorSource = sErrSource
    '(ALWAYS comment out two of the three options below)
    '*** EITHER display the error here
    'DisplayError lErrNum, sErrSource, sErrDesc
    '*** OR raise the error to the calling procedure
    Err.Raise lErrNum, sErrSource, sErrDesc
    '*** OR ignore the error and continue
    'Resume Next
End Sub
Private Sub ConnectODBC()
    '***********************************************************************************************
    '   Purpose: Connects to the database for ODBC access mode
    '   Created on : 1/9/2001
    '   Created by : Gordon
    '   Modified on :
    '***********************************************************************************************
    
    '-------------------------------
    '*** Generic Procedure Code
    '-------------------------------
    On Error GoTo errHandler
    Dim sModuleAndProcName As String
    sModuleAndProcName = "dbAccess - ConnectODBC Sub"
    
    If m_Cn Is Nothing Then

        Set m_Cn = New ADODB.Connection

        m_Cn.Open m_strDSN, m_strUserName  ', "resubew123"
        
        '***    set command timeout
        m_Cn.CommandTimeout = 1000
        
        '***    set lock mode
        SetLockMode m_intLockMode
        
        '***    set isolation mode
        SetIsolationMode m_intIsolationMode
        '***    Default mode for transaction control is not in transaction
        '***    We should use the NotInTransaction = 0 in the enum
        '***    Change it will cause dll benary compatibility broken
        m_intTransMode = NotInTransaction
    End If
    
    Exit Sub
    
    '------------------------------------------------------------
    '*** Generic Error Handling Code (ensure Exit Proc is above)
    '------------------------------------------------------------
errHandler:
    If Err.Number = 0 Then Exit Sub
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim sErrSource As String
    
    lErrNum = Err.Number
    sErrDesc = Err.Description
    sErrSource = "[" & App.Title & " - " & sModuleAndProcName & " - Build " & App.Major & "." & App.Minor & "." & App.Revision
    If Erl > 0 Then sErrSource = sErrSource & " - Line No " & str(Erl)
    If InStr(1, Err.Source, "-->") > 0 Then
        '*** This error has already been handled by our code
        sErrSource = sErrSource & "]" & vbNewLine & "  --> " & Err.Source
    Else
        '*** Newly generated error, log it here.
        sErrSource = sErrSource & "]" & vbNewLine & "  --> [Source: " & Err.Source & "]"
        On Error Resume Next
        LogError lErrNum, sErrSource, sErrDesc
    End If
    '***    CUSTOM error handler
    m_strErrorCode = CStr(lErrNum)
    m_strErrorDesc = sErrDesc
    m_strErrorSource = sErrSource
    '(ALWAYS comment out two of the three options below)
    '*** EITHER display the error here
    'DisplayError lErrNum, sErrSource, sErrDesc
    '*** OR raise the error to the calling procedure
    Err.Raise lErrNum, sErrSource, sErrDesc
    '*** OR ignore the error and continue
    'Resume Next
End Sub
Private Sub ConnectHTTP()
    '***********************************************************************************************
    '   Purpose: Connects to the database for HTTP access mode.
    '           We only use this as a precursor for a transaction.
    '   Created on : 1/9/2001
    '   Created by : Gordon
    '   Modified on :
    '***********************************************************************************************
    
    '-------------------------------
    '*** Generic Procedure Code
    '-------------------------------
    On Error GoTo errHandler
    Dim sModuleAndProcName As String
    sModuleAndProcName = "dbAccess - ConnectHTTP Sub"
    
    Dim rs As ADODB.Recordset
    Dim strURL As String
    
    '***    validate the properties
    ValidateOperation
    
    '***    fetch the data
    strURL = "http://" & m_strURL & "/Common/Server/ExecuteSQL.asp" & _
            "?DSN=" & URLEncode(m_strDSN) & _
            "&SQL=" & URLEncode("CONNECT")
'            "&LockMode=" & URLEncode(m_intLockMode) & _
'            "&IsolationLevel=" & URLEncode(m_intIsolationMode)
'
    Set rs = New ADODB.Recordset
    rs.Open strURL
    
    '***    TODO check the SQL worked!!
    
    Set rs = Nothing
    
    Exit Sub
    
    '------------------------------------------------------------
    '*** Generic Error Handling Code (ensure Exit Proc is above)
    '------------------------------------------------------------
errHandler:
    If Err.Number = 0 Then Exit Sub
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim sErrSource As String
    
    lErrNum = Err.Number
    sErrDesc = Err.Description
    sErrSource = "[" & App.Title & " - " & sModuleAndProcName & " - Build " & App.Major & "." & App.Minor & "." & App.Revision
    If Erl > 0 Then sErrSource = sErrSource & " - Line No " & str(Erl)
    If InStr(1, Err.Source, "-->") > 0 Then
        '*** This error has already been handled by our code
        sErrSource = sErrSource & "]" & vbNewLine & "  --> " & Err.Source
    Else
        '*** Newly generated error, log it here.
        sErrSource = sErrSource & "]" & vbNewLine & "  --> [Source: " & Err.Source & "]"
        On Error Resume Next
        LogError lErrNum, sErrSource, sErrDesc
    End If
    '***    CUSTOM error handler
    m_strErrorCode = CStr(lErrNum)
    m_strErrorDesc = sErrDesc
    m_strErrorSource = sErrSource
    '(ALWAYS comment out two of the three options below)
    '*** EITHER display the error here
    'DisplayError lErrNum, sErrSource, sErrDesc
    '*** OR raise the error to the calling procedure
    Err.Raise lErrNum, sErrSource, sErrDesc
    '*** OR ignore the error and continue
    'Resume Next
End Sub

Private Sub ValidateOperation()
    '***********************************************************************************************
    '   Purpose: Validates the operation requested. Checks if sufficient properties have been
    '           set etc.
    '   Created on : 1/9/2001
    '   Created by : Gordon
    '   Modified on :
    '***********************************************************************************************
    
    '-------------------------------
    '*** Generic Procedure Code
    '-------------------------------
    On Error GoTo errHandler
    Dim sModuleAndProcName As String
    sModuleAndProcName = "dbAccess - ValidateOperation sub"
    
    '***    for ODBC check User Name and Password.
    If m_intAccessMode = ODBC Then
        If m_strUserName = "" Then
            Err.Raise -100, sModuleAndProcName, "User name has not been set."
        End If
        If m_strPassword = "" Then
            Err.Raise -101, sModuleAndProcName, "User password has not been set."
        End If
    Else
        If m_strURL = "" Then
            Err.Raise -102, sModuleAndProcName, "URL has not been set."
        End If
    End If
    '***    check DSN
    If m_strDSN = "" Then
        Err.Raise -103, sModuleAndProcName, "DSN has not been set."
    End If
    '***    check serial column and table
    If m_strSerialCol <> "" And m_strSerialTable = "" Then
        Err.Raise -104, sModuleAndProcName, "Serial column has been set but no table has been set."
    End If
    
    Exit Sub
    '------------------------------------------------------------
    '*** Generic Error Handling Code (ensure Exit Proc is above)
    '------------------------------------------------------------
errHandler:
    If Err.Number = 0 Then Exit Sub
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim sErrSource As String
    
    lErrNum = Err.Number
    sErrDesc = Err.Description
    sErrSource = "[" & App.Title & " - " & sModuleAndProcName & " - Build " & App.Major & "." & App.Minor & "." & App.Revision
    If Erl > 0 Then sErrSource = sErrSource & " - Line No " & str(Erl)
    If InStr(1, Err.Source, "-->") > 0 Then
        '*** This error has already been handled by our code
        sErrSource = sErrSource & "]" & vbNewLine & "  --> " & Err.Source
    Else
        '*** Newly generated error, log it here.
        sErrSource = sErrSource & "]" & vbNewLine & "  --> [Source: " & Err.Source & "]"
        On Error Resume Next
        LogError lErrNum, sErrSource, sErrDesc
    End If
    '***    CUSTOM error handler
    m_strErrorCode = CStr(lErrNum)
    m_strErrorDesc = sErrDesc
    m_strErrorSource = sErrSource
    '(ALWAYS comment out two of the three options below)
    '*** EITHER display the error here
    'DisplayError lErrNum, sErrSource, sErrDesc
    '*** OR raise the error to the calling procedure
    Err.Raise lErrNum, sErrSource, sErrDesc
    '*** OR ignore the error and continue
    'Resume Next
End Sub
Private Function URLEncode(ByVal strValue As String) As String
'***********************************************************************************************
' Purpose: Encodes strings using the x-www-form-urlencoded format
' AZaz not encoded
' 0-9 not encoded
' [SPACE] becomes +
' Everything else is encoded as %xy where xy is
' the hexadecimal value of the ascii value of the letter
'Created on : 23/11/2000
'Created by : Darren
'Modified on :
'***********************************************************************************************

'-------------------------------
'*** Generic Procedure Code
'-------------------------------
1        On Error GoTo errHandler
2        Dim sModuleAndProcName As String
3        sModuleAndProcName = "OESalesAdmin basGlobal - URLEncode Function"
     
    '-------------------------------
    '*** Specific Procedure Code
    '-------------------------------

4     Dim lngCount As Long
5     Dim strLetter As String
6     Dim strOutput As String
7     Dim strHexValue As String
  
      ' Loop through every character(letter) of the string
8     For lngCount = 1 To Len(strValue)
  
        ' Get the letter to check
9       strLetter = Mid(strValue, lngCount, 1)
        ' Check to see if we have to encode the letter
10      If strLetter = " " Then
          ' Replace spaces with plus signs
11        strOutput = strOutput & "+"
12      ElseIf IsLetterAlphaNumeric(strLetter) = False Then
          ' Encode the string
13        strHexValue = Hex(Asc(strLetter))
          ' Pad the Hex value to a length of two characters
14        If Len(strHexValue) = 1 Then strHexValue = "0" & strHexValue
          ' Add the HexValue to the string
15        strOutput = strOutput & "%" & strHexValue
16      Else
          ' Do not need to encode alphanumeric characters
17        strOutput = strOutput & strLetter
18      End If
    
19    Next
  
20    URLEncode = strOutput
  
    '------------------------------------------------------------
    '*** Generic Error Handling Code (ensure Exit Proc is above)
    '------------------------------------------------------------
errHandler:
  If Err.Number = 0 Then Exit Function
  Dim lErrNum As Long, sErrDesc As String, sErrSource As String
  lErrNum = Err.Number
  sErrDesc = Err.Description
  sErrSource = "[" & App.Title & " - " & sModuleAndProcName & " - Build " & App.Major & "." & App.Minor & "." & App.Revision
  If Erl > 0 Then sErrSource = sErrSource & " - Line No " & str(Erl)
  If InStr(1, Err.Source, "-->") > 0 Then
    '*** This error has already been handled by our code
    sErrSource = sErrSource & "]" & vbNewLine & "  --> " & Err.Source
  Else
    '*** Newly generated error, log it here.
    sErrSource = sErrSource & "]" & vbNewLine & "  --> [Source: " & Err.Source & "]"
    On Error Resume Next
    LogError lErrNum, sErrSource, sErrDesc
  End If
  '***    CUSTOM error handler
    m_strErrorCode = CStr(lErrNum)
    m_strErrorDesc = sErrDesc
    m_strErrorSource = sErrSource
  '(ALWAYS comment out two of the three options below)
  '*** EITHER display the error here
    'DisplayError lErrNum, sErrSource, sErrDesc
  '*** OR raise the error to the calling procedure
    Err.Raise lErrNum, sErrSource, sErrDesc
  '*** OR ignore the error and continue
    'Resume Next

End Function

Private Function URLDecode(ByVal strValue As String) As String
'-------------------------------
'*** Generic Procedure Code
'-------------------------------
1        On Error GoTo errHandler
2        Dim sModuleAndProcName As String
3        sModuleAndProcName = "OESalesAdmin basGlobal - URLDecode Function"
     
    '-------------------------------
    '*** Specific Procedure Code
    '-------------------------------

4     Dim nPos As Integer
5     Dim strBefore As String
6     Dim strAfter As String
7     Dim strLetter As String
8     Dim strTemp As String
9     Dim lngASCII As Integer
  
      ' Replace all "+" signs
10    strTemp = Replace(strValue, "+", " ")
  
      ' Replace all occurences of a hex representation with its ascii value
11    nPos = InStr(1, strTemp, "%")
12    Do While nPos <> 0
  
        ' Each occurence is of the form %XY where XY is the Hex Value of the letter
13      strBefore = Mid(strTemp, 1, nPos - 1)
14      strAfter = Mid(strTemp, nPos + 1 + 2) ' after the %XY part
        
        ' Get the two letters that comprise the Hex Value
15      strLetter = Mid(strTemp, nPos + 1, 2)
    
        ' Now convert the Hex value to a number (this is the ASCII value of the encoded letter)
16      lngASCII = CLng("&h" & strLetter)
    
        ' Get the Unencoded value of the Letter
17      strLetter = Chr(lngASCII)
    
18      strTemp = strBefore & strLetter & strAfter
19      nPos = InStr(1, strTemp, "%")
20    Loop
  
21    URLDecode = strTemp
  
    '------------------------------------------------------------
    '*** Generic Error Handling Code (ensure Exit Proc is above)
    '------------------------------------------------------------
errHandler:
  If Err.Number = 0 Then Exit Function
  Dim lErrNum As Long, sErrDesc As String, sErrSource As String
  lErrNum = Err.Number
  sErrDesc = Err.Description
  sErrSource = "[" & App.Title & " - " & sModuleAndProcName & " - Build " & App.Major & "." & App.Minor & "." & App.Revision
  If Erl > 0 Then sErrSource = sErrSource & " - Line No " & str(Erl)
  If InStr(1, Err.Source, "-->") > 0 Then
    '*** This error has already been handled by our code
    sErrSource = sErrSource & "]" & vbNewLine & "  --> " & Err.Source
  Else
    '*** Newly generated error, log it here.
    sErrSource = sErrSource & "]" & vbNewLine & "  --> [Source: " & Err.Source & "]"
    On Error Resume Next
    LogError lErrNum, sErrSource, sErrDesc
  End If
  '***    CUSTOM error handler
    m_strErrorCode = CStr(lErrNum)
    m_strErrorDesc = sErrDesc
    m_strErrorSource = sErrSource
  '(ALWAYS comment out two of the three options below)
  '*** EITHER display the error here
    'DisplayError lErrNum, sErrSource, sErrDesc
  '*** OR raise the error to the calling procedure
    Err.Raise lErrNum, sErrSource, sErrDesc
  '*** OR ignore the error and continue
    'Resume Next
  
End Function

' a-zA-Z True
' 0-9 True
' Everything else is False
Private Function IsLetterAlphaNumeric(ByVal strLetter As String) As Boolean
'-------------------------------
'*** Generic Procedure Code
'-------------------------------
1        On Error GoTo errHandler
2        Dim sModuleAndProcName As String
3        sModuleAndProcName = "OESalesAdmin basGlobal - IsLetterAlphaNumeric Function"
     
    '-------------------------------
    '*** Specific Procedure Code
    '-------------------------------
  
      ' Check length of the passed string
4     If Len(strLetter) > 1 Then
5       IsLetterAlphaNumeric = False
6       Exit Function
7     End If
  
      ' IsNumeric
8     If IsNumeric(strLetter) Then
9       IsLetterAlphaNumeric = True
10      Exit Function
11    End If
  
      ' Is a letter?
12    If (Asc(strLetter) >= 65 And Asc(strLetter) <= 90) Or _
    (Asc(strLetter) >= 97 And Asc(strLetter) <= 122) Then
    
13      IsLetterAlphaNumeric = True
14      Exit Function
    
15    End If
    
16    IsLetterAlphaNumeric = False
  
    '------------------------------------------------------------
    '*** Generic Error Handling Code (ensure Exit Proc is above)
    '------------------------------------------------------------
errHandler:
  If Err.Number = 0 Then Exit Function
  Dim lErrNum As Long, sErrDesc As String, sErrSource As String
  lErrNum = Err.Number
  sErrDesc = Err.Description
  sErrSource = "[" & App.Title & " - " & sModuleAndProcName & " - Build " & App.Major & "." & App.Minor & "." & App.Revision
  If Erl > 0 Then sErrSource = sErrSource & " - Line No " & str(Erl)
  If InStr(1, Err.Source, "-->") > 0 Then
    '*** This error has already been handled by our code
    sErrSource = sErrSource & "]" & vbNewLine & "  --> " & Err.Source
  Else
    '*** Newly generated error, log it here.
    sErrSource = sErrSource & "]" & vbNewLine & "  --> [Source: " & Err.Source & "]"
    On Error Resume Next
    LogError lErrNum, sErrSource, sErrDesc
  End If
  '***    CUSTOM error handler
    m_strErrorCode = CStr(lErrNum)
    m_strErrorDesc = sErrDesc
    m_strErrorSource = sErrSource
  '(ALWAYS comment out two of the three options below)
  '*** EITHER display the error here
    'DisplayError lErrNum, sErrSource, sErrDesc
  '*** OR raise the error to the calling procedure
    Err.Raise lErrNum, sErrSource, sErrDesc
  '*** OR ignore the error and continue
    'Resume Next

End Function

Private Function FieldToString(objRS As ADODB.Recordset, ByVal varIndex As Variant) _
                    As String

  Dim varValue As Variant
  
  varValue = objRS(varIndex)

  If Not IsNull(varValue) Then
    FieldToString = Trim$(CStr(varValue))
  Else
    FieldToString = ""
  End If
    
End Function

Private Sub Class_Initialize()
    m_intLockMode = NotWait
    m_intIsolationMode = CommittedRead
    m_intAccessMode = ODBC
    m_lngSerialNumber = 0
End Sub

Private Sub Class_Terminate()
    Set m_Cn = Nothing
End Sub
