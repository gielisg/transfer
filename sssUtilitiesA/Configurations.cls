VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Configurations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Base 1

Public Type ComponentConfiguration
    BusinessUnitCode As String
    EntityNarr As String
    FileName As String
End Type

Private m_objUser As CUser
'***    array that stores the various components (dlls, assemblies in .NET) that contain
'***    the functionality for the various entities in the system.
'Private ma_Components() As String
'Private m_intArrayCount As Integer
'Private m_intArraySize As Integer
Private m_col As Collection

Property Let User(objUser As CUser)
    Set m_objUser = objUser
End Property
Public Sub LoadComponents()
    '***********************************************************************************************
    'Purpose    : Loads the components for an business unit
    'Created by: Gordon
    '***********************************************************************************************

    '-------------------------------
    '*** Generic Procedure Code
    '-------------------------------
    On Error GoTo ERROR_HANDLER
    
    Dim strModuleAndProcName As String
    strModuleAndProcName = "Configurations - LoadComponents Sub"
    Dim rs As ADODB.Recordset
    Dim dbAccess As dbAccess
    Dim strSQL As String
    Dim intCounter As Integer
    Dim varBusinessUnitCodes As Variant
    Dim idx As Integer
    
    Set m_col = New Collection
    
    '***    some validations
    If m_objUser Is Nothing Then Err.Raise vbObjectError + 100, strModuleAndProcName, "User not set."
    
    GoTo EXIT_PROC
    
    '***    fetch the components for business units
    Set dbAccess = m_objUser.GetDbAccessObject(m_objUser.DefaultBusinessUnitCode)
    GoTo EXIT_PROC
    strSQL = "select bus_unit_code, entity_narr, file_name " & _
                "from components_cnfg a, entities_cnfg b, ent_comp_bu_rel c " & _
                "where a.component_id = c.component_id " & _
                "and b.entity_id = c.entity_id "
    
    Set rs = dbAccess.LoadRS(strSQL)

    
    While Not rs.EOF

        Dim ComponentConfigA As ComponentConfiguration
        ComponentConfigA.BusinessUnitCode = FieldToString(rs, "bus_unit_code")
        ComponentConfigA.EntityNarr = FieldToString(rs, "entity_narr")
        ComponentConfigA.FileName = FieldToString(rs, "file_name")
        m_col.Add ComponentConfigA
        
        rs.MoveNext
    Wend
    strSQL = "select entity_narr, file_name " & _
                "from components_cnfg a, entities_cnfg b, ent_comp_rel c " & _
                "where a.component_id = c.component_id " & _
                "and b.entity_id = c.entity_id "
    
    Set rs = dbAccess.LoadRS(strSQL)
    
    '***    get the array of business units available to this user
    varBusinessUnitCodes = m_objUser.BusinessUnits.Keys
    
    intCounter = m_objUser.BusinessUnits.Count
    While Not rs.EOF
        For idx = 1 To intCounter
            
            '***    set the data if there is not already a record
            If Exists(varBusinessUnitCodes(idx - 1), FieldToString(rs, "entity_narr")) Then
                GoTo SKIP_RECORD
            End If
            
            Dim ComponentConfigB As ComponentConfiguration
            ComponentConfigB.BusinessUnitCode = varBusinessUnitCodes(idx - 1)
            ComponentConfigB.EntityNarr = FieldToString(rs, "entity_narr")
            ComponentConfigB.FileName = FieldToString(rs, "file_name")
            m_col.Add ComponentConfigB
            
SKIP_RECORD:
            
        Next
        rs.MoveNext
    Wend
    
EXIT_PROC:
    
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
Public Function Exists(ByVal strBusinessUnitCode As String, ByVal strEntity As String) As Boolean
    '***********************************************************************************************
    'Purpose    : Indicates whether a value exists for a business unit/entity
    'Created by: Gordon
    '***********************************************************************************************

    '-------------------------------
    '*** Generic Procedure Code
    '-------------------------------
    On Error GoTo ERROR_HANDLER
    
    Dim strModuleAndProcName As String
    strModuleAndProcName = "Configurations - Exists Function"
    Dim blnFound As Boolean
    Dim idx As Integer
    Dim ComponentConfig As ComponentConfiguration
    
    '***    scroll through the array
    For idx = 1 To m_col.Count
        ComponentConfig = m_col(idx)
        If ComponentConfig.BusinessUnitCode = strBusinessUnitCode _
        And ComponentConfig.EntityNarr = strEntity _
        Then
            blnFound = True
            Exit For
        End If
    Next
    
    Exists = blnFound
    
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
Public Function GetComponent(ByVal strEntity As String, Optional ByVal strBusinessUnitCode As String) As String
    '***********************************************************************************************
    'Purpose    : Get the component name for a business unit/entity
    'Created by: Gordon
    '***********************************************************************************************

    '-------------------------------
    '*** Generic Procedure Code
    '-------------------------------
    On Error GoTo ERROR_HANDLER
    
    Dim strModuleAndProcName As String
    strModuleAndProcName = "Configurations - GetComponent Function"
    Dim blnFound As Boolean
    Dim idx As Integer
    Dim ComponentConfig As ComponentConfiguration
    
    '***    use the current business unit code if none passed
    If strBusinessUnitCode = "" Then strBusinessUnitCode = m_objUser.CurrentBusinessUnitCode
    
    '***    scroll through the array
    For idx = 1 To m_col.Count
        ComponentConfig = m_col(idx)
        If ComponentConfig.BusinessUnitCode = strBusinessUnitCode _
        And ComponentConfig.EntityNarr = strEntity _
        Then
            GetComponent = ComponentConfig.FileName
        End If
    Next
    
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


