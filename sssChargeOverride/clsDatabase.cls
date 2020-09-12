VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public objApplication As Application
Dim strSQL As String
Dim colParameters As Collection

Public Function GetChargeOverride(blnOther As Boolean) As Recordset
'*************************************************************************************
'Purpose    : Get Account charge override
'Inputs     :
'Outputs    :
'Version: $SSSVersion:
'Created on : 13/03/2001
'Created by : Patrick
'*************************************************************************************

'-------------------------------
'*** Generic Procedure Code
'-------------------------------
    On Error GoTo errHandler
    Dim sModuleAndProcName As String
    sModuleAndProcName = "clsDatabase - GetChargeOverride Procedure"
    
'-------------------------------
'*** Specific Procedure Code
'-------------------------------
        
    Set colParameters = New Collection
    
    strSQL = "execute procedure ss_chgoverride(p_debtor_code = ?"
    colParameters.Add objApplication.ContactCode
    
    If objApplication.Mode = typService Then
        strSQL = strSQL & ", p_sp_cn_ref = ?"
        colParameters.Add objApplication.SpCnRef
    End If
    
    If Not blnOther Then
        If objApplication.Mode = typContact Then
            strSQL = strSQL & ", p_debtor_only = 'Y'"
        Else
            strSQL = strSQL & ", p_service_only = 'Y'"
        End If
    End If
        
    strSQL = strSQL & ", resume_flg = 'Y')"
    
    objApplication.dbAccess.Parameters = colParameters
    
    Set GetChargeOverride = objApplication.dbAccess.loadrs(strSQL)
       
Exit Function
'------------------------------------------------------------
'*** Generic Error Handling Code (ensure Exit Proc is above)
'------------------------------------------------------------
errHandler:
    If Err.Number = 0 Then Exit Function
    
    Dim lErrNum As Long, sErrDesc As String, sErrSource As String
    lErrNum = Err.Number
    sErrDesc = Err.Description
    sErrSource = "[" & App.Title & " - " & sModuleAndProcName & " - Build " & App.Major & "." & App.Minor & "." & App.Revision
    If Erl > 0 Then sErrSource = sErrSource & " - Line No " & CStr(Erl)
    If InStr(1, Err.Source, "-->") > 0 Then
        '*** This error has already been handled by our code
        sErrSource = sErrSource & "]" & vbNewLine & "  --> " & Err.Source
    Else
        '*** Newly generated error, log it here.
        sErrSource = sErrSource & "]" & vbNewLine & "  --> [Source: " & Err.Source & "]"
        On Error Resume Next
        LogError lErrNum, sErrSource, sErrDesc
    End If
    '(ALWAYS comment out two of the three options below)
    '*** EITHER display the error here
    'DisplayError lErrNum, sErrSource, sErrDesc
    '*** OR raise the error to the calling procedure
    Err.Raise lErrNum, sErrSource, sErrDesc
    '*** OR ignore the error and continue
    'Resume Next
End Function

Public Function GetAccountPackage(SpCnRef As Long) As Recordset
'*************************************************************************************
'Purpose    : Get Account package associated
'Inputs     :
'Outputs    :
'Version: $SSSVersion:
'Created on : 13/03/2001
'Created by : Patrick
'*************************************************************************************

'-------------------------------
'*** Generic Procedure Code
'-------------------------------
    On Error GoTo errHandler
    Dim sModuleAndProcName As String
    sModuleAndProcName = "clsDatabase - GetAccountPackage Procedure"
    
'-------------------------------
'*** Specific Procedure Code
'-------------------------------
        
    Set colParameters = New Collection
    
    strSQL = "execute procedure ss_account_package(p_debtor_code = ?"
    colParameters.Add objApplication.ContactCode
    
    If objApplication.Mode = typService Then
        strSQL = strSQL & ", p_sp_cn_ref = ?"
        colParameters.Add objApplication.SpCnRef
    ElseIf SpCnRef > 0 Then
        strSQL = strSQL & ", p_sp_cn_ref = "
        colParameters.Add SpCnRef
    End If
    
    strSQL = strSQL & ", resume_flg = 'Y')"
    
    objApplication.dbAccess.Parameters = colParameters
    
    Set GetAccountPackage = objApplication.dbAccess.loadrs(strSQL)
       
Exit Function
'------------------------------------------------------------
'*** Generic Error Handling Code (ensure Exit Proc is above)
'------------------------------------------------------------
errHandler:
    If Err.Number = 0 Then Exit Function
    
    Dim lErrNum As Long, sErrDesc As String, sErrSource As String
    lErrNum = Err.Number
    sErrDesc = Err.Description
    sErrSource = "[" & App.Title & " - " & sModuleAndProcName & " - Build " & App.Major & "." & App.Minor & "." & App.Revision
    If Erl > 0 Then sErrSource = sErrSource & " - Line No " & CStr(Erl)
    If InStr(1, Err.Source, "-->") > 0 Then
        '*** This error has already been handled by our code
        sErrSource = sErrSource & "]" & vbNewLine & "  --> " & Err.Source
    Else
        '*** Newly generated error, log it here.
        sErrSource = sErrSource & "]" & vbNewLine & "  --> [Source: " & Err.Source & "]"
        On Error Resume Next
        LogError lErrNum, sErrSource, sErrDesc
    End If
    '(ALWAYS comment out two of the three options below)
    '*** EITHER display the error here
    'DisplayError lErrNum, sErrSource, sErrDesc
    '*** OR raise the error to the calling procedure
    Err.Raise lErrNum, sErrSource, sErrDesc
    '*** OR ignore the error and continue
    'Resume Next
End Function

'Public Function GetAvailChargesTree() As ADODB.Recordset
''***********************************************************************************************
''Purpose    : Get charges tree for a connection
''Inputs     :
''Outputs    :
''Version: $SSSVersion:
''Created on : 13/03/2001
''Created by : Patrick
''***********************************************************************************************
'
''-------------------------------
''*** Generic Procedure Code
''-------------------------------
'    On Error GoTo errHandler
'    Dim sModuleAndProcName As String
'    sModuleAndProcName = "clsDatabase - GetAvailChargesTree Procedure"
'
''-------------------------------
''*** Specific Procedure Code
''-------------------------------
'
'    Set colParameters = New Collection
'
'    strSQL = "execute procedure ss_avail_chgtree(resume_flg = 'Y')"
'
'    objApplication.dbAccess.Parameters = colParameters
'
'    Set GetAvailChargesTree = objApplication.dbAccess.LoadRS(strSQL)
'
'Exit Function
''------------------------------------------------------------
''*** Generic Error Handling Code (ensure Exit Proc is above)
''------------------------------------------------------------
'errHandler:
'    If Err.Number = 0 Then Exit Function
'
'    Dim lErrNum As Long, sErrDesc As String, sErrSource As String
'    lErrNum = Err.Number
'    sErrDesc = Err.Description
'    sErrSource = "[" & App.Title & " - " & sModuleAndProcName & " - Build " & App.Major & "." & App.Minor & "." & App.Revision
'    If Erl > 0 Then sErrSource = sErrSource & " - Line No " & CStr(Erl)
'    If InStr(1, Err.Source, "-->") > 0 Then
'        '*** This error has already been handled by our code
'        sErrSource = sErrSource & "]" & vbNewLine & "  --> " & Err.Source
'    Else
'        '*** Newly generated error, log it here.
'        sErrSource = sErrSource & "]" & vbNewLine & "  --> [Source: " & Err.Source & "]"
'        On Error Resume Next
'        LogError lErrNum, sErrSource, sErrDesc
'    End If
'    '(ALWAYS comment out two of the three options below)
'    '*** EITHER display the error here
'    'DisplayError lErrNum, sErrSource, sErrDesc
'    '*** OR raise the error to the calling procedure
'    Err.Raise lErrNum, sErrSource, sErrDesc
'    '*** OR ignore the error and continue
'    'Resume Next
'End Function
'
Public Function GetChargeLastUsed(ChargeCode As String) As ADODB.Recordset
'***********************************************************************************************
'Purpose    : name
'Inputs     :
'Outputs    :
'Version: $SSSVersion:
'Created on : 13/03/2001
'Created by : Patrick
'***********************************************************************************************

'-------------------------------
'*** Generic Procedure Code
'-------------------------------
    On Error GoTo errHandler
    Dim sModuleAndProcName As String
    sModuleAndProcName = "clsDatabase - GetChargeLastUsed Procedure"
    
'-------------------------------
'*** Specific Procedure Code
'-------------------------------
    
    Set colParameters = New Collection
    
    strSQL = "execute procedure ss_cont_chg_last(p_debtor_code = ?, p_chg_code = ?, resume_flg = 'Y')"
    colParameters.Add objApplication.ContactCode
    colParameters.Add ChargeCode
    
    objApplication.dbAccess.Parameters = colParameters
    
    Set GetChargeLastUsed = objApplication.dbAccess.loadrs(strSQL)
    
Exit Function
'------------------------------------------------------------
'*** Generic Error Handling Code (ensure Exit Proc is above)
'------------------------------------------------------------
errHandler:
    If Err.Number = 0 Then Exit Function
    
    Dim lErrNum As Long, sErrDesc As String, sErrSource As String
    lErrNum = Err.Number
    sErrDesc = Err.Description
    sErrSource = "[" & App.Title & " - " & sModuleAndProcName & " - Build " & App.Major & "." & App.Minor & "." & App.Revision
    If Erl > 0 Then sErrSource = sErrSource & " - Line No " & CStr(Erl)
    If InStr(1, Err.Source, "-->") > 0 Then
        '*** This error has already been handled by our code
        sErrSource = sErrSource & "]" & vbNewLine & "  --> " & Err.Source
    Else
        '*** Newly generated error, log it here.
        sErrSource = sErrSource & "]" & vbNewLine & "  --> [Source: " & Err.Source & "]"
        On Error Resume Next
        LogError lErrNum, sErrSource, sErrDesc
    End If
    '(ALWAYS comment out two of the three options below)
    '*** EITHER display the error here
    'DisplayError lErrNum, sErrSource, sErrDesc
    '*** OR raise the error to the calling procedure
    Err.Raise lErrNum, sErrSource, sErrDesc
    '*** OR ignore the error and continue
    'Resume Next


End Function

Public Function GetSaveChargeOverride(PackageCode As Long, PackageOption As Long, _
    ChargeCode As String, StartDate As Date, EndDate As Date, _
    ChargeAmount As Double, Markup As Double, SpCnRef As Long, BillNarrative As String) As Recordset
'*************************************************************************************
'Purpose    : Save charge override
'Inputs     :
'Outputs    :
'Version: $SSSVersion:
'Created on : 13/03/2001
'Created by : Patrick
'*************************************************************************************

'-------------------------------
'*** Generic Procedure Code
'-------------------------------
    On Error GoTo errHandler
    Dim sModuleAndProcName As String
    sModuleAndProcName = "clsDatabase - GetSaveChargeOverride Procedure"
    
'-------------------------------
'*** Specific Procedure Code
'-------------------------------
        
    Set colParameters = New Collection
    
    strSQL = "execute procedure sp_new_chgoveride (p_debtor_code = ?, p_pkg_code  = ?"
    colParameters.Add objApplication.ContactCode
    colParameters.Add PackageCode
    
    If objApplication.Mode = typService Then
        strSQL = strSQL & ", p_sp_cn_ref = ?"
        colParameters.Add objApplication.SpCnRef
    ElseIf SpCnRef <> 0 Then
        strSQL = strSQL & ", p_sp_cn_ref = ?"
        colParameters.Add SpCnRef
    End If
    
    If PackageOption > 0 Then
        strSQL = strSQL & ", p_option = ?"
        colParameters.Add PackageOption
    End If
    
    strSQL = strSQL & ", p_chg_code = ?, p_chgoverridewefr = ?, p_chgoverrideweto = ?"
        
    colParameters.Add ChargeCode
    colParameters.Add szConvertDateToInformix(StartDate, False)
    colParameters.Add szConvertDateToInformix(EndDate, False)
    
    If ChargeAmount <> 0 Then
        strSQL = strSQL & ", p_chg_amount = ?"
        colParameters.Add ChargeAmount
    End If

    If Markup <> 0 Then
        strSQL = strSQL & ", p_nt_cost_comp = ?"
        colParameters.Add Markup
    End If
    
    If BillNarrative <> "" Then
        strSQL = strSQL & ", p_ChgOverrideBNarr = ?"
        colParameters.Add BillNarrative
    End If
    
    strSQL = strSQL & ", resume_flg = 'Y', tran_flg = 'Y')"
    
    objApplication.dbAccess.Parameters = colParameters
    
    Set GetSaveChargeOverride = objApplication.dbAccess.loadrs(strSQL)
       
Exit Function
'------------------------------------------------------------
'*** Generic Error Handling Code (ensure Exit Proc is above)
'------------------------------------------------------------
errHandler:
    If Err.Number = 0 Then Exit Function
    
    Dim lErrNum As Long, sErrDesc As String, sErrSource As String
    lErrNum = Err.Number
    sErrDesc = Err.Description
    sErrSource = "[" & App.Title & " - " & sModuleAndProcName & " - Build " & App.Major & "." & App.Minor & "." & App.Revision
    If Erl > 0 Then sErrSource = sErrSource & " - Line No " & CStr(Erl)
    If InStr(1, Err.Source, "-->") > 0 Then
        '*** This error has already been handled by our code
        sErrSource = sErrSource & "]" & vbNewLine & "  --> " & Err.Source
    Else
        '*** Newly generated error, log it here.
        sErrSource = sErrSource & "]" & vbNewLine & "  --> [Source: " & Err.Source & "]"
        On Error Resume Next
        LogError lErrNum, sErrSource, sErrDesc
    End If
    '(ALWAYS comment out two of the three options below)
    '*** EITHER display the error here
    'DisplayError lErrNum, sErrSource, sErrDesc
    '*** OR raise the error to the calling procedure
    Err.Raise lErrNum, sErrSource, sErrDesc
    '*** OR ignore the error and continue
    'Resume Next
End Function

Public Function EndChargeOverride(ChargeOverrideID As Long, EndDate As Date) As Boolean
'*************************************************************************************
'Purpose    :
'Inputs     :
'Outputs    :
'Version: $SSSVersion:
'Created on : 13/03/2001
'Created by : Patrick
'*************************************************************************************

'-------------------------------
'*** Generic Procedure Code
'-------------------------------
    On Error GoTo errHandler
    Dim sModuleAndProcName As String
    sModuleAndProcName = "clsDatabase - EndChargeOverride Procedure"
    
'-------------------------------
'*** Specific Procedure Code
'-------------------------------
            
    Set colParameters = New Collection
    
    strSQL = "update chgoverride set chgoverrideweto = ? " & _
        "where chgoverrideid = ?"
    
    colParameters.Add szConvertDateToInformix(EndDate, False)
    colParameters.Add ChargeOverrideID
    
    objApplication.dbAccess.Parameters = colParameters
    objApplication.dbAccess.Executesql strSQL
    
    EndChargeOverride = True
Exit Function
'------------------------------------------------------------
'*** Generic Error Handling Code (ensure Exit Proc is above)
'------------------------------------------------------------
errHandler:
    If Err.Number = 0 Then Exit Function
    
    Dim lErrNum As Long, sErrDesc As String, sErrSource As String
    lErrNum = Err.Number
    sErrDesc = Err.Description
    sErrSource = "[" & App.Title & " - " & sModuleAndProcName & " - Build " & App.Major & "." & App.Minor & "." & App.Revision
    If Erl > 0 Then sErrSource = sErrSource & " - Line No " & CStr(Erl)
    If InStr(1, Err.Source, "-->") > 0 Then
        '*** This error has already been handled by our code
        sErrSource = sErrSource & "]" & vbNewLine & "  --> " & Err.Source
    Else
        '*** Newly generated error, log it here.
        sErrSource = sErrSource & "]" & vbNewLine & "  --> [Source: " & Err.Source & "]"
        On Error Resume Next
        LogError lErrNum, sErrSource, sErrDesc
    End If
    '(ALWAYS comment out two of the three options below)
    '*** EITHER display the error here
    'DisplayError lErrNum, sErrSource, sErrDesc
    '*** OR raise the error to the calling procedure
    Err.Raise lErrNum, sErrSource, sErrDesc
    '*** OR ignore the error and continue
    'Resume Next
End Function

Public Function GetUpdateChargeOverride(ChargeOverrideID As Long, StartDate As Date, _
    EndDate As Date, ChargeAmount As Double, Markup As Double, BillNarrative As String) As Recordset
'*************************************************************************************
'Purpose    : name
'Inputs     :
'Outputs    :
'Version: $SSSVersion:
'Created on : 13/03/2001
'Created by : Patrick
'*************************************************************************************

'-------------------------------
'*** Generic Procedure Code
'-------------------------------
    On Error GoTo errHandler
    Dim sModuleAndProcName As String
    sModuleAndProcName = "clsDatabase - GetUpdateChargeOverride Procedure"
    
'-------------------------------
'*** Specific Procedure Code
'-------------------------------
        
    Set colParameters = New Collection
    
    strSQL = "execute procedure sp_upd_chgoveride(p_chgoverrideid = ?" & _
        ", p_chgoverridewefr = ?, p_chgoverrideweto = ?, p_chg_amount = ?"
        
    colParameters.Add ChargeOverrideID
    colParameters.Add szConvertDateToInformix(StartDate, False)
    colParameters.Add szConvertDateToInformix(EndDate, False)
    colParameters.Add ChargeAmount
    
    If Markup <> 0 Then
        strSQL = strSQL & ", p_nt_cost_comp = ?"
        colParameters.Add Markup
    End If
    
    If BillNarrative <> "" Then
        strSQL = strSQL & ", p_ChgOverrideBNarr = ?"
        colParameters.Add BillNarrative
    End If
    
    strSQL = strSQL & ", resume_flg = 'Y', tran_flg = 'Y')"
    
    objApplication.dbAccess.Parameters = colParameters
    
    Set GetUpdateChargeOverride = objApplication.dbAccess.loadrs(strSQL)
       
Exit Function
'------------------------------------------------------------
'*** Generic Error Handling Code (ensure Exit Proc is above)
'------------------------------------------------------------
errHandler:
    If Err.Number = 0 Then Exit Function
    
    Dim lErrNum As Long, sErrDesc As String, sErrSource As String
    lErrNum = Err.Number
    sErrDesc = Err.Description
    sErrSource = "[" & App.Title & " - " & sModuleAndProcName & " - Build " & App.Major & "." & App.Minor & "." & App.Revision
    If Erl > 0 Then sErrSource = sErrSource & " - Line No " & CStr(Erl)
    If InStr(1, Err.Source, "-->") > 0 Then
        '*** This error has already been handled by our code
        sErrSource = sErrSource & "]" & vbNewLine & "  --> " & Err.Source
    Else
        '*** Newly generated error, log it here.
        sErrSource = sErrSource & "]" & vbNewLine & "  --> [Source: " & Err.Source & "]"
        On Error Resume Next
        LogError lErrNum, sErrSource, sErrDesc
    End If
    '(ALWAYS comment out two of the three options below)
    '*** EITHER display the error here
    'DisplayError lErrNum, sErrSource, sErrDesc
    '*** OR raise the error to the calling procedure
    Err.Raise lErrNum, sErrSource, sErrDesc
    '*** OR ignore the error and continue
    'Resume Next
End Function

Public Function GetDeleteChargeOverride(ChargeOverrideID As Long) As Recordset
'*************************************************************************************
'Purpose    : name
'Inputs     :
'Outputs    :
'Version: $SSSVersion:
'Created on : 13/03/2001
'Created by : Patrick
'*************************************************************************************

'-------------------------------
'*** Generic Procedure Code
'-------------------------------
    On Error GoTo errHandler
    Dim sModuleAndProcName As String
    sModuleAndProcName = "clsDatabase - GetDeleteChargeOverride Procedure"
    
'-------------------------------
'*** Specific Procedure Code
'-------------------------------
        
    Set colParameters = New Collection
    
    strSQL = "execute procedure sdchgoverride(p_chgoverrideid = ?, tran_flg = 'Y')"
    
    colParameters.Add ChargeOverrideID
    objApplication.dbAccess.Parameters = colParameters
    
    Set GetDeleteChargeOverride = objApplication.dbAccess.loadrs(strSQL)
       
Exit Function
'------------------------------------------------------------
'*** Generic Error Handling Code (ensure Exit Proc is above)
'------------------------------------------------------------
errHandler:
    If Err.Number = 0 Then Exit Function
    
    Dim lErrNum As Long, sErrDesc As String, sErrSource As String
    lErrNum = Err.Number
    sErrDesc = Err.Description
    sErrSource = "[" & App.Title & " - " & sModuleAndProcName & " - Build " & App.Major & "." & App.Minor & "." & App.Revision
    If Erl > 0 Then sErrSource = sErrSource & " - Line No " & CStr(Erl)
    If InStr(1, Err.Source, "-->") > 0 Then
        '*** This error has already been handled by our code
        sErrSource = sErrSource & "]" & vbNewLine & "  --> " & Err.Source
    Else
        '*** Newly generated error, log it here.
        sErrSource = sErrSource & "]" & vbNewLine & "  --> [Source: " & Err.Source & "]"
        On Error Resume Next
        LogError lErrNum, sErrSource, sErrDesc
    End If
    '(ALWAYS comment out two of the three options below)
    '*** EITHER display the error here
    'DisplayError lErrNum, sErrSource, sErrDesc
    '*** OR raise the error to the calling procedure
    Err.Raise lErrNum, sErrSource, sErrDesc
    '*** OR ignore the error and continue
    'Resume Next
End Function

Public Function GetCharge(PackageCode As Long, PackageOption As Long) As Recordset
'*************************************************************************************
'Purpose    : name
'Inputs     :
'Outputs    :
'Version: $SSSVersion:
'Created on : 13/03/2001
'Created by : Patrick
'*************************************************************************************

'-------------------------------
'*** Generic Procedure Code
'-------------------------------
    On Error GoTo errHandler
    Dim sModuleAndProcName As String
    sModuleAndProcName = "clsDatabase - GetCharge Procedure"
    
'-------------------------------
'*** Specific Procedure Code
'-------------------------------
        
    Set colParameters = New Collection
    
    strSQL = "execute procedure ss_ovrrd_chg_codes(p_pkg_code = ?"
    colParameters.Add PackageCode
    
    If PackageOption <> 0 Then
        strSQL = strSQL & ", p_option = ?"
        colParameters.Add PackageOption
    End If
    
    strSQL = strSQL & ", resume_flg = 'Y')"
    
    objApplication.dbAccess.Parameters = colParameters
    
    Set GetCharge = objApplication.dbAccess.loadrs(strSQL)
       
Exit Function
'------------------------------------------------------------
'*** Generic Error Handling Code (ensure Exit Proc is above)
'------------------------------------------------------------
errHandler:
    If Err.Number = 0 Then Exit Function
    
    Dim lErrNum As Long, sErrDesc As String, sErrSource As String
    lErrNum = Err.Number
    sErrDesc = Err.Description
    sErrSource = "[" & App.Title & " - " & sModuleAndProcName & " - Build " & App.Major & "." & App.Minor & "." & App.Revision
    If Erl > 0 Then sErrSource = sErrSource & " - Line No " & CStr(Erl)
    If InStr(1, Err.Source, "-->") > 0 Then
        '*** This error has already been handled by our code
        sErrSource = sErrSource & "]" & vbNewLine & "  --> " & Err.Source
    Else
        '*** Newly generated error, log it here.
        sErrSource = sErrSource & "]" & vbNewLine & "  --> [Source: " & Err.Source & "]"
        On Error Resume Next
        LogError lErrNum, sErrSource, sErrDesc
    End If
    '(ALWAYS comment out two of the three options below)
    '*** EITHER display the error here
    'DisplayError lErrNum, sErrSource, sErrDesc
    '*** OR raise the error to the calling procedure
    Err.Raise lErrNum, sErrSource, sErrDesc
    '*** OR ignore the error and continue
    'Resume Next
End Function

Public Function GetPackageCharge(PackageCode As Long, PackageOption As Long) As Recordset
'*************************************************************************************
'Purpose    : name
'Inputs     :
'Outputs    :
'Version: $SSSVersion:
'Created on : 13/03/2001
'Created by : Patrick
'*************************************************************************************

'-------------------------------
'*** Generic Procedure Code
'-------------------------------
    On Error GoTo errHandler
    Dim sModuleAndProcName As String
    sModuleAndProcName = "clsDatabase - GetPackageCharge Procedure"
    
'-------------------------------
'*** Specific Procedure Code
'-------------------------------
        
    Set colParameters = New Collection
    
    strSQL = "select chg_code, chg_narr" & _
        " from charge_code " & _
        " where chg_code in (select chg_code from pkg_chgh where pkg_code = ?"
        
    colParameters.Add PackageCode
    
    If PackageOption <> 0 Then
        strSQL = strSQL & " and option = ?"
        colParameters.Add PackageOption
    End If
    
    strSQL = strSQL & " and current between start_date and end_date)"
    strSQL = strSQL & " order by chg_narr"
           
    objApplication.dbAccess.Parameters = colParameters
    
    Set GetPackageCharge = objApplication.dbAccess.loadrs(strSQL)
       
Exit Function
'------------------------------------------------------------
'*** Generic Error Handling Code (ensure Exit Proc is above)
'------------------------------------------------------------
errHandler:
    If Err.Number = 0 Then Exit Function
    
    Dim lErrNum As Long, sErrDesc As String, sErrSource As String
    lErrNum = Err.Number
    sErrDesc = Err.Description
    sErrSource = "[" & App.Title & " - " & sModuleAndProcName & " - Build " & App.Major & "." & App.Minor & "." & App.Revision
    If Erl > 0 Then sErrSource = sErrSource & " - Line No " & CStr(Erl)
    If InStr(1, Err.Source, "-->") > 0 Then
        '*** This error has already been handled by our code
        sErrSource = sErrSource & "]" & vbNewLine & "  --> " & Err.Source
    Else
        '*** Newly generated error, log it here.
        sErrSource = sErrSource & "]" & vbNewLine & "  --> [Source: " & Err.Source & "]"
        On Error Resume Next
        LogError lErrNum, sErrSource, sErrDesc
    End If
    '(ALWAYS comment out two of the three options below)
    '*** EITHER display the error here
    'DisplayError lErrNum, sErrSource, sErrDesc
    '*** OR raise the error to the calling procedure
    Err.Raise lErrNum, sErrSource, sErrDesc
    '*** OR ignore the error and continue
    'Resume Next
End Function

Public Function GetServiceList() As ADODB.Recordset
'***********************************************************************************************
'Purpose    : Get call serice for the account
'Inputs     :
'Outputs    :
'Version: $SSSVersion:
'Created on : 04/09/2001
'Created by : Patrick
'***********************************************************************************************

'-------------------------------
'*** Generic Procedure Code
'-------------------------------
    On Error GoTo errHandler
    Dim sModuleAndProcName As String
    sModuleAndProcName = "clsDatabase - GetServiceList Procedure"
    
'-------------------------------
'*** Specific Procedure Code
'-------------------------------
    
    Set colParameters = New Collection
    
    strSQL = "execute procedure ss_serv_list2(p_subscriber_code = ?, resume_flg = 'Y',isolation_flg = 'D')"
    colParameters.Add objApplication.ContactCode
    
    objApplication.dbAccess.Parameters = colParameters
    
    Set GetServiceList = objApplication.dbAccess.loadrs(strSQL)

Exit Function
'------------------------------------------------------------
'*** Generic Error Handling Code (ensure Exit Proc is above)
'------------------------------------------------------------
errHandler:
    If Err.Number = 0 Then Exit Function
    
    Dim lErrNum As Long, sErrDesc As String, sErrSource As String
    lErrNum = Err.Number
    sErrDesc = Err.Description
    sErrSource = "[" & App.Title & " - " & sModuleAndProcName & " - Build " & App.Major & "." & App.Minor & "." & App.Revision
    If Erl > 0 Then sErrSource = sErrSource & " - Line No " & CStr(Erl)
    If InStr(1, Err.Source, "-->") > 0 Then
        '*** This error has already been handled by our code
        sErrSource = sErrSource & "]" & vbNewLine & "  --> " & Err.Source
    Else
        '*** Newly generated error, log it here.
        sErrSource = sErrSource & "]" & vbNewLine & "  --> [Source: " & Err.Source & "]"
        On Error Resume Next
        LogError lErrNum, sErrSource, sErrDesc
    End If
    '(ALWAYS comment out two of the three options below)
    '*** EITHER display the error here
    'DisplayError lErrNum, sErrSource, sErrDesc
    '*** OR raise the error to the calling procedure
    Err.Raise lErrNum, sErrSource, sErrDesc
    '*** OR ignore the error and continue
    'Resume Next

End Function


