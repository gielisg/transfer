Attribute VB_Name = "modLoadRS"
Option Explicit
Private colConnection As New Collection

Public Function LoadRS(ByVal strMode As String, _
    ByVal strDSN As String, _
    ByVal strSQL As String, _
    Optional ByVal intLockMode As Integer = 0, _
    Optional ByVal intIsolationLevel As Integer = 0, _
    Optional ByVal strURL As String, _
    Optional ByVal strServer As String) As ADODB.Recordset
    
    On Error GoTo errHandler
    Dim sModuleAndProcName As String
    sModuleAndProcName = "LoadRS "

Dim rs As ADODB.Recordset, cn As New ADODB.Connection
Dim obj As Object

    If strDSN = "" Or strSQL = "" Then
        MsgBox "Dsn and SQL string must be supplied"
        Exit Function
    End If
     strMode = UCase(strMode)
    
    If intLockMode < 0 Then
        MsgBox "Invalid Lock mode"
        Exit Function
    End If
    
    If intIsolationLevel <> 0 And intIsolationLevel <> 1 Then
        MsgBox "Invalid Isolation level"
        Exit Function
    End If

    If strMode <> "HTTP" And strMode <> "ODBC" And strMode <> "DCOM" Then
        MsgBox "Incorrect mode: must be ODBC, HTTP or DCOM"
        Exit Function
    End If
    
    If strMode = "HTTP" And strURL = "" Then
        MsgBox "URL must be supplied for HTTP mode"
        Exit Function
    End If

     If strMode = "ODBC" Then
        Set LoadRS = LoadRSODBC(strDSN, strSQL, intLockMode, intIsolationLevel)
        Exit Function
    ElseIf strMode = "DCOM" Then
        If strServer <> "" Then
            Set obj = CreateNewObject("LoadRsDS.LoadRs", True, strServer)
        Else
            Set obj = CreateNewObject("LoadRsDS.LoadRs", False, strServer)
        End If
        Set LoadRS = obj.ExecuteSQL(strSQL, strDSN, intLockMode, intIsolationLevel)
        Exit Function
    End If
     On Error GoTo errHandler
     
    sModuleAndProcName = "LoadRS"
    strURL = "http://" & strURL & "/Common/Server/ExecuteSQL.asp"
    strURL = strURL & "?DSN=" & URLEncode(strDSN)
    strURL = strURL & "&SQL=" & URLEncode(strSQL)
    strURL = strURL & "&LockMode=" & URLEncode(intLockMode)
    strURL = strURL & "&IsolationLevel=" & URLEncode(intIsolationLevel)
    
'    Screen.MousePointer = vbHourglass
    
    Set rs = New ADODB.Recordset
    rs.Open strURL
    
'    Screen.MousePointer = vbNormal
        
    Set LoadRS = rs
    Set rs = Nothing
    
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
    If Erl > 0 Then sErrSource = sErrSource & " - Line No " & str(Erl)
    If InStr(1, Err.Source, "-->") > 0 Then
        '*** This error has already been handled by our code
        sErrSource = sErrSource & "]" & vbNewLine & "  --> " & Err.Source
    Else
        '*** Newly generated error, log it here.
        sErrSource = sErrSource & "]" & vbNewLine & "  --> [Source: " & Err.Source & "]"
        On Error Resume Next
        'LogError lErrNum, sErrSource, sErrDesc
    End If
    '(ALWAYS comment out two of the three options below)
    '*** EITHER display the error here
        'DisplayError lErrNum, sErrSource, sErrDesc
    '*** OR raise the error to the calling procedure
        Err.Raise lErrNum, sErrSource, sErrDesc
    '*** OR ignore the error and continue
        'Resume Nexttl, colLookup)
      
End Function

Private Function LoadRSODBC(ByVal strDSN As String, _
    ByVal strSQL As String, _
    ByVal intLockMode As Integer, _
    ByVal intIsolationLevel As Integer _
    ) As ADODB.Recordset

Dim sModuleAndProcName As String
Dim rs As ADODB.Recordset
Dim cn As ADODB.Connection
Dim idx As Integer

    On Error GoTo errHandler
    sModuleAndProcName = "LoadRSODBC"
    
    ' Find if a connection already exists for this DSN
    For idx = 1 To colConnection.Count Step 2
        If colConnection.Item(idx) = strDSN Then
            Set cn = colConnection.Item(idx + 1)
            Exit For
        End If
    Next idx
    
    If cn Is Nothing Then
        Set cn = New ADODB.Connection
        cn.Open strDSN
        'cn.CursorLocation = adUseClient
        ' add the connection to collection for reuse
        cn.CommandTimeout = 1000
        If intLockMode = 0 Then
            cn.Execute "SET LOCK MODE TO WAIT"
        Else
            cn.Execute "SET LOCK MODE TO WAIT " & intLockMode
        End If
        If intIsolationLevel = 0 Then
            cn.Execute "SET ISOLATION TO COMMITTED READ"
        Else
            cn.Execute "SET ISOLATION TO DIRTY READ"
        End If
        colConnection.Add strDSN
        colConnection.Add cn
    End If

'    Screen.MousePointer = vbHourglass
    
    'Set rs = New ADODB.Recordset
    'Set rs.ActiveConnection = cn
    'rs.open strSQL
    'Set rs.ActiveConnection = Nothing
    Set rs = cn.Execute(strSQL)
'    Screen.MousePointer = vbNormal
        
    Set LoadRSODBC = rs
    Set rs = Nothing
    Exit Function
'------------------------------------------------------------
'*** Generic Error Handling Code (ensure Exit Proc is above)
'------------------------------------------------------------
errHandler:
    If Err.Number = 0 Then Exit Function
    ' Release the connection to avoid possible locks
    Set cn = Nothing
    For idx = 1 To colConnection.Count Step 2
        If colConnection.Item(idx) = strDSN Then
            colConnection.Remove idx ' removes the dsn name
            colConnection.Remove idx ' removes the connection
            Exit For
        End If
    Next idx

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
        'LogError lErrNum, sErrSource, sErrDesc
    End If
    '(ALWAYS comment out two of the three options below)
    '*** EITHER display the error here
        'DisplayError lErrNum, sErrSource, sErrDesc
    '*** OR raise the error to the calling procedure
        Err.Raise lErrNum, sErrSource, sErrDesc
    '*** OR ignore the error and continue
        'Resume Nexttl, colLookup)
End Function

Public Function URLEncode(ByVal strValue As String) As String
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
  '(ALWAYS comment out two of the three options below)
  '*** EITHER display the error here
    'DisplayError lErrNum, sErrSource, sErrDesc
  '*** OR raise the error to the calling procedure
    Err.Raise lErrNum, sErrSource, sErrDesc
  '*** OR ignore the error and continue
    'Resume Next

End Function

Public Function URLDecode(ByVal strValue As String) As String
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
  '(ALWAYS comment out two of the three options below)
  '*** EITHER display the error here
    'DisplayError lErrNum, sErrSource, sErrDesc
  '*** OR raise the error to the calling procedure
    Err.Raise lErrNum, sErrSource, sErrDesc
  '*** OR ignore the error and continue
    'Resume Next

End Function

Public Function FieldToString(objRS As ADODB.Recordset, ByVal varIndex As Variant) _
                    As String

  Dim varValue As Variant
  
  varValue = objRS(varIndex)

  If Not IsNull(varValue) Then
    FieldToString = Trim$(CStr(varValue))
  Else
    FieldToString = ""
  End If
    
End Function
Public Function VarTypeString(ByVal var As Variant) As String
    
    Dim iType As Integer
    
    iType = VarType(var)
    
    Select Case iType
    Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency
        VarTypeString = "integer"
    Case vbString, vbVariant
        VarTypeString = "string"
    Case vbDate
        VarTypeString = "date"
    Case Else
        VarTypeString = "string"
    End Select
    
End Function


