Attribute VB_Name = "CommonModules"
Option Explicit
Option Base 1

'*** Application Definitions
Global gsServerPath As String
Global gsAppTitle As String
Global gsAppVersion As String
Global gsUserID As String
Global gsPassword As String
Global gsUserKey As String
Global glErrNum As Long
Global gsErrDesc As String
Global gsErrSource As String
Global gRDSDataSpace As Object
Global gsPWD As String
Global Const VB_NULL_DATE = "00:00:00"
Global Const VB_NULL_LONG = -2 ^ 31
Global Const VB_NULL_INTEGER = -2 ^ 15
Global Const VB_NULL_DOUBLE = -10 ^ 99
'*** Win32 API Declarations
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'*********For ErrorHandling*********
' Define your custom errors here.  Be sure to use numbers
' greater than 512, to avoid conflicts with OLE error numbers.
Public Const MyObjectError1 = 1000
Public Const MyObjectError2 = 1010
Public Const MyObjectErrorN = 1234
Public Const MyUnhandledError = 9999
'*********End************************

Global Const DB_CONNECT = "DSN=Primus"

Private m_strDSN As String


Public Function szConvertDateToInformix(ByVal dtVBDate As Date, Optional bDateTime = True, Optional DefaultFormat As Variant) As String
'***********************************************************************************************
'Purpose    : Converts VB Date variable into an Informix Date/DateTime.
'Inputs     : dtVBDate - Date to be converted, a valid VB Date variable.
'             bDateTime - True - will convert into DateTime variable.
'                         False - will convert into Date Variable
'             DefaultFormat - An optional argument, specifying a forced format (e.g. "dd-mon-yy")
'Outputs    : String containing the converted date in Informix format.
'Version: $SSSVersion:                                                                                                1.147 $
'Created on : 10/12/1999
'Created by : 10/12/1999 Anand
'***********************************************************************************************

'-------------------------------
'*** Generic Procedure Code
'-------------------------------
1     On Error GoTo errHandler
2     Dim sModuleAndProcName As String
3     sModuleAndProcName = "CommonModules - szConvertDateToInformix Function"

'-------------------------------
'*** Specific Procedure Code
'-------------------------------
4     If TypeName(bDateTime) = "Boolean" Or IsMissing(bDateTime) Then
5         If IsMissing(DefaultFormat) Then
6             Select Case bDateTime
                 Case True
8                     szConvertDateToInformix = Format(dtVBDate, "yyyy-MM-dd hh:mm:ss")
9                 Case False
10                    szConvertDateToInformix = Format(dtVBDate, "dd/MM/yyyy")
11            End Select
12        Else
13            szConvertDateToInformix = Format(dtVBDate, DefaultFormat)
14        End If
15    Else
16        Err.Raise MyObjectErrorN, "Date Conversion", "bDateTime argument not a boolean."
17    End If
  
18    Exit Function


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

Public Sub LogErrorToEventLog(ByVal lngErrNo As Long, ByVal strSource As String, _
                    ByVal strDescription As String, Optional ByVal strNotes As String = "", _
                    Optional lngEventType As Long)
'***********************************************************************************************
'Purpose    : logs an error to the system application log
'Inputs     : lngErrNo - Error number
'             strSource - Error Source
'             strDescription - Error description
'             strNotes - Error note
'             lngEventType - Event type (vbLogEventTypeError 1 Error.
'                                       vbLogEventTypeWarning 2 Warning.
'                                       vbLogEventTypeInformation 4 Information.)
'Outputs    : none
'Version: $SSSVersion:                                                                                                1.147 $
'Created on : 05/04/2001
'Created by : Patrick
'***********************************************************************************************

'-------------------------------
'*** Generic Procedure Code
'-------------------------------
1     On Error GoTo errHandler
2     Dim sModuleAndProcName As String
3     sModuleAndProcName = "CommonModules - LogErrorToEventLog sub"

'-------------------------------
'*** Specific Procedure Code
'-------------------------------
    Select Case lngEventType
        Case vbLogEventTypeError, vbLogEventTypeInformation, vbLogEventTypeWarning
        
            App.LogEvent vbNewLine & "Error Number:" & lngErrNo & vbNewLine & vbNewLine & _
            "Description:" & vbNewLine & strDescription & vbNewLine & vbNewLine & _
            "Source: " & vbNewLine & strSource & vbNewLine & vbNewLine & _
            "Notes: " & vbNewLine & strNotes, _
            lngEventType
        Case Else

            App.LogEvent vbNewLine & "Error Number:" & lngErrNo & vbNewLine & vbNewLine & _
            "Description:" & vbNewLine & strDescription & vbNewLine & vbNewLine & _
            "Source: " & vbNewLine & strSource & vbNewLine & vbNewLine & _
            "Notes: " & vbNewLine & strNotes
    End Select
Exit Sub

'------------------------------------------------------------
'*** Generic Error Handling Code (ensure Exit Proc is above)
'------------------------------------------------------------
errHandler:
  If Err.Number = 0 Then Exit Sub
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
'    Err.Raise lErrNum, sErrSource, sErrDesc
  '*** OR ignore the error and continue
    Resume Next
End Sub
Public Sub SetDSN(DSN As String)
    m_strDSN = DSN
End Sub

Public Function GetNextNo(szLabel As String) As Long
'***********************************************************************************************
'Purpose    : Gets No. identified by szLabel from the registry, increments
'it and returns the no.
'Inputs     : szLabel - Key of the No. in registry.
'Outputs    : 100000 if the key is not found, else incremented no.
'Version: $SSSVersion:                   1.147 $
'Created on : 06/01/2000
'Created by : 06/01/2000 Anand
'***********************************************************************************************

'-------------------------------
'*** Generic Procedure Code
'-------------------------------
1     On Error GoTo errHandler
2     Dim sModuleAndProcName As String
3     sModuleAndProcName = "CommonModules - GetNextNo Function"

'-------------------------------
'*** Specific Procedure Code
'-------------------------------

4     Dim iNextNo As Long
      Dim DSN As String
      Dim rs As Object
      Dim Con As Object
      Dim strSQL As String
      Dim blnTran As Boolean

      If Trim(m_strDSN) <> "" Then
        DSN = m_strDSN
      Else
        DSN = GetSetting("Selcomm", "Configuration", "AutonumDSN", "")
      End If
      
      If DSN <> "" Then
          strSQL = "SELECT count(*) from autonum WHERE autonum_code =?"
          SubstitutePlaceholders strSQL, UCase$(szLabel), "string"
          Set Con = CreateObject("ADODB.CONNECTION")
          Con.Open DSN
          Con.Execute "SET LOCK MODE TO WAIT 300"
          Set rs = Con.Execute(strSQL)
          blnTran = False
          If Not rs.EOF Then
             If rs(0) > 0 Then
                Con.BeginTrans
                blnTran = True
                strSQL = "Update autonum set  last_num_used =  last_num_used + 1 " & _
                "WHERE autonum_code =?"
                SubstitutePlaceholders strSQL, UCase$(szLabel), "string"

                Con.Execute strSQL
                strSQL = "SELECT  last_num_used from autonum WHERE autonum_code =?"
                SubstitutePlaceholders strSQL, UCase$(szLabel), "string"
                Set rs = Con.Execute(strSQL)
                GetNextNo = rs(0)
                Con.CommitTrans
                blnTran = False
                Set Con = Nothing
                Exit Function
            End If
        End If
    Set Con = Nothing
    End If
  'Get settings from registry.
5     iNextNo = CLng(GetSetting("Selcomm", "AutoGenNo", szLabel, 99999))

  'See if the settings exist (if not GetSetting returns the default value)
6     If iNextNo = 99999 Then
      'Return 100000
7         GetNextNo = 100000
8     Else
      'Increment No and return
9         GetNextNo = iNextNo + 1
10    End If

  'Save new No. to registry
  'GetNextNo = iNextNo

11    SaveSetting "Selcomm", "AutoGenNo", szLabel, CStr(GetNextNo)

12    Exit Function

'------------------------------------------------------------
'*** Generic Error Handling Code (ensure Exit Proc is above)
'------------------------------------------------------------
errHandler:
  If blnTran Then Con.RollbackTrans
  Set Con = Nothing
  If Err.Number = 0 Then Exit Function
  Dim lErrNum As Long, sErrDesc As String, sErrSource As String
  lErrNum = Err.Number
  sErrDesc = Err.Description
  sErrSource = "[" & App.Title & " - " & sModuleAndProcName & " - Build " & _
App.Major & "." & App.Minor & "." & App.Revision
  If Erl > 0 Then sErrSource = sErrSource & " - Line No " & str(Erl)
  If InStr(1, Err.Source, "-->") > 0 Then
    '*** This error has already been handled by our code
    sErrSource = sErrSource & "]" & vbNewLine & "  --> " & Err.Source
  Else
    '*** Newly generated error, log it here.
    sErrSource = sErrSource & "]" & vbNewLine & "  --> [Source: " & _
Err.Source & "]"
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

Private Function GetErrorTextFromResource(ErrorNum As Long) _
          As String
'------------------------------------------------------------
'Purpose:
'Author:
'Created:
'Revised:
'------------------------------------------------------------

'-------------------------------
'*** Generic Procedure Code
'-------------------------------
1     On Error GoTo errHandler
2     Dim sModuleAndProcName As String
3     sModuleAndProcName = "CommonModules - GetErrorTextFromResource Function"

'-------------------------------
'*** Specific Procedure Code
'-------------------------------
4     Dim strMsg As String
  
  ' this function will retrieve an error description from a resource
  ' file (.RES).  The ErrorNum is the index of the string
  ' in the resource file.  Called by RaiseError
  
  ' get the string from a resource file
5     GetErrorTextFromResource = LoadResString(ErrorNum)
  
6     Exit Function

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


Public Sub RaiseError(ErrorNumber As Long, Optional Source As String, Optional Description As String)
      Dim strErrorText As String


      'there are a number of methods for retrieving the error
      'message.  The following method uses a resource file to
      'retrieve strings indexed by the error number you are
      'raising.
     ' strErrorText = GetErrorTextFromResource(ErrorNumber)

    If IsMissing(Source) Then
        Source = "Missing source"
    End If

      'raise an error back to the client
      Err.Raise vbObjectError + ErrorNumber, Source, strErrorText


End Sub

Public Function isStoredProcedure(SQL As String) As Boolean
'------------------------------------------------------------
'Purpose: See if they are executing a stored procedure. Our standard is to
'         return a zero as the first return.
'Author:
'Created:
'Revised:
'------------------------------------------------------------

'-------------------------------
'*** Generic Procedure Code
'-------------------------------
1     On Error GoTo errHandler
2     Dim sModuleAndProcName As String
3     sModuleAndProcName = "CommonModules - isStoredProcedure Function"

'-------------------------------
'*** Specific Procedure Code
'-------------------------------
4     If InStr(1, LCase(SQL), "execute procedure") = 0 Then
5         isStoredProcedure = False
6     Else
7         isStoredProcedure = True
       
8     End If
  
9     Exit Function

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


Public Sub SubstitutePlaceholders(ByRef strSQL As String, colVal As Variant, DataType As String)
'***********************************************************************************************
'Purpose    : Substitues "?" in a string with ColVal depending upon DataType.
'Inputs     : strSQL - String representing the sql where "?" needs to be replaced.
'             ColVal - A variant storing value that is copied in the place of "?"
'             DataType - A string representing datatype of ColVal.
'Outputs    : None.
'Version: $SSSVersion:                                                                                             1.147 $
'Created on : 31/12/1999
'Created by : 31/12/1999 Anand
'Revised    : 04/03/2000 Anand.
'           Line 58: To strip the "?" character off colval, to avoid inserting "?" in strSQL
'***********************************************************************************************

'-------------------------------
'*** Generic Procedure Code
'-------------------------------
1     On Error GoTo errHandler
2     Dim sModuleAndProcName As String
3     sModuleAndProcName = "CommonModules - SubstitutePlaceholders Procedure"

'-------------------------------
'*** Specific Procedure Code
'-------------------------------
4     Dim sStr As String
5     Dim str2 As String
6     Dim idx As Integer
7     Dim kdx As Integer
  
  'strSQL = strSQL & "              " Nobody knows why it was done!
  
  'if ColVal doesnot contain anything, its value is not Null/"", if it is declared as String * .
  'So initialise with spaces if there is no value in the variable.
8     If Not IsNull(colVal) Then  'for values that are not strings
9         If CStr(colVal) <> "" Then  'for values that are strings!
10            If Asc(CStr(colVal)) = 0 Then
              'initalise with Spaces.
11                colVal = " "
12            End If
13        End If
14    End If
  
  'Find "?" in the string, if cant find raise error
15    idx = InStr(1, strSQL, "?", 0)
16    If idx = 0 Then
17        Err.Raise 100000, "", "Invalid number of placeholders passed in SQL string."
18    End If
  
19    sStr = LCase(DataType)
  
  '//Modified by ASP 08/01/2000
20    Select Case Trim$(sStr)
  
      'Numbers and boolean values
          Case "long", "integer", "boolean"
22            If Not (IsNull(colVal)) Then
23                str2 = Trim$(CStr(colVal))
24            Else
25                str2 = "NULL"
26            End If
           
      'Convert to Informix Date format
27        Case "date"
          'ColVal is a VB Date in this case
28            If IsNull(colVal) Or colVal = CDate(VB_NULL_DATE) Then
29                str2 = "NULL"
30            Else
              'Informix expects date as 'dd/MM/yyyy'
31                str2 = Chr(39) & Trim$(szConvertDateToInformix(colVal, False)) & Chr(39)
32            End If

      'Convert to Informix DateTime format
33        Case "datetime"
          'ColVal is a VB Date in this case
34            If IsNull(colVal) Or colVal = CDate(VB_NULL_DATE) Then
35                str2 = "NULL"
36            Else
              'Informix expects datetime as "yyyy-MM-dd hh:mi:ss"
37                str2 = Chr(34) & Trim$(szConvertDateToInformix(colVal)) & Chr(34)
38            End If
          
      'Convert to Informix Date Format
39        Case "datestring"
          'ColVal is a VB String
40            If Trim$(CStr(colVal)) = "" Or colVal = VB_NULL_DATE Then
41                str2 = "NULL"
42            Else
43                str2 = Chr(39) & Trim$(CStr(colVal)) & Chr(39)
44            End If
      
      'Convert to Informix DateTime Format
45        Case "datetimestring"
          'ColVal is a VB String
46            If Trim$(CStr(colVal)) = "" Or colVal = VB_NULL_DATE Then
47                str2 = "NULL"
48            Else
49                str2 = Chr(34) & Trim$(CStr(colVal)) & Chr(34)
50            End If
          
      'Normal string
51        Case Else
52            If Trim$(CStr(colVal)) = "" Then
53                str2 = "NULL"
54            Else
55              str2 = Trim$(CStr(colVal))
                '//Following code is to replce " character into informix compatible format
56              str2 = Replace(str2, Chr(34), _
                    Chr(34) + " || " + _
                    Chr(34) + Chr(34) + Chr(34) + Chr(34) + _
                    " || " + Chr(34))
                '//Code Ends
                    
57              str2 = Chr(34) & Trim$(str2) & Chr(34)
58            End If
          
59    End Select
  
     'replce "= NULL" by "IS NULL" in the where clause
60   If Trim(UCase(str2)) = "NULL" Then
61      If Not isStoredProcedure(strSQL) Then
62          idx = InStrRev(Mid(strSQL, 1, InStr(1, strSQL, "?")), "=")
63          If idx <> 0 Then
                If Trim(Mid(strSQL, idx + 1, InStr(1, strSQL, "?") - idx)) = "" Then
64                  Mid(strSQL, idx, 1) = " "
65                  str2 = " IS NULL"
                End If
66          End If
67      End If
68   End If
     
     'Replace Carriage Return and Line feed char (vbCrLf) with " ",
     'to avoid error thrown by Informix ODBC driver.
     str2 = Replace(str2, vbCrLf, " ")
     
     'Put an escape character (130 - a non printable character) in place of "?".
69    strSQL = Replace(strSQL, "?", Trim$(Replace(str2, "?", Chr(130))), , 1)
      
     'If all placeholders have been replaced then replace the escape character
     'with "?" again so the final string that goes to the query remains intact.
70    If InStr(1, strSQL, "?") <= 0 Then strSQL = Replace(strSQL, Chr(130), "?")

71    Exit Sub

'------------------------------------------------------------
'*** Generic Error Handling Code (ensure Exit Proc is above)
'------------------------------------------------------------
errHandler:
  If Err.Number = 0 Then Exit Sub
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
End Sub

Public Function GetServer() As String
'***********************************************************************************************
'Purpose    : Get servername from registry entry and return it.
'Inputs     : None.
'Outputs    : Returns the server name used for creating objects
'Version: $SSSVersion:                                                                                             1.147 $
'Created on : 31/12/1999
'Created by : 31/12/1999 Anand
'***********************************************************************************************
'------------------------------------------------------------
'Purpose:
'Author:
'Created:
'Revised:
'------------------------------------------------------------

'-------------------------------
'*** Generic Procedure Code
'-------------------------------
1     On Error GoTo errHandler
2     Dim sModuleAndProcName As String
3     sModuleAndProcName = "CommonModules - GetServer Function"

'-------------------------------
'*** Specific Procedure Code
'-------------------------------
4     Dim szServerReg As String
      'Read Registry and Return ServerName
5     szServerReg = GetSetting("Selcomm", "Objects", "ServerName", "")
      'Return registry value.
6       If szServerReg <> "" Then
7       'Get Registry value
8           GetServer = szServerReg
9       Else
10          'Return Default Value
11          GetServer = "NTDEV"
12      End If
13     Exit Function

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

Public Function GetAppServer() As String
'***********************************************************************************************
'Purpose    : Get servername from registry entry and return it.
'Inputs     : None.
'Outputs    : Returns the server name used for creating objects
'Version: $SSSVersion:                                                                                             1.147 $
'Created on : 31/12/1999
'Created by : 31/12/1999 Anand
'***********************************************************************************************
'------------------------------------------------------------
'Purpose:
'Author:
'Created:
'Revised:
'------------------------------------------------------------

'-------------------------------
'*** Generic Procedure Code
'-------------------------------
    ' This is common initialisation code for the error handler
    Const ksMethod As String = "CreateServerObject"
    Dim sStatusInfo As String
    On Error GoTo ErrorHandler
    ' End error handler init

'-------------------------------
'*** Specific Procedure Code
'-------------------------------
4     Dim sResult As String

      'Read Registry and Return ServerName
5     sResult = GetSetting("Selcomm", "Objects", "ServerName", "")
      'Return registry value.
      
      GetAppServer = sResult
      
13     Exit Function

'------------------------------------------------------------
'*** Generic Error Handling Code (ensure Exit Proc is above)
'------------------------------------------------------------
ErrorHandler:

    If Err.Number = 0 Then
        Exit Function
    End If
    
    Call HandleError(App.EXEName, ksMethod, Err.Number, Err.Source, Err.Description, Erl, sStatusInfo, False)

End Function

Public Function GetQueueServer() As String
'***********************************************************************************************
'Purpose    : Get queue servername from registry entry and return it.
'Inputs     : None.
'Outputs    : Returns the server name used for creating objects
'Created on : 5/2/00
'Created by : Tharanga
'***********************************************************************************************
'------------------------------------------------------------
'Purpose:
'Author:
'Created:
'Revised:
'------------------------------------------------------------

'-------------------------------
'*** Generic Procedure Code
'-------------------------------
1     On Error GoTo errHandler
2     Dim sModuleAndProcName As String
3     sModuleAndProcName = "CommonModules - GetQueueServer Function"

'-------------------------------
'*** Specific Procedure Code
'-------------------------------
4     Dim szServerReg As String
  'Read Registry and Return ServerName
5     szServerReg = GetSetting("Selcomm", "Configuration", "MSMQServerName", "")
6     If szServerReg <> "" Then
      'Return registry value.
7         GetQueueServer = szServerReg
8     Else
9         SaveSetting "Selcomm", "Configuration", "MSMQServerName", "NTDEV"
      'return Default value
10        GetQueueServer = "NTDEV"
11    End If
12    Exit Function

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



Public Function SetConnectString(Optional szDSN As String) As String
'***********************************************************************************************
'Purpose    : Get DSN from registry entry and return it.
'Inputs     : szDSN - A string representing Default DSN, if calling program wants to pass it.
'Outputs    : A string containing DSN to be passed to the recordset in order to connect to the database.
'Version: $SSSVersion:                                                                                             1.147 $
'Created on : 31/12/1999
'Created by : 31/12/1999 Anand
'Revised    : 04/03/2000 Anand
'            To use the DSN if passed first to connect, else read registry, else use constant.
'***********************************************************************************************

'-------------------------------
'*** Generic Procedure Code
'-------------------------------
1     On Error GoTo errHandler
2     Dim sModuleAndProcName As String
3     sModuleAndProcName = "CommonModules - SetConnectString Function"

'-------------------------------
'*** Specific Procedure Code
'-------------------------------
4     Dim szDSNReg As String, szCatReg As String
  
5     If szDSN <> "" Then
        'IF DSN is passed from client use it
6        SetConnectString = "Data Source=" + szDSN
      
      'if the value is in the registry.
7     ElseIf GetSetting("Selcomm", "StartUp", "DSN", "") <> "" Then
      'Read Registry and Return ServerName
8         szDSNReg = GetSetting("Selcomm", "StartUp", "DSN", "")
9         szCatReg = GetSetting("Selcomm", "StartUp", "Initial Catalog")
10        SetConnectString = "Data Source=" + szDSNReg + _
          ";Initial Catalog=" + szCatReg + ";"
11    Else
      'Default Value
12        SetConnectString = DB_CONNECT
13    End If

14    Exit Function


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


Public Sub SetStartAndEndDatesByOptions(Flag As Integer, StartDate As String, EndDate As String)
'------------------------------------------------------------
'*****This is to set the start and end date of any period user passed
'*****Eg.  Today    - StartDate = 21/01/2000;   EndDate = 21/01/2000
'****      Last seven days  - StartDate = 14/01/2000;   EndDate = 21/01/2000
'****      This Month StatDate  = 01/01/2000;   EndDate = 31/01/2000
'****      Last Month StartDate = 01/12/1999;   EndDate = 31/12/1999
'Developed By   - Tharanga (For Primus OrderEntry)
'Date Developed - 21/01/2000
'------------------------------------------------------------
'Modifactions:
'1  Modified By patrick 10/03/2000
'   Add Option 5 for All date, set from 01/01/1900 to 01/01/2100
'2  Modified By Patrick 14/03/2000
'   Change This Week startdate from "StartDate = CStr(date)" to "StartDate = CStr(Date - 7)"
'   Change This Week EndDate from "EndDate = CStr(Date -7 )" to "EndDate = CStr(Date)"
'-------------------------------
'*** Generic Procedure Code
'-------------------------------
1     On Error GoTo errHandler
2     Dim sModuleAndProcName As String
3     sModuleAndProcName = "CommonModules - SetStartAndEndDatesByOptions Procedure"

'-------------------------------
'*** Specific Procedure Code
'-------------------------------
4     If Flag = 1 Then        '   Today
5         StartDate = CStr(Date)
6         EndDate = CStr(Date)
7     ElseIf Flag = 2 Then    '   Last Seven Days
8         StartDate = CStr(Date - 7)
9         EndDate = CStr(Date)
10    ElseIf Flag = 3 Then    '   This Month
11        EndDate = SetThisMonth(0)
12        StartDate = SetThisMonth(1)
13    ElseIf Flag = 4 Then    '   Last Month
14        StartDate = SetLastMonth(1)
15        EndDate = SetLastMonth(0)
16    ElseIf Flag = 5 Then
17        StartDate = #1/1/1900#
18        EndDate = #1/1/2100#
19    End If
  
20    Exit Sub

'------------------------------------------------------------
'*** Generic Error Handling Code (ensure Exit Proc is above)
'------------------------------------------------------------
errHandler:
  If Err.Number = 0 Then Exit Sub
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
End Sub

Public Function SetThisMonth(Strt As Integer) As String
'------------------------------------------------------------
'Purpose:
'Author:
'Created:
'Revised:
'------------------------------------------------------------

'-------------------------------
'*** Generic Procedure Code
'-------------------------------
1     On Error GoTo errHandler
2     Dim sModuleAndProcName As String
3     sModuleAndProcName = "CommonModules - SetThisMonth Function"

'-------------------------------
'*** Specific Procedure Code
'-------------------------------
4     Dim Mth As Integer
5     Dim Dte As Date

6     Dte = Date          '   get the current date
7     Mth = Month(Dte)    '   get the current month

8     If Strt = 0 Then    '   get the end data
9         While Month(Dte) = Mth  '   while the month is same
10            Dte = DateAdd("d", 1, Dte)
11        Wend
12        Dte = Dte - 1
13    End If
  
14    If Strt = 1 Then    '   get the start date
15        While Month(Dte) = Mth
16            Dte = DateAdd("d", -1, Dte)
17        Wend
18        Dte = Dte + 1
19    End If
  
  
20    SetThisMonth = CStr(Dte)

21    Exit Function

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

Public Function SetLastMonth(Strt As Integer)
'------------------------------------------------------------
'Purpose:
'Author:
'Created:
'Revised:
'------------------------------------------------------------

'-------------------------------
'*** Generic Procedure Code
'-------------------------------
1     On Error GoTo errHandler
2     Dim sModuleAndProcName As String
3     sModuleAndProcName = "CommonModules - SetLastMonth Function"

'-------------------------------
'*** Specific Procedure Code
'-------------------------------
4     Dim Mth As Integer
5     Dim Dte As Date

6     Dte = DateAdd("m", -1, Date) ' get the date one month before the current
7     Mth = Month(Dte) ' get the month of that date

8     If Strt = 0 Then    ' get the end date
9         While Month(Dte) = Mth  ' while the month is same
10            Dte = DateAdd("d", 1, Dte)
11        Wend
      
12        Dte = Dte - 1
13    End If
  
14    If Strt = 1 Then    ' get the start date
15        While Month(Dte) = Mth
16            Dte = DateAdd("d", -1, Dte)
17        Wend
18        Dte = Dte + 1
19    End If
  
  
20    SetLastMonth = CStr(Dte)
    
    
21    Exit Function

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

Public Sub DisplayError(ByVal ErrorNum As Long, ByVal errorSource As String, ByVal errorDesc As String)
'------------------------------------------------------------
'Purpose: Provides a central error display service
'Author:  Alistair Cook
'Created: 30 Sep 1999
'Revised: Anand Pansare (24 Feb 2000) -
'               Added Call to LogError
'               Deleted resetting of global variables, as they are not going to be used anymore.
'Revised: March 7th 2000 - Errors between 55000 and 56000 are
'         displayed with the description only to cater for
'         validation messages.
'------------------------------------------------------------

On Error Resume Next

Err.Clear

'--------------------------------
'*** Play the error sound
'--------------------------------
  'Play sndERROR
  
'--------------------------------
'*** Log the error.
'--------------------------------
  LogError ErrorNum, errorSource, errorDesc
  
'--------------------------------
'*** Display the error message
'--------------------------------
  Screen.MousePointer = vbDefault
  If ErrorNum >= 55000 And ErrorNum <= 56000 Then
    'This is a custom raised validation error, display the error description only
    MsgBox errorDesc, vbInformation, App.Title
  Else
    'This is an actual error, display the full error information
    'MsgBox "Error: " & vbNewLine & vbNewLine & _
      "An unexpected message was received from the following code:" _
        & vbNewLine & vbNewLine & "       " & _
      errorSource & "." _
        & vbNewLine & vbNewLine & _
      "Error Number: " & Trim(str(ErrorNum)) _
        & vbNewLine & vbNewLine & _
      "Error Description: " & errorDesc & "." _
        & vbNewLine & vbNewLine & _
      "An entry has been placed in the ErrorLog file." _
        & vbNewLine & vbNewLine & _
      "Please click on OK to continue using this application." _
      , vbInformation, App.Title
    If InStr(errorDesc, vbNewLine) = 0 Then
        errorDesc = errorDesc & vbNewLine
    End If
    If MsgBox("Error : " & Trim(str(ErrorNum)) _
        & ":" & _
      " " & Mid(errorDesc, 1, InStr(errorDesc, vbNewLine)) & "." _
      & "Please click on Yes for more Details" _
      , vbYesNo, App.Title) = vbYes Then
        
    MsgBox "Error: " & vbNewLine & vbNewLine & _
      "An unexpected message was received from the following code:" _
        & vbNewLine & vbNewLine & "       " & _
      errorSource & "." _
        & vbNewLine & vbNewLine & _
      "Error Number: " & Trim(str(ErrorNum)) _
        & vbNewLine & vbNewLine & _
      "Error Description: " & errorDesc & "." _
        & vbNewLine & vbNewLine & _
      "An entry has been placed in the ErrorLog file." _
        & vbNewLine & vbNewLine & _
      "Please click on OK to continue using this application." _
      , vbInformation, App.Title
      
    End If
  End If


End Sub

Public Function DisplayMessage(Message As String, Style As String)
'------------------------------------------------------------
'Purpose: Provides a central message display service
'Author:  Alistair Cook
'Created: 21 Oct 1999
'Revised:
'------------------------------------------------------------

'-------------------------------
'*** Generic Procedure Code
'-------------------------------
1     On Error GoTo errHandler
2     Dim sModuleAndProcName As String
3     sModuleAndProcName = "CommonModules - DisplayMessage Function"

'-------------------------------
'*** Specific Procedure Code
'-------------------------------

      '--------------------------------
      '*** Display the message
      '--------------------------------
4     Screen.MousePointer = vbDefault
5     If Style = "" Or Style = "0" Then Style = vbInformation
6     DisplayMessage = MsgBox(Trim$(Message), Style, App.Title)
  
7     Exit Function

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

Public Function CreateServerObject(ByVal ClassName As String, Optional ByVal Init As Boolean = False) As Object

    ' This is common initialisation code for the error handler
    Const ksMethod As String = "CreateServerObject"
    Dim sStatusInfo As String
    On Error GoTo ErrorHandler
    ' End error handler init
    
    Static sRemoteServerName As String
    Static bInitialised As Boolean
    
    If Init Or Not (bInitialised) Then
        sRemoteServerName = GetAppServer
        bInitialised = True
    End If
    
    If sRemoteServerName = "" Then
        Set CreateServerObject = CreateObject(ClassName)
    Else
        Set CreateServerObject = CreateObject(ClassName, sRemoteServerName)
    End If
    
    Exit Function
    
ErrorHandler:

    If Err.Number = 0 Then
        Exit Function
    End If
    
    Call HandleError(App.EXEName, ksMethod, Err.Number, Err.Source, Err.Description, Erl, "", False)
'    Call HandleError(Err.Number, Err.Source, Err.Description, False)
    
End Function


Public Function CreateNewObject(ByVal ObjectName As String, ByVal CreateOnRemotePC As Boolean, ByVal RemoteServerName As String) As Object
'------------------------------------------------------------
'Purpose: Creates the server COM object on the local PC or on
'         another PC on the network using DCOM or via a Web
'         Server using RDS.  Uses late binding to avoid object
'         compatibility errors.  The Data Access Components 2.0
'         or later must be installed on the local machine for
'         RDS to work.   The passed object must be registered
'         on the local machine for DCOM to work.
'Author:  Alistair Cook
'Created: 30 Nov 1999
'Revised:
'------------------------------------------------------------

'-------------------------------
'*** Generic Procedure Code
'-------------------------------
1     On Error GoTo errHandler
2     Dim sStatusInfo As String  'Variable data can be stored in this string within the Specific Procedure Code section to assist in error resolution
3     Dim sModuleAndProcName As String
4     sModuleAndProcName = "CommonModules - CreateNewObject Function"

'-------------------------------
'*** Specific Procedure Code
'-------------------------------
    
      '-----------------------------------------------------------------------
      '*** Populate sStatusInfo variable with passed info for error resolution
      '-----------------------------------------------------------------------
5     sStatusInfo = "Object Name: " & ObjectName & ", Remote Server Name: " & RemoteServerName & ", CreateOnRemotePC: "
6     If CreateOnRemotePC = True Then
7       sStatusInfo = sStatusInfo & "True"
8     Else
9       sStatusInfo = sStatusInfo & "False"
10    End If
  
      '----------------------------------------------------------------------
      '*** Create the passed ObjectName either Locally or via RDS or via DCOM
      '----------------------------------------------------------------------
11    If CreateOnRemotePC = False Then
        'Create the Server Object on the local machine
12      Set CreateNewObject = CreateObject(ObjectName)
13    ElseIf InStr(RemoteServerName, "http") > 0 Then
        'Create the Server Object on the remote machine using RDS via the web server
14      If gRDSDataSpace Is Nothing Then
15        Set gRDSDataSpace = CreateObject("RDS.DataSpace")
16      End If
17      Set CreateNewObject = gRDSDataSpace.CreateObject(ObjectName, RemoteServerName)
18    Else
        'Create the Server Object on the remote machine using DCOM
19      Set CreateNewObject = CreateObject(ObjectName, RemoteServerName)
20    End If
21    DoEvents

22    Exit Function

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
    If sStatusInfo <> "" Then sErrDesc = sErrDesc & vbNewLine & "  <Status Info for " & App.Title & " - " & sModuleAndProcName & ">   " & sStatusInfo
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


Private Function GetLocalComputerName() As String
'------------------------------------------------------------
'Purpose: Retrieves the name of the computer on which this
'         object is running.  This can be used for passing
'         to the error handling object.  Requires the
'         GetComputerName WIN32 API declaration.
'Author:  Alistair Cook
'Created: 3 Dec 1999
'Revised:
'------------------------------------------------------------

'-------------------------------
'*** Generic Procedure Code
'-------------------------------
1     On Error GoTo errHandler
2     Dim sModuleAndProcName As String
3     sModuleAndProcName = "CommonModules - GetLocalComputerName Function"

'-------------------------------
'*** Specific Procedure Code
'-------------------------------

4     Dim sName As String
5     Dim lSize As Long
6     Dim lResult As Long
  
7     On Error Resume Next
8     GetLocalComputerName = ""
9     sName = Space(64)
10    lSize = 64
  '*** Call the WIN32 API function
11    lResult = GetComputerName(sName, lSize)
12    If lResult = 1 Then
13        GetLocalComputerName = Left$(sName, lSize)
14    End If

15    Exit Function
  
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


Public Function InDevelopmentMode() As Boolean
'------------------------------------------------------------
'Purpose:
'Author:
'Created:
'Revised:
'------------------------------------------------------------

'-------------------------------
'*** Generic Procedure Code
'-------------------------------
1     On Error GoTo errHandler
2     Dim sModuleAndProcName As String
3     sModuleAndProcName = "CommonModules - InDevelopmentMode Function"

'-------------------------------
'*** Specific Procedure Code
'-------------------------------
4     If InStr(gsServerPath, "http") = 0 Then
5       InDevelopmentMode = True
6       If gsUserID = "" Then
7         gsUserID = "a"
8       End If
9       If gsPassword = "" Then
10        gsPassword = "x"
11      End If
12      If gsUserKey = "" Then
13        gsUserKey = Chr$(50) + Chr$(52) + Chr$(54) + Chr$(56) + Chr$(62) _
              + Chr$(63) + Chr$(64) + Chr$(65) + Chr$(66) + Chr$(67) + _
              Chr$(68) + Chr$(70) + Chr$(71) + Chr$(72) + Chr$(73) + _
              Chr$(74) + Chr$(75) + Chr$(77) + Chr$(78) + Chr$(79) + _
              Chr$(80) + Chr$(81) + Chr$(82) + Chr$(83) + Chr$(84) + Chr$(85)
14      End If
15    Else
16      InDevelopmentMode = False
17    End If
  
18    Exit Function

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


Public Sub LogError(ErrorNum As Long, errorSource As String, errorDesc As String, _
    Optional ErrorFile As String, Optional LogToEvent As Boolean = False)
'------------------------------------------------------------
'Purpose: Provides a central error logging service
'Author:  Alistair Cook
'Created: 30 Sep 1999
'Revised: 04/05/01 Patrick add optional parameter for error log file name
'         07/05/01 Patrick add optional parameter for log error to nt event log
'------------------------------------------------------------

  Dim lFileNumber As Long
  Dim ErrorLogName As String
  
  On Error Resume Next
    
'--------------------------------
'*** Write error to Log File
'--------------------------------
  If Trim(ErrorFile) <> "" Then
    ErrorLogName = "C:\" & Trim(ErrorFile)
  Else
    ErrorLogName = "C:\ErrorLog.txt"
  End If
  lFileNumber = FreeFile
  Open ErrorLogName For Append As lFileNumber
  Print #lFileNumber, Format$(Now, "dd mmm yyyy hh:nn:ss am/pm") _
        & vbNewLine & "      " & errorSource & "" _
        & vbNewLine & "  Err Num:  " & Trim(str(ErrorNum)) _
        & vbNewLine & "  Err Desc: " & errorDesc & "." & vbNewLine
  Close lFileNumber
  
  If LogToEvent Then
    LogErrorToEventLog ErrorNum, errorSource, errorDesc, "", vbLogEventTypeError
  End If

End Sub



Public Sub LogDebugInfo(DebugInfo As String)
'------------------------------------------------------------
'Purpose: Provides a central debug logging service. Logs info
'         to C:\Debug.txt only if the file exists.
'Author:  Alistair Cook
'Created: 1 March 2000
'Revised:
'------------------------------------------------------------

  Dim lFileNumber As Long
  Dim DebugFileName As String
  
  On Error Resume Next
    
'--------------------------------
'*** Write error to Log File
'--------------------------------
  DebugFileName = "C:\Debug.txt"
  '*** If the file exists
  If Dir(DebugFileName) <> "" Then
    '*** Write the passed info to the file
    lFileNumber = FreeFile
    Open DebugFileName For Append As lFileNumber
    Print #lFileNumber, _
          Format$(Now, "dd mmm yyyy hh:nn:ss am/pm") & vbNewLine & _
          "    " & DebugInfo & vbNewLine
    Close lFileNumber
  End If
  
End Sub

Public Function AdoToArray(rsAdo As Object) As Variant
'------------------------------------------------------------
'Purpose: Converts an ADO Recorset into a variant array
'         containing the recordset structure and data
'Author:
'Created: 13 Jan 2000
'Revised:
'------------------------------------------------------------

'-------------------------------
'*** Generic Procedure Code
'-------------------------------
1     On Error GoTo errHandler
2     Dim sModuleAndProcName As String
3     sModuleAndProcName = "CommonModules - AdoToArray Function"

'-------------------------------
'*** Specific Procedure Code
'-------------------------------
4     Dim arrDataStructure() As Variant
5     Dim arrData() As Variant
6     Dim arrDataAndStructure(0 To 1) As Variant
7     Dim i As Long

      '*** Build the ADO Data Structure array
8     If rsAdo.Fields.Count = 0 Then
9       Err.Raise 40000, "", "The returned recordset does not contain any fields."
10    End If
11    ReDim arrDataStructure(0 To rsAdo.Fields.Count - 1, 0 To 5)
12    For i = 0 To rsAdo.Fields.Count - 1
13      arrDataStructure(i, 0) = rsAdo.Fields(i).Name
14      arrDataStructure(i, 1) = rsAdo.Fields(i).Type
15      arrDataStructure(i, 2) = rsAdo.Fields(i).DefinedSize
16      arrDataStructure(i, 3) = rsAdo.Fields(i).Attributes
17      arrDataStructure(i, 4) = rsAdo.Fields(i).NumericScale
18      arrDataStructure(i, 5) = rsAdo.Fields(i).Precision
19    Next i

      '*** Build the Data array
20    If rsAdo.BOF And rsAdo.EOF Then
21      ReDim arrData(0 To 0, 0 To 0)
22      arrData(0, 0) = "No Data"
23    Else
24      rsAdo.MoveFirst
25      arrData() = rsAdo.GetRows()
26    End If

      '*** Add the two arrays into one
27    arrDataAndStructure(0) = arrDataStructure()
28    arrDataAndStructure(1) = arrData()
  
29    AdoToArray = arrDataAndStructure()

30    Exit Function

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



Public Function ArrayToADO(arrDataAndStructure) As Object
'------------------------------------------------------------
'Purpose: Builds and returns an ADO Recordset object based on
'         structure & data info in the passed array
'Author:
'Created: 13 Jan 2000
'Revised:
'------------------------------------------------------------

'-------------------------------
'*** Generic Procedure Code
'-------------------------------
1     On Error GoTo errHandler
2     Dim sModuleAndProcName As String
3     sModuleAndProcName = "CommonModules - ArrayToADO Function"

'-------------------------------
'*** Specific Procedure Code
'-------------------------------
4     Dim rsAdo As Object
5     Dim lFieldNum As Long
6     Dim lRecNum As Long
7     Dim i As Long
  
      '*** Create the ADO Recordset from the structure stored in the array
8     Set rsAdo = CreateObject("ADODB.Recordset")
9     For i = 0 To UBound(arrDataAndStructure(0), 1)
        '*** If the field name is blank then set it to the ordinal number
10      If arrDataAndStructure(0)(i, 0) = "" Then
11       arrDataAndStructure(0)(i, 0) = Trim$(str(i))
12      End If
13      rsAdo.Fields.Append arrDataAndStructure(0)(i, 0), arrDataAndStructure(0)(i, 1), arrDataAndStructure(0)(i, 2), arrDataAndStructure(0)(i, 3)
14      rsAdo.Fields(i).NumericScale = arrDataAndStructure(0)(i, 4)
15      rsAdo.Fields(i).Precision = arrDataAndStructure(0)(i, 5)
16      DoEvents
17    Next i
18    rsAdo.Open

      '*** Populate the ADO Recordset from the data stored in the array
19    If arrDataAndStructure(1)(0, 0) <> "No Data" Then
20      For lRecNum = 0 To UBound(arrDataAndStructure(1), 2)
21        rsAdo.AddNew
22        For lFieldNum = 0 To UBound(arrDataAndStructure(1), 1)
23          rsAdo.Fields(lFieldNum) = arrDataAndStructure(1)(lFieldNum, lRecNum)
24        Next lFieldNum
25        rsAdo.Update
26        DoEvents
27      Next lRecNum
      'rsAdo.MoveFirst
28    End If
29    DoEvents
  
      '*** Return the recordset to the calling procedure
30    Set ArrayToADO = rsAdo
  
31    Exit Function

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



Public Function GetServerPath(ParentLocation As String) As String
'------------------------------------------------------------
'Purpose:
'Author:
'Created:
'Revised:
'------------------------------------------------------------

'-------------------------------
'*** Generic Procedure Code
'-------------------------------
1     On Error GoTo errHandler
2     Dim sModuleAndProcName As String
3     sModuleAndProcName = "CommonModules - GetServerPath Function"

'-------------------------------
'*** Specific Procedure Code
'-------------------------------

4     Dim lPosition As Long

  ' Find the position of the last separator character.
5     For lPosition = Len(ParentLocation) To 1 Step -1
6        If Mid$(ParentLocation, lPosition, 1) = "/" Or _
        Mid$(ParentLocation, lPosition, 1) = "\" Then Exit For
7     Next lPosition
  
  ' Strip the name of the current .vbd file.
8     GetServerPath = Left$(ParentLocation, lPosition)
  
9     Exit Function
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

Public Function GetSystemKey(ByVal UserID As String) As String
'------------------------------------------------------------
'Purpose: This will give a system key to the user
'Author:
'Created:
'Revised:
'------------------------------------------------------------

'-------------------------------
'*** Generic Procedure Code
'-------------------------------
1     On Error GoTo errHandler
2     Dim sModuleAndProcName As String
3     sModuleAndProcName = "CommonModules - GetSystemKey Function"

'-------------------------------
'*** Specific Procedure Code
'-------------------------------
    
4     Dim SystemKey As String
  
5     SystemKey = Chr$(50) + Chr$(52) + Chr$(54) + Chr$(56) + Chr$(62) _
        + Chr$(63) + Chr$(64) + Chr$(65) + Chr$(66) + Chr$(67) + _
        Chr$(68) + Chr$(70) + Chr$(71) + Chr$(72) + Chr$(73) + _
        Chr$(74) + Chr$(75) + Chr$(77) + Chr$(78) + Chr$(79) + _
        Chr$(80) + Chr$(81) + Chr$(82) + Chr$(83) + Chr$(84) + Chr$(85)

6     GetSystemKey = SystemKey
    
7     Exit Function
  
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

Public Sub CheckNumeric(objControl As Control, KeyCd As Integer, WholeDigs As Integer, DecimalDigs As Integer)
'------------------------------------------------------------
'Purpose:   To Check if the Key pressed (KeyCd) is a valid key for numeric format
'           defined by No of whole digits (WholeDigs) and No of decimal digits (DecimalDigs)
'           If it is not valid, KeyCd is set to 0, thus "unaccepting" the key.
'           IMPORTANT: - Call this function from Keypress event of the control (objControl)
'Author:    Anand Pansare
'Created:   03/03/2000
'Revised:
'------------------------------------------------------------

'-------------------------------
'*** Generic Procedure Code
'-------------------------------
1   On Error GoTo errHandler
2   Dim sModuleAndProcName As String
3   sModuleAndProcName = "CommonModules - CheckNumeric"

'-------------------------------
'*** Specific Procedure Code
'-------------------------------
4   Dim NoSplits() As String
5   Dim iPos As Integer

    'Supress all printable characters other than numeric
6   If Chr(KeyCd) Like "#" Or (Chr(KeyCd) Like "." And DecimalDigs > 0) Then

7       iPos = objControl.SelStart
        'If "." is pressed, then see if it is valid.
8       If Chr(KeyCd) Like "." Then
9           If InStr(1, objControl.Text, ".") > 0 Or iPos = 0 Then KeyCd = 0
            'If a normal numeric character then check for validity of the character,
            'in whole part as well as decimal part.
10      ElseIf iPos > 0 Then
11          NoSplits = Split(objControl.Text, ".")
            'Whole part
12          If iPos <= Len(NoSplits(0)) Then
13              If Len(NoSplits(0)) >= WholeDigs Then KeyCd = 0
            'Decimal Part
14          ElseIf iPos >= Len(NoSplits(0)) + 1 Then
15              If UBound(NoSplits) = 1 And Len(NoSplits(1)) >= DecimalDigs Then KeyCd = 0
16          End If
17      End If

18  ElseIf KeyCd >= 32 And KeyCd < 127 Then
19      KeyCd = 0
20  End If

21  Exit Sub

'------------------------------------------------------------
'*** Generic Error Handling Code (ensure Exit Proc is above)
'------------------------------------------------------------
errHandler:
    If Err.Number = 0 Then Exit Sub
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
End Sub

Function CreateErrorObject() As Object
'------------------------------------------------------------
'Purpose: Creates an instance of the standard error handling
'         object on the local PC. Uses late binding to avoid
'         object compatibility errors.
'Author:  Alistair Cook
'Created: 30 Nov 1999
'Revised:
'------------------------------------------------------------

'-------------------------------
'*** Generic Procedure Code
'-------------------------------
  On Error GoTo errHandler
  Dim CurrentProcName As String
'  CurrentProcName = objApp.AppTitle & " CreateErrorObject Function " & objApp.AppVersion

'-------------------------------
'*** Specific Procedure Code
'-------------------------------
 ' Set CreateErrorObject = CreateObject("SelcommApplication.SelcommErrHandler")
  DoEvents

  Exit Function

'-------------------------------
'*** Generic Error Handling Code
'    (Ensure Exit Proc is above)
'-------------------------------
errHandler:
  Err.Clear
  'An error is already being handled, do not
  'propagate any more errors that occur here.
End Function

Public Function SetConnectString2(Optional szDSN As String) As String
'***********************************************************************************************
'Purpose    : This is a workaround for the problem which ODBC returns db not found and cursor error.
'             It works exactly the same as SetConnectString except it does not use Initial Catalog.
'Inputs     : szDSN - A string representing Default DSN, if calling program wants to pass it.
'Outputs    : A string containing DSN to be passed to the recordset in order to connect to the database.
'Version: $SSSVersion:                                                                                             1.147 $
'Created on : 31/12/1999
'Cretaed by: David
'***********************************************************************************************

'-------------------------------
'*** Generic Procedure Code
'-------------------------------
1     On Error GoTo errHandler
2     Dim sModuleAndProcName As String
3     sModuleAndProcName = "CommonModules - SetConnectString Function"

'-------------------------------
'*** Specific Procedure Code
'-------------------------------
4     Dim szDSNReg As String, szCatReg As String
  
5     If szDSN <> "" Then
        'IF DSN is passed from client use it
6        SetConnectString2 = "Data Source=" + szDSN
      
      'if the value is in the registry.
7     ElseIf GetSetting("Selcomm", "StartUp", "DSN", "") <> "" Then
      'Read Registry and Return ServerName
8         szDSNReg = GetSetting("Selcomm", "StartUp", "DSN", "")
9         szCatReg = GetSetting("Selcomm", "StartUp", "Initial Catalog")
10        SetConnectString2 = "Data Source=" + szDSNReg
11    Else
      'Default Value
12        SetConnectString2 = DB_CONNECT
13    End If

14    Exit Function


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

Public Function LoadXML(ByVal strURL As String) As Object

'***********************************************************************************************
'Purpose    : Loads an XML file and returns it
'
'Inputs     : strURL - the URL to an xml document
'Outputs    : An XML document ("MSXML.DomDocument")
'Version: $SSSVersion:                                                                                             1.147 $
'Created on : 19/9/2000
'Created by: Darren
'***********************************************************************************************

'-------------------------------
'*** Generic Procedure Code
'-------------------------------
1     On Error GoTo errHandler
2     Dim sModuleAndProcName As String
3     sModuleAndProcName = "CommonModules - LoadXML Function"

  Dim docXML As Object
  
  Set docXML = CreateObject("MSXML.DomDocument")
  
  docXML.async = False
  docXML.Load strURL
  
  If docXML.parsed = False Then
    Err.Raise vbObjectError + 6000, "LoadXML", "Failed to Load the XML file"
  End If
  
  Set LoadXML = docXML
  Set docXML = Nothing
  
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

Sub AddXMLHeader(ByRef sXML As String)

    sXML = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & "?>"
    
End Sub

Function ElementStart(Label As String) As String
    
    ElementStart = "<" & Label & ">"
    
End Function

Function ElementEnd(Label As String) As String
    
    ElementEnd = "</" & Label & ">"

End Function

Function Element(Label, Data) As String
    
    Element = "<" & Label & ">" & Data & "</" & Label & ">"
    
End Function

Sub AddElement(ByRef XML As String, Label As String, Data As String)
    
    If Trim$(Data) = "" Then
        XML = XML & "<" & Trim$(Label) & "></" & Trim$(Label) & ">"
    Else
        XML = XML & "<" & Trim$(Label) & ">" & Trim$(Data) & "</" & Trim$(Label) & ">"
    End If

End Sub

Public Sub HandleError(ByVal Module As String, ByVal Method As String, ByVal ENumber As Long, ByVal ESource As String, ByVal EDescription As String, ByVal ELine As Long, ByVal StatusInfo As String, ByVal Display As Boolean)
    
  Dim lErrNum As Long
  Dim sErrDesc As String
  Dim sErrSource As String
  
  lErrNum = ENumber
  sErrDesc = EDescription
  sErrSource = "[" & App.Title & " - " & Module & " - " & Method & " - Build " & App.Major & "." & App.Minor & "." & App.Revision
  If ELine > 0 Then sErrSource = sErrSource & " - Line No " & str(Erl)
  If InStr(1, ESource, "-->") > 0 Then
    '*** This error has already been handled by our code
    sErrSource = sErrSource & "]" & vbNewLine & "  --> " & ESource
  Else
    '*** Newly generated error, log it here.
    If StatusInfo <> "" Then
      sErrDesc = sErrDesc & vbNewLine & "  <Status Info for " & App.Title & " - " & Module & " - " & Method & ">   " & StatusInfo
    End If
    sErrSource = sErrSource & "]" & vbNewLine & "  --> [Source: " & Err.Source & "]"
    On Error Resume Next
    LogError lErrNum, sErrSource, sErrDesc
  End If
  If Display Then
    DisplayError lErrNum, sErrSource, sErrDesc
  Else
    Err.Raise lErrNum, sErrSource, sErrDesc
  End If
  '(ALWAYS comment out two of the three options below)
  '*** EITHER display the error here
    'DisplayError lErrNum, sErrSource, sErrDesc
  '*** OR raise the error to the calling procedure
''    Err.Raise lErrNum, sErrSource, sErrDesc
  '*** OR ignore the error and continue
    'Resume Next

End Sub

Public Function DecodegNo(ByVal strValue As String) As String
On Error GoTo Err_Handler
Dim sModuleAndProcName As String
sModuleAndProcName = "CommonModules - " & "DecodegNo"

1     Dim nPos As Integer
2     Dim strBefore As String
3     Dim strAfter As String
4     Dim strLetter As String
5     Dim strTemp As String
6     Dim lngASCII As Integer
  
        'Replace all "+" signs
7       strTemp = Replace(strValue, "+", " ")
  
        'Replace all occurences of a hex representation with its ascii value
8       nPos = InStr(1, strTemp, "%")

9       Do While nPos <> 0
  
            'Each occurence is of the form %XY where XY is the Hex Value of the letter
10          strBefore = Mid(strTemp, 1, nPos - 1)
11          strAfter = Mid(strTemp, nPos + 1 + 2) ' after the %XY part
        
            'Get the two letters that comprise the Hex Value
12          strLetter = Mid(strTemp, nPos + 1, 2)
    
            'Now convert the Hex value to a number (this is the ASCII value of the encoded letter)
13          lngASCII = CLng("&h" & strLetter)
    
            'Get the Unencoded value of the Letter
14          strLetter = (lngASCII)
    
15          strTemp = strBefore & strLetter & strAfter
16          nPos = InStr(1, strTemp, "%")
17      Loop
  
18      DecodegNo = strTemp

Exit Function
Err_Handler:

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
    
    'Either Dispaly the error
    'DisplayError lErrNum, sErrSource, sErrDesc
    
    'OR ignore the error and continue
    'Resume Next
    
    'OR Raise the error to the calling procedure
    Err.Raise lErrNum, sErrSource, sErrDesc

    
End Function

Public Function ConvgNo(ByVal strValue As String) As String
On Error GoTo Err_Handler
Dim sModuleAndProcName As String
sModuleAndProcName = "CommonModules - " & "ConvgNo"


Dim idx As Long
Dim jdx As Long
   
1       For idx = 1 To Len(strValue)
                    
2       jdx = Mid(strValue, idx, 1)
            
3           If CInt(jdx) = Chr$(50) - 1 Then
4               jdx = Chr$(48)
5           ElseIf CInt(jdx) = Chr$(48) + Chr$(50) Then
6               jdx = Chr$(49)
7           ElseIf idx Mod Chr$(50) = Chr$(48) Then
8                jdx = CInt(jdx) / 2
9           End If
            
10          ConvgNo = CStr(ConvgNo) & CStr(jdx)
                            
11          Next
    
    
Exit Function
Err_Handler:

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
    
    'Either Dispaly the error
    'DisplayError lErrNum, sErrSource, sErrDesc
    
    'OR ignore the error and continue
    'Resume Next
    
    'OR Raise the error to the calling procedure
    Err.Raise lErrNum, sErrSource, sErrDesc
    
End Function

Public Function IsValidURL(InputURL As String) As Boolean
    On Error Resume Next
    Dim chkURL As String
    Dim objHTTP As Object
    Dim sHTML As String
    
    IsValidURL = False
    chkURL = InputURL
    
    If LCase(Left(chkURL, 7)) <> "http://" Then chkURL = "http://" & chkURL
    
    Set objHTTP = CreateObject("Microsoft.XMLHTTP")
    
    objHTTP.Open "GET", chkURL, False
    objHTTP.send
    sHTML = objHTTP.statusText
    
    'some browser retunr 'Ok'
    If Not (Err Or UCase(sHTML) <> "OK") Then IsValidURL = True
    
    Set objHTTP = Nothing
End Function

Public Function IsVowle(str As String) As Boolean
    If Trim(str) = "" Then
        IsVowle = False
    Else
        IsVowle = InStr("AaEeIiOoUu", Mid(str, 1, 1))
    End If

End Function

Public Sub SwitchTranFlag(strSQLString As String)
    'when in transation control mode set to in transaction, this function is called
    'it search for parameter tran_flg, if found and value is 'Y', change it to 'N'
    Dim strParameter As String
    Dim arrParameter() As String
    Dim lngLeftBracketPos As Long
    Dim lngRightBracketPos As Long
    Dim idx As Integer
    Dim blnFound As Boolean
    
    If Trim(strSQLString) = "" Then Exit Sub
    
    lngLeftBracketPos = InStr(strSQLString, "(")
    lngRightBracketPos = InStrRev(strSQLString, ")")
    If lngLeftBracketPos = 0 Or lngRightBracketPos = 0 Then Exit Sub
    
    strParameter = Mid(strSQLString, lngLeftBracketPos + 1, lngRightBracketPos - lngLeftBracketPos - 1)

    arrParameter = Split(strParameter, ",")
    
    For idx = 0 To UBound(arrParameter)
        If UCase(Replace(arrParameter(idx), " ", "")) = "TRAN_FLG='Y'" Then
            arrParameter(idx) = "tran_flg = 'N'"
            blnFound = True
            'should be no more than one
            Exit For
        End If
    Next
    
    If Not blnFound Then Exit Sub
    
    'rebuld the parameter string
    strParameter = ""
    For idx = 0 To UBound(arrParameter)
        If strParameter = "" Then
            strParameter = arrParameter(idx)
        Else
            strParameter = strParameter & "," & arrParameter(idx)
        End If
    Next
            
    'rebuild the SQL
    strSQLString = Mid(strSQLString, 1, lngLeftBracketPos) & strParameter & _
        Mid(strSQLString, lngRightBracketPos, Len(strSQLString) - lngRightBracketPos + 1)
End Sub

'***Patrick add function to handle copy all, copy selection and print
'*** all and print select for vsf control
Public Sub vsfCopyAll(vsf As Object, DisTitle As String, _
    Optional HiddenRow As Boolean = True, Optional HiddenCol As Boolean = True, _
    Optional ssb As Object, Optional frm As Object)
    
    On Error GoTo errHandler
    
    Dim SelectString As String
    Dim idxRow As Long
    Dim idxCol As Integer
    Dim blnFirstCell As Boolean
    Dim strCellContent As String
    
    If Not frm Is Nothing Then
        frm.MousePointer = vbHourglass
    End If
    If Not ssb Is Nothing Then
        ssb.Panels(1) = "Copying all to clipboard..."
    End If
    
    blnFirstCell = True
    With vsf
        'get all row not hidden
        For idxRow = 0 To .Rows - 1
            If Not .RowHidden(idxRow) Then
                For idxCol = 0 To .Cols - 1
                    If Not .ColHidden(idxCol) Then
                        strCellContent = Replace(.TextMatrix(idxRow, idxCol), vbNewLine, " ")
                        strCellContent = Replace(strCellContent, vbLf, " ")
                        strCellContent = Replace(strCellContent, vbCr, " ")
                        
                        If blnFirstCell Then
                            SelectString = SelectString & strCellContent
                            blnFirstCell = False
                        Else
                            SelectString = SelectString & vbTab & strCellContent
                        End If
                    End If
                Next
                
                If HiddenCol Then
                    For idxCol = 0 To .Cols - 1
                        If .ColHidden(idxCol) Then
                            strCellContent = Replace(.TextMatrix(idxRow, idxCol), vbNewLine, " ")
                            strCellContent = Replace(strCellContent, vbLf, " ")
                            strCellContent = Replace(strCellContent, vbCr, " ")
                            SelectString = SelectString & vbTab & strCellContent
                        End If
                    Next
                End If
                
                SelectString = SelectString & vbNewLine
                blnFirstCell = True
            End If
        Next
        
        If HiddenRow Then
            For idxRow = 0 To .Rows - 1
                If .RowHidden(idxRow) Then
                    For idxCol = 0 To .Cols - 1
                        If Not .ColHidden(idxCol) Then
                            strCellContent = Replace(.TextMatrix(idxRow, idxCol), vbNewLine, " ")
                            strCellContent = Replace(strCellContent, vbLf, " ")
                            strCellContent = Replace(strCellContent, vbCr, " ")
                            
                            If blnFirstCell Then
                                SelectString = SelectString & strCellContent
                                blnFirstCell = False
                            Else
                                SelectString = SelectString & vbTab & strCellContent
                            End If
                        End If
                    Next
                    
                    If HiddenCol Then
                        For idxCol = 0 To .Cols - 1
                            If .ColHidden(idxCol) Then
                                strCellContent = Replace(.TextMatrix(idxRow, idxCol), vbNewLine, " ")
                                strCellContent = Replace(strCellContent, vbLf, " ")
                                strCellContent = Replace(strCellContent, vbCr, " ")
                                SelectString = SelectString & vbTab & strCellContent
                            End If
                        Next
                    End If
                    
                    SelectString = SelectString & vbNewLine
                    blnFirstCell = True
                End If
            Next
        End If
    End With
    
    Clipboard.Clear
    Clipboard.SetText SelectString
    
    If Not frm Is Nothing Then
        frm.MousePointer = vbDefault
    End If
    If Not ssb Is Nothing Then
        ssb.Panels(1) = "All copied to clipboard."
    End If
Exit Sub

errHandler:
    If MsgBox("Failed to copy to the clipboard.", vbExclamation + vbRetryCancel, DisTitle) = vbRetry Then
        Resume
    End If
    If Not frm Is Nothing Then
        frm.MousePointer = vbDefault
    End If
End Sub

Public Sub vsfCopySelection(vsf As Object, DisTitle As String, _
    Optional HiddenCol As Boolean = True, Optional ssb As Object, _
    Optional frm As Object)
    
    On Error GoTo errHandler
    
    Dim SelectString As String
    Dim idxRow As Long
    Dim idxCol As Integer
    Dim blnFirstCell As Boolean
    Dim strCellContent As String
    
    If Not frm Is Nothing Then
        frm.MousePointer = vbDefault
    End If
    
    If Not ssb Is Nothing Then
        ssb.Panels(1) = "Copying selection to clipboard..."
    End If
    
    blnFirstCell = True
    With vsf
        If .SelectedRows = 0 Then
            MsgBox "You must select some row.", vbInformation + vbOKOnly, DisTitle
        Else
            'column header
            For idxRow = 0 To .FixedRows - 1
                For idxCol = 0 To .Cols - 1
                    If Not .ColHidden(idxCol) Then
                        strCellContent = Replace(.TextMatrix(idxRow, idxCol), vbNewLine, " ")
                        strCellContent = Replace(strCellContent, vbLf, " ")
                        strCellContent = Replace(strCellContent, vbCr, " ")
                        
                        If blnFirstCell Then
                            SelectString = SelectString & strCellContent
                            blnFirstCell = False
                        Else
                            SelectString = SelectString & vbTab & strCellContent
                        End If
                    End If
                Next
                
                If HiddenCol Then
                    For idxCol = 0 To .Cols - 1
                        If .ColHidden(idxCol) Then
                            strCellContent = Replace(.TextMatrix(idxRow, idxCol), vbNewLine, " ")
                            strCellContent = Replace(strCellContent, vbLf, " ")
                            strCellContent = Replace(strCellContent, vbCr, " ")
                            
                            SelectString = SelectString & vbTab & strCellContent
                        End If
                    Next
                End If
                
                SelectString = SelectString & vbNewLine
                blnFirstCell = True
            Next
            
            For idxRow = 0 To .SelectedRows - 1
                For idxCol = 0 To .Cols - 1
                    If Not .ColHidden(idxCol) Then
                        strCellContent = Replace(.TextMatrix(.SelectedRow(idxRow), idxCol), vbNewLine, " ")
                        strCellContent = Replace(strCellContent, vbLf, " ")
                        strCellContent = Replace(strCellContent, vbCr, " ")
                        
                        If blnFirstCell Then
                            SelectString = SelectString & strCellContent
                            blnFirstCell = False
                        Else
                            SelectString = SelectString & vbTab & strCellContent
                        End If
                    End If
                Next
                
                If HiddenCol Then
                    For idxCol = 0 To .Cols - 1
                        If .ColHidden(idxCol) Then
                            strCellContent = Replace(.TextMatrix(.SelectedRow(idxRow), idxCol), vbNewLine, " ")
                            strCellContent = Replace(strCellContent, vbLf, " ")
                            strCellContent = Replace(strCellContent, vbCr, " ")
                            
                            SelectString = SelectString & vbTab & strCellContent
                        End If
                    Next
                End If
                
                SelectString = SelectString & vbNewLine
                blnFirstCell = True
            Next
            
            Clipboard.Clear
            Clipboard.SetText SelectString
            
            If Not ssb Is Nothing Then
                ssb.Panels(1) = "Selection copied to clipboard."
            End If
        End If
    End With
    If Not frm Is Nothing Then
        frm.MousePointer = vbDefault
    End If
    
Exit Sub
errHandler:
    If MsgBox("Failed to copy to the clipboard.", vbExclamation + vbRetryCancel, DisTitle) = vbRetry Then
        Resume
    End If
    If Not frm Is Nothing Then
        frm.MousePointer = vbDefault
    End If
End Sub

Public Sub vsfPrintAll(vsf As Object, Optional ByVal DocName As Variant, _
    Optional ByVal ShowDialog As Variant, Optional ByVal Orientation As Variant, _
    Optional ByVal MarginLR As Variant, Optional ByVal MarginTB As Variant, _
    Optional ssb As Object, Optional frm As Object)
    
    On Error GoTo errHandler
    
    If Not frm Is Nothing Then
        frm.MousePointer = vbDefault
    End If
    
    If Not ssb Is Nothing Then
        ssb.Panels(1) = "Printing all..."
    End If
    
    vsf.PrintGrid DocName, ShowDialog, Orientation, MarginLR, MarginTB
    
    If Not frm Is Nothing Then
        frm.MousePointer = vbDefault
    End If
    If Not ssb Is Nothing Then
        ssb.Panels(1) = "All printed."
    End If

Exit Sub
errHandler:
    If Not frm Is Nothing Then
        frm.MousePointer = vbDefault
    End If
    MsgBox "Error occur when print all.", vbInformation + vbOKOnly, DocName
End Sub

Public Sub vsfPrintSelection(vsf As Object, Optional ByVal DocName As Variant, _
    Optional ByVal ShowDialog As Variant, Optional ByVal Orientation As Variant, _
    Optional ByVal MarginLR As Variant, Optional ByVal MarginTB As Variant, _
    Optional ssb As Object, Optional frm As Object)
    
    On Error GoTo errHandler
    
    Dim idxRow As Long
    Dim hl As Integer
    Dim tr As Long
    Dim lc As Long
    Dim rd As Integer
     
    If Not frm Is Nothing Then
        frm.MousePointer = vbDefault
    End If
    
    If Not ssb Is Nothing Then
        ssb.Panels(1) = "Printing selection..."
    End If
     
    'note this function has a little problem
    'when the total was selected with another record(s)
    'even the selected records are not hidden
    'the printgrid method only print the header and the total only
    
    With vsf
        If .SelectedRows = 0 Then
            MsgBox "You must select some row.", vbInformation + vbOKOnly, DocName
        Else
            ' save current settings
            hl = .HighLight
            tr = .TopRow
            lc = .LeftCol
            rd = .Redraw
            .HighLight = 0
            .Redraw = 0 'flexRDNone
    
            'hide non-selected rows
            .RowHidden(-1) = True
            For idxRow = 0 To .FixedRows
                .RowHidden(idxRow) = False
            Next
            For idxRow = 0 To .SelectedRows - 1
                .RowHidden(.SelectedRow(idxRow)) = False
            Next
            ' scroll to top left corner
            .TopRow = .FixedRows
            .LeftCol = .FixedCols
            
            .PrintGrid DocName, ShowDialog, Orientation, MarginLR, MarginTB
            ' restore control
            .RowHidden(-1) = False
            .TopRow = tr
            .LeftCol = lc
            .HighLight = hl
            .Redraw = rd
        End If
    End With
    If Not frm Is Nothing Then
        frm.MousePointer = vbDefault
    End If
    If Not ssb Is Nothing Then
        ssb.Panels(1) = "Selection printed."
    End If
    
Exit Sub
errHandler:
    If Not frm Is Nothing Then
        frm.MousePointer = vbDefault
    End If
    MsgBox "Error occur when print selection.", vbInformation + vbOKOnly, DocName
    
End Sub
'*** end of copy and print

Public Function IPAddressInRange(ip_address As String, ip_address_range As String) As Boolean
  ' Determines if ip_address is in ip_address_range
  ' ip_address assumed to be in format x.x.x.x.  e.g. 10.1.30.20
  ' ip_address_range assumed to be in format x.x.x.x or x.x.x.x-x.x.x.x.  e.g. 10.1.30.20 or
  ' 10.1.30.1-10.1.30.255
  'Scott R
  Dim start_ip_address As String
  Dim end_ip_address As String
  Dim b1 As Byte, b2 As Byte, b3 As Byte, b4 As Byte
  Dim start_b1 As Byte, start_b2 As Byte, start_b3 As Byte, start_b4 As Byte
  Dim end_b1 As Byte, end_b2 As Byte, end_b3 As Byte, end_b4 As Byte
  Dim pos As Long
  
  ' Assume it's not in the range
  IPAddressInRange = False
  
  ' See if the range is indeed a range
  pos = InStr(ip_address_range, "-")
  
  ' Assumed be a single address, so can direct compare
  If (pos = 0) Then
    If (ip_address = ip_address_range) Then
      IPAddressInRange = True
    End If
  
  ' Assumed be a range, so need to parse addresses to compare
  Else
    start_ip_address = Left(ip_address_range, pos - 1)
    end_ip_address = Mid(ip_address_range, pos + 1)
    
    If (SplitIPAddress(ip_address, b1, b2, b3, b4)) Then
      If (SplitIPAddress(start_ip_address, start_b1, start_b2, start_b3, start_b4)) Then
        If (SplitIPAddress(end_ip_address, end_b1, end_b2, end_b3, end_b4)) Then
          If (start_b1 <= b1) And (b1 <= end_b1) And _
             (start_b2 <= b2) And (b2 <= end_b2) And _
             (start_b3 <= b3) And (b3 <= end_b3) And _
             (start_b4 <= b4) And (b4 <= end_b4) Then
            IPAddressInRange = True
          End If
        End If
      End If
    End If
  End If

End Function

Public Function SplitIPAddress(ByVal ip_address As String, ByRef b1 As Byte, ByRef b2 As Byte, _
                        ByRef b3 As Byte, ByRef b4 As Byte)
  'Scott R
  Dim pos As Integer
  Dim s As String
  
  SplitIPAddress = False
  
  ' Extract Byte 1
  pos = InStr(ip_address, ".")
  
  If (pos = 0) Then
      Exit Function
  Else
    s = Left(ip_address, pos - 1)
    ip_address = Mid(ip_address, pos + 1)
    
    If (Not IsNumeric(s)) Then
      Exit Function
    End If
    
    b1 = s
  End If

  ' Extract Byte 2
  pos = InStr(ip_address, ".")
  
  If (pos = 0) Then
      Exit Function
  Else
    s = Left(ip_address, pos - 1)
    ip_address = Mid(ip_address, pos + 1)
    
    If (Not IsNumeric(s)) Then
      Exit Function
    End If
    
    b2 = s
  End If

  ' Extract Byte 3
  pos = InStr(ip_address, ".")
  
  If (pos = 0) Then
      Exit Function
  Else
    s = Left(ip_address, pos - 1)
    ip_address = Mid(ip_address, pos + 1)
    
    If (Not IsNumeric(s)) Then
      Exit Function
    End If
    
    b3 = s
  End If
  
  ' Extract Byte 4
  If (Not IsNumeric(ip_address)) Then
    Exit Function
  End If
  
  b4 = ip_address
  
  SplitIPAddress = True

End Function

Public Function RemoverInvalidCharacters(strIn As String) As String
    Dim idx As Integer
    
    For idx = 0 To 31
        strIn = Replace(strIn, Chr(idx), "")
    Next
    
    For idx = 128 To 255
        strIn = Replace(strIn, Chr(idx), "")
    Next
    
    RemoverInvalidCharacters = strIn
End Function

Function sf_implode(colFields As Collection, ByVal strDelim As String, ByVal strEncapsulateChar) As String
    'mirror the equivalent SP sf_implode
    'Code by Scott Robertson
    Dim strData As String
    Dim intIndex As Long
    Dim strField As Variant
    
    sf_implode = ""
    strData = ""
    
    If colFields.Count = 0 Then
      Exit Function
    End If
    
    ' Look through each rcord in the collection
    For Each strField In colFields
        ' Convert single occurances of strEncapsulateChar to double occurances
        strField = Replace(strField, strEncapsulateChar, strEncapsulateChar & strEncapsulateChar)
        
        ' If there's any occurances of p_encapsulate_char or p_delimiter, quote the string using p_encapsulate_char
        If (InStr(strField, strDelim) > 0) Or (InStr(strField, strEncapsulateChar) > 0) Then
          strField = strEncapsulateChar & strField & strEncapsulateChar
        End If
        
        ' Add delimter after previous record
        If (strData <> "") Then
          strData = strData & strDelim
        End If
        
        ' Add the field
        strData = strData & strField
    Next
    
    sf_implode = strData
  
End Function

Function sf_explode(ByVal strData As String, ByVal strDelim As String, ByVal strEncapsulateChar) As String()
    'mirror the equivalent SP sf_explode
    'Code by Scott Robertson
    Dim arrData() As String
    Dim intPos As Long
    
    Dim strChar As String
    Dim strField As String
    
    Dim strNextChar As String
    Dim boolEncapsulated As Boolean
    Dim boolAtEndOfField As Boolean
    
    Dim intFieldNo
  
    ' Initialise
    intFieldNo = 0
    intPos = 1
    boolAtEndOfField = False
    boolEncapsulated = False
    strField = ""
      
    ' Add a delimiter to the data to make it easier to process
    strData = strData & strDelim
      
    ' Loop through the string
    While intPos <= Len(strData)
      
        strChar = Mid(strData, intPos, 1)
    
        If (intPos < Len(strData)) Then
            strNextChar = Mid(strData, intPos + 1, 1)
        Else
            strNextChar = ""
            ' Force an end to the encapsulation.
            ' This is designed for the case where a field starts with a quote but is not correctly ended - e.g.
            boolEncapsulated = False
        End If
    
        ' Whilst still at the start of the field, check to see if we have an
        If (intPos = 1) Then
        ' String is encapsulated
            If (strChar = strEncapsulateChar) Then
                boolEncapsulated = True
                strChar = ""
            End If
    
        End If
    
        ' We have a character to process
        If (strChar <> "") Then
            ' Processing without encapsulation
            If (Not boolEncapsulated) Then
                ' A delimiter indicates the end of the record
                If (strChar = strDelim) Then
                    boolAtEndOfField = True
        
                ' Anything else is a regular character to add to the field
                Else
                  strField = strField & strChar
                End If

            ' Processing with encapsulation
            Else
                ' Quote needs special handling
                If (strChar = strEncapsulateChar) Then
                    ' Two quotes in a row get replaced with a single quote
                    If (strNextChar = strEncapsulateChar) Then
                        ' Skip the first quote
                        intPos = intPos + 1
                        strField = strField & strChar
                      
                    ' A Quote followed by anything else is deemed to indicate the end of
                    ' encapsulation.  Normally we'd be expecting a delimiter but it
                    '  doesn't matter what we actually get...
                    Else
                        boolEncapsulated = False
          
                    End If
          
                ' Anything else is a regular character to add to the field
                Else
                    strField = strField & strChar
                End If
            End If
        End If
    
        ' Add the completed field to the array
        If (boolAtEndOfField) Then
            intFieldNo = intFieldNo + 1
            ReDim Preserve arrData(intFieldNo)
      
            arrData(intFieldNo) = strField
            
            ' Skip past the field just found and returned
            strData = Mid(strData, intPos + 1, Len(strData))
            
            ' Reset everything in order to process the next field
            intPos = 0
            boolAtEndOfField = False
            boolEncapsulated = False
            strField = ""
      
        End If
    
        ' Skip to the next character
        intPos = intPos + 1

    Wend
  
    sf_explode = arrData
  
End Function

Public Function HasNonPrintableChar(str As String) As Boolean
    'check for string contains non-printable cahracters (32-127 condsider as valid)
    Dim idx As Integer
    
    For idx = 1 To Len(str)
        If Asc(Mid(str, idx, 1)) < 32 Or Asc(Mid(str, idx, 1)) > 127 Then
            HasNonPrintableChar = True
            Exit Function
        End If
    Next
End Function
