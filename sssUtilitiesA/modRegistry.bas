Attribute VB_Name = "modRegistry"
Option Explicit

' Registry.bas
'Version History
'10/11/97 HR V1.0.2
'   - Created this module
'11/11/97 HR V1.1.1
'   - QueryValueEx() was returning string which was one character too long

Private Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type

Private Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Global Const REG_SZ As Long = 1
Global Const REG_EXPAND_SZ = 2
Global Const REG_BINARY = 3
Global Const REG_DWORD As Long = 4
Global Const REG_MULTI_SZ = 7

Global Const HKEY_CLASSES_ROOT = &H80000000
Global Const HKEY_CURRENT_USER = &H80000001
Global Const HKEY_LOCAL_MACHINE = &H80000002
Global Const HKEY_USERS = &H80000003
Global Const HKEY_PERFORMANCE_DATA = &H80000004
Global Const HKEY_CURRENT_CONFIG = &H80000005
Global Const HKEY_DYN_DATA = &H80000006

Global Const ERROR_NONE = 0
Global Const ERROR_BADDB = 1
Global Const ERROR_BADKEY = 2
Global Const ERROR_CANTOPEN = 3
Global Const ERROR_CANTREAD = 4
Global Const ERROR_CANTWRITE = 5
Global Const ERROR_OUTOFMEMORY = 6
Global Const ERROR_INVALID_PARAMETER = 7
Global Const ERROR_ACCESS_DENIED = 8
Global Const ERROR_INVALID_PARAMETERS = 87
Global Const ERROR_NO_MORE_ITEMS = 259

Global Const KEY_ALL_ACCESS = &H3F
Global Const REG_OPTION_NON_VOLATILE = 0

Global Const REG_OPTION_VOLATILE = 1           ' Key is not preserved when system is rebooted
'Const REG_OPTION_NON_VOLATILE = 0       ' Key is preserved when system is rebooted                                  ' KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Global Const SYNCHRONIZE = &H100000
Global Const READ_CONTROL = &H20000
Global Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Global Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Global Const STANDARD_RIGHTS_ALL = &H1F0000
Global Const KEY_QUERY_VALUE = &H1
Global Const KEY_SET_VALUE = &H2
Global Const KEY_CREATE_SUB_KEY = &H4
Global Const KEY_ENUMERATE_SUB_KEYS = &H8
Global Const KEY_NOTIFY = &H10
Global Const KEY_CREATE_LINK = &H20
Global Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Global Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Global Const KEY_EXECUTE = (KEY_READ)
'Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Global Const ERROR_MORE_DATA = 234
'Const ERROR_NO_MORE_ITEMS = &H103
Global Const ERROR_KEY_NOT_FOUND = &H2

Dim Security As SECURITY_ATTRIBUTES

Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hkey As Long) As Long

Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias _
"RegCreateKeyExA" (ByVal hkey As Long, ByVal lpSubKey As String, _
ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions _
As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes _
As Long, phkResult As Long, lpdwDisposition As Long) As Long

Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias _
"RegOpenKeyExA" (ByVal hkey As Long, ByVal lpSubKey As String, _
ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As _
Long) As Long

Declare Function RegQueryValueExString Lib "advapi32.dll" Alias _
"RegQueryValueExA" (ByVal hkey As Long, ByVal lpValueName As _
String, ByVal lpReserved As Long, lpType As Long, ByVal lpData _
As String, lpcbData As Long) As Long

Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias _
"RegQueryValueExA" (ByVal hkey As Long, ByVal lpValueName As _
String, ByVal lpReserved As Long, lpType As Long, lpData As _
Long, lpcbData As Long) As Long

Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias _
"RegQueryValueExA" (ByVal hkey As Long, ByVal lpValueName As _
String, ByVal lpReserved As Long, lpType As Long, ByVal lpData _
As Long, lpcbData As Long) As Long

Declare Function RegSetValueExString Lib "advapi32.dll" Alias _
"RegSetValueExA" (ByVal hkey As Long, ByVal lpValueName As String, _
ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As _
String, ByVal cbData As Long) As Long

Declare Function RegSetValueExLong Lib "advapi32.dll" Alias _
"RegSetValueExA" (ByVal hkey As Long, ByVal lpValueName As String, _
ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, _
ByVal cbData As Long) As Long

Private Declare Function ExpandEnvironmentStrings Lib "kernel32" Alias "ExpandEnvironmentStringsA" (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long
Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hkey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hkey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hkey As Long, ByVal lpSubKey As String) As Long
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hkey As Long, ByVal lpValueName As String) As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, Source As Any, ByVal numBytes As Long)
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

' SetValueEx and QueryValueEx Wrapper Functions
Public Function SetValueEx(ByVal hkey As Long, sValueName As String, _
                            lType As Long, vValue As Variant) As Long
    Dim lValue As Long
    Dim sValue As String
    Select Case lType
        Case REG_SZ
            sValue = vValue & Chr$(0)
            SetValueEx = RegSetValueExString(hkey, sValueName, 0&, _
                                           lType, sValue, Len(sValue))
        Case REG_DWORD
            lValue = vValue
            SetValueEx = RegSetValueExLong(hkey, sValueName, 0&, _
                lType, lValue, 4)
    End Select
End Function

Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As _
        String, vValue As Variant) As Long

    Dim nch As Long
    Dim lrc As Long
    Dim lType As Long
    Dim lValue As Long
    Dim sValue As String

    On Error GoTo QueryValueExError

    ' Determine the size and type of data to be read
    lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, nch)
    If lrc <> ERROR_NONE Then Error 5

    Select Case lType
        ' For strings
        Case REG_SZ:
            sValue = String(nch, 0)     ' pad string with nulls
            lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, nch)

            If lrc = ERROR_NONE Then
                If nch = 0 Then nch = 1
                vValue = Left$(sValue, nch - 1) ' convert zero terminated string
            Else
                vValue = Empty
            End If
        ' For DWORDS
        Case REG_DWORD:
            lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, lValue, nch)
            If lrc = ERROR_NONE Then vValue = lValue
        Case Else
            'all other data types not supported
            lrc = -1
    End Select

QueryValueExExit:

    QueryValueEx = lrc
    Exit Function

QueryValueExError:
    Resume QueryValueExExit
End Function


' ================  PUBLIC FUNCTIONS ================================


Public Sub CreateNewKey(lPredefinedKey As Long, sNewKeyName As String)
' Example CreateNewKey HKEY_CURRENT_USER, "TestKey\SubKey1\SubKey2"
       
    Dim hNewKey As Long         'handle to the new key
    Dim lRetVal As Long         'result of the RegCreateKeyEx function

    lRetVal = RegCreateKeyEx(lPredefinedKey, sNewKeyName, 0&, _
                 vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, _
                 0&, hNewKey, lRetVal)
    RegCloseKey (hNewKey)
End Sub

Public Sub SetKeyValue(lPredefinedKey As Long, sKeyName As String, sValueName As String, _
            vValueSetting As Variant, lValueType As Long)
' Example SetKeyValue HK_LOCAL_MACHINE, "TestKey\SubKey1", "StringValue", "Hello", REG_SZ
    
    Dim lRetVal As Long         'result of the SetValueEx function
    Dim hkey As Long            'handle of open key

    'open the specified key. Use RegCreateKey if want to create if doesn't exist.
    ' lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
    RegCreateKeyEx lPredefinedKey, sKeyName, 0&, _
                vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, _
                0&, hkey, lRetVal
    lRetVal = SetValueEx(hkey, sValueName, lValueType, vValueSetting)
    RegCloseKey (hkey)
End Sub

Public Function QueryValue(lPredefinedKey As Long, sKeyName As String, _
        sValueName As String) As Variant
    
    Dim lRetVal As Long         'result of the API functions
    Dim hkey As Long            'handle of opened key
    Dim vValue As Variant       'setting of queried value

    QueryValue = ""             ' default if key not found
    lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, _
                    KEY_ALL_ACCESS, hkey)
    lRetVal = QueryValueEx(hkey, sValueName, vValue)
    RegCloseKey (hkey)
    QueryValue = vValue

End Function

'****************************************************************************
'       Check to see if Registry key exists
'       Inputs: None
'       Class Properties: Classname.hkey, Classname.keyroot, Classname.subkey
'       Return: True if key exists
'****************************************************************************
Public Function KeyExists(lPredefinedKey As Long, sKeyName As String) As Boolean
    Dim handle As Long
    Dim ret As Long
    
    If RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_READ, handle) Then
        KeyExists = False
        Exit Function
    End If
    KeyExists = True
End Function

'****************************************************************************
'       Create a key in the registry
'       Inputs: KeyName
'       Class Properties: if Input Empty Classname.subkey
'       Return: 0 if successful
'****************************************************************************
Public Function CreateKey(lPredefinedKey As Long, sKeyName As String) As String
    Dim handle As Long
    Dim disp As Long
    Dim RetVal As Long
    
    RetVal = RegCreateKeyEx(lPredefinedKey, sKeyName, 0, "", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, handle, disp)
    If RetVal Then Exit Function
    RegCloseKey (handle)
    CreateKey = RetVal
End Function

'****************************************************************************
'       Delete a key from the registry
'       Inputs: SubKey
'       Class Properties: Classname.hkey, Classname.keyroot
'       Returns: 0 if successful
'****************************************************************************
Public Function DeleteKey(lPredefinedKey As Long, sKeyName As String) As Long
    Dim RetVal As Long
    Dim handle As Long
    
    RetVal = RegDeleteKey(lPredefinedKey, sKeyName)
    If RetVal Then Exit Function
    RegCloseKey (handle)
    DeleteKey = RetVal
End Function
'****************************************************************************
'       Delete the value of a key
'       Inputs: Value Name
'       Class Properties: Classname.hkey, Classname.keyroot, Classname.subkey
'       Return: 0 if successful
'****************************************************************************
Public Function DeleteValue(lPredefinedKey As Long, sKeyName As String, sValueName As String) As Long
    Dim RetVal As Long
    Dim handle As Long
    
    RetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, handle)
    If RetVal <> 0 Then 'Operation Failed
        DeleteValue = RetVal
        Exit Function
    End If
    DeleteValue = RegDeleteValue(handle, sValueName)
    RegCloseKey (handle)
End Function

'****************************************************************************
'       Enumerate Value Names under a given key
'       Inputs: Key Root, Key Name
'       Return: a collection of strings
'       Source: Slightly modified from www.vb2themax.com EnumRegistryKeys
'****************************************************************************
Public Function EnumRegistryKeys(lPredefinedKey As Long, sKeyName As String) As _
                Collection
    Dim handle As Long
    Dim length As Long
    Dim Index As Long
    Dim subkeyName As String
    Dim fFiletime As FILETIME
         ' initialize the result collection
         Set EnumRegistryKeys = New Collection
         
         ' Open the key, exit if not found
         If Len(sKeyName) Then
             If RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_READ, handle) Then Exit Function
             ' in all case the subsequent functions use hKey
             lPredefinedKey = handle
         End If
         
         Do
             ' this is the max length for a key name
             length = 260
             subkeyName = Space$(length)
             ' get the N-th key, exit the loop if not found
             If RegEnumKeyEx(lPredefinedKey, Index, subkeyName, length, 0, "", vbNull, fFiletime) = ERROR_NO_MORE_ITEMS Then Exit Do
             ' add to the result collection
             subkeyName = Left$(subkeyName, InStr(subkeyName, vbNullChar) - 1)
             EnumRegistryKeys.Add subkeyName, subkeyName
             ' prepare to query for next key
             Index = Index + 1
         Loop
        
         ' Close the key, if it was actually opened
         If handle Then RegCloseKey handle
        
End Function

'****************************************************************************
'       Enumerate values under a given registry key
'       Inputs: Key Root, Key Name
'       Return: a collection, where each element of the collection
'               is a 2-element array of Variants:
'               element(0) is the value name, element(1) is the value's value
'       Source: Slightly Modified from www.vb2themax.com EnumRegistryValues
'****************************************************************************
Function EnumRegistryValues(lPredefinedKey As Long, sKeyName As String) As _
    Collection
    Dim handle As Long
    Dim Index As Long
    Dim valueType As Long
    Dim Name As String
    Dim nameLen As Long
    Dim resLong As Long
    Dim resString As String
    Dim length As Long
    Dim valueInfo(0 To 1) As Variant
    Dim RetVal As Long
    Dim i As Integer
    Dim vTemp As Variant
    
    ' initialize the result
    Set EnumRegistryValues = New Collection
    
    ' Open the key, exit if not found.
    If Len(sKeyName) Then
        If RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_READ, handle) Then Exit Function
        ' in all cases, subsequent functions use hKey
        lPredefinedKey = handle
    End If
    
    Do
        ' this is the max length for a key name
        nameLen = 260
        Name = Space$(nameLen)
        ' prepare the receiving buffer for the value
        length = 4096
        ReDim resBinary(0 To length - 1) As Byte
        
        ' read the value's name and data
        ' exit the loop if not found
        RetVal = RegEnumValue(lPredefinedKey, Index, Name, nameLen, ByVal 0&, valueType, _
            resBinary(0), length)
        
        ' enlarge the buffer if you need more space
        If RetVal = ERROR_MORE_DATA Then
            ReDim resBinary(0 To length - 1) As Byte
            RetVal = RegEnumValue(lPredefinedKey, Index, Name, nameLen, ByVal 0&, _
                valueType, resBinary(0), length)
        End If
        ' exit the loop if any other error (typically, no more values)
        If RetVal Then Exit Do
        
        ' retrieve the value's name
        valueInfo(0) = Left$(Name, nameLen)
        
        ' return a value corresponding to the value type
        Select Case valueType
            
            Case REG_DWORD
                CopyMemory resLong, resBinary(0), 4
                valueInfo(1) = resLong
            
            Case REG_SZ
                ' copy everything but the trailing null char
                If length <> 0 Then
                    resString = Space$(length - 1)
                    CopyMemory ByVal resString, resBinary(0), length - 1
                    valueInfo(1) = resString
                Else
                    valueInfo(1) = ""
                End If
                
            Case REG_EXPAND_SZ
                ' copy everything but the trailing null char
                ' expand the environment variable to it's value
                ' Ignore a Blank String
                If length <> 0 Then
                    resString = Space$(length - 1)
                    CopyMemory ByVal resString, resBinary(0), length - 1
                    length = ExpandEnvironmentStrings(resString, resString, Len(resString))
                    valueInfo(1) = TrimNull(resString)
                Else
                    valueInfo(1) = ""
                End If

            Case REG_BINARY
                ' shrink the buffer if necessary
                If length < UBound(resBinary) + 1 Then
                    ReDim Preserve resBinary(0 To length - 1) As Byte
                End If
                 'Convert to display as string like this: 00 01 01 00 01
                    For i = 0 To UBound(resBinary)
                         resString = resString & " " & Format(Trim(Hex(resBinary(i))), "0#")
                    Next i
                    valueInfo(1) = LTrim(resString) 'Get rid of leading space
            
            Case REG_MULTI_SZ
                ' copy everything but the 2 trailing null chars
                resString = Space$(length - 2)
                CopyMemory ByVal resString, resBinary(0), length - 2
                
                'convert from null-delimited (vbNullChar) stream of strings
                'to comma delimited stream of strings
                'The listview likes it better that way
                resString = Replace(resString, vbNullChar, ",", , , vbBinaryCompare)
                valueInfo(1) = resString
            
            Case Else
                ' Unsupported value type - do nothing
        End Select
        
        ' add the array to the result collection
        ' the element's key is the value's name
        EnumRegistryValues.Add valueInfo, valueInfo(0)
        
        Index = Index + 1
    Loop
   
    ' Close the key, if it was actually opened
    If handle Then RegCloseKey handle
        
End Function

'****************************************************************************
' Trim to first Null character
' Inputs: String with null characaters
' Return: String up to where first null character occured
'****************************************************************************
Public Function TrimNull(item As String) As String
    Dim pos As Integer
        pos = InStr(item, Chr$(0))
        If pos Then item = Left$(item, pos - 1)
        TrimNull = item
End Function

'****************************************************************************
'       Read a Registry value
'
'       Inputs: Use KeyName = "" for the Default value
'                If the value isn't there, it returns the DefaultValue
'                argument passed in, or Empty if the argument has been omitted
'       Return: Variant
'
'               REG_DWORD: Long
'               REG_SZ: String
'               REG_EXPAND_SZ: String with Expanded Environment variable
'               REG_BINARY: Byte Array
'               REG_MULTI_SZ: null-delimited (vbNullChar) stream of strings
'                   (VB6 users can use Split to convert to an array of string)
'                    Split(expression[, delimiter[, count[, compare]]])
'       Source: Slightly modified from www.vb2themax GetRegistryValue
'****************************************************************************
Public Function GetRegistryValue(lPredefinedKey As Long, sKeyName As String, ByVal ValueName As String, Optional DefaultValue As Variant) As Variant
    Dim handle As Long
    Dim resLong As Long
    Dim resString As String
    Dim TestString As String
    Dim resBinary() As Byte
    Dim length As Long
    Dim RetVal As Long
    Dim valueType As Long
    
        ' Prepare the default result
        GetRegistryValue = IIf(IsMissing(DefaultValue), Empty, DefaultValue)
        
        ' Open the key, exit if not found.
        If RegOpenKeyEx(lPredefinedKey, sKeyName, REG_OPTION_NON_VOLATILE, KEY_READ, handle) Then
           'Don 't overwrite the default value!
           'GetRegistryValue = CVar("Error!")
           Exit Function
        End If
        
        ' prepare a 1K receiving resBinary
        length = 1024
        ReDim resBinary(0 To length - 1) As Byte
        
        ' read the registry key
        RetVal = RegQueryValueEx(handle, ValueName, 0, valueType, resBinary(0), _
            length)
        ' if resBinary was too small, try again
        If RetVal = ERROR_MORE_DATA Then
            ' enlarge the resBinary, and read the value again
            ReDim resBinary(0 To length - 1) As Byte
            RetVal = RegQueryValueEx(handle, ValueName, 0, valueType, resBinary(0), _
                length)
        End If
        
        'Added 11/5/01 Don Kiser
        If RetVal = ERROR_KEY_NOT_FOUND Then
                 RegCloseKey (handle)
                 Exit Function
        End If
        
        ' return a value corresponding to the value type
        Select Case valueType
            Case REG_DWORD
                CopyMemory resLong, resBinary(0), 4
                GetRegistryValue = resLong
            
            Case REG_SZ
                ' copy everything but the trailing null char
                ' Ignore Blank Strings
                If length <> 0 Then
                    resString = Space$(length - 1)
                    CopyMemory ByVal resString, resBinary(0), length - 1
                    GetRegistryValue = resString
                End If
            
            Case REG_EXPAND_SZ
                ' copy everything but the trailing null char
                ' expand the environment variable to it's value
                ' Ignore a Blank String
                If length <> 0 Then
                    resString = Space$(length - 1)
                    CopyMemory ByVal resString, resBinary(0), length - 1
                    
                    length = ExpandEnvironmentStrings(resString, resString, Len(resString))
                    GetRegistryValue = Left$(resString, length)
                    
                End If
            
            Case REG_BINARY
                ' resize the result resBinary
                If length <> UBound(resBinary) + 1 Then
                    ReDim Preserve resBinary(0 To length - 1) As Byte
                End If
                GetRegistryValue = resBinary()
            
            Case REG_MULTI_SZ
                ' copy everything but the 2 trailing null chars
                resString = Space$(length - 2)
                CopyMemory ByVal resString, resBinary(0), length - 2
                'A nonexistant value for REG_MULTI_SZ will return a string of nulls
                'with a length = 1022
                'This is because at the beginging of the routine we define Length = 1024
                ' resString = Space$(length -2) = 1022
                'So If we trims all nulls and are left with an empty string then
                'the value doesn't exist so the defualt value is returned
                'Set resstring to a temporary variable because trimnull will truncate it
                TestString = resString
                If Len(TrimNull(TestString)) > 0 Then GetRegistryValue = resString
                
            Case Else
                ' Unsupported value type - do nothing
                ' Shouldn't ever get here
        End Select
        
        ' close the registry key
     RegCloseKey (handle)
   
End Function

'****************************************************************************
'       Write or Create a Registry value
'
'       Inputs: ValueName, Value, Data Type
'       Class Properties: Classname.hkey, Classname.Keyroot, Classname.subkey
'       Return: True if successful
'
'       Use KeyName = "" for the default value
'       Supports:
'       REG_DWORD      -Integer or Long
'       REG_SZ         -String
'       REG_EXPAND_SZ  -String with Environment Variable Ex. %SystemDrive%
'       REG_BINARY     -an array of binary
'       REG_MULTI_SZ   -Null delimited String with double null terminator
'       Source: Slightly modified from www.vb2themax.com SetRegistryValue
'****************************************************************************
Public Function SetRegistryValue(lPredefinedKey As Long, sKeyName As String, ByVal ValueName As String, Value As Variant, DType As Long) As Boolean
    Dim handle As Long
    Dim lngValue As Long
    Dim strValue As String
    Dim binValue() As Byte
    Dim length As Long
    Dim RetVal As Long
    
    ' Open the key, exit if not found
    If RegOpenKeyEx(lPredefinedKey, sKeyName, REG_OPTION_NON_VOLATILE, KEY_WRITE, handle) Then
       SetRegistryValue = False 'CVar("Error!")
       Exit Function
    End If

    ' three cases, according to the data type passed
    Select Case DType
        Case REG_DWORD
            lngValue = Value
            RetVal = RegSetValueEx(handle, ValueName, 0, REG_DWORD, lngValue, 4)
        Case REG_SZ
            strValue = Value
            RetVal = RegSetValueEx(handle, ValueName, 0, REG_SZ, ByVal strValue, _
                Len(strValue))
        Case REG_BINARY
            binValue = Value
            length = UBound(binValue) - LBound(binValue) + 1
            RetVal = RegSetValueEx(handle, ValueName, 0, REG_BINARY, _
                                   binValue(LBound(binValue)), length)
        Case REG_EXPAND_SZ
            strValue = Value
            RetVal = RegSetValueEx(handle, ValueName, 0, REG_EXPAND_SZ, ByVal strValue, _
                Len(strValue))
        
        Case REG_MULTI_SZ
            strValue = Value
            RetVal = RegSetValueEx(handle, ValueName, 0, REG_MULTI_SZ, ByVal strValue, _
                Len(strValue))
        
        Case Else
            ' Unsupported value type - do nothing
            ' Shouldn't ever get here
    End Select
    
    ' Close the key and signal success
     RegCloseKey (handle)
    ' signal success if the value was written correctly
    SetRegistryValue = (RetVal = 0)
    
End Function


