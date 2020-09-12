Attribute VB_Name = "CommonAPI"
Option Explicit

'*** Win32 API Declarations
Declare Function SetWindowLong Lib "user32" Alias _
    "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long
    
Declare Function GetWindowLong Lib "user32" Alias _
    "GetWindowLongA" (ByVal hWnd As Long, _
    ByVal nIndex As Long) As Long
    
Declare Sub Sleep Lib "kernel32" (ByVal milliseconds As Long)

Declare Function GetTempFileName Lib "kernel32" Alias _
    "GetTempFileNameA" (ByVal lpszPath As String, _
    ByVal lpPrefixString As String, ByVal wUnique As Long, _
    ByVal lpTempFileName As String) As Long
    
Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal _
    nBufferLength As Long, ByVal lpBuffer As String) As Long
 
Declare Function SendMessage Lib "user32" Alias _
    "SendMessageA" (ByVal hWnd As Long, _
    ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
  
Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" _
    (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, _
    ByVal lParam As Long) As Long
 
Declare Function GetSystemMenu Lib "user32" _
    (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Declare Function GetMenuItemCount Lib "user32" _
    (ByVal hMenu As Long) As Long
Declare Function RemoveMenu Lib "user32" _
    (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare Function DrawMenuBar Lib "user32" _
    (ByVal hWnd As Long) As Long
 
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

'API Declarations and type for print object
Public Const CCHDEVICENAME = 32
Public Const CCHFORMNAME = 32
Public Const PRINTER_ACCESS_ADMINISTER = &H4
Public Const PRINTER_ACCESS_USE = &H8

Public Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, pDefault As Any) As Long
Public Declare Function ClosePrinter Lib "winspool.drv" _
        (ByVal HPrinter As Long) As Long

Public Declare Function EnumJobs Lib "winspool.drv" Alias "EnumJobsA" _
  (ByVal HPrinter As Long, ByVal FirstJob As Long, ByVal NoJobs As Long, _
   ByVal Level As Long, pJob As Any, ByVal cdBuf As Long, pcbNeeded As Long, _
   pcReturned As Long) As Long

Public Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long

Public Declare Sub CopyMem Lib "kernel32.dll" Alias "RtlMoveMemory" (pTo As Any, uFrom As Any, ByVal lSize As Long)
Public Declare Function lstrlenW Lib "kernel32.dll" (ByVal lpString As Long) As Long

Public Declare Function HeapAlloc Lib "kernel32.dll" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function GetProcessHeap Lib "kernel32.dll" () As Long
Public Declare Function HeapFree Lib "kernel32.dll" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long

Public Declare Function GetAddrOf Lib "kernel32" Alias "MulDiv" (nNumber As Any, Optional ByVal nNumerator As Long = 1, Optional ByVal nDenominator As Long = 1) As Long
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFOEX) As Long

Public Type OSVERSIONINFOEX
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128 ' Maintenance string for PSS usage
        wServicePackMajor As Integer
        wServicePackMinor As Integer
        wSuiteMask As Integer
        wProductType As Byte
        wReserved As Byte
End Type

Public Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Public Type JOB_INFO_1_API
    JobId As Long
    pPrinterName As Long
    pMachineName As Long
    pUserName As Long
    pDocument As Long
    pDatatype As Long
    pStatus As Long
    Status As Long
    Priority As Long
    Position As Long
    TotalPages As Long
    PagesPrinted As Long
    Submitted As SYSTEMTIME
End Type

'API Declarations for print object end

'API for Local system date time setting
Public Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long
Public Declare Function GetUserDefaultLCID Lib "kernel32" () As Long
Public Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
'API for Local system date time setting end


Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As _
   MEMORYSTATUS)

Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type

Const WM_CLOSE = &H10
 
'get file info & type
Public Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Public Declare Function SHGetFileInfoA Lib "shell32" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Public Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Public Declare Function DestroyIcon Lib "user32" (ByVal hIcon&) As Long
Public Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl&, ByVal i&, ByVal hDCDest&, ByVal x&, ByVal y&, ByVal Flags&) As Long

Public Type SHFILEINFO

  hIcon As Long
  iIcon As Long
  dwAttributes As Long
  szDisplayName As String * 260
  szTypeName As String * 80

End Type

'menu handler for popup menu in modal form fired by popup menu
Public Declare Function CreatePopupMenu Lib "user32" () As Long
Public Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, ByVal nReserved As Long, ByVal hWnd As Long, ByVal lprc As Any) As Long
Public Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Public Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function EnableMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long
'end of menu handler
 
'get current pointer
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Type POINTAPI
    x As Long
    y As Long
End Type
'end of get pointer

Public Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Declare Function GetComputerNameA Lib "kernel32" (ByVal sBuffer As String, lSize As Long) As Long

'set window pos
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'end of window pos

'set combobox dropdown height
Public Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
'end of set combobox dropdown height

Public Sub DisableX(frm As Form)
'*----------------------------------------------------------*
'* Name       : DisableX                                    *
'*----------------------------------------------------------*
'* Purpose    : Disables the close button ('X') on form.    *
'*----------------------------------------------------------*
'* Parameters : frm    Required. Form to disable 'X'-button *
'*----------------------------------------------------------*
'* Description: This function disables the X-button on a    *
'*            : form, to keep the user from closing a form  *
'*            : that way, but keeps the min & max buttons.  *
'*----------------------------------------------------------*
    Const MF_BYPOSITION = &H400&
    Const MF_DISABLED = &H2&

    Dim hMenu As Long, nCount As Long

    'Get handle to system menu
    hMenu = GetSystemMenu(frm.hWnd, 0)

    'Get number of items in menu
    nCount = GetMenuItemCount(hMenu)

    'Remove last item from system menu (last item is 'Close')
    Call RemoveMenu(hMenu, nCount - 1, MF_DISABLED Or MF_BYPOSITION)

    'Redraw menu
    DrawMenuBar frm.hWnd

End Sub


'*** start of NumericEdit***
'Create Numeric-Input Text Boxes
'To set a text box as numeric-input only
'simple Call NumericEdit(txtControl) at form load

Sub NumericEdit(TheControl As Control)
    Const ES_NUMBER = &H2000&
    Const GWL_STYLE = (-16)
    Dim x As Long
    Dim Estyle As Long
    Estyle = GetWindowLong(TheControl.hWnd, GWL_STYLE)
    Estyle = Estyle Or ES_NUMBER
    x = SetWindowLong(TheControl.hWnd, GWL_STYLE, Estyle)
End Sub
'*** end of NumericEdit ***

'*** start of Creates a temporary (0 byte) file in the \TEMP directory
' and returns its name

Public Function GetTempFile(Optional Prefix As String) As String
    Dim TempFile As String
    Dim TempPath As String
    Const MAX_PATH = 260
    
    ' get the path of the \TEMP directory
    TempPath = Space$(MAX_PATH)
    GetTempPath Len(TempPath), TempPath
    ' trim off characters in excess
    TempPath = Left$(TempPath, InStr(TempPath & vbNullChar, vbNullChar) - 1)
    
    ' get the name of a temporary file in that path, with a given prefix
    TempFile = Space$(MAX_PATH)
    GetTempFileName TempPath, Prefix, 0, TempFile
    GetTempFile = Left$(TempFile, InStr(TempFile & vbNullChar, vbNullChar) - 1)

End Function
'*** end of create temp file ***

'*** start of Quicky clear the treeview identified by the hWnd parameter

Sub ClearTreeViewNodes(ByVal hWnd As Long)
    Const WM_SETREDRAW As Long = &HB
    Const TV_FIRST As Long = &H1100
    Const TVM_GETNEXTITEM As Long = (TV_FIRST + 10)
    Const TVM_DELETEITEM As Long = (TV_FIRST + 1)
    Const TVGN_ROOT As Long = &H0

    Dim hItem As Long
    
    ' lock the window update to avoid flickering
    SendMessageLong hWnd, WM_SETREDRAW, False, &O0

    ' clear the treeview
    Do
        hItem = SendMessageLong(hWnd, TVM_GETNEXTITEM, TVGN_ROOT, 0)
        If hItem <= 0 Then Exit Do
        SendMessageLong hWnd, TVM_DELETEITEM, &O0, hItem
    Loop
    
    ' unlock the window
    SendMessageLong hWnd, WM_SETREDRAW, True, &O0
End Sub

Sub SetComboDropDownWidth(ComboBox As ComboBox, ByVal lWidth As Long)
    Const CB_SETDROPPEDWIDTH = &H160
    
    SendMessage ComboBox.hWnd, CB_SETDROPPEDWIDTH, lWidth, ByVal 0&
End Sub
Sub ComboBoxOpenList(ComboBox As ComboBox, Optional showIt As Boolean = True)
    Const CB_SHOWDROPDOWN = &H14F
    SendMessage ComboBox.hWnd, CB_SHOWDROPDOWN, showIt, ByVal 0&
End Sub

'*** Start: Handle printer object through API
Public Function TrimStr(strName As String) As String
    'Finds a null then trims the string
    Dim x As Integer

    x = InStr(strName, vbNullChar)
    If x > 0 Then TrimStr = Left(strName, x - 1) Else TrimStr = strName
End Function


Public Function LPSTRtoSTRING(ByVal lngPointer As Long) As String
    Dim lngLength As Long

    'Get number of characters in string
    lngLength = lstrlenW(lngPointer) * 2
    'Initialize string so we have something to copy the string into
    LPSTRtoSTRING = String(lngLength, 0)
    'Copy the string
    CopyMem ByVal StrPtr(LPSTRtoSTRING), ByVal lngPointer, lngLength
    'Convert to Unicode
    LPSTRtoSTRING = TrimStr(StrConv(LPSTRtoSTRING, vbUnicode))
End Function

'*** End: Handle printer object through API

'set listbox's extent with greater than its width, add the horizontal scroll bar
Sub SetListBoxHorizontalExtent(ByVal lb As Control, ByVal newWidth As Long)
    Const LB_SETHORIZONTALEXTENT = &H194
    
    SendMessage lb.hWnd, LB_SETHORIZONTALEXTENT, newWidth, ByVal 0&
End Sub

'Set Treeview control's tooltip on-off
Sub SetTreeViewToolTip(tv As Control, bln As Boolean)

    Const TVS_NOTOOLTIPS = &H80
    Const GWL_STYLE = (-16)
    
    If bln Then
        SetWindowLong tv.hWnd, GWL_STYLE, GetWindowLong(tv.hWnd, _
            GWL_STYLE) And (Not TVS_NOTOOLTIPS)
    Else
        SetWindowLong tv.hWnd, GWL_STYLE, GetWindowLong(tv.hWnd, _
            GWL_STYLE) Or TVS_NOTOOLTIPS
    End If
End Sub

'Local system date format string
Function GetLocalSystemDateTimeFormatString(Optional blnDateOnly As Boolean = False, Optional blnLongDate As Boolean = False, Optional blnSystem As Boolean = False) As String
    Dim str As String
    Dim LCID As Long
    
    Const LOCALE_SSHORTDATE As Long = &H1F 'short date format string
    Const LOCALE_SLONGDATE As Long = &H20 'long date format string
    Const LOCALE_STIMEFORMAT As Long = &H1003
    
    'get the user's Locale ID
    If blnSystem Then
        LCID = GetSystemDefaultLCID
    Else
        LCID = GetUserDefaultLCID
    End If
    
    If blnLongDate Then
        str = GetUserLocaleInfo(LCID, LOCALE_SLONGDATE)
    Else
        str = GetUserLocaleInfo(LCID, LOCALE_SSHORTDATE)
    End If
    
    If blnDateOnly Then
        GetLocalSystemDateTimeFormatString = str
    Else
        GetLocalSystemDateTimeFormatString = str & " " & GetUserLocaleInfo(LCID, LOCALE_STIMEFORMAT)
    End If
End Function

Public Function GetUserLocaleInfo(ByVal dwLocaleID As Long, ByVal dwLCType As Long) As String

   Dim sReturn As String
   Dim r As Long

  'call the function passing the Locale type
  'variable to retrieve the required size of
  'the string buffer needed
   r = GetLocaleInfo(dwLocaleID, dwLCType, sReturn, Len(sReturn))
    
  'if successful..
   If r Then
    
     'pad the buffer with spaces
      sReturn = Space$(r)
       
     'and call again passing the buffer
      r = GetLocaleInfo(dwLocaleID, dwLCType, sReturn, Len(sReturn))
     
     'if successful (r > 0)
      If r Then
      
        'r holds the size of the string
        'including the terminating null
         GetUserLocaleInfo = Left$(sReturn, r - 1)
      
      End If
   End If
End Function

Sub SetButtonToggle(ctl As Control, bln As Boolean)
    Const BM_SETSTATE = &HF3
    If Not TypeOf ctl Is CommandButton Then Exit Sub
    SendMessage ctl.hWnd, BM_SETSTATE, CInt(bln), 0&
End Sub

Function IsTerminalSession() As Boolean
    Dim lngMetrics As Long
    
    Const SM_REMOTESESSION = &H1000
    
    lngMetrics = GetSystemMetrics(SM_REMOTESESSION)
    
    If lngMetrics Then
        IsTerminalSession = True
    Else
        IsTerminalSession = False
    End If
    
End Function

Public Function IsThisTerminalServer() As Boolean

    Dim lRV As Long
    Dim osv As OSVERSIONINFOEX
    Dim bRV As Boolean
    
    Const VER_SERVER_NT As Long = &H80000000
    Const VER_WORKSTATION_NT As Long = &H40000000
    Const VER_SUITE_SMALLBUSINESS As Long = &H1
    Const VER_SUITE_ENTERPRISE As Long = &H2
    Const VER_SUITE_BACKOFFICE As Long = &H4
    Const VER_SUITE_COMMUNICATIONS As Long = &H8
    Const VER_SUITE_TERMINAL As Long = &H10
    Const VER_SUITE_SMALLBUSINESS_RESTRICTED As Long = &H20
    Const VER_SUITE_EMBEDDEDNT As Long = &H40
    Const VER_SUITE_DATACENTER As Long = &H80
    Const VER_SUITE_SINGLEUSERTS As Long = &H100
    Const VER_SUITE_PERSONAL As Long = &H200
    Const VER_SUITE_BLADE As Long = &H400
    
    osv.dwOSVersionInfoSize = Len(osv)
     
    lRV = GetVersionEx(osv)
     
    If CBool(osv.wSuiteMask And VER_SUITE_TERMINAL) Then
        bRV = True
    Else
        bRV = False
    End If
     
    IsThisTerminalServer = bRV
End Function

'**** start of Determine the number of visible items in a ListView control (list and report mode)
Function ListViewGetVisibleCount(lv As Control) As Long
    Const LVM_FIRST = &H1000
    Const LVM_GETCOUNTPERPAGE = (LVM_FIRST + 40)

   ListViewGetVisibleCount = SendMessage(lv.hWnd, LVM_GETCOUNTPERPAGE, 0&, _
       ByVal 0&)
End Function
'*** end of Determine the number of visible items in a ListView control

'*** Get Local computer name
Function GetLocalComputerName() As String

Dim strString As String
    'Create a buffer
  strString = String(255, Chr$(0))
  'Get the computer name
  GetComputerNameA strString, 255
  'Remove the unnecessary Chr$(0)
  strString = Left$(strString, InStr(1, strString, Chr$(0)) - 1)
  
  GetLocalComputerName = strString
End Function

Function ListBoxGetSelCount(lst As ListBox) As Long
    Const LB_GETSELCOUNT = &H190
    ListBoxGetSelCount = SendMessage(lst.hWnd, LB_GETSELCOUNT, ByVal CLng(0), ByVal CLng(0))
End Function

Public Sub SetWindowPosition(ByVal lngPointer As Long, blnOnTop As Boolean)
    Const HWND_BOTTOM As Long = 1
    Const HWND_BROADCAST As Long = &HFFFF&
    Const HWND_MESSAGE As Long = -3
    Const HWND_NOTOPMOST As Long = -2
    Const HWND_TOP As Long = 0
    Const HWND_TOPMOST As Long = -1
    
    Const SWP_ASYNCWINDOWPOS As Long = &H4000
    Const SWP_DEFERERASE As Long = &H2000
    Const SWP_FRAMECHANGED As Long = &H20
    Const SWP_DRAWFRAME As Long = SWP_FRAMECHANGED
    Const SWP_HIDEWINDOW As Long = &H80
    Const SWP_NOACTIVATE As Long = &H10
    Const SWP_NOCOPYBITS As Long = &H100
    Const SWP_NOMOVE As Long = &H2
    Const SWP_NOOWNERZORDER As Long = &H200
    Const SWP_NOREDRAW As Long = &H8
    Const SWP_NOREPOSITION As Long = SWP_NOOWNERZORDER
    Const SWP_NOSENDCHANGING As Long = &H400
    Const SWP_NOSIZE As Long = &H1
    Const SWP_NOZORDER As Long = &H4
    Const SWP_SHOWWINDOW As Long = &H40
   
    If blnOnTop Then
        SetWindowPos lngPointer, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
    Else
        SetWindowPos lngPointer, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE + SWP_NOOWNERZORDER
        SetWindowPos lngPointer, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE + SWP_NOOWNERZORDER
    End If
End Sub

Sub SetComboMaxLength(ComboBox As ComboBox, ByVal lMaxLength As Long)
    Const CB_LIMITTEXT = &H141
    
    SendMessageLong ComboBox.hWnd, CB_LIMITTEXT, lMaxLength, ByVal 0&
End Sub

Sub SetProgressbarBarColour(ctl As Control, BackColour As Long)
    Const PBM_SETBARCOLOR = 1033
    
    SendMessageLong ctl.hWnd, PBM_SETBARCOLOR, 0, ByVal BackColour
End Sub

Sub SetTextBoxUpperCase(ctl As Control)
    Const ES_UPPERCASE = &H8&
    Const GWL_STYLE = (-16)
    
    SetWindowLong ctl.hWnd, GWL_STYLE, GetWindowLong(ctl.hWnd, GWL_STYLE) Or ES_UPPERCASE
End Sub

Sub FormRemoveTitleBar(f As Form, ShowTitle As Boolean)
    Const GWL_STYLE = (-16)
    Const WS_CAPTION = &HC00000
    Const SWP_FRAMECHANGED = &H20
    Const SWP_NOMOVE = &H2
    Const SWP_NOZORDER = &H4
    Const SWP_NOSIZE = &H1

    Dim Style As Long
    ' Get window's current style bits.
    Style = GetWindowLong(f.hWnd, GWL_STYLE)
    ' Set the style bit for the title on or off.
    If ShowTitle Then
        Style = Style Or WS_CAPTION
    Else
        Style = Style And Not WS_CAPTION
    End If
    ' Send the new style to the window.
    SetWindowLong f.hWnd, GWL_STYLE, Style
    ' Repaint the window.
    SetWindowPos f.hWnd, 0, 0, 0, 0, 0, SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOSIZE
End Sub
