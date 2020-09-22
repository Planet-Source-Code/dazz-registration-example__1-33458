Attribute VB_Name = "Registry"
Public Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const WM_PASTE = &H302

Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Dim hKey As Long, MainKeyHandle As Long
Dim rtn As Long, lBuffer As Long, sBuffer As String

'Constants for Registry Keys
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_DYN_DATA = &H80000006
Public Const HKEY_CURRENT_USER = &H80000001
' other constants used in API calls
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const FLAGS1 = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_NOTOPMOST = -2
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_CREATE_LINK = &H20
Public Const KEY_READ = KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY
Public Const KEY_WRITE = KEY_SET_VALUE Or KEY_CREATE_SUB_KEY
Public Const KEY_EXECUTE = KEY_READ
Public Const KEY_ALL_ACCESS = KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK
Public Const ERROR_SUCCESS = 0&
Public Const REG_NONE = 0      ' No value type
Public Const REG_SZ = 1        ' Unicode nul terminated string
Public Const REG_EXPAND_SZ = 2 ' Unicode nul terminated string (with environment variable references)
Public Const REG_BINARY = 3    ' Free form binary
Public Const REG_DWORD = 4     ' 32-bit number
Public Const REG_DWORD_LITTLE_ENDIAN = 4 ' 32-bit number (same as REG_DWORD)
Public Const REG_DWORD_BIG_ENDIAN = 5    ' 32-bit number
Public Const REG_LINK = 6                ' Symbolic Link (unicode)
Public Const REG_MULTI_SZ = 7            ' Multiple Unicode strings
Public Const REG_OPTION_NON_VOLATILE = &H0
Public Const REG_CREATED_NEW_KEY = &H1

' Declare API calls for Registry access
Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Any, phkResult As Long, lpdwDisposition As Long) As Long
Declare Function RegFlushKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegSetValueEx_String Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Declare Function RegSetValueEx_DWord Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Long, ByVal cbData As Long) As Long
Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long

Public Declare Function GetWindowText Lib "User32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "User32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Public Declare Function EnumWindows Lib "User32" (ByVal lpEnumFunc As Long, ByVal lParam As Any) As Long
Public Declare Function GetClassName Lib "User32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function EnumChildWindows Lib "User32" (ByVal hwndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Any) As Long
Public Declare Function ShowWindow Lib "User32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function BringWindowToTop Lib "User32" (ByVal hWnd As Long) As Long
Public Declare Function CascadeWindows Lib "User32" (ByVal hwndParent As Long, ByVal wHow As Long, lpRect As RECT, ByVal cKids As Long, lpkids As Long) As Integer

'APIs for Spying Menus:
Public Declare Function GetMenu Lib "User32" (ByVal hWnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "User32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemID Lib "User32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuItemInfo Lib "User32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal B As Long, lpMenuItemInfo As MENUITEMINFO) As Long
Public Declare Function GetSubMenu Lib "User32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function PostMessage Lib "User32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wid As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String '* 255
    cch As Long
End Type
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Const WM_COMMAND = &H111
Public Const MIIM_TYPE = &H10
Public Const MFT_STRING = &H0&


'Public Const WM_SETFOCUS = &H7     Messages for:
Public Const WM_CLOSE = &H10                    'Closing window
Public Const SW_SHOW = 5                        'showing window
Public Const WM_SETTEXT = &HC                   'Setting text of child window
Public Const WM_GETTEXT = &HD                   'Getting text of child window
Public Const WM_GETTEXTLENGTH = &HE
Public Const EM_GETPASSWORDCHAR = &HD2          'Checking if its a password field or not
Public Const BM_CLICK = &HF5                    'Clicking a button
Public Const SW_Maximize = 3
Public Const SW_Minimize = 6
Public Const SW_HIDE = 0
Public Const SW_RESTORE = 9
Public Const WM_MDICASCADE = &H227              'Cascading windows
Public Const MDITILE_HORIZONTAL = &H1
Public Const MDITILE_SKIPDISABLED = &H2
Public Const WM_MDITILE = &H226

Public VCount As Integer, ICount As Integer
Public SpyHwnd As Long
Public jPath As String
Public jData As String


Private Type APPBARDATA
        cbSize As Long
        hWnd As Long
        uCallbackMessage As Long
        uEdge As Long
        rc As RECT
        lParam As Long '  message specific
End Type

Private Declare Function SHAppBarMessage Lib "shell32.dll" (ByVal dwMessage As Long, pData As APPBARDATA) As Long
Private Declare Function GetSystemMetrics Lib "User32" (ByVal nIndex As Long) As Long
Private Declare Function SetRect Lib "User32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetDC Lib "User32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "User32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long

Private Const WM_USER = &H400
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOACTIVATE = &H10
Private Const SM_CYSCREEN = 1
Private Const SM_CXSCREEN = 0
Private Const ABM_NEW = &H0&
Private Const ABM_REMOVE = &H1&
Private Const ABM_QUERYPOS = &H2&
Private Const ABM_SETPOS = &H3&
Private Const ABM_GETSTATE = &H4&
Private Const ABM_GETTASKBARPOS = &H5&
Private Const ABM_ACTIVATE = &H6&          'lParam == TRUE/FALSE means activate/deactivate
Private Const ABM_GETAUTOHIDEBAR = &H7&
Private Const ABM_SETAUTOHIDEBAR = &H8&
Private Const ABE_LEFT = 0
Private Const ABE_TOP = 1
Private Const ABE_RIGHT = 2
Private Const ABE_BOTTOM = 3
Private Const WU_LOGPIXELSX = 88
Private Const WU_LOGPIXELSY = 90
Private Const nTwipsPerInch = 1440
Private Const GWL_STYLE = (-16)
'Public Const REG_DWORD = 4
Public Enum jPosition
    jBottom = ABE_BOTTOM
    jtop = ABE_TOP
End Enum

Private jABD As APPBARDATA
' using Win API calls.
'
Function ExistKey(ByVal Root As Long, ByVal Key As String) As Boolean
' Check whether a key exists or not.
Dim lResult As Long
Dim Keyhandle As Long
    
    ' Try to open the key...
    lResult = RegOpenKeyEx(Root, Key, 0, KEY_READ, Keyhandle)
    
    ' If the key exists, close it (because its just a test)
    If lResult = ERROR_SUCCESS Then RegCloseKey Keyhandle
    
    ' return the value true or false
    ExistKey = (lResult = ERROR_SUCCESS)
End Function

Function GetValue(Root As Long, Key As String, field As String, Value As Variant) As Boolean
' Read a value from a specified key
' The key is set as: Root, key and name
Dim lResult As Long
Dim Keyhandle As Long
Dim dwType As Long
Dim zw As Long
Dim bufsize As Long
Dim Buffer As String
Dim i As Integer
Dim tmp As String

    ' Open the key
    lResult = RegOpenKeyEx(Root, Key, 0, KEY_READ, Keyhandle)
    GetValue = (lResult = ERROR_SUCCESS) ' success?
    
    If lResult <> ERROR_SUCCESS Then Exit Function ' Key doesn't exist
    ' Get the value
    lResult = RegQueryValueEx(Keyhandle, field, 0&, dwType, _
              ByVal 0&, bufsize)
    GetValue = (lResult = ERROR_SUCCESS) ' Success?
        
    If lResult <> ERROR_SUCCESS Then Exit Function ' Name doesn't exist
 
    Select Case dwType
        Case REG_SZ       ' Zero terminated string
            Buffer = Space(bufsize + 1)
            lResult = RegQueryValueEx(Keyhandle, field, 0&, dwType, ByVal Buffer, bufsize)
            GetValue = (lResult = ERROR_SUCCESS)
            If lResult <> ERROR_SUCCESS Then Exit Function ' Error
            Value = Buffer
            
        Case REG_DWORD     ' 32-Bit Number   !!!! Word
            bufsize = 4      ' = 32 Bit
            lResult = RegQueryValueEx(Keyhandle, field, 0&, dwType, zw, bufsize)
            GetValue = (lResult = ERROR_SUCCESS)
            If lResult <> ERROR_SUCCESS Then Exit Function ' Error
            Value = zw
   
        Case REG_BINARY     ' Binary
            Buffer = Space(bufsize + 1)
            lResult = RegQueryValueEx(Keyhandle, field, 0&, dwType, ByVal Buffer, bufsize)
            GetValue = (lResult = ERROR_SUCCESS)
            If lResult <> ERROR_SUCCESS Then Exit Function ' Error
            Value = ""
            For i = 1 To bufsize
                tmp = Hex(Asc(Mid(Buffer, i, 1)))
                If Len(tmp) = 1 Then tmp = "0" + tmp
                Value = Value + tmp + " "
            Next i
        ' Here is space for other data types
    End Select
  
    If lResult = ERROR_SUCCESS Then RegCloseKey Keyhandle
    GetValue = True
    
End Function

Function CreateKey(Root As Long, newkey As String, Class As String) As Boolean
Dim lResult As Long
Dim Keyhandle As Long
Dim Action As Long

    lResult = RegCreateKeyEx(Root, newkey, 0, Class, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, Keyhandle, Action)
    If lResult = ERROR_SUCCESS Then
        If RegFlushKey(Keyhandle) = ERROR_SUCCESS Then RegCloseKey Keyhandle
    Else
        CreateKey = False
        Exit Function
    End If
    CreateKey = (Action = REG_CREATED_NEW_KEY)
    
End Function

Function SetValue(Root As Long, Key As String, field As String, Value As Variant) As Boolean
Dim lResult As Long
Dim Keyhandle As Long
Dim S As String
Dim l As Long
    
    lResult = RegOpenKeyEx(Root, Key, 0, KEY_ALL_ACCESS, Keyhandle)
    If lResult <> ERROR_SUCCESS Then
        SetValue = False
        Exit Function
    End If
 
    Select Case VarType(Value)
        Case vbInteger, vbLong
            l = CLng(Value)
            lResult = RegSetValueEx_DWord(Keyhandle, field, 0, REG_DWORD, l, 4)
        Case vbString
            S = CStr(Value)
            lResult = RegSetValueEx_String(Keyhandle, field, 0, REG_SZ, S, Len(S) + 1)    ' +1 for trailing 00
        
        ' Here is space for other data types
    End Select
    
    RegCloseKey Keyhandle
    SetValue = (lResult = ERROR_SUCCESS)
    
End Function

Function DeleteKey(Root As Long, Key As String) As Boolean
Dim lResult As Long

    lResult = RegDeleteKey(Root, Key)
    DeleteKey = (lResult = ERROR_SUCCESS)
End Function



Function gettext(iHwnd As Long) As String
    Dim Textlen As Long
    Dim Text As String

    Textlen = SendMessage(iHwnd, WM_GETTEXTLENGTH, 0, 0)
    If Textlen = 0 Then
        gettext = ">No text for this class<"
        Exit Function
    End If
    Textlen = Textlen + 1
    Text = Space(Textlen)
    Textlen = SendMessage(iHwnd, WM_GETTEXT, Textlen, ByVal Text)
    'The 'ByVal' keyword is necessary or you'll get an invalid page fault
    'and the app crashes, and takes VB with it.
    gettext = Right(Text, Textlen)

End Function

Public Function ConvertTwipsToPixels(nTwips As Long, nDirection As Long) As Integer
    Dim hDC As Long
    Dim nPixelsPerInch As Long
       
    hDC = GetDC(0)
    If (nDirection = 0) Then       'Horizontal
        nPixelsPerInch = GetDeviceCaps(hDC, WU_LOGPIXELSX)
    Else                            'Vertical
        nPixelsPerInch = GetDeviceCaps(hDC, WU_LOGPIXELSY)
    End If
    
    hDC = ReleaseDC(0, hDC)
    ConvertTwipsToPixels = (nTwips / nTwipsPerInch) * nPixelsPerInch
End Function

Public Function GetSettingString(hKey As Long, strPath As String, strValue As String, Optional Default As String) As String
Dim hCurKey As Long
Dim lValueType As Long
Dim strbuffer As String
Dim lDataBufferSize As Long
Dim intZeroPos As Integer
Dim lRegResult As Long
' Set up default value
If Not IsEmpty(Default) Then
  GetSettingString = Default
Else
  GetSettingString = ""
End If

' Open the key and get length of string
lRegResult = RegOpenKey(hKey, strPath, hCurKey)
lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, lValueType, ByVal 0&, lDataBufferSize)

If lRegResult = ERROR_SUCCESS Then

  If lValueType = REG_SZ Then
    ' initialise string buffer and retrieve string
    strbuffer = String(lDataBufferSize, " ")
    lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, 0&, ByVal strbuffer, lDataBufferSize)
    
    ' format string
    intZeroPos = InStr(strbuffer, Chr$(0))
    If intZeroPos > 0 Then
      GetSettingString = Right$(strbuffer, intZeroPos)
        
    Else
      GetSettingString = strbuffer
    End If

  End If

Else
  ' there is a problem
End If

lRegResult = RegCloseKey(hCurKey)
End Function
Public Sub SaveSettingString(hKey As Long, strPath As String, strValue As String, strData As String)
Dim hCurKey As Long
Dim lRegResult As Long

lRegResult = RegCreateKey(hKey, strPath, hCurKey)

lRegResult = RegSetValueEx(hCurKey, strValue, 0, REG_SZ, ByVal strData, Len(strData))

If lRegResult <> ERROR_SUCCESS Then
  'there is a problem
End If

lRegResult = RegCloseKey(hCurKey)
End Sub
Public Sub DeleteValue(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String)
Dim hCurKey As Long
Dim lRegResult As Long

lRegResult = RegOpenKey(hKey, strPath, hCurKey)

lRegResult = RegDeleteValue(hCurKey, strValue)

lRegResult = RegCloseKey(hCurKey)

End Sub
Private Sub ParseKey(KeyName As String, Keyhandle As Long)
    
rtn = InStr(KeyName, "\") 'return if "\" is contained in the Keyname

If Right(KeyName, 5) <> "HKEY_" Or Right(KeyName, 1) = "\" Then 'if the is a "\" at the end of the Keyname then
   MsgBox "Incorrect Format:" + Chr(10) + Chr(10) + KeyName 'display error to the user
   Exit Sub 'exit the procedure
ElseIf rtn = 0 Then 'if the Keyname contains no "\"
   Keyhandle = GetMainKeyHandle(KeyName)
   KeyName = "" 'leave Keyname blank
Else 'otherwise, Keyname contains "\"
   Keyhandle = GetMainKeyHandle(Right(KeyName, rtn - 1)) 'seperate the Keyname
   KeyName = Right(KeyName, Len(KeyName) - rtn)
End If

End Sub
Function GetStringValue(SubKey As String, Entry As String)

Call ParseKey(SubKey, MainKeyHandle)

If MainKeyHandle Then
   rtn = RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_READ, hKey) 'open the key
   If rtn = ERROR_SUCCESS Then 'if the key could be opened then
      sBuffer = Space(255)     'make a buffer
      lBufferSize = Len(sBuffer)
      rtn = RegQueryValueEx(hKey, Entry, 0, REG_SZ, sBuffer, lBufferSize) 'get the value from the registry
      If rtn = ERROR_SUCCESS Then 'if the value could be retreived then
         rtn = RegCloseKey(hKey)  'close the key
         sBuffer = Trim(sBuffer)
         GetStringValue = Right(sBuffer, Len(sBuffer) - 1) 'return the value to the user
      Else                        'otherwise, if the value couldnt be retreived
         GetStringValue = "Error" 'return Error to the user
         If DisplayErrorMsg = True Then 'if the user wants errors displayed then
            MsgBox ErrorMsg(rtn)  'tell the user what was wrong
         End If
      End If
   Else 'otherwise, if the key couldnt be opened
      GetStringValue = "Error"       'return Error to the user
      If DisplayErrorMsg = True Then 'if the user wants errors displayed then
         MsgBox ErrorMsg(rtn)        'tell the user what was wrong
      End If
   End If
End If

End Function
Function GetMainKeyHandle(MainKeyName As String) As Long

Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS = &H80000003
Const HKEY_PERFORMANCE_DATA = &H80000004
Const HKEY_CURRENT_CONFIG = &H80000005
Const HKEY_DYN_DATA = &H80000006
   
Select Case MainKeyName
       Case "HKEY_CLASSES_ROOT"
            GetMainKeyHandle = HKEY_CLASSES_ROOT
       Case "HKEY_CURRENT_USER"
            GetMainKeyHandle = HKEY_CURRENT_USER
       Case "HKEY_LOCAL_MACHINE"
            GetMainKeyHandle = HKEY_LOCAL_MACHINE
       Case "HKEY_USERS"
            GetMainKeyHandle = HKEY_USERS
       Case "HKEY_PERFORMANCE_DATA"
            GetMainKeyHandle = HKEY_PERFORMANCE_DATA
       Case "HKEY_CURRENT_CONFIG"
            GetMainKeyHandle = HKEY_CURRENT_CONFIG
       Case "HKEY_DYN_DATA"
            GetMainKeyHandle = HKEY_DYN_DATA
End Select

End Function

Function ErrorMsg(lErrorCode As Long) As String
    
'If an error does accurr, and the user wants error messages displayed, then
'display one of the following error messages

Select Case lErrorCode
       Case 1009, 1015
            GetErrorMsg = "The Registry Database is corrupt!"
       Case 2, 1010
            GetErrorMsg = "Bad Key Name"
       Case 1011
            GetErrorMsg = "Can't Open Key"
       Case 4, 1012
            GetErrorMsg = "Can't Read Key"
       Case 5
            GetErrorMsg = "Access to this key is denied"
       Case 1013
            GetErrorMsg = "Can't Write Key"
       Case 8, 14
            GetErrorMsg = "Out of memory"
       Case 87
            GetErrorMsg = "Invalid Parameter"
       Case 234
            GetErrorMsg = "There is more data than the buffer has been allocated to hold."
       Case Else
            GetErrorMsg = "Undefined Error Code:  " & Str$(lErrorCode)
End Select

End Function
Public Function GetBinary(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String, bArray() As Byte) As Boolean
'How to use this function:
'Dim bArray() As Byte
'If GetBinary(KEY, PATH, VALUE, bArray()) = True Then
'   MsgBox StrConv(bArray, vbUnicode)
'End If
    Dim lResult As Long, lValueType As Long, lBuf As Long
    Dim lDataBufSize As Long, R As Long, keyhand As Long
    R = RegOpenKey(hKey, strPath, keyhand)
    ' Get length/data type
    lDataBufSize = 0
    ReDim bArray(1 To 1) As Byte
    lResult = RegQueryValueEx(keyhand, strValueName, 0&, lValueType, bArray(1), lDataBufSize)
    If lResult > 0 And lValueType = REG_BINARY Then
        ReDim bArray(1 To lDataBufSize) As Byte
        lResult = RegQueryValueEx(keyhand, strValueName, 0&, lValueType, bArray(1), lDataBufSize)
        If lResult = ERROR_SUCCESS Then GetBinary = True
    End If
    R = RegCloseKey(keyhand)
End Function
Public Function SaveBinary(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String, bStart As Byte, bLen As Long) As Boolean
'How to use this function:
'Dim bArray(1 To 3) As Byte
'SaveBinary Key, Path, Value, bArray(1), 3
    Dim lResult As Long
    Dim keyhand As Long
    Dim R As Long
    R = RegCreateKey(hKey, strPath, keyhand)
    lResult = RegSetValueEx(keyhand, strValueName, 0&, REG_BINARY, bStart, bLen)
    If lResult = ERROR_SUCCESS Then SaveBinary = True
    R = RegCloseKey(keyhand)
End Function
Public Function CheckKey(hKey As Long, strPath As String, ByVal strValueName As String) As String ' this function returns if a valuse exists or not

    Dim lResult As Long
    Dim lValueType As Long
    Dim lBuf As Long
    Dim lDataBufSize As Long
    Dim R As Long
    Dim keyhand As Long
    R = RegOpenKey(hKey, strPath, keyhand)
    lDataBufSize = 4
    lResult = RegQueryValueEx(keyhand, strValueName, 0&, lValueType, lBuf, lDataBufSize)


    If lResult = ERROR_SUCCESS Then
        CheckKey = "No"
        
    Else
        CheckKey = "Yes"
    End If
End Function


Function getdword(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String) As Long
Dim lResult As Long
Dim lValueType As Long
Dim lBuf As Long
Dim lDataBufSize As Long
Dim R As Long
Dim keyhand As Long

R = RegOpenKey(hKey, strPath, keyhand)

 ' Get length/data type
lDataBufSize = 4
    
lResult = RegQueryValueEx(keyhand, strValueName, 0&, lValueType, lBuf, lDataBufSize)

If lResult = ERROR_SUCCESS Then
    If lValueType = REG_DWORD Then
        getdword = lBuf
    End If
'Else
'    Call errlog("GetDWORD-" & strPath, False)
End If

R = RegCloseKey(keyhand)
    
End Function
Function SaveDword(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String, ByVal lData As Long)
    Dim lResult As Long
    Dim keyhand As Long
    Dim R As Long
    R = RegCreateKey(hKey, strPath, keyhand)
    lResult = RegSetValueEx(keyhand, strValueName, 0&, REG_DWORD, lData, 4)
    'If lResult <> error_success Then Call errlog("SetDWORD", False)
    R = RegCloseKey(keyhand)
End Function






