Attribute VB_Name = "Registry"
Option Explicit


Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Public Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long

Public Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey As Long, ByVal lpClass As String, ByVal lpcbClass As Long, ByVal lpReserved As Long, ByVal lpcSubKeys As Long, ByVal lpcbMaxSubKeyLen As Long, ByVal lpcbMaxClassLen As Long, ByVal lpcValues As Long, ByVal lpcbMaxValueNameLen As Long, ByVal lpcbMaxValueLen As Long, ByVal lpcbSecurityDescriptor As Long, ByVal lpftLastWriteTime As Long) As Long
Public Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, ByVal lpcbClass As Long, ByVal lpftLastWriteTime As Long) As Long


Public Declare Function RegQueryValueEx2 Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByVal lpType As Long, ByVal lpData As Long, ByVal lpcbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As Long, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Public Declare Function RegSetValueEx2 Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Public Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Public Const REG_DWORD = 4                      ' 32-bit number
Public Const REG_OPTION_NON_VOLATILE = 0       ' Key is preserved when system is rebooted


Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long


Public Const KEY_ALL_ACCESS = &H2003F
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_DYN_DATA = &H80000006
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const HKEY_USERS = &H80000003
Public Const ERROR_SUCCESS = 0
Public Const REG_SZ = 1

            
Public Function GetProfileString(ByVal Section$, ByVal Key$, ByVal File$, Optional ByVal DefaultValue) As String
    Dim sBuffer As String * 2048, sValue$
    If IsMissing(DefaultValue) Then DefaultValue = ""
    If vbString <> VarType(DefaultValue) Then DefaultValue = CStr(DefaultValue)
    If GetPrivateProfileString(Section, Key, "", sBuffer, 2048, File) > 0 Then
        Dim iNull As Integer
        iNull = InStr(sBuffer, Chr$(0))
        If iNull > 0 Then sValue = Trim$(Left$(sBuffer, iNull - 1))
    Else
        sValue = DefaultValue
    End If
    GetProfileString = sValue
End Function

Public Function GetProfileInt(ByVal Section$, ByVal Key$, ByVal File$, Optional ByVal DefaultValue) As Long
    If IsMissing(DefaultValue) Then DefaultValue = 0
    If Not IsNumeric(DefaultValue) Then DefaultValue = 0
    If vbLong <> VarType(DefaultValue) Then DefaultValue = CLng(DefaultValue)
    GetProfileInt = GetPrivateProfileInt(Section, Key, DefaultValue, File)
End Function

Public Sub PutProfileString(ByVal Section$, ByVal Key$, ByVal Value$, ByVal File$)
    WritePrivateProfileString Section, Key, Value, File
End Sub

Public Sub PutProfileInt(ByVal Section$, ByVal Key$, ByVal Value As Long, ByVal File$)
    WritePrivateProfileString Section, Key, CStr(Value), File
End Sub

Public Function GetClassesSetting(ByVal Section As String, ByVal Key As String, Optional ByVal DefaultValue, Optional ByRef ErrorCode) As String
    GetClassesSetting = GetRegSetting(HKEY_CLASSES_ROOT, Section, Key, DefaultValue, ErrorCode)
End Function

Public Sub SaveClassesSetting(ByVal Section As String, ByVal Key As String, ByVal Value As String, Optional ByRef ErrorCode)
    SaveRegSetting HKEY_CLASSES_ROOT, Section, Key, Value, ErrorCode
End Sub

Public Sub SaveRegSetting(ByVal KeyRoot As Long, ByVal Key As String, ByVal Variable As String, ByVal Value As String, Optional ByRef ErrorCode)
    ' set a zero error code, by default
    If Not IsMissing(ErrorCode) Then ErrorCode = 0
    
    ' open/create the given key
    Dim rc As Long, hKey As Long
    hKey = 0
    rc = RegCreateKey(KeyRoot, Key, hKey)
    If rc <> ERROR_SUCCESS Then hKey = 0  ' make sure this handle is nullified

    If rc = ERROR_SUCCESS Then
        rc = RegSetValue(hKey, Variable, REG_SZ, Value, 0)
    End If

    ' close the registry key, if opened
    If hKey <> 0 Then RegCloseKey hKey

    ' on failure, record the error code, if requested
    If rc <> ERROR_SUCCESS Then
        If Not IsMissing(ErrorCode) Then ErrorCode = rc
    End If
End Sub

Public Function GetRegSetting(ByVal KeyRoot As Long, ByVal Key As String, ByVal Variable As String, Optional ByVal DefaultValue, Optional ByRef ErrorCode) As String
    ' set a zero error code, by default
    If Not IsMissing(ErrorCode) Then ErrorCode = 0
    
    ' open the given key
    Dim rc As Long, hKey As Long
    hKey = 0
    rc = RegOpenKeyEx(KeyRoot, Key, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    If rc <> ERROR_SUCCESS Then hKey = 0  ' make sure this handle is nullified

    Dim lValueType As Long, sValue As String, lValueLen As Long
    If rc = ERROR_SUCCESS Then
        lValueLen = 1024
        sValue = String$(lValueLen, 0)
        rc = RegQueryValueEx(hKey, Variable, 0, lValueType, sValue, lValueLen)
    End If
    ' assume:  lValueType = REG_SZ
    
    ' truncate at null-terminator
    If rc = ERROR_SUCCESS Then
        If lValueLen > 0 Then
            If Chr$(0) = Mid$(sValue, lValueLen, 1) Then lValueLen = lValueLen - 1
        End If
        sValue = Left$(sValue, lValueLen)
    End If
    
    ' close the registry key, if opened
    If hKey <> 0 Then RegCloseKey hKey

    ' record the return value, on success
    ' otherwise, record the error code and return the given default value
    If rc = ERROR_SUCCESS Then
        GetRegSetting = sValue
    Else
        If Not IsMissing(ErrorCode) Then ErrorCode = rc
        If IsMissing(DefaultValue) Then
            GetRegSetting = ""
        Else
            GetRegSetting = CStr(DefaultValue)
        End If
    End If

End Function

Public Function ExitsKey(ByVal hKey As Long, ByVal sName As String) As Boolean
  Dim lRes As Long, hKey2 As Long
  lRes = RegOpenKeyEx(hKey, sName, 0, KEY_ALL_ACCESS, hKey2)
  If lRes = ERROR_SUCCESS Then
    ExitsKey = True: RegCloseKey hKey2
  Else
    ExitsKey = False
  End If
End Function

Public Sub MkRemoveKey(ByVal hKey As Long, ByVal sName As String)
  Dim lRes As Long, hKey2 As Long
  lRes = RegOpenKeyEx(hKey, sName, 0, KEY_ALL_ACCESS, hKey2)
  If lRes = ERROR_SUCCESS Then
  
    Dim lNSubKeys As Long, lMaxlen As Long, l As Long
    lRes = RegQueryInfoKey(hKey2, vbNullString, 0&, 0&, VarPtr(lNSubKeys), VarPtr(lMaxlen), 0&, 0&, 0&, 0&, 0&, 0&)
    If lRes <> ERROR_SUCCESS Then
      RegCloseKey hKey2
      Err.Raise vbObjectError + 1, "MkRemoveKey", "Ошибка запроса информации по ключу '" & sName & "' в реестре: '" & Win32ErrInfo(lRes) & "'"
    End If
    
    Dim sTmp As String, lLenIn As Long
    For l = lNSubKeys - 1 To 0 Step -1
      sTmp = String$(lMaxlen + 1, vbNullChar)
      lLenIn = lMaxlen + 1
      lRes = RegEnumKeyEx(hKey2, l, sTmp, VarPtr(lLenIn), 0&, vbNullString, 0&, 0&)
      If lRes <> ERROR_SUCCESS Then
        RegCloseKey hKey2
        Err.Raise vbObjectError + 1, "MkRemoveKey", "Ошибка енумерации по ключу '" & sName & "' в реестре: '" & Win32ErrInfo(lRes) & "'"
      Else
        If lLenIn > 0 Then
          If Chr$(0) = Mid$(sTmp, lLenIn, 1) Then lLenIn = lLenIn - 1
        End If
        sTmp = Left$(sTmp, lLenIn)
        MkRemoveKey hKey2, sTmp
      End If
    Next l
  
    RegCloseKey hKey2
    lRes = RegDeleteKey(hKey, sName)
    If lRes <> ERROR_SUCCESS Then _
      Err.Raise vbObjectError + 1, "MkRemoveKey", "Ошибка удаления ключа '" & sName & "' в реестре: '" & Win32ErrInfo(lRes) & "'"
  Else
    Err.Raise vbObjectError + 1, "MkRemoveKey", "Ошибка открытия ключа '" & sName & "' в реестре: '" & Win32ErrInfo(lRes) & "'"
  End If
End Sub
