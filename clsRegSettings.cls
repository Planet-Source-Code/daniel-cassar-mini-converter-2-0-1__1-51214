VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRegSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
' Win32 Registry functions
'
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As Any, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey As Long, ByVal lpClass As String, lpcbClass As Long, lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, lpftLastWriteTime As Any) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As Any) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
'
' Constants for Windows 32-bit Registry API
'
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_USERS = &H80000003
Private Const HKEY_PERFORMANCE_DATA = &H80000004
Private Const HKEY_CURRENT_CONFIG = &H80000005
Private Const HKEY_DYN_DATA = &H80000006
'
' Reg result codes
'
Private Const REG_CREATED_NEW_KEY = &H1                      ' New Registry Key created
Private Const REG_OPENED_EXISTING_KEY = &H2                      ' Existing Key opened
'
' Reg Create Type Values...
'
Private Const REG_OPTION_RESERVED = 0           ' Parameter is reserved
Private Const REG_OPTION_NON_VOLATILE = 0       ' Key is preserved when system is rebooted
Private Const REG_OPTION_VOLATILE = 1           ' Key is not preserved when system is rebooted
Private Const REG_OPTION_CREATE_LINK = 2        ' Created key is a symbolic link
Private Const REG_OPTION_BACKUP_RESTORE = 4     ' open for backup or restore
'
' Reg Key Security Options
'
Private Const DELETE = &H10000
Private Const READ_CONTROL = &H20000
Private Const WRITE_DAC = &H40000
Private Const WRITE_OWNER = &H80000
Private Const SYNCHRONIZE = &H100000
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Private Const STANDARD_RIGHTS_EXECUTE = (READ_CONTROL)
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const SPECIFIC_RIGHTS_ALL = &HFFFF
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Private Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Private Const KEY_EXECUTE = ((KEY_READ) And (Not SYNCHRONIZE))

Private Const ERROR_SUCCESS = 0&
Private Const ERROR_MORE_DATA = 234
Private Const ERROR_NO_MORE_ITEMS = 259

Private Const REG_SZ = 1                         ' Unicode nul terminated string
'
' Private member variables
'
Private m_Company As String
Private m_AppName As String
'
' Private class constants
'
Private Const defCompany As String = "VB and VBA Program Settings"

' ********************************************
'  Initialize and Terminate
' ********************************************
Private Sub Class_Initialize()
   m_Company = defCompany
   m_AppName = App.ProductName
End Sub

' ********************************************
'  Public Properties
' ********************************************
Public Property Let Company(ByVal NewVal As String)
   If Len(NewVal) Then
      m_Company = Trim(NewVal)
   Else
      m_Company = defCompany
   End If
End Property

Public Property Get Company() As String
   Company = m_Company
End Property

Public Property Let AppName(ByVal NewVal As String)
   If Len(NewVal) Then
      m_AppName = Trim(NewVal)
   Else
      m_AppName = App.ProductName
   End If
End Property

Public Property Get AppName() As String
   AppName = m_AppName
End Property

' ********************************************
'  Public Methods
' ********************************************
Public Function DeleteSetting(ByVal Section As String, Optional ByVal Key As String = "") As Boolean
   ' Section   Required. String expression containing the name of the section where the key setting
   '           is being deleted. If only section is provided, the specified section is deleted along
   '           with all related key settings.
   ' Key       Optional. String expression containing the name of the key setting being deleted.
   Dim nRet As Long
   Dim hKey As Long

   If Len(Key) Then
      ' Open key
      nRet = RegOpenKeyEx(HKEY_CURRENT_USER, SubKey(Section), 0&, KEY_ALL_ACCESS, hKey)
      If nRet = ERROR_SUCCESS Then
         ' Set appropriate value for default query
         If Key = "*" Then Key = vbNullString
         ' Delete the requested value
         nRet = RegDeleteValue(hKey, Key)
         Call RegCloseKey(hKey)
      End If
   Else
      ' Open parent key
      nRet = RegOpenKeyEx(HKEY_CURRENT_USER, SubKey(), 0&, KEY_ALL_ACCESS, hKey)
      If nRet = ERROR_SUCCESS Then
         ' Attempt to delete whole section
         nRet = RegDeleteKey(hKey, Section)
         Call RegCloseKey(hKey)
      End If
   End If
   DeleteSetting = (nRet = ERROR_SUCCESS)
End Function

Public Function GetAllSettings(ByVal Section As String) As Variant
   ' Section   Required. String expression containing the name of the section
   '           to enumerate.
   Dim nRet As Long
   Dim hKey As Long
   Dim nMaxValueNameLen As Long
   Dim nMaxValueLen As Long
   Dim nValueNameLen As Long
   Dim nValueLen As Long
   Dim nType As Long
   Dim nIndex As Long
   Dim nStrings As Long
   Dim ValueName As String
   Dim Value As String
   Dim Values() As String
   Dim Results() As String
   Dim i As Long

   ' Open key
   nRet = RegOpenKeyEx(HKEY_CURRENT_USER, SubKey(Section), 0&, KEY_READ, hKey)
   If nRet = ERROR_SUCCESS Then
   
      ' Get a quick snapshot of what we're facing.
      nRet = RegQueryInfoKey(hKey, vbNullString, ByVal 0&, ByVal 0&, ByVal 0&, ByVal 0&, ByVal 0&, ByVal 0&, nMaxValueNameLen, nMaxValueLen, ByVal 0&, ByVal 0&)
      If nRet = ERROR_SUCCESS Then
      
         ' Allocate buffers.
         ValueName = Space(nMaxValueNameLen + 1)
         Value = Space(nMaxValueLen + 1)
         
         ' Get value names and associated values.
         Do
            nValueNameLen = Len(ValueName)
            nValueLen = Len(Value)
            nRet = RegEnumValue(hKey, nIndex, ValueName, nValueNameLen, ByVal 0&, nType, ByVal Value, nValueLen)
            If nRet = ERROR_SUCCESS Then
               ' Only return string values.
               If nType = REG_SZ Then
                  ReDim Preserve Values(0 To 1, 0 To nStrings) As String
                  Values(0, nStrings) = Left$(ValueName, nValueNameLen)
                  Values(1, nStrings) = Left$(Value, nValueLen - 1)
                  nStrings = nStrings + 1
               End If
               nIndex = nIndex + 1
            Else 'ERROR_NO_MORE_ITEMS
               Exit Do
            End If
         Loop
         
         ' Transpose array to match VB's output, and
         ' return Results if any were obtained.
         If nStrings >= 1 Then
            ReDim Results(0 To nStrings - 1, 0 To 1) As String
            For i = 0 To nStrings - 1
               Results(i, 0) = Values(0, i)
               Results(i, 1) = Values(1, i)
            Next i
            GetAllSettings = Results
         End If
      End If
      
      Call RegCloseKey(hKey)
   End If
End Function

Public Function GetSetting(ByVal Section As String, ByVal Key As String, Optional ByVal Default As String = "") As String
   ' Section   Required. String expression containing the name of the section where the key setting is found.
   '           If omitted, key setting is assumed to be in default subkey.
   ' Key       Required. String expression containing the name of the key setting to return.
   ' Default   Optional. Expression containing the value to return if no value is set in the key setting.
   '           If omitted, default is assumed to be a zero-length string ("").
   Dim nRet As Long
   Dim hKey As Long
   Dim nType As Long
   Dim nBytes As Long
   Dim Buffer As String
   
   ' Assume failure and set return to Default
   GetSetting = Default

   ' Open key
   nRet = RegOpenKeyEx(HKEY_CURRENT_USER, SubKey(Section), 0&, KEY_ALL_ACCESS, hKey)
   If nRet = ERROR_SUCCESS Then
      ' Set appropriate value for default query
      If Key = "*" Then Key = vbNullString
      
      ' Determine how large the buffer needs to be
      nRet = RegQueryValueEx(hKey, Key, 0&, nType, ByVal Buffer, nBytes)
      If nRet = ERROR_SUCCESS Then
         ' Build buffer and get data
         If nBytes > 0 Then
            Buffer = Space(nBytes)
            nRet = RegQueryValueEx(hKey, Key, 0&, nType, ByVal Buffer, Len(Buffer))
            If nRet = ERROR_SUCCESS Then
               ' Trim NULL and return successful query!
               GetSetting = Left(Buffer, nBytes - 1)
            End If
         End If
      Call RegCloseKey(hKey)
      End If
   End If
End Function

Public Function SaveSetting(ByVal Section As String, ByVal Key As String, ByVal Setting As String) As Boolean
   ' Section   Required. String expression containing the name of the section where the key setting is being saved.
   ' Key       Required. String expression containing the name of the key setting being saved.
   ' Setting   Required. Expression containing the value that key is being set to.
   Dim nRet As Long
   Dim hKey As Long
   Dim nResult As Long
   
   ' Open (or create and open) key
   nRet = RegCreateKeyEx(HKEY_CURRENT_USER, SubKey(Section), 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, ByVal 0&, hKey, nResult)
   If nRet = ERROR_SUCCESS Then
      ' Set appropriate value for default query
      If Key = "*" Then Key = vbNullString
      ' Null-terminate setting, in case it's empty.
      ' Strange mirroring can occur otherwise.
      Setting = Setting & vbNullChar
      ' Write new value to registry
      nRet = RegSetValueEx(hKey, Key, 0&, REG_SZ, ByVal Setting, Len(Setting))
      Call RegCloseKey(hKey)
   End If
   SaveSetting = (nRet = ERROR_SUCCESS)
End Function

' ********************************************
'  Private Methods
' ********************************************
Private Function SubKey(Optional ByVal Section As String = "") As String
   ' Build SubKey from known values
   SubKey = "Software\" & m_Company & "\" & m_AppName
   If Len(Section) Then
      SubKey = SubKey & "\" & Section
   End If
End Function


