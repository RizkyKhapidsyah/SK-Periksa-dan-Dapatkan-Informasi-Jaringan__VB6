VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRegistryAccess"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = " Encapsulate reading and writing to registry in the HKEY_LOCAL_MACHINE hive.\r\n"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"357C51DA0230"
Option Explicit

Private Const HKEY_LOCAL_MACHINE       As Long = &H80000002
Private Const ERROR_SUCCESS            As Long = 0&
Private Const ERROR_NO_MORE_ITEMS      As Long = 259&
Private Const MAX_PATH                 As Long = 260&

Private Const REG_NONE                 As Long = 0&
Private Const REG_SZ                   As Long = 1&
Private Const REG_EXPAND_SZ            As Long = 2&
Private Const REG_BINARY               As Long = 3&
Private Const REG_DWORD                As Long = 4&
Private Const REG_DWORD_LITTLE_ENDIAN  As Long = 4&
Private Const REG_DWORD_BIG_ENDIAN     As Long = 5&
Private Const REG_LINK                 As Long = 6&
Private Const REG_MULTI_SZ             As Long = 7&
Private Const REG_RESOURCE_LIST        As Long = 8&
Private Const REG_FULL_RESOURCE_DESCRIPTOR   As Long = 9&

Private Const REG_CREATED_NEW_KEY      As Long = &H1&
Private Const REG_OPENED_EXISTING_KEY  As Long = &H2&
Private Const REG_OPTION_BACKUP_RESTORE   As Long = 4&
Private Const REG_OPTION_NON_VOLATILE  As Long = 0&
Private Const REG_OPTION_VOLATILE      As Long = 1&

Private Const STANDARD_RIGHTS_ALL      As Long = &H1F0000
Private Const STANDARD_RIGHTS_EXECUTE  As Long = &H20000
Private Const STANDARD_RIGHTS_READ     As Long = &H20000
Private Const STANDARD_RIGHTS_REQUIRED As Long = &HF0000
Private Const STANDARD_RIGHTS_WRITE    As Long = &H20000
Private Const SYNCHRONIZE              As Long = &H100000
Private Const KEY_CREATE_LINK          As Long = &H20&
Private Const KEY_CREATE_SUB_KEY       As Long = &H4&
Private Const KEY_ENUMERATE_SUB_KEYS   As Long = &H8&
Private Const KEY_EVENT                As Long = &H1&
Private Const KEY_NOTIFY               As Long = &H10&
Private Const KEY_QUERY_VALUE          As Long = &H1&
Private Const KEY_SET_VALUE            As Long = &H2&
Private Const KEY_WRITE                As Long = (STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And Not SYNCHRONIZE
Private Const KEY_READ                 As Long = (STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And Not SYNCHRONIZE
Private Const KEY_EXECUTE              As Long = KEY_READ&
Private Const KEY_ALL_ACCESS           As Long = (STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And Not SYNCHRONIZE

Private Type SECURITY_ATTRIBUTES
   nLength                             As Long
   lpSecurityDescriptor                As Long
   bInheritHandle                      As Long
End Type

Private Type FILETIME
   dwLowDateTime                       As Long
   dwHighDateTime                      As Long
End Type

Private Declare Function RegDeleteValue _
   Lib "advapi32.dll" Alias "RegDeleteValueA" _
   (ByVal hKey As Long, _
    ByVal lpValueName As String) As Long
Private Declare Function RegDeleteKey _
   Lib "advapi32.dll" Alias "RegDeleteKeyA" _
   (ByVal hKey As Long, _
   ByVal lpSubKey As String) As Long
Private Declare Function RegOpenKeyEx _
   Lib "advapi32.dll" Alias "RegOpenKeyExA" _
   (ByVal hKey As Long, _
    ByVal lpSubKey As String, _
    ByVal ulOptions As Long, _
    ByVal samDesired As Long, _
    phkResult As Long) As Long
Private Declare Function RegOpenKey _
   Lib "advapi32.dll" Alias "RegOpenKeyA" _
   (ByVal hKey As Long, _
    ByVal lpSubKey As String, _
    phkResult As Long) As Long
Private Declare Function RegCloseKey _
   Lib "advapi32.dll" _
   (ByVal hKey As Long) As Long
Private Declare Function RegSetValueEx _
   Lib "advapi32.dll" Alias "RegSetValueExA" _
   (ByVal hKey As Long, _
    ByVal lpValueName As String, _
    ByVal Reserved As Long, _
    ByVal dwType As Long, _
    lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegCreateKeyEx _
   Lib "advapi32.dll" Alias "RegCreateKeyExA" _
   (ByVal hKey As Long, _
    ByVal lpSubKey As String, _
    ByVal Reserved As Long, _
    ByVal lpClass As String, _
    ByVal dwOptions As Long, _
    ByVal samDesired As Long, _
    ByVal lpSecurityAttributes As Long, _
    phkResult As Long, _
    lpdwDisposition As Long) As Long
Private Declare Function RegCreateKey _
   Lib "advapi32.dll" Alias "RegCreateKeyA" _
   (ByVal hKey As Long, _
    ByVal lpSubKey As String, _
    phkResult As Long) As Long
Private Declare Function RegQueryValueEx _
   Lib "advapi32.dll" Alias "RegQueryValueExA" _
   (ByVal hKey As Long, _
    ByVal lpValueName As String, _
    ByVal lpReserved As Long, _
    lpType As Long, _
    lpData As Any, _
    lpcbData As Long) As Long
Private Declare Function RegEnumKeyEx _
   Lib "advapi32.dll" Alias "RegEnumKeyExA" _
   (ByVal hKey As Long, _
    ByVal dwIndex As Long, _
    ByVal lpName As String, _
    lpcbName As Long, _
    ByVal lpReserved As Long, _
    ByVal lpClass As String, _
    lpcbClass As Long, _
    lpftLastWriteTime As FILETIME) As Long
Private Declare Function RegEnumValue _
   Lib "advapi32.dll" Alias "RegEnumValueA" _
   (ByVal hKey As Long, _
    ByVal dwIndex As Long, _
    ByVal lpValueName As String, _
    lpcbValueName As Long, _
    ByVal lpReserved As Long, _
    lpType As Long, _
    lpData As Any, _
    lpcbData As Long) As Long

' *******************************************************
' Routine Name : (PUBLIC in CLASS) EnumerateRegistryValues
' Written By   : L.J. Johnson
' Inputs       : xi_strSubKey:String                  - Name of the key to delete the value from
'              : xi_lngRegistryHive:enumRegistryHives - (Optional) ID of the hive to use (defaults to HKEY_LOCAL_MACHINE)
' Ouputs       : Variant -- Returns variant array
'              :     which contains a ValueName, Value, and DataType
'              :     for each entry
' Description  : Retrieves variant array for each key of a subkey
' *******************************************************
Public Function EnumerateRegistryValues(ByVal xi_strSubKey As String, _
                                        Optional ByVal xi_lngRegistryHive As Long) As Variant
On Error Resume Next                   ' Don't accept errors here
   Const p_constProcName               As String = "DeleteRegistryValue"
   Dim p_lngRtn                        As Long
   Dim p_lngKey                        As Long
   Dim p_lngKeyHwnd                    As Long
   Dim p_lngIndex                      As Long
   Dim p_lngNumItems                   As Long
   Dim p_strValueName                  As String
   Dim p_lngValueNameLen               As Long
   Dim p_vntValue                      As String
   Dim p_lngValueLen                   As Long
   Dim p_lngValueType                  As Long
   Dim p_strValueType                  As String
   Dim p_dblTempValue                  As Double
   Dim p_avntTmpRow                    As Variant
   Dim p_avntTmpRtn()                  As Variant

   ' ------------------------------------------
   ' Delete under the passed hive ID (defaults
   '     to HKEY_LOCAL_MACHINE)
   ' ------------------------------------------
   If xi_lngRegistryHive = 0 Then
      p_lngKey = HKEY_LOCAL_MACHINE
   Else
      p_lngKey = xi_lngRegistryHive
   End If
   
   ' ------------------------------------------
   ' Open the key with rights to enumerate the
   '     subkeys
   ' ------------------------------------------
   p_lngRtn = RegOpenKeyEx(p_lngKey, xi_strSubKey, 0&, KEY_READ, p_lngKeyHwnd)
   If p_lngRtn <> ERROR_SUCCESS Then
      Exit Function
   Else
      ' ---------------------------------------
      ' p_lngKeyHwnd is handle to the open key
      ' ---------------------------------------
   End If
   
   ' ------------------------------------------
   ' Start with an index value of zero
   ' ------------------------------------------
   p_lngIndex = 0
   p_lngNumItems = 0
   
   ' ------------------------------------------
   ' Loop thru the possible keys
   ' ------------------------------------------
   Do
      p_lngValueNameLen = MAX_PATH
      p_strValueName = Space$(p_lngValueNameLen)
      p_lngValueLen = MAX_PATH
      p_vntValue = Space$(p_lngValueLen)
      
      p_lngRtn = RegEnumValue(p_lngKeyHwnd, p_lngIndex, p_strValueName, p_lngValueNameLen, 0&, p_lngValueType, ByVal p_vntValue, p_lngValueLen)
      
      If p_lngRtn = ERROR_NO_MORE_ITEMS Then
         ' ------------------------------------
         ' Everything is OK, just at end of items
         ' ------------------------------------
         Exit Do
      ElseIf p_lngRtn <> ERROR_SUCCESS Then
         ' ------------------------------------
         ' Some other error
         ' ------------------------------------
         Debug.Print Err.Number
         Exit Do
      Else
         ' ------------------------------------
         ' Get the values for this item
         ' ------------------------------------
         p_strValueName = Mid$(p_strValueName, 1, p_lngValueNameLen)
         
         If p_lngValueType = REG_DWORD Then
            p_dblTempValue = Asc(Mid$(p_vntValue, 1, 1)) + &H100& * Asc(Mid$(p_vntValue, 2, 1)) + &H10000 * Asc(Mid$(p_vntValue, 3, 1)) + &H1000000 * CDbl(Asc(Mid$(p_vntValue, 4, 1)))
            If p_dblTempValue > &H7FFFFFFF Then
                p_vntValue = "&H" & Hex$(p_dblTempValue - 4294967296#)
            Else
                p_vntValue = "&H" & Hex$(p_dblTempValue)
            End If
         End If
         
         Select Case p_lngValueType
            Case REG_NONE
               p_strValueType = "REG_NONE"
            Case REG_SZ
               p_strValueType = "REG_SZ"
            Case REG_EXPAND_SZ
               p_strValueType = "REG_EXPAND_SZ"
            Case REG_BINARY
               p_strValueType = "REG_BINARY"
            Case REG_DWORD
               p_strValueType = "REG_DWORD"
            Case REG_DWORD_LITTLE_ENDIAN
               p_strValueType = "REG_DWORD_LITTLE_ENDIAN"
            Case REG_DWORD_BIG_ENDIAN
               p_strValueType = "REG_DWORD_BIG_ENDIAN"
            Case REG_LINK
               p_strValueType = "REG_LINK"
            Case REG_MULTI_SZ
               p_strValueType = "REG_MULTI_SZ"
            Case REG_RESOURCE_LIST
               p_strValueType = "REG_RESOURCE_LIST"
            Case Else
               p_strValueType = "UNKNOWN_TYPE"
         End Select
         
         p_avntTmpRow = Array(p_strValueName, p_vntValue, p_strValueType)
         
         p_lngNumItems = p_lngNumItems + 1
         ReDim Preserve p_avntTmpRtn(1 To p_lngNumItems)
         p_avntTmpRtn(p_lngNumItems) = p_avntTmpRow
      End If
      
      ' ---------------------------------------
      ' Increment the index
      ' ---------------------------------------
      p_lngIndex = p_lngIndex + 1
      
   Loop
   
   ' ------------------------------------------
   ' Close the key
   ' ------------------------------------------
   p_lngRtn = RegCloseKey(p_lngKeyHwnd)
   
   ' ------------------------------------------
   ' Set the return value
   ' ------------------------------------------
   EnumerateRegistryValues = p_avntTmpRtn
   
End Function

' *******************************************************
' Routine Name : (PUBLIC in CLASS) EnumerateRegistryKeys
' Written By   : L.J. Johnson
' Inputs       : xi_strSubKey:String                  - Name of the key to delete the value from
'              : xi_lngRegistryHive:enumRegistryHives - (Optional) ID of the hive to use (defaults to HKEY_LOCAL_MACHINE)
' Ouputs       : Variant -- Returns variant array
'              :     which contains a KeyName and KeyClass
'              :     for each entry
' Description  : Retrieves variant array for each key of a subkey
' *******************************************************
Public Function EnumerateRegistryKeys(ByVal xi_strSubKey As String, _
                                      Optional ByVal xi_lngRegistryHive As Long) As Variant
On Error Resume Next                   ' Don't accept errors here
   Const p_constProcName               As String = "DeleteRegistryValue"
   Dim p_lngRtn                        As Long
   Dim p_lngKey                        As Long
   Dim p_lngValueType                  As Long
   Dim p_lngKeyHwnd                    As Long
   Dim p_lngIndex                      As Long
   Dim p_lngNumItems                   As Long
   Dim p_strValueName                    As String
   Dim p_lngValueNameLen                 As Long
   Dim p_vntValue                  As String
   Dim p_plngValueLen               As Long
   Dim p_typFileTime                   As FILETIME
   Dim p_avntTmpRow                    As Variant
   Dim p_avntTmpRtn()                  As Variant
   
   ' ------------------------------------------
   ' Delete under the passed hive ID (defaults
   '     to HKEY_LOCAL_MACHINE)
   ' ------------------------------------------
   If xi_lngRegistryHive = 0 Then
      p_lngKey = HKEY_LOCAL_MACHINE
   Else
      p_lngKey = xi_lngRegistryHive
   End If
   
   ' ------------------------------------------
   ' Open the key with rights to enumerate the
   '     subkeys
   ' ------------------------------------------
   p_lngRtn = RegOpenKeyEx(p_lngKey, xi_strSubKey, 0&, KEY_ENUMERATE_SUB_KEYS, p_lngKeyHwnd)
   If p_lngRtn <> ERROR_SUCCESS Then
      Exit Function
   Else
      ' ---------------------------------------
      ' p_lngKeyHwnd is handle to the open key
      ' ---------------------------------------
   End If
   
   ' ------------------------------------------
   ' Start with an index value of zero
   ' ------------------------------------------
   p_lngIndex = 0
   p_lngNumItems = 0
   
   ' ------------------------------------------
   ' Loop thru the possible keys
   ' ------------------------------------------
   Do
      p_lngValueNameLen = MAX_PATH
      p_strValueName = Space$(p_lngValueNameLen)
      p_plngValueLen = MAX_PATH
      p_vntValue = Space$(p_plngValueLen)
      p_lngRtn = RegEnumKeyEx(p_lngKeyHwnd, p_lngIndex, p_strValueName, p_lngValueNameLen, 0&, p_vntValue, p_plngValueLen, p_typFileTime)
      
      If p_lngRtn = ERROR_NO_MORE_ITEMS Then
         ' ------------------------------------
         ' Everything is OK, just at end of items
         ' ------------------------------------
         Exit Do
      ElseIf p_lngRtn <> ERROR_SUCCESS Then
         ' ------------------------------------
         ' Some other error
         ' ------------------------------------
         Debug.Print Err.Number
         Exit Do
      Else
         ' ------------------------------------
         ' Get the values for this item
         ' ------------------------------------
         p_strValueName = Mid$(p_strValueName, 1, p_lngValueNameLen)
         p_vntValue = Mid$(p_vntValue, 1, p_plngValueLen)
         p_avntTmpRow = Array(p_strValueName, p_vntValue)
         
         p_lngNumItems = p_lngNumItems + 1
         ReDim Preserve p_avntTmpRtn(1 To p_lngNumItems)
         p_avntTmpRtn(p_lngNumItems) = p_avntTmpRow
      End If
      
      ' ---------------------------------------
      ' Increment the index
      ' ---------------------------------------
      p_lngIndex = p_lngIndex + 1
      
   Loop
   
   ' ------------------------------------------
   ' Close the key
   ' ------------------------------------------
   p_lngRtn = RegCloseKey(p_lngKeyHwnd)
   
   ' ------------------------------------------
   ' Set the return value
   ' ------------------------------------------
   EnumerateRegistryKeys = p_avntTmpRtn
   
End Function

' *******************************************************
' Routine Name : (PUBLIC in CLASS) DeleteRegistryValue
' Written By   : L.J. Johnson
' Inputs       : xi_strSubKey:String                  - Name of the key to delete the value from
'              : xi_strValueName:String               - Name of the value to delete
'              : xi_lngRegistryHive:enumRegistryHives - (Optional) ID of the hive to use (defaults to HKEY_LOCAL_MACHINE)
' Ouputs       : Boolen -- Success or failure deleting value
' Description  : Delete a value in the registry
' *******************************************************
Public Function DeleteRegistryValue(ByVal xi_strSubKey As String, _
                                    ByVal xi_strValueName As String, _
                                    Optional ByVal xi_lngRegistryHive As Long) As Boolean
On Error Resume Next                   ' Don't accept errors here
   Const p_constProcName               As String = "DeleteRegistryValue"
   Dim p_lngRtn                        As Long
   Dim p_lngKey                        As Long
   Dim p_lngValueType                  As Long
   Dim p_lngKeyHwnd                    As Long
   
   ' ------------------------------------------
   ' Default to failure
   ' ------------------------------------------
   DeleteRegistryValue = False
   
   ' ------------------------------------------
   ' Delete under the passed hive ID (defaults
   '     to HKEY_LOCAL_MACHINE)
   ' ------------------------------------------
   If xi_lngRegistryHive = 0 Then
      p_lngKey = HKEY_LOCAL_MACHINE
   Else
      p_lngKey = xi_lngRegistryHive
   End If
   
   ' ------------------------------------------
   ' Open the key with all rights
   ' ------------------------------------------
   p_lngRtn = RegOpenKeyEx(p_lngKey, xi_strSubKey, 0&, KEY_ALL_ACCESS, p_lngKeyHwnd)
   If p_lngRtn <> ERROR_SUCCESS Then
      Exit Function
   Else
      ' ---------------------------------------
      ' p_lngKeyHwnd is handle to the open key
      ' ---------------------------------------
   End If
   
   ' ------------------------------------------
   ' Delete the value
   ' ------------------------------------------
   p_lngRtn = RegDeleteValue(p_lngKeyHwnd, xi_strValueName)
   If p_lngRtn <> ERROR_SUCCESS Then
      DeleteRegistryValue = False
   Else
      DeleteRegistryValue = True
   End If

   ' ------------------------------------------
   ' Close the key
   ' ------------------------------------------
   p_lngRtn = RegCloseKey(p_lngKeyHwnd)
   
End Function

' *******************************************************
' Routine Name : (PUBLIC in CLASS) DeleteRegistryKey
' Written By   : L.J. Johnson
' Inputs       : xi_strSubKey:String                  - Name of the key to delete
'              : xi_lngRegistryHive:enumRegistryHives - (Optional) ID of the hive to use (defaults to HKEY_LOCAL_MACHINE)
' Ouputs       : Boolen -- Success or failure deleting key
' Description  : Delete a subkey in the registry
'              : NOTE: Can't delete keys with subkeys on NT!
' *******************************************************
Public Function DeleteRegistryKey(ByVal xi_strSubKey As String, _
                                  Optional ByVal xi_lngRegistryHive As Long) As Boolean
On Error Resume Next                   ' Don't accept errors here
   Const p_constProcName               As String = "DeleteRegistryKey"
   Dim p_lngRtn                        As Long
   Dim p_lngKey                        As Long
   Dim p_lngKeyHwnd                    As Long
   Dim p_lngDisposition                As Long
   Dim p_lngLoop                       As Long
   Dim p_strSubKey                     As String
   Dim p_strToDelete                   As String
   
   ' ------------------------------------------
   ' Default to failure
   ' ------------------------------------------
   DeleteRegistryKey = False
   
   ' ------------------------------------------
   ' Don't accept a blank subkey
   ' ------------------------------------------
   If Len(Trim$(xi_strSubKey)) < 1 Then
      Exit Function
   End If
   
   ' ------------------------------------------
   ' Create under the passed hive ID (defaults
   '     to HKEY_LOCAL_MACHINE)
   ' ------------------------------------------
   If xi_lngRegistryHive = 0 Then
      p_lngKey = HKEY_LOCAL_MACHINE
   Else
      p_lngKey = xi_lngRegistryHive
   End If
   
   ' ------------------------------------------
   ' Get all but the last portion of the
   '     subkey
   ' ------------------------------------------
   If InStr(1, xi_strSubKey, "\", vbTextCompare) > 0 Then
      For p_lngLoop = Len(xi_strSubKey) To 1 Step -1
         If Mid$(xi_strSubKey, p_lngLoop, 1) = "\" Then
            p_strSubKey = Mid$(xi_strSubKey, 1, p_lngLoop - 1)
            p_strToDelete = Mid$(xi_strSubKey, p_lngLoop + 1)
            Exit For
         End If
      Next p_lngLoop
   Else
      Exit Function
   End If
   
   ' ------------------------------------------
   ' Open the key with all rights
   ' ------------------------------------------
   p_lngRtn = RegOpenKeyEx(p_lngKey, p_strSubKey, 0&, KEY_ALL_ACCESS, p_lngKeyHwnd)
   If p_lngRtn <> ERROR_SUCCESS Then
      Exit Function
   Else
      ' ---------------------------------------
      ' p_lngKeyHwnd is handle to the open key
      ' ---------------------------------------
   End If
   
   ' ------------------------------------------
   ' Delete the subkey
   ' ------------------------------------------
   p_lngRtn = RegDeleteKey(p_lngKeyHwnd, p_strToDelete)
   If p_lngRtn <> ERROR_SUCCESS Then
      DeleteRegistryKey = False
   Else
      DeleteRegistryKey = True
   End If
   
   ' ------------------------------------------
   ' Close the key
   ' ------------------------------------------
   p_lngRtn = RegCloseKey(p_lngKeyHwnd)
   
End Function

' *******************************************************
' Routine Name : (PUBLIC in CLASS) SetRegistryValue
' Written By   : L.J. Johnson
' Inputs       : xi_strSubKey:String     - Name of the key to set the value under
'              : xi_strValueName:String  - Name of the value to set under the subkey
'              : xi_strValue:String      - Value to set
'              : xi_lngRegistryHive:enumRegistryHives   - (Optional) ID of the hive to use (defaults to HKEY_LOCAL_MACHINE)
'              : xi_lngValueType:enumRegistryValueTypes - (Optional) Type of value to set (defaults to REG_SZ)
' Ouputs       : N/A
' Example(s)   : strSubKey = "Software\MyProg\Services\Template Location"
'              : strValueName = "UNC Path"
'              : strValue = "C:\Template"
'              : SetRegistryValue strSubKey, strValueName, strValue
' Description  : Set a Registry value -- defaults to HKEY_LOCAL_MACHINE
'              :     and values of type REG_SZ (string)
' *******************************************************
Public Sub SetRegistryValue(ByVal xi_strSubKey As String, _
                            ByVal xi_strValueName As String, _
                            ByVal xi_strValue As String, _
                            Optional ByVal xi_lngRegistryHive As Long, _
                            Optional ByVal xi_lngValueType As Long)
On Error Resume Next                   ' Don't accept errors here
   Const p_constProcName               As String = "SetRegistryValue"
   Dim p_strData                       As String
   Dim p_lngRtn                        As Long
   Dim p_lngKey                        As Long
   Dim p_lngValueType                  As Long
   Dim p_lngKeyHwnd                    As Long
   Dim p_lngDisposition                As Long
   Dim p_abytValue()                   As Byte
   
   ' ------------------------------------------
   ' Create under the passed hive ID (defaults
   '     to HKEY_LOCAL_MACHINE)
   ' ------------------------------------------
   If xi_lngRegistryHive = 0 Then
      p_lngKey = HKEY_LOCAL_MACHINE
   Else
      p_lngKey = xi_lngRegistryHive
   End If
   
   ' ------------------------------------------
   ' Create value under the passed hive ID
   '     (defaults to REG_SZ -- regular string)
   ' ------------------------------------------
   If xi_lngValueType = 0 Then
      p_lngValueType = REG_SZ
   Else
      p_lngValueType = xi_lngValueType
   End If
   
   ' ------------------------------------------
   ' RegCreateKey will create key if doesn't
   '     exist, or open it if it does
   ' ------------------------------------------
   p_lngRtn = RegCreateKeyEx(p_lngKey, xi_strSubKey, 0&, REG_SZ, REG_OPTION_NON_VOLATILE, KEY_SET_VALUE, 0&, p_lngKeyHwnd, p_lngDisposition)
   If p_lngRtn <> ERROR_SUCCESS Then
      Exit Sub
   Else
      ' ---------------------------------------
      ' p_lngKeyHwnd is handle to the open key
      ' ---------------------------------------
      If p_lngDisposition = REG_CREATED_NEW_KEY Then
         Debug.Print "Created new key"
      ElseIf p_lngDisposition = REG_OPENED_EXISTING_KEY Then
         Debug.Print "Opened existing key"
      Else
         Debug.Print "Unknown disposition."
      End If
   End If
   
   If p_lngValueType = REG_SZ Then
      ' ------------------------------------------
      ' Convert the string (after conversion to
      '     ANSI) to byte array
      ' ------------------------------------------
      p_abytValue = StrConv(xi_strValue, vbFromUnicode)
       
      ' ------------------------------------------
      ' Now set the string value
      ' ------------------------------------------
      p_lngRtn = RegSetValueEx(p_lngKeyHwnd, xi_strValueName, 0&, p_lngValueType, p_abytValue(0), Len(xi_strValue))
      If p_lngRtn <> ERROR_SUCCESS Then
         Debug.Print Err.LastDllError
      End If
   ElseIf p_lngValueType = REG_DWORD Then
      ' ------------------------------------------
      ' Now set the numeric value
      ' ------------------------------------------
      'p_strData = "&H" & Hex$(xi_strValue) & "&"

      p_lngRtn = RegSetValueEx(p_lngKeyHwnd, xi_strValueName, 0&, p_lngValueType, CLng(xi_strValue), 4&)
      If p_lngRtn <> ERROR_SUCCESS Then
         Debug.Print Err.LastDllError
      End If
   End If
   
   ' ------------------------------------------
   ' Close the handle to the subkey
   ' ------------------------------------------
   p_lngRtn = RegCloseKey(p_lngKeyHwnd)

On Error GoTo 0
End Sub

' *******************************************************
' Routine Name : (PUBLIC in CLASS) GetRegistryValue
' Written By   : L.J. Johnson
' Inputs       : xi_strSubKey:String     - Name of the key to get the value from
'              : xi_strValueName:String  - Name of the value to get from the subkey
'              : xi_lngRegistryHive:enumRegistryHives - (Optional) ID of the hive to use (defaults to HKEY_LOCAL_MACHINE)
' Ouputs       : Variant: Value requested from registry
' Example(s)   : strSubKey = "Software\MyProg\Services\Template Location"
'              : strValueName = "UNC Path"
'              : Debug.Print GetRegistryValue(strSubKey, strValueName)
' Description  : Retrieve a Registry value -- defaults to HKEY_LOCAL_MACHINE
' *******************************************************
Public Function GetRegistryValue(ByVal xi_strSubKey As String, _
                                 ByVal xi_strValueName As String, _
                                 Optional ByVal xi_lngRegistryHive As Long) As Variant
On Error Resume Next                   ' Don't accept errors here
   Const p_constProcName               As String = "GetRegistryValue"
   Dim p_strRtnValue                   As String
   Dim p_lngRtn                        As Long
   Dim p_lngKey                        As Long
   Dim p_lngKeyHwnd                    As Long
   Dim p_lngRtnValue                   As Long
   Dim p_lngValueType                  As Long
   Dim p_vntValue                      As Variant
   Dim p_dblTempValue                  As Double
   Dim p_lngPos                        As Long
   
   ' ------------------------------------------
   ' Create under the passed hive ID (defaults
   '     to HKEY_LOCAL_MACHINE)
   ' ------------------------------------------
   If xi_lngRegistryHive = 0 Then
      p_lngKey = HKEY_LOCAL_MACHINE
   Else
      p_lngKey = xi_lngRegistryHive
   End If
   
   
   p_lngRtn = RegOpenKey(p_lngKey, xi_strSubKey, p_lngKeyHwnd)
   If p_lngRtn <> ERROR_SUCCESS Then
      Exit Function
   Else
      ' ---------------------------------------
      ' p_lngKeyHwnd is handle to the open key
      ' ---------------------------------------
   End If
   
   ' ------------------------------------------
   ' Set the variables to pass
   ' ------------------------------------------
   p_strRtnValue = String(MAX_PATH, Chr$(0))
   p_lngRtnValue = Len(p_strRtnValue)
   
   ' ------------------------------------------
   ' Get the value
   ' ------------------------------------------
   p_lngRtn = RegQueryValueEx(p_lngKeyHwnd, xi_strValueName, 0&, p_lngValueType, ByVal p_strRtnValue, p_lngRtnValue)
   
   ' ------------------------------------------
   ' Close the handle to the subkey
   ' ------------------------------------------
   p_lngRtn = RegCloseKey(p_lngKeyHwnd)
   
   ' ------------------------------------------
   ' Convert byte array into string
   ' ------------------------------------------
   If p_lngValueType = REG_DWORD Then
      p_vntValue = p_strRtnValue
      
      p_dblTempValue = Asc(Mid$(p_vntValue, 1, 1)) + &H100& * Asc(Mid$(p_vntValue, 2, 1)) + &H10000 * Asc(Mid$(p_vntValue, 3, 1)) + &H1000000 * CDbl(Asc(Mid$(p_vntValue, 4, 1)))
      If p_dblTempValue > &H7FFFFFFF Then
          p_vntValue = "&H" & Hex$(p_dblTempValue - 4294967296#)
      Else
          p_vntValue = "&H" & Hex$(p_dblTempValue)
      End If
   Else
      ' ---------------------------------------
      ' Must be a string - trim any trailing null
      ' ---------------------------------------
      p_lngPos = InStr(1, p_strRtnValue, Chr$(0))
      If p_lngPos > 0 Then
         p_vntValue = Mid$(p_strRtnValue, 1, p_lngPos - 1)
      End If
   
   End If
   
   ' ------------------------------------------
   ' Set the return value
   ' ------------------------------------------
   GetRegistryValue = p_vntValue

On Error GoTo 0
End Function
