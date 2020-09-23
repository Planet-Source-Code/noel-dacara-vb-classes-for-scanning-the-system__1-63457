VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAdvanceRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Author         : Noel A. Dacara (noeldacara@yahoo.com)
'Filename       : Advance Registry.cls (cAdvanceRegistry Class Module)
'Description    : Advance registry functions for VB programmers
'Date           : Wednesday, February 2, 2005, 11:02 PM
'Last Update    : Thursday, November 24, 2005, 02:14 PM

'You can freely use and distribute this class or upload these codes on any site
'provided that the original credits are kept unmodified.

'User Caution   :
'Incorrectly editing the registry may severely damage your system.
'Before using this module, you should back up any valued data on your computer.

'Keep note that :
'If the remote registry is on a system that is running Windows NT 4.0, 2000, or XP,
'you must run the code from an account that has permission to read that registry.

'Be carefull in using the DeleteTree function. Backup your registry for more assurance.
'The author will not be held liable for any damages due to improper use of this module.

'Binary typed registry data are returned in an array of ascii numbers depending on how
'the data is written in the registry. Binary in unicode format can be read as strings.

'Most arguments are on a variant type (sorry???)...
'IsMissing command only works for variant type variables, so I decided to use variants.

'Credits goes to:
'Makers of the great Win32 Programmer's Reference, don't know who you are but thanks.
'Christoph von Wittich (Christoph@ActiveVB.de), author of ApiViewer 2004 for the APIs

'Sample usage   :
'Dim KeyArray() As String
'If REG.EnumerateKeys(KeyArray, HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows") Then
'   For i = LBound(KeyArray) To UBound(KeyArray)
'       Debug.Print KeyArray(i)
'   Next i
'End If

'API declarations
Private Declare Function ExpandEnvironmentStrings Lib "kernel32.dll" Alias "ExpandEnvironmentStringsA" (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegConnectRegistry Lib "advapi32.dll" Alias "RegConnectRegistryA" (ByVal lpMachineName As String, ByVal hKey As Long, ByRef phkResult As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES, ByRef phkResult As Long, ByRef lpdwDisposition As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByRef lpData As Any, ByRef lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Any, ByVal cbData As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'Modified API declarations
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Long, ByRef Source As Long, ByVal Length As Long)
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, ByRef lpcbValueName As Long, ByVal lpReserved As Long, ByVal lpType As Long, ByVal lpData As Byte, ByVal lpcbData As Long) As Long
Private Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey As Long, ByVal lpClass As String, ByRef lpcbClass As Long, ByVal lpReserved As Long, ByRef lpcSubKeys As Long, ByRef lpcbMaxSubKeyLen As Long, ByRef lpcbMaxClassLen As Long, ByRef lpcValues As Long, ByRef lpcbMaxValueNameLen As Long, ByRef lpcbMaxValueLen As Long, ByRef lpcbSecurityDescriptor As Long, ByRef lpftLastWriteTime As FILETIME) As Long
Private Declare Function RegQueryValueExByte Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal lKey As Long, ByVal lpValueName As String, ByVal lReserved As Long, ByRef lpType As Long, ByRef lpData As Byte, ByRef lpcbData As Long) As Long
Private Declare Function RegQueryValueExDouble Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal lKey As Long, ByVal lpValueName As String, ByVal lReserved As Long, ByRef lpType As Long, ByRef lpData As Double, ByRef lpcbData As Long) As Long
Private Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal lKey As Long, ByVal lpValueName As String, ByVal lReserved As Long, ByRef lpType As Long, ByRef lpData As Long, ByRef lpcbData As Long) As Long
Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal lKey As Long, ByVal lpValueName As String, ByVal lReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegSetValueExByte Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Byte, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExDouble Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Double, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Long, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long

'Registry security option constants
Private Const KEY_CREATE_LINK           As Long = &H20
Private Const KEY_CREATE_SUB_KEY        As Long = &H4
Private Const KEY_ENUMERATE_SUB_KEYS    As Long = &H8
Private Const KEY_NOTIFY                As Long = &H10
Private Const KEY_QUERY_VALUE           As Long = &H1
Private Const KEY_SET_VALUE             As Long = &H2
Private Const STANDARD_RIGHTS_ALL       As Long = &H1F0000
Private Const SYNCHRONIZE               As Long = &H100000
Private Const KEY_ALL_ACCESS            As Long = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))

'Registry create type values
Private Const REG_OPTION_NON_VOLATILE   As Long = 0     'Key is preserved on system reboot
Private Const REG_OPTION_RESERVED       As Long = 0     'Parameter is reserved for future use

'Registry create/open disposition
Private Const REG_CREATED_NEW_KEY       As Long = &H1   'New registry key was created
Private Const MAX_CLASS_NAME            As Long = 255

'Registry return constants
Private Const ERROR_ACCESS_DENIED       As Long = 5&    'Access to registry key was denied
Private Const ERROR_MORE_DATA           As Long = 234   'More data is available
Private Const ERROR_NO_MORE_ITEMS       As Long = 259&  'No data is available
Private Const ERROR_SUCCESS             As Long = 0&    'Operation was completed successfully

'Other constaints
Private Const CLASS_NAME                As String = ".cAdvanceRegistry"

'API types
Private Type FILETIME
    dwLowDateTime                       As Long
    dwHighDateTime                      As Long
End Type

Private Type SECURITY_ATTRIBUTES
    nLength                             As Long
    lpSecurityDescriptor                As Long
    bInheritHandle                      As Long
End Type

'Public enums
Enum ERegistryRoots
    HKEY_CLASSES_ROOT = &H80000000  'Contains file association mappings
    HKEY_CURRENT_USER = &H80000001  'Contains information about the current user
    HKEY_LOCAL_MACHINE = &H80000002 'Contains computer specific information
    HKEY_USERS = &H80000003         'Contains individual preferences for each users
End Enum

#If False Then
    'Trick to preserve casing of these variables when used in VB IDE
    Private HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_LOCAL_MACHINE, HKEY_USERS
#End If

'Common registry data types
Enum ERegistryDataTypes
    REG_BINARY = 3                  'Binary data in any form
    REG_DWORD = 4                   '32-bit number
    REG_DWORD_BIG_ENDIAN = 5        '32-bit number in big-endian format
    REG_DWORD_LITTLE_ENDIAN = 4     '32-bit number in little-endian format
    REG_EXPAND_SZ = 2               'Contains unexpanded references to environment variables (Ex. %SystemRoot%=C:\Windows)
    REG_MULTI_SZ = 7                'String that contain lists or multiple values
    REG_QWORD = 11                  '64-bit number
    REG_QWORD_LITTLE_ENDIAN = 11    '64-bit number in little-endian format
    REG_SZ = 1                      'Standard string
    
    'Read-only purpose data types
    REG_FULL_RESOURCE_DESCRIPTOR = 9    '
    REG_LINK = 6                        'Unicode symbolic link
    REG_NONE = 0                        'No defined value type
    REG_RESOURCE_LIST = 8               'A device-driver resource list
    REG_RESOURCE_REQUIREMENTS_LIST = 10 '
End Enum

#If False Then
    'Trick to preserve casing of these variables when used in VB IDE
    Private REG_BINARY, REG_DWORD, REG_DWORD_BIG_ENDIAN, REG_DWORD_LITTLE_ENDIAN
    Private REG_EXPAND_SZ, REG_MULTI_SZ, REG_QWORD, REG_QWORD_LITTLE_ENDIAN
    Private REG_SZ, REG_FULL_RESOURCE_DESCRIPTOR, REG_LINK, REG_NONE
    Private REG_RESOURCE_LIST, REG_RESOURCE_REQUIREMENTS_LIST
#End If

'Variable declarations
Private m_ClassRoot         As ERegistryRoots
Private m_DataType          As ERegistryDataTypes
Private m_Default           As Variant
Private m_Key               As String
Private m_MachineName       As String
Private m_RemoteClassRoot   As ERegistryRoots
Private m_RemoteConnection  As Boolean
Private m_RemoteKey         As Long
Private m_ValueName         As String

Property Get ClassRoot() As ERegistryRoots
Attribute ClassRoot.VB_Description = "Returns/sets a value to determine the registry key root type used."
'Returns/sets a value to determine the registry key root type used.
    ClassRoot = m_ClassRoot
End Property

Property Let ClassRoot(Value As ERegistryRoots)
    If Value = 0 Then
        Err.Raise -1, App.EXEName & CLASS_NAME, "Invalid registry class root value."
        Exit Property
    End If
    
    m_ClassRoot = Value
End Property

Property Get DataType() As ERegistryDataTypes
Attribute DataType.VB_Description = "Returns a value to determine the data type used or set the data type to be used."
'Returns a value to determine the data type used or set the data type to be used.
    DataType = m_DataType
End Property

Property Let DataType(Value As ERegistryDataTypes)
    m_DataType = Value
End Property

Property Get Default() As Variant
Attribute Default.VB_Description = "Returns/sets the default registry value to be used incase of problems."
'Returns/sets the default registry value to be used incase of problems.
    Default = m_Default
End Property

Property Let Default(Value As Variant)
    m_Default = Value
End Property

Property Get Key() As String
Attribute Key.VB_Description = "Returns/sets a value to determines the registry key/section to be used."
'Returns/sets a value to determines the registry key/section to be used.
    Key = m_Key
End Property

Property Let Key(Value As String)
    m_Key = Value
End Property

Property Get MachineName() As String
Attribute MachineName.VB_Description = "Returns/sets the name of computer where to establish a remote registry connection."
'Returns/sets the name of computer where to establish a remote registry connection.
    MachineName = m_MachineName
End Property

Property Let MachineName(Value As String)
    Call ValidateMachineName(Value) 'make sure name is in the proper format
    m_MachineName = Value
End Property

Property Get Value() As Variant
Attribute Value.VB_Description = "Returns/sets the value of a registry data."
'Returns/sets the value of a registry data.
    Value = Me.ValueEx(m_ClassRoot, m_Key, m_ValueName)
End Property

Property Let Value(Value As Variant)
    Me.ValueEx(m_ClassRoot, m_Key, m_ValueName) = Value
End Property

Property Get ValueEx(Optional ClassRoot, Optional Key, Optional ValueName) As Variant
Attribute ValueEx.VB_Description = "Returns/sets the value of the specified registry data."
'Returns/sets the value of the specified registry data.
    If IsMissing(ClassRoot) Then
        ClassRoot = m_ClassRoot
    End If
    
    If IsMissing(Key) Then
        Key = m_Key
    End If
    
    If IsMissing(ValueName) Then
        ValueName = m_ValueName
    End If
    
    Dim p_Key       As Long
    Dim p_Ret       As Long
    Dim p_DataLen   As Long
    Dim p_DataType  As Long
    
    If m_RemoteConnection Then
        ClassRoot = m_RemoteKey
    End If
    
    p_Ret = RegOpenKeyEx(ClassRoot, Key, REG_OPTION_NON_VOLATILE, KEY_QUERY_VALUE, p_Key)
    p_Ret = RegQueryValueExLong(p_Key, ValueName, REG_OPTION_RESERVED, p_DataType, 0&, p_DataLen)
    
    If p_Ret And Not p_Ret = ERROR_MORE_DATA Then
        ValueEx = m_Default
        Exit Property
    End If
    
    m_DataType = p_DataType
    Select Case p_DataType
        Case REG_SZ, REG_MULTI_SZ, REG_EXPAND_SZ
            Dim p_Buffer As String
            p_Buffer = String$(p_DataLen, 0)
            
            p_Ret = RegQueryValueExString(p_Key, ValueName, REG_OPTION_RESERVED, p_DataType, p_Buffer, p_DataLen)
            
            If p_DataType = REG_EXPAND_SZ Then
                ValueEx = ExpandEnvironmentString(p_Buffer)
            Else
                If p_DataLen > 0 Then
                    ValueEx = Left$(p_Buffer, p_DataLen - 1)
                End If
            End If
        Case REG_DWORD, REG_DWORD_LITTLE_ENDIAN, REG_DWORD_BIG_ENDIAN
            Dim p_Dword As Long
            
            p_Ret = RegQueryValueExLong(p_Key, ValueName, REG_OPTION_RESERVED, p_DataType, p_Dword, p_DataLen)
            
            If p_DataType = REG_DWORD Then
                ValueEx = CLng(p_Dword)
            Else
                ValueEx = GetBigEndianValue(p_Dword) 'REG_DWORD_BIG_ENDIAN
            End If
        Case REG_QWORD, REG_QWORD_LITTLE_ENDIAN
            Dim p_Qword As Double
            
            p_Ret = RegQueryValueExDouble(p_Key, ValueName, REG_OPTION_RESERVED, p_DataType, p_Qword, p_DataLen)
            ValueEx = CDbl(p_Qword)
        Case Else
            Dim p_Binary() As Byte
            If p_DataLen > 0 Then
                ReDim p_Binary(p_DataLen - 1) 'without the terminating null character
                
                p_Ret = RegQueryValueExByte(p_Key, ValueName, REG_OPTION_RESERVED, p_DataType, p_Binary(0), p_DataLen)
            End If
            ValueEx = p_Binary 'Return as an array of characters
    End Select
    
    Call RegCloseKey(p_Key)
End Property

Property Let ValueEx(Optional ClassRoot, Optional Key, Optional ValueName, Value As Variant)
    If IsMissing(ClassRoot) Then
        ClassRoot = m_ClassRoot
    End If
    
    If IsMissing(Key) Then
        Key = m_Key
    End If
    
    If IsMissing(ValueName) Then
        ValueName = m_ValueName
    End If
    
    Dim p_Key       As Long
    Dim p_Len       As Long
    Dim p_Ret       As Long
    Dim p_SecAttrib As SECURITY_ATTRIBUTES
    
    If m_RemoteConnection Then
        ClassRoot = m_RemoteKey
    End If
    
    p_Ret = RegCreateKeyEx(ClassRoot, Key, REG_OPTION_RESERVED, "", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, p_SecAttrib, p_Key, 0&)
    
    If p_Ret = ERROR_SUCCESS Then
        Select Case m_DataType
            Case REG_BINARY
                Dim p_Byte() As Byte
                If VarType(Value) = vbArray + vbByte Or vbString Then 'Can also accept strings
                    p_Byte = Value
                Else
                    Err.Raise m_DataType, App.EXEName & CLASS_NAME, "Cannot set the specified value in the registry using the defined data type."
                End If
                p_Len = UBound(p_Byte) - LBound(p_Byte) + 1
                
                p_Ret = RegSetValueExByte(p_Key, ValueName, REG_OPTION_RESERVED, m_DataType, p_Byte(0), p_Len)
            Case REG_SZ, REG_MULTI_SZ, REG_EXPAND_SZ
                Dim p_String As String
                p_String = Value
                
                If m_DataType = REG_MULTI_SZ Then
                    'Replace all linefeeds with null characters
                    p_String = Replace$(p_String, vbNewLine, vbNullChar)
                    
                    'Windows® Registry Editor does not allow REG_MULTI_SZ
                    'typed data to contain empty lines so, so as we...
                    While InStr(1, p_String, String$(2, vbNullChar)) > 0&
                        p_String = Replace$(p_String, String$(2, vbNullChar), vbNullChar)
                    Wend
                    
                    'REG_MULTI_SZ typed data is terminated by two null characters
                    If Right$(p_String, 1) = vbNullChar Then
                        p_String = p_String & vbNullChar
                    End If
                End If
                
                p_String = p_String & vbNullChar 'Terminate with null character
                p_Len = Len(p_String)
                
                p_Ret = RegSetValueExString(p_Key, ValueName, REG_OPTION_RESERVED, m_DataType, p_String, p_Len)
            Case REG_DWORD, 5 'REG_DWORD_LITTLE_ENDIAN, REG_DWORD_BIG_ENDIAN
                If VarType(Value) = vbInteger Or vbLong Then
                    Dim p_Dword As Long
                    p_Dword = CLng(Value)
                    
                    p_Len = 4& '32-bits
                    
                    p_Ret = RegSetValueExLong(p_Key, ValueName, REG_OPTION_RESERVED, m_DataType, p_Dword, p_Len)
                Else
                    Err.Raise m_DataType, App.EXEName & CLASS_NAME, "Cannot set the specified value in the registry using the defined data type."
                End If
            Case REG_QWORD 'REG_QWORD_LITTLE_ENDIAN
                If VarType(Value) = vbInteger Or vbLong Or vbDouble Then
                    Dim p_Qword As Double
                    p_Qword = CDbl(Value)
                    
                    p_Len = 8& '64-bits
                    
                    p_Ret = RegSetValueExDouble(p_Key, ValueName, REG_OPTION_RESERVED, m_DataType, p_Qword, p_Len)
                Else
                    Err.Raise m_DataType, App.EXEName & CLASS_NAME, "Cannot set the specified value in the registry using the defined data type."
                End If
            Case Else
                Err.Raise m_DataType, App.EXEName & CLASS_NAME, "The specified data type is either invalid or not supported for write purpose."
        End Select
        
        Call RegCloseKey(p_Key)
    Else
        Err.Raise ClassRoot, App.EXEName & CLASS_NAME, "Unable to open/create registry key: '" & Key & "' for setting new value of '" & ValueName & "' to '" & Value & "'"
    End If
End Property

Property Get ValueName() As String
Attribute ValueName.VB_Description = "Returns/sets a name to identify a particular registry data."
'Returns/sets a name to identify a particular registry data.
    ValueName = m_ValueName
End Property

Property Let ValueName(Value As String)
    m_ValueName = Value
End Property

'Public procedures
Function CreateKey(Optional ClassRoot, Optional Key) As Boolean
Attribute CreateKey.VB_Description = "Creates a new registry key from a specified path of a registry root."
'Creates a new registry key from a specified path of a registry root.
    If IsMissing(ClassRoot) Then
        ClassRoot = m_ClassRoot
    End If
    
    If IsMissing(Key) Then
        Key = m_Key
    End If
    
    Dim p_Key       As Long
    Dim p_Ret       As Long
    Dim p_SecAttrib As SECURITY_ATTRIBUTES
    
    If m_RemoteConnection Then
        ClassRoot = m_RemoteKey
    End If
    
    p_Ret = RegCreateKeyEx(ClassRoot, Key, REG_OPTION_RESERVED, Empty, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, p_SecAttrib, p_Key, 0&)
    
    If p_Ret = ERROR_SUCCESS Then
        CreateKey = True
        Call RegCloseKey(p_Key)
    Else
        Err.Raise ClassRoot, App.EXEName & CLASS_NAME, "Unable to create registry key: '" & Key & "'"
    End If
End Function

Function DeleteAllData(Optional ClassRoot, Optional Key) As Boolean
Attribute DeleteAllData.VB_Description = "Remove and clear all data of the specified registry key."
'Remove and clear all data of the specified registry key.
    If IsMissing(ClassRoot) Then
        ClassRoot = m_ClassRoot
    End If
    
    If IsMissing(Key) Then
        Key = m_Key
    End If
    
    Dim p_DataArray()   As String
    Dim p_Ctr           As Long
    
    If Me.EnumerateData(p_DataArray, ClassRoot, Key) Then
        For p_Ctr = LBound(p_DataArray) To UBound(p_DataArray)
            Me.DeleteData ClassRoot, Key, p_DataArray(p_Ctr)
        Next p_Ctr
        
        DeleteAllData = True
    End If
End Function

Function DeleteData(Optional ClassRoot, Optional Key, Optional ValueName) As Boolean
Attribute DeleteData.VB_Description = "Removes the specified data on a particular registry key and class root."
'Removes the specified data on a particular registry key and class root.
    If IsMissing(ClassRoot) Then
        ClassRoot = m_ClassRoot
    End If
    
    If IsMissing(Key) Then
        Key = m_Key
    End If
    
    If IsMissing(ValueName) Then
        ValueName = m_ValueName
    End If
    
    Dim p_Key As Long
    Dim p_Ret As Long
    
    If m_RemoteConnection Then
        ClassRoot = m_RemoteKey
    End If
    
    p_Ret = RegOpenKeyEx(ClassRoot, Key, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, p_Key)
    
    If p_Ret = ERROR_SUCCESS Then
        p_Ret = RegDeleteValue(p_Key, ValueName)
        
        If p_Ret = ERROR_SUCCESS Then
            DeleteData = True
        'Else
            'Err.Raise ClassRoot, App.EXEName & CLASS_NAME, "Unable to delete '" & ValueName & "' in '" & Key & "'"
        End If
        
        Call RegCloseKey(p_Key)
    Else
        Err.Raise ClassRoot, App.EXEName & CLASS_NAME, "Unable to open registry key '" & Key & "' for delete access."
    End If
End Function

Function DeleteKey(Optional ClassRoot, Optional Key) As Boolean
Attribute DeleteKey.VB_Description = "Removes the specified key on a particular class root in the registry."
'Removes the specified key on a particular class root in the registry.
    If IsMissing(ClassRoot) Then
        ClassRoot = m_ClassRoot
    End If
    
    If IsMissing(Key) Then
        Key = m_Key
    End If
    
    Dim p_Key       As Long
    Dim p_ChildKey  As String
    Dim p_ParentKey As String
    Dim p_Ret       As Long
    
    If m_RemoteConnection Then
        ClassRoot = m_RemoteKey
    End If
    
    p_ParentKey = GetParentKey(Key)  'Get parent key
    p_Ret = RegOpenKeyEx(ClassRoot, p_ParentKey, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, p_Key)
    
    If p_Ret = ERROR_SUCCESS Then
        If p_ParentKey = Empty Then
            p_ParentKey = Key
        Else
            p_ChildKey = Mid$(Key, Len(p_ParentKey) + 2)
        End If
        
        p_Ret = RegDeleteKey(p_Key, p_ChildKey)
        
        If p_Ret = ERROR_SUCCESS Then
            DeleteKey = True
        'Else
            'Err.Raise ClassRoot, App.EXEName & CLASS_NAME, "Unable to delete registry key: '" & Key & "'"
        End If
        
        Call RegCloseKey(p_Key)
    Else
        Err.Raise ClassRoot, App.EXEName & CLASS_NAME, "Unable to open registry key '" & p_ParentKey & "' for delete access."
    End If
End Function

Function DeleteTree(Optional ClassRoot, Optional Key) As Boolean
Attribute DeleteTree.VB_Description = "Remove the specified registry key and all of its subkeys."
'Remove the specified registry key and all of its subkeys.
    If IsMissing(ClassRoot) Then
        ClassRoot = m_ClassRoot
    End If
    
    If IsMissing(Key) Then
        Key = m_Key
    End If
    
    Dim p_KeyArray() As String
    Dim p_CurrentKey As String
    Dim p_HasSubKey As Boolean
    
    p_CurrentKey = Key 'Set current key
    
    Do
        If p_HasSubKey Then
            p_CurrentKey = p_CurrentKey & "\" & p_KeyArray(0) 'Go deeper
        Else
            'Check if current key is a subkey
            If Not p_CurrentKey = Key Then
                p_CurrentKey = GetParentKey(p_CurrentKey)
            End If
        End If
        
        'Ensure that the system can still process events other than this function
        DoEvents
        Call DeleteSubKeys(ClassRoot, p_CurrentKey)
        
        'Determine and get if current key has any subkeys
        p_HasSubKey = Me.EnumerateKeys(p_KeyArray, ClassRoot, p_CurrentKey)
        
        If p_CurrentKey = Key And Not p_HasSubKey Then
            Exit Do 'This is where we get out the loop
        End If
    Loop 'Does it looks like an infinite loop? hehe...
    
    If Me.DeleteKey(ClassRoot, Key) Then 'Attempt to delete the main key
        DeleteTree = True 'If successful, then so as this function...
    End If
End Function

Function EnumerateData(DataArray() As String, Optional ClassRoot, Optional Key) As Boolean
Attribute EnumerateData.VB_Description = "Enumerate the list of valuenames found on a specified registry key."
'Enumerate the list of valuenames found on a specified registry key.
    If IsMissing(ClassRoot) Then
        ClassRoot = m_ClassRoot
    End If
    
    If IsMissing(Key) Then
        Key = m_Key
    End If
    
    Dim p_Key As Long
    Dim p_Ret As Long
    
    If m_RemoteConnection Then
        ClassRoot = m_RemoteKey
    End If
    
    p_Ret = RegOpenKeyEx(ClassRoot, Key, REG_OPTION_NON_VOLATILE, KEY_QUERY_VALUE, p_Key)
    
    If p_Ret = ERROR_SUCCESS Then
        Dim p_Buffer    As String
        Dim p_DataCount As Long
        Dim p_DataLen   As Long
        Dim p_Index     As Long
        Dim p_MaxLen    As Long
        
        Dim FT As FILETIME
        p_Ret = RegQueryInfoKey(p_Key, "", 0&, REG_OPTION_RESERVED, 0&, 0&, 0&, p_DataCount, p_MaxLen, 0&, 0&, FT)
        
        If p_DataCount > 0 Then
            ReDim DataArray(0 To (p_DataCount - 1)) As String
            
            Do
                p_DataLen = p_MaxLen + 1
                p_Buffer = String$(p_DataLen, 0)
                p_Ret = RegEnumValue(p_Key, p_Index, p_Buffer, p_DataLen, REG_OPTION_RESERVED, 0&, 0&, 0&)
                
                If p_Ret = ERROR_SUCCESS Then
                    DataArray(p_Index) = Left$(p_Buffer, p_DataLen)
                 End If
                
                p_Index = p_Index + 1
            Loop While (p_Ret = ERROR_SUCCESS) And (p_Index < p_DataCount)
            
            EnumerateData = True
        'Else
            'Err.Raise ClassRoot, App.EXEName & CLASS_NAME, "Registry key '" & Key & "' does not contain any data."
        End If
        
        Call RegCloseKey(p_Key)
    Else
        Err.Raise ClassRoot, App.EXEName & CLASS_NAME, "Unable to open registry key '" & Key & "' for data query access."
    End If
End Function

Function EnumerateKeys(KeyArray() As String, Optional ClassRoot, Optional Key) As Boolean
Attribute EnumerateKeys.VB_Description = "Enumerate the list of subkeys found on a specified registry key."
'Enumerate the list of subkeys found on a specified registry key.
    If IsMissing(ClassRoot) Then
        ClassRoot = m_ClassRoot
    End If
    
    If IsMissing(Key) Then
        Key = m_Key
    End If
    
    Dim p_Key As Long
    Dim p_Ret As Long
    
    If m_RemoteConnection Then
        ClassRoot = m_RemoteKey
    End If
    
    p_Ret = RegOpenKeyEx(ClassRoot, Key, REG_OPTION_NON_VOLATILE, KEY_QUERY_VALUE + KEY_ENUMERATE_SUB_KEYS, p_Key)
    
    If p_Ret = ERROR_SUCCESS Then
        Dim p_Buffer    As String
        Dim p_Index     As Long
        Dim p_KeyCount  As Long
        Dim p_MaxLen    As Long
        Dim p_NullPos   As Long
        
        Dim FT As FILETIME
        p_Ret = RegQueryInfoKey(p_Key, "", 0&, REG_OPTION_RESERVED, p_KeyCount, p_MaxLen, 0&, 0&, 0&, 0&, 0&, FT)
        
        If p_KeyCount > 0 Then
            ReDim KeyArray(0 To (p_KeyCount - 1)) As String
            p_MaxLen = p_MaxLen + 1 'for the terminating null character
            
            Do
                p_Buffer = String$(p_MaxLen, 0)
                p_Ret = RegEnumKey(p_Key, p_Index, p_Buffer, p_MaxLen)
                
                If p_Ret = ERROR_SUCCESS Then
                    p_NullPos = InStr(1, p_Buffer, vbNullChar)
                    KeyArray(p_Index) = Left$(p_Buffer, p_NullPos - 1)
                End If
                
                p_Index = p_Index + 1
            Loop While (p_Ret = ERROR_SUCCESS) And (p_Index < p_KeyCount)
            
            EnumerateKeys = True
        'Else
            'Err.Raise ClassRoot, App.EXEName & CLASS_NAME, "Registry key '" & Key & "' does not contain any subkeys."
        End If
        
        Call RegCloseKey(p_Key)
    Else
        Err.Raise ClassRoot, App.EXEName & CLASS_NAME, "Unable to open registry key '" & Key & "' for key query access."
    End If
End Function

Sub Export(File As String, Optional Overwrite As Boolean, Optional ClassRoot, Optional Key)
Attribute Export.VB_Description = "Attempt to export the contents of a registry key and all of its subkeys to a file."
'Attempt to export the contents of a registry key and all of its subkeys to a file.
    If Not Overwrite Then
        If Not Len(Dir$(File)) = 0 Then
            Exit Sub 'Do not overwrite existing file
        End If
    End If
    
    If IsMissing(Key) Then
        Key = m_Key
    End If
    
    If IsMissing(ClassRoot) Then
        ClassRoot = m_ClassRoot
    End If
    
    If IsNumeric(ClassRoot) Then
        Select Case ClassRoot
            Case HKEY_CLASSES_ROOT
                ClassRoot = "HKEY_CLASSES_ROOT"
            Case HKEY_CURRENT_USER
                ClassRoot = "HKEY_CURRENT_USER"
            Case HKEY_LOCAL_MACHINE
                ClassRoot = "HKEY_LOCAL_MACHINE"
            Case HKEY_USERS
                ClassRoot = "HKEY_USERS"
            Case 0
                Err.Raise -1, App.EXEName & CLASS_NAME, "Registry root class not accepted."
                Exit Sub
            Case Else
                Err.Raise ClassRoot, App.EXEName & CLASS_NAME, "Registry root class not supported."
                Exit Sub
        End Select
    End If
    
    Dim p_Param As String
    p_Param = ClassRoot & "\" & Key 'Build registry path
    
    'Export via Windows® built-in Registry Editor
    ShellExecute 0&, "", "regedit", "/e """ & File & """ " & p_Param, App.Path, vbHide
End Sub

Function HasSubKey(Optional ClassRoot, Optional Key) As Boolean
Attribute HasSubKey.VB_Description = "Returns a value to determine if the registry path supplied contains subkeys."
'Returns a value to determine if the registry path supplied contains subkeys.
    If IsMissing(ClassRoot) Then
        ClassRoot = m_ClassRoot
    End If
    
    If IsMissing(Key) Then
        Key = m_Key
    End If
    
    Dim p_Key As Long
    Dim p_Ret As Long
    
    If m_RemoteConnection Then
        ClassRoot = m_RemoteKey
    End If
    
    p_Ret = RegOpenKeyEx(ClassRoot, Key, REG_OPTION_NON_VOLATILE, KEY_ENUMERATE_SUB_KEYS, p_Key)
    
    If p_Ret = ERROR_SUCCESS Then
        Dim p_Buffer As String * MAX_CLASS_NAME
        
        p_Ret = RegEnumKey(p_Key, 0&, p_Buffer, MAX_CLASS_NAME) 'Attempt to get subkey
        
        If p_Ret = ERROR_SUCCESS Then
            HasSubKey = True
        End If
        
        Call RegCloseKey(p_Key)
    Else
        Err.Raise ClassRoot, App.EXEName & CLASS_NAME, "Unable to open registry key '" & Key & "' for enumerate subkeys access."
    End If
End Function

Sub Import(File As String)
Attribute Import.VB_Description = "Attempt to import a valid registry file to the registry."
'Attempt to import a valid registry file to the registry.
    File = Trim$(File) 'remove trailing and leading spaces
    
    If Not Len(Dir$(File)) = 0 Then
        ShellExecute 0&, "", "regedit", "/s """ & File & """", App.Path, vbHide
    Else
        Err.Raise 53, App.EXEName & CLASS_NAME
    End If
End Sub

Function KeyExists(Optional ClassRoot, Optional Key) As Boolean
Attribute KeyExists.VB_Description = "Returns/sets a value to determine if a key exists on the registry."
'Returns a value to determine if a key exists on the registry.
    If IsMissing(ClassRoot) Then
        ClassRoot = m_ClassRoot
    End If
    
    If IsMissing(Key) Then
        Key = m_Key
    End If
    
    Dim p_Key As Long
    Dim p_Ret As Long
    
    If m_RemoteConnection Then
        ClassRoot = m_RemoteKey
    End If
    
    p_Ret = RegOpenKeyEx(ClassRoot, Key, REG_OPTION_NON_VOLATILE, KEY_QUERY_VALUE, p_Key)
    
    If p_Ret = ERROR_SUCCESS Or p_Ret = ERROR_ACCESS_DENIED Then
        KeyExists = True
        Call RegCloseKey(p_Key)
    End If
End Function

Function RemoteConnect(Optional MachineName, Optional ClassRoot) As Boolean
Attribute RemoteConnect.VB_Description = "Establish a remote registry connection."
'Establish a remote registry connection.
    If IsMissing(MachineName) Then
        MachineName = m_MachineName
    End If
    
    If IsMissing(ClassRoot) Then
        ClassRoot = m_ClassRoot
    End If
    
    Dim p_Key As Long
    Dim p_Ret As Long
    
    p_Ret = RegConnectRegistry(MachineName, ClassRoot, p_Key)
    
    If p_Ret = ERROR_SUCCESS Then
        m_RemoteKey = p_Key
        m_RemoteClassRoot = ClassRoot
        m_RemoteConnection = True
        RemoteConnect = True
    Else
        Err.Raise ClassRoot, App.EXEName & CLASS_NAME, "Unable to establish remote registry connection: '" & MachineName & "'"
    End If
End Function

Function RemoteDisconnect() As Boolean
Attribute RemoteDisconnect.VB_Description = "Disconnect from remote registry connection."
'Disconnect from remote registry connection.
    If m_RemoteConnection Then
        Dim p_Key As Long
        Dim p_Ret As Long
        
        p_Ret = RegCloseKey(m_RemoteKey)
        p_Ret = RegConnectRegistry("", m_RemoteClassRoot, p_Key) 'Connect from local system
        
        If p_Ret = ERROR_SUCCESS Then
            RegCloseKey p_Key 'Then close registry handle
            
            m_RemoteKey = 0&
            m_RemoteClassRoot = 0&
            m_RemoteConnection = False
            RemoteDisconnect = True
        Else
            Err.Raise m_RemoteClassRoot, App.EXEName & CLASS_NAME, "Unable to disconnect from remote registry connection."
        End If
    End If
End Function

Function ValueNameExists(Optional ClassRoot, Optional Key, Optional ValueName) As Boolean
Attribute ValueNameExists.VB_Description = "Returns/sets a value to determine if the specified valuename exists in a particular key."
'Returns a value to determine if the specified valuename exists in a particular key.
    If IsMissing(ClassRoot) Then
        ClassRoot = m_ClassRoot
    End If
    
    If IsMissing(Key) Then
        Key = m_Key
    End If
    
    If IsMissing(ValueName) Then
        ValueName = m_ValueName
    End If
    
    Dim p_Key As Long
    Dim p_Ret As Long
    
    If m_RemoteConnection Then
        ClassRoot = m_RemoteKey
    End If
    
    p_Ret = RegOpenKeyEx(ClassRoot, Key, REG_OPTION_NON_VOLATILE, KEY_QUERY_VALUE, p_Key)
    
    If p_Ret = ERROR_SUCCESS Then
        p_Ret = RegQueryValueExString(p_Key, ValueName, REG_OPTION_RESERVED, 0&, "", 0&)
        
        If p_Ret = ERROR_SUCCESS Or p_Ret = ERROR_MORE_DATA Then
            ValueNameExists = True
        End If
        
        Call RegCloseKey(p_Key)
    Else
        Err.Raise ClassRoot, App.EXEName & CLASS_NAME, "Unable to open registry key '" & Key & "' for key query access."
    End If
End Function

'Private procedures
Private Sub DeleteSubKeys(ByVal ClassRoot As String, Key As String)
'Deletes only subkeys with no child keys
    Dim p_KeyArray() As String
    
    If Me.EnumerateKeys(p_KeyArray, ClassRoot, Key) Then
        Dim p_Ctr As Long
        
        For p_Ctr = LBound(p_KeyArray) To UBound(p_KeyArray)
            'Attempt to delete every subkey found
            Me.DeleteKey ClassRoot, Key & "\" & p_KeyArray(p_Ctr)
        Next p_Ctr
    End If
End Sub

Private Function ExpandEnvironmentString(Value As String) As String
    Dim p_Buffer As String
    Dim p_Len As Long
    
    'Get length of expanded string
    p_Buffer = ""
    p_Len = ExpandEnvironmentStrings(Value, p_Buffer, p_Len)
    
    'Expand string
    p_Buffer = String$(p_Len, 0)
    p_Len = ExpandEnvironmentStrings(Value, p_Buffer, p_Len)
    
    If p_Len > 0 Then
        'Without the terminating null character
        ExpandEnvironmentString = Left$(p_Buffer, p_Len - 1)
    End If
End Function

Private Function GetParentKey(ByVal Value As String) As String
    If InStr(1, Value, "\") Then
        GetParentKey = Left$(Value, InStrRev(Value, "\") - 1) 'Get parent key
    End If
End Function

Private Function GetBigEndianValue(ByVal Value As Long) As Long
    CopyMemory ByVal VarPtr(GetBigEndianValue) + 3, Value, 1&
    CopyMemory ByVal VarPtr(GetBigEndianValue) + 2, ByVal VarPtr(Value) + 1, 1&
    CopyMemory ByVal VarPtr(GetBigEndianValue) + 1, ByVal VarPtr(Value) + 2, 1&
    'Debug.Print "GetBigEndianValue", GetBigEndianValue, Value
    CopyMemory GetBigEndianValue, ByVal VarPtr(Value) + 3, 1&
End Function

Private Sub ValidateMachineName(Value As String)
    If Not Len(Value) = 0 Then
        If InStr(1, Value, "\\") = 0 Then
            Value = "\\" & Value
        End If
    End If
End Sub

'Class Procedure
Private Sub Class_Initialize()
    'Unless these properties are set, these would be their default values
    m_ClassRoot = HKEY_CURRENT_USER
    m_DataType = REG_SZ
    
    'This is the required path format in writing software settings in the registry
    m_Key = "Software\" & App.CompanyName & "\" & App.ProductName
End Sub

'Created by Noel A. Dacara | Copyright © 2003-2005 Davao City, Philippines
