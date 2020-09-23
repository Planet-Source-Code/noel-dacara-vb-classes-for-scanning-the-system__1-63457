VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cScanRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Author         : Noel A. Dacara (noeldacara@yahoo.com)
'Filename       : Scan Registry.cls (cScanRegistry Class Module)
'Description    : Scan the registry for keys and data
'Date           : Friday, July 15, 2005, 01:58 AM
'Last Update    : Thursday, November 24, 2005, 02:22 PM

'You can freely use and distribute this class or upload these codes on any site
'provided that the original credits are kept unmodified.

'User Caution   :
'Incorrectly editing the registry may severely damage your system.
'Before using this module, you should backup any valued data on your computer.

'Keep note that :
'The Filter feature for keys is only applicable for 1st level keys
'after the specified start path. Start path is of course not included.
'Preceding subkeys are automatically included if permitted with ScanSubkeys property.

'Filter feature could also include wildcards and can be as many as you want.
'Separate filter patterns with the pipe character (|). Ex: Run*|File.???

'In custom scanning for file extensions, filter feature for keys is discarded.

'Scanning Tips  :
'To experience the maximum scanning speed of your machine with this module,
'it is advisable that you close other unnecessary programs running in the system.

'API Declarations
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long

'Modified API Declarations
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, ByRef lpcbValueName As Long, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As Long, ByVal lpcbData As Long) As Long
Private Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey As Long, ByVal lpClass As String, ByRef lpcbClass As Long, ByVal lpReserved As Long, ByRef lpcSubKeys As Long, ByRef lpcbMaxSubKeyLen As Long, ByRef lpcbMaxClassLen As Long, ByRef lpcValues As Long, ByRef lpcbMaxValueNameLen As Long, ByRef lpcbMaxValueLen As Long, ByRef lpcbSecurityDescriptor As Long, ByRef lpftLastWriteTime As FILETIME) As Long

'API Constants
Private Const ERROR_SUCCESS             As Long = 0&
Private Const KEY_CREATE_LINK           As Long = &H20
Private Const KEY_CREATE_SUB_KEY        As Long = &H4
Private Const KEY_ENUMERATE_SUB_KEYS    As Long = &H8
Private Const KEY_NOTIFY                As Long = &H10
Private Const KEY_QUERY_VALUE           As Long = &H1
Private Const KEY_SET_VALUE             As Long = &H2
Private Const MAX_PATH                  As Long = 260
Private Const REG_OPTION_NON_VOLATILE   As Long = 0
Private Const REG_OPTION_RESERVED       As Long = 0
Private Const STANDARD_RIGHTS_ALL       As Long = &H1F0000
Private Const SYNCHRONIZE               As Long = &H100000

Private Const KEY_ALL_ACCESS            As Long = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))

Private Type FILETIME
    dwLowDateTime   As Long
    dwHighDateTime  As Long
End Type

'User-Defined API Enum
Enum ECustomRegistryPaths
    SCAN_ADDREMOVELISTS = 0     'Add/Remove Programs
    SCAN_CUSTOMCONTROLS = 1     'Custom Controls
    SCAN_FILEEXTENSIONS = 2     'File Extensions
    SCAN_HELPRESOURCES = 3      'Help and Resources
    SCAN_SHAREDDLLS = 4         'Shared DLLs
    SCAN_SHELLFOLDERS = 5       'Registry Integrity
    SCAN_SOFTWAREPATHS = 6      'Software Locations
    SCAN_STARTUPKEYS = 7        'Startup Programs
    SCAN_WINDOWSFONTS = 8       'Windows Fonts
End Enum

#If False Then
    'Trick to preserve casing of these variables when used in VB IDE
    Private SCAN_ADDREMOVELISTS, SCAN_CUSTOMCONTROLS, SCAN_FILEEXTENSIONS
    Private SCAN_HELPRESOURCES, SCAN_SHAREDDLLS, SCAN_SHELLFOLDERS
    Private SCAN_SOFTWAREPATHS, SCAN_STARTUPKEYS, SCAN_WINDOWSFONTS
#End If

Enum EScanRegistryRoots
    SCAN_CLASSES_ROOT = &H80000000  'These 4 registry roots would be enough...
    SCAN_CURRENT_USER = &H80000001
    SCAN_LOCAL_MACHINE = &H80000002
    SCAN_USERS = &H80000003
End Enum

#If False Then
    'Trick to preserve casing of these variables when used in VB IDE
    Private SCAN_CLASSES_ROOT, SCAN_CURRENT_USER, SCAN_LOCAL_MACHINE, SCAN_USERS
#End If

'Variable Declarations
Private m_Cancel        As Boolean
Private m_CustomPath    As ECustomRegistryPaths
Private m_CustomScan    As Boolean
Private m_FilterData()  As String
Private m_FilterDataB   As Boolean
Private m_FilterKeys()  As String
Private m_FilterKeysB   As Boolean
Private m_ClassRoot     As EScanRegistryRoots
Private m_ScanDeep      As Integer
Private m_Scanning      As Boolean
Private m_ScanSubKeys   As Boolean
Private m_StartDeep     As Integer
Private m_ScanPath      As String
Private m_TotalData     As Long
Private m_TotalKeys     As Long

Event CurrentData(Value As String, Key As String, Root As EScanRegistryRoots, Delete As Boolean)
Attribute CurrentData.VB_Description = "Occurs when a registry data is scanned or processed."
'Occurs when a registry data is scanned or processed.
Event CurrentKey(Key As String, Root As EScanRegistryRoots, Delete As Boolean)
Attribute CurrentKey.VB_Description = "Occurs when a registry key has been scanned or processed for data."
'Occurs when a registry key has been scanned or processed for data.
Event DoneScanning(TotalData As Long, TotalKeys As Long)
Attribute DoneScanning.VB_Description = "Occurs when scanning is done either successfully or cancelled."
'Occurs when scanning is done either successfully or cancelled.

#If False Then
    'Trick to preserve casing of these variables when used in VB IDE
    Private Value, Key, Delete, Root, TotalData, TotalKeys
#End If

Property Get ClassRoot() As EScanRegistryRoots
Attribute ClassRoot.VB_Description = "Returns/sets a value to determine the registry key root type to scan."
'Returns/sets a value to determine the registry key root type to scan.
    ClassRoot = m_ClassRoot
End Property

Property Let ClassRoot(Value As EScanRegistryRoots)
    m_ClassRoot = Value
End Property

Property Get CustomScanPath() As ECustomRegistryPaths
Attribute CustomScanPath.VB_Description = "Returns/sets whether to scan custom registry keys for data."
'Returns/sets whether to scan custom registry keys for data.
    CustomScanPath = m_CustomPath
End Property

Property Let CustomScanPath(Value As ECustomRegistryPaths)
    m_CustomPath = Value
End Property

Property Get FilterData() As String
Attribute FilterData.VB_Description = "Returns/sets the filter pattern to scan for specific data."
'Returns/sets the filter pattern to scan for specific data.
    FilterData = Join(m_FilterData, "|")
End Property

Property Let FilterData(Value As String)
    m_FilterData = Split(LCase$(Value), "|")
    m_FilterDataB = (Not UBound(m_FilterData) = -1)
End Property

Property Get FilterKeys() As String
Attribute FilterKeys.VB_Description = "Returns/sets the filter pattern to scan for specific keys."
'Returns/sets the filter pattern to scan for specific keys.
    FilterKeys = Join(m_FilterKeys, "|")
End Property

Property Let FilterKeys(Value As String)
    m_FilterKeys = Split(LCase$(Value), "|")
    m_FilterKeysB = (Not UBound(m_FilterKeys) = -1)
End Property

Property Get ScanPath() As String
Attribute ScanPath.VB_Description = "Returns/sets the registry key where to start scanning."
'Returns/sets the registry key where to start scanning.
    ScanPath = m_ScanPath
End Property

Property Let ScanPath(Value As String)
    m_ScanPath = Value
End Property

Property Get ScanDeep() As Integer
Attribute ScanDeep.VB_Description = "Returns/sets a value to determine how deep from the startpath to scan."
'Returns/sets a value to determine how deep from the startpath to scan.
    ScanDeep = m_ScanDeep
End Property

Property Let ScanDeep(Value As Integer)
    m_ScanDeep = Value
End Property

Property Get Scanning() As Boolean
Attribute Scanning.VB_Description = "Returns True if class is currently on scanning state."
'Returns True if class is currently on scanning state.
    Scanning = m_Scanning
End Property

Property Get ScanSubKeys() As Boolean
Attribute ScanSubKeys.VB_Description = "Returns/sets whether to include subkeys when scanning."
'Returns/sets whether to include subkeys when scanning.
    ScanSubKeys = m_ScanSubKeys
End Property

Property Let ScanSubKeys(Value As Boolean)
    m_ScanSubKeys = Value
End Property

Property Get TotalData() As String
Attribute TotalData.VB_Description = "Returns the current number of data that has been scanned."
'Returns the current number of data that has been scanned.
    TotalData = m_TotalData
End Property

Property Get TotalKeys() As Long
Attribute TotalKeys.VB_Description = "Returns the current number of keys that has been scanned."
'Returns the current number of keys that has been scanned.
    TotalKeys = m_TotalKeys
End Property

Sub BeginCustomScan(Optional CustomScanPath)
Attribute BeginCustomScan.VB_Description = "Start scanning for keys and values on the predefined registry locations."
'Start scanning for keys and values on the predefined registry locations.
    If Not IsMissing(CustomScanPath) Then
        Me.CustomScanPath = CustomScanPath
    End If
    
    Dim p_FilterKeys    As String
    Dim p_ScanSubKeys   As Boolean
    
    p_FilterKeys = Me.FilterKeys
    p_ScanSubKeys = Me.ScanSubKeys
    Me.ScanSubKeys = True
    
    m_CustomScan = True
    Call InitializeCustomScanning
    m_CustomScan = False
    
    Me.FilterKeys = p_FilterKeys
    Me.ScanSubKeys = p_ScanSubKeys
    
    Call FinalizeScanning
End Sub

Sub BeginScanning(Optional ClassRoot, Optional ScanPath)
Attribute BeginScanning.VB_Description = "Start scanning for keys and values on the specified registry key."
'Start scanning for keys and values on the specified registry key.
    If Not IsMissing(ClassRoot) Then
        Me.ClassRoot = ClassRoot
    End If
    
    If Not IsMissing(ScanPath) Then
        Me.ScanPath = ScanPath
    End If
    
    Call InitializeScanning
    Call GetData(Me.ScanPath)
    
    If m_ScanSubKeys Then
        GetKeys (Me.ScanPath)
    End If
    
    If Not m_CustomScan Then
        Call FinalizeScanning
    End If
End Sub

Sub CancelScanning()
Attribute CancelScanning.VB_Description = "Abort the whole scanning process that has been started."
'Abort the whole scanning process that has been started.
    m_Cancel = True
End Sub

Sub ValidateKey(Key As String)
Attribute ValidateKey.VB_Description = "Validate a registry key from common typographical mistakes."
'Validate a registry key from common typographical mistakes.
    While InStr(1, Key, "\\")
        Key = Replace$(Key, "\\", "\")
    Wend
    
    If Key Like "*\" Then
        Key = Left$(Key, Len(Key) - 1)
    End If
    
    If Key Like "\*" Then
        Key = Mid$(Key, 2)
    End If
End Sub

Function DeleteValue(Root As EScanRegistryRoots, Path As String, Value As String) As Boolean
Attribute DeleteValue.VB_Description = "Delete a registry data specified. Returns True if successful."
'Delete a registry data specified. Returns True if successful.
    Dim p_Key As Long
    Dim p_Ret As Long
    
    p_Ret = RegOpenKeyEx(Root, Path, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, p_Key)
    
    If p_Ret = ERROR_SUCCESS Then
        p_Ret = RegDeleteValue(p_Key, Value)
        RegCloseKey p_Key
    End If
    
    DeleteValue = (p_Ret = ERROR_SUCCESS)
End Function

Function DeleteKey(Root As EScanRegistryRoots, Path As String) As Boolean
Attribute DeleteKey.VB_Description = "Attempt to delete registry key specified. Returns True if successful."
'Delete a registry key specified. Returns True if successful.
'Key must have no subkeys inorder for it to be deleted.
    Dim p_Key       As Long
    Dim p_ParentKey As String
    Dim p_Ret       As Long
   
    If Not InStr(1, Path, "\") = 0 Then
        p_ParentKey = Left$(Path, InStrRev(Path, "\") - 1)
    End If
    
    p_Ret = RegOpenKeyEx(Root, p_ParentKey, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, p_Key)
    
    If p_Ret = ERROR_SUCCESS Then
        p_ParentKey = Right$(Path, Len(Path) - Len(p_ParentKey) - 1)
        p_Ret = RegDeleteKey(p_Key, p_ParentKey)
        
        RegCloseKey p_Key
    End If
    
    DeleteKey = (p_Ret = ERROR_SUCCESS)
End Function

'Private Procedures
Private Sub InitializeCustomScanning()
    Dim p_ClassRoots()    As EScanRegistryRoots
    Dim p_ScanDeeps()   As Integer
    Dim p_ScanPaths()   As String
    Dim p_Ctr           As Integer
    
    Select Case m_CustomPath
        Case SCAN_ADDREMOVELISTS, SCAN_FILEEXTENSIONS, SCAN_HELPRESOURCES, SCAN_SHAREDDLLS, SCAN_SOFTWAREPATHS
            ReDim p_ClassRoots(0 To 0) As EScanRegistryRoots
            ReDim p_ScanDeeps(0 To 0) As Integer
            ReDim p_ScanPaths(0 To 0) As String
            
            p_ClassRoots(0) = SCAN_LOCAL_MACHINE
            p_ScanDeeps(0) = 1
            
            Select Case m_CustomPath
                Case SCAN_ADDREMOVELISTS
                    p_ScanPaths(0) = "Software\Microsoft\Windows\CurrentVersion\Uninstall"
                Case SCAN_FILEEXTENSIONS
                    p_ClassRoots(0) = SCAN_CLASSES_ROOT
                    p_ScanPaths(0) = ""
                    Me.FilterKeys = ".*|"
                Case SCAN_HELPRESOURCES
                    p_ScanPaths(0) = "Software\Microsoft\Windows\Help"
                Case SCAN_SHAREDDLLS
                    p_ScanPaths(0) = "Software\Microsoft\Windows\CurrentVersion\SharedDLLs"
                Case SCAN_SOFTWAREPATHS
                    p_ScanPaths(0) = "Software\Microsoft\Windows\CurrentVersion\App Paths"
            End Select
        Case SCAN_SHELLFOLDERS, SCAN_CUSTOMCONTROLS, SCAN_WINDOWSFONTS
            ReDim p_ClassRoots(0 To 1) As EScanRegistryRoots
            ReDim p_ScanDeeps(0 To 1) As Integer
            ReDim p_ScanPaths(0 To 1) As String
            
            Select Case m_CustomPath
                Case SCAN_CUSTOMCONTROLS
                    p_ClassRoots(0) = SCAN_CLASSES_ROOT
                    p_ClassRoots(1) = SCAN_CLASSES_ROOT
                    
                    p_ScanDeeps(0) = 0
                    p_ScanDeeps(1) = 0
                    
                    p_ScanPaths(0) = "CLSID"
                    p_ScanPaths(1) = "TypeLib"
                Case SCAN_SHELLFOLDERS
                    p_ClassRoots(0) = SCAN_CURRENT_USER
                    p_ClassRoots(1) = SCAN_LOCAL_MACHINE
                    
                    p_ScanDeeps(0) = 1
                    p_ScanDeeps(1) = 1
                    
                    p_ScanPaths(0) = "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
                    p_ScanPaths(1) = "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
                Case SCAN_WINDOWSFONTS
                    p_ClassRoots(0) = SCAN_LOCAL_MACHINE
                    p_ClassRoots(1) = SCAN_LOCAL_MACHINE
                    
                    p_ScanDeeps(0) = 1
                    p_ScanDeeps(1) = 1
                    
                    p_ScanPaths(0) = "Software\Microsoft\Windows\CurrentVersion\Fonts"
                    p_ScanPaths(1) = "Software\Microsoft\Windows NT\CurrentVersion\Fonts"
            End Select
        Case SCAN_STARTUPKEYS
            ReDim p_ClassRoots(0 To 9) As EScanRegistryRoots
            ReDim p_ScanDeeps(0 To 9) As Integer
            ReDim p_ScanPaths(0 To 9) As String
            
            For p_Ctr = 0 To 4
                p_ClassRoots(p_Ctr) = SCAN_CURRENT_USER
                p_ClassRoots(p_Ctr + 5) = SCAN_LOCAL_MACHINE
                
                p_ScanDeeps(p_Ctr) = 1
                p_ScanDeeps(p_Ctr + 5) = 1
                
                Select Case p_Ctr
                    Case 0: p_ScanPaths(p_Ctr) = "Software\Microsoft\Windows\CurrentVersion\Run"
                    Case 1: p_ScanPaths(p_Ctr) = "Software\Microsoft\Windows\CurrentVersion\RunOnce"
                    Case 2: p_ScanPaths(p_Ctr) = "Software\Microsoft\Windows\CurrentVersion\RunServices"
                    Case 3: p_ScanPaths(p_Ctr) = "Software\Microsoft\Windows\CurrentVersion\RunServicesOnce"
                    Case 4: p_ScanPaths(p_Ctr) = "Software\Microsoft\Windows NT\CurrentVersion\Windows"
                End Select
                
                p_ScanPaths(p_Ctr + 5) = p_ScanPaths(p_Ctr)
            Next p_Ctr
            
            p_ScanPaths(9) = "Software\Microsoft\Windows NT\CurrentVersion\Winlogon"
    End Select
    
    Call StartCustomScanning(p_ClassRoots, p_ScanPaths)
End Sub

Private Sub StartCustomScanning(p_ClassRoots() As EScanRegistryRoots, p_ScanPaths() As String)
    Dim p_Ctr As Integer
    
    m_TotalData = 0&
    m_TotalKeys = 0&
    m_Cancel = False 'Ensure that we can still scan successfully
    
    m_Scanning = True
    For p_Ctr = 0 To UBound(p_ClassRoots)
        If m_Cancel Then
            Exit Sub
        End If
        
        Me.ClassRoot = p_ClassRoots(p_Ctr)
        Me.ScanPath = p_ScanPaths(p_Ctr)
        
        Me.BeginScanning
    Next p_Ctr
End Sub

Private Sub InitializeScanning()
    If Not m_CustomScan Then
        m_TotalData = 0&
        m_TotalKeys = 0&
        m_Scanning = True
    End If
    
    m_Cancel = False 'Ensure that we can still scan successfully
    
    Call ValidateKey(Me.ScanPath)
    m_StartDeep = KeyDeep(ScanPath)
End Sub

Private Sub FinalizeScanning()
    m_Scanning = False
    RaiseEvent DoneScanning((m_TotalData), (m_TotalKeys))
    m_Cancel = False
End Sub

Private Function KeyDeep(Key As String) As Integer
    Dim i As Integer
    
    If Len(Key) Then
        KeyDeep = 1
    End If
    
    Do
        i = InStr(i + 1, Key, "\")
        If i Then
            KeyDeep = KeyDeep + 1 'Count number of separators
        End If
    Loop While i
End Function

'Registry Scanning Procedures
Private Sub GetKeys(Path As String)
    If m_Cancel Then
        Exit Sub
    End If
    
    If Not m_ScanDeep = 0 Then
        If KeyDeep(Path) - m_StartDeep > m_ScanDeep Then
            Exit Sub
        End If
    End If
    
    Dim Ctr As Long
    If m_FilterKeysB Then
        If KeyDeep(Path) - m_StartDeep = 1 Then
            For Ctr = LBound(m_FilterKeys) To UBound(m_FilterKeys)
                'LCase$ for no case-sensitive filter search
                If LCase$(Mid$(Path, InStrRev(Path, "\") + 1)) Like m_FilterKeys(Ctr) Then
                    Exit For
                Else
                    If Ctr = UBound(m_FilterKeys) Then
                        Exit Sub
                    End If
                End If
            Next Ctr
        End If
    End If
    
    Dim p_Key As Long
    Dim p_Ret As Long
    
    p_Ret = RegOpenKeyEx(m_ClassRoot, Path, REG_OPTION_NON_VOLATILE, KEY_QUERY_VALUE + KEY_ENUMERATE_SUB_KEYS, p_Key)
    
    If p_Ret = ERROR_SUCCESS Then
        Dim p_Buffer    As String
        Dim p_FT        As FILETIME
        Dim p_Index     As Long
        Dim p_Keys()    As String
        Dim p_MaxLen    As Long
        Dim p_NullPos   As Long
        
        Ctr = 0&
        p_Ret = RegQueryInfoKey(p_Key, "", 0&, REG_OPTION_RESERVED, Ctr, p_MaxLen, 0&, 0&, 0&, 0&, 0&, p_FT)
        
        If Ctr > 0 Then
            ReDim p_Keys(0 To (Ctr - 1)) As String
            p_MaxLen = p_MaxLen + 1 'for the terminating null character
            
            Do
                DoEvents 'Ensure the system processes events other than this function
                If m_Cancel Then
                    Call RegCloseKey(p_Key)
                    Exit Sub
                End If
                
                p_Buffer = String$(p_MaxLen, 0)
                p_Ret = RegEnumKey(p_Key, p_Index, p_Buffer, p_MaxLen)
                
                If p_Ret = ERROR_SUCCESS Then
                    p_NullPos = InStr(1, p_Buffer, vbNullChar)
                    p_Buffer = Left$(p_Buffer, p_NullPos - 1)
                    
                    If Len(Path) = 0 Then
                        p_Keys(p_Index) = p_Buffer
                    Else
                        p_Keys(p_Index) = Path & "\" & p_Buffer 'Write the whole path
                    End If
                End If
                
                p_Index = p_Index + 1
            Loop While (p_Ret = ERROR_SUCCESS) And (p_Index < Ctr)
        End If
        
        Call RegCloseKey(p_Key)
    End If
    
    If Ctr > 0 Then
        For Ctr = 0 To UBound(p_Keys)
            DoEvents 'Ensure the system can process other events
            If m_Cancel Then
                Exit Sub
            End If
            
            p_Buffer = p_Keys(Ctr)
            
            GetData p_Buffer
            GetKeys p_Buffer
        Next Ctr
    End If
End Sub

Private Sub GetData(Path As String)
    If m_Cancel Then
        Exit Sub
    End If
    
    If Not m_ScanDeep = 0 Then
        If KeyDeep(Path) - m_StartDeep > m_ScanDeep Then
            Exit Sub
        End If
    End If
    
    Dim Ctr As Long
    If m_FilterKeysB Then
        If KeyDeep(Path) - m_StartDeep = 1 Then
            For Ctr = LBound(m_FilterKeys) To UBound(m_FilterKeys)
                'LCase$ for no case-sensitive filter search
                If LCase$(Mid$(Path, InStrRev(Path, "\") + 1)) Like m_FilterKeys(Ctr) Then
                    Exit For
                Else
                    If Ctr = UBound(m_FilterKeys) Then
                        Exit Sub
                    End If
                End If
            Next Ctr
        End If
    End If
    
    Dim p_Delete    As Boolean
    Dim p_Key       As Long
    Dim p_Ret       As Long
    
    p_Delete = False
    m_TotalKeys = m_TotalKeys + 1
    RaiseEvent CurrentKey((Path), (m_ClassRoot), p_Delete)
    
    If p_Delete Then
        If Me.DeleteKey(m_ClassRoot, Path) Then
            Exit Sub 'If successful, we have no need to continue below
        End If
    End If
    
    p_Ret = RegOpenKeyEx(m_ClassRoot, Path, REG_OPTION_NON_VOLATILE, KEY_QUERY_VALUE, p_Key)
    
    If p_Ret = ERROR_SUCCESS Then
        Dim p_Buffer    As String
        Dim p_FT        As FILETIME
        Dim p_Index     As Long
        Dim p_MaxLen    As Long
        Dim p_Size      As Long
        Dim p_Values()  As String
        
        Ctr = 0&
        p_Ret = RegQueryInfoKey(p_Key, "", 0&, REG_OPTION_RESERVED, 0&, 0&, 0&, Ctr, p_MaxLen, 0&, 0&, p_FT)
        
        If Ctr > 0 Then
            ReDim p_Values(0 To (Ctr - 1)) As String
            
            Do
                DoEvents 'Ensure the system can process other events
                If m_Cancel Then
                    Call RegCloseKey(p_Key)
                    Exit Sub
                End If
                
                p_Size = p_MaxLen + 1
                p_Buffer = String$(p_Size, 0)
                p_Ret = RegEnumValue(p_Key, (p_Index), p_Buffer, p_Size, REG_OPTION_RESERVED, 0&, 0&, 0&)
                
                If p_Ret = ERROR_SUCCESS Then
                    p_Values(p_Index) = Left$(p_Buffer, p_Size)
                End If
                
                p_Index = p_Index + 1
            Loop While (p_Ret = ERROR_SUCCESS) And (p_Index < Ctr)
        End If
        
        Call RegCloseKey(p_Key)
    End If
    
    If Ctr > 0 Then
        For Ctr = 0 To UBound(p_Values)
            DoEvents 'Ensure the system can process other events
            If m_Cancel Then
                Exit Sub
            End If
            
            If m_FilterDataB Then
                For p_Index = LBound(m_FilterData) To UBound(m_FilterData)
                    'LCase$ for no case-sensitive filter search
                    If LCase$(p_Values(Ctr)) Like m_FilterData(p_Index) Then
                        Exit For
                    Else
                        If p_Index = UBound(m_FilterData) Then
                            GoTo Nxt_Data
                        End If
                    End If
                Next p_Index
            End If
            
            p_Delete = False
            m_TotalData = m_TotalData + 1
            RaiseEvent CurrentData((p_Values(Ctr)), (Path), (m_ClassRoot), p_Delete)
            
            If p_Delete Then
                Call Me.DeleteValue(m_ClassRoot, Path, p_Values(Ctr))
            End If
Nxt_Data:
        Next Ctr
    End If
End Sub

'Class Procedures
Private Sub Class_Initialize()
    'Unless these values are set, these would be their default values
    m_ClassRoot = SCAN_CURRENT_USER
    m_ScanDeep = 0&
    m_ScanPath = "Software"
End Sub

'Created by Noel A. Dacara | Copyright © 2003-2005 Davao City, Philippines
