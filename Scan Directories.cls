VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cScanDirectories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Author         : Noel A. Dacara (noeldacara@yahoo.com)
'Filename       : Scan Directories.cls (cScanDirectories Class Module)
'Description    : Scan directories for specific files and folders
'Date           : Tuesday, August 02, 2005, 09:14 AM
'Last Update    : Thursday, November 24, 2005, 12:41 PM

'You can freely use and distribute this class or upload these codes on any site
'provided that the original credits are kept unmodified.

'Keep note that :
'I only introduce the 13 common special directories for the custom scan paths.
'You can still enter other special directories not included by specifying the
'constant CSIDL of the folder. Ex. CustomScanPath = &H27 <- CSIDL_MYPICTURES..

'If the startpath is empty, it will automatically be set to the current directory.
'Sample Filter Pattern: *.frm|*.bas|class1.cls|project1.prj

'Scanning Tips  :
'To experience the maximum scanning speed of your machine with this module,
'it is advisable that you close other unnecessary programs running in the system.

'API Declarations
Private Declare Function FindClose Lib "kernel32.dll" (ByVal hFindFile As Long) As Long
Private Declare Function FindFirstFile Lib "kernel32.dll" Alias "FindFirstFileA" (ByVal lpFileName As String, ByRef lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32.dll" Alias "FindNextFileA" (ByVal hFindFile As Long, ByRef lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function GetShortPathName Lib "kernel32.dll" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function SHGetFolderPath Lib "ShFolder" Alias "SHGetFolderPathA" (ByVal hWnd As Long, ByVal CSIDL As Long, ByVal TOKENHANDLE As Long, ByVal flags As Long, ByVal lpPath As String) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long

'API Constants
Private Const INVALID_HANDLE_VALUE  As Long = -1
Private Const MAX_PATH              As Long = 260

'API Types
Private Type FILETIME
    dwLowDateTime   As Long
    dwHighDateTime  As Long
End Type

Private Type SHORTITEMID
    cb      As Long
    abID    As Integer
End Type

Private Type ITEMIDLIST
    mkid    As SHORTITEMID
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes    As Long
    ftCreationTime      As FILETIME
    ftLastAccessTime    As FILETIME
    ftLastWriteTime     As FILETIME
    nFileSizeHigh       As Long
    nFileSizeLow        As Long
    dwReserved0         As Long
    dwReserved1         As Long
    cFileName           As String * MAX_PATH
    cAlternate          As String * 14
End Type

'User-Defined API Enum
Enum ECustomSystemPaths
    CSIDL_DESKTOP = &H0                   'Desktop
    CSIDL_PERSONAL = &H5                  'My Documents
    CSIDL_STARTUP = &H7                   'User startup
    CSIDL_RECENT = &H8                    'Recent files
    CSIDL_STARTMENU = &HB                 'Start menu
    CSIDL_FONTS = &H14                    'Fonts
    CSIDL_COMMON_STARTMENU = &H16         'All user start menu
    CSIDL_COMMON_STARTUP = &H18           'All user startup
    CSIDL_COMMON_DESKTOPDIRECTORY = &H19  'All user desktop
    CSIDL_WINDOWS = &H24                  'Windows
    CSIDL_SYSTEM = &H25                   'System
    CSIDL_PROGRAM_FILES = &H26            'Program files
    CSIDL_TEMPORARYFILES = &HFF '???      'User Temporary files
End Enum

#If False Then
    'Trick to preserve casing of these variables when used in VB IDE
    Private CSIDL_DESKTOP, CSIDL_PERSONAL, CSIDL_STARTUP
    Private CSIDL_RECENT, CSIDL_STARTMENU, CSIDL_FONTS, CSIDL_COMMON_STARTMENU
    Private CSIDL_COMMON_STARTUP, CSIDL_COMMON_DESKTOPDIRECTORY, CSIDL_WINDOWS
    Private CSIDL_SYSTEM, CSIDL_PROGRAM_FILES, CSIDL_TEMPORARYFILES
#End If

'Variable Declarations
Private m_Attributes        As VbFileAttribute
Private m_Cancel            As Boolean
Private m_CurrentDirectory  As String
Private m_CustomScan        As Boolean
Private m_CustomScanPath    As ECustomSystemPaths
Private m_Filters()         As String
Private m_ScanDeep          As Integer
Private m_Scanning          As Boolean
Private m_StartDeep         As Integer
Private m_StartPath         As String
Private m_SubDirectories    As Boolean
Private m_TotalFiles        As Long
Private m_TotalFolders      As Long

Event CurrentFile(File As String, Path As String, Delete As Boolean)
Attribute CurrentFile.VB_Description = "Occurs everytime a file is processed."
'Occurs everytime a file is processed.
Event CurrentFolder(Path As String, Cancel As Boolean, Delete As Boolean)
Attribute CurrentFolder.VB_Description = "Occurs when a folder is scanned for files."
'Occurs when a folder is scanned for files.
Event DoneScanning(TotalFolders As Long, TotalFiles As Long)
Attribute DoneScanning.VB_Description = "Occurs when scanning process is done."
'Occurs when scanning process is done.

#If False Then
    'Trick to preserve casing of these variables when used in VB IDE
    Private File, Path, Delete, Cancel, TotalFolders, TotalFiles
#End If

Property Get Attributes() As VbFileAttribute
Attribute Attributes.VB_Description = "Returns/sets the attribute of files to be returned while scanning directories."
'Returns/sets the attribute of files to be returned while scanning directories.
    Attributes = m_Attributes
End Property

Property Let Attributes(Value As VbFileAttribute)
    m_Attributes = Value
End Property

Property Get CurrentDirectory() As String
Attribute CurrentDirectory.VB_Description = "Returns the current directory being processed."
'Returns the current directory being processed.
    CurrentDirectory = m_CurrentDirectory
End Property

Property Get CustomScan() As Boolean
Attribute CustomScan.VB_Description = "Returns/sets whether to scan in the predefined special directories."
'Returns/sets whether to scan in the predefined special directories.
    CustomScan = m_CustomScan
End Property

Property Let CustomScan(Value As Boolean)
    m_CustomScan = Value
End Property

Property Get CustomScanPath() As ECustomSystemPaths
Attribute CustomScanPath.VB_Description = "Returns/sets the custom scan path to scan for files and directory."
'Returns/sets the custom scan path to scan for files and directory.
    CustomScanPath = m_CustomScanPath
End Property

Property Let CustomScanPath(Value As ECustomSystemPaths)
    m_CustomScanPath = Value
End Property

Property Get Filter() As String
Attribute Filter.VB_Description = "Returns/sets the filter patterns used to select files to be processed."
'Returns/sets the filter patterns used to select files to be processed.
    Filter = Join(m_Filters(), Chr(124)) 'Return as one string
End Property

Property Let Filter(Value As String)
    If Len(Trim$(Value)) = 0 Then
        Value = "*" 'Search for all files
    End If
    m_Filters = Split(LCase$(Value), Chr(124)) 'Save as an array of filter patterns
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

Property Get StartPath() As String
Attribute StartPath.VB_Description = "Returns/sets the path where to start scanning for files and folders."
'Returns/sets the path where to start scanning for files and folders.
    StartPath = m_StartPath
End Property

Property Let StartPath(Value As String)
    If Len(Value) = 1 Then  'Maybe a drive letter
        Value = Value & ":" 'So append a colon...
    ElseIf Len(Value) = 0 Then
        Value = CurDir$() 'If empty then set the current directory instead
    End If
    
    m_StartPath = Value
End Property

Property Get SubDirectories() As Boolean
Attribute SubDirectories.VB_Description = "Returns/sets whether to include subdirectories in scanning for files."
'Returns/sets whether to include subdirectories in scanning for files.
    SubDirectories = m_SubDirectories
End Property

Property Let SubDirectories(Value As Boolean)
    m_SubDirectories = Value
End Property

Property Get TotalFiles() As Long
Attribute TotalFiles.VB_Description = "Returns the current number of files processed."
'Returns the current number of files processed.
    TotalFiles = m_TotalFiles
End Property

Property Get TotalFolders() As Long
Attribute TotalFolders.VB_Description = "Returns the current number of folders processed."
'Returns the current number of folders processed.
    TotalFolders = m_TotalFolders
End Property

'Public Procedures
Sub BeginScanning(Optional StartPath As String, Optional Filter As String)
Attribute BeginScanning.VB_Description = "Start scanning for the specified files according to filter pattern."
'Start scanning for the specified files according to filter pattern.
    If Len(Filter) Then
        Me.Filter = Filter
    End If
    
    If Len(StartPath) Then
        Me.StartPath = StartPath
    End If
    
    If m_CustomScan Then
        Dim p_StartPath As String
        p_StartPath = Me.StartPath 'Save startpath
        
        Me.StartPath = GetDirectory(Me.CustomScanPath)
    End If
    
    m_TotalFiles = Empty 'Reset number of files and folders processed
    m_TotalFolders = Empty
    
    m_Cancel = False
    If KeyExists(Me.StartPath) Then 'Only scan if path exists
        Call StartScanning(Me.StartPath)
    End If
    
    If m_CustomScan Then
        Me.StartPath = p_StartPath 'Restore startpath
    End If
    m_Cancel = False 'Ensure we can still successully scan on future events
End Sub

Sub CancelScanning()
Attribute CancelScanning.VB_Description = "Abort the whole scanning process up to the current folder processed."
'Abort the whole scanning process up to the current folder processed.
    m_Cancel = True
End Sub

'Private Procedures
Private Function GetDirectory(CustomScanPath As ECustomSystemPaths) As String
    Dim p_Buffer    As String * MAX_PATH
    Dim t_IDList    As ITEMIDLIST
    Dim p_NullPos   As Integer
    Dim p_Result    As Long
    
    p_Result = SHGetSpecialFolderLocation(0&, CustomScanPath, t_IDList)
    
    If Not p_Result = 0 Then
        p_Result = SHGetFolderPath(0&, CustomScanPath, 0&, 0&, p_Buffer)
        
        If p_Result = 0 Then
            p_Result = SHGetPathFromIDList(ByVal t_IDList.mkid.cb, ByVal p_Buffer)
        End If
    Else
        p_Result = SHGetPathFromIDList(ByVal t_IDList.mkid.cb, ByVal p_Buffer)
    End If
    
    p_NullPos = InStr(1, p_Buffer, vbNullChar)
    If p_NullPos > 0 Then
        GetDirectory = Left$(p_Buffer, p_NullPos - 1)
    End If
    
    If Len(GetDirectory) = 0 Then
        p_Buffer = String$(MAX_PATH, 0) 'Reinitialize string buffer
        
        Select Case CustomScanPath
            Case CSIDL_SYSTEM
                Call GetSystemDirectory(p_Buffer, MAX_PATH)
            Case CSIDL_TEMPORARYFILES
                Call GetTempPath(MAX_PATH, p_Buffer)
            Case CSIDL_WINDOWS, CSIDL_PROGRAM_FILES
                Call GetWindowsDirectory(p_Buffer, MAX_PATH)
                GetDirectory = Trim$(Environ$("ProgramFiles")) 'Not sure if everyone has this
        End Select
        
        If CustomScanPath = CSIDL_PROGRAM_FILES Then
            If Len(GetDirectory) = 0 Then
                GetDirectory = Left$(p_Buffer, 3) & "Program Files"
            End If
        Else
            GetDirectory = Left$(p_Buffer, InStr(1, p_Buffer, vbNullChar) - 1)
            If GetDirectory Like "*\" Then
                GetDirectory = Left$(GetDirectory, Len(GetDirectory) - 1)
            End If
        End If
    End If
End Function

Private Function GetShortPath(LongPath As String) As String
    Dim p_Buffer As String * MAX_PATH
    Dim p_Ret As Long
    
    p_Ret = GetShortPathName(LongPath, p_Buffer, MAX_PATH)
    
    GetShortPath = Left$(p_Buffer, p_Ret)
End Function

Private Function KeyDeep(Key As String) As Integer
    Dim i As Integer
    
    If Len(Key) Then
        KeyDeep = 1
    End If
    
    Do
        i = InStr(i + 1, Key, "\")
        If Not i = 0 Then
            KeyDeep = KeyDeep + 1 'Count number of separators
        End If
    Loop While Not i = 0
End Function

Private Function KeyExists(ByVal Path As String) As Boolean
    If Not (Path Like "*\") Then
        Path = Path & "\"
    End If
    Path = Path & "nul"
    
    If Not Len(Path) = 0 Then
        If Not Len(Dir$(Path)) = 0 Then
            KeyExists = True 'Path/folder exists...
        End If
    End If
End Function

Private Sub StartScanning(Path As String)
    If Attributes And vbDirectory Then
        Attributes = Attributes - vbDirectory 'Valid file pattern only
    End If
    
    m_StartDeep = KeyDeep(Path)
    m_Scanning = True
    Call ScanForData(Path) 'Start scanning
    m_Scanning = False
    
    RaiseEvent DoneScanning((m_TotalFolders), (m_TotalFiles)) 'Done scanning on specified folder
End Sub

Private Sub ScanForData(Path As String)
    If m_Cancel Then
        Exit Sub 'Attempt to cancel the whole scanning process
    End If
    
    If Not m_ScanDeep = 0 Then
        If KeyDeep(Path) - m_StartDeep > m_ScanDeep Then
            Exit Sub
        End If
    End If
    
    Dim p_Delete    As Boolean
    
    m_TotalFolders = m_TotalFolders + 1
    m_CurrentDirectory = Path
    RaiseEvent CurrentFolder((Path), m_Cancel, p_Delete)
    
    If p_Delete Then
        On Local Error Resume Next
        
        RmDir Path 'Attempt to remove directory
        If Err.Number = 0 Then
            Exit Sub
        End If
        
        Err.Clear
        On Error GoTo 0 'Cancel error handler
    End If
                        
    Dim i           As Long
    Dim p_Data      As String
    Dim p_Find      As Long
    Dim p_FindData  As WIN32_FIND_DATA
    
    If Len(Path) - 2 > MAX_PATH Then
        p_Find = FindFirstFile(GetShortPath(Path & "\*"), p_FindData)
    Else
        p_Find = FindFirstFile(Path & "\*", p_FindData)
    End If
    
    If Not p_Find = INVALID_HANDLE_VALUE Then
        Do
            DoEvents
            If m_Cancel Then
                FindClose p_Find 'Close search handle before exit
                Exit Sub 'Cancel the whole scanning process
            End If
            
            p_Data = p_FindData.cFileName
            p_Data = Left$(p_Data, InStr(1, p_Data, vbNullChar) - 1)
            
            p_Delete = False
            If Not (p_Data = "." Or p_Data = "..") Then
                If p_FindData.dwFileAttributes And vbDirectory Then
                    If m_SubDirectories Then
                        ScanForData Path & "\" & p_Data
                    End If
                ElseIf p_FindData.dwFileAttributes Or Attributes Then
                    For i = 0 To UBound(m_Filters) 'Loop through all filter patterns
                        If LCase$(p_Data) Like m_Filters(i) Then
                            m_TotalFiles = m_TotalFiles + 1
                            RaiseEvent CurrentFile((p_Data), (Path), p_Delete) 'Send file to event
                            
                            If p_Delete Then
                                On Local Error Resume Next
                                
                                p_Delete = False
                                Kill Path & "\" & p_Data
                                
                                On Error GoTo 0 'Cancel error handler
                            End If
                        End If
                    Next i
                End If
            End If
        Loop While FindNextFile(p_Find, p_FindData)
        
        FindClose p_Find 'Close search handle
    End If
End Sub

'Class Procedures
Private Sub Class_Initialize()
    'Unless these values are set, these would be their default values
    m_Attributes = vbNormal Or vbHidden Or vbReadOnly Or vbArchive Or vbSystem
    Filter = "*"
    m_StartPath = CurDir$()
    m_SubDirectories = True
End Sub

'Created by Noel A. Dacara | Copyright © 2003-2005 Davao City, Philippines

