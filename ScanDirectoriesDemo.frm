VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Scan Directories Demo"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6795
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   6795
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5790
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ScanDirectoriesDemo.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3765
      Left            =   180
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1575
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   6641
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      TextBackground  =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Path/Filename"
         Object.Width           =   8820
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "CRC32"
         Object.Width           =   1942
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Scan Options:"
      Height          =   1140
      Left            =   180
      TabIndex        =   7
      Top             =   5445
      Width           =   6405
      Begin VB.ComboBox Combo5 
         Height          =   315
         Left            =   1665
         Style           =   2  'Dropdown List
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   630
         Width           =   4560
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Custom Scan"
         Height          =   195
         Left            =   2610
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   270
         Width           =   1300
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   5760
         Style           =   1  'Simple Combo
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   "Combo2"
         Top             =   225
         Width           =   465
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Include Subdirectories"
         Height          =   195
         Left            =   180
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   270
         Width           =   2000
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Custom Scan Path:"
         Height          =   195
         Left            =   180
         TabIndex        =   17
         Top             =   675
         Width           =   1365
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Scan Deepness:"
         Height          =   195
         Left            =   4500
         TabIndex        =   10
         Top             =   270
         Width           =   1155
      End
   End
   Begin VB.ComboBox Combo3 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   1395
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "Combo3"
      Top             =   1125
      Width           =   5190
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1395
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "Combo2"
      Top             =   675
      Width           =   5190
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Begin Scanning"
      Default         =   -1  'True
      Height          =   510
      Left            =   180
      TabIndex        =   1
      Top             =   6750
      Width           =   6405
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1395
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "Combo1"
      Top             =   225
      Width           =   5190
   End
   Begin VB.Label NUMFOLDERS 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6435
      TabIndex        =   14
      Top             =   7380
      Width           =   105
   End
   Begin VB.Label NUMFILES 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6435
      TabIndex        =   13
      Top             =   7650
      Width           =   105
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Current number of Files Scanned:"
      Height          =   195
      Left            =   225
      TabIndex        =   12
      Top             =   7650
      Width           =   2415
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Current number of Folders Scanned:"
      Height          =   195
      Left            =   225
      TabIndex        =   11
      Top             =   7380
      Width           =   2625
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "File Attributes:"
      Height          =   195
      Left            =   225
      TabIndex        =   6
      Top             =   1170
      Width           =   1065
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Filter Settings:"
      Height          =   195
      Left            =   225
      TabIndex        =   4
      Top             =   720
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Directory Path:"
      Height          =   195
      Left            =   225
      TabIndex        =   2
      Top             =   270
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Watch this area here--^ ...it contains all the form objects that supports events

Dim WithEvents SCANDIR As cScanDirectories
Attribute SCANDIR.VB_VarHelpID = -1
'Take note of the declaration above!!!

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private t As Single
Dim CRC32 As cCRC32

'##################################################################################

Private Sub SCANDIR_CurrentFile(File As String, Path As String, Delete As Boolean)
'   Tips: You can perform search and destroy operations here!
'         You can perform checksum checks and other file validation you want.
    
    ListView1.ListItems.Add(, , File).SubItems(1) = CRC32.FileChecksum(Path & "\" & File)
    NUMFILES = SCANDIR.TotalFiles
    
    ListView1.ListItems(ListView1.ListItems.Count).Selected = True
    ListView1.SelectedItem.EnsureVisible
End Sub

Private Sub SCANDIR_CurrentFolder(Path As String, Cancel As Boolean, Delete As Boolean)
'   Tips: You can perform search and destroy operations here!
'         Delete method here only supports for empty folders and without subfolders.
    
    'If Path = App.Path & "\New Folder" Then Delete = True
    
    ListView1.ListItems.Add(, , Mid$(Path, 4), , 1).SubItems(1) = "Drive " & Left$(Path, 2)
    NUMFOLDERS = SCANDIR.TotalFolders
    
    ListView1.ListItems(ListView1.ListItems.Count).Selected = True
    ListView1.SelectedItem.ForeColor = vbBlue
    ListView1.SelectedItem.ToolTipText = ListView1.SelectedItem.Text
    ListView1.SelectedItem.EnsureVisible
End Sub

Private Sub SCANDIR_DoneScanning(TotalFolders As Long, TotalFiles As Long)
    Command1.Caption = "Begin Scanning"

    MsgBox "Total Folders Scanned: " & TotalFolders & vbNewLine & _
           "Total Files Scanned  : " & TotalFiles & vbNewLine & vbNewLine & _
           "Total Scan Time: " & Timer - t & " seconds.", vbInformation, "Done Scanning"
End Sub

'##################################################################################

'Below are the usual form procedures

Private Sub Check2_Click()
    If Check2.Value = vbChecked Then
        Combo1.Locked = True
        Combo1.BackColor = vbButtonFace
        Combo5.Locked = False
        Combo5.BackColor = vbWindowBackground
    Else
        Combo1.Locked = False
        Combo1.BackColor = vbWindowBackground
        Combo5.Locked = True
        Combo5.BackColor = vbButtonFace
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        SCANDIR.CancelScanning
        Unload Me
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SCANDIR.CancelScanning
    
    Set SCANDIR = Nothing
    Set CRC32 = Nothing
    
    'Provide a link to my code in PSC
    'If you use this on your other submissions, please give me some credits
    Dim URL As String
    URL = Dir$(App.Path & "\@PSC_ReadMe_*.txt")
    
    If Len(URL) > 0 Then
        Dim f As Integer
        f = FreeFile
        
        On Error Resume Next
        Open App.Path & "\" & URL For Input As f
            URL = Input(LOF(f), 1) 'Get contents of file
        Close 1
        
        f = InStrRev(URL, "http://")
        URL = Mid$(URL, f, InStr(f, URL, vbCrLf) - f)
        
        MsgBox "I would like to here from you about my work so that I can improve it in the future." & vbNewLine & "Your comments or any suggestions are good but your votes would be much better.", vbInformation, "PLEASE DONT FORGET TO VOTE"
        ShellExecute Me.hWnd, "open", URL, "", "", vbNormalFocus
    End If
End Sub

Private Sub Command1_Click()
    If Command1.Caption = "Begin Scanning" Then
        ListView1.ListItems.Clear
        
        t = Timer
        SCANDIR.StartPath = Combo1.Text
        SCANDIR.Filter = Combo2.Text
        SCANDIR.SubDirectories = (Check1.Value = vbChecked)
        SCANDIR.ScanDeep = Combo4.Text
        SCANDIR.CustomScan = (Check2.Value = vbChecked)
        SCANDIR.CustomScanPath = Val(Combo5.Text)
        
        Command1.Caption = "Abort Scanning"
        
        SCANDIR.BeginScanning
    Else
        SCANDIR.CancelScanning
        Command1.Caption = "Begin Scanning"
    End If
End Sub

Private Sub Form_Load()
    Set SCANDIR = New cScanDirectories
    Set CRC32 = New cCRC32
    
    'You can also search the current drive simply by leaving the start path empty
    'Or a whole drive by specifying only the drive letter...
    Combo1.Text = "C:\Program Files"
    
    Combo2.AddItem "*.com"
    Combo2.AddItem "*.exe"
    Combo2.AddItem "*.dll"
    Combo2.AddItem "*.ocx"
    Combo2.AddItem "*.scr"
    Combo2.AddItem "*.vbs"
    Combo2.AddItem "*.com|*.exe|*.dll|*.ocx|*.scr|*.vbs"
    Combo2.AddItem "*.*" 'All files with extension
    Combo2.AddItem "*" 'All files (includes files w/o extension)
    
    Combo2.ListIndex = Combo2.ListCount - 3
    
    Combo3.Text = "Default to search for all files (can be changed)"
    Combo4.Text = 0
    
    Combo5.AddItem ECustomSystemPaths.CSIDL_COMMON_DESKTOPDIRECTORY & "    SCAN_ALLDESKTOP"
    Combo5.AddItem ECustomSystemPaths.CSIDL_COMMON_STARTMENU & "    SCAN_ALLSTARTMENU"
    Combo5.AddItem ECustomSystemPaths.CSIDL_COMMON_STARTUP & "    SCAN_ALLSTARTUP"
    Combo5.AddItem ECustomSystemPaths.CSIDL_FONTS & "    SCAN_FONTS"
    Combo5.AddItem ECustomSystemPaths.CSIDL_PROGRAM_FILES & "    SCAN_PROGRAMFILES"
    Combo5.AddItem ECustomSystemPaths.CSIDL_SYSTEM & "    SCAN_SYSTEMDIR"
    Combo5.AddItem ECustomSystemPaths.CSIDL_TEMPORARYFILES & "  SCAN_TEMPFILESDIR"
    Combo5.AddItem ECustomSystemPaths.CSIDL_DESKTOP & "      SCAN_USERDESKTOP"
    Combo5.AddItem ECustomSystemPaths.CSIDL_PERSONAL & "      SCAN_USERDOCUMENTS"
    Combo5.AddItem ECustomSystemPaths.CSIDL_RECENT & "      SCAN_USERRECENTS"
    Combo5.AddItem ECustomSystemPaths.CSIDL_STARTMENU & "    SCAN_USERSTARTMENU"
    Combo5.AddItem ECustomSystemPaths.CSIDL_STARTUP & "      SCAN_USERSTARTUP"
    Combo5.AddItem ECustomSystemPaths.CSIDL_WINDOWS & "    SCAN_WINDOWSDIR"
    
    Combo5.ListIndex = 0
    
    Check1.Value = vbChecked
    Check2.Value = vbChecked
    Check2.Value = vbUnchecked 'To trigger the click event
End Sub
