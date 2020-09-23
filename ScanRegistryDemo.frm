VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Scan Registry Demo"
   ClientHeight    =   8505
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5790
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   5790
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4920
      Top             =   5145
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
            Picture         =   "ScanRegistryDemo.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4350
      Left            =   180
      TabIndex        =   19
      Top             =   1530
      Width           =   5460
      _ExtentX        =   9631
      _ExtentY        =   7673
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   5469
      EndProperty
   End
   Begin VB.ComboBox Combo6 
      Height          =   315
      Left            =   4035
      Style           =   1  'Simple Combo
      TabIndex        =   4
      Text            =   "Combo6"
      Top             =   1050
      Width           =   1600
   End
   Begin VB.ComboBox Combo5 
      Height          =   315
      Left            =   1305
      Style           =   1  'Simple Combo
      TabIndex        =   3
      Text            =   "Combo5"
      Top             =   1050
      Width           =   1600
   End
   Begin VB.Frame Frame1 
      Caption         =   "Scan Options"
      Height          =   1185
      Left            =   180
      TabIndex        =   10
      Top             =   5985
      Width           =   5460
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   1575
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   675
         Width           =   3705
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Custom Scanning"
         Height          =   195
         Left            =   1890
         TabIndex        =   6
         Top             =   270
         Width           =   1590
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   4860
         Style           =   1  'Simple Combo
         TabIndex        =   7
         Text            =   "Combo3"
         Top             =   225
         Width           =   420
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Include Subkeys"
         Height          =   195
         Left            =   135
         TabIndex        =   5
         Top             =   270
         Width           =   1500
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Custom Scan Path:"
         Height          =   195
         Left            =   135
         TabIndex        =   16
         Top             =   720
         Width           =   1365
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Scan Deepness:"
         Height          =   195
         Left            =   3645
         TabIndex        =   11
         Top             =   270
         Width           =   1155
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Begin Scanning"
      Default         =   -1  'True
      Height          =   420
      Left            =   180
      TabIndex        =   0
      Top             =   7335
      Width           =   5460
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1305
      Style           =   1  'Simple Combo
      TabIndex        =   2
      Text            =   "Combo2"
      Top             =   630
      Width           =   4335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   180
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   225
      Width           =   5460
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Filter Keys:"
      Height          =   195
      Left            =   3120
      TabIndex        =   18
      Top             =   1095
      Width           =   810
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Filter Data:"
      Height          =   195
      Left            =   210
      TabIndex        =   17
      Top             =   1095
      Width           =   810
   End
   Begin VB.Label NUMDATA 
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
      Left            =   5520
      TabIndex        =   15
      Top             =   8190
      Width           =   105
   End
   Begin VB.Label NUMKEYS 
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
      Left            =   5520
      TabIndex        =   14
      Top             =   7920
      Width           =   105
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Current number of Data Scanned:"
      Height          =   195
      Left            =   180
      TabIndex        =   13
      Top             =   8190
      Width           =   2445
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Current number of Keys Scanned:"
      Height          =   195
      Left            =   180
      TabIndex        =   12
      Top             =   7920
      Width           =   2445
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Registry Path:"
      Height          =   195
      Left            =   225
      TabIndex        =   9
      Top             =   675
      Width           =   1035
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Watch this area here--^ ...it contains all the form objects that supports events

Dim WithEvents SCANREG As cScanRegistry
Attribute SCANREG.VB_VarHelpID = -1
'Take note of the declaration above!!!

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private t As Single
Dim REG As cAdvanceRegistry

'##################################################################################

Private Sub SCANREG_CurrentData(Value As String, Key As String, Root As EScanRegistryRoots, Delete As Boolean)
'   Tips: You can perform search and destroy operations here.

    'Debug.Print Key & "\" & Value
    If Len(Value) = 0 Then
        'Default Key Data doesn't have value names (Empty)
        'so let's make it as (Default) here for clarification...
        ListView1.ListItems.Add(, , "(Default)").SubItems(1) = REG.ValueEx(Root, Key, Value)
    Else
        ListView1.ListItems.Add(, , Value).SubItems(1) = REG.ValueEx(Root, Key, Value)
    End If
    
    NUMDATA = SCANREG.TotalData
    
    On Error Resume Next
    ListView1.ListItems(ListView1.ListItems.Count).Selected = True
    ListView1.SelectedItem.ListSubItems(1).ToolTipText = ListView1.SelectedItem.SubItems(1)
    ListView1.SelectedItem.EnsureVisible
End Sub

Private Sub SCANREG_CurrentKey(Key As String, Root As EScanRegistryRoots, Delete As Boolean)
'   Tips: You can perform search and destroy operations here.
    Dim r As String
    Select Case Root
        Case SCAN_CLASSES_ROOT
            r = "CLASSES_ROOT"
        Case SCAN_CURRENT_USER
            r = "CURRENT_USER"
        Case SCAN_LOCAL_MACHINE
            r = "LOCAL_MACHINE"
        Case SCAN_USERS
            r = "USERS"
    End Select
    
    ListView1.ListItems.Add(, , r, , 1).SubItems(1) = Key
    
    NUMKEYS = SCANREG.TotalKeys
    
    On Error Resume Next
    ListView1.ListItems(ListView1.ListItems.Count).Selected = True
    ListView1.SelectedItem.ForeColor = vbBlue
    ListView1.SelectedItem.ListSubItems(1).ForeColor = vbRed
    ListView1.SelectedItem.ListSubItems(1).ToolTipText = ListView1.SelectedItem.SubItems(1)
    ListView1.SelectedItem.EnsureVisible
End Sub

Private Sub SCANREG_DoneScanning(TotalData As Long, TotalKeys As Long)
    Command1.Caption = "Begin Scanning"
    
    MsgBox "Total Keys Scanned: " & TotalKeys & vbNewLine & _
           "Total Data Scanned: " & TotalData & vbNewLine & vbNewLine & _
           "Total Scan Time: " & Timer - t & " seconds.", vbInformation, "Done Scanning"
End Sub

'##################################################################################

Private Sub Check2_Click()
    If Check2.Value = vbChecked Then
        Combo1.Locked = True
        Combo1.BackColor = vbButtonFace
        Combo2.Locked = True
        Combo2.BackColor = vbButtonFace
        Combo3.Locked = True
        Combo3.BackColor = vbButtonFace
        Combo4.Locked = False
        Combo4.BackColor = vbWindowBackground
        Check1.Enabled = False
    Else
        Combo1.Locked = False
        Combo1.BackColor = vbWindowBackground
        Combo2.Locked = False
        Combo2.BackColor = vbWindowBackground
        Combo3.Locked = False
        Combo3.BackColor = vbWindowBackground
        Combo4.Locked = True
        Combo4.BackColor = vbButtonFace
        Check1.Enabled = True
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        SCANREG.CancelScanning
        Unload Me
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SCANREG.CancelScanning
    
    Set SCANREG = Nothing
    Set REG = Nothing
    
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
        ShellExecute Me.hwnd, "open", URL, "", "", vbNormalFocus
    End If
End Sub

Private Sub Command1_Click()
    If Command1.Caption = "Begin Scanning" Then
        ListView1.ListItems.Clear
        t = Timer
    
        Select Case Combo1.ListIndex
            Case 0: SCANREG.ClassRoot = SCAN_CLASSES_ROOT
            Case 1: SCANREG.ClassRoot = SCAN_CURRENT_USER
            Case 2: SCANREG.ClassRoot = SCAN_LOCAL_MACHINE
        End Select
        
        SCANREG.FilterData = Combo5.Text
        SCANREG.FilterKeys = Combo6.Text
        SCANREG.ScanPath = Combo2.Text
        SCANREG.ScanSubKeys = (Check1.Value = vbChecked)
        SCANREG.ScanDeep = Combo3.Text
        SCANREG.CustomScanPath = Combo4.ListIndex
        
        Command1.Caption = "Abort Scanning"
        
        If Check2.Value = vbChecked Then
            SCANREG.BeginCustomScan
        Else
            SCANREG.BeginScanning
        End If
    Else
        Command1.Caption = "Begin Scanning"
        SCANREG.CancelScanning
    End If
End Sub

Private Sub Form_Load()
    Set SCANREG = New cScanRegistry
    Set REG = New cAdvanceRegistry
    
    Combo1.AddItem "SCAN_CLASSES_ROOT", 0
    Combo1.AddItem "SCAN_CURRENT_USER", 1
    Combo1.AddItem "SCAN_LOCAL_MACHINE", 2

    Combo1.ListIndex = 1 'Select SCAN_CURRENT_USER on load

    Combo2.Text = "Software\Microsoft"
    Combo3.Text = 2 'Scan through all subdirectories
    
    Combo4.AddItem "SCAN_ADDREMOVELISTS", 0
    Combo4.AddItem "SCAN_CUSTOMCONTROLS", 1
    Combo4.AddItem "SCAN_FILEEXTENSIONS", 2
    Combo4.AddItem "SCAN_HELPRESOURCES", 3
    Combo4.AddItem "SCAN_SHAREDDLLS", 4
    Combo4.AddItem "SCAN_SHELLFOLDERS", 5
    Combo4.AddItem "SCAN_SOFTWAREPATHS", 6
    Combo4.AddItem "SCAN_STARTUPKEYS", 7
    Combo4.AddItem "SCAN_WINDOWSFONTS", 8
    
    Combo4.ListIndex = 0
    
    Combo5.Text = Empty
    Combo6.Text = Empty
    
    Check1.Value = vbChecked
    Check2.Value = vbChecked
End Sub
