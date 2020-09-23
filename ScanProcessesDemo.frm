VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Scan Processes Demo"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7215
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   7215
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6240
      Top             =   3975
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Refresh List of Processes"
      Default         =   -1  'True
      Height          =   465
      Left            =   225
      TabIndex        =   0
      Top             =   5580
      Width           =   6810
   End
   Begin VB.Frame Frame1 
      Caption         =   "Scan Options:"
      Height          =   555
      Left            =   225
      TabIndex        =   3
      Top             =   4860
      Width           =   4380
      Begin VB.CheckBox Check1 
         Caption         =   "Show System Processes"
         Height          =   240
         Left            =   180
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   225
         Width           =   2040
      End
      Begin VB.Label NUMPROC 
         AutoSize        =   -1  'True
         Caption         =   "00"
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
         Left            =   4005
         TabIndex        =   6
         Top             =   240
         Width           =   210
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Number of Processes:"
         Height          =   195
         Left            =   2340
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Terminate Process"
      Height          =   465
      Left            =   4725
      TabIndex        =   1
      Top             =   4950
      Width           =   2310
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4560
      Left            =   180
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   225
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   8043
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Process Name"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Process Path"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Process ID"
         Object.Width           =   1676
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Watch this area here--^ ...it contains all the form objects that supports events

Dim WithEvents SCANPROC As cScanProcesses
Attribute SCANPROC.VB_VarHelpID = -1
'Take note of the declaration above!!!

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Dim m_Time As Single
Dim FILEICON As cGetFileIcon

'##################################################################################

Private Sub SCANPROC_CurrentProcess(File As String, Path As String, ID As Long, Terminate As Boolean)
'   Tips: You can perform checksum checks here indiviually for each file...
    Dim p_HasImage As Boolean
    If Path <> "SYSTEM" Then
        p_HasImage = True
        On Error Resume Next
         ImageList1.ListImages(File).Tag = "" 'Just to test if this item exists
         If Not Err.Number = 0 Then
            ImageList1.ListImages.Add , File, FILEICON.Icon(Path & File, SmallIcon)
            Err.Clear
        End If
    End If
    
    Dim lsv As ListItem
    If p_HasImage Then
        Set lsv = ListView1.ListItems.Add(, , File, , File)
    Else
        Set lsv = ListView1.ListItems.Add(, , File)
    End If
    lsv.SubItems(1) = Path
    lsv.SubItems(2) = ID
    lsv.ListSubItems(2).ForeColor = vbBlue
    lsv.Selected = True
    lsv.EnsureVisible
End Sub

Private Sub SCANPROC_DoneScanning(TotalProcesses As Integer)
    Dim p_Elapsed As Single
    p_Elapsed = Timer - m_Time
    'MsgBox "Total Number of Process Detected: " & TotalProcesses & vbNewLine & vbNewLine & "Total Scan Time: " & p_Elapsed, vbInformation, "Done Scanning"
    Debug.Print "Total Number of Process Detected: " & TotalProcesses & vbNewLine & "Total Scan Time: " & p_Elapsed & vbNewLine
    NUMPROC = TotalProcesses
End Sub

'##################################################################################

Private Sub Check1_Click()
    ListView1.ListItems.Clear
    
    SCANPROC.SystemProcesses = (Check1.Value = vbChecked)
    m_Time = Timer
    SCANPROC.BeginScanning
End Sub

Private Sub Command1_Click()
    If MsgBox("Are you sure to terminate this process?", vbExclamation + vbYesNoCancel + vbDefaultButton2, "Terminate Process") = vbYes Then
        If SCANPROC.TerminateProcess(ListView1.SelectedItem.SubItems(2)) = True Then
            Call Check1_Click
        End If
    End If
End Sub

Private Sub Command2_Click()
    Call Check1_Click
End Sub

Private Sub Form_Initialize()
    Set SCANPROC = New cScanProcesses
    Set FILEICON = New cGetFileIcon
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        SCANPROC.CancelScanning
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Call Check1_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SCANPROC.CancelScanning
    
    Set SCANPROC = Nothing
    Set FILEICON = Nothing
    
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

