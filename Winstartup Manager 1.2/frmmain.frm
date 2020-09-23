VERSION 5.00
Begin VB.Form frmmain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Winstartup Manager 1.2"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   6135
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   6135
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4185
      Left            =   120
      TabIndex        =   9
      Top             =   930
      Width           =   5880
      Begin VB.ListBox lstName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   2175
         ItemData        =   "frmmain.frx":27A2
         Left            =   3990
         List            =   "frmmain.frx":27A4
         TabIndex        =   12
         ToolTipText     =   "You can always press the DEL key on your keyboard to delete the selected item or [ CTRL + C  ] to copy the selected item ."
         Top             =   345
         Width           =   1755
      End
      Begin VB.ListBox lstCmdLine 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   1005
         ItemData        =   "frmmain.frx":27A6
         Left            =   105
         List            =   "frmmain.frx":27A8
         TabIndex        =   11
         ToolTipText     =   "Command line and executable path"
         Top             =   2790
         Width           =   5655
      End
      Begin VB.PictureBox mainpic 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   2175
         Left            =   105
         ScaleHeight     =   2145
         ScaleWidth      =   3825
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   345
         Width           =   3855
         Begin VB.OptionButton chk1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "HKEY_CURRENT_USER    Run"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   -195
            MaskColor       =   &H00FFFFFF&
            MouseIcon       =   "frmmain.frx":27AA
            MousePointer    =   99  'Custom
            TabIndex        =   1
            Top             =   90
            Width           =   4100
         End
         Begin VB.OptionButton chk2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "HKEY_CURRENT_USER    Run Once"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   -195
            MaskColor       =   &H00FFFFFF&
            MouseIcon       =   "frmmain.frx":28FC
            MousePointer    =   99  'Custom
            TabIndex        =   2
            Top             =   345
            Width           =   4100
         End
         Begin VB.OptionButton chk3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "HKEY_LOCAL_MACHINE  Run"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   -195
            MaskColor       =   &H00FFFFFF&
            MouseIcon       =   "frmmain.frx":2A4E
            MousePointer    =   99  'Custom
            TabIndex        =   3
            Top             =   600
            Width           =   4100
         End
         Begin VB.OptionButton chk4 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "HKEY_LOCAL_MACHINE  Run Once"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   -195
            MaskColor       =   &H00FFFFFF&
            MouseIcon       =   "frmmain.frx":2BA0
            MousePointer    =   99  'Custom
            TabIndex        =   4
            Top             =   855
            Width           =   4100
         End
         Begin VB.OptionButton chk5 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "HKEY_LOCAL_MACHINE  Run Services"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   -195
            MaskColor       =   &H00FFFFFF&
            MouseIcon       =   "frmmain.frx":2CF2
            MousePointer    =   99  'Custom
            TabIndex        =   5
            Top             =   1095
            Width           =   4100
         End
         Begin VB.OptionButton chk6 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "All User Startup"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   -195
            MaskColor       =   &H00FFFFFF&
            MouseIcon       =   "frmmain.frx":2E44
            MousePointer    =   99  'Custom
            TabIndex        =   6
            Top             =   1365
            Width           =   4100
         End
         Begin VB.OptionButton chk7 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "Win.ini (Manual Edit)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   -195
            MaskColor       =   &H00FFFFFF&
            MouseIcon       =   "frmmain.frx":2F96
            MousePointer    =   99  'Custom
            TabIndex        =   8
            Top             =   1875
            Width           =   4100
         End
         Begin VB.OptionButton chk8 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "System.ini (Manual Edit)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   -195
            MaskColor       =   &H00FFFFFF&
            MouseIcon       =   "frmmain.frx":30E8
            MousePointer    =   99  'Custom
            TabIndex        =   7
            Top             =   1620
            Width           =   4100
         End
      End
      Begin VB.FileListBox filelistbox 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   2175
         Left            =   3990
         System          =   -1  'True
         TabIndex        =   13
         ToolTipText     =   "You can always press the DEL key on your keyboard to delete the selected item or [ CTRL + C  ] to copy the selected item ."
         Top             =   345
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Registry Run Section Key Names"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   135
         TabIndex        =   17
         Top             =   120
         Width           =   3765
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Executable Filename"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   4020
         TabIndex        =   16
         Top             =   120
         Width           =   1680
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Executable Path  [Start In]"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   120
         TabIndex        =   15
         Top             =   2550
         Width           =   5580
      End
      Begin VB.Label lblinf 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   14
         Top             =   3855
         Width           =   5640
      End
   End
   Begin VB.Label lbldate 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   5190
      Width           =   5880
   End
   Begin VB.Image imgregg 
      Height          =   435
      Left            =   2865
      Picture         =   "frmmain.frx":323A
      Stretch         =   -1  'True
      Top             =   315
      Width           =   525
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   810
      Left            =   120
      Picture         =   "frmmain.frx":3544
      Stretch         =   -1  'True
      Top             =   90
      Width           =   5880
   End
   Begin VB.Menu mnufle 
      Caption         =   "&File"
      Begin VB.Menu mnuCopyToClipboard 
         Caption         =   "Copy To Clipboard"
      End
      Begin VB.Menu mnuDeleteEntry 
         Caption         =   "Delete Selected Item"
      End
      Begin VB.Menu sep 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCreateBackup 
         Caption         =   "Create Backup"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRestoreBackup 
         Caption         =   "Restore Backup"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuview 
      Caption         =   "&View"
      Begin VB.Menu mnualwaysontop 
         Caption         =   "Always on top"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuAlwayscreatebackuponexit 
         Caption         =   "Always Create Backup On Exit"
      End
      Begin VB.Menu mnuRestoreaprevoussavedbackup 
         Caption         =   "Restore Previous Saved Backup"
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
      Begin VB.Menu mnucontact 
         Caption         =   "Contact"
      End
      Begin VB.Menu mnufeedback 
         Caption         =   "Feedback"
      End
      Begin VB.Menu mnuvisithomepage 
         Caption         =   "Visit homepage"
      End
      Begin VB.Menu mnusepa 
         Caption         =   "-"
      End
      Begin VB.Menu mnuabout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long

Dim reg As CRegistry
Dim env As CEnvironment

Dim hKey As Long, LCount As Long, i As Long

Private Sub chk1_Click()

On Error Resume Next

If chk1.Value = True Then
On Error Resume Next

lblinf.Caption = chk1.Caption
lstName.Clear
lstCmdLine.Clear
filelistbox.FileName = vbNullString
filelistbox.Visible = False
lstName.Visible = True
EnumRegRun1
End If

End Sub

Private Sub chk1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
lstName.SetFocus
End If

End Sub

Private Sub chk2_Click()

On Error Resume Next

If chk2.Value = True Then

On Error Resume Next
lblinf.Caption = chk2.Caption
lstName.Clear
lstCmdLine.Clear
filelistbox.FileName = vbNullString
filelistbox.Visible = False
lstName.Visible = True
EnumRegRun2
End If

End Sub

Private Sub chk2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
lstName.SetFocus
End If
End Sub

Private Sub chk3_Click()

On Error Resume Next

If chk3.Value = True Then

On Error Resume Next
lblinf.Caption = chk3.Caption
lstName.Clear
lstCmdLine.Clear
filelistbox.FileName = vbNullString
filelistbox.Visible = False
lstName.Visible = True
EnumRegRun3
End If

End Sub

Private Sub chk3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
lstName.SetFocus
End If
End Sub

Private Sub chk4_Click()

On Error Resume Next


If chk4.Value = True Then

On Error Resume Next
lblinf.Caption = chk4.Caption
lstName.Clear
lstCmdLine.Clear
filelistbox.FileName = vbNullString
filelistbox.Visible = False
lstName.Visible = True
EnumRegRun4
End If

End Sub

Private Sub chk4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
lstName.SetFocus
End If
End Sub

Private Sub chk5_Click()

On Error Resume Next

If chk5.Value = True Then

On Error Resume Next
lblinf.Caption = chk5.Caption
lstName.Clear
lstCmdLine.Clear
filelistbox.FileName = vbNullString
filelistbox.Visible = False
lstName.Visible = True
EnumRegRun5
End If

End Sub


Private Sub chk5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
lstName.SetFocus
End If
End Sub

Private Sub chk6_Click()

On Error Resume Next


Dim selfile As Variant
selfile = filelistbox.Selected(0)

lstName.Clear
lstCmdLine.Clear
lblinf.Caption = chk6.Caption
lstName.Visible = False

filelistbox.Visible = True
filelistbox.FileName = CheckFolderID(Common_StartUp)
lstCmdLine.AddItem CheckFolderID(Common_StartUp)

filelistbox.Selected(selfile) = True

End Sub

Private Sub chk6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
filelistbox.SetFocus
End If
End Sub

Private Sub chk7_Click()

On Error Resume Next
lblinf.Caption = chk7.Caption
lstName.Clear
lstCmdLine.Clear
lstName.Visible = True
MousePointer = 11
ShellExecute 0, "open", "notepad.exe", env.WindowsDirectory & "\win.ini", "", 1
MousePointer = 0

End Sub

Private Sub chk8_Click()

On Error Resume Next
frmmain.filelistbox.Visible = False
lblinf.Caption = chk8.Caption
lstName.Clear
lstCmdLine.Clear
lstName.Visible = True
MousePointer = 11
ShellExecute 0, "open", "notepad.exe", env.WindowsDirectory & "\system.ini", "", 1
MousePointer = 0

End Sub

Private Sub filelistbox_Click()
Dim tmp
tmp = filelistbox.FileName
End Sub

Private Sub filelistbox_KeyDown(KeyCode As Integer, Shift As Integer)

On Error Resume Next

Dim tmp

tmp = filelistbox.FileName
If KeyCode = vbKeyDelete Then

mnuDeletestartupfiles

Else

If KeyCode = "{CTRL + C}" Then
Clipboard.Clear
Clipboard.SetText (tmp), 1

Else

If filelistbox.ListCount = 0 Then
Exit Sub

End If
End If
End If
End Sub

Private Sub filelistbox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If filelistbox.ListCount = 0 Then
Exit Sub
Else

If Button = 2 Then
PopupMenu mnufle
End If
End If

End Sub

Private Sub Form_Initialize()
On Error Resume Next

Dim X As Variant
X = InitCommonControls
End Sub

Private Sub Form_Load()
On Error Resume Next
Me.Height = 6300
Me.Width = 6210

checkinstances
readsetting
lbldate.Caption = " Last backup " & GetSetting(App.EXEName, "settings", "BackupDate", 0)

Set reg = New CRegistry
Set env = New CEnvironment

lblinf.Caption = chk1.Caption
mainpic.SetFocus
firstrun

End Sub
Sub readsetting()
Dim checkback
checkback = GetSetting(App.EXEName, "settings", "AlwaysBackup", 0)
mnuAlwayscreatebackuponexit.Checked = checkback

End Sub

Sub checkinstances()
On Error Resume Next

If App.PrevInstance = True Then
'MsgBox Me.Caption & " Is already running", vbInformation, "No more instances"

End
  
End If
End Sub
Sub firstrun()

Dim f As Variant
f = GetSetting(App.EXEName, "settings", "FirstRun", 0)
If f = "Done" Then
Exit Sub

Else
Me.Hide

frmwait.Show

End If

End Sub

Private Sub lstCmdLine_Click()
On Error Resume Next
lstName.ListIndex = lstCmdLine.ListIndex
End Sub

Private Sub lstCmdLine_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyDelete Then
mnuDeleteEntry_Click

End If
End Sub

Private Sub lstCmdLine_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = vbRightButton And lstCmdLine.ListCount > 0 Then

End If
End Sub

Private Sub lstName_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next

Dim tmp1
tmp1 = lstName.Text

If KeyCode = vbKeyDelete Then
mnuDeleteEntry_Click

Else

If KeyCode = "{CTRL + C}" Then
Clipboard.Clear
Clipboard.SetText (tmp1), 1

End If

End If

End Sub

Private Sub lstName_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If lstName.ListCount = 0 Then
Exit Sub
Else
If Button = 2 Then
PopupMenu mnufle
End If
End If

End Sub

Private Sub mnuAlwayscreatebackuponexit_Click()
If mnuAlwayscreatebackuponexit.Checked = True Then
mnuAlwayscreatebackuponexit.Checked = False

SaveSetting App.EXEName, "Settings", "AlwaysBackup", "False"
Else
SaveSetting App.EXEName, "Settings", "AlwaysBackup", "True"
mnuAlwayscreatebackuponexit.Checked = True

'Unload Me
End If

End Sub

Private Sub mnucontact_Click()
On Error Resume Next
ShellExecute hwnd, "open", "mailto:kegham_d@hotmail.com", vbNullString, vbNullString, 1

End Sub

Private Sub mnuCopyToClipboard_Click()
On Error Resume Next
Clipboard.Clear
Clipboard.SetText lstName.Text

End Sub

Private Sub mnuCreateBackup_Click()
On Error Resume Next

lstCmdLine.Clear
lstName.Clear
txtcmdline.Text = vbNullString
lblinf.Caption = "Creating backup"
createbackup

End Sub
Private Sub Form_Unload(Cancel As Integer)

On Error Resume Next
Dim backupopt As Variant

backupopt = GetSetting(App.EXEName, "settings", "AlwaysBackup", 0)
If backupopt = "True" Then

createbackup

Set env = Nothing ' Clear everything and end here!
Set reg = Nothing

End If
Exit Sub

Unload Me
End

End Sub

Private Sub lstName_Click()
On Error Resume Next

txtcmdline.Text = lstName.Text
lstCmdLine.ListIndex = lstName.ListIndex

End Sub

Private Sub mnuabout_Click()
On Error Resume Next
frmAbout.Show
End Sub

Private Sub mnualwaysontop_Click()
On Error Resume Next

Dim RetVal As Variant

If mnualwaysontop.Checked = False Then
mnualwaysontop.Checked = True
RetVal = SetWindowPos(frmmain.hwnd, -1, 0, 0, 0, 0, 3)

Else

If mnualwaysontop.Checked = True Then
mnualwaysontop.Checked = False
RetVal = SetWindowPos(frmmain.hwnd, -2, 0, 0, 0, 0, 3)

Else
End If
End If

End Sub

Private Sub mnuDeleteEntry_Click()
On Error Resume Next

Dim sel As Integer
Dim selregentry As Variant
Dim check As OptionButton

sel = lstCmdLine.ListIndex
selregentry = check

If selregentry = check6 And lstName.Visible = False Then
mnuDeletestartupfiles
chk6_Click

Else


If lstName.ListCount > 0 Then
Dim ask As String
ask = MsgBox("You are about to remove this item from execution." & vbCrLf & "Item Name: " & lstName.Text, vbYesNo, "Please confirm removing it if you sure")
If ask = vbYes Then

'Begin checking which option button clicked

If selregentry = check1 Then
ModDel.DeleteValue ERegRoot.eRegRoot_HKeyCurrentUser, "Software\Microsoft\Windows\CurrentVersion\Run", lstName.Text
lstCmdLine.RemoveItem sel
lstName.RemoveItem sel
lstCmdLine.Clear
lstName.Clear
txtcmdline.Text = vbNullString
chk1_Click


Else

If selregentry = check2 Then
ModDel.DeleteValue ERegRoot.eRegRoot_HKeyCurrentUser, "Software\Microsoft\Windows\CurrentVersion\RunOnce", lstName.Text
lstCmdLine.RemoveItem sel
lstName.RemoveItem sel
lstCmdLine.Clear
lstName.Clear
txtcmdline.Text = vbNullString
chk2_Click


Else

If selregentry = check3 Then
ModDel.DeleteValue ERegRoot.eRegRoot_HKeyLocalMachine, "Software\Microsoft\Windows\CurrentVersion\Run", lstName.Text
lstCmdLine.RemoveItem sel
lstName.RemoveItem sel
lstCmdLine.Clear
lstName.Clear
txtcmdline.Text = vbNullString
chk3_Click


Else

If selregentry = check4 Then
ModDel.DeleteValue ERegRoot.eRegRoot_HKeyLocalMachine, "Software\Microsoft\Windows\CurrentVersion\RunOnce", lstName.Text
lstCmdLine.RemoveItem sel
lstName.RemoveItem sel
lstCmdLine.Clear
lstName.Clear
txtcmdline.Text = vbNullString
chk4_Click


Else

If selregentry = check5 Then
ModDel.DeleteValue ERegRoot.eRegRoot_HKeyLocalMachine, "Software\Microsoft\Windows\CurrentVersion\RunServices", lstName.Text
lstCmdLine.RemoveItem sel
lstName.RemoveItem sel
lstCmdLine.Clear
lstName.Clear
txtcmdline.Text = vbNullString
chk5_Click

End If
End If
End If
End If
End If
End If
End If
End If

End Sub

Sub mnuDeletestartupfiles()

If filelistbox.ListCount = 0 Then
Exit Sub
Else

Dim askfiledel As Variant
Dim startupfolder As Variant
Dim sfile As Variant
askfiledel = MsgBox("Are you sure you want to delete the selected file from startup directory", vbYesNo, "Confirm please to delete")
If askfiledel = vbNo Then
Exit Sub
Else

startupfolder = CheckFolderID(Common_StartUp)
sfile = startupfolder & "\" & (filelistbox.FileName)
On Error Resume Next

Kill (sfile)
filelistbox.Refresh

chk6_Click

Exit Sub
End If
End If

End Sub

Private Sub mnufeedback_Click()
On Error Resume Next
frmfeedback.Show

End Sub

Private Sub mnuRestoreaprevoussavedbackup_Click()
On Error Resume Next

Dim askrestore
askrestore = MsgBox("This will restore a pre saved backup do you wish to continue", vbInformation + vbYesNo, "Yes to restore no to not")
If askrestore = vbYes Then

lstName.Clear
lstCmdLine.Clear
txtcmdline.Text = vbNullString
restorebackup

Exit Sub
End If
End Sub


Private Sub mnuvisithomepage_Click()
On Error Resume Next
ShellExecute hwnd, "open", "http://www.vbdotlb.connect.to", vbNullString, vbNullString, 1

End Sub

Sub createbackup()

On Error Resume Next

'Cancel = 1

Dim curdate As Variant
Dim f As Variant
Dim curtime As Variant

curdate = Date
curtime = Time
f = FreeFile

Shell "regedit /e HKCU_Run.reg HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run"
Shell "regedit /e HKCU_RunOnce.reg HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\RunOnce"
Shell "regedit /e HKLM_Run.reg HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Run"
Shell "regedit /e HKLM_RunOnce.reg HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\RunOnce"
Shell "regedit /e HKLM_RunServices.reg HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\RunServices"

On Error Resume Next

SaveSetting App.EXEName, "Settings", "BackupDate", "Date " & Date & " Time:" & Time

'Or you can create a free file
'*****************************************
'Close
'Open "Backup date.txt" For Output As #f
'Print #f, "Last backup date was:"
'Print #f, curdate
'Print #f, ""
'Print #f, "At time"
'Print #f, curtime
'Close

'MsgBox "A new backup has been created successfully for later restore", vbInformation, "Auto backup enabled"

chk1_Click


End Sub
Sub restorebackup()
On Error Resume Next

Shell "regedit /is HKCU_Run.reg HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run"
Shell "regedit /is HKCU_RunOnce.reg HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\RunOnce"
Shell "regedit /is HKLM_Run.reg HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Run"
Shell "regedit /is HKLM_RunOnce.reg HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\RunOnce"
Shell "regedit /is HKLM_RunServices.reg HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\RunServices"
lblinf.Caption = "Restoring complete"

chk1_Click
MsgBox "A pre saved Backup has been restored", vbInformation, "Selected job Done"

'End


Exit Sub
End Sub

Sub EnumRegRun1()
On Error Resume Next
    
hKey = OpenKey(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run")
LCount = GetCount(hKey, Values)
For i = 0 To LCount - 1

lstName.AddItem EnumValue(hKey, i)
lstCmdLine.AddItem GetKeyValue(hKey, EnumValue(hKey, i))

Next i
lstName.ListIndex = 0
lstCmdLine.ListIndex = 0
    
End Sub

Sub EnumRegRun2()
On Error Resume Next

hKey = OpenKey(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\RunOnce")
LCount = GetCount(hKey, Values)
For i = 0 To LCount - 1
        
lstName.AddItem EnumValue(hKey, i)
lstCmdLine.AddItem GetKeyValue(hKey, EnumValue(hKey, i))
Next i
lstName.ListIndex = 0
lstCmdLine.ListIndex = 0
    
End Sub

Sub EnumRegRun3()
On Error Resume Next

hKey = OpenKey(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run")
LCount = GetCount(hKey, Values)
For i = 0 To LCount - 1
lstName.AddItem EnumValue(hKey, i)
lstCmdLine.AddItem GetKeyValue(hKey, EnumValue(hKey, i))
Next i
lstName.ListIndex = 0
lstCmdLine.ListIndex = 0

End Sub

Sub EnumRegRun4()
On Error Resume Next

hKey = OpenKey(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce")
LCount = GetCount(hKey, Values)
For i = 0 To LCount - 1
lstName.AddItem EnumValue(hKey, i)
lstCmdLine.AddItem GetKeyValue(hKey, EnumValue(hKey, i))

Next i
lstName.ListIndex = 0
lstCmdLine.ListIndex = 0

End Sub

Sub EnumRegRun5()
On Error Resume Next

hKey = OpenKey(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunServices")
LCount = GetCount(hKey, Values)
For i = 0 To LCount - 1
lstName.AddItem EnumValue(hKey, i)
lstCmdLine.AddItem GetKeyValue(hKey, EnumValue(hKey, i))
Next i
lstName.ListIndex = 0
lstCmdLine.ListIndex = 0
End Sub
