VERSION 5.00
Begin VB.Form frmwait 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " "
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4200
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3750
      Top             =   -15
   End
   Begin VB.Line Line1 
      X1              =   -165
      X2              =   4305
      Y1              =   765
      Y2              =   765
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Registry important sections backup"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   45
      TabIndex        =   1
      Top             =   60
      Width           =   4110
   End
   Begin VB.Shape sh 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   3  'Vertical Line
      Height          =   225
      Left            =   45
      Top             =   420
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Creating first runtime backup please wait ..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   15
      TabIndex        =   0
      Top             =   810
      Width           =   4185
   End
   Begin VB.Shape shmain 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   45
      Top             =   1815
      Width           =   4080
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   45
      Shape           =   4  'Rounded Rectangle
      Top             =   405
      Width           =   4110
   End
End
Attribute VB_Name = "frmwait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long

Private Sub Form_Initialize()
On Error Resume Next

Dim X As Variant
X = InitCommonControls

End Sub

Private Sub Form_Load()
Me.Height = 1455

Timer1.Enabled = True

End Sub

Private Sub Timer1_Timer()

On Error Resume Next

sh.Visible = True
sh.Width = sh.Width + 10
pb.Width = sh.Width
If sh.Width > shmain.Width Then
SaveSetting App.EXEName, "Settings", "FirstRun", "Done"

frmmain.createbackup
Timer1.Enabled = False
Unload Me
frmmain.lbldate.Caption = " First backup created" & GetSetting(App.EXEName, "settings", "BackupDate", 0)
frmmain.Show
End If

End Sub

