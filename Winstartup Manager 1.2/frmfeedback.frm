VERSION 5.00
Begin VB.Form frmfeedback 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "To feedback"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5865
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmfeedback.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   5865
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1410
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmfeedback.frx":000C
      Top             =   810
      Width           =   5880
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00E0E0E0&
      X1              =   -60
      X2              =   5835
      Y1              =   795
      Y2              =   795
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      X1              =   -15
      X2              =   5880
      Y1              =   2220
      Y2              =   2220
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   810
      Left            =   0
      Picture         =   "frmfeedback.frx":0237
      Stretch         =   -1  'True
      Top             =   15
      Width           =   5880
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   " To feedback please read what's the developer of Winstartup Manager says."
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   0
      TabIndex        =   1
      Top             =   2250
      Width           =   5880
   End
End
Attribute VB_Name = "frmfeedback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


