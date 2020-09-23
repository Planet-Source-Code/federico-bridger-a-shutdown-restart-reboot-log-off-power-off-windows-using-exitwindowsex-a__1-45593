VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Exit Windows"
   ClientHeight    =   3345
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4830
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3345
   ScaleWidth      =   4830
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      Picture         =   "frmMain.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton cmdAccept 
      Caption         =   "&Accept"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      Picture         =   "frmMain.frx":0294
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Choose your option"
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4575
      Begin VB.OptionButton optPoweroff 
         Caption         =   "Power Off (ATX)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   6
         Top             =   1560
         Width           =   1575
      End
      Begin VB.OptionButton optLogoff 
         Caption         =   "Logoff"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   5
         Top             =   1200
         Width           =   1575
      End
      Begin VB.OptionButton optReboot 
         Caption         =   "Reboot"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   4
         Top             =   840
         Width           =   1575
      End
      Begin VB.OptionButton optShutdown 
         Caption         =   "Shutdown"
         Height          =   255
         Left            =   2040
         TabIndex        =   3
         Top             =   480
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.PictureBox picLogo 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   600
         Picture         =   "frmMain.frx":03DE
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   1
         Top             =   600
         Width           =   480
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         Caption         =   "Developed By Federico Bridger (Rosario, Arg.)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   675
         Left            =   240
         TabIndex        =   2
         Top             =   1080
         Width           =   1395
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAccept_Click()

  Dim cExitWindows As New clsExitWindows
  
  If optShutdown.Value Then

    cExitWindows.ExitWindows WE_SHUTDOWN

  ElseIf optLogoff.Value Then

    cExitWindows.ExitWindows WE_LOGOFF

  ElseIf optReboot.Value Then

    cExitWindows.ExitWindows WE_REBOOT

  ElseIf optPoweroff.Value Then

    cExitWindows.ExitWindows WE_POWEROFF

  End If
  
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub
