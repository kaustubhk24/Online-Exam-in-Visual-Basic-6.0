VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Enabled         =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   2640
      TabIndex        =   1
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1320
      Top             =   2280
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   1508
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   615
      Left            =   960
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
tmr1.Enabled = True
If LTrim(tmr1.Caption) = LTrim("0:0:05") Then
   MsgBox " timup"
   End If
   
  End Sub
