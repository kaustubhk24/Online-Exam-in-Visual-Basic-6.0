VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmmain 
   BackColor       =   &H00000000&
   Caption         =   "frmmain"
   ClientHeight    =   6120
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7845
   BeginProperty Font 
      Name            =   "Symbol"
      Size            =   18
      Charset         =   2
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6120
   ScaleWidth      =   7845
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   800
      Left            =   4800
      Top             =   4440
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   5400
      Top             =   4440
   End
   Begin MSComctlLib.ProgressBar ProgressBar2 
      Height          =   300
      Left            =   3840
      TabIndex        =   1
      Top             =   5040
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
      Max             =   6
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000007&
      Caption         =   "ARE            YOU           READY !!!......"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   2880
      Width           =   6615
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Loading........."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   " WELCOME                    TO                          ONLINEEXAMINATION"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   2175
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   5415
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim L%
Dim status As Boolean

Private Sub Form_DblClick()
Unload Me
Database.Show
End Sub

Private Sub Form_Load()
frmmain.WindowState = 2
L = 0
Label1.Left = L
status = True
ProgressBar2.Visible = False
Label3.Visible = False
End Sub







Private Sub Timer1_Timer()
Static i As Integer
If status Then
   L = L + 150
   If Label1.Left + Label1.Width > frmmain.Width Then
      status = False
   End If
Else
   L = L - 150
   If Label1.Left < 0 Then
      status = True
   End If
End If
Label1.Left = L
i = i + 1
If i Mod 1 = 0 Then
   Label1.Enabled = Not Label1.Enabled
End If
End Sub

Private Sub Timer2_Timer()
'Unload Me
'Database.Visible = True
ProgressBar2.Visible = True
Label3.Visible = True
If ProgressBar2.Value = 6 Then
   'frmmain.Hide
   Unload Me
   SELECTUSER.Show
Else
   ProgressBar2.Value = ProgressBar2.Value + 1
End If
End Sub


