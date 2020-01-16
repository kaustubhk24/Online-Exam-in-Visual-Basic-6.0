VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form login 
   BackColor       =   &H00FFC0FF&
   Caption         =   "ONLINE-EXAMINATION"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   2940
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdcon 
      BackColor       =   &H00FFC0FF&
      Caption         =   "&CONTINUE- ->"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4560
      Width           =   2775
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2280
      Top             =   1320
   End
   Begin VB.CommandButton cmdcre 
      BackColor       =   &H00FFC0FF&
      Caption         =   " &NEW-USER CREATION -ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2400
      Width           =   1815
   End
   Begin VB.TextBox txtupw 
      BackColor       =   &H00FFC0FF&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   6360
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   3840
      Width           =   2295
   End
   Begin VB.TextBox txtunm 
      BackColor       =   &H00FFC0FF&
      Height          =   375
      Left            =   6360
      TabIndex        =   5
      Top             =   3000
      Width           =   2295
   End
   Begin VB.TextBox txtuid 
      BackColor       =   &H00FFC0FF&
      Height          =   375
      Left            =   6360
      TabIndex        =   2
      Top             =   2160
      Width           =   2295
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFC0FF&
      Caption         =   "WELCOME TO  ON LINE EXAMINATION SYSTEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2760
      TabIndex        =   9
      Top             =   1440
      Width           =   5295
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   1320
      Picture         =   "login.frx":0000
      Top             =   5880
      Width           =   7335
   End
   Begin VB.Line Line8 
      X1              =   480
      X2              =   9120
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Line Line7 
      X1              =   9120
      X2              =   9120
      Y1              =   1200
      Y2              =   5640
   End
   Begin VB.Line Line6 
      X1              =   480
      X2              =   9120
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line5 
      X1              =   480
      X2              =   480
      Y1              =   1200
      Y2              =   5640
   End
   Begin VB.Label password 
      BackColor       =   &H00FFC0FF&
      Caption         =   " PASSWORD"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   4
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0FF&
      Caption         =   "USER- NAME"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0FF&
      Caption         =   "USER- ID"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0FF&
      Caption         =   "WELCOME TO ONLIINE EXAMINATION"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9495
   End
   Begin VB.Image Image2 
      Height          =   300
      Left            =   1320
      Picture         =   "login.frx":086D
      Top             =   720
      Width           =   7335
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Variant
Dim STR As String
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Private Sub cmdcon_Click()
If txtupw.Text = "" Then
MsgBox " enter password"
txtupw.SetFocus
Exit Sub
End If
STR = " SELECT * from user where userid ='" & txtuid.Text & "'" & " And userpsw = '" & txtupw.Text & "'"
rs.Open STR, cn, adOpenDynamic, adLockOptimistic
If rs.EOF And rs.BOF Then
 MsgBox " enter password or name correctly "
 txtuid.SetFocus
 Exit Sub
 Else
 main.userid = txtuid.Text
 main.username = txtunm.Text


 If main.OPT = 3 Then
 MDIForm1.MN_FAC.Enabled = False
 MDIForm1.MN_EXAM.Enabled = False
 MDIForm1.MN_REPO.Enabled = False
 MDIForm1.Show
 ElseIf main.OPT = 2 Then
 MDIForm1.MN_ADMIN.Enabled = False
 MDIForm1.MN_EXAM.Enabled = False
 MDIForm1.Show
 ElseIf main.OPT = 1 Then
 cmdcre.Enabled = True
 MODULE.Show
  End If
 End If
 End Sub

Private Sub cmdcre_Click()
REGISTER.Show
End Sub


Private Sub Form_Activate()
If main.OPT = 1 Then
  MsgBox "  ENTER USERID  in numbers from 1 to 10000 "
  Exit Sub
  End If
End Sub


Private Sub Form_Load()
Set cn = New ADODB.Connection
cn.Provider = "microsoft.jet.oledb.3.51"
cn.Open "C:/exam.mdb"
Set rs = New ADODB.Recordset
Dim mypanel As Panel
StatusBar1.Panels.Clear
Set mypanel = StatusBar1.Panels.Add(1, , , sbrDate)
mypanel.AutoSize = sbrNoAutoSize
mypanel.Bevel = sbrInset
Set mypanel = StatusBar1.Panels.Add(1, , , sbrTime)
mypanel.AutoSize = sbrNoAutoSize
mypanel.Bevel = sbrInset
mypanel.Alignment = sbrLeft
Set mypanel = StatusBar1.Panels.Add(3)
StatusBar1.Panels(3).Text = "enter id ,name and password correctly"
StatusBar1.Panels(3).AutoSize = sbrSpring
If main.OPT = 2 Or main.OPT = 3 Then
cmdcre.Enabled = False
End If
End Sub



Private Sub Timer1_Timer()
Label5.BackColor = RGB(Rnd * 221, Rnd * 223, Rnd * 222)
Label5.Move a
a = a + 100
If a = 12000 Then
a = 100
End If
End Sub
