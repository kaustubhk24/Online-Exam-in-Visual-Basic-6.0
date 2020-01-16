VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form MODULE 
   BackColor       =   &H00FFC0C0&
   Caption         =   "ONLINE-EXAMINATION"
   ClientHeight    =   6105
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8940
   LinkTopic       =   "Form2"
   ScaleHeight     =   6105
   ScaleWidth      =   8940
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.ListBox rqt 
      BackColor       =   &H00FFC0C0&
      Height          =   450
      ItemData        =   "MODULE.frx":0000
      Left            =   6360
      List            =   "MODULE.frx":001C
      TabIndex        =   5
      Top             =   3720
      Width           =   735
   End
   Begin VB.ListBox noq 
      BackColor       =   &H00FFC0C0&
      Height          =   450
      ItemData        =   "MODULE.frx":003E
      Left            =   6360
      List            =   "MODULE.frx":0054
      TabIndex        =   4
      Top             =   2760
      Width           =   1455
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   5730
      Width           =   8940
      _ExtentX        =   15769
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txt1 
      Height          =   495
      Left            =   3000
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   1200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdcon 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&CONTINUE- ->"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "PRESS CONTINUE  TO ENTER INTO EXAM PAPER"
      Top             =   4560
      Width           =   2055
   End
   Begin VB.TextBox txtmodno 
      BackColor       =   &H00FFC0C0&
      Height          =   375
      Left            =   6360
      TabIndex        =   3
      Text            =   " "
      Top             =   2040
      Width           =   2055
   End
   Begin VB.ComboBox cmbmod 
      BackColor       =   &H00FFC0C0&
      Height          =   315
      ItemData        =   "MODULE.frx":0070
      Left            =   6360
      List            =   "MODULE.frx":0072
      TabIndex        =   2
      ToolTipText     =   "SELECT MODULE NUMBER FOR EXAM"
      Top             =   1320
      Width           =   2535
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFC0C0&
      Caption         =   "MINUTES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7200
      TabIndex        =   12
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFC0C0&
      Caption         =   "REQUIRED-TIME"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   11
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "NO-OF QUESTION"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   10
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "SELECT SUBJECT-NAME,NO OF QUESTION AND TIME"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   9
      Top             =   0
      Width           =   7815
   End
   Begin VB.Image Image3 
      Height          =   1590
      Left            =   480
      Picture         =   "MODULE.frx":0074
      Top             =   3960
      Width           =   12000
   End
   Begin VB.Image Image2 
      Height          =   1590
      Left            =   480
      Picture         =   "MODULE.frx":14AA
      Top             =   840
      Width           =   12000
   End
   Begin VB.Image Image1 
      Height          =   1590
      Left            =   480
      Picture         =   "MODULE.frx":28E0
      Top             =   2400
      Width           =   12000
   End
   Begin VB.Line Line4 
      X1              =   9120
      X2              =   9120
      Y1              =   6240
      Y2              =   480
   End
   Begin VB.Line Line3 
      X1              =   240
      X2              =   9120
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line Line2 
      X1              =   240
      X2              =   9120
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   240
      Y1              =   480
      Y2              =   6240
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "SUBJECT-CODE"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SUBJECT NAME"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   0
      ToolTipText     =   "PRESS TAB KEY"
      Top             =   1320
      Width           =   2175
   End
End
Attribute VB_Name = "MODULE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim STR As String

Private Sub cmbmod_Click()
If cmbmod.Text = "" Then
MsgBox " select module name "
cmbmod.SetFocus
End If
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
txt1.Text = cmbmod.Text
STR = "select * from module where msub ='" & txt1.Text & "'"
rs.Open STR, cn, adOpenDynamic, adLockOptimistic
If rs.EOF And rs.BOF Then
MsgBox " select nodule name correctly"
End If
txtmodno.Text = rs!mcode
main.modcode = rs!mcode
NOQ.SetFocus
End Sub

Private Sub cmbmod_LostFocus()
If cmbmod.Text = "" Then
MsgBox " select module name "
cmbmod.SetFocus
End If
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
txt1.Text = cmbmod.Text
STR = "select * from module where msub ='" & txt1.Text & "'"
rs.Open STR, cn, adOpenDynamic, adLockOptimistic
If rs.EOF And rs.BOF Then
MsgBox " select nodule name correctly"
End If
txtmodno.Text = rs!mcode
main.modcode = rs!mcode
End Sub

Private Sub cmdcon_Click()
INSTRUCTION.Show
End Sub



Private Sub Form_Load()
Set cn = New ADODB.Connection
cn.Provider = "microsoft.jet.oledb.3.51"
cn.Open "C:/exam.mdb"
Set rs = New ADODB.Recordset
rs.Open "module", cn, adOpenDynamic, adLockOptimistic
 While rs.EOF <> True
    cmbmod.AddItem rs!msub
    rs.MoveNext
Wend
rs.Close
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
StatusBar1.Panels(3).Text = "select  module correctly"
StatusBar1.Panels(3).AutoSize = sbrSpring
End Sub
Private Sub noq_Click()
main.NOQ = NOQ.Text
'Print main.noq

End Sub

Private Sub rqt_Click()
main.RT = rqt.Text
' Print main.RT
End Sub

Private Sub rqt_GotFocus()
If main.NOQ = 0 Then
   NOQ.SetFocus
   End If
End Sub
