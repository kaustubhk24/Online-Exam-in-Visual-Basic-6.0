VERSION 5.00
Object = "{47780D65-7A5E-11D4-8B28-008048B7D6C6}#3.0#0"; "ttt.ocx"
Begin VB.Form QUESTIONPAPER 
   BackColor       =   &H00FFFFC0&
   Caption         =   "ONLINE-EXAMINATION                                                              QUESTION PAPER"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   FillColor       =   &H00FFFF80&
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFC0&
   LinkTopic       =   "Form3"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin Project1.tmr tmr1 
      Height          =   375
      Left            =   7800
      TabIndex        =   25
      Top             =   600
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Enabled         =   0   'False
      Caption         =   " 0: 0: 10"
   End
   Begin VB.TextBox ch 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   24
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox marks 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   23
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Timer tmobj 
      Interval        =   1000
      Left            =   3120
      Top             =   1200
   End
   Begin VB.CommandButton cmdres 
      BackColor       =   &H00FFC0C0&
      Caption         =   " &RESULT"
      Height          =   495
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "TO GET RESULT OF EXAM"
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton cmdexit 
      BackColor       =   &H00FFC0C0&
      Caption         =   "E&XIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "EXIT FROM EXAM"
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton cmdnext 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&NEXT- - ->"
      Height          =   495
      Left            =   720
      MaskColor       =   &H8000000A&
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "CONTINUE TO NEXT QUESTION"
      Top             =   5760
      Width           =   1815
   End
   Begin VB.OptionButton optd 
      BackColor       =   &H00FFC0C0&
      Caption         =   "D"
      Height          =   495
      Left            =   360
      TabIndex        =   14
      Top             =   5160
      Width           =   495
   End
   Begin VB.OptionButton optc 
      BackColor       =   &H00FFC0C0&
      Caption         =   "C"
      Height          =   375
      Left            =   360
      TabIndex        =   13
      Top             =   4680
      Width           =   495
   End
   Begin VB.OptionButton optb 
      BackColor       =   &H00FFC0C0&
      Caption         =   "B"
      Height          =   495
      Left            =   360
      TabIndex        =   12
      Top             =   4080
      Width           =   495
   End
   Begin VB.OptionButton opta 
      BackColor       =   &H00FFC0C0&
      Caption         =   "A"
      Height          =   375
      Left            =   360
      TabIndex        =   11
      Top             =   3600
      Width           =   495
   End
   Begin VB.TextBox txtqno 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   8
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox txtdate 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6000
      TabIndex        =   5
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox txtmod 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox txtuid 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FF0000&
      Caption         =   "                        QUESTION PAPER"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   375
      Left            =   240
      TabIndex        =   22
      Top             =   0
      Width           =   9015
   End
   Begin VB.Label lbld 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   18
      Top             =   5160
      Width           =   8175
   End
   Begin VB.Label lblc 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   17
      Top             =   4680
      Width           =   8175
   End
   Begin VB.Label lblb 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   16
      Top             =   4080
      Width           =   8175
   End
   Begin VB.Label lbla 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   15
      Top             =   3600
      Width           =   8175
   End
   Begin VB.Label lblqnm 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   360
      TabIndex        =   10
      Top             =   1920
      Width           =   8775
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFC0C0&
      Caption         =   "QUESTION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   1680
      Width           =   8775
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFC0C0&
      Caption         =   "QUESTION NO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "TIME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      TabIndex        =   6
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "DATE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5400
      TabIndex        =   4
      Top             =   600
      Width           =   615
   End
   Begin VB.Line Line5 
      X1              =   240
      X2              =   9240
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "MODULENO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "USERNO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   855
   End
   Begin VB.Line Line4 
      X1              =   240
      X2              =   9240
      Y1              =   6360
      Y2              =   6360
   End
   Begin VB.Line Line3 
      X1              =   240
      X2              =   240
      Y1              =   480
      Y2              =   6360
   End
   Begin VB.Line Line2 
      X1              =   9240
      X2              =   9240
      Y1              =   480
      Y2              =   6360
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   9240
      Y1              =   480
      Y2              =   480
   End
End
Attribute VB_Name = "QUESTIONPAPER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mark As Integer
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim rs1 As ADODB.Recordset
Dim no(10) As Integer
Dim unm As String
Dim i, k As Integer
Dim modcode As String
Dim h, s, m As Integer
Dim res As Integer
Dim STR, ans, qsans As String
Dim tm As Timer
Private Sub cmdexit_Click()
ans = MsgBox("are you sure to exit", vbYesNo)
If ans = vbYes Then
End
End If
End Sub

Private Sub cmdnext_Click()
If opta.Value = False And optb.Value = False And optc.Value = False And optd.Value = False Then
MsgBox " select any answer  and then click next"
Exit Sub
End If
Call checkans
If i = 10 Then
cmdres.SetFocus
cmdnext.Enabled = False
ans = Empty
End If
Call navi
'If i = 11 Then cmdnext.Enabled = False
End Sub

Private Sub cmdnext_GotFocus()
If opta.Value = False And optb.Value = False And optc.Value = False And optd.Value = False Then
MsgBox " select any answer  and then click next"
End If
End Sub

Private Sub cmdres_Click()
Dim nanme As String
marks.Text = mark
ch.Text = chk
rs1.AddNew
'unm = main.username
rs1!unm = main.username
rs1!uid = main.userid
rs1!marks = Val(marks.Text)
rs1!timetaken = tmr1.Caption
rs1!Date = CDate(txtdate.Text)
rs1!modcode = txtmod.Text
rs1.Update
rs1.Requery
RESULT.Show
End Sub


Private Sub Form_Activate()
tmr1.Enabled = True
i = 1
k = 0
End Sub

Private Sub Form_Initialize()
i = 0
End Sub

Private Sub Form_Load()
Set cn = New ADODB.Connection
Set rs = New ADODB.Recordset
Set rs1 = New ADODB.Recordset
Set rs2 = New ADODB.Recordset
cn.Provider = "microsoft.jet.oledb.3.51"
cn.Open "z:/exam/exam.mdb"
txtuid.Text = main.userid
txtmod.Text = main.modcode
txtdate.Text = Date
no(0) = Int(Rnd * 10) + 1
STR = " select * from qbank where mcode ='" & Trim(txtmod.Text) & "'and qno=" & no(0)
rs.Open STR, cn, adOpenDynamic, adLockOptimistic
rs1.Open "result", cn, adOpenDynamic, adLockOptimistic
txtqno.Text = rs!qno
lblqnm.Caption = rs!qnm
lbla.Caption = rs!opt1
lblb.Caption = rs!opt2
lblc.Caption = rs!opt3
lbld.Caption = rs!opt4
qsnas = rs!qans
s = 0
m = 0
h = 0
Call checkans
End Sub


Public Sub navi()
Dim j As Integer

If i = 10 Then Exit Sub
     k = k + 1
lb:
     rn = Int(Rnd * 10) + 1
          
     For m = 0 To k
        If no(m) = rn Then
         GoTo lb
        End If
        Next m
        no(0 + k) = rn
        rs.Close
        STR = " select * from qbank where mcode ='" & Trim(txtmod.Text) & "'and qno=" & rn
        '" slect * from qbank where mcode = '" & Trim(txtmod.Text)"' and qno = "&rn
        rs.Open STR, cn, adOpenDynamic, adLockOptimistic
        txtqno.Text = rs!qno
        lblqnm.Caption = rs!qnm
        lblqnm.Caption = rs!qnm
        lbla.Caption = rs!opt1
        lblb.Caption = rs!opt2
        lblc.Caption = rs!opt3
        lbld.Caption = rs!opt4
        qsans = rs!qans
        i = i + 1
        Exit Sub
        End Sub


Public Sub checkans()
  If opta.Value = True Then
        ans = "a"
        chk = chk + 1
        ElseIf optb.Value = True Then
            ans = "b"
            chk = chk + 1
            ElseIf optc.Value = True Then
             ans = "c"
              chk = chk + 1
              ElseIf optd.Value = True Then
               ans = "d"
               chk = chk + 1
  End If
               If ans = qsans Then
                mark = mark + 1
                Print mark
                End If
       opta.Value = False
       optb.Value = False
       optc.Value = False
       optd.Value = False
       End Sub
       
Private Sub tmobj_Timer()
Dim t As String
t = Mid(tmr1.Caption, 4, 2)
Print t
'If Val(t) = 30 Then
If Val(t) = 1 Then
Call cmdres_Click

End If
End Sub
