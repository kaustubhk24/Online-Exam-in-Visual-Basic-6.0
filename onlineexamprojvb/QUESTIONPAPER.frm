VERSION 5.00
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
      Size            =   13.5
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
   Begin VB.TextBox marks 
      Enabled         =   0   'False
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
      TabIndex        =   28
      Top             =   1200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Timer t 
      Interval        =   1000
      Left            =   6600
      Top             =   1200
   End
   Begin VB.Timer tmobj 
      Interval        =   1000
      Left            =   6000
      Top             =   1200
   End
   Begin VB.CommandButton cmdres 
      BackColor       =   &H00FFC0C0&
      Caption         =   " &RESULT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   14
      Top             =   5160
      Width           =   495
   End
   Begin VB.OptionButton optc 
      BackColor       =   &H00FFC0C0&
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   13
      Top             =   4680
      Width           =   495
   End
   Begin VB.OptionButton optb 
      BackColor       =   &H00FFC0C0&
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   12
      Top             =   4080
      Width           =   495
   End
   Begin VB.OptionButton opta 
      BackColor       =   &H00FFC0C0&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      Width           =   375
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
      Left            =   5400
      TabIndex        =   5
      Top             =   600
      Width           =   855
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
      Left            =   3360
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
      Width           =   975
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "STATAUS"
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
      Left            =   2400
      TabIndex        =   35
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label TOTQ 
      BackColor       =   &H00FFC0C0&
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
      Left            =   4920
      TabIndex        =   34
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "OUTOF"
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
      Left            =   4080
      TabIndex        =   33
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label qno2 
      BackColor       =   &H00FFC0C0&
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
      Left            =   3480
      TabIndex        =   32
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label s6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8760
      TabIndex        =   31
      Top             =   1200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label m6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      TabIndex        =   30
      Top             =   1200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label h6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      TabIndex        =   29
      Top             =   1200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   8760
      TabIndex        =   27
      Top             =   600
      Width           =   135
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
      Caption         =   ":    "
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   8400
      TabIndex        =   26
      Top             =   600
      Width           =   135
   End
   Begin VB.Label sec 
      BackColor       =   &H00000000&
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   8880
      TabIndex        =   25
      Top             =   600
      Width           =   255
   End
   Begin VB.Label min 
      BackColor       =   &H00000000&
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   8520
      TabIndex        =   24
      Top             =   600
      Width           =   255
   End
   Begin VB.Label hor 
      BackColor       =   &H00000000&
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   8160
      TabIndex        =   23
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FF0000&
      Caption         =   "                        QUESTION PAPER"
      BeginProperty Font 
         Name            =   "Times New Roman"
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
      Alignment       =   2  'Center
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
      Caption         =   "REMAINING-TIME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   6
      Top             =   600
      Width           =   1695
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
      Left            =   4800
      TabIndex        =   4
      Top             =   600
      Width           =   495
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
      Left            =   2280
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
      Y1              =   120
      Y2              =   6000
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
Dim unm, grade As String
Dim i, k, qsta As Integer
Dim modcode As String
Dim hor1, sec1, second, min1, q As Integer
Dim s2, m2, h2 As Integer
Dim res, qno1 As Integer
Dim STR, ans, qsans As String
Dim tm As Timer
Private Sub cmdexit_Click()
ans = MsgBox("are you sure to exit", vbYesNo)
If ans = vbYes Then
End
End If
End Sub

Private Sub cmdnext_Click()
qsta = qsta + 1
main.qsta = qsta
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
'ch.Text = chk
rs1.AddNew
'unm = main.username
q = main.NOQ
rs1!unm = main.username
rs1!uid = main.userid
rs1!marks = Val(marks.Text)
rs1!timetaken = h6.Caption + ":" + m6.Caption + ":" + s6.Caption
rs1!Date = CDate(txtdate.Text)
rs1!modcode = txtmod.Text
rs1!sec = second
rs1!NOQ = q
rs1!noqatmp = main.qsta
rs1.Update
rs1.Requery
RESULT.Show
End Sub


Private Sub Form_Activate()
'tmr1.Enabled = True
i = 1
k = 0
End Sub

Private Sub Form_Initialize()
second = 0
i = 0
End Sub

Private Sub Form_Load()
Set cn = New ADODB.Connection
Set rs = New ADODB.Recordset
Set rs1 = New ADODB.Recordset
Set rs2 = New ADODB.Recordset
cn.Provider = "microsoft.jet.oledb.3.51"
cn.Open "C:/exam.mdb"
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
qno1 = 1
s2 = 60
m2 = main.RT
QUESTIONPAPER.Show
TOTQ.Caption = main.NOQ
qno2.Caption = qno1
Call checkans
End Sub


Public Sub navi()
Dim j As Integer

If i = main.NOQ Then Exit Sub
      qno1 = qno1 + 1
      qno2.Caption = qno1
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
                'Print mark
                End If
       opta.Value = False
       optb.Value = False
       optc.Value = False
       optd.Value = False
       End Sub
 


Private Sub Form_LostFocus()
main.second = second
main.qsta = qsta
End Sub

Private Sub t_Timer()
 hor1 = 0
 sec1 = sec1 + 1
 second = second + 1
 main.second = second
 Print second
    If sec1 = 60 Then
        sec1 = 0
       min1 = min1 + 1
       If min1 = 60 Then
          min1 = 0
          hor1 = hor1 + 1
           If hor1 = 24 Then
              hor1 = 0
            End If
       End If
     End If
     
     s2 = s2 - 1
     If s2 = 0 Then
        s2 = 60
          m2 = m2 - 1
          If m2 = 0 Then
             m2 = 60
             h2 = h2 - 1
                If h2 = 0 Then
                  h = 24
                 End If
            End If
       End If
   h6.Caption = hor1
   m6.Caption = min1
   s6.Caption = sec1
   sec.Caption = s2
    If main.RT = 1 Then
    m2 = "00"
    End If
   min.Caption = m2
   hor.Caption = h2
  End Sub

Private Sub tmobj_Timer()
Dim t As String
If Val(min1) = main.RT Then
Call cmdres_Click
End If
End Sub
