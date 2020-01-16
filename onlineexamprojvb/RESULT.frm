VERSION 5.00
Begin VB.Form RESULT 
   BackColor       =   &H00000000&
   Caption         =   "ONLINE-EXAMINATION"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdexit 
      BackColor       =   &H0000FF00&
      Caption         =   "&EXIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1200
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5040
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2280
      TabIndex        =   8
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label qsta 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   6480
      TabIndex        =   24
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Label Label12 
      BackColor       =   &H00000000&
      Caption         =   "NO-QUES-ATTEMPTED"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   4200
      TabIndex        =   23
      Top             =   3840
      Width           =   2055
   End
   Begin VB.Label RESNO 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   6480
      TabIndex        =   22
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label Label11 
      BackColor       =   &H00000000&
      Caption         =   "RESULT-NO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   4200
      TabIndex        =   21
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   840
      TabIndex        =   20
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label txtunm1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   720
      TabIndex        =   19
      Top             =   0
      Width           =   3615
   End
   Begin VB.Label TXTGRD 
      BackColor       =   &H00000000&
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   6480
      TabIndex        =   18
      Top             =   5280
      Width           =   2175
   End
   Begin VB.Label txtt 
      BackColor       =   &H00000000&
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   6480
      TabIndex        =   17
      Top             =   4320
      Width           =   2175
   End
   Begin VB.Label txttmk 
      BackColor       =   &H00000000&
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   6480
      TabIndex        =   16
      Top             =   4800
      Width           =   2175
   End
   Begin VB.Label txtmod 
      BackColor       =   &H00000000&
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   6480
      TabIndex        =   15
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Label txtuid 
      BackColor       =   &H00000000&
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   6480
      TabIndex        =   14
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label txtunm2 
      BackColor       =   &H00000000&
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   6480
      TabIndex        =   13
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label txtdate 
      BackColor       =   &H00000000&
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   6480
      TabIndex        =   12
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label NOQ1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   6480
      TabIndex        =   11
      Top             =   3360
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "no-of question"
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
      Height          =   255
      Left            =   4200
      TabIndex        =   10
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label10 
      BackColor       =   &H00000000&
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
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   4200
      TabIndex        =   7
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackColor       =   &H00000000&
      Caption         =   "-MARK LIST"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   4440
      TabIndex        =   6
      Top             =   0
      Width           =   3135
   End
   Begin VB.Line Line4 
      BorderColor     =   &H0000FF00&
      X1              =   240
      X2              =   8760
      Y1              =   6360
      Y2              =   6360
   End
   Begin VB.Line Line3 
      BorderColor     =   &H0000FF00&
      X1              =   8760
      X2              =   8760
      Y1              =   720
      Y2              =   6360
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0000FF00&
      X1              =   240
      X2              =   240
      Y1              =   720
      Y2              =   6360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      X1              =   240
      X2              =   8760
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
      Caption         =   "GRADE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   4320
      TabIndex        =   5
      Top             =   5280
      Width           =   975
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      Caption         =   "TOTAL TIME-TAKEN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   4200
      TabIndex        =   4
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   " MARKS OBTAINED"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   4200
      TabIndex        =   3
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "MODULECODE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   4200
      TabIndex        =   2
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "STUDENT-ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   4200
      TabIndex        =   1
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label LBLNAME 
      BackColor       =   &H00000000&
      Caption         =   "STUDENT -NAME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   4200
      TabIndex        =   0
      Top             =   1920
      Width           =   1695
   End
End
Attribute VB_Name = "RESULT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim STR As String
Dim str1 As String
Dim a As Integer
Private Sub cmdexit_Click()
End
End Sub

Private Sub Form_Activate()
Dim cn1 As ADODB.Connection
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim rs1 As ADODB.Recordset
End Sub

Private Sub Form_Load()
Set cn = New ADODB.Connection
cn.Provider = "microsoft.jet.oledb.3.51"
'cn1.Provider = "microsoft.jet.oledb.3.51"
cn.Open "C:/exam.mdb"
'cn1
Set rs = New ADODB.Recordset
Text1.Text = main.userid
STR = "select * from result where uid ='" & Text1.Text & "'"
rs.Open STR, cn, adOpenDynamic, adLockOptimistic
rs.MoveLast
txtdate.Caption = rs!Date
txtunm1.Caption = main.username
txtunm2.Caption = main.username
txtuid.Caption = main.userid
txtmod.Caption = rs!modcode
'txtinfo.Caption = "a"
txttmk.Caption = rs!marks
TXTGRD.Caption = "A"
txtt.Caption = rs!timetaken
NOQ1.Caption = rs!NOQ
Label1.Caption = rs!sec
RESNO.Caption = rs!RESNO
qsta.Caption = rs!noqatmp
'txtnoq.Text

End Sub

