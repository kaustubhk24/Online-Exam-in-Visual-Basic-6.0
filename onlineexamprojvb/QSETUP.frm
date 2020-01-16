VERSION 5.00
Begin VB.Form ADMIN_QPAPER 
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
   Begin VB.CommandButton cmdsave 
      Caption         =   "&SAVE"
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
      Left            =   4320
      TabIndex        =   26
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "&CLOSE"
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
      Left            =   6720
      TabIndex        =   25
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton cmdmodi 
      Caption         =   "&MODIFY"
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
      Left            =   3120
      TabIndex        =   24
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton cmddelete 
      Caption         =   "&DELETE"
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
      Left            =   5520
      TabIndex        =   23
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&ADD"
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
      Left            =   1920
      TabIndex        =   22
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton cmdlast 
      Caption         =   ">>"
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
      Left            =   7200
      TabIndex        =   21
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton cmdnext 
      Caption         =   ">"
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
      Left            =   6480
      TabIndex        =   20
      Top             =   4800
      Width           =   735
   End
   Begin VB.CommandButton cmdpre 
      Caption         =   "<"
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
      Left            =   5760
      TabIndex        =   19
      Top             =   4800
      Width           =   735
   End
   Begin VB.CommandButton cmdfirst 
      Caption         =   "<<"
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
      Left            =   5160
      TabIndex        =   18
      Top             =   4800
      Width           =   615
   End
   Begin VB.TextBox txtqmarks 
      Height          =   375
      Left            =   3960
      TabIndex        =   17
      Top             =   4800
      Width           =   855
   End
   Begin VB.TextBox txtqans 
      Height          =   375
      Left            =   1560
      TabIndex        =   15
      Top             =   4800
      Width           =   855
   End
   Begin VB.TextBox txtopt4 
      Height          =   375
      Left            =   1560
      TabIndex        =   14
      Top             =   4320
      Width           =   6615
   End
   Begin VB.TextBox txtopt3 
      Height          =   375
      Left            =   1560
      TabIndex        =   13
      Top             =   3840
      Width           =   6615
   End
   Begin VB.TextBox txtopt2 
      Height          =   375
      Left            =   1560
      TabIndex        =   12
      Top             =   3360
      Width           =   6615
   End
   Begin VB.TextBox txtopt1 
      Height          =   375
      Left            =   1560
      TabIndex        =   11
      Top             =   2880
      Width           =   6615
   End
   Begin VB.TextBox txtqnm 
      Height          =   1335
      Left            =   480
      TabIndex        =   10
      Top             =   1320
      Width           =   8535
   End
   Begin VB.TextBox txtqno 
      Height          =   375
      Left            =   7560
      TabIndex        =   9
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox txtmcode 
      Height          =   375
      Left            =   2400
      TabIndex        =   8
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label9 
      Caption         =   "MARKS"
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
      Left            =   2880
      TabIndex        =   16
      Top             =   4920
      Width           =   735
   End
   Begin VB.Line Line4 
      X1              =   240
      X2              =   9120
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line3 
      X1              =   9120
      X2              =   9120
      Y1              =   6480
      Y2              =   480
   End
   Begin VB.Line Line2 
      X1              =   240
      X2              =   9120
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   240
      Y1              =   480
      Y2              =   6480
   End
   Begin VB.Label Label8 
      Caption         =   "QNAS"
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
      Left            =   480
      TabIndex        =   7
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label Label7 
      Caption         =   "OPT-D"
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
      Left            =   480
      TabIndex        =   6
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "OPT-C"
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
      Left            =   480
      TabIndex        =   5
      Top             =   3840
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "OPT-B"
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
      Left            =   480
      TabIndex        =   4
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "OPT-A"
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
      Left            =   480
      TabIndex        =   3
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "QUSTION-NAME"
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
      Left            =   3600
      TabIndex        =   2
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "QUESTION-NO"
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
      Left            =   5760
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "MODULE-CODE"
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
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   1695
   End
End
Attribute VB_Name = "ADMIN_QPAPER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Sub Form_Load()
 ' Text1.Text = Mid$(str(Time), 3, 2)
'end Sub

Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim nm As String
Dim id As String
Dim pass As String
Dim info As String
Private Sub cmddel_Click()
rs.Delete
rs.MoveNext
If rs.EOF Then rs.MoveLast
display
End Sub

Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub cmdfirst_Click()
rs.MoveFirst
display
End Sub

Private Sub cmdlast_Click()
rs.MoveLast
display
End Sub

Private Sub cmdmodi_Click()
id = InputBox("Enter the user id", "modification")
rs.MoveFirst
Do While Not rs.EOF
    If rs.Fields(0) = id Then
        If rs.EditMode = adEditNone Then
        txtmcode.Text = rs!mcode
        txtqnm.Text = rs!qnm
        txtqno.Text = rs!qno
        txtopt1.Text = rs!opt1
        txtopt2.Text = rs!opt2
        txtopt3.Text = rs!opt3
        txtopt4.Text = rs!opt4
        txtqans.Text = rs!qans
        txtqmarks.Text = rs!qmarks

        End If
        Exit Do
    Else
        rs.MoveNext
    End If
Loop
cmdsave.SetFocus
End Sub

Private Sub cmdnext_Click()
rs.MoveNext
If rs.EOF Then rs.MoveLast
display
End Sub

Private Sub cmdpre_Click()
rs.MovePrevious
If rs.BOF Then rs.MoveFirst
display
End Sub
Private Sub cmdsave_Click()
rs.Close
rs.Open "qbank", cn, adOpenDynamic, adLockOptimistic
rs.AddNew
 rs!mcode = txtmcode.Text
 rs!qno = txtqno.Text
 rs!qnm = txtqnm.Text
 rs!opt1 = txtopt1.Text
 rs!opt2 = txtopt2.Text
 rs!opt3 = txtopt3.Text
 rs!opt4 = txtopt4.Text
 rs!qans = txtqans.Text
 rs!qmarks = txtqmarks.Text
 rs.Update
End Sub

Private Sub Command1_Click()
txtmcode.Text = ""
txtqnm.Text = ""
txtqno.Text = ""
txtopt1.Text = ""
txtopt2.Text = ""
txtopt3.Text = ""
txtopt4.Text = ""
txtqans.Text = ""
txtqmarks.Text = ""
txtmcode.SetFocus
End Sub
Private Sub Form_Activate()
Set rs = New ADODB.Recordset
rs.Open "qbank", cn, adOpenDynamic, adLockOptimistic
display
End Sub
Private Sub Form_Load()
Set cn = New ADODB.Connection
cn.Provider = "microsoft.jet.oledb.3.51"
cn.Open "C:/exam.mdb"
End Sub
Public Sub display()
txtmcode.Text = rs!mcode
txtqnm.Text = rs!qnm
txtqno.Text = rs!qno
txtopt1.Text = rs!opt1
txtopt2.Text = rs!opt2
txtopt3.Text = rs!opt3
txtopt4.Text = rs!opt4
txtqans.Text = rs!qans
'txtqmarks.Text = rs!qmarks
End Sub

