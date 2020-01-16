VERSION 5.00
Begin VB.Form ADMIN_LOG 
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
   Begin VB.CommandButton cmddel 
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
      Height          =   615
      Left            =   1440
      TabIndex        =   16
      Top             =   5160
      Width           =   1095
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
      Height          =   615
      Left            =   2520
      TabIndex        =   15
      Top             =   5160
      Width           =   1095
   End
   Begin VB.TextBox txtinf 
      Height          =   495
      Left            =   6360
      TabIndex        =   14
      Top             =   3360
      Width           =   2175
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
      Height          =   615
      Left            =   6000
      TabIndex        =   12
      Top             =   5160
      Width           =   1215
   End
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
      Height          =   615
      Left            =   3600
      TabIndex        =   11
      Top             =   5160
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
      Height          =   615
      Left            =   4800
      TabIndex        =   10
      Top             =   5160
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
      Left            =   7440
      TabIndex        =   9
      Top             =   4320
      Width           =   975
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
      TabIndex        =   8
      Top             =   4320
      Width           =   975
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
      Left            =   5520
      TabIndex        =   7
      Top             =   4320
      Width           =   975
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
      Left            =   4560
      TabIndex        =   6
      Top             =   4320
      Width           =   975
   End
   Begin VB.TextBox txtupw 
      Height          =   495
      Left            =   6360
      TabIndex        =   5
      Text            =   " "
      Top             =   2640
      Width           =   2175
   End
   Begin VB.TextBox txtunm 
      Height          =   495
      Left            =   6360
      TabIndex        =   4
      Text            =   " "
      Top             =   1920
      Width           =   2175
   End
   Begin VB.TextBox txtuid 
      Height          =   495
      Left            =   6360
      TabIndex        =   3
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label4 
      Caption         =   "USERINFORMATION"
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
      Left            =   3600
      TabIndex        =   13
      Top             =   3360
      Width           =   2295
   End
   Begin VB.Line Line4 
      X1              =   480
      X2              =   9120
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line3 
      X1              =   9120
      X2              =   9120
      Y1              =   6120
      Y2              =   840
   End
   Begin VB.Line Line2 
      X1              =   480
      X2              =   9120
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Line Line1 
      X1              =   480
      X2              =   480
      Y1              =   840
      Y2              =   6120
   End
   Begin VB.Label Label3 
      Caption         =   "USER-PASSWORD"
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
      Left            =   3600
      TabIndex        =   2
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "USER-NAME"
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
      Left            =   3600
      TabIndex        =   1
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "USER-ID"
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
      Left            =   3600
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "ADMIN_LOG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
        txtuid.Text = rs.Fields(0)
        txtunm.Text = rs.Fields(1)
        txtupw.Text = rs.Fields(2)
        txtinf.Text = rs.Fields(3)
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
rs.Open "user", cn, adOpenDynamic, adLockOptimistic
rs.AddNew
rs!userid = txtuid.Text
rs!username = txtunm.Text
rs!userpsw = txtupw.Text
rs!userinfo = txtinf.Text
rs.Update
End Sub

Private Sub Command1_Click()
txtuid.Text = ""
txtunm.Text = ""
txtupw.Text = ""
txtinf.Text = ""
txtuid.SetFocus
End Sub

Private Sub Form_Activate()
Set rs = New ADODB.Recordset
rs.Open "user", cn, adOpenDynamic, adLockOptimistic
display
End Sub
Private Sub Form_Load()
Set cn = New ADODB.Connection
cn.Provider = "microsoft.jet.oledb.3.51"
cn.Open "C:/exam.mdb"
End Sub
Public Sub display()
txtuid.Text = rs!userid
txtunm.Text = rs!username
txtupw.Text = rs!userpsw
txtinf.Text = rs!userinfo
End Sub
