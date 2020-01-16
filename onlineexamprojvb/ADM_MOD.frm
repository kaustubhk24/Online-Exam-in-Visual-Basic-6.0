VERSION 5.00
Begin VB.Form ADMIN_MOD 
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
   Begin VB.CommandButton Command2 
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
      Left            =   960
      TabIndex        =   14
      Top             =   5280
      Width           =   1095
   End
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
      Height          =   495
      Left            =   2040
      TabIndex        =   13
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "&CLOSE"
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
      Left            =   6480
      TabIndex        =   12
      Top             =   5280
      Width           =   1575
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "&SAVE"
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
      Left            =   4920
      TabIndex        =   11
      Top             =   5280
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&INSERT"
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
      Left            =   3360
      TabIndex        =   10
      Top             =   5280
      Width           =   1575
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
      Left            =   6720
      TabIndex        =   9
      Top             =   4200
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
      Left            =   5760
      TabIndex        =   8
      Top             =   4200
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
      Left            =   4800
      TabIndex        =   7
      Top             =   4200
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
      Left            =   3840
      TabIndex        =   6
      Top             =   4200
      Width           =   975
   End
   Begin VB.TextBox txtmno 
      Height          =   495
      Left            =   6240
      TabIndex        =   5
      Text            =   " "
      Top             =   3120
      Width           =   2055
   End
   Begin VB.TextBox txtmname 
      Height          =   495
      Left            =   6240
      TabIndex        =   4
      Text            =   " "
      Top             =   2400
      Width           =   2055
   End
   Begin VB.TextBox txtmcode 
      Height          =   495
      Left            =   6240
      TabIndex        =   3
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "MODULE-NUMBER"
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
      Left            =   3120
      TabIndex        =   2
      Top             =   3240
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "MODULE--NAME"
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
      Left            =   3120
      TabIndex        =   1
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "MODULE--CODE"
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
      Left            =   3120
      TabIndex        =   0
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Line Line4 
      X1              =   360
      X2              =   8880
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Line Line3 
      X1              =   8880
      X2              =   8880
      Y1              =   360
      Y2              =   6120
   End
   Begin VB.Line Line2 
      X1              =   360
      X2              =   360
      Y1              =   360
      Y2              =   6120
   End
   Begin VB.Line Line1 
      X1              =   360
      X2              =   8880
      Y1              =   360
      Y2              =   360
   End
End
Attribute VB_Name = "ADMIN_MOD"
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
rs.Open "module", cn, adOpenDynamic, adLockOptimistic
rs.AddNew
 rs!mcode = txtmcode.Text
 rs!msub = txtmname.Text
 rs!mno = txtmno.Text
 rs.Update
End Sub

Private Sub Command1_Click()
txtmcode.Text = ""
txtmname.Text = ""
txtmno.Text = ""
txtmcode.SetFocus
End Sub



Private Sub Command2_Click()
'rs!mcode = txtmcode.Text
 rs!msub = txtmname.Text
 rs!mno = txtmno.Text
 End Sub

Private Sub Form_Activate()
Set rs = New ADODB.Recordset
rs.Open "module", cn, adOpenDynamic, adLockOptimistic
display
End Sub
Private Sub Form_Load()
Set cn = New ADODB.Connection
cn.Provider = "microsoft.jet.oledb.3.51"
cn.Open "C:/exam.mdb"
End Sub
Public Sub display()
txtmcode.Text = rs!mcode
txtmname.Text = rs!msub
txtmno.Text = rs!mno
End Sub

