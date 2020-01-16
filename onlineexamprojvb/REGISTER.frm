VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form REGISTER 
   BackColor       =   &H00C0C0FF&
   Caption         =   "ONLINE-EXAMINATION"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   FillColor       =   &H00000080&
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdcnf 
      BackColor       =   &H00FF8080&
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1800
      Width           =   1575
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   2820
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtinfo 
      Height          =   495
      Left            =   7320
      TabIndex        =   6
      ToolTipText     =   "ENTER INFORMATION"
      Top             =   4320
      Width           =   1695
   End
   Begin VB.CommandButton CMDREG 
      BackColor       =   &H00FF8080&
      Caption         =   "&REGISTER ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton cmdmodi 
      BackColor       =   &H00FF8080&
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
      Height          =   600
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CommandButton cmdexam 
      BackColor       =   &H00FF8080&
      Caption         =   "&EXAM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton cmdexit 
      BackColor       =   &H00FF8080&
      Caption         =   "E&XIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3600
      Width           =   1575
   End
   Begin VB.TextBox txtpsw 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   7320
      PasswordChar    =   "*"
      TabIndex        =   5
      ToolTipText     =   "ENTER PASSWORD"
      Top             =   3480
      Width           =   1695
   End
   Begin VB.TextBox txtname 
      Height          =   495
      Left            =   7320
      TabIndex        =   4
      Text            =   " "
      ToolTipText     =   "ENTER NAME "
      Top             =   2640
      Width           =   1695
   End
   Begin VB.TextBox txtid 
      Height          =   495
      Left            =   7320
      TabIndex        =   3
      Text            =   " "
      ToolTipText     =   "ENTER ID CORRECTLY"
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Line Line4 
      X1              =   9240
      X2              =   9240
      Y1              =   5280
      Y2              =   960
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   9240
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   120
      Y1              =   960
      Y2              =   5280
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   9240
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0FF&
      Caption         =   " QUALIFIACATION"
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
      Left            =   4440
      TabIndex        =   12
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0FF&
      Caption         =   " PASSWORD"
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
      Left            =   4440
      TabIndex        =   2
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0FF&
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
      Height          =   255
      Left            =   4440
      TabIndex        =   1
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0FF&
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
      Height          =   255
      Left            =   4440
      TabIndex        =   0
      Top             =   1920
      Width           =   1215
   End
End
Attribute VB_Name = "REGISTER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim con As ADODB.Connection
Dim STR As String


Private Sub cmdcnf_Click()
On Error GoTo xy
Set rs = New ADODB.Recordset
rs.Open "user", con, adOpenDynamic, adLockPessimistic
rs.AddNew
rs!userid = Val(txtid.Text)
 rs!username = txtname.Text
 rs!userpsw = txtpsw.Text
 rs!userinfo = Val(Text4)
 main.info = txtinfo.Text
rs.Update
Exit Sub
'Resume Next
xy:
MsgBox " YOUR USER-ID IS ALREADY CREATED  TAKE SOME OTHER NUMBER "
'End If
StatusBar1.Panels(3).Text = "YOUR ID IS  CREATED "
'rs.Close
'Set con = New ADODB.Connection
'con.Provider = "microsoft.jet.oledb.3.51"
'con.Open "z:/exam/exam.mdb"
'Set rs = New ADODB.Recordset
' STR = " SELECT * from user where userid <>'" & txtid.Text & "'"
 ''rs.Open STR, con, adOpenDynamic, adLockOptimistic
 'If rs.EOF And rs.BOF Then
 'Text1.Text = rs!userid
 'MsgBox " YOUR USER-ID IS ALREADY CREATED  TAKE SOME OTHER NUMBER "
 'CMDSAVE.SetFocus
 'Exit Sub
 'Else
 'Text1.Text = rs!userid
 'MsgBox " YOUR ID IS ACCEPTED--THEN CLICK SAVE "
'  End If
End Sub

Private Sub cmdexam_Click()
login.Show
End Sub

Private Sub cmdexit_Click()
End
End Sub
Private Sub CMDREG_Click()
txtid.Text = ""
txtname.Text = ""
txtpsw.Text = ""
txtinfo.Text = ""
txtid.SetFocus
'rs.AddNew
StatusBar1.Panels(3).Text = "CLICK ON SAVE BUTTON AFTER  DETAILS "
End Sub

Private Sub cmdsave_Click()
On Error GoTo xy
Set rs = New ADODB.Recordset
rs.Open "user", con, adOpenDynamic, adLockPessimistic
rs.AddNew
rs!userid = Val(txtid.Text)
 rs!username = txtname.Text
 rs!userpsw = txtpsw.Text
 rs!userinfo = Val(Text4)
rs.Update
StatusBar1.Panels(3).Text = "YOUR ID IS  CREATED "

xy:
 MsgBox " YOUR USER-ID IS ALREADY CREATED  TAKE SOME OTHER NUMBER "
 End Sub



Private Sub Form_Activate()
MsgBox " ENTER USER ID BETWEEN 1 TO 10000 "
CMDREG.SetFocus
End Sub

Private Sub Form_Load()
Set con = New ADODB.Connection
con.Provider = "microsoft.jet.oledb.3.51"
con.Open "C:/exam.mdb"
'Set rs = New ADODB.Recordset
'rs.Open "user", con, adOpenDynamic, adLockPessimistic
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
StatusBar1.Panels(3).Text = "WELCOME TO ON LINE EXAMINATION "
StatusBar1.Panels(3).AutoSize = sbrSpring
'CMDREG.SetFocus

End Sub



Private Sub txtinfo_GotFocus()
'If txtinfo.Text = "" Then
 ' MsgBox " ENTER USER INFO THEN  CLICK SAVE BUTTON"
  'End If
End Sub
Private Sub txtname_Change()
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 32 Then Exit Sub
  If IsNumeric(Chr(KeyAscii)) Or _
  (KeyAscii >= 33 And KeyAscii <= 64) Or _
  (KeyAscii >= 91 And KeyAscii <= 96) Or _
  (KeyAscii >= 123 And KeyAscii <= 126) Then
  KeyAscii = 0
StatusBar1.Panels(3).Text = "user name should be charecter"
End If
End Sub

Private Sub txtname_GotFocus()
If txtid.Text = "" Then
  MsgBox " ENTER USER ID THEN USER NAME"
  End If
End Sub

Private Sub txtname_LostFocus()
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 32 Then Exit Sub
  If IsNumeric(Chr(KeyAscii)) Or _
  (KeyAscii >= 33 And KeyAscii <= 64) Or _
  (KeyAscii >= 91 And KeyAscii <= 96) Or _
  (KeyAscii >= 123 And KeyAscii <= 126) Then
  KeyAscii = 0
StatusBar1.Panels(3).Text = "user name should be charecter"
End If
End Sub


Private Sub txtpsw_GotFocus()
If txtname.Text = "" Then
  MsgBox " ENTER USER NAME THEN USER INFORMATION"
  End If
End Sub
