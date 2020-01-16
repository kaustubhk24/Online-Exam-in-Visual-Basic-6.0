VERSION 5.00
Begin VB.Form SELECTUSER 
   BackColor       =   &H008080FF&
   Caption         =   "ONLINE-EXAMINATION"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "&CONTINUE- ->"
      BeginProperty Font 
         Name            =   "Roman"
         Size            =   15.75
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4200
      Width           =   2895
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ADMINISATRATOR"
      Height          =   495
      Left            =   3120
      TabIndex        =   3
      Top             =   3240
      Width           =   3015
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "STAFF"
      Height          =   495
      Left            =   3120
      TabIndex        =   2
      Top             =   2520
      Width           =   3015
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "STUDENT"
      Height          =   495
      Left            =   3120
      TabIndex        =   1
      Top             =   1680
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "SELECT USER TYPE"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
   Begin VB.Line Line8 
      X1              =   720
      X2              =   720
      Y1              =   600
      Y2              =   6120
   End
   Begin VB.Line Line7 
      X1              =   8640
      X2              =   8640
      Y1              =   6120
      Y2              =   600
   End
   Begin VB.Line Line6 
      X1              =   720
      X2              =   8640
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Line Line5 
      X1              =   720
      X2              =   8640
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line4 
      X1              =   840
      X2              =   8520
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Line Line3 
      X1              =   8520
      X2              =   8520
      Y1              =   720
      Y2              =   6000
   End
   Begin VB.Line Line2 
      X1              =   840
      X2              =   8520
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line1 
      X1              =   840
      X2              =   840
      Y1              =   720
      Y2              =   6000
   End
End
Attribute VB_Name = "SELECTUSER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
If Option1.Value = True Then
 main.OPT = 1
 ElseIf Option2.Value = True Then
  main.OPT = 2
  ElseIf Option3.Value = True Then
  main.OPT = 3
  End If
  login.Show
 End Sub
