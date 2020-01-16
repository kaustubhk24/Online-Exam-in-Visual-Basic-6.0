VERSION 5.00
Begin VB.Form INSTRUCTION 
   BackColor       =   &H00C0FFC0&
   Caption         =   "ONLINE-EXAMINATION"
   ClientHeight    =   3195
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "CONTINUE- ->"
      BeginProperty Font 
         Name            =   "Roman"
         Size            =   12
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5520
      Width           =   2295
   End
   Begin VB.Label RT 
      BackColor       =   &H0080FF80&
      Height          =   375
      Left            =   5760
      TabIndex        =   8
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFC0C0&
      Caption         =   "NEXT- ->"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label NOQ 
      BackColor       =   &H0080FF80&
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
      Left            =   3360
      TabIndex        =   6
      Top             =   3600
      Width           =   375
   End
   Begin VB.Label Label5 
      BackColor       =   &H0080FF80&
      Caption         =   $"INSTRUCTION.frx":0000
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
      Left            =   1200
      TabIndex        =   5
      Top             =   3600
      Width           =   7215
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080FF80&
      Caption         =   "3 . SELECT ANSWER BY SELECTING CONCERN RADIO BUTTON     THEN                       BUTTON TO SEE THE NEXT QUESTION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1200
      TabIndex        =   4
      Top             =   2520
      Width           =   7215
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080FF80&
      Caption         =   "1  .QUESTION PAPER IS IN OBJECTIVE TYPE  "
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
      Left            =   1200
      TabIndex        =   3
      Top             =   1200
      Width           =   7215
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FF80&
      Caption         =   "2 .  SELECT MODULE NAME TO ENTER  IN TO EXAM "
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
      Left            =   1200
      TabIndex        =   2
      Top             =   1800
      Width           =   7215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "INSTRUCTIONS"
      BeginProperty Font 
         Name            =   "Roman"
         Size            =   24
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin VB.Line Line4 
      X1              =   720
      X2              =   720
      Y1              =   6240
      Y2              =   360
   End
   Begin VB.Line Line3 
      X1              =   8640
      X2              =   720
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line Line2 
      X1              =   8640
      X2              =   8640
      Y1              =   360
      Y2              =   6240
   End
   Begin VB.Line Line1 
      X1              =   720
      X2              =   8640
      Y1              =   360
      Y2              =   360
   End
End
Attribute VB_Name = "INSTRUCTION"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload login

QUESTIONPAPER.Show
End Sub

Private Sub Form_Load()
RT.Caption = main.RT
NOQ.Caption = main.NOQ
End Sub
