VERSION 5.00
Begin VB.Form ADMIN 
   Caption         =   "                                                               ADMINISTRATION--FORM"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   6795
   ScaleWidth      =   9480
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdexit 
      Caption         =   "EXIT"
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
      Left            =   7320
      TabIndex        =   12
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CommandButton Command 
      Caption         =   "SAVE"
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
      Left            =   5880
      TabIndex        =   11
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "INSERT"
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
      Left            =   4440
      TabIndex        =   10
      Top             =   5160
      Width           =   1455
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
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   6360
      TabIndex        =   5
      Text            =   " "
      Top             =   3360
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   6360
      TabIndex        =   4
      Text            =   " "
      Top             =   2520
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   6360
      TabIndex        =   3
      Top             =   1680
      Width           =   2175
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
      Left            =   3960
      TabIndex        =   2
      Top             =   3480
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
      Left            =   3960
      TabIndex        =   1
      Top             =   2640
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
      Left            =   3960
      TabIndex        =   0
      Top             =   1680
      Width           =   1215
   End
End
Attribute VB_Name = "ADMIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
