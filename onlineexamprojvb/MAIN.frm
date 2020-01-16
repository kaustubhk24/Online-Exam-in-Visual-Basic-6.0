VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00C0FFFF&
   Caption         =   "ONLINE-EXAMINATION"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4680
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu MN_MAIN 
      Caption         =   "&MAIN"
      Begin VB.Menu MN_FAC 
         Caption         =   "&FACULTY"
         Begin VB.Menu MNU_MOD 
            Caption         =   "&MODULE_INFORMATION"
            Index           =   1
         End
         Begin VB.Menu MNU_QPAPER 
            Caption         =   "&QPAPER_INFORMATION"
            Index           =   2
         End
      End
      Begin VB.Menu MNUSEP 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu MN_ADMIN 
         Caption         =   "&ADMINISTRATOR"
      End
   End
   Begin VB.Menu MN_EXAM 
      Caption         =   "&EXAMINATION"
   End
   Begin VB.Menu MN_REPO 
      Caption         =   "&REPORTS"
      Begin VB.Menu MOD_REP 
         Caption         =   "&MODULE_REPORTS"
         Index           =   1
      End
      Begin VB.Menu MNU_SAP2 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu MNU_DAILY 
         Caption         =   "&USER_REPORTS"
      End
   End
   Begin VB.Menu MNU_EXIT 
      Caption         =   "E&XIT"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MN_ADMIN_Click()
ADMIN_LOG.Show
End Sub

Private Sub MN_EXAM_Click()
 INSTRUCTION.Show
End Sub

Private Sub MNU_DAILY_Click()
DataReport3.Show
End Sub

Private Sub MNU_EXIT_Click()
End
End Sub

Private Sub MNU_MOD_Click(Index As Integer)
ADMIN_MOD.Show
End Sub

Private Sub MNU_QPAPER_Click(Index As Integer)
ADMIN_QPAPER.Show
End Sub

Private Sub MNU_TOT_Click()
DataReport1.Show
End Sub

Private Sub MOD_REP_Click(Index As Integer)
DataReport2.Show
End Sub
