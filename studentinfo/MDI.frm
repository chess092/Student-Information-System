VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Mail us at-pkdag@indiatimes.com"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu sis 
      Caption         =   "Student_ Info_ System"
      Index           =   0
      Begin VB.Menu open 
         Caption         =   "Open SIS"
         Index           =   1
         Shortcut        =   ^O
      End
      Begin VB.Menu pass 
         Caption         =   "Change Password"
         Index           =   3
         Shortcut        =   ^C
      End
      Begin VB.Menu update 
         Caption         =   "Update Year"
         Index           =   4
         Shortcut        =   ^U
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
         Index           =   2
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu special 
      Caption         =   "Special"
      Index           =   5
      Begin VB.Menu calendar 
         Caption         =   "Calendar"
         Index           =   52
         Shortcut        =   {F1}
      End
      Begin VB.Menu cal 
         Caption         =   "Calculator"
         Index           =   51
         Shortcut        =   {F2}
      End
   End
   Begin VB.Menu about 
      Caption         =   "About"
      Index           =   6
      Begin VB.Menu cred 
         Caption         =   "Credits"
         Index           =   62
         Shortcut        =   {F5}
      End
      Begin VB.Menu help 
         Caption         =   "Help"
         Index           =   61
         Shortcut        =   {F6}
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cal_Click(Index As Integer)
Call Shell(App.Path & "\calc.exe", vbNormalFocus)

End Sub

Private Sub calendar_Click(Index As Integer)
Form3.Show
End Sub

Private Sub cred_Click(Index As Integer)
Form5.Show
End Sub

Private Sub exit_Click(Index As Integer)
End
End Sub

Private Sub help_Click(Index As Integer)
Form7.Show
End Sub

Private Sub MDIForm_Terminate()
End
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
End
End Sub

Private Sub open_Click(Index As Integer)
If Index = 1 Then

Form2.Show
End If
End Sub

Private Sub pass_Click(Index As Integer)
If Index = 3 Then
password.Show
End If
End Sub

Private Sub update_Click(Index As Integer)
Form4.Show
End Sub
