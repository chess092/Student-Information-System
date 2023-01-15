VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Update year"
   ClientHeight    =   2625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5490
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   5490
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   3
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1 year UP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2520
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Year updation form"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   240
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "Enter master password:"
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
      TabIndex        =   1
      Top             =   1080
      Width           =   2055
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs3 As Recordset
Dim con3 As Connection
Dim rs4 As Recordset
Dim con4 As Connection

Private Sub Command1_Click()
If MsgBox("Are you sure?", vbYesNo) = vbYes Then
Set rs4 = New Recordset
Set con4 = New Connection
con4.open "Provider=Microsoft.Jet.OLEDB.4.0;User ID=admin;Data Source=" & App.Path & "\student.mdb;Persist Security Info=False"
rs4.open "select year from studac", con4, adOpenDynamic, adLockOptimistic
Do While Not rs4.EOF
  Select Case rs4!Year
  Case "1st":

   rs4.update Array("year"), Array("2nd")
  Case "2nd":

   rs4.update Array("year"), Array("3rd")
  Case "3rd":

   rs4.update Array("year"), Array("4th")
  Case "4th":
  
  rs4.update Array("year"), Array("Ex")
  End Select
rs4.MoveNext
Loop
MsgBox "Successfully increased by 1 year"
Form2.Refresh
Form1.Refresh
Else

End If
Text1.Text = ""
Command1.Enabled = False
Me.Hide
End Sub

Private Sub Command3_Click()
Text1.Text = ""
Command1.Enabled = False
Me.Hide
End Sub

Private Sub Form_Load()
Set rs3 = New Recordset
Set con3 = New Connection
con3.open "Provider=Microsoft.Jet.OLEDB.4.0;User ID=admin;Data Source=" & App.Path & "\student.mdb;Persist Security Info=False"
rs3.open "select * from passw", con3, adOpenDynamic, adLockOptimistic
rs3.MoveLast
Command1.Enabled = 0
'Command2.Enabled = 0


End Sub

Private Sub Text1_Change()
If Text1.Text = rs3!pass Then
Command1.Enabled = True
'Command2.Enabled = True

End If
End Sub
