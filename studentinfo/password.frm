VERSION 5.00
Begin VB.Form password 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change password"
   ClientHeight    =   3180
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1878.849
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1290
      TabIndex        =   1
      Top             =   615
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   360
      TabIndex        =   4
      Top             =   2640
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2160
      TabIndex        =   5
      Top             =   2640
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1005
      Width           =   2325
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Change password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "Reenter new password:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Enter new password:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label lblLabels 
      Caption         =   "User name:"
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   630
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "Old password:"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   1020
      Width           =   1080
   End
End
Attribute VB_Name = "password"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs3 As Recordset
Dim con3 As Connection
Option Explicit

Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    'set the global var to false
    'to denote a failed login
    Text1.Text = ""
    Text2.Text = ""
    txtUserName.Text = ""
    txtPassword = ""
    LoginSucceeded = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    'check for correct password
    If txtUserName.Text = rs3!uname And txtPassword = rs3!pass Then
        'place code to here to pass the
        'success to the calling sub
        'setting a global var is the easiest
        If Text1.Text = Text2.Text And Text2.Text <> "" Then
        rs3.Update Array("pass"), Array(Text2.Text)
        LoginSucceeded = True
        MsgBox "Your password has been changed successfully"
        Me.Hide
        Else
        MsgBox "Mistake in new password! Retype"
        Text1.SetFocus
        End If
        
    Else
        MsgBox "Invalid Username or Password, try again!", , "Login"
        txtUserName.SetFocus
        SendKeys "{Home}+{End}"
    End If
End Sub

Private Sub Form_Load()
Set rs3 = New Recordset
Set con3 = New Connection
con3.Open "Provider=Microsoft.Jet.OLEDB.4.0;User ID=admin;Data Source=E:\vbcodes\Student\stud2\student.mdb;Persist Security Info=False"
rs3.Open "select * from passw", con3, adOpenDynamic, adLockOptimistic

End Sub
