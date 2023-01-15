VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Form2"
   ClientHeight    =   6360
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10140
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   6360
   ScaleWidth      =   10140
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   8760
      Top             =   840
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "E&xit"
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
      Left            =   6240
      MaskColor       =   &H00008000&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5640
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Enter"
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
      Left            =   4320
      MaskColor       =   &H00008000&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5640
      Width           =   1455
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   6480
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   4320
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3960
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Height          =   1575
      Left            =   3360
      TabIndex        =   3
      Top             =   3600
      Width           =   5175
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Select Year"
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
         Left            =   3120
         TabIndex        =   5
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Select Branch"
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
         Left            =   600
         TabIndex        =   4
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFC0C0&
      Height          =   2775
      Left            =   2760
      TabIndex        =   31
      Top             =   0
      Width           =   6615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   23
      Left            =   7080
      TabIndex        =   30
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   22
      Left            =   6600
      TabIndex        =   29
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   21
      Left            =   6120
      TabIndex        =   28
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   20
      Left            =   5640
      TabIndex        =   27
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   19
      Left            =   5160
      TabIndex        =   26
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   18
      Left            =   4680
      TabIndex        =   25
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   17
      Left            =   8040
      TabIndex        =   24
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   16
      Left            =   7440
      TabIndex        =   23
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   15
      Left            =   7200
      TabIndex        =   22
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   14
      Left            =   6720
      TabIndex        =   21
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   13
      Left            =   6240
      TabIndex        =   20
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   12
      Left            =   5760
      TabIndex        =   19
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   11
      Left            =   5280
      TabIndex        =   18
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   10
      Left            =   4800
      TabIndex        =   17
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   9
      Left            =   4320
      TabIndex        =   16
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   8
      Left            =   3840
      TabIndex        =   15
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   7
      Left            =   3480
      TabIndex        =   14
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   6
      Left            =   7320
      TabIndex        =   13
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   5
      Left            =   6840
      TabIndex        =   12
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   4
      Left            =   6360
      TabIndex        =   11
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   3
      Left            =   5880
      TabIndex        =   10
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   5400
      TabIndex        =   9
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   4920
      TabIndex        =   8
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   4440
      TabIndex        =   7
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()








'Call txt0

If Combo1.ListIndex < 0 Or Combo2.ListIndex < 0 Then
MsgBox "Please select Branch & Year"
Form1.Hide
Else
Select Case Form2.Combo1.Text
 Case "IT":
  Select Case Form2.Combo2.Text
   Case "1st":
    Form1.Data1.RecordSource = "select name,addr,dob,roll,phone,email,branch,year,sex from studac where branch='IT' and year='1st' order by roll"
    Form1.Data2.RecordSource = "select name,roll,class10,class12,sem1,sem2 from studac where branch='IT' and year='1st' order by roll "
    Case "2nd":
    Form1.Data1.RecordSource = "select name,addr,dob,roll,phone,email,branch,year,sex from studac where branch='IT' and year='2nd' order by roll "
    Form1.Data2.RecordSource = "select name,roll,class10,class12,sem1,sem2,sem3,sem4 from studac where branch='IT' and year='2nd'order by roll"
    Case "3rd":
    Form1.Data1.RecordSource = "select name,addr,dob,roll,phone,email,branch,year,sex from studac where branch='IT' and year='3rd' order by roll "
    Form1.Data2.RecordSource = "select name,roll,class10,class12,sem1,sem2,sem3,sem4,sem5,sem6 from studac where branch='IT' and year='3rd' order by roll"
    Case "4th":
    Form1.Data1.RecordSource = "select name,addr,dob,roll,phone,email,branch,year,sex from studac where branch='IT' and year='4th' order by roll "
    Form1.Data2.RecordSource = "select name,roll,class10,class12,sem1,sem2,sem3,sem4,sem5,sem6,sem7,sem8 from studac where branch='IT' and year='4th' order by roll"
  End Select
  Case "CSE":
  Select Case Form2.Combo2.Text
   Case "1st":
    Form1.Data1.RecordSource = "select name,addr,dob,roll,phone,email,branch,year,sex from studac where branch='CSE' and year='1st'  order by roll"
    Form1.Data2.RecordSource = "select name,roll,class10,class12,sem1,sem2 from studac where branch='CSE' and year='1st' order by roll "
    Case "2nd":
    Form1.Data1.RecordSource = "select name,addr,dob,roll,phone,email,branch,year,sex from studac where branch='CSE' and year='2nd' order by roll "
    Form1.Data2.RecordSource = "select name,roll,class10,class12,sem1,sem2,sem3,sem4 from studac where branch='CSE' and year='2nd' order by roll "
    Case "3rd":
    Form1.Data1.RecordSource = "select name,addr,dob,roll,phone,email,branch,year,sex from studac where branch='CSE' and year='3rd' order by roll "
    Form1.Data2.RecordSource = "select name,roll,class10,class12,sem1,sem2,sem3,sem4,sem5,sem6 from studac where branch='CSE' and year='3rd' order by roll "
    Case "4th":
    Form1.Data1.RecordSource = "select name,addr,dob,roll,phone,email,branch,year,sex from studac where branch='CSE' and year='4th' order by roll"
    Form1.Data2.RecordSource = "select name,roll,class10,class12,sem1,sem2,sem3,sem4,sem5,sem6,sem7,sem8 from studac where branch='CSE' and year='4th'  order by roll"
  End Select
  Case "ECE":
  Select Case Form2.Combo2.Text
   Case "1st":
    Form1.Data1.RecordSource = "select name,addr,dob,roll,phone,email,branch,year,sex from studac where branch='ECE' and year='1st'  order by roll"
     Form1.Data2.RecordSource = "select name,roll,class10,class12,sem1,sem2 from studac where branch='ECE' and year='1st'  order by roll"
    Case "2nd":
    Form1.Data1.RecordSource = "select name,addr,dob,roll,phone,email,branch,year,sex from studac where branch='ECE' and year='2nd'  order by roll"
    Form1.Data2.RecordSource = "select name,roll,class10,class12,sem1,sem2,sem3,sem4 from studac where branch='ECE' and year='2nd'  order by roll"
    Case "3rd":
    Form1.Data1.RecordSource = "select name,addr,dob,roll,phone,email,branch,year,sex from studac where branch='ECE' and year='3rd'  order by roll"
    Form1.Data2.RecordSource = "select name,roll,class10,class12,sem1,sem2,sem3,sem4,sem5,sem6 from studac where branch='ECE' and year='3rd'  order by roll"
    Case "4th":
    Form1.Data1.RecordSource = "select name,addr,dob,roll,phone,email,branch,year,sex from studac where branch='ECE' and year='4th'  order by roll"
    Form1.Data2.RecordSource = "select name,roll,class10,class12,sem1,sem2,sem3,sem4,sem5,sem6,sem7,sem8 from studac where branch='ECE' and year='4th' order by roll "
  End Select
End Select
Form1.Data1.Refresh
Form1.Data2.Refresh


Form1.Command3.Enabled = False
Form1.Command4.Enabled = False
Form1.Command10.Enabled = False
Form1.Command7.Enabled = 0
Form1.Command8.Enabled = 0
Form1.Command1.Enabled = 1
Form1.Command2.Enabled = 1
Form1.Command9.Enabled = 1
'------
        Form1.Text9.BackColor = vbWhite
        Form1.Text10.BackColor = vbWhite
        Form1.Text11.BackColor = vbWhite
        Form1.Text12.BackColor = vbWhite
        Form1.Text13.BackColor = vbWhite
        Form1.Text14.BackColor = vbWhite
        Form1.Text15.BackColor = vbWhite
        Form1.Text16.BackColor = vbWhite


'------
Form1.Text1.Text = ""
Form1.Text2.Text = ""
Form1.Text3.Text = ""
Form1.Text6.Text = ""
Form1.Text7.Text = ""
Form1.Text8.Text = ""
Form1.Text9.Text = ""
Form1.Text10.Text = ""
Form1.Text11.Text = ""
Form1.Text12.Text = ""
Form1.Text13.Text = ""
Form1.Text14.Text = ""
Form1.Text15.Text = ""
Form1.Text16.Text = ""
Form1.Text20.Text = ""
Form1.Text21.Text = ""
Form1.Text22.Text = ""
Form1.Text23.Text = ""
'*********
Form1.Label2.Caption = "Branch-" & Form2.Combo1.Text & " # Year-" & Form2.Combo2.Text
Form1.Label22.Caption = "Branch-" & Form2.Combo1.Text & " # Year-" & Form2.Combo2.Text
Form1.Label44.Caption = "Branch-" & Form2.Combo1.Text & " # Year-" & Form2.Combo2.Text

Form1.DBGrid3.Visible = False


'**********
Me.Hide
Form1.Text4.Enabled = False
Form1.Text5.Enabled = False
Select Case Combo2.ListIndex
Case 0: Form1.Text7.Enabled = 1
        Form1.Text8.Enabled = 1
        Form1.Text9.Enabled = 1
        Form1.Text10.Enabled = 1
        Form1.Text11.Enabled = False
        Form1.Text12.Enabled = False
        Form1.Text13.Enabled = False
        Form1.Text14.Enabled = False
        Form1.Text15.Enabled = False
        Form1.Text16.Enabled = False
        Form1.Text11.BackColor = vbYellow
        Form1.Text12.BackColor = vbYellow
        Form1.Text13.BackColor = vbYellow
        Form1.Text14.BackColor = vbYellow
        Form1.Text15.BackColor = vbYellow
        Form1.Text16.BackColor = vbYellow
        Form1.Text11.Text = "0"
        Form1.Text12.Text = "0"
        Form1.Text13.Text = "0"
        Form1.Text14.Text = "0"
        Form1.Text15.Text = "0"
        Form1.Text16.Text = "0"
 Case 1:
        Form1.Text7.Enabled = 1
        Form1.Text8.Enabled = 1
        Form1.Text9.Enabled = 1
        Form1.Text10.Enabled = 1
        Form1.Text11.Enabled = 1
        Form1.Text12.Enabled = 1
        Form1.Text13.Enabled = False
        Form1.Text14.Enabled = False
        Form1.Text15.Enabled = False
        Form1.Text16.Enabled = False
        Form1.Text13.BackColor = vbYellow
        Form1.Text14.BackColor = vbYellow
        Form1.Text15.BackColor = vbYellow
        Form1.Text16.BackColor = vbYellow
        Form1.Text11.BackColor = vbWhite
        Form1.Text12.BackColor = vbWhite
        Form1.Text13.Text = "0"
        Form1.Text14.Text = "0"
        Form1.Text15.Text = "0"
        Form1.Text16.Text = "0"
 Case 2:
        Form1.Text7.Enabled = 1
        Form1.Text8.Enabled = 1
        Form1.Text9.Enabled = 1
        Form1.Text10.Enabled = 1
        Form1.Text11.Enabled = 1
        Form1.Text12.Enabled = 1
        Form1.Text13.Enabled = 1
        Form1.Text14.Enabled = 1
        Form1.Text15.Enabled = False
        Form1.Text16.Enabled = False
        Form1.Text13.BackColor = vbWhite
        Form1.Text14.BackColor = vbWhite
        Form1.Text15.BackColor = vbYellow
        Form1.Text16.BackColor = vbYellow
        Form1.Text11.BackColor = vbWhite
        Form1.Text12.BackColor = vbWhite
        Form1.Text15.Text = "0"
        Form1.Text16.Text = "0"
    Case 3:
          Form1.Text7.Enabled = 1
        Form1.Text8.Enabled = 1
        Form1.Text9.Enabled = 1
        Form1.Text10.Enabled = 1
        Form1.Text11.Enabled = 1
        Form1.Text12.Enabled = 1
        Form1.Text13.Enabled = 1
        Form1.Text14.Enabled = 1
        Form1.Text15.Enabled = 1
        Form1.Text16.Enabled = 1
        Form1.Text15.BackColor = vbWhite
        Form1.Text16.BackColor = vbWhite
End Select
Form1.Text4.Text = Form2.Combo1.List(Form2.Combo1.ListIndex)
Form1.Text5.Text = Form2.Combo2.List(Form2.Combo2.ListIndex)

Form1.Show
Form1.WindowState = 2
End If

End Sub

Private Sub Command2_Click()
'End
Form1.Hide
Form2.Hide
MDIForm1.Enabled = True
End Sub

Private Sub Form_Load()
Label4.Visible = False
Form1.Data1.DatabaseName = App.Path & "\student.mdb"
Form1.Data2.DatabaseName = App.Path & "\student.mdb"
Form1.Data3.DatabaseName = App.Path & "\student.mdb"
sis = "STUDENTINFORMATIONSYSTEM"
For i = 0 To 23
Label3(i).Caption = Mid(sis, i + 1, 1)
Next
Combo1.AddItem "IT"
Combo1.AddItem "CSE"
Combo1.AddItem "ECE"

Combo2.AddItem "1st"
Combo2.AddItem "2nd"
Combo2.AddItem "3rd"
Combo2.AddItem "4th"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form1.Hide
End Sub

Private Sub Timer1_Timer()
For i = 0 To 23
Col = RGB(225 * Rnd + i, 225 * Rnd + i, Rnd * 225 + i)
Label3(i).ForeColor = Col
Next
Timer1.Enabled = True
End Sub
