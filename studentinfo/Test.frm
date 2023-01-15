VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form6 
   Caption         =   "Form6"
   ClientHeight    =   4890
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5760
   LinkTopic       =   "Form6"
   ScaleHeight     =   4890
   ScaleWidth      =   5760
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2280
      Top             =   4200
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   4920
      TabIndex        =   21
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox Text21 
      Height          =   285
      Left            =   3600
      TabIndex        =   20
      Text            =   "overall"
      Top             =   4200
      Width           =   855
   End
   Begin VB.TextBox Text20 
      Height          =   285
      Left            =   3600
      TabIndex        =   19
      Text            =   "sem tot"
      Top             =   3840
      Width           =   855
   End
   Begin VB.TextBox Text19 
      Height          =   285
      Left            =   3600
      TabIndex        =   18
      Text            =   "s8"
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox Text18 
      Height          =   285
      Left            =   3600
      TabIndex        =   17
      Text            =   "s7"
      Top             =   3120
      Width           =   855
   End
   Begin VB.TextBox Text17 
      Height          =   285
      Left            =   3600
      TabIndex        =   16
      Text            =   "s6"
      Top             =   2760
      Width           =   855
   End
   Begin VB.TextBox Text16 
      Height          =   285
      Left            =   3480
      TabIndex        =   15
      Text            =   "s5"
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox Text15 
      Height          =   285
      Left            =   3480
      TabIndex        =   14
      Text            =   "s4"
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox Text14 
      Height          =   285
      Left            =   3480
      TabIndex        =   13
      Text            =   "s3"
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox Text13 
      Height          =   285
      Left            =   3480
      TabIndex        =   12
      Text            =   "s2"
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox Text12 
      Height          =   285
      Left            =   3480
      TabIndex        =   11
      Text            =   "s1"
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox Text11 
      Height          =   285
      Left            =   3360
      TabIndex        =   10
      Text            =   "12"
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   3360
      TabIndex        =   9
      Text            =   "10"
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   1080
      TabIndex        =   8
      Text            =   "email"
      Top             =   4320
      Width           =   735
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   1080
      TabIndex        =   7
      Text            =   "sex"
      Top             =   3840
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   1080
      TabIndex        =   6
      Text            =   "dob"
      Top             =   3360
      Width           =   855
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Text            =   "roll"
      Top             =   2760
      Width           =   1455
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   960
      TabIndex        =   4
      Text            =   "yr"
      Top             =   2280
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Text            =   "br"
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   960
      TabIndex        =   2
      Text            =   "ph"
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Text            =   "addr"
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Text            =   "name"
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs5 As Recordset
Dim con5 As Connection

Private Sub Command1_Click()
Set rs5 = New Recordset
Set con5 = New Connection
con5.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\WINDOWS\Desktop\StudentInfo\student.mdb;Persist Security Info=False"
rs5.open "select * from studac", con5, adOpenDynamic, adLockOptimistic
For i = 1 To 70
rs5.AddNew Array("name", "addr", "phone", "branch", "year", "roll", "sex", "class10", "class12", "sem1", "sem2", "sem3", "sem4", "sem5", "sem6", "sem7", "sem8", "DoB", "Email", "semstotal", "overalltotal"), Array(Text1.Text, Text2.Text, Val(Text3.Text) + i, Text4.Text, Text5.Text, Val(Text6.Text) + i, Text8.Text, Text10.Text, Text11.Text, Val(Text12.Text) + i, Val(Text13.Text) + i, Val(Text14.Text) + i, Val(Text15.Text) + i, Val(Text16.Text) + i, Val(Text17.Text) + i, Val(Text18.Text) + i, Val(Text19.Text) + i, Text7.Text, Text9.Text, Val(Text12.Text) + i + Val(Text13.Text) + i + Val(Text14.Text) + i + Val(Text15.Text) + i + Val(Text16.Text) + i + Val(Text17.Text) + i + Val(Text18.Text) + i + Val(Text19.Text) + i, Val(Text12.Text) + i + Val(Text13.Text) + i + Val(Text14.Text) + i + Val(Text15.Text) + i + Val(Text16.Text) + i + Val(Text17.Text) + i + Val(Text18.Text) + i + Val(Text19.Text) + i + Val(Text10.Text) + Val(Text11.Text))
rs5.MoveNext
Next
MsgBox "OK"
End Sub

