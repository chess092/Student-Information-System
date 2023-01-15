VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8025
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11790
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8025
   ScaleWidth      =   11790
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   29
      Top             =   7620
      Width           =   11790
      _ExtentX        =   20796
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   1
            Object.Width           =   10160
            TextSave        =   "8/24/02"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   1
            Object.Width           =   10160
            TextSave        =   "9:33 PM"
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6855
      Left            =   240
      TabIndex        =   30
      Top             =   720
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   12091
      _Version        =   393216
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "ADD/MODIFY"
      TabPicture(0)   =   "main.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Text23"
      Tab(0).Control(1)=   "Text22"
      Tab(0).Control(2)=   "Text21"
      Tab(0).Control(3)=   "Text20"
      Tab(0).Control(4)=   "Command10"
      Tab(0).Control(5)=   "Command9"
      Tab(0).Control(6)=   "Command8"
      Tab(0).Control(7)=   "Command7"
      Tab(0).Control(8)=   "Command2"
      Tab(0).Control(9)=   "Command6"
      Tab(0).Control(10)=   "Command5"
      Tab(0).Control(11)=   "Command4"
      Tab(0).Control(12)=   "Command3"
      Tab(0).Control(13)=   "Command1"
      Tab(0).Control(14)=   "Text1"
      Tab(0).Control(15)=   "Text2"
      Tab(0).Control(16)=   "Text3"
      Tab(0).Control(17)=   "Text4"
      Tab(0).Control(18)=   "Text5"
      Tab(0).Control(19)=   "Text6"
      Tab(0).Control(20)=   "List1"
      Tab(0).Control(21)=   "Text7"
      Tab(0).Control(22)=   "Text8"
      Tab(0).Control(23)=   "Text9"
      Tab(0).Control(24)=   "Text10"
      Tab(0).Control(25)=   "Text11"
      Tab(0).Control(26)=   "Text12"
      Tab(0).Control(27)=   "Text13"
      Tab(0).Control(28)=   "Text14"
      Tab(0).Control(29)=   "Text16"
      Tab(0).Control(30)=   "Text15"
      Tab(0).Control(31)=   "Label51"
      Tab(0).Control(32)=   "Label50"
      Tab(0).Control(33)=   "Label49"
      Tab(0).Control(34)=   "Label48"
      Tab(0).Control(35)=   "Label47"
      Tab(0).Control(36)=   "Label3"
      Tab(0).Control(37)=   "Label4"
      Tab(0).Control(38)=   "Label5"
      Tab(0).Control(39)=   "Label6"
      Tab(0).Control(40)=   "Label7"
      Tab(0).Control(41)=   "Label8"
      Tab(0).Control(42)=   "Label9"
      Tab(0).Control(43)=   "Label10"
      Tab(0).Control(44)=   "Label11"
      Tab(0).Control(45)=   "Label12"
      Tab(0).Control(46)=   "Label13"
      Tab(0).Control(47)=   "Label14"
      Tab(0).Control(48)=   "Label15"
      Tab(0).Control(49)=   "Label16"
      Tab(0).Control(50)=   "Label17"
      Tab(0).Control(51)=   "Label18"
      Tab(0).Control(52)=   "Label19"
      Tab(0).ControlCount=   53
      TabCaption(1)   =   "SEARCH"
      TabPicture(1)   =   "main.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label23"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label24"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label25"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label26"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label27"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label28"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label29"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label30"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label31"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label32"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label33"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label34"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Label35"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Label36"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Label37"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Label38"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Label39"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Label40"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Label41"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Label42"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "Label43"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Label44"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "Label20"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "Label21"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "Label52"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "Label53"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "Label54"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "Label55"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "DBGrid3"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "Frame1"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "List4"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "Check1"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "Check2"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "Check3"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "Check4"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "Check5"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).Control(36)=   "Check6"
      Tab(1).Control(36).Enabled=   0   'False
      Tab(1).Control(37)=   "Check7"
      Tab(1).Control(37).Enabled=   0   'False
      Tab(1).Control(38)=   "Check8"
      Tab(1).Control(38).Enabled=   0   'False
      Tab(1).Control(39)=   "Text17"
      Tab(1).Control(39).Enabled=   0   'False
      Tab(1).Control(40)=   "Text18"
      Tab(1).Control(40).Enabled=   0   'False
      Tab(1).Control(41)=   "Command15"
      Tab(1).Control(41).Enabled=   0   'False
      Tab(1).Control(42)=   "Check9"
      Tab(1).Control(42).Enabled=   0   'False
      Tab(1).Control(43)=   "Check10"
      Tab(1).Control(43).Enabled=   0   'False
      Tab(1).Control(44)=   "Check11"
      Tab(1).Control(44).Enabled=   0   'False
      Tab(1).Control(45)=   "Check12"
      Tab(1).Control(45).Enabled=   0   'False
      Tab(1).Control(46)=   "Check13"
      Tab(1).Control(46).Enabled=   0   'False
      Tab(1).Control(47)=   "Check14"
      Tab(1).Control(47).Enabled=   0   'False
      Tab(1).Control(48)=   "Check15"
      Tab(1).Control(48).Enabled=   0   'False
      Tab(1).Control(49)=   "Check16"
      Tab(1).Control(49).Enabled=   0   'False
      Tab(1).Control(50)=   "Text19"
      Tab(1).Control(50).Enabled=   0   'False
      Tab(1).Control(51)=   "Data3"
      Tab(1).Control(51).Enabled=   0   'False
      Tab(1).Control(52)=   "Check17"
      Tab(1).Control(52).Enabled=   0   'False
      Tab(1).Control(53)=   "Command16"
      Tab(1).Control(53).Enabled=   0   'False
      Tab(1).Control(54)=   "Command17"
      Tab(1).Control(54).Enabled=   0   'False
      Tab(1).Control(55)=   "Option1"
      Tab(1).Control(55).Enabled=   0   'False
      Tab(1).Control(56)=   "Option2"
      Tab(1).Control(56).Enabled=   0   'False
      Tab(1).Control(57)=   "Check18"
      Tab(1).Control(57).Enabled=   0   'False
      Tab(1).Control(58)=   "Check19"
      Tab(1).Control(58).Enabled=   0   'False
      Tab(1).Control(59)=   "Check20"
      Tab(1).Control(59).Enabled=   0   'False
      Tab(1).Control(60)=   "Check21"
      Tab(1).Control(60).Enabled=   0   'False
      Tab(1).Control(61)=   "Check22"
      Tab(1).Control(61).Enabled=   0   'False
      Tab(1).Control(62)=   "Check23"
      Tab(1).Control(62).Enabled=   0   'False
      Tab(1).Control(63)=   "Frame2"
      Tab(1).Control(63).Enabled=   0   'False
      Tab(1).ControlCount=   64
      TabCaption(2)   =   "ACADEMIC INFO."
      TabPicture(2)   =   "main.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label22"
      Tab(2).Control(1)=   "DBGrid2"
      Tab(2).Control(2)=   "Data2"
      Tab(2).Control(3)=   "Command13"
      Tab(2).Control(4)=   "Command14"
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "NON-ACADEMIC INFO."
      TabPicture(3)   =   "main.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Command12"
      Tab(3).Control(1)=   "Command11"
      Tab(3).Control(2)=   "Data1"
      Tab(3).Control(3)=   "DBGrid1"
      Tab(3).Control(4)=   "Label2"
      Tab(3).ControlCount=   5
      Begin VB.Frame Frame2 
         Caption         =   "Database"
         Height          =   975
         Left            =   3720
         TabIndex        =   128
         Top             =   1560
         Width           =   2055
         Begin VB.CheckBox Check25 
            Caption         =   "Check25"
            Height          =   195
            Left            =   120
            TabIndex        =   130
            Top             =   600
            Width           =   255
         End
         Begin VB.CheckBox Check24 
            Caption         =   "Check24"
            Height          =   195
            Left            =   120
            TabIndex        =   129
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label57 
            Caption         =   "All students"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   480
            TabIndex        =   132
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label56 
            Caption         =   "Current branch-year."
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   360
            TabIndex        =   131
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.CheckBox Check23 
         Caption         =   "Check23"
         Height          =   195
         Left            =   10080
         TabIndex        =   122
         Top             =   1320
         Width           =   255
      End
      Begin VB.CheckBox Check22 
         Caption         =   "Check22"
         Height          =   255
         Left            =   10080
         TabIndex        =   121
         Top             =   960
         Width           =   255
      End
      Begin VB.CheckBox Check21 
         Caption         =   "Check21"
         Height          =   195
         Left            =   10080
         TabIndex        =   120
         Top             =   600
         Width           =   255
      End
      Begin VB.CheckBox Check20 
         Caption         =   "Check20"
         Height          =   255
         Left            =   9240
         TabIndex        =   119
         Top             =   1320
         Width           =   255
      End
      Begin VB.TextBox Text23 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -67080
         TabIndex        =   20
         Top             =   5880
         Width           =   1455
      End
      Begin VB.TextBox Text22 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -67080
         TabIndex        =   19
         Top             =   5370
         Width           =   1455
      End
      Begin VB.TextBox Text21 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -73680
         TabIndex        =   8
         Top             =   5040
         Width           =   2775
      End
      Begin VB.TextBox Text20 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -73680
         TabIndex        =   7
         Top             =   3960
         Width           =   1095
      End
      Begin VB.CheckBox Check19 
         Caption         =   "Check19"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   8400
         TabIndex        =   110
         Top             =   1920
         Width           =   255
      End
      Begin VB.CheckBox Check18 
         Caption         =   "Check18"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   6240
         TabIndex        =   109
         Top             =   1920
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Option2"
         Height          =   255
         Left            =   2040
         TabIndex        =   106
         Top             =   2640
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   255
         Left            =   240
         TabIndex        =   105
         Top             =   2640
         Width           =   255
      End
      Begin VB.CommandButton Command17 
         BackColor       =   &H00FFFFC0&
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
         Left            =   10320
         TabIndex        =   103
         Top             =   2400
         Width           =   855
      End
      Begin VB.CommandButton Command16 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Back"
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
         Left            =   10320
         TabIndex        =   102
         Top             =   1800
         Width           =   855
      End
      Begin VB.CheckBox Check17 
         Caption         =   "Check17"
         Height          =   255
         Left            =   8040
         TabIndex        =   100
         Top             =   1320
         Width           =   255
      End
      Begin VB.Data Data3 
         Caption         =   "Data3"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   2160
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   4800
         Width           =   2295
      End
      Begin VB.TextBox Text19 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   96
         Top             =   2160
         Width           =   1815
      End
      Begin VB.CheckBox Check16 
         Caption         =   "Check16"
         Height          =   195
         Left            =   6960
         TabIndex        =   87
         Top             =   1320
         Width           =   255
      End
      Begin VB.CheckBox Check15 
         Caption         =   "Check15"
         Height          =   195
         Left            =   5880
         TabIndex        =   86
         Top             =   1320
         Width           =   255
      End
      Begin VB.CheckBox Check14 
         Caption         =   "Check14"
         Height          =   255
         Left            =   4680
         TabIndex        =   85
         Top             =   1320
         Width           =   255
      End
      Begin VB.CheckBox Check13 
         Caption         =   "Check13"
         Height          =   255
         Left            =   3480
         TabIndex        =   84
         Top             =   1320
         Width           =   255
      End
      Begin VB.CheckBox Check12 
         Caption         =   "Check12"
         Height          =   255
         Left            =   9240
         TabIndex        =   83
         Top             =   960
         Width           =   255
      End
      Begin VB.CheckBox Check11 
         Caption         =   "Check11"
         Height          =   255
         Left            =   8040
         TabIndex        =   82
         Top             =   960
         Width           =   255
      End
      Begin VB.CheckBox Check10 
         Caption         =   "Check10"
         Height          =   255
         Left            =   6960
         TabIndex        =   81
         Top             =   960
         Width           =   255
      End
      Begin VB.CheckBox Check9 
         Caption         =   "Check9"
         Height          =   255
         Left            =   5880
         TabIndex        =   80
         Top             =   960
         Width           =   255
      End
      Begin VB.CommandButton Command15 
         Caption         =   "SEARCH"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   79
         Top             =   2640
         Width           =   1095
      End
      Begin VB.TextBox Text18 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   78
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox Text17 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   77
         Top             =   2160
         Width           =   975
      End
      Begin VB.CheckBox Check8 
         Caption         =   "Check8"
         Height          =   255
         Left            =   4680
         TabIndex        =   75
         Top             =   960
         Width           =   255
      End
      Begin VB.CheckBox Check7 
         Caption         =   "Check7"
         Height          =   255
         Left            =   3480
         TabIndex        =   73
         Top             =   960
         Width           =   255
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Check6"
         Height          =   255
         Left            =   9240
         TabIndex        =   66
         Top             =   600
         Width           =   255
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Check5"
         Height          =   255
         Left            =   8040
         TabIndex        =   65
         Top             =   600
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Check4"
         Height          =   255
         Left            =   6960
         TabIndex        =   64
         Top             =   600
         Width           =   255
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Check3"
         Height          =   255
         Left            =   5880
         TabIndex        =   63
         Top             =   600
         Width           =   255
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Check2"
         Height          =   255
         Left            =   4680
         TabIndex        =   62
         Top             =   600
         Width           =   255
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Left            =   3480
         TabIndex        =   61
         Top             =   600
         Width           =   255
      End
      Begin VB.ListBox List4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1020
         Left            =   360
         TabIndex        =   59
         Top             =   720
         Width           =   2895
      End
      Begin VB.CommandButton Command14 
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
         Height          =   375
         Left            =   -65280
         TabIndex        =   58
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Back"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -66600
         TabIndex        =   57
         Top             =   840
         Width           =   1095
      End
      Begin VB.Data Data2 
         Caption         =   "Data2"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   -74520
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   960
         Width           =   2655
      End
      Begin MSDBGrid.DBGrid DBGrid2 
         Bindings        =   "main.frx":0070
         Height          =   5175
         Left            =   -74880
         OleObjectBlob   =   "main.frx":0084
         TabIndex        =   56
         Top             =   1560
         Width           =   11295
      End
      Begin VB.CommandButton Command12 
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
         Left            =   -65520
         TabIndex        =   54
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Back"
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
         Left            =   -66960
         TabIndex        =   53
         Top             =   840
         Width           =   1095
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -69960
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   3360
         Width           =   2775
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "main.frx":0A57
         Height          =   5055
         Left            =   -74880
         OleObjectBlob   =   "main.frx":0A6B
         TabIndex        =   52
         Top             =   1680
         Width           =   11295
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Save m&odi"
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
         Left            =   -65280
         TabIndex        =   25
         Top             =   3840
         Width           =   1095
      End
      Begin VB.CommandButton Command9 
         Caption         =   "&Modify"
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
         Left            =   -65280
         TabIndex        =   26
         Top             =   3240
         Width           =   1095
      End
      Begin VB.CommandButton Command8 
         Caption         =   "<<Prev"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -73680
         TabIndex        =   50
         Top             =   5880
         Width           =   975
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Next>>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -72480
         TabIndex        =   49
         Top             =   5880
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Delete"
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
         Left            =   -65280
         TabIndex        =   23
         Top             =   2040
         Width           =   1095
      End
      Begin VB.CommandButton Command6 
         Caption         =   "&Exit"
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
         Left            =   -65280
         TabIndex        =   28
         Top             =   5040
         Width           =   1095
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Back"
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
         Left            =   -65280
         TabIndex        =   27
         Top             =   4440
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Sa&ve del"
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
         Left            =   -65280
         TabIndex        =   24
         Top             =   2640
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Save add"
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
         Left            =   -65280
         TabIndex        =   22
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   -65280
         TabIndex        =   21
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         DataField       =   "name"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -73680
         TabIndex        =   1
         Top             =   1020
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         DataField       =   "addr"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -73680
         TabIndex        =   2
         Top             =   1500
         Width           =   2775
      End
      Begin VB.TextBox Text3 
         DataField       =   "phone"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -73680
         TabIndex        =   3
         Top             =   1980
         Width           =   2775
      End
      Begin VB.TextBox Text4 
         DataField       =   "branch"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -73680
         TabIndex        =   4
         Top             =   2460
         Width           =   1095
      End
      Begin VB.TextBox Text5 
         DataField       =   "year"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -73680
         TabIndex        =   5
         Top             =   2940
         Width           =   1095
      End
      Begin VB.TextBox Text6 
         DataField       =   "roll"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -73680
         TabIndex        =   6
         Top             =   3420
         Width           =   1095
      End
      Begin VB.ListBox List1 
         DataField       =   "sex"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   -73680
         TabIndex        =   31
         Top             =   4500
         Width           =   1095
      End
      Begin VB.TextBox Text7 
         DataField       =   "class10"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -67080
         TabIndex        =   9
         Top             =   660
         Width           =   1455
      End
      Begin VB.TextBox Text8 
         DataField       =   "class12"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -67080
         TabIndex        =   10
         Top             =   1110
         Width           =   1455
      End
      Begin VB.TextBox Text9 
         DataField       =   "sem1"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -67080
         TabIndex        =   11
         Top             =   1620
         Width           =   1455
      End
      Begin VB.TextBox Text10 
         DataField       =   "sem2"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -67080
         TabIndex        =   12
         Top             =   2100
         Width           =   1455
      End
      Begin VB.TextBox Text11 
         DataField       =   "sem3"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -67080
         TabIndex        =   13
         Top             =   2580
         Width           =   1455
      End
      Begin VB.TextBox Text12 
         DataField       =   "sem4"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -67080
         TabIndex        =   14
         Top             =   3060
         Width           =   1455
      End
      Begin VB.TextBox Text13 
         DataField       =   "sem5"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -67080
         TabIndex        =   15
         Top             =   3540
         Width           =   1455
      End
      Begin VB.TextBox Text14 
         DataField       =   "sem6"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -67080
         TabIndex        =   16
         Top             =   4020
         Width           =   1455
      End
      Begin VB.TextBox Text16 
         DataField       =   "sem8"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -67080
         TabIndex        =   18
         Top             =   4860
         Width           =   1455
      End
      Begin VB.TextBox Text15 
         DataField       =   "sem7"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -67080
         TabIndex        =   17
         Top             =   4440
         Width           =   1455
      End
      Begin VB.Frame Frame1 
         Caption         =   "Select search criteria"
         Height          =   615
         Left            =   6000
         TabIndex        =   111
         Top             =   1680
         Width           =   4095
         Begin VB.Label Label45 
            Caption         =   "Exact match"
            BeginProperty Font 
               Name            =   "Lucida Sans"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   480
            TabIndex        =   112
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label46 
            Caption         =   "Appx. match"
            BeginProperty Font 
               Name            =   "Lucida Sans"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   2640
            TabIndex        =   113
            Top             =   240
            Width           =   1335
         End
      End
      Begin MSDBGrid.DBGrid DBGrid3 
         Bindings        =   "main.frx":143A
         Height          =   3615
         Left            =   120
         OleObjectBlob   =   "main.frx":144E
         TabIndex        =   127
         Top             =   3120
         Width           =   11295
      End
      Begin VB.Label Label55 
         Caption         =   "OverallTotal"
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
         Left            =   10320
         TabIndex        =   126
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label54 
         Caption         =   "SemsTotal"
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
         Left            =   10320
         TabIndex        =   125
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label53 
         Caption         =   "DoB"
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
         Left            =   10320
         TabIndex        =   124
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label52 
         Caption         =   "Email"
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
         Left            =   9480
         TabIndex        =   123
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label51 
         Caption         =   "(MM/DD/YY)"
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
         Left            =   -72360
         TabIndex        =   118
         Top             =   4080
         Width           =   1335
      End
      Begin VB.Label Label50 
         Caption         =   "Overall Total"
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
         Left            =   -68520
         TabIndex        =   117
         Top             =   5880
         Width           =   1335
      End
      Begin VB.Label Label49 
         Caption         =   "Sems Total"
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
         Left            =   -68520
         TabIndex        =   116
         Top             =   5400
         Width           =   1215
      End
      Begin VB.Label Label48 
         Caption         =   "Email-ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74640
         TabIndex        =   115
         Top             =   5040
         Width           =   855
      End
      Begin VB.Label Label47 
         Caption         =   "* DOB "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74640
         TabIndex        =   114
         Top             =   4080
         Width           =   735
      End
      Begin VB.Label Label21 
         Caption         =   "Descending"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   108
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label20 
         Caption         =   "Ascending"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   107
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label44 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label44"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   6480
         TabIndex        =   104
         Top             =   2520
         Width           =   3015
      End
      Begin VB.Label Label43 
         Caption         =   "Sex"
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
         Left            =   8280
         TabIndex        =   101
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label42 
         Caption         =   "Enter upper range"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   255
         Left            =   1920
         TabIndex        =   99
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label41 
         Caption         =   "Enter lower range"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   98
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label40 
         Caption         =   "Enter value"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   255
         Left            =   1200
         TabIndex        =   97
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label39 
         Caption         =   "Sem8"
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
         Left            =   7200
         TabIndex        =   95
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label38 
         Caption         =   "Sem7"
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
         Left            =   6120
         TabIndex        =   94
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label37 
         Caption         =   "Sem6"
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
         Left            =   4920
         TabIndex        =   93
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label36 
         Caption         =   "Sem5"
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
         Left            =   3720
         TabIndex        =   92
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label35 
         Caption         =   "Sem4"
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
         Left            =   9480
         TabIndex        =   91
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label34 
         Caption         =   "Sem3"
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
         Left            =   8280
         TabIndex        =   90
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label33 
         Caption         =   "Sem2"
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
         Left            =   7200
         TabIndex        =   89
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label32 
         Caption         =   "Sem1"
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
         Left            =   6120
         TabIndex        =   88
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label31 
         Caption         =   "Class 12"
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
         Left            =   4920
         TabIndex        =   76
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label30 
         Caption         =   "Class 10"
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
         Left            =   3720
         TabIndex        =   74
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label29 
         Caption         =   "Year"
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
         Left            =   9480
         TabIndex        =   72
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label28 
         Caption         =   "Branch"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8280
         TabIndex        =   71
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label27 
         Caption         =   "Phone"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7200
         TabIndex        =   70
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label26 
         Caption         =   "Roll"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6120
         TabIndex        =   69
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label25 
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4920
         TabIndex        =   68
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label24 
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   67
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label23 
         Caption         =   "SELECT THE SEARCH OPTION"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   60
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         Caption         =   "Label22"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   495
         Left            =   -70680
         TabIndex        =   55
         Top             =   840
         Width           =   3015
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   -71160
         TabIndex        =   51
         Top             =   840
         Width           =   3255
      End
      Begin VB.Label Label3 
         Caption         =   "* Name"
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
         Left            =   -74640
         TabIndex        =   48
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "* Address"
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
         Left            =   -74640
         TabIndex        =   47
         Top             =   1620
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Branch"
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
         Left            =   -74640
         TabIndex        =   46
         Top             =   2580
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Year"
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
         Left            =   -74640
         TabIndex        =   45
         Top             =   3060
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "* Roll"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74640
         TabIndex        =   44
         Top             =   3540
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Phone No"
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
         Left            =   -74640
         TabIndex        =   43
         Top             =   2100
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "* Sex"
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
         Left            =   -74640
         TabIndex        =   42
         Top             =   4500
         Width           =   495
      End
      Begin VB.Label Label10 
         Caption         =   "Class 10 (%)"
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
         Left            =   -68520
         TabIndex        =   41
         Top             =   660
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "Class 12 (%)"
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
         Left            =   -68520
         TabIndex        =   40
         Top             =   1140
         Width           =   1215
      End
      Begin VB.Label Label12 
         Caption         =   "1st SEM"
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
         Left            =   -68520
         TabIndex        =   39
         Top             =   1620
         Width           =   1455
      End
      Begin VB.Label Label13 
         Caption         =   "2nd SEM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -68520
         TabIndex        =   38
         Top             =   2100
         Width           =   1335
      End
      Begin VB.Label Label14 
         Caption         =   "3rd SEM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -68520
         TabIndex        =   37
         Top             =   2580
         Width           =   1095
      End
      Begin VB.Label Label15 
         Caption         =   "4th SEM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -68520
         TabIndex        =   36
         Top             =   3060
         Width           =   975
      End
      Begin VB.Label Label16 
         Caption         =   "5th SEM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -68520
         TabIndex        =   35
         Top             =   3540
         Width           =   1095
      End
      Begin VB.Label Label17 
         Caption         =   "6th SEM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -68520
         TabIndex        =   34
         Top             =   4020
         Width           =   975
      End
      Begin VB.Label Label18 
         Caption         =   "7th SEM"
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
         Left            =   -68520
         TabIndex        =   33
         Top             =   4500
         Width           =   855
      End
      Begin VB.Label Label19 
         Caption         =   "8th SEM"
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
         Left            =   -68520
         TabIndex        =   32
         Top             =   4980
         Width           =   975
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Student Information System"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs1 As Recordset
Dim rs2 As Recordset
Dim con1 As Connection
Dim con2 As Connection
Dim f As Integer

Private Sub Adodc1_WillMove(ByVal adReason As ADODB.EventReasonEnum, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

End Sub

Private Sub Check18_Click()
If Check18.Value = 1 Then
Check18.Enabled = False
Check19.Enabled = True
Check19.Value = 0
End If
End Sub

Private Sub Check19_Click()
If Check19.Value = 1 Then
Check18.Enabled = 1
Check19.Enabled = 0
Check18.Value = 0
End If
End Sub

Private Sub Check24_Click()

If Check24.Value = 1 Then
Check24.Enabled = False
Check25.Enabled = True
Check25.Value = 0
End If
End Sub

Private Sub Check25_Click()
 If Check25.Value = 1 Then
 Check24.Enabled = 1
Check25.Enabled = 0
Check24.Value = 0
End If
End Sub

Private Sub Command1_Click()
'Adodc1.Visible = False
Call txt1
Text6.Enabled = True
Command3.Enabled = True
Command2.Enabled = True
Command9.Enabled = True
Command1.Enabled = 0
Command4.Enabled = 0
Command7.Enabled = 0
Command8.Enabled = 0
Command10.Enabled = 0
Form1.Text1.Text = ""
Form1.Text2.Text = ""
Form1.Text3.Text = ""
Form1.Text6.Text = ""
Text20.Text = ""
Text21.Text = ""
If Text7.Enabled = True Then
Text7.Text = ""
End If
If Text8.Enabled = True Then
Text8.Text = ""
End If
If Text9.Enabled = True Then
Text9.Text = ""
End If
If Text10.Enabled = True Then
Text10.Text = ""
End If
If Text11.Enabled = True Then
Text11.Text = ""
End If
If Text12.Enabled = True Then
Text12.Text = ""
End If
If Text13.Enabled = True Then
Text13.Text = ""
End If
If Text14.Enabled = True Then
Text14.Text = ""
End If
If Text15.Enabled = True Then
Text15.Text = ""
End If
If Text16.Enabled = True Then
Text16.Text = ""
End If
Form1.Text22.Text = ""
Form1.Text23.Text = ""
'=======
'Form1.Text7.Text = 0
'Form1.Text8.Text = 0
'Form1.Text9.Text = 0
'Form1.Text10.Text = 0
'Form1.Text11.Text = 0
'Form1.Text12.Text = 0
'Form1.Text13.Text = 0
'Form1.Text14.Text = 0
'Form1.Text15.Text = 0
'Form1.Text16.Text = 0
'=========
'Form1.Text1.Enabled = 1
'Form1.Text2.Enabled = 1
'Form1.Text3.Enabled = 1
'Form1.Text6.Enabled = 1
'Form1.Text7.Enabled = 1
'Form1.Text8.Enabled = 1
'Form1.Text9.Enabled = 1
'Form1.Text10.Enabled = 1
'Form1.Text11.Enabled = 1
'Form1.Text12.Enabled = 1
'Form1.Text13.Enabled = 1
'Form1.Text14.Enabled = 1
'Form1.Text15.Enabled = 1
'Form1.Text16.Enabled = 1
Call txt1
End Sub

Private Sub Command10_Click()

ps = Len(Text21.Text)
at = 0
dot = 0
For i = 1 To ps
If Mid(Text21.Text, i, 1) = "@" Then
at = at + 1
End If
If Mid(Text21.Text, i, 1) = "." Then
dot = dot + 1
End If
Next
If ps > 0 And (dot = 0 Or at <> 1) Then
MsgBox "Invalid Email ID"
ElseIf IsDate(Text20.Text) = False Then
MsgBox "Invalid DOB ! Please re enter DOB"
Text20.Text = ""
Text20.SetFocus
ElseIf (Text3.Text <> "" And Len(Text3.Text) < 6) Then
MsgBox "Enter a valid phone no"
ElseIf Text1.Text = "" Or Text2.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Text7.Text = "" Or Text8.Text = "" Or Text9.Text = "" _
Or Text10.Text = "" Or (Text11.Text = "" And Text11.Enabled = True) Or (Text12.Text = "" And Text12.Enabled = True) Or (Text13.Text = "" And Text13.Enabled = True) Or (Text14.Text = "" And Text14.Enabled = True) Or (Text15.Text = "" And Text15.Enabled = True) Or (Text16.Text = "" And Text16.Enabled = True) Or List1.Text = "" Then
MsgBox "Please fill all the fields properly"
Else
If Text21.Text = "" Then
Text21.Text = "Nil"
End If
If Text3.Text = "" Then
'Text3.Text = "Nil"
rs1.update Array("name", "addr", "phone", "branch", "year", "roll", "sex", "class10", "class12", "sem1", "sem2", "sem3", "sem4", "sem5", "sem6", "sem7", "sem8", "DoB", "Email", "SemsTotal", "Overalltotal"), Array(Text1.Text, Text2.Text, "Nil", Text4.Text, Text5.Text, Text6.Text, List1.Text, Text7.Text, Text8.Text, Text9.Text, Text10.Text, Text11.Text, Text12.Text, Text13.Text, Text14.Text, Text15.Text, Text16.Text, Text20.Text, Text21.Text, Text22.Text, Text23.Text)
Else
rs1.update Array("name", "addr", "phone", "branch", "year", "roll", "sex", "class10", "class12", "sem1", "sem2", "sem3", "sem4", "sem5", "sem6", "sem7", "sem8", "DoB", "Email", "SemsTotal", "Overalltotal"), Array(Text1.Text, Text2.Text, Text3.Text, Text4.Text, Text5.Text, Text6.Text, List1.Text, Text7.Text, Text8.Text, Text9.Text, Text10.Text, Text11.Text, Text12.Text, Text13.Text, Text14.Text, Text15.Text, Text16.Text, Text20.Text, Text21.Text, Text22.Text, Text23.Text)
End If
MsgBox "Successfully Updated"
Command9.Enabled = 1
Command10.Enabled = 0
Command7.Enabled = 0
Command8.Enabled = 0
'Form1.Text1.Text = ""
'Form1.Text2.Text = ""
'Form1.Text3.Text = ""
'Form1.Text6.Text = ""
'Form1.Text7.Text = ""
'Form1.Text8.Text = ""
'Form1.Text9.Text = ""
'Form1.Text10.Text = ""
'Form1.Text11.Text = ""
'Form1.Text12.Text = ""
'Form1.Text13.Text = ""
'Form1.Text14.Text = ""
'Form1.Text15.Text = ""
'Form1.Text16.Text = ""
Text6.Enabled = 1
End If
Data1.Refresh
Data2.Refresh
Call txt0
'===
'rs1.Close
'con1.Close
'===
End Sub

Private Sub Command11_Click()
Me.Hide
Form2.WindowState = 2
Form2.Show
End Sub

Private Sub Command12_Click()
'End
Form1.Hide
Form2.Hide
MDIForm1.Enabled = True
End Sub

Private Sub Command13_Click()
Me.Hide
Form2.WindowState = 2
Form2.Show
End Sub

Private Sub Command14_Click()
'End
Form1.Hide
Form2.Hide
MDIForm1.Enabled = True
End Sub

Private Sub Command15_Click()
DBGrid3.Visible = 1
Data3.Refresh
If List4.ListIndex < 0 Then
MsgBox "Please select a search option", vbCritical
ElseIf (Text17.Text = "" And Text17.Visible = True) Or _
(Text18.Text = "" And Text18.Visible = True) Or _
(Text19.Text = "" And Text19.Visible = True) Then
MsgBox "Please enter the range/value"
Else
Select Case Form2.Combo1.Text
 Case "IT":
  Select Case Form2.Combo2.Text
    Case "1st":
     st2 = "branch='IT' and year='1st'"
    Case "2nd":
     st2 = "branch='IT' and year='2nd'"
    Case "3rd":
    st2 = "branch='IT' and year='3rd'"
     Case "4th":
     st2 = "branch='IT' and year='4th'"
     End Select
  Case "CSE":
  Select Case Form2.Combo2.Text
   Case "1st": st2 = "branch='CSE' and year='1st'"
    Case "2nd": st2 = "branch='CSE' and year='2nd'"
    Case "3rd": st2 = "branch='CSE' and year='3rd'"
    Case "4th": st2 = "branch='CSE' and year='4th'"
  End Select
  Case "ECE":
  Select Case Form2.Combo2.Text
   Case "1st": st2 = "branch='ECE' and year='1st'"
    Case "2nd": st2 = "branch='ECE' and year='2nd'"
    Case "3rd": st2 = "branch='ECE' and year='3rd'"
    Case "4th": st2 = "branch='ECE' and year='4th'"
   End Select
End Select
Select Case List4.ListIndex
    Case 0:
            If Check18.Value = 1 Then
            st1 = "name=" & "'" & Text19.Text & "'"
            ElseIf Check19.Value = 1 Then
            st1 = " name like " & "'" & Text19.Text & "*'"
            End If
    Case 1:
            If Check18.Value = 1 Then
            st1 = "addr=" & "'" & Text19.Text & "'"
            ElseIf Check19.Value = 1 Then
            st1 = " addr like " & "'" & Text19.Text & "*'"
            End If
    Case 2: st1 = "roll between " & Text17.Text & " And " & Text18.Text
    Case 3:
             If Check18.Value = 1 Then
            st1 = "phone=" & "'" & Text19.Text & "'"
            ElseIf Check19.Value = 1 Then
            st1 = " phone like " & "'" & Text19.Text & "*'"
            End If
    Case 4: st1 = "class10 between " & Text17.Text & " And " & Text18.Text
    Case 5: st1 = "class12 between " & Text17.Text & " And " & Text18.Text
    Case 6: st1 = "sem1 between " & Text17.Text & " And " & Text18.Text
    Case 7: st1 = "sem2 between " & Text17.Text & " And " & Text18.Text
    Case 8: st1 = "sem3 between " & Text17.Text & " And " & Text18.Text
    Case 9: st1 = "sem4 between " & Text17.Text & " And " & Text18.Text
    Case 10: st1 = "sem5 between " & Text17.Text & " And " & Text18.Text
    Case 11: st1 = "sem6 between " & Text17.Text & " And " & Text18.Text
    Case 12: st1 = "sem7 between " & Text17.Text & " And " & Text18.Text
    Case 13: st1 = "sem8 between " & Text17.Text & " And " & Text18.Text
     Case 14: st1 = "sex=" & "'" & Text19.Text & "'"
     Case 15:
            If IsDate(Text17.Text) = True And IsDate(Text18.Text) = True Then
             st1 = "dob between " & "#" & Text17.Text & "#" & " And " & "#" & Text18.Text & "#"
             Else
             MsgBox "Invalid date"
             st1 = ""
             Text17.Text = ""
             Text18.Text = ""
             DBGrid3.ClearFields
             GoTo leave
            End If
     Case 16:
            If Check18.Value = 1 Then
            st1 = "email=" & "'" & Text19.Text & "'"
            ElseIf Check19.Value = 1 Then
            st1 = " email like " & "'" & Text19.Text & "*'"
            End If
      Case 17: st1 = "semstotal between " & Text17.Text & " And " & Text18.Text
      Case 18: st1 = "overalltotal between " & Text17.Text & " And " & Text18.Text
     'st = "select * from studac where branch='IT' and year='1st'and phone=" & "'" & Text19.Text & "'"
   End Select
   st3 = ""
   If Check1.Value = 1 Then
   st3 = st3 & " name,"
   End If
   If Check2.Value = 1 Then
   st3 = st3 & " addr,"
   End If
   If Check3.Value = 1 Then
   st3 = st3 & " roll,"
   End If
   If Check21.Value = 1 Then
   st3 = st3 & " dob,"
   End If
   If Check4.Value = 1 Then
   st3 = st3 & " phone,"
   End If
   If Check20.Value = 1 Then
   st3 = st3 & " email,"
   End If
   If Check5.Value = 1 Then
   st3 = st3 & " branch,"
   End If
   If Check6.Value = 1 Then
   st3 = st3 & " year,"
   End If
   If Check7.Value = 1 Then
   st3 = st3 & " class10,"
   End If
   If Check8.Value = 1 Then
   st3 = st3 & " class12,"
   End If
   If Check9.Value = 1 Then
   st3 = st3 & "sem1,"
   End If
   If Check10.Value = 1 Then
   st3 = st3 & " sem2,"
   End If
   If Check11.Value = 1 Then
   st3 = st3 & " sem3,"
   End If
   If Check12.Value = 1 Then
   st3 = st3 & " sem4,"
   End If
   If Check13.Value = 1 Then
   st3 = st3 & " sem5,"
   End If
   If Check14.Value = 1 Then
   st3 = st3 & " sem6,"
   End If
   If Check15.Value = 1 Then
   st3 = st3 & " sem7,"
   End If
   If Check16.Value = 1 Then
   st3 = st3 & " sem8,"
   End If
   If Check22.Value = 1 Then
   st3 = st3 & " semstotal,"
   End If
   If Check23.Value = 1 Then
   st3 = st3 & " overalltotal,"
   End If
   If Check17.Value = 1 Then
   st3 = st3 & " sex,"
   End If
   z = Len(st3)
   st4 = ""
   If z > 0 Then
    For i = 1 To z - 1
    st4 = st4 & Mid(st3, i, 1)
    Next
   End If

   Select Case List4.ListIndex
   Case 0: str5 = " order by name"
            
   Case 1: str5 = " order by addr"
   Case 2: str5 = " order by roll"
   Case 3: str5 = " order by phone"
   Case 4: str5 = " order by class10"
   Case 5: str5 = " order by class12"
   Case 6: str5 = " order by sem1"
   Case 7: str5 = " order by sem2"
   Case 8: str5 = " order by sem3"
   Case 9: str5 = " order by sem4"
   Case 10: str5 = " order by sem5"
   Case 11: str5 = " order by sem6"
   Case 12: str5 = " order by sem7"
   Case 13: str5 = " order by sem8"
   Case 14: str5 = " order by sex"
   Case 15: str5 = " order by dob"
   Case 16: str5 = " order by email"
   Case 17: str5 = " order by semstotal"
   Case 18: str5 = " order by overalltotal"
   End Select
 If Option2.Value = True Then
 str5 = str5 & " desc"
 End If
   
   If z > 0 Then
    st = "select " & st4 & " from studac where " & st2 & " and " & st1 & str5
    If Check25.Value = 1 Then
      st = "select " & st4 & " from studac where (year='1st' or year='2nd' or year='3rd' or year='4th') and " & st1 & str5
     
    End If
    Data3.RecordSource = st
   End If
Form1.Data3.Refresh
End If


leave:
End Sub

Private Sub Command16_Click()
DBGrid3.Visible = False
Me.Hide
Form2.WindowState = 2
Form2.Show
End Sub

Private Sub Command17_Click()
'End
Form1.Hide
Form2.Hide
MDIForm1.Enabled = True
End Sub

Private Sub Command2_Click()

'Adodc1.Visible = 0
Command4.Enabled = True
Command2.Enabled = 0
Command3.Enabled = 0
Command1.Enabled = 1
Command9.Enabled = True
Command10.Enabled = 0
Set rs1 = New Recordset
Set con1 = New Connection
con1.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\student.mdb;Persist Security Info=False"
Select Case Text4.Text
Case "IT":
  Select Case Text5.Text
  Case "1st":
       rs1.open "select * from studac where branch='IT' and year='1st'", con1, adOpenDynamic, adLockOptimistic
  Case "2nd":
       rs1.open "select * from studac where branch='IT' and year='2nd'", con1, adOpenDynamic, adLockOptimistic
  Case "3rd":
       rs1.open "select * from studac where branch='IT' and year='3rd'", con1, adOpenDynamic, adLockOptimistic
  Case "4th":
       rs1.open "select * from studac where branch='IT' and year='4th'", con1, adOpenDynamic, adLockOptimistic
  End Select
Case "ECE":
       'rs1.Open "select * from studac where branch='ECE'", con1, adOpenDynamic, adLockOptimistic
    Select Case Text5.Text
  Case "1st":
       rs1.open "select * from studac where branch='ECE' and year='1st'", con1, adOpenDynamic, adLockOptimistic
  Case "2nd":
       rs1.open "select * from studac where branch='ECE' and year='2nd'", con1, adOpenDynamic, adLockOptimistic
  Case "3rd":
       rs1.open "select * from studac where branch='ECE' and year='3rd'", con1, adOpenDynamic, adLockOptimistic
  Case "4th":
       rs1.open "select * from studac where branch='ECE' and year='4th'", con1, adOpenDynamic, adLockOptimistic
  End Select
Case "CSE":
       'rs1.Open "select * from studac where branch='CSE'", con1, adOpenDynamic, adLockOptimistic
  Select Case Text5.Text
  Case "1st":
       rs1.open "select * from studac where branch='CSE' and year='1st'", con1, adOpenDynamic, adLockOptimistic
  Case "2nd":
       rs1.open "select * from studac where branch='CSE' and year='2nd'", con1, adOpenDynamic, adLockOptimistic
  Case "3rd":
       rs1.open "select * from studac where branch='CSE' and year='3rd'", con1, adOpenDynamic, adLockOptimistic
  Case "4th":
       rs1.open "select * from studac where branch='CSE' and year='4th'", con1, adOpenDynamic, adLockOptimistic
  End Select
End Select
'rs1.AddNew Array("name", "addr", "phone", "branch", "year", "roll", "sex", "class10", "class12", "sem1", "sem2", "sem3", "sem4", "sem5", "sem6", "sem7", "sem8"), Array(Text1.Text, Text2.Text, Text3.Text, Text4.Text, Text5.Text, Text6.Text, List1.Text, Text7.Text, Text8.Text, Text9.Text, Text10.Text, Text11.Text, Text12.Text, Text13.Text, Text14.Text, Text15.Text, Text16.Text)
'rs1.MoveFirst
Do While Not rs1.EOF
flag = flag + 1
rs1.MoveNext
Loop
If flag > 0 Then
rs1.MoveFirst
Text1.Text = rs1!Name
Text2.Text = rs1!addr
'Text3.Text = rs1!phone
If rs1!phone = "Nil" Then
Text3.Text = ""
Else
Text3.Text = rs1!phone
End If
Text4.Text = rs1!branch
Text5.Text = rs1!Year
Text6.Text = rs1!roll
List1.Text = rs1!sex
Text7.Text = rs1!class10
Text8.Text = rs1!class12
Text9.Text = rs1!sem1
Text10.Text = rs1!sem2
Text11.Text = rs1!sem3
Text12.Text = rs1!sem4
Text13.Text = rs1!sem5
Text14.Text = rs1!sem6
Text15.Text = rs1!sem7
Text16.Text = rs1!sem8
Command7.Enabled = 1
Command8.Enabled = 1
Else
MsgBox "No record exists"
Command4.Enabled = 0
End If
Call txt1
End Sub

Private Sub Command3_Click()

pp:
ps = Len(Text21.Text)
at = 0
dot = 0
For i = 1 To ps
If Mid(Text21.Text, i, 1) = "@" Then
at = at + 1
End If
If Mid(Text21.Text, i, 1) = "." Then
dot = dot + 1
End If
Next
If ps > 0 And (dot = 0 Or at <> 1) Then
MsgBox "Invalid Email ID"

ElseIf IsDate(Text20.Text) = False Then
MsgBox "Invalid DOB ! Please re enter DOB"
Text20.Text = ""
Text20.SetFocus
ElseIf (Text3.Text <> "" And Len(Text3.Text) < 6) Then
MsgBox "Enter a valid phone no"
ElseIf Text20.Text = "" Or Text1.Text = "" Or Text2.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Text7.Text = "" Or Text8.Text = "" Or Text9.Text = "" _
Or Text10.Text = "" Or (Text11.Text = "" And Text11.Enabled = True) Or (Text12.Text = "" And Text12.Enabled = True) Or (Text13.Text = "" And Text13.Enabled = True) Or (Text14.Text = "" And Text14.Enabled = True) Or (Text15.Text = "" And Text15.Enabled = True) Or (Text16.Text = "" And Text16.Enabled = True) Or List1.Text = "" Then
MsgBox "Please fill all the fields properly"
Command3.Enabled = True
Else
If MsgBox("Are you sure to save?", vbOKCancel) = vbOK Then
Set rs1 = New Recordset
Set con1 = New Connection
con1.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\student.mdb;Persist Security Info=False"
rs1.open "select * from studac", con1, adOpenDynamic, adLockOptimistic
Do While Not rs1.EOF
If rs1!branch = Text4.Text And rs1!Year = Text5.Text And rs1!roll = Text6.Text Then
MsgBox "The student of same class,year and roll already exists"
Text6.Text = ""
rs1.Close
con1.Close
GoTo pp
End If
rs1.MoveNext
Loop
If Text21.Text = "" Then
Text21.Text = "Nil"
End If
If Text3.Text = "" Then
'Text3.Text = "Nil"
rs1.AddNew Array("name", "addr", "phone", "branch", "year", "roll", "sex", "class10", "class12", "sem1", "sem2", "sem3", "sem4", "sem5", "sem6", "sem7", "sem8", "DoB", "Email", "SemsTotal", "Overalltotal"), Array(Text1.Text, Text2.Text, "Nil", Text4.Text, Text5.Text, Text6.Text, List1.Text, Text7.Text, Text8.Text, Text9.Text, Text10.Text, Text11.Text, Text12.Text, Text13.Text, Text14.Text, Text15.Text, Text16.Text, Text20.Text, Text21.Text, Text22.Text, Text23.Text)
Else
rs1.AddNew Array("name", "addr", "phone", "branch", "year", "roll", "sex", "class10", "class12", "sem1", "sem2", "sem3", "sem4", "sem5", "sem6", "sem7", "sem8", "DoB", "Email", "SemsTotal", "Overalltotal"), Array(Text1.Text, Text2.Text, Text3.Text, Text4.Text, Text5.Text, Text6.Text, List1.Text, Text7.Text, Text8.Text, Text9.Text, Text10.Text, Text11.Text, Text12.Text, Text13.Text, Text14.Text, Text15.Text, Text16.Text, Text20.Text, Text21.Text, Text22.Text, Text23.Text)
End If
'rs1.AddNew Array("name", "addr", "phone", "branch", "year", "roll", "sex"), Array(Text1.Text, Text2.Text, Text3.Text, Text4.Text, Text5.Text, Text6.Text, List1.Text)
MsgBox "Successfully added"
Command3.Enabled = False
Command1.Enabled = 1
Call txt0
'rs1.Close
'con1.Close
Else: Command3.Enabled = True
End If
End If
Data1.Refresh
Data2.Refresh
'End If

End Sub

Private Sub Command4_Click()

'rs1.update Array("name", "addr", "phone", "branch", "year", "roll", "sex", "class10", "class12", "sem1", "sem2", "sem3", "sem4", "sem5", "sem6", "sem7", "sem8", "DoB", "Email", "Sems Total", "Overall total"), Array("x", "x", 0, "x", "0", "0", "x", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 / 0 / 0, " ", 0, 0)
If MsgBox("Are you sure to delete?", vbYesNo) = vbYes Then
rs1.Delete adAffectCurrent
MsgBox "Successfully deleted"
End If
Command2.Enabled = 1
Command4.Enabled = 0
Command7.Enabled = 0
Command8.Enabled = 0
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
Data1.Refresh
Data2.Refresh
Call txt0
'=====
rs1.Close
con1.Close
'=====
End Sub

Private Sub Command5_Click()
Me.Hide
Form2.Show
Form2.WindowState = 2
Call txt0
'rs1.Close
'con1.Close
End Sub

Private Sub Command6_Click()
'End
Form1.Hide
Form2.Hide
MDIForm1.Enabled = True
End Sub

Private Sub Command7_Click()
'rs1.MoveFirst
rs1.MoveNext
If rs1.EOF = True Then
MsgBox "End of file"
Command7.Enabled = False
Command8.Enabled = True
rs1.MoveLast
Else
Command8.Enabled = True
Text1.Text = rs1!Name
Text2.Text = rs1!addr
If rs1!phone = "Nil" Then
Text3.Text = ""
Else
Text3.Text = rs1!phone
End If
Text4.Text = rs1!branch
Text5.Text = rs1!Year
Text6.Text = rs1!roll
List1.Text = rs1!sex
Text7.Text = rs1!class10
Text8.Text = rs1!class12
Text9.Text = rs1!sem1
Text10.Text = rs1!sem2
Text11.Text = rs1!sem3
Text12.Text = rs1!sem4
Text13.Text = rs1!sem5
Text14.Text = rs1!sem6
Text15.Text = rs1!sem7
Text16.Text = rs1!sem8
Text20.Text = rs1!dob
'Text21.Text = rs1!email
If rs1!email = "Nil" Then
Text21.Text = ""
Else
Text21.Text = rs1!email
End If
Text22.Text = rs1!semstotal
Text23.Text = rs1!overalltotal

End If
End Sub

Private Sub Command8_Click()
'rs1.MoveLast

rs1.MovePrevious
If rs1.BOF = True Then
MsgBox "Begin of file"
Command8.Enabled = False
Command7.Enabled = True
rs1.MoveFirst
Else
Command7.Enabled = True
Text1.Text = rs1!Name
Text2.Text = rs1!addr
'Text3.Text = rs1!phone
If rs1!phone = "Nil" Then
Text3.Text = ""
Else
Text3.Text = rs1!phone
End If
Text4.Text = rs1!branch
Text5.Text = rs1!Year
Text6.Text = rs1!roll
List1.Text = rs1!sex
Text7.Text = rs1!class10
Text8.Text = rs1!class12
Text9.Text = rs1!sem1
Text10.Text = rs1!sem2
Text11.Text = rs1!sem3
Text12.Text = rs1!sem4
Text13.Text = rs1!sem5
Text14.Text = rs1!sem6
Text15.Text = rs1!sem7
Text16.Text = rs1!sem8
Text20.Text = rs1!dob
'Text21.Text = rs1!email
If rs1!email = "Nil" Then
Text21.Text = ""
Else
Text21.Text = rs1!email
End If
Text22.Text = rs1!semstotal
Text23.Text = rs1!overalltotal

End If
End Sub

Private Sub Command9_Click()

'Form1.Text1.Enabled = 1
'Form1.Text2.Enabled = 1
'Form1.Text3.Enabled = 1
''Form1.Text6.Enabled = 1
'Form1.Text7.Enabled = 1
'Form1.Text8.Enabled = 1
'Form1.Text9.Enabled = 1
'Form1.Text10.Enabled = 1
'Form1.Text11.Enabled = 1
'Form1.Text12.Enabled = 1
'Form1.Text13.Enabled = 1
'Form1.Text14.Enabled = 1
'Form1.Text15.Enabled = 1
'Form1.Text16.Enabled = 1
Text6.Enabled = False

'Adodc1.Visible = 0
Command4.Enabled = 0
Command2.Enabled = 1
Command3.Enabled = 0
Command1.Enabled = 1
Command9.Enabled = 0
Command10.Enabled = 1
Set rs1 = New Recordset
Set con1 = New Connection
con1.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\student.mdb;Persist Security Info=False"
Select Case Text4.Text
Case "IT":
  Select Case Text5.Text
  Case "1st":
       rs1.open "select * from studac where branch='IT' and year='1st'", con1, adOpenDynamic, adLockOptimistic
  Case "2nd":
       rs1.open "select * from studac where branch='IT' and year='2nd'", con1, adOpenDynamic, adLockOptimistic
  Case "3rd":
       rs1.open "select * from studac where branch='IT' and year='3rd'", con1, adOpenDynamic, adLockOptimistic
  Case "4th":
       rs1.open "select * from studac where branch='IT' and year='4th'", con1, adOpenDynamic, adLockOptimistic
  End Select
Case "ECE":
       'rs1.Open "select * from studac where branch='ECE'", con1, adOpenDynamic, adLockOptimistic
    Select Case Text5.Text
  Case "1st":
       rs1.open "select * from studac where branch='ECE' and year='1st'", con1, adOpenDynamic, adLockOptimistic
  Case "2nd":
       rs1.open "select * from studac where branch='ECE' and year='2nd'", con1, adOpenDynamic, adLockOptimistic
  Case "3rd":
       rs1.open "select * from studac where branch='ECE' and year='3rd'", con1, adOpenDynamic, adLockOptimistic
  Case "4th":
       rs1.open "select * from studac where branch='ECE' and year='4th'", con1, adOpenDynamic, adLockOptimistic
  End Select
Case "CSE":
       'rs1.Open "select * from studac where branch='CSE'", con1, adOpenDynamic, adLockOptimistic
  Select Case Text5.Text
  Case "1st":
       rs1.open "select * from studac where branch='CSE' and year='1st'", con1, adOpenDynamic, adLockOptimistic
  Case "2nd":
       rs1.open "select * from studac where branch='CSE' and year='2nd'", con1, adOpenDynamic, adLockOptimistic
  Case "3rd":
       rs1.open "select * from studac where branch='CSE' and year='3rd'", con1, adOpenDynamic, adLockOptimistic
  Case "4th":
       rs1.open "select * from studac where branch='CSE' and year='4th'", con1, adOpenDynamic, adLockOptimistic
  End Select
End Select
'rs1.AddNew Array("name", "addr", "phone", "branch", "year", "roll", "sex", "class10", "class12", "sem1", "sem2", "sem3", "sem4", "sem5", "sem6", "sem7", "sem8"), Array(Text1.Text, Text2.Text, Text3.Text, Text4.Text, Text5.Text, Text6.Text, List1.Text, Text7.Text, Text8.Text, Text9.Text, Text10.Text, Text11.Text, Text12.Text, Text13.Text, Text14.Text, Text15.Text, Text16.Text)
'rs1.MoveFirst
Do While Not rs1.EOF
flag = flag + 1
rs1.MoveNext
Loop
If flag > 0 Then
rs1.MoveFirst
Text1.Text = rs1!Name
Text2.Text = rs1!addr
'Text3.Text = rs1!phone
If rs1!phone = "Nil" Then
Text3.Text = ""
Else
Text3.Text = rs1!phone
End If
Text4.Text = rs1!branch
Text5.Text = rs1!Year
Text6.Text = rs1!roll
List1.Text = rs1!sex
Text7.Text = rs1!class10
Text8.Text = rs1!class12
Text9.Text = rs1!sem1
Text10.Text = rs1!sem2
Text11.Text = rs1!sem3
Text12.Text = rs1!sem4
Text13.Text = rs1!sem5
Text14.Text = rs1!sem6
Text15.Text = rs1!sem7
Text16.Text = rs1!sem8
Text20.Text = rs1!dob
'Text21.Text = rs1!email
If rs1!email = "Nil" Then
Text21.Text = ""
Else
Text21.Text = rs1!email
End If
Text22.Text = rs1!semstotal
Text23.Text = rs1!overalltotal

Command7.Enabled = 1
Command8.Enabled = 1
Else
MsgBox "No record exists"
Command10.Enabled = 0
Command9.Enabled = 1
End If

Call txt1


End Sub

Private Sub Form_Load()
Call txt0
Check24.Value = 1
'************
Text23.Text = Val(Text7.Text) + Val(Text8.Text) + Val(Text9.Text) + Val(Text10.Text) _
+ Val(Text11.Text) + Val(Text12.Text) + Val(Text13.Text) + Val(Text14.Text) + Val(Text15.Text) + Val(Text16.Text)

Text22.Enabled = False
Text23.Enabled = False

'*************
Check18.Value = 1
Check19.Enabled = 0
Check1.Value = 1
Check3.Value = 1
Check5.Value = 1
Check6.Value = 1
Option1.Value = True
Data3.Visible = False
Text17.Visible = False
Text18.Visible = False
Text19.Visible = False
Label40.Visible = False
Label41.Visible = False
Label42.Visible = False
List4.AddItem "By Name"
List4.AddItem "By Address"
List4.AddItem "By Roll no"
List4.AddItem "By Phone"
List4.AddItem "By Class 10 marks"
List4.AddItem "By Class 12 marks"
List4.AddItem "By Semester 1 marks"
List4.AddItem "By Semester 2 marks"
List4.AddItem "By Semester 3 marks"
List4.AddItem "By Semester 4 marks"
List4.AddItem "By Semester 5 marks"
List4.AddItem "By Semester 6 marks"
List4.AddItem "By Semester 7 marks"
List4.AddItem "By Semester 8 marks"
List4.AddItem "By Sex"
'-------
List4.AddItem "By Date of Birth"
List4.AddItem "By Email"
List4.AddItem "By Semester total"
List4.AddItem "By overall total"



 

Data1.Visible = False
Data2.Visible = False

Command3.Enabled = 0
Command4.Enabled = 0
Command10.Enabled = 0
Command7.Enabled = 0
Command8.Enabled = 0
List1.AddItem "M"
List1.AddItem "F"
'Adodc1.Visible = False
'Text4.Text = Form2.Combo1.List(Form2.Combo1.ListIndex)
'Text5.Text = Form2.Combo2.List(Form2.Combo2.ListIndex)
'=========
Call txt0
End Sub

Private Sub List4_Click()
Text17.Text = ""
Text18.Text = ""
Text19.Text = ""
If List4.Text = "By Name" Or List4.Text = "By Address" Or List4.Text = "By Phone" Or List4.Text = "By Sex" Or List4.Text = "By Email" Then
Text19.Visible = True
Text17.Visible = 0
Text18.Visible = 0
Label40.Visible = 1
Label41.Visible = False
Label42.Visible = False
Check19.Enabled = 1
ElseIf List4.Text = "By Roll no" _
 Or List4.Text = "By Class 10 marks" _
 Or List4.Text = "By Class 12 marks" _
 Or List4.Text = "By Semester 1 marks" _
 Or List4.Text = "By Semester 2 marks" _
 Or List4.Text = "By Semester 3 marks" _
Or List4.Text = "By Semester 4 marks" _
 Or List4.Text = "By Semester 5 marks" _
Or List4.Text = "By Semester 6 marks" _
Or List4.Text = "By Semester 7 marks" _
Or List4.Text = "By Semester 8 marks" _
Or List4.Text = "By Date of Birth" _
Or List4.Text = "By Semester total" _
Or List4.Text = "By overall total" Then

Check19.Enabled = 0
Check18.Value = 1
Text17.Visible = True
Text18.Visible = True
Text19.Visible = 0
Label40.Visible = 0
Label41.Visible = 1
Label42.Visible = 1
End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)

Data1.Refresh
Data2.Refresh
End Sub

Private Sub Text1_Change()
x = Len(Text1.Text)
For i = 1 To x
If UCase(Mid(Text1.Text, i, 1)) >= "A" And UCase(Mid(Text1.Text, i, 1)) <= "Z" Or _
Mid(Text1.Text, i, 1) = "." Or Mid(Text1.Text, i, 1) = " " Then
Else
MsgBox "Enter only character"
Text1.Text = ""
End If
Next
End Sub

Private Sub Text10_Change()
If IsNumeric(Text10.Text) = False And Len(Text10.Text) >= 1 Then
MsgBox "Enter only number"
Text10.Text = ""
End If
Text22.Text = Val(Text9.Text) + Val(Text10.Text) _
+ Val(Text11.Text) + Val(Text12.Text) + Val(Text13.Text) + Val(Text14.Text) + Val(Text15.Text) + Val(Text16.Text)

End Sub

Private Sub Text11_Change()
If IsNumeric(Text11.Text) = False And Len(Text11.Text) >= 1 Then
MsgBox "Enter only number"
Text11.Text = ""
End If
Text22.Text = Val(Text9.Text) + Val(Text10.Text) _
+ Val(Text11.Text) + Val(Text12.Text) + Val(Text13.Text) + Val(Text14.Text) + Val(Text15.Text) + Val(Text16.Text)

End Sub

Private Sub Text12_Change()
If IsNumeric(Text12.Text) = False And Len(Text12.Text) >= 1 Then
MsgBox "Enter only number"
Text12.Text = ""
End If
Text22.Text = Val(Text9.Text) + Val(Text10.Text) _
+ Val(Text11.Text) + Val(Text12.Text) + Val(Text13.Text) + Val(Text14.Text) + Val(Text15.Text) + Val(Text16.Text)

End Sub

Private Sub Text13_Change()
If IsNumeric(Text13.Text) = False And Len(Text13.Text) >= 1 Then
MsgBox "Enter only number"
Text13.Text = ""
End If
Text22.Text = Val(Text9.Text) + Val(Text10.Text) _
+ Val(Text11.Text) + Val(Text12.Text) + Val(Text13.Text) + Val(Text14.Text) + Val(Text15.Text) + Val(Text16.Text)

End Sub

Private Sub Text14_Change()
If IsNumeric(Text14.Text) = False And Len(Text14.Text) >= 1 Then
MsgBox "Enter only number"
Text14.Text = ""
End If
Text22.Text = Val(Text9.Text) + Val(Text10.Text) _
+ Val(Text11.Text) + Val(Text12.Text) + Val(Text13.Text) + Val(Text14.Text) + Val(Text15.Text) + Val(Text16.Text)

End Sub

Private Sub Text15_Change()
If IsNumeric(Text15.Text) = False And Len(Text15.Text) >= 1 Then
MsgBox "Enter only number"
Text15.Text = ""
End If
Text22.Text = Val(Text9.Text) + Val(Text10.Text) _
+ Val(Text11.Text) + Val(Text12.Text) + Val(Text13.Text) + Val(Text14.Text) + Val(Text15.Text) + Val(Text16.Text)

End Sub

Private Sub Text16_Change()
If IsNumeric(Text16.Text) = False And Len(Text16.Text) >= 1 Then
MsgBox "Enter only number"
Text16.Text = ""
End If
Text22.Text = Val(Text9.Text) + Val(Text10.Text) _
+ Val(Text11.Text) + Val(Text12.Text) + Val(Text13.Text) + Val(Text14.Text) + Val(Text15.Text) + Val(Text16.Text)

End Sub

Private Sub Text17_Change()
If List4.Text = "By Date of Birth" Then
  v = Len(Text17.Text)
  For i = 1 To v
  If (UCase(Mid(Text17.Text, i, 1)) >= 0 And UCase(Mid(Text17.Text, i, 1)) <= 9) Or Mid(Text17.Text, i, 1) = "/" Then
  Else
  MsgBox "Enter only INTEGER NUMBER and / for DoB"
  Text17.Text = ""
  End If
  Next
  '---------
ElseIf List4.Text = "By Roll no" Then
x = Len(Text17.Text)
For i = 1 To x
If UCase(Mid(Text17.Text, i, 1)) >= 0 And UCase(Mid(Text17.Text, i, 1)) <= 9 Then
Else
MsgBox "Enter only INTEGER NUMBER"
Text17.Text = ""
End If
Next
Else
If IsNumeric(Text17.Text) = False And Len(Text17.Text) >= 1 Then
MsgBox "Enter only numbers"
Text17.Text = ""
End If
End If
End Sub

Private Sub Text18_Change()
If List4.Text = "By Date of Birth" Then
y = Len(Text18.Text)
For i = 1 To y
If (UCase(Mid(Text18.Text, i, 1)) >= 0 And UCase(Mid(Text18.Text, i, 1)) <= 9) Or Mid(Text18.Text, i, 1) = "/" Then
Else
MsgBox "Enter only INTEGER NUMBER or / for DoB"
Text18.Text = ""
End If
Next
'-----
ElseIf List4.Text = "By Roll no" Then
x = Len(Text18.Text)
For i = 1 To x
If UCase(Mid(Text18.Text, i, 1)) >= 0 And UCase(Mid(Text18.Text, i, 1)) <= 9 Then
Else
MsgBox "Enter only INTEGER NUMBER"
Text18.Text = ""
End If
Next
Else
If IsNumeric(Text18.Text) = False And Len(Text18.Text) >= 1 Then
MsgBox "Enter only numbers"
Text18.Text = ""
End If
End If
End Sub

Private Sub Text21_Change()
x = Len(Text21.Text)
For i = 1 To x
If UCase(Mid(Text21.Text, i, 1)) >= 0 And UCase(Mid(Text21.Text, i, 1)) <= 9 Or _
Mid(Text21.Text, i, 1) = "." Or Mid(Text21.Text, i, 1) = "_" Or Mid(Text21.Text, i, 1) = "-" Or Mid(Text21.Text, i, 1) = "@" _
Or UCase(Mid(Text21.Text, i, 1)) >= "A" And UCase(Mid(Text21.Text, i, 1)) <= "Z" Then
Else
MsgBox "Enter only Valid Email Address"
Text21.Text = ""
End If
Next
End Sub

Private Sub Text22_Change()
Text22.Text = Val(Text9.Text) + Val(Text10.Text) _
+ Val(Text11.Text) + Val(Text12.Text) + Val(Text13.Text) + Val(Text14.Text) + Val(Text15.Text) + Val(Text16.Text)
Text23.Text = Val(Text7.Text) + Val(Text8.Text) + Val(Text22.Text)
End Sub

Private Sub Text23_Change()
Text23.Text = Val(Text7.Text) + Val(Text8.Text) + Val(Text9.Text) + Val(Text10.Text) _
+ Val(Text11.Text) + Val(Text12.Text) + Val(Text13.Text) + Val(Text14.Text) + Val(Text15.Text) + Val(Text16.Text)

End Sub

Private Sub Text3_Change()
x = Len(Text3.Text)
If x <= 7 Then
For i = 1 To x
If UCase(Mid(Text3.Text, i, 1)) >= 0 And UCase(Mid(Text3.Text, i, 1)) <= 9 Then
Else
MsgBox "Enter only 6/7 digit NUMBER"
Text3.Text = ""
End If
Next
Else
MsgBox "Cannot exceed 7 digit"
Text3.Text = ""
End If
End Sub

Private Sub Text6_Change()
x = Len(Text6.Text)
For i = 1 To x
If UCase(Mid(Text6.Text, i, 1)) >= 0 And UCase(Mid(Text6.Text, i, 1)) <= 9 Then
Else
MsgBox "Enter only INTEGER NUMBERS"
Text6.Text = ""
End If
Next
End Sub

Private Sub Text7_Change()
If (IsNumeric(Text7.Text) = False And Len(Text7.Text) >= 1) Or Val(Text7.Text) > 100 Then
MsgBox "Enter only percentage number"
Text7.Text = ""
End If
Text22.Text = Val(Text9.Text) + Val(Text10.Text) _
+ Val(Text11.Text) + Val(Text12.Text) + Val(Text13.Text) + Val(Text14.Text) + Val(Text15.Text) + Val(Text16.Text)

Text23.Text = Val(Text7.Text) + Val(Text8.Text) + Val(Text22.Text)
End Sub

Private Sub Text8_Change()
If (IsNumeric(Text8.Text) = False And Len(Text8.Text) >= 1) Or Val(Text8.Text) > 100 Then
MsgBox "Enter only number"
Text8.Text = ""
End If
Text22.Text = Val(Text9.Text) + Val(Text10.Text) _
+ Val(Text11.Text) + Val(Text12.Text) + Val(Text13.Text) + Val(Text14.Text) + Val(Text15.Text) + Val(Text16.Text)

Text23.Text = Val(Text7.Text) + Val(Text8.Text) + Val(Text22.Text)
End Sub

Private Sub Text9_Change()
If IsNumeric(Text9.Text) = False And Len(Text9.Text) >= 1 Then
MsgBox "Enter only number"
Text9.Text = ""
End If
Text22.Text = Val(Text9.Text) + Val(Text10.Text) _
+ Val(Text11.Text) + Val(Text12.Text) + Val(Text13.Text) + Val(Text14.Text) + Val(Text15.Text) + Val(Text16.Text)

End Sub

Public Function txt0()
Form1.Text1.Enabled = False
Form1.Text2.Enabled = False
Form1.Text3.Enabled = False
Form1.Text6.Enabled = False
'Form1.Text7.Enabled = False
'Form1.Text8.Enabled = False
'Form1.Text9.Enabled = False
'Form1.Text10.Enabled = False
'Form1.Text11.Enabled = False
'Form1.Text12.Enabled = False
'Form1.Text13.Enabled = False
'Form1.Text14.Enabled = False
'Form1.Text15.Enabled = False
'Form1.Text16.Enabled = False
Form1.Text20.Enabled = False
Form1.Text21.Enabled = False
'Form1.Text22.Enabled = False
'Form1.Text23.Enabled = False
End Function

Public Function txt1()
Form1.Text1.Enabled = True
Form1.Text2.Enabled = True
Form1.Text3.Enabled = True
Form1.Text6.Enabled = True
'Form1.Text7.Enabled = True
'Form1.Text8.Enabled = True
'Form1.Text9.Enabled = True
'Form1.Text10.Enabled = True
'Form1.Text11.Enabled = True
'Form1.Text12.Enabled = True
'Form1.Text13.Enabled = True
'Form1.Text14.Enabled = True
'Form1.Text15.Enabled = True
'Form1.Text16.Enabled = True
Form1.Text20.Enabled = True
Form1.Text21.Enabled = True
'Form1.Text22.Enabled = True
'Form1.Text23.Enabled = True
End Function
