VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form cleanup 
   BackColor       =   &H00C0C0C0&
   Caption         =   "People Management"
   ClientHeight    =   7380
   ClientLeft      =   135
   ClientTop       =   1440
   ClientWidth     =   11700
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7380
   ScaleWidth      =   11700
   Begin VB.Frame mergeframe 
      BackColor       =   &H000000C0&
      Caption         =   "MERGE NAME CLEANUP"
      ForeColor       =   &H0000FFFF&
      Height          =   6015
      Left            =   9360
      TabIndex        =   89
      Top             =   7200
      Visible         =   0   'False
      Width           =   5655
      Begin VB.ListBox packagelist 
         Height          =   1425
         Left            =   1320
         MultiSelect     =   1  'Simple
         Sorted          =   -1  'True
         TabIndex        =   103
         Top             =   3360
         Width           =   3015
      End
      Begin VB.CommandButton Command15 
         BackColor       =   &H000000C0&
         Caption         =   "MERGE NAMES"
         Height          =   495
         Left            =   120
         TabIndex        =   102
         Top             =   5400
         Width           =   5415
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H0000FFFF&
         Height          =   1860
         Left            =   120
         TabIndex        =   91
         Top             =   1200
         Width           =   5415
         Begin VB.TextBox mlname 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1080
            MaxLength       =   60
            TabIndex        =   100
            Top             =   645
            Width           =   4150
         End
         Begin VB.TextBox mfname 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1080
            MaxLength       =   60
            TabIndex        =   98
            Top             =   240
            Width           =   4150
         End
         Begin VB.TextBox midnumber 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3480
            MaxLength       =   20
            TabIndex        =   94
            Top             =   1320
            Width           =   1575
         End
         Begin VB.TextBox mssn 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1680
            MaxLength       =   11
            TabIndex        =   93
            Top             =   1320
            Width           =   1455
         End
         Begin VB.TextBox mbirthdate 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   120
            MaxLength       =   10
            TabIndex        =   92
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label Label39 
            BackStyle       =   0  'Transparent
            Caption         =   "Correct Name L,F"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   495
            Left            =   120
            TabIndex        =   101
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label38 
            BackStyle       =   0  'Transparent
            Caption         =   "Correct Name F L"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   615
            Left            =   120
            TabIndex        =   99
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label Label37 
            BackStyle       =   0  'Transparent
            Caption         =   "ID Number:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   375
            Left            =   3480
            TabIndex        =   97
            Top             =   1050
            Width           =   1815
         End
         Begin VB.Label Label35 
            BackStyle       =   0  'Transparent
            Caption         =   "Social Security #:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   375
            Left            =   1680
            TabIndex        =   96
            Top             =   1080
            Width           =   2190
         End
         Begin VB.Label Label34 
            BackStyle       =   0  'Transparent
            Caption         =   "Birthdate"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   375
            Left            =   90
            TabIndex        =   95
            Top             =   1095
            Width           =   2190
         End
      End
      Begin VB.Label mergestatus 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   104
         Top             =   4920
         Width           =   5415
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   $"cleanup.frx":0000
         ForeColor       =   &H0000FFFF&
         Height          =   855
         Left            =   120
         TabIndex        =   90
         Top             =   480
         Width           =   5415
      End
   End
   Begin Crystal.CrystalReport REPORT 
      Left            =   3240
      Top             =   7080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame rapsheetframe 
      BackColor       =   &H00800000&
      Caption         =   "RAP SHEET"
      ForeColor       =   &H0000FFFF&
      Height          =   6855
      Left            =   960
      TabIndex        =   66
      Top             =   360
      Visible         =   0   'False
      Width           =   11490
      Begin VB.TextBox DISPOSITIONDESCRIPTION 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7800
         MaxLength       =   50
         TabIndex        =   73
         Top             =   5760
         Width           =   2500
      End
      Begin VB.CommandButton Command13 
         BackColor       =   &H0080FFFF&
         Caption         =   "&CLEAR"
         Height          =   495
         Left            =   5880
         Style           =   1  'Graphical
         TabIndex        =   88
         Top             =   6240
         Width           =   1215
      End
      Begin VB.CommandButton Command12 
         BackColor       =   &H0080FFFF&
         Caption         =   "CLO&SE"
         Height          =   495
         Left            =   10200
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   6240
         Width           =   1215
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H0080FFFF&
         Caption         =   "&PRINT"
         Height          =   495
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   6240
         Width           =   1215
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H0080FFFF&
         Caption         =   "&DELETE"
         Height          =   495
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   6240
         Width           =   1215
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H0080FFFF&
         Caption         =   "&CHANGE"
         Height          =   495
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   6240
         Width           =   1215
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H0080FFFF&
         Caption         =   "&ADD"
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   6240
         Width           =   1215
      End
      Begin VB.TextBox dispositiondate 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   10380
         MaxLength       =   10
         TabIndex        =   74
         Top             =   4680
         Width           =   1035
      End
      Begin VB.ListBox disposition 
         Height          =   1035
         Left            =   7800
         TabIndex        =   72
         Top             =   4680
         Width           =   2500
      End
      Begin VB.TextBox charge 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4440
         MaxLength       =   100
         TabIndex        =   71
         Top             =   4680
         Width           =   3195
      End
      Begin VB.TextBox warrantnumber 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2880
         MaxLength       =   20
         TabIndex        =   70
         Top             =   4680
         Width           =   1515
      End
      Begin VB.TextBox casenumber 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   69
         Top             =   4680
         Width           =   1515
      End
      Begin VB.TextBox arrestdate 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         MaxLength       =   10
         TabIndex        =   68
         Top             =   4680
         Width           =   1155
      End
      Begin MSComctlLib.ListView rapsheet 
         Height          =   3975
         Left            =   75
         TabIndex        =   67
         Top             =   360
         Width           =   11340
         _ExtentX        =   20003
         _ExtentY        =   7011
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Arrest Date"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Case#"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Warrant#"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Charge"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Disposition"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Description"
            Object.Width           =   2558
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Date"
            Object.Width           =   1940
         EndProperty
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   10380
         TabIndex        =   85
         Top             =   4440
         Width           =   1095
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Disposition"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   7800
         TabIndex        =   84
         Top             =   4440
         Width           =   1095
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Charge"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   4440
         TabIndex        =   83
         Top             =   4440
         Width           =   1095
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "Warrant#"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   2880
         TabIndex        =   82
         Top             =   4440
         Width           =   1095
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Case#"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   1320
         TabIndex        =   81
         Top             =   4440
         Width           =   1095
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Arrest Date"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   80
         Top             =   4440
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&RAP SHEET"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   6960
      Width           =   1275
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&LINE UP"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   6960
      Width           =   1275
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&UPDATE OPEN SCREENS"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   6975
      Width           =   2350
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   3660
      Top             =   -270
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C0C0&
      Height          =   1755
      Left            =   5520
      TabIndex        =   60
      Top             =   5640
      Width           =   3735
      Begin VB.CommandButton Command5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Associate with Picture File"
         Height          =   660
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   960
         Width           =   1620
      End
      Begin VB.Image mugshot 
         BorderStyle     =   1  'Fixed Single
         Height          =   1485
         Left            =   120
         Stretch         =   -1  'True
         Top             =   120
         Width           =   1800
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Height          =   2340
      Left            =   0
      TabIndex        =   56
      Top             =   4560
      Width           =   5415
      Begin VB.TextBox fbinumber 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2400
         MaxLength       =   20
         TabIndex        =   9
         Top             =   1950
         Width           =   1950
      End
      Begin VB.TextBox birthplace 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2430
         MaxLength       =   50
         TabIndex        =   4
         Top             =   345
         Width           =   2790
      End
      Begin VB.TextBox birthdate 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         MaxLength       =   10
         TabIndex        =   3
         Top             =   360
         Width           =   1950
      End
      Begin VB.TextBox DLSTATE 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1650
         MaxLength       =   3
         TabIndex        =   6
         Top             =   1185
         Width           =   465
      End
      Begin VB.TextBox ssn 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1650
         MaxLength       =   11
         TabIndex        =   5
         Top             =   780
         Width           =   1950
      End
      Begin VB.TextBox dl 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2295
         MaxLength       =   20
         TabIndex        =   7
         Top             =   1185
         Width           =   1305
      End
      Begin VB.TextBox idnumber 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         MaxLength       =   20
         TabIndex        =   8
         Top             =   1920
         Width           =   1950
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "FBI Number:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   2400
         TabIndex        =   87
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "Birthplace"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   2400
         TabIndex        =   86
         Top             =   120
         Width           =   2190
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Birthdate"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   90
         TabIndex        =   61
         Top             =   135
         Width           =   2190
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Social Security #:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   90
         TabIndex        =   59
         Top             =   765
         Width           =   2190
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Driver's License State/Number:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   480
         Left            =   105
         TabIndex        =   58
         Top             =   1110
         Width           =   1560
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "ID Number:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   120
         TabIndex        =   57
         Top             =   1650
         Width           =   1815
      End
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&PRINT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   6105
      Width           =   2350
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Height          =   3390
      Left            =   5475
      TabIndex        =   44
      Top             =   2160
      Width           =   6180
      Begin VB.ListBox resident 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1110
         ItemData        =   "cleanup.frx":00D9
         Left            =   75
         List            =   "cleanup.frx":00E0
         Sorted          =   -1  'True
         TabIndex        =   20
         Top             =   375
         Width           =   1110
      End
      Begin VB.TextBox age 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   660
         MaxLength       =   4
         TabIndex        =   24
         Top             =   1620
         Width           =   735
      End
      Begin VB.TextBox ht 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2520
         MaxLength       =   10
         TabIndex        =   25
         Top             =   1620
         Width           =   1170
      End
      Begin VB.TextBox wt 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4740
         MaxLength       =   10
         TabIndex        =   26
         Top             =   1620
         Width           =   1335
      End
      Begin VB.TextBox hair 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   660
         MaxLength       =   10
         TabIndex        =   27
         Top             =   1995
         Width           =   2385
      End
      Begin VB.TextBox eyes 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3990
         MaxLength       =   10
         TabIndex        =   28
         Top             =   2010
         Width           =   2100
      End
      Begin VB.TextBox peculiarities 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   60
         MaxLength       =   50
         TabIndex        =   29
         Top             =   2550
         Width           =   6000
      End
      Begin VB.ListBox RACE 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1110
         Left            =   1245
         Sorted          =   -1  'True
         TabIndex        =   21
         Top             =   375
         Width           =   1950
      End
      Begin VB.ListBox SEX 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1110
         Left            =   3255
         Sorted          =   -1  'True
         TabIndex        =   22
         Top             =   375
         Width           =   930
      End
      Begin VB.ListBox ETHNICITY 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1110
         Left            =   4245
         Sorted          =   -1  'True
         TabIndex        =   23
         Top             =   375
         Width           =   1815
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Resident"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   60
         TabIndex        =   54
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Race"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1260
         TabIndex        =   53
         Top             =   105
         Width           =   1815
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Sex"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   3225
         TabIndex        =   52
         Top             =   135
         Width           =   1815
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Age:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   60
         TabIndex        =   51
         Top             =   1620
         Width           =   600
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Ethnicity"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   4275
         TabIndex        =   50
         Top             =   135
         Width           =   1815
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Height:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   1695
         TabIndex        =   49
         Top             =   1620
         Width           =   1815
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Weight:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   3900
         TabIndex        =   48
         Top             =   1620
         Width           =   1815
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Hair:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   45
         TabIndex        =   47
         Top             =   2010
         Width           =   1815
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Eyes:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   3390
         TabIndex        =   46
         Top             =   2010
         Width           =   1815
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Peculiarities:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   45
         TabIndex        =   45
         Top             =   2325
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Height          =   2130
      Left            =   5475
      TabIndex        =   39
      Top             =   -45
      Width           =   6180
      Begin VB.TextBox wzipcode 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3075
         MaxLength       =   10
         TabIndex        =   18
         Top             =   1740
         Width           =   800
      End
      Begin VB.TextBox wstate 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2535
         MaxLength       =   2
         TabIndex        =   17
         Top             =   1740
         Width           =   450
      End
      Begin VB.TextBox hzipcode 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3045
         MaxLength       =   10
         TabIndex        =   13
         Top             =   735
         Width           =   800
      End
      Begin VB.TextBox hstate 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2520
         MaxLength       =   2
         TabIndex        =   12
         Top             =   735
         Width           =   450
      End
      Begin VB.TextBox haddress 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   60
         MaxLength       =   60
         TabIndex        =   10
         Top             =   390
         Width           =   5865
      End
      Begin VB.TextBox haddress2 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   60
         MaxLength       =   30
         TabIndex        =   11
         Top             =   735
         Width           =   2400
      End
      Begin VB.TextBox hphone 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4525
         MaxLength       =   20
         TabIndex        =   14
         Top             =   735
         Width           =   1400
      End
      Begin VB.TextBox waddress 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   60
         MaxLength       =   60
         TabIndex        =   15
         Top             =   1350
         Width           =   5865
      End
      Begin VB.TextBox waddress2 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   60
         MaxLength       =   30
         TabIndex        =   16
         Top             =   1740
         Width           =   2400
      End
      Begin VB.TextBox wphone 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4525
         MaxLength       =   20
         TabIndex        =   19
         Top             =   1740
         Width           =   1400
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Home Address"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   60
         TabIndex        =   43
         Top             =   135
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Phone:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   3885
         TabIndex        =   42
         Top             =   750
         Width           =   1815
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Work Address"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   60
         TabIndex        =   41
         Top             =   1110
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Phone:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   3915
         TabIndex        =   40
         Top             =   1785
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   4545
      Left            =   15
      TabIndex        =   35
      Top             =   -30
      Width           =   5415
      Begin MSComctlLib.ListView NAMELIST 
         Height          =   2640
         Left            =   60
         TabIndex        =   62
         Top             =   360
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   4657
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   3616
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "SSN"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ID#"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Birthdate"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Width           =   2
         EndProperty
      End
      Begin VB.TextBox ALIAS 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1215
         MaxLength       =   60
         TabIndex        =   2
         Top             =   4065
         Width           =   4150
      End
      Begin VB.TextBox fname 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1215
         MaxLength       =   60
         TabIndex        =   0
         Top             =   3120
         Width           =   4150
      End
      Begin VB.TextBox lname 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1215
         MaxLength       =   60
         TabIndex        =   1
         Top             =   3585
         Width           =   4150
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Alias"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   135
         TabIndex        =   55
         Top             =   4050
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Correct Name F L"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   615
         Left            =   135
         TabIndex        =   37
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Correct Name L,F"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   495
         Left            =   120
         TabIndex        =   36
         Top             =   3540
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&DELETE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   9330
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   6555
      Width           =   2350
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&SAVE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   9330
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   5640
      Width           =   2350
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&MERGE NAMES"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   6960
      Width           =   1515
   End
End
Attribute VB_Name = "cleanup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mugshotfile As String, painted As Integer, inpeople(999) As String, inidx As Long, itmx As ListItem

Private Sub ARRESTDATE_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
    If Len(arrestdate) = 1 Or Len(arrestdate) = 4 Then
        Call sendslash
    End If
End If

End Sub

Private Sub birthdate_Change()
If IsDate(birthdate) Then
    age = DateDiff("yyyy", CDate(birthdate), CDate(Date$))
End If
End Sub

Private Sub birthdate_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
    If Len(birthdate) = 1 Or Len(birthdate) = 4 Then
        Call sendslash
    End If
End If
End Sub

Private Sub Command1_Click()
If frmLogin.rmsedit = "0" And frmLogin.rmssupervisor = "0" And frmLogin.SUPERVISOR = "0" Then
    MsgBox "Insufficient access authority.", 48, "Genesis Error Log"
    Exit Sub
End If
If UCase(frmLogin.txtUserName) = "DEMO" And UCase(frmLogin.txtPassword) = "DEMO" Then
    msg = MsgBox("Not available in DEMO version.", 48, "Genesis Information Log")
    Screen.MousePointer = 0
    Exit Sub
End If
If NAMELIST.SelectedItem Is Nothing Then
    Exit Sub
End If
Screen.MousePointer = 11
Dim DB As Database, RS As Recordset
On Error GoTo oderror
od:
Set DB = OpenDatabase(nwl + "lawsuite.mdb")
Set itmx = NAMELIST.ListItems(NAMELIST.SelectedItem.index)
ssql = ""
If ssn > "" Then
    ssql = ssql + " and ssn = '" + ssn + "'"
End If
If idnumber > "" Then
    ssql = ssql + " and idnumber = '" + idnumber + "'"
End If
If IsDate(birthdate) Then
    ssql = ssql + " and birthdate = #" + birthdate + "#"
End If
Set RS = DB.OpenRecordset("select * from people where dpnamelf = " + Chr$(34) + itmx + Chr$(34) + ssql)
If Not RS.EOF Then
    RS.MoveFirst
    RS.Edit
Else
    RS.AddNew
End If
On Error Resume Next
RS("dpname") = fname
RS("dpnamelf") = lname
RS("dphaddress") = haddress
RS("dphaddress2") = haddress2
RS("hstate") = hstate
RS("hzipcode") = hzipcode
RS("wstate") = wstate
RS("wzipcode") = wzipcode
RS("dphphone") = hphone
RS("dpwaddress") = waddress
RS("dpwaddress2") = waddress2
RS("dpwphone") = wphone
If resident.ListIndex > -1 Then
    RS("resident") = resident.List(resident.ListIndex)
End If
If race.ListIndex > -1 Then
    RS("race") = race.List(race.ListIndex)
End If
If sex.ListIndex > -1 Then
    RS("sex") = sex.List(sex.ListIndex)
End If
If ethnicity.ListIndex > -1 Then
    RS("ethnicity") = ethnicity.List(ethnicity.ListIndex)
End If
RS("age") = age
RS("height") = ht
RS("weight") = wt
RS("hair") = hair
RS("dl") = dl
RS("dlstate") = DLSTATE
RS("ssn") = ssn
If IsDate(birthdate) Then
    RS("birthdate") = birthdate
Else
    RS("birthdate") = Null
End If
RS("idnumber") = idnumber
RS("fbinumber") = fbinumber
RS("birthplace") = birthplace
RS("alias") = alias
If mugshotfile > "" Then
    RS("mugshot") = mugshotfile
Else
    RS("mugshot") = nwl + itmx + ssn + idnumber + ".bmp"
    SavePicture mugshot.Picture, RS("mugshot")
End If
RS("eyes") = eyes
RS("peculiarities") = peculiarities
RS.Update
DB.Close
If fname <> itmx Then
    Call loadnames
End If
Call clearroutine
Screen.MousePointer = 0
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If
End Sub

Private Sub Command10_Click()
msg = MsgBox("Are you sure you want to delete this line?", 4, "Genesis Information Log")
If msg <> 6 Then
    Exit Sub
End If
If Not IsDate(arrestdate) Or (casenumber = "" And warrantnumber = "") Or charge = "" Then
    MsgBox "Arrest Date, Case or Warrant #, and Charge must be entered.", 48, "Genesis Error Log"
    Exit Sub
End If
If rapsheet.SelectedItem Is Nothing Then
    MsgBox "An item must be selected to change.", 48, "Genesis Error Log"
    Exit Sub
End If
Set itmx = rapsheet.ListItems(rapsheet.SelectedItem.index)
Dim DB As Database, RS As Recordset
Set DB = OpenDatabase(nwl + "rapsheet.mdb")
ssql = ""
If ssn > "" Then
    ssql = ssql + " and ssn = '" + ssn + "'"
End If
If idnumber > "" Then
    ssql = ssql + " and idnumber = '" + idnumber + "'"
End If
If IsDate(birthdate) Then
    ssql = ssql + " and birthdate = #" + birthdate + "#"
End If
If itmx.SubItems(1) > "" Then
    ssql = ssql + " and casenumber = " + Chr$(34) + itmx.SubItems(1) + Chr$(34)
End If
If itmx.SubItems(2) > "" Then
    ssql = ssql + " and warrantnumber = " + Chr$(34) + itmx.SubItems(2) + Chr$(34)
End If
Set RS = DB.OpenRecordset("select * from rapsheet where lname = " + Chr$(34) + lname + Chr$(34) + " " + ssql + " and arrestdate = #" + itmx + "# and charge = " + Chr$(34) + itmx.SubItems(3) + Chr$(34))
If Not RS.EOF Then
    RS.MoveFirst
    RS.Delete
End If
DB.Close
rapsheet.ListItems.Remove rapsheet.SelectedItem.index
Call CLEARRAP

End Sub

Private Sub Command11_Click()
inp = InputBox("Enter S for rap sheet or L for letter.", "Genesis Information Log", "S")
inp = UCase(inp)
If inp <> "S" And inp <> "L" Then
    Exit Sub
End If
If inp = "S" Then
    If rapsheet.ListItems.Count = 0 Then
        MsgBox "There are no charges for this individual.  A rap sheet cannot be generated.", 48, "Genesis Error Log"
        Exit Sub
    End If
End If
idtag = lname + Space$(60 - Len(lname))
If Not IsNull(birthdate) Then
    idtag = idtag + Format$(birthdate, "mmddyyyy")
Else
    idtag = idtag + Space$(8)
End If
If Not IsNull(idnumber) Then
    idtag = idtag + idnumber + Space$(20 - Len(idnumber))
Else
    idtag = idtag + Space$(20)
End If
If Not IsNull(ssn) Then
    idtag = idtag + ssn + Space$(11 - Len(ssn))
Else
    idtag = idtag + Space$(11)
End If
Dim DB As Database, RS As Recordset
Set DB = OpenDatabase(nwl + "LAWSUITE.MDB")
ssql = ""
If ssn > "" Then
    ssql = ssql + " and ssn = '" + ssn + "'"
End If
If idnumber > "" Then
    ssql = ssql + " and idnumber = '" + idnumber + "'"
End If
If IsDate(birthdate) Then
    ssql = ssql + " and birthdate = #" + birthdate + "#"
End If
Set DB = OpenDatabase(nwl + "lawsuite.mdb")
Set RS = DB.OpenRecordset("select * from people where dpnamelf = " + Chr$(34) + lname + Chr$(34) + ssql)
If Not RS.EOF Then
    RS.Edit
    RS("IDTAG") = idtag
    RS.Update
End If
Set DB = OpenDatabase(nwl + "rapsheet.mdb")
Set RS = DB.OpenRecordset("select * from rapsheet where lname = " + Chr$(34) + lname + Chr$(34) + ssql)
While Not RS.EOF
    RS.Edit
    RS("IDTAG") = idtag
    RS.Update
    RS.MoveNext
Wend
RS.Close
DB.Close
report.SelectionFormula = ""
report.SelectionFormula = "{PEOPLE.idtag} = " + Chr$(34) + idtag + Chr$(34) + " AND NOT ISNULL({RAPSHEET.DISPOSITION})"
If inp = "S" Then
    report.ReportFileName = nwl + "RAPSHEET.RPT"
Else
    report.ReportFileName = nwl + "RAPletter.RPT"
End If
report.Action = 1
End Sub

Private Sub Command12_Click()
rapsheetframe.Visible = False
End Sub

Private Sub Command13_Click()
Call CLEARRAP
End Sub

Private Sub Command14_Click()
If frmLogin.SUPERVISOR = "0" Then
    MsgBox "Insufficient access authority.", 48, "Genesis Error Log"
    Exit Sub
End If
If NAMELIST.SelectedItem Is Nothing Then
    MsgBox "All names to merge must be highlighted in the NAME listbox.", 48, "Genesis Error Log"
    Exit Sub
End If
packagelist.clear
packagelist.AddItem "People Management Records"
If nwb > "" Then
    packagelist.AddItem "Genesis Booking Report"
End If
If nwi > "" Then
    packagelist.AddItem "Genesis Incident Report"
End If
If nws > "" Then
    packagelist.AddItem "Genesis Service Call"
End If
If nwl > "" Then
    packagelist.AddItem "Genesis Rap Sheet"
End If
If nwr > "" Then
    packagelist.AddItem "Genesis Restraining Order"
End If
If nww > "" Then
    packagelist.AddItem "Genesis Warrant Manager"
End If
mergeframe.Left = 5500
mergeframe.Top = 360
mergeframe.Visible = True
End Sub

Private Sub Command15_Click()
Dim DB As Database, RS As Recordset, tfld(9) As String, nfld(9) As String, tidx As Long, tsql As String
inidx = 0
If mlname = "" Or mfname = "" Then
    MsgBox "Both F/L and L/F Names must be entered.", 48, "Genesis Error Log"
    Exit Sub
End If
For t = 1 To NAMELIST.ListItems.Count
    If NAMELIST.ListItems(t).Selected Then
        inidx = inidx + 1
        inpeople(inidx) = NAMELIST.ListItems(t)
    End If
Next t
For t = 0 To packagelist.ListCount - 1
    If packagelist.Selected(t) Then
        Select Case packagelist.List(t)
            Case "People Management Records"

                tbname = "people"
                dbname = nwl + "lawsuite.mdb"
                tidx = 5
                tfld(1) = "dpnamelf"
                tfld(2) = "dpname"
                tfld(3) = "ssn"
                tfld(4) = "birthdate"
                tfld(5) = "idnumber"
                nfld(1) = mlname
                nfld(2) = mfname
                nfld(3) = mssn
                nfld(4) = mbirthdate
                nfld(5) = midnumber
                Call runtsql(CStr(dbname), CStr(tbname), tidx, tfld(), nfld())
                Call loadnames
               
            Case "Genesis Incident Report"
                
                dbname = nwi + "incident.mdb"
                tbname = "badcheck"
                tidx = 1
                tfld(1) = "cname"
                nfld(1) = mlname
                Call runtsql(CStr(dbname), CStr(tbname), tidx, tfld(), nfld())
                tfld(1) = "vname"
                Call runtsql(CStr(dbname), CStr(tbname), tidx, tfld(), nfld())
                tdx = 4
                tfld(1) = "sname"
                tfld(2) = "ssn"
                tfld(3) = "sbirthdate"
                tfld(4) = "idnumber"
                nfld(1) = mlname
                nfld(2) = mssn
                nfld(3) = mbirthdate
                nfld(4) = midnumber
                Call runtsql(CStr(dbname), CStr(tbname), tidx, tfld(), nfld())
                
                tbname = "noncriminal"
                tidx = 1
                tfld(1) = "cname"
                nfld(1) = mlname
                Call runtsql(CStr(dbname), CStr(tbname), tidx, tfld(), nfld())
                tfld(1) = "driver0"
                Call runtsql(CStr(dbname), CStr(tbname), tidx, tfld(), nfld())
                tfld(1) = "driver1"
                Call runtsql(CStr(dbname), CStr(tbname), tidx, tfld(), nfld())
                tfld(1) = "owner0"
                Call runtsql(CStr(dbname), CStr(tbname), tidx, tfld(), nfld())
                tfld(1) = "owner1"
                Call runtsql(CStr(dbname), CStr(tbname), tidx, tfld(), nfld())
                                
                tbname = "incidentreportc"
                tidx = 1
                tfld(1) = "cname"
                nfld(1) = mlname
                Call runtsql(CStr(dbname), CStr(tbname), tidx, tfld(), nfld())
            
                tbname = "incidentreportv"
                tidx = 1
                tfld(1) = "vname"
                nfld(1) = mlname
                Call runtsql(CStr(dbname), CStr(tbname), tidx, tfld(), nfld())
            
                tbname = "incidentreports"
                tidx = 2
                tfld(1) = "sname"
                nfld(1) = mlname
                tfld(2) = "sbirthdate"
                nfld(2) = mbirthdate
                Call runtsql(CStr(dbname), CStr(tbname), tidx, tfld(), nfld())
            
                tbname = "supplemental"
                tidx = 2
                tfld(1) = "name1"
                nfld(1) = mlname
                tfld(2) = "birthdate1"
                nfld(2) = mbirthdate
                Call runtsql(CStr(dbname), CStr(tbname), tidx, tfld(), nfld())
            
                tfld(1) = "name2"
                tfld(2) = "birthdate2"
                Call runtsql(CStr(dbname), CStr(tbname), tidx, tfld(), nfld())
            
            Case "Genesis Booking Report"
                dbname = nwb + "booking.mdb"
                tbname = "booking"
                tidx = 4
                tfld(1) = "sname"
                nfld(1) = mlname
                tfld(2) = "ssn"
                nfld(2) = mssn
                tfld(3) = "idnumber"
                nfld(3) = midnumber
                tfld(4) = "sbirthdate"
                nfld(4) = mbirthdate
                Call runtsql(CStr(dbname), CStr(tbname), tidx, tfld(), nfld())
            
            Case "Genesis Service Call"
                dbname = nws + "service.mdb"
                tbname = "service"
                tidx = 1
                tfld(1) = "compsubj"
                nfld(1) = mlname
                Call runtsql(CStr(dbname), CStr(tbname), tidx, tfld(), nfld())
                        
            Case "Genesis Rap Sheet"
                dbname = nwl + "rapsheet.mdb"
                tbname = "rapsheet"
                tidx = 4
                tfld(1) = "lname"
                nfld(1) = mlname
                tfld(2) = "ssn"
                nfld(2) = mssn
                tfld(3) = "idnumber"
                nfld(3) = midnumber
                tfld(4) = "birthdate"
                nfld(4) = mbirthdate
                Call runtsql(CStr(dbname), CStr(tbname), tidx, tfld(), nfld())
                        
            Case "Genesis Restraining Order"
                dbname = nwr + "ro.mdb"
                tbname = "rorder"
                tidx = 1
                tfld(1) = "plaintiff"
                nfld(1) = mlname
                Call runtsql(CStr(dbname), CStr(tbname), tidx, tfld(), nfld())
                tfld(1) = "defendant"
                Call runtsql(CStr(dbname), CStr(tbname), tidx, tfld(), nfld())
                                    
            Case "Genesis Warrant Manager"
                dbname = nww + "warrant.mdb"
                tbname = "warrantinfo"
                tidx = 4
                tfld(1) = "wname"
                nfld(1) = mlname
                tfld(2) = "birthdate"
                nfld(2) = mbirthdate
                tfld(3) = "ssn"
                nfld(3) = mssn
                tfld(4) = "idnumber"
                nfld(4) = midnumber
                Call runtsql(CStr(dbname), CStr(tbname), tidx, tfld(), nfld())
                tidx = 1
                tfld(1) = "plaintiff"
                Call runtsql(CStr(dbname), CStr(tbname), tidx, tfld(), nfld())
                tfld(1) = "defendant"
                Call runtsql(CStr(dbname), CStr(tbname), tidx, tfld(), nfld())
        End Select
    End If
Next t
MsgBox "Name Merge completed successfully.", 48, "Genesis Information Log"
mergeframe.Visible = False

End Sub

Private Sub Command2_Click()
If frmLogin.rmsdelete = "0" And frmLogin.rmssupervisor = "0" And frmLogin.SUPERVISOR = "0" Then
    MsgBox "Insufficient access authority.", 48, "Genesis Error Log"
    Exit Sub
End If
If UCase(frmLogin.txtUserName) = "DEMO" And UCase(frmLogin.txtPassword) = "DEMO" Then
    msg = MsgBox("Not available in DEMO version.", 48, "Genesis Information Log")
    Screen.MousePointer = 0
    Exit Sub
End If
If NAMELIST.SelectedItem Is Nothing Then
    Exit Sub
End If
Set itmx = NAMELIST.ListItems(NAMELIST.SelectedItem.index)
msg = MsgBox("Are You Sure?", 4, "Genesis Information Log")
If msg <> 6 Then
    Exit Sub
End If
Screen.MousePointer = 11
Dim DB As Database, RS As Recordset
On Error GoTo oderror
od:
ssql = ""
If ssn > "" Then
    ssql = ssql + " and ssn = '" + ssn + "'"
End If
If idnumber > "" Then
    ssql = ssql + " and idnumber = '" + idnumber + "'"
End If
If IsDate(birthdate) Then
    ssql = ssql + " and birthdate = #" + birthdate + "#"
End If
Set DB = OpenDatabase(nwl + "lawsuite.mdb")
Set RS = DB.OpenRecordset("select * from people where dpnamelf = " + Chr$(34) + itmx + Chr$(34) + ssql)
If Not RS.EOF Then
    RS.MoveFirst
    RS.Delete
End If
On Error Resume Next
DB.Close
NAMELIST.ListItems.Remove NAMELIST.SelectedItem.index
'Call clearroutine
'Call loadnames
Screen.MousePointer = 0
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If
End Sub

Private Sub Command3_Click()
If frmLogin.rmsprint = "0" And frmLogin.rmssupervisor = "0" And frmLogin.SUPERVISOR = "0" Then
    MsgBox "Insufficient access authority.", 48, "Genesis Error Log"
    Exit Sub
End If
Printer.FontName = "Times New Roman"
Printer.FontSize = 18
Printer.FontBold = True
Printer.Print "PERSONAL INFORMATION SHEET";
Printer.FontSize = 12
Printer.Print Tab(100); Date$
Printer.FontBold = False
Printer.Print
Printer.Print
Printer.FontBold = True
Printer.Print "NAME:";
Printer.FontBold = False
Printer.Print Tab(30); fname
Printer.FontBold = True
Printer.Print "ALIAS:";
Printer.FontBold = False
Printer.Print Tab(30); alias
Printer.Print
Printer.FontBold = True
Printer.Print "HOME ADDRESS:";
Printer.FontBold = False
Printer.Print Tab(30); haddress
If haddress2 > "" Then
    Printer.Print Tab(30); haddress2 + " " + hstate + " " + hzipcode
End If
Printer.FontBold = True
Printer.Print "HOME PHONE:";
Printer.FontBold = False
Printer.Print Tab(30); hphone
Printer.FontBold = True
Printer.Print "WORK ADDRESS:";
Printer.FontBold = False
Printer.Print Tab(30); waddress
If waddress2 > "" Then
    Printer.Print Tab(30); waddress2 + " " + wstate + " " + wzipcode
End If
Printer.FontBold = True
Printer.Print "WORK PHONE:";
Printer.FontBold = False
Printer.Print Tab(30); wphone
Printer.Print
Printer.FontBold = True
Printer.Print "RESIDENT STATUS:";
Printer.FontBold = False
If resident.ListIndex > -1 Then
    Printer.Print Tab(30); resident.List(resident.ListIndex)
Else
    Printer.Print
End If
Printer.FontBold = True
Printer.Print "RACE:";
Printer.FontBold = False
If race.ListIndex > -1 Then
    Printer.Print Tab(30); race.List(race.ListIndex)
Else
    Printer.Print
End If
Printer.FontBold = True
Printer.Print "SEX:";
Printer.FontBold = False
If sex.ListIndex > -1 Then
    Printer.Print Tab(30); sex.List(sex.ListIndex)
Else
    Printer.Print
End If
Printer.FontBold = True
Printer.Print "ETHNICITY:";
Printer.FontBold = False
If ethnicity.ListIndex > -1 Then
    Printer.Print Tab(30); ethnicity.List(ethnicity.ListIndex)
Else
    Printer.Print
End If
Printer.FontBold = True
Printer.Print "AGE:";
Printer.FontBold = False
Printer.Print Tab(30); age
Printer.FontBold = True
Printer.Print "HEIGHT:";
Printer.FontBold = False
Printer.Print Tab(30); ht
Printer.FontBold = True
Printer.Print "WEIGHT:";
Printer.FontBold = False
Printer.Print Tab(30); wt
Printer.FontBold = True
Printer.Print "HAIR:";
Printer.FontBold = False
Printer.Print Tab(30); hair
Printer.FontBold = True
Printer.Print "EYES:";
Printer.FontBold = False
Printer.Print Tab(30); eyes
Printer.FontBold = True
Printer.Print "PECULIARITES:";
Printer.FontBold = False
Printer.Print Tab(30); peculiarities
Printer.Print
Printer.FontBold = True
Printer.Print "DATE OF BIRTH:";
Printer.FontBold = False
Printer.Print Tab(30); birthdate
Printer.FontBold = True
Printer.Print "SSN:";
Printer.FontBold = False
Printer.Print Tab(30); ssn
Printer.FontBold = True
Printer.Print "DRIVER'S LICENSE:";
Printer.FontBold = False
Printer.Print Tab(30); DLSTATE + " " + dl
Printer.FontBold = True
Printer.Print "ID NUMBER:";
Printer.FontBold = False
Printer.Print Tab(30); idnumber
Printer.Print


Printer.PaintPicture mugshot.Picture, Printer.CurrentX, Printer.CurrentY
Printer.EndDoc

End Sub

Private Sub Command4_Click()
lineup.Show
End Sub

Private Sub Command5_Click()
hcd = CurDir
cd.Filter = "Bitmaps (*.bmp)|*.bmp|JPG Files (*.jpg)|*.jpg|GIF Files (*.gif)|*.gif"
cd.ShowOpen
If cd.FileName > "" Then
    mugshot.Picture = LoadPicture(cd.FileName)
    mugshotfile = cd.FileName
End If
ChDir hcd
End Sub

Private Sub Command6_Click()
If frmLogin.rmsedit = "0" And frmLogin.rmssupervisor = "0" And frmLogin.SUPERVISOR = "0" Then
    MsgBox "Insufficient access authority.", 48, "Genesis Error Log"
    Exit Sub
End If
Call Command1_Click
For t% = 0 To Forms.Count - 1
    Select Case LCase(Forms(t%).Name)
        Case "badcheck"
            If badcheck.cname = lname Then
                badcheck.caddress = haddress
                badcheck.CCITY = haddress2
                badcheck.CSTATE = hstate
                badcheck.CZIPCODE = hzipcode
                badcheck.cphone = hphone
            End If
            If badcheck.vname = lname Then
                badcheck.vaddress = haddress + " " + haddress2
                badcheck.vhphone = hphone
                badcheck.vwphone = wphone
                badcheck.vrace.ListIndex = race.ListIndex
                badcheck.vsex.ListIndex = sex.ListIndex
                badcheck.vage = age
            End If
            If badcheck.sname = lname Then
                badcheck.saddress = haddress
                badcheck.scity = haddress2
                badcheck.sstate = hstate
                badcheck.szipcode = hzipcode
                badcheck.srace.ListIndex = race.ListIndex
                badcheck.ssex.ListIndex = sex.ListIndex
                badcheck.sage = age
                badcheck.sheight = ht
                badcheck.sweight = wt
                badcheck.SHAIR = hair
                badcheck.SEYES = eyes
                badcheck.drivers = dl
                badcheck.driversstate = DLSTATE
                badcheck.ssn = ssn
                badcheck.sbirthdate = birthdate
                If mugshotfile > "" Then
                    badcheck.mugshot.Picture = LoadPicture(mugshotfile)
                End If
            End If
        Case "booking"
            If booking.sname = lname Then
                booking.saddress = haddress + " " + haddress2
                booking.srace.ListIndex = race.ListIndex
                booking.ssex.ListIndex = sex.ListIndex
                booking.sage = age
                booking.SHT = ht
                booking.sweight = wt
                booking.SHAIR = hair
                booking.SEYES = eyes
                booking.driverslicense = dl
                booking.driverslicensestate = DLSTATE
                booking.ssn = ssn
                booking.sbirthdate = birthdate
                booking.speculiarities = peculiarities
                booking.idnumber = idnumber
                If mugshotfile > "" Then
                    booking.mugshot.Picture = LoadPicture(mugshotfile)
                End If
            End If
        Case "incident"
            If incident.vsname(2) = lname Then
                incident.address(2) = haddress
                incident.City(2) = haddress2
                incident.State(2) = hstate
                incident.zipcode(2) = hzipcode
                incident.race(2).ListIndex = race.ListIndex
                incident.sex(2).ListIndex = sex.ListIndex
                incident.age(2) = age
                incident.ht(1) = ht
                incident.weight(1) = wt
                incident.hair(1) = hair
                incident.eyes(1) = eyes
                incident.birthdate = birthdate
                incident.peculiarities(1) = peculiarities
                If mugshotfile > "" Then
                    incident.mugshot.Picture = LoadPicture(mugshotfile)
                End If
            End If
        Case "sinciden"
            For yy% = 0 To 1
                If sinciden.vsname(yy%) = lname Then
                    sinciden.address(yy%) = haddress
                    sinciden.City(yy%) = haddress2
                    sinciden.State(yy%) = hstate
                    sinciden.zipcode(yy%) = hzipcode
                    sinciden.race(yy%).ListIndex = race.ListIndex
                    sinciden.sex(yy%).ListIndex = sex.ListIndex
                    sinciden.age(yy%) = age
                    sinciden.ht(yy%) = ht
                    sinciden.weight(yy%) = wt
                    sinciden.hair(yy%) = hair
                    sinciden.eyes(yy%) = eyes
                    sinciden.birthdate(yy%) = birthdate
                    sinciden.peculiarities(yy%) = peculiarities
                    If mugshotfile > "" Then
                        sinciden.mugshot(yy%).Picture = LoadPicture(mugshotfile)
                    End If
                End If
            Next yy%
        Case "noncrim"
            If noncrim.cname = lname Then
                noncrim.caddress = haddress
                noncrim.CCITY = haddress2
                noncrim.CSTATE = hstate
                noncrim.CZIPCODE = hzipcode
            End If
            For yy% = 0 To 1
                If noncrim.driver(yy%) = lname Then
                    noncrim.driveraddress1(yy%) = haddress
                    noncrim.driveraddress2(yy%) = haddress2
                    noncrim.DRIVERSTATE(yy%) = hstate
                    noncrim.DRIVERZIPCODE(yy%) = hzipcode
                    noncrim.driverdl(yy%) = dl
                End If
                If noncrim.owner(yy%) = lname Then
                    noncrim.owneraddress1(yy%) = haddress
                    noncrim.owneraddress2(yy%) = haddress2
                    noncrim.OWNERSTATE(yy%) = hstate
                    noncrim.OWNERZIPCODE(yy%) = hzipcode
                End If
            Next yy%
        Case "ro"
            If ro.plaintiff = lname Then
                ro.plaintiffaddress(0) = haddress
                ro.plaintiffaddress(1) = haddress2
                ro.plaintiffstate = hstate
                ro.plaintiffzipcode = hzipcode
            End If
            If ro.defendant = lname Then
                ro.defendantaddress(0) = haddress
                ro.defendantaddress(1) = haddress2
                ro.defendantstate = hstate
                ro.defendantzipcode = hzipcode
                If mugshotfile > "" Then
                    ro.mugshot.Picture = LoadPicture(mugshotfile)
                End If
            End If
        Case "service"
            If service.compsubj = lname Then
                service.Address1 = haddress
                service.Address2 = haddress2
                service.State = hstate
                service.zipcode = hzipcode
                service.phone = hphone
                If mugshotfile > "" Then
                    service.mugshot.Picture = LoadPicture(mugshotfile)
                End If
            End If
        Case "civil"
            If CIVIL.serviceof = fname Then
                CIVIL.sohomeaddress = haddress
                CIVIL.sohomeaddress2 = haddress2
                CIVIL.sohomestate = hstate
                CIVIL.sohomezipcode = hzipcode
                CIVIL.sohomephone = hphone
                CIVIL.soworkaddress = waddress
                CIVIL.soworkaddress2 = waddress2
                CIVIL.soworkstate = wstate
                CIVIL.soworkzipcode = wzipcode
                CIVIL.soworkphone = wphone
                If mugshotfile > "" Then
                    CIVIL.mugshot.Picture = LoadPicture(mugshotfile)
                End If
            End If
        Case "warrant"
            If warrant.wname = lname Then
                warrant.address(0) = haddress
                warrant.address(1) = haddress2
                warrant.address(2) = hstate
                warrant.address(3) = hzipcode
                warrant.ssn = snn
                warrant.birthdate = birthdate
                warrant.idnumber = idnumber
                warrant.hair = hair
                warrant.eyes = eyes
                warrant.ht = ht
                warrant.weight = wt
                If race.ListIndex > -1 Then
                    If Left$(race.List(race.ListIndex), 1) = "W" Then
                        warrant.caucasian = True
                    End If
                    If Left$(race.List(race.ListIndex), 1) = "B" Then
                        warrant.africanamerican = True
                    End If
                    If Left$(race.List(race.ListIndex), 1) = "O" Then
                        warrant.Oriental = True
                    End If
                    If Left$(race.List(race.ListIndex), 1) = "A" Then
                        warrant.Other = True
                        warrant.otherrace = "Asian"
                    End If
                End If
                If sex.ListIndex > -1 Then
                    If Left$(sex.List(sex.ListIndex), 1) = "F" Then
                        warrant.female = True
                    End If
                    If Left$(sex.List(sex.ListIndex), 1) = "M" Then
                        warrant.male = True
                    End If
                End If
                If mugshotfile > "" Then
                    warrant.mugshot.Picture = LoadPicture(mugshotfile)
                End If
            End If
            Case "frmbookingreport"
                If frmBookingReport.sname = lname Then
                    frmBookingReport.saddress = haddress
                    frmBookingReport.scity = haddress2
                    frmBookingReport.sstate = hstate
                    frmBookingReport.szipcode = hzipcode
                    frmBookingReport.ssn = snn
                    frmBookingReport.sbirthdate = birthdate
                    frmBookingReport.IdNum = idnumber
                    frmBookingReport.SHAIR = hair
                    frmBookingReport.SEYES = eyes
                    frmBookingReport.SHT = ht
                    frmBookingReport.sweight = wt
                    If race.ListIndex > -1 Then
                        frmBookingReport.srace.BoundText = race.List(race.ListIndex)
                    End If
                    If sex.ListIndex > -1 Then
                        frmBookingReport.ssex.BoundText = sex.List(sex.ListIndex)
                    End If
                    If ethnicity.ListIndex > -1 Then
                        frmBookingReport.sethnicity.BoundText = ethnicity.List(ethnicity.ListIndex)
                    End If
                    frmBookingReport.alias = alias
                    If mugshotfile > "" Then
                        frmBookingReport.mugshot.Picture = LoadPicture(mugshotfile)
                    End If
                End If
        End Select
Next t%
Unload Me
End Sub

Private Sub Command7_Click()
If frmLogin.rmsprint = "0" And frmLogin.rmsreport = "0" And frmLogin.rmssupervisor = "0" And frmLogin.SUPERVISOR = "0" Then
    MsgBox "Insufficient access authority.", 48, "Genesis Error Log"
    Exit Sub
End If
If lname = "" Then
    MsgBox "A name must be selected/entered for Correct Name L,F in order to access the Rap Sheet.", 48, "Genesis Error Log"
    Exit Sub
End If
Dim DB As Database, RS As Recordset
Set DB = OpenDatabase(nwl + "rapsheet.mdb")
ssql = ""
If ssn > "" Then
    ssql = ssql + " and ssn = '" + ssn + "'"
End If
If idnumber > "" Then
    ssql = ssql + " and idnumber = '" + idnumber + "'"
End If
If IsDate(birthdate) Then
    ssql = ssql + " and birthdate = #" + birthdate + "#"
End If
Set RS = DB.OpenRecordset("select * from rapsheet where lname = " + Chr$(34) + lname + Chr$(34) + " " + ssql)
rapsheet.ListItems.clear
While Not RS.EOF
    Set itmx = rapsheet.ListItems.add(, , RS("arrestdate"))
    itmx.SubItems(1) = RS("casenumber")
    itmx.SubItems(2) = RS("warrantnumber")
    itmx.SubItems(3) = RS("charge")
    If Not IsNull(RS("disposition")) Then
        itmx.SubItems(4) = RS("disposition")
    End If
    If Not IsNull(RS("dispositiondescription")) Then
        itmx.SubItems(5) = RS("dispositiondescription")
    End If
    If Not IsNull(RS("dispositiondate")) Then
        itmx.SubItems(6) = RS("dispositiondate")
    End If
    RS.MoveNext
Wend
DB.Close
rapsheetframe.Left = 100
rapsheetframe.Top = 500
rapsheetframe.Visible = True
arrestdate.SetFocus
End Sub

Private Sub Command8_Click()
If Not IsDate(arrestdate) Or (casenumber = "" And warrantnumber = "") Or charge = "" Then
    MsgBox "Arrest Date, Case or Warrant #, and Charge must be entered.", 48, "Genesis Error Log"
    Exit Sub
End If
Set itmx = rapsheet.ListItems.add(, , arrestdate)
itmx.SubItems(1) = casenumber
itmx.SubItems(2) = warrantnumber
itmx.SubItems(3) = charge
If disposition.ListIndex > -1 Then
    itmx.SubItems(4) = disposition.List(disposition.ListIndex)
End If
itmx.SubItems(5) = DISPOSITIONDESCRIPTION
itmx.SubItems(6) = dispositiondate
Dim DB As Database, RS As Recordset
Set DB = OpenDatabase(nwl + "rapsheet.mdb")
ssql = ""
If ssn > "" Then
    ssql = ssql + " and ssn = '" + ssn + "'"
End If
If idnumber > "" Then
    ssql = ssql + " and idnumber = '" + idnumber + "'"
End If
If IsDate(birthdate) Then
    ssql = ssql + " and birthdate = #" + birthdate + "#"
End If
Set RS = DB.OpenRecordset("select * from rapsheet where lname = " + Chr$(34) + lname + Chr$(34) + " " + ssql)
RS.AddNew
RS("lname") = lname
RS("ssn") = ssn
RS("idnumber") = idnumber
If IsDate(birthdate) Then
    RS("birthdate") = CDate(birthdate)
End If
RS("arrestdate") = CDate(arrestdate)
RS("casenumber") = casenumber
RS("warrantnumber") = warrantnumber
RS("charge") = charge
If disposition.ListIndex > -1 Then
    RS("disposition") = disposition.List(disposition.ListIndex)
Else
    RS("disposition") = Null
End If
RS("dispositiondescription") = DISPOSITIONDESCRIPTION
If IsDate(dispositiondate) Then
    RS("dispositiondate") = CDate(dispositiondate)
Else
    RS("dispositiondate") = Null
End If
RS.Update
DB.Close
Call CLEARRAP
End Sub

Private Sub Command9_Click()
If Not IsDate(arrestdate) Or (casenumber = "" And warrantnumber = "") Or charge = "" Then
    MsgBox "Arrest Date, Case or Warrant #, and Charge must be entered.", 48, "Genesis Error Log"
    Exit Sub
End If
If rapsheet.SelectedItem Is Nothing Then
    MsgBox "An item must be selected to change.", 48, "Genesis Error Log"
    Exit Sub
End If
Set itmx = rapsheet.ListItems(rapsheet.SelectedItem.index)
Dim DB As Database, RS As Recordset
Set DB = OpenDatabase(nwl + "rapsheet.mdb")
ssql = ""
If ssn > "" Then
    ssql = ssql + " and ssn = '" + ssn + "'"
End If
If idnumber > "" Then
    ssql = ssql + " and idnumber = '" + idnumber + "'"
End If
If IsDate(birthdate) Then
    ssql = ssql + " and birthdate = #" + birthdate + "#"
End If
If itmx.SubItems(1) > "" Then
    ssql = ssql + " and casenumber = " + Chr$(34) + itmx.SubItems(1) + Chr$(34)
End If
If itmx.SubItems(2) > "" Then
    ssql = ssql + " and warrantnumber = " + Chr$(34) + itmx.SubItems(2) + Chr$(34)
End If
Set RS = DB.OpenRecordset("select * from rapsheet where lname = " + Chr$(34) + lname + Chr$(34) + " " + ssql + " and arrestdate = #" + itmx + "# and charge = " + Chr$(34) + itmx.SubItems(3) + Chr$(34))
If Not RS.EOF Then
    RS.MoveFirst
    RS.Edit
    RS("lname") = lname
    RS("ssn") = ssn
    RS("idnumber") = idnumber
    If IsDate(birthdate) Then
        RS("birthdate") = CDate(birthdate)
    End If
    RS("arrestdate") = CDate(arrestdate)
    RS("casenumber") = casenumber
    RS("warrantnumber") = warrantnumber
    RS("charge") = charge
    If disposition.ListIndex > -1 Then
        RS("disposition") = disposition.List(disposition.ListIndex)
    Else
        RS("disposition") = Null
    End If
    RS("dispositiondescription") = DISPOSITIONDESCRIPTION
    If IsDate(dispositiondate) Then
        RS("dispositiondate") = CDate(dispositiondate)
    Else
        RS("dispositiondate") = Null
    End If
    RS.Update
End If
DB.Close
itmx = arrestdate
itmx.SubItems(1) = casenumber
itmx.SubItems(2) = warrantnumber
itmx.SubItems(3) = charge
If disposition.ListIndex > -1 Then
    itmx.SubItems(4) = disposition.List(disposition.ListIndex)
Else
    itmx.SubItems(4) = ""
End If
itmx.SubItems(5) = DISPOSITIONDESCRIPTION
itmx.SubItems(6) = dispositiondate
Call CLEARRAP
End Sub

Private Sub disposition_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    disposition.ListIndex = -1
End If
End Sub

Private Sub dispositiondate_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
If Len(dispositiondate) = 1 Or Len(dispositiondate) = 4 Then
    Call sendslash
End If
End If
End Sub

Private Sub Form_Load()
If Dir(nwl + "rapsheet.mdb") = "" Then
    Command7.Visible = False
End If
disposition.clear
disposition.AddItem "Dismissed"
disposition.AddItem "Guilty"
disposition.AddItem "Not Guilty"
disposition.AddItem "No Contest"
disposition.AddItem "Null Process"
mugshotfile = ""
painted = 0
On Error GoTo 0
cleanup.Top = 0
cleanup.Left = 0
cleanup.Width = 11805
cleanup.Height = 7860
Call loadnames
resident.clear
race.clear
sex.clear
ethnicity.clear
resident.AddItem "J - Jurisdiction"
resident.AddItem "S - State"
resident.AddItem "U - Unknown"
resident.AddItem "O - Out of State"
race.AddItem "White"
race.AddItem "Black"
race.AddItem "Indian - American Indian/Alaskan Native"
race.AddItem "Asian/Pacific Islander"
race.AddItem "Unknown"
sex.AddItem "Male"
sex.AddItem "Female"
sex.AddItem "Unknown"
ethnicity.AddItem "Hispanic Origin"
ethnicity.AddItem "Not of Hispanic Origin"
ethnicity.AddItem "Unknown"
End Sub

Private Sub Text2_Change()

End Sub
Private Sub loadnames()
Screen.MousePointer = 11
Dim DB As Database, RS As Recordset
On Error GoTo oderror
od:
Set DB = OpenDatabase(nwl + "lawsuite.mdb")
Set RS = DB.OpenRecordset("select * from people WHERE DPNAMELF IS NOT NULL AND DPNAME > '' order by dpnamelf")
NAMELIST.ListItems.clear
If Not RS.EOF Then
    RS.MoveFirst
    While Not RS.EOF
        Set itmx = NAMELIST.ListItems.add(, , RS("dpnamelf"))
        If Not IsNull(RS("ssn")) Then
            itmx.SubItems(1) = RS("ssn")
        End If
        If Not IsNull(RS("idnumber")) Then
            itmx.SubItems(2) = RS("idnumber")
        End If
        If Not IsNull(RS("birthdate")) Then
            itmx.SubItems(3) = RS("birthdate")
        End If
        itmx.SubItems(4) = RS("dpnamelf")
        RS.MoveNext
    Wend
End If
While Not NAMELIST.SelectedItem Is Nothing
    NAMELIST.SelectedItem = Nothing
Wend
On Error Resume Next
DB.Close
Screen.MousePointer = 0
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
    Resume
End If
End Sub

Private Sub Form_Paint()
painted = painted + 1
If painted <> 1 Then
    Exit Sub
End If
If fname > "" And lname = "" Then
    For t% = 1 To NAMELIST.ListItems.Count
        If NAMELIST.ListItems(t%) = fname Then
            NAMELIST.ListItems(t%).Selected = True
            Call namelist_ItemClick(itmx)
        Else
            NAMELIST.ListItems(t%).Selected = False
        End If
    Next t%
Else
If lname > "" And fname = "" Then
    For t% = 1 To NAMELIST.ListItems.Count
        Set itmx = NAMELIST.ListItems(t%)
        If itmx.SubItems(4) = lname Then
            NAMELIST.ListItems(t%).Selected = True
            Call namelist_ItemClick(itmx)
        Else
            NAMELIST.ListItems(t%).Selected = False
        End If
    Next t%
End If
End If
If Not NAMELIST.SelectedItem Is Nothing Then
    NAMELIST.SelectedItem.EnsureVisible
End If

End Sub

Private Sub clearroutine()
mugshotfile = ""
mugshot.Picture = LoadPicture()
alias = ""
ssn = ""
birthdate = ""
dl = ""
DLSTATE = ""
idnumber = ""
fbinumber = ""
birthplace = ""
fname = ""
lname = ""
haddress = ""
haddress2 = ""
hstate = ""
hzipcode = ""
wstate = ""
wzipcode = ""
hphone = ""
waddress = ""
waddress2 = ""
wphone = ""
resident.ListIndex = -1
race.ListIndex = -1
sex.ListIndex = -1
ethnicity.ListIndex = -1
age = ""
ht = ""
wt = ""
hair = ""
eyes = ""
peculiarities = ""
End Sub

Private Sub newname_Change()

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set cleanup = Nothing
End Sub

Private Sub lname_LostFocus()
If lname.Text > "" And NAMELIST.SelectedItem Is Nothing Then
    For t% = 1 To NAMELIST.ListItems.Count
        If NAMELIST.ListItems(t%) = lname Then
            NAMELIST.ListItems(t%).Selected = True
            Call namelist_ItemClick(NAMELIST.ListItems(t%))
            NAMELIST.ListItems(t%).EnsureVisible
            t% = NAMELIST.ListItems.Count
        End If
    Next t%
End If
End Sub

Private Sub mbirthdate_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
    If Len(mbirthdate) = 1 Or Len(mbirthdate) = 4 Then
        Call sendslash
    End If
End If
End Sub

Private Sub namelist_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
NAMELIST.SortKey = ColumnHeader.index - 1
If NAMELIST.SortOrder = lvwAscending Then
    NAMELIST.SortOrder = lvwDescending
Else
    NAMELIST.SortOrder = lvwAscending
End If
NAMELIST.Sorted = True
End Sub

Private Sub namelist_ItemClick(ByVal Item As MSComctlLib.ListItem)
If NAMELIST.SelectedItem Is Nothing Then
    Exit Sub
End If
Set itmx = NAMELIST.ListItems(NAMELIST.SelectedItem.index)
Dim DB As Database, RS As Recordset
On Error GoTo oderror
od:
ssql = ""
If ssn > "" Then
    ssql = ssql + " and ssn = '" + ssn + "'"
End If
If idnumber > "" Then
    ssql = ssql + " and idnumber = '" + idnumber + "'"
End If
If IsDate(birthdate) Then
    ssql = ssql + " and birthdate = #" + birthdate + "#"
End If
Set DB = OpenDatabase(nwl + "lawsuite.mdb")
Set RS = DB.OpenRecordset("select * from people where dpnamelf = " + Chr$(34) + itmx + Chr$(34) + ssql)
On Error Resume Next
Call clearroutine
If Not RS.EOF Then
    RS.MoveFirst
    fname = RS("dpname")
    Call setpopup(Mid$(Str(fname), 2), "F")
    If Not IsNull(RS("dpnamelf")) Then
        lname = RS("dpnamelf")
    End If
    If Not IsNull(RS("dphaddress")) Then
        haddress = RS("dphaddress")
    End If
    If Not IsNull(RS("dphaddress2")) Then
        haddress2 = RS("dphaddress2")
    End If
    If Not IsNull(RS("Hstate")) Then
        hstate = RS("HSTATE")
    End If
    If Not IsNull(RS("HZIPCODE")) Then
        hzipcode = RS("HZIPCODE")
    End If
    If Not IsNull(RS("dphphone")) Then
        hphone = RS("dphphone")
    End If
    If Not IsNull(RS("dpwaddress")) Then
        waddress = RS("dpwaddress")
    End If
    If Not IsNull(RS("dpwaddress2")) Then
        waddress2 = RS("dpwaddress2")
    End If
    If Not IsNull(RS("Wstate")) Then
        wstate = RS("WSTATE")
    End If
    If Not IsNull(RS("WZIPCODE")) Then
        wzipcode = RS("WZIPCODE")
    End If
    If Not IsNull(RS("dpwphone")) Then
        wphone = RS("dpwphone")
    End If
    If Not IsNull(RS("resident")) Then
        For t% = 0 To resident.ListCount - 1
            If resident.List(t%) = RS("resident") Or Left(resident.List(t%), 1) = RS("resident") Then
                resident.ListIndex = t%
                t% = resident.ListCount - 1
            End If
        Next t%
    End If
    If Not IsNull(RS("race")) Then
        For t% = 0 To race.ListCount - 1
            If race.List(t%) = RS("race") Or Left(race.List(t%), 1) = RS("race") Then
                race.ListIndex = t%
                t% = race.ListCount - 1
            End If
        Next t%
    End If
    If Not IsNull(RS("sex")) Then
        For t% = 0 To sex.ListCount - 1
            If sex.List(t%) = RS("sex") Or Left(sex.List(t%), 1) = RS("sex") Then
                sex.ListIndex = t%
                t% = sex.ListCount - 1
            End If
        Next t%
    End If
    If Not IsNull(RS("ethnicity")) Then
        For t% = 0 To ethnicity.ListCount - 1
            If ethnicity.List(t%) = RS("ethnicity") Or Left(ethnicity.List(t%), 1) = RS("ethnicity") Then
                ethnicity.ListIndex = t%
                t% = ethnicity.ListCount - 1
            End If
        Next t%
    End If
    If Not IsNull(RS("age")) Then
        age = RS("age")
    End If
    If Not IsNull(RS("height")) Then
        ht = RS("height")
    End If
    If Not IsNull(RS("weight")) Then
        wt = RS("weight")
    End If
    If Not IsNull(RS("hair")) Then
        hair = RS("hair")
    End If
    If Not IsNull(RS("eyes")) Then
        eyes = RS("eyes")
    End If
    If Not IsNull(RS("peculiarities")) Then
        peculiarities = RS("peculiarities")
    End If
    If Not IsNull(RS("alias")) Then
        alias = RS("alias")
    End If
    If Not IsNull(RS("ssn")) Then
        ssn = RS("ssn")
    End If
    If Not IsNull(RS("birthdate")) Then
        birthdate = RS("birthdate")
    End If
    If Not IsNull(RS("dl")) Then
        dl = RS("dl")
    End If
    If Not IsNull(RS("dlstate")) Then
        DLSTATE = RS("dlstate")
    End If
    If Not IsNull(RS("idnumber")) Then
        idnumber = RS("idnumber")
    End If
    If Not IsNull(RS("fbinumber")) Then
        fbinumber = RS("fbinumber")
    End If
    If Not IsNull(RS("birthplace")) Then
        birthplace = RS("birthplace")
    End If
    If Not IsNull(RS("mugshot")) Then
        mugshot.Picture = LoadPicture(RS("mugshot"))
        mugshotfile = RS("mugshot")
    Else
        mugshot.Picture = LoadPicture()
    End If
End If
DB.Close
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If

End Sub


Private Sub rapsheet_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
rapsheet.SortKey = ColumnHeader.index - 1
If rapsheet.SortOrder = lvwAscending Then
    rapsheet.SortOrder = lvwDescending
Else
    rapsheet.SortOrder = lvwAscending
End If
rapsheet.Sorted = True
End Sub

Private Sub rapsheet_ItemClick(ByVal Item As MSComctlLib.ListItem)
Set itmx = rapsheet.ListItems(rapsheet.SelectedItem.index)
arrestdate = itmx
casenumber = itmx.SubItems(1)
warrantnumber = itmx.SubItems(2)
charge = itmx.SubItems(3)
If itmx.SubItems(4) > "" Then
    For t% = 1 To disposition.ListCount
        If disposition.List(t%) = itmx.SubItems(4) Then
            disposition.ListIndex = t%
            t% = disposition.ListCount - 1
        End If
    Next t%
End If
DISPOSITIONDESCRIPTION = itmx.SubItems(5)
dispositiondate = itmx.SubItems(6)
End Sub

Private Sub sendslash()
SendKeys "/"
End Sub
Private Sub CLEARRAP()
arrestdate = ""
casenumber = ""
warrantnumber = ""
charge = ""
disposition.ListIndex = -1
dispositiondate = ""
DISPOSITIONDESCRIPTION = ""
End Sub
Private Sub OLDPRINT()
Printer.FontName = "Times New Roman"
Printer.FontSize = 16
Printer.FontBold = True
Printer.Print "Charge History"
Printer.Print Date$
Printer.FontSize = 10
Printer.Print
Printer.Print
Printer.FontBold = True
Printer.Print "NAME:";
Printer.FontBold = False
Printer.Print Tab(18); Left(lname, 37); Tab(80);
Printer.FontBold = True
Printer.Print "ID#:";
Printer.FontBold = False
Printer.Print Tab(100); idnumber
Printer.FontBold = True
Printer.Print
Printer.Print "ADDRESS:";
Printer.FontBold = False
Printer.Print Tab(18); Left(haddress, 37); Tab(80);
Printer.FontBold = True
Printer.Print "SSN:";
Printer.FontBold = False
Printer.Print Tab(100); ssn
Printer.Print
Printer.Print Tab(18); Left(haddress2 + " " + hstate + " " + hzipcode, 37); Tab(80);
Printer.FontBold = True
Printer.Print "DL#:";
Printer.FontBold = False
Printer.Print Tab(100); dl
Printer.Print
Printer.FontBold = True
If race.ListIndex > -1 Then
    Printer.Print "RACE:";
    Printer.FontBold = False
    Printer.Print Tab(18); race.List(race.ListIndex);
Else
    Printer.Print "RACE:";
End If
Printer.FontBold = False
If sex.ListIndex > -1 Then
    Printer.Print Tab(35);
    Printer.FontBold = True
    Printer.Print "SEX:";
    Printer.FontBold = False
    Printer.Print Tab(45); sex.List(sex.ListIndex); Tab(80);
    Printer.FontBold = True
    Printer.Print "FBI#:";
    Printer.FontBold = False
    Printer.Print Tab(100); fbinumber
Else
    Printer.Print Tab(35);
    Printer.FontBold = True
    Printer.Print "SEX:";
    Printer.FontBold = False
    Printer.Print Tab(80);
    Printer.FontBold = True
    Printer.Print "FBI#:";
    Printer.FontBold = False
    Printer.Print Tab(100); fbinumber
End If
Printer.FontBold = True
Printer.Print
Printer.Print "ALIAS:";
Printer.FontBold = False
Printer.Print Tab(18); Left(alias, 37)
Printer.FontBold = True
Printer.Print
Printer.Print "BIRTHDATE:";
Printer.FontBold = False
Printer.Print Tab(18); birthdate; Tab(80);
Printer.FontBold = True
Printer.Print "BIRTHPLACE:";
Printer.FontBold = False
Printer.Print Tab(100); birthplace
Printer.FontBold = True
Printer.Print
Printer.Print "HEIGHT:";
Printer.FontBold = False
Printer.Print Tab(15); ht; Tab(25);
Printer.FontBold = True
Printer.Print "WEIGHT:";
Printer.FontBold = False
Printer.Print Tab(40); wt; Tab(50);
Printer.FontBold = True
Printer.Print "HAIR:";
Printer.FontBold = False
Printer.Print Tab(60); hair; Tab(75);
Printer.FontBold = True
Printer.Print "EYES:";
Printer.FontBold = False
Printer.Print Tab(85); eyes
Printer.FontBold = True
Printer.Print
Printer.Print "OTHER DESCRIPTION:";
Printer.FontBold = False
Printer.Print Tab(30); peculiarities
Printer.Print
Printer.Print
Printer.Print
Printer.Print
Printer.Print
Printer.Print
Printer.FontSize = 10
Printer.FontUnderline = True
Printer.Print "ARREST DATE"; Tab(20); "CASE NO."; Tab(40); "WARRANT NO."; Tab(60); "CHARGE"; Tab(110); "DISPOSITION"; Tab(130); "DATE"
Printer.FontUnderline = False
For t% = 1 To rapsheet.ListItems.Count
    Set itmx = rapsheet.ListItems(t%)
    Printer.Print itmx; Tab(20); itmx.SubItems(1); Tab(40); itmx.SubItems(2); Tab(60); itmx.SubItems(3); Tab(110); itmx.SubItems(4); Tab(130); itmx.SubItems(6)
    Printer.Print Tab(110); itmx.SubItems(5)
Next t%
Printer.EndDoc

End Sub

Private Sub settsql(tfld As String, tsql As String)
tsql = ""
For tt = 1 To inidx
    If tsql = "" Then
        tsql = tfld + " = " + Chr$(34) + inpeople(tt) + Chr$(34)
    Else
        tsql = tsql + " OR " + tfld + " = " + Chr$(34) + inpeople(tt) + Chr$(34)
    End If
Next tt
End Sub
Private Sub runtsql(dbname As String, tbname As String, tidx As Long, tfld() As String, nfld() As String)
Dim FSQL As String
Call settsql(tfld(1), FSQL)
Dim DB As Database, RS As Recordset
Set DB = OpenDatabase(dbname)
Set RS = DB.OpenRecordset("select * from " + tbname + " where " + FSQL)
mergestatus = UCase(tbname) + " " + FSQL
mergestatus.Refresh
While Not RS.EOF
    RS.Edit
    For t = 1 To tidx
        If InStr(UCase(tfld(t)), "DATE") And nfld(t) = "" Then
        Else
            RS(tfld(t)) = nfld(t)
        End If
    Next t
    RS.Update
    RS.MoveNext
Wend
RS.Close
DB.Close



End Sub
