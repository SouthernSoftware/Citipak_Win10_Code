VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form badcheck 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Genesis Fraudulent Check"
   ClientHeight    =   7080
   ClientLeft      =   195
   ClientTop       =   945
   ClientWidth     =   11760
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7080
   ScaleWidth      =   11760
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C0C0&
      Height          =   2040
      Left            =   60
      TabIndex        =   99
      Top             =   5100
      Width           =   11655
      Begin VB.TextBox checkamount 
         Height          =   285
         Left            =   4440
         MaxLength       =   10
         TabIndex        =   56
         Top             =   330
         Width           =   1245
      End
      Begin VB.CommandButton spellck 
         Caption         =   "Spelling"
         Height          =   240
         Left            =   135
         TabIndex        =   107
         Top             =   870
         Width           =   945
      End
      Begin VB.TextBox comments 
         Height          =   465
         Left            =   1095
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   59
         Top             =   645
         Width           =   10515
      End
      Begin VB.ListBox approvingofficer 
         Height          =   645
         Left            =   7725
         TabIndex        =   66
         Top             =   1350
         Width           =   2415
      End
      Begin VB.ListBox reportingofficer 
         Height          =   645
         Left            =   3645
         TabIndex        =   64
         Top             =   1350
         Width           =   2415
      End
      Begin VB.TextBox reportingofficernumber 
         Height          =   285
         Left            =   6165
         TabIndex        =   65
         Top             =   1335
         Width           =   855
      End
      Begin VB.TextBox approvingofficernumber 
         Height          =   285
         Left            =   10215
         TabIndex        =   67
         Top             =   1350
         Width           =   900
      End
      Begin VB.TextBox checknumbers 
         Height          =   285
         Left            =   3060
         MaxLength       =   50
         TabIndex        =   55
         Top             =   330
         Width           =   1245
      End
      Begin VB.TextBox bankname 
         Height          =   285
         Left            =   5940
         MaxLength       =   50
         TabIndex        =   57
         Top             =   330
         Width           =   2800
      End
      Begin VB.TextBox status 
         Height          =   285
         Left            =   8820
         MaxLength       =   50
         TabIndex        =   58
         Top             =   330
         Width           =   2800
      End
      Begin VB.CheckBox theft 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Theft"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   150
         TabIndex        =   60
         Top             =   1155
         Width           =   975
      End
      Begin VB.CheckBox recovery 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Recovery"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1230
         TabIndex        =   61
         Top             =   1155
         Width           =   1215
      End
      Begin VB.CheckBox active 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Active"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   165
         TabIndex        =   62
         Top             =   1590
         Width           =   975
      End
      Begin VB.CheckBox cleared 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cleared by Arrest"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1245
         TabIndex        =   63
         Top             =   1590
         Width           =   2055
      End
      Begin VB.TextBox jurisdiction 
         Height          =   285
         Left            =   2160
         MaxLength       =   1
         TabIndex        =   54
         Top             =   315
         Width           =   765
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   "Check Amount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   4440
         TabIndex        =   108
         Top             =   120
         Width           =   1680
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Jurisdiction:     1 City   2 County 3 State   4 Out of State"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   495
         Left            =   120
         TabIndex        =   106
         Top             =   135
         Width           =   2775
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "COMMENTS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   45
         TabIndex        =   105
         Top             =   660
         Width           =   2775
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "APPROVING OFFICER          NUMBER"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   7710
         TabIndex        =   104
         Top             =   1155
         Width           =   3975
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "REPORTING OFFICER          NUMBER"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   3630
         TabIndex        =   103
         Top             =   1155
         Width           =   3855
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Check Number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   3060
         TabIndex        =   102
         Top             =   120
         Width           =   1680
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "Bank"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   5940
         TabIndex        =   101
         Top             =   120
         Width           =   1680
      End
      Begin VB.Label label45 
         BackStyle       =   0  'Transparent
         Caption         =   "Return Status"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   8820
         TabIndex        =   100
         Top             =   120
         Width           =   1680
      End
      Begin VB.Shape Shape2 
         Height          =   510
         Left            =   60
         Top             =   120
         Width           =   2940
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Height          =   1695
      Left            =   45
      TabIndex        =   84
      Top             =   3405
      Width           =   11650
      Begin VB.TextBox saddress 
         Height          =   285
         Left            =   3480
         MaxLength       =   100
         TabIndex        =   30
         Top             =   375
         Width           =   3495
      End
      Begin VB.ComboBox sname 
         Height          =   315
         Left            =   45
         TabIndex        =   29
         Top             =   330
         Width           =   3375
      End
      Begin VB.TextBox sage 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   8955
         MaxLength       =   4
         TabIndex        =   36
         Top             =   360
         Width           =   525
      End
      Begin VB.ListBox SSEX 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   8055
         TabIndex        =   35
         Top             =   225
         Width           =   885
      End
      Begin VB.ListBox SRACE 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   7020
         TabIndex        =   34
         Top             =   225
         Width           =   975
      End
      Begin VB.TextBox sbirthdate 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   9630
         TabIndex        =   37
         Top             =   360
         Width           =   1005
      End
      Begin VB.TextBox sweight 
         Height          =   285
         Left            =   690
         MaxLength       =   15
         TabIndex        =   43
         Top             =   1305
         Width           =   630
      End
      Begin VB.TextBox sheight 
         Height          =   285
         Left            =   30
         MaxLength       =   15
         TabIndex        =   42
         Top             =   1305
         Width           =   600
      End
      Begin VB.TextBox seyes 
         Height          =   285
         Left            =   2310
         MaxLength       =   15
         TabIndex        =   45
         Top             =   1305
         Width           =   800
      End
      Begin VB.TextBox shair 
         Height          =   285
         Left            =   1425
         MaxLength       =   15
         TabIndex        =   44
         Top             =   1305
         Width           =   800
      End
      Begin VB.CheckBox suspect 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Suspect"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   10635
         TabIndex        =   38
         Top             =   120
         Width           =   900
      End
      Begin VB.CheckBox wanted 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Wanted"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   10650
         TabIndex        =   39
         Top             =   345
         Width           =   900
      End
      Begin VB.CheckBox warrant 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Warrant"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   10635
         TabIndex        =   40
         Top             =   585
         Width           =   900
      End
      Begin VB.CheckBox arrest 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Arrest"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   10635
         TabIndex        =   41
         Top             =   825
         Width           =   900
      End
      Begin VB.TextBox totalarrested 
         Height          =   285
         Left            =   6870
         MaxLength       =   15
         TabIndex        =   50
         Top             =   1305
         Width           =   975
      End
      Begin VB.OptionButton nearyes 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Yes"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   8430
         TabIndex        =   51
         Top             =   1320
         Width           =   735
      End
      Begin VB.OptionButton nearno 
         BackColor       =   &H00C0C0C0&
         Caption         =   "No"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   9390
         TabIndex        =   52
         Top             =   1305
         Width           =   615
      End
      Begin VB.TextBox dateofoffense 
         Height          =   285
         Left            =   10485
         MaxLength       =   15
         TabIndex        =   53
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox drivers 
         Height          =   285
         Left            =   3615
         MaxLength       =   15
         TabIndex        =   47
         Top             =   1305
         Width           =   1155
      End
      Begin VB.TextBox ssn 
         Height          =   285
         Left            =   4815
         MaxLength       =   15
         TabIndex        =   48
         Top             =   1305
         Width           =   975
      End
      Begin VB.TextBox idnumber 
         Height          =   285
         Left            =   5865
         MaxLength       =   20
         TabIndex        =   49
         Top             =   1305
         Width           =   975
      End
      Begin VB.TextBox driversstate 
         Height          =   285
         Left            =   3165
         MaxLength       =   3
         TabIndex        =   46
         Top             =   1305
         Width           =   405
      End
      Begin VB.TextBox scity 
         Height          =   285
         Left            =   3495
         MaxLength       =   30
         TabIndex        =   31
         Top             =   675
         Width           =   1995
      End
      Begin VB.TextBox sstate 
         Height          =   285
         Left            =   5565
         MaxLength       =   2
         TabIndex        =   32
         Top             =   675
         Width           =   390
      End
      Begin VB.TextBox szipcode 
         Height          =   285
         Left            =   6045
         MaxLength       =   10
         TabIndex        =   33
         Top             =   675
         Width           =   930
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   3420
         TabIndex        =   98
         Top             =   150
         Width           =   975
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Subject Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   30
         TabIndex        =   97
         Top             =   150
         Width           =   2055
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Age"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   9015
         TabIndex        =   96
         Top             =   150
         Width           =   495
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Birthdate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   9615
         TabIndex        =   95
         Top             =   150
         Width           =   1095
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Height"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   30
         TabIndex        =   94
         Top             =   1050
         Width           =   1095
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Weight"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   705
         TabIndex        =   93
         Top             =   1050
         Width           =   1095
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Hair"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1425
         TabIndex        =   92
         Top             =   1050
         Width           =   1095
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Eyes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   2310
         TabIndex        =   91
         Top             =   1050
         Width           =   735
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Arrested"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   6750
         TabIndex        =   90
         Top             =   1065
         Width           =   1320
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "Arrest Near Offense Scene"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   8070
         TabIndex        =   89
         Top             =   1065
         Width           =   2400
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Offense Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   10470
         TabIndex        =   88
         Top             =   1065
         Width           =   1200
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "DL State/Number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   3225
         TabIndex        =   87
         Top             =   1050
         Width           =   1560
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "SSN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   4830
         TabIndex        =   86
         Top             =   1050
         Width           =   735
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "ID Num"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   5865
         TabIndex        =   85
         Top             =   1050
         Width           =   1005
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Height          =   1035
      Left            =   60
      TabIndex        =   78
      Top             =   2460
      Width           =   11655
      Begin VB.TextBox vhphone 
         Height          =   285
         Left            =   7020
         MaxLength       =   15
         TabIndex        =   24
         Top             =   360
         Width           =   1095
      End
      Begin VB.ComboBox vname 
         Height          =   315
         Left            =   45
         TabIndex        =   19
         Top             =   330
         Width           =   3375
      End
      Begin VB.TextBox vaddress 
         Height          =   285
         Left            =   3525
         MaxLength       =   100
         TabIndex        =   20
         Top             =   360
         Width           =   3495
      End
      Begin VB.TextBox vwphone 
         Height          =   285
         Left            =   8175
         MaxLength       =   15
         TabIndex        =   25
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox vage 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   11145
         MaxLength       =   4
         TabIndex        =   28
         Top             =   360
         Width           =   450
      End
      Begin VB.ListBox vsex 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   10230
         TabIndex        =   27
         Top             =   225
         Width           =   885
      End
      Begin VB.ListBox vrace 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   9270
         TabIndex        =   26
         Top             =   225
         Width           =   975
      End
      Begin VB.TextBox vcity 
         Height          =   285
         Left            =   3525
         MaxLength       =   30
         TabIndex        =   21
         Top             =   690
         Width           =   1995
      End
      Begin VB.TextBox vstate 
         Height          =   285
         Left            =   5610
         MaxLength       =   2
         TabIndex        =   22
         Top             =   690
         Width           =   390
      End
      Begin VB.TextBox vzipcode 
         Height          =   285
         Left            =   6090
         MaxLength       =   10
         TabIndex        =   23
         Top             =   690
         Width           =   930
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Victim Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   30
         TabIndex        =   83
         Top             =   150
         Width           =   2055
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Home Phone"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   7020
         TabIndex        =   82
         Top             =   150
         Width           =   1095
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   3525
         TabIndex        =   81
         Top             =   150
         Width           =   975
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Age"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   11190
         TabIndex        =   80
         Top             =   150
         Width           =   495
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Work Phone"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   8145
         TabIndex        =   79
         Top             =   150
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Height          =   2055
      Left            =   6075
      TabIndex        =   71
      Top             =   435
      Width           =   5610
      Begin VB.TextBox incidentdate 
         Height          =   285
         Left            =   4305
         MaxLength       =   10
         TabIndex        =   11
         Top             =   345
         Width           =   1095
      End
      Begin VB.TextBox incidentlocation 
         Height          =   285
         Left            =   75
         MaxLength       =   70
         TabIndex        =   10
         Top             =   345
         Width           =   4080
      End
      Begin VB.TextBox caddress 
         Height          =   285
         Left            =   60
         MaxLength       =   100
         TabIndex        =   14
         Top             =   1410
         Width           =   3735
      End
      Begin VB.ComboBox cname 
         Height          =   315
         Left            =   45
         TabIndex        =   13
         Top             =   870
         Width           =   3375
      End
      Begin VB.TextBox cphone 
         Height          =   285
         Left            =   4290
         MaxLength       =   15
         TabIndex        =   18
         Top             =   1725
         Width           =   1110
      End
      Begin VB.TextBox incidenttime 
         Height          =   285
         Left            =   4305
         MaxLength       =   10
         TabIndex        =   12
         Top             =   855
         Width           =   1095
      End
      Begin VB.TextBox ccity 
         Height          =   285
         Left            =   60
         MaxLength       =   30
         TabIndex        =   15
         Top             =   1740
         Width           =   1995
      End
      Begin VB.TextBox cstate 
         Height          =   285
         Left            =   2145
         MaxLength       =   2
         TabIndex        =   16
         Top             =   1740
         Width           =   390
      End
      Begin VB.TextBox czipcode 
         Height          =   285
         Left            =   2610
         MaxLength       =   10
         TabIndex        =   17
         Top             =   1740
         Width           =   1200
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   4305
         TabIndex        =   77
         Top             =   135
         Width           =   1020
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Incident Location"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   30
         TabIndex        =   76
         Top             =   135
         Width           =   2055
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   45
         TabIndex        =   75
         Top             =   1215
         Width           =   975
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   4305
         TabIndex        =   74
         Top             =   1515
         Width           =   975
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   4305
         TabIndex        =   73
         Top             =   645
         Width           =   1695
      End
      Begin VB.Label label44 
         BackStyle       =   0  'Transparent
         Caption         =   "Complainant Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   45
         TabIndex        =   72
         Top             =   660
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   2085
      Left            =   3255
      TabIndex        =   69
      Top             =   435
      Width           =   2730
      Begin VB.CheckBox highway 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Highway"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   420
         Width           =   1275
      End
      Begin VB.CheckBox commercial 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Commercial"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   780
         Width           =   1200
      End
      Begin VB.CheckBox scvstation 
         BackColor       =   &H00C0C0C0&
         Caption         =   "SCV. Station"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1140
         Width           =   1230
      End
      Begin VB.CheckBox chainstore 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Chain Store"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1440
         TabIndex        =   5
         Top             =   420
         Width           =   1215
      End
      Begin VB.CheckBox residence 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Residence"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1440
         TabIndex        =   6
         Top             =   780
         Width           =   1095
      End
      Begin VB.CheckBox bank 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Bank"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1440
         TabIndex        =   7
         Top             =   1140
         Width           =   1095
      End
      Begin VB.CheckBox other 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1500
         Width           =   225
      End
      Begin VB.TextBox otherspecify 
         Height          =   285
         Left            =   360
         MaxLength       =   30
         TabIndex        =   9
         Top             =   1500
         Width           =   2055
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   "Check Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   60
         TabIndex        =   70
         Top             =   165
         Width           =   1680
      End
   End
   Begin Crystal.CrystalReport REPORT 
      Left            =   8400
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      ReportFileName  =   "SERVICE"
      Destination     =   1
      PrintFileLinesPerPage=   60
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   68
      Top             =   0
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   741
      ButtonWidth     =   1773
      ButtonHeight    =   582
      Appearance      =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save   "
            Object.ToolTipText     =   "Save Case Number"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Clear   "
            Object.ToolTipText     =   "Clear All Fields"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete   "
            Object.ToolTipText     =   "Delete Case Number"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Print   "
            Object.ToolTipText     =   "Print Fraudulent Check Report"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit   "
            Object.ToolTipText     =   "Exit Fraudulent Check"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.ComboBox casenumber 
      Height          =   315
      Left            =   1455
      TabIndex        =   1
      Top             =   435
      Width           =   1770
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7200
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "badcheck.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "badcheck.frx":0454
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "badcheck.frx":08A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "badcheck.frx":0CFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "badcheck.frx":1150
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "badcheck.frx":15A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "badcheck.frx":19F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "badcheck.frx":1E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "badcheck.frx":22A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "badcheck.frx":26F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "badcheck.frx":2B48
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "badcheck.frx":2F9C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image mugshot 
      BorderStyle     =   1  'Fixed Single
      Height          =   1590
      Left            =   840
      Stretch         =   -1  'True
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "CASE NUMBER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   30
      TabIndex        =   0
      Top             =   465
      Width           =   1455
   End
End
Attribute VB_Name = "badcheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FROMXREF As Integer, casesetup As String, nametype As Integer

Private Sub address1_GotFocus()
Dim db As Database, rs As Recordset
On Error Resume Next
If Address1 = "" And Address2 = "" And compsubj > "" Then
    On Error GoTo oderror
od:
    Set db = OpenDatabase(nwl + "LAWSUITE.MDB")
    Set rs = db.OpenRecordset("SELECT * FROM PEOPLE WHERE DPNAMElf = " + Chr$(34) + compsubj + Chr$(34))
    If Not rs.EOF Then
        rs.MoveFirst
        Address1 = rs("DPHADDRESS")
        Address2 = rs("DPHADDRESS2")
        If phone = "" Then
            phone = rs("DPHPHONE")
        End If
    End If
End If
On Error Resume Next
db.Close
On Error GoTo 0
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If
    
End Sub

Private Sub approvingofficernumber_GotFocus()
If approvingofficernumber > "" Or approvingofficer = "" Then
    Exit Sub
End If
Dim db As DAO.Database, rs As DAO.Recordset
Set db = DAO.OpenDatabase(nwl + "LAWSUITE.MDB")
Set rs = db.OpenRecordset("SELECT PROFIDNUM FROM PROFESSIONALS WHERE PROFNAME = '" + approvingofficer + "' AND TYPE = 'D'")
If Not rs.EOF Then
    rs.MoveFirst
    If Not IsNull(rs("PROFIDNUM")) Then
        approvingofficernumber = rs("PROFIDNUM")
    End If
End If
db.Close
End Sub

Private Sub casenumber_Change()
If FROMXREF = 1 Then
    Call findrtn
    FROMXREF = 0
End If
End Sub

Private Sub CaseNumber_Click()
Call findrtn
casenumber.Refresh
End Sub

Private Sub completed_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
If Len(completed) = 1 Then
    SendKeys ":"
End If
End If

End Sub

Private Sub completed_LostFocus()
If completed > "" And Not IsDate(completed) Then
    msg = MsgBox("Invalid time entered.", 48, "Genesis Error Log")
    completed.SetFocus
End If

End Sub

Private Sub compsubj_LostFocus()
If compsubj > "" And InStr(compsubj, ",") = 0 Then
    msg = MsgBox("All names in the Service Call system should be entered in the format last name + comma + firstname.", 48, "Invalid Data Format")
    compsubj.SetFocus
End If

End Sub

Private Sub escort_LostFocus()
If escort = 0 Then
    warrant.SetFocus
End If
End Sub

Private Sub casenumber_GotFocus()
If casenumber = "" Then
    Dim db As Database, rs, rs2 As Recordset, mi As String
    On Error GoTo oderror
od:
    Set db = OpenDatabase(nwi + "INCIDENT.MDB")
    If Mid$(casesetup, 5, 1) = "0" Then
        If Len(casesetup) > 5 Then
            likepattern$ = String$(9, "?") + Mid$(casesetup, 6)
        Else
            likepattern$ = String$(9, "?")
        End If
    Else
        If Len(casesetup) > 5 Then
            likepattern$ = String$(10, "?") + Mid$(casesetup, 6)
        Else
            likepattern$ = String$(10, "?")
        End If
    End If
    likepattern$ = likepattern$ + String$(12 - Len(likepattern$), " ")
    Set rs = db.OpenRecordset("SELECT MAX(INCIDENTNUMBER) AS MI FROM INCIDENTREPORTC WHERE INCIDENTNUMBER LIKE '" + likepattern$ + "'")
    If Not rs.EOF Then
        rs.MoveFirst
        If Not IsNull(rs("MI")) Then
            mi = rs("mi")
            tempi$ = ""
            If Mid$(mi, 3, 1) = "-" Then
                starty% = 6
            Else
                starty% = 5
            End If
            For yy% = starty% To Len(mi)
                If InStr("0123456789", Mid$(mi, yy%, 1)) > 0 Then
                    tempi$ = tempi$ + Mid$(mi, yy%, 1)
                End If
            Next yy%
            Set rs2 = db.OpenRecordset("SELECT max(INCIDENTNUMBER) as bi froM badcheck WHERE INCIDENTNUMBER LIKE '" + likepattern$ + "'")
            If Not rs2.EOF Then
                rs2.MoveFirst
                If Not IsNull(rs2("bi")) Then
                    tempb$ = ""
                    If Mid$(rs2("bi"), 3, 1) = "-" Then
                        starty% = 6
                    Else
                        starty% = 5
                    End If
                    For yy% = starty% To Len(rs2("bi"))
                        If InStr("0123456789", Mid$(rs2("bi"), yy%, 1)) > 0 Then
                            tempb$ = tempb$ + Mid$(rs2("bi"), yy%, 1)
                        End If
                    Next yy%
                    If Val(tempb$) > Val(tempi$) Then
                        mi = rs2("bi")
                        tempi$ = tempb$
                    End If
                End If
            End If
            Select Case Left$(casesetup, 1)
                Case "1"
                    If Left$(mi, 2) <> Right$(Date$, 2) Then
                        tempi$ = "00000"
                    End If
                    Select Case Mid$(casesetup, 3, 1)
                        Case "1"
                            Select Case Val(Mid$(casesetup, 4, 1))
                                Case 0
                                    Select Case Len(casesetup)
                                        Case 5
                                            casenumber = Right$(Date$, 2) + "-" + Left$(Date$, 2) + Format$(Val(tempi$) + 1, "00000")
                                        Case Else
                                            casenumber = Right$(Date$, 2) + "-" + Left$(Date$, 2) + Format$(Val(tempi$) + 1, "00000") + Mid$(casesetup, 6)
                                    End Select
                                Case 1
                                    Select Case Len(casesetup)
                                        Case 5
                                            casenumber = Right$(Date$, 2) + "-" + Right$(Date$, 2) + Format$(Val(tempi$) + 1, "00000")
                                        Case Else
                                            casenumber = Right$(Date$, 2) + "-" + Right$(Date$, 2) + Format$(Val(tempi$) + 1, "00000") + Mid$(casesetup, 6)
                                    End Select
                            End Select
                        Case "0"
                            Select Case Val(Mid$(casesetup, 4, 1))
                                Case 0
                                    Select Case Len(casesetup)
                                        Case 5
                                            casenumber = Right$(Date$, 2) + Left$(Date$, 2) + Format$(Val(tempi$) + 1, "00000")
                                        Case Else
                                            casenumber = Right$(Date$, 2) + Left$(Date$, 2) + Format$(Val(tempi$) + 1, "00000") + Mid$(casesetup, 6)
                                    End Select
                                Case 1
                                    Select Case Len(casesetup)
                                        Case 5
                                            casenumber = Right$(Date$, 2) + Right$(Date$, 2) + Format$(Val(tempi$) + 1, "00000")
                                        Case Else
                                            casenumber = Right$(Date$, 2) + Right$(Date$, 2) + Format$(Val(tempi$) + 1, "00000") + Mid$(casesetup, 6)
                                    End Select
                            End Select
                    End Select
                Case "0"
                    If Left$(mi, 2) <> Left$(Date$, 2) Then
                        tempi$ = "00000"
                    End If
                    Select Case Mid$(casesetup, 3, 1)
                        Case "1"
                            Select Case Val(Mid$(casesetup, 4, 1))
                                Case 0
                                    Select Case Len(casesetup)
                                        Case 5
                                            casenumber = Left$(Date$, 2) + "-" + Left$(Date$, 2) + Format$(Val(tempi$) + 1, "00000")
                                        Case Else
                                            casenumber = Left$(Date$, 2) + "-" + Left$(Date$, 2) + Format$(Val(tempi$) + 1, "00000") + Mid$(casesetup, 6)
                                    End Select
                                Case 1
                                    Select Case Len(casesetup)
                                        Case 5
                                            casenumber = Left$(Date$, 2) + "-" + Right$(Date$, 2) + Format$(Val(tempi$) + 1, "00000")
                                        Case Else
                                            casenumber = Left$(Date$, 2) + "-" + Right$(Date$, 2) + Format$(Val(tempi$) + 1, "00000") + Mid$(casesetup, 6)
                                    End Select
                            End Select
                        Case "0"
                            Select Case Val(Mid$(casesetup, 4, 1))
                                Case 0
                                    Select Case Len(casesetup)
                                        Case 5
                                            casenumber = Left$(Date$, 2) + Left$(Date$, 2) + Format$(Val(tempi$) + 1, "00000")
                                        Case Else
                                            casenumber = Left$(Date$, 2) + Left$(Date$, 2) + Format$(Val(tempi$) + 1, "00000") + Mid$(casesetup, 6)
                                    End Select
                                Case 1
                                    Select Case Len(casesetup)
                                        Case 5
                                            casenumber = Left$(Date$, 2) + Right$(Date$, 2) + Format$(Val(tempi$) + 1, "00000")
                                        Case Else
                                            casenumber = Left$(Date$, 2) + Right$(Date$, 2) + Format$(Val(tempi$) + 1, "00000") + Mid$(casesetup, 6)
                                    End Select
                            End Select
                    End Select
            End Select
    End If
    On Error Resume Next
    db.Close
End If
End If
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If


End Sub

Private Sub cname_Click()
Dim db As Database, ds As Recordset
On Error GoTo oderror
od:
Call setpopup(cname, "L")
Set db = OpenDatabase(nwl + "LAWSUITE.MDB")
Set ds = db.OpenRecordset("select * from PEOPLE where dpnameLF = " + Chr$(34) + cname + Chr$(34))
If Not ds.EOF Then
   ds.MoveFirst
    If Not IsNull(ds("dphaddress")) Then
        caddress = ds("dphaddress")
    Else
        caddress = ""
        End If
    If Not IsNull(ds("dphaddress2")) Then
        CCITY = ds("dphaddress2")
    End If
    If Not IsNull(ds("Hstate")) Then
        CSTATE = ds("Hstate")
    End If
    If Not IsNull(ds("Hzipcode")) Then
        CZIPCODE = ds("Hzipcode")
    End If
    If Not IsNull(ds("dphphone")) And ds("dphphone") <> "" Then
        cphone = ds("dphphone")
    Else
        cphone = ""
    End If
End If
Exit Sub
db.Close
On Error GoTo 0
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
    Resume
End If

End Sub

Private Sub cname_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 And Shift = vbCtrlMask Then
    If nametype = 0 Then
        cleanup.fname = Me.ActiveControl.Text
        cleanup.lname = ""
    Else
        cleanup.lname = Me.ActiveControl.Text
        cleanup.fname = ""
    End If
    cleanup.Show
End If
End Sub

Private Sub cname_LostFocus()
If Len(cname) > 50 Then
    cname = Left$(cname, 50)
End If
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Comments_KeyDown(KeyCode As Integer, Shift As Integer)
If (Shift = vbCtrlMask) And (KeyCode = vbKeyF2) Then
        Call SpellCk_Click
    End If
End Sub

Private Sub dateofoffense_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
If Len(dateofoffense) = 1 Or Len(dateofoffense) = 4 Then
    SendKeys "/"
End If
End If

End Sub

Private Sub Form_Load()
nametype = 1
Me.Top = 0
Me.Left = 0
Me.Height = 7600
Me.Width = 11800
For t% = 0 To Forms.Count - 1
    If Forms(t%).Name = "xref" Then
        FROMXREF = 1
        t% = Forms.Count - 1
    End If
Next t%
srace.clear
srace.AddItem "White"
srace.AddItem "Black"
srace.AddItem "Indian"
srace.AddItem "Asian/Pacific Islander"
srace.AddItem "Unknown"
vrace.clear
vrace.AddItem "White"
vrace.AddItem "Black"
vrace.AddItem "Indian"
vrace.AddItem "Asian/Pacific Islander"
vrace.AddItem "Unknown"
ssex.clear
ssex.AddItem "Female"
ssex.AddItem "Male"
ssex.AddItem "Unknown"
vsex.clear
vsex.AddItem "Female"
vsex.AddItem "Male"
vsex.AddItem "Unknown"
casesetup = "10001"
Open nwi + "caseset.tag" For Input As #1
Line Input #1, ABC$
Close #1
If ABC$ > "" Then
    casesetup = ABC$
End If
Call clearrtn
Call loadcase
Call loadofficer
Call loadname
End Sub

Private Sub loadcase()
Dim db As Database, rs As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nwi + "incident.mdb")
Set rs = db.OpenRecordset("select incidentnumber from badcheck order by incidentnumber desc")
If Not rs.EOF Then
    rs.MoveFirst
End If
casenumber.clear
While Not rs.EOF
    casenumber.AddItem rs("incidentnumber")
    rs.MoveNext
Wend
db.Close
On Error Resume Next
casenumber.SetFocus
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If
End Sub
Private Sub clearrtn()
mugshot.Picture = LoadPicture()
casenumber = ""
cname = ""
caddress = ""
CCITY = ""
CSTATE = ""
CZIPCODE = ""
cphone = ""
IncidentDate = ""
IncidentTime = ""
dateofoffense = ""
incidentlocation = ""
highway = 0
Commercial = 0
scvstation = 0
chainstore = 0
residence = 0
bank = 0
Other = 0
otherspecify = ""
vname = ""
vaddress = ""
vcity = ""
vstate = ""
vzipcode = ""
vrace.ListIndex = -1
vsex.ListIndex = -1
vage = ""
vhphone = ""
vwphone = ""
suspect = 0
wanted = 0
warrant = 0
arrest = 0
sname = ""
saddress = ""
scity = ""
sstate = ""
szipcode = ""
srace.ListIndex = -1
ssex.ListIndex = -1
sage = ""
drivers = ""
driversstate = ""
ssn = ""
idnumber = ""
sbirthdate = ""
sheight = ""
sweight = ""
SHAIR = ""
SEYES = ""
totalarrested = ""
nearyes = False
nearno = False
checknumbers = ""
checkamount = ""
jurisdiction = ""
bankname = ""
status = ""
comments.Text = ""
theft = 0
recovery = 0
active = 0
cleared = 0
reportingofficer.ListIndex = -1
reportingofficernumber = ""
approvingofficer.ListIndex = -1
approvingofficernumber = ""
Call loadcase
Call loadofficer
Call loadname
Call casenumber_GotFocus
End Sub
Private Sub loadofficer()
Dim db As Database, rs As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nwl + "lawsuite.mdb")
Set rs = db.OpenRecordset("select profname from professionals where type = 'D' order by profname")
If Not rs.EOF Then
    rs.MoveFirst
End If
reportingofficer.clear
approvingofficer.clear
While Not rs.EOF
    reportingofficer.AddItem rs("profname")
    approvingofficer.AddItem rs("profname")
    rs.MoveNext
Wend
db.Close
On Error Resume Next
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If
End Sub
Private Sub loadname()
Dim db As Database, rs As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nwl + "LAWSUITE.mdb")
Set rs = db.OpenRecordset("select DPnamelf from people order by DPnamElf")
If Not rs.EOF Then
    rs.MoveFirst
End If
cname.clear
vname.clear
sname.clear
While Not rs.EOF
    cname.AddItem rs("DPnamelf")
    vname.AddItem rs("DPnamelf")
    sname.AddItem rs("DPnamelf")
    rs.MoveNext
Wend
db.Close
On Error Resume Next
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If
End Sub

Private Sub Form_LostFocus()
goingelsewhere = False
End Sub

Private Sub Form_Paint()
SetAlwaysOnTop Me

End Sub

Private Sub Form_Unload(Cancel As Integer)
goingelsewhere = False
Set badcheck = Nothing
End Sub

Private Sub incidentdate_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
If Len(IncidentDate) = 1 Or Len(IncidentDate) = 4 Then
    SendKeys "/"
End If
End If


End Sub

Private Sub incidenttime_Change()
If KeyAscii <> 8 Then
If Len(IncidentTime) = 2 Then
    SendKeys ":"
End If
End If

End Sub

Private Sub received_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
If Len(received) = 1 Then
    SendKeys ":"
End If
End If

End Sub

Private Sub received_LostFocus()
If received > "" And Not IsDate(received) Then
    msg = MsgBox("Invalid time entered.", 48, "Genesis Error Log")
    received.SetFocus
End If
End Sub

Private Sub reportingofficernumber_GotFocus()
If reportingofficernumber > "" Or reportingofficer = "" Then
    Exit Sub
End If
Dim db As DAO.Database, rs As DAO.Recordset
Set db = DAO.OpenDatabase(nwl + "LAWSUITE.MDB")
Set rs = db.OpenRecordset("SELECT PROFIDNUM FROM PROFESSIONALS WHERE PROFNAME = '" + reportingofficer + "' AND TYPE = 'D'")
If Not rs.EOF Then
    rs.MoveFirst
    If Not IsNull(rs("PROFIDNUM")) Then
        reportingofficernumber = rs("PROFIDNUM")
    End If
End If
db.Close
End Sub

Private Sub sbirthdate_Change()
If IsDate(sbirthdate) Then
    sage = DateDiff("yyyy", CDate(sbirthdate), CDate(Date$))
End If
End Sub

Private Sub sbirthdate_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
If Len(sbirthdate) = 1 Or Len(sbirthdate) = 4 Then
    SendKeys "/"
End If
End If

End Sub

Private Sub sname_Click()
Dim db As Database, ds As Recordset
On Error GoTo oderror
od:
Call setpopup(sname, "L")
Set db = OpenDatabase(nwl + "LAWSUITE.MDB")
Set ds = db.OpenRecordset("select * from PEOPLE where dpnameLF = " + Chr$(34) + sname + Chr$(34))
If Not ds.EOF Then
   ds.MoveFirst
    If Not IsNull(ds("dphaddress")) Then
        saddress = ds("dphaddress")
    Else
        saddress = ""
    End If
    If Not IsNull(ds("dphaddress2")) Then
        scity = ds("dphaddress2")
    End If
    If Not IsNull(ds("Hstate")) Then
        sstate = ds("Hstate")
    End If
    If Not IsNull(ds("Hzipcode")) Then
        szipcode = ds("Hzipcode")
    End If
    If Not IsNull(ds("dphphone")) And ds("dphphone") <> "" Then
        sphone = ds("dphphone")
    Else
        sphone = ""
    End If
End If
Exit Sub
db.Close
On Error GoTo 0
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If

End Sub

Private Sub sname_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 And Shift = vbCtrlMask Then
    If nametype = 0 Then
        cleanup.fname = Me.ActiveControl.Text
        cleanup.lname = ""
    Else
        cleanup.lname = Me.ActiveControl.Text
        cleanup.fname = ""
    End If
    cleanup.Show
End If
End Sub

Private Sub sname_LostFocus()
If Len(sname) > 50 Then
    sname = Left$(sname, 50)
End If

End Sub


Private Sub SpellCk_Click()
BeginSpellCheck comments.Text, comments
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button
    Case "Save   "
        If UCase(frmLogin.txtUserName) = "DEMO" And UCase(frmLogin.txtPassword) = "DEMO" Then
            msg = MsgBox("Not available in DEMO version.", 48, "Genesis Information Log")
            Screen.MousePointer = 0
            Exit Sub
        End If
        Screen.MousePointer = 11
        Call savertn
        Screen.MousePointer = 0
    Case "Clear   "
        Screen.MousePointer = 11
        Call clearrtn
        Screen.MousePointer = 0
    Case "Delete   "
        If UCase(frmLogin.txtUserName) = "DEMO" And UCase(frmLogin.txtPassword) = "DEMO" Then
            msg = MsgBox("Not available in DEMO version.", 48, "Genesis Information Log")
            Screen.MousePointer = 0
            Exit Sub
        End If
        Screen.MousePointer = 11
        Call deletertn
        Screen.MousePointer = 0
    Case "Print   "
        Screen.MousePointer = 11
        Call printrtn
        Screen.MousePointer = 0
    Case "Exit   "
        Unload Me
End Select
End Sub
Private Sub savertn()
Dim db As Database, rs As Recordset
If casenumber = "" Then
    msg = MsgBox("A case number must be entered.", 48, "Genesis Error Log")
    Exit Sub
End If
If Not IsDate(IncidentDate) Then
    msg = MsgBox("Incident Date must be entered and must be a valid date.", 48, "Genesis Error Log")
    Exit Sub
End If
If Not IsDate(dateofoffense) Then
    msg = MsgBox("Date of Offense must be entered and must be a valid date.", 48, "Genesis Error Log")
    Exit Sub
End If
On Error GoTo oderror
od:
Set db = OpenDatabase(nwi + "incident.mdb")
Set rs = db.OpenRecordset("select * from badcheck where incidentnumber =" + Chr$(34) + casenumber + Chr$(34))
If Not rs.EOF Then
    rs.MoveFirst
    rs.Edit
Else
    rs.AddNew
End If
rs("incidentnumber") = casenumber + Space$(12 - Len(casenumber))
rs("cname") = cname
rs("caddress") = caddress
rs("ccity") = CCITY
rs("cstate") = CSTATE
rs("czipcode") = CZIPCODE
rs("cphone") = cphone
rs("incidentdate") = IncidentDate
rs("incidenttime") = IncidentTime
rs("dateofoffense") = dateofoffense
rs("incidentlocation") = incidentlocation
rs("highway") = highway
rs("commercial") = Commercial
rs("scvstation") = scvstation
rs("chainstore") = chainstore
rs("residence") = residence
rs("bank") = bank
rs("other") = Other
rs("otherspecify") = otherspecify
rs("vname") = vname
rs("vaddress") = vaddress
rs("vcity") = vcity
rs("vstate") = vstate
rs("vzipcode") = vzipcode
If vrace.ListIndex > -1 Then
    rs("vrace") = Left$(vrace.List(vrace.ListIndex), 1)
Else
    rs("vrace") = ""
End If
If vsex.ListIndex > -1 Then
    rs("vsex") = Left$(vsex.List(vsex.ListIndex), 1)
Else
    rs("vsex") = ""
End If
rs("vage") = vage
rs("vhphone") = vhphone
rs("vwphone") = vwphone
rs("suspect") = suspect
rs("wanted") = wanted
rs("warrant") = warrant
rs("arrest") = arrest
rs("sname") = sname
rs("saddress") = saddress
rs("scity") = scity
rs("sstate") = sstate
rs("szipcode") = szipcode
If srace.ListIndex > -1 Then
    rs("srace") = Left$(srace.List(srace.ListIndex), 1)
Else
    rs("srace") = ""
End If
If ssex.ListIndex > -1 Then
    rs("ssex") = Left$(ssex.List(ssex.ListIndex), 1)
Else
    rs("ssex") = ""
End If
rs("sage") = sage
rs("drivers") = drivers
rs("driversstate") = driversstate
rs("ssn") = ssn
rs("idnumber") = idnumber
If IsDate(sbirthdate) Then
    rs("sdateofbirth") = sbirthdate
Else
    rs("sdateofbirth") = Null
End If
rs("sheight") = sheight
rs("sweight") = sweight
rs("shair") = SHAIR
rs("seyes") = SEYES
rs("totalarrested") = Val(totalarrested)
rs("nearyes") = nearyes
rs("nearno") = nearno
rs("checknumbers") = checknumbers
rs("checkamount") = checkamount
rs("jurisdiction") = jurisdiction
rs("bankname") = bankname
rs("status") = status
rs("comments") = comments.Text
rs("theft") = theft
rs("recovery") = recovery
rs("active") = active
rs("cleared") = cleared
If reportingofficer.ListIndex > -1 Then
    rs("reportingofficer") = reportingofficer.List(reportingofficer.ListIndex)
Else
    rs("reportingofficer") = ""
End If
rs("reportingofficerUNIT") = reportingofficernumber
If approvingofficer.ListIndex > -1 Then
    rs("approvingofficer") = approvingofficer.List(approvingofficer.ListIndex)
Else
    rs("approvingofficer") = ""
End If
rs("approvingofficerUNIT") = approvingofficernumber
'CES Code
rs("userfullname") = frmLogin.userfullname
rs("userid") = frmLogin.userid
rs("ORINUMBER") = frmLogin.orinumber
rs("udate") = Format$(Now, "mm/dd/yyyy")
rs("utime") = Format$(Now, "hh:mm:ss")
'********
rs.Update
On Error Resume Next
Set db = OpenDatabase(nwl + "lawsuite.mdb")
'----- OFFICERS
If reportingofficer > "" Then
    Set rs = db.OpenRecordset("select profidnum,profname, type from professionals where profname =" + Chr$(34) + reportingofficer + Chr$(34))
    If rs.EOF Then
        rs.AddNew
    Else
        rs.MoveFirst
        rs.Edit
    End If
    rs("profname") = reportingofficer
    rs("profidnum") = reportingofficernumber
    If rs.EOF Then
        reportingofficer.AddItem reportingofficer
        approvingofficer.AddItem reportingofficer
    End If
    rs("type") = "D"
    rs.Update
End If
If approvingofficer > "" Then
    Set rs = db.OpenRecordset("select profidnum,profname, type from professionals where profname =" + Chr$(34) + approvingofficer + Chr$(34))
    If rs.EOF Then
        rs.AddNew
    Else
        rs.MoveFirst
        rs.Edit
    End If
    rs("profname") = approvingofficer
    rs("profidnum") = approvingofficernumber
    If rs.EOF Then
        reportingofficer.AddItem approvingofficer
        approvingofficer.AddItem approvingofficer
    End If
    rs("type") = "D"
    rs.Update
End If

Set rs = db.OpenRecordset("select * from people where dpnamelf = " + Chr$(34) + cname + Chr$(34))
If Not rs.EOF Then
    rs.MoveFirst
    rs.Edit
Else
    rs.AddNew
End If
rs("dpnamelf") = cname
rs("dphaddress") = caddress
rs("dphaddress2") = CCITY
rs("dstate") = CSTATE
rs("dzipcode") = CZIPCODE
rs("dpsort") = Left$(cname, 15)
If cphone > "" Then
    rs("dphphone") = cphone
End If
hoLdname = cname
osort1$ = ""
If Left$(hoLdname, 1) = " " Then
    hoLdname = Mid$(hoLdname, 2)
End If
If InStr(hoLdname, " CORP") > 0 Or InStr(hoLdname, ",INC") > 0 Or InStr(hoLdname, "COMPANY") > 0 Or InStr(hoLdname, "INC.") > 0 Then
    osort1$ = hoLdname
End If
tso$ = hoLdname
If InStr(tso$, " et al") > 0 Then
    tso$ = Left$(tso$, InStr(tso$, " et al") - 1)
End If
If InStr(tso$, " et. al.") > 0 Then
    tso$ = Left$(tso$, InStr(tso$, " et. al.") - 1)
End If
If InStr(tso$, ",et al") > 0 Then
    tso$ = Left$(tso$, InStr(tso$, ",et al") - 1)
End If
If InStr(tso$, ",et. al.") > 0 Then
    tso$ = Left$(tso$, InStr(tso$, ",et. al.") - 1)
End If
If Right$(tso$, 1) = "," Then
    tso$ = Left$(tso$, Len(tso$) - 1)
End If
If InStr(tso$, "&") > 0 Then
    tso$ = Left$(tso$, InStr(tso$, "&") - 1)
End If
If Right$(tso$, 1) = "," Then
    tso$ = Left$(tso$, Len(tso$) - 1)
End If
firstspace% = 0
While Right$(tso$, 1) = " " And Len(tso$) > 1
    tso$ = Left$(tso$, Len(tso$) - 1)
Wend
For tt% = 1 To Len(tso$)
    If Mid$(tso$, tt%, 1) = "," Then
        firstspace% = tt%
        tt% = Len(tso$)
    End If
Next tt%
If firstspace% = 0 Then
    If osort1$ = "" Then
        osort1$ = tso$
    End If
    GoTo rsupdate1
End If
tempsort$ = Mid$(tso$, firstspace% + 1)
If Left$(tempsort$, 1) = " " Then
    tempsort$ = Mid$(tempsort$, 2)
End If
tso$ = Left$(tso$, firstspace% - 1)
If Right$(tso$, 1) = " " Then
    tso$ = Left$(tso$, Len(tso$) - 1)
End If
tempsort$ = tempsort$ + " " + tso$
If osort1$ = "" Then
    osort1$ = tempsort$
End If
If InStr(osort1$, "JR.") Then
    If Mid$(osort1$, InStr(osort1$, "JR.") + 3, 1) = " " Then
        osort1$ = Left$(osort1$, InStr(osort1$, "JR.") - 1) + Mid$(osort1$, InStr(osort1$, "JR.") + 4) + ", JR."
    Else
        osort1$ = Left$(osort1$, InStr(osort1$, "JR.") - 1) + Mid$(osort1$, InStr(osort1$, "JR.") + 3) + ", JR."
End If
End If
If InStr(osort1$, "SR.") Then
    If Mid$(osort1$, InStr(osort1$, "SR.") + 3, 1) = " " Then
        osort1$ = Left$(osort1$, InStr(osort1$, "SR.") - 1) + Mid$(osort1$, InStr(osort1$, "SR.") + 4) + ", SR."
    Else
        osort1$ = Left$(osort1$, InStr(osort1$, "SR.") - 1) + Mid$(osort1$, InStr(osort1$, "SR.") + 3) + ", SR."
End If
End If
If InStr(osort1$, "III") Then
    If Mid$(osort1$, InStr(osort1$, "III") + 3, 1) = " " Then
        osort1$ = Left$(osort1$, InStr(osort1$, "III") - 1) + Mid$(osort1$, InStr(osort1$, "III") + 4) + ", III"
    Else
        osort1$ = Left$(osort1$, InStr(osort1$, "III") - 1) + Mid$(osort1$, InStr(osort1$, "III") + 3) + ", III"
    End If
End If
If InStr(osort1$, "IV") Then
    If Mid$(osort1$, InStr(osort1$, "IV") + 2, 1) = " " Then
        osort1$ = Left$(osort1$, InStr(osort1$, "IV") - 1) + Mid$(osort1$, InStr(osort1$, "IV") + 3) + ", III"
    Else
        osort1$ = Left$(osort1$, InStr(osort1$, "IV") - 1) + Mid$(osort1$, InStr(osort1$, "IV") + 2) + ", III"
    End If
End If
If Left$(osort1$, 1) = " " Then
    osort1$ = Mid$(osort1$, 2)
End If
rsupdate1:
rs("dpname") = osort1$
rs.Update
Set rs = db.OpenRecordset("select * from people where dpnamelf = " + Chr$(34) + vname + Chr$(34))
If Not rs.EOF Then
    rs.MoveFirst
    rs.Edit
Else
    rs.AddNew
End If
rs("dpnamelf") = vname
rs("dphaddress") = vaddress
rs("dphaddress2") = vcity
rs("dstate") = vstate
rs("dzipcode") = vzipcode
rs("dpsort") = Left$(vname, 15)
If vhphone > "" Then
    rs("dphphone") = vhphone
End If
If vwphone > "" Then
    rs("dpwphone") = vwphone
End If
rs("race") = Left$(vrace.List(vrace.ListIndex), 1)
rs("sex") = Left$(vsex.List(vsex.ListIndex), 1)
rs("age") = vage
hoLdname = vname
osort1$ = ""
If Left$(hoLdname, 1) = " " Then
    hoLdname = Mid$(hoLdname, 2)
End If
If InStr(hoLdname, " CORP") > 0 Or InStr(hoLdname, ",INC") > 0 Or InStr(hoLdname, "COMPANY") > 0 Or InStr(hoLdname, "INC.") > 0 Then
    osort1$ = hoLdname
End If
tso$ = hoLdname
If InStr(tso$, " et al") > 0 Then
    tso$ = Left$(tso$, InStr(tso$, " et al") - 1)
End If
If InStr(tso$, " et. al.") > 0 Then
    tso$ = Left$(tso$, InStr(tso$, " et. al.") - 1)
End If
If InStr(tso$, ",et al") > 0 Then
    tso$ = Left$(tso$, InStr(tso$, ",et al") - 1)
End If
If InStr(tso$, ",et. al.") > 0 Then
    tso$ = Left$(tso$, InStr(tso$, ",et. al.") - 1)
End If
If Right$(tso$, 1) = "," Then
    tso$ = Left$(tso$, Len(tso$) - 1)
End If
If InStr(tso$, "&") > 0 Then
    tso$ = Left$(tso$, InStr(tso$, "&") - 1)
End If
If Right$(tso$, 1) = "," Then
    tso$ = Left$(tso$, Len(tso$) - 1)
End If
firstspace% = 0
While Right$(tso$, 1) = " " And Len(tso$) > 1
    tso$ = Left$(tso$, Len(tso$) - 1)
Wend
For tt% = 1 To Len(tso$)
    If Mid$(tso$, tt%, 1) = "," Then
        firstspace% = tt%
        tt% = Len(tso$)
    End If
Next tt%
If firstspace% = 0 Then
    If osort1$ = "" Then
        osort1$ = tso$
    End If
    GoTo rsupdate2
End If
tempsort$ = Mid$(tso$, firstspace% + 1)
If Left$(tempsort$, 1) = " " Then
    tempsort$ = Mid$(tempsort$, 2)
End If
tso$ = Left$(tso$, firstspace% - 1)
If Right$(tso$, 1) = " " Then
    tso$ = Left$(tso$, Len(tso$) - 1)
End If
tempsort$ = tempsort$ + " " + tso$
If osort1$ = "" Then
    osort1$ = tempsort$
End If
If InStr(osort1$, "JR.") Then
    If Mid$(osort1$, InStr(osort1$, "JR.") + 3, 1) = " " Then
        osort1$ = Left$(osort1$, InStr(osort1$, "JR.") - 1) + Mid$(osort1$, InStr(osort1$, "JR.") + 4) + ", JR."
    Else
        osort1$ = Left$(osort1$, InStr(osort1$, "JR.") - 1) + Mid$(osort1$, InStr(osort1$, "JR.") + 3) + ", JR."
End If
End If
If InStr(osort1$, "SR.") Then
    If Mid$(osort1$, InStr(osort1$, "SR.") + 3, 1) = " " Then
        osort1$ = Left$(osort1$, InStr(osort1$, "SR.") - 1) + Mid$(osort1$, InStr(osort1$, "SR.") + 4) + ", SR."
    Else
        osort1$ = Left$(osort1$, InStr(osort1$, "SR.") - 1) + Mid$(osort1$, InStr(osort1$, "SR.") + 3) + ", SR."
End If
End If
If InStr(osort1$, "III") Then
    If Mid$(osort1$, InStr(osort1$, "III") + 3, 1) = " " Then
        osort1$ = Left$(osort1$, InStr(osort1$, "III") - 1) + Mid$(osort1$, InStr(osort1$, "III") + 4) + ", III"
    Else
        osort1$ = Left$(osort1$, InStr(osort1$, "III") - 1) + Mid$(osort1$, InStr(osort1$, "III") + 3) + ", III"
    End If
End If
If InStr(osort1$, "IV") Then
    If Mid$(osort1$, InStr(osort1$, "IV") + 2, 1) = " " Then
        osort1$ = Left$(osort1$, InStr(osort1$, "IV") - 1) + Mid$(osort1$, InStr(osort1$, "IV") + 3) + ", III"
    Else
        osort1$ = Left$(osort1$, InStr(osort1$, "IV") - 1) + Mid$(osort1$, InStr(osort1$, "IV") + 2) + ", III"
    End If
End If
If Left$(osort1$, 1) = " " Then
    osort1$ = Mid$(osort1$, 2)
End If
rsupdate2:
rs("dpname") = osort1$
rs.Update
ssql = ""
If ssn > "" Then
    ssql = ssql + " and ssn = '" + ssn + "'"
End If
If IsDate(BIRTHDATE) Then
    ssql = ssql + " and birthdate = #" + BIRTHDATE + "#"
End If
If idnumber > "" Then
    ssql = ssql + " and idnumber = '" + idnumber + "'"
End If
Set rs = db.OpenRecordset("select * from people where dpnamelf = " + Chr$(34) + sname + Chr$(34) + ssql)
If Not rs.EOF Then
    rs.MoveFirst
    rs.Edit
Else
    rs.AddNew
End If
rs("dpnamelf") = sname
rs("dphaddress") = saddress
rs("dphaddress2") = scity
rs("dstate") = sstate
rs("dzipcode") = szipcode
rs("dpsort") = Left$(sname, 15)
If sphone > "" Then
    rs("dphphone") = sphone
    rs("resident") = Left$(sresident.List(sresident.ListIndex), 1)
End If
rs("HEIGHT") = sheight
rs("WEIGHT") = sweight
rs("HAIR") = SHAIR
rs("EYES") = SEYES
rs("race") = Left$(srace.List(srace.ListIndex), 1)
rs("sex") = Left$(ssex.List(ssex.ListIndex), 1)
rs("age") = sage
If IsDate(sbirthdate) Then
    rs("birthdate") = sbirthdate
End If
rs("ssn") = ssn
rs("idnumber") = idnumber
rs("dl") = drivers
rs("dlstate") = driversstate
hoLdname = sname
osort1$ = ""
If Left$(hoLdname, 1) = " " Then
    hoLdname = Mid$(hoLdname, 2)
End If
If InStr(hoLdname, " CORP") > 0 Or InStr(hoLdname, ",INC") > 0 Or InStr(hoLdname, "COMPANY") > 0 Or InStr(hoLdname, "INC.") > 0 Then
    osort1$ = hoLdname
End If
tso$ = hoLdname
If InStr(tso$, " et al") > 0 Then
    tso$ = Left$(tso$, InStr(tso$, " et al") - 1)
End If
If InStr(tso$, " et. al.") > 0 Then
    tso$ = Left$(tso$, InStr(tso$, " et. al.") - 1)
End If
If InStr(tso$, ",et al") > 0 Then
    tso$ = Left$(tso$, InStr(tso$, ",et al") - 1)
End If
If InStr(tso$, ",et. al.") > 0 Then
    tso$ = Left$(tso$, InStr(tso$, ",et. al.") - 1)
End If
If Right$(tso$, 1) = "," Then
    tso$ = Left$(tso$, Len(tso$) - 1)
End If
If InStr(tso$, "&") > 0 Then
    tso$ = Left$(tso$, InStr(tso$, "&") - 1)
End If
If Right$(tso$, 1) = "," Then
    tso$ = Left$(tso$, Len(tso$) - 1)
End If
firstspace% = 0
While Right$(tso$, 1) = " " And Len(tso$) > 1
    tso$ = Left$(tso$, Len(tso$) - 1)
Wend
For tt% = 1 To Len(tso$)
    If Mid$(tso$, tt%, 1) = "," Then
        firstspace% = tt%
        tt% = Len(tso$)
    End If
Next tt%
If firstspace% = 0 Then
    If osort1$ = "" Then
        osort1$ = tso$
    End If
    GoTo rsupdate3
End If
tempsort$ = Mid$(tso$, firstspace% + 1)
If Left$(tempsort$, 1) = " " Then
    tempsort$ = Mid$(tempsort$, 2)
End If
tso$ = Left$(tso$, firstspace% - 1)
If Right$(tso$, 1) = " " Then
    tso$ = Left$(tso$, Len(tso$) - 1)
End If
tempsort$ = tempsort$ + " " + tso$
If osort1$ = "" Then
    osort1$ = tempsort$
End If
If InStr(osort1$, "JR.") Then
    If Mid$(osort1$, InStr(osort1$, "JR.") + 3, 1) = " " Then
        osort1$ = Left$(osort1$, InStr(osort1$, "JR.") - 1) + Mid$(osort1$, InStr(osort1$, "JR.") + 4) + ", JR."
    Else
        osort1$ = Left$(osort1$, InStr(osort1$, "JR.") - 1) + Mid$(osort1$, InStr(osort1$, "JR.") + 3) + ", JR."
End If
End If
If InStr(osort1$, "SR.") Then
    If Mid$(osort1$, InStr(osort1$, "SR.") + 3, 1) = " " Then
        osort1$ = Left$(osort1$, InStr(osort1$, "SR.") - 1) + Mid$(osort1$, InStr(osort1$, "SR.") + 4) + ", SR."
    Else
        osort1$ = Left$(osort1$, InStr(osort1$, "SR.") - 1) + Mid$(osort1$, InStr(osort1$, "SR.") + 3) + ", SR."
End If
End If
If InStr(osort1$, "III") Then
    If Mid$(osort1$, InStr(osort1$, "III") + 3, 1) = " " Then
        osort1$ = Left$(osort1$, InStr(osort1$, "III") - 1) + Mid$(osort1$, InStr(osort1$, "III") + 4) + ", III"
    Else
        osort1$ = Left$(osort1$, InStr(osort1$, "III") - 1) + Mid$(osort1$, InStr(osort1$, "III") + 3) + ", III"
    End If
End If
If InStr(osort1$, "IV") Then
    If Mid$(osort1$, InStr(osort1$, "IV") + 2, 1) = " " Then
        osort1$ = Left$(osort1$, InStr(osort1$, "IV") - 1) + Mid$(osort1$, InStr(osort1$, "IV") + 3) + ", III"
    Else
        osort1$ = Left$(osort1$, InStr(osort1$, "IV") - 1) + Mid$(osort1$, InStr(osort1$, "IV") + 2) + ", III"
    End If
End If
If Left$(osort1$, 1) = " " Then
    osort1$ = Mid$(osort1$, 2)
End If
rsupdate3:
rs("dpname") = osort1$
rs.Update

db.Close
Call clearrtn
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume
End If
End Sub
Private Sub findrtn()
If casenumber = "" Then
    msg = MsgBox("A valid case number must be entered.", 48, "Genesis Error Log")
    Exit Sub
End If
On Error Resume Next
Dim db As Database, rs As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nwi + "incident.mdb")
Set rs = db.OpenRecordset("select * from badcheck where incidentnumber =" + Chr$(34) + casenumber + Chr$(34))
If Not rs.EOF Then
    rs.MoveFirst
Else
    On Error Resume Next
    db.Close
    msg = MsgBox("Case Number not found.", 48, "Genesis Information Log")
    Exit Sub
End If
On Error Resume Next
casenumber = rs("incidentnumber")
cname = rs("cname")
caddress = rs("caddress")
If Not IsNull(rs("ccity")) Then
    CCITY = rs("ccity")
End If
If Not IsNull(rs("cstate")) Then
    CSTATE = rs("cstate")
End If
If Not IsNull(rs("czipcode")) Then
    CZIPCODE = rs("czipcode")
End If
cphone = rs("cphone")
IncidentDate = rs("incidentdate")
IncidentTime = rs("incidenttime")
dateofoffense = rs("dateofoffense")
incidentlocation = rs("incidentlocation")
highway = rs("highway")
Commercial = rs("commercial")
scvstation = rs("scvstation")
chainstore = rs("chainstore")
residence = rs("residence")
bank = rs("bank")
Other = rs("other")
otherspecify = rs("otherspecify")
vname = rs("vname")
vaddress = rs("vaddress")
If Not IsNull(rs("vcity")) Then
    vcity = rs("vcity")
End If
If Not IsNull(rs("vstate")) Then
    vstate = rs("vstate")
End If
If Not IsNull(rs("vzipcode")) Then
    vzipcode = rs("vzipcode")
End If
vrace.ListIndex = -1
For t% = 0 To vrace.ListCount - 1
    If rs("vrace") = Left$(vrace.List(t%), 1) Then
        vrace.ListIndex = t%
        t% = vrace.ListCount - 1
    End If
Next t%
vsex.ListIndex = -1
For t% = 0 To vsex.ListCount - 1
    If rs("vsex") = Left$(vsex.List(t%), 1) Then
        vsex.ListIndex = t%
        t% = vsex.ListCount - 1
    End If
Next t%
vage = rs("vage")
vhphone = rs("vhphone")
vwphone = rs("vwphone")
suspect = rs("suspect")
wanted = rs("wanted")
warrant = rs("warrant")
arrest = rs("arrest")
sname = rs("sname")
saddress = rs("saddress")
If Not IsNull(rs("scity")) Then
    scity = rs("scity")
End If
If Not IsNull(rs("sstate")) Then
    sstate = rs("sstate")
End If
If Not IsNull(rs("szipcode")) Then
    szipcode = rs("szipcode")
End If
srace.ListIndex = -1
For t% = 0 To srace.ListCount - 1
    If rs("srace") = Left$(srace.List(t%), 1) Then
        srace.ListIndex = t%
        t% = srace.ListCount - 1
    End If
Next t%
ssex.ListIndex = -1
For t% = 0 To ssex.ListCount - 1
    If rs("ssex") = Left$(ssex.List(t%), 1) Then
        ssex.ListIndex = t%
        t% = ssex.ListCount - 1
    End If
Next t%
sage = rs("sage")
drivers = rs("drivers")
driversstate = rs("driversstate")
ssn = rs("ssn")
idnumber = rs("idnumber")
sbirthdate = rs("sdateofbirth")
sheight = rs("sheight")
sweight = rs("sweight")
SHAIR = rs("shair")
SEYES = rs("seyes")
totalarrested = rs("totalarrested")
nearyes = rs("nearyes")
nearno = rs("nearno")
checknumbers = rs("checknumbers")
checkamount = rs("checkamount")
jurisdiction = rs("jurisdiction")
bankname = rs("bankname")
status = rs("status")
comments.Text = rs("comments")
theft = rs("theft")
recovery = rs("recovery")
active = rs("active")
cleared = rs("cleared")
reportingofficer.ListIndex = -1
For t% = 0 To reportingofficer.ListCount - 1
    If rs("reportingofficer") = reportingofficer.List(t%) Then
        reportingofficer.ListIndex = t%
        t% = reportingofficer.ListCount - 1
    End If
Next t%
reportingofficernumber = rs("reportingofficerUNIT")
approvingofficer.ListIndex = -1
For t% = 0 To approvingofficer.ListCount - 1
    If rs("approvingofficer") = approvingofficer.List(t%) Then
        approvingofficer.ListIndex = t%
        t% = approvingofficer.ListCount - 1
    End If
Next t%
approvingofficernumber = rs("approvingofficerUNIT")
ssql = "select mugshot from people where dpnamelf = '" + sname + "'"
If ssn > "" Then
    ssql = ssql + " and ssn = '" + ssn + "'"
End If
If idnumber > "" Then
    ssql = ssql + " and idnumber = '" + idnumber + "'"
End If
If IsDate(sbirthdate) Then
    ssql = ssql + " and birthdate = #" + sbirthdate + "#"
End If
Set db = OpenDatabase(nwl + "lawsuite.mdb")
Set rs = db.OpenRecordset(ssql)
If Not rs.EOF Then
    rs.MoveFirst
    If Not IsNull(rs("mUgshot")) Then
        mugshot.Picture = LoadPicture(rs("mugshot"))
    End If
End If
db.Close
On Error GoTo 0
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If
End Sub
Private Sub deletertn()
If casenumber = "" Then
    msg = MsgBox("A valid case number must be entered.", 48, "Genesis Error Log")
    Exit Sub
End If
msg = MsgBox("Are you sure you want to delete this record?", 4, "Genesis Information Log")
If msg = 7 Then
    Exit Sub
End If
Dim db As Database, rs As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nwi + "incident.mdb")
Set rs = db.OpenRecordset("select * from badcheck where incidentnumber =" + Chr$(34) + casenumber + Chr$(34))
If Not rs.EOF Then
    rs.MoveFirst
    rs.Delete
End If
db.Close
On Error Resume Next
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If
End Sub
Private Sub printrtn()
report.ReportFileName = nwi + "badcheck.RPT"
report.SelectionFormula = "{badcheck.incidentNUMBER} = '" + casenumber + "'"
report.Action = 1
End Sub

Private Sub vname_Click()
Dim db As Database, ds As Recordset
On Error GoTo oderror
od:
Call setpopup(vname, "L")
Set db = OpenDatabase(nwl + "LAWSUITE.MDB")
Set ds = db.OpenRecordset("select * from PEOPLE where dpnameLF = " + Chr$(34) + vname + Chr$(34))
If Not ds.EOF Then
   ds.MoveFirst
    If Not IsNull(ds("dphaddress")) Then
        vaddress = ds("dphaddress")
    Else
        vaddress = ""
    End If
    If Not IsNull(ds("dphaddress2")) Then
        vcity = ds("dphaddress2")
    End If
    If Not IsNull(ds("Hstate")) Then
        vstate = ds("Hstate")
    End If
    If Not IsNull(ds("Hzipcode")) Then
        vzipcode = ds("Hzipcode")
    End If
    If Not IsNull(ds("dphphone")) And ds("dphphone") <> "" Then
        vphone = ds("dphphone")
    Else
        vphone = ""
    End If
End If
Exit Sub
db.Close
On Error GoTo 0
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If

End Sub

Private Sub vname_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 And Shift = vbCtrlMask Then
    If nametype = 0 Then
        cleanup.fname = Me.ActiveControl.Text
        cleanup.lname = ""
    Else
        cleanup.lname = Me.ActiveControl.Text
        cleanup.fname = ""
    End If
    cleanup.Show
End If
End Sub

Private Sub vname_LostFocus()
If Len(vname) > 50 Then
    vname = Left$(vname, 50)
End If

End Sub
