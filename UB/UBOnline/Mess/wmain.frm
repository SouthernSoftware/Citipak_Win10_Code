VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form warrant 
   Appearance      =   0  'Flat
   BackColor       =   &H00800000&
   Caption         =   "Genesis Warrant Book"
   ClientHeight    =   7590
   ClientLeft      =   75
   ClientTop       =   885
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7590
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame issueframe 
      Caption         =   "Issue Warrant Information Frame"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   4035
      Left            =   11520
      TabIndex        =   147
      Top             =   5280
      Visible         =   0   'False
      Width           =   8340
      Begin VB.CommandButton SpellCk 
         Caption         =   "Spelling"
         Height          =   195
         Index           =   2
         Left            =   7125
         TabIndex        =   159
         Top             =   2130
         Width           =   1125
      End
      Begin VB.CommandButton SpellCk 
         Caption         =   "Spelling"
         Height          =   195
         Index           =   1
         Left            =   3000
         TabIndex        =   157
         Top             =   2160
         Width           =   1125
      End
      Begin VB.TextBox mdesc 
         Height          =   300
         Left            =   2700
         MaxLength       =   30
         TabIndex        =   158
         Top             =   1695
         Width           =   2655
      End
      Begin VB.TextBox offensedate 
         Height          =   300
         Left            =   5535
         MaxLength       =   20
         TabIndex        =   154
         Top             =   1230
         Width           =   2655
      End
      Begin VB.TextBox defendant 
         Height          =   300
         Left            =   5190
         MaxLength       =   50
         TabIndex        =   153
         Top             =   765
         Width           =   3000
      End
      Begin VB.TextBox plaintiff 
         Height          =   300
         Left            =   5190
         MaxLength       =   50
         TabIndex        =   152
         Top             =   330
         Width           =   3000
      End
      Begin RichTextLib.RichTextBox facts 
         Height          =   1140
         Left            =   4275
         TabIndex        =   161
         Top             =   2295
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   2011
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"wmain.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox offense 
         Height          =   1140
         Left            =   150
         TabIndex        =   160
         Top             =   2340
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   2011
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"wmain.frx":0084
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.OptionButton Municipality 
         Caption         =   "Municipality"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   1260
         TabIndex        =   156
         Top             =   1740
         Width           =   1470
      End
      Begin VB.OptionButton county 
         Caption         =   "County"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   210
         TabIndex        =   155
         Top             =   1740
         Width           =   1290
      End
      Begin VB.CommandButton Command24 
         Caption         =   "CANCEL"
         Height          =   480
         Left            =   6165
         TabIndex        =   163
         Top             =   3495
         Width           =   2085
      End
      Begin VB.CommandButton Command23 
         Caption         =   "PRINT WARRANT"
         Height          =   480
         Left            =   180
         TabIndex        =   162
         Top             =   3510
         Width           =   2085
      End
      Begin VB.TextBox judge 
         Height          =   300
         Left            =   1050
         MaxLength       =   50
         TabIndex        =   151
         Top             =   1245
         Width           =   3000
      End
      Begin VB.TextBox ind 
         Height          =   300
         Left            =   1050
         MaxLength       =   50
         TabIndex        =   150
         Top             =   810
         Width           =   3000
      End
      Begin VB.TextBox court 
         Height          =   300
         Left            =   1050
         MaxLength       =   50
         TabIndex        =   149
         Top             =   375
         Width           =   3000
      End
      Begin VB.Label Label41 
         Caption         =   "Offense Dates:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   4245
         TabIndex        =   170
         Top             =   1275
         Width           =   1680
      End
      Begin VB.Label Label40 
         Caption         =   "Defendant:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   4245
         TabIndex        =   169
         Top             =   810
         Width           =   1680
      End
      Begin VB.Label Label39 
         Caption         =   "Plaintiff:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   4245
         TabIndex        =   168
         Top             =   375
         Width           =   1680
      End
      Begin VB.Label Label38 
         Caption         =   "Supporting Facts:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   4320
         TabIndex        =   167
         Top             =   2055
         Width           =   1680
      End
      Begin VB.Label Label37 
         Caption         =   "Offense Description:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   195
         TabIndex        =   166
         Top             =   2085
         Width           =   1680
      End
      Begin VB.Label Label36 
         Caption         =   "Judge:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   210
         TabIndex        =   165
         Top             =   1260
         Width           =   1680
      End
      Begin VB.Label Label29 
         Caption         =   "Ind#:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   225
         TabIndex        =   164
         Top             =   840
         Width           =   1680
      End
      Begin VB.Label Label25 
         Caption         =   "Court:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   225
         TabIndex        =   148
         Top             =   420
         Width           =   1680
      End
   End
   Begin VB.Frame Frame100 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   8370
      Left            =   15
      TabIndex        =   56
      Top             =   0
      Width           =   11655
      Begin VB.Frame lookupframe 
         BackColor       =   &H00800000&
         Caption         =   "Look Up Frame"
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
         Height          =   2895
         Left            =   11520
         TabIndex        =   104
         Top             =   5760
         Visible         =   0   'False
         Width           =   6495
         Begin MSComctlLib.ListView lookuplist 
            Height          =   2535
            Left            =   120
            TabIndex        =   105
            Top             =   240
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   4471
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Warrant"
               Object.Width           =   2822
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Log Date"
               Object.Width           =   2293
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Name"
               Object.Width           =   5292
            EndProperty
         End
      End
      Begin VB.Frame POFRAME 
         BackColor       =   &H00800000&
         Caption         =   "Print Options"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   5520
         Left            =   5160
         TabIndex        =   61
         Top             =   7560
         Visible         =   0   'False
         Width           =   7605
         Begin VB.CommandButton Command22 
            Caption         =   "Issue Arrest Warrant"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3840
            TabIndex        =   81
            Top             =   4305
            Width           =   3500
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Issue Bench Warrant"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   105
            TabIndex        =   71
            Top             =   4320
            Width           =   3500
         End
         Begin VB.CommandButton Command21 
            Caption         =   "Pending Non-County Warrant Summary"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3840
            TabIndex        =   80
            Top             =   3850
            Width           =   3500
         End
         Begin VB.CommandButton Command20 
            Caption         =   "Pending Warrant Summary by Area"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3840
            TabIndex        =   78
            Top             =   2950
            Width           =   3500
         End
         Begin VB.CommandButton Command19 
            Caption         =   "Pending County Warrant Summary"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3840
            TabIndex        =   79
            Top             =   3400
            Width           =   3500
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Current Outstanding Warrant"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   105
            TabIndex        =   62
            Top             =   250
            Width           =   3500
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Sent Warrants"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3840
            TabIndex        =   73
            Top             =   700
            Width           =   3500
         End
         Begin VB.CommandButton rowb 
            Caption         =   "Record of Warrants"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3840
            TabIndex        =   72
            Top             =   250
            Width           =   3500
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Served 4D Warrant Report"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   105
            TabIndex        =   70
            Top             =   3850
            Width           =   3500
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Recalled Warrant Report"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   105
            TabIndex        =   69
            Top             =   3400
            Width           =   3500
         End
         Begin VB.CommandButton swbdbd 
            Caption         =   "Served Warrant By Deputy By Date"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   105
            TabIndex        =   68
            Top             =   2950
            Width           =   3500
         End
         Begin VB.CommandButton owbd 
            Caption         =   "Oustanding Warrant By Deputy"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   105
            TabIndex        =   67
            Top             =   2500
            Width           =   3500
         End
         Begin VB.CommandButton swbd 
            Caption         =   "Served Warrant By Date"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   105
            TabIndex        =   66
            Top             =   2050
            Width           =   3500
         End
         Begin VB.TextBox numcopies 
            Height          =   285
            Left            =   1770
            MaxLength       =   2
            TabIndex        =   82
            Text            =   "01"
            Top             =   5025
            Width           =   375
         End
         Begin VB.CommandButton ow 
            Caption         =   "Outstanding Warrant"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   105
            TabIndex        =   65
            Top             =   1600
            Width           =   3500
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Close"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   510
            Left            =   6570
            TabIndex        =   87
            Top             =   4860
            Width           =   735
         End
         Begin VB.CommandButton dow 
            Caption         =   "Current Oustanding Warrant-Detail"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   105
            TabIndex        =   63
            Top             =   700
            Width           =   3500
         End
         Begin VB.CommandButton Command7 
            Caption         =   "Outstanding Bench Warrant"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   105
            TabIndex        =   64
            Top             =   1150
            Width           =   3500
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   525
            Left            =   2685
            TabIndex        =   89
            Top             =   4860
            Width           =   1695
            Begin VB.OptionButton oarea1 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Order By Area 1"
               Height          =   195
               Left            =   30
               TabIndex        =   83
               Top             =   0
               Width           =   1815
            End
            Begin VB.OptionButton oarea2 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Order By Area 2"
               Height          =   195
               Left            =   45
               TabIndex        =   84
               Top             =   255
               Width           =   1815
            End
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   525
            Left            =   4545
            TabIndex        =   88
            Top             =   4860
            Width           =   1575
            Begin VB.OptionButton owarant 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Warrant Order"
               Height          =   195
               Left            =   0
               TabIndex        =   86
               Top             =   240
               Width           =   1815
            End
            Begin VB.OptionButton oalpha 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Alpha Order"
               Height          =   195
               Left            =   0
               TabIndex        =   85
               Top             =   0
               Width           =   1815
            End
         End
         Begin VB.CommandButton Command9 
            Caption         =   "Returned Warrants"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3840
            TabIndex        =   74
            Top             =   1150
            Width           =   3500
         End
         Begin VB.CommandButton Command10 
            Caption         =   "Pending Warrant Summary"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3840
            TabIndex        =   75
            Top             =   1600
            Width           =   3500
         End
         Begin VB.CommandButton Command14 
            Caption         =   "Alpha Pending Warrant Summary"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3840
            TabIndex        =   76
            Top             =   2050
            Width           =   3500
         End
         Begin VB.CommandButton Command15 
            Caption         =   "Pending Warrant Summary by Address"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3840
            TabIndex        =   77
            Top             =   2500
            Width           =   3500
         End
         Begin VB.Label Label20 
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Number of Copies:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   165
            TabIndex        =   90
            Top             =   5040
            Width           =   1575
         End
      End
      Begin VB.Frame icframe 
         BackColor       =   &H00808080&
         Caption         =   "Issued By Cleanup Frame"
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
         Height          =   2535
         Left            =   0
         TabIndex        =   98
         Top             =   7635
         Visible         =   0   'False
         Width           =   7215
         Begin VB.TextBox changeissued 
            Height          =   285
            Left            =   120
            MaxLength       =   75
            TabIndex        =   102
            Top             =   2130
            Width           =   5415
         End
         Begin VB.CommandButton Command16 
            BackColor       =   &H00808080&
            Caption         =   "Close"
            Height          =   615
            Left            =   5640
            Style           =   1  'Graphical
            TabIndex        =   101
            Top             =   1800
            Width           =   1455
         End
         Begin VB.CommandButton Command17 
            BackColor       =   &H00808080&
            Caption         =   "Change To"
            Height          =   615
            Left            =   5640
            Style           =   1  'Graphical
            TabIndex        =   100
            Top             =   240
            Width           =   1455
         End
         Begin VB.ListBox ic 
            Height          =   1620
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   99
            Top             =   240
            Width           =   5415
         End
         Begin VB.Label Label28 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Change To"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   270
            Left            =   120
            TabIndex        =   103
            Top             =   1890
            Width           =   2010
         End
      End
      Begin VB.Frame ccframe 
         BackColor       =   &H00800000&
         Caption         =   "Charge Cleanup Frame"
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
         Height          =   2535
         Left            =   1080
         TabIndex        =   106
         Top             =   2040
         Visible         =   0   'False
         Width           =   7215
         Begin VB.ListBox cc 
            Height          =   1620
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   110
            Top             =   240
            Width           =   5415
         End
         Begin VB.CommandButton Command12 
            Caption         =   "Change To"
            Height          =   615
            Left            =   5640
            TabIndex        =   109
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton Command13 
            Caption         =   "Close"
            Height          =   615
            Left            =   5640
            TabIndex        =   108
            Top             =   1800
            Width           =   1455
         End
         Begin VB.TextBox changeto 
            Height          =   285
            Left            =   120
            MaxLength       =   75
            TabIndex        =   107
            Top             =   2130
            Width           =   5415
         End
         Begin VB.Label Label27 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Change To"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Left            =   120
            TabIndex        =   111
            Top             =   1920
            Width           =   2010
         End
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00C0C0C0&
         Height          =   2445
         Left            =   75
         TabIndex        =   134
         Top             =   5055
         Width           =   8280
         Begin VB.CheckBox ccounty 
            BackColor       =   &H00C0C0C0&
            Caption         =   "County"
            ForeColor       =   &H00800000&
            Height          =   345
            Left            =   4350
            TabIndex        =   51
            Top             =   1995
            Width           =   930
         End
         Begin VB.CheckBox NODISP 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Never Display on Current List"
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   6180
            TabIndex        =   49
            Top             =   1575
            Width           =   1650
         End
         Begin VB.CheckBox keeplist 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Keep on Current List Until Served"
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   4350
            TabIndex        =   48
            Top             =   1560
            Width           =   1680
         End
         Begin VB.TextBox witness 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   780
            Left            =   1650
            MaxLength       =   50
            TabIndex        =   45
            Top             =   1560
            Width           =   2610
         End
         Begin VB.TextBox whenarrested 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   1650
            MaxLength       =   20
            TabIndex        =   44
            Top             =   1245
            Width           =   1290
         End
         Begin VB.CheckBox recalled 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Recalled On"
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
            Height          =   300
            Left            =   4785
            TabIndex        =   46
            Top             =   1170
            Width           =   1395
         End
         Begin VB.TextBox recalldate 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   6240
            TabIndex        =   47
            Top             =   1200
            Width           =   1830
         End
         Begin VB.ComboBox officer 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   345
            Left            =   1650
            Sorted          =   -1  'True
            TabIndex        =   39
            Top             =   810
            Width           =   3030
         End
         Begin VB.TextBox sentto 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   4740
            MaxLength       =   30
            TabIndex        =   40
            Top             =   315
            Width           =   1710
         End
         Begin VB.TextBox senton 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   6615
            MaxLength       =   10
            TabIndex        =   41
            Top             =   300
            Width           =   1425
         End
         Begin VB.ComboBox assignedto 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   345
            Left            =   1635
            Sorted          =   -1  'True
            TabIndex        =   38
            Top             =   315
            Width           =   3045
         End
         Begin VB.TextBox origination 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   6060
            MaxLength       =   30
            TabIndex        =   43
            Top             =   840
            Width           =   1995
         End
         Begin VB.TextBox returnedon 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   4770
            MaxLength       =   10
            TabIndex        =   42
            Top             =   855
            Width           =   1170
         End
         Begin VB.Label Label9 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Prosecuting Witness(es):"
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
            Left            =   90
            TabIndex        =   142
            Top             =   1575
            Width           =   1230
         End
         Begin VB.Label Label10 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "When Arrested:"
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
            Height          =   315
            Left            =   60
            TabIndex        =   141
            Top             =   1275
            Width           =   2055
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Arresting Officer:"
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
            Height          =   285
            Left            =   45
            TabIndex        =   140
            Top             =   810
            Width           =   2655
         End
         Begin VB.Label Label11 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Warrant Sent To:"
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
            Height          =   405
            Left            =   4755
            TabIndex        =   139
            Top             =   105
            Width           =   1710
         End
         Begin VB.Label Label12 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "On:"
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
            Height          =   330
            Left            =   6630
            TabIndex        =   138
            Top             =   75
            Width           =   315
         End
         Begin VB.Label Label21 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Assigned To:"
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
            Left            =   45
            TabIndex        =   137
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label22 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Origination:"
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
            Height          =   315
            Left            =   6075
            TabIndex        =   136
            Top             =   630
            Width           =   1095
         End
         Begin VB.Label Label26 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Returned On:"
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
            Height          =   330
            Left            =   4755
            TabIndex        =   135
            Top             =   615
            Width           =   1380
         End
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H00C0C0C0&
         Height          =   5790
         Left            =   8370
         TabIndex        =   143
         Top             =   525
         Width           =   3210
         Begin VB.CommandButton SpellCk 
            Caption         =   "Spelling"
            Height          =   195
            Index           =   3
            Left            =   1965
            TabIndex        =   50
            Top             =   120
            Width           =   1125
         End
         Begin VB.CommandButton SpellCk 
            Caption         =   "Spelling"
            Height          =   195
            Index           =   0
            Left            =   2025
            TabIndex        =   52
            Top             =   1785
            Width           =   1125
         End
         Begin VB.TextBox remarks 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   1425
            Left            =   60
            MaxLength       =   250
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   54
            Top             =   1965
            Width           =   3075
         End
         Begin VB.TextBox iddata 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   1440
            Left            =   60
            MaxLength       =   100
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   53
            Top             =   315
            Width           =   3075
         End
         Begin VB.CheckBox danger 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Subject Dangerous"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   300
            Left            =   675
            TabIndex        =   55
            Top             =   3390
            Width           =   1965
         End
         Begin VB.Image MUGSHOT 
            BorderStyle     =   1  'Fixed Single
            Height          =   1995
            Left            =   615
            Stretch         =   -1  'True
            Top             =   3690
            Width           =   1995
         End
         Begin VB.Label Label14 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Remarks:"
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
            Height          =   360
            Left            =   60
            TabIndex        =   145
            Top             =   1755
            Width           =   900
         End
         Begin VB.Label Label15 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Identifying Data:"
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
            Height          =   315
            Left            =   30
            TabIndex        =   144
            Top             =   120
            Width           =   1950
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00C0C0C0&
         Caption         =   "OPTIONAL DESCRIPTION INFORMATION"
         ForeColor       =   &H00800000&
         Height          =   1185
         Left            =   3675
         TabIndex        =   131
         Top             =   3885
         Width           =   4680
         Begin VB.TextBox AREA1 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   915
            MaxLength       =   20
            TabIndex        =   36
            Top             =   345
            Width           =   3630
         End
         Begin VB.TextBox AREA2 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Left            =   915
            MaxLength       =   20
            TabIndex        =   37
            Top             =   735
            Width           =   3630
         End
         Begin VB.Label Label35 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Area 1:"
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
            Height          =   225
            Left            =   45
            TabIndex        =   133
            Top             =   420
            Width           =   1095
         End
         Begin VB.Label Label34 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Area 2:"
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
            Height          =   225
            Left            =   30
            TabIndex        =   132
            Top             =   765
            Width           =   765
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00C0C0C0&
         Height          =   3150
         Left            =   3660
         TabIndex        =   122
         Top             =   525
         Width           =   4665
         Begin VB.CommandButton Command18 
            Caption         =   "CleanUp"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3840
            TabIndex        =   29
            Top             =   675
            Width           =   750
         End
         Begin VB.ComboBox charge 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   345
            Left            =   1005
            Sorted          =   -1  'True
            TabIndex        =   26
            Top             =   195
            Width           =   2760
         End
         Begin VB.ComboBox issuedby 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   345
            Left            =   1005
            Sorted          =   -1  'True
            TabIndex        =   28
            Top             =   720
            Width           =   2775
         End
         Begin VB.TextBox birthdate 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   1020
            MaxLength       =   10
            TabIndex        =   30
            Top             =   1170
            Width           =   1035
         End
         Begin VB.TextBox casenumber 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   990
            MaxLength       =   20
            TabIndex        =   31
            Top             =   1545
            Width           =   3555
         End
         Begin VB.TextBox docket 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   990
            MaxLength       =   20
            TabIndex        =   32
            Top             =   1890
            Width           =   3540
         End
         Begin VB.TextBox dss 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   975
            MaxLength       =   20
            TabIndex        =   33
            Top             =   2280
            Width           =   1560
         End
         Begin VB.TextBox courtdate 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   3615
            MaxLength       =   10
            TabIndex        =   34
            Top             =   2265
            Width           =   915
         End
         Begin VB.TextBox guardian 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   990
            MaxLength       =   40
            TabIndex        =   35
            Top             =   2685
            Width           =   3570
         End
         Begin VB.CommandButton Command11 
            Caption         =   "CleanUp"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3810
            TabIndex        =   27
            Top             =   195
            Width           =   750
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Charge:"
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
            Height          =   330
            Left            =   90
            TabIndex        =   130
            Top             =   225
            Width           =   1575
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Issued By:"
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
            Height          =   570
            Left            =   90
            TabIndex        =   129
            Top             =   675
            Width           =   915
         End
         Begin VB.Label Label7 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Birthdate:"
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
            Height          =   270
            Left            =   90
            TabIndex        =   128
            Top             =   1185
            Width           =   1365
         End
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Case#:"
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
            Height          =   270
            Left            =   255
            TabIndex        =   127
            Top             =   1575
            Width           =   810
         End
         Begin VB.Label Label16 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Docket#:"
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
            Height          =   285
            Left            =   60
            TabIndex        =   126
            Top             =   1890
            Width           =   1575
         End
         Begin VB.Label Label17 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "DSS#:"
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
            Height          =   285
            Left            =   90
            TabIndex        =   125
            Top             =   2250
            Width           =   735
         End
         Begin VB.Label Label18 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Court Date:"
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
            Left            =   2595
            TabIndex        =   124
            Top             =   2265
            Width           =   1290
         End
         Begin VB.Label Label19 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Custodial Guardian:"
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
            Left            =   60
            TabIndex        =   123
            Top             =   2595
            Width           =   1215
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0C0&
         Height          =   4545
         Left            =   75
         TabIndex        =   112
         Top             =   525
         Width           =   3600
         Begin VB.TextBox address 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   3
            Left            =   2850
            MaxLength       =   10
            TabIndex        =   6
            Top             =   405
            Width           =   705
         End
         Begin VB.TextBox address 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   2
            Left            =   2385
            MaxLength       =   2
            TabIndex        =   5
            Top             =   405
            Width           =   420
         End
         Begin VB.TextBox idnumber 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   1245
            MaxLength       =   15
            TabIndex        =   21
            Top             =   3810
            Width           =   795
         End
         Begin VB.TextBox ssn 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   45
            MaxLength       =   15
            TabIndex        =   20
            Top             =   3810
            Width           =   1155
         End
         Begin VB.TextBox ht 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   2070
            MaxLength       =   10
            TabIndex        =   22
            Top             =   3810
            Width           =   705
         End
         Begin VB.TextBox weight 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   2820
            MaxLength       =   10
            TabIndex        =   23
            Top             =   3810
            Width           =   705
         End
         Begin VB.TextBox eyes 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   2295
            MaxLength       =   10
            TabIndex        =   25
            Top             =   4185
            Width           =   1230
         End
         Begin VB.TextBox hair 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   540
            MaxLength       =   10
            TabIndex        =   24
            Top             =   4185
            Width           =   1215
         End
         Begin VB.Frame Frame8 
            Appearance      =   0  'Flat
            BackColor       =   &H00800000&
            Caption         =   "SEX"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFFF&
            Height          =   850
            Left            =   30
            TabIndex        =   115
            Top             =   2730
            Width           =   1750
            Begin VB.OptionButton female 
               Appearance      =   0  'Flat
               BackColor       =   &H00800000&
               Caption         =   "Female"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   345
               Left            =   120
               TabIndex        =   14
               Top             =   480
               Width           =   1170
            End
            Begin VB.OptionButton male 
               Appearance      =   0  'Flat
               BackColor       =   &H00800000&
               Caption         =   "Male"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   345
               Left            =   120
               TabIndex        =   13
               Top             =   195
               Value           =   -1  'True
               Width           =   1095
            End
         End
         Begin VB.Frame Frame2 
            Appearance      =   0  'Flat
            BackColor       =   &H00800000&
            Caption         =   "RACE"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFFF&
            Height          =   1935
            Left            =   15
            TabIndex        =   114
            Top             =   765
            Width           =   1750
            Begin VB.OptionButton indian 
               Appearance      =   0  'Flat
               BackColor       =   &H00800000&
               Caption         =   "Indian-Amer."
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   25
               TabIndex        =   10
               Top             =   1150
               Width           =   1600
            End
            Begin VB.TextBox otherrace 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   330
               Left            =   960
               MaxLength       =   20
               TabIndex        =   12
               Top             =   1680
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.OptionButton other 
               Appearance      =   0  'Flat
               BackColor       =   &H00800000&
               Caption         =   "Unknown"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   345
               Left            =   25
               TabIndex        =   11
               Top             =   1440
               Width           =   1150
            End
            Begin VB.OptionButton Oriental 
               Appearance      =   0  'Flat
               BackColor       =   &H00800000&
               Caption         =   "Asian/Pacific Isl"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   25
               TabIndex        =   9
               Top             =   840
               Width           =   1635
            End
            Begin VB.OptionButton caucasian 
               Appearance      =   0  'Flat
               BackColor       =   &H00800000&
               Caption         =   "Caucasian"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   330
               Left            =   25
               TabIndex        =   7
               Top             =   240
               Value           =   -1  'True
               Width           =   1150
            End
            Begin VB.OptionButton africanamerican 
               Appearance      =   0  'Flat
               BackColor       =   &H00800000&
               Caption         =   "AfricanAmerican"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   375
               Left            =   25
               TabIndex        =   8
               Top             =   480
               Width           =   1700
            End
         End
         Begin VB.TextBox address 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   0
            Left            =   900
            MaxLength       =   40
            TabIndex        =   3
            Top             =   120
            Width           =   2655
         End
         Begin VB.Frame Frame3 
            Appearance      =   0  'Flat
            BackColor       =   &H00800000&
            Caption         =   "TYPE"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFFF&
            Height          =   2805
            Left            =   1800
            TabIndex        =   113
            Top             =   780
            Width           =   1755
            Begin VB.OptionButton w4d 
               Appearance      =   0  'Flat
               BackColor       =   &H00800000&
               Caption         =   "4D Warrant"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   255
               Left            =   120
               TabIndex        =   19
               Top             =   2340
               Width           =   1455
            End
            Begin VB.OptionButton regw 
               Appearance      =   0  'Flat
               BackColor       =   &H00800000&
               Caption         =   "Warrant"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   255
               Left            =   120
               TabIndex        =   18
               Top             =   2040
               Width           =   1455
            End
            Begin VB.OptionButton benchgs 
               Appearance      =   0  'Flat
               BackColor       =   &H00800000&
               Caption         =   "General Sessions Bench Warrant"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   690
               Left            =   120
               TabIndex        =   17
               Top             =   1275
               Width           =   1530
            End
            Begin VB.OptionButton benchm 
               Appearance      =   0  'Flat
               BackColor       =   &H00800000&
               Caption         =   "Magistrate Bench Warrant"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   510
               Left            =   120
               TabIndex        =   16
               Top             =   735
               Width           =   1545
            End
            Begin VB.OptionButton benchfc 
               Appearance      =   0  'Flat
               BackColor       =   &H00800000&
               Caption         =   "Family Court Bench Warrant"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   480
               Left            =   120
               TabIndex        =   15
               Top             =   240
               Value           =   -1  'True
               Width           =   1600
            End
         End
         Begin VB.TextBox address 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   1
            Left            =   60
            MaxLength       =   40
            TabIndex        =   4
            Top             =   405
            Width           =   2265
         End
         Begin VB.Label Label24 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "ID#:"
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
            Height          =   240
            Left            =   1260
            TabIndex        =   146
            Top             =   3600
            Width           =   1215
         End
         Begin VB.Label Label33 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Eyes"
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
            Height          =   240
            Left            =   1860
            TabIndex        =   121
            Top             =   4185
            Width           =   720
         End
         Begin VB.Label Label23 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "SSN:"
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
            Height          =   240
            Left            =   60
            TabIndex        =   120
            Top             =   3600
            Width           =   1215
         End
         Begin VB.Label Label30 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Height"
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
            Height          =   240
            Left            =   2070
            TabIndex        =   119
            Top             =   3600
            Width           =   735
         End
         Begin VB.Label Label31 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Weight"
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
            Height          =   240
            Left            =   2790
            TabIndex        =   118
            Top             =   3600
            Width           =   735
         End
         Begin VB.Label Label32 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Hair"
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
            Height          =   240
            Left            =   90
            TabIndex        =   117
            Top             =   4185
            Width           =   720
         End
         Begin VB.Label Label13 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Address:"
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
            Left            =   60
            TabIndex        =   116
            Top             =   120
            Width           =   1335
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00800000&
         BorderStyle     =   0  'None
         Caption         =   "Frame6"
         Height          =   1080
         Left            =   8400
         TabIndex        =   91
         Top             =   6360
         Width           =   3165
         Begin VB.CommandButton Command8 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Searc&h"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1590
            TabIndex        =   97
            Top             =   765
            Width           =   1500
         End
         Begin VB.CommandButton clearbutton 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "&Clear"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   60
            TabIndex        =   96
            Top             =   765
            Width           =   1500
         End
         Begin VB.CommandButton lookupbutton 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "&Lookup"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1605
            TabIndex        =   95
            Top             =   45
            Width           =   1500
         End
         Begin VB.CommandButton printbutton 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "&Print Options"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   75
            TabIndex        =   94
            Top             =   405
            Width           =   1500
         End
         Begin VB.CommandButton deletebutton 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "&Delete"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1605
            TabIndex        =   93
            Top             =   405
            Width           =   1500
         End
         Begin VB.CommandButton savebutton 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "&Save"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   45
            TabIndex        =   92
            Top             =   45
            Width           =   1500
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00800000&
         BorderStyle     =   0  'None
         Caption         =   "Frame7"
         Height          =   465
         Left            =   75
         TabIndex        =   57
         Top             =   60
         Width           =   11490
         Begin VB.ComboBox warrant 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   345
            Left            =   1290
            Sorted          =   -1  'True
            TabIndex        =   0
            Top             =   60
            Width           =   2535
         End
         Begin VB.ComboBox wname 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   345
            Left            =   7095
            Sorted          =   -1  'True
            TabIndex        =   2
            Top             =   60
            Width           =   4395
         End
         Begin VB.TextBox logdate 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Left            =   5025
            MaxLength       =   10
            TabIndex        =   1
            Top             =   60
            Width           =   1125
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "NAME:"
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
            Height          =   480
            Left            =   6360
            TabIndex        =   60
            Top             =   60
            Width           =   615
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "LOG DATE:"
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
            Height          =   480
            Left            =   3900
            TabIndex        =   59
            Top             =   45
            Width           =   1095
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "WARRANT #:"
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
            Height          =   480
            Left            =   15
            TabIndex        =   58
            Top             =   60
            Width           =   1395
         End
      End
      Begin Crystal.CrystalReport report 
         Left            =   2295
         Top             =   7215
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
   End
End
Attribute VB_Name = "warrant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim recentry As Long, justfound As Integer, sedit, sprint, sreport, sbrowse, sdelete, ssupervisor As Integer, FROMXREF As Integer
Dim itmx As ListItem, nametype As Integer

Private Sub assignedtoload()
Dim db As Database, ds As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nwl + "lawsuite.mdb")
Set ds = db.OpenRecordset("select distinct profname from professionals where type = 'D'")
assignedto.clear
officer.clear
If Not ds.EOF Then
    ds.MoveFirst
End If
While Not ds.EOF
    If Not IsNull(ds("profname")) Then
        assignedto.AddItem ds("profname")
        officer.AddItem ds("profname")
    End If
    ds.MoveNext
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

Private Sub printoutstanding()
Screen.MousePointer = 11
REPORT.Destination = 0
REPORT.ReportFileName = nww + "OUTSTAND.RPT"
REPORT.SelectionFormula = "{WARRANTINFO.OUTSTANDING} = '1' and {warrantinfo.warrant} <> 'DELETED'"
REPORT.Action = 1
Screen.MousePointer = 0
End Sub




Private Sub address_GotFocus(index As Integer)
Dim db As Database, rs As Recordset
On Error Resume Next
If address(0) = "" And index = 0 Then
    On Error GoTo oderror
od:
    Set db = OpenDatabase(nwl + "LAWSUITE.MDB")
    Set rs = db.OpenRecordset("SELECT * FROM PEOPLE WHERE DPNAMElf = " + Chr$(34) + wname + Chr$(34))
    If Not rs.EOF Then
        rs.MoveFirst
        address(0) = rs("DPHADDRESS")
        address(1) = rs("DPHADDRESS2")
        address(2) = rs("hstate")
        address(3) = rs("hzipcode")
        If Not IsNull(rs("height")) Then
            ht = rs("height")
        End If
        If Not IsNull(rs("weight")) Then
            weight = rs("weight")
        End If
        If Not IsNull(rs("hair")) Then
            hair = rs("hair")
        End If
        If Not IsNull(rs("eyes")) Then
            eyes = rs("eyes")
        End If
        If Not IsNull(rs("ssn")) Then
            ssn = rs("ssn")
        End If
        If Not IsNull(rs("idnumber")) Then
            idnumber = rs("idnumber")
        End If
        If Not IsNull(rs("birthdate")) Then
            birthdate = rs("birthdate")
        End If
        If Not IsNull(rs("race")) Then
            Select Case rs("race")
                Case "C", "W", "White"
                    caucasian = True
                Case "B", "Black"
                    africanamerican = True
                Case "O", "A", "Asian/Pacific Islander"
                    Oriental = True
                Case "I", "Indian - American Indian/Alaskan Native"
                    indian = True
                Case "U", "Unknown"
                    other = True
            End Select
        End If
        If Not IsNull(rs("sex")) Then
            Select Case rs("sex")
                Case "F", "Female"
                    female = True
                Case "M", "Male"
                    male = True
            End Select
        End If
        If Not IsNull(rs("mugshot")) Then
            mugshot.Picture = LoadPicture(rs("mugshot"))
        Else
            mugshot.Picture = LoadPicture()
        End If
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

Private Sub assignedto_keypress(KeyAscii As Integer)
If Len(assignedto) = 40 Then
    KeyAscii = 0
End If
End Sub

Private Sub benchfc_Click()
If benchfc.Value = True Then
    docket.Enabled = False
    dss.Enabled = False
    guardian.Enabled = False
    courtdate.Enabled = False
    Label16.Enabled = False
    Label17.Enabled = False
    Label18.Enabled = False
    Label19.Enabled = False
End If
End Sub

Private Sub benchgs_Click()
If benchgs.Value = True Then
    docket.Enabled = False
    dss.Enabled = False
    guardian.Enabled = False
    courtdate.Enabled = False
    Label16.Enabled = False
    Label17.Enabled = False
    Label18.Enabled = False
    Label19.Enabled = False
End If

End Sub

Private Sub benchm_Click()
If benchm.Value = True Then
    docket.Enabled = False
    dss.Enabled = False
    guardian.Enabled = False
    courtdate.Enabled = False
    Label16.Enabled = False
    Label17.Enabled = False
    Label18.Enabled = False
    Label19.Enabled = False
End If

End Sub

Private Sub breakup(temp As String, maxlen As Long, T1 As String, T2 As String, T3 As String, T4 As String)
If InStr(temp, " ") = 0 Then
    GoSub nospace
Else
    GoSub hasspace
End If
Exit Sub
nospace:
T1 = ""
T2 = ""
T3 = ""
T4 = ""
ctr% = 1
While Printer.TextWidth(T1) <= maxlen And ctr% <= Len(temp)
    T1 = T1 + Mid$(temp, ctr%, 1)
    ctr% = ctr% + 1
Wend
T1 = Left$(T1, Len(T1) - 1)
If ctr% <= Len(temp) Then
    temp = Mid$(temp, ctr%)
Else
    Return
End If
ctr% = 1
While Printer.TextWidth(T2) <= maxlen And ctr% <= Len(temp)
    T2 = T2 + Mid$(temp, ctr%, 1)
    ctr% = ctr% + 1
Wend
T2 = Left$(T2, Len(T2) - 1)
If ctr% <= Len(temp) Then
    temp = Mid$(temp, ctr%)
Else
    Return
End If

ctr% = 1
While Printer.TextWidth(T3) <= maxlen And ctr% <= Len(temp)
    T3 = T3 + Mid$(temp, ctr%, 1)
    ctr% = ctr% + 1
Wend
T3 = Left$(T3, Len(T3) - 1)
If ctr% <= Len(temp) Then
    temp = Mid$(temp, ctr%)
Else
    Return
End If

ctr% = 1
While Printer.TextWidth(T4) <= maxlen And ctr% <= Len(temp)
    T4 = T4 + Mid$(temp, ctr%, 1)
    ctr% = ctr% + 1
Wend
T4 = Left$(T4, Len(T4) - 1)
If ctr% <= Len(temp) Then
    temp = Mid$(temp, ctr%)
Else
    Return
End If

Return
hasspace:
T1 = ""
T2 = ""
T3 = ""
T4 = ""
TP$ = temp
While Printer.TextWidth(T1) <= maxlen And Len(TP$) > 0
    Last$ = T1
    If InStr(TP$, " ") = 0 Then
        T1 = T1 + TP$
        Last$ = T1
        TP$ = ""
    Else
        T1 = T1 + Left$(TP$, InStr(TP$, " "))
        TP$ = Mid$(TP$, InStr(TP$, " ") + 1)
    End If
Wend
T1 = Last$
Last$ = ""
temp = Mid$(temp, Len(T1) + 1)
If Left$(temp, 1) = " " Then
    temp = Mid$(temp, 2)
End If
TP$ = temp
While Printer.TextWidth(T2) <= maxlen And Len(TP$) > 0
    Last$ = T2
    If InStr(TP$, " ") = 0 Then
        T2 = T2 + TP$
        Last$ = T2
        TP$ = ""
    Else
        T2 = T2 + Left$(TP$, InStr(TP$, " "))
        TP$ = Mid$(TP$, InStr(TP$, " ") + 1)
    End If
Wend
T2 = Last$
Last$ = ""
temp = Mid$(temp, Len(T2) + 1)
If Left$(temp, 1) = " " Then
    temp = Mid$(temp, 2)
End If
TP$ = temp
While Printer.TextWidth(T3) <= maxlen And Len(TP$) > 0
    Last$ = T3
    If InStr(TP$, " ") = 0 Then
        T3 = T3 + TP$
        Last$ = T3
        TP$ = ""
    Else
        T3 = T3 + Left$(TP$, InStr(TP$, " "))
        TP$ = Mid$(TP$, InStr(TP$, " ") + 1)
    End If
Wend
T3 = Last$
Last$ = ""
temp = Mid$(temp, Len(T3) + 1)
If Left$(temp, 1) = " " Then
    temp = Mid$(temp, 2)
End If
TP$ = temp
While Printer.TextWidth(T4) <= maxlen And Len(TP$) > 0
    Last$ = T4
    If InStr(TP$, " ") = 0 Then
        T4 = T4 + TP$
        Last$ = T4
        TP$ = ""
    Else
        T4 = T4 + Left$(TP$, InStr(TP$, " "))
        TP$ = Mid$(TP$, InStr(TP$, " ") + 1)
    End If
Wend
T4 = Last$
Last$ = ""
temp = Mid$(temp, Len(T4) + 1)
If Left$(temp, 1) = " " Then
    temp = Mid$(temp, 2)
End If
Return
End Sub

Private Sub birthdate_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
If Len(birthdate) = 1 Or Len(birthdate) = 4 Then
    SendKeys "/"
End If
End If
End Sub

Private Sub BIRTHDATE_LostFocus()
If IsDate(birthdate) Then
    birthdate = Format$(birthdate, "mm/dd/yyyy")
End If

End Sub

Private Sub charge_KeyPress(KeyAscii As Integer)
If Len(charge) = 75 Then
    KeyAscii = 0
End If
End Sub

Private Sub chargeload()
Dim db As Database, ds As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nww + "warrant.mdb")
Set ds = db.OpenRecordset("select distinct charge from warrantinfo")
charge.clear
If Not ds.EOF Then
    ds.MoveFirst
End If
While Not ds.EOF
    If ds("charge") <> " " Then
        charge.AddItem ds("charge")
    End If
    ds.MoveNext
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

Private Sub clearbutton_Click()
Call nullfields
End Sub

Private Sub Command1_Click()
Screen.MousePointer = 11
If Val(numcopies) = 0 Then
    CopiesToPrinter = 1
Else
    CopiesToPrinter = Val(numcopies)
End If
REPORT.Destination = 0
REPORT.ReportFileName = nww + "recalled.RPT"
REPORT.SelectionFormula = ""
If oalpha Then
    If oarea1 Then
        REPORT.SortFields(1) = "+{warrantinfo.wNAME}"
        REPORT.SortFields(0) = "+{warrantinfo.area1}"
    Else
    If oarea2 Then
        REPORT.SortFields(1) = "+{warrantinfo.wNAME}"
        REPORT.SortFields(0) = "+{warrantinfo.area2}"
    Else
        REPORT.SortFields(0) = "+{warrantinfo.wNAME}"
        REPORT.SortFields(1) = ""
    End If
    End If
Else
    If oarea1 Then
        REPORT.SortFields(1) = "+{warrantinfo.warrant}"
        REPORT.SortFields(0) = "+{warrantinfo.area1}"
    Else
    If oarea2 Then
        REPORT.SortFields(1) = "+{warrantinfo.warrant}"
        REPORT.SortFields(0) = "+{warrantinfo.area2}"
    Else
        REPORT.SortFields(0) = "+{warrantinfo.warrant}"
        REPORT.SortFields(1) = ""
    End If
    End If
End If
REPORT.Action = 1
Screen.MousePointer = 0
'POFRAME.Visible = False
End Sub

Private Sub Command10_Click()
Screen.MousePointer = 11
If Val(numcopies) = 0 Then
    CopiesToPrinter = 1
Else
    CopiesToPrinter = Val(numcopies)
End If
REPORT.Destination = 0
tyy$ = Right$(Date$, 4)
ty = Val(tyy$)
tmm$ = Left$(Date$, 2)
tm = Val(tmm$)
tdd$ = Mid$(Date$, 4, 2)
td = Val(tdd$)
If tm - 1 < 1 Then
    ty = ty - 1
    tm = 12 + tm - 1
Else
    tm = tm - 1
End If
If tm = 2 And td > 28 Then
    td = 28
End If
If tm = 4 Or tm = 6 Or tm = 9 Or tm = 11 Then
    If td = 31 Then
        td = 30
    End If
End If
tyy$ = Mid$(Str$(ty), 2)
tmm$ = Mid$(Str$(tm), 2)
tdd$ = Mid$(Str$(td), 2)
REPORT.SelectionFormula = "({warrantinfo.logdate} > DATE(" + tyy$ + "," + tmm$ + "," + tdd$ + ") and {warrantinfo.nodisp} = 0 and (IsNull({warrantinfo.sentto}) OR (NOT ISNULL({warrantinfo.sentto}) AND NOT ISNULL({warrantinfo.RETURNEDON}))) and IsNull({warrantinfo.whenarrested}) and {warrantinfo.recalled}  <> 1) or ({warrantinfo.keeplist} = 1 and IsNull({warrantinfo.whenarrested}))"
REPORT.ReportFileName = nww + "pws.rpt"
REPORT.SortFields(0) = ""
REPORT.SortFields(1) = ""
REPORT.SortFields(2) = ""
REPORT.SortFields(3) = ""
On Error GoTo 0
REPORT.Action = 1
Screen.MousePointer = 0
'POFRAME.Visible = False
Exit Sub

End Sub

Private Sub Command11_Click()
Call loadcc
ccframe.Top = 1000
ccframe.Left = 1000
ccframe.Visible = True
cc.SetFocus
End Sub

Private Sub Command12_Click()
If cc.ListIndex = -1 Or changeto = "" Then
    Exit Sub
End If
msg = MsgBox("All warrants with a charge of " + cc.List(cc.ListIndex) + " will be changed to have a charge of " + changeto + ".  Are you sure?", 4, "Genesis Information Log")
If msg <> 6 Then
    Exit Sub
End If
Dim db As Database, ds As Recordset
Set db = OpenDatabase(nww + "warrant.mdb")
Set ds = db.OpenRecordset("select charge from warrantinfo where charge = " + Chr$(34) + cc.List(cc.ListIndex) + Chr$(34))
If Not ds.EOF Then
    ds.MoveFirst
    While Not ds.EOF
        ds.Edit
        ds("charge") = changeto
        ds.Update
        ds.MoveNext
    Wend
End If
db.Close
Call loadcc
changeto = ""
End Sub

Private Sub Command13_Click()
HC = charge
Call chargeload
charge = HC
ccframe.Visible = False
End Sub

Private Sub Command14_Click()
Screen.MousePointer = 11
If Val(numcopies) = 0 Then
    CopiesToPrinter = 1
Else
    CopiesToPrinter = Val(numcopies)
End If
REPORT.Destination = 0
tyy$ = Right$(Date$, 4)
ty = Val(tyy$)
tmm$ = Left$(Date$, 2)
tm = Val(tmm$)
tdd$ = Mid$(Date$, 4, 2)
td = Val(tdd$)
If tm - 1 < 1 Then
    ty = ty - 1
    tm = 12 + tm - 1
Else
    tm = tm - 1
End If
If tm = 2 And td > 28 Then
    td = 28
End If
If tm = 4 Or tm = 6 Or tm = 9 Or tm = 11 Then
    If td = 31 Then
        td = 30
    End If
End If
tyy$ = Mid$(Str$(ty), 2)
tmm$ = Mid$(Str$(tm), 2)
tdd$ = Mid$(Str$(td), 2)
REPORT.SelectionFormula = "({warrantinfo.logdate} > DATE(" + tyy$ + "," + tmm$ + "," + tdd$ + ") and {warrantinfo.nodisp} = 0 and (IsNull({warrantinfo.sentto}) OR (NOT ISNULL({warrantinfo.sentto}) AND NOT ISNULL({warrantinfo.RETURNEDON}))) and IsNull({warrantinfo.whenarrested}) and {warrantinfo.recalled}  <> 1) or ({warrantinfo.keeplist} = 1 and IsNull({warrantinfo.whenarrested}))"
REPORT.ReportFileName = nww + "pwsa.rpt"
REPORT.SortFields(0) = ""
REPORT.SortFields(1) = ""
REPORT.SortFields(2) = ""
REPORT.SortFields(3) = ""
On Error GoTo 0
REPORT.Action = 1
Screen.MousePointer = 0
'POFRAME.Visible = False
Exit Sub

End Sub

Private Sub Command15_Click()
Screen.MousePointer = 11
If Val(numcopies) = 0 Then
    CopiesToPrinter = 1
Else
    CopiesToPrinter = Val(numcopies)
End If
REPORT.Destination = 0
tyy$ = Right$(Date$, 4)
ty = Val(tyy$)
tmm$ = Left$(Date$, 2)
tm = Val(tmm$)
tdd$ = Mid$(Date$, 4, 2)
td = Val(tdd$)
If tm - 1 < 1 Then
    ty = ty - 1
    tm = 12 + tm - 1
Else
    tm = tm - 1
End If
If tm = 2 And td > 28 Then
    td = 28
End If
If tm = 4 Or tm = 6 Or tm = 9 Or tm = 11 Then
    If td = 31 Then
        td = 30
    End If
End If
tyy$ = Mid$(Str$(ty), 2)
tmm$ = Mid$(Str$(tm), 2)
tdd$ = Mid$(Str$(td), 2)
addr1 = InputBox("Enter the beginning portion of the first address field to be printed.", "Genesis Information Log", "")
If addr1 = "" Then
    Screen.MousePointer = 0
    Exit Sub
End If
addr2 = InputBox("Enter the beginning portion of the city/state field to be printed.", "Genesis Information Log", "")
If addr2 = "" Then
    Screen.MousePointer = 0
    Exit Sub
End If
REPORT.SelectionFormula = "({warrantinfo.address} [1 to " + Mid$(Str$(Len(addr1)), 2) + "] = '" + addr1 + "' and {warrantinfo.address2} [1 to " + Mid$(Str$(Len(addr2)), 2) + "] = '" + addr2 + "') and (({warrantinfo.logdate} > DATE(" + tyy$ + "," + tmm$ + "," + tdd$ + ") and {warrantinfo.nodisp} = 0 and (IsNull({warrantinfo.sentto}) OR (NOT ISNULL({warrantinfo.sentto}) AND NOT ISNULL({warrantinfo.RETURNEDON}))) and IsNull({warrantinfo.whenarrested}) and {warrantinfo.recalled}  <> 1) or ({warrantinfo.keeplist} = 1 and IsNull({warrantinfo.whenarrested})))"
REPORT.ReportFileName = nww + "pws.rpt"
REPORT.SortFields(0) = ""
REPORT.SortFields(1) = ""
REPORT.SortFields(2) = ""
REPORT.SortFields(3) = ""
On Error GoTo 0
REPORT.Action = 1
Screen.MousePointer = 0
'POFRAME.Visible = False
Exit Sub

End Sub

Private Sub Command16_Click()
HI = issuedby
Call issuedbyload
issuedby = HI
icframe.Visible = False

End Sub

Private Sub Command17_Click()
If ic.ListIndex = -1 Or changeissued = "" Then
    Exit Sub
End If
msg = MsgBox("All warrants Issued By " + ic.List(ic.ListIndex) + " will be changed to " + changeissued + ".  Are you sure?", 4, "Genesis Information Log")
If msg <> 6 Then
    Exit Sub
End If
Dim db As Database, ds As Recordset
Set db = OpenDatabase(nww + "warrant.mdb")
Set ds = db.OpenRecordset("select issuedby from warrantinfo where issuedby = " + Chr$(34) + ic.List(ic.ListIndex) + Chr$(34))
If Not ds.EOF Then
    ds.MoveFirst
    While Not ds.EOF
        ds.Edit
        ds("issuedby") = changeissued
        ds.Update
        ds.MoveNext
    Wend
End If
db.Close
Call loadic
changeissued = ""

End Sub

Private Sub Command18_Click()
Call loadic
icframe.Left = 255
icframe.Top = 345
icframe.Visible = True
ic.SetFocus

End Sub

Private Sub Command19_Click()
Screen.MousePointer = 11
If Val(numcopies) = 0 Then
    CopiesToPrinter = 1
Else
    CopiesToPrinter = Val(numcopies)
End If
REPORT.Destination = 0
tyy$ = Right$(Date$, 4)
ty = Val(tyy$)
tmm$ = Left$(Date$, 2)
tm = Val(tmm$)
tdd$ = Mid$(Date$, 4, 2)
td = Val(tdd$)
If tm - 1 < 1 Then
    ty = ty - 1
    tm = 12 + tm - 1
Else
    tm = tm - 1
End If
If tm = 2 And td > 28 Then
    td = 28
End If
If tm = 4 Or tm = 6 Or tm = 9 Or tm = 11 Then
    If td = 31 Then
        td = 30
    End If
End If
tyy$ = Mid$(Str$(ty), 2)
tmm$ = Mid$(Str$(tm), 2)
tdd$ = Mid$(Str$(td), 2)
REPORT.SelectionFormula = "({warrantinfo.ccounty} = 1 and {warrantinfo.logdate} > DATE(" + tyy$ + "," + tmm$ + "," + tdd$ + ") and {warrantinfo.nodisp} = 0 and (IsNull({warrantinfo.sentto}) OR (NOT ISNULL({warrantinfo.sentto}) AND NOT ISNULL({warrantinfo.RETURNEDON}))) and IsNull({warrantinfo.whenarrested}) and {warrantinfo.recalled}  <> 1) or ({warrantinfo.ccounty} = 1 and {warrantinfo.keeplist} = 1 and IsNull({warrantinfo.whenarrested}))"
REPORT.ReportFileName = nww + "pwsc.rpt"
REPORT.SortFields(0) = ""
REPORT.SortFields(1) = ""
REPORT.SortFields(2) = ""
REPORT.SortFields(3) = ""
On Error GoTo 0
REPORT.Action = 1
Screen.MousePointer = 0
'POFRAME.Visible = False
Exit Sub

End Sub

Private Sub Command2_Click()
StartDate = InputBox("Enter starting date for report.", "Genesis Information Log", "")
If Not IsDate(StartDate) Then
    msg = MsgBox("Invalid start date.", 48, "Genesis Error Log")
    Exit Sub
End If
EndDate = InputBox("Enter ending date for report.", "Genesis Information Log", "")
If Not IsDate(EndDate) Then
    msg = MsgBox("Invalid end date.", 48, "Genesis Error Log")
    Exit Sub
End If
Dim db As Database, ds As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nww + "warrant.mdb")
Set ds = db.OpenRecordset("select * from passthru")
If ds.EOF Then
    ds.AddNew
Else
    ds.Edit
End If
ds("todate") = EndDate
ds("fromdate") = StartDate
ds.Update
db.Close
On Error Resume Next
Screen.MousePointer = 11
If Val(numcopies) = 0 Then
    CopiesToPrinter = 1
Else
    CopiesToPrinter = Val(numcopies)
End If
REPORT.Destination = 0
REPORT.ReportFileName = nww + "fourd.RPT"
If oalpha Then
    If oarea1 Then
        REPORT.SortFields(1) = "+{warrantinfo.wNAME}"
        REPORT.SortFields(0) = "+{warrantinfo.area1}"
    Else
    If oarea2 Then
        REPORT.SortFields(1) = "+{warrantinfo.wNAME}"
        REPORT.SortFields(0) = "+{warrantinfo.area2}"
    Else
        REPORT.SortFields(1) = "+{warrantinfo.wNAME}"
        REPORT.SortFields(0) = "+{warrantinfo.wNAME}"
    End If
    End If
Else
    If oarea1 Then
        REPORT.SortFields(1) = "+{warrantinfo.warrant}"
        REPORT.SortFields(0) = "+{warrantinfo.area1}"
    Else
    If oarea2 Then
        REPORT.SortFields(1) = "+{warrantinfo.warrant}"
        REPORT.SortFields(0) = "+{warrantinfo.area2}"
    Else
        REPORT.SortFields(0) = "+{warrantinfo.warrant}"
        REPORT.SortFields(1) = ""
    End If
    End If
End If
StartDate = Format$(StartDate, "mm/dd/yyyy")
starty = Right$(StartDate, 4)
startm = Left$(StartDate, 2)
startd = Mid$(StartDate, 4, 2)
EndDate = Format$(EndDate, "mm/dd/yyyy")
endy = Right$(EndDate, 4)
endm = Left$(EndDate, 2)
endd = Mid$(EndDate, 4, 2)
REPORT.SelectionFormula = "not isnull({warrantinfo.whenarrested}) and {warrantinfo.type} = '4D Warrant' and {WARRANTINFO.whenarrested} >= date(" + starty + "," + startm + "," + startd + ") and {WARRANTINFO.whenarrested} <= date(" + endy + "," + endm + "," + endd + ")"
REPORT.Action = 1
Screen.MousePointer = 0
'POFRAME.Visible = False

Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If

End Sub

Private Sub Command20_Click()
Screen.MousePointer = 11
If Val(numcopies) = 0 Then
    CopiesToPrinter = 1
Else
    CopiesToPrinter = Val(numcopies)
End If
REPORT.Destination = 0
tyy$ = Right$(Date$, 4)
ty = Val(tyy$)
tmm$ = Left$(Date$, 2)
tm = Val(tmm$)
tdd$ = Mid$(Date$, 4, 2)
td = Val(tdd$)
If tm - 1 < 1 Then
    ty = ty - 1
    tm = 12 + tm - 1
Else
    tm = tm - 1
End If
If tm = 2 And td > 28 Then
    td = 28
End If
If tm = 4 Or tm = 6 Or tm = 9 Or tm = 11 Then
    If td = 31 Then
        td = 30
    End If
End If
tyy$ = Mid$(Str$(ty), 2)
tmm$ = Mid$(Str$(tm), 2)
tdd$ = Mid$(Str$(td), 2)
addr1 = InputBox("Enter the Area Description 1 value.", "Genesis Information Log", "")
If addr1 = "" Then
    Screen.MousePointer = 0
    Exit Sub
End If
addr2 = InputBox("Enter the Area Description 2 value.", "Genesis Information Log", "")
If addr2 = "" Then
    Screen.MousePointer = 0
    Exit Sub
End If
REPORT.SelectionFormula = "({warrantinfo.area1} = '" + addr1 + "' and {warrantinfo.area2} = '" + addr2 + "') and (({warrantinfo.logdate} > DATE(" + tyy$ + "," + tmm$ + "," + tdd$ + ") and {warrantinfo.nodisp} = 0 and (IsNull({warrantinfo.sentto}) OR (NOT ISNULL({warrantinfo.sentto}) AND NOT ISNULL({warrantinfo.RETURNEDON}))) and IsNull({warrantinfo.whenarrested}) and {warrantinfo.recalled}  <> 1) or ({warrantinfo.keeplist} = 1 and IsNull({warrantinfo.whenarrested})))"
REPORT.ReportFileName = nww + "pwsbarea.rpt"
REPORT.SortFields(0) = ""
REPORT.SortFields(1) = ""
REPORT.SortFields(2) = ""
REPORT.SortFields(3) = ""
On Error GoTo 0
REPORT.Action = 1
Screen.MousePointer = 0
'POFRAME.Visible = False
Exit Sub

End Sub

Private Sub Command21_Click()
Screen.MousePointer = 11
If Val(numcopies) = 0 Then
    CopiesToPrinter = 1
Else
    CopiesToPrinter = Val(numcopies)
End If
REPORT.Destination = 0
tyy$ = Right$(Date$, 4)
ty = Val(tyy$)
tmm$ = Left$(Date$, 2)
tm = Val(tmm$)
tdd$ = Mid$(Date$, 4, 2)
td = Val(tdd$)
If tm - 1 < 1 Then
    ty = ty - 1
    tm = 12 + tm - 1
Else
    tm = tm - 1
End If
If tm = 2 And td > 28 Then
    td = 28
End If
If tm = 4 Or tm = 6 Or tm = 9 Or tm = 11 Then
    If td = 31 Then
        td = 30
    End If
End If
tyy$ = Mid$(Str$(ty), 2)
tmm$ = Mid$(Str$(tm), 2)
tdd$ = Mid$(Str$(td), 2)
REPORT.SelectionFormula = "{warrantinfo.ccounty} = 0 and (({warrantinfo.logdate} > DATE(" + tyy$ + "," + tmm$ + "," + tdd$ + ") and {warrantinfo.nodisp} = 0 and (IsNull({warrantinfo.sentto}) OR (NOT ISNULL({warrantinfo.sentto}) AND NOT ISNULL({warrantinfo.RETURNEDON}))) and IsNull({warrantinfo.whenarrested}) and {warrantinfo.recalled}  <> 1) or ({warrantinfo.keeplist} = 1 and IsNull({warrantinfo.whenarrested})))"
REPORT.ReportFileName = nww + "pwsnc.rpt"
REPORT.SortFields(0) = ""
REPORT.SortFields(1) = ""
REPORT.SortFields(2) = ""
REPORT.SortFields(3) = ""
On Error GoTo 0
REPORT.Action = 1
Screen.MousePointer = 0
'POFRAME.Visible = False
Exit Sub


End Sub

Private Sub Command22_Click()
REPORT.ReportFileName = nww + "iawarrant.rpt"
REPORT.SelectionFormula = "{warrantinfo.warrant} = '" + warrant + "'"
If defendant = "" Then
    defendant = wname
End If
issueframe.Left = 2000
issueframe.Top = 2500
issueframe.Visible = True
court.SetFocus

End Sub

Private Sub Command23_Click()
If UCase(frmLogin.txtUserName) = "DEMO" And UCase(frmLogin.txtPassword) = "DEMO" Then
    msg = MsgBox("Not available in DEMO version.", 48, "Genesis Information Log")
    Screen.MousePointer = 0
    Exit Sub
End If
If sedit = 0 And ssupervisor = 0 Then
    msg = MsgBox("You have insufficient authority to save.", 48, "Genesis Error Log")
    Exit Sub
End If
If warrant = "" Then
    msg = MsgBox("Warrant# must be entered.", 48, "Genesis Error Log")
    warrant.SetFocus
    Exit Sub
End If
If Len(warrant) > 20 Then
    msg = MsgBox("Warrant# must be 20 characters or less.", 48, "Genesis Error Log")
    warrant.SetFocus
    Exit Sub
End If

If Not IsDate(logdate) Then
    msg = MsgBox("Log Date must be a valid date.", 48, "Genesis Error Log")
    logdate.SetFocus
    Exit Sub
End If
If courtdate > "" And Not IsDate(courtdate) Then
    msg = MsgBox("Court Date must be a valid date.", 48, "Genesis Error Log")
    courtdate.SetFocus
    Exit Sub
End If
If whenarrested > "" And Not IsDate(whenarrested) Then
    msg = MsgBox("When Arrested Date must be a valid date.", 48, "Genesis Error Log")
    whenarrested.SetFocus
    Exit Sub
End If
If birthdate > "" And Not IsDate(birthdate) Then
    msg = MsgBox("Birthdate must be a valid date.", 48, "Genesis Error Log")
    birthdate.SetFocus
    Exit Sub
End If
If senton > "" And Not IsDate(senton) Then
    msg = MsgBox("WARRANT SENT ON must be a valid date.", 48, "Genesis Error Log")
    senton.SetFocus
    Exit Sub
End If
If returnedon > "" And Not IsDate(returnedon) Then
    msg = MsgBox("WARRANT RETURNED ON must be a valid date.", 48, "Genesis Error Log")
    returnedon.SetFocus
    Exit Sub
End If
logdate = Format$(logdate, "mm/dd/yyyy")
If courtdate > "" Then
    courtdate = Format$(courtdate, "mm/dd/yyyy")
End If
If whenarrested > "" Then
    If whenarrested > "" Then
        whenarrested = Format$(whenarrested, "mm/dd/yyyy")
    End If
End If
If birthdate > "" Then
    If birthdate > "" Then
        birthdate = Format$(birthdate, "mm/dd/yyyy")
    End If
End If
If senton > "" Then
    If senton > "" Then
        senton = Format$(senton, "mm/dd/yyyy")
    End If
End If
If returnedon > "" Then
    If returnedon > "" Then
        returnedon = Format$(returnedon, "mm/dd/yyyy")
    End If
End If
If wname = "" Then
    msg = MsgBox("Name must be entered.", 48, "Genesis Error Log")
    wname.SetFocus
    Exit Sub
End If
If charge = "" Then
    msg = MsgBox("Charge must be entered.", 48, "Genesis Error Log")
    charge.SetFocus
    Exit Sub
End If
If issuedby = "" Then
    msg = MsgBox("Issued By must be entered.", 48, "Genesis Error Log")
    issuedby.SetFocus
    Exit Sub
End If
Screen.MousePointer = 11
Call savertn
If mugshot.Picture > 0 Then
    SavePicture mugshot.Picture, "c:\mug.jpg"
Else
    FileCopy nwl + "blank.jpg", "c:\mug.jpg"
End If
REPORT.SelectionFormula = "{WARRANTINFO.WARRANT} = '" + warrant + "'"
REPORT.Destination = 1
REPORT.Action = 1

If ReportFileName = nww + "IBWARRANT.RPT" Then
    MsgBox "Flip paper to print other side.", vbOKOnly, "Genesis Information Log"
    ReportFileName = nww + "IBWARRANT2.RPT"
    REPORT.Action = 1
End If

Screen.MousePointer = 0
End Sub

Private Sub Command24_Click()
issueframe.Visible = False

End Sub

Private Sub Command3_Click()
Screen.MousePointer = 11
If Val(numcopies) = 0 Then
    CopiesToPrinter = 1
Else
    CopiesToPrinter = Val(numcopies)
End If
REPORT.Destination = 0
REPORT.ReportFileName = nww + "sent.RPT"
REPORT.SelectionFormula = ""
If oalpha Then
    If oarea1 Then
        REPORT.SortFields(1) = "+{warrantinfo.wNAME}"
        REPORT.SortFields(0) = "+{warrantinfo.area1}"
    Else
    If oarea2 Then
        REPORT.SortFields(1) = "+{warrantinfo.wNAME}"
        REPORT.SortFields(0) = "+{warrantinfo.area2}"
    Else
        REPORT.SortFields(0) = "+{warrantinfo.wNAME}"
        REPORT.SortFields(1) = ""
    End If
    End If
Else
    If oarea1 Then
        REPORT.SortFields(1) = "+{warrantinfo.warrant}"
        REPORT.SortFields(0) = "+{warrantinfo.area1}"
    Else
    If oarea2 Then
        REPORT.SortFields(1) = "+{warrantinfo.warrant}"
        REPORT.SortFields(0) = "+{warrantinfo.area2}"
    Else
        REPORT.SortFields(0) = "+{warrantinfo.warrant}"
        REPORT.SortFields(1) = ""
    End If
    End If
End If
REPORT.Action = 1
Screen.MousePointer = 0
'POFRAME.Visible = False
End Sub

Private Sub Command4_Click()
'chris
inp = InputBox("Please Select Report Type                                             Enter 'S' for Standard or Enter 'C' for Custom", "Genesis Information Log", "C")
inp = UCase(inp)
If inp <> "S" And inp <> "C" Then
    msg = MsgBox("          Invalid Entry", vbOKOnly, "Genesis Error Log")
    Exit Sub
End If
Screen.MousePointer = 11
If Val(numcopies) = 0 Then
    CopiesToPrinter = 1
Else
    CopiesToPrinter = Val(numcopies)
End If
REPORT.Destination = 0
tyy$ = Right$(Date$, 4)
ty = Val(tyy$)
tmm$ = Left$(Date$, 2)
tm = Val(tmm$)
tdd$ = Mid$(Date$, 4, 2)
td = Val(tdd$)
If tm - 1 < 1 Then
    ty = ty - 1
    tm = 12 + tm - 1
Else
    tm = tm - 1
End If
If tm = 2 And td > 28 Then
    td = 28
End If
If tm = 4 Or tm = 6 Or tm = 9 Or tm = 11 Then
    If td = 31 Then
        td = 30
    End If
End If
tyy$ = Mid$(Str$(ty), 2)
tmm$ = Mid$(Str$(tm), 2)
tdd$ = Mid$(Str$(td), 2)
REPORT.SelectionFormula = "({warrantinfo.logdate} > DATE(" + tyy$ + "," + tmm$ + "," + tdd$ + ") and {warrantinfo.nodisp} = 0 and (IsNull({warrantinfo.sentto}) OR (NOT ISNULL({warrantinfo.sentto}) AND NOT ISNULL({warrantinfo.RETURNEDON}))) and IsNull({warrantinfo.whenarrested}) and {warrantinfo.recalled}  <> 1) or ({warrantinfo.keeplist} = 1 and IsNull({warrantinfo.whenarrested}))"
If inp = "S" Then
    REPORT.ReportFileName = nww + "unservcu.RPT"
Else
If inp = "C" Then
    REPORT.ReportFileName = nww + "unservc2.rpt"
End If
End If
If oalpha Then
    If oarea1 Then
        REPORT.SortFields(1) = "+{warrantinfo.wNAME}"
        REPORT.SortFields(0) = "+{warrantinfo.area1}"
    Else
    If oarea2 Then
        REPORT.SortFields(1) = "+{warrantinfo.wNAME}"
        REPORT.SortFields(0) = "+{warrantinfo.area2}"
    Else
        REPORT.SortFields(0) = "+{warrantinfo.wNAME}"
        REPORT.SortFields(1) = ""
    End If
    End If
Else
    If oarea1 Then
        REPORT.SortFields(1) = "+{warrantinfo.warrant}"
        REPORT.SortFields(0) = "+{warrantinfo.area1}"
    Else
    If oarea2 Then
        REPORT.SortFields(1) = "+{warrantinfo.warrant}"
        REPORT.SortFields(0) = "+{warrantinfo.area2}"
    Else
        REPORT.SortFields(0) = "+{warrantinfo.warrant}"
        REPORT.SortFields(1) = ""
    End If
    End If
End If
On Error GoTo 0
REPORT.Action = 1
Screen.MousePointer = 0
'POFRAME.Visible = False
Exit Sub
End Sub

Private Sub Command5_Click()
POFRAME.Visible = False
End Sub

Private Sub Command6_Click()
REPORT.ReportFileName = nww + "ibwarrant.rpt"
REPORT.SelectionFormula = "{warrantinfo.warrant} = '" + warrant + "'"
If defendant = "" Then
    defendant = wname
End If
issueframe.Left = 2000
issueframe.Top = 2500
issueframe.Visible = True
court.SetFocus
End Sub

Private Sub Command7_Click()
'chris
inp = InputBox("Please Select Report Type                                             Enter 'S' for Standard or Enter 'C' for Custom", "Genesis Information Log", "C")
inp = UCase(inp)
If inp <> "S" And inp <> "C" Then
    msg = MsgBox("          Invalid Entry", vbOKOnly, "Genesis Error Log")
    Exit Sub
    End If
Screen.MousePointer = 11
If Val(numcopies) = 0 Then
    CopiesToPrinter = 1
Else
    CopiesToPrinter = Val(numcopies)
End If
REPORT.Destination = 0
REPORT.SelectionFormula = "{warrantinfo.recalled} <> 1 and {warrantinfo.nodisp} = 0 and (IsNull({warrantinfo.sentto}) OR (NOT ISNULL({warrantinfo.sentto}) AND NOT ISNULL({warrantinfo.RETURNEDON}))) AND ISNULL({warrantinfo.whenarrested}) and ({warrantinfo.type} = 'Bench - Magistrate' or {warrantinfo.type} = 'Bench - Family Court' or {warrantinfo.type} = 'Bench - General Sessions' or {warrantinfo.type} = '4D Warrant')"
If inp = "S" Then
    REPORT.ReportFileName = nww + "bench.RPT"
Else
If inp = "C" Then
    REPORT.ReportFileName = nww + "bench2.RPT"
End If
End If
REPORT.SortFields(0) = ""
REPORT.SortFields(1) = ""
REPORT.SortFields(2) = ""
REPORT.SortFields(3) = ""
On Error GoTo 0
REPORT.Action = 1
Screen.MousePointer = 0
'POFRAME.Visible = False
Exit Sub
End Sub

Private Sub Command8_Click()
If lookupframe.Visible = True Then
    lookupframe.Visible = False
    Command8.Caption = "Searc&h"
    Exit Sub
End If
inp = InputBox("Enter data to search for in IDENTIFYING DATA and REMARKS.", "Genesis Information Log", "")
If inp = "" Then
    Exit Sub
End If
Command8.Caption = "Close Searc&h"
Screen.MousePointer = 11
Dim db As Database, ds As Recordset
lookuplist.ListItems.clear
On Error GoTo oderror
od:
Set db = OpenDatabase(nww + "warrant.mdb")
Set ds = db.OpenRecordset("select warrant,logdate,wname from warrantinfo where remarks like '*" + inp + "*' or iddata like '*" + inp + "*' order by wname,logdate,warrant asc")
If Not ds.EOF Then
    ds.MoveFirst
Else
    msg = MsgBox("No items available for Search.", 48, "Genesis Error Log")
    Screen.MousePointer = 0
    On Error Resume Next
    Command8.Caption = "Searc&h"
    Exit Sub
End If
While Not ds.EOF
    Set itmx = lookuplist.ListItems.add(, , ds("warrant"))
    itmx.SubItems(1) = ds("logdate")
    itmx.SubItems(2) = ds("wname")
    ds.MoveNext
Wend
lookupframe.Left = 120
lookupframe.Top = 4000
lookupframe.Visible = True
Screen.MousePointer = 0
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

Private Sub Command9_Click()
Screen.MousePointer = 11
If Val(numcopies) = 0 Then
    CopiesToPrinter = 1
Else
    CopiesToPrinter = Val(numcopies)
End If
REPORT.Destination = 0
REPORT.ReportFileName = nww + "returned.RPT"
REPORT.SelectionFormula = ""
If oalpha Then
    If oarea1 Then
        REPORT.SortFields(1) = "+{warrantinfo.wNAME}"
        REPORT.SortFields(0) = "+{warrantinfo.area1}"
    Else
    If oarea2 Then
        REPORT.SortFields(1) = "+{warrantinfo.wNAME}"
        REPORT.SortFields(0) = "+{warrantinfo.area2}"
    Else
        REPORT.SortFields(0) = "+{warrantinfo.wNAME}"
        REPORT.SortFields(1) = ""
    End If
    End If
Else
    If oarea1 Then
        REPORT.SortFields(1) = "+{warrantinfo.warrant}"
        REPORT.SortFields(0) = "+{warrantinfo.area1}"
    Else
    If oarea2 Then
        REPORT.SortFields(1) = "+{warrantinfo.warrant}"
        REPORT.SortFields(0) = "+{warrantinfo.area2}"
    Else
        REPORT.SortFields(0) = "+{warrantinfo.warrant}"
        REPORT.SortFields(1) = ""
    End If
    End If
End If
REPORT.Action = 1
Screen.MousePointer = 0
'POFRAME.Visible = False

End Sub

Private Sub courtdate_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
If Len(courtdate) = 1 Or Len(courtdate) = 4 Then
    SendKeys "/"
End If
End If

End Sub

Private Sub courtdate_LostFocus()
If IsDate(courtdate) Then
    courtdate = Format$(courtdate, "mm/dd/yyyy")
End If

End Sub

Private Sub Data1_Validate(Action As Integer, Save As Integer)

End Sub

Private Sub deletebutton_Click()
If UCase(frmLogin.txtUserName) = "DEMO" And UCase(frmLogin.txtPassword) = "DEMO" Then
    msg = MsgBox("Not available in DEMO version.", 48, "Genesis Information Log")
    Screen.MousePointer = 0
    Exit Sub
End If
If sdelete = 0 And ssupervisor = 0 Then
    msg = MsgBox("You have insufficient authority to delete.", 48, "Genesis Error Log")
    Exit Sub
End If
If warrant = "" Then
    msg = MsgBox("Warrant# must be entered.", 48, "Genesis Error Log")
    warrant.SetFocus
End If
msg = MsgBox("Are you sure you wish to delete this record?", 4, "Genesis Information Log")
If msg = 7 Then
    Exit Sub
End If
Screen.MousePointer = 11
Dim db As Database, ds As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nww + "warrant.mdb")
Set ds = db.OpenRecordset("select * from warrantinfo where warrant = '" + warrant + "'")
If Not ds.EOF Then
    ds.MoveFirst
    ds.Delete
End If
db.Close
On Error Resume Next
Call wnameload
Call wload
Call chargeload
Call issuedbyload
Call assignedtoload
Call nullfields
justfound = 0
warrant.SetFocus
Screen.MousePointer = 0
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If
End Sub

Private Sub exitbutton_Click()
Unload wmain
End Sub

Private Sub findrec()
Dim db As Database, ds, rs As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nww + "warrant.mdb")
Set ds = db.OpenRecordset("select * from warrantinfo where warrant = '" + warrant + "'")
On Error Resume Next
If ds.EOF Then
    Exit Sub
End If
ds.MoveFirst
warrant = ds("warrant")
logdate = ds("logdate")
danger.Value = ds("danger")
recalled.Value = ds("recalled")
keeplist.Value = ds("keeplist")
NODISP.Value = ds("nodisp")
If Not IsNull(ds("recalldate")) Then
    recalldate = ds("recalldate")
Else
    recalldate = ""
End If
If Not IsNull(ds("docket")) Then
    docket = ds("docket")
Else
    docket = ""
End If
If Not IsNull(ds("dss")) Then
    dss = ds("dss")
Else
    dss = ""
End If
If Not IsNull(ds("guardian")) Then
    guardian = ds("guardian")
Else
    guardian = ""
End If
If Not IsNull(ds("courtdate")) Then
    courtdate = ds("courtdate")
Else
    courtdate = ""
End If
If Not IsNull(ds("remarks")) Then
    remarks = ds("remarks")
Else
    remarks = ""
End If
If Not IsNull(ds("iddata")) Then
    iddata = ds("iddata")
Else
    iddata = ""
End If
If Not IsNull(ds("origination")) Then
    origination = ds("origination")
Else
    origination = ""
End If
If Not IsNull(ds("address")) Then
    address(0) = ds("address")
Else
    address(0) = ""
End If
If Not IsNull(ds("address2")) Then
    address(1) = ds("address2")
Else
    address(1) = ""
End If
If Not IsNull(ds("state")) Then
    address(2) = ds("state")
Else
    address(2) = ""
End If
If Not IsNull(ds("zipcode")) Then
    address(3) = ds("zipcode")
Else
    address(3) = ""
End If
If Not IsNull(ds("whenarrested")) Then
    whenarrested = ds("whenarrested")
Else
    whenarrested = ""
End If
If Not IsNull(ds("sentto")) Then
     sentto = ds("sentto")
Else
     sentto = ""
End If
If Not IsNull(ds("senton")) Then
     senton = ds("senton")
Else
     senton = ""
End If
If Not IsNull(ds("returnedon")) Then
     returnedon = ds("returnedon")
Else
     returnedon = ""
End If
If Not IsNull(ds("whenarrested")) Then
    whenarrested = ds("whenarrested")
Else
    whenarrested = ""
End If
If Not IsNull(ds("witness")) Then
    witness = ds("witness")
Else
    witness = ""
End If
If Not IsNull(ds("area1")) Then
    AREA1 = ds("area1")
Else
    AREA1 = ""
End If
If Not IsNull(ds("area2")) Then
    AREA2 = ds("area2")
Else
    AREA2 = ""
End If
If Not IsNull(ds("ssn")) Then
    ssn = ds("ssn")
Else
    ssn = ""
End If
If Not IsNull(ds("idnumber")) Then
    idnumber = ds("idnumber")
Else
    idnumber = ""
End If
ccounty = ds("ccounty")
If Not IsNull(ds("height")) Then
    ht = ds("height")
Else
    ht = ""
End If
If Not IsNull(ds("weight")) Then
    weight = ds("weight")
Else
    weight = ""
End If
If Not IsNull(ds("hair")) Then
    hair = ds("hair")
Else
    hair = ""
End If
If Not IsNull(ds("eyes")) Then
    eyes = ds("eyes")
Else
    eyes = ""
End If

If Not IsNull(ds("casenumber")) Then
    casenumber = ds("casenumber")
Else
    casenumber = ""
End If
If Not IsNull(ds("birthdate")) Then
    birthdate = ds("birthdate")
Else
    birthdate = ""
End If
wname = ds("wname")
charge = ds("charge")
If Not IsNull(ds("officer")) Then
    officer = ds("officer")
Else
    officer = ""
End If
If Not IsNull(ds("assignedto")) Then
    assignedto = ds("assignedto")
Else
    assignedto = ""
End If
issuedby = ds("issuedby")
If ds("race") = "Caucasian" Or ds("race") = "White" Then
    caucasian.Value = True
Else
If ds("race") = "African-American" Or ds("race") = "Black" Then
    africanamerican.Value = True
Else
If ds("race") = "Oriental" Or ds("race") = "Asian/Pacific Islander" Then
    Oriental.Value = True
Else
If ds("race") = "Indian - American Indian/Alaskan Native" Then
    indian.Value = True
Else
If InStr(UCase(ds("RACE")), "HISPANIC") > 0 Then
    caucasian.Value = True
Else
    other.Value = True
End If
End If
End If
End If
End If
If ds("type") = "Bench - Family Court" Then
    benchfc.Value = True
Else
If ds("type") = "Bench - Magistrate" Then
    benchm.Value = True
Else
If ds("type") = "Bench - General Sessions" Then
    benchgs.Value = True
Else
If ds("type") = "Warrant" Then
    regw.Value = True
Else
    w4d.Value = True
End If
End If
End If
End If
If ds("sex") = "Male" Then
    male.Value = True
Else
    female.Value = True
End If
recentry = ds("entry")
If Not IsNull(ds("court")) Then
    court = ds("court")
End If
If Not IsNull(ds("ind")) Then
    ind = ds("ind")
End If
If Not IsNull(ds("judge")) Then
    judge = ds("judge")
End If
If Not IsNull(ds("plaintiff")) Then
    plaintiff = ds("plaintiff")
End If
If Not IsNull(ds("mdesc")) Then
    mdesc = ds("mdesc")
End If
If Not IsNull(ds("offensedate")) Then
    offensedate = ds("offensedate")
End If
If Not IsNull(ds("defendant")) Then
    defendant = ds("defendant")
End If
If Not IsNull(ds("county")) Then
    county = ds("county")
End If
If Not IsNull(ds("municipality")) Then
    Municipality = ds("municipality")
End If
If Not IsNull(ds("offense")) Then
    offense.Text = ds("offense")
End If
If Not IsNull(ds("facts")) Then
    facts.Text = ds("facts")
End If

ssql = ""
If idnumber > "" Then
    ssql = ssql + " and idnumber = '" + idnumber + "'"
End If
If IsDate(birthdate) Then
    ssql = ssql + " and birthdate = #" + birthdate + "#"
End If
If ssn > "" Then
    ssql = ssql + " and ssn = '" + ssn + "'"
End If
Set db = OpenDatabase(nwl + "LAWSUITE.MDB")
Set rs = db.OpenRecordset("SELECT * FROM PEOPLE WHERE DPNAMElf = " + Chr$(34) + wname + Chr$(34) + ssql)
If Not rs.EOF Then
    rs.MoveFirst
    If Not IsNull(rs("mugshot")) Then
        mugshot.Picture = LoadPicture(rs("mugshot"))
    Else
        mugshot.Picture = LoadPicture()
    End If
End If
justfound = 1
db.Close
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If
End Sub

Private Sub dow_Click()
Screen.MousePointer = 11
If Val(numcopies) = 0 Then
    CopiesToPrinter = 1
Else
    CopiesToPrinter = Val(numcopies)
End If
REPORT.Destination = 0
tyy$ = Right$(Date$, 4)
ty = Val(tyy$)
tmm$ = Left$(Date$, 2)
tm = Val(tmm$)
tdd$ = Mid$(Date$, 4, 2)
td = Val(tdd$)
If tm - 1 < 1 Then
    ty = ty - 1
    tm = 12 + tm - 1
Else
    tm = tm - 1
End If
If tm = 2 And td > 28 Then
    td = 28
End If
If tm = 4 Or tm = 6 Or tm = 9 Or tm = 11 Then
    If td = 31 Then
        td = 30
    End If
End If
tyy$ = Mid$(Str$(ty), 2)
tmm$ = Mid$(Str$(tm), 2)
tdd$ = Mid$(Str$(td), 2)
REPORT.SelectionFormula = "({warrantinfo.logdate} > DATE(" + tyy$ + "," + tmm$ + "," + tdd$ + ") and {warrantinfo.nodisp} = 0 and (IsNull({warrantinfo.sentto}) OR (NOT ISNULL({warrantinfo.sentto}) AND NOT ISNULL({warrantinfo.RETURNEDON}))) and IsNull ({warrantinfo.whenarrested}) and {warrantinfo.recalled}  <> 1) or ({warrantinfo.keeplist} = 1 and IsNull ({warrantinfo.whenarrested}))"
REPORT.ReportFileName = nww + "unservcd.RPT"
If oalpha Then
    If oarea1 Then
        REPORT.SortFields(3) = "+{warrantinfo.wNAME}"
        REPORT.SortFields(1) = "+{warrantinfo.address}"
        REPORT.SortFields(2) = "+{warrantinfo.address2}"
        REPORT.SortFields(0) = "+{warrantinfo.area1}"
    Else
    If oarea2 Then
        REPORT.SortFields(3) = "+{warrantinfo.wNAME}"
        REPORT.SortFields(1) = "+{warrantinfo.address}"
        REPORT.SortFields(2) = "+{warrantinfo.address2}"
        REPORT.SortFields(0) = "+{warrantinfo.area2}"
    Else
        REPORT.SortFields(0) = "+{warrantinfo.wNAME}"
        REPORT.SortFields(1) = ""
    End If
    End If
Else
    If oarea1 Then
        REPORT.SortFields(3) = "+{warrantinfo.warrant}"
        REPORT.SortFields(1) = "+{warrantinfo.address}"
        REPORT.SortFields(2) = "+{warrantinfo.address2}"
        REPORT.SortFields(0) = "+{warrantinfo.area1}"
    Else
    If oarea2 Then
        REPORT.SortFields(3) = "+{warrantinfo.warrant}"
        REPORT.SortFields(1) = "+{warrantinfo.address}"
        REPORT.SortFields(2) = "+{warrantinfo.address2}"
        REPORT.SortFields(0) = "+{warrantinfo.area2}"
    Else
        REPORT.SortFields(0) = "+{warrantinfo.warrant}"
        REPORT.SortFields(1) = ""
    End If
    End If
End If
On Error GoTo 0
REPORT.Action = 1
Screen.MousePointer = 0
'POFRAME.Visible = False
Exit Sub

End Sub

Private Sub facts_KeyDown(KeyCode As Integer, Shift As Integer)
If (Shift = vbCtrlMask) And (KeyCode = vbKeyF2) Then
        Call SpellCk_Click(2)
    End If
End Sub

Private Sub Form_Load()
nametype = 1
For t% = 0 To Forms.Count - 1
    If Forms(t%).Name = "xref" Then
        FROMXREF = 1
        t% = Forms.Count - 1
    End If
Next t%
oalpha = True
oarea1 = False
oarea2 = False
sedit = frmLogin.wedit
sprint = frmLogin.wprint
sreport = frmLogin.wreport
sbrowse = frmLogin.wbrowse
sdelete = frmLogin.wdelete
ssupervisor = frmLogin.wsupervisor
On Error Resume Next
Call wnameload
Call wload
Call chargeload
Call issuedbyload
Call assignedtoload
Call loadsystem
justfound = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set wmain = Nothing
End Sub

Private Sub issuedby_KeyPress(KeyAscii As Integer)
If Len(issuedby) = 40 Then
    KeyAscii = 0
End If

End Sub

Private Sub issuedbyload()
Dim db As Database, ds As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nww + "warrant.mdb")
Set ds = db.OpenRecordset("select distinct issuedby from warrantinfo")
issuedby.clear
If Not ds.EOF Then
    ds.MoveFirst
End If
While Not ds.EOF
    If ds("issuedby") <> " " Then
        issuedby.AddItem ds("issuedby")
    End If
    ds.MoveNext
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

Private Sub keeplist_Click()
If keeplist = 1 Then
    NODISP = 0
End If
End Sub

Private Sub logdate_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
If Len(logdate) = 1 Or Len(logdate) = 4 Then
    SendKeys "/"
End If
End If
End Sub

Private Sub logdate_LostFocus()
If logdate = "" Then
    logdate = Date$
End If
If IsDate(logdate) Then
    logdate = Format$(logdate, "mm/dd/yyyy")
End If
End Sub

Private Sub lookupbutton_Click()
If lookupframe.Visible = True Then
    lookupframe.Visible = False
    lookupbutton.Caption = "&Lookup"
    Exit Sub
End If
lookupbutton.Caption = "Close &Lookup"
Screen.MousePointer = 11
Dim db As Database, ds As Recordset
lookuplist.ListItems.clear
On Error GoTo oderror
od:
Set db = OpenDatabase(nww + "warrant.mdb")
Set ds = db.OpenRecordset("select warrant,logdate,wname from warrantinfo order by wname,logdate,warrant asc")
On Error Resume Next
If Not ds.EOF Then
    ds.MoveFirst
Else
    msg = MsgBox("No items available for lookup.", 48, "Genesis Error Log")
    Screen.MousePointer = 0
    lookupbutton.Caption = "&Lookup"
    Exit Sub
End If
While Not ds.EOF
    Set itmx = lookuplist.ListItems.add(, , ds("warrant"))
    itmx.SubItems(1) = ds("logdate")
    itmx.SubItems(2) = ds("wname")
    ds.MoveNext
Wend
lookupframe.Left = 120
lookupframe.Top = 4400
lookupframe.Visible = True
Screen.MousePointer = 0
db.Close
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If
End Sub

Private Sub nullfields()
mugshot.Picture = LoadPicture()
court = ""
judge = ""
plaintiff = ""
mdesc = ""
offensedate = ""
defendant = ""
ind = ""
county = False
Municipality = False
facts.Text = ""
offense.Text = ""
warrant = ""
oalpha = True
oarea1 = False
oarea2 = False
logdate = ""
danger.Value = 0
recalled.Value = 0
keeplist.Value = 0
NODISP.Value = 0
recalldate = ""
docket = ""
dss = ""
guardian = ""
courtdate = ""
remarks = ""
iddata = ""
address(0) = ""
address(1) = ""
address(2) = ""
address(3) = ""
origination = ""
whenarrested = ""
sentto = ""
senton = ""
returnedon = ""
witness = ""
casenumber = ""
AREA1 = ""
AREA2 = ""
ssn = ""
idnumber = ""
ccounty = 0
ht = ""
weight = ""
hair = ""
eyes = ""
birthdate = ""
wname = ""
charge = ""
issuedby = ""
officer = ""
otherrace = ""
assignedto = ""
caucasian.Value = True
male.Value = True
benchfc.Value = True
End Sub

Private Sub lookuplist_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
lookuplist.SortKey = ColumnHeader.index - 1
If lookuplist.SortOrder = lvwAscending Then
    lookuplist.SortOrder = lvwDescending
Else
    lookuplist.SortOrder = lvwAscending
End If
lookuplist.Sorted = True

End Sub

Private Sub lookuplist_ItemClick(ByVal Item As MSComctlLib.ListItem)
If lookuplist.SelectedItem Is Nothing Then
    Exit Sub
End If
Set itmx = lookuplist.ListItems(lookuplist.SelectedItem.index)
warrant = itmx
Call findrec
On Error GoTo 0
lookupframe.Visible = False
lookupbutton.Caption = "&Lookup"
Command8.Caption = "Searc&h"
End Sub

Private Sub NODISP_Click()
If NODISP = 1 Then
    keeplist = 0
End If
End Sub

Private Sub offense_KeyDown(KeyCode As Integer, Shift As Integer)
If (Shift = vbCtrlMask) And (KeyCode = vbKeyF2) Then
        Call SpellCk_Click(1)
    End If
End Sub

Private Sub officer_KeyPress(KeyAscii As Integer)
If Len(officer) = 40 Then
    KeyAscii = 0
End If
End Sub
Private Sub otherrace_GotFocus()
If other.Value = False Then
    male.SetFocus
End If
End Sub

Private Sub ow_Click()
'chris
inp = InputBox("Please Select Report Type                                             Enter 'S' for Standard or Enter 'C' for Custom", "Genesis Information Log", "C")
inp = UCase(inp)
If inp <> "S" And inp <> "C" Then
    msg = MsgBox("          Invalid Entry", vbOKOnly, "Genesis Error Log")
    Exit Sub
End If
Screen.MousePointer = 11
If Val(numcopies) = 0 Then
    CopiesToPrinter = 1
Else
    CopiesToPrinter = Val(numcopies)
End If
REPORT.Destination = 0
REPORT.SelectionFormula = "{warrantinfo.recalled} <> 1 and {warrantinfo.nodisp} = 0 and (IsNull({warrantinfo.sentto}) OR (NOT ISNULL({warrantinfo.sentto}) AND NOT ISNULL({warrantinfo.RETURNEDON}))) AND ISNULL({warrantinfo.whenarrested})"
If inp = "S" Then
    REPORT.ReportFileName = nww + "unserved.RPT"
Else
If inp = "C" Then
    REPORT.ReportFileName = nww + "unserve2.RPT"
End If
End If

If oalpha Then
    If oarea1 Then
        REPORT.SortFields(0) = "+{warrantinfo.wNAME}"
        REPORT.SortFields(1) = "+{warrantinfo.area1}"
    Else
    If oarea2 Then
        REPORT.SortFields(0) = "+{warrantinfo.wNAME}"
        REPORT.SortFields(1) = "+{warrantinfo.area2}"
    Else
        REPORT.SortFields(0) = "+{warrantinfo.wNAME}"
        REPORT.SortFields(1) = ""
    End If
    End If
Else
    If oarea1 Then
        REPORT.SortFields(0) = "+{warrantinfo.warrant}"
        REPORT.SortFields(1) = "+{warrantinfo.area1}"
    Else
    If oarea2 Then
        REPORT.SortFields(0) = "+{warrantinfo.warrant}"
        REPORT.SortFields(1) = "+{warrantinfo.area2}"
    Else
        REPORT.SortFields(0) = "+{warrantinfo.warrant}"
        REPORT.SortFields(1) = ""
    End If
    End If
End If
On Error GoTo 0
REPORT.Action = 1
Screen.MousePointer = 0
'POFRAME.Visible = False
Exit Sub
End Sub

Private Sub owbd_Click()
'chris
inp = InputBox("Please Select Report Type                                             Enter 'S' for Standard or Enter 'C' for Custom", "Genesis Information Log", "C")
inp = UCase(inp)
If inp <> "S" And inp <> "C" Then
    msg = MsgBox("          Invalid Entry", vbOKOnly, "Genesis Error Log")
    Exit Sub
End If
Screen.MousePointer = 11
If Val(numcopies) = 0 Then
    CopiesToPrinter = 1
Else
    CopiesToPrinter = Val(numcopies)
End If
REPORT.Destination = 0
REPORT.SelectionFormula = "(IsNull({warrantinfo.whenarrested}) OR totext({warrantinfo.whenarrested}) = '') and {warrantinfo.nodisp} = 0 and (IsNull({warrantinfo.sentto}) OR (NOT ISNULL({warrantinfo.sentto}) AND NOT ISNULL({warrantinfo.RETURNEDON})))"
If inp = "S" Then
    REPORT.ReportFileName = nww + "unservdp.RPT"
Else
If inp = "C" Then
    REPORT.ReportFileName = nww + "unservd2.RPT"
End If
End If


If oalpha Then
    If oarea1 Then
        REPORT.SortFields(0) = "+{warrantinfo.wNAME}"
        REPORT.SortFields(1) = "+{warrantinfo.area1}"
    Else
    If oarea2 Then
        REPORT.SortFields(0) = "+{warrantinfo.wNAME}"
        REPORT.SortFields(1) = "+{warrantinfo.area2}"
    Else
        REPORT.SortFields(0) = "+{warrantinfo.wNAME}"
        REPORT.SortFields(1) = ""
    End If
    End If
Else
    If oarea1 Then
        REPORT.SortFields(0) = "+{warrantinfo.warrant}"
        REPORT.SortFields(1) = "+{warrantinfo.area1}"
    Else
    If oarea2 Then
        REPORT.SortFields(0) = "+{warrantinfo.warrant}"
        REPORT.SortFields(1) = "+{warrantinfo.area2}"
    Else
        REPORT.SortFields(0) = "+{warrantinfo.warrant}"
        REPORT.SortFields(1) = ""
    End If
    End If
End If
On Error GoTo 0
REPORT.Action = 1
Screen.MousePointer = 0
'POFRAME.Visible = False
Exit Sub

End Sub


Private Sub printbutton_click()
If sprint = 0 And ssupervisor = 0 Then
    msg = MsgBox("You have insufficient authority to print.", 48, "Genesis Error Log")
    Exit Sub
End If
POFRAME.Left = 1950
POFRAME.Top = 1020
POFRAME.Visible = True
End Sub

Private Sub regw_Click()
If regw.Value = True Then
    docket.Enabled = False
    dss.Enabled = False
    guardian.Enabled = False
    courtdate.Enabled = False
    Label16.Enabled = False
    Label17.Enabled = False
    Label18.Enabled = False
    Label19.Enabled = False
End If

End Sub

Private Sub remarks_KeyDown(KeyCode As Integer, Shift As Integer)
If (Shift = vbCtrlMask) And (KeyCode = vbKeyF2) Then
        Call SpellCk_Click(0)
    End If
End Sub

Private Sub returnedon_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
If Len(returnedon) = 1 Or Len(returnedon) = 4 Then
    SendKeys "/"
End If
End If

End Sub

Private Sub returnedon_LostFocus()
If IsDate(returnedon) Then
    returnedon = Format$(returnedon, "mm/dd/yyyy")
End If

End Sub

Private Sub rowb_Click()
StartDate = InputBox("Enter starting date for report.", "Genesis Information Log", "")
If Not IsDate(StartDate) Then
    msg = MsgBox("Invalid start date.", 48, "Genesis Error Log")
    Exit Sub
End If
EndDate = InputBox("Enter ending date for report.", "Genesis Information Log", "")
If Not IsDate(EndDate) Then
    msg = MsgBox("Invalid end date.", 48, "Genesis Error Log")
    Exit Sub
End If
Dim db As Database, ds As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nww + "warrant.mdb")
Set ds = db.OpenRecordset("select * from passthru")
If ds.EOF Then
    ds.AddNew
Else
    ds.Edit
End If
On Error Resume Next
ds("todate") = EndDate
ds("fromdate") = StartDate
ds.Update
Screen.MousePointer = 11
db.Close
If Val(numcopies) = 0 Then
    CopiesToPrinter = 1
Else
    CopiesToPrinter = Val(numcopies)
End If
REPORT.Destination = 0
REPORT.ReportFileName = nww + "record.rpt"
If oalpha Then
    If oarea1 Then
        REPORT.SortFields(0) = "+{warrantinfo.wNAME}"
        REPORT.SortFields(1) = "+{warrantinfo.area1}"
    Else
    If oarea2 Then
        REPORT.SortFields(0) = "+{warrantinfo.wNAME}"
        REPORT.SortFields(1) = "+{warrantinfo.area2}"
    Else
        REPORT.SortFields(0) = "+{warrantinfo.wNAME}"
        REPORT.SortFields(1) = ""
    End If
    End If
Else
    If oarea1 Then
        REPORT.SortFields(0) = "+{warrantinfo.warrant}"
        REPORT.SortFields(1) = "+{warrantinfo.area1}"
    Else
    If oarea2 Then
        REPORT.SortFields(0) = "+{warrantinfo.warrant}"
        REPORT.SortFields(1) = "+{warrantinfo.area2}"
    Else
        REPORT.SortFields(0) = "+{warrantinfo.warrant}"
        REPORT.SortFields(1) = ""
    End If
    End If
End If
StartDate = Format$(StartDate, "mm/dd/yyyy")
starty = Right$(StartDate, 4)
startm = Left$(StartDate, 2)
startd = Mid$(StartDate, 4, 2)
EndDate = Format$(EndDate, "mm/dd/yyyy")
endy = Right$(EndDate, 4)
endm = Left$(EndDate, 2)
endd = Mid$(EndDate, 4, 2)
REPORT.SelectionFormula = "{WARRANTINFO.logdate} >= date(" + starty + "," + startm + "," + startd + ") and {WARRANTINFO.logdate} <= date(" + endy + "," + endm + "," + endd + ")"
REPORT.Action = 1
Screen.MousePointer = 0
'POFRAME.Visible = False

Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If

End Sub

Private Sub savebutton_Click()
If UCase(frmLogin.txtUserName) = "DEMO" And UCase(frmLogin.txtPassword) = "DEMO" Then
    msg = MsgBox("Not available in DEMO version.", 48, "Genesis Information Log")
    Screen.MousePointer = 0
    Exit Sub
End If
If sedit = 0 And ssupervisor = 0 Then
    msg = MsgBox("You have insufficient authority to save.", 48, "Genesis Error Log")
    Exit Sub
End If
If warrant = "" Then
    msg = MsgBox("Warrant# must be entered.", 48, "Genesis Error Log")
    warrant.SetFocus
    Exit Sub
End If
If Len(warrant) > 20 Then
    msg = MsgBox("Warrant# must be 20 characters or less.", 48, "Genesis Error Log")
    warrant.SetFocus
    Exit Sub
End If

If Not IsDate(logdate) Then
    msg = MsgBox("Log Date must be a valid date.", 48, "Genesis Error Log")
    logdate.SetFocus
    Exit Sub
End If
If courtdate > "" And Not IsDate(courtdate) Then
    msg = MsgBox("Court Date must be a valid date.", 48, "Genesis Error Log")
    courtdate.SetFocus
    Exit Sub
End If
If whenarrested > "" And Not IsDate(whenarrested) Then
    msg = MsgBox("When Arrested Date must be a valid date.", 48, "Genesis Error Log")
    whenarrested.SetFocus
    Exit Sub
End If
If birthdate > "" And Not IsDate(birthdate) Then
    msg = MsgBox("Birthdate must be a valid date.", 48, "Genesis Error Log")
    birthdate.SetFocus
    Exit Sub
End If
If senton > "" And Not IsDate(senton) Then
    msg = MsgBox("WARRANT SENT ON must be a valid date.", 48, "Genesis Error Log")
    senton.SetFocus
    Exit Sub
End If
If returnedon > "" And Not IsDate(returnedon) Then
    msg = MsgBox("WARRANT RETURNED ON must be a valid date.", 48, "Genesis Error Log")
    returnedon.SetFocus
    Exit Sub
End If
logdate = Format$(logdate, "mm/dd/yyyy")
If courtdate > "" Then
    courtdate = Format$(courtdate, "mm/dd/yyyy")
End If
If whenarrested > "" Then
    If whenarrested > "" Then
        whenarrested = Format$(whenarrested, "mm/dd/yyyy")
    End If
End If
If birthdate > "" Then
    If birthdate > "" Then
        birthdate = Format$(birthdate, "mm/dd/yyyy")
    End If
End If
If senton > "" Then
    If senton > "" Then
        senton = Format$(senton, "mm/dd/yyyy")
    End If
End If
If returnedon > "" Then
    If returnedon > "" Then
        returnedon = Format$(returnedon, "mm/dd/yyyy")
    End If
End If
If wname = "" Then
    msg = MsgBox("Name must be entered.", 48, "Genesis Error Log")
    wname.SetFocus
    Exit Sub
End If
If charge = "" Then
    msg = MsgBox("Charge must be entered.", 48, "Genesis Error Log")
    charge.SetFocus
    Exit Sub
End If
If issuedby = "" Then
    msg = MsgBox("Issued By must be entered.", 48, "Genesis Error Log")
    issuedby.SetFocus
    Exit Sub
End If
Screen.MousePointer = 11
Call savertn
Call wnameload
Call wload
Call chargeload
Call issuedbyload
Call assignedtoload
Call nullfields
justfound = 0
warrant.SetFocus
Screen.MousePointer = 0
Exit Sub
End Sub

Private Sub senton_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
If Len(senton) = 1 Or Len(senton) = 4 Then
    SendKeys "/"
End If
End If

End Sub

Private Sub senton_LostFocus()
If IsDate(senton) Then
    senton = Format$(senton, "mm/dd/yyyy")
End If

End Sub

Private Sub SpellCk_Click(index As Integer)
If index = 0 Then BeginSpellCheck remarks, remarks
If index = 1 Then BeginSpellCheck offense.Text, offense
If index = 2 Then BeginSpellCheck facts.Text, facts
If index = 3 Then BeginSpellCheck iddata, iddata
End Sub

Private Sub swbd_Click()
StartDate = InputBox("Enter starting date for report.", "Genesis Information Log", "")
If Not IsDate(StartDate) Then
    msg = MsgBox("Invalid start date.", 48, "Genesis Error Log")
    Exit Sub
End If
EndDate = InputBox("Enter ending date for report.", "Genesis Information Log", "")
If Not IsDate(EndDate) Then
    msg = MsgBox("Invalid end date.", 48, "Genesis Error Log")
    Exit Sub
End If
Dim db As Database, ds As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nww + "warrant.mdb")
Set ds = db.OpenRecordset("select * from passthru")
If ds.EOF Then
    ds.AddNew
Else
    ds.Edit
End If
On Error Resume Next
ds("todate") = EndDate
ds("fromdate") = StartDate
ds.Update
Screen.MousePointer = 11
db.Close
If Val(numcopies) = 0 Then
    CopiesToPrinter = 1
Else
    CopiesToPrinter = Val(numcopies)
End If
REPORT.Destination = 0
REPORT.ReportFileName = nww + "servedd.RPT"
If oalpha Then
    If oarea1 Then
        REPORT.SortFields(0) = "+{warrantinfo.wNAME}"
        REPORT.SortFields(1) = "+{warrantinfo.area1}"
    Else
    If oarea2 Then
        REPORT.SortFields(0) = "+{warrantinfo.wNAME}"
        REPORT.SortFields(1) = "+{warrantinfo.area2}"
    Else
        REPORT.SortFields(0) = "+{warrantinfo.wNAME}"
        REPORT.SortFields(1) = ""
    End If
    End If
Else
    If oarea1 Then
        REPORT.SortFields(0) = "+{warrantinfo.warrant}"
        REPORT.SortFields(1) = "+{warrantinfo.area1}"
    Else
    If oarea2 Then
        REPORT.SortFields(0) = "+{warrantinfo.warrant}"
        REPORT.SortFields(1) = "+{warrantinfo.area2}"
    Else
        REPORT.SortFields(0) = "+{warrantinfo.warrant}"
        REPORT.SortFields(1) = ""
    End If
    End If
End If
StartDate = Format$(StartDate, "mm/dd/yyyy")
starty = Right$(StartDate, 4)
startm = Left$(StartDate, 2)
startd = Mid$(StartDate, 4, 2)
EndDate = Format$(EndDate, "mm/dd/yyyy")
endy = Right$(EndDate, 4)
endm = Left$(EndDate, 2)
endd = Mid$(EndDate, 4, 2)
REPORT.SelectionFormula = "{WARRANTINFO.whenarrested} >= date(" + starty + "," + startm + "," + startd + ") and {WARRANTINFO.whenarrested} <= date(" + endy + "," + endm + "," + endd + ")"
REPORT.Action = 1
Screen.MousePointer = 0
'POFRAME.Visible = False


    Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If
End Sub

Private Sub swbdbd_Click()
StartDate = InputBox("Enter starting date for report.", "Genesis Information Log", "")
If Not IsDate(StartDate) Then
    msg = MsgBox("Invalid start date.", 48, "Genesis Error Log")
    Exit Sub
End If
EndDate = InputBox("Enter ending date for report.", "Genesis Information Log", "")
If Not IsDate(EndDate) Then
    msg = MsgBox("Invalid end date.", 48, "Genesis Error Log")
    Exit Sub
End If
Dim db As Database, ds As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nww + "warrant.mdb")
Set ds = db.OpenRecordset("select * from passthru")
If ds.EOF Then
    ds.AddNew
Else
    ds.Edit
End If
On Error Resume Next
ds("todate") = EndDate
ds("fromdate") = StartDate
ds.Update
Screen.MousePointer = 11
db.Close
If Val(numcopies) = 0 Then
    CopiesToPrinter = 1
Else
    CopiesToPrinter = Val(numcopies)
End If
REPORT.Destination = 0
REPORT.ReportFileName = nww + "serveddd.RPT"
If oalpha Then
    If oarea1 Then
        REPORT.SortFields(0) = "+{warrantinfo.wNAME}"
        REPORT.SortFields(1) = "+{warrantinfo.area1}"
    Else
    If oarea2 Then
        REPORT.SortFields(0) = "+{warrantinfo.wNAME}"
        REPORT.SortFields(1) = "+{warrantinfo.area2}"
    Else
        REPORT.SortFields(0) = "+{warrantinfo.wNAME}"
        REPORT.SortFields(1) = ""
    End If
    End If
Else
    If oarea1 Then
        REPORT.SortFields(0) = "+{warrantinfo.warrant}"
        REPORT.SortFields(1) = "+{warrantinfo.area1}"
    Else
    If oarea2 Then
        REPORT.SortFields(0) = "+{warrantinfo.warrant}"
        REPORT.SortFields(1) = "+{warrantinfo.area2}"
    Else
        REPORT.SortFields(0) = "+{warrantinfo.warrant}"
        REPORT.SortFields(1) = ""
    End If
    End If
End If
StartDate = Format$(StartDate, "mm/dd/yyyy")
starty = Right$(StartDate, 4)
startm = Left$(StartDate, 2)
startd = Mid$(StartDate, 4, 2)
EndDate = Format$(EndDate, "mm/dd/yyyy")
endy = Right$(EndDate, 4)
endm = Left$(EndDate, 2)
endd = Mid$(EndDate, 4, 2)
REPORT.SelectionFormula = "{WARRANTINFO.whenarrested} >= date(" + starty + "," + startm + "," + startd + ") and {WARRANTINFO.whenarrested} <= date(" + endy + "," + endm + "," + endd + ")"
REPORT.Action = 1
Screen.MousePointer = 0
'POFRAME.Visible = False
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If
End Sub

Private Sub Text1_Change()

End Sub

Private Sub w4d_Click()
If w4d.Value = True Then
    docket.Enabled = True
    dss.Enabled = True
    guardian.Enabled = True
    courtdate.Enabled = True
    Label16.Enabled = True
    Label17.Enabled = True
    Label18.Enabled = True
    Label19.Enabled = True
End If

End Sub

Private Sub warrant_Change()
If FROMXREF = 1 Then
    Call findrec
    FROMXREF = 0
End If
End Sub

Private Sub warrant_Click()
If warrant > "" Then
    Call findrec
End If
End Sub

Private Sub warrant_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And warrant > "" Then
    Call findrec
End If
End Sub

Private Sub warrant_LostFocus()
If Len(warrant) > 20 Then
    warrant = Left$(warrant, 20)
End If
If warrant > "" Then
    Call findrec
End If
End Sub

Private Sub whenarrested_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
If Len(whenarrested) = 1 Or Len(whenarrested) = 4 Then
    SendKeys "/"
End If
End If

End Sub

Private Sub whenarrested_LostFocus()
If IsDate(whenarrested) Then
    whenarrested = Format$(whenarrested, "mm/dd/yyyy")
End If

End Sub

Private Sub wname_click()
If wname = "" Then
    Exit Sub
End If
Dim db As Database, ds As Recordset

Call setpopup(wname, "L")
On Error GoTo oderror
od:
Set db = OpenDatabase(nww + "warrant.mdb")
If warrant > "" Then
    Set ds = db.OpenRecordset("select recalldate,keeplist, recalled, sex,race,birthdate,logdate,iddata, address,address2,state,zipcode, danger from warrantinfo where warrant <> '" + warrant + "' and wname = '" + wname + "' order by logdate desc")
Else
    msg = MsgBox("Would you like to see a selection list of all warrants in the system for this person?", 4, "Genesis Information Log")
    If msg = 6 Then
        GoSub gethistory
        On Error Resume Next
        Exit Sub
    End If
    Set ds = db.OpenRecordset("select keeplist, recalled, sex,race,birthdate,logdate,iddata, address,address2,state,zipcode, danger,recalldate from warrantinfo where wname = '" + wname + "' order by logdate desc")
End If
If ds.EOF Then
    Exit Sub
End If
ds.MoveFirst
On Error Resume Next
msg = MsgBox("This subject has one or more prior warrants in the system.  Would you like to drag the race, sex, birthdate, address,address2,state,zipcode, identifying data, and danger flag forward from the latest warrant?", 4, "Genesis Information Log")
If msg = 7 Then
    Exit Sub
End If
If Not IsNull(ds("address")) Then
    address(0) = ds("address")
Else
    address(0) = ""
End If
If Not IsNull(ds("address2")) Then
    address(1) = ds("address2")
Else
    address(1) = ""
End If
If Not IsNull(ds("state")) Then
    address(2) = ds("state")
Else
    address(2) = ""
End If
If Not IsNull(ds("zipcode")) Then
    address(3) = ds("zipcode")
Else
    address(3) = ""
End If
If Not IsNull(ds("iddata")) Then
    iddata = ds("iddata")
Else
    iddata = ""
End If
danger.Value = ds("danger")
recalled.Value = ds("recalled")
keeplist.Value = ds("keeplist")
If Not IsNull(ds("recalldate")) Then
    recalldate = ds("recalldate")
Else
    recalldate = ""
End If
If ds("race") = "Caucasian" Or ds("race") = "White" Then
    caucasian.Value = True
Else
If ds("race") = "African-American" Or ds("race") = "Black" Then
    africanamerican.Value = True
Else
If ds("race") = "Oriental" Or ds("race") = "Asian/Pacific Islander" Then
    Oriental.Value = True
Else
If ds("race") = "Indian - American Indian/Alaskan Native" Then
    indian.Value = True
Else
If InStr(UCase(ds("RACE")), "HISPANIC") > 0 Then
    caucasian.Value = True
Else
    other.Value = True
End If
End If
End If
End If
End If
If Not IsNull(ds("birthdate")) Then
    birthdate = ds("birthdate")
End If
If ds("sex") = "Female" Then
    female.Value = True
Else
    male.Value = True
End If
If justfound = 0 Then
    benchfc.SetFocus
Else
    birthdate.SetFocus
End If
NODISP = ds("nodisp")
Set db = OpenDatabase(nwl + "LAWSUITE.MDB")
Set rs = db.OpenRecordset("SELECT * FROM PEOPLE WHERE DPNAMElf = " + Chr$(34) + wname + Chr$(34))
If Not rs.EOF Then
    rs.MoveFirst
    address(0) = rs("DPHADDRESS")
    address(1) = rs("DPHADDRESS2")
    address(2) = rs("hstate")
    address(3) = rs("hzipcode")
    If Not IsNull(rs("height")) Then
        ht = rs("height")
    End If
    If Not IsNull(rs("weight")) Then
        weight = rs("weight")
    End If
    If Not IsNull(rs("hair")) Then
        hair = rs("hair")
    End If
    If Not IsNull(rs("eyes")) Then
        eyes = rs("eyes")
    End If
    If Not IsNull(rs("ssn")) Then
        ssn = rs("ssn")
    End If
    If Not IsNull(rs("idnumber")) Then
        idnumber = rs("idnumber")
    End If
    If Not IsNull(rs("birthdate")) Then
        birthdate = rs("birthdate")
    End If
    If Not IsNull(rs("race")) Then
        Select Case rs("race")
            Case "C", "W", "White"
                caucasian = True
            Case "B", "Black"
                africanamerican = True
            Case "O", "A", "Asian/Pacific Islander"
                Oriental = True
            Case "I", "Indian - American Indian/Alaskan Native"
                indian = True
            Case "U", "Unknown"
                other = True
        End Select
    End If
    If Not IsNull(rs("sex")) Then
        Select Case rs("sex")
            Case "F", "Female"
                female = True
            Case "M", "Male"
                male = True
        End Select
    End If
    If Not IsNull(rs("mugshot")) Then
        mugshot.Picture = LoadPicture(rs("mugshot"))
    Else
        mugshot.Picture = LoadPicture()
    End If
End If
db.Close
Exit Sub
gethistory:
lookuplist.ListItems.clear
Set ds = db.OpenRecordset("select WNAME, warrant, logdate, charge from warrantinfo where wname = '" + wname + "' order by warrant")
If ds.EOF Then
    Return
End If
ds.MoveFirst
While Not ds.EOF
    Set itmx = lookuplist.ListItems.add(, , ds("warrant"))
    itmx.SubItems(1) = ds("logdate")
    itmx.SubItems(2) = ds("wname")
    ds.MoveNext
Wend
lookupframe.Top = 1000
lookupframe.Left = 1000
lookupframe.Visible = True
lookuplist.SetFocus
Return
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
    'Resume
End If
End Sub

Private Sub wname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call wname_click
End If
End Sub

Private Sub wnameload()
Dim db As Database, ds As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nwl + "LAWSUITE.mdb")
Set ds = db.OpenRecordset("select dPNAMElf FROM PEOPLE")
wname.clear
If Not ds.EOF Then
    ds.MoveFirst
End If
While Not ds.EOF
    If ds("DPnamelf") <> " " Then
        wname.AddItem ds("DPnamelf")
    End If
    ds.MoveNext
Wend
db.Close
On Error Resume Next
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
    Resume
End If
End Sub

Private Sub wload()
Dim db As Database, ds As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nww + "warrant.mdb")
Set ds = db.OpenRecordset("select distinct warrant from warrantinfo")
warrant.clear
If Not ds.EOF Then
    ds.MoveFirst
End If
While Not ds.EOF
    warrant.AddItem ds("warrant")
    ds.MoveNext
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
Private Sub loadsystem()
Dim db As Database, rs As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nww + "warrant.mdb")
Set rs = db.OpenRecordset("select * from system")
On Error Resume Next
If rs.EOF Then
    msg = MsgBox("Enter Sheriff Information on System Tab.", 48, "Genesis Information Log")
    db.Close
    Exit Sub
End If
rs.MoveFirst
office = rs("office")
sheriffaddress = rs("sheriffaddress")
sheriffaddress2 = rs("sheriffaddress2")
sheriffphone = rs("sheriffphone")
county = rs("county")
db.Close
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If

End Sub
Private Sub loadcc()
Dim db As Database, ds As Recordset
Set db = OpenDatabase(nww + "warrant.mdb")
Set ds = db.OpenRecordset("select distinct charge from warrantinfo")
cc.clear
If Not ds.EOF Then
    ds.MoveFirst
End If
While Not ds.EOF
    If ds("charge") <> " " Then
        cc.AddItem ds("charge")
    End If
    ds.MoveNext
Wend
db.Close

End Sub

Private Sub wname_KeyUp(KeyCode As Integer, Shift As Integer)
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

Private Sub wname_LostFocus()
'If wname > "" Then
'    Call wname_click
'End If
End Sub
Private Sub loadic()
Dim db As Database, ds As Recordset
Set db = OpenDatabase(nww + "warrant.mdb")
Set ds = db.OpenRecordset("select distinct issuedby from warrantinfo")
ic.clear
If Not ds.EOF Then
    ds.MoveFirst
End If
While Not ds.EOF
    If ds("issuedby") <> " " Then
        ic.AddItem ds("issuedby")
    End If
    ds.MoveNext
Wend
db.Close

End Sub
Private Sub savertn()
Dim db As Database, ds, rs As Recordset, ds2 As Recordset
On Error GoTo oderror1
od1:
Set db = OpenDatabase(nww + "warrant.mdb")
Set ds = db.OpenRecordset("select * from warrantinfo where warrant = '" + warrant + "'")
If ds.EOF Then
    ds.AddNew
    Set ds2 = db.OpenRecordset("select entry from warrantinfo order by entry desc")
    If ds2.EOF Then
        Entry% = 1
    Else
        ds2.MoveFirst
        Entry% = ds2("entry") + 1
    End If
Else
    ds.MoveFirst
    ds.Edit
    Entry% = ds("entry")
End If
On Error GoTo 0
ds("warrant") = Left$(warrant, 20)
ds("logdate") = logdate
ds("danger") = danger.Value
ds("recalled") = recalled.Value
ds("keeplist") = keeplist.Value
If Not IsDate(recalldate) Then
    ds("RECALLDATE") = Null
Else
    ds("recalldate") = recalldate
End If
ds("docket") = docket
ds("dss") = dss
ds("guardian") = guardian
If Not IsDate(courtdate) Then
    ds("COURTDATE") = Null
Else
    ds("courtdate") = courtdate
End If
ds("remarks") = remarks
ds("iddata") = iddata
ds("address") = address(0)
ds("address2") = address(1)
ds("state") = address(2)
ds("zipcode") = address(3)
ds("origination") = origination
If Not IsDate(senton) Then
    ds("SENTON") = Null
Else
    ds("senton") = senton
End If
If Not IsDate(returnedon) Then
    ds("returnedon") = Null
Else
    ds("returnedon") = returnedon
End If
If Not IsDate(whenarrested) Then
    ds("WHENARRESTED") = Null
Else
    ds("whenarrested") = whenarrested
End If
If sentto = "" Then
    ds("SENTTO") = Null
Else
    ds("sentto") = sentto
End If
ds("witness") = witness
ds("casenumber") = casenumber
ds("ssn") = ssn
ds("idnumber") = idnumber
ds("ccounty") = ccounty
ds("area1") = AREA1
ds("area2") = AREA2
ds("height") = ht
ds("weight") = weight
ds("hair") = hair
ds("eyes") = eyes
If Not IsDate(birthdate) Then
    ds("BIRTHDATE") = Null
Else
    ds("birthdate") = birthdate
End If
ds("wname") = Left$(wname, 40)
ds("charge") = Left$(charge, 75)
ds("issuedby") = Left$(issuedby, 40)
ds("officer") = Left$(officer, 40)
ds("ASSIGNEDTO") = Left$(assignedto, 40)
If benchfc.Value = True Then
    ds("type") = "Bench - Family Court"
Else
If benchm.Value = True Then
    ds("type") = "Bench - Magistrate"
Else
If benchgs.Value = True Then
    ds("type") = "Bench - General Sessions"
Else
If regw.Value = True Then
    ds("type") = "Warrant"
Else
    ds("type") = "4D Warrant"
End If
End If
End If
End If
If caucasian.Value = True Then
    ds("race") = "White"
Else
If africanamerican.Value = True Then
    ds("race") = "Black"
Else
If Oriental.Value = True Then
    ds("race") = "Asian/Pacific Islander"
Else
If indian.Value = True Then
    ds("race") = "Indian - American Indian/Alaskan Native"
Else
    ds("race") = "Unknown"
End If
End If
End If
End If
If male.Value = True Then
    ds("sex") = "Male"
Else
    ds("sex") = "Female"
End If
ds("entry") = Entry%
ds("NODISP") = NODISP
ds("court") = court
ds("judge") = judge
ds("defendant") = defendant
ds("plaintiff") = plaintiff
ds("mdesc") = mdesc
ds("offensedate") = offensedate
ds("ind") = ind
ds("county") = county
ds("municipality") = Municipality
ds("offense") = offense.Text
ds("facts") = facts.Text
'CES Code
ds("userfullname") = frmLogin.UserFullName
ds("userid") = frmLogin.UserID
ds("ORINUMBER") = frmLogin.ORINumber
ds("udate") = Format$(Now, "mm/dd/yyyy")
ds("utime") = Format$(Now, "hh:mm:ss")
'********
ds.Update
On Error GoTo oderror2
od2:
ssql = ""
If ssn > "" Then
    ssql = ssql + " and ssn = '" + ssn + "'"
End If
If IsDate(birthdate) Then
    ssql = ssql + " and birthdate = #" + birthdate + "#"
End If
If idnumber > "" Then
    ssql = ssql + " and idnumber = '" + idnumber + "'"
End If
Set db = OpenDatabase(nwl + "lawsuite.mdb")
Set rs = db.OpenRecordset("select * from people where dpnamelf =" + Chr$(34) + wname + Chr$(34) + ssql)
If rs.EOF Then
    rs.AddNew
Else
    rs.MoveFirst
    rs.Edit
End If
On Error GoTo 0
rs("dpnamelf") = wname
rs("dphaddress") = address(0)
rs("dphaddress2") = address(1)
rs("hstate") = address(2)
rs("hzipcode") = address(3)
rs("ssn") = ssn
If IsDate(birthdate) Then
    rs("birthdate") = birthdate
End If
rs("idnumber") = idnumber
rs("dpsort") = Left$(wname, 15)
rs("height") = ht
rs("weight") = weight
rs("hair") = hair
rs("eyes") = eyes
If caucasian.Value = True Then
    rs("race") = "White"
Else
If africanamerican.Value = True Then
    rs("race") = "Black"
Else
If Oriental.Value = True Then
    rs("race") = "Asian/Pacific Islander"
Else
If indian.Value = True Then
    rs("race") = "Indian - American Indian/Alaskan Native"
Else
    rs("race") = "Unknown"
End If
End If
End If
End If
If male.Value = True Then
    rs("sex") = "Male"
Else
    rs("sex") = "Female"
End If
hoLdname = wname
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
    GoTo rsupdate
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
rsupdate:
rs("dpname") = osort1$
On Error GoTo 0
rs.Update
If Not IsDate(whenarrested) Or ccounty.Value = 0 Then
    Exit Sub
End If
If Dir(nwl + "rapsheet.mdb") = "" Then
    Exit Sub
End If
Set db = OpenDatabase(nwl + "rapsheet.mdb")
If casenumber > "" Then
    ssql = ssql + " and casenumber = " + Chr$(34) + casenumber + Chr$(34)
End If
ssql = ssql + " and warrantnumber = " + Chr$(34) + warrant + Chr$(34)
Set rs = db.OpenRecordset("select * from rapsheet where lname = " + Chr$(34) + wname + Chr$(34) + " " + ssql + " and arrestdate = #" + whenarrested + "# and charge = " + Chr$(34) + charge + Chr$(34))
If rs.EOF Then
    rs.AddNew
Else
    rs.MoveFirst
    rs.Edit
End If
rs("lname") = wname
rs("ssn") = ssn
rs("idnumber") = idnumber
If IsDate(birthdate) Then
    rs("birthdate") = CDate(birthdate)
End If
rs("arrestdate") = CDate(whenarrested)
rs("casenumber") = casenumber
rs("warrantnumber") = warrant
rs("charge") = charge
rs.Update
db.Close
Exit Sub
oderror1:
If Err > 3200 Then
    Resume od1
Else
    Resume Next
End If
oderror2:
If Err > 3200 Then
    Resume od2
Else
    Resume Next
End If

End Sub
