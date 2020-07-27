VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form ro 
   Appearance      =   0  'Flat
   BackColor       =   &H00000080&
   Caption         =   "Genesis Restraining Order Manager"
   ClientHeight    =   7380
   ClientLeft      =   255
   ClientTop       =   1575
   ClientWidth     =   11685
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "romain.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7380
   ScaleWidth      =   11685
   Begin VB.Frame Frame100 
      Height          =   7500
      Left            =   15
      TabIndex        =   35
      Top             =   15
      Width           =   11700
      Begin VB.CheckBox moop 
         Caption         =   "Check1"
         Height          =   255
         Left            =   8880
         TabIndex        =   30
         Top             =   6840
         Width           =   255
      End
      Begin VB.TextBox defendantaddress 
         BackColor       =   &H00000080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Index           =   1
         Left            =   6825
         MaxLength       =   60
         TabIndex        =   17
         Top             =   1680
         Width           =   3480
      End
      Begin VB.TextBox plaintiffaddress 
         BackColor       =   &H00000080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Index           =   1
         Left            =   1000
         MaxLength       =   60
         TabIndex        =   6
         Top             =   1680
         Width           =   4200
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   8520
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   6570
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.CommandButton printbutton 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "&Outstanding Report"
         Height          =   350
         Left            =   9550
         TabIndex        =   34
         Top             =   6930
         Width           =   2000
      End
      Begin VB.CommandButton deletebutton 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "&Delete"
         Height          =   350
         Left            =   9550
         TabIndex        =   33
         Top             =   6555
         Width           =   2000
      End
      Begin VB.CommandButton savebutton 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "&Save"
         Height          =   350
         Left            =   9550
         TabIndex        =   31
         Top             =   5805
         Width           =   2000
      End
      Begin VB.TextBox effectivetime 
         BackColor       =   &H00000080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   7440
         MaxLength       =   10
         TabIndex        =   28
         Top             =   6120
         Width           =   1095
      End
      Begin VB.TextBox effectivedate 
         BackColor       =   &H00000080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   5650
         MaxLength       =   10
         TabIndex        =   27
         Top             =   6135
         Width           =   1575
      End
      Begin VB.TextBox expiration 
         BackColor       =   &H00000080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   5650
         MaxLength       =   10
         TabIndex        =   29
         Top             =   6775
         Width           =   1575
      End
      Begin VB.TextBox casenumber 
         BackColor       =   &H00000080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   5880
         MaxLength       =   20
         TabIndex        =   1
         Top             =   200
         Width           =   1665
      End
      Begin VB.TextBox motiondate 
         BackColor       =   &H00000080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   8350
         MaxLength       =   10
         TabIndex        =   2
         Top             =   200
         Width           =   1215
      End
      Begin VB.TextBox hearingdate 
         BackColor       =   &H00000080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   10440
         MaxLength       =   10
         TabIndex        =   3
         Top             =   200
         Width           =   1215
      End
      Begin VB.TextBox checkdon 
         BackColor       =   &H00000080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   420
         Left            =   8760
         MaxLength       =   60
         MultiLine       =   -1  'True
         TabIndex        =   26
         Top             =   5325
         Width           =   2730
      End
      Begin VB.CheckBox checkd 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   5655
         TabIndex        =   25
         Top             =   5100
         Width           =   220
      End
      Begin VB.CheckBox checkc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   50
         TabIndex        =   14
         Top             =   6525
         Width           =   220
      End
      Begin VB.TextBox checkbwhere 
         BackColor       =   &H00000080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   5880
         MaxLength       =   60
         MultiLine       =   -1  'True
         TabIndex        =   24
         Top             =   4350
         Width           =   5625
      End
      Begin VB.CheckBox checkb 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   5655
         TabIndex        =   23
         Top             =   3360
         Width           =   220
      End
      Begin VB.CheckBox checka 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   50
         TabIndex        =   13
         Top             =   5565
         Width           =   220
      End
      Begin VB.TextBox threat 
         BackColor       =   &H00000080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   960
         Left            =   50
         MaxLength       =   200
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   4560
         Width           =   5040
      End
      Begin VB.TextBox occurredin 
         BackColor       =   &H00000080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   50
         MaxLength       =   20
         TabIndex        =   11
         Top             =   3720
         Width           =   4980
      End
      Begin VB.CheckBox check4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   50
         TabIndex        =   10
         Top             =   2760
         Width           =   220
      End
      Begin VB.TextBox plaintiffstate 
         BackColor       =   &H00000080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   3360
         MaxLength       =   2
         TabIndex        =   8
         Top             =   2265
         Width           =   525
      End
      Begin VB.TextBox plaintiffcounty 
         BackColor       =   &H00000080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   45
         MaxLength       =   20
         TabIndex        =   7
         Top             =   2265
         Width           =   3195
      End
      Begin VB.TextBox plaintiffaddress 
         BackColor       =   &H00000080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Index           =   0
         Left            =   1000
         MaxLength       =   60
         TabIndex        =   5
         Top             =   1260
         Width           =   4200
      End
      Begin VB.ComboBox index 
         BackColor       =   &H00000080&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   720
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   195
         Width           =   4200
      End
      Begin VB.ComboBox plaintiff 
         BackColor       =   &H00000080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   60
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   840
         Width           =   5200
      End
      Begin VB.ComboBox defendant 
         BackColor       =   &H00000080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   5670
         Sorted          =   -1  'True
         TabIndex        =   15
         Top             =   840
         Width           =   4620
      End
      Begin VB.TextBox defendantemployaddr 
         BackColor       =   &H00000080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   7000
         MaxLength       =   60
         TabIndex        =   22
         Top             =   3060
         Width           =   4575
      End
      Begin VB.TextBox defendantemploy 
         BackColor       =   &H00000080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   7000
         MaxLength       =   60
         TabIndex        =   21
         Top             =   2685
         Width           =   4575
      End
      Begin VB.TextBox defendantstate 
         BackColor       =   &H00000080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   9480
         MaxLength       =   2
         TabIndex        =   19
         Top             =   2265
         Width           =   495
      End
      Begin VB.TextBox defendantcounty 
         BackColor       =   &H00000080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   5685
         MaxLength       =   20
         TabIndex        =   18
         Top             =   2265
         Width           =   3525
      End
      Begin VB.TextBox defendantaddress 
         BackColor       =   &H00000080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Index           =   0
         Left            =   6825
         MaxLength       =   60
         TabIndex        =   16
         Top             =   1260
         Width           =   3480
      End
      Begin VB.Data Data2 
         Caption         =   "Data2"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   8400
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   7200
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "&Clear"
         Height          =   350
         Left            =   9550
         TabIndex        =   32
         Top             =   6180
         Width           =   2000
      End
      Begin VB.TextBox plaintiffzipcode 
         BackColor       =   &H00000080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   4035
         MaxLength       =   10
         TabIndex        =   9
         Top             =   2265
         Width           =   1185
      End
      Begin VB.TextBox defendantzipcode 
         BackColor       =   &H00000080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   10350
         MaxLength       =   10
         TabIndex        =   20
         Top             =   2280
         Width           =   1185
      End
      Begin Crystal.CrystalReport report 
         Left            =   9960
         Top             =   6840
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         ReportFileName  =   "c:\genesis\prod\ro\ineffect.rpt"
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileLinesPerPage=   60
      End
      Begin VB.Image MUGSHOT 
         BorderStyle     =   1  'Fixed Single
         Height          =   1410
         Left            =   10320
         Stretch         =   -1  'True
         Top             =   645
         Width           =   1245
      End
      Begin VB.Label Label25 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Magistrate's Order of Protection"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   495
         Left            =   7440
         TabIndex        =   62
         Top             =   6555
         Width           =   2055
      End
      Begin VB.Label Label24 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Time:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   7440
         TabIndex        =   61
         Top             =   5880
         Width           =   540
      End
      Begin VB.Label Label23 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Effective Date:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   5650
         TabIndex        =   60
         Top             =   5880
         Width           =   1575
      End
      Begin VB.Label Label22 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Expiration Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   5650
         TabIndex        =   59
         Top             =   6555
         Width           =   1695
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Case Number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   495
         Left            =   5040
         TabIndex        =   58
         Top             =   120
         Width           =   960
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Motion Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   495
         Left            =   7635
         TabIndex        =   57
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Hearing Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   495
         Left            =   9600
         TabIndex        =   56
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label21 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "A copy of this Order shall be served on the following law enforcement agencies:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   690
         Left            =   5880
         TabIndex        =   55
         Top             =   5085
         Width           =   5535
      End
      Begin VB.Label Label20 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   $"romain.frx":030A
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   795
         Left            =   255
         TabIndex        =   54
         Top             =   6525
         Width           =   5520
      End
      Begin VB.Label Label19 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   $"romain.frx":0392
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   1260
         Left            =   5880
         TabIndex        =   53
         Top             =   3390
         Width           =   5655
      End
      Begin VB.Label Label18 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   $"romain.frx":0447
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   945
         Left            =   255
         TabIndex        =   52
         Top             =   5520
         Width           =   5145
      End
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "The Defendant has committed the following acts which constitute Harassment or Stalking:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   615
         Left            =   45
         TabIndex        =   51
         Top             =   4080
         Width           =   4755
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "The Harrassment or Stalking, as described herein, occurred in                                              County, S.C."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   510
         Left            =   120
         TabIndex        =   50
         Top             =   3240
         Width           =   4755
      End
      Begin VB.Label Label15 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "The Defendant is a nonresident of this state or cannot be found."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   570
         Left            =   255
         TabIndex        =   49
         Top             =   2640
         Width           =   4260
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "State"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   3375
         TabIndex        =   48
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "County"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   45
         TabIndex        =   47
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Plaintiff Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   495
         Left            =   45
         TabIndex        =   46
         Top             =   1260
         Width           =   1095
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Index"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   50
         TabIndex        =   45
         Top             =   200
         Width           =   1815
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Defendant"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   5655
         TabIndex        =   44
         Top             =   585
         Width           =   5655
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Plaintiff"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   45
         TabIndex        =   43
         Top             =   585
         Width           =   1815
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   5655
         TabIndex        =   42
         Top             =   3090
         Width           =   5655
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Defendant Employment"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   495
         Left            =   5655
         TabIndex        =   41
         Top             =   2610
         Width           =   1335
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "State"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   495
         Left            =   9510
         TabIndex        =   40
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "County"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   495
         Left            =   5655
         TabIndex        =   39
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Defendant Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   495
         Left            =   5655
         TabIndex        =   38
         Top             =   1260
         Width           =   1095
      End
      Begin VB.Label Label26 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Zipcode"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   4050
         TabIndex        =   37
         Top             =   2040
         Width           =   1110
      End
      Begin VB.Label Label27 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Zipcode"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   10365
         TabIndex        =   36
         Top             =   2055
         Width           =   1110
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         Height          =   6750
         Left            =   30
         Top             =   600
         Width           =   5385
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FFFFFF&
         Height          =   6765
         Left            =   5490
         Top             =   600
         Width           =   6105
      End
   End
End
Attribute VB_Name = "ro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nametype, FROMXREF As Integer
Dim sedit, sprint, sreport, sbrowse, sdelete, ssupervisor As Integer
Private Sub casenumber_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If plaintiff > "" And defendant > "" And casenumber > "" Then
        Call findrec
    End If
End If

End Sub

Private Sub casenumber_LostFocus()
If plaintiff > "" And defendant > "" And casenumber > "" Then
    Call findrec
End If

End Sub

Private Sub Command1_Click()
Call nullfields
End Sub

Private Sub Command6_Click()
Screen.MousePointer = 11
If sedit = 1 Then
    Dim db As Database, ds As Recordset
    On Error GoTo oderror
od:
    Set db = OpenDatabase(nwr + "ro.mdb")
    Set ds = db.OpenRecordset("select * from system")
    If ds.EOF Then
        ds.AddNew
    Else
        ds.MoveFirst
        ds.Edit
    End If
    On Error Resume Next
    ds("sheriffaddress") = sheriffaddress
    ds("sheriffaddress2") = sheriffaddress2
    ds("sheriffphone") = sheriffphone
    ds("sheriff") = sheriff
    ds("county") = county
    ds("office") = office
    ds.Update
Else
    msg = MsgBox("You have insufficient authority for this operation.", 48, "Genesis Error Log")
End If
On Error Resume Next
db.Close
Screen.MousePointer = 0
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If

End Sub

Private Sub defendant_Click()
If defendant = "" Then
    Exit Sub
End If
Dim db As Database, ds As Recordset
On Error GoTo oderror
od:
Call setpopup(defendant, "L")
Set db = OpenDatabase(nwl + "lawsuite.mdb")
Set ds = db.OpenRecordset("select * from people WHERE dpnamelf = " + Chr$(34) + defendant + Chr$(34))
If Not ds.EOF Then
    ds.MoveFirst
    If Not IsNull(ds("dphaddress")) Then
        defendantaddress(0) = ds("dphaddress")
    End If
    If Not IsNull(ds("dphaddress2")) Then
        defendantaddress(1) = ds("dphaddress2")
    End If
    If Not IsNull(ds("hstate")) Then
        defendantstate = ds("hstate")
    End If
    If Not IsNull(ds("hzipcode")) Then
        defendantzipcode = ds("hzipcode")
    End If
    If Not IsNull(ds("mugshot")) Then
        MUGSHOT.Picture = LoadPicture(ds("mugshot"))
    End If
End If
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

Private Sub defendant_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If plaintiff > "" And defendant > "" And casenumber > "" Then
        Call findrec
    End If
End If

End Sub

Private Sub DEFENDANT_KeyUp(KeyCode As Integer, Shift As Integer)
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

Private Sub defendant_LostFocus()
If defendant > "" And InStr(defendant, ",") = 0 Then
    msg = MsgBox("All names in the Restraining Order system should be entered in the format last name + comma + firstname.", 48, "Invalid Data Format")
    defendant.SetFocus
End If
If plaintiff > "" And defendant > "" And casenumber > "" Then
    Call findrec
End If

End Sub

Private Sub defendantaddress_GotFocus(index As Integer)
If defendantaddress(index) > "" Or defendant = "" Then
    Exit Sub
End If
Dim db As Database, ds As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nwl + "LAWSUITE.mdb")
Set ds = db.OpenRecordset("select * FROM PEOPLE WHERE DPNAMElf = " + Chr$(34) + defendant + Chr$(34))
If Not ds.EOF Then
    ds.MoveFirst
    If Not IsNull(ds("DPHaddress")) Then
        defendantaddress(0) = ds("DPHaddress")
        defendantaddress(1) = ds("DPHADDRESS2")
    End If
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
If plaintiff = "" Or defendant = "" Or casenumber = "" Then
    msg = MsgBox("PLAINTIFF, DEFENDANT, and CASENUMBER must all be entered.", 48, "Missing Data")
    Exit Sub
End If
msg = MsgBox("Are you sure you wish to delete this record?", 4, "Genesis Information Log")
If msg = 7 Then
    Exit Sub
End If
Screen.MousePointer = 11
Dim db As Database, ds As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nwr + "ro.mdb")
Set ds = db.OpenRecordset("select * from rorder where plaintiff = " + Chr$(34) + plaintiff + Chr$(34) + " and defendant = " + Chr$(34) + defendant + Chr$(34) + " and casenumber = " + Chr$(34) + casenumber + Chr$(34))
If Not ds.EOF Then
    ds.MoveFirst
    ds.Delete
End If
On Error Resume Next
db.Close
Call nullfields
Call loadindex
Call LOADPEOPLE
plaintiff.SetFocus
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
Unload romain
End
End Sub

Private Sub findrec()
Dim db As Database, ds As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nwr + "ro.mdb")
Set ds = db.OpenRecordset("select * from rorder where plaintiff = " + Chr$(34) + plaintiff + Chr$(34) + " and defendant = " + Chr$(34) + defendant + Chr$(34) + " and casenumber = " + Chr$(34) + casenumber + Chr$(34))
If Not ds.EOF Then
    ds.MoveFirst
Else
    On Error Resume Next
    db.Close
    Exit Sub
End If
plaintiff = ds("plaintiff")
defendant = ds("defendant")
casenumber = ds("casenumber")
If Not IsNull(ds("motiondate")) Then
    motiondate = ds("motiondate")
Else
    motiondate = ""
End If
If Not IsNull(ds("hearingdate")) Then
    hearingdate = ds("hearingdate")
Else
    hearingdate = ""
End If
If Not IsNull(ds("plaintiffaddress")) Then
    plaintiffaddress(0) = ds("plaintiffaddress")
Else
    plaintiffaddress(0) = ""
End If
If Not IsNull(ds("plaintiffaddress2")) Then
    plaintiffaddress(1) = ds("plaintiffaddress2")
Else
    plaintiffaddress(1) = ""
End If
If Not IsNull(ds("plaintiffcounty")) Then
    plaintiffcounty = ds("plaintiffcounty")
Else
    plaintiffcounty = ""
End If
If Not IsNull(ds("plaintiffstate")) Then
    plaintiffstate = ds("plaintiffstate")
Else
    plaintiffstate = ""
End If
If Not IsNull(ds("plaintiffzipcode")) Then
    plaintiffzipcode = ds("plaintiffzipcode")
Else
    plaintiffzipcode = ""
End If
If Not IsNull(ds("defendantaddress")) Then
    defendantaddress(0) = ds("defendantaddress")
Else
    defendantaddress(0) = ""
End If
If Not IsNull(ds("defendantaddress2")) Then
    defendantaddress(1) = ds("defendantaddress2")
Else
    defendantaddress(1) = ""
End If
If Not IsNull(ds("defendantcounty")) Then
    defendantcounty = ds("defendantcounty")
Else
    defendantcounty = ""
End If
If Not IsNull(ds("defendantstate")) Then
    defendantstate = ds("defendantstate")
Else
    defendantstate = ""
End If
If Not IsNull(ds("defendantzipcode")) Then
    defendantzipcode = ds("defendantzipcode")
Else
    defendantzipcode = ""
End If
If Not IsNull(ds("defendantemploy")) Then
    defendantemploy = ds("defendantemploy")
Else
    defendantemploy = ""
End If
If Not IsNull(ds("defendantemployaddr")) Then
    defendantemployaddr = ds("defendantemployaddr")
Else
    defendantemployaddr = ""
End If
check4.Value = ds("check4")
If Not IsNull(ds("occurredin")) Then
    occurredin = ds("occurredin")
Else
    occurredin = ""
End If
If Not IsNull(ds("threat")) Then
    threat = ds("threat")
Else
    threat = ""
End If
checka.Value = ds("checka")
checkb.Value = ds("checkb")
If Not IsNull(ds("checkbwhere")) Then
    checkbwhere = ds("checkbwhere")
Else
    checkbwhere = ""
End If
checkc.Value = ds("checkc")
checkd.Value = ds("checkd")
If Not IsNull(ds("checkdon")) Then
    checkdon = ds("checkdon")
Else
    checkdon = ""
End If
expiration = ds("expiration")
If Not IsNull(ds("effectivedate")) Then
    effectivedate = ds("effectivedate")
End If
If Not IsNull(ds("effectivetime")) Then
    effectivetime = ds("effectivetime")
Else
    effectivetime = ""
End If
moop = 0
On Error Resume Next
If Not IsNull(ds("moop")) Then
    moop = ds("moop")
End If
db.Close



Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If
End Sub

Private Sub effectivedate_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
If Len(effectivedate) = 1 Or Len(effectivedate) = 4 Then
    SendKeys "/"
End If
End If
End Sub

Private Sub effectivetime_KeyPress(KeyAscii As Integer)
If Len(effectivetime) = 1 Then
    SendKeys ":"
End If

End Sub

Private Sub expiration_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
If Len(expiration) = 1 Or Len(expiration) = 4 Then
    SendKeys "/"
End If
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
On Error Resume Next
sedit = frmLogin.redit
sprint = frmLogin.rprint
sreport = frmLogin.rreport
sbrowse = frmLogin.rbrowse
sdelete = frmLogin.rdelete
ssupervisor = frmLogin.rsupervisor
Call nullfields
Call loadindex
Call LOADPEOPLE
Call loadsystem
Me.Top = 0
Me.Left = 0
Me.Height = 7700
Me.Width = 11700
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set ro = Nothing
End Sub

Private Sub hearingdate_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
If Len(hearingdate) = 1 Or Len(hearingdate) = 4 Then
    SendKeys "/"
End If
End If
End Sub

Private Sub index_Change()
If FROMXREF = 1 Then
    Call index_Click
    FROMXREF = 0
End If
End Sub

Private Sub index_Click()
If index = "" Then
    Exit Sub
End If
Dim db As Database, ds As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nwr + "ro.mdb")
Set ds = db.OpenRecordset("select * from rorder where plaintiff = " + Chr$(34) + Left$(index, InStr(index, " VS. ") - 1) + Chr$(34) + " and defendant = " + Chr$(34) + Mid$(index, InStr(index, " VS. ") + 5, InStr(index, "CASE#") - InStr(index, " VS. ") - 7) + Chr$(34) + " and casenumber = " + Chr$(34) + Mid$(index, InStr(index, "CASE#:") + 7) + Chr$(34))
If Not ds.EOF Then
    ds.MoveFirst
    Call nullfields
Else
    On Error Resume Next
    db.Close
    Exit Sub
End If
plaintiff = ds("plaintiff")
defendant = ds("defendant")
casenumber = ds("casenumber")
If Not IsNull(ds("motiondate")) Then
    motiondate = ds("motiondate")
Else
    motiondate = ""
End If
If Not IsNull(ds("hearingdate")) Then
    hearingdate = ds("hearingdate")
Else
    hearingdate = ""
End If
If Not IsNull(ds("plaintiffaddress")) Then
    plaintiffaddress(0) = ds("plaintiffaddress")
Else
    plaintiffaddress(0) = ""
End If
If Not IsNull(ds("plaintiffaddress2")) Then
    plaintiffaddress(1) = ds("plaintiffaddress2")
Else
    plaintiffaddress(1) = ""
End If
If Not IsNull(ds("plaintiffcounty")) Then
    plaintiffcounty = ds("plaintiffcounty")
Else
    plaintiffcounty = ""
End If
If Not IsNull(ds("plaintiffstate")) Then
    plaintiffstate = ds("plaintiffstate")
Else
    plaintiffstate = ""
End If
If Not IsNull(ds("plaintiffzipcode")) Then
    plaintiffzipcode = ds("plaintiffzipcode")
Else
    plaintiffzipcode = ""
End If
If Not IsNull(ds("defendantaddress")) Then
    defendantaddress(0) = ds("defendantaddress")
Else
    defendantaddress(0) = ""
End If
If Not IsNull(ds("defendantaddress2")) Then
    defendantaddress(1) = ds("defendantaddress2")
Else
    defendantaddress(1) = ""
End If
If Not IsNull(ds("defendantcounty")) Then
    defendantcounty = ds("defendantcounty")
Else
    defendantcounty = ""
End If
If Not IsNull(ds("defendantstate")) Then
    defendantstate = ds("defendantstate")
Else
    defendantstate = ""
End If
If Not IsNull(ds("defendantzipcode")) Then
    defendantzipcode = ds("defendantzipcode")
Else
    defendantzipcode = ""
End If
If Not IsNull(ds("defendantemploy")) Then
    defendantemploy = ds("defendantemploy")
Else
    defendantemploy = ""
End If
If Not IsNull(ds("defendantemployaddr")) Then
    defendantemployaddr = ds("defendantemployaddr")
Else
    defendantemployaddr = ""
End If
check4.Value = ds("check4")
If Not IsNull(ds("occurredin")) Then
    occurredin = ds("occurredin")
Else
    occurredin = ""
End If
If Not IsNull(ds("threat")) Then
    threat = ds("threat")
Else
    threat = ""
End If
checka.Value = ds("checka")
checkb.Value = ds("checkb")
If Not IsNull(ds("checkbwhere")) Then
    checkbwhere = ds("checkbwhere")
Else
    checkbwhere = ""
End If
checkc.Value = ds("checkc")
checkd.Value = ds("checkd")
If Not IsNull(ds("checkdon")) Then
    checkdon = ds("checkdon")
Else
    checkdon = ""
End If
expiration = ds("expiration")
If Not IsNull(ds("effectivedate")) Then
   effectivedate = ds("effectivedate")
End If
If Not IsNull(ds("effectivetime")) Then
    effectivetime = ds("effectivetime")
Else
    effectivetime = ""
End If
moop = 0
On Error Resume Next
If Not IsNull(ds("moop")) Then
    moop = ds("moop")
End If
Set db = OpenDatabase(nwl + "lawsuite.mdb")
Set rs = db.OpenRecordset("select mugshot from people where dpnamelf = " + Chr$(34) + defendant + Chr$(34) + " and not mugshot is null")
If Not rs.EOF Then
    rs.MoveFirst
    MUGSHOT.Picture = LoadPicture(rs("mugshot"))
Else
    MUGSHOT.Picture = LoadPicture()
End If

db.Close


Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If
End Sub

Private Sub loadindex()
Dim db As Database, ds As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nwr + "ro.mdb")
Set ds = db.OpenRecordset("select plaintiff,defendant,casenumber from rorder order by plaintiff,defendant,casenumber")
index.clear
If Not ds.EOF Then
    ds.MoveFirst
End If
While Not ds.EOF
    index.AddItem ds("plaintiff") + " VS. " + ds("defendant") + "  CASE#: " + ds("casenumber")
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

Private Sub nullfields()
MUGSHOT.Picture = LoadPicture()
plaintiff = ""
index = ""
defendant = ""
casenumber = ""
motiondate = ""
hearingdate = ""
plaintiffaddress(0) = ""
plaintiffaddress(1) = ""
plaintiffcounty = ""
plaintiffstate = ""
plaintiffzipcode = ""
defendantaddress(0) = ""
defendantaddress(1) = ""
defendantcounty = ""
defendantstate = ""
defendantzipcode = ""
defendantemploy = ""
defendantemployaddr = ""
check4.Value = 0
occurredin = ""
threat = ""
checka.Value = 0
checkb.Value = 0
checkbwhere = ""
checkc.Value = 0
checkd.Value = 0
checkdon = ""
effectivedate = ""
effectivetime = ""
expiration = ""
moop = 0
End Sub

Private Sub motiondate_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
If Len(motiondate) = 1 Or Len(motiondate) = 4 Then
    SendKeys "/"
End If
End If

End Sub

Private Sub plaintiff_Click()
If plaintiff = "" Then
    Exit Sub
End If
Dim db As Database, ds As Recordset
On Error GoTo oderror
od:
Call setpopup(plaintiff, "L")
Set db = OpenDatabase(nwl + "lawsuite.mdb")
Set ds = db.OpenRecordset("select * from people WHERE dpnamelf = " + Chr$(34) + plaintiff + Chr$(34))
If Not ds.EOF Then
    ds.MoveFirst
If Not ds.EOF Then
    ds.MoveFirst
    If Not IsNull(ds("dphaddress")) Then
        plaintiffaddress(0) = ds("dphaddress")
    End If
    If Not IsNull(ds("dphaddress2")) Then
        plaintiffaddress(1) = ds("dphaddress2")
    End If
    If Not IsNull(ds("hstate")) Then
        plaintiffstate = ds("hstate")
    End If
    If Not IsNull(ds("hzipcode")) Then
        plaintiffzipcode = ds("hzipcode")
    End If
End If
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

Private Sub plaintiff_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If plaintiff > "" And defendant > "" And casenumber > "" Then
        Call findrec
    End If
End If
End Sub

Private Sub PLAINTIFF_KeyUp(KeyCode As Integer, Shift As Integer)
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

Private Sub plaintiff_LostFocus()
If plaintiff > "" And InStr(plaintiff, ",") = 0 Then
    msg = MsgBox("All names in the Restraining Order system should be entered in the format last name + comma + firstname.", 48, "Invalid Data Format")
    plaintiff.SetFocus
End If
If plaintiff > "" And defendant > "" And casenumber > "" Then
    Call findrec
End If
End Sub

Private Sub plaintiffaddress_GotFocus(index As Integer)
If plaintiffaddress(index) > "" Or plaintiff = "" Then
    Exit Sub
End If
Dim db As Database, ds As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nwl + "LAWSUITE.mdb")
Set ds = db.OpenRecordset("select * FROM PEOPLE WHERE DPNAMElf = " + Chr$(34) + plaintiff + Chr$(34))
If Not ds.EOF Then
    ds.MoveFirst
    If Not IsNull(ds("DPHaddress")) Then
        plaintiffaddress(0) = ds("DPHaddress")
    End If
    If Not IsNull(ds("DPHaddress2")) Then
        plaintiffaddress(1) = ds("DPHaddress2")
    End If
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

Private Sub printbutton_click()
If sprint = 0 And ssupervisor = 0 Then
    msg = MsgBox("You have insufficient authority to print.", 48, "Genesis Error Log")
    Exit Sub
End If
Screen.MousePointer = 11
inp = UCase(inp)
report.Destination = 0
report.ReportFileName = nwr + "ineffect.rpt"
yy$ = Right$(Date$, 4)
MM$ = Left$(Date$, 2)
dd$ = Mid$(Date$, 4, 2)
report.SelectionFormula = "{rorder.expiration} >= date(" + yy$ + "," + MM$ + "," + dd$ + ")"
report.Action = 1
Screen.MousePointer = 0
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
If plaintiff = "" Then
    msg = MsgBox("Plaintiff is a required field.", 48, "Genesis Error Log")
    plaintiff.SetFocus
    Exit Sub
End If
If defendant = "" Then
    msg = MsgBox("Defendant is a required field.", 48, "Genesis Error Log")
    defendant.SetFocus
    Exit Sub
End If
If casenumber = "" Then
    msg = MsgBox("Case Number is a required field.", 48, "Genesis Error Log")
    casenumber.SetFocus
    Exit Sub
End If
If Not IsDate(expiration) Then
    msg = MsgBox("Expiration Date must be a valid date.", 48, "Genesis Error Log")
    expiration.SetFocus
    Exit Sub
End If
expirationdate = Format$(expirationdate, "mm/dd/yyyy")
If motiondate > "" And Not IsDate(motiondate) Then
    msg = MsgBox("Invalid date in focus field.", 48, "Genesis Error Log")
    motiondate.SetFocus
    Exit Sub
End If
If motiondate > "" Then
    motiondate = Format$(motiondate, "mm/dd/yyyy")
End If
If hearingdate > "" And Not IsDate(hearingdate) Then
    msg = MsgBox("Invalid date in focus field.", 48, "Genesis Error Log")
    hearingdate.SetFocus
    Exit Sub
End If
If hearingdate > "" Then
    hearingdate = Format$(hearingdate, "mm/dd/yyyy")
End If
If Not IsDate(effectivedate) Then
    msg = MsgBox("Invalid date in focus field.", 48, "Genesis Error Log")
    effectivedate.SetFocus
    Exit Sub
End If
effectivedate = Format$(effectivedate, "mm/dd/yyyy")
Dim db As Database, ds As Recordset
On Error GoTo oderror1
od1:
Set db = OpenDatabase(nwr + "ro.mdb")
Set ds = db.OpenRecordset("select * from rorder where plaintiff = " + Chr$(34) + plaintiff + Chr$(34) + " and defendant = " + Chr$(34) + defendant + Chr$(34) + " and casenumber = " + Chr$(34) + casenumber + Chr$(34))
If Not ds.EOF Then
    ds.MoveFirst
        ds.Edit
Else
        ds.AddNew
End If
Screen.MousePointer = 11
ds("plaintiff") = plaintiff
ds("defendant") = defendant
ds("casenumber") = casenumber
If IsDate(motiondate) Then
   ds("motiondate") = motiondate
Else
    ds("motiondate") = Null
End If
If IsDate(hearingdate) Then
    ds("hearingdate") = hearingdate
Else
    ds("hearingdate") = Null
End If
ds("plaintiffaddress") = plaintiffaddress(0)
ds("plaintiffaddress2") = plaintiffaddress(1)
ds("plaintiffcounty") = plaintiffcounty
ds("plaintiffstate") = plaintiffstate
ds("plaintiffzipcode") = plaintiffzipcode
ds("defendantaddress") = defendantaddress(0)
ds("defendantaddress2") = defendantaddress(1)
ds("defendantcounty") = defendantcounty
ds("defendantstate") = defendantstate
ds("defendantzipcode") = defendantzipcode
ds("defendantemploy") = defendantemploy
ds("defendantemployaddr") = defendantemployaddr
ds("check4") = check4.Value
ds("occurredin") = occurredin
ds("threat") = threat
ds("checka") = checka.Value
ds("checkb") = checkb.Value
ds("checkbwhere") = checkbwhere
ds("checkc") = checkc.Value
ds("checkd") = checkd.Value
ds("checkdon") = checkdon
ds("expiration") = expiration
If IsDate(effectivedate) Then
    ds("effectivedate") = effectivedate
Else
    ds("effectivedate") = Null
End If
ds("effectivetime") = effectivetime
ds("moop") = moop
'CES Code
ds("userfullname") = frmLogin.UserFullName
ds("userid") = frmLogin.UserID
ds("ORINUMBER") = frmLogin.orinumber
ds("udate") = Format$(Now, "mm/dd/yyyy")
ds("utime") = Format$(Now, "hh:mm:ss")
'********
ds.Update
On Error Resume Next
On Error GoTo oderror2
od2:
Set db = OpenDatabase(nwl + "lawsuite.mdb")
Set rs = db.OpenRecordset("select * from people where dpnamelf =" + Chr$(34) + plaintiff + Chr$(34))
If rs.EOF Then
    rs.AddNew
Else
    rs.MoveFirst
    rs.Edit
End If
rs("dpnamelf") = plaintiff
rs("dphaddress") = plaintiffaddress(0)
rs("dphaddress2") = plaintiffaddress(1)
rs("dpsort") = Left$(plaintiff, 15)
rs("hstate") = plaintiffstate
rs("hzipcode") = plaintiffzipcode

hoLdname = plaintiff
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
rs.Update
Set rs = db.OpenRecordset("select * from people where dpnamelf =" + Chr$(34) + defendant + Chr$(34))
If rs.EOF Then
    rs.AddNew
Else
    rs.MoveFirst
    rs.Edit
End If
rs("dpnamelf") = defendant
rs("dphaddress") = defendantaddress(0)
rs("dphaddress2") = defendantaddress(1)
rs("hstate") = defendantstate
rs("hzipcode") = defendantzipcode
rs("dpsort") = Left$(defendant, 15)
hoLdname = defendant
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
db.Close
On Error Resume Next
Call nullfields
Call loadindex
Call LOADPEOPLE
plaintiff.SetFocus
Screen.MousePointer = 0
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

Private Sub LOADPEOPLE()
Dim db As Database, ds As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nwl + "LAWSUITE.mdb")
Set ds = db.OpenRecordset("select DPNAMElf FROM PEOPLE ORDER BY DPNAMElf")
plaintiff.clear
defendant.clear
If Not ds.EOF Then
    ds.MoveFirst
End If
While Not ds.EOF
    plaintiff.AddItem ds("DPNAMElf")
    defendant.AddItem ds("DPNAMElf")
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
Set db = OpenDatabase(nwr + "ro.mdb")
Set rs = db.OpenRecordset("select * from system")
If rs.EOF Then
    On Error Resume Next
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
On Error Resume Next
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If
End Sub

