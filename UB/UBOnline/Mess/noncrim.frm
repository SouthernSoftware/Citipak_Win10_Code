VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form noncrim 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Genesis Non-Criminal Police Response"
   ClientHeight    =   7365
   ClientLeft      =   270
   ClientTop       =   1305
   ClientWidth     =   11655
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7365
   ScaleWidth      =   11655
   Begin VB.Frame Frame1 
      Height          =   5715
      Left            =   8085
      TabIndex        =   111
      Top             =   1665
      Width           =   3570
      Begin VB.CommandButton SPellCk 
         Caption         =   "Spelling"
         Height          =   195
         Index           =   1
         Left            =   2085
         TabIndex        =   121
         Top             =   2640
         Width           =   1440
      End
      Begin VB.CommandButton SPellCk 
         Caption         =   "Spelling"
         Height          =   195
         Index           =   0
         Left            =   2085
         TabIndex        =   120
         Top             =   1785
         Width           =   1440
      End
      Begin VB.ListBox approvingofficer 
         Height          =   645
         Left            =   0
         TabIndex        =   66
         Top             =   5025
         Width           =   2415
      End
      Begin VB.ListBox reportingofficer 
         Height          =   645
         Left            =   0
         TabIndex        =   64
         Top             =   4140
         Width           =   2415
      End
      Begin VB.TextBox reportingofficernumber 
         Height          =   285
         Left            =   2520
         TabIndex        =   65
         Top             =   4140
         Width           =   765
      End
      Begin VB.TextBox approvingofficernumber 
         Height          =   285
         Left            =   2520
         TabIndex        =   67
         Top             =   5025
         Width           =   765
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00800000&
         Caption         =   "Form309/SR21 provided?"
         ForeColor       =   &H0000FFFF&
         Height          =   615
         Left            =   0
         TabIndex        =   113
         Top             =   165
         Width           =   2055
         Begin VB.OptionButton yes309 
            BackColor       =   &H00800000&
            Caption         =   "Yes"
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   360
            TabIndex        =   6
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton no309 
            BackColor       =   &H00800000&
            Caption         =   "No"
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   1200
            TabIndex        =   57
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00800000&
         Caption         =   "Personal Injuries?"
         ForeColor       =   &H0000FFFF&
         Height          =   615
         Left            =   2085
         TabIndex        =   112
         Top             =   165
         Width           =   1455
         Begin VB.OptionButton injno 
            BackColor       =   &H00800000&
            Caption         =   "No"
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   720
            TabIndex        =   59
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton injyes 
            BackColor       =   &H00800000&
            Caption         =   "Yes"
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   58
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.TextBox insurance 
         Height          =   285
         Index           =   0
         Left            =   960
         MaxLength       =   50
         TabIndex        =   60
         Top             =   960
         Width           =   2600
      End
      Begin VB.TextBox insurance 
         Height          =   285
         Index           =   1
         Left            =   960
         MaxLength       =   50
         TabIndex        =   61
         Top             =   1440
         Width           =   2600
      End
      Begin RichTextLib.RichTextBox narrative 
         Height          =   1095
         Left            =   0
         TabIndex        =   63
         Top             =   2805
         Width           =   3570
         _ExtentX        =   6297
         _ExtentY        =   1931
         _Version        =   393217
         TextRTF         =   $"noncrim.frx":0000
      End
      Begin RichTextLib.RichTextBox remarks 
         Height          =   615
         Left            =   0
         TabIndex        =   62
         Top             =   1965
         Width           =   3570
         _ExtentX        =   6297
         _ExtentY        =   1085
         _Version        =   393217
         TextRTF         =   $"noncrim.frx":00D5
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "SUPERVISOR                      NUMBER"
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
         Left            =   0
         TabIndex        =   119
         Top             =   4785
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
         Left            =   0
         TabIndex        =   118
         Top             =   3900
         Width           =   3855
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Insurance Company 1"
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
         Index           =   2
         Left            =   0
         TabIndex        =   117
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Insurance Company 2"
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
         Index           =   3
         Left            =   0
         TabIndex        =   116
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "REMARKS:"
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
         Left            =   0
         TabIndex        =   115
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "NARRATIVE:"
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
         Left            =   0
         TabIndex        =   114
         Top             =   2640
         Width           =   2055
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Vehicle Two"
      ForeColor       =   &H00800000&
      Height          =   5715
      Left            =   3975
      TabIndex        =   96
      Top             =   1665
      Width           =   4080
      Begin VB.TextBox OWNERSTATE 
         Height          =   285
         Index           =   1
         Left            =   2715
         MaxLength       =   2
         TabIndex        =   52
         Top             =   4365
         Width           =   390
      End
      Begin VB.TextBox OWNERZIPCODE 
         Height          =   285
         Index           =   1
         Left            =   3195
         MaxLength       =   10
         TabIndex        =   53
         Top             =   4350
         Width           =   765
      End
      Begin VB.TextBox DRIVERSTATE 
         Height          =   285
         Index           =   1
         Left            =   2715
         MaxLength       =   2
         TabIndex        =   45
         Top             =   2595
         Width           =   390
      End
      Begin VB.TextBox DRIVERZIPCODE 
         Height          =   285
         Index           =   1
         Left            =   3195
         MaxLength       =   10
         TabIndex        =   46
         Top             =   2595
         Width           =   765
      End
      Begin VB.TextBox vin 
         Height          =   285
         Index           =   1
         Left            =   1425
         MaxLength       =   20
         TabIndex        =   38
         Top             =   510
         Width           =   2500
      End
      Begin VB.TextBox tagSTATE 
         Height          =   285
         Index           =   1
         Left            =   1425
         MaxLength       =   20
         TabIndex        =   37
         Top             =   165
         Width           =   2500
      End
      Begin VB.TextBox make 
         Height          =   285
         Index           =   1
         Left            =   1425
         MaxLength       =   20
         TabIndex        =   39
         Top             =   870
         Width           =   2500
      End
      Begin VB.TextBox model 
         Height          =   285
         Index           =   1
         Left            =   1425
         MaxLength       =   20
         TabIndex        =   40
         Top             =   1215
         Width           =   2500
      End
      Begin VB.TextBox color 
         Height          =   285
         Index           =   1
         Left            =   1425
         MaxLength       =   20
         TabIndex        =   41
         Top             =   1560
         Width           =   2500
      End
      Begin VB.TextBox driveraddress1 
         Height          =   285
         Index           =   1
         Left            =   1425
         MaxLength       =   50
         TabIndex        =   43
         Top             =   2265
         Width           =   2500
      End
      Begin VB.TextBox driveraddress2 
         Height          =   285
         Index           =   1
         Left            =   390
         MaxLength       =   50
         TabIndex        =   44
         Top             =   2595
         Width           =   2250
      End
      Begin VB.TextBox driverphone 
         Height          =   285
         Index           =   1
         Left            =   1425
         MaxLength       =   50
         TabIndex        =   47
         Top             =   2985
         Width           =   2500
      End
      Begin VB.TextBox driverdl 
         Height          =   285
         Index           =   1
         Left            =   1425
         MaxLength       =   50
         TabIndex        =   48
         Top             =   3315
         Width           =   2500
      End
      Begin VB.TextBox owneraddress1 
         Height          =   285
         Index           =   1
         Left            =   1425
         MaxLength       =   50
         TabIndex        =   50
         Top             =   4020
         Width           =   2500
      End
      Begin VB.TextBox owneraddress2 
         Height          =   285
         Index           =   1
         Left            =   390
         MaxLength       =   50
         TabIndex        =   51
         Top             =   4350
         Width           =   2250
      End
      Begin VB.TextBox ownerphone 
         Height          =   285
         Index           =   1
         Left            =   1425
         MaxLength       =   50
         TabIndex        =   54
         Top             =   4710
         Width           =   2500
      End
      Begin VB.TextBox damaged 
         Height          =   285
         Index           =   1
         Left            =   1425
         MaxLength       =   30
         TabIndex        =   55
         Top             =   5070
         Width           =   2500
      End
      Begin VB.TextBox estcost 
         Height          =   285
         Index           =   1
         Left            =   1425
         MaxLength       =   100
         TabIndex        =   56
         Top             =   5400
         Width           =   2500
      End
      Begin VB.ComboBox driver 
         Height          =   315
         Index           =   1
         Left            =   1425
         TabIndex        =   42
         Top             =   1920
         Width           =   2535
      End
      Begin VB.ComboBox owner 
         Height          =   315
         Index           =   1
         Left            =   1425
         TabIndex        =   49
         Top             =   3660
         Width           =   2535
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "VIN #"
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
         Index           =   1
         Left            =   105
         TabIndex        =   110
         Top             =   585
         Width           =   2055
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Tag State/#"
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
         Index           =   1
         Left            =   105
         TabIndex        =   109
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Make"
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
         Index           =   1
         Left            =   105
         TabIndex        =   108
         Top             =   945
         Width           =   2055
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
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
         Index           =   1
         Left            =   105
         TabIndex        =   107
         Top             =   1290
         Width           =   2055
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Color"
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
         Index           =   1
         Left            =   105
         TabIndex        =   106
         Top             =   1635
         Width           =   2055
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Driver's Name"
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
         Index           =   1
         Left            =   105
         TabIndex        =   105
         Top             =   1995
         Width           =   2055
      End
      Begin VB.Label Label22 
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
         Index           =   1
         Left            =   120
         TabIndex        =   104
         Top             =   2340
         Width           =   2055
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Phone #"
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
         Index           =   1
         Left            =   105
         TabIndex        =   103
         Top             =   3045
         Width           =   2055
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "DL#"
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
         Index           =   1
         Left            =   105
         TabIndex        =   102
         Top             =   3390
         Width           =   2055
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Owner's Name"
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
         Index           =   1
         Left            =   105
         TabIndex        =   101
         Top             =   3735
         Width           =   2055
      End
      Begin VB.Label Label29 
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
         Index           =   1
         Left            =   105
         TabIndex        =   100
         Top             =   4095
         Width           =   2055
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Phone #"
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
         Index           =   1
         Left            =   105
         TabIndex        =   99
         Top             =   4785
         Width           =   2055
      End
      Begin VB.Label Label34 
         BackStyle       =   0  'Transparent
         Caption         =   "Damaged Area"
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
         Index           =   1
         Left            =   105
         TabIndex        =   98
         Top             =   5145
         Width           =   2055
      End
      Begin VB.Label Label35 
         BackStyle       =   0  'Transparent
         Caption         =   "Est. Cost"
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
         Index           =   1
         Left            =   105
         TabIndex        =   97
         Top             =   5490
         Width           =   2055
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Vehicle One"
      ForeColor       =   &H00800000&
      Height          =   5730
      Left            =   0
      TabIndex        =   81
      Top             =   1665
      Width           =   3975
      Begin VB.TextBox OWNERSTATE 
         Height          =   285
         Index           =   0
         Left            =   2640
         MaxLength       =   2
         TabIndex        =   32
         Top             =   4365
         Width           =   390
      End
      Begin VB.TextBox OWNERZIPCODE 
         Height          =   285
         Index           =   0
         Left            =   3135
         MaxLength       =   10
         TabIndex        =   33
         Top             =   4365
         Width           =   765
      End
      Begin VB.TextBox DRIVERSTATE 
         Height          =   285
         Index           =   0
         Left            =   2655
         MaxLength       =   2
         TabIndex        =   25
         Top             =   2595
         Width           =   390
      End
      Begin VB.TextBox DRIVERZIPCODE 
         Height          =   285
         Index           =   0
         Left            =   3135
         MaxLength       =   10
         TabIndex        =   26
         Top             =   2595
         Width           =   765
      End
      Begin VB.TextBox tagSTATE 
         Height          =   285
         Index           =   0
         Left            =   1365
         MaxLength       =   20
         TabIndex        =   17
         Top             =   150
         Width           =   2500
      End
      Begin VB.TextBox vin 
         Height          =   285
         Index           =   0
         Left            =   1365
         MaxLength       =   20
         TabIndex        =   18
         Top             =   495
         Width           =   2500
      End
      Begin VB.TextBox make 
         Height          =   285
         Index           =   0
         Left            =   1365
         MaxLength       =   20
         TabIndex        =   19
         Top             =   855
         Width           =   2500
      End
      Begin VB.TextBox model 
         Height          =   285
         Index           =   0
         Left            =   1365
         MaxLength       =   20
         TabIndex        =   20
         Top             =   1215
         Width           =   2500
      End
      Begin VB.TextBox color 
         Height          =   285
         Index           =   0
         Left            =   1365
         MaxLength       =   20
         TabIndex        =   21
         Top             =   1545
         Width           =   2500
      End
      Begin VB.TextBox driveraddress1 
         Height          =   285
         Index           =   0
         Left            =   1365
         MaxLength       =   50
         TabIndex        =   23
         Top             =   2265
         Width           =   2500
      End
      Begin VB.TextBox driveraddress2 
         Height          =   285
         Index           =   0
         Left            =   345
         MaxLength       =   30
         TabIndex        =   24
         Top             =   2595
         Width           =   2250
      End
      Begin VB.TextBox driverphone 
         Height          =   285
         Index           =   0
         Left            =   1350
         MaxLength       =   50
         TabIndex        =   27
         Top             =   2955
         Width           =   2500
      End
      Begin VB.TextBox driverdl 
         Height          =   285
         Index           =   0
         Left            =   1365
         MaxLength       =   50
         TabIndex        =   28
         Top             =   3315
         Width           =   2500
      End
      Begin VB.TextBox owneraddress1 
         Height          =   285
         Index           =   0
         Left            =   1365
         MaxLength       =   50
         TabIndex        =   30
         Top             =   4005
         Width           =   2500
      End
      Begin VB.TextBox owneraddress2 
         Height          =   285
         Index           =   0
         Left            =   315
         MaxLength       =   50
         TabIndex        =   31
         Top             =   4365
         Width           =   2250
      End
      Begin VB.TextBox ownerphone 
         Height          =   285
         Index           =   0
         Left            =   1365
         MaxLength       =   50
         TabIndex        =   34
         Top             =   4695
         Width           =   2500
      End
      Begin VB.TextBox damaged 
         Height          =   285
         Index           =   0
         Left            =   1350
         MaxLength       =   30
         TabIndex        =   35
         Top             =   5055
         Width           =   2500
      End
      Begin VB.TextBox estcost 
         Height          =   285
         Index           =   0
         Left            =   1365
         MaxLength       =   100
         TabIndex        =   36
         Top             =   5400
         Width           =   2500
      End
      Begin VB.ComboBox driver 
         Height          =   315
         Index           =   0
         Left            =   1365
         TabIndex        =   22
         Top             =   1905
         Width           =   2535
      End
      Begin VB.ComboBox owner 
         Height          =   315
         Index           =   0
         Left            =   1365
         TabIndex        =   29
         Top             =   3660
         Width           =   2535
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Tag State/#"
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
         Index           =   0
         Left            =   45
         TabIndex        =   95
         Top             =   210
         Width           =   1575
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "VIN #"
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
         Index           =   0
         Left            =   45
         TabIndex        =   94
         Top             =   555
         Width           =   2055
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Make"
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
         Index           =   0
         Left            =   45
         TabIndex        =   93
         Top             =   915
         Width           =   2055
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
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
         Index           =   0
         Left            =   45
         TabIndex        =   92
         Top             =   1260
         Width           =   2055
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Color"
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
         Index           =   0
         Left            =   45
         TabIndex        =   91
         Top             =   1605
         Width           =   2055
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Driver's Name"
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
         Index           =   0
         Left            =   45
         TabIndex        =   90
         Top             =   1965
         Width           =   2055
      End
      Begin VB.Label Label22 
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
         Index           =   0
         Left            =   45
         TabIndex        =   89
         Top             =   2310
         Width           =   2055
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Phone #"
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
         Index           =   0
         Left            =   45
         TabIndex        =   88
         Top             =   3015
         Width           =   2055
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "DL#"
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
         Index           =   0
         Left            =   45
         TabIndex        =   87
         Top             =   3360
         Width           =   2055
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Owner's Name"
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
         Index           =   0
         Left            =   45
         TabIndex        =   86
         Top             =   3705
         Width           =   2055
      End
      Begin VB.Label Label29 
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
         Index           =   0
         Left            =   45
         TabIndex        =   85
         Top             =   4065
         Width           =   2055
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Phone #"
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
         Index           =   0
         Left            =   45
         TabIndex        =   84
         Top             =   4755
         Width           =   2055
      End
      Begin VB.Label Label34 
         BackStyle       =   0  'Transparent
         Caption         =   "Damaged Area"
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
         Index           =   0
         Left            =   45
         TabIndex        =   83
         Top             =   5115
         Width           =   2055
      End
      Begin VB.Label Label35 
         BackStyle       =   0  'Transparent
         Caption         =   "Est. Cost"
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
         Index           =   0
         Left            =   45
         TabIndex        =   82
         Top             =   5460
         Width           =   2055
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1260
      Left            =   0
      TabIndex        =   69
      Top             =   420
      Width           =   11625
      Begin VB.TextBox incidentdate 
         Height          =   285
         Left            =   4575
         MaxLength       =   10
         TabIndex        =   1
         Top             =   120
         Width           =   1095
      End
      Begin VB.TextBox caddress 
         Height          =   285
         Left            =   4695
         MaxLength       =   30
         TabIndex        =   12
         Top             =   930
         Width           =   2085
      End
      Begin VB.ComboBox cname 
         Height          =   315
         Left            =   1230
         TabIndex        =   11
         Top             =   915
         Width           =   2640
      End
      Begin VB.TextBox cphone 
         Height          =   285
         Left            =   10350
         MaxLength       =   15
         TabIndex        =   16
         Top             =   915
         Width           =   1260
      End
      Begin VB.TextBox received 
         Height          =   285
         Left            =   5655
         MaxLength       =   10
         TabIndex        =   5
         Top             =   600
         Width           =   975
      End
      Begin VB.CheckBox privprop 
         BackColor       =   &H00800000&
         Caption         =   "Private Property"
         ForeColor       =   &H0080FFFF&
         Height          =   255
         Left            =   8055
         TabIndex        =   2
         Top             =   120
         Width           =   1455
      End
      Begin VB.CheckBox lessthan 
         BackColor       =   &H00800000&
         Caption         =   "Less than $1,000.00"
         ForeColor       =   &H0080FFFF&
         Height          =   255
         Left            =   9735
         TabIndex        =   3
         Top             =   120
         Width           =   1935
      End
      Begin VB.TextBox incidentlocation 
         Height          =   285
         Left            =   15
         MaxLength       =   70
         TabIndex        =   4
         Top             =   600
         Width           =   4935
      End
      Begin VB.ComboBox casenumber 
         Height          =   315
         Left            =   1455
         TabIndex        =   0
         Top             =   120
         Width           =   2415
      End
      Begin VB.TextBox consumed 
         Height          =   285
         Left            =   7935
         MaxLength       =   10
         TabIndex        =   8
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox completed 
         Height          =   285
         Left            =   6735
         MaxLength       =   10
         TabIndex        =   7
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox units 
         Height          =   285
         Left            =   9135
         MaxLength       =   20
         TabIndex        =   9
         Top             =   615
         Width           =   1095
      End
      Begin VB.TextBox zone 
         Height          =   285
         Left            =   10335
         MaxLength       =   10
         TabIndex        =   10
         Top             =   600
         Width           =   1275
      End
      Begin VB.TextBox CCITY 
         Height          =   285
         Left            =   6855
         MaxLength       =   30
         TabIndex        =   13
         Top             =   915
         Width           =   1575
      End
      Begin VB.TextBox CSTATE 
         Height          =   285
         Left            =   8490
         MaxLength       =   2
         TabIndex        =   14
         Top             =   915
         Width           =   390
      End
      Begin VB.TextBox CZIPCODE 
         Height          =   285
         Left            =   8940
         MaxLength       =   10
         TabIndex        =   15
         Top             =   900
         Width           =   765
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Completed     Consumed"
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
         Left            =   6735
         TabIndex        =   80
         Top             =   390
         Width           =   2415
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
         Left            =   3870
         TabIndex        =   79
         Top             =   915
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
         Left            =   9765
         TabIndex        =   78
         Top             =   915
         Width           =   975
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Time:   Received"
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
         Left            =   5055
         TabIndex        =   77
         Top             =   390
         Width           =   1575
      End
      Begin VB.Label label44 
         BackStyle       =   0  'Transparent
         Caption         =   "Complainant"
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
         Left            =   15
         TabIndex        =   76
         Top             =   915
         Width           =   2055
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   "TRAFFIC ACCIDENT"
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
         Left            =   6135
         TabIndex        =   75
         Top             =   120
         Width           =   2400
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
         Left            =   15
         TabIndex        =   74
         Top             =   390
         Width           =   2055
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
         Left            =   15
         TabIndex        =   73
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label2 
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
         Left            =   3975
         TabIndex        =   72
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "Units Resp"
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
         Left            =   9135
         TabIndex        =   71
         Top             =   390
         Width           =   1695
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   "Zone#"
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
         Left            =   10335
         TabIndex        =   70
         Top             =   390
         Width           =   1695
      End
   End
   Begin Crystal.CrystalReport REPORT 
      Left            =   11400
      Top             =   7320
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
      Width           =   11655
      _ExtentX        =   20558
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
            Object.ToolTipText     =   "Print Service Call Report"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit   "
            Object.ToolTipText     =   "Exit Service Call Repor"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   11400
      Top             =   7200
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
            Picture         =   "noncrim.frx":01AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "noncrim.frx":05FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "noncrim.frx":0A52
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "noncrim.frx":0EA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "noncrim.frx":12FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "noncrim.frx":174E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "noncrim.frx":1BA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "noncrim.frx":1FF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "noncrim.frx":244A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "noncrim.frx":289E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "noncrim.frx":2CF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "noncrim.frx":3146
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "noncrim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FROMXREF As Integer, nametype As Integer

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

Private Sub caddress_GotFocus()
Dim db As Database, rs As Recordset
On Error Resume Next
If caddress = "" And cname > "" Then
    On Error GoTo oderror
od:
    Set db = OpenDatabase(nwl + "LAWSUITE.MDB")
    Set rs = db.OpenRecordset("SELECT * FROM PEOPLE WHERE DPNAMElf = " + Chr$(34) + cname + Chr$(34))
    If Not rs.EOF Then
        rs.MoveFirst
        caddress = rs("DPHADDRESS")
        If Not IsNull(rs("DPHADDRESS2")) Then
            CCITY = rs("DPHADDRESS2")
        End If
        If Not IsNull(rs("hstate")) Then
            CSTATE = rs("hstate")
        End If
        If Not IsNull(rs("hzipcode")) Then
            CZIPCODE = rs("hzipcode")
        End If
        If cphone = "" Then
            cphone = rs("DPHPHONE")
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

Private Sub casenumber_LostFocus()
If casenumber = "SAVE" Then
    For t% = 0 To casenumber.ListCount - 1
        casenumber = casenumber.List(t%)
        Call findrtn
        Me.Refresh
        Call savertn
    Next t%
End If
End Sub

Private Sub cname_Click()
If cname = "" Then
    Exit Sub
End If
Call setpopup(cname, "L")
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

Private Sub cname_LostFocus()
If Len(cname) > 50 Then
    cname = Left$(cname, 50)
End If
If cname > "" And InStr(cname, ",") = 0 Then
    msg = MsgBox("All names in the Non-Criminal Police Response system should be entered in the format last name + comma + firstname.", 48, "Invalid Data Format")
    cname.SetFocus
End If
End Sub

Private Sub driver_Click(index As Integer)
If driver(index) = "" Then
    Exit Sub
End If
Call setpopup(driver(index), "L")
End Sub

Private Sub driver_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)
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

Private Sub driver_LostFocus(index As Integer)
If Len(driver(index)) > 50 Then
    driver(index) = Left$(driver(index), 50)
End If
If driver(index) > "" And InStr(driver(index), ",") = 0 Then
    msg = MsgBox("All names in the Non-Criminal Police Response system should be entered in the format last name + comma + firstname.", 48, "Invalid Data Format")
    driver(index).SetFocus
End If

End Sub

Private Sub driveraddress1_GotFocus(index As Integer)
Dim db As Database, rs As Recordset
On Error Resume Next
If driveraddress1(index) = "" And driver(index) > "" Then
    On Error GoTo oderror
od:
    Set db = OpenDatabase(nwl + "LAWSUITE.MDB")
    Set rs = db.OpenRecordset("SELECT * FROM PEOPLE WHERE DPNAMElf = " + Chr$(34) + driver(index) + Chr$(34))
    If Not rs.EOF Then
        rs.MoveFirst
        driveraddress1(index) = rs("DPHADDRESS")
        driveraddress2(index) = rs("DPHADDRESS2")
        If Not IsNull(rs("hstate")) Then
            DRIVERSTATE(index) = rs("hstate")
        End If
        If Not IsNull(rs("hzipcode")) Then
            DRIVERZIPCODE(index) = rs("hzipcode")
        End If
        If driverphone(index) = "" Then
            driverphone(index) = rs("DPHPHONE")
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

Private Sub Form_Load()
nametype = 1
Me.Top = 0
Me.Left = 0
Me.Height = 7850
Me.Width = 11750
For t% = 0 To Forms.Count - 1
    If Forms(t%).Name = "xref" Then
        FROMXREF = 1
        t% = Forms.Count - 1
    End If
Next t%
On Error Resume Next
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
Set rs = db.OpenRecordset("select CASEnumber from NONCRIMINAL order by CASEnumber")
If Not rs.EOF Then
    rs.MoveFirst
End If
casenumber.clear
While Not rs.EOF
    casenumber.AddItem rs("casenumber")
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
Private Sub clearrtn()
casenumber = ""
cname = ""
caddress = ""
CCITY = ""
CSTATE = ""
CZIPCODE = ""
cphone = ""
IncidentDate = ""
privprop = 0
lessthan = 0
incidentlocation = ""
received = ""
completed = ""
comsumed = ""
units = ""
zone = ""
For t% = 0 To 1
    tagSTATE(t%) = ""
    VIN(t%) = ""
    make(t%) = ""
    model(t%) = ""
    color(t%) = ""
    driver(t%) = ""
    driveraddress1(t%) = ""
    driveraddress2(t%) = ""
    DRIVERSTATE(t%) = ""
    DRIVERZIPCODE(t%) = ""
    driverphone(t%) = ""
    driverdl(t%) = ""
    owner(t%) = ""
    owneraddress1(t%) = ""
    owneraddress2(t%) = ""
    OWNERSTATE(t%) = ""
    OWNERZIPCODE(t%) = ""
    ownerphone(t%) = ""
    damaged(t%) = ""
    estcost(t%) = ""
    insurance(t%) = ""
Next t%
yes309 = True
no309 = False
injyes = False
injno = True
REMARKS.Text = ""
narrative.Text = ""
reportingofficer.ListIndex = -1
reportingofficernumber = ""
approvingofficer.ListIndex = -1
approvingofficernumber = ""
Call loadcase
Call loadofficer
Call loadname

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
driver(0).clear
driver(1).clear
owner(0).clear
owner(1).clear
While Not rs.EOF
    cname.AddItem rs("DPnamelf")
    For t% = 0 To 1
        driver(t%).AddItem rs("DPnamelf")
        owner(t%).AddItem rs("DPnamelf")
    Next t%
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
Set noncrim = Nothing
End Sub

Private Sub incidentdate_GotFocus()
If IncidentDate = "" Then
    IncidentDate = Format$(Date$, "mm/dd/yyyy")
End If
End Sub

Private Sub incidentdate_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
If Len(IncidentDate) = 1 Or Len(IncidentDate) = 4 Then
    SendKeys "/"
End If
End If


End Sub

Private Sub NARRATIVE_KeyDown(KeyCode As Integer, Shift As Integer)
If (Shift = vbCtrlMask) And (KeyCode = vbKeyF2) Then
    Call SpellCk_Click(2)
End If
End Sub

Private Sub owner_Click(index As Integer)
If owner(index) = "" Then
    Exit Sub
End If
Call setpopup(owner(index), "L")
End Sub

Private Sub owner_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)
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

Private Sub owner_LostFocus(index As Integer)
If Len(owner(index)) > 50 Then
    owner(index) = Left$(owner(index), 50)
End If
If owner(index) > "" And InStr(owner(index), ",") = 0 Then
    msg = MsgBox("All names in the Non-Criminal Police Response system should be entered in the format last name + comma + firstname.", 48, "Invalid Data Format")
    owner(index).SetFocus
End If

End Sub

Private Sub owneraddress1_GotFocus(index As Integer)
Dim db As Database, rs As Recordset
On Error Resume Next
If owneraddress1(index) = "" And owner(index) > "" Then
    On Error GoTo oderror
od:
    Set db = OpenDatabase(nwl + "LAWSUITE.MDB")
    Set rs = db.OpenRecordset("SELECT * FROM PEOPLE WHERE DPNAMElf = " + Chr$(34) + owner(index) + Chr$(34))
    If Not rs.EOF Then
        rs.MoveFirst
        owneraddress1(index) = rs("DPHADDRESS")
        owneraddress2(index) = rs("DPHADDRESS2")
        If Not IsNull(rs("hstate")) Then
            OWNERSTATE(index) = rs("hstate")
        End If
        If Not IsNull(rs("hzipcode")) Then
            OWNERZIPCODE(index) = rs("hzipcode")
        End If
        If ownerphone(index) = "" Then
            ownerphone(index) = rs("DPHPHONE")
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

Private Sub remarks_KeyDown(KeyCode As Integer, Shift As Integer)
If (Shift = vbCtrlMask) And (KeyCode = vbKeyF2) Then
        Call SpellCk_Click(0)
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

Private Sub SpellCk_Click(index As Integer)
If index = 0 Then BeginSpellCheck REMARKS.Text, REMARKS
If index = 1 Then BeginSpellCheck narrative.Text, narrative

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
If Not IsDate(received) Then
    msg = MsgBox("Received must be entered and must be a valid time.", 48, "Genesis Error Log")
    Exit Sub
End If
If Not IsDate(completed) Then
    msg = MsgBox("Completed must be entered and must be a valid time.", 48, "Genesis Error Log")
    Exit Sub
End If
On Error GoTo oderror
od:
Set db = OpenDatabase(nwi + "incident.mdb")
Set rs = db.OpenRecordset("select * from noncriminal where casenumber =" + Chr$(34) + casenumber + Chr$(34))
If Not rs.EOF Then
    rs.MoveFirst
    rs.Edit
Else
    rs.AddNew
End If
rs("casenumber") = casenumber
rs("cname") = cname
rs("caddress") = caddress
rs("ccity") = CCITY
rs("cstate") = CSTATE
rs("czipcode") = CZIPCODE
rs("cphone") = cphone
rs("incidentdate") = IncidentDate
rs("received") = received
rs("consumed") = consumed
rs("completed") = completed
rs("incidentlocation") = incidentlocation
rs("remarks") = REMARKS.Text
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
rs("privprop") = privprop
rs("lessthan") = lessthan
rs("units") = Val(units)
rs("zone") = zone
For t% = 0 To 1
    rs("TAGSTATE" + Mid$(Str$(t%), 2)) = tagSTATE(t%)
    rs("vin" + Mid$(Str$(t%), 2)) = VIN(t%)
    rs("make" + Mid$(Str$(t%), 2)) = make(t%)
    rs("model" + Mid$(Str$(t%), 2)) = model(t%)
    rs("color" + Mid$(Str$(t%), 2)) = color(t%)
    rs("driver" + Mid$(Str$(t%), 2)) = driver(t%)
    rs("driver" + Mid$(Str$(t%), 2) + "address1") = driveraddress1(t%)
    rs("driver" + Mid$(Str$(t%), 2) + "address2") = driveraddress2(t%)
    rs("driver" + Mid$(Str$(t%), 2) + "state") = DRIVERSTATE(t%)
    rs("driver" + Mid$(Str$(t%), 2) + "zipcode") = DRIVERZIPCODE(t%)
    rs("driver" + Mid$(Str$(t%), 2) + "phone") = driverphone(t%)
    rs("driver" + Mid$(Str$(t%), 2) + "dl") = driverdl(t%)
    rs("owner" + Mid$(Str$(t%), 2)) = owner(t%)
    rs("owner" + Mid$(Str$(t%), 2) + "address1") = owneraddress1(t%)
    rs("owner" + Mid$(Str$(t%), 2) + "address2") = owneraddress2(t%)
    rs("owner" + Mid$(Str$(t%), 2) + "state") = OWNERSTATE(t%)
    rs("owner" + Mid$(Str$(t%), 2) + "zipcode") = OWNERZIPCODE(t%)
    rs("owner" + Mid$(Str$(t%), 2) + "phone") = ownerphone(t%)
    rs("damaged" + Mid$(Str$(t%), 2)) = damaged(t%)
    rs("estcost" + Mid$(Str$(t%), 2)) = Val(estcost(t%))
    rs("insurance" + Mid$(Str$(t%), 2)) = insurance(t%)
Next t%
rs("yes309") = yes309
rs("no309") = no309
rs("injyes") = injyes
rs("injno") = injno
rs("remarks") = REMARKS.Text
rs("narrative") = narrative.Text
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
        followupofficer.AddItem reportingofficer
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
        followupofficer.AddItem approvingofficer
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

For yy% = 0 To 1
    Set rs = db.OpenRecordset("select * from people where dpnamelf = " + Chr$(34) + driver(yy%) + Chr$(34))
    If Not rs.EOF Then
        rs.MoveFirst
        rs.Edit
    Else
        rs.AddNew
    End If
    rs("dpnamelf") = driver(yy%)
    rs("dphaddress") = driveraddress1(yy%)
    rs("dphaddress2") = driveraddress2(yy%)
    rs("dstate") = DRIVERSTATE(yy%)
    rs("dzipcode") = DRIVERZIPCODE(yy%)
    rs("dpsort") = Left$(driver(yy%), 15)
    rs("dphphone") = driverphone(yy%)
    hoLdname = driver(yy%)
    rs("dl") = driverdl(yy%)
    
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
Next yy%

For yy% = 0 To 1
    Set rs = db.OpenRecordset("select * from people where dpnamelf = " + Chr$(34) + owner(yy%) + Chr$(34))
    If Not rs.EOF Then
        rs.MoveFirst
        rs.Edit
    Else
        rs.AddNew
    End If
    rs("dpnamelf") = owner(yy%)
    rs("dphaddress") = owneraddress1(yy%)
    rs("dphaddress2") = owneraddress2(yy%)
    rs("dstate") = OWNERSTATE(yy%)
    rs("dzipcode") = OWNERZIPCODE(yy%)
    rs("dpsort") = Left$(owner(yy%), 15)
    rs("dphphone") = ownerphone(yy%)
    hoLdname = owner(yy%)
    
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
Next yy%

db.Close
Call clearrtn
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
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
Set rs = db.OpenRecordset("select * from noncriminal where casenumber =" + Chr$(34) + casenumber + Chr$(34))
If Not rs.EOF Then
    rs.MoveFirst
Else
    On Error Resume Next
    db.Close
    msg = MsgBox("Case Number not found.", 48, "Genesis Information Log")
    Exit Sub
End If
casenumber = rs("casenumber")
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
received = rs("received")
consumed = rs("consumed")
completed = rs("completed")
incidentlocation = rs("incidentlocation")
REMARKS.Text = rs("remarks")
privprop = rs("privprop")
lessthan = rs("lessthan")
units = rs("units")
zone = rs("zone")
For t% = 0 To 1
    tagSTATE(t%) = rs("TAGSTATE" + Mid$(Str$(t%), 2))
    VIN(t%) = rs("vin" + Mid$(Str$(t%), 2))
    make(t%) = rs("make" + Mid$(Str$(t%), 2))
    model(t%) = rs("model" + Mid$(Str$(t%), 2))
    color(t%) = rs("color" + Mid$(Str$(t%), 2))
    driver(t%) = rs("driver" + Mid$(Str$(t%), 2))
    driveraddress1(t%) = rs("driver" + Mid$(Str$(t%), 2) + "address1")
    driveraddress2(t%) = rs("driver" + Mid$(Str$(t%), 2) + "address2")
    If Not IsNull(rs("driver" + Mid$(Str$(t%), 2) + "state")) Then
        DRIVERSTATE(t%) = rs("driver" + Mid$(Str$(t%), 2) + "state")
    End If
    If Not IsNull(rs("driver" + Mid$(Str$(t%), 2) + "zipcode")) Then
        DRIVERZIPCODE(t%) = rs("driver" + Mid$(Str$(t%), 2) + "zipcode")
    End If
    driverphone(t%) = rs("driver" + Mid$(Str$(t%), 2) + "phone")
    driverdl(t%) = rs("driver" + Mid$(Str$(t%), 2) + "dl")
    owner(t%) = rs("owner" + Mid$(Str$(t%), 2))
    owneraddress1(t%) = rs("owner" + Mid$(Str$(t%), 2) + "address1")
    owneraddress2(t%) = rs("owner" + Mid$(Str$(t%), 2) + "address2")
    If Not IsNull(rs("owner" + Mid$(Str$(t%), 2) + "state")) Then
        OWNERSTATE(t%) = rs("owner" + Mid$(Str$(t%), 2) + "state")
    End If
    If Not IsNull(rs("owner" + Mid$(Str$(t%), 2) + "zipcode")) Then
        OWNERZIPCODE(t%) = rs("owner" + Mid$(Str$(t%), 2) + "zipcode")
    End If
    ownerphone(t%) = rs("owner" + Mid$(Str$(t%), 2) + "phone")
    damaged(t%) = rs("damaged" + Mid$(Str$(t%), 2))
    estcost(t%) = rs("estcost" + Mid$(Str$(t%), 2))
    insurance(t%) = rs("insurance" + Mid$(Str$(t%), 2))
Next t%
yes309 = rs("yes309")
no309 = rs("no309")
injyes = rs("injyes")
injno = rs("injno")
REMARKS.Text = rs("remarks")
narrative.Text = rs("narrative")
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
Set rs = db.OpenRecordset("select * from noncriminal where casenumber =" + Chr$(34) + casenumber + Chr$(34))
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
REPORT.ReportFileName = nwi + "noncrim.RPT"
REPORT.SelectionFormula = "{noncriminal.caseNUMBER} = '" + casenumber + "'"
REPORT.Action = 1
End Sub
