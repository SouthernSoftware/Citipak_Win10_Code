VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Search 
   BackColor       =   &H00808000&
   Caption         =   "Search Incident Database"
   ClientHeight    =   7110
   ClientLeft      =   585
   ClientTop       =   1215
   ClientWidth     =   10710
   LinkTopic       =   "Form1"
   ScaleHeight     =   7110
   ScaleWidth      =   10710
   Begin Crystal.CrystalReport reportg 
      Left            =   240
      Top             =   6105
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame findlistframe 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   240
      TabIndex        =   55
      Top             =   600
      Visible         =   0   'False
      Width           =   10475
      Begin MSComctlLib.ListView findlist 
         Height          =   4575
         Left            =   120
         TabIndex        =   57
         Top             =   120
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   8070
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Incident#"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Offense Date"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Offense 1"
            Object.Width           =   4762
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Victim 1"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Subject 1"
            Object.Width           =   4410
         EndProperty
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4800
         TabIndex        =   56
         Top             =   4725
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Clear"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3000
      TabIndex        =   32
      Top             =   6000
      Width           =   4815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Generate Report"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   34
      Top             =   6720
      Width           =   10455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Search"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1440
      TabIndex        =   33
      Top             =   6360
      Width           =   7815
   End
   Begin MSComctlLib.ListView ucr 
      Height          =   975
      Left            =   7800
      TabIndex        =   31
      Top             =   4800
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1720
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   5080
      EndProperty
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   3720
      TabIndex        =   53
      Top             =   5040
      Width           =   3375
      Begin VB.OptionButton wns 
         BackColor       =   &H00808000&
         Caption         =   "Similar To"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   2280
         TabIndex        =   30
         Top             =   120
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton wne 
         BackColor       =   &H00808000&
         Caption         =   "Exact Match"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   29
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.TextBox wn 
      Height          =   285
      Left            =   3720
      TabIndex        =   28
      Top             =   4755
      Width           =   3375
   End
   Begin MSComctlLib.ListView bm 
      Height          =   975
      Left            =   240
      TabIndex        =   27
      Top             =   4800
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1720
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   5080
      EndProperty
   End
   Begin MSComctlLib.ListView ro 
      Height          =   975
      Left            =   7800
      TabIndex        =   26
      Top             =   3360
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1720
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   4939
      EndProperty
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00808000&
      Caption         =   "Disposition"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   975
      Left            =   4080
      TabIndex        =   49
      Top             =   3120
      Width           =   3375
      Begin VB.CheckBox cex 
         BackColor       =   &H00808000&
         Caption         =   "Exceptionally Cleared"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   25
         Top             =   550
         Width           =   1815
      End
      Begin VB.CheckBox dun 
         BackColor       =   &H00808000&
         Caption         =   "Unfounded"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   550
         Width           =   1095
      End
      Begin VB.CheckBox dcl 
         BackColor       =   &H00808000&
         Caption         =   "Adm. Closed"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   1440
         TabIndex        =   23
         Top             =   240
         Width           =   1335
      End
      Begin VB.CheckBox dac 
         BackColor       =   &H00808000&
         Caption         =   "Active"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   255
      TabIndex        =   48
      Top             =   4125
      Width           =   3375
      Begin VB.OptionButton pe 
         BackColor       =   &H00808000&
         Caption         =   "Exact Match"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   15
         TabIndex        =   20
         Top             =   60
         Width           =   1335
      End
      Begin VB.OptionButton ps 
         BackColor       =   &H00808000&
         Caption         =   "Similar To"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   2295
         TabIndex        =   21
         Top             =   60
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.TextBox p 
      Height          =   285
      Left            =   105
      TabIndex        =   19
      Top             =   3855
      Width           =   3585
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   7200
      TabIndex        =   46
      Top             =   2280
      Width           =   3375
      Begin VB.OptionButton sne 
         BackColor       =   &H00808000&
         Caption         =   "Exact Match"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   16
         Top             =   120
         Width           =   1335
      End
      Begin VB.OptionButton sns 
         BackColor       =   &H00808000&
         Caption         =   "Similar To"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   2280
         TabIndex        =   17
         Top             =   120
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.TextBox sn 
      Height          =   285
      Left            =   7200
      TabIndex        =   15
      Top             =   1995
      Width           =   3375
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   3720
      TabIndex        =   44
      Top             =   2280
      Width           =   3375
      Begin VB.OptionButton vne 
         BackColor       =   &H00808000&
         Caption         =   "Exact Match"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   13
         Top             =   120
         Width           =   1335
      End
      Begin VB.OptionButton vns 
         BackColor       =   &H00808000&
         Caption         =   "Similar To"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   2280
         TabIndex        =   14
         Top             =   120
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.TextBox vn 
      Height          =   285
      Left            =   3720
      TabIndex        =   12
      Top             =   1995
      Width           =   3375
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   4440
      TabIndex        =   41
      Top             =   840
      Width           =   3375
      Begin VB.OptionButton ile 
         BackColor       =   &H00808000&
         Caption         =   "Exact Match"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   6
         Top             =   120
         Width           =   1335
      End
      Begin VB.OptionButton ils 
         BackColor       =   &H00808000&
         Caption         =   "Similar To"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   2280
         TabIndex        =   7
         Top             =   120
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.TextBox il 
      Height          =   285
      Left            =   4440
      TabIndex        =   5
      Top             =   555
      Width           =   3375
   End
   Begin VB.TextBox idr2 
      Height          =   285
      Left            =   2520
      TabIndex        =   4
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox idr1 
      Height          =   285
      Left            =   2520
      TabIndex        =   3
      Top             =   555
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   240
      TabIndex        =   37
      Top             =   2280
      Width           =   3375
      Begin VB.OptionButton cns 
         BackColor       =   &H00808000&
         Caption         =   "Similar To"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   2280
         TabIndex        =   11
         Top             =   120
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton cne 
         BackColor       =   &H00808000&
         Caption         =   "Exact Match"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   10
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.TextBox cn 
      Height          =   285
      Left            =   240
      TabIndex        =   9
      Top             =   1995
      Width           =   3375
   End
   Begin VB.TextBox inr2 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox inr1 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   550
      Width           =   1695
   End
   Begin MSComctlLib.ListView MAJORMINOR 
      Height          =   540
      Left            =   90
      TabIndex        =   18
      Top             =   3270
      Width           =   3585
      _ExtentX        =   6324
      _ExtentY        =   953
      View            =   3
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
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
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   5733
      EndProperty
   End
   Begin MSComctlLib.ListView o 
      Height          =   1005
      Left            =   8040
      TabIndex        =   8
      Top             =   480
      Width           =   2610
      _ExtentX        =   4604
      _ExtentY        =   1773
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
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
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   5080
      EndProperty
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   6
      X1              =   7440
      X2              =   7440
      Y1              =   4440
      Y2              =   5880
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   5
      X1              =   3360
      X2              =   3360
      Y1              =   4440
      Y2              =   5880
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   4
      X1              =   7680
      X2              =   7680
      Y1              =   3000
      Y2              =   4440
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   3
      X1              =   3840
      X2              =   3840
      Y1              =   3000
      Y2              =   4440
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   2
      X1              =   7150
      X2              =   7150
      Y1              =   1560
      Y2              =   3000
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   1
      X1              =   3675
      X2              =   3675
      Y1              =   1560
      Y2              =   3000
   End
   Begin VB.Line Line3 
      BorderColor     =   &H0000FFFF&
      X1              =   7920
      X2              =   7920
      Y1              =   120
      Y2              =   1560
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0000FFFF&
      X1              =   4320
      X2              =   4320
      Y1              =   120
      Y2              =   1560
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   0
      X1              =   2280
      X2              =   2280
      Y1              =   120
      Y2              =   1560
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H0000FFFF&
      Height          =   1455
      Left            =   -15
      Top             =   4455
      Width           =   10695
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0000FFFF&
      Height          =   1455
      Left            =   15
      Top             =   3000
      Width           =   10695
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000FFFF&
      Height          =   1455
      Left            =   0
      Top             =   1560
      Width           =   10695
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FFFF&
      Height          =   1455
      Left            =   15
      Top             =   120
      Width           =   10695
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "UCR Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   7800
      TabIndex        =   54
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Witness Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3720
      TabIndex        =   52
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Bias Motivation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   51
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Reporting Officer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   7800
      TabIndex        =   50
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Property"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   15
      TabIndex        =   47
      Top             =   3015
      Width           =   1815
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Subject's Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   7200
      TabIndex        =   45
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Victim's Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3720
      TabIndex        =   43
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Offense"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   8040
      TabIndex        =   42
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label6 
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
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   4440
      TabIndex        =   40
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "to"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3240
      TabIndex        =   39
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Incident Date Range"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2400
      TabIndex        =   38
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Complainant's Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   36
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "to"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   960
      TabIndex        =   35
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Incident Number Range"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "Search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim qe, qv, qs, qp, qa, qw As String

Private Sub Command1_Click()
Dim itmx As ListItem
QC = "not incidentnumber is null "
qv = "" '"not incidentnumber is null "
qs = "" '"not incidentnumber is null "
qo = "" '"not incidentnumber is null "
qisu = "" '"not incidentnumber is null "
qsu = "" '"not incidentnumber is null "
qsus = "" '"not incidentnumber is null "
If (inr1 > "" And inr2 = "") Or (inr1 = "" And inr2 > "") Then
    msg = MsgBox("Both fields for Incident Number Range must be entered.", 48, "Genesis Error Log")
    Exit Sub
End If
If (idr1 > "" And idr2 = "") Or (idr1 = "" And idr2 > "") Then
    msg = MsgBox("Both fields for Incident Date Range must be entered.", 48, "Genesis Error Log")
    Exit Sub
End If
If (idr1 > "" And Not IsDate(idr1)) Or (idr2 > "" And Not IsDate(idr2)) Then
    msg = MsgBox("Incident Date Range is invalid.", 48, "Genesis Error Log")
    Exit Sub
End If
If inr1 > "" Then
    QC = QC + " AND incidentnumber >= '" + inr1 + "' and incidentnumber <= " + Chr$(34) + inr2 + Chr$(34)
End If
If idr1 > "" Then
    QC = QC + " and dateofoffense1 between #" + idr1 + "# and #" + idr2 + "#"
End If
If il > "" Then
    If ile Then
        QC = QC + " and incidentlocation = " + Chr$(34) + il + Chr$(34)
    Else
        QC = QC + " and incidentlocation like '*" + il + "*'"
    End If
End If
If Not (o.SelectedItem Is Nothing) Then
    For yy% = 1 To o.ListItems.Count
        If o.ListItems(yy%).Selected Then
            QC = QC + " and (offense1 = '" + o.ListItems(yy%) + "' or offense2 = '" + o.ListItems(yy%) + "' or offense3 = '" + o.ListItems(yy%) + "')"
        End If
    Next yy%
End If
If cn > "" Then
    If cne Then
        QC = QC + " and (cname = " + cn + " or incidentnumber in (select incidentnumber from supplemental where  ((complainant1 = 1 and name1 = " + cn + ") or (complainant2 = 1 and name2 = " + cn + "))))"
    Else
        QC = QC + " and (cname like '*" + cn + "*' or incidentnumber in (select incidentnumber from supplemental where ((complainant1 = 1 and name1 like '*" + cn + "*') or (complainant2 = 1 and name2 like '*" + cn + "*'))))"
    End If
End If
If sn > "" Then
    If sne Then
        If qs > "" Then
            qs = qs + " and (sname = " + sn + " or incidentnumber in (select incidentnumber from supplemental where  ((subject1 = 1 and name1 = " + sn + ") or (subject2 = 1 and name2 = " + sn + "))))"
        Else
            qs = "(sname = " + sn + " or incidentnumber in (select incidentnumber from supplemental where  ((subject1 = 1 and name1 = " + sn + ") or (subject2 = 1 and name2 = " + sn + "))))"
        End If
    Else
        If qs > "" Then
            qs = qs + " and (sname like '*" + sn + "*' or incidentnumber in (select incidentnumber from supplemental where ((subject1 = 1 and name1 like '*" + sn + "*') or (subject2 = 1 and name2 like '*" + sn + "*'))))"
        Else
            qs = "(sname like '*" + sn + "*' or incidentnumber in (select incidentnumber from supplemental where ((subject1 = 1 and name1 like '*" + sn + "*') or (subject2 = 1 and name2 like '*" + sn + "*'))))"
        End If
    End If
End If
If vn > "" Then
    If vne Then
        If qv > "" Then
            qv = qv + " and (vname = " + vn + " or incidentnumber in (select incidentnumber from supplemental where  ((victim1 = 1 and name1 = " + vn + ") or (victim2 = 1 and name2 = " + vn + "))))"
        Else
            qv = "(vname = " + vn + " or incidentnumber in (select incidentnumber from supplemental where  ((victim1 = 1 and name1 = " + vn + ") or (victim2 = 1 and name2 = " + vn + "))))"
        End If
    Else
        If qv > "" Then
            qv = qv + " and (vname like '*" + vn + "*' or incidentnumber in (select incidentnumber from supplemental where ((victim1 = 1 and name1 like '*" + vn + "*') or (victim2 = 1 and name2 like '*" + vn + "*'))))"
        Else
            qv = "(vname like '*" + vn + "*' or incidentnumber in (select incidentnumber from supplemental where ((victim1 = 1 and name1 like '*" + vn + "*') or (victim2 = 1 and name2 like '*" + vn + "*'))))"
        End If
    End If
End If
If wn > "" Then
    If wne Then
        If qsu > "" Then
            qsu = qsu + " and ((typeother1 = 'WITNESS' and name1 = " + wn + ") or (typeother2 = 'WITNESS' and name2 = " + wn + "))"
        Else
            qsu = "((typeother1 = 'WITNESS' and name1 = " + wn + ") or (typeother2 = 'WITNESS' and name2 = " + wn + "))"
        End If
    Else
        If qsu > "" Then
            qsu = qsu + " and ((typeother1 = 'WITNESS' and name1 like '*" + wn + "*') or (typeother2 = 'WITNESS' and name2 like '*" + wn + "*'))"
        Else
            qsu = "((typeother1 = 'WITNESS' and name1 like '*" + wn + "*') or (typeother2 = 'WITNESS' and name2 like '*" + wn + "*'))"
        End If
    End If
End If
If p > "" Then
    If pe Then
        If qo > "" Then
            qo = qo + " and ((type1 = " + p + " or type2 = " + p + " or type3 = " + p + " or type4 = " + p + " or type5 = " + p + " or type6 = " + p + ") or incidentnumber in (select incidentnumber from supplemental where (type1 = " + p + " or type2 = " + p + " or type3 = " + p + " or type4 = " + p + " or type5 = " + p + " or type6 = " + p + "))"
        Else
            qo = "((type1 = " + p + " or type2 = " + p + " or type3 = " + p + " or type4 = " + p + " or type5 = " + p + " or type6 = " + p + ") or incidentnumber in (select incidentnumber from supplemental where (type1 = " + p + " or type2 = " + p + " or type3 = " + p + " or type4 = " + p + " or type5 = " + p + " or type6 = " + p + "))"
        End If
    Else
        If qo > "" Then
            qo = qo + " and ((type1 like '*" + p + " or type2 like '*" + p + " or type3 like '*" + p + " or type4 like '*" + p + " or type5 like '*" + p + " or type6 like '*" + p + ") or incidentnumber in (select incidentnumber from supplemental where (type1 like '*" + p + " or type2 like '*" + p + " or type3 like '*" + p + " or type4 like '*" + p + " or type5 like '*" + p + " or type6 like '*" + p + "*'))"
        Else
            qo = "((type1 like '*" + p + " or type2 like '*" + p + " or type3 like '*" + p + " or type4 like '*" + p + " or type5 like '*" + p + " or type6 like '*" + p + ") or incidentnumber in (select incidentnumber from supplemental where (type1 like '*" + p + " or type2 like '*" + p + " or type3 like '*" + p + " or type4 like '*" + p + " or type5 like '*" + p + " or type6 like '*" + p + "*'))"
        End If
    End If
End If
If dac Then
    If qo > "" Then
        qo = qo + " and active = 'X'"
    Else
        qo = "active = 'X'"
    End If
End If
If dcl Then
    If qo > "" Then
        qo = qo + " and admclosed = 'X'"
    Else
        qo = "admclosed = 'X'"
    End If
End If
If dun Then
    If qo > "" Then
        qo = qo + " and unfounded = 'X'"
    Else
        qo = "unfounded = 'X'"
    End If
End If
If dex Then
    If qo > "" Then
        qo = qo + " and (exclearover18 = 'X' or exclearunder18 = 'X')"
    Else
        qo = "(exclearover18 = 'X' or exclearunder18 = 'X')"
    End If
End If
FOUNDT% = 0
For t% = 1 To ro.ListItems.Count
    If ro.ListItems(t%).Selected Then
        FOUNDT% = FOUNDT% + 1
        If FOUNDT% = 1 Then
            If qo > "" Then
                qo = qo + " and ((reportingofficer1 = '" + ro.ListItems(t%) + "' or reportingofficer2 = '" + ro.ListItems(t%) + "')"
            Else
                qo = "((reportingofficer1 = '" + ro.ListItems(t%) + "' or reportingofficer2 = '" + ro.ListItems(t%) + "')"
            End If
        Else
            qo = qo + " or  (reportingofficer1 = '" + ro.ListItems(t%) + "' or reportingofficer2 = '" + ro.ListItems(t%) + "')"
        End If
    End If
Next t%
If FOUNDT% > 0 And (InStr(qo, " and ") > 0 Or InStr(qo, " or ") > 0) Then
    qo = qo + ")"
End If
FOUNDB% = 0
For t% = 1 To bm.ListItems.Count
    If bm.ListItems(t%).Selected Then
        FOUNDB% = FOUNDB% + 1
        tb$ = Mid$(bm.ListItems(t%), InStr(bm.ListItems(t%), "(") + 1, 2)
        If FOUNDB% = 1 Then
            If qo > "" Then
                qo = qo + " and (bias = " + Chr$(34) + tb$ + Chr$(34)
            Else
                qo = qo + " bias = " + Chr$(34) + tb$ + Chr$(34)
            End If
        Else
            qo = " or (bias = " + Chr$(34) + tb$ + Chr$(34)
        End If
    End If
Next t%
If FOUNDB% > 0 And (InStr(qo, " and ") > 0 Or InStr(qo, " or ") > 0) Then
    qo = qo + ")"
End If
foundu% = 0
For t% = 1 To ucr.ListItems.Count
    If ucr.ListItems(t%).Selected Then
        foundu% = foundu% + 1
        uc$ = Mid$(ucr.ListItems(t%), InStr(ucr.ListItems(t%), "(") + 1, 3)
        If foundu% = 1 Then
            If qisu > "" Then
                If foundu% = 1 Then
                    qisu = qisu + " and ((ucr1 = '" + uc$ + "' or ucr2 = '" + uc$ + "' or ucr3 = '" + uc$ + "' or ucr4 = '" + uc$ + "' or ucr5 = '" + uc$ + "' or ucr6 = '" + uc$ + "' or ucr7 = '" + uc$ + "' or ucr8 = '" + uc$ + "' or ucr9 = '" + uc$ + "' or ucr10 = '" + uc$ + "')"
                End If
            Else
                qisu = "((ucr1 = '" + uc$ + "' or ucr2 = '" + uc$ + "' or ucr3 = '" + uc$ + "' or ucr4 = '" + uc$ + "' or ucr5 = '" + uc$ + "' or ucr6 = '" + uc$ + "' or ucr7 = '" + uc$ + "' or ucr8 = '" + uc$ + "' or ucr9 = '" + uc$ + "' or ucr10 = '" + uc$ + "')"
            End If
        Else
            qisu = qisu + " or (ucr1 = '" + uc$ + "' or ucr2 = '" + uc$ + "' or ucr3 = '" + uc$ + "' or ucr4 = '" + uc$ + "' or ucr5 = '" + uc$ + "' or ucr6 = '" + uc$ + "' or ucr7 = '" + uc$ + "' or ucr8 = '" + uc$ + "' or ucr9 = '" + uc$ + "' or ucr10 = '" + uc$ + "')"
        End If
    End If
Next t%
If foundu% > 0 And (InStr(qisu, " and ") > 0 Or InStr(qisu, " or ") > 0) Then
    qisu = qisu + ")"
End If
If QC = "" And qv = "" And qs = "" And qo = "" And qisu = "" And qsu = "" And MAJORMINOR.SelectedItem Is Nothing Then
    msg = MsgBox("No criteria selected.", 48, "Genesis Error Log")
    Exit Sub
End If
findlist.ListItems.clear
Dim db As Database, rs, rs1, rsg As Recordset, bigq, addstring As String
On Error GoTo oderror
od:
Set db = OpenDatabase(nwi + "incident.mdb")
bigq = ""
If QC > "" Then
    bigq = "select incidentnumber from incidentreportc where " + QC
End If
If qv > "" Then
    If bigq > "" Then
        bigq = bigq + " and incidentnumber in (select incidentnumber from incidentreportv where " + qv + ")"
    Else
        bigq = "select incidentnumber from incidentreportv where " + qv
    End If
End If
If qs > "" Then
    If bigq > "" Then
        bigq = bigq + " and incidentnumber in (select incidentnumber from incidentreports where " + qs + ")"
    Else
        bigq = "select incidentnumber from incidentreports where " + qs
    End If
End If
If qo > "" Then
    If bigq > "" Then
        bigq = bigq + " and incidentnumber in (select incidentnumber from incidentreporto where " + qo + ")"
    Else
        bigq = "select incidentnumber from incidentreporto where " + qo
    End If
End If
If qisu > "" Then
    If bigq > "" Then
        bigq = bigq + " and incidentnumber in (select incidentnumber from incidentsupport where " + qisu + ")"
    Else
        bigq = "select incidentnumber from incidentsupport where " + qisu
    End If
End If
If qsu > "" Then
    If bigq > "" Then
        bigq = bigq + " and incidentnumber in (select incidentnumber from supplemental where " + qsu + ")"
    Else
        bigq = "select incidentnumber from supplemental where " + qsu
    End If
End If
If bigq = "" Then
    If MAJORMINOR.SelectedItem Is Nothing Then
        msg = MsgBox("No criteria entered.", 48, "Genesis Error Log")
        Screen.MousePointer = 0
        Exit Sub
    Else
        bigq = "SELECT INCIDENTNUMBER FROM INCIDENTREPORTO"
    End If
End If
Set rs = db.OpenRecordset(bigq + "  order by incidentnumber")
If rs.EOF Then
    On Error Resume Next
    msg = MsgBox("No selections found that match your criteria.", 48, "Genesis Error Log")
    db.Close
    Exit Sub
End If
rs.MoveFirst
addstring = ""
Set rsg = db.OpenRecordset("select * from greport")
If Not rsg.EOF Then
    rsg.MoveFirst
    While Not rsg.EOF
        rsg.Delete
        rsg.MoveNext
    Wend
End If
Set rsg = db.OpenRecordset("select * from greport")
findlist.ListItems.clear
While Not rs.EOF
    If MAJORMINOR.SelectedItem Is Nothing Then
    Else
        Set rs1 = db.OpenRecordset("select major1, minor1,major2, minor2,major3, minor3,major4, minor4,major5, minor5,major6, minor6 from incidentreporto where incidentnumber = " + Chr$(34) + rs("incidentnumber") + Chr$(34))
        foundmatch% = 0
        COUNTSEL% = 0
        For MM% = 1 To MAJORMINOR.ListItems.Count
            If MAJORMINOR.ListItems(MM%).Selected Then
                COUNTSEL% = COUNTSEL% + 1
                tmajor$ = Left$(MAJORMINOR.ListItems(MM%), InStr(MAJORMINOR.ListItems(MM%), "***") - 1)
                tminor$ = Mid$(MAJORMINOR.ListItems(MM%), InStr(MAJORMINOR.ListItems(MM%), "***") + 3)
                For MM2% = 1 To 6
                    tmm$ = Mid$(Str$(MM2%), 2)
                    If Not IsNull(rs1("major" + tmm$)) And Not IsNull(rs1("minor" + tmm$)) Then
                        If tmajor$ = rs1("major" + tmm$) And tminor$ = rs1("minor" + tmm$) Then
                            foundmatch% = 1
                            MM2% = 6
                            MM% = MAJORMINOR.ListItems.Count
                        End If
                    End If
                Next MM2%
            End If
        Next MM%
        If foundmatch% = 0 And COUNTSEL% > 0 Then
            GoTo loopagain
        End If
    End If
    Set rs1 = db.OpenRecordset("select dateofoffense1, offense1 from incidentreportc where incidentnumber = " + Chr$(34) + rs("incidentnumber") + Chr$(34))
    rs1.MoveFirst
    If rs1.EOF Then
        GoTo wendloop
    End If
    rs1.MoveFirst
    rsg.AddNew
    rsg("incidentnumber") = Format$(rs("incidentnumber"), "@@@@@@@@@@@@")
    rsg("incidentdate") = Format$(rs1("dateofoffense1"), "mm/dd/yyyy")
    rsg("offense") = rs1("offense1")
    Set rs1 = db.OpenRecordset("select vname from incidentreportv where incidentnumber = " + Chr$(34) + rs("incidentnumber") + Chr$(34))
    rs1.MoveFirst
    If Not rs1.EOF Then
        rs1.MoveFirst
        rsg("victim") = rs1("vname")
    End If
    Set rs1 = db.OpenRecordset("select excleardate from incidentreporto where incidentnumber = " + Chr$(34) + rs("incidentnumber") + Chr$(34))
    rs1.MoveFirst
    If Not rs1.EOF Then
        rs1.MoveFirst
        If Not IsNull(rs1("excleardate")) Then
            rsg("offense") = "*CLEAR" + Format$(rs1("excleardate"), "mmddyyyy") + " " + Left$(rsg("offense"), 35)
        End If
    End If
    Set rs1 = db.OpenRecordset("select sname from incidentreports where incidentnumber = " + Chr$(34) + rs("incidentnumber") + Chr$(34))
    rs1.MoveFirst
    If Not rs1.EOF Then
        rs1.MoveFirst
        rsg("subject") = rs1("sname")
    End If
wendloop:
    On Error Resume Next
    Set itmx = findlist.ListItems.add(, , rsg("incidentnumber"))
    itmx.SubItems(1) = rsg("incidentdate")
    itmx.SubItems(2) = rsg("offense")
    itmx.SubItems(3) = rsg("victim")
    itmx.SubItems(4) = rsg("subject")
    rsg.Update
loopagain:
    rs.MoveNext
Wend
For yy% = 1 To findlist.ListItems.Count
    findlist.ListItems(yy%).Selected = False
Next yy%
findlist.SelectedItem = Nothing
findlistframe.Top = 500
findlistframe.Left = 250
findlistframe.Visible = True
db.Close
On Error Resume Next
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume
End If
End Sub

Private Sub Command2_Click()
If findlist.ListItems.Count = 0 Then
    msg = MsgBox("A search must be conducted prior to generating a report.", 48, "Genesis Error Log")
    Exit Sub
End If
If findlist.ListItems.Count = 0 Then
    msg = MsgBox("A search must have results prior to generating a report.", 48, "Genesis Error Log")
    Exit Sub
End If
Dim tempf As String
tempf = ""
For t% = 1 To findlist.ListItems.Count
    tempf = tempf + "{event.incidentnumber} = '" + Left$(findlist.ListItems(t%), 12) + "' or "
Next t%
inp = InputBox("Enter Title of Report to be generated.", "Genesis Information Log", "")
Dim db As Database, rs As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nwi + "incident.mdb")
Set rs = db.OpenRecordset("select * from title")
If rs.EOF Then
    rs.AddNew
Else
    rs.Delete
    rs.AddNew
End If
rs("title") = inp
rs.Update
db.Close
On Error Resume Next
tempf = Left$(tempf, Len(tempf) - 3)
reportg.SelectionFormula = ""  'tempf
reportg.ReportFileName = nwi + "greport.rpt"
reportg.WindowShowZoomCtl = True
reportg.Action = 1
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If
End Sub

Private Sub Command3_Click()
findlistframe.Visible = False
End Sub

Private Sub Command4_Click()
inr1 = ""
idr1 = ""
inr = ""
idr2 = ""
il = ""
cn = ""
vn = ""
sn = ""
wn = ""
p = ""

For t% = 1 To ro.ListItems.Count
    ro.ListItems(t%).Selected = False
Next t%
ro.SelectedItem = Nothing
For t% = 1 To bm.ListItems.Count
    bm.ListItems(t%).Selected = False
Next t%
bm.SelectedItem = Nothing
For t% = 1 To ucr.ListItems.Count
    ucr.ListItems(t%).Selected = False
Next t%
ucr.SelectedItem = Nothing
For yy% = 1 To MAJORMINOR.ListItems.Count
    MAJORMINOR.ListItems(yy%).Selected = False
Next yy%
MAJORMINOR.SelectedItem = Nothing
For yy% = 1 To o.ListItems.Count
    o.ListItems(yy%).Selected = False
Next yy%
o.SelectedItem = Nothing
End Sub

Private Sub findlist_Click()
If findlist.SelectedItem Is Nothing Then
    msg = MsgBox("You have to select a record first.", 48, "Genesis Error Log")
    Exit Sub
End If
incident.incidentnumber = findlist.ListItems(findlist.SelectedItem.index)
incident.optimer.Enabled = False
Call incident.getincident
incident.optimer.Enabled = True
Unload Search
End Sub

Private Sub findlist_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
findlist.SortKey = ColumnHeader.index - 1
If findlist.SortOrder = lvwAscending Then
    findlist.SortOrder = lvwDescending
Else
    findlist.SortOrder = lvwAscending
End If
findlist.Sorted = True

End Sub

Private Sub Form_Load()
On Error Resume Next
Dim itmx As ListItem, db As Database, rs As Recordset
On Error GoTo oderror1
od1:
Set db = OpenDatabase(nwl + "lawsuite.mdb")
Set rs = db.OpenRecordset("select profname from professionals where type = 'D' order by profname")
ro.ListItems.clear
If Not rs.EOF Then
    rs.MoveFirst
End If
While Not rs.EOF
    Set itmx = ro.ListItems.add(, , rs("profname"))
    itmx.Selected = False
    rs.MoveNext
Wend
ro.SelectedItem = Nothing
On Error Resume Next
db.Close
On Error GoTo oderror2
od2:
Set db = OpenDatabase(nwi + "incident.mdb")
Set rs = db.OpenRecordset("select code from codes where type = 'bias' order by code")
bm.ListItems.clear
If Not rs.EOF Then
    rs.MoveFirst
End If
While Not rs.EOF
    Set itmx = bm.ListItems.add(, , rs("code"))
    itmx.Selected = False
    rs.MoveNext
Wend
bm.SelectedItem = Nothing
Set rs = db.OpenRecordset("SELECT MAJOR,MINOR FROM PGROUP")
MAJORMINOR.ListItems.clear
If Not rs.EOF Then
    rs.MoveFirst
    While Not rs.EOF
        Set itmx = MAJORMINOR.ListItems.add(, , rs("major") + "***" + rs("minor"))
        rs.MoveNext
    Wend
End If
For yy% = 1 To MAJORMINOR.ListItems.Count
    MAJORMINOR.ListItems(yy%).Selected = False
Next yy%
MAJORMINOR.SelectedItem = Nothing
Set rs = db.OpenRecordset("select DISTINCT OFFENSE from OFFENSE order by OFFENSE")
o.ListItems.clear
If Not rs.EOF Then
    rs.MoveFirst
    While Not rs.EOF
        Set itmx = o.ListItems.add(, , rs("OFFENSE"))
        rs.MoveNext
    Wend
End If
For yy% = 1 To o.ListItems.Count
    o.ListItems(yy%).Selected = False
Next yy%
o.SelectedItem = Nothing
Set rs = db.OpenRecordset("select code from ucr order by code")
ucr.ListItems.clear
If Not rs.EOF Then
    rs.MoveFirst
End If
While Not rs.EOF
    Set itmx = ucr.ListItems.add(, , rs("code"))
    itmx.Selected = False
    rs.MoveNext
Wend
ucr.SelectedItem = Nothing
On Error Resume Next
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

Private Sub Form_Unload(Cancel As Integer)
Set Search = Nothing
End Sub
