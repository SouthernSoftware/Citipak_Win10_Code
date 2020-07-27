VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "grid32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form CIVIL 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808000&
   Caption         =   "Genesis Civil Service version 2.0"
   ClientHeight    =   7380
   ClientLeft      =   60
   ClientTop       =   1140
   ClientWidth     =   11865
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CIVIL.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7380
   ScaleWidth      =   11865
   WindowState     =   2  'Maximized
   Begin VB.Frame infoframe 
      Height          =   6690
      Left            =   120
      TabIndex        =   84
      Top             =   480
      Width           =   11700
      Begin VB.Frame receiptframe 
         BackColor       =   &H00808000&
         Caption         =   "RECEIPT-----------Received From Defendant/Plaintiff/Other"
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
         Height          =   2895
         Left            =   4800
         TabIndex        =   189
         Top             =   2880
         Visible         =   0   'False
         Width           =   5310
         Begin VB.CheckBox fromplaintiff 
            BackColor       =   &H00808000&
            Caption         =   "From Plaintiff"
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
            Height          =   270
            Left            =   2880
            TabIndex        =   188
            Top             =   510
            Width           =   2010
         End
         Begin VB.CheckBox fromdefendant 
            BackColor       =   &H00808000&
            Caption         =   "From Defendant"
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
            Height          =   270
            Left            =   405
            TabIndex        =   187
            Top             =   510
            Width           =   2010
         End
         Begin VB.CommandButton Command13 
            Caption         =   "Close"
            Height          =   255
            Left            =   3960
            TabIndex        =   194
            Top             =   2600
            Width           =   975
         End
         Begin VB.ComboBox othername 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   315
            Left            =   360
            Sorted          =   -1  'True
            TabIndex        =   190
            Top             =   1080
            Width           =   4515
         End
         Begin VB.TextBox otheraddress1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   285
            Left            =   345
            MaxLength       =   60
            TabIndex        =   191
            Top             =   1575
            Width           =   4530
         End
         Begin VB.TextBox otheraddress2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   285
            Left            =   345
            MaxLength       =   60
            TabIndex        =   192
            Top             =   2055
            Width           =   4530
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Print"
            Height          =   255
            Left            =   360
            TabIndex        =   193
            Top             =   2600
            Width           =   975
         End
      End
      Begin VB.Frame paymentframe 
         Caption         =   "Execution Payments                           Balance:"
         Height          =   5595
         Left            =   7560
         TabIndex        =   211
         Top             =   5520
         Visible         =   0   'False
         Width           =   10920
         Begin VB.CommandButton pl 
            Caption         =   "Letter"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   9240
            TabIndex        =   226
            Top             =   5220
            Width           =   750
         End
         Begin VB.TextBox principal 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   285
            Left            =   4755
            TabIndex        =   217
            Top             =   4875
            Width           =   1150
         End
         Begin VB.CommandButton propb 
            Caption         =   "Appropriate Payments"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   222
            Top             =   5220
            Width           =   1875
         End
         Begin VB.CommandButton closepay 
            Caption         =   "Close"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   10080
            TabIndex        =   227
            Top             =   5220
            Width           =   750
         End
         Begin VB.TextBox DATEPAID 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   285
            Left            =   120
            MaxLength       =   10
            TabIndex        =   213
            Top             =   4875
            Width           =   1150
         End
         Begin VB.TextBox AMOUNT 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   285
            Left            =   1270
            TabIndex        =   214
            Top             =   4875
            Width           =   1150
         End
         Begin VB.TextBox RECEIPT 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   285
            Left            =   2420
            TabIndex        =   215
            Top             =   4875
            Width           =   1150
         End
         Begin VB.TextBox CHECK 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   285
            Left            =   3600
            TabIndex        =   216
            Top             =   4875
            Width           =   1150
         End
         Begin VB.CommandButton addpay 
            Caption         =   "Add"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   6720
            TabIndex        =   223
            Top             =   5220
            Width           =   750
         End
         Begin VB.CommandButton removepay 
            Caption         =   "Remove"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   7560
            TabIndex        =   224
            Top             =   5220
            Width           =   750
         End
         Begin VB.TextBox commiss 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   285
            Left            =   5925
            TabIndex        =   218
            Top             =   4875
            Width           =   1150
         End
         Begin VB.TextBox inter 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   285
            Left            =   7080
            TabIndex        =   219
            Top             =   4875
            Width           =   1150
         End
         Begin VB.TextBox remarks 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   285
            Left            =   9480
            MaxLength       =   50
            TabIndex        =   221
            Top             =   4875
            Width           =   1275
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Receipt"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   8400
            TabIndex        =   225
            Top             =   5220
            Width           =   750
         End
         Begin VB.TextBox eservicefee 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   285
            Left            =   8280
            MaxLength       =   50
            TabIndex        =   220
            Top             =   4875
            Width           =   1150
         End
         Begin MSGrid.Grid expaygrid 
            Height          =   4005
            Left            =   150
            TabIndex        =   212
            Top             =   570
            Width           =   10545
            _Version        =   65536
            _ExtentX        =   18600
            _ExtentY        =   7064
            _StockProps     =   77
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderStyle     =   0
            Rows            =   1
            Cols            =   9
            FixedRows       =   0
            FixedCols       =   0
            ScrollBars      =   2
            HighLight       =   0   'False
         End
         Begin Crystal.CrystalReport report 
            Left            =   3840
            Top             =   2640
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            Destination     =   1
            PrintFileLinesPerPage=   60
         End
         Begin VB.Label Label43 
            BackStyle       =   0  'Transparent
            Caption         =   $"CIVIL.frx":0442
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   4560
            Left            =   120
            TabIndex        =   232
            Top             =   360
            Width           =   10620
         End
         Begin VB.Shape Shape1 
            Height          =   255
            Left            =   120
            Top             =   300
            Width           =   10695
         End
         Begin VB.Line Line2 
            X1              =   105
            X2              =   105
            Y1              =   480
            Y2              =   5025
         End
         Begin VB.Line Line3 
            X1              =   10800
            X2              =   10800
            Y1              =   510
            Y2              =   5055
         End
         Begin VB.Label possprin 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   4900
            TabIndex        =   231
            Top             =   0
            Width           =   850
         End
         Begin VB.Label posscomm 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   6120
            TabIndex        =   230
            Top             =   0
            Width           =   855
         End
         Begin VB.Label possint 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   7320
            TabIndex        =   229
            Top             =   0
            Width           =   855
         End
         Begin VB.Label POSSTOTAL 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   9720
            TabIndex        =   228
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.Frame likeframe 
         BackColor       =   &H00808000&
         Caption         =   "LIKE RESULTS ------Click Item to Access"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   5115
         Left            =   10950
         TabIndex        =   206
         Top             =   7680
         Visible         =   0   'False
         Width           =   10425
         Begin VB.CommandButton closebutton 
            Caption         =   "Close"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   9240
            TabIndex        =   207
            Top             =   4750
            Width           =   1140
         End
         Begin MSComctlLib.ListView likelist 
            Height          =   4335
            Left            =   120
            TabIndex        =   208
            Top             =   240
            Width           =   10200
            _ExtentX        =   17992
            _ExtentY        =   7646
            View            =   3
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
               Text            =   "Service Of"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Date Received"
               Object.Width           =   2205
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Iteration"
               Object.Width           =   1499
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Home Address"
               Object.Width           =   7056
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Work Address"
               Object.Width           =   3528
            EndProperty
         End
      End
      Begin VB.Frame remarksframe 
         BackColor       =   &H00808000&
         Caption         =   "Remarks and Reasons"
         ForeColor       =   &H00FFFFFF&
         Height          =   3900
         Left            =   11580
         TabIndex        =   199
         Top             =   7065
         Visible         =   0   'False
         Width           =   8400
         Begin VB.CommandButton rclose 
            Caption         =   "Close"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   7320
            TabIndex        =   204
            Top             =   3600
            Width           =   930
         End
         Begin VB.TextBox nsreason 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   285
            Left            =   120
            MaxLength       =   255
            TabIndex        =   202
            Top             =   2040
            Width           =   8175
         End
         Begin VB.TextBox wremarks 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   285
            Left            =   120
            MaxLength       =   255
            TabIndex        =   201
            Top             =   1400
            Width           =   8175
         End
         Begin VB.TextBox premarks 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   285
            Left            =   90
            MaxLength       =   255
            TabIndex        =   200
            Top             =   675
            Width           =   8175
         End
         Begin RichTextLib.RichTextBox levy 
            Height          =   855
            Left            =   120
            TabIndex        =   203
            Top             =   2640
            Width           =   8175
            _ExtentX        =   14420
            _ExtentY        =   1508
            _Version        =   393217
            TextRTF         =   $"CIVIL.frx":139A
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label LEVYL 
            BackStyle       =   0  'Transparent
            Caption         =   $"CIVIL.frx":141C
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   2220
            Left            =   120
            TabIndex        =   205
            Top             =   405
            Width           =   2415
         End
      End
      Begin VB.Frame mprintframe 
         BackColor       =   &H00808000&
         Caption         =   "Print Options"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   6120
         Left            =   3720
         TabIndex        =   162
         Top             =   600
         Visible         =   0   'False
         Width           =   2220
         Begin VB.OptionButton dos 
            BackColor       =   &H00808000&
            Caption         =   "Determination of Sale"
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
            Height          =   285
            Left            =   50
            TabIndex        =   175
            Top             =   4920
            Visible         =   0   'False
            Width           =   2100
         End
         Begin VB.OptionButton status 
            BackColor       =   &H00808000&
            Caption         =   "Status Letter"
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
            Height          =   285
            Left            =   50
            TabIndex        =   209
            Top             =   1680
            Width           =   2100
         End
         Begin VB.OptionButton Partial 
            BackColor       =   &H00808000&
            Caption         =   "Partial Satisfaction"
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
            Height          =   285
            Left            =   50
            TabIndex        =   169
            Top             =   2760
            Visible         =   0   'False
            Width           =   2100
         End
         Begin VB.OptionButton levyp 
            BackColor       =   &H00808000&
            Caption         =   "Notice of Levy"
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
            Height          =   285
            Left            =   50
            TabIndex        =   174
            Top             =   4560
            Visible         =   0   'False
            Width           =   2100
         End
         Begin VB.OptionButton nrl 
            BackColor       =   &H00808000&
            Caption         =   "No Response Letter"
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
            Height          =   285
            Left            =   50
            TabIndex        =   173
            Top             =   4200
            Visible         =   0   'False
            Width           =   2100
         End
         Begin VB.OptionButton RL 
            BackColor       =   &H00808000&
            Caption         =   "Reminder Letter"
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
            Height          =   285
            Left            =   50
            TabIndex        =   172
            Top             =   3840
            Visible         =   0   'False
            Width           =   2100
         End
         Begin VB.OptionButton nullaex 
            BackColor       =   &H00808000&
            Caption         =   "Nulla Bona"
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
            Height          =   285
            Left            =   50
            TabIndex        =   171
            Top             =   3480
            Visible         =   0   'False
            Width           =   2100
         End
         Begin VB.OptionButton sat 
            BackColor       =   &H00808000&
            Caption         =   "Satisfaction"
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
            Height          =   285
            Left            =   50
            TabIndex        =   170
            Top             =   3120
            Visible         =   0   'False
            Width           =   2100
         End
         Begin VB.OptionButton asb 
            BackColor       =   &H00808000&
            Caption         =   "Execution Account"
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
            Height          =   285
            Left            =   50
            TabIndex        =   168
            Top             =   2415
            Visible         =   0   'False
            Width           =   2100
         End
         Begin VB.OptionButton preceipt 
            BackColor       =   &H00808000&
            Caption         =   "Receipt"
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
            Height          =   285
            Left            =   50
            TabIndex        =   166
            Top             =   1320
            Width           =   2100
         End
         Begin VB.OptionButton epwb 
            BackColor       =   &H00808000&
            Caption         =   "Property Worksheet"
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
            Height          =   285
            Left            =   50
            TabIndex        =   167
            Top             =   2040
            Visible         =   0   'False
            Width           =   2100
         End
         Begin VB.CommandButton goprint 
            Caption         =   "Print"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   176
            Top             =   5640
            Width           =   690
         End
         Begin VB.OptionButton affidavit 
            BackColor       =   &H00808000&
            Caption         =   "Affidavit"
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
            Height          =   285
            Left            =   50
            TabIndex        =   165
            Top             =   975
            Width           =   2100
         End
         Begin VB.OptionButton letter 
            BackColor       =   &H00808000&
            Caption         =   "Letter"
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
            Height          =   285
            Left            =   50
            TabIndex        =   164
            Top             =   600
            Width           =   2100
         End
         Begin VB.OptionButton worksheet 
            BackColor       =   &H00808000&
            Caption         =   "Worksheet"
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
            Height          =   285
            Left            =   50
            TabIndex        =   163
            Top             =   240
            Value           =   -1  'True
            Width           =   2100
         End
         Begin VB.CommandButton closeprint 
            Caption         =   "Close"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1440
            TabIndex        =   177
            Top             =   5640
            Width           =   690
         End
      End
      Begin VB.Frame INDEXFRAME 
         BackColor       =   &H00808000&
         Caption         =   "INDEX OF PAPERS"
         Height          =   5415
         Left            =   1605
         TabIndex        =   184
         Top             =   1000
         Visible         =   0   'False
         Width           =   9375
         Begin MSComctlLib.ListView alllist 
            Height          =   4695
            Left            =   60
            TabIndex        =   235
            Top             =   240
            Width           =   9225
            _ExtentX        =   16272
            _ExtentY        =   8281
            View            =   3
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
            NumItems        =   6
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Category"
               Object.Width           =   1940
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Service Of"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Date Received"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Iteration"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Assigned To"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Paper Type"
               Object.Width           =   5292
            EndProperty
         End
         Begin VB.CommandButton Command11 
            Caption         =   "Close"
            Height          =   300
            Left            =   7680
            TabIndex        =   186
            Top             =   5040
            Width           =   1575
         End
         Begin VB.CommandButton Command12 
            Caption         =   "Back 1 Year"
            Height          =   300
            Left            =   3840
            TabIndex        =   185
            Top             =   5040
            Width           =   1575
         End
      End
      Begin VB.Frame mframe 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   650
         Left            =   5430
         TabIndex        =   240
         Top             =   135
         Width           =   700
         Begin VB.Image MUGSHOT 
            BorderStyle     =   1  'Fixed Single
            Height          =   645
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   705
         End
      End
      Begin VB.TextBox checkd 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   7560
         TabIndex        =   26
         Top             =   1830
         Width           =   705
      End
      Begin VB.TextBox total 
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   7920
         MaxLength       =   100
         TabIndex        =   183
         TabStop         =   0   'False
         Top             =   6300
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.CommandButton INDEXBUTTON 
         Caption         =   "IDX"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   325
         Left            =   50
         TabIndex        =   182
         Top             =   225
         Width           =   400
      End
      Begin VB.TextBox INTEREST 
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4680
         MaxLength       =   100
         TabIndex        =   180
         TabStop         =   0   'False
         Top             =   6270
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.CheckBox nulla 
         Caption         =   "Nulla Bona"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   9270
         TabIndex        =   70
         Top             =   5475
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox receiptd 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   5760
         TabIndex        =   25
         Top             =   1830
         Width           =   945
      End
      Begin VB.CommandButton remarksbutton 
         Caption         =   "&6 Remarks"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   10545
         TabIndex        =   77
         Top             =   5895
         Width           =   1050
      End
      Begin VB.TextBox datereceived 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   8130
         MaxLength       =   10
         TabIndex        =   2
         Top             =   225
         Width           =   1215
      End
      Begin VB.ComboBox iteration 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   315
         Left            =   9450
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   225
         Width           =   630
      End
      Begin VB.TextBox courttime 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   9360
         MaxLength       =   10
         TabIndex        =   18
         Top             =   1470
         Width           =   840
      End
      Begin VB.TextBox defendantsort 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   10005
         MaxLength       =   15
         TabIndex        =   29
         Top             =   2205
         Width           =   1650
      End
      Begin VB.TextBox plaintiffsort 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   10005
         MaxLength       =   15
         TabIndex        =   41
         Top             =   3255
         Width           =   1650
      End
      Begin VB.CommandButton paybutton 
         Caption         =   "Payments"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9360
         TabIndex        =   71
         Top             =   6150
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.TextBox balance 
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   7920
         MaxLength       =   100
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   5895
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.TextBox perday 
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   6240
         MaxLength       =   100
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   5910
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.TextBox commission 
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4680
         MaxLength       =   100
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   5910
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.TextBox judgementamount 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   8160
         MaxLength       =   100
         TabIndex        =   65
         Top             =   5445
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.TextBox intrate 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   300
         Left            =   2160
         MaxLength       =   100
         TabIndex        =   62
         Top             =   5445
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.CheckBox ivd 
         Caption         =   "IVD  Custodian:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   3705
         TabIndex        =   23
         Top             =   1845
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.CommandButton likebutton 
         Caption         =   "&1 Like"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   10560
         TabIndex        =   72
         Top             =   4410
         Width           =   1050
      End
      Begin VB.TextBox relationship 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   4100
         MaxLength       =   120
         TabIndex        =   59
         Top             =   5055
         Width           =   2490
      End
      Begin VB.TextBox locationserved 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   7440
         MaxLength       =   60
         TabIndex        =   60
         Top             =   5055
         Width           =   3000
      End
      Begin VB.TextBox personserved 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   720
         MaxLength       =   60
         TabIndex        =   58
         Top             =   5070
         Width           =   2100
      End
      Begin VB.TextBox servicetime 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   6000
         MaxLength       =   10
         TabIndex        =   57
         Top             =   4650
         Width           =   1140
      End
      Begin VB.CheckBox nonservice 
         Alignment       =   1  'Right Justify
         Caption         =   "Non-Service"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   8280
         TabIndex        =   55
         Top             =   4650
         Width           =   1320
      End
      Begin VB.CheckBox served 
         Alignment       =   1  'Right Justify
         Caption         =   "Served"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   8640
         TabIndex        =   54
         Top             =   4350
         Width           =   945
      End
      Begin VB.TextBox assignedon 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   6000
         MaxLength       =   10
         TabIndex        =   53
         Top             =   4305
         Width           =   1170
      End
      Begin VB.ComboBox assignedto 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   315
         Left            =   1320
         TabIndex        =   52
         Top             =   4305
         Width           =   3285
      End
      Begin VB.CommandButton printbutton 
         Caption         =   "&5 Print"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   10545
         TabIndex        =   76
         Top             =   5610
         Width           =   1050
      End
      Begin VB.CommandButton deletebutton 
         Caption         =   "&4 Delete"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   10545
         TabIndex        =   75
         Top             =   5295
         Width           =   1050
      End
      Begin VB.CommandButton clearbutton 
         Caption         =   "&3 Clear"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   10545
         TabIndex        =   74
         Top             =   4995
         Width           =   1050
      End
      Begin VB.CommandButton savebutton 
         Caption         =   "&2 Save"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   10560
         TabIndex        =   73
         Top             =   4695
         Width           =   1050
      End
      Begin VB.TextBox pworkaddress 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   1320
         MaxLength       =   60
         TabIndex        =   47
         Top             =   3930
         Width           =   3750
      End
      Begin VB.TextBox phomeaddress 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   1320
         MaxLength       =   60
         TabIndex        =   42
         Top             =   3615
         Width           =   3750
      End
      Begin VB.ComboBox professional 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   315
         Left            =   9240
         Sorted          =   -1  'True
         TabIndex        =   27
         Top             =   1845
         Width           =   2445
      End
      Begin VB.TextBox dworkaddress 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   1320
         MaxLength       =   60
         TabIndex        =   35
         Top             =   2895
         Width           =   3750
      End
      Begin VB.TextBox dhomeaddress 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   1305
         MaxLength       =   60
         TabIndex        =   30
         Top             =   2565
         Width           =   3750
      End
      Begin VB.TextBox soworkaddress 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   1320
         MaxLength       =   60
         TabIndex        =   10
         Top             =   1125
         Width           =   3750
      End
      Begin VB.TextBox sohomeaddress 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   1320
         MaxLength       =   60
         TabIndex        =   5
         Top             =   825
         Width           =   3750
      End
      Begin VB.TextBox servicefee 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   720
         TabIndex        =   20
         Top             =   1845
         Width           =   660
      End
      Begin VB.TextBox serviceofsort 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   300
         Left            =   10215
         MaxLength       =   15
         TabIndex        =   4
         Top             =   225
         Width           =   1455
      End
      Begin VB.TextBox daystorespond 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   11070
         MaxLength       =   15
         TabIndex        =   19
         Top             =   1470
         Width           =   585
      End
      Begin VB.TextBox casenumber 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   750
         MaxLength       =   20
         TabIndex        =   15
         Top             =   1440
         Width           =   1056
      End
      Begin VB.TextBox pworkphone 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   315
         Left            =   10005
         MaxLength       =   20
         TabIndex        =   51
         Top             =   3945
         Width           =   1665
      End
      Begin VB.TextBox phomephone 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   315
         Left            =   10005
         MaxLength       =   20
         TabIndex        =   46
         Top             =   3600
         Width           =   1665
      End
      Begin VB.TextBox dworkphone 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   315
         Left            =   10005
         MaxLength       =   20
         TabIndex        =   39
         Top             =   2895
         Width           =   1665
      End
      Begin VB.TextBox dhomephone 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   315
         Left            =   10005
         MaxLength       =   20
         TabIndex        =   34
         Top             =   2550
         Width           =   1665
      End
      Begin VB.TextBox soworkphone 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   315
         Left            =   10005
         MaxLength       =   20
         TabIndex        =   14
         Top             =   1125
         Width           =   1665
      End
      Begin VB.TextBox sohomephone 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   315
         Left            =   10005
         MaxLength       =   20
         TabIndex        =   9
         Top             =   825
         Width           =   1665
      End
      Begin VB.TextBox estpayoffdate 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   315
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   66
         Top             =   5910
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.TextBox judgementdate 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   315
         Left            =   6000
         MaxLength       =   10
         TabIndex        =   64
         Top             =   5445
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox datesatisfied 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   315
         Left            =   3750
         MaxLength       =   10
         TabIndex        =   63
         Top             =   5445
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox apptdate 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   315
         Left            =   480
         MaxLength       =   10
         TabIndex        =   61
         Top             =   5445
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox servicedate 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   315
         Left            =   2520
         MaxLength       =   10
         TabIndex        =   56
         Top             =   4650
         Width           =   1215
      End
      Begin VB.TextBox feedate 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   315
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   21
         Top             =   1830
         Width           =   1140
      End
      Begin VB.CheckBox bill 
         Caption         =   "Bill"
         ForeColor       =   &H00808000&
         Height          =   240
         Left            =   3705
         TabIndex        =   22
         Top             =   1860
         Width           =   675
      End
      Begin VB.TextBox custodian 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   5400
         MaxLength       =   100
         TabIndex        =   24
         Top             =   1830
         Visible         =   0   'False
         Width           =   2865
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Case"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1720
         TabIndex        =   233
         Top             =   1470
         Width           =   385
      End
      Begin MSDBCtls.DBCombo serviceof 
         Bindings        =   "CIVIL.frx":1585
         DataSource      =   "Data1"
         Height          =   315
         Left            =   480
         TabIndex        =   0
         Top             =   225
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   556
         _Version        =   393216
         ForeColor       =   8421376
         ListField       =   "dpname"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDBCtls.DBCombo PLAINTIFF 
         Bindings        =   "CIVIL.frx":1599
         DataSource      =   "Data1"
         Height          =   315
         Left            =   1320
         TabIndex        =   40
         Top             =   3270
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   556
         _Version        =   393216
         ForeColor       =   8421376
         ListField       =   "dpname"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDBCtls.DBCombo DEFENDANT 
         Bindings        =   "CIVIL.frx":15AD
         DataSource      =   "Data1"
         Height          =   315
         Left            =   1320
         TabIndex        =   28
         Top             =   2220
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   556
         _Version        =   393216
         ForeColor       =   8421376
         ListField       =   "dpname"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox dhomeaddress2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   5175
         MaxLength       =   60
         TabIndex        =   31
         Top             =   2565
         Width           =   2340
      End
      Begin VB.TextBox dhomestate 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   7590
         MaxLength       =   2
         TabIndex        =   32
         Top             =   2565
         Width           =   405
      End
      Begin VB.TextBox dhomezipcode 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   8100
         MaxLength       =   10
         TabIndex        =   33
         Top             =   2565
         Width           =   930
      End
      Begin VB.TextBox dworkaddress2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   5175
         MaxLength       =   60
         TabIndex        =   36
         Top             =   2895
         Width           =   2340
      End
      Begin VB.TextBox dworkstate 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   7590
         MaxLength       =   2
         TabIndex        =   37
         Top             =   2895
         Width           =   405
      End
      Begin VB.TextBox dworkzipcode 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   8100
         MaxLength       =   10
         TabIndex        =   38
         Top             =   2895
         Width           =   930
      End
      Begin VB.TextBox phomeaddress2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   5175
         MaxLength       =   60
         TabIndex        =   43
         Top             =   3615
         Width           =   2340
      End
      Begin VB.TextBox phomestate 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   7590
         MaxLength       =   2
         TabIndex        =   44
         Top             =   3615
         Width           =   405
      End
      Begin VB.TextBox phomezipcode 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   8100
         MaxLength       =   10
         TabIndex        =   45
         Top             =   3615
         Width           =   930
      End
      Begin VB.TextBox pworkaddress2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   5175
         MaxLength       =   60
         TabIndex        =   48
         Top             =   3930
         Width           =   2340
      End
      Begin VB.TextBox pworkstate 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   7590
         MaxLength       =   2
         TabIndex        =   49
         Top             =   3930
         Width           =   405
      End
      Begin VB.TextBox pworkzipcode 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   8100
         MaxLength       =   10
         TabIndex        =   50
         Top             =   3930
         Width           =   930
      End
      Begin VB.TextBox soworkzipcode 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   8100
         MaxLength       =   10
         TabIndex        =   13
         Top             =   1125
         Width           =   930
      End
      Begin VB.TextBox soworkstate 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   7590
         MaxLength       =   2
         TabIndex        =   12
         Top             =   1125
         Width           =   405
      End
      Begin VB.TextBox soworkaddress2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   5175
         MaxLength       =   60
         TabIndex        =   11
         Top             =   1125
         Width           =   2340
      End
      Begin VB.TextBox sohomezipcode 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   8100
         MaxLength       =   10
         TabIndex        =   8
         Top             =   825
         Width           =   930
      End
      Begin VB.TextBox sohomestate 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   7590
         MaxLength       =   2
         TabIndex        =   7
         Top             =   825
         Width           =   405
      End
      Begin VB.TextBox sohomeaddress2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   5175
         MaxLength       =   60
         TabIndex        =   6
         Top             =   825
         Width           =   2340
      End
      Begin VB.TextBox courtdate 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   315
         Left            =   7560
         MaxLength       =   10
         TabIndex        =   17
         Top             =   1470
         Width           =   1140
      End
      Begin VB.ComboBox papertype 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   315
         Left            =   2760
         Sorted          =   -1  'True
         TabIndex        =   16
         Top             =   1455
         Width           =   3945
      End
      Begin VB.ListBox daterlist 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   450
         Left            =   6405
         TabIndex        =   1
         Top             =   135
         Width           =   1620
      End
      Begin VB.CheckBox corporate 
         Caption         =   "Corporate     Title:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   210
         Left            =   1320
         TabIndex        =   242
         Top             =   540
         Width           =   1695
      End
      Begin VB.TextBox title 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   300
         Left            =   3000
         MaxLength       =   50
         TabIndex        =   243
         Top             =   520
         Width           =   1890
      End
      Begin VB.CheckBox armedforces 
         Caption         =   "Armed Forces"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   210
         Left            =   50
         TabIndex        =   241
         Top             =   540
         Width           =   1260
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Plaintiff:                               Home Address:                         Work Address:                      Assigned to:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   1365
         Left            =   0
         TabIndex        =   181
         Top             =   3270
         Width           =   1590
      End
      Begin VB.Label receiptl 
         BackStyle       =   0  'Transparent
         Caption         =   $"CIVIL.frx":15C1
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   2775
         Left            =   5280
         TabIndex        =   138
         Top             =   1830
         Width           =   2385
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   $"CIVIL.frx":17C0
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   3585
         Left            =   8985
         TabIndex        =   122
         Top             =   855
         Width           =   1140
      End
      Begin VB.Label Label54 
         BackStyle       =   0  'Transparent
         Caption         =   $"CIVIL.frx":18CD
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   435
         Left            =   2160
         TabIndex        =   121
         Top             =   1425
         Width           =   8970
      End
      Begin VB.Label Label40 
         BackStyle       =   0  'Transparent
         Caption         =   $"CIVIL.frx":19E2
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   780
         Left            =   3600
         TabIndex        =   87
         Top             =   5820
         Visible         =   0   'False
         Width           =   4380
      End
      Begin VB.Label Label34 
         BackStyle       =   0  'Transparent
         Caption         =   $"CIVIL.frx":1AD8
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   885
         Left            =   0
         TabIndex        =   86
         Top             =   5445
         Visible         =   0   'False
         Width           =   8235
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   $"CIVIL.frx":1C63
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   540
         Left            =   0
         TabIndex        =   85
         Top             =   4845
         Width           =   7560
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "Service/Non-Service Date:                                    Time of Service:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   300
         Left            =   30
         TabIndex        =   83
         Top             =   4650
         Width           =   5970
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Magistrate:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   8280
         TabIndex        =   82
         Top             =   1875
         Width           =   1230
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   $"CIVIL.frx":1D49
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   2565
         Left            =   0
         TabIndex        =   81
         Top             =   840
         Width           =   1590
      End
      Begin VB.Label feel 
         BackStyle       =   0  'Transparent
         Caption         =   "Fee:                  Fee Date:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   300
         Left            =   45
         TabIndex        =   80
         Top             =   1830
         Width           =   2490
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   $"CIVIL.frx":1E38
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   255
         Left            =   480
         TabIndex        =   79
         Top             =   0
         Width           =   10935
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   60
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7000
      Visible         =   0   'False
      Width           =   1140
   End
   Begin TabDlg.SSTab maintab 
      Height          =   7080
      Left            =   0
      TabIndex        =   78
      Top             =   -15
      Width           =   11865
      _ExtentX        =   20929
      _ExtentY        =   12488
      _Version        =   393216
      Tabs            =   7
      TabsPerRow      =   7
      TabHeight       =   794
      TabMaxWidth     =   2955
      BackColor       =   8421376
      ForeColor       =   4210688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   " &Magistrate Papers"
      TabPicture(0)   =   "CIVIL.frx":1ED2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "&Writs/Other Papers"
      TabPicture(1)   =   "CIVIL.frx":1EEE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Family Court Papers"
      TabPicture(2)   =   "CIVIL.frx":1F0A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Executions"
      TabPicture(3)   =   "CIVIL.frx":1F26
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      TabCaption(4)   =   "&Reports/Checks"
      TabPicture(4)   =   "CIVIL.frx":1F42
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "fromdate"
      Tab(4).Control(1)=   "todate"
      Tab(4).Control(2)=   "Frame2"
      Tab(4).Control(3)=   "Label2"
      Tab(4).ControlCount=   4
      TabCaption(5)   =   "&Outstanding Papers/Search"
      TabPicture(5)   =   "CIVIL.frx":1F5E
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "sname"
      Tab(5).Control(1)=   "scasenumber"
      Tab(5).Control(2)=   "Command14"
      Tab(5).Control(3)=   "accessbutton"
      Tab(5).Control(4)=   "outstandinglist"
      Tab(5).Control(5)=   "Labelo"
      Tab(5).ControlCount=   6
      TabCaption(6)   =   "&System"
      TabPicture(6)   =   "CIVIL.frx":1F7A
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "odep"
      Tab(6).Control(1)=   "ocou"
      Tab(6).Control(2)=   "oatt"
      Tab(6).Control(3)=   "omag"
      Tab(6).Control(4)=   "Command15"
      Tab(6).Control(5)=   "List2"
      Tab(6).Control(6)=   "List1"
      Tab(6).Control(7)=   "autoprint"
      Tab(6).Control(8)=   "nextreceipt"
      Tab(6).Control(9)=   "Command9"
      Tab(6).Control(10)=   "lnf"
      Tab(6).Control(11)=   "fnf"
      Tab(6).Control(12)=   "county"
      Tab(6).Control(13)=   "office"
      Tab(6).Control(14)=   "Command6"
      Tab(6).Control(15)=   "treasureraddress2"
      Tab(6).Control(16)=   "treasureraddress1"
      Tab(6).Control(17)=   "treasurer"
      Tab(6).Control(18)=   "papertypelist"
      Tab(6).Control(19)=   "Command3"
      Tab(6).Control(20)=   "Command2"
      Tab(6).Control(21)=   "Command1"
      Tab(6).Control(22)=   "UPDSHERIFF"
      Tab(6).Control(23)=   "dprof"
      Tab(6).Control(24)=   "aeprof"
      Tab(6).Control(25)=   "profname"
      Tab(6).Control(26)=   "profaddr1"
      Tab(6).Control(27)=   "profaddr2"
      Tab(6).Control(28)=   "profphone"
      Tab(6).Control(29)=   "exintrate"
      Tab(6).Control(30)=   "excommrate1"
      Tab(6).Control(31)=   "exonfirst"
      Tab(6).Control(32)=   "excommrate2"
      Tab(6).Control(33)=   "sheriff"
      Tab(6).Control(34)=   "sheriffphone"
      Tab(6).Control(35)=   "sheriffaddress2"
      Tab(6).Control(36)=   "sheriffaddress"
      Tab(6).Control(37)=   "FROMXREF"
      Tab(6).Control(38)=   "Label65"
      Tab(6).Control(39)=   "Label49"
      Tab(6).Control(40)=   "Label56"
      Tab(6).Control(41)=   "Label61"
      Tab(6).ControlCount=   42
      Begin VB.OptionButton odep 
         Caption         =   "Deputy"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   240
         Left            =   -64545
         TabIndex        =   239
         Top             =   900
         Width           =   1125
      End
      Begin VB.OptionButton ocou 
         Caption         =   "Court"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   240
         Left            =   -65850
         TabIndex        =   238
         Top             =   900
         Width           =   1305
      End
      Begin VB.OptionButton oatt 
         Caption         =   "Attorney"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   240
         Left            =   -67320
         TabIndex        =   237
         Top             =   900
         Width           =   1305
      End
      Begin VB.OptionButton omag 
         Caption         =   "Magistrate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   240
         Left            =   -68895
         TabIndex        =   236
         Top             =   900
         Value           =   -1  'True
         Width           =   1305
      End
      Begin VB.CommandButton Command15 
         Caption         =   "List"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -64080
         TabIndex        =   117
         Top             =   5715
         Width           =   810
      End
      Begin VB.TextBox fromdate 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   300
         Left            =   -74760
         TabIndex        =   125
         Top             =   1020
         Width           =   1215
      End
      Begin VB.TextBox todate 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   300
         Left            =   -74760
         TabIndex        =   126
         Top             =   1545
         Width           =   1215
      End
      Begin VB.TextBox sname 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -74835
         TabIndex        =   197
         Top             =   6600
         Width           =   2655
      End
      Begin VB.TextBox scasenumber 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -71880
         TabIndex        =   196
         Top             =   6600
         Width           =   2655
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Search All Tabs"
         Height          =   375
         Left            =   -69000
         TabIndex        =   195
         Top             =   6600
         Width           =   2055
      End
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -65950
         TabIndex        =   119
         Top             =   6300
         Width           =   2415
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -68880
         TabIndex        =   118
         Top             =   6300
         Width           =   2415
      End
      Begin VB.CheckBox autoprint 
         Caption         =   "Auto-Print"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   375
         Left            =   -68910
         TabIndex        =   108
         Top             =   2820
         Width           =   1335
      End
      Begin VB.TextBox nextreceipt 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   345
         Left            =   -67440
         MaxLength       =   15
         TabIndex        =   112
         Top             =   4485
         Width           =   1005
      End
      Begin VB.CommandButton Command9 
         Caption         =   "ARCHIVE DATA "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -68880
         TabIndex        =   120
         Top             =   6585
         Width           =   5610
      End
      Begin VB.OptionButton lnf 
         Caption         =   "Last Name First"
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   -71520
         TabIndex        =   94
         Top             =   2880
         Width           =   2055
      End
      Begin VB.OptionButton fnf 
         Caption         =   "First Name First"
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   -74745
         TabIndex        =   93
         Top             =   2880
         Width           =   2175
      End
      Begin VB.TextBox county 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   345
         Left            =   -73560
         MaxLength       =   50
         TabIndex        =   100
         Top             =   6240
         Width           =   4350
      End
      Begin VB.TextBox office 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   345
         Left            =   -73560
         MaxLength       =   50
         TabIndex        =   95
         Top             =   4200
         Width           =   4350
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Add/Update"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   -64560
         TabIndex        =   113
         Top             =   4560
         Width           =   1260
      End
      Begin VB.TextBox treasureraddress2 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   345
         Left            =   -67440
         MaxLength       =   75
         TabIndex        =   111
         Top             =   4035
         Width           =   4215
      End
      Begin VB.TextBox treasureraddress1 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   345
         Left            =   -67440
         MaxLength       =   75
         TabIndex        =   110
         Top             =   3645
         Width           =   4215
      End
      Begin VB.TextBox treasurer 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   345
         Left            =   -67440
         MaxLength       =   75
         TabIndex        =   109
         Top             =   3255
         Width           =   4215
      End
      Begin VB.ComboBox papertypelist 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   345
         Left            =   -68835
         Sorted          =   -1  'True
         TabIndex        =   114
         Top             =   5280
         Width           =   5550
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Add/Edit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -66480
         TabIndex        =   115
         Top             =   5715
         Width           =   810
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -65280
         TabIndex        =   116
         Top             =   5715
         Width           =   810
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Add/Update"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   -70560
         TabIndex        =   92
         Top             =   2520
         Width           =   1380
      End
      Begin VB.CommandButton UPDSHERIFF 
         Caption         =   "Add/Update"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   -70680
         TabIndex        =   101
         Top             =   6650
         Width           =   1500
      End
      Begin VB.CommandButton accessbutton 
         Caption         =   "Access"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -64920
         TabIndex        =   178
         Top             =   6600
         Width           =   1695
      End
      Begin VB.CommandButton dprof 
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -64155
         TabIndex        =   107
         Top             =   2505
         Width           =   810
      End
      Begin VB.CommandButton aeprof 
         Caption         =   "Add/Edit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -65325
         TabIndex        =   106
         Top             =   2505
         Width           =   810
      End
      Begin VB.ComboBox profname 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   345
         Left            =   -67680
         Sorted          =   -1  'True
         TabIndex        =   102
         Top             =   1260
         Width           =   4350
      End
      Begin VB.TextBox profaddr1 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   345
         Left            =   -67680
         MaxLength       =   75
         TabIndex        =   103
         Top             =   1645
         Width           =   4350
      End
      Begin VB.TextBox profaddr2 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   345
         Left            =   -67680
         MaxLength       =   75
         TabIndex        =   104
         Top             =   2040
         Width           =   4350
      End
      Begin VB.TextBox profphone 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   330
         Left            =   -67680
         MaxLength       =   15
         TabIndex        =   105
         Top             =   2415
         Width           =   2145
      End
      Begin VB.TextBox exintrate 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   345
         Left            =   -72075
         MaxLength       =   15
         TabIndex        =   88
         Top             =   1050
         Width           =   1005
      End
      Begin VB.TextBox excommrate1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   360
         Left            =   -72090
         MaxLength       =   15
         TabIndex        =   89
         Top             =   1560
         Width           =   1005
      End
      Begin VB.TextBox exonfirst 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   345
         Left            =   -69885
         MaxLength       =   15
         TabIndex        =   90
         Top             =   1560
         Width           =   555
      End
      Begin VB.TextBox excommrate2 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   345
         Left            =   -72090
         MaxLength       =   15
         TabIndex        =   91
         Top             =   2190
         Width           =   1005
      End
      Begin VB.Frame Frame2 
         Caption         =   "Reports and Checks"
         ForeColor       =   &H00404000&
         Height          =   6550
         Left            =   -73320
         TabIndex        =   124
         Top             =   480
         Width           =   10185
         Begin VB.CommandButton Command7 
            Caption         =   "Print Checks"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   8520
            TabIndex        =   247
            Top             =   6120
            Width           =   1455
         End
         Begin VB.OptionButton epic 
            Caption         =   "Execution Principal and Interest Checks"
            ForeColor       =   &H00808000&
            Height          =   840
            Left            =   8040
            TabIndex        =   246
            Top             =   1905
            Width           =   1995
         End
         Begin VB.OptionButton ecc 
            Caption         =   "Execution Commission Checks"
            ForeColor       =   &H00808000&
            Height          =   840
            Left            =   8040
            TabIndex        =   245
            Top             =   945
            Width           =   2000
         End
         Begin VB.OptionButton Sfc 
            Caption         =   "Service Fee Checks"
            ForeColor       =   &H00808000&
            Height          =   480
            Left            =   8040
            TabIndex        =   244
            Top             =   225
            Width           =   1755
         End
         Begin VB.OptionButton opsr 
            Caption         =   "Officer Paper Status Report"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   255
            Left            =   3960
            TabIndex        =   149
            Top             =   1080
            Width           =   3660
         End
         Begin VB.OptionButton owbor 
            Caption         =   "Outstanding Writ Papers By Officer Report"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   495
            Left            =   3960
            TabIndex        =   154
            Top             =   3720
            Width           =   3660
         End
         Begin VB.OptionButton ofcbor 
            Caption         =   "Outstanding Family Court Papers By Officer Report"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   495
            Left            =   3960
            TabIndex        =   153
            Top             =   3120
            Width           =   3660
         End
         Begin VB.OptionButton bsfl 
            Caption         =   "Billed Service Fee Listing"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   255
            Left            =   120
            TabIndex        =   146
            Top             =   6240
            Width           =   3375
         End
         Begin VB.OptionButton RLOG 
            Caption         =   "Receipt Log"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   255
            Left            =   3960
            TabIndex        =   160
            Top             =   6240
            Width           =   2220
         End
         Begin VB.OptionButton cbr 
            Caption         =   "Check Balancing Report"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   255
            Left            =   3960
            TabIndex        =   158
            Top             =   5600
            Width           =   3800
         End
         Begin VB.OptionButton aer 
            Caption         =   "Active Executions Report"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   255
            Left            =   3960
            TabIndex        =   156
            Top             =   4680
            Width           =   3800
         End
         Begin VB.OptionButton aprbdr 
            Caption         =   "All Papers Received"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   255
            Left            =   100
            TabIndex        =   127
            Top             =   240
            Width           =   2535
         End
         Begin VB.OptionButton ealbdr 
            Caption         =   "Execution Assignment List"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   255
            Left            =   105
            TabIndex        =   141
            Top             =   3960
            Width           =   3135
         End
         Begin VB.OptionButton walbdr 
            Caption         =   "Writ Assignment List"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   255
            Left            =   120
            TabIndex        =   134
            Top             =   1650
            Width           =   3015
         End
         Begin VB.OptionButton falbdr 
            Caption         =   "Family Court Assignment List"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   255
            Left            =   120
            TabIndex        =   136
            Top             =   2400
            Width           =   3495
         End
         Begin VB.OptionButton malbdr 
            Caption         =   "Magistrate Assignment List"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   255
            Left            =   120
            TabIndex        =   139
            Top             =   3200
            Width           =   3360
         End
         Begin VB.OptionButton elbdr 
            Caption         =   "Execution Listing (Standard Only)"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   495
            Left            =   120
            TabIndex        =   140
            Top             =   3480
            Width           =   3855
         End
         Begin VB.OptionButton oepbor 
            Caption         =   "Outstanding Execution Papers By Officer Report"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   495
            Left            =   3960
            TabIndex        =   152
            Top             =   2520
            Width           =   3660
         End
         Begin VB.OptionButton cl 
            Caption         =   "Check Listing"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   255
            Left            =   3960
            TabIndex        =   159
            Top             =   5915
            Width           =   2000
         End
         Begin VB.CommandButton rprintbutton 
            Caption         =   "Print Report"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   6495
            TabIndex        =   161
            Top             =   6120
            Width           =   1440
         End
         Begin VB.OptionButton al 
            Caption         =   "Attorney Listing"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   255
            Left            =   3960
            TabIndex        =   157
            Top             =   5160
            Width           =   3800
         End
         Begin VB.OptionButton sdrbdr 
            Caption         =   "Service/Non-Service Detail Report"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   255
            Left            =   120
            TabIndex        =   131
            Top             =   780
            Width           =   3780
         End
         Begin VB.OptionButton mlbdr 
            Caption         =   "Magistrate Listing (Standard or By Court Date)"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   615
            Left            =   100
            TabIndex        =   137
            Top             =   2640
            Width           =   3855
         End
         Begin VB.OptionButton fclbdr 
            Caption         =   "Family Court Listing (Standard or By Court Date)"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   495
            Left            =   120
            TabIndex        =   135
            Top             =   1920
            Width           =   3855
         End
         Begin VB.OptionButton nullar 
            Caption         =   "Nulla Bona Executions Report"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   255
            Left            =   3960
            TabIndex        =   155
            Top             =   4440
            Width           =   3800
         End
         Begin VB.OptionButton ivdsopr 
            Caption         =   "IVD Service of Process Report"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   255
            Left            =   3960
            TabIndex        =   148
            Top             =   600
            Width           =   3525
         End
         Begin VB.OptionButton nmsfrrbdr 
            Caption         =   "Non-Magistrate Service Fee Receipts Report"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   615
            Left            =   120
            TabIndex        =   145
            Top             =   5640
            Width           =   3060
         End
         Begin VB.OptionButton msfrrbdr 
            Caption         =   "Magistrate Service Fee Receipts Report"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   495
            Left            =   120
            TabIndex        =   144
            Top             =   5160
            Width           =   3195
         End
         Begin VB.OptionButton sfrrbdr 
            Caption         =   "Service Fee Receipts Report"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   255
            Left            =   100
            TabIndex        =   143
            Top             =   4800
            Width           =   3375
         End
         Begin VB.OptionButton ompbor 
            Caption         =   "Outstanding Magistrate Papers By Officer Report"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   480
            Left            =   3960
            TabIndex        =   151
            Top             =   1920
            Width           =   3660
         End
         Begin VB.OptionButton opbor 
            Caption         =   "Outstanding Papers by Officer Rpt (excluding Executions)"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   600
            Left            =   3960
            TabIndex        =   150
            Top             =   1320
            Width           =   3800
         End
         Begin VB.OptionButton fcsopr 
            Caption         =   "Family Court Service of Process Rpt"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   360
            Left            =   3960
            TabIndex        =   147
            Top             =   240
            Width           =   3900
         End
         Begin VB.OptionButton erlbdr 
            Caption         =   "Execution Receipts Listing"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   255
            Left            =   100
            TabIndex        =   142
            Top             =   4440
            Width           =   3375
         End
         Begin VB.OptionButton wlbdr 
            Caption         =   "Writ Listing (Standard or By Court Date)"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   495
            Left            =   100
            TabIndex        =   132
            Top             =   1175
            Width           =   3735
         End
         Begin VB.OptionButton opsrbdr 
            Caption         =   "Officer Paper Service Report"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   255
            Left            =   120
            TabIndex        =   129
            Top             =   525
            Width           =   3495
         End
      End
      Begin VB.TextBox sheriff 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   345
         Left            =   -73560
         MaxLength       =   75
         TabIndex        =   99
         Top             =   5760
         Width           =   4350
      End
      Begin VB.TextBox sheriffphone 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   345
         Left            =   -73560
         MaxLength       =   15
         TabIndex        =   98
         Top             =   5310
         Width           =   2145
      End
      Begin VB.TextBox sheriffaddress2 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   345
         Left            =   -73560
         MaxLength       =   75
         TabIndex        =   97
         Top             =   4920
         Width           =   4350
      End
      Begin VB.TextBox sheriffaddress 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   345
         Left            =   -73560
         MaxLength       =   75
         TabIndex        =   96
         Top             =   4560
         Width           =   4350
      End
      Begin MSGrid.Grid outstandinglist 
         Height          =   5055
         Left            =   -74880
         TabIndex        =   179
         Top             =   960
         Width           =   11655
         _Version        =   65536
         _ExtentX        =   20558
         _ExtentY        =   8916
         _StockProps     =   77
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Rows            =   1
         Cols            =   8
         FixedRows       =   0
         FixedCols       =   0
         ScrollBars      =   2
         HighLight       =   0   'False
      End
      Begin VB.Label FROMXREF 
         Height          =   255
         Left            =   -69480
         TabIndex        =   234
         Top             =   3000
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "  Date Range                                                               to"
         Height          =   975
         Left            =   -74880
         TabIndex        =   210
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Labelo 
         Caption         =   $"CIVIL.frx":1F96
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6120
         Left            =   -74880
         TabIndex        =   198
         Top             =   495
         Width           =   11655
      End
      Begin VB.Label Label65 
         BackStyle       =   0  'Transparent
         Caption         =   $"CIVIL.frx":3638
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   3615
         Left            =   -68925
         TabIndex        =   133
         Top             =   1300
         Width           =   1185
      End
      Begin VB.Label Label49 
         BackStyle       =   0  'Transparent
         Caption         =   $"CIVIL.frx":3764
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   1455
         Left            =   -74775
         TabIndex        =   128
         Top             =   1080
         Width           =   5115
      End
      Begin VB.Label Label56 
         BackStyle       =   0  'Transparent
         Caption         =   $"CIVIL.frx":3985
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   2505
         Left            =   -74820
         TabIndex        =   123
         Top             =   4110
         Width           =   1455
      End
      Begin VB.Label Label61 
         BackStyle       =   0  'Transparent
         Caption         =   $"CIVIL.frx":3A8C
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   5805
         Left            =   -74745
         TabIndex        =   130
         Top             =   585
         Width           =   11415
      End
   End
End
Attribute VB_Name = "CIVIL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sofn, soln, somi, soo As String, fee As Single, cnumber, rnumber As String, FROMG As Integer, dbname As String, stopspool As Integer, TP As String, LASTR, LASTD, LASTA As String, FROMP As Integer, nametype As Integer, sfpay As Boolean
Dim itmx As ListItem
Dim SEARCHTYPE As Integer
Dim SAVEERR, CSERVICEOF As Integer
Dim lname, LPHONE As String, ltab, LAFFTYPE As Integer
Dim FROMLF As Integer
Dim totalinterest, totalcommission, totalpayments As Single
Dim procdate As Date
Dim bprincip, bcommiss, BINTer As Single
Dim sedit, sprint, sreport, sbrowse, sdelete, ssupervisor As Integer
Dim lastserviceof As String
Dim paydate(99) As String, payamount(99), payi(99), payc(99), paysf(99) As Single
Dim blnSavePressed As Boolean


      Private Function PtrCtoVbString(add As Long) As String

          Dim sTemp As String * 512, X As Long

          X = lstrcpy(sTemp, add)
          If (InStr(1, sTemp, Chr(0)) = 0) Then
               PtrCtoVbString = ""
          Else
               PtrCtoVbString = Left(sTemp, InStr(1, sTemp, Chr(0)) - 1)
          End If
      End Function

      Private Sub SetDefaultPrinter(ByVal PrinterName As String, _
          ByVal DriverName As String, ByVal PrinterPort As String)
          Dim DeviceLine As String
          Dim r As Long
          Dim l As Long
          DeviceLine = PrinterName & "," & DriverName & "," & PrinterPort
          ' Store the new printer information in the [WINDOWS] section of
          ' the WIN.INI file for the DEVICE= item
          r = WriteProfileString("windows", "Device", DeviceLine)
          ' Cause all applications to reload the INI file:
          l = SendMessage(HWND_BROADCAST, WM_WININICHANGE, 0, "windows")
      End Sub

      Private Sub Win95SetDefaultPrinter(pname As String)
          Dim Handle As Long          'handle to printer
          Dim PrinterName As String
          Dim pd As PRINTER_DEFAULTS
          Dim X As Long
          Dim need As Long            ' bytes needed
          Dim pi5 As PRINTER_INFO_5   ' your PRINTER_INFO structure
          Dim LastError As Long

          ' determine which printer was selected
          PrinterName = pname
          ' none - exit
          If PrinterName = "" Then
              Exit Sub
          End If

          ' set the PRINTER_DEFAULTS members
          pd.pDatatype = 0&
          pd.DesiredAccess = PRINTER_ALL_ACCESS Or pd.DesiredAccess

          ' Get a handle to the printer
          X = OpenPrinter(PrinterName, Handle, pd)
          ' failed the open
          If X = False Then
              'error handler code goes here
              Exit Sub
          End If

          ' Make an initial call to GetPrinter, requesting Level 5
          ' (PRINTER_INFO_5) information, to determine how many bytes
          ' you need
          X = GetPrinter(Handle, 5, ByVal 0&, 0, need)
          ' don't want to check Err.LastDllError here - it's supposed
          ' to fail
          ' with a 122 - ERROR_INSUFFICIENT_BUFFER
          ' redim t as large as you need
          ReDim t((need \ 4)) As Long

          ' and call GetPrinter for keepers this time
          X = GetPrinter(Handle, 5, t(0), need, need)
          ' failed the GetPrinter
          If X = False Then
              'error handler code goes here
              Exit Sub
          End If

          ' set the members of the pi5 structure for use with SetPrinter.
          ' PtrCtoVbString copies the memory pointed at by the two string
          ' pointers contained in the t() array into a Visual Basic string.
          ' The other three elements are just DWORDS (long integers) and
          ' don't require any conversion
          pi5.pPrinterName = PtrCtoVbString(t(0))
          pi5.pPortName = PtrCtoVbString(t(1))
          pi5.Attributes = t(2)
          pi5.DeviceNotSelectedTimeout = t(3)
          pi5.TransmissionRetryTimeout = t(4)

          ' this is the critical flag that makes it the default printer
          pi5.Attributes = PRINTER_ATTRIBUTE_DEFAULT

          ' call SetPrinter to set it
          X = SetPrinter(Handle, 5, pi5, 0)
          ' failed the SetPrinter
          If X = False Then
              MsgBox "SetPrinterFailed. Error code: " & Err.LastDllError
              Exit Sub
          End If

          ' and close the handle
          ClosePrinter (Handle)

      End Sub

      Private Sub GetDriverAndPort(ByVal Buffer As String, DriverName As _
          String, PrinterPort As String)

          Dim iDriver As Integer
          Dim iPort As Integer
          DriverName = ""
          PrinterPort = ""

          'The driver name is first in the string terminated by a comma
          iDriver = InStr(Buffer, ",")
          If iDriver > 0 Then

              'Strip out the driver name
              DriverName = Left(Buffer, iDriver - 1)

              'The port name is the second entry after the driver name
              'separated by commas.
              iPort = InStr(iDriver + 1, Buffer, ",")

              If iPort > 0 Then
                  'Strip out the port name
                  PrinterPort = Mid(Buffer, iDriver + 1, _
                  iPort - iDriver - 1)
              End If
          End If
      End Sub

      Private Sub ParseList(lstCtl As Control, ByVal Buffer As String)
          Dim i As Integer

          Dim s As String

          Do
              i = InStr(Buffer, Chr(0))
              If i > 0 Then
                  s = Left(Buffer, i - 1)
                  If Len(Trim(s)) Then lstCtl.AddItem s
                  Buffer = Mid(Buffer, i + 1)
              Else
                  If Len(Trim(Buffer)) Then lstCtl.AddItem Buffer
                  Buffer = ""
              End If
          Loop While i > 0
      End Sub

      Private Sub WinNTSetDefaultPrinter()
          Dim Buffer As String
          Dim DeviceName As String
          Dim DriverName As String
          Dim PrinterPort As String
          Dim PrinterName As String
          Dim r As Long
              Buffer = Space(1024)
              PrinterName = pname
              r = GetProfileString("PrinterPorts", PrinterName, "", Buffer, Len(Buffer))

              'Parse the driver name and port name out of the buffer
              GetDriverAndPort Buffer, DriverName, PrinterPort

              If DriverName <> "" And PrinterPort <> "" Then
                  SetDefaultPrinter pname, DriverName, PrinterPort
              End If
                End Sub



Private Sub commissandint()
Dim tempdate As String, totsi As Integer
Dim TUN, TOV, TINT1, TINT2 As Single
If Not IsDate(judgementdate) Or Not IsDate(estpayoffdate) Then
    Exit Sub
End If
If CDate(judgementdate) = CDate(estpayoffdate) Then
    Exit Sub
End If
If Val(judgementamount) = 0 Then
    Exit Sub
End If
If Not IsDate(estpayoffdate) Then
    estpayoffdate = Format$(Date$, "yyyy/mm/dd")
Else
    estpayoffdate = Format$(CVDate(estpayoffdate), "yyyy/mm/dd")
End If
Screen.MousePointer = 11
Call setintervals(totsi, estpayoffdate)
perday = 0
INTEREST = 0
totalinterest = 0
totalpayments = 0
totalcommission = 0
commission = 0
sf = 0
total = 0
tb = 0
prevprin = Val(judgementamount)
T2 = Val(judgementamount) * (Val(intrate) / 365)
T2 = Val(Format$(T2, "######0.00"))
T2 = T2 * DateDiff("d", judgementdate, paydate(1))
T2 = Val(Format$(T2, "######0.00"))
If Val(judgementamount) + T2 > Val(exonfirst) Then
    commission = (Val(exonfirst) * Val(excommrate1))
    commission = commission + ((prevprin + T2 - Val(exonfirst)) * Val(excommrate2))
Else
    commission = (prevprin + T2) * Val(excommrate1)
End If
commission = Val(Format$(commission, "####0.00"))
totalcommission = commission
For t% = 1 To totsi
    tempdate = Format$(paydate(t%), "mm/dd/yyyy")
    If (paydate(t%)) <= (estpayoffdate) Then
        T1 = (prevprin * (Val(intrate) * (DateDiff("d", (paydate(t% - 1)), tempdate)) / 365) / DateDiff("d", (paydate(t% - 1)), tempdate))
        perday = T1
        perday = Format(perday, "####0.00")
        T2 = prevprin * (Val(intrate) / 365)
        T2 = Val(Format$(T2, "######0.00"))
        T2 = T2 * DateDiff("d", (paydate(t% - 1)), tempdate)
        T2 = Val(Format$(T2, "######0.00"))
        INTEREST = Val(INTEREST) + T2 - payi(t% - 1)
        totalpayments = totalpayments + payi(t% - 1)
        totalinterest = totalinterest + T2
        TUN = 0
        TOV = 0
        TINT1 = 0
        TINT2 = 0
        If Val(judgementamount) < Val(exonfirst) Then
            If Val(judgementamount) + tb + T2 < Val(exonfirst) Then
                TUN = T2
                TOV = 0
                TINT1 = Val(excommrate1)
                TINT2 = 0
            Else
            If Val(judgementamount) + tb < Val(exonfirst) Then
                TUN = (Val(exonfirst) - Val(judgementamount) - tb)
                TOV = T2 - TUN
                TINT1 = Val(excommrate1)
                TINT2 = Val(excommrate2)
            Else
                TUN = 0
                TOV = T2
                TINT1 = 0
                TINT2 = Val(excommrate2)
            End If
            End If
        Else
            TUN = 0
            TOV = T2
            TINT1 = 0
            TINT2 = Val(excommrate2)
        End If
        tb = tb + T2
        commission = commission + (TUN * TINT1) + (TOV * TINT2)
        totalcommission = totalcommission + (TUN * TINT1) + (TOV * TINT2)
        commission = commission - payc(t% - 1)
        totalpayments = totalpayments + payc(t% - 1)
        prevprin = prevprin - payamount(t%)
        totalpayments = totalpayments + payamount(t% - 1)
        totalpayments = totalpayments + paysf(t% - 1)
        TOTSF = TOTSF + paysf(t% - 1)
    Else
        t% = totsi
    End If
Next t%
If payi(totsi) > 0 Or payc(totsi) > 0 Then
    INTEREST = Val(INTEREST) - payi(totsi)
    commission = commission - payc(totsi)
    totalpayments = totalpayments + payi(totsi)
    totalpayments = totalpayments + payc(totsi)
    totalpayments = totalpayments + payamount(totsi)
    totalpayments = totalpayments + paysf(totsi)
    TOTSF = TOTSF + paysf(totsi)
End If
If prevprin > 0 And CVDate(paydate(t% - 1)) <> CVDate(tempdate) Then
    T1 = (prevprin * (Val(intrate) * (DateDiff("d", (paydate(t% - 1)), tempdate)) / 365) / DateDiff("d", (paydate(t% - 1)), tempdate))
Else
If prevprin = 0 Then
    T1 = 0
End If
End If
perday = T1
perday = Format(perday, "####0.00")

commission = Format(commission, "####0.00")
INTEREST = Format(INTEREST, "####0.00")
possprin = Format$(prevprin, "######0.00")
posscomm = Format$(commission, "######0.00")
possint = Format$(INTEREST, "######0.00")
If sfpay Then
    POSSTOTAL = Val(possprin) + Val(posscomm) + Val(possint) + Val(servicefee) - TOTSF
Else
    POSSTOTAL = Val(possprin) + Val(posscomm) + Val(possint)
End If
POSSTOTAL = Format$(POSSTOTAL, "######0.00")
commission = Format(commission, "####0.00")
balance = Format$(prevprin, "#######0.00")
commission.Refresh
INTEREST.Refresh
total = total - TOTSF
total.Refresh
possprin.Refresh
possint.Refresh
POSSTOTAL.Refresh
balance.Refresh
Screen.MousePointer = 0
End Sub

Private Sub loaddeputy()
assignedto.clear
Dim db As Database, ds As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nwl + "LAWSUITE.MDB")
Set ds = db.OpenRecordset("select profname from professionals where type = 'D'")
On Error Resume Next
If Not ds.EOF Then
    ds.MoveFirst
End If
While Not ds.EOF
    assignedto.AddItem ds("profname")
    ds.MoveNext
Wend
On Error GoTo 0
db.Close
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If

End Sub


Private Sub loaddp()
Dim db As Database, ds As Recordset
Data1.DatabaseName = nwl + "lawsuite.mdb"
Data1.Refresh
Data1.RecordSource = "select dpname from PEOPLE order by dpsort"
Data1.Refresh
serviceof.DataField = dpname
serviceof.Refresh
plaintiff.DataField = dpname
plaintiff.Refresh
defendant.DataField = dpname
defendant.Refresh
On Error GoTo oderror
od:
Set db = OpenDatabase(nwl + "LAWSUITE.MDB")
FIRSTTIME% = 0
dpstart:
db.Close
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If

End Sub

Private Sub loadpapertype()
papertype.clear
papertypelist.clear
Dim db As Database, ds As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nwc + dbname)
Set ds = db.OpenRecordset("select papertype from papers")
If Not ds.EOF Then
    ds.MoveFirst
End If
While Not ds.EOF
    papertype.AddItem ds("papertype")
    papertypelist.AddItem ds("papertype")
    ds.MoveNext
Wend
db.Close
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If

End Sub

Private Sub loadpay()
expaygrid.Rows = 1
expaygrid.Row = 0
For t% = 0 To 8
    expaygrid.Col = t%
    expaygrid.Text = ""
Next t%
On Error GoTo oderror
Dim db As Database, ds As Recordset
od:
Set db = OpenDatabase(nwc + dbname)
Set ds = db.OpenRecordset("select * from executionspay where serviceof = " + Chr$(34) + serviceof + Chr$(34) + " and datereceived = #" + datereceived + "# and iteration = " + Chr$(34) + iteration + Chr$(34) + " order by datepaid desc")
If Not ds.EOF Then
    ds.MoveFirst
End If
While Not ds.EOF
    If Not IsNull(ds("receipt")) Then
        r$ = ds("receipt")
    Else
        r$ = ""
    End If
    If Not IsNull(ds("check")) Then
        c$ = ds("check")
    Else
        c$ = ""
    End If
    If Not IsNull(ds("payrem")) Then
        pr$ = ds("payrem")
    Else
        pr$ = ""
    End If
    If Not IsNull(ds("commiss")) Then
        cc$ = ds("COMMISS")
    Else
        cc$ = ""
    End If
    If Not IsNull(ds("inter")) Then
        ii$ = ds("INTER")
    Else
        ii$ = ""
    End If
    If Not IsNull(ds("principal")) Then
        p$ = ds("principal")
    Else
        p$ = ""
    End If
    If Not IsNull(ds("SERVICEFEE")) Then
        sf$ = ds("SERVICEFEE")
    Else
        sf$ = ""
    End If
    expaygrid.AddItem Str$(ds("datepaid")) + Chr$(9) + Str$(ds("amount")) + Chr$(9) + r$ + Chr$(9) + c$ + Chr$(9) + p$ + Chr$(9) + cc$ + Chr$(9) + ii$ + Chr$(9) + sf$ + Chr$(9) + pr$, expaygrid.Rows - 1
    ds.MoveNext
Wend
Call commissandint
db.Close
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If

End Sub


Private Sub loadprof()
If maintab.Tab = 6 Then
    If omag.Value = True Then
        TP = "M"
    End If
    If oatt.Value = True Then
        TP = "A"
    End If
    If ocou.Value = True Then
        TP = "C"
    End If
    If odep.Value = True Then
        TP = "D"
    End If
End If
If maintab.Tab = 0 Then
    TP = "M"
End If
If maintab.Tab = 1 Then
    TP = "A"
End If
If maintab.Tab = 2 Then
    TP = "C"
End If
If maintab.Tab = 3 Then
    TP = "A"
End If
profname.clear
professional.clear
profaddr1 = ""
profaddr2 = ""
profphone = ""
Dim db As Database, ds As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nwl + "LAWSUITE.MDB")
Set ds = db.OpenRecordset("select profname from professionals where type = " + Chr$(34) + TP + Chr$(34))
If Not ds.EOF Then
    ds.MoveFirst
End If
While Not ds.EOF
    profname.AddItem ds("profname")
    professional.AddItem ds("PROFNAME")
    ds.MoveNext
Wend
db.Close
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If

End Sub


Private Sub loadsystem()
Dim db As Database, ds, rs As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nwc + dbname)
Set ds = db.OpenRecordset("select * from system")
If ds.EOF Then
    msg = MsgBox("Update the Sheriff's Office and Execution Information on the SYSTEM tab.", 48, "Genesis Information Log")
    exintrate = ".14"
    excommrate1 = ".075"
    exonfirst = "500"
    excommrate2 = ".03"
    sheriffaddress = ""
    sheriffaddress2 = ""
    sheriffphone = ""
    sheriff = ""
    treasurer = ""
    treasureraddress1 = ""
    treasureraddress2 = ""
    nextreceipt = "1"
'    prepareprinter = 0
    county = ""
    office = ""
    fnf = False
    lnf = True
'    checkprint = 0
    autoprint = 0
    regularprinter.ListIndex = -1
    moneyprinter.ListIndex = -1
Else
    ds.MoveFirst
    exintrate = ds("exintrate")
    excommrate1 = ds("excommrate1")
    excommrate2 = ds("excommrate2")
    exonfirst = ds("exonfirst")
    sheriffaddress = ds("sheriffaddress")
    sheriffaddress2 = ds("sheriffaddress2")
    sheriffphone = ds("sheriffphone")
    sheriff = ds("sheriff")
    county = ds("county")
    office = ds("office")
    fnf = ds("fnf")
    lnf = ds("lnf")
    If Not IsNull(ds("treasurer")) Then
        treasurer = ds("treasurer")
    Else
        treasurer = ""
    End If
    nextreceipt = ds("startreceipt")
'    prepareprinter = ds("prepareprinter")
    If Not IsNull(ds("treasureraddress1")) Then
        treasureraddress1 = ds("treasureraddress1")
    Else
        treasureraddress1 = ""
    End If
    If Not IsNull(ds("treasureraddress2")) Then
        treasureraddress2 = ds("treasureraddress2")
    Else
        treasureraddress2 = ""
    End If
'    checkprint = ds("checkprint")
    autoprint = ds("autoprint")
    On Error GoTo n3
    Open "rp.dat" For Input As #1
    Line Input #1, rp$
    Close #1
    GoTo n4
n3:
    Open "rp.dat" For Output As #1
    Print #1, ds("regularprinter")
    Close #1
n4:
    If rp$ = "" Then
        rp$ = ds("regularprinter")
    End If
    For t% = 0 To List1.ListCount - 1
        If List1.List(t%) = rp$ Then
            List1.ListIndex = t%
            t% = List1.ListCount
        End If
    Next t%
    On Error GoTo n1
    Open "mp.dat" For Input As #1
    Line Input #1, mp$
    Close #1
    GoTo n2
n1:
    Open "mp.dat" For Output As #1
    Print #1, ds("moneyprinter")
    Close #1
n2:
    If mp$ = "" Then
        mp$ = ds("moneyprinter")
    End If
    For t% = 0 To List2.ListCount - 1
        If List2.List(t%) = mp$ Then
            List2.ListIndex = t%
            t% = List2.ListCount
        End If
    Next t%
End If
If exintrate = 0 And excommrate1 = 0 And excommrate2 = 0 And exonfirst = 0 Then
    msg = MsgBox("Update the Sheriff's Office and Execution Information on the SYSTEM tab.", 48, "Genesis Information Log")
    exintrate = ".14"
    excommrate1 = ".075"
    exonfirst = "500"
    excommrate2 = ".03"
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

Private Sub SENDCHAR(CH As String)
SendKeys CH
End Sub

Private Sub sendclosepara()
SendKeys "{)}"
End Sub
Private Sub senddash()
SendKeys "-"
End Sub

Private Sub sendopenpara()
SendKeys "{(}"
End Sub
Private Sub sendslash()
SendKeys "/"
End Sub

Private Sub sendspace()
SendKeys " "
End Sub

Private Sub setintervals(totsi As Integer, setdate As Date)
Dim holddate As String, holdamount, holdi, holdc As Single
For si% = 0 To 99
    paydate(si%) = ""
    payamount(si%) = 0
    paysf(si%) = 0
    payi(si%) = 0
    payc(si%) = 0
Next si%
totsi = 1
paydate(0) = "0001/01/01"
paydate(1) = Format$(setdate, "yyyy/mm/dd")
a = paydate(2)
a = paydate(3)
a = paydate(4)
a = paydate(5)
a = paydate(6)
For si% = 1 To expaygrid.Rows
    expaygrid.Row = si% - 1
    expaygrid.Col = 0
    If expaygrid.Text = "" Then
        GoTo LOOP5
    End If
    expaygrid.Col = 4
    holdamount = Val(expaygrid.Text)
    expaygrid.Col = 5
    holdc = Val(expaygrid.Text)
    expaygrid.Col = 6
    holdi = Val(expaygrid.Text)
    expaygrid.Col = 7
    holdsf = Val(expaygrid.Text)
    expaygrid.Col = 0
    holddate = Format$(CVDate(expaygrid.Text), "yyyy/mm/dd")
    For si2% = totsi To 0 Step -1
        If holddate > (paydate(si2%)) Then
            For si3% = totsi To si2% + 1 Step -1
                paydate(si3% + 1) = paydate(si3%)
                payamount(si3% + 1) = payamount(si3%)
                paysf(si3% + 1) = paysf(si3%)
                payi(si3% + 1) = payi(si3%)
                payc(si3% + 1) = payc(si3%)
            Next si3%
            paydate(si2% + 1) = holddate
            payamount(si2% + 1) = holdamount
            paysf(si2% + 1) = holdsf
            payi(si2% + 1) = holdi
            payc(si2% + 1) = holdc
            si2% = 0
        Else
        If holddate = (paydate(si2%)) Then
            payamount(si2%) = payamount(si2%) + holdamount
            paysf(si2%) = paysf(si2%) + holdsf
            payi(si2%) = payi(si2%) + holdi
            payc(si2%) = payc(si2%) + holdc
            totsi = totsi - 1
            si2% = 0
        End If
        End If
    Next si2%
    totsi = totsi + 1
LOOP5:
Next si%
paydate(0) = (Format$(judgementdate, "yyyy/mm/dd"))
a = paydate(1)
a = paydate(2)
a = paydate(3)
a = paydate(4)
a = paydate(5)
a = paydate(6)
End Sub

Private Sub accessbutton_Click()
HST = SEARCHTYPE
outstandinglist.Col = 0
If outstandinglist.Text = "Magistrate" Then
    maintab.Tab = 0
End If
If outstandinglist.Text = "Family Court" Then
    maintab.Tab = 2
End If
If outstandinglist.Text = "Writ/Other" Then
    maintab.Tab = 1
End If
If outstandinglist.Text = "Executions" Then
    maintab.Tab = 3
End If
Call clearbutton_Click
outstandinglist.Col = 1
serviceof = outstandinglist.Text
outstandinglist.Col = 2
datereceived = outstandinglist.Text
outstandinglist.Col = 3
iteration = outstandinglist.Text
Call iteration_Click
SEARCHTYPE = HST
End Sub

Private Sub addpay_Click()
If Not IsDate(DATEPAID) Or (IsDate(DATEPAID) And CVDate(DATEPAID) < CVDate(judgementdate)) Then
    msg = MsgBox("An invalid date has been entered.", 48, "Genesis Error Log")
    Exit Sub
End If
If Val(amount) = 0 Then
    msg = MsgBox("Entry in AMOUNT field must be numeric and not equal to zero.", 48, "Genesis Error Log")
    amount.SetFocus
    Exit Sub
End If
If Val(Format$(Val(commiss) + Val(inter) + Val(principal) + Val(eservicefee), "00000000.00")) > Val(amount) And Val(amount) >= 0 Then
    msg = MsgBox("All field amount totals exceed payment amount.", 48, "Genesis Error Log")
    commiss.SetFocus
    Exit Sub
End If
For t% = 1 To expaygrid.Rows
    expaygrid.Row = t% - 1
    expaygrid.Col = 0
    If expaygrid.Text > "" Then
        testdate = CVDate(expaygrid.Text)
        If testdate > CVDate(DATEPAID) Then
            msg = MsgBox("Executions must be entered in chronological order.", 48, "Genesis Error Log")
            Exit Sub
        End If
    End If
Next t%
expaygrid.AddItem DATEPAID + Chr$(9) + amount + Chr$(9) + receipt + Chr$(9) + check + Chr$(9) + principal + Chr$(9) + commiss + Chr$(9) + inter + Chr$(9) + eservicefee + Chr$(9) + remarks, expaygrid.Rows - 1
expaygrid.Row = expaygrid.Rows - 1
If CVDate(DATEPAID) > CVDate(estpayoffdate) Then
    estpayoffdate = DATEPAID
End If
LASTD = DATEPPAID
LASTA = amount
LASTR = receipt
DATEPAID = ""
principal = ""
amount = ""
remarks = ""
receipt = ""
check = ""
commiss = ""
inter = ""
eservicefee = ""
Call commissandint
If autoprint = 1 Then
    FROMP = 1
    If List2.ListCount > 1 And List2.ListIndex > -1 Then
        Call defaultprinter(List2.List(List2.ListIndex))
    End If
    If List1.ListIndex > -1 And List2.ListIndex > -1 Then
        If List1.List(List1.ListIndex) <> List2.List(List2.ListIndex) And prepareprinter = 1 Then
            msg = MsgBox("Prepare for Receipt/Check Printing.", 48, "Genesis Error Log")
        End If
    End If
    If expaygrid.Row < 0 Then
        msg = MsgBox("A payment row must be selected.", 48, "Genesis Error Log")
        Exit Sub
    End If
    expaygrid.Col = 0
    If Not IsDate(expaygrid.Text) Then
        expaygrid.Row = expaygrid.Row - 1
    End If
    expaygrid.Col = 1
    fee = Val(expaygrid.Text)
    expaygrid.Col = 3
    cnumber = expaygrid.Text
    expaygrid.Col = 2
    FROMG = 1
    rnumber = expaygrid.Text
    receiptframe.Left = 1000
    receiptframe.Top = 2000
    Call LOADOTHER
    receiptframe.Visible = True
    If fromdefendant = 0 And fromplaintiff = 0 And othername = "" Then
        fromdefendant = 1
    End If
    Screen.MousePointer = 0
'    othername.SetFocus
End If
DATEPAID.SetFocus
End Sub


Private Sub aeprof_Click()
If Val(frmLogin.CSUPERVISOR(0)) = 1 And Val(frmLogin.CSUPERVISOR(1)) = 1 And Val(frmLogin.CSUPERVISOR(2)) = 1 And Val(frmLogin.CSUPERVISOR(3)) = 1 Then
    a = 1
Else
    msg = MsgBox("Your USER ID does not have sufficient access to perform this task.", 48, "Genesis Information Log")
    Exit Sub
End If
Screen.MousePointer = 11
Dim db As Database, ds As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nwl + "LAWSUITE.MDB")
If omag.Value = True Then
    TP = "M"
End If
If oatt.Value = True Then
    TP = "A"
End If
If ocou.Value = True Then
    TP = "C"
End If
If odep.Value = True Then
    TP = "D"
End If
If profname > "" Then
   Set ds = db.OpenRecordset("select * from professionals where profname = " + Chr$(34) + profname + Chr$(34) + " and type = " + Chr$(34) + TP + Chr$(34))
       If ds.EOF Then
           ds.AddNew
       Else
           ds.MoveFirst
           ds.Edit
       End If
       ds("profname") = profname
       ds("profaddr1") = profaddr1
       ds("profaddr2") = profaddr2
       ds("profphone") = profphone
       ds("type") = TP
       ds.Update
End If
On Error GoTo 0
db.Close
Call loadprof
Screen.MousePointer = 0
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If

End Sub






Private Sub alllist_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
alllist.SortKey = ColumnHeader.index - 1
If alllist.SortOrder = lvwAscending Then
    alllist.SortOrder = lvwDescending
Else
    alllist.SortOrder = lvwAscending
End If
alllist.Sorted = True
End Sub

Private Sub alllist_ItemClick(ByVal Item As MSComctlLib.ListItem)
Set itmx = alllist.ListItems(alllist.SelectedItem.index)
If itmx = "Magistrate" Then
    maintab.Tab = 0
End If
If itmx = "Family Court" Then
    maintab.Tab = 2
End If
If itmx = "Writ/Other" Then
    maintab.Tab = 1
End If
If itmx = "Executions" Then
    maintab.Tab = 3
End If
Call clearbutton_Click
serviceof = itmx.SubItems(1)
datereceived = itmx.SubItems(2)
iteration = itmx.SubItems(3)
Call iteration_Click
procdate = "12/31/9999"
indexframe.Visible = False
sohomeaddress.SetFocus

End Sub

Private Sub AMOUNT_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    receipt.SetFocus
End If

End Sub

Private Sub apptdate_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
If Len(apptdate) = 1 Or Len(apptdate) = 4 Then
    Call sendslash
End If
End If
If maintab.Tab = 3 Then
If KeyAscii = 13 Then
    intrate.SetFocus
End If
End If

End Sub


Private Sub arcbutton_Click()
If Val(frmLogin.CSUPERVISOR(0)) = 1 And Val(frmLogin.CSUPERVISOR(1)) = 1 And Val(frmLogin.CSUPERVISOR(2)) = 1 And Val(frmLogin.CSUPERVISOR(3)) = 1 Then
    a = 1
Else
    msg = MsgBox("Your USER ID does not have sufficient access to perform this task.", 48, "Genesis Information Log")
    Exit Sub
End If
back6$ = DateAdd("m", -6, Date$)
inp = InputBox("Enter the last date of completed papers to be kept.  (All served and non-service papers received prior to that date will be moved to the archive file.  NOTE: Only paid executions are eligible for archive.)", "Genesis Information Log", back6$)
If Not IsDate(inp) Then
    msg = MsgBox("Date entered for archive was not valid.  Process halted.", 48, "Genesis Error Log")
    Exit Sub
End If
Screen.MousePointer = 11
Dim db As Database, db2 As Database, ds As Recordset, ds2 As Recordset, ds3 As Recordset
Set db = OpenDatabase(nwc + "CIVIL.MDB")
Set db2 = OpenDatabase(nwc + "arccivil.mdb")
ct% = 0
arccount.Visible = True
Set ds = db.OpenRecordset("select * from magistrate where datereceived < #" + inp + "# and (served = '1' or nonservice = '1')")
If Not ds.EOF Then
    ds.MoveFirst
End If
While Not ds.EOF
    Set ds2 = db2.OpenRecordset("select * from magistrate where serviceof = " + Chr$(34) + ds("serviceof") + Chr$(34) + " and datereceived = #" + Str$(ds("datereceived")) + "# and iteration = " + Chr$(34) + ds("iteration") + Chr$(34))
    If ds2.EOF Then
        ds2.AddNew
    Else
        ds2.MoveFirst
        ds2.Edit
    End If
    For yy% = 0 To ds.Fields.Count - 1
        ds2(yy%) = ds(yy%)
    Next yy%
    ct% = ct% + 1
    arccount.Caption = Mid$(Str$(ct%), 2)
    arccount.Refresh
    ds.Delete
    ds.MoveNext
Wend
Set ds = db.OpenRecordset("select * from familycourt where datereceived < #" + inp + "# and (served = '1' or nonservice = '1')")
If Not ds.EOF Then
    ds.MoveFirst
End If
While Not ds.EOF
    Set ds2 = db2.OpenRecordset("select * from familycourt where serviceof = " + Chr$(34) + ds("serviceof") + Chr$(34) + " and datereceived = #" + Str$(ds("datereceived")) + "# and iteration = " + Chr$(34) + ds("iteration") + Chr$(34))
    If ds2.EOF Then
        ds2.AddNew
    Else
        ds2.MoveFirst
        ds2.Edit
    End If
    For yy% = 0 To ds.Fields.Count - 1
        ds2(yy%) = ds(yy%)
    Next yy%
    ds2.Update
    ct% = ct% + 1
    arccount.Caption = Mid$(Str$(ct%), 2)
    arccount.Refresh
    ds.Delete
    ds.MoveNext
Wend
Set ds = db.OpenRecordset("select * from writother where datereceived < #" + inp + "# and (served = '1' or nonservice = '1')")
If Not ds.EOF Then
    ds.MoveFirst
End If
While Not ds.EOF
    Set ds2 = db2.OpenRecordset("select * from writother where serviceof = " + Chr$(34) + ds("serviceof") + Chr$(34) + " and datereceived = #" + Str$(ds("datereceived")) + "# and iteration = " + Chr$(34) + ds("iteration") + Chr$(34))
    If ds2.EOF Then
        ds2.AddNew
    Else
        ds2.MoveFirst
        ds2.Edit
    End If
    For yy% = 0 To ds.Fields.Count - 1
        ds2(yy%) = ds(yy%)
    Next yy%
    ds2.Update
    ct% = ct% + 1
    arccount.Caption = Mid$(Str$(ct%), 2)
    arccount.Refresh
    ds.Delete
    ds.MoveNext
Wend
Set ds = db.OpenRecordset("select * from executions where datereceived < #" + inp + "# and (served = '1' or nonservice = '1') and balance = 0")
If Not ds.EOF Then
    ds.MoveFirst
End If
While Not ds.EOF
    Set ds2 = db2.OpenRecordset("select * from executions where serviceof = " + Chr$(34) + ds("serviceof") + Chr$(34) + " and datereceived = #" + Str$(ds("datereceived")) + "# and iteration = " + Chr$(34) + ds("iteration") + Chr$(34))
    If ds2.EOF Then
        ds2.AddNew
    Else
        ds2.MoveFirst
        ds2.Edit
    End If
    For yy% = 0 To ds.Fields.Count - 1
        ds2(yy%) = ds(yy%)
    Next yy%
    
    ds2.Update
    Set ds3 = db.OpenRecordset("select * from executionspay where serviceof = " + Chr$(34) + ds("serviceof") + Chr$(34) + " and datereceived = #" + Str$(ds("datereceived")) + "# and iteration = " + Chr$(34) + ds("iteration") + Chr$(34))
    If Not ds3.EOF Then
        ds3.MoveFirst
    End If
    While Not ds3.EOF
        For yy% = 0 To ds3.Fields.Count - 1
            ds2(yy%) = ds3(yy%)
        Next yy%
        ds2.Update
        ds3.Delete
        ds3.MoveNext
    Wend
    ct% = ct% + 1
    arccount.Caption = Mid$(Str$(ct%), 2)
    arccount.Refresh
    ds.Delete
    ds.MoveNext
Wend
Screen.MousePointer = 0
arccount.Visible = False
db.Close

    
    
End Sub

Private Sub assignedon_GotFocus()
If KeyAscii <> 8 Then
If assignedto > "" And assignedon = "" Then
    assignedon = Format$(Date$, "mm/dd/yyyy")
End If
End If
End Sub



Private Sub assignedon_KeyPress(KeyAscii As Integer)
If Len(assignedon) = 1 Or Len(assignedon) = 4 Then
    Call sendslash
End If
If KeyAscii = 13 Then
    served.SetFocus
End If

End Sub


Private Sub assignedto_Click()
infoframe.Refresh

End Sub



Private Sub assignedto_keypress(KeyAscii As Integer)
If KeyAscii = 34 Then
    msg = MsgBox("The " + Chr$(34) + " character is not allowed.  Use the ' character instead.", 48, "Genesis Error Log")
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
    assignedon.SetFocus
End If

End Sub

Private Sub autoprint_Click()
Dim db As Database, rs As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nwc + dbname)
Set rs = db.OpenRecordset("select autoprint from system")
If rs.EOF Then
    rs.AddNew
Else
    rs.MoveFirst
    rs.Edit
End If
rs("autoprint") = autoprint.Value
rs.Update
On Error GoTo 0
db.Close
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If


End Sub

Private Sub balance_Change()
If sfpay Then
    total = Str$(Val(commission) + Val(balance) + Val(INTEREST) + Val(servicefee))
Else
    total = Str$(Val(commission) + Val(balance) + Val(INTEREST))
End If
End Sub

Private Sub bill_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   receiptd.SetFocus
End If

End Sub

Private Sub casenumber_KeyPress(KeyAscii As Integer)
If KeyAscii = 34 Then
    msg = MsgBox("The " + Chr$(34) + " character is not allowed.  Use the ' character instead.", 48, "Genesis Error Log")
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
    papertype.SetFocus
End If

End Sub

Private Sub CHECK_KeyPress(KeyAscii As Integer)
If KeyAscii = 34 Then
    msg = MsgBox("The " + Chr$(34) + " character is not allowed.  Use the ' character instead.", 48, "Genesis Error Log")
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
    principal.SetFocus
End If

End Sub

Private Sub checkd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    professional.SetFocus
End If
End Sub

Private Sub checkprint_Click()
Dim db As Database, rs As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nwc + dbname)
Set rs = db.OpenRecordset("select checkprint from system")
If rs.EOF Then
    rs.AddNew
Else
    rs.MoveFirst
    rs.Edit
End If
rs("checkprint") = checkprint.Value
rs.Update
db.Close
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If


End Sub

Private Sub closebutton_Click()
likeframe.Visible = False
End Sub

Private Sub closepay_Click()
paymentframe.Visible = False
DATEPAID = ""
principal = ""
amount = ""
remarks = ""
eservicefee = ""
receipt = ""
check = ""
commiss = ""
inter = ""
Call commissandint
End Sub

Private Sub colorchange_Click()
colorchg.Show
End Sub

Private Sub Command1_Click()
Screen.MousePointer = 11
If Val(frmLogin.CSUPERVISOR(3)) = 1 Then
    Dim db As Database, ds As Recordset
    On Error GoTo oderror
od:
    Set db = OpenDatabase(nwc + dbname)
    Set ds = db.OpenRecordset("select * from system")
    If ds.EOF Then
        ds.AddNew
    Else
        ds.MoveFirst
        ds.Edit
    End If
    ds("exintrate") = exintrate
    ds("excommrate1") = excommrate1
    ds("excommrate2") = excommrate2
    ds("exonfirst") = exonfirst
    ds.Update
Else
    msg = MsgBox("You have insufficient authority for this operation.", 48, "Genesis Error Log")
End If
On Error GoTo 0
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

Private Sub closeprint_Click()
mprintframe.Visible = False
If SEARCHTYPE = 1 Then
    SEARCHTYPE = 0
    maintab.Tab = 5
Else
    SEARCHTYPE = 0
End If
End Sub

Private Sub Command10_Click()
End Sub

Private Sub Command11_Click()
procdate = "12/31/9999"
indexframe.Visible = False
End Sub

Private Sub Command12_Click()
Call INDEXBUTTON_Click
End Sub

Private Sub Command13_Click()
receiptframe.Visible = False
If FROMG = 0 Then
    Call clearbutton_Click
End If
FROMG = 0

End Sub

Private Sub Command14_Click()
SEARCHTYPE = 1
    Labelo = "SEARCH ALL RESULTS" + Mid$(Labelo, 19)
    If sname = "" And scasenumber = "" Then
        msg = MsgBox("Either name or case number fields must have entry.", 48, "Genesis Error Log")
        Exit Sub
    End If
    While Right$(sname, 1) = " "
        sname = Left$(sname, Len(sname) - 1)
    Wend
    If InStr(sname, " ") > 0 Then
        SNAME1$ = Left$(sname, InStr(sname, " ") - 1)
        SNAME2$ = Mid$(sname, InStr(sname, " ") + 1)
    Else
        SNAME1$ = sname
        SNAME2$ = ""
    End If
    Screen.MousePointer = 11
    outstandinglist.Rows = 1
    For t% = 0 To 7
        outstandinglist.Col = t%
        outstandinglist.Text = ""
    Next t%
    outstandinglist.ColWidth(0) = 900
    outstandinglist.ColWidth(1) = 2400
    outstandinglist.ColWidth(2) = 1000
    outstandinglist.ColWidth(3) = 400
    outstandinglist.ColWidth(4) = 2000
    outstandinglist.ColWidth(5) = 2000
    outstandinglist.ColWidth(6) = 1500
    outstandinglist.ColWidth(7) = 1100
    On Error GoTo oderror
od:
    Set db = OpenDatabase(nwc + dbname)
    If sname > "" Then
        If scasenumber > "" Then
            If InStr(sname, " ") > 0 Then
                Set ds = db.OpenRecordset("select CASENUMBER, serviceof,datereceived,iteration,DEFENDANT, PLAINTIFF,papertype from magistrate where ((serviceof like '*" + SNAME1$ + "*' AND SERVICEOF LIKE '*" + SNAME2$ + "*') or (defendant like '*" + SNAME1$ + "*' AND DEFENDANT LIKE '*" + SNAME2$ + "*') or (plaintiff like '*" + SNAME1$ + "*' AND PLAINTIFF LIKE '*" + SNAME2$ + "*')) and casenumber like '*" + scasenumber + "*' order by datereceived,SERVICEOF,iteration")
            Else
                Set ds = db.OpenRecordset("select CASENUMBER, serviceof,datereceived,iteration,DEFENDANT, PLAINTIFF,papertype from magistrate where (serviceof like '*" + sname + "*' or defendant like '*" + sname + "*' or plaintiff like '*" + sname + "*') and casenumber like '*" + scasenumber + "*' order by datereceived,SERVICEOF,iteration")
            End If
        Else
            Set ds = db.OpenRecordset("select CASENUMBER, serviceof,datereceived,iteration,DEFENDANT, PLAINTIFF,papertype from magistrate where (serviceof like '*" + SNAME1$ + "*' AND SERVICEOF LIKE '*" + SNAME2$ + "*') or (defendant like '*" + SNAME1$ + "*' AND DEFENDANT LIKE '*" + SNAME2$ + "*') or (plaintiff like '*" + SNAME1$ + "*' AND PLAINTIFF LIKE '*" + SNAME2$ + "*') order by datereceived,SERVICEOF,iteration")
        End If
    Else
    If scasenumber > "" Then
        Set ds = db.OpenRecordset("select CASENUMBER, serviceof,datereceived,iteration,DEFENDANT, PLAINTIFF,papertype from magistrate where casenumber like '*" + scasenumber + "*' order by datereceived,SERVICEOF,iteration")
    End If
    End If
    ct% = 0
    If Not ds.EOF Then
        ds.MoveFirst
    End If
    While Not ds.EOF
        outstandinglist.AddItem "Magistrate" + Chr$(9) + ds("serviceof") + Chr$(9) + Str$(ds("datereceived")) + Chr$(9) + ds("iteration") + Chr$(9) + ds("plaintiff") + Chr$(9) + ds("defendant") + Chr$(9) + ds("papertype") + Chr$(9) + ds("casenumber"), ct%
        outstandinglist.RowHeight(ct%) = 400
        ct% = ct% + 1
        ds.MoveNext
    Wend
    If sname > "" Then
        If scasenumber > "" Then
            If InStr(sname, " ") > 0 Then
                Set ds = db.OpenRecordset("select CASENUMBER, serviceof,datereceived,iteration,DEFENDANT, PLAINTIFF,papertype from FAMILYCOURT where ((serviceof like '*" + SNAME1$ + "*' AND SERVICEOF LIKE '*" + SNAME2$ + "*') or (defendant like '*" + SNAME1$ + "*' AND DEFENDANT LIKE '*" + SNAME2$ + "*') or (plaintiff like '*" + SNAME1$ + "*' AND PLAINTIFF LIKE '*" + SNAME2$ + "*')) and casenumber like '*" + scasenumber + "*' order by datereceived,SERVICEOF,iteration")
            Else
                Set ds = db.OpenRecordset("select CASENUMBER, serviceof,datereceived,iteration,DEFENDANT, PLAINTIFF,papertype from FAMILYCOURT where (serviceof like '*" + sname + "*' or defendant like '*" + sname + "*' or plaintiff like '*" + sname + "*') and casenumber like '*" + scasenumber + "*' order by datereceived,SERVICEOF,iteration")
            End If
        Else
            Set ds = db.OpenRecordset("select CASENUMBER, serviceof,datereceived,iteration,DEFENDANT, PLAINTIFF,papertype from FAMILYCOURT where (serviceof like '*" + SNAME1$ + "*' AND SERVICEOF LIKE '*" + SNAME2$ + "*') or (defendant like '*" + SNAME1$ + "*' AND DEFENDANT LIKE '*" + SNAME2$ + "*') or (plaintiff like '*" + SNAME1$ + "*' AND PLAINTIFF LIKE '*" + SNAME2$ + "*') order by datereceived,SERVICEOF,iteration")
        End If
    Else
    If scasenumber > "" Then
        Set ds = db.OpenRecordset("select CASENUMBER, serviceof,datereceived,iteration,DEFENDANT, PLAINTIFF,papertype from familycourt where casenumber like '*" + scasenumber + "*' order by datereceived,SERVICEOF,iteration")
    End If
    End If
    ct% = 0
    If Not ds.EOF Then
        ds.MoveFirst
    End If
    While Not ds.EOF
        outstandinglist.AddItem "Family Court" + Chr$(9) + ds("serviceof") + Chr$(9) + Str$(ds("datereceived")) + Chr$(9) + ds("iteration") + Chr$(9) + ds("plaintiff") + Chr$(9) + ds("defendant") + Chr$(9) + ds("papertype") + Chr$(9) + ds("casenumber"), ct%
        outstandinglist.RowHeight(ct%) = 400
        ct% = ct% + 1
        ds.MoveNext
    Wend
    If sname > "" Then
        If scasenumber > "" Then
            If InStr(sname, " ") > 0 Then
                Set ds = db.OpenRecordset("select CASENUMBER, serviceof,datereceived,iteration,DEFENDANT, PLAINTIFF,papertype from WRITOTHER where ((serviceof like '*" + SNAME1$ + "*' AND SERVICEOF LIKE '*" + SNAME2$ + "*') or (defendant like '*" + SNAME1$ + "*' AND DEFENDANT LIKE '*" + SNAME2$ + "*') or (plaintiff like '*" + SNAME1$ + "*' AND PLAINTIFF LIKE '*" + SNAME2$ + "*')) and casenumber like '*" + scasenumber + "*' order by datereceived,SERVICEOF,iteration")
            Else
                Set ds = db.OpenRecordset("select CASENUMBER, serviceof,datereceived,iteration,DEFENDANT, PLAINTIFF,papertype from WRITOTHER where (serviceof like '*" + sname + "*' or defendant like '*" + sname + "*' or plaintiff like '*" + sname + "*') and casenumber like '*" + scasenumber + "*' order by datereceived,SERVICEOF,iteration")
            End If
        Else
            Set ds = db.OpenRecordset("select CASENUMBER, serviceof,datereceived,iteration,DEFENDANT, PLAINTIFF,papertype from WRITOTHER where (serviceof like '*" + SNAME1$ + "*' AND SERVICEOF LIKE '*" + SNAME2$ + "*') or (defendant like '*" + SNAME1$ + "*' AND DEFENDANT LIKE '*" + SNAME2$ + "*') or (plaintiff like '*" + SNAME1$ + "*' AND PLAINTIFF LIKE '*" + SNAME2$ + "*') order by datereceived,SERVICEOF,iteration")
        End If
    Else
    If scasenumber > "" Then
        Set ds = db.OpenRecordset("select CASENUMBER, serviceof,datereceived,iteration,DEFENDANT, PLAINTIFF,papertype from writother where casenumber like '*" + scasenumber + "*' order by datereceived,SERVICEOF,iteration")
    End If
    End If
    ct% = 0
    If Not ds.EOF Then
        ds.MoveFirst
    End If
    While Not ds.EOF
        outstandinglist.AddItem "Writ/Other" + Chr$(9) + ds("serviceof") + Chr$(9) + Str$(ds("datereceived")) + Chr$(9) + ds("iteration") + Chr$(9) + ds("plaintiff") + Chr$(9) + ds("defendant") + Chr$(9) + ds("papertype") + Chr$(9) + ds("casenumber"), ct%
        outstandinglist.RowHeight(ct%) = 400
        ct% = ct% + 1
        ds.MoveNext
    Wend
    If sname > "" Then
        If scasenumber > "" Then
            If InStr(sname, " ") > 0 Then
                Set ds = db.OpenRecordset("select CASENUMBER, serviceof,datereceived,iteration,DEFENDANT, PLAINTIFF,papertype from EXECUTIONS where ((serviceof like '*" + SNAME1$ + "*' AND SERVICEOF LIKE '*" + SNAME2$ + "*') or (defendant like '*" + SNAME1$ + "*' AND DEFENDANT LIKE '*" + SNAME2$ + "*') or (plaintiff like '*" + SNAME1$ + "*' AND PLAINTIFF LIKE '*" + SNAME2$ + "*')) and casenumber like '*" + scasenumber + "*' order by datereceived,SERVICEOF,iteration")
            Else
                Set ds = db.OpenRecordset("select CASENUMBER, serviceof,datereceived,iteration,DEFENDANT, PLAINTIFF,papertype from EXECUTIONS where (serviceof like '*" + sname + "*' or defendant like '*" + sname + "*' or plaintiff like '*" + sname + "*') and casenumber like '*" + scasenumber + "*' order by datereceived,SERVICEOF,iteration")
            End If
        Else
            Set ds = db.OpenRecordset("select CASENUMBER, serviceof,datereceived,iteration,DEFENDANT, PLAINTIFF,papertype from EXECUTIONS where (serviceof like '*" + SNAME1$ + "*' AND SERVICEOF LIKE '*" + SNAME2$ + "*') or (defendant like '*" + SNAME1$ + "*' AND DEFENDANT LIKE '*" + SNAME2$ + "*') or (plaintiff like '*" + SNAME1$ + "*' AND PLAINTIFF LIKE '*" + SNAME2$ + "*') order by datereceived,SERVICEOF,iteration")
        End If
    Else
    If scasenumber > "" Then
        Set ds = db.OpenRecordset("select CASENUMBER, serviceof,datereceived,iteration,DEFENDANT, PLAINTIFF,papertype from executions where casenumber like '*" + scasenumber + "*' order by datereceived,SERVICEOF,iteration")
    End If
    End If
    ct% = 0
    If Not ds.EOF Then
        ds.MoveFirst
    End If
    While Not ds.EOF
        outstandinglist.AddItem "Executions" + Chr$(9) + ds("serviceof") + Chr$(9) + Str$(ds("datereceived")) + Chr$(9) + ds("iteration") + Chr$(9) + ds("plaintiff") + Chr$(9) + ds("defendant") + Chr$(9) + ds("papertype") + Chr$(9) + ds("casenumber"), ct%
        outstandinglist.RowHeight(ct%) = 400
        ct% = ct% + 1
        ds.MoveNext
    Wend
    If outstandinglist.Rows > 1 Then
        outstandinglist.Rows = outstandinglist.Rows - 1
    End If
    outstandinglist.Row = 0
    outstandinglist.Col = 0
    Screen.MousePointer = 0
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

Private Sub Command15_Click()
Dim db As Database, rs As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nwc + dbname)
linect% = 0
pgct% = 0
If List2.ListCount > 1 And List2.ListIndex > -1 Then
    Call defaultprinter(List2.List(List2.ListIndex))
End If
If List1.ListIndex > -1 And List2.ListIndex > -1 Then
    If List1.List(List1.ListIndex) <> List2.List(List2.ListIndex) And prepareprinter = 1 Then
        msg = MsgBox("Prepare for Receipt/Check Printing.", 48, "Genesis Error Log")
    End If
End If
GoSub header
Set rs = db.OpenRecordset("select papertype from papers order by papertype")
If Not rs.EOF Then
    rs.MoveFirst
    While Not rs.EOF
        Printer.Print rs("papertype")
        linect% = linect% + 1
        rs.MoveNext
        If linect% > 50 Then
            Printer.NewPage
            GoSub header
        End If
    Wend
End If
db.Close
Printer.EndDoc
On Error GoTo 0
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If

header:
pgct% = pgct% + 1
Printer.FontName = "Times New Roman"
Printer.FontSize = 18
Printer.FontBold = True
Printer.Print "Paper Type Listing"
Printer.FontSize = 12
Printer.Print Date$; Tab(100); "Page: "; pgct%
Printer.Print
Printer.Print
Printer.Print
Printer.FontBold = False
linect% = 5
Return
End Sub

Private Sub Command2_Click()
If Val(frmLogin.CSUPERVISOR(0)) = 1 And Val(frmLogin.CSUPERVISOR(1)) = 1 And Val(frmLogin.CSUPERVISOR(2)) = 1 And Val(frmLogin.CSUPERVISOR(3)) = 1 Then
    a = 1
Else
    msg = MsgBox("Your USER ID does not sufficient authority for this operation.", 48, "Genesis Error Log")
    Exit Sub
End If
If papertypelist = "" Then
    Exit Sub
End If
Dim db As Database, rs As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nwc + dbname)
Set rs = db.OpenRecordset("select papertype from papers where papertype = " + Chr$(34) + papertypelist + Chr$(34))
If Not rs.EOF Then
    rs.MoveFirst
    While Not rs.EOF
        rs.Delete
        rs.MoveNext
    Wend
    papertypelist.RemoveItem papertypelist.ListIndex
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

Private Sub Command3_Click()
If Val(frmLogin.CSUPERVISOR(0)) = 1 And Val(frmLogin.CSUPERVISOR(1)) = 1 And Val(frmLogin.CSUPERVISOR(2)) = 1 And Val(frmLogin.CSUPERVISOR(3)) = 1 Then
    a = 1
Else
    msg = MsgBox("Your USER ID does not sufficient authority for this operation.", 48, "Genesis Error Log")
    Exit Sub
End If
If papertypelist = "" Then
    Exit Sub
End If
Dim db As Database, rs As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nwc + dbname)
Set rs = db.OpenRecordset("select papertype from papers where papertype = " + Chr$(34) + papertypelist + Chr$(34))
If rs.EOF Then
    rs.AddNew
    rs("papertype") = papertypelist
    papertypelist.AddItem papertypelist
    rs.Update
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

Private Sub Command4_Click()
If List2.ListCount > 1 And List2.ListIndex > -1 Then
    Call defaultprinter(List2.List(List2.ListIndex))
End If
If List1.ListIndex > -1 And List2.ListIndex > -1 Then
    If List1.List(List1.ListIndex) <> List2.List(List2.ListIndex) And prepareprinter = 1 Then
        msg = MsgBox("Prepare for Receipt/Check Printing.", 48, "Genesis Error Log")
    End If
End If
If expaygrid.Row < 0 Then
    msg = MsgBox("A payment row must be selected.", 48, "Genesis Error Log")
    Exit Sub
End If
expaygrid.Col = 1
fee = Val(expaygrid.Text)
If fee = 0 Then
    msg = MsgBox("A valid payment row must be selected.", 48, "Genesis Error Log")
    Exit Sub
End If
expaygrid.Col = 3
cnumber = expaygrid.Text
expaygrid.Col = 2
FROMG = 1
FROMP = 1
rnumber = expaygrid.Text
receiptframe.Left = 1000
receiptframe.Top = 2000
Call LOADOTHER
If fromdefendant = 0 And fromplaintiff = 0 And othername = "" Then
    fromdefendant = 1
End If
receiptframe.Visible = True
Screen.MousePointer = 0
'othername.SetFocus
End Sub

Private Sub Command5_Click()
If othername > "" Then
    fromdefendant = 0
    fromplaintiff = 0
End If
Dim dt As String
If FROMP = 1 Then
    expaygrid.Col = 1
    If expaygrid.Row = expaygrid.Rows Then
        expaygrid.Row = expaygrid.Row - 1
    End If
    fee = Val(expaygrid.Text)
    expaygrid.Col = 3
    cnumber = expaygrid.Text
    expaygrid.Col = 2
    rnumber = expaygrid.Text
    expaygrid.Col = 0
    dt = expaygrid.Text
Else
    If Not IsDate(feedate) Then
        feedate = datereceived
    End If
    dt = feedate
End If
If fromdefendant = 1 Then
    Call printreceipt(fee, rnumber, cnumber, defendant, dhomeaddress, dhomeaddress2, dhomestate, dhomezipcode, dt)
    Call printreceipt(fee, rnumber, cnumber, defendant, dhomeaddress, dhomeaddress2, dhomestate, dhomezipcode, dt)
Else
If fromplaintiff = 1 Then
    Call printreceipt(fee, rnumber, cnumber, plaintiff, phomeaddress, phomeaddress2, phomestate, phomezipcode, dt)
    Call printreceipt(fee, rnumber, cnumber, plaintiff, phomeaddress, phomeaddress2, phomestate, phomezipcode, dt)
Else
    Call printreceipt(fee, rnumber, cnumber, othername, otheraddress1, otheraddress2, "", "", dt)
    Call printreceipt(fee, rnumber, cnumber, othername, otheraddress1, otheraddress2, "", "", dt)
End If
End If
receiptframe.Visible = False
If FROMG = 0 Then
    Call clearbutton_Click
End If
FROMG = 0
FROMP = 0
Screen.MousePointer = 0
othername = ""
otheraddress1 = ""
otheraddress2 = ""
End Sub

Private Sub Command6_Click()
Screen.MousePointer = 11
If Val(frmLogin.CPRINT(0)) = 1 And Val(frmLogin.CPRINT(1)) = 1 And Val(frmLogin.CPRINT(2)) = 1 And Val(frmLogin.CPRINT(3)) = 1 Then
    Dim db As Database, rs As Recordset
    On Error GoTo oderror
od:
    Set db = OpenDatabase(nwc + dbname)
    Set ds = db.OpenRecordset("select * from system")
    If ds.EOF Then
        ds.AddNew
    Else
        ds.MoveFirst
        ds.Edit
    End If
    If treasurer > "" Then
        ds("treasurer") = treasurer
        ds("treasureraddress1") = treasureraddress1
        ds("treasureraddress2") = treasureraddress2
    End If
    ds("startreceipt") = Val(nextreceipt)
    ds.Update
Else
    msg = MsgBox("You have insufficient authority for this operation.", 48, "Genesis Error Log")
End If
On Error GoTo 0
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

Private Sub Command7_Click()
If Val(frmLogin.CSUPERVISOR(0)) = 0 Or Val(frmLogin.CSUPERVISOR(1)) = 0 Or Val(frmLogin.CSUPERVISOR(2)) = 0 Or Val(frmLogin.CSUPERVISOR(3)) = 0 Then
    msg = MsgBox("You have insufficient access authority for this function.", 48, "Genesis Error Log")
    Exit Sub
End If
If Not IsDate(fromdate) Or Not IsDate(todate) Then
    msg = MsgBox("A valid date range must be entered.", 48, "Genesis Error Log")
    Exit Sub
End If
If List2.ListCount > 1 And List2.ListIndex > -1 Then
    Call defaultprinter(List2.List(List2.ListIndex))
End If
If List1.ListIndex > -1 And List2.ListIndex > -1 Then
    If List1.List(List1.ListIndex) <> List2.List(List2.ListIndex) And prepareprinter = 1 Then
        msg = MsgBox("Prepare for Receipt/Check Printing.", 48, "Genesis Error Log")
    End If
End If
Dim fd, td As Date
fd = CVDate(fromdate)
td = CVDate(todate)
If fd > td Then
    msg = MsgBox("From Date cannot be later than To Date", 48, "Genesis Error Log")
    Exit Sub
End If
Dim db, db2 As Database, rs As Recordset, rs2, rs3 As Recordset, totamt As Single, b, hm, tm, m, hth, tth, th, h, t, o, cents As Integer, buildword As String
totamt = 0
buildword = ""
On Error GoTo oderror
od:
Set db = OpenDatabase(nwc + dbname)
If Sfc Then
    Set rs = db.OpenRecordset("Select sum(servicefee) as sf from magistrate where feedate between #" + fromdate + "# and #" + todate + "#")
    If Not rs.EOF Then
        rs.MoveFirst
        If Not IsNull(rs("sf")) Then
            totamt = totamt + rs("sf")
        End If
    End If
    Set rs = db.OpenRecordset("Select sum(servicefee) as sf from familycourt where feedate between #" + fromdate + "# and #" + todate + "#")
    If Not rs.EOF Then
        rs.MoveFirst
        If Not IsNull(rs("sf")) Then
            totamt = totamt + rs("sf")
        End If
    End If
    Set rs = db.OpenRecordset("Select sum(servicefee) as sf from writother where feedate between #" + fromdate + "# and #" + todate + "#")
    If Not rs.EOF Then
        rs.MoveFirst
        If Not IsNull(rs("sf")) Then
            totamt = totamt + rs("sf")
        End If
    End If
    Set rs = db.OpenRecordset("Select sum(servicefee) as sf from executions where feedate between #" + fromdate + "# and #" + todate + "#")
    If Not rs.EOF Then
        rs.MoveFirst
        If Not IsNull(rs("sf")) Then
            totamt = totamt + rs("sf")
        End If
    End If
    If totamt = 0 Then
        msg = MsgBox("No service fees found for dates specified.", 48, "Genesis Error Log")
        db.Close
        Exit Sub
    End If
    Set rs3 = db.OpenRecordset("select treasurer, treasureraddress1, treasureraddress2 from system")
    If rs3.EOF Then
        msg = MsgBox("No treasurer information found on System tab.", 48, "Genesis Error Log")
        db.Close
        Exit Sub
    End If
    If IsNull(rs3("treasurer")) Or IsNull(rs3("treasureraddress1")) Or IsNull(rs3("treasureraddress2")) Then
        msg = MsgBox("No treasurer information found on System tab.", 48, "Genesis Error Log")
        db.Close
        Exit Sub
    End If
    rs3.MoveFirst
    inp = InputBox("Enter check number.", "Genesis Information Log", "")
    If inp = "" Or Val(inp) = 0 Then
        msg = MsgBox("Invalid check number entered.", 48, "Genesis Error Log")
        db.Close
        Exit Sub
    End If
    Printer.FontName = "Courier New"
    Printer.FontSize = 10
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print Tab(10); "SERVICE FEES: "; Tab(30); fromdate; " through "; todate
    Printer.Print Tab(10); "CHECK DATE:  "; Tab(30); Format$(Date$, "mm/dd/yyyy")
    Printer.Print Tab(10); "AMOUNT PAID: "; Tab(30); Format$(totamt, "$#########0.00")
    For xt% = 1 To 26
        Printer.Print
    Next xt%
    Printer.Print Tab(63); Format$(Date$, "mm/dd/yyyy"); Tab(87); Format$(totamt, "$#########0.00")
    Printer.Print
    ta$ = Format$(totamt, "0000000000.00")
    GoSub makeword
    Printer.Print Tab(9); buildword
    Printer.Print
    Printer.Print Tab(10); rs3("treasurer")
    Printer.Print Tab(10); rs3("treasureraddress1")
    Printer.Print Tab(10); rs3("treasureraddress2")
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print Tab(10); "SERVICE FEES: "; Tab(30); fromdate; " through "; todate
    Printer.Print Tab(10); "CHECK DATE:  "; Tab(30); Format$(Date$, "mm/dd/yyyy")
    Printer.Print Tab(10); "AMOUNT PAID: "; Tab(30); Format$(totamt, "$#########0.00")
    Printer.EndDoc
    Set rs = db.OpenRecordset("select * from checks where checknumber = '" + inp + "' and checkdate = #" + Date$ + "#")
    If rs.EOF Then
        rs.AddNew
    Else
        rs.MoveFirst
        rs.Edit
    End If
    rs("payto") = rs3("treasurer")
    rs("checknumber") = Val(inp)
    rs("type") = "SERVICE FEES"
    rs("fromdate") = fd
    rs("todate") = td
    rs("amount") = totamt
    rs("checkdate") = Format$(Date$, "mm/dd/yyyy")
    rs.Update
End If
If ecc Then
    Set rs = db.OpenRecordset("Select sum(commiss) as sf from executionspay where NOT PAYREM LIKE 'WRITEOFF*' AND NOT PAYREM LIKE 'writeoff*' and datepaid between #" + fromdate + "# and #" + todate + "#")
    If Not rs.EOF Then
        rs.MoveFirst
        If Not IsNull(rs("sf")) Then
            totamt = rs("sf")
        Else
            totamt = 0
        End If
    Else
        msg = MsgBox("No execution checks found for dates specified.", 48, "Genesis Error Log")
        db.Close
        Exit Sub
    End If
    If totamt = 0 Then
        msg = MsgBox("No commissions found for dates specified.", 48, "Genesis Error Log")
        db.Close
        Exit Sub
    End If
    Set rs3 = db.OpenRecordset("select treasurer, treasureraddress1, treasureraddress2 from system")
    If rs3.EOF Then
        msg = MsgBox("No treasurer information found on System tab.", 48, "Genesis Error Log")
        db.Close
        Exit Sub
    End If
    If IsNull(rs3("treasurer")) Or IsNull(rs3("treasureraddress1")) Or IsNull(rs3("treasureraddress2")) Then
        msg = MsgBox("No treasurer information found on System tab.", 48, "Genesis Error Log")
        db.Close
        Exit Sub
    End If
    rs3.MoveFirst
    inp = InputBox("Enter check number.", "Genesis Information Log", "")
    If inp = "" Or Val(inp) = 0 Then
        msg = MsgBox("Invalid check number entered.", 48, "Genesis Error Log")
        db.Close
        Exit Sub
    End If
    Printer.FontName = "Courier New"
    Printer.FontSize = 10
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print Tab(10); "COMMISSION: "; Tab(30); fromdate; " through "; todate
    Printer.Print Tab(10); "CHECK DATE:  "; Tab(30); Format$(Date$, "mm/dd/yyyy")
    Printer.Print Tab(10); "AMOUNT PAID: "; Tab(30); Format$(totamt, "$#########0.00")
    For xt% = 1 To 26
        Printer.Print
    Next xt%
    Printer.Print Tab(63); Format$(Date$, "mm/dd/yyyy"); Tab(87); Format$(totamt, "$#########0.00")
    ta$ = Format$(totamt, "0000000000.00")
    GoSub makeword
    Printer.Print
    Printer.Print Tab(9); buildword
    Printer.Print
    Printer.Print Tab(10); rs3("treasurer")
    Printer.Print Tab(10); rs3("treasureraddress1")
    Printer.Print Tab(10); rs3("treasureraddress2")
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print Tab(10); "COMMISSION: "; Tab(30); fromdate; " through "; todate
    Printer.Print Tab(10); "CHECK DATE:  "; Tab(30); Format$(Date$, "mm/dd/yyyy")
    Printer.Print Tab(10); "AMOUNT PAID: "; Tab(30); Format$(totamt, "$#########0.00")
    Printer.EndDoc
    Set rs = db.OpenRecordset("select * from checks where checknumber = '" + inp + "' and checkdate = #" + Date$ + "#")
    If rs.EOF Then
        rs.AddNew
    Else
        rs.MoveFirst
        rs.Edit
    End If
    rs("payto") = rs3("treasurer")
    rs("checknumber") = Val(inp)
    rs("type") = "COMMISSION  "
    rs("fromdate") = fd
    rs("todate") = td
    rs("amount") = totamt
    rs("checkdate") = Format$(Date$, "mm/dd/yyyy")
    rs.Update
End If
If epic Then
    Screen.MousePointer = 11
    so$ = ""
    DR$ = ""
    IT$ = ""
    Set rs = db.OpenRecordset("Select SERVICEFEE,serviceof, datereceived, iteration, principal, inter from executionspay where NOT PAYREM LIKE 'WRITEOFF*' AND NOT PAYREM LIKE 'writeoff*' and datepaid between #" + fromdate + "# and #" + todate + "# order by serviceof, datereceived, iteration")
    If Not rs.EOF Then
        rs.MoveFirst
        so$ = rs("serviceof")
        DR$ = rs("datereceived")
        IT$ = rs("iteration")
        totamt = 0
        inp = InputBox("Enter starting check number.", "Genesis Information Log", "")
        If inp = "" Or Val(inp) = 0 Then
            msg = MsgBox("Invalid check number entered.", 48, "Genesis Error Log")
            db.Close
            Exit Sub
        End If
        chk% = Val(inp)
    Else
        msg = MsgBox("No execution checks found for dates specified.", 48, "Genesis Error Log")
        Screen.MousePointer = 0
        db.Close
        Exit Sub
    End If
    While Not rs.EOF
      If rs("serviceof") <> so$ Or rs("datereceived") <> CVDate(rs("datereceived")) Or IT$ <> rs("iteration") Then
        Set rs2 = db.OpenRecordset("select * from executions where serviceof = " + Chr$(34) + so$ + Chr$(34) + " and datereceived = #" + DR$ + "# and iteration = " + Chr$(34) + IT$ + Chr$(34))
        If rs2.EOF Then
            msg = MsgBox("No valid execution information found for plaintiff for SERVICE OF: " + rs("serviceof") + "  DATE RECEIVED: " + Str$(rs("datereceived")) + "  ITERATION: " + rs("iteration"), 48, "Genesis Error Log")
            GoTo loopepic
        End If
        rs2.MoveFirst
        rs2.Edit
        If IsNull(rs2("phomeaddress")) And IsNull(rs2("pworkaddress")) Then
            rs2("phomeaddress") = "   "
            rs2("phomeaddress2") = "   "
            rs2("phomestate") = " "
            rs2("phomezipcode") = " "
        End If
        If IsNull(rs2("phomeaddress2")) Then
            rs2("phomeaddress2") = "   "
            rs2("phomestate") = " "
            rs2("phomezipcode") = " "
        End If
        If IsNull(rs2("pworkaddress2")) Then
            rs2("pworkaddress2") = "   "
            rs2("pworkstate") = " "
            rs2("pworkzipcode") = " "
        End If
        Printer.FontName = "Courier New"
        Printer.FontSize = 10
        Printer.Print
        Printer.Print
        Printer.Print
        Printer.Print Tab(10); "CASE NUMBER: "; Tab(30); rs2("casenumber")
        Printer.Print Tab(10); "DEFENDANT:   "; Tab(30); rs2("defendant")
        Printer.Print Tab(10); "PLAINTIFF:   "; Tab(30); rs2("plaintiff")
        Printer.Print Tab(10); "CHECK DATE:  "; Tab(30); Format$(Date$, "mm/dd/yyyy")
        Printer.Print Tab(10); "AMOUNT PAID: "; Tab(30); Format$(totamt, "$#########0.00")
        For xt% = 1 To 24
            Printer.Print
        Next xt%
        Printer.Print Tab(63); Format$(Date$, "mm/dd/yyyy"); Tab(87); Format$(totamt, "$#########0.00")
        ta$ = Format$(totamt, "0000000000.00")
        GoSub makeword
        Printer.Print
        Printer.Print Tab(9); buildword
        Printer.Print
        If rs2("PROFESSIONAL") = "" Or Left$(rs2("PROFESSIONAL"), 7) = "** WARN" Then
            Printer.Print Tab(10); rs2("plaintiff")
            If IsNull(rs2("phomeaddress")) Or rs2("phomeaddress") = "" Then
                Printer.Print Tab(10); rs2("pworkaddress")
                Printer.Print Tab(10); rs2("pworkaddress2") + " " + rs2("pworkstate") + " " + rs2("pworkzipcode")
            Else
                Printer.Print Tab(10); rs2("phomeaddress")
                Printer.Print Tab(10); rs2("phomeaddress2") + " " + rs2("phomestate") + " " + rs2("phomezipcode")
            End If
        Else
            Set db2 = OpenDatabase(nwl + "lawsuite.mdb")
            Set rs3 = db2.OpenRecordset("select * from professionals where type = 'A' and profname = " + Chr$(34) + rs2("professional") + Chr$(34))
            
            Printer.Print Tab(10); rs2("pROFESSIONAL")
            If Not rs3.EOF Then
                rs3.MoveFirst
                If IsNull(rs3("profaddr1")) Then
                    Printer.Print
                Else
                    Printer.Print Tab(10); rs3("profaddr1")
                End If
                If IsNull(rs3("profaddr2")) Then
                    Printer.Print
                Else
                    Printer.Print Tab(10); rs3("profaddr2")
                End If
            Else
                Printer.Print
                Printer.Print
            End If
        End If
        Printer.Print
        Printer.Print
        Printer.Print
        Printer.Print
        Printer.Print
        Printer.Print
        Printer.Print
        Printer.Print
        Printer.Print
        Printer.Print
        Printer.Print
        Printer.Print Tab(10); "CASE NUMBER: "; Tab(30); rs2("casenumber")
        Printer.Print Tab(10); "DEFENDANT:   "; Tab(30); rs2("defendant")
        Printer.Print Tab(10); "PLAINTIFF:   "; Tab(30); rs2("plaintiff")
        Printer.Print Tab(10); "CHECK DATE:  "; Tab(30); Format$(Date$, "mm/dd/yyyy")
        Printer.Print Tab(10); "AMOUNT PAID: "; Tab(30); Format$(totamt, "$#########0.00")
        Printer.EndDoc

        Set rs3 = db.OpenRecordset("select * from checks where checknumber = '" + Mid$(Str$(chk%), 2) + "' and checkdate = #" + Date$ + "#")
        If rs3.EOF Then
            rs3.AddNew
        Else
            rs3.MoveFirst
            rs3.Edit
        End If
        If rs2("PROFESSIONAL") = "" Then
            rs3("PAYTO") = rs2("plaintiff")
        Else
            rs3("PAYTO") = rs2("pROFESSIONAL")
        End If

        rs3("payto") = rs2("plaintiff")
        rs3("checknumber") = chk%
        rs3("type") = rs2("casenumber")
        rs3("fromdate") = fd
        rs3("todate") = td
        rs3("amount") = totamt
        rs3("checkdate") = Format$(Date$, "mm/dd/yyyy")
        rs3.Update
        chk% = chk% + 1
        so$ = rs("serviceof")
        DR$ = rs("datereceived")
        IT$ = rs("iteration")
        totamt = 0
      End If
      totamt = totamt + rs("principal") + rs("inter")
      If Not IsNull(rs("SERVICEFEE")) Then
        totamt = totamt + rs("SERVICEFEE")
      End If
loopepic:
        rs.MoveNext
    Wend
End If
On Error GoTo 0
If Not epic Then
    GoTo getoutpc
End If
Set rs2 = db.OpenRecordset("select * from executions where serviceof = " + Chr$(34) + so$ + Chr$(34) + " and datereceived = #" + DR$ + "# and iteration = " + Chr$(34) + IT$ + Chr$(34))
If rs2.EOF Then
    msg = MsgBox("No valid execution information found for plaintiff for SERVICE OF: " + rs("serviceof") + "  DATE RECEIVED: " + Str$(rs("datereceived")) + "  ITERATION: " + rs("iteration"), 48, "Genesis Error Log")
    GoTo loopepic
End If
rs2.MoveFirst
rs2.Edit
If IsNull(rs2("phomeaddress")) And IsNull(rs2("pworkaddress")) Then
    rs2("phomeaddress") = "   "
    rs2("phomeaddress2") = "   "
    rs2("phomestate") = " "
    rs2("phomezipcode") = " "
End If
If IsNull(rs2("phomeaddress2")) Then
    rs2("phomeaddress2") = "   "
    rs2("phomestate") = " "
    rs2("phomezipcode") = " "
End If
If IsNull(rs2("pworkaddress2")) Then
    rs2("pworkaddress2") = "   "
    rs2("pworkstate") = " "
    rs2("pworkzipcode") = " "
End If
Printer.FontName = "Courier New"
Printer.FontSize = 10
Printer.Print
Printer.Print
Printer.Print
Printer.Print Tab(10); "CASE NUMBER: "; Tab(30); rs2("casenumber")
Printer.Print Tab(10); "DEFENDANT:   "; Tab(30); rs2("defendant")
Printer.Print Tab(10); "PLAINTIFF:   "; Tab(30); rs2("plaintiff")
Printer.Print Tab(10); "CHECK DATE:  "; Tab(30); Format$(Date$, "mm/dd/yyyy")
Printer.Print Tab(10); "AMOUNT PAID: "; Tab(30); Format$(totamt, "$#########0.00")
For xt% = 1 To 24
    Printer.Print
Next xt%
Printer.Print Tab(63); Format$(Date$, "mm/dd/yyyy"); Tab(87); Format$(totamt, "$#########0.00")
ta$ = Format$(totamt, "0000000000.00")
GoSub makeword
Printer.Print
Printer.Print Tab(9); buildword
Printer.Print
If rs2("PROFESSIONAL") = "" Or Left$(rs2("PROFESSIONAL"), 7) = "** WARN" Then
    Printer.Print Tab(10); rs2("plaintiff")
    If IsNull(rs2("phomeaddress")) Or rs2("phomeaddress") = "" Then
        Printer.Print Tab(10); rs2("pworkaddress")
        Printer.Print Tab(10); rs2("pworkaddress2") + " " + rs2("pworkstate") + " " + rs2("pworkzipcode")
    Else
        Printer.Print Tab(10); rs2("phomeaddress")
        Printer.Print Tab(10); rs2("phomeaddress2") + " " + rs2("phomestate") + " " + rs2("phomezipcode")
    End If
Else
    Set db2 = OpenDatabase(nwl + "lawsuite.mdb")
    Set rs3 = db2.OpenRecordset("select * from professionals where type = 'A' and profname = " + Chr$(34) + rs2("professional") + Chr$(34))
    Printer.Print Tab(10); rs2("pROFESSIONAL")
    If Not rs3.EOF Then
        rs3.MoveFirst
        If IsNull(rs3("profaddr1")) Then
            Printer.Print
        Else
            Printer.Print Tab(10); rs3("profaddr1")
        End If
        If IsNull(rs3("profaddr2")) Then
            Printer.Print
        Else
            Printer.Print Tab(10); rs3("profaddr2")
        End If
    Else
        Printer.Print
        Printer.Print
    End If
End If
Printer.Print
Printer.Print
Printer.Print
Printer.Print
Printer.Print
Printer.Print
Printer.Print
Printer.Print
Printer.Print
Printer.Print
Printer.Print
Printer.Print Tab(10); "CASE NUMBER: "; Tab(30); rs2("casenumber")
Printer.Print Tab(10); "DEFENDANT:   "; Tab(30); rs2("defendant")
Printer.Print Tab(10); "PLAINTIFF:   "; Tab(30); rs2("plaintiff")
Printer.Print Tab(10); "CHECK DATE:  "; Tab(30); Format$(Date$, "mm/dd/yyyy")
Printer.Print Tab(10); "AMOUNT PAID: "; Tab(30); Format$(totamt, "$#########0.00")
Printer.EndDoc

Set rs3 = db.OpenRecordset("select * from checks where checknumber = '" + Mid$(Str$(chk%), 2) + "' and checkdate = #" + Date$ + "#")
If rs3.EOF Then
    rs3.AddNew
Else
    rs3.MoveFirst
    rs3.Edit
End If
If rs2("PROFESSIONAL") = "" Then
    rs3("PAYTO") = rs2("plaintiff")
Else
    rs3("PAYTO") = rs2("pROFESSIONAL")
End If
rs3("payto") = rs2("plaintiff")
rs3("checknumber") = chk%
rs3("type") = rs2("casenumber")
rs3("fromdate") = fd
rs3("todate") = td
rs3("amount") = totamt
rs3("checkdate") = Format$(Date$, "mm/dd/yyyy")
rs3.Update
getoutpc:
db.Close
Screen.MousePointer = 0
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If

makeword:
buildword = ""
b = Val(Left$(ta$, 1))
hm = Val(Mid$(ta$, 2, 1))
tm = Val(Mid$(ta$, 3, 1))
m = Val(Mid$(ta$, 4, 1))
hth = Val(Mid$(ta$, 5, 1))
tth = Val(Mid$(ta$, 6, 1))
th = Val(Mid$(ta$, 7, 1))
h = Val(Mid$(ta$, 8, 1))
t = Val(Mid$(ta$, 9, 1))
o = Val(Mid$(ta$, 10, 1))
cents = Val(Right$(ta$, 2))
If b > 0 Then
    Select Case b
        Case 1
            buildword = buildword + " One Billion"
        Case 2
            buildword = buildword + " Two Billion"
        Case 3
            buildword = buildword + " Three Billion"
        Case 4
            buildword = buildword + " Four Billion"
        Case 5
            buildword = buildword + " Five Billion"
        Case 6
            buildword = buildword + " Six Billion"
        Case 7
            buildword = buildword + " Seven Billion"
        Case 8
            buildword = buildword + " Eight Billion"
        Case 9
            buildword = buildword + " Nine Billion"
    End Select
End If
If hm > 0 Then
    Select Case hm
        Case 1
            buildword = buildword + " One Hundred Million"
        Case 2
            buildword = buildword + " Two Hundred Million"
        Case 3
            buildword = buildword + " Three Hundred Million"
        Case 4
            buildword = buildword + " Four Hundred Million"
        Case 5
            buildword = buildword + " Five Hundred Million"
        Case 6
            buildword = buildword + " Six Hundred Million"
        Case 7
            buildword = buildword + " Seven Hundred Million"
        Case 8
            buildword = buildword + " Eight Hundred Million"
        Case 9
            buildword = buildword + " Nine Hundred Million"
    End Select
End If
If tm > 0 Then
    Select Case tm
        Case 1
            buildword = buildword + " Ten Million"
        Case 2
            buildword = buildword + " Twenty Million"
        Case 3
            buildword = buildword + " Thirty Million"
        Case 4
            buildword = buildword + " Forty Million"
        Case 5
            buildword = buildword + " Fifty Million"
        Case 6
            buildword = buildword + " Sixty Million"
        Case 7
            buildword = buildword + " Seventy Million"
        Case 8
            buildword = buildword + " Eighty Million"
        Case 9
            buildword = buildword + " Ninety Million"
    End Select
End If
If m > 0 Then
    Select Case m
        Case 1
            buildword = buildword + " One Million"
        Case 2
            buildword = buildword + " Two Million"
        Case 3
            buildword = buildword + " Three Million"
        Case 4
            buildword = buildword + " Four Million"
        Case 5
            buildword = buildword + " Five Million"
        Case 6
            buildword = buildword + " Six Million"
        Case 7
            buildword = buildword + " Seven Million"
        Case 8
            buildword = buildword + " Eight Million"
        Case 9
            buildword = buildword + " Nine Million"
    End Select
End If
If hth > 0 Then
    Select Case hth
        Case 1
            buildword = buildword + " One Hundred Thousand"
        Case 2
            buildword = buildword + " Two Hundred Thousand"
        Case 3
            buildword = buildword + " Three Hundred Thousand"
        Case 4
            buildword = buildword + " Four Hundred Thousand"
        Case 5
            buildword = buildword + " Five Hundred Thousand"
        Case 6
            buildword = buildword + " Six Hundred Thousand"
        Case 7
            buildword = buildword + " Seven Hundred Thousand"
        Case 8
            buildword = buildword + " Eight Hundred Thousand"
        Case 9
            buildword = buildword + " Nine Hundred Thousand"
    End Select
End If
If tth > 0 Then
    Select Case tth
        Case 1
            buildword = buildword + " Ten Thousand"
        Case 2
            buildword = buildword + " Twenty Thousand"
        Case 3
            buildword = buildword + " Thirty Thousand"
        Case 4
            buildword = buildword + " Forty Thousand"
        Case 5
            buildword = buildword + " Fifty Thousand"
        Case 6
            buildword = buildword + " Sixty Thousand"
        Case 7
            buildword = buildword + " Seventy Thousand"
        Case 8
            buildword = buildword + " Eighty Thousand"
        Case 9
            buildword = buildword + " Ninety Thousand"
    End Select
End If
If th > 0 Then
    Select Case th
        Case 1
            buildword = buildword + " One Thousand"
        Case 2
            buildword = buildword + " Two Thousand"
        Case 3
            buildword = buildword + " Three Thousand"
        Case 4
            buildword = buildword + " Four Thousand"
        Case 5
            buildword = buildword + " Five Thousand"
        Case 6
            buildword = buildword + " Six Thousand"
        Case 7
            buildword = buildword + " Seven Thousand"
        Case 8
            buildword = buildword + " Eight Thousand"
        Case 9
            buildword = buildword + " Nine Thousand"
    End Select
End If
If h > 0 Then
    Select Case h
        Case 1
            buildword = buildword + " One Hundred"
        Case 2
            buildword = buildword + " Two Hundred"
        Case 3
            buildword = buildword + " Three Hundred"
        Case 4
            buildword = buildword + " Four Hundred"
        Case 5
            buildword = buildword + " Five Hundred"
        Case 6
            buildword = buildword + " Six Hundred"
        Case 7
            buildword = buildword + " Seven Hundred"
        Case 8
            buildword = buildword + " Eight Hundred"
        Case 9
            buildword = buildword + " Nine Hundred"
    End Select
End If
If t > 0 Then
    Select Case t
        Case 1
            buildword = buildword + " Ten"
        Case 2
            buildword = buildword + " Twenty"
        Case 3
            buildword = buildword + " Thirty"
        Case 4
            buildword = buildword + " Forty"
        Case 5
            buildword = buildword + " Fifty"
        Case 6
            buildword = buildword + " Sixty"
        Case 7
            buildword = buildword + " Seventy"
        Case 8
            buildword = buildword + " Eighty"
        Case 9
            buildword = buildword + " Ninety"
    End Select
End If
Select Case o
    Case 0
        If buildword = "" Then
            buildword = "Zero Dollars"
        End If
    Case 1
        buildword = buildword + " One Dollar"
    Case 2
        buildword = buildword + " Two Dollars"
    Case 3
        buildword = buildword + " Three Dollars"
    Case 4
        buildword = buildword + " Four Dollars"
    Case 5
        buildword = buildword + " Five Dollars"
    Case 6
        buildword = buildword + " Six Dollars"
    Case 7
        buildword = buildword + " Seven Dollars"
    Case 8
        buildword = buildword + " Eight Dollars"
    Case 9
        buildword = buildword + " Nine Dollars"
End Select
buildword = buildword + " & " + Mid$(Str$(cents), 2) + "/100"
Return
End Sub

Private Sub Command8_Click()
Dim db As Database, rs As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nwc + dbname)
Select Case maintab.Tab
    Case 0
        Set rs = db.OpenRecordset("select casenumber from magistrate order by casenumber desc")
    Case 1
        Set rs = db.OpenRecordset("select casenumber from writother order by casenumber desc")
    Case 2
        Set rs = db.OpenRecordset("select casenumber from familycourt order by casenumber desc")
    Case 3
        Set rs = db.OpenRecordset("select casenumber from executions order by casenumber desc")
    Case Else
        db.Close
        Exit Sub
End Select
If Not rs.EOF Then
    rs.MoveFirst
Else
    msg = MsgBox("Highest Case Number is 0.", 48, "Genesis Infomation Log")
    db.Close
    Exit Sub
End If
msg = MsgBox("Highest Case Number is " + rs("casenumber") + ".", 48, "Genesis Error Log")
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

Private Sub Command9_Click()
inp = InputBox("Enter last date to archive through.", "Archive Date Entry", "")
ARCHIVEDATE = inp
If Not IsDate(ARCHIVEDATE) Then
    msg = MsgBox("Invalid archive date.", 48, "Genesis Error Log")
    Exit Sub
End If
msg = MsgBox("Are you sure you wish to archive all data up to " + ARCHIVEDATE + "?", 4, "Genesis Information Log")
If msg = 7 Then
    Exit Sub
End If
Screen.MousePointer = 11
'msg = MsgBox("Beginning archival of Executions.", 48, "Genesis Information Log")
TimedMsgBox "Beginning archival of Executions.", 0 'RLB code


Dim db1, db2 As Database, rs1, rs2, rs3, rs4 As Recordset
Set db1 = OpenDatabase(nwc + "civil.mdb")
Set db2 = OpenDatabase(nwc + "arccivil.mdb")
Set rs1 = db1.OpenRecordset("select * from executions where datereceived < #" + Format$(ARCHIVEDATE, "mm/dd/yyyy") + "# and balance <= 0")
If Not rs1.EOF Then
    rs1.MoveFirst
End If
While Not rs1.EOF
    DoEvents 'rlb
    Set rs2 = db2.OpenRecordset("select * from executions where serviceof = " + Chr$(34) + rs1("serviceof") + Chr$(34) + " and datereceived = #" + Format$(rs1("datereceived"), "mm/dd/yyyy") + "# and iteration = '" + rs1("iteration") + "' and balance <= 0")
    If rs2.EOF Then
        rs2.AddNew
    Else
        rs2.MoveFirst
        rs2.Edit
    End If
    Set rs3 = db1.OpenRecordset("select * from executionspay where serviceof = " + Chr$(34) + rs1("serviceof") + Chr$(34) + " and datereceived = #" + Format$(rs1("datereceived"), "mm/dd/yyyy") + "# and iteration = '" + rs1("iteration") + "'")
    If Not rs3.EOF Then
        rs3.MoveFirst
        While Not rs3.EOF
            rs3.Edit
            Set rs4 = db2.OpenRecordset("select * from executionspay where serviceof = " + Chr$(34) + rs3("serviceof") + Chr$(34) + " and datereceived = #" + Format$(rs3("datereceived"), "mm/dd/yyyy") + "# and iteration = '" + rs3("iteration") + "'")
            If rs4.EOF Then
                rs4.AddNew
            Else
                rs4.MoveFirst
                rs4.Edit
            End If
            For t% = 0 To rs3.Fields.Count - 1
                rs4(t%) = rs3(t%)
            Next t%
            rs4.Update
            rs3.Delete
            rs3.MoveNext
        Wend
    End If
    For t% = 0 To rs1.Fields.Count - 1
        rs2(t%) = rs1(t%)
    Next t%
    rs2.Update
    rs1.Delete
    rs1.MoveNext
Wend
'msg = MsgBox("Executions completed.  Beginning archival of Execution Payments.", 48, "Genesis Information Log")
TimedMsgBox "Executions Completed.  Beginning archival of Executiuon Payments.", 0 'RLB

'msg = MsgBox("Execution Payments completed.  Beginning archival of Family Court.", 48, "Genesis Information Log")
TimedMsgBox "Execution Payments completed.  Beginning archival of Family Court.", 0 'RLB
Set rs1 = db1.OpenRecordset("select * from familycourt where datereceived < #" + Format$(ARCHIVEDATE, "mm/dd/yyyy") + "#")
If Not rs1.EOF Then
    rs1.MoveFirst
End If
While Not rs1.EOF
    DoEvents 'rlb
    Set rs2 = db2.OpenRecordset("select * from familycourt where serviceof = " + Chr$(34) + rs1("serviceof") + Chr$(34) + " and datereceived = #" + Format$(rs1("datereceived"), "mm/dd/yyyy") + "# and iteration = '" + rs1("iteration") + "'")
    If rs2.EOF Then
        rs2.AddNew
    Else
        rs2.MoveFirst
        rs2.Edit
    End If
    For t% = 0 To rs1.Fields.Count - 1
        rs2(t%) = rs1(t%)
    Next t%
    rs2.Update
    rs1.Delete
    rs1.MoveNext
Wend
'msg = MsgBox("Family Court completed.  Beginning archival of Magistrate.", 48, "Genesis Information Log")
TimedMsgBox "Family Court completed.  Beginning archival of Magistrate.", 0 'RLB
Set rs1 = db1.OpenRecordset("select * from magistrate where datereceived < #" + Format$(ARCHIVEDATE, "mm/dd/yyyy") + "#")
If Not rs1.EOF Then
    rs1.MoveFirst
End If
While Not rs1.EOF
    DoEvents 'rlb
    Set rs2 = db2.OpenRecordset("select * from magistrate where serviceof = " + Chr$(34) + rs1("serviceof") + Chr$(34) + " and datereceived = #" + Format$(rs1("datereceived"), "mm/dd/yyyy") + "# and iteration = '" + rs1("iteration") + "'")
    If rs2.EOF Then
        rs2.AddNew
    Else
        rs2.MoveFirst
        rs2.Edit
    End If
    For t% = 0 To rs1.Fields.Count - 1
        rs2(t%) = rs1(t%)
    Next t%
    rs2.Update
    rs1.Delete
    rs1.MoveNext
Wend
'msg = MsgBox("Magistrate completed.  Beginning archival of Receipts.", 48, "Genesis Information Log")
TimedMsgBox "Magistrate completed.  Beginning archival of Receipts.", 0 'RLB
Set rs1 = db1.OpenRecordset("select * from receipt where datereceived < #" + Format$(ARCHIVEDATE, "mm/dd/yyyy") + "#")
If Not rs1.EOF Then
    rs1.MoveFirst
End If
While Not rs1.EOF
    DoEvents 'rlb
    Set rs2 = db2.OpenRecordset("select * from receipt where serviceof = " + Chr$(34) + rs1("serviceof") + Chr$(34) + " and datereceived = #" + Format$(rs1("datereceived"), "mm/dd/yyyy") + "# and iteration = '" + rs1("iteration") + "'")
    If rs2.EOF Then
        rs2.AddNew
    Else
        rs2.MoveFirst
        rs2.Edit
    End If
    For t% = 0 To rs1.Fields.Count - 1
        rs2(t%) = rs1(t%)
    Next t%
    rs2.Update
    rs1.Delete
    rs1.MoveNext
Wend
'msg = MsgBox("Receipts completed.  Beginning archival of Writ/Other.", 48, "Genesis Information Log")
TimedMsgBox "Receipts completed.  Beginning archival of Writ/Other.", 0 'RLB
Set rs1 = db1.OpenRecordset("select * from writother where datereceived < #" + Format$(ARCHIVEDATE, "mm/dd/yyyy") + "#")
If Not rs1.EOF Then
    rs1.MoveFirst
End If
While Not rs1.EOF
    DoEvents 'rlb
    Set rs2 = db2.OpenRecordset("select * from writother where serviceof = " + Chr$(34) + rs1("serviceof") + Chr$(34) + " and datereceived = #" + Format$(rs1("datereceived"), "mm/dd/yyyy") + "# and iteration = '" + rs1("iteration") + "'")
    If rs2.EOF Then
        rs2.AddNew
    Else
        rs2.MoveFirst
        rs2.Edit
    End If
    For t% = 0 To rs1.Fields.Count - 1
        rs2(t%) = rs1(t%)
    Next t%
    rs2.Update
    rs1.Delete
    rs1.MoveNext
Wend
Unload Dialog
msg = MsgBox("Writ/Other completed.  Archival process finished successfully.", 48, "Genesis Information Log")
db1.Close
db2.Close
Screen.MousePointer = 0
Exit Sub
End Sub

Private Sub COMMISS_GotFocus()
Call commissandint
End Sub


Private Sub commiss_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    inter.SetFocus
End If

End Sub

Private Sub commission_Change()
If sfpay Then
    total = Str$(Val(commission) + Val(balance) + Val(INTEREST) + Val(servicefee))
Else
    total = Str$(Val(commission) + Val(balance) + Val(INTEREST))
End If
End Sub

Private Sub commission_GotFocus()
Call commissandint
End Sub




Private Sub datereceived_DropDown()
a = 1
End Sub

Private Sub county_KeyPress(KeyAscii As Integer)
If KeyAscii = 34 Then
    msg = MsgBox("The " + Chr$(34) + " character is not allowed.  Use the ' character instead.", 48, "Genesis Error Log")
    KeyAscii = 0
    Exit Sub
End If

End Sub

Private Sub courtdate_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
If Len(courtdate) = 1 Or Len(courtdate) = 4 Then
    Call sendslash
End If
End If
If KeyAscii = 13 Then
    courttime.SetFocus
End If

End Sub


Private Sub courttime_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
If Len(courttime) = 1 Then
    Call sendcolon
End If
End If
If KeyAscii = 13 Then
    daystorespond.SetFocus
End If

End Sub

Private Sub custodian_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    professional.SetFocus
End If
End Sub

Private Sub DATEPAID_GotFocus()
DATEPAID = Format$(Date$, "mm/dd/yyyy")
End Sub

Private Sub DATEPAID_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
If Len(DATEPAID) = 1 Or Len(DATEPAID) = 4 Then
    Call sendslash
End If
End If
If KeyAscii = 13 Then
    amount.SetFocus
End If


End Sub

Private Sub datereceived_GotFocus()
If Len(serviceof) > 60 Then
    msg = MsgBox("Maximum length of 60 has been exceeded for SERVICE OF entry.  This entry will be truncated.", 48, "Genesis Error Log")
    serviceof = Left$(serviceof, 60)
End If
If serviceof > "" Then
sohomeaddress.Refresh
sohomeaddress2.Refresh
sohomestate.Refresh
sohomezipcode.Refresh
soworkaddress.Refresh
soworkaddress2.Refresh
soworkstate.Refresh
soworkzipcode.Refresh
casenumber.Refresh
armedforces.Refresh
corporate.Refresh
title.Refresh
papertype.Refresh
daystorespond.Refresh
servicefee.Refresh
bill.Refresh
feedate.Refresh
ivd.Refresh
custodian.Refresh
receipt.Refresh
defendant.Refresh
Label12.Refresh
'Label16.Refresh
'Label3.Refresh
'Label4.Refresh
feel.Refresh
Label6.Refresh
infoframe.Refresh
If maintab.Tab < 4 Then
    If daterlist.ListCount = 0 Then
        datereceived.SetFocus
    End If
    If datereceived = "" And iteration = "" Then
        Call serviceof_Click(0)
    End If
End If
End If

End Sub

Private Sub datereceived_keypress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 9 Then
    If datereceived <> "" Then
        Call daterlist_click
    End If
End If
If KeyAscii <> 8 Then
If Len(datereceived) = 1 Or Len(datereceived) = 4 Then
    Call sendslash
End If
End If
If KeyAscii = 13 Then
    iteration.SetFocus
End If
End Sub

Private Sub daterlist_click()
On Error Resume Next
If daterlist.ListIndex = -1 Then
    Exit Sub
End If
datereceived.Text = Format$(daterlist.List(daterlist.ListIndex), "mm/dd/yyyy")
If serviceof = "" Then
    Exit Sub
End If
If daterlist.ListCount = 1 Then
    iteration = "1"
End If
If iteration = "" Then
   Exit Sub
End If
If datereceived <> "" Then
    If Not IsDate(datereceived) Then
        msg = MsgBox("Filter entry in DATE RECEIVED is not a valid date.", 48, "Genesis Error Log")
        datereceived.SetFocus
        Exit Sub
    End If
End If
If Val(iteration) = 0 Then
   msg = MsgBox("Filter entry in ITERATION is not a valid number.", 48, "Genesis Error Log")
   iteration.SetFocus
   Exit Sub
End If
If maintab.Tab > 3 Then
    GoSub tab1
End If
On maintab.Tab + 1 GoSub tab1, tab2, tab3, tab4
Dim db As Database, ds As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nwc + dbname)
Set ds = db.OpenRecordset("select * from " + TP + " where serviceof = " + Chr$(34) + serviceof + Chr$(34) + " and datereceived = #" + datereceived + "#")
If ds.EOF Then
    db.Close
    Exit Sub
End If
ds.MoveFirst
If maintab.Tab = 2 Then
    If Not IsNull(ds("osce")) Then
        servicefee = ds("osce")
    Else
        servicefee = ""
    End If
    If Not IsNull(ds("IVD")) Then
        ivd.Value = ds("ivd")
    Else
        ivd.Value = 0
    End If
    If Not IsNull(ds("fl1")) Then
        custodian = ds("fl1")
    Else
        custodian = ""
    End If
End If
If maintab.Tab <> 2 Then
    If Not IsNull(ds("receiptnum")) Then
        receiptd = ds("receiptnum")
    Else
        receiptd = ""
    End If
    If Not IsNull(ds("checknum")) Then
        checkd = ds("checknum")
    Else
        checkd = ""
    End If
End If
serviceof = ds("serviceof")
iteration = ds("iteration")
datereceived = Format$(ds("datereceived"), "mm/dd/yyyy")
serviceofsort = ds("serviceofsort")
If Not IsNull(ds("CASENUMBER")) Then
    casenumber = ds("casenumber")
Else
    casenumber = "UNKNOWN"
End If
If Not IsNull(ds("fs2")) Then
    armedforces.Value = Val(ds("fs2"))
Else
    armedforces.Value = 0
End If
If Not IsNull(ds("fs1")) Then
    corporate.Value = Val(ds("fs1"))
Else
    corporate.Value = 0
End If
If Not IsNull(ds("fl1")) Then
    title = ds("fl1")
Else
    title = ""
End If
If Not IsNull(ds("sohomeaddress")) Then
    sohomeaddress = ds("sohomeaddress")
Else
    sohomeaddress = ""
End If
If Not IsNull(ds("sohomeaddress2")) Then
    sohomeaddress2 = ds("sohomeaddress2")
Else
    sohomeaddress2 = ""
End If
If Not IsNull(ds("sohomestate")) Then
    sohomestate = ds("sohomestate")
Else
    sohomestate = ""
End If
If Not IsNull(ds("sohomezipcode")) Then
    sohomezipcode = ds("sohomezipcode")
Else
    sohomezipcode = ""
End If
If Not IsNull(ds("sohomephone")) And ds("sohomephone") <> "" Then
    sohomephone = ds("sohomephone")
Else
    sohomephone = ""
End If
If Not IsNull(ds("soworkaddress")) Then
    soworkaddress = ds("soworkaddress")
Else
    soworkaddress = ""
End If
If Not IsNull(ds("soworkaddress2")) Then
    soworkaddress2 = ds("soworkaddress2")
Else
    soworkaddress2 = ""
End If
If Not IsNull(ds("soworkstate")) Then
    soworkstate = ds("soworkstate")
Else
    soworkstate = ""
End If
If Not IsNull(ds("soworkzipcode")) Then
    soworkzipcode = ds("soworkzipcode")
Else
    soworkzipcode = ""
End If
If Not IsNull(ds("soworkphone")) And ds("soworkphone") <> "" Then
    soworkphone = ds("soworkphone")
Else
    soworkphone = ""
End If
papertype = ds("papertype")
If Not IsNull(ds("courtdate")) Then
    courtdate = Format$(ds("courtdate"), "mm/dd/yyyy")
Else
    courtdate = ""
End If
If Not IsNull(ds("courttime")) Then
    courttime = ds("courttime")
Else
    courttime = ""
End If
If Not IsNull(ds("daystorespond")) Then
    daystorespond = ds("daystorespond")
Else
    daystorespond = ""
End If
If maintab.Tab <> 2 Then
    If Not IsNull(ds("servicefee")) Then
        servicefee = ds("servicefee")
    Else
        servicefee = ""
    End If
    If Not IsNull(ds("bill")) Then
        bill = ds("bill")
    Else
        bill = 0
    End If
    If Not IsNull(ds("feedate")) Then
        feedate = ds("feedate")
    Else
        feedate = ""
    End If
End If

defendant = ds("defendant")
defendantsort = ds("defendantsort")
If Not IsNull(ds("dhomeaddress")) Then
    dhomeaddress = ds("dhomeaddress")
Else
    dhomeaddress = ""
End If
If Not IsNull(ds("dhomeaddress2")) Then
    dhomeaddress2 = ds("dhomeaddress2")
Else
    dhomeaddress2 = ""
End If
If Not IsNull(ds("dhomestate")) Then
    dhomestate = ds("dhomestate")
Else
    dhomestate = ""
End If
If Not IsNull(ds("dhomezipcode")) Then
    dhomezipcode = ds("dhomezipcode")
Else
    dhomezipcode = ""
End If
If Not IsNull(ds("dhomephone")) And ds("dhomephone") <> "" Then
    dhomephone = ds("dhomephone")
Else
    dhomephone = ""
End If
If Not IsNull(ds("dworkaddress")) Then
    dworkaddress = ds("dworkaddress")
Else
    dworkaddress = ""
End If
If Not IsNull(ds("dworkaddress2")) Then
    dworkaddress2 = ds("dworkaddress2")
Else
    dworkaddress2 = ""
End If
If Not IsNull(ds("dworkstate")) Then
    dworkstate = ds("dworkstate")
Else
    dworkstate = ""
End If
If Not IsNull(ds("dworkzipcode")) Then
    dworkzipcode = ds("dworkzipcode")
Else
    dworkzipcode = ""
End If
If Not IsNull(ds("dworkphone")) And ds("dworkphone") <> "" Then
    dworkphone = ds("dworkphone")
Else
    dworkphone = ""
End If
plaintiff = ds("plaintiff")
plaintiffsort = ds("plaintiffsort")
If Not IsNull(ds("phomeaddress")) Then
    phomeaddress = ds("phomeaddress")
Else
    phomeaddress = ""
End If
If Not IsNull(ds("phomeaddress2")) Then
    phomeaddress2 = ds("phomeaddress2")
Else
    phomeaddress2 = ""
End If
If Not IsNull(ds("phomestate")) Then
    phomestate = ds("phomestate")
Else
    phomestate = ""
End If
If Not IsNull(ds("phomezipcode")) Then
    phomezipcode = ds("phomezipcode")
Else
    phomezipcode = ""
End If
If Not IsNull(ds("phomephone")) And ds("phomephone") <> "" Then
    phomephone = ds("phomephone")
Else
    phomephone = ""
End If
If Not IsNull(ds("pworkaddress")) Then
    pworkaddress = ds("pworkaddress")
Else
    pworkaddress = ""
End If
If Not IsNull(ds("pworkaddress2")) Then
    pworkaddress2 = ds("pworkaddress2")
Else
    pworkaddress2 = ""
End If
If Not IsNull(ds("pworkstate")) Then
    pworkstate = ds("pworkstate")
Else
    pworkstate = ""
End If
If Not IsNull(ds("pworkzipcode")) Then
    pworkzipcode = ds("pworkzipcode")
Else
    pworkzipcode = ""
End If
If Not IsNull(ds("pworkphone")) And ds("pworkphone") <> "" Then
    pworkphone = ds("pworkphone")
Else
    pworkphone = ""
End If
If Not IsNull(ds("assignedto")) Then
    assignedto = ds("assignedto")
Else
    assignedto = ""
End If
If Not IsNull(ds("assignedon")) Then
    assignedon = ds("assignedon")
Else
    assignedon = ""
End If
If Not IsNull(ds("served")) Then
    served.Value = Val(ds("served"))
Else
    served.Value = 0
End If
If Not IsNull(ds("nonservice")) Then
    nonservice.Value = Val(ds("nonservice"))
Else
    nonservice.Value = 0
End If
If Not IsNull(ds("nsreason")) Then
    nsreason = ds("nsreason")
Else
    nsreason = ""
End If
If Not IsNull(ds("premarks")) Then
    premarks = ds("premarks")
Else
    premarks = ""
End If
If maintab.Tab = 3 Then
    If Not IsNull(ds("levy")) Then
        levy.Text = ds("levy")
    Else
        levy.Text = ""
    End If
End If
If Not IsNull(ds("wremarks")) Then
    wremarks = ds("wremarks")
Else
    wremarks = ""
End If
If Not IsNull(ds("servicedate")) Then
    servicedate = Format$(ds("servicedate"), "mm/dd/yyyy")
Else
    servicedate = ""
End If
If Not IsNull(ds("servicetime")) Then
    servicetime = ds("servicetime")
Else
    servicetime = ""
End If
If Not IsNull(ds("personserved")) Then
    personserved = ds("personserved")
Else
    personserved = ""
End If
If Not IsNull(ds("locationserved")) Then
    locationserved = ds("locationserved")
Else
    locationserved = ""
End If
If Not IsNull(ds("relationship")) Then
    relationship = ds("relationship")
Else
    relationship = ""
End If
If Not IsNull(ds("professional")) Then
    professional = ds("professional")
Else
    professional = ""
End If
If maintab.Tab = 3 Then
        If Not IsNull(ds("apptdate")) Then
            apptdate = Format$(ds("apptdate"), "mm/dd/yyyy")
        Else
            apptdate = ""
        End If
        If Not IsNull(ds("intrate")) Then
                intrate = ds("intrate")
        Else
                intrate = ""
        End If
        If Not IsNull(ds("datesatisfied")) Then
                datesatisfied = Format$(ds("datesatisfied"), "mm/dd/yyyy")
        Else
                datesatisfied = ""
        End If
        If Not IsNull(ds("judgementdate")) Then
                judgementdate = Format$(ds("judgementdate"), "mm/dd/yyyy")
        Else
                judgementdate = ""
        End If
        If Not IsNull(ds("judgementamount")) Then
                judgementamount = ds("judgementamount")
        Else
                judgementamount = ""
        End If
        'If Not IsNull(ds("estpayoffdate")) Then
        '        estpayoffdate = Format$(ds("estpayoffdate"), "mm/dd/yyyy")
        'Else
        '        estpayoffdate = ""
        'End If
        estpayoffdate = Format$(Date$, "mm/dd/yyyy")
        If Not IsNull(ds("nulla")) Then
            nulla.Value = ds("nulla")
        Else
            nulla.Value = 0
        End If
        commission = ds("COMMISSION")
        If Not IsNull(ds("INTEREST")) Then
            INTEREST = ds("INTERest")
        Else
            INTEREST = 0
        End If
        perday = ds("PERDAY")
        If Not IsNull(ds("totalinterest")) Then
            totalinterest = ds("totalinterest")
        Else
            totalinterest = 0
        End If
        If Not IsNull(ds("totalcommission")) Then
            totalcommission = ds("totalcommission")
        Else
            totalcommission = 0
        End If
        If Not IsNull(ds("totalpayments")) Then
            totalpayments = ds("totalpayments")
        Else
            totalpayments = 0
        End If
        Call loadpay
End If
lastserviceof = serviceof
CSERVICEOF = 0
db.Close
infoframe.Refresh
On Error GoTo 0
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If

tab1:
TP = "magistrate"
Return
tab2:
TP = "writother"
Return
tab3:
TP = "familycourt"
Return
tab4:
TP = "executions"
Return

End Sub

Private Sub datesatisfied_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
If Len(datesatisfied) = 1 Or Len(datesatisfied) = 4 Then
    Call sendslash
End If
End If
If maintab.Tab = 3 Then
If KeyAscii = 13 Then
    judgementdate.SetFocus
End If
End If

End Sub


Private Sub daystorespond_KeyPress(KeyAscii As Integer)
If KeyAscii = 34 Then
    msg = MsgBox("The " + Chr$(34) + " character is not allowed.  Use the ' character instead.", 48, "Genesis Error Log")
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
    servicefee.SetFocus
End If

End Sub



Private Sub defendant_Click(AREA As Integer)
infoframe.Refresh
Dim db As Database, ds As Recordset
On Error GoTo oderror
od:
Call setpopup(defendant, "F")
Set db = OpenDatabase(nwl + "LAWSUITE.MDB")
Set ds = db.OpenRecordset("select * from PEOPLE where dpname = " + Chr$(34) + defendant + Chr$(34))
If Not ds.EOF Then
   ds.MoveFirst
    defendantsort = ds("dpsort")
    If Not IsNull(ds("dphaddress")) Then
        dhomeaddress = ds("dphaddress")
    Else
        dhomeaddress = ""
    End If
    If Not IsNull(ds("dphaddress2")) Then
        dhomeaddress2 = ds("dphaddress2")
    Else
        dhomeaddress2 = ""
    End If
    If Not IsNull(ds("hstate")) Then
        dhomestate = ds("hstate")
    Else
        dhomestate = ""
    End If
    If Not IsNull(ds("hzipcode")) Then
        dhomezipcode = ds("hzipcode")
    Else
        dhomezipcode = ""
    End If
    If Not IsNull(ds("dpwaddress")) Then
        dworkaddress = ds("dpwaddress")
    Else
        dworkaddress = ""
    End If
    If Not IsNull(ds("dpwaddress2")) Then
        dworkaddress2 = ds("dpwaddress2")
    Else
        dworkaddress2 = ""
    End If
    If Not IsNull(ds("wstate")) Then
        dworkstate = ds("wstate")
    Else
        dworkstate = ""
    End If
    If Not IsNull(ds("wzipcode")) Then
        dworkzipcode = ds("wzipcode")
    Else
        dworkzipcode = ""
    End If
    If Not IsNull(ds("dphphone")) And ds("dphphone") <> "" Then
        dhomephone = ds("dphphone")
    Else
        dhomephone = ""
    End If
    If Not IsNull(ds("dpwphone")) And ds("dpwphone") <> "" Then
        dworkphone = ds("dpwphone")
    Else
        dworkphone = ""
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

Private Sub defendant_GotFocus()
If Len(defendant) > 60 Then
    msg = MsgBox("Maximum length of 60 has been exceeded for DEFENDANT entry.  This entry will be truncated.", 48, "Genesis Error Log")
    defendant = Left$(defendant, 60)
End If
If defendant = "" Then
    defendant = serviceof
End If
End Sub



Private Sub defendant_KeyPress(KeyAscii As Integer)
If KeyAscii = 34 Then
    msg = MsgBox("The " + Chr$(34) + " character is not allowed.  Use the ' character instead.", 48, "Genesis Error Log")
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
    defendantsort.SetFocus
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
If Len(defendant) > 60 Then
    msg = MsgBox("SERVICE OF can be no more than 60 characters.  The data will be truncated.", 48, "Genesis Error Log")
    defendant = Left$(defendant, 60)
End If

End Sub

Private Sub defendantsort_GotFocus()
If defendant = serviceof Then
    defendantsort = serviceofsort
    dhomeaddress = sohomeaddress
    dhomeaddress2 = sohomeaddress2
    dhomestate = sohomestate
    dhomezipcode = sohomezipcode
    dhomephone = sohomephone
    dworkaddress = soworkaddress
    dworkaddress2 = soworkaddress2
    dworkstate = soworkstate
    dworkzipcode = soworkzipcode
    dworkphone = soworkphone
    Exit Sub
End If
If defendantsort > "" Then
    Exit Sub
End If
Dim db As Database, rs As Recordset, ff, LF As Integer, HS As String
ff = 0
LF = 1
HS = ""
On Error GoTo oderror
od:
Set db = OpenDatabase(nwc + dbname)
Set rs = db.OpenRecordset("select fnf,lnf from system")
If Not rs.EOF Then
    rs.MoveFirst
    If rs("fNf") = True Then
        ff = 1
        LF = 0
    End If
End If
db.Close
Call setsort(ff, LF, defendant, HS)
defendantsort = HS
On Error GoTo 0
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If

End Sub


Private Sub defendantsort_KeyPress(KeyAscii As Integer)
If KeyAscii = 34 Then
    msg = MsgBox("The " + Chr$(34) + " character is not allowed.  Use the ' character instead.", 48, "Genesis Error Log")
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
    dhomeaddress.SetFocus
End If

End Sub

Private Sub deletebutton_Click()
If UCase(frmLogin.txtUserName) = "DEMO" And UCase(frmLogin.txtPassword) = "DEMO" Then
    msg = MsgBox("Not available in DEMO version.", 48, "Genesis Information Log")
    Screen.MousePointer = 0
    Exit Sub
End If
SEARCHTYPE = 0
If Val(frmLogin.CDELETE(0)) = 1 And Val(frmLogin.CDELETE(1)) = 1 And Val(frmLogin.CDELETE(2)) = 1 And Val(frmLogin.CDELETE(3)) = 1 Then
    a = 1
Else
    msg = MsgBox("Your USER ID does not have sufficient access to perform this task.", 48, "Genesis Information Log")
    Exit Sub
End If
If serviceof = "" Then
    msg = MsgBox("Invalid entry in SERVICE OF field.", 48, "Genesis Error Log")
    serviceof.SetFocus
    Exit Sub
End If
If Not IsDate(datereceived) Then
    msg = MsgBox("Invalid entry in DATE RECEIVED field.", 48, "Genesis Error Log")
    datereceived.SetFocus
    Exit Sub
End If
If iteration = "" Then
    msg = MsgBox("Invalid entry in ITERATION field.", 48, "Genesis Error Log")
    iteration.SetFocus
    Exit Sub
End If
If datereceived = "" Then
    Exit Sub
End If
msg = MsgBox("Are you sure you wish to delete this record?", 4, "Genesis Information Log")
If msg = 7 Then
    Exit Sub
End If
Screen.MousePointer = 11
If maintab.Tab > 3 Then
    GoSub tab1
End If
On maintab.Tab + 1 GoSub tab1, tab2, tab3, tab4
Dim db As Database, ds As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nwc + dbname)
Set ds = db.OpenRecordset("select * from " + TP + " where serviceof = " + Chr$(34) + serviceof + Chr$(34) + " and datereceived = #" + datereceived + "# AND ITERATION = " + Chr$(34) + iteration + Chr$(34))
If ds.EOF Then
    Screen.MousePointer = 0
    db.Close
    Exit Sub
Else
    ds.MoveFirst
    ds.Delete
End If
If TP = "executions" Then
    Set ds = db.OpenRecordset("select * from executionspay where serviceof = " + Chr$(34) + serviceof + Chr$(34) + " and datereceived = #" + datereceived + "# AND ITERATION = " + Chr$(34) + iteration + Chr$(34))
    If ds.EOF Then
        Screen.MousePointer = 0
        db.Close
        Exit Sub
    Else
        ds.MoveFirst
        ds.Delete
    End If
End If
lastserviceof = ""
Call clearit
db.Close
hold$ = papertype
Call loadpapertype
papertype = hold$
hold$ = professional
Call loadprof
professional = hold$
hold$ = assignedto
Call loaddeputy
assignedto = hold$
Screen.MousePointer = 0
On Error GoTo 0
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If

tab1:
TP = "magistrate"
Return
tab2:
TP = "writother"
Return
tab3:
TP = "familycourt"
Return
tab4:
TP = "executions"
Return
End Sub

Private Sub dhomeaddress_KeyPress(KeyAscii As Integer)
If KeyAscii = 34 Then
    msg = MsgBox("The " + Chr$(34) + " character is not allowed.  Use the ' character instead.", 48, "Genesis Error Log")
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
    dhomeaddress2.SetFocus
End If

End Sub

Private Sub dhomeaddress2_KeyPress(KeyAscii As Integer)
If KeyAscii = 34 Then
    msg = MsgBox("The " + Chr$(34) + " character is not allowed.  Use the ' character instead.", 48, "Genesis Error Log")
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
    dhomestate.SetFocus
End If

End Sub

Private Sub dhomephone_GotFocus()
On Error Resume Next
If Len(dhomephone) = 0 Then
    Dim db As Database, rs As Recordset
    On Error GoTo oderror
od:
    Set db = OpenDatabase(nwl + "LAWSUITE.MDB")
    Set rs = db.OpenRecordset("SELECT AREACODE FROM DEFAULTS")
    If rs.EOF Then
        db.Close
        Exit Sub
    End If
    rs.MoveFirst
    If IsNull(rs("AREACODE")) Then
        db.Close
        Exit Sub
    End If
    If Len(rs("AREACODE")) <> 3 Then
        db.Close
        Exit Sub
    End If
    Call sendopenpara
    Call SENDCHAR(Left$(rs("AREACODE"), 1))
    Call SENDCHAR(Mid$(rs("AREACODE"), 2, 1))
    Call SENDCHAR(Mid$(rs("AREACODE"), 3, 1))
    Call SENDEND
    db.Close
End If
On Error GoTo 0
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If


End Sub

Private Sub dhomephone_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
If Len(dhomephone) = 3 Then
    Call sendclosepara
End If
If Len(dhomephone) = 4 Then
    Call sendspace
End If
If Len(dhomephone) = 8 Then
    Call senddash
End If
If Len(dhomephone) = 13 Then
    Call sendspace
End If
End If
If KeyAscii = 13 Then
    dworkaddress.SetFocus
End If

End Sub


Private Sub dhomephone_LostFocus()
If Len(dhomephone) = 5 Or Len(dhomephone) = 6 Then
    dhomephone = ""
End If

End Sub

Private Sub dhomestate_KeyPress(KeyAscii As Integer)
If KeyAscii = 34 Then
    msg = MsgBox("The " + Chr$(34) + " character is not allowed.  Use the ' character instead.", 48, "Genesis Error Log")
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
    dhomezipcode.SetFocus
End If

End Sub

Private Sub dhomezipcode_KeyPress(KeyAscii As Integer)
If KeyAscii = 34 Then
    msg = MsgBox("The " + Chr$(34) + " character is not allowed.  Use the ' character instead.", 48, "Genesis Error Log")
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
    dhomephone.SetFocus
End If

End Sub

Private Sub dprof_Click()
If Val(frmLogin.CSUPERVISOR(0)) = 1 And Val(frmLogin.CSUPERVISOR(1)) = 1 And Val(frmLogin.CSUPERVISOR(2)) = 1 And Val(frmLogin.CSUPERVISOR(3)) = 1 Or Val(frmLogin.SUPERVISOR) = 1 Then
    a = 1
Else
    msg = MsgBox("Your USER ID does not have sufficient access to perform this task.", 48, "Genesis Information Log")
    Exit Sub
End If
Dim db As Database, ds As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nwl + "LAWSUITE.MDB")
If omag.Value = True Then
    TP = "M"
End If
If oatt.Value = True Then
    TP = "A"
End If
If ocou.Value = True Then
    TP = "C"
End If
If odep.Value = True Then
    TP = "D"
End If
If profname > "" Then
   Set ds = db.OpenRecordset("select * from professionals where profname = " + Chr$(34) + profname + Chr$(34) + " and type = " + Chr$(34) + TP + Chr$(34))
       If Not ds.EOF Then
           ds.MoveFirst
           ds.Delete
       End If
End If
On Error GoTo 0
db.Close
Call loadprof
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If

End Sub

Private Sub dragbutton_Click()

End Sub



Private Sub duser_Click()
If Val(frmLogin.CSUPERVISOR(0)) = 1 And Val(frmLogin.CSUPERVISOR(1)) = 1 And Val(frmLogin.CSUPERVISOR(2)) = 1 And Val(frmLogin.CSUPERVISOR(3)) = 1 Then
    a = 1
Else
    msg = MsgBox("Your USER ID does not have sufficient access to perform this task.", 48, "Genesis Information Log")
    Exit Sub
End If
If userid = "" Or password = "" Then
    msg = MsgBox("Both USER ID and PASSWORD must be entered.", 48, "Genesis Error Log")
    Exit Sub
End If
flipuser$ = ""
For t% = Len(userid) To 1 Step -1
    flipuser$ = flipuser$ + Mid$(userid, t%, 1)
Next t%
flippw$ = ""
For t% = Len(password) To 1 Step -1
    flippw$ = flippw$ + Mid$(password, t%, 1)
Next t%
Dim db As Database, ds As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nwc + dbname)
Set ds = db.OpenRecordset("select * from security where userid = " + Chr$(34) + flipuser$ + Chr$(34) + " and password = " + Chr$(34) + flippw$ + Chr$(34))
If Not ds.EOF Then
    ds.MoveFirst
    ds.Delete
Else
    msg = MsgBox("Entry not found.", 48, "Genesis Error Log")
    db.Close
    Exit Sub
End If
userid = ""
password = ""
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

Private Sub dworkaddress_KeyPress(KeyAscii As Integer)
If KeyAscii = 34 Then
    msg = MsgBox("The " + Chr$(34) + " character is not allowed.  Use the ' character instead.", 48, "Genesis Error Log")
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
    dworkaddress2.SetFocus
End If

End Sub

Private Sub dworkaddress2_KeyPress(KeyAscii As Integer)
If KeyAscii = 34 Then
    msg = MsgBox("The " + Chr$(34) + " character is not allowed.  Use the ' character instead.", 48, "Genesis Error Log")
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
    dworkstate.SetFocus
End If

End Sub

Private Sub dworkphone_GotFocus()
On Error Resume Next
If Len(dworkphone) = 0 Then
    Dim db As Database, rs As Recordset
    On Error GoTo oderror
od:
    Set db = OpenDatabase(nwl + "LAWSUITE.MDB")
    Set rs = db.OpenRecordset("SELECT AREACODE FROM DEFAULTS")
    If rs.EOF Then
        db.Close
        Exit Sub
    End If
    rs.MoveFirst
    If IsNull(rs("AREACODE")) Then
        db.Close
        Exit Sub
    End If
    If Len(rs("AREACODE")) <> 3 Then
        db.Close
        Exit Sub
    End If
    Call sendopenpara
    Call SENDCHAR(Left$(rs("AREACODE"), 1))
    Call SENDCHAR(Mid$(rs("AREACODE"), 2, 1))
    Call SENDCHAR(Mid$(rs("AREACODE"), 3, 1))
    Call SENDEND
    db.Close
End If

On Error GoTo 0

Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If


End Sub

Private Sub dworkphone_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
If Len(dworkphone) = 3 Then
    Call sendclosepara
End If
If Len(dworkphone) = 4 Then
    Call sendspace
End If
If Len(dworkphone) = 8 Then
    Call senddash
End If
If Len(dworkphone) = 13 Then
    Call sendspace
End If
End If
If KeyAscii = 13 Then
    plaintiff.SetFocus
End If


End Sub


Private Sub dworkphone_LostFocus()
If Len(dworkphone) = 5 Or Len(dworkphone) = 6 Then
    dworkphone = ""
End If

End Sub

Private Sub dworkstate_KeyPress(KeyAscii As Integer)
If KeyAscii = 34 Then
    msg = MsgBox("The " + Chr$(34) + " character is not allowed.  Use the ' character instead.", 48, "Genesis Error Log")
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
    dworkzipcode.SetFocus
End If
End Sub

Private Sub dworkzipcode_KeyPress(KeyAscii As Integer)
If KeyAscii = 34 Then
    msg = MsgBox("The " + Chr$(34) + " character is not allowed.  Use the ' character instead.", 48, "Genesis Error Log")
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
    dworkphone.SetFocus
End If
End Sub

Private Sub erlbdr_Click()
If erlbdr.Value = True Then
    fromdate.SetFocus
End If
End Sub

Private Sub estpayoffdate_GotFocus()
If Not IsDate(estpayoffdate) Then
    estpayoffdate = Format$(Date, "mm/dd/yyyy")
End If
End Sub

Private Sub estpayoffdate_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
If Len(estpayoffdate) = 1 Or Len(estpayoffdate) = 4 Then
    Call sendslash
End If
End If
If maintab.Tab = 3 Then
If KeyAscii = 13 Then
    nulla.SetFocus
End If
End If

End Sub

Private Sub estpayoffdate_LostFocus()
If Not IsDate(estpayoffdate) Then
    estpayoffdate = ""
End If
Call commissandint
End Sub

Private Sub feedate_GotFocus()
If feedate = "" And Val(servicefee) > 0 And (maintab.Tab = 0 Or maintab.Tab = 1 Or maintab.Tab = 3) Then
    feedate = datereceived
End If
End Sub

Private Sub feedate_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
If Len(feedate) = 1 Or Len(feedate) = 4 Then
    Call sendslash
End If
End If
If KeyAscii = 13 Then
   bill.SetFocus
End If

End Sub

Private Sub fnf_Click()
Call setfnln
End Sub

Private Sub Form_Load()
nametype = 0
Call checkcivil
SEARCHTYPE = 0
For t% = 0 To Forms.Count - 1
    If Forms(t%).Name = "xref" Then
        FROMXREF = "1"
        t% = Forms.Count - 1
    End If
Next t%
SAVEERR = 0
FROMG = 0
CSERVICEOF = 0
Dim r As Long
Dim Buffer As String
On Error Resume Next
FROMLF = 0
Kill "holdm"
Kill "holdw"
Kill "holdf"
Kill "holde"
On Error GoTo 0
Buffer = Space(8192)
r = GetProfileString("PrinterPorts", vbNullString, "", Buffer, Len(Buffer))
ParseList List1, Buffer
ParseList List2, Buffer
procdate = "12/31/9999"
On Error Resume Next
lname = ""
LAFFTYPE = 0
LPHONE = ""
a$ = ""
Open frmLogin.txtUserName + ".pro" For Input As #1
Line Input #1, a$
If FROMXREF = "0" Then
    If a$ > "" Then
        ltab = Val(a$)
    Else
        ltab = -1
    End If
Else
    ltab = maintab.Tab
End If
Line Input #1, d$
If d$ > "" Then
    lname = d$
End If
Line Input #1, b$
If b$ > "" Then
    LAFFTYPE = Val(b$)
End If
Line Input #1, c$
If c$ > "" Then
    LPHONE = c$
End If
Close #1
If lname = "" Then
    lname = "Civil Process Officer"
End If
If Val(a$) = 1 Then
    sfpay = True
Else
    sfpay = False
End If
If sfpay Then
    Label40 = "Commission               Current                 Principal Balance:                   Per Day:              Balance:  Interest                                                            Balance:                         Total (w/ Service Fee):"
    Label43 = "DATE PAID    AMOUNT      RECEIPT      CHECK       PRINCIPAL     COMMISSION INTEREST    SERVICE FEE  REMARKS" + Space$(3720) + "DATE PAID    AMOUNT     RECEIPT      CHECK          PRINCIPAL  COMMISSION INTEREST   SERVICE FEE REMARKS"
    eservicefee.Visible = True
Else
    Label40 = "Commission               Current                 Principal Balance:                   Per Day:              Balance:  Interest                                                            Balance:                    Total (w/out Service Fee):"
    Label43 = "DATE PAID    AMOUNT      RECEIPT      CHECK       PRINCIPAL     COMMISSION INTEREST                       REMARKS" + Space$(3720) + "DATE PAID    AMOUNT     RECEIPT      CHECK          PRINCIPAL  COMMISSION INTEREST                       REMARKS"
    eservicefee.Visible = True
End If
If mainform.mcurrent.checked = True Then
    dbname = "civil.mdb"
Else
    dbname = "arccivil.mdb"
End If
stopspool = 0
lastserviceof = ""
maintab.Tab = 0
Call loadpapertype
Call loadprof
Call loaddeputy
Call loadsystem
Call loaddp
omag.Value = True
Call loadprof
sofn = ""
soln = ""
somi = ""
soo = ""
infoframe.Top = 480
infoframe.Left = 50
infoframe.Visible = True
FROMP = 0
If mainform.marchived.checked = True Then
    rprintbutton.Enabled = False
    Command7.Enabled = False
    printbutton.Enabled = False
Else
    rprintbutton.Enabled = True
    Command7.Enabled = True
    printbutton.Enabled = True
End If
If ltab > -1 Then
    maintab.Tab = ltab
End If
If Val(frmLogin.CBROWSE(0)) = 0 Then
    If Val(frmLogin.CBROWSE(1)) = 0 Then
        If Val(frmLogin.CBROWSE(2)) = 0 Then
            If Val(frmLogin.CBROWSE(3)) = 0 Then
                maintab.Tab = 6
            Else
                maintab.Tab = 3
            End If
        Else
            maintab.Tab = 2
        End If
    Else
        maintab.Tab = 1
    End If
Else
If ltab = -1 Then
    maintab.Tab = 0
End If
End If
On Error GoTo nocolor
Dim aa, bb As Long, tofrom(1000, 2) As String, ct As Integer
ct = 0
Open "cc.tag" For Input As #1
While Not EOF(1)
    Input #1, aa, bb
    ct = ct + 1
    tofrom(ct, 1) = aa
    tofrom(ct, 2) = bb
Wend
Close #1
On Error GoTo errrtn
For i = 0 To CIVIL.Controls.Count - 1
    a = CIVIL.Controls(i).Name
    For j = 1 To ct
        If CIVIL.Controls(i).ForeColor = tofrom(j, 1) Then
            CIVIL.Controls(i).ForeColor = tofrom(j, 2)
            CIVIL.Controls(i).Refresh
            j = ct
        End If
    Next j
ni:
Next i
getoutc:
maintab.Tab = 0
Exit Sub
nocolor:
    Resume getoutc
errrtn:
If Err = 438 Or Err = 458 Then
    Resume ni
End If
Resume Next
End Sub

Private Sub Form_Paint()
serviceof.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
Screen.MousePointer = 0
Set CIVIL = Nothing
End Sub

Private Sub fromdate_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
If Len(fromdate) = 1 Or Len(fromdate) = 4 Then
    Call sendslash
End If
End If
If KeyAscii = 13 Then
    todate.SetFocus
End If


End Sub

Private Sub goprint_Click()
On Error GoTo 0
Dim db, db2 As Database, ds, rs, rs2 As Recordset
If Val(frmLogin.CPRINT(maintab.Tab)) = 0 And Val(frmLogin.CSUPERVISOR(maintab.Tab)) = 0 Then
    msg = MsgBox("You have insufficient authority to print.", 48, "Genesis Error Log")
    Exit Sub
End If
FROMP = 0
If levyp.Value = True Then
    If levy.Text = "" Then
        msg = MsgBox("You must use the REMARKS button to enter the verbiage for the Notice of Levy.", 48, "Genesis Error Log")
        Exit Sub
    End If
End If
If affidavit.Value = True Then
    If served.Value = False And nonservice.Value = False Then
        If Dir(nwc + "c?affidav.rpt") = "" Then
            msg = MsgBox("Neither SERVED not NON-SERVICE has been checked. ", 48, "Genesis Error Log")
            Exit Sub
        End If
    End If
    If served.Value = True And (personserved = "" Or locationserved = "") Then
        If Dir(nwc + "c?affidav.rpt") = "" Then
            msg = MsgBox("Both PERSON SERVED and LOCATION SERVED must be entered.", 48, "Genesis Error Log")
            Exit Sub
        End If
    End If
    If served.Value = True And (Not IsDate(servicedate) Or servicetime = "") Then
        If Dir(nwc + "c?affidav.rpt") = "" Then
            msg = MsgBox("Both SERVICE DATE and TIME must be entered.", 48, "Genesis Error Log")
            Exit Sub
        End If
    End If
    If nonservice.Value = True And Not IsDate(servicedate) Then
        If Dir(nwc + "c?affidav.rpt") = "" Then
            msg = MsgBox("SERVICE DATE must be entered.", 48, "Genesis Error Log")
            Exit Sub
        End If
    End If
End If
If serviceof = "" Or Not IsDate(datereceived) Then
    msg = MsgBox("SERVICE OF and DATE RECEIVED must be entered and valid.", 48, "Genesis Error Log")
    Exit Sub
End If
FROMG = 1
SAVEERR = 0
If UCase(frmLogin.txtUserName) = "DEMO" And UCase(frmLogin.txtPassword) = "DEMO" Then
Else
    Call savebutton_Click
End If
If SAVEERR = 1 Then
    Screen.MousePointer = 0
    Exit Sub
End If
Screen.MousePointer = 11
Y$ = Right$(Format$(datereceived, "mmddyyyy"), 4)
m$ = Left$(Format$(datereceived, "mmddyyyy"), 2)
d$ = Mid$(Format$(datereceived, "mmddyyyy"), 3, 2)
If maintab.Tab = 0 Then
    report.SelectionFormula = "{magistrate.serviceof} = " + Chr$(34) + serviceof + Chr$(34) + " and {magistrate.datereceived} = date(" + Y$ + "," + m$ + "," + d$ + ") AND {MAGISTRATE.ITERATION} = " + Chr$(34) + iteration + Chr$(34)
End If
If maintab.Tab = 1 Then
    report.SelectionFormula = "{writother.serviceof} = " + Chr$(34) + serviceof + Chr$(34) + " and {writother.datereceived} = date(" + Y$ + "," + m$ + "," + d$ + ") AND {WRITOTHER.ITERATION} = " + Chr$(34) + iteration + Chr$(34)
End If
If maintab.Tab = 2 Then
    report.SelectionFormula = "{familycourt.serviceof} = " + Chr$(34) + serviceof + Chr$(34) + " and {familycourt.datereceived} = date(" + Y$ + "," + m$ + "," + d$ + ") AND {FAMILYCOURT.ITERATION} = " + Chr$(34) + iteration + Chr$(34)
End If
If maintab.Tab = 3 Then
    report.SelectionFormula = "{executions.serviceof} = " + Chr$(34) + serviceof + Chr$(34) + " and {executions.datereceived} = date(" + Y$ + "," + m$ + "," + d$ + ") AND {EXECUTIONS.ITERATION} = " + Chr$(34) + iteration + Chr$(34)
End If
If levyp.Value = True Then
    If List1.ListCount > 1 And List1.ListIndex > -1 Then
        Call defaultprinter(List1.List(List1.ListIndex))
    End If
    If List1.ListIndex > -1 And List2.ListIndex > -1 Then
        If List1.List(List1.ListIndex) <> List2.List(List2.ListIndex) And prepareprinter = 1 Then
            msg = MsgBox("Prepare for Worksheet/Affidavit/Letter/Report Printing.", 48, "Genesis Error Log")
        End If
    End If
    If maintab.Tab = 3 Then
        report.Destination = crptToPrinter
        report.CopiesToPrinter = 1
        report.ReportFileName = nwc + "levy.rpt"
        report.Action = 1
    End If
End If
If status.Value = True Then
    If List1.ListCount > 1 And List1.ListIndex > -1 Then
        Call defaultprinter(List1.List(List1.ListIndex))
    End If
    If List1.ListIndex > -1 And List2.ListIndex > -1 Then
        If List1.List(List1.ListIndex) <> List2.List(List2.ListIndex) And prepareprinter = 1 Then
            msg = MsgBox("Prepare for Worksheet/Affidavit/Letter/Report Printing.", 48, "Genesis Error Log")
        End If
    End If
    Select Case maintab.Tab
        Case 0
            inp = InputBox("Should this letter be addressed to the Plaintiff or to the Magistrate? (P/M)", "Genesis Information Log", "P")
            inp = UCase(inp)
            If inp <> "M" And inp <> "P" Then
                msg = MsgBox("Invalid selection.", 48, "Genesis Error Log")
                Screen.MousePointer = 0
                Exit Sub
            End If
            If inp = "M" Then
                On Error GoTo oderror1
od1:
                Set db = OpenDatabase(nwl + "lawsuite.mdb")
                Set rs = db.OpenRecordset("select * from professionals where profname = " + Chr$(34) + professional + Chr$(34) + " and type = 'M'")
                If rs.EOF Then
                    inp1 = InputBox("Enter the first address line for the Magistrate.", "Genesis Information Log", "")
                    inp2 = InputBox("Enter the second address line for the Magistrate.", "Genesis Information Log", "")
                    rs.AddNew
                    rs("profname") = professional
                    rs("type") = "M"
                    rs("profaddr1") = inp1
                    rs("profaddr2") = inp2
                    rs.Update
                Else
                    rs.MoveFirst
                    inp1 = ""
                    inp2 = ""
                    If Not IsNull(rs("profaddr1")) Then
                        inp1 = rs("profaddr1")
                    End If
                    If Not IsNull(rs("profaddr2")) Then
                        inp2 = rs("profaddr2")
                    End If
                    If inp1 = "" And inp2 = "" Then
                        inp1 = InputBox("Enter the first address line for the Magistrate.", "Genesis Information Log", "")
                        inp2 = InputBox("Enter the second address line for the Magistrate.", "Genesis Information Log", "")
                        rs.Edit
                        rs("profaddr1") = inp1
                        rs("profaddr2") = inp2
                        rs.Update
                    End If
                End If
                db.Close
                GoSub letterheader
                Printer.Print Tab(10); professional
                Printer.Print Tab(10); inp1
                Printer.Print Tab(10); inp2
                GoSub letterbody
                db.Close
            Else
                GoSub letterheader
                Printer.Print Tab(10); plaintiff
                Printer.Print Tab(10); phomeaddress
                Printer.Print Tab(10); phomeaddress2 + " " + phomestate + " " + phomezipcode
                GoSub letterbody
                db.Close
            End If
        Case 1, 3
            inp = InputBox("Should this letter be addressed to the Plaintiff or to the Attorney? (P/A)", "Genesis Information Log", "P")
            inp = UCase(inp)
            If inp <> "A" And inp <> "P" Then
                msg = MsgBox("Invalid selection.", 48, "Genesis Error Log")
                Screen.MousePointer = 0
                Exit Sub
            End If
            If inp = "A" Then
                On Error GoTo oderror2
od2:
                Set db = OpenDatabase(nwl + "lawsuite.mdb")
                Set rs = db.OpenRecordset("select * from professionals where profname = " + Chr$(34) + professional + Chr$(34) + " and type = 'A'")
                If rs.EOF Then
                    inp1 = InputBox("Enter the first address line for the Attorney.", "Genesis Information Log", "")
                    inp2 = InputBox("Enter the second address line for the Attorney.", "Genesis Information Log", "")
                    rs.AddNew
                    rs("profname") = professional
                    rs("type") = "A"
                    rs("profaddr1") = inp1
                    rs("profaddr2") = inp2
                    rs.Update
                Else
                    rs.MoveFirst
                    inp1 = ""
                    inp2 = ""
                    If Not IsNull(rs("profaddr1")) Then
                        inp1 = rs("profaddr1")
                    End If
                    If Not IsNull(rs("profaddr2")) Then
                        inp2 = rs("profaddr2")
                    End If
                    If inp1 = "" And inp2 = "" Then
                        inp1 = InputBox("Enter the first address line for the Attorney.", "Genesis Information Log", "")
                        inp2 = InputBox("Enter the second address line for the Attorney.", "Genesis Information Log", "")
                        rs.Edit
                        rs("profaddr1") = inp1
                        rs("profaddr2") = inp2
                        rs.Update
                    End If
                End If
                db.Close
                GoSub letterheader
                Printer.Print Tab(10); professional
                Printer.Print Tab(10); inp1
                Printer.Print Tab(10); inp2
                GoSub letterbody
            Else
                GoSub letterheader
                Printer.Print Tab(10); plaintiff
                Printer.Print Tab(10); phomeaddress
                Printer.Print Tab(10); phomeaddress2 + " " + phomestate + " " + phomezipcode
                GoSub letterbody
            End If
        Case 2
            inp = InputBox("Should this letter be addressed to the Plaintiff or to the Court? (P/C)", "Genesis Information Log", "P")
            inp = UCase(inp)
            If inp <> "C" And inp <> "P" Then
                msg = MsgBox("Invalid selection.", 48, "Genesis Error Log")
                Screen.MousePointer = 0
                Exit Sub
            End If
            If inp = "A" Then
                On Error GoTo oderror3
od3:
                Set db = OpenDatabase(nwl + "lawsuite.mdb")
                Set rs = db.OpenRecordset("select * from professionals where profname = " + Chr$(34) + professional + Chr$(34) + " and type = 'C'")
                If rs.EOF Then
                    inp1 = InputBox("Enter the first address line for the Court.", "Genesis Information Log", "")
                    inp2 = InputBox("Enter the second address line for the Court.", "Genesis Information Log", "")
                    rs.AddNew
                    rs("profname") = professional
                    rs("type") = "C"
                    rs("profaddr1") = inp1
                    rs("profaddr2") = inp2
                    rs.Update
                Else
                    rs.MoveFirst
                    inp1 = ""
                    inp2 = ""
                    If Not IsNull(rs("profaddr1")) Then
                        inp1 = rs("profaddr1")
                    End If
                    If Not IsNull(rs("profaddr2")) Then
                        inp2 = rs("profaddr2")
                    End If
                    If inp1 = "" And inp2 = "" Then
                        inp1 = InputBox("Enter the first address line for the Court.", "Genesis Information Log", "")
                        inp2 = InputBox("Enter the second address line for the Court.", "Genesis Information Log", "")
                        rs.Edit
                        rs("profaddr1") = inp1
                        rs("profaddr2") = inp2
                        rs.Update
                    End If
                End If
                db.Close
                GoSub letterheader
                Printer.Print Tab(10); professional
                Printer.Print Tab(10); inp1
                Printer.Print Tab(10); inp2
                GoSub letterbody
            Else
                GoSub letterheader
                Printer.Print Tab(10); plaintiff
                Printer.Print Tab(10); phomeaddress
                Printer.Print Tab(10); phomeaddress2 + " " + phomestate + " " + phomezipcode
                GoSub letterbody
            End If
    End Select
End If
If worksheet.Value = True Then
    If List1.ListCount > 1 And List1.ListIndex > -1 Then
        Call defaultprinter(List1.List(List1.ListIndex))
    End If
    If List1.ListIndex > -1 And List2.ListIndex > -1 Then
        If List1.List(List1.ListIndex) <> List2.List(List2.ListIndex) And prepareprinter = 1 Then
            msg = MsgBox("Prepare for Worksheet/Affidavit/Letter/Report Printing.", 48, "Genesis Error Log")
        End If
    End If
    If mugshot.Picture > 0 Then
        SavePicture mugshot.Picture, "c:\mug.jpg"
    Else
        FileCopy nwl + "blank.jpg", "c:\mug.jpg"
    End If
    If maintab.Tab = 0 Then
        report.ReportFileName = nwc + "mworksht.rpt"
        report.Destination = crptToPrinter
        report.CopiesToPrinter = 1
        report.Action = 1
    End If
    If maintab.Tab = 1 Then
        report.Destination = crptToPrinter
        report.CopiesToPrinter = 1
        report.ReportFileName = nwc + "wworksht.rpt"
        report.Action = 1
    End If
    If maintab.Tab = 2 Then
        report.Destination = crptToPrinter
        report.CopiesToPrinter = 1
        report.ReportFileName = nwc + "fworksht.rpt"
        report.Action = 1
    End If
    If maintab.Tab = 3 Then
        report.Destination = crptToPrinter
        report.CopiesToPrinter = 1
        report.ReportFileName = nwc + "eworksht.rpt"
        report.Action = 1
    End If
End If
If preceipt.Value = True Then
    If List2.ListCount > 1 And List2.ListIndex > -1 Then
        Call defaultprinter(List2.List(List2.ListIndex))
    End If
    If List1.ListIndex > -1 And List2.ListIndex > -1 Then
        If List1.List(List1.ListIndex) <> List2.List(List2.ListIndex) And prepareprinter = 1 Then
            msg = MsgBox("Prepare for Receipt/Check Printing.", 48, "Genesis Error Log")
        End If
    End If

    fee = Val(servicefee)
    rnumber = receiptd
    cnumber = checkd
    receiptframe.Left = 1000
    receiptframe.Top = 2000
    Call LOADOTHER
    If fromdefendant = 0 And fromplaintiff = 0 And othername = "" Then
        fromdefendant = 1
    End If
    receiptframe.Visible = True
    Screen.MousePointer = 0
    othername.SetFocus
    Exit Sub
End If
If letter.Value = True Then
    If List1.ListCount > 1 And List1.ListIndex > -1 Then
        Call defaultprinter(List1.List(List1.ListIndex))
    End If
    If List1.ListIndex > -1 And List2.ListIndex > -1 Then
        If List1.List(List1.ListIndex) <> List2.List(List2.ListIndex) And prepareprinter = 1 Then
            msg = MsgBox("Prepare for Worksheet/Affidavit/Letter/Report Printing.", 48, "Genesis Error Log")
        End If
    End If
    On Error GoTo oderror4
od4:
    Set db = OpenDatabase(nwc + dbname)
    Set ds = db.OpenRecordset("select * from passthru")
    If ds.EOF Then
        ds.AddNew
    Else
        ds.MoveFirst
        ds.Edit
    End If
    If Len(lname) > 30 Then
        lname = Left$(lname, 30)
    Else
        lname = lname + Space$(30 - Len(lname))
    End If
    ds("textstring") = lname + Space$(20)
    If maintab.Tab = 3 And LPHONE > "" Then
        If Len(LPHONE) > 20 Then
            LPHONE = Left$(LPHONE, 20)
        Else
            LPHONE = LPHONE + Space$(20 - Len(LPHONE))
        End If
        ds("TEXTSTRING") = lname + LPHONE
    End If
    ds.Update
    db.Close
    If maintab.Tab = 0 Then
        report.Destination = crptToPrinter
        report.CopiesToPrinter = 1
        report.ReportFileName = nwc + "mletter.rpt"
        report.Action = 1
    End If
    If maintab.Tab = 1 Then
        report.Destination = crptToPrinter
        report.CopiesToPrinter = 1
        report.ReportFileName = nwc + "wletter.rpt"
        report.Action = 1
    End If
    If maintab.Tab = 2 Then
        report.Destination = crptToPrinter
        report.CopiesToPrinter = 1
        report.ReportFileName = nwc + "fletter.rpt"
        report.Action = 1
    End If
    If maintab.Tab = 3 Then
        report.Destination = crptToPrinter
        report.CopiesToPrinter = 1
        report.ReportFileName = nwc + "eletter.rpt"
        report.Action = 1
    End If
End If
If RL.Value = True Then
    If List1.ListCount > 1 And List1.ListIndex > -1 Then
        Call defaultprinter(List1.List(List1.ListIndex))
    End If
    If List1.ListIndex > -1 And List2.ListIndex > -1 Then
        If List1.List(List1.ListIndex) <> List2.List(List2.ListIndex) And prepareprinter = 1 Then
            msg = MsgBox("Prepare for Worksheet/Affidavit/Letter/Report Printing.", 48, "Genesis Error Log")
        End If
    End If
    On Error GoTo oderror5
od5:
    Set db = OpenDatabase(nwc + dbname)
    Set ds = db.OpenRecordset("select * from passthru")
    If ds.EOF Then
        ds.AddNew
    Else
        ds.MoveFirst
        ds.Edit
    End If
    ds("textstring") = lname
    ds.Update
    db.Close
    report.Destination = crptToPrinter
    report.CopiesToPrinter = 1
    report.ReportFileName = nwc + "eletterr.rpt"
    report.Action = 1
End If
If nrl.Value = True Then
    If List1.ListCount > 1 And List1.ListIndex > -1 Then
        Call defaultprinter(List1.List(List1.ListIndex))
    End If
    If List1.ListIndex > -1 And List2.ListIndex > -1 Then
        If List1.List(List1.ListIndex) <> List2.List(List2.ListIndex) And prepareprinter = 1 Then
            msg = MsgBox("Prepare for Worksheet/Affidavit/Letter/Report Printing.", 48, "Genesis Error Log")
        End If
    End If
    On Error GoTo oderror6
od6:
    Set db = OpenDatabase(nwc + dbname)
    Set ds = db.OpenRecordset("select * from passthru")
    If ds.EOF Then
        ds.AddNew
    Else
        ds.MoveFirst
        ds.Edit
    End If
    ds("textstring") = lname
    ds.Update
    db.Close
    report.Destination = crptToPrinter
    report.CopiesToPrinter = 1
    report.ReportFileName = nwc + "elettern.rpt"
    report.Action = 1
End If
If affidavit.Value = True Then
    If List1.ListCount > 1 And List1.ListIndex > -1 Then
        Call defaultprinter(List1.List(List1.ListIndex))
    End If
    If List1.ListIndex > -1 And List2.ListIndex > -1 Then
        If List1.List(List1.ListIndex) <> List2.List(List2.ListIndex) And prepareprinter = 1 Then
            msg = MsgBox("Prepare for Worksheet/Affidavit/Letter/Report Printing.", 48, "Genesis Error Log")
        End If
    End If
    
    If maintab.Tab = 0 Then
        report.Destination = crptToPrinter
        report.CopiesToPrinter = 1
        If Dir(nwc + "CMAFFIDAV.RPT") > "" Then
            report.ReportFileName = nwc + "Cmaffidav.rpt"
        Else
            report.ReportFileName = nwc + "maffidav.rpt"
        End If
        report.Action = 1
    End If
    If maintab.Tab = 1 Then
        report.Destination = crptToPrinter
        report.CopiesToPrinter = 1
        If Dir(nwc + "CWAFFIDAV.RPT") > "" Then
            report.ReportFileName = nwc + "CWaffidav.rpt"
        Else
            report.ReportFileName = nwc + "Waffidav.rpt"
        End If

        report.Action = 1
    End If
    If maintab.Tab = 2 Then
        report.Destination = crptToPrinter
        report.CopiesToPrinter = 1
        If Dir(nwc + "CFAFFIDAV.RPT") > "" Then
            report.ReportFileName = nwc + "CFaffidav.rpt"
        Else
            report.ReportFileName = nwc + "Faffidav.rpt"
        End If
        report.Action = 1
    End If
    If maintab.Tab = 3 Then
        report.Destination = crptToPrinter
        report.CopiesToPrinter = 1
        If Dir(nwc + "CEAFFIDAV.RPT") > "" Then
            report.ReportFileName = nwc + "CEaffidav.rpt"
        Else
            report.ReportFileName = nwc + "Eaffidav.rpt"
        End If
        report.Action = 1
    End If
End If
If epwb.Value = True Then
    If List1.ListCount > 1 And List1.ListIndex > -1 Then
        Call defaultprinter(List1.List(List1.ListIndex))
    End If
    If List1.ListIndex > -1 And List2.ListIndex > -1 Then
        If List1.List(List1.ListIndex) <> List2.List(List2.ListIndex) And prepareprinter = 1 Then
            msg = MsgBox("Prepare for Worksheet/Affidavit/Letter/Report Printing.", 48, "Genesis Error Log")
        End If
    End If

    report.Destination = crptToPrinter
    report.CopiesToPrinter = 1
    report.ReportFileName = nwc + "epw1.rpt"
    report.SelectionFormula = "{executions.serviceof} = " + Chr$(34) + serviceof + Chr$(34) + " and {executions.datereceived} = Date(" + Y$ + "," + m$ + "," + d$ + ") and {executions.iteration} = " + Chr$(34) + iteration + Chr$(34)
    report.Action = 1
    report.SelectionFormula = ""
    report.ReportFileName = nwc + "epw2.rpt"
    report.Action = 1
End If
If asb.Value = True Then
    If List1.ListCount > 1 And List1.ListIndex > -1 Then
        Call defaultprinter(List1.List(List1.ListIndex))
    End If
    If List1.ListIndex > -1 And List2.ListIndex > -1 Then
        If List1.List(List1.ListIndex) <> List2.List(List2.ListIndex) And prepareprinter = 1 Then
            msg = MsgBox("Prepare for Worksheet/Affidavit/Letter/Report Printing.", 48, "Genesis Error Log")
        End If
    End If
    On Error GoTo oderror7
od7:
    Set db = OpenDatabase(nwc + dbname)
    Set rs = db.OpenRecordset("select * from executions where serviceof = " + Chr$(34) + serviceof + Chr$(34) + " and datereceived = #" + Format$(datereceived, "mm/dd/yyyy") + "# and iteration = " + Chr$(34) + iteration + Chr$(34))
    If Not rs.EOF Then
        rs.MoveFirst
        Printer.FontName = "Times New Roman"
        Printer.FontSize = 14
        Printer.FontBold = True
        Printer.Print "Execution Account Statement"; Tab(70); "As Of  ";
        Printer.FontBold = False
        Printer.Print rs("estpayoffdate")
        Printer.FontSize = 10
        Printer.Print Format$(Date$, "mm/dd/yyyy")
        Printer.Print
        Printer.FontBold = True
        Printer.FontUnderline = True
        Printer.Print "Defendant";
        Printer.FontBold = False
        Printer.FontUnderline = False
        Printer.Print Tab(20); Left$(rs("defendant"), 30);
        Printer.FontBold = True
        Printer.FontUnderline = True
        Printer.Print Tab(55); "Judgement Date";
        Printer.FontBold = False
        Printer.FontUnderline = False
        Printer.Print Tab(85); Format$(rs("judgementdate"), "mm/dd/yyyy");
        Printer.FontBold = True
        Printer.FontUnderline = True
        Printer.Print Tab(95); "Principal Balance";
        Printer.FontBold = False
        Printer.FontUnderline = False
        Printer.Print Tab(125); Format$(rs("balance"), "$#########0.00")
        Printer.Print
        Printer.FontBold = True
        Printer.FontUnderline = True
        Printer.Print "Case Number";
        Printer.FontBold = False
        Printer.FontUnderline = False
        Printer.Print Tab(20); rs("casenumber");
        Printer.FontBold = True
        Printer.FontUnderline = True
        Printer.Print Tab(55); "Judgement Amount";
        Printer.FontBold = False
        Printer.FontUnderline = False
        Printer.Print Tab(85); Format$(rs("judgementamount"), "$#########0.00");
        Printer.FontBold = True
        Printer.FontUnderline = True
        Printer.Print Tab(95); "Interest Balance";
        Printer.FontBold = False
        Printer.FontUnderline = False
        Printer.Print Tab(125); Format$(rs("INTEREST"), "$#########0.00")
        Printer.Print
        Printer.FontBold = True
        Printer.FontUnderline = True
        Printer.Print Tab(55); "Service Fee:";
        Printer.FontBold = False
        Printer.FontUnderline = False
        If sfpay Then
            Printer.Print Tab(85); Format$(rs("servicefee"), "$#########0.00");
        Else
            Printer.Print Tab(85); "N/A"
        End If
        Printer.FontBold = True
        Printer.FontUnderline = True
        Printer.Print Tab(95); "Commission Balance";
        Printer.FontBold = False
        Printer.FontUnderline = False
        Printer.Print Tab(125); Format$(rs("commission"), "$#########0.00")
        Printer.Print
        Printer.FontBold = True
        Printer.FontUnderline = True
        Set rs2 = db.OpenRecordset("select * from executionspay where serviceof = " + Chr$(34) + serviceof + Chr$(34) + " and datereceived = #" + Format$(datereceived, "mm/dd/yyyy") + "# and iteration = " + Chr$(34) + iteration + Chr$(34))
        TOTSF = 0
        If Not rs2.EOF Then
            rs2.MoveFirst
            While Not rs2.EOF
                If Not IsNull(rs2("servicefee")) Then
                    TOTSF = TOTSF + rs2("SERVICEFEE")
                End If
                rs2.MoveNext
            Wend
        End If
        Printer.Print Tab(95); "TOTAL DUE";
        Printer.FontBold = False
        Printer.FontUnderline = False
        If sfpay Then
            Printer.Print Tab(125); Format$(rs("commission") + rs("INTEREST") + rs("balance") + rs("SERVICEFEE") - TOTSF, "$#########0.00")
        Else
            Printer.Print Tab(125); Format$(rs("commission") + rs("INTEREST") + rs("balance"), "$#########0.00")
        End If
        Printer.Print
        Printer.Print
        Printer.FontBold = True
        Printer.FontSize = 12
        Printer.Print "PAYMENT HISTORY"
        Printer.Print
        Printer.FontBold = True
        Printer.FontUnderline = True
        Printer.FontSize = 10
        Printer.Print "Date Paid"; Tab(20); "Receipt"; Tab(40); "Amount"; Tab(60); "Principal Portion"; Tab(80); "Interest Portion"; Tab(97); "Commission Portion"; Tab(120); "Service Fee"
        Printer.FontUnderline = False
        Set rs2 = db.OpenRecordset("select * from executionspay where serviceof = " + Chr$(34) + serviceof + Chr$(34) + " and datereceived = #" + Format$(datereceived, "mm/dd/yyyy") + "# and iteration = " + Chr$(34) + iteration + Chr$(34))
        If Not rs2.EOF Then
            rs2.MoveFirst
            While Not rs2.EOF
                If Not IsNull(rs2("servicefee")) Then
                    Printer.Print Format$(rs2("datepaid"), "mm/dd/yyyy"); Tab(20); rs2("receipt"); Tab(40); Format$(rs2("amount"), "$#########0.00"); Tab(60); Format$(rs2("principal"), "$#########0.00"); Tab(80); Format$(rs2("inter"), "$#########0.00"); Tab(97); Format$(rs2("commiss"), "$#########0.00"); Tab(120); rs2("servicefee")
                Else
                    Printer.Print Format$(rs2("datepaid"), "mm/dd/yyyy"); Tab(20); rs2("receipt"); Tab(40); Format$(rs2("amount"), "$#########0.00"); Tab(60); Format$(rs2("principal"), "$#########0.00"); Tab(80); Format$(rs2("inter"), "$#########0.00"); Tab(97); Format$(rs2("commiss"), "$#########0.00")
                End If
                rs2.MoveNext
            Wend
        End If
        Printer.EndDoc
    End If
    db.Close
End If
If sat.Value = True Then
    If List1.ListCount > 1 And List1.ListIndex > -1 Then
        Call defaultprinter(List1.List(List1.ListIndex))
    End If
    If List1.ListIndex > -1 And List2.ListIndex > -1 Then
        If List1.List(List1.ListIndex) <> List2.List(List2.ListIndex) And prepareprinter = 1 Then
            msg = MsgBox("Prepare for Worksheet/Affidavit/Letter/Report Printing.", 48, "Genesis Error Log")
        End If
    End If
    Set db = OpenDatabase(nwc + dbname)
    Set rs2 = db.OpenRecordset("select sum(amount) as ap from executionspay where serviceof = " + Chr$(34) + serviceof + Chr$(34) + " and datereceived = #" + Format$(datereceived, "mm/dd/yyyy") + "# and iteration = " + Chr$(34) + iteration + Chr$(34) + " and (payrem like 'WRITEOFF*' or payrem like 'writeoff*')")
    totap = 0
    If Not rs2.EOF Then
        rs2.MoveFirst
        If Not IsNull(rs2("ap")) Then
            totap = rs2("ap")
        End If
    End If
    Set ds = db.OpenRecordset("select * from passthru")
    If ds.EOF Then
        ds.AddNew
    Else
        ds.MoveFirst
        ds.Edit
    End If
    If totap > 0 Then
        ds("textstring") = "(Includes a write-off amount of " + Format$(totap, "$######0.00") + ")"
    Else
        ds("textstring") = ""
    End If
    ds.Update
    db.Close
    report.Destination = crptToPrinter
    report.CopiesToPrinter = 1
    If sfpay Then
        report.ReportFileName = nwc + "satisfy.rpt"
    Else
        report.ReportFileName = nwc + "satisfy2.rpt"
    End If
    report.SelectionFormula = "{executions.serviceof} = " + Chr$(34) + serviceof + Chr$(34) + " and {executions.datereceived} = Date(" + Y$ + "," + m$ + "," + d$ + ") and {executions.iteration} = " + Chr$(34) + iteration + Chr$(34)
    report.Action = 1
End If
If Partial.Value = True Then
    If List1.ListCount > 1 And List1.ListIndex > -1 Then
        Call defaultprinter(List1.List(List1.ListIndex))
    End If
    If List1.ListIndex > -1 And List2.ListIndex > -1 Then
        If List1.List(List1.ListIndex) <> List2.List(List2.ListIndex) And prepareprinter = 1 Then
            msg = MsgBox("Prepare for Worksheet/Affidavit/Letter/Report Printing.", 48, "Genesis Error Log")
        End If
    End If
    Set db = OpenDatabase(nwc + dbname)
    Set rs2 = db.OpenRecordset("select sum(amount) as ap from executionspay where serviceof = " + Chr$(34) + serviceof + Chr$(34) + " and datereceived = #" + Format$(datereceived, "mm/dd/yyyy") + "# and iteration = " + Chr$(34) + iteration + Chr$(34) + " and (payrem like 'WRITEOFF*' or payrem like 'writeoff*')")
    totap = 0
    If Not rs2.EOF Then
        rs2.MoveFirst
        If Not IsNull(rs2("ap")) Then
            totap = rs2("ap")
        End If
    End If
    Set ds = db.OpenRecordset("select * from passthru")
    If ds.EOF Then
        ds.AddNew
    Else
        ds.MoveFirst
        ds.Edit
    End If
    If totap > 0 Then
        ds("textstring") = "(Includes a write-off amount of " + Format$(totap, "$######0.00") + ")"
    Else
        ds("textstring") = ""
    End If
    ds.Update
    db.Close
    report.Destination = crptToPrinter
    report.CopiesToPrinter = 1
    If sfpay Then
        report.ReportFileName = nwc + "psatisfy.rpt"
    Else
        report.ReportFileName = nwc + "psatisfy2.rpt"
    End If
    report.SelectionFormula = "{executions.serviceof} = " + Chr$(34) + serviceof + Chr$(34) + " and {executions.datereceived} = Date(" + Y$ + "," + m$ + "," + d$ + ") and {executions.iteration} = " + Chr$(34) + iteration + Chr$(34)
    report.Action = 1
End If
If nullaex.Value = True Then
    If List1.ListCount > 1 And List1.ListIndex > -1 Then
        Call defaultprinter(List1.List(List1.ListIndex))
    End If
    If List1.ListIndex > -1 And List2.ListIndex > -1 Then
        If List1.List(List1.ListIndex) <> List2.List(List2.ListIndex) And prepareprinter = 1 Then
            msg = MsgBox("Prepare for Worksheet/Affidavit/Letter/Report Printing.", 48, "Genesis Error Log")
        End If
    End If
    report.Destination = crptToPrinter
    report.CopiesToPrinter = 1
    If sfpay Then
        report.ReportFileName = nwc + "nullaex.rpt"
    Else
        report.ReportFileName = nwc + "nullaex2.rpt"
    End If
    report.SelectionFormula = "{executions.serviceof} = " + Chr$(34) + serviceof + Chr$(34) + " and {executions.datereceived} = Date(" + Y$ + "," + m$ + "," + d$ + ") and {executions.iteration} = " + Chr$(34) + iteration + Chr$(34)
    report.Action = 1
End If
If dos.Value = True Then
    If List1.ListCount > 1 And List1.ListIndex > -1 Then
        Call defaultprinter(List1.List(List1.ListIndex))
    End If
    If List1.ListIndex > -1 And List2.ListIndex > -1 Then
        If List1.List(List1.ListIndex) <> List2.List(List2.ListIndex) And prepareprinter = 1 Then
            msg = MsgBox("Prepare for Worksheet/Affidavit/Letter/Report Printing.", 48, "Genesis Error Log")
        End If
    End If
        linect% = 0
        Printer.FontName = "Times New Roman"
        Printer.FontSize = 16
        Printer.FontBold = True
        Printer.Print
        Call cp(office)
        Printer.FontSize = 12
        Call cp(sheriff)
        Printer.FontBold = False
        Printer.FontSize = 10
        Printer.Print
        Printer.Print
        Printer.Print
        Printer.Print Tab(10); Format$(Date$, "mm/dd/yyyy")
        Printer.Print
        Printer.Print
        If professional = "" Then
            inp = InputBox("Enter the name to address the letter.", "Genesis Information Log", "")
        Else
            inp = professional
        End If
        Printer.Print Tab(10); inp
        If professional > "" Then
            On Error GoTo oderror8
od8:
            Set db2 = OpenDatabase(nwl + "lawsuite.mdb")
            Set rs2 = db2.OpenRecordset("select * from professionals where profname = " + Chr$(34) + professional + Chr$(34) + " and type = 'A'")
            If Not rs2.EOF Then
                rs2.MoveFirst
                rs2.Edit
            Else
                rs2.AddNew
            End If
            If IsNull(rs2("profaddr1")) Or rs2("profaddr1") = "" Then
                inp1 = InputBox("Enter the first address line for the Attorney.", "Genesis Information Log", "")
            Else
                inp1 = rs2("profaddr1")
            End If
            If IsNull(rs2("profaddr2")) Or rs2("profaddr2") = "" Then
                inp2 = InputBox("Enter the second address line for the Attorney.", "Genesis Information Log", "")
            Else
                inp2 = rs2("profaddr2")
            End If
            rs2("profname") = professional
            rs2("type") = "A"
            rs2("profaddr1") = inp1
            rs2("profaddr2") = inp2
            rs2.Update
            db2.Close
        Else
            inp1 = InputBox("Enter the first address line for the letter.", "Genesis Information Log", "")
            inp2 = InputBox("Enter the second address line for the letter.", "Genesis Information Log", "")
        End If
        Printer.Print Tab(10); inp1
        Printer.Print Tab(10); inp2
        Printer.Print
        Printer.Print
        Printer.Print Tab(10); "RE:"; Tab(20); plaintiff
        Printer.Print Tab(20); "vs."
        Printer.Print Tab(20); defendant
        Printer.Print Tab(20); "OUR FILE # "; casenumber
        Printer.Print
        Printer.Print
        Printer.Print Tab(10); "Dear " + inp + ":"
        Printer.Print
        Printer.Print Tab(10); "Enclosed will be the above referenced Execution Against Property.  This office has levied on the following property"
        Printer.Print Tab(10); "belonging to the defendant:"
        Printer.Print
        linect% = 25
        holdwidth = Printer.Width - 2500
        temp$ = levy.Text
        lasttemp2$ = ""
        While temp$ > ""
            TEMP2$ = temp$
            If holdwidth < Printer.TextWidth(TEMP2$) Then
                TEMP2$ = ""
                While Printer.TextWidth(TEMP2$) <= holdwidth
                    FOUNDSPACE% = 0
                    For t% = Len(TEMP2$) + 1 To Len(temp$)
                        If Mid$(temp$, t%, 1) = " " Then
                            FOUNDSPACE% = t%
                            t% = Len(temp$)
                        End If
                    Next t%
                    lasttemp2$ = TEMP2$
                    If FOUNDSPACE% = 0 Then
                        FOUNDSPACE% = Len(temp$)
                    End If
                    TEMP2$ = TEMP2$ + Mid$(temp$, Len(TEMP2$) + 1, FOUNDSPACE% - Len(TEMP2$))
                Wend
                TEMP2$ = lasttemp2$
            End If
            temp$ = Mid$(temp$, Len(TEMP2$) + 1)
            If Left$(TEMP2$, 1) = " " Then
                TEMP2$ = Mid$(TEMP2$, 2)
            End If
            Printer.Print Tab(10); TEMP2$
            linect% = linect% + 1
        Wend
        Printer.Print
        Printer.Print Tab(10); "If you would like to proceed with a Sheriff's Sale, please forward the $25.00 sale fee, and a legal description will be"
        Printer.Print Tab(10); "forwarded to you.  You will be responsible for placing an ad in the newspaper in the legal section.  The cost of the ad"
        Printer.Print Tab(10); "usually runs from $300 - $500.  We will schedule it for the next available sale date.  If we have not heard from you"
        Printer.Print Tab(10); "within 30 days from the date of this letter, we will assume you do not wish to hold the sale, and at that time our files"
        Printer.Print Tab(10); "will be closed.  Upon settlement, our Sheriff's fees due are " + Format$(commission, "$####0.00") + ", which are based on the amount of the judgement."
        Printer.Print
        Printer.Print Tab(10); "If this office may be of assistance to you in any other matters, please feel free to contact me at ";
        If LPHONE > "" Then
            Printer.Print LPHONE + "."
        Else
            Printer.Print sheriffphone + "."
        End If
        Printer.Print
        Printer.Print
        Printer.Print Tab(10); "Sincerely,"
        Printer.Print
        Printer.Print
        Printer.Print
        Printer.Print
        Printer.Print
        Printer.Print Tab(10); lname
        Printer.Print Tab(10); "Civil Division"
        Printer.Print Tab(10); office
        Printer.Print
        Printer.Print Tab(10); "Enclosure (1)"
        linect% = linect% + 21
        For gg% = linect% + 1 To 62
            Printer.Print
        Next gg%
        Printer.FontBold = True
        Call cp(sheriffaddress + "   •   " + sheriffaddress2)
        Call cp("Telephone: " + sheriffphone)
        Printer.EndDoc
End If
FROMG = 0
Screen.MousePointer = 0
Exit Sub
letterheader:
On Error GoTo oderror9
od9:
Set db = OpenDatabase(nwc + dbname)
Set rs = db.OpenRecordset("select office, sheriffaddress, sheriffaddress2, sheriffphone from system")
rs.MoveFirst
Printer.FontName = "Times New Roman"
Printer.FontSize = 12
Printer.FontBold = True
Printer.Print
Printer.Print
Printer.Print
Printer.Print Tab(10); rs("office"); Tab(90); Date$
Printer.Print Tab(10); rs("sheriffaddress")
Printer.Print Tab(10); rs("sheriffaddress2")
Printer.Print Tab(10); rs("sheriffphone")
Printer.Print
Printer.Print
Printer.Print
Printer.Print
Return

letterbody:
Printer.FontBold = False
Printer.Print
Printer.Print
Printer.Print Tab(10); "RE:"; Tab(30); "Case Number "; casenumber
Printer.Print
Printer.Print Tab(10); "SERVICE OF:"; Tab(30); serviceof
Printer.Print
Printer.Print
Printer.Print
Printer.Print
Printer.Print Tab(10); "The " + rs("office") + " is sending you this letter to inform you of the status of the civil "
Printer.Print
Printer.Print Tab(10); "document(s) filed for service on the person named above.  Several attempts have been made to serve"
Printer.Print
Printer.Print Tab(10); "this person without success."
Printer.Print
Printer.Print
Printer.Print
Printer.Print Tab(10); "If you have any further information that may help service, please contact our civil division.  We will"
Printer.Print
Printer.Print Tab(10); "continue trying to serve this person for you or return the document(s) with affidavit of non service"
Printer.Print
Printer.Print Tab(10); "after 30 days."
Printer.EndDoc
Return
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
oderror3:
If Err > 3200 Then
    Resume od3
Else
    Resume Next
End If
oderror4:
If Err > 3200 Then
    Resume od4
Else
    Resume Next
End If
oderror5:
If Err > 3200 Then
    Resume od5
Else
    Resume Next
End If
oderror6:
If Err > 3200 Then
    Resume od6
Else
    Resume Next
End If
oderror7:
If Err > 3200 Then
    Resume od7
Else
    Resume Next
End If
oderror8:
If Err > 3200 Then
    Resume od8
Else
    Resume Next
End If
oderror9:
If Err > 3200 Then
    Resume od9
Else
    Resume Next
End If


End Sub



Private Sub INDEXBUTTON_Click()
Dim db As Database, ds As Recordset, qs1, qs2, qs3, qs4 As String, holdproc As Date
Screen.MousePointer = 11
If procdate = "12/31/9999" Then
    procdate = DateAdd("yyyy", -1, CVDate(Date$))
    qs1 = "select serviceof,datereceived,iteration,assignedto,papertype from magistrate where datereceived >= #" + Format$(procdate, "mm/dd/yyyy") + "# order by datereceived desc,SERVICEOF,iteration asc"
    qs2 = "select serviceof,datereceived,iteration,assignedto,papertype from familycourt where datereceived >= #" + Format$(procdate, "mm/dd/yyyy") + "# order by datereceived desc,SERVICEOF,iteration asc"
    qs3 = "select serviceof,datereceived,iteration,assignedto,papertype from executions where datereceived >= #" + Format$(procdate, "mm/dd/yyyy") + "# order by datereceived desc,SERVICEOF,iteration asc"
    qs4 = "select serviceof,datereceived,iteration,assignedto,papertype from writother where datereceived >= #" + Format$(procdate, "mm/dd/yyyy") + "# order by datereceived desc,SERVICEOF,iteration asc"
Else
    holdproc = DateAdd("d", -1, procdate)
    procdate = DateAdd("yyyy", -1, procdate)
    qs1 = "select serviceof,datereceived,iteration,assignedto,papertype from magistrate where datereceived between #" + Format$(holdproc, "mm/dd/yyyy") + "# and #" + Format$(procdate, "mm/dd/yyyy") + "# order by datereceived desc,SERVICEOF,iteration asc"
    qs2 = "select serviceof,datereceived,iteration,assignedto,papertype from familycourt where datereceived between #" + Format$(holdproc, "mm/dd/yyyy") + "# and #" + Format$(procdate, "mm/dd/yyyy") + "# order by datereceived desc,SERVICEOF,iteration asc"
    qs3 = "select serviceof,datereceived,iteration,assignedto,papertype from executions where datereceived between #" + Format$(holdproc, "mm/dd/yyyy") + "# and #" + Format$(procdate, "mm/dd/yyyy") + "# order by datereceived desc,SERVICEOF,iteration asc"
    qs4 = "select serviceof,datereceived,iteration,assignedto,papertype from writother where datereceived between #" + Format$(holdproc, "mm/dd/yyyy") + "# and #" + Format$(procdate, "mm/dd/yyyy") + "# order by datereceived desc,SERVICEOF,iteration asc"
End If
On Error GoTo oderror
od:
alllist.ListItems.clear
Set db = OpenDatabase(nwc + dbname)
If maintab.Tab = 0 Then
    Set ds = db.OpenRecordset(qs1)
    ct% = 0
    If Not ds.EOF Then
        ds.MoveFirst
    End If
    While Not ds.EOF
        ds.Edit
        If IsNull(ds("assignedto")) Then
            ds("assignedto") = ""
        End If
        Set itmx = alllist.ListItems.add(, , "Magistrate")
        itmx.SubItems(1) = ds("serviceof")
        itmx.SubItems(2) = Str$(ds("datereceived"))
        itmx.SubItems(3) = ds("iteration")
        itmx.SubItems(4) = ds("assignedto")
        itmx.SubItems(5) = ds("papertype")
        ct% = ct% + 1
        ds.MoveNext
    Wend
End If
If maintab.Tab = 2 Then
    Set ds = db.OpenRecordset(qs2)
    If Not ds.EOF Then
        ds.MoveFirst
    End If
    While Not ds.EOF
        Set itmx = alllist.ListItems.add(, , "Family Court")
        itmx.SubItems(1) = ds("serviceof")
        itmx.SubItems(2) = Str$(ds("datereceived"))
        itmx.SubItems(3) = ds("iteration")
        itmx.SubItems(4) = ds("assignedto")
        itmx.SubItems(5) = ds("papertype")
        ct% = ct% + 1
        ds.MoveNext
    Wend
End If
If maintab.Tab = 3 Then
    Set ds = db.OpenRecordset(qs3)
    If Not ds.EOF Then
        ds.MoveFirst
    End If
    While Not ds.EOF
        Set itmx = alllist.ListItems.add(, , "Executions")
        itmx.SubItems(1) = ds("serviceof")
        itmx.SubItems(2) = Str$(ds("datereceived"))
        itmx.SubItems(3) = ds("iteration")
        itmx.SubItems(4) = ds("assignedto")
        itmx.SubItems(5) = ds("papertype")
        ct% = ct% + 1
        ds.MoveNext
    Wend
End If
If maintab.Tab = 1 Then
    Set ds = db.OpenRecordset(qs4)
    If Not ds.EOF Then
        ds.MoveFirst
    End If
    While Not ds.EOF
        Set itmx = alllist.ListItems.add(, , "Writ/Other")
        itmx.SubItems(1) = ds("serviceof")
        itmx.SubItems(2) = Str$(ds("datereceived"))
        itmx.SubItems(3) = ds("iteration")
        itmx.SubItems(4) = ds("assignedto")
        itmx.SubItems(5) = ds("papertype")
        ct% = ct% + 1
        ds.MoveNext
    Wend
End If
On Error GoTo 0
db.Close
indexframe.Left = 480
indexframe.Top = 360
indexframe.Visible = True
indexframe.Refresh
Screen.MousePointer = 0
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If


End Sub

Private Sub inter_GotFocus()
Call commissandint
End Sub

Private Sub inter_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    remarks.SetFocus
End If

End Sub

Private Sub INTEREST_Change()
If sfpay Then
    total = Str$(Val(commission) + Val(balance) + Val(INTEREST) + Val(servicefee))
Else
    total = Str$(Val(commission) + Val(balance) + Val(INTEREST))
End If
End Sub

Private Sub intrate_GotFocus()
If intrate = "" Then
    intrate = exintrate
End If
End Sub


Private Sub intrate_KeyPress(KeyAscii As Integer)
If maintab.Tab = 3 Then
If KeyAscii = 13 Then
    datesatisfied.SetFocus
End If
End If

End Sub

Private Sub intrate_LostFocus()
Call commissandint
End Sub

Private Sub iteration_Change()
If FROMXREF = "1" Then
    Call iteration_Click
    FROMXREF = "0"
End If
End Sub

Private Sub iteration_Click()
On Error Resume Next
If serviceof = "" Then
    Exit Sub
End If
If datereceived = "" Then
   Exit Sub
End If
If Not IsDate(datereceived) Then
    msg = MsgBox("Filter entry in DATE RECEIVED is not a valid date.", 48, "Genesis Error Log")
    datereceived.SetFocus
    Exit Sub
End If
If Val(iteration) = 0 Then
   msg = MsgBox("Filter entry in ITERATION is not a valid number.", 48, "Genesis Error Log")
   iteration.SetFocus
   Exit Sub
End If
If maintab.Tab > 3 Then
    GoSub tab1
End If
On maintab.Tab + 1 GoSub tab1, tab2, tab3, tab4
Dim db As Database, ds As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nwc + dbname)
Set ds = db.OpenRecordset("select * from " + TP + " where serviceof = " + Chr$(34) + serviceof + Chr$(34) + " and datereceived = #" + datereceived + "# AND ITERATION = " + Chr$(34) + iteration + Chr$(34))
If ds.EOF Then
    db.Close
    Exit Sub
End If
ds.MoveFirst
If maintab.Tab = 2 Then
    If Not IsNull(ds("osce")) Then
        servicefee = ds("osce")
    Else
        servicefee = ""
    End If
    If Not IsNull(ds("IVD")) Then
        ivd.Value = ds("ivd")
    Else
        ivd.Value = 0
    End If
    If Not IsNull(ds("fl1")) Then
        custodian = ds("fl1")
    Else
        custodian = ""
    End If
End If
If maintab.Tab <> 2 Then
    If Not IsNull(ds("receiptnum")) Then
        receiptd = ds("receiptnum")
    Else
        receiptd = ""
    End If
    If Not IsNull(ds("checknum")) Then
        checkd = ds("checknum")
    Else
        checkd = ""
    End If
End If
serviceof = ds("serviceof")
datereceived = Format$(ds("datereceived"), "mm/dd/yyyy")
serviceofsort = ds("serviceofsort")
If Not IsNull(ds("CASENUMBER")) Then
    casenumber = ds("casenumber")
Else
    casenumber = "UNKNOWN"
End If
If Not IsNull(ds("fs2")) Then
    armedforces.Value = Val(ds("fs2"))
Else
    armedforces.Value = 0
End If
If Not IsNull(ds("fs1")) Then
    corporate.Value = Val(ds("fs1"))
Else
    corporate.Value = 0
End If
If Not IsNull(ds("fl1")) Then
    title = ds("fl1")
Else
    title = ""
End If

If Not IsNull(ds("sohomeaddress")) Then
    sohomeaddress = ds("sohomeaddress")
Else
    sohomeaddress = ""
End If
If Not IsNull(ds("sohomeaddress2")) Then
    sohomeaddress2 = ds("sohomeaddress2")
Else
    sohomeaddress2 = ""
End If
If Not IsNull(ds("sohomestate")) Then
    sohomestate = ds("sohomestate")
Else
    sohomestate = ""
End If
If Not IsNull(ds("sohomezipcode")) Then
    sohomezipcode = ds("sohomezipcode")
Else
    sohomezipcode = ""
End If
If Not IsNull(ds("sohomephone")) And ds("sohomephone") <> "" Then
    sohomephone = ds("sohomephone")
Else
    sohomephone = ""
End If
If Not IsNull(ds("soworkaddress")) Then
    soworkaddress = ds("soworkaddress")
Else
    soworkaddress = ""
End If
If Not IsNull(ds("soworkaddress2")) Then
    soworkaddress2 = ds("soworkaddress2")
Else
    soworkaddress2 = ""
End If
If Not IsNull(ds("soworkstate")) Then
    soworkstate = ds("soworkstate")
Else
    soworkstate = ""
End If
If Not IsNull(ds("soworkzipcode")) Then
    soworkzipcode = ds("soworkzipcode")
Else
    soworkzipcode = ""
End If
If Not IsNull(ds("soworkphone")) And ds("soworkphone") <> "" Then
    soworkphone = ds("soworkphone")
Else
    soworkphone = ""
End If
papertype = ds("papertype")
If Not IsNull(ds("courtdate")) Then
    courtdate = Format$(ds("courtdate"), "mm/dd/yyyy")
Else
    courtdate = ""
End If
If Not IsNull(ds("courttime")) Then
    courttime = ds("courttime")
Else
    courttime = ""
End If
If Not IsNull(ds("daystorespond")) Then
    daystorespond = ds("daystorespond")
Else
    daystorespond = ""
End If
If maintab.Tab <> 2 Then
    If Not IsNull(ds("servicefee")) Then
        servicefee = ds("servicefee")
    Else
        servicefee = ""
    End If
    If Not IsNull(ds("bill")) Then
        bill = ds("bill")
    Else
        bill = 0
    End If
    If Not IsNull(ds("FEEDATE")) Then
        feedate = ds("FEEDATE")
    Else
        feedate = ""
    End If
End If
defendant = ds("defendant")
defendantsort = ds("defendantsort")
If Not IsNull(ds("dhomeaddress")) Then
    dhomeaddress = ds("dhomeaddress")
Else
    dhomeaddress = ""
End If
If Not IsNull(ds("dhomeaddress2")) Then
    dhomeaddress2 = ds("dhomeaddress2")
Else
    dhomeaddress2 = ""
End If
If Not IsNull(ds("dhomestate")) Then
    dhomestate = ds("dhomestate")
Else
    dhomestate = ""
End If
If Not IsNull(ds("dhomezipcode")) Then
    dhomezipcode = ds("dhomezipcode")
Else
    dhomezipcode = ""
End If
If Not IsNull(ds("dhomephone")) And ds("dhomephone") <> "" Then
    dhomephone = ds("dhomephone")
Else
    dhomephone = ""
End If
If Not IsNull(ds("dworkaddress")) Then
    dworkaddress = ds("dworkaddress")
Else
    dworkaddress = ""
End If
If Not IsNull(ds("dworkaddress2")) Then
    dworkaddress2 = ds("dworkaddress2")
Else
    dworkaddress2 = ""
End If
If Not IsNull(ds("dworkstate")) Then
    dworkstate = ds("dworkstate")
Else
    dworkstate = ""
End If
If Not IsNull(ds("dworkzipcode")) Then
    dworkzipcode = ds("dworkzipcode")
Else
    dworkzipcode = ""
End If
If Not IsNull(ds("dworkphone")) And ds("dworkphone") <> "" Then
    dworkphone = ds("dworkphone")
Else
    dworkphone = ""
End If
plaintiff = ds("plaintiff")
plaintiffsort = ds("plaintiffsort")
If Not IsNull(ds("phomeaddress")) Then
    phomeaddress = ds("phomeaddress")
Else
    phomeaddress = ""
End If
If Not IsNull(ds("phomeaddress2")) Then
    phomeaddress2 = ds("phomeaddress2")
Else
    phomeaddress2 = ""
End If
If Not IsNull(ds("phomestate")) Then
    phomestate = ds("phomestate")
Else
    phomestate = ""
End If
If Not IsNull(ds("phomezipcode")) Then
    phomezipcode = ds("phomezipcode")
Else
    phomezipcode = ""
End If
If Not IsNull(ds("phomephone")) And ds("phomephone") <> "" Then
    phomephone = ds("phomephone")
Else
    phomephone = ""
End If
If Not IsNull(ds("pworkaddress")) Then
    pworkaddress = ds("pworkaddress")
Else
    pworkaddress = ""
End If
If Not IsNull(ds("pworkaddress2")) Then
    pworkaddress2 = ds("pworkaddress2")
Else
    pworkaddress2 = ""
End If
If Not IsNull(ds("pworkstate")) Then
    pworkstate = ds("pworkstate")
Else
    pworkstate = ""
End If
If Not IsNull(ds("pworkzipcode")) Then
    pworkzipcode = ds("pworkzipcode")
Else
    pworkzipcode = ""
End If
If Not IsNull(ds("pworkphone")) And ds("pworkphone") <> "" Then
    pworkphone = ds("pworkphone")
Else
    pworkphone = ""
End If
If Not IsNull(ds("assignedto")) Then
    assignedto = ds("assignedto")
Else
    assignedto = ""
End If
If Not IsNull(ds("assignedon")) Then
    assignedon = ds("assignedon")
Else
    assignedon = ""
End If
If Not IsNull(ds("served")) Then
    served.Value = Val(ds("served"))
Else
    served.Value = 0
End If
If Not IsNull(ds("nonservice")) Then
    nonservice.Value = Val(ds("nonservice"))
Else
    nonservice.Value = 0
End If
If Not IsNull(ds("nsreason")) Then
    nsreason = ds("nsreason")
Else
    nsreason = ""
End If
If Not IsNull(ds("premarks")) Then
    premarks = ds("premarks")
Else
    premarks = ""
End If
If maintab.Tab = 3 Then
    If Not IsNull(ds("levy")) Then
        levy.Text = ds("levy")
    Else
        levy.Text = ""
    End If
End If
If Not IsNull(ds("wremarks")) Then
    wremarks = ds("wremarks")
Else
    wremarks = ""
End If
If Not IsNull(ds("servicedate")) Then
    servicedate = Format$(ds("servicedate"), "mm/dd/yyyy")
Else
    servicedate = ""
End If
If Not IsNull(ds("servicetime")) Then
    servicetime = ds("servicetime")
Else
    servicetime = ""
End If
If Not IsNull(ds("personserved")) Then
    personserved = ds("personserved")
Else
    personserved = ""
End If
If Not IsNull(ds("locationserved")) Then
    locationserved = ds("locationserved")
Else
    locationserved = ""
End If
If Not IsNull(ds("relationship")) Then
    relationship = ds("relationship")
Else
    relationship = ""
End If
If Not IsNull(ds("professional")) Then
    professional = ds("professional")
Else
    professional = ""
End If
If maintab.Tab = 3 Then
        If Not IsNull(ds("apptdate")) Then
                apptdate = Format$(ds("apptdate"), "mm/dd/yyyy")
        Else
                apptdate = ""
        End If
        If Not IsNull(ds("intrate")) Then
                intrate = ds("intrate")
        Else
                intrate = ""
        End If
        If Not IsNull(ds("datesatisfied")) Then
                datesatisfied = Format$(ds("datesatisfied"), "mm/dd/yyyy")
        Else
                datesatisfied = ""
        End If
        If Not IsNull(ds("judgementdate")) Then
                judgementdate = Format$(ds("judgementdate"), "mm/dd/yyyy")
        Else
                judgementdate = ""
        End If
        If Not IsNull(ds("judgementamount")) Then
                judgementamount = ds("judgementamount")
        Else
                judgementamount = ""
        End If
        'If Not IsNull(ds("estpayoffdate")) Then
        '        estpayoffdate = Format$(ds("estpayoffdate"), "mm/dd/yyyy")
        'Else
        '        estpayoffdate = ""
        'End If
        estpayoffdate = Format$(Date$, "mm/dd/yyyy")
        If Not IsNull(ds("nulla")) Then
            nulla.Value = ds("nulla")
        Else
            nulla.Value = 0
        End If
        commission = ds("COMMISSION")
        INTEREST = ds("INTERest")
        perday = ds("PERDAY")
        If Not IsNull(ds("totalinterest")) Then
            totalinterest = ds("totalinterest")
        Else
            totalinterest = 0
        End If
        If Not IsNull(ds("totalcommission")) Then
            totalcommission = ds("totalcommission")
        Else
            totalcommission = 0
        End If
        If Not IsNull(ds("totalpayments")) Then
            totalpayments = ds("totalpayments")
        Else
            totalpayments = 0
        End If
        
        Call loadpay
End If
lastserviceof = serviceof
CSERVICEOF = 0
infoframe.Refresh
Set db = OpenDatabase(nwl + "LAWSUITE.MDB")
Set rs = db.OpenRecordset("SELECT * FROM PEOPLE WHERE DPNAME = " + Chr$(34) + serviceof + Chr$(34) + " AND NOT MUGSHOT IS NULL")
If Not rs.EOF Then
    rs.MoveFirst
    mugshot.Picture = LoadPicture(rs("MUGSHOT"))
Else
    mugshot.Picture = LoadPicture()
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

tab1:
TP = "magistrate"
Return
tab2:
TP = "writother"
Return
tab3:
TP = "familycourt"
Return
tab4:
TP = "executions"
Return
End Sub

Private Sub iteration_GotFocus()
If maintab.Tab > 3 Then
    GoSub tab1
End If
On maintab.Tab + 1 GoSub tab1, tab2, tab3, tab4
hold$ = iteration
iteration.clear
iteration = hold$
Dim db As Database, ds As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nwc + dbname)
If IsDate(datereceived) Then
   If serviceof > "" Then
         Set ds = db.OpenRecordset("select distinct iteration from " + TP + " where serviceof = " + Chr$(34) + serviceof + Chr$(34) + " and datereceived = #" + datereceived + "# order by iteration")
   Else
         Set ds = db.OpenRecordset("select distinct iteration from " + TP + " order by iteration")
   End If
Else
   If serviceof > "" Then
         Set ds = db.OpenRecordset("select distinct iteration from " + TP + " where serviceof = " + Chr$(34) + serviceof + Chr$(34) + " order by iteration")
   Else
        db.Close
        On Error GoTo 0
         Exit Sub
   End If
End If
If Not ds.EOF Then
    ds.MoveFirst
End If
While Not ds.EOF
    iteration.AddItem ds("iteration")
    ds.MoveNext
Wend
If iteration.ListCount = 0 Then
    iteration = "1"
Else
    If iteration.ListCount >= 9 Then
        msg = MsgBox("Maximum instances of 10 papers served on SERVICE OF party on DATE RECEIVED.  Unable to process without change of date or service of.", 48, "Genesis Error Log")
        iteration = ""
        iteration.SetFocus
        Exit Sub
    Else
        iteration = Mid$(Str$(iteration.List(iteration.ListCount - 1) + 1), 2)
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

tab1:
TP = "magistrate"
Return
tab2:
TP = "writother"
Return
tab3:
TP = "familycourt"
Return
tab4:
TP = "executions"
Return
End Sub


Private Sub iteration_keypress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 9 Then
    If iteration > "" Then
        Call iteration_Click
    End If
End If
If KeyAscii = 13 Then
    serviceofsort.SetFocus
End If

End Sub

Private Sub ivd_Click()
On Error Resume Next
If maintab.Tab <> 2 And receiptd.Visible = True Then
    receiptd.SetFocus
Else
    custodian.SetFocus
End If

End Sub

Private Sub ivd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If maintab.Tab <> 2 And receiptd.Visible = True Then
    receiptd.SetFocus
Else
    custodian.SetFocus
End If
End If

End Sub

Private Sub judgementamount_KeyPress(KeyAscii As Integer)
If maintab.Tab = 3 Then
If KeyAscii = 13 Then
    estpayoffdate.SetFocus
End If
End If

End Sub

Private Sub judgementdate_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
If Len(judgementdate) = 1 Or Len(judgementdate) = 4 Then
    Call sendslash
End If
End If
If maintab.Tab = 3 Then
If KeyAscii = 13 Then
    judgementamount.SetFocus
End If
End If

End Sub


Private Sub levy_KeyPress(KeyAscii As Integer)
If KeyAscii = 34 Then
    msg = MsgBox("The " + Chr$(34) + " character is not allowed.  Use the ' character instead.", 48, "Genesis Error Log")
    KeyAscii = 0
    Exit Sub
End If

End Sub

Private Sub likelist_ItemClick(ByVal Item As MSComctlLib.ListItem)
If likelist.SelectedItem.index = 0 Then
'If likelist.SelectedItem.index = 0 Then
    msg = MsgBox("No selection has been made.", 48, "Genesis Error Log")
    Exit Sub
End If

Set itmx = likelist.ListItems(likelist.SelectedItem.index)
serviceof = itmx
datereceived = itmx.SubItems(1)
iteration = itmx.SubItems(2)
Call iteration_Click
likeframe.Visible = False
End Sub

Private Sub List1_Click()
If List1.ListIndex = -1 Then
    Exit Sub
End If
Open "rp.dat" For Output As #1
Print #1, List1.List(List1.ListIndex)
Close #1
'Dim db As Database, rs As Recordset
'Set db = OpenDatabase(nwc + dbname)
'Set rs = db.OpenRecordset("select regularprinter from system")
'If rs.EOF Then
'    rs.AddNew
'Else
'    rs.MoveFirst
'    rs.Edit
'End If
'rs("regularprinter") = List1.List(List1.ListIndex)
'rs.Update
'db.Close

End Sub

Private Sub List2_Click()
If List2.ListIndex = -1 Then
    Exit Sub
End If
Open "mp.dat" For Output As #1
Print #1, List2.List(List2.ListIndex)
Close #1
'Dim db As Database, rs As Recordset
'Set db = OpenDatabase(nwc + dbname)
'Set rs = db.OpenRecordset("select moneyprinter from system")
'If rs.EOF Then
'    rs.AddNew
'Else
'    rs.MoveFirst
'    rs.Edit
'End If
'rs("moneyprinter") = List2.List(List2.ListIndex)
'rs.Update
'db.Close

End Sub

Private Sub lnf_Click()
Call setfnln
End Sub

Private Sub locationserved_GotFocus()
If personserved > "" Then
        If locationserved = "" Then
            locationserved = sohomeaddress + " " + sohomeaddress2 + " " + sohomestate + " " + sohomezipcode
        End If
End If
End Sub

Private Sub locationserved_KeyPress(KeyAscii As Integer)
If KeyAscii = 34 Then
    msg = MsgBox("The " + Chr$(34) + " character is not allowed.  Use the ' character instead.", 48, "Genesis Error Log")
    KeyAscii = 0
    Exit Sub
End If
If maintab.Tab = 3 Then
If KeyAscii = 13 Then
    apptdate.SetFocus
End If
End If
End Sub

Private Sub maintab_Click(PreviousTab As Integer)
SEARCHTYPE = 0
Dim db As Database, ds As Recordset, rs As Recordset
Select Case maintab.Tab
    Case 0, 1, 2, 3
        If Val(frmLogin.CBROWSE(maintab.Tab)) = 0 Then
            msg = MsgBox("You do not sufficient authority to view the tab selected.", 48, "Genesis Error Log")
            maintab.Tab = PreviousTab
        End If
    Case 5
        If Val(frmLogin.CBROWSE(0)) = 0 Or Val(frmLogin.CBROWSE(1)) = 0 Or Val(frmLogin.CBROWSE(2)) = 0 Or Val(frmLogin.CBROWSE(3)) = 0 Then
            msg = MsgBox("You do not sufficient authority to view the tab selected.", 48, "Genesis Error Log")
            maintab.Tab = PreviousTab
        End If
End Select
On Error Resume Next
Call holdlast(PreviousTab)
Call clearbutton_Click
Call floodlast(maintab.Tab)
If maintab.Tab > 3 Then
    infoframe.Visible = False
    If maintab.Tab = 6 Then
        exintrate.SetFocus
        If omag.Value = True Then
           Call loadprof
        End If
    End If
    If maintab.Tab = 4 Then
        On Error GoTo oderror1
od1:
        'Set db = OpenDatabase(nwc + dbname)
        'Set rs = db.OpenRecordset("select checkprint from system")
        'If Not rs.EOF Then
        '    rs.MoveFirst
        '    If rs("checkprint") = 0 Then
        '        Sfc.Enabled = False
        '        ecc.Enabled = False
        '        epic.Enabled = False
        '        Command7.Enabled = False
        '    Else
                Sfc.Enabled = True
                ecc.Enabled = True
                epic.Enabled = True
                Command7.Enabled = True
        '    End If
        'Else
        '    Sfc.Enabled = False
        '    ecc.Enabled = False
        '    epic.Enabled = False
        '    Command7.Enabled = False
        'End If
        'db.Close
    End If
Else
    serviceof.SetFocus
    On maintab.Tab + 1 GoSub tab1, tab2, tab3, tab4
    infoframe.Top = 450
    infoframe.Left = 50
    infoframe.Visible = True
    serviceof.SetFocus
End If
If maintab.Tab = 2 Then
    receiptl.Visible = False
    receiptd.Visible = False
    checkd.Visible = False
Else
    receiptl.Visible = True
    receiptd.Visible = True
    checkd.Visible = True
End If
If maintab.Tab = 5 Then
    SEARCHTYPE = 1
    Labelo = "OUTSTANDING PAPERS" + Mid$(Labelo, 19)
    Screen.MousePointer = 11
    outstandinglist.Rows = 1
    For t% = 0 To 7
        outstandinglist.Col = t%
        outstandinglist.Text = ""
    Next t%
    outstandinglist.ColWidth(0) = 1000
    outstandinglist.ColWidth(1) = 2400
    outstandinglist.ColWidth(2) = 1100
    outstandinglist.ColWidth(3) = 600
    outstandinglist.ColWidth(4) = 2400
    outstandinglist.ColWidth(5) = 2400
    outstandinglist.ColWidth(6) = 1500
    outstandinglist.ColWidth(7) = 1
    On Error GoTo oderror2
od2:
    Set db = OpenDatabase(nwc + dbname)
    Set ds = db.OpenRecordset("select CASENUMBER, serviceof,datereceived,iteration,assignedto,papertype from magistrate where served <> '1' and nonservice <> '1' order by datereceived,SERVICEOF,iteration")
    ct% = 0
    If Not ds.EOF Then
        ds.MoveFirst
    End If
    While Not ds.EOF
        outstandinglist.AddItem "Magistrate" + Chr$(9) + ds("serviceof") + Chr$(9) + Str$(ds("datereceived")) + Chr$(9) + ds("iteration") + Chr$(9) + ds("assignedto") + Chr$(9) + ds("papertype") + Chr$(9) + ds("casenumber"), ct%
        outstandinglist.RowHeight(ct%) = 400
        ct% = ct% + 1
        ds.MoveNext
    Wend
    Set ds = db.OpenRecordset("select casenumber, serviceof,datereceived,iteration,assignedto,papertype from familycourt where served <> '1' and nonservice <> '1' order by datereceived,SERVICEOF,iteration")
    If Not ds.EOF Then
        ds.MoveFirst
    End If
    While Not ds.EOF
        outstandinglist.AddItem "Family Court" + Chr$(9) + ds("serviceof") + Chr$(9) + Str$(ds("datereceived")) + Chr$(9) + ds("iteration") + Chr$(9) + ds("assignedto") + Chr$(9) + ds("papertype") + Chr$(9) + ds("casenumber"), ct%
        outstandinglist.RowHeight(ct%) = 400
        ct% = ct% + 1
        ds.MoveNext
    Wend
    Set ds = db.OpenRecordset("select casenumber, serviceof,datereceived,iteration,assignedto,papertype from executions where served <> '1' and nonservice <> '1' order by datereceived,SERVICEOF,iteration")
    If Not ds.EOF Then
        ds.MoveFirst
    End If
    While Not ds.EOF
        outstandinglist.AddItem "Executions" + Chr$(9) + ds("serviceof") + Chr$(9) + Str$(ds("datereceived")) + Chr$(9) + ds("iteration") + Chr$(9) + ds("assignedto") + Chr$(9) + ds("papertype") + Chr$(9) + ds("casenumber"), ct%
        outstandinglist.RowHeight(ct%) = 400
        ct% = ct% + 1
        ds.MoveNext
    Wend
    Set ds = db.OpenRecordset("select casenumber,serviceof,datereceived,iteration,assignedto,papertype from writother where served <> '1' and nonservice <> '1' order by datereceived,SERVICEOF,iteration")
    If Not ds.EOF Then
        ds.MoveFirst
    End If
    While Not ds.EOF
        outstandinglist.AddItem "Writ/Other" + Chr$(9) + ds("serviceof") + Chr$(9) + Str$(ds("datereceived")) + Chr$(9) + ds("iteration") + Chr$(9) + ds("assignedto") + Chr$(9) + ds("papertype") + Chr$(9) + ds("casenumber"), ct%
        outstandinglist.RowHeight(ct%) = 400
        ct% = ct% + 1
        ds.MoveNext
    Wend
    outstandinglist.Rows = outstandinglist.Rows - 1
    outstandinglist.Row = 0
    outstandinglist.Col = 0
    Screen.MousePointer = 0
    db.Close
End If
If maintab.Tab = 6 Then
    Call loadprof
End If
If maintab.Tab < 4 Then
    Set db = OpenDatabase(nwl + "LAWSUITE.MDB")
    Set rs = db.OpenRecordset("SELECT * FROM PEOPLE WHERE DPNAME = " + Chr$(34) + serviceof + Chr$(34) + " AND NOT MUGSHOT IS NULL")
    If Not rs.EOF Then
        rs.MoveFirst
        mugshot.Picture = LoadPicture(rs("MUGSHOT"))
    Else
        mugshot.Picture = LoadPicture()
    End If
End If
On Error Resume Next
db.Close
On Error GoTo 0
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
tab1:
Label20.Caption = "Magistrate:"
feel.Caption = "Fee:                  Fee Date:"
apptdate.Visible = False
intrate.Visible = False
datesatisfied.Visible = False
judgementdate.Visible = False
judgementamount.Visible = False
epwb.Visible = False
sat.Visible = False
Partial.Visible = False
nullaex.Visible = False
RL.Visible = False
nrl.Visible = False
dos.Visible = False
levyp.Visible = False
asb.Visible = False
estpayoffdate.Visible = False
nulla.Visible = False
commission.Visible = False
INTEREST.Visible = False
total.Visible = False
perday.Visible = False
balance.Visible = False
'usedlabel.Left = 15
'usedlabel.Top = 5400
'usedlabel.Visible = True
paybutton.Visible = False
ivd.Visible = False
custodian.Visible = False
feedate.Visible = True
bill.Visible = True
Label34.Visible = False
Label35.Visible = False
Label36.Visible = False
Label37.Visible = False
Label38.Visible = False
Label39.Visible = False
Label40.Visible = False
Label67.Visible = False
Label41.Visible = False
Label42.Visible = False
holddep$ = assignedto
Call loaddeputy
assignedto = holddep$
holdprof$ = professional
Call loadprof
professional = holdprof$
Return
tab2:
Label20.Caption = "Attorney:"
ivd.Visible = False
custodian.Visible = False
feedate.Visible = True
bill.Visible = True
feel.Caption = "Fee:                  Fee Date:"
epwb.Visible = False
sat.Visible = False
Partial.Visible = False
nullaex.Visible = False
RL.Visible = False
nrl.Visible = False
dos.Visible = False
levyp.Visible = False
asb.Visible = False
apptdate.Visible = False
intrate.Visible = False
datesatisfied.Visible = False
judgementdate.Visible = False
judgementamount.Visible = False
estpayoffdate.Visible = False
nulla.Visible = False

commission.Visible = False
INTEREST.Visible = False
total.Visible = False
perday.Visible = False
balance.Visible = False
'usedlabel.Left = 15
'usedlabel.Top = 5400
'usedlabel.Visible = True
paybutton.Visible = False
Label34.Visible = False
Label35.Visible = False
Label36.Visible = False
Label37.Visible = False
Label38.Visible = False
Label39.Visible = False
Label40.Visible = False
Label67.Visible = False
Label41.Visible = False
Label42.Visible = False
holddep$ = assignedto
Call loaddeputy
assignedto = holddep$
holdprof$ = professional
Call loadprof
professional = holdprof$
Return
tab3:
Label20.Caption = "Court:"
feel.Caption = "OSCE:"
ivd.Visible = True
custodian.Visible = True
feedate.Visible = False
bill.Visible = False
apptdate.Visible = False
intrate.Visible = False
datesatisfied.Visible = False
judgementdate.Visible = False
judgementamount.Visible = False
estpayoffdate.Visible = False
nulla.Visible = False
epwb.Visible = False
sat.Visible = False
Partial.Visible = False
nullaex.Visible = False
RL.Visible = False
nrl.Visible = False
dos.Visible = False
levyp.Visible = False
asb.Visible = False
commission.Visible = False
INTEREST.Visible = False
total.Visible = False
perday.Visible = False
balance.Visible = False
'usedlabel.Left = 15
'usedlabel.Top = 5400
'usedlabel.Visible = True
paybutton.Visible = False
Label34.Visible = False
Label35.Visible = False
Label36.Visible = False
Label37.Visible = False
Label38.Visible = False
Label39.Visible = False
Label40.Visible = False
Label67.Visible = False
Label41.Visible = False
Label42.Visible = False
holddep$ = assignedto
Call loaddeputy
assignedto = holddep$
holdprof$ = professional
Call loadprof
professional = holdprof$
Return
tab4:
Label20.Caption = "Attorney:"
ivd.Visible = False
custodian.Visible = False

feedate.Visible = True
bill.Visible = True
feel.Caption = "Fee:                  Fee Date:"
apptdate.Visible = True
intrate.Visible = True
datesatisfied.Visible = True
judgementdate.Visible = True
judgementamount.Visible = True
estpayoffdate.Visible = True
nulla.Visible = True
epwb.Visible = True
sat.Visible = True
Partial.Visible = True
nullaex.Visible = True
RL.Visible = True
nrl.Visible = True
dos.Visible = True
levyp.Visible = True
asb.Visible = True
commission.Visible = True
INTEREST.Visible = True
total.Visible = True
perday.Visible = True
balance.Visible = True
'usedlabel.Visible = False
paybutton.Visible = True
Label34.Visible = True
Label35.Visible = True
Label36.Visible = True
Label37.Visible = True
Label38.Visible = True
Label39.Visible = True
Label40.Visible = True
Label67.Visible = True
Label41.Visible = True
Label42.Visible = True
holddep$ = assignedto
Call loaddeputy
assignedto = holddep$
holdprof$ = professional
Call loadprof
professional = holdprof$
Return


End Sub



Private Sub MUGSHOT_Click()
If mugshot.Height = 650 Then
    mframe.Height = 2600
    mframe.Width = 2800
    mugshot.Height = 2600
    mugshot.Width = 2800
Else
    mframe.Height = 650
    mframe.Width = 700
    mugshot.Height = 650
    mugshot.Width = 700
End If
End Sub

Private Sub nonservice_Click()
On Error Resume Next
servicedate.SetFocus
End Sub

Private Sub nonservice_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    servicedate.SetFocus
End If

End Sub

Private Sub nsreason_KeyPress(KeyAscii As Integer)
If KeyAscii = 34 Then
    msg = MsgBox("The " + Chr$(34) + " character is not allowed.  Use the ' character instead.", 48, "Genesis Error Log")
    KeyAscii = 0
    Exit Sub
End If

End Sub

Private Sub oatt_Click()
Call loadprof

End Sub

Private Sub ocou_Click()
Call loadprof

End Sub


Private Sub odep_Click()
Call loadprof

End Sub


Private Sub office_KeyPress(KeyAscii As Integer)
If KeyAscii = 34 Then
    msg = MsgBox("The " + Chr$(34) + " character is not allowed.  Use the ' character instead.", 48, "Genesis Error Log")
    KeyAscii = 0
    Exit Sub
End If

End Sub

Private Sub omag_Click()
Call loadprof
End Sub

Private Sub opsrbdr_Click()
If opsrbdr.Value = True Then
    fromdate.SetFocus
End If
End Sub

Private Sub Option1_Click()

End Sub

Private Sub otheraddress1_KeyPress(KeyAscii As Integer)
If KeyAscii = 34 Then
    msg = MsgBox("The " + Chr$(34) + " character is not allowed.  Use the ' character instead.", 48, "Genesis Error Log")
    KeyAscii = 0
    Exit Sub
End If

End Sub

Private Sub otheraddress2_KeyPress(KeyAscii As Integer)
If KeyAscii = 34 Then
    msg = MsgBox("The " + Chr$(34) + " character is not allowed.  Use the ' character instead.", 48, "Genesis Error Log")
    KeyAscii = 0
    Exit Sub
End If

End Sub

Private Sub othername_Click()
If othername.ListIndex = -1 Then
    Exit Sub
End If
If othername > "" Then
    fromdefendant = 0
    fromplaintiff = 0
End If
Dim db As Database, rs As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nwl + "LAWSUITE.MDB")
Set rs = db.OpenRecordset("select * from professionals where type = 'A' and profname = " + Chr$(34) + othername + Chr$(34))
If Not rs.EOF Then
    rs.MoveFirst
    If Not IsNull(rs("profaddr1")) Then
        otheraddress1 = rs("profaddr1")
    End If
    If Not IsNull(rs("profaddr2")) Then
        otheraddress2 = rs("profaddr2")
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

Private Sub othername_GotFocus()
'If fromdefendant Or fromplaintiff Then
'    Exit Sub
'End If
If othername > "" Then
    Exit Sub
End If
If maintab.Tab = 1 Or maintab.Tab = 3 Then
    othername = professional
    Dim db As Database, rs As Recordset
    On Error GoTo oderror
od:
    Set db = OpenDatabase(nwl + "LAWSUITE.MDB")
    Set rs = db.OpenRecordset("select * from professionals where type = 'A' and profname = " + Chr$(34) + othername + Chr$(34))
    If Not rs.EOF Then
        rs.MoveFirst
        If Not IsNull(rs("profaddr1")) Then
            otheraddress1 = rs("profaddr1")
        End If
        If Not IsNull(rs("profaddr2")) Then
            otheraddress2 = rs("profaddr2")
        End If
    End If
    db.Close
End If
On Error GoTo 0
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If


End Sub

Private Sub othername_KeyPress(KeyAscii As Integer)
If KeyAscii = 34 Then
    msg = MsgBox("The " + Chr$(34) + " character is not allowed.  Use the ' character instead.", 48, "Genesis Error Log")
    KeyAscii = 0
    Exit Sub
End If

End Sub


Private Sub othername_LostFocus()
If othername > "" Then
    fromdefendant = 0
    fromplaintiff = 0
End If

End Sub

Private Sub papertype_GotFocus()
If maintab.Tab = 3 Then
    If papertype = "" Then
        papertype = "EXECUTION"
    End If
End If
End Sub

Private Sub papertype_KeyPress(KeyAscii As Integer)
If KeyAscii = 34 Then
    msg = MsgBox("The " + Chr$(34) + " character is not allowed.  Use the ' character instead.", 48, "Genesis Error Log")
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
    courtdate.SetFocus
End If

End Sub

Private Sub papertype_LostFocus()
If Len(papertype) > 75 Then
    msg = MsgBox("PAPER TYPE can be no more than 75 characters.  The data will be truncated.", 48, "Genesis Error Log")
    papertype = Left$(papertype, 75)
End If

End Sub

Private Sub paybutton_Click()
If serviceof = "" Then
    msg = MsgBox("No SERVICE OF value has been entered.", 48, "Genesis Error Log")
    serviceof.SetFocus
    Exit Sub
End If
If datereceived = "" Then
    msg = MsgBox("DATE RECEIVED is invalid or empty.", 48, "Genesis Error Log")
    datereceived.SetFocus
    Exit Sub
Else
    If Not IsDate(datereceived) Then
        msg = MsgBox("DATE RECEIVED is invalid or empty.", 48, "Genesis Error Log")
        datereceived.SetFocus
        Exit Sub
End If
End If
If Val(iteration) = 0 Then
    msg = MsgBox("ITERATION is invalid or empty.", 48, "Genesis Error Log")
    iteration.SetFocus
    Exit Sub
End If
paymentframe.Left = 50
paymentframe.Top = 150
For t% = 0 To 8
    expaygrid.ColWidth(t%) = 1150
Next t%
paymentframe.Visible = True
expaygrid.Col = 0
expaygrid.Row = 0
Call commissandint
DATEPAID.SetFocus
End Sub

Private Sub perday_GotFocus()
Call commissandint
End Sub


Private Sub personserved_GotFocus()
If personserved = "" And served.Value = 1 Then
    personserved = serviceof
End If
End Sub

Private Sub personserved_KeyPress(KeyAscii As Integer)
If KeyAscii = 34 Then
    msg = MsgBox("The " + Chr$(34) + " character is not allowed.  Use the ' character instead.", 48, "Genesis Error Log")
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
    relationship.SetFocus
End If

End Sub

Private Sub phomeaddress_KeyPress(KeyAscii As Integer)
If KeyAscii = 34 Then
    msg = MsgBox("The " + Chr$(34) + " character is not allowed.  Use the ' character instead.", 48, "Genesis Error Log")
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
    phomeaddress2.SetFocus
End If

End Sub

Private Sub phomeaddress2_KeyPress(KeyAscii As Integer)
If KeyAscii = 34 Then
    msg = MsgBox("The " + Chr$(34) + " character is not allowed.  Use the ' character instead.", 48, "Genesis Error Log")
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
    phomestate.SetFocus
End If

End Sub

Private Sub phomephone_GotFocus()
If Len(phomephone) = 0 Then
Dim db As Database, rs As Recordset
    On Error GoTo oderror
od:
    Set db = OpenDatabase(nwl + "LAWSUITE.MDB")
    Set rs = db.OpenRecordset("SELECT AREACODE FROM DEFAULTS")
    If rs.EOF Then
        db.Close
        Exit Sub
    End If
    rs.MoveFirst
    If IsNull(rs("AREACODE")) Then
        db.Close
        Exit Sub
    End If
    If Len(rs("AREACODE")) <> 3 Then
        db.Close
        Exit Sub
    End If
    Call sendopenpara
    Call SENDCHAR(Left$(rs("AREACODE"), 1))
    Call SENDCHAR(Mid$(rs("AREACODE"), 2, 1))
    Call SENDCHAR(Mid$(rs("AREACODE"), 3, 1))
    Call SENDEND
    db.Close
End If
On Error GoTo 0
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If


End Sub

Private Sub phomephone_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
If Len(phomephone) = 3 Then
    Call sendclosepara
End If
If Len(phomephone) = 4 Then
    Call sendspace
End If
If Len(phomephone) = 8 Then
    Call senddash
End If
If Len(phomephone) = 13 Then
    Call sendspace
End If
End If
If KeyAscii = 13 Then
    pworkaddress.SetFocus
End If

End Sub


Private Sub phomephone_LostFocus()
If Len(phomephone) = 5 Or Len(phomephone) = 6 Then
    phomephone = ""
End If

End Sub

Private Sub phomestate_KeyPress(KeyAscii As Integer)
If KeyAscii = 34 Then
    msg = MsgBox("The " + Chr$(34) + " character is not allowed.  Use the ' character instead.", 48, "Genesis Error Log")
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
    phomezipcode.SetFocus
End If
End Sub

Private Sub phomezipcode_KeyPress(KeyAscii As Integer)
If KeyAscii = 34 Then
    msg = MsgBox("The " + Chr$(34) + " character is not allowed.  Use the ' character instead.", 48, "Genesis Error Log")
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
    phomephone.SetFocus
End If
End Sub

Private Sub pl_Click()
If expaygrid.Row < 0 Then
    msg = MsgBox("A payment row must be selected.", 48, "Genesis Error Log")
    Exit Sub
End If
expaygrid.Col = 1
fee = Val(expaygrid.Text)
If fee = 0 Then
    msg = MsgBox("A valid payment row must be selected.", 48, "Genesis Error Log")
    Exit Sub
End If
expaygrid.Col = 3
cnumber = expaygrid.Text
expaygrid.Col = 2
rnumber = expaygrid.Text
expaygrid.Col = 0
dtp = expaygrid.Text
Dim db, db2 As Database, rs, rs2 As Recordset
If List1.ListCount > 1 And List1.ListIndex > -1 Then
    Call defaultprinter(List1.List(List1.ListIndex))
End If
If List1.ListIndex > -1 And List2.ListIndex > -1 Then
    If List1.List(List1.ListIndex) <> List2.List(List2.ListIndex) And prepareprinter = 1 Then
        msg = MsgBox("Prepare for Worksheet/Affidavit/Letter/Report Printing.", 48, "Genesis Error Log")
    End If
End If
Printer.FontName = "Times New Roman"
Printer.FontSize = 16
Printer.FontBold = True
Printer.Print
Call cp(office)
Printer.FontSize = 12
Call cp(sheriff)
Printer.FontBold = False
Printer.FontSize = 10
Printer.Print
Printer.Print
Printer.Print
Printer.Print Tab(10); Format$(Date$, "mm/dd/yyyy")
Printer.Print
Printer.Print
If professional = "" Then
    inp = InputBox("Enter the name to address the letter.", "Genesis Information Log", "")
Else
    inp = professional
End If
Printer.Print Tab(10); inp
If professional > "" Then
    On Error GoTo oderror
od:
    Set db2 = OpenDatabase(nwl + "lawsuite.mdb")
    Set rs2 = db2.OpenRecordset("select * from professionals where profname = " + Chr$(34) + professional + Chr$(34) + " and type = 'A'")
    If Not rs2.EOF Then
        rs2.MoveFirst
        rs2.Edit
    Else
        rs2.AddNew
    End If
    If IsNull(rs2("profaddr1")) Or rs2("profaddr1") = "" Then
        inp1 = InputBox("Enter the first address line for the Attorney.", "Genesis Information Log", "")
    Else
        inp1 = rs2("profaddr1")
    End If
    If IsNull(rs2("profaddr2")) Or rs2("profaddr2") = "" Then
        inp2 = InputBox("Enter the second address line for the Attorney.", "Genesis Information Log", "")
    Else
        inp2 = rs2("profaddr1")
    End If
    rs2("profname") = professional
    rs2("type") = "A"
    rs2("profaddr1") = inp1
    rs2("profaddr2") = inp2
    rs2.Update
    db2.Close
Else
    inp1 = InputBox("Enter the first address line for the letter.", "Genesis Information Log", "")
    inp2 = InputBox("Enter the second address line for the letter.", "Genesis Information Log", "")
End If
Printer.Print Tab(10); inp1
Printer.Print Tab(10); inp2
Printer.Print
Printer.Print
Printer.Print Tab(10); "RE:"; Tab(20); plaintiff
Printer.Print Tab(20); "vs."
Printer.Print Tab(20); defendant
Printer.Print Tab(20); "OUR FILE # "; casenumber
Printer.Print
Printer.Print
Printer.Print Tab(10); "Dear " + inp + ":"
Printer.Print
Printer.Print Tab(10); "Enclosed is a payment of " + Format$(fee, "$#######0.00") + " on the above referenced Execution Against Property.  As we receive the others,"
Printer.Print Tab(10); "our office will forward them to you."
Printer.Print
Printer.Print Tab(10); "If this office may be of assistance to you in any other matters, please feel free to contact me at ";
If LPHONE > "" Then
    Printer.Print LPHONE + "."
Else
    Printer.Print sheriffphone + "."
End If
Printer.Print
Printer.Print
Printer.Print Tab(10); "Sincerely,"
Printer.Print
Printer.Print
Printer.Print
Printer.Print
Printer.Print
Printer.Print Tab(10); lname
Printer.Print Tab(10); "Civil Division"
Printer.Print Tab(10); office
Printer.Print
Printer.Print Tab(10); "Enclosure"
linect% = 39
For gg% = linect% + 1 To 62
    Printer.Print
Next gg%
Printer.FontBold = True
Call cp(sheriffaddress + "   •   " + sheriffaddress2)
Call cp("Telephone: " + sheriffphone)
Printer.EndDoc
On Error GoTo 0
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If




End Sub

Private Sub plaintiff_Click(AREA As Integer)
infoframe.Refresh
On Error GoTo oderror
Dim db As Database, ds As Recordset
od:
Call setpopup(plaintiff, "F")

Set db = OpenDatabase(nwl + "LAWSUITE.MDB")
Set ds = db.OpenRecordset("select * from PEOPLE where dpname = " + Chr$(34) + plaintiff + Chr$(34))
If Not ds.EOF Then
   ds.MoveFirst
    plaintiffsort = ds("dpsort")
    If Not IsNull(ds("dphaddress")) Then
        phomeaddress = ds("dphaddress")
    Else
        phomeaddress = ""
    End If
    If Not IsNull(ds("dphaddress2")) Then
        phomeaddress2 = ds("dphaddress2")
    Else
        phomeaddress2 = ""
    End If
    If Not IsNull(ds("hstate")) Then
        phomestate = ds("hstate")
    Else
        phomestate = ""
    End If
    If Not IsNull(ds("hzipcode")) Then
        phomezipcode = ds("hzipcode")
    Else
        phomezipcode = ""
    End If
    If Not IsNull(ds("dpwaddress")) Then
        pworkaddress = ds("dpwaddress")
    Else
        pworkaddress = ""
    End If
    If Not IsNull(ds("dpwaddress2")) Then
        pworkaddress2 = ds("dpwaddress2")
    Else
        pworkaddress2 = ""
    End If
    If Not IsNull(ds("wstate")) Then
        pworkstate = ds("wstate")
    Else
        pworkstate = ""
    End If
    If Not IsNull(ds("wzipcode")) Then
        pworkzipcode = ds("wzipcode")
    Else
        pworkzipcode = ""
    End If
    If Not IsNull(ds("dphphone")) And ds("dphphone") <> "" Then
        phomephone = ds("dphphone")
    Else
        phomephone = ""
    End If
    If Not IsNull(ds("dpwphone")) And ds("dpwphone") <> "" Then
        pworkphone = ds("dpwphone")
    Else
        pworkphone = ""
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

Private Sub plaintiff_KeyPress(KeyAscii As Integer)
If KeyAscii = 34 Then
    msg = MsgBox("The " + Chr$(34) + " character is not allowed.  Use the ' character instead.", 48, "Genesis Error Log")
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
    plaintiffsort.SetFocus
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
If Len(plaintiff) > 60 Then
    msg = MsgBox("Maximum length of 60 has been exceeded for PLAINTIFF entry.  This entry will be truncated.", 48, "Genesis Error Log")
    plaintiff = Left$(plaintiff, 60)
End If

End Sub

Private Sub plaintiffsort_GotFocus()
If plaintiff = serviceof Then
    plaintiffsort = serviceofsort
    phomeaddress = sohomeaddress
    phomeaddress2 = sohomeaddress2
    phomestate = sohomestate
    phomezipcode = sohomezipcode
    phomephone = sohomephone
    pworkaddress = soworkaddress
    pworkaddress2 = soworkaddress2
    pworkstate = soworkstate
    pworkzipcode = soworkzipcode
    pworkphone = soworkphone
    Exit Sub
End If
If plaintiffsort > "" Then
    Exit Sub
End If
Dim db As Database, rs As Recordset, ff, LF As Integer, HS As String
ff = 0
LF = 1
HS = ""
On Error GoTo oderror
od:
Set db = OpenDatabase(nwc + dbname)
Set rs = db.OpenRecordset("select fnf,lnf from system")
If Not rs.EOF Then
    rs.MoveFirst
    If rs("fNf") = True Then
        ff = 1
        LF = 0
    End If
End If
db.Close
Call setsort(ff, LF, plaintiff, HS)
plaintiffsort = HS
On Error GoTo 0
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If


End Sub

Private Sub plaintiffsort_KeyPress(KeyAscii As Integer)
If KeyAscii = 34 Then
    msg = MsgBox("The " + Chr$(34) + " character is not allowed.  Use the ' character instead.", 48, "Genesis Error Log")
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
    phomeaddress.SetFocus
End If

End Sub

Private Sub premarks_KeyPress(KeyAscii As Integer)
If KeyAscii = 34 Then
    msg = MsgBox("The " + Chr$(34) + " character is not allowed.  Use the ' character instead.", 48, "Genesis Error Log")
    KeyAscii = 0
    Exit Sub
End If

End Sub

Private Sub prepareprinter_Click()
Dim db As Database, rs As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nwc + dbname)
Set rs = db.OpenRecordset("select prepareprinter from system")
If rs.EOF Then
    rs.AddNew
Else
    rs.MoveFirst
    rs.Edit
End If
rs("prepareprinter") = prepareprinter.Value
rs.Update
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

Private Sub principal_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    commiss.SetFocus
End If

End Sub

Private Sub printbutton_click()
If Val(frmLogin.CPRINT(maintab.Tab)) = 0 And Val(frmLogin.CSUPERVISOR(maintab.Tab)) = 0 Then
    msg = MsgBox("Your USER ID does not have sufficient access to perform this task.", 48, "Genesis Information Log")
    Exit Sub
End If
mprintframe.Left = 8500
mprintframe.Top = 50
mprintframe.Visible = True
goprint.SetFocus
End Sub


Private Sub profaddr1_KeyPress(KeyAscii As Integer)
If KeyAscii = 34 Then
    msg = MsgBox("The " + Chr$(34) + " character is not allowed.  Use the ' character instead.", 48, "Genesis Error Log")
    KeyAscii = 0
    Exit Sub
End If

End Sub

Private Sub profaddr2_KeyPress(KeyAscii As Integer)
If KeyAscii = 34 Then
    msg = MsgBox("The " + Chr$(34) + " character is not allowed.  Use the ' character instead.", 48, "Genesis Error Log")
    KeyAscii = 0
    Exit Sub
End If

End Sub

Private Sub professional_KeyPress(KeyAscii As Integer)
If KeyAscii = 34 Then
    msg = MsgBox("The " + Chr$(34) + " character is not allowed.  Use the ' character instead.", 48, "Genesis Error Log")
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
    defendant.SetFocus
End If

End Sub

Private Sub profname_Click()
If omag.Value = True Then
    TP = "M"
End If
If oatt.Value = True Then
    TP = "A"
End If
If ocou.Value = True Then
    TP = "C"
End If
If odep.Value = True Then
    TP = "D"
End If
Dim db As Database, ds As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nwl + "LAWSUITE.MDB")
Set ds = db.OpenRecordset("select * from professionals where profname = " + Chr$(34) + profname + Chr$(34) + " and type = " + Chr$(34) + TP + Chr$(34))
If ds.EOF Then
   profaddr1 = ""
   profaddr2 = ""
   profphone = ""
Else
    ds.MoveFirst
End If
profname = ds("profname")
If Not IsNull(ds("profaddr1")) Then
    profaddr1 = ds("profaddr1")
Else
    profaddr1 = ""
End If
If Not IsNull(ds("profaddr2")) Then
    profaddr2 = ds("profaddr2")
Else
    profaddr2 = ""
End If
If Not IsNull(ds("profphone")) And ds("profphone") <> "" Then
    profphone = ds("profphone")
Else
    profphone = ""
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

Private Sub profname_KeyPress(KeyAscii As Integer)
If KeyAscii = 34 Then
    msg = MsgBox("The " + Chr$(34) + " character is not allowed.  Use the ' character instead.", 48, "Genesis Error Log")
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 9 Or KeyAscii = 13 Then
    Call profname_Click
End If
End Sub


Private Sub profname_LostFocus()
If Len(profname) > 50 Then
    msg = MsgBox("Length of PROFESSIONAL NAME exceeds the maximum of 50 characters and has been truncated.", 48, "Genesis Error Log")
    profname = Left$(profname, 50)
End If
End Sub


Private Sub propb_Click()
If Not IsDate(DATEPAID) Then
    msg = MsgBox("An invalid date has been entered.", 48, "Genesis Error Log")
    Exit Sub
End If
If CVDate(DATEPAID) < CVDate(judgementdate) Then
    msg = MsgBox("An invalid date has been entered.", 48, "Genesis Error Log")
    Exit Sub
End If
If Val(amount) = 0 Then
    msg = MsgBox("Unable to perform operation on zero amount.", 48, "Genesis Error Log")
    Exit Sub
End If
Call backdate
bprincip = Val(bprincip)
bcommiss = Val(bcommiss)
BINTer = Val(BINTer)
perprin = bprincip / (bprincip + bcommiss + BINTer)
percomm = bcommiss / (bprincip + bcommiss + BINTer)
perint = BINTer / (bprincip + bcommiss + BINTer)
Pprin = bprincip
'For ct% = 1 To expaygrid.Rows
'    expaygrid.Row = ct% - 1
'    expaygrid.Col = 4
'    Pprin = Pprin - Val(expaygrid.Text)
'Next ct%
pcomm = bcommiss
'For ct% = 1 To expaygrid.Rows
'    expaygrid.Row = ct% - 1
'    expaygrid.Col = 5
'    pcomm = pcomm - Val(expaygrid.Text)
'Next ct%
pint = BINTer
'For ct% = 1 To expaygrid.Rows
'    expaygrid.Row = ct% - 1
'    expaygrid.Col = 6
'    pint = pint - Val(expaygrid.Text)
'Next ct%
If sfpay Then
    totall = Pprin + pint + pcomm + Val(servicefee)
Else
    totall = Pprin + pint + pcomm
End If
totamt = Val(amount)
If Val(Format(totall, "#######.##")) < Val(Format(totamt, "#######.##")) Then
    msg = MsgBox("Amount greater than possible total principal, interest, and commission of " + Str$(Pprin + pint + pcomm) + ". Change amount or enter as 2 separate transactions.", 48, "Genesis Error Log")
    Exit Sub
End If
extra = 0
'If Val(Format(totall, "#######.##")) = Val(Format(totamt, "#######.##")) Then
'    AMOUNT = Val(AMOUNT) - Val(servicefee)
'End If
tempprin = perprin * Val(amount)
tempcomm = percomm * Val(amount)
tempint = perint * Val(amount)
If tempprin > Pprin Then
    extrap = tempprin - Pprin
    tempprin = Pprin
    extrac = extrap * (percomm / (percomm + perint))
    extrai = extrap - extrac
    tempcomm = tempcomm + extrac
    If tempcomm > pcomm Then
        extrai = extrai + tempcomm - pcomm
        tempcomm = pcomm
    End If
    tempint = tempint + extrai
    If tempint > pint Then
        extrac = extrac + tempint - pint
        tempint = pint
    End If
Else
    If tempcomm > pcomm Then
        extrac = tempcomm - pcomm
        tempcomm = pcomm
        tempint = tempint + extrac
    Else
        If tempint > pint Then
            extrai = tempint - pint
            tempint = pint
            tempcomm = tempcomm + extrai
        End If
    End If
End If
tempprin = Val(Format$(tempprin, "######0.00"))
tempcomm = Val(Format$(tempcomm, "######0.00"))
tempint = Val(Format$(tempint, "######0.00"))
If CSng(tempprin + tempcomm + tempint) <> Val(amount) Then
    extraa = Val(amount) - tempprin - tempcomm - tempint
    If tempprin + extraa <= Pprin Then
        tempprin = tempprin + extraa
        extraa = 0
    End If
    If tempcomm + extraa <= pcomm Then
        tempcomm = tempcomm + extraa
        extraa = 0
    End If
    If tempint + extraa <= pint Then
        tempint = tempint + extraa
        extraa = 0
    End If
End If
If Val(Format(totall, "#######.##")) = Val(Format(totamt, "#######.##")) Then
    eservicefee = Val(servicefee)
End If
'AMOUNT = AMOUNT + Val(servicefee)
principal = Format$(tempprin, "######0.00")
commiss = Format$(tempcomm, "######0.00")
inter = Format$(tempint, "######0.00")
possprin = Format$(Pprin, "######0.00")
posscomm = Format$(pcomm, "######0.00")
possint = Format$(pint, "######0.00")
If sfpay Then
    POSSTOTAL = Val(possprin) + Val(posscomm) + Val(possint) + Val(servicefee)
Else
    POSSTOTAL = Val(possprin) + Val(posscomm) + Val(possint)
End If
POSSTOTAL = Format$(POSSTOTAL, "######0.00")

End Sub

Private Sub pworkaddress_KeyPress(KeyAscii As Integer)
If KeyAscii = 34 Then
    msg = MsgBox("The " + Chr$(34) + " character is not allowed.  Use the ' character instead.", 48, "Genesis Error Log")
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
    pworkaddress2.SetFocus
End If

End Sub

Private Sub pworkaddress2_KeyPress(KeyAscii As Integer)
If KeyAscii = 34 Then
    msg = MsgBox("The " + Chr$(34) + " character is not allowed.  Use the ' character instead.", 48, "Genesis Error Log")
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
    pworkstate.SetFocus
End If

End Sub

Private Sub pworkphone_GotFocus()
If Len(pworkphone) = 0 Then
    Dim db As Database, rs As Recordset
    On Error GoTo oderror
od:
    Set db = OpenDatabase(nwl + "LAWSUITE.MDB")
    Set rs = db.OpenRecordset("SELECT AREACODE FROM DEFAULTS")
    If rs.EOF Then
        db.Close
        Exit Sub
    End If
    rs.MoveFirst
    If IsNull(rs("AREACODE")) Then
        db.Close
        Exit Sub
    End If
    If Len(rs("AREACODE")) <> 3 Then
        db.Close
        Exit Sub
    End If
    Call sendopenpara
    Call SENDCHAR(Left$(rs("AREACODE"), 1))
    Call SENDCHAR(Mid$(rs("AREACODE"), 2, 1))
    Call SENDCHAR(Mid$(rs("AREACODE"), 3, 1))
    Call SENDEND
    db.Close
End If
On Error GoTo 0
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If


End Sub

Private Sub pworkphone_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
If Len(pworkphone) = 3 Then
    Call sendclosepara
End If
If Len(pworkphone) = 4 Then
    Call sendspace
End If
If Len(pworkphone) = 8 Then
    Call senddash
End If
If Len(pworkphone) = 13 Then
    Call sendspace
End If
End If
If KeyAscii = 13 Then
    assignedto.SetFocus
End If

End Sub


Private Sub pworkphone_LostFocus()
If Len(pworkphone) = 5 Or Len(pworkphone) = 6 Then
    pworkphone = ""
End If

End Sub

Private Sub reason_GotFocus()

End Sub


Private Sub pworkstate_KeyPress(KeyAscii As Integer)
If KeyAscii = 34 Then
    msg = MsgBox("The " + Chr$(34) + " character is not allowed.  Use the ' character instead.", 48, "Genesis Error Log")
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
    pworkzipcode.SetFocus
End If

End Sub

Private Sub pworkzipcode_KeyPress(KeyAscii As Integer)
If KeyAscii = 34 Then
    msg = MsgBox("The " + Chr$(34) + " character is not allowed.  Use the ' character instead.", 48, "Genesis Error Log")
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
    pworkphone.SetFocus
End If

End Sub

Private Sub rclose_Click()
remarksframe.Visible = False
printbutton.SetFocus
'remarksbutton.SetFocus
End Sub

Private Sub RECEIPT_GotFocus()
If receipt = "" And Val(amount) > 0 Then
    Dim db As Database, rs As Recordset
    On Error GoTo oderror
od:
    Set db = OpenDatabase(nwc + dbname)
    Set rs = db.OpenRecordset("select startreceipt from system")
    If rs.EOF Then
        db.Close
        Exit Sub
    Else
        rs.MoveFirst
        If IsNull(rs("startreceipt")) Or rs("startreceipt") = 0 Then
            db.Close
            Exit Sub
        End If
        receipt = rs("startreceipt")
        rs.Edit
        rs("startreceipt") = Val(receipt) + 1
        rs.Update
        db.Close
    End If
End If
On Error GoTo 0
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If



End Sub

Private Sub RECEIPT_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    check.SetFocus
End If

End Sub

Private Sub receiptd_GotFocus()
If receiptd = "" And Val(servicefee) > 0 And maintab.Tab <> 2 Then
    Dim db As Database, rs As Recordset
    On Error GoTo oderror
od:
    Set db = OpenDatabase(nwc + dbname)
    Set rs = db.OpenRecordset("select startreceipt from system")
    If rs.EOF Then
        db.Close
        Exit Sub
    Else
        rs.MoveFirst
        If IsNull(rs("startreceipt")) Or rs("startreceipt") = 0 Then
            db.Close
            Exit Sub
        End If
        receiptd = rs("startreceipt")
        rs.Edit
        rs("startreceipt") = Val(receiptd) + 1
        rs.Update
        db.Close
    End If
End If
On Error GoTo 0
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If


End Sub

Private Sub receiptd_KeyPress(KeyAscii As Integer)
If KeyAscii = 34 Then
    msg = MsgBox("The " + Chr$(34) + " character is not allowed.  Use the ' character instead.", 48, "Genesis Error Log")
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
    If maintab.Tab <> 2 Then
        checkd.SetFocus
    Else
        professional.SetFocus
    End If
End If

End Sub

Private Sub relationship_GotFocus()
If personserved > "" Then
    If personserved <> serviceof Then
        If relationship = "" Then
            relationship = "A PERSON OF AGE AND DISCRETION RESIDING WITH " + serviceof
        End If
    End If
End If
End Sub


Private Sub relationship_KeyPress(KeyAscii As Integer)
If KeyAscii = 34 Then
    msg = MsgBox("The " + Chr$(34) + " character is not allowed.  Use the ' character instead.", 48, "Genesis Error Log")
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
    locationserved.SetFocus
End If

End Sub

Private Sub remarks_KeyPress(KeyAscii As Integer)
If KeyAscii = 34 Then
    msg = MsgBox("The " + Chr$(34) + " character is not allowed.  Use the ' character instead.", 48, "Genesis Error Log")
    KeyAscii = 0
    Exit Sub
End If

End Sub

Private Sub remarksbutton_Click()
remarksframe.Top = 1250
remarksframe.Left = 350
If maintab.Tab = 3 Then
    LEVYL = "Proof of Service Remarks:                                                                                                           Worksheet Remarks:                                                                                        Non-Service Reason                                                                                        Levy Verbiage"
    levy.Visible = True
Else
    LEVYL = "Proof of Service Remarks:                                                                                                           Worksheet Remarks:                                                                                        Non-Service Reason                                                                                        "
    levy.Visible = False
End If
remarksframe.Visible = True
premarks.SetFocus
End Sub

Private Sub removepay_Click()
On Error GoTo blankout
expaygrid.RemoveItem expaygrid.Row
Call commissandint
Exit Sub
blankout:
expaygrid.Row = 0
For uu% = 0 To 8
    expaygrid.Col = uu%
    expaygrid.Text = ""
Next uu%
Call commissandint
Resume Next

End Sub

Private Sub rprintbutton_Click()
Dim edx As Integer, inp As String
If al.Value = True Then
    If Val(frmLogin.CREPORT(0)) = 0 And Val(frmLogin.CREPORT(1)) = 0 And Val(frmLogin.CREPORT(2)) = 0 And Val(frmLogin.CREPORT(3)) = 0 Then
        msg = MsgBox("Your USER ID does not have sufficient access to perform this task.", 48, "Genesis Information Log")
        Exit Sub
    End If
End If
If bsfl.Value = True Or cl.Value = True Or opsrbdr.Value = True Or sdrbdr.Value = True Or aprbdr.Value = True Or sfrrbdr.Value = True Or RLOG.Value = True Or cbr.Value = True Or opbor.Value = True Or opsr.Value = True Then
    If Val(frmLogin.CREPORT(0)) = 0 Or Val(frmLogin.CREPORT(1)) = 0 Or Val(frmLogin.CREPORT(2)) = 0 Or Val(frmLogin.CREPORT(3)) = 0 Then
        msg = MsgBox("Your USER ID does not have sufficient access to perform this task.", 48, "Genesis Information Log")
        Exit Sub
    End If
End If
If wlbdr.Value = True Or walbdr.Value = True Or owbor.Value = True Then
    If Val(frmLogin.CREPORT(1)) = 0 Then
        msg = MsgBox("Your USER ID does not have sufficient access to perform this task.", 48, "Genesis Information Log")
        Exit Sub
    End If
End If
If malbdr.Value = True Or mlbdr.Value = True Or msfrrbdr.Value = True Or ompbor.Value = True Then
    If Val(frmLogin.CREPORT(0)) = 0 Then
        msg = MsgBox("Your USER ID does not have sufficient access to perform this task.", 48, "Genesis Information Log")
        Exit Sub
    End If
End If
If ofcbor.Value = True Or falbdr.Value = True Or fclbdr.Value = True Or fcsopr.Value = True Or ivdsopr.Value = True Then
    If Val(frmLogin.CREPORT(2)) = 0 Then
        msg = MsgBox("Your USER ID does not have sufficient access to perform this task.", 48, "Genesis Information Log")
        Exit Sub
    End If
End If
If ealbdr.Value = True Or elbdr.Value = True Or erlbdr.Value = True Or nullar.Value = True Or aer.Value = True Or oepbor.Value = True Then
    If Val(frmLogin.CREPORT(3)) = 0 Then
        msg = MsgBox("Your USER ID does not have sufficient access to perform this task.", 48, "Genesis Information Log")
        Exit Sub
    End If
End If
If nmsfrrbdr.Value = True Then
    If Val(frmLogin.CREPORT(1)) = 0 Or Val(frmLogin.CREPORT(2)) = 0 Or Val(frmLogin.CREPORT(3)) = 0 Then
        msg = MsgBox("Your USER ID does not have sufficient access to perform this task.", 48, "Genesis Information Log")
        Exit Sub
    End If
End If
If List1.ListCount > 1 And List1.ListIndex > -1 Then
    Call defaultprinter(List1.List(List1.ListIndex))
End If
If List1.ListIndex > -1 And List2.ListIndex > -1 Then
    If List1.List(List1.ListIndex) <> List2.List(List2.ListIndex) And prepareprinter = 1 Then
        msg = MsgBox("Prepare for Worksheet/Affidavit/Letter/Report Printing.", 48, "Genesis Error Log")
    End If
End If

Screen.MousePointer = 11
Dim NOREPORT As Integer
NOREPORT = 0
Dim db As Database, ds As Recordset, rs1, ds2 As Recordset, ds3 As Recordset, ds4 As Recordset
Dim ps, pns, rs, rns, s, ns, pn, rn, n As Integer, cty As String
Dim tps, tpns, trs, trns, ts, tns, tpn, trn, tn As Integer
On Error GoTo oderror
od:
Set db = OpenDatabase(nwc + dbname)
Set ds = db.OpenRecordset("select office from system")
If ds.EOF Then
    msg = MsgBox("Incomplete Sheriff information.", 48, "Genesis Error Log")
    Screen.MousePointer = 0
    db.Close
    Exit Sub
End If
ds.MoveFirst
If Not IsNull(ds("office")) Then
    cty = ds("office")
Else
    msg = MsgBox("Incomplete Sheriff information.", 48, "Genesis Error Log")
    Screen.MousePointer = 0
    db.Close
    Exit Sub
End If
If al.Value = True Then
    report.ReportFileName = nwc + "al.rpt"
    report.Destination = crptToWindow
    report.CopiesToPrinter = 1
    report.SelectionFormula = ""
    report.Action = 1
    Screen.MousePointer = 0
    db.Close
    Exit Sub
End If
If cl.Value = True Then
    If Not IsDate(fromdate) Or Not IsDate(todate) Then
        msg = MsgBox("An invalid date exists in the DATE RANGE CRITERIA frame.", 48, "Genesis Error Log")
        Screen.MousePointer = 0
        db.Close
        Exit Sub
    End If
    On Error GoTo oderror1
od1:
    Set db = OpenDatabase(nwc + dbname)
    Set ds = db.OpenRecordset("select * from passthru")
    If ds.EOF Then
        ds.AddNew
    Else
        ds.MoveFirst
        ds.Edit
    End If
    ds("textstring") = "Date Criteria: " + fromdate + " through " + todate
    ds.Update
    db.Close
    fy$ = Right$(Format$(fromdate, "mmddyyyy"), 4)
    fm$ = Left$(Format$(fromdate, "mmddyyyy"), 2)
    fd$ = Mid$(Format$(fromdate, "mmddyyyy"), 3, 2)
    ty$ = Right$(Format$(todate, "mmddyyyy"), 4)
    tm$ = Left$(Format$(todate, "mmddyyyy"), 2)
    td$ = Mid$(Format$(todate, "mmddyyyy"), 3, 2)
    report.ReportFileName = nwc + "cl.rpt"
    report.SelectionFormula = "{checks.fromdate} >= date(" + fy$ + "," + fm$ + "," + fd$ + ") and {checks.todate} <= date(" + ty$ + "," + tm$ + "," + td$ + ")"
    report.Destination = crptToWindow
    report.CopiesToPrinter = 1
    report.Action = 1
    Screen.MousePointer = 0
    Exit Sub
End If
If bsfl.Value = True Then
    If Not IsDate(fromdate) Or Not IsDate(todate) Then
        msg = MsgBox("An invalid date exists in the DATE RANGE CRITERIA frame.", 48, "Genesis Error Log")
        Screen.MousePointer = 0
        db.Close
        Exit Sub
    End If
    On Error GoTo oderror2
od2:
    Set db = OpenDatabase(nwc + dbname)
    Set ds = db.OpenRecordset("select * from passthru")
    If ds.EOF Then
        ds.AddNew
    Else
        ds.MoveFirst
        While Not ds.EOF
            ds.Delete
            ds.MoveNext
        Wend
        ds.AddNew
    End If
    ds("textstring") = "Date Criteria: " + fromdate + " through " + todate
    ds.Update
    fy$ = Right$(Format$(fromdate, "mmddyyyy"), 4)
    fm$ = Left$(Format$(fromdate, "mmddyyyy"), 2)
    fd$ = Mid$(Format$(fromdate, "mmddyyyy"), 3, 2)
    ty$ = Right$(Format$(todate, "mmddyyyy"), 4)
    tm$ = Left$(Format$(todate, "mmddyyyy"), 2)
    td$ = Mid$(Format$(todate, "mmddyyyy"), 3, 2)
    db.Close
    report.ReportFileName = nwc + "bille.rpt"
    report.SelectionFormula = "{executions.bill} = 1"
    report.Destination = crptToWindow
    report.CopiesToPrinter = 1
    report.Action = 1
    report.ReportFileName = nwc + "billw.rpt"
    report.SelectionFormula = "{writother.bill} = 1"
    report.Destination = crptToWindow
    report.CopiesToPrinter = 1
    report.Action = 1
    report.ReportFileName = nwc + "billm.rpt"
    report.SelectionFormula = "{magistrate.bill} = 1"
    report.Destination = crptToWindow
    report.CopiesToPrinter = 1
    report.Action = 1
    Screen.MousePointer = 0
    Exit Sub
End If
If opsrbdr.Value = True Then
    If Not IsDate(fromdate) Or Not IsDate(todate) Then
        msg = MsgBox("An invalid date exists in the DATE RANGE CRITERIA frame.", 48, "Genesis Error Log")
        Screen.MousePointer = 0
        db.Close
        Exit Sub
    End If
    Call printroutines("ofsrbdrrtn", "")
    If NOREPORT = 1 Then
        msg = MsgBox("No data exists for report criteria.", 48, "Genesis Information Log")
        Screen.MousePointer = 0
        db.Close
        Exit Sub
    End If
    On Error GoTo oderror3
od3:
    Set db = OpenDatabase(nwc + dbname)
    Set ds = db.OpenRecordset("select * from passthru")
    If ds.EOF Then
        ds.AddNew
    Else
        ds.MoveFirst
        While Not ds.EOF
            ds.Delete
            ds.MoveNext
        Wend
        ds.AddNew
    End If
    ds("textstring") = "Date Criteria: " + fromdate + " through " + todate
    ds.Update
    fy$ = Right$(Format$(fromdate, "mmddyyyy"), 4)
    fm$ = Left$(Format$(fromdate, "mmddyyyy"), 2)
    fd$ = Mid$(Format$(fromdate, "mmddyyyy"), 3, 2)
    ty$ = Right$(Format$(todate, "mmddyyyy"), 4)
    tm$ = Left$(Format$(todate, "mmddyyyy"), 2)
    td$ = Mid$(Format$(todate, "mmddyyyy"), 3, 2)
    db.Close
    report.ReportFileName = nwc + "ofsrbdr.rpt"
    report.SelectionFormula = ""
    report.Destination = crptToWindow
    report.CopiesToPrinter = 1
    report.Action = 1
    Screen.MousePointer = 0
    Exit Sub
End If
If sdrbdr.Value = True Then
    inp = InputBox("Choose report option:  S = Service Detail Report   N = Non-Service Detail Report", "Genesis Information Log", "S")
    inp = UCase(inp)
    If inp <> "S" And inp <> "N" Then
        msg = MsgBox("Invalid option selected.", 48, "Genesis Error Log")
        Exit Sub
    End If
    If Not IsDate(fromdate) Or Not IsDate(todate) Then
        msg = MsgBox("An invalid date exists in the DATE RANGE CRITERIA frame.", 48, "Genesis Error Log")
        Screen.MousePointer = 0
        db.Close
        Exit Sub
    End If
    Call printroutines("sdrbdrrtn", inp)
    If NOREPORT = 1 Then
        msg = MsgBox("No data exists for report criteria.", 48, "Genesis Information Log")
        Screen.MousePointer = 0
        db.Close
        Exit Sub
    End If
    On Error GoTo oderror4
od4:
    Set db = OpenDatabase(nwc + dbname)
    Set ds = db.OpenRecordset("select * from passthru")
    If ds.EOF Then
        ds.AddNew
    Else
        ds.MoveFirst
        ds.Edit
    End If
    ds("textstring") = "Date Criteria: " + fromdate + " through " + todate
    ds.Update
    fy$ = Right$(Format$(fromdate, "mmddyyyy"), 4)
    fm$ = Left$(Format$(fromdate, "mmddyyyy"), 2)
    fd$ = Mid$(Format$(fromdate, "mmddyyyy"), 3, 2)
    ty$ = Right$(Format$(todate, "mmddyyyy"), 4)
    tm$ = Left$(Format$(todate, "mmddyyyy"), 2)
    td$ = Mid$(Format$(todate, "mmddyyyy"), 3, 2)
    db.Close
    If inp = "S" Then
        report.ReportFileName = nwc + "sdrbdr.rpt"
    Else
        report.ReportFileName = nwc + "nsdrbdr.rpt"
    End If
    report.Destination = crptToWindow
    report.CopiesToPrinter = 1
    report.SelectionFormula = ""
    report.Action = 1
    Screen.MousePointer = 0
    Exit Sub
End If
If wlbdr.Value = True Then
    If Not IsDate(fromdate) Or Not IsDate(todate) Then
        msg = MsgBox("An invalid date exists in the DATE RANGE CRITERIA frame.", 48, "Genesis Error Log")
        Screen.MousePointer = 0
        db.Close
        Exit Sub
    End If
    If Not IsDate(fromdate) Or Not IsDate(todate) Then
        msg = MsgBox("An invalid date exists in the DATE RANGE CRITERIA frame.", 48, "Genesis Error Log")
        Screen.MousePointer = 0
        db.Close
        Exit Sub
    End If
    inp = InputBox("Enter S for Standard Report or C for Court Date Report.", "Genesis Information Log", "S")
    inp = UCase(inp)
    If inp <> "S" And inp <> "C" Then
        msg = MsgBox("Invalid entry.", 48, "Genesis Error Log")
        Screen.MousePointer = 0
        db.Close
        Exit Sub
    End If
    If inp = "C" Then
        On Error GoTo oderror5
od5:
        Set db = OpenDatabase(nwc + dbname)
        Set ds = db.OpenRecordset("select * from passthru")
        If ds.EOF Then
            ds.AddNew
        Else
            ds.MoveFirst
            ds.Edit
        End If
        ds("textstring") = "Date Criteria: " + fromdate + " through " + todate
        ds.Update
        db.Close
        fy$ = Right$(Format$(fromdate, "mmddyyyy"), 4)
        fm$ = Left$(Format$(fromdate, "mmddyyyy"), 2)
        fd$ = Mid$(Format$(fromdate, "mmddyyyy"), 3, 2)
        ty$ = Right$(Format$(todate, "mmddyyyy"), 4)
        tm$ = Left$(Format$(todate, "mmddyyyy"), 2)
        td$ = Mid$(Format$(todate, "mmddyyyy"), 3, 2)
        report.ReportFileName = nwc + "wc.rpt"
        report.Destination = crptToWindow
        report.CopiesToPrinter = 1
        report.SelectionFormula = "{writother.courtdate} >= DATE(" + fy$ + "," + fm$ + "," + fd$ + ") and {writother.courtdate} <= DATE(" + ty$ + "," + tm$ + "," + td$ + ")"
        report.Action = 1
        Screen.MousePointer = 0
        Exit Sub
    Else
        Call printroutines("wlbdrrtn", inp)
        If NOREPORT = 1 Then
            msg = MsgBox("No data exists for report criteria.", 48, "Genesis Information Log")
            Screen.MousePointer = 0
            db.Close
            Exit Sub
        End If
        On Error GoTo oderror6
od6:
        Set db = OpenDatabase(nwc + dbname)
        Set ds = db.OpenRecordset("select * from passthru")
        If ds.EOF Then
            ds.AddNew
        Else
            ds.MoveFirst
            ds.Edit
        End If
        ds("textstring") = "Date Criteria: " + fromdate + " through " + todate
        ds.Update
        fy$ = Right$(Format$(fromdate, "mmddyyyy"), 4)
        fm$ = Left$(Format$(fromdate, "mmddyyyy"), 2)
        fd$ = Mid$(Format$(fromdate, "mmddyyyy"), 3, 2)
        ty$ = Right$(Format$(todate, "mmddyyyy"), 4)
        tm$ = Left$(Format$(todate, "mmddyyyy"), 2)
        td$ = Mid$(Format$(todate, "mmddyyyy"), 3, 2)
        db.Close
        report.ReportFileName = nwc + "wlbdr.rpt"
        report.Destination = crptToWindow
        report.CopiesToPrinter = 1
        report.SelectionFormula = ""
        report.Action = 1
        Screen.MousePointer = 0
        Exit Sub
    End If
End If
If aprbdr.Value = True Then
    If Not IsDate(fromdate) Or Not IsDate(todate) Then
        msg = MsgBox("An invalid date exists in the DATE RANGE CRITERIA frame.", 48, "Genesis Error Log")
        Screen.MousePointer = 0
        db.Close
        Exit Sub
    End If
    Call printroutines("aprbdrrtn", "")
    If NOREPORT = 1 Then
        msg = MsgBox("No data exists for report criteria.", 48, "Genesis Information Log")
        Screen.MousePointer = 0
        db.Close
        Exit Sub
    End If
    On Error GoTo oderror7
od7:
    Set db = OpenDatabase(nwc + dbname)
    Set ds = db.OpenRecordset("select * from passthru")
    If ds.EOF Then
        ds.AddNew
    Else
        ds.MoveFirst
        ds.Edit
    End If
    ds("textstring") = "Date Criteria: " + fromdate + " through " + todate
    ds.Update
    fy$ = Right$(Format$(fromdate, "mmddyyyy"), 4)
    fm$ = Left$(Format$(fromdate, "mmddyyyy"), 2)
    fd$ = Mid$(Format$(fromdate, "mmddyyyy"), 3, 2)
    ty$ = Right$(Format$(todate, "mmddyyyy"), 4)
    tm$ = Left$(Format$(todate, "mmddyyyy"), 2)
    td$ = Mid$(Format$(todate, "mmddyyyy"), 3, 2)
    db.Close
    report.ReportFileName = nwc + "albdr.rpt"
    report.Destination = crptToWindow
    report.CopiesToPrinter = 1
    report.SelectionFormula = ""
    report.Action = 1
    Screen.MousePointer = 0
    Exit Sub
End If
If walbdr.Value = True Then
    If Not IsDate(fromdate) Or Not IsDate(todate) Then
        msg = MsgBox("An invalid date exists in the DATE RANGE CRITERIA frame.", 48, "Genesis Error Log")
        Screen.MousePointer = 0
        db.Close
        Exit Sub
    End If
    Call printroutines("walbdrrtn", "")
    If NOREPORT = 1 Then
        msg = MsgBox("No data exists for report criteria.", 48, "Genesis Information Log")
        Screen.MousePointer = 0
        db.Close
        Exit Sub
    End If
    On Error GoTo oderror8
od8:
    Set db = OpenDatabase(nwc + dbname)
    Set ds = db.OpenRecordset("select * from passthru")
    If ds.EOF Then
        ds.AddNew
    Else
        ds.MoveFirst
        ds.Edit
    End If
    ds("textstring") = "Date Criteria: " + fromdate + " through " + todate
    ds.Update
    db.Close
    fy$ = Right$(Format$(fromdate, "mmddyyyy"), 4)
    fm$ = Left$(Format$(fromdate, "mmddyyyy"), 2)
    fd$ = Mid$(Format$(fromdate, "mmddyyyy"), 3, 2)
    ty$ = Right$(Format$(todate, "mmddyyyy"), 4)
    tm$ = Left$(Format$(todate, "mmddyyyy"), 2)
    td$ = Mid$(Format$(todate, "mmddyyyy"), 3, 2)
    report.ReportFileName = nwc + "walbdr.rpt"
    report.Destination = crptToWindow
    report.CopiesToPrinter = 1
    report.SelectionFormula = ""
    report.Action = 1
    Screen.MousePointer = 0
    Exit Sub
End If
If malbdr.Value = True Then
    If Not IsDate(fromdate) Or Not IsDate(todate) Then
        msg = MsgBox("An invalid date exists in the DATE RANGE CRITERIA frame.", 48, "Genesis Error Log")
        Screen.MousePointer = 0
        db.Close
        Exit Sub
    End If
    Call printroutines("malbdrrtn", "")
    If NOREPORT = 1 Then
        msg = MsgBox("No data exists for report criteria.", 48, "Genesis Information Log")
        Screen.MousePointer = 0
        db.Close
        Exit Sub
    End If
    On Error GoTo oderror9
od9:
    Set db = OpenDatabase(nwc + dbname)
    Set ds = db.OpenRecordset("select * from passthru")
    If ds.EOF Then
        ds.AddNew
    Else
        ds.MoveFirst
        ds.Edit
    End If
    ds("textstring") = "Date Criteria: " + fromdate + " through " + todate
    ds.Update
    db.Close
    fy$ = Right$(Format$(fromdate, "mmddyyyy"), 4)
    fm$ = Left$(Format$(fromdate, "mmddyyyy"), 2)
    fd$ = Mid$(Format$(fromdate, "mmddyyyy"), 3, 2)
    ty$ = Right$(Format$(todate, "mmddyyyy"), 4)
    tm$ = Left$(Format$(todate, "mmddyyyy"), 2)
    td$ = Mid$(Format$(todate, "mmddyyyy"), 3, 2)
    report.ReportFileName = nwc + "malbdr.rpt"
    report.Destination = crptToWindow
    report.CopiesToPrinter = 1
    report.SelectionFormula = ""
    report.Action = 1
    Screen.MousePointer = 0
    Exit Sub
End If
If falbdr.Value = True Then
    If Not IsDate(fromdate) Or Not IsDate(todate) Then
        msg = MsgBox("An invalid date exists in the DATE RANGE CRITERIA frame.", 48, "Genesis Error Log")
        Screen.MousePointer = 0
        db.Close
        Exit Sub
    End If
    Call printroutines("falbdrrtn", "")
    If NOREPORT = 1 Then
        msg = MsgBox("No data exists for report criteria.", 48, "Genesis Information Log")
        Screen.MousePointer = 0
        db.Close
        Exit Sub
    End If
    On Error GoTo oderror10
od10:
    Set db = OpenDatabase(nwc + dbname)
    Set ds = db.OpenRecordset("select * from passthru")
    If ds.EOF Then
        ds.AddNew
    Else
        ds.MoveFirst
        ds.Edit
    End If
    ds("textstring") = "Date Criteria: " + fromdate + " through " + todate
    ds.Update
    db.Close
    fy$ = Right$(Format$(fromdate, "mmddyyyy"), 4)
    fm$ = Left$(Format$(fromdate, "mmddyyyy"), 2)
    fd$ = Mid$(Format$(fromdate, "mmddyyyy"), 3, 2)
    ty$ = Right$(Format$(todate, "mmddyyyy"), 4)
    tm$ = Left$(Format$(todate, "mmddyyyy"), 2)
    td$ = Mid$(Format$(todate, "mmddyyyy"), 3, 2)
    report.ReportFileName = nwc + "falbdr.rpt"
    report.Destination = crptToWindow
    report.CopiesToPrinter = 1
    report.SelectionFormula = ""
    report.Action = 1
    Screen.MousePointer = 0
    Exit Sub
End If
If ealbdr.Value = True Then
    If Not IsDate(fromdate) Or Not IsDate(todate) Then
        msg = MsgBox("An invalid date exists in the DATE RANGE CRITERIA frame.", 48, "Genesis Error Log")
        Screen.MousePointer = 0
        db.Close
        Exit Sub
    End If
    Call printroutines("ealbdrrtn", "")
    If NOREPORT = 1 Then
        msg = MsgBox("No data exists for report criteria.", 48, "Genesis Information Log")
        Screen.MousePointer = 0
        db.Close
        Exit Sub
    End If
    On Error GoTo oderror11
od11:
    Set db = OpenDatabase(nwc + dbname)
    Set ds = db.OpenRecordset("select * from passthru")
    If ds.EOF Then
        ds.AddNew
    Else
        ds.MoveFirst
        ds.Edit
    End If
    ds("textstring") = "Date Criteria: " + fromdate + " through " + todate
    ds.Update
    db.Close
    fy$ = Right$(Format$(fromdate, "mmddyyyy"), 4)
    fm$ = Left$(Format$(fromdate, "mmddyyyy"), 2)
    fd$ = Mid$(Format$(fromdate, "mmddyyyy"), 3, 2)
    ty$ = Right$(Format$(todate, "mmddyyyy"), 4)
    tm$ = Left$(Format$(todate, "mmddyyyy"), 2)
    td$ = Mid$(Format$(todate, "mmddyyyy"), 3, 2)
    report.ReportFileName = nwc + "ealbdr.rpt"
    report.Destination = crptToWindow
    report.CopiesToPrinter = 1
    report.SelectionFormula = ""
    report.Action = 1
    Screen.MousePointer = 0
    Exit Sub
End If
If fclbdr.Value = True Then
    If Not IsDate(fromdate) Or Not IsDate(todate) Then
        msg = MsgBox("An invalid date exists in the DATE RANGE CRITERIA frame.", 48, "Genesis Error Log")
        Screen.MousePointer = 0
        db.Close
        Exit Sub
    End If
    inp = InputBox("Enter S for Standard Report or C for Court Date Report.", "Genesis Information Log", "S")
    inp = UCase(inp)
    If inp <> "S" And inp <> "C" Then
        msg = MsgBox("Invalid entry.", 48, "Genesis Error Log")
        Screen.MousePointer = 0
        db.Close
        Exit Sub
    End If
    If inp = "C" Then
        On Error GoTo oderror12
od12:
        Set db = OpenDatabase(nwc + dbname)
        Set ds = db.OpenRecordset("select * from passthru")
        If ds.EOF Then
            ds.AddNew
        Else
            ds.MoveFirst
            ds.Edit
        End If
        ds("textstring") = "Date Criteria: " + fromdate + " through " + todate
        ds.Update
        db.Close
        fy$ = Right$(Format$(fromdate, "mmddyyyy"), 4)
        fm$ = Left$(Format$(fromdate, "mmddyyyy"), 2)
        fd$ = Mid$(Format$(fromdate, "mmddyyyy"), 3, 2)
        ty$ = Right$(Format$(todate, "mmddyyyy"), 4)
        tm$ = Left$(Format$(todate, "mmddyyyy"), 2)
        td$ = Mid$(Format$(todate, "mmddyyyy"), 3, 2)
        report.ReportFileName = nwc + "fcc.rpt"
        report.Destination = crptToWindow
        report.CopiesToPrinter = 1
        report.SelectionFormula = "{familycourt.courtdate} >= DATE(" + fy$ + "," + fm$ + "," + fd$ + ") and {familycourt.courtdate} <= DATE(" + ty$ + "," + tm$ + "," + td$ + ")"
        report.Action = 1
        Screen.MousePointer = 0
        Exit Sub
    Else
        Call printroutines("fclbdrrtn", inp)
        If NOREPORT = 1 Then
            msg = MsgBox("No data exists for report criteria.", 48, "Genesis Information Log")
            Screen.MousePointer = 0
            db.Close
            Exit Sub
        End If
        On Error GoTo oderror13
OD13:
        Set db = OpenDatabase(nwc + dbname)
        Set ds = db.OpenRecordset("select * from passthru")
        If ds.EOF Then
            ds.AddNew
        Else
            ds.MoveFirst
            ds.Edit
        End If
        ds("textstring") = "Date Criteria: " + fromdate + " through " + todate
        ds.Update
        db.Close
        fy$ = Right$(Format$(fromdate, "mmddyyyy"), 4)
        fm$ = Left$(Format$(fromdate, "mmddyyyy"), 2)
        fd$ = Mid$(Format$(fromdate, "mmddyyyy"), 3, 2)
        ty$ = Right$(Format$(todate, "mmddyyyy"), 4)
        tm$ = Left$(Format$(todate, "mmddyyyy"), 2)
        td$ = Mid$(Format$(todate, "mmddyyyy"), 3, 2)
        report.ReportFileName = nwc + "fclbdr.rpt"
        report.Destination = crptToWindow
        report.CopiesToPrinter = 1
        report.SelectionFormula = ""
        report.Action = 1
        Screen.MousePointer = 0
        Exit Sub
    End If
End If
If mlbdr.Value = True Then
    If Not IsDate(fromdate) Or Not IsDate(todate) Then
        msg = MsgBox("An invalid date exists in the DATE RANGE CRITERIA frame.", 48, "Genesis Error Log")
        Screen.MousePointer = 0
        db.Close
        Exit Sub
    End If
    If Not IsDate(fromdate) Or Not IsDate(todate) Then
        msg = MsgBox("An invalid date exists in the DATE RANGE CRITERIA frame.", 48, "Genesis Error Log")
        Screen.MousePointer = 0
        db.Close
        Exit Sub
    End If
    inp = InputBox("Enter S for Standard Report or C for Court Date Report.", "Genesis Information Log", "S")
    inp = UCase(inp)
    If inp <> "S" And inp <> "C" Then
        msg = MsgBox("Invalid entry.", 48, "Genesis Error Log")
        Screen.MousePointer = 0
        db.Close
        Exit Sub
    End If
    If inp = "C" Then
        On Error GoTo oderror14
OD14:
        Set db = OpenDatabase(nwc + dbname)
        Set ds = db.OpenRecordset("select * from passthru")
        If ds.EOF Then
            ds.AddNew
        Else
            ds.MoveFirst
            ds.Edit
        End If
        ds("textstring") = "Date Criteria: " + fromdate + " through " + todate
        ds.Update
        db.Close
        fy$ = Right$(Format$(fromdate, "mmddyyyy"), 4)
        fm$ = Left$(Format$(fromdate, "mmddyyyy"), 2)
        fd$ = Mid$(Format$(fromdate, "mmddyyyy"), 3, 2)
        ty$ = Right$(Format$(todate, "mmddyyyy"), 4)
        tm$ = Left$(Format$(todate, "mmddyyyy"), 2)
        td$ = Mid$(Format$(todate, "mmddyyyy"), 3, 2)
        report.ReportFileName = nwc + "mc.rpt"
        report.Destination = crptToWindow
        report.CopiesToPrinter = 1
        report.SelectionFormula = "{magistrate.courtdate} >= DATE(" + fy$ + "," + fm$ + "," + fd$ + ") and {magistrate.courtdate} <= DATE(" + ty$ + "," + tm$ + "," + td$ + ")"
        report.Action = 1
        Screen.MousePointer = 0
        Exit Sub
    Else
        Call printroutines("mlbdrrtn", inp)
        If NOREPORT = 1 Then
            msg = MsgBox("No data exists for report criteria.", 48, "Genesis Information Log")
            Screen.MousePointer = 0
            db.Close
            Exit Sub
        End If
        On Error GoTo oderror15
OD15:
        Set db = OpenDatabase(nwc + dbname)
        Set ds = db.OpenRecordset("select * from passthru")
        If ds.EOF Then
            ds.AddNew
        Else
            ds.MoveFirst
            ds.Edit
        End If
        ds("textstring") = "Date Criteria: " + fromdate + " through " + todate
        ds.Update
        db.Close
        fy$ = Right$(Format$(fromdate, "mmddyyyy"), 4)
        fm$ = Left$(Format$(fromdate, "mmddyyyy"), 2)
        fd$ = Mid$(Format$(fromdate, "mmddyyyy"), 3, 2)
        ty$ = Right$(Format$(todate, "mmddyyyy"), 4)
        tm$ = Left$(Format$(todate, "mmddyyyy"), 2)
        td$ = Mid$(Format$(todate, "mmddyyyy"), 3, 2)
        report.ReportFileName = nwc + "Mlbdr.rpt"
        report.Destination = crptToWindow
        report.CopiesToPrinter = 1
        report.SelectionFormula = ""
        report.Action = 1
        Screen.MousePointer = 0
        Exit Sub
    End If
End If
If elbdr.Value = True Then
    If Not IsDate(fromdate) Or Not IsDate(todate) Then
        msg = MsgBox("An invalid date exists in the DATE RANGE CRITERIA frame.", 48, "Genesis Error Log")
        Screen.MousePointer = 0
        db.Close
        Exit Sub
    End If
        Call printroutines("elbdrrtn", "")
        If NOREPORT = 1 Then
            msg = MsgBox("No data exists for report criteria.", 48, "Genesis Information Log")
            Screen.MousePointer = 0
            db.Close
            Exit Sub
        End If
        On Error GoTo oderror16
OD16:
        Set db = OpenDatabase(nwc + dbname)
        Set ds = db.OpenRecordset("select * from passthru")
        If ds.EOF Then
            ds.AddNew
        Else
            ds.MoveFirst
            ds.Edit
        End If
        ds("textstring") = "Date Criteria: " + fromdate + " through " + todate
        ds.Update
        db.Close
        fy$ = Right$(Format$(fromdate, "mmddyyyy"), 4)
        fm$ = Left$(Format$(fromdate, "mmddyyyy"), 2)
        fd$ = Mid$(Format$(fromdate, "mmddyyyy"), 3, 2)
        ty$ = Right$(Format$(todate, "mmddyyyy"), 4)
        tm$ = Left$(Format$(todate, "mmddyyyy"), 2)
        td$ = Mid$(Format$(todate, "mmddyyyy"), 3, 2)
        report.ReportFileName = nwc + "elbdr.rpt"
        report.Destination = crptToWindow
        report.CopiesToPrinter = 1
        report.SelectionFormula = ""
        report.Action = 1
        Screen.MousePointer = 0
        Exit Sub
End If
If erlbdr.Value = True Then
    If Not IsDate(fromdate) Or Not IsDate(todate) Then
        msg = MsgBox("An invalid date exists in the DATE RANGE CRITERIA frame.", 48, "Genesis Error Log")
        Screen.MousePointer = 0
        db.Close
        Exit Sub
    End If
    On Error GoTo oderror17
OD17:
    Set db = OpenDatabase(nwc + dbname)
    Set ds = db.OpenRecordset("select * from passthru")
    If ds.EOF Then
        ds.AddNew
    Else
        ds.MoveFirst
        ds.Edit
    End If
    ds("textstring") = "Date Criteria: " + fromdate + " through " + todate
    ds.Update
    db.Close
    fy$ = Right$(Format$(fromdate, "mmddyyyy"), 4)
    fm$ = Left$(Format$(fromdate, "mmddyyyy"), 2)
    fd$ = Mid$(Format$(fromdate, "mmddyyyy"), 3, 2)
    ty$ = Right$(Format$(todate, "mmddyyyy"), 4)
    tm$ = Left$(Format$(todate, "mmddyyyy"), 2)
    td$ = Mid$(Format$(todate, "mmddyyyy"), 3, 2)
    report.ReportFileName = nwc + "erlbdr.rpt"
    report.SelectionFormula = "{executionspay.datepaid} >= date(" + fy$ + "," + fm$ + "," + fd$ + ") and {executionspay.datepaid} <= date(" + ty$ + "," + tm$ + "," + td$ + ")"
    report.Destination = crptToWindow
    report.CopiesToPrinter = 1
    report.Action = 1
    Screen.MousePointer = 0
    Exit Sub
End If
If nullar.Value = True Then
    report.ReportFileName = nwc + "NULLAR.RPT"
    report.Destination = crptToWindow
    report.CopiesToPrinter = 1
    report.SelectionFormula = ""
    report.Action = 1
    Screen.MousePointer = 0
    db.Close
    Exit Sub
End If
If aer.Value = True Then
    report.ReportFileName = nwc + "ael.RPT"
    report.Destination = crptToWindow
    report.CopiesToPrinter = 1
    report.SelectionFormula = "{executions.balance} > 0 and {executions.nulla} <> 1"
    report.Action = 1
    Screen.MousePointer = 0
    db.Close
    Exit Sub
End If
If sfrrbdr.Value = True Then
    If Not IsDate(fromdate) Or Not IsDate(todate) Then
        msg = MsgBox("An invalid date exists in the DATE RANGE CRITERIA frame.", 48, "Genesis Error Log")
        Screen.MousePointer = 0
        db.Close
        Exit Sub
    End If
    On Error GoTo oderror18
OD18:
    Set db = OpenDatabase(nwc + dbname)
    Set ds = db.OpenRecordset("select * from passthru")
    If ds.EOF Then
        ds.AddNew
    Else
        ds.MoveFirst
        ds.Edit
    End If
    ds("textstring") = "Date Criteria: " + fromdate + " through " + todate
    ds.Update
    Set ds2 = db.OpenRecordset("select * from sfrrbdra")
    If Not ds2.EOF Then
        ds2.MoveFirst
        While Not ds2.EOF
            ds2.Delete
            ds2.MoveNext
        Wend
    End If
    Set ds = db.OpenRecordset("select serviceof, datereceived, servicefee, FEEDATE, receiptnum from magistrate where servicefee > 0 and feedate between #" + Format$(fromdate, "mm/dd/yyyy") + "# and #" + Format$(todate, "mm/dd/yyyy") + "#")
    If Not ds.EOF Then
        ds.MoveFirst
    End If
    While Not ds.EOF
        Set ds2 = db.OpenRecordset("select * from sfrrbdra")
        ds2.AddNew
        ds2("Service Of") = ds("serviceof")
        ds2("Date Received") = ds("feedate")
        ds2("Service Fee") = ds("servicefee")
        ds2("receiptnum") = ds("receiptnum")
        ds2.Update
        ds.MoveNext
    Wend
    Set ds = db.OpenRecordset("select serviceof, datereceived, servicefee, FEEDATE, receiptnum from writother where servicefee > 0 and feedate between #" + Format$(fromdate, "mm/dd/yyyy") + "# and #" + Format$(todate, "mm/dd/yyyy") + "#")
    If Not ds.EOF Then
        ds.MoveFirst
    End If
    While Not ds.EOF
        Set ds2 = db.OpenRecordset("select * from sfrrbdra")
        ds2.AddNew
        ds2("Service Of") = ds("serviceof")
        ds2("Date Received") = ds("feedate")
        ds2("Service Fee") = ds("servicefee")
        ds2("receiptnum") = ds("receiptnum")
        ds2.Update
        ds.MoveNext
    Wend
    Set ds = db.OpenRecordset("select serviceof, datereceived, servicefee, FEEDATE, receiptnum from executions where servicefee > 0 and feedate between #" + Format$(fromdate, "mm/dd/yyyy") + "# and #" + Format$(todate, "mm/dd/yyyy") + "#")
    If Not ds.EOF Then
        ds.MoveFirst
    End If
    While Not ds.EOF
        Set ds2 = db.OpenRecordset("select * from sfrrbdra")
        ds2.AddNew
        ds2("Service Of") = ds("serviceof")
        ds2("Date Received") = ds("feedate")
        ds2("Service Fee") = ds("servicefee")
        ds2("receiptnum") = ds("receiptnum")
        ds2.Update
        ds.MoveNext
    Wend
    db.Close
    report.SelectionFormula = ""
    report.Destination = crptToWindow
    report.CopiesToPrinter = 1
    report.ReportFileName = nwc + "sfrrbdra.rpt"
    report.Action = 1
    Screen.MousePointer = 0
    Exit Sub
End If
If RLOG.Value = True Then
    If Not IsDate(fromdate) Or Not IsDate(todate) Then
        msg = MsgBox("An invalid date exists in the DATE RANGE CRITERIA frame.", 48, "Genesis Error Log")
        Screen.MousePointer = 0
        db.Close
        Exit Sub
    End If
    fy$ = Right$(Format$(fromdate, "mmddyyyy"), 4)
    fm$ = Left$(Format$(fromdate, "mmddyyyy"), 2)
    fd$ = Mid$(Format$(fromdate, "mmddyyyy"), 3, 2)
    ty$ = Right$(Format$(todate, "mmddyyyy"), 4)
    tm$ = Left$(Format$(todate, "mmddyyyy"), 2)
    td$ = Mid$(Format$(todate, "mmddyyyy"), 3, 2)
    On Error GoTo oderror19
od19:
    Set ds = db.OpenRecordset("select * from receipt where datereceiVED between #" + Format$(fromdate, "mm/dd/yyyy") + "# and #" + Format$(todate, "mm/dd/yyyy") + "# and receiptnum is not null order by receiptnum")
    If Not ds.EOF Then
        ds.MoveFirst
    End If
    If Not ds.EOF Then
        ds.MoveFirst
        hl = ds("receiptnum")
        hd = ds("datereceiVED")
        ds.MoveNext
        While Not ds.EOF
            If ds("receiptnum") - hl < 100 And ds("receiptnum") - 1 <> hl Then
                For at! = hl + 1 To ds("receiptnum") - 1
                    Set ds2 = db.OpenRecordset("select * from receipt")
                    ds2.AddNew
                    ds2("serviceof") = "VOID"
                    ds2("datereceived") = hd
                    ds2("iteration") = ""
                    ds2("casenumber") = "VOID"
                    ds2("papertype") = "VOID"
                    ds2("servicefee") = Null
                    ds2("receiptnum") = Mid$(Str$(at!), 2)
                    ds2("datereceipt") = hd
                    ds2("checknum") = "VOID"
                    ds2("defendant") = "VOID"
                    ds2("plaintiff") = "VOID"
                    ds2("from") = "VOID"
                    ds2("fromaddress1") = "VOID"
                    ds2("fromaddress2") = "VOID"
                    ds2.Update
                Next at!
            End If
            hl = ds("receiptnum")
            hd = ds("datereceipt")
            ds.MoveNext
        Wend
    End If
    On Error GoTo oderror20
od20:
    Set db = OpenDatabase(nwc + dbname)
    Set ds = db.OpenRecordset("select * from passthru")
    If ds.EOF Then
        ds.AddNew
    Else
        ds.MoveFirst
        While Not ds.EOF
            ds.Delete
            ds.MoveNext
        Wend
        ds.AddNew
    End If
    ds("textstring") = "Date Criteria: " + fromdate + " through " + todate
    ds.Update
    db.Close
    report.SelectionFormula = "{receipt.datereceiVED} >= date(" + fy$ + "," + fm$ + "," + fd$ + ") and {receipt.datereceiVED} <= date(" + ty$ + "," + tm$ + "," + td$ + ")"
    report.Destination = crptToWindow
    report.CopiesToPrinter = 1
    report.ReportFileName = nwc + "rl.rpt"
    report.Action = 1
    Screen.MousePointer = 0
    Exit Sub
End If
If cbr.Value = True Then
    If Not IsDate(fromdate) Or Not IsDate(todate) Then
        msg = MsgBox("An invalid date exists in the DATE RANGE CRITERIA frame.", 48, "Genesis Error Log")
        Screen.MousePointer = 0
        db.Close
        Exit Sub
    End If
    On Error GoTo oderror21
od21:
    Set db = OpenDatabase(nwc + dbname)
    Set ds = db.OpenRecordset("select * from passthru")
    If ds.EOF Then
        ds.AddNew
    Else
        ds.MoveFirst
        ds.Edit
    End If
    ds("textstring") = "Date Criteria: " + fromdate + " through " + todate
    ds.Update
    fy$ = Right$(Format$(fromdate, "mmddyyyy"), 4)
    fm$ = Left$(Format$(fromdate, "mmddyyyy"), 2)
    fd$ = Mid$(Format$(fromdate, "mmddyyyy"), 3, 2)
    ty$ = Right$(Format$(todate, "mmddyyyy"), 4)
    tm$ = Left$(Format$(todate, "mmddyyyy"), 2)
    td$ = Mid$(Format$(todate, "mmddyyyy"), 3, 2)
    Set ds2 = db.OpenRecordset("select * from cbr")
    If Not ds2.EOF Then
        ds2.MoveFirst
        While Not ds2.EOF
            ds2.Delete
            ds2.MoveNext
        Wend
    End If
    Set ds = db.OpenRecordset("select serviceof, datereceived, servicefee, casenumber, defendant, plaintiff, feedate from magistrate where servicefee > 0 and feedate between #" + Format$(fromdate, "mm/dd/yyyy") + "# and #" + Format$(todate, "mm/dd/yyyy") + "#")
    If Not ds.EOF Then
        ds.MoveFirst
    End If
    While Not ds.EOF
        Set ds2 = db.OpenRecordset("select * from cbr")
        ds2.AddNew
        ds2("Service Of") = ds("serviceof")
        ds2("Date Received") = ds("feedate")
        ds2("Service Fee") = ds("servicefee")
        ds2("Case Number") = ds("casenumber")
        ds2("Defendant") = ds("defendant")
        ds2("Plaintiff") = ds("plaintiff")
        ds2("Fee Received") = ds("feedate")
        ds2.Update
        ds.MoveNext
    Wend
    Set ds = db.OpenRecordset("select serviceof, datereceived, servicefee, casenumber, defendant, plaintiff, feedate from WRITOTHER where servicefee > 0 and feedate between #" + Format$(fromdate, "mm/dd/yyyy") + "# and #" + Format$(todate, "mm/dd/yyyy") + "#")
    If Not ds.EOF Then
        ds.MoveFirst
    End If
    While Not ds.EOF
        Set ds2 = db.OpenRecordset("select * from cbr")
        ds2.AddNew
        ds2("Service Of") = ds("serviceof")
        ds2("Date Received") = ds("feedate")
        ds2("Service Fee") = ds("servicefee")
        ds2("Case Number") = ds("casenumber")
        ds2("Defendant") = ds("defendant")
        ds2("Plaintiff") = ds("plaintiff")
        ds2("Fee Received") = ds("feedate")
        ds2.Update
        ds.MoveNext
    Wend
    Set ds = db.OpenRecordset("select serviceof, datereceived, servicefee, casenumber, defendant, plaintiff, feedate from EXECUTIONS where servicefee > 0 and feedate between #" + Format$(fromdate, "mm/dd/yyyy") + "# and #" + Format$(todate, "mm/dd/yyyy") + "#")
    If Not ds.EOF Then
        ds.MoveFirst
    End If
    While Not ds.EOF
        Set ds2 = db.OpenRecordset("select * from cbr")
        ds2.AddNew
        ds2("Service Of") = ds("serviceof")
        ds2("Date Received") = ds("feedate")
        ds2("Service Fee") = ds("servicefee")
        ds2("Case Number") = ds("casenumber")
        ds2("Defendant") = ds("defendant")
        ds2("Plaintiff") = ds("plaintiff")
        ds2("Fee Received") = ds("feedate")
        ds2.Update
        ds.MoveNext
    Wend
    db.Close
    report.SelectionFormula = ""
    report.Destination = crptToWindow
    report.CopiesToPrinter = 2
    report.ReportFileName = nwc + "cbr1.rpt"
    report.Action = 1
    report.Destination = crptToWindow
    report.CopiesToPrinter = 1
    report.SelectionFormula = "{executionspay.datepaid} >= date(" + fy$ + "," + fm$ + "," + fd$ + ") and {executionspay.datepaid} <= date(" + ty$ + "," + tm$ + "," + td$ + ")"
    report.ReportFileName = nwc + "cbr2.rpt"
    report.Action = 1
    Screen.MousePointer = 0
    Exit Sub
End If
If msfrrbdr.Value = True Then
    If Not IsDate(fromdate) Or Not IsDate(todate) Then
        msg = MsgBox("An invalid date exists in the DATE RANGE CRITERIA frame.", 48, "Genesis Error Log")
        Screen.MousePointer = 0
        db.Close
        Exit Sub
    End If
    On Error GoTo oderror22
od22:
    Set db = OpenDatabase(nwc + dbname)
    Set ds = db.OpenRecordset("select * from passthru")
    If ds.EOF Then
        ds.AddNew
    Else
        ds.MoveFirst
        ds.Edit
    End If
    ds("textstring") = "Date Criteria: " + fromdate + " through " + todate
    ds.Update
    db.Close
    fy$ = Right$(Format$(fromdate, "mmddyyyy"), 4)
    fm$ = Left$(Format$(fromdate, "mmddyyyy"), 2)
    fd$ = Mid$(Format$(fromdate, "mmddyyyy"), 3, 2)
    ty$ = Right$(Format$(todate, "mmddyyyy"), 4)
    tm$ = Left$(Format$(todate, "mmddyyyy"), 2)
    td$ = Mid$(Format$(todate, "mmddyyyy"), 3, 2)
    report.Destination = crptToWindow
    report.CopiesToPrinter = 1
    report.ReportFileName = nwc + "sfrrbdrm.rpt"
    report.SelectionFormula = "{magistrate.FEEDATE} >= date(" + fy$ + "," + fm$ + "," + fd$ + ") and {magistrate.FEEDATE} <= date(" + ty$ + "," + tm$ + "," + td$ + ")"
    report.Action = 1
    Screen.MousePointer = 0
    Exit Sub
End If
If nmsfrrbdr.Value = True Then
    If Not IsDate(fromdate) Or Not IsDate(todate) Then
        msg = MsgBox("An invalid date exists in the DATE RANGE CRITERIA frame.", 48, "Genesis Error Log")
        Screen.MousePointer = 0
        db.Close
        Exit Sub
    End If
    On Error GoTo oderror23
od23:
    Set db = OpenDatabase(nwc + dbname)
    Set ds = db.OpenRecordset("select * from passthru")
    If ds.EOF Then
        ds.AddNew
    Else
        ds.MoveFirst
        ds.Edit
    End If
    ds("textstring") = "Date Criteria: " + fromdate + " through " + todate
    ds.Update
    db.Close
    fy$ = Right$(Format$(fromdate, "mmddyyyy"), 4)
    fm$ = Left$(Format$(fromdate, "mmddyyyy"), 2)
    fd$ = Mid$(Format$(fromdate, "mmddyyyy"), 3, 2)
    ty$ = Right$(Format$(todate, "mmddyyyy"), 4)
    tm$ = Left$(Format$(todate, "mmddyyyy"), 2)
    td$ = Mid$(Format$(todate, "mmddyyyy"), 3, 2)
    report.Destination = crptToWindow
    report.CopiesToPrinter = 1
    report.ReportFileName = nwc + "sfrrbdrw.rpt"
    report.SelectionFormula = "{writother.FEEDATE} >= date(" + fy$ + "," + fm$ + "," + fd$ + ") and {writother.FEEDATE} <= date(" + ty$ + "," + tm$ + "," + td$ + ")"
    report.Action = 1
    report.ReportFileName = nwc + "sfrrbdre.rpt"
    report.SelectionFormula = "{executions.FEEDATE} >= date(" + fy$ + "," + fm$ + "," + fd$ + ") and {executions.FEEDATE} <= date(" + ty$ + "," + tm$ + "," + td$ + ")"
    report.Action = 1
    Screen.MousePointer = 0
    Exit Sub
End If
If fcsopr.Value = True Then
    If Not IsDate(fromdate) Or Not IsDate(todate) Then
        msg = MsgBox("An invalid date exists in the DATE RANGE CRITERIA frame.", 48, "Genesis Error Log")
        Screen.MousePointer = 0
        db.Close
        Exit Sub
    End If
    On Error GoTo oderror24
od24:
    Set db = OpenDatabase(nwc + dbname)
    Set ds = db.OpenRecordset("select * from passthru")
    If ds.EOF Then
        ds.AddNew
    Else
        ds.MoveFirst
        ds.Edit
    End If
    ds("textstring") = "Date Criteria: " + fromdate + " through " + todate
    ds.Update
    db.Close
    fy$ = Right$(Format$(fromdate, "mmddyyyy"), 4)
    fm$ = Left$(Format$(fromdate, "mmddyyyy"), 2)
    fd$ = Mid$(Format$(fromdate, "mmddyyyy"), 3, 2)
    ty$ = Right$(Format$(todate, "mmddyyyy"), 4)
    tm$ = Left$(Format$(todate, "mmddyyyy"), 2)
    td$ = Mid$(Format$(todate, "mmddyyyy"), 3, 2)
    report.ReportFileName = nwc + "fcsopr.rpt"
    report.Destination = crptToWindow
    report.CopiesToPrinter = 1
    report.SelectionFormula = "{familycourt.served} = '1' and ({familycourt.servicedate} >= date(" + fy$ + "," + fm$ + "," + fd$ + ") and {familycourt.servicedate} <= date(" + ty$ + "," + tm$ + "," + td$ + "))"
    report.Action = 1
    Screen.MousePointer = 0
    Exit Sub
End If
If ivdsopr.Value = True Then
    If Not IsDate(fromdate) Or Not IsDate(todate) Then
        msg = MsgBox("An invalid date exists in the DATE RANGE CRITERIA frame.", 48, "Genesis Error Log")
        Screen.MousePointer = 0
        db.Close
        Exit Sub
    End If
    On Error GoTo oderror25
od25:
    Set db = OpenDatabase(nwc + dbname)
    Set ds = db.OpenRecordset("select * from passthru")
    If ds.EOF Then
        ds.AddNew
    Else
        ds.MoveFirst
        ds.Edit
    End If
    ds("textstring") = "Date Criteria: " + fromdate + " through " + todate
    ds.Update
    db.Close
    fy$ = Right$(Format$(fromdate, "mmddyyyy"), 4)
    fm$ = Left$(Format$(fromdate, "mmddyyyy"), 2)
    fd$ = Mid$(Format$(fromdate, "mmddyyyy"), 3, 2)
    ty$ = Right$(Format$(todate, "mmddyyyy"), 4)
    tm$ = Left$(Format$(todate, "mmddyyyy"), 2)
    td$ = Mid$(Format$(todate, "mmddyyyy"), 3, 2)
    report.ReportFileName = nwc + "FDsopr.rpt"
    report.Destination = crptToWindow
    report.CopiesToPrinter = 1
    report.SelectionFormula = "{familycourt.served} = '1' and ({familycourt.servicedate} >= date(" + fy$ + "," + fm$ + "," + fd$ + ") and {familycourt.servicedate} <= date(" + ty$ + "," + tm$ + "," + td$ + ")) and {familycourt.ivd} = 1"
    report.Action = 1
    Screen.MousePointer = 0
    Exit Sub
End If
If opsr.Value = True Then
    If Not IsDate(fromdate) Or Not IsDate(todate) Then
        msg = MsgBox("An invalid date exists in the DATE RANGE CRITERIA frame.", 48, "Genesis Error Log")
        Screen.MousePointer = 0
        db.Close
        Exit Sub
    End If
    inp = InputBox("Enter A for all officers or the Name of a specific officer.", "Genesis Information Log", "A")
    If inp = "" Then
        inp = "A"
    End If
    inp = UCase(inp)
    Call printroutines("opsrrtn", inp)
    If NOREPORT = 1 Then
        msg = MsgBox("No data exists for report criteria.", 48, "Genesis Information Log")
        Screen.MousePointer = 0
        db.Close
        Exit Sub
    End If
    Set db = OpenDatabase(nwc + dbname)
    Set ds = db.OpenRecordset("select * from passthru")
    If ds.EOF Then
        ds.AddNew
    Else
        ds.MoveFirst
        ds.Edit
    End If
    ds("textstring") = "Date Criteria: " + fromdate + " through " + todate
    ds.Update
    db.Close
    report.ReportFileName = nwc + "opsr.rpt"
    report.SelectionFormula = ""
    report.Destination = crptToWindow
    report.CopiesToPrinter = 1
    report.Action = 1
    Screen.MousePointer = 0
    Exit Sub
End If
If opbor.Value = True Then
    inp = InputBox("Enter A for all officers or the Name of a specific officer.", "Genesis Information Log", "A")
    If inp = "" Then
        inp = "A"
    End If
    inp = UCase(inp)
    Call printroutines("opborrtn", inp)
    If NOREPORT = 1 Then
        msg = MsgBox("No data exists for report criteria.", 48, "Genesis Information Log")
        Screen.MousePointer = 0
        db.Close
        Exit Sub
    End If
    report.ReportFileName = nwc + "opbor.rpt"
    report.SelectionFormula = ""
    report.Destination = crptToWindow
    report.CopiesToPrinter = 1
    report.Action = 1
    Screen.MousePointer = 0
    Exit Sub
End If
If ompbor.Value = True Then
    inp = InputBox("Enter A for all officers or the Name of a specific officer.", "Genesis Information Log", "A")
    If inp = "" Then
        inp = "A"
    End If
    inp = UCase(inp)
    Call printroutines("ompborrtn", inp)
    If NOREPORT = 1 Then
        msg = MsgBox("No data exists for report criteria.", 48, "Genesis Information Log")
        Screen.MousePointer = 0
        db.Close
        Exit Sub
    End If
    report.ReportFileName = nwc + "ompbor.rpt"
    report.Destination = crptToWindow
    report.CopiesToPrinter = 1
    report.SelectionFormula = ""
    report.Action = 1
    Screen.MousePointer = 0
    Exit Sub
End If
If oepbor.Value = True Then
    inp = InputBox("Enter A for all officers or the Name of a specific officer.", "Genesis Information Log", "A")
    If inp = "" Then
        inp = "A"
    End If
    inp = UCase(inp)
    Call printroutines("oepborrtn", inp)
    If NOREPORT = 1 Then
        msg = MsgBox("No data exists for report criteria.", 48, "Genesis Information Log")
        Screen.MousePointer = 0
        db.Close
        Exit Sub
    End If
    report.ReportFileName = nwc + "oepbor.rpt"
    report.Destination = crptToWindow
    report.CopiesToPrinter = 1
    report.SelectionFormula = ""
    report.Action = 1
    Screen.MousePointer = 0
    Exit Sub
End If
If ofcbor.Value = True Then
    inp = InputBox("Enter A for all officers or the Name of a specific officer.", "Genesis Information Log", "A")
    If inp = "" Then
        inp = "A"
    End If
    inp = UCase(inp)
    Call printroutines("ofcpborrtn", inp)
    If NOREPORT = 1 Then
        msg = MsgBox("No data exists for report criteria.", 48, "Genesis Information Log")
        Screen.MousePointer = 0
        db.Close
        Exit Sub
    End If
    report.ReportFileName = nwc + "ofcpbor.rpt"
    report.Destination = crptToWindow
    report.CopiesToPrinter = 1
    report.SelectionFormula = ""
    report.Action = 1
    Screen.MousePointer = 0
    Exit Sub
End If
'chris
If owbor.Value = True Then
    inp = InputBox("Enter A for all officers or the Name of a specific officer.", "Genesis Information Log", "A")
    If inp = "" Then
        inp = "A"
    End If
    inp = UCase(inp)
    Call printroutines("owpborrtn", inp)
    If NOREPORT = 1 Then
        msg = MsgBox("No data exists for report criteria.", 48, "Genesis Information Log")
        Screen.MousePointer = 0
        db.Close
        Exit Sub
    End If
    report.ReportFileName = nwc + "owpbor.rpt"
    report.Destination = crptToWindow
    report.CopiesToPrinter = 1
    report.SelectionFormula = ""
    report.Action = 1
    Screen.MousePointer = 0
    Exit Sub
End If
'chris
Screen.MousePointer = 0
db.Close
On Error GoTo 0
Exit Sub
oderror:
If Err > 3200 Then
    If Err = 20525 Then
        MsgBox "Error opening report", vbOKOnly, "Geneiss Error Log"
        Exit Sub
    End If
    Resume od
Else
    Resume Next
End If
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
oderror3:
If Err > 3200 Then
    Resume od3
Else
    Resume Next
End If
oderror4:
If Err > 3200 Then
    Resume od4
Else
    Resume Next
End If
oderror5:
If Err > 3200 Then
    Resume od5
Else
    Resume Next
End If
oderror6:
If Err > 3200 Then
    Resume od6
Else
    Resume Next
End If
oderror7:
If Err > 3200 Then
    Resume od7
Else
    Resume Next
End If
oderror8:
If Err > 3200 Then
    Resume od8
Else
    Resume Next
End If
oderror9:
If Err > 3200 Then
    Resume od9
Else
    Resume Next
End If
oderror10:
If Err > 3200 Then
    Resume od10
Else
    Resume Next
End If
oderror11:
If Err > 3200 Then
    Resume od11
Else
    Resume Next
End If
oderror12:
If Err > 3200 Then
    Resume od12
Else
    Resume Next
End If
oderror13:
If Err > 3200 Then
    Resume OD13
Else
    Resume Next
End If
oderror14:
If Err > 3200 Then
    Resume OD14
Else
    Resume Next
End If
oderror15:
If Err > 3200 Then
    Resume OD15
Else
    Resume Next
End If
oderror16:
If Err > 3200 Then
    Resume OD16
Else
    Resume Next
End If
oderror17:
If Err > 3200 Then
    Resume OD17
Else
    Resume Next
End If
oderror18:
If Err > 3200 Then
    Resume OD18
Else
    Resume Next
End If
oderror19:
If Err > 3200 Then
    Resume od19
Else
    Resume Next
End If
oderror20:
If Err > 3200 Then
    Resume od20
Else
    Resume Next
End If
oderror21:
If Err > 3200 Then
    Resume od21
Else
    Resume Next
End If
oderror22:
If Err > 3200 Then
    Resume od22
Else
    Resume Next
End If
oderror23:
If Err > 3200 Then
    Resume od23
Else
    Resume Next
End If
oderror24:
If Err > 3200 Then
    Resume od24
Else
    Resume Next
End If
oderror25:
If Err > 3200 Then
    Resume od25
Else
    Resume Next
End If
End Sub

Private Sub savebutton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    blnSavePressed = True
End Sub

Private Sub served_Click()
On Error Resume Next
nonservice.SetFocus
End Sub

Private Sub served_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    nonservice.SetFocus
End If

End Sub

Private Sub servicedate_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
If Len(servicedate) = 1 Or Len(servicedate) = 4 Then
    Call sendslash
End If
End If
If KeyAscii = 13 Then
    servicetime.SetFocus
End If

End Sub


Private Sub servicefee_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If maintab.Tab = 2 Then
        ivd.SetFocus
    Else
        feedate.SetFocus
    End If
End If

End Sub

Private Sub serviceof_Click(AREA As Integer)
On Error Resume Next
If serviceof = "" Then
    Exit Sub
End If
Call setpopup(serviceof, "F")
CSERVICEOF = 1
If maintab.Tab > 3 Then
    GoSub tab1
End If
On maintab.Tab + 1 GoSub tab1, tab2, tab3, tab4
daterlist.clear
Dim db As Database, ds As Recordset

On Error GoTo oderror
od:
Set db = OpenDatabase(nwc + dbname)
If serviceof > "" Then
        Set ds = db.OpenRecordset("select distinct datereceived from " + TP + " where serviceof = " + Chr$(34) + serviceof + Chr$(34) + " order by datereceived")
Else
        Set ds = db.OpenRecordset("select distinct datereceived from " + TP + " order by datereceived")
End If
If Not ds.EOF Then
    ds.MoveFirst
End If
While Not ds.EOF
    daterlist.AddItem ds("datereceived")
    ds.MoveNext
Wend
If datereceived = "" Then
    datereceived = Format$(Date$, "mm/dd/yyyy")
End If
If paymentframe.Visible = True Or indexframe.Visible = True Or FROMLF = 1 Then
    db.Close
    Exit Sub
End If

If datereceived = "" Then
    db.Close
    Exit Sub
End If
If iteration = "" Then
    db.Close
    Exit Sub
End If
If Not IsDate(datereceived) Then
   msg = MsgBox("Filter entry in DATE RECEIVED is not a valid date.", 48, "Genesis Error Log")
   datereceived.SetFocus
   db.Close
   Exit Sub
End If
If Val(iteration) = 0 Then
   msg = MsgBox("Filter entry in ITERATION is not a valid number.", 48, "Genesis Error Log")
   iteration.SetFocus
   db.Close
   Exit Sub
End If
If maintab.Tab > 3 Then
    GoSub tab1
End If
On maintab.Tab + 1 GoSub tab1, tab2, tab3, tab4
Set ds = db.OpenRecordset("select * from " + TP + " where serviceof = " + Chr$(34) + serviceof + Chr$(34) + " and datereceived = #" + datereceived + "# and iteration = " + Chr$(34) + iteration + Chr$(34))
If ds.EOF Then
    db.Close
    Exit Sub
End If
ds.MoveFirst
If maintab.Tab = 2 Then
    If Not IsNull(ds("osce")) Then
        servicefee = ds("osce")
    Else
        servicefee = ""
    End If
    If Not IsNull(ds("IVD")) Then
        ivd.Value = ds("ivd")
    Else
        ivd.Value = 0
    End If
    If Not IsNull(ds("fl1")) Then
        custodian = ds("fl1")
    Else
        custodian = ""
    End If
End If
If maintab.Tab <> 2 Then
    If Not IsNull(ds("receiptnum")) Then
        receiptd = ds("receiptnum")
    Else
        receiptd = ""
    End If
    If Not IsNull(ds("checknum")) Then
        checkd = ds("checknum")
    Else
        checkd = ""
    End If
End If
serviceof = ds("serviceof")
datereceived = Format$(ds("datereceived"), "mm/dd/yyyy")
iteration = ds("iteration")
serviceofsort = ds("serviceofsort")
If Not IsNull(ds("CASENUMBER")) Then
    casenumber = ds("casenumber")
Else
    casenumber = "UNKNOWN"
End If
If Not IsNull(ds("fs2")) Then
    armedforces.Value = Val(ds("fs2"))
Else
    armedforces.Value = 0
End If
If Not IsNull(ds("fs1")) Then
    corporate.Value = Val(ds("fs1"))
Else
    corporate.Value = 0
End If
If Not IsNull(ds("fl1")) Then
    title = ds("fl1")
Else
    title = ""
End If

If Not IsNull(ds("sohomeaddress")) Then
    sohomeaddress = ds("sohomeaddress")
Else
    sohomeaddress = ""
End If
If Not IsNull(ds("sohomeaddress2")) Then
    sohomeaddress2 = ds("sohomeaddress2")
Else
    sohomeaddress2 = ""
End If
If Not IsNull(ds("sohomestate")) Then
    sohomestate = ds("sohomestate")
Else
    sohomestate = ""
End If
If Not IsNull(ds("sohomezipcode")) Then
    sohomezipcode = ds("sohomezipcode")
Else
    sohomezipcode = ""
End If
If Not IsNull(ds("sohomephone")) And ds("sohomephone") <> "" Then
    sohomephone = ds("sohomephone")
Else
    sohomephone = ""
End If
If Not IsNull(ds("soworkaddress")) Then
    soworkaddress = ds("soworkaddress")
Else
    soworkaddress = ""
End If
If Not IsNull(ds("soworkaddress2")) Then
    soworkaddress2 = ds("soworkaddress2")
Else
    soworkaddress2 = ""
End If
If Not IsNull(ds("soworkstate")) Then
    soworkstate = ds("soworkstate")
Else
    soworkstate = ""
End If
If Not IsNull(ds("soworkzipcode")) Then
    soworkzipcode = ds("soworkzipcode")
Else
    soworkzipcode = ""
End If
If Not IsNull(ds("soworkphone")) And ds("soworkphone") <> "" Then
    soworkphone = ds("soworkphone")
Else
    soworkphone = ""
End If
papertype = ds("papertype")
If Not IsNull(ds("courtdate")) Then
    courtdate = Format$(ds("courtdate"), "mm/dd/yyyy")
Else
    courtdate = ""
End If
If Not IsNull(ds("courttime")) Then
    courttime = ds("courttime")
Else
    courttime = ""
End If
If Not IsNull(ds("daystorespond")) Then
    daystorespond = ds("daystorespond")
Else
    daystorespond = ""
End If
If maintab.Tab <> 2 Then
    If Not IsNull(ds("servicefee")) Then
        servicefee = ds("servicefee")
    Else
        servicefee = ""
    End If
    If Not IsNull(ds("bill")) Then
        bill = ds("bill")
    Else
        bill = 0
    End If
    If Not IsNull(ds("feedate")) Then
        feedate = ds("feedate")
    Else
        feedate = ""
    End If
End If
defendant = ds("defendant")
defendantsort = ds("defendantsort")
If Not IsNull(ds("dhomeaddress")) Then
    dhomeaddress = ds("dhomeaddress")
Else
    dhomeaddress = ""
End If
If Not IsNull(ds("dhomeaddress2")) Then
    dhomeaddress2 = ds("dhomeaddress2")
Else
    dhomeaddress2 = ""
End If
If Not IsNull(ds("dhomestate")) Then
    dhomestate = ds("dhomestate")
Else
    dhomestate = ""
End If
If Not IsNull(ds("dhomezipcode")) Then
    dhomezipcode = ds("dhomezipcode")
Else
    dhomezipcode = ""
End If
If Not IsNull(ds("dhomephone")) And ds("dhomephone") <> "" Then
    dhomephone = ds("dhomephone")
Else
    dhomephone = ""
End If
If Not IsNull(ds("dworkaddress")) Then
    dworkaddress = ds("dworkaddress")
Else
    dworkaddress = ""
End If
If Not IsNull(ds("dworkaddress2")) Then
    dworkaddress2 = ds("dworkaddress2")
Else
    dworkaddress2 = ""
End If
If Not IsNull(ds("dworkstate")) Then
    dworkstate = ds("dworkstate")
Else
    dworkstate = ""
End If
If Not IsNull(ds("dworkzipcode")) Then
    dworkzipcode = ds("dworkzipcode")
Else
    dworkzipcode = ""
End If
If Not IsNull(ds("dworkphone")) And ds("dworkphone") <> "" Then
    dworkphone = ds("dworkphone")
Else
    dworkphone = ""
End If
plaintiff = ds("plaintiff")
plaintiffsort = ds("plaintiffsort")
If Not IsNull(ds("phomeaddress")) Then
    phomeaddress = ds("phomeaddress")
Else
    phomeaddress = ""
End If
If Not IsNull(ds("phomeaddress2")) Then
    phomeaddress2 = ds("phomeaddress2")
Else
    phomeaddress2 = ""
End If
If Not IsNull(ds("phomestate")) Then
    phomestate = ds("phomestate")
Else
    phomestate = ""
End If
If Not IsNull(ds("phomezipcode")) Then
    phomezipcode = ds("phomezipcode")
Else
    phomezipcode = ""
End If
If Not IsNull(ds("phomephone")) And ds("phomephone") <> "" Then
    phomephone = ds("phomephone")
Else
    phomephone = ""
End If
If Not IsNull(ds("pworkaddress")) Then
    pworkaddress = ds("pworkaddress")
Else
    pworkaddress = ""
End If
If Not IsNull(ds("pworkaddress2")) Then
    pworkaddress2 = ds("pworkaddress2")
Else
    pworkaddress2 = ""
End If
If Not IsNull(ds("pworkstate")) Then
    pworkstate = ds("pworkstate")
Else
    pworkstate = ""
End If
If Not IsNull(ds("pworkzipcode")) Then
    pworkzipcode = ds("pworkzipcode")
Else
    pworkzipcode = ""
End If
If Not IsNull(ds("pworkphone")) And ds("pworkphone") <> "" Then
    pworkphone = ds("pworkphone")
Else
    pworkphone = ""
End If
If Not IsNull(ds("assignedto")) Then
    assignedto = ds("assignedto")
Else
    assignedto = ""
End If
If Not IsNull(ds("assignedon")) Then
    assignedon = ds("assignedon")
Else
    assignedon = ""
End If
If Not IsNull(ds("served")) Then
    served.Value = Val(ds("served"))
Else
    served.Value = 0
End If
If Not IsNull(ds("nonservice")) Then
    nonservice.Value = Val(ds("nonservice"))
Else
    nonservice.Value = 0
End If
If Not IsNull(ds("nsreason")) Then
    nsreason = ds("nsreason")
Else
    nsreason = ""
End If
If Not IsNull(ds("premarks")) Then
    premarks = ds("premarks")
Else
    premarks = ""
End If
If maintab.Tab = 3 Then
    If Not IsNull(ds("levy")) Then
        levy.Text = ds("levy")
    Else
        levy.Text = ""
    End If
End If
If Not IsNull(ds("wremarks")) Then
    wremarks = ds("wremarks")
Else
    wremarks = ""
End If

If Not IsNull(ds("servicedate")) Then
    servicedate = Format$(ds("servicedate"), "mm/dd/yyyy")
Else
    servicedate = ""
End If
If Not IsNull(ds("servicetime")) Then
    servicetime = ds("servicetime")
Else
    servicetime = ""
End If
If Not IsNull(ds("personserved")) Then
    personserved = ds("personserved")
Else
    personserved = ""
End If
If Not IsNull(ds("locationserved")) Then
    locationserved = ds("locationserved")
Else
    locationserved = ""
End If
If Not IsNull(ds("relationship")) Then
    relationship = ds("relationship")
Else
    relationship = ""
End If
If Not IsNull(ds("professional")) Then
    professional = ds("professional")
Else
    professional = ""
End If
If maintab.Tab = 3 Then
        If Not IsNull(ds("apptdate")) Then
                apptdate = Format$(ds("apptdate"), "mm/dd/yyyy")
        Else
                apptdate = ""
        End If
        If Not IsNull(ds("intrate")) Then
                intrate = ds("intrate")
        Else
                intrate = ""
        End If
        If Not IsNull(ds("datesatisfied")) Then
                datesatisfied = Format$(ds("datesatisfied"), "mm/dd/yyyy")
        Else
                datesatisfied = ""
        End If
        If Not IsNull(ds("judgementdate")) Then
                judgementdate = Format$(ds("judgementdate"), "mm/dd/yyyy")
        Else
                judgementdate = ""
        End If
        If Not IsNull(ds("judgementamount")) Then
                judgementamount = ds("judgementamount")
        Else
                judgementamount = ""
        End If
        'If Not IsNull(ds("estpayoffdate")) Then
        '        estpayoffdate = Format$(ds("estpayoffdate"), "mm/dd/yyyy")
        'Else
        '        estpayoffdate = ""
        'End If
        estpayoffdate = Format$(Date$, "mm/dd/yyyy")
        If Not IsNull(ds("nulla")) Then
            nulla.Value = ds("nulla")
        Else
            nulla.Value = 0
        End If
        commission = ds("COMMISSION")
        INTEREST = ds("INTERest")
        perday = ds("PERDAY")
        balance = ds("BALANCE")
        If Not IsNull(ds("totalinterest")) Then
            totalinterest = ds("totalinterest")
        Else
            totalinterest = 0
        End If
        If Not IsNull(ds("totalcommission")) Then
            totalcommission = ds("totalcommission")
        Else
            totalcommission = 0
        End If
        If Not IsNull(ds("totalpayments")) Then
            totalpayments = ds("totalpayments")
        Else
            totalpayments = 0
        End If
        Call loadpay
End If
lastserviceof = serviceof
CSERVICEOF = 0
infoframe.Refresh
db.Close
On Error GoTo 0
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If


tab1:
TP = "magistrate"
Return
tab2:
TP = "writother"
Return
tab3:
TP = "familycourt"
Return
tab4:
TP = "executions"
Return


End Sub

Private Sub serviceof_GotFocus()
a = 1
End Sub

Private Sub serviceof_KeyUp(KeyCode As Integer, Shift As Integer)
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

Private Sub serviceof_LostFocus()
If Len(serviceof) > 60 Then
    msg = MsgBox("SERVICE OF can be no more than 60 characters.  The data will be truncated.", 48, "Genesis Error Log")
    serviceof = Left$(serviceof, 60)
End If
If maintab.Tab = 3 Then
    FROMLF = 1
Else
    FROMLF = 0
End If
If serviceof > "" And CSERVICEOF = 1 Then
    Call serviceof_Click(0)
End If

End Sub

Private Sub serviceofsort_GotFocus()
If serviceofsort > "" Then
    Exit Sub
End If
Dim db As Database, rs As Recordset, ff, LF As Integer, HS As String
ff = 0
LF = 1
HS = ""
On Error GoTo oderror
od:
Set db = OpenDatabase(nwc + dbname)
Set rs = db.OpenRecordset("select fnf,lnf from system")
If Not rs.EOF Then
    rs.MoveFirst
    If rs("fNf") = True Then
        ff = 1
        LF = 0
    End If
End If
db.Close
Call setsort(ff, LF, serviceof, HS)
serviceofsort = HS
On Error GoTo 0
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If



End Sub

Private Sub serviceofsort_KeyPress(KeyAscii As Integer)
If KeyAscii = 34 Then
    msg = MsgBox("The " + Chr$(34) + " character is not allowed.  Use the ' character instead.", 48, "Genesis Error Log")
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
    sohomeaddress.SetFocus
End If

End Sub

Private Sub servicetime_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
If Len(servicetime) = 1 Then
    Call sendcolon
End If
End If
If KeyAscii = 13 Then
    personserved.SetFocus
End If

End Sub

Private Sub sfrrbdr_Click()
If sfrrbdr.Value = True Then
    fromdate.SetFocus
End If

End Sub

Private Sub likebutton_Click()
SEARCHTYPE = 0
If serviceof = "" And sohomeaddress = "" And soworkaddress = "" And sohomeaddress2 = "" And sohomestate = "" And sohomezipcode = "" And soworkaddress2 = "" And soworkstate = "" And soworkzipcode = "" Then
    msg = MsgBox("You must enter a partial or complete name or address in the SERVICE OF, HOME, or WORK ADDRESS fields to do a LIKE lookup.", 48, "Genesis Error Log")
    serviceof.SetFocus
    Exit Sub
End If
Dim wrds(10), wrdha(10), wrdwa(10), wrdha2(10), wrdwa2(10), wrdha3(10), wrdwa3(10), wrdha4(10), wrdwa4(10), twrd As String, ctpos, ctwrd As Long
For t% = 1 To 10
    wrds(t%) = ""
    wrdha(t%) = ""
    wrdwa(t%) = ""
Next t%
ctwrd = 0
twrd = serviceof
If InStr(twrd, " ") > 0 Then
    While InStr(twrd, " ") > 0 And ctwrd < 11
        ctwrd = ctwrd + 1
        wrds(ctwrd) = Left$(twrd, InStr(twrd, " ") - 1)
        twrd = Mid$(twrd, InStr(twrd, " ") + 1)
    Wend
Else
If twrd > "" Then
    wrds(1) = twrd
End If
End If
ctwrd = 0
twrd = sohomeaddress
If InStr(twrd, " ") > 0 Then
    While InStr(twrd, " ") > 0 And ctwrd < 11
        ctwrd = ctwrd + 1
        wrdha(ctwrd) = Left$(twrd, InStr(twrd, " ") - 1)
        twrd = Mid$(twrd, InStr(twrd, " ") + 1)
    Wend
Else
If twrd > "" Then
    wrdha(1) = twrd
End If
End If
ctwrd = 0
twrd = sohomeaddress2
If InStr(twrd, " ") > 0 Then
    While InStr(twrd, " ") > 0 And ctwrd < 11
        ctwrd = ctwrd + 1
        wrdha2(ctwrd) = Left$(twrd, InStr(twrd, " ") - 1)
        twrd = Mid$(twrd, InStr(twrd, " ") + 1)
    Wend
Else
If twrd > "" Then
    wrdha2(1) = twrd
End If
End If
ctwrd = 0
twrd = sohomestate
If InStr(twrd, " ") > 0 Then
    While InStr(twrd, " ") > 0 And ctwrd < 11
        ctwrd = ctwrd + 1
        wrdha3(ctwrd) = Left$(twrd, InStr(twrd, " ") - 1)
        twrd = Mid$(twrd, InStr(twrd, " ") + 1)
    Wend
Else
If twrd > "" Then
    wrdha3(1) = twrd
End If
End If
ctwrd = 0
twrd = sohomezipcode
If InStr(twrd, " ") > 0 Then
    While InStr(twrd, " ") > 0 And ctwrd < 11
        ctwrd = ctwrd + 1
        wrdha4(ctwrd) = Left$(twrd, InStr(twrd, " ") - 1)
        twrd = Mid$(twrd, InStr(twrd, " ") + 1)
    Wend
Else
If twrd > "" Then
    wrdha4(1) = twrd
End If
End If
ctwrd = 0
twrd = soworkaddress
If InStr(twrd, " ") > 0 Then
    While InStr(twrd, " ") > 0 And ctwrd < 11
        ctwrd = ctwrd + 1
        wrdwa(ctwrd) = Left$(twrd, InStr(twrd, " ") - 1)
        twrd = Mid$(twrd, InStr(twrd, " ") + 1)
    Wend
Else
If twrd > "" Then
    wrdwa(1) = twrd
End If
End If
ctwrd = 0
twrd = soworkaddress2
If InStr(twrd, " ") > 0 Then
    While InStr(twrd, " ") > 0 And ctwrd < 11
        ctwrd = ctwrd + 1
        wrdwa2(ctwrd) = Left$(twrd, InStr(twrd, " ") - 1)
        twrd = Mid$(twrd, InStr(twrd, " ") + 1)
    Wend
Else
If twrd > "" Then
    wrdwa2(1) = twrd
End If
End If
ctwrd = 0
twrd = soworkstate
If InStr(twrd, " ") > 0 Then
    While InStr(twrd, " ") > 0 And ctwrd < 11
        ctwrd = ctwrd + 1
        wrdwa3(ctwrd) = Left$(twrd, InStr(twrd, " ") - 1)
        twrd = Mid$(twrd, InStr(twrd, " ") + 1)
    Wend
Else
If twrd > "" Then
    wrdwa3(1) = twrd
End If
End If
ctwrd = 0
twrd = soworkzipcode
If InStr(twrd, " ") > 0 Then
    While InStr(twrd, " ") > 0 And ctwrd < 11
        ctwrd = ctwrd + 1
        wrdwa4(ctwrd) = Left$(twrd, InStr(twrd, " ") - 1)
        twrd = Mid$(twrd, InStr(twrd, " ") + 1)
    Wend
Else
If twrd > "" Then
    wrdwa4(1) = twrd
End If
End If
If maintab.Tab > 3 Then
    GoSub tab1
End If
Screen.MousePointer = 11
On maintab.Tab + 1 GoSub tab1, tab2, tab3, tab4
Dim db As Database, ds As Recordset, tsql As String
tsql = ""
On Error GoTo oderror
od:
Set db = OpenDatabase(nwc + dbname)
For ctwrd = 1 To 10
    If wrds(ctwrd) > "" Then
        If tsql = "" Then
            tsql = "where serviceof like '*" + wrds(ctwrd) + "*'"
        Else
            tsql = tsql + " and serviceof like '*" + wrds(ctwrd) + "*'"
        End If
    End If
Next ctwrd
For ctwrd = 1 To 10
    If wrdha(ctwrd) > "" Then
        If tsql = "" Then
            tsql = "where sohomeaddress like '*" + wrdha(ctwrd) + "*'"
        Else
            tsql = tsql + " and sohomeaddress like '*" + wrdha(ctwrd) + "*'"
        End If
    End If
Next ctwrd
For ctwrd = 1 To 10
    If wrdha2(ctwrd) > "" Then
        If tsql = "" Then
            tsql = "where sohomeaddress2 like '*" + wrdha2(ctwrd) + "*'"
        Else
            tsql = tsql + " and sohomeaddress2 like '*" + wrdha2(ctwrd) + "*'"
        End If
    End If
Next ctwrd
For ctwrd = 1 To 10
    If wrdha3(ctwrd) > "" Then
        If tsql = "" Then
            tsql = "where sohomestate like '*" + wrdha3(ctwrd) + "*'"
        Else
            tsql = tsql + " and sohomestate like '*" + wrdha3(ctwrd) + "*'"
        End If
    End If
Next ctwrd
For ctwrd = 1 To 10
    If wrdha4(ctwrd) > "" Then
        If tsql = "" Then
            tsql = "where sohomezipcode like '*" + wrdha4(ctwrd) + "*'"
        Else
            tsql = tsql + " and sohomezipcode like '*" + wrdha4(ctwrd) + "*'"
        End If
    End If
Next ctwrd
For ctwrd = 1 To 10
    If wrdwa(ctwrd) > "" Then
        If tsql = "" Then
            tsql = "where soworkaddress like '*" + wrdwa(ctwrd) + "*'"
        Else
            tsql = tsql + " and soworkaddress like '*" + wrdwa(ctwrd) + "*'"
        End If
    End If
Next ctwrd
For ctwrd = 1 To 10
    If wrdwa2(ctwrd) > "" Then
        If tsql = "" Then
            tsql = "where soworkaddress2 like '*" + wrdwa2(ctwrd) + "*'"
        Else
            tsql = tsql + " and soworkaddress2 like '*" + wrdwa2(ctwrd) + "*'"
        End If
    End If
Next ctwrd
For ctwrd = 1 To 10
    If wrdwa3(ctwrd) > "" Then
        If tsql = "" Then
            tsql = "where soworkstate like '*" + wrdwa3(ctwrd) + "*'"
        Else
            tsql = tsql + " and soworkstate like '*" + wrdwa3(ctwrd) + "*'"
        End If
    End If
Next ctwrd
For ctwrd = 1 To 10
    If wrdwa4(ctwrd) > "" Then
        If tsql = "" Then
            tsql = "where soworkzipcode like '*" + wrdwa4(ctwrd) + "*'"
        Else
            tsql = tsql + " and soworkzipcode like '*" + wrdwa4(ctwrd) + "*'"
        End If
    End If
Next ctwrd
Set ds = db.OpenRecordset("select * from " + TP + " " + tsql + " order by serviceofsort,SERVICEOF,DATERECEIVED,ITERATION")
If ds.EOF Then
    msg = MsgBox("No eligible records for retrieval on this tab.", 48, "Genesis Error Log")
    db.Close
    Screen.MousePointer = 0
    Exit Sub
End If
ds.MoveFirst
likelist.ListItems.clear
While Not ds.EOF
    Set itmx = likelist.ListItems.add(, , ds("serviceof"))
    itmx.SubItems(1) = Format$(ds("DATERECEIVED"), "mm/dd/yyyy")
    itmx.SubItems(2) = ds("iteration")
    If Not IsNull(ds("sohomeaddress")) Then
        itmx.SubItems(3) = ds("sohomeaddress")
    End If
    If Not IsNull(ds("sohomeaddress2")) Then
        itmx.SubItems(3) = itmx.SubItems(3) + " " + ds("sohomeaddress2")
    End If
    If Not IsNull(ds("sohomestate")) Then
        itmx.SubItems(3) = itmx.SubItems(3) + " " + ds("sohomestate")
    End If
    If Not IsNull(ds("sohomezipcode")) Then
        itmx.SubItems(3) = itmx.SubItems(3) + " " + ds("sohomezipcode")
    End If
    If Not IsNull(ds("soworkaddress")) Then
        itmx.SubItems(4) = ds("soworkaddress")
    End If
    If Not IsNull(ds("soworkaddress2")) Then
        itmx.SubItems(4) = itmx.SubItems(4) + " " + ds("soworkaddress2")
    End If
    If Not IsNull(ds("soworkstate")) Then
        itmx.SubItems(3) = itmx.SubItems(3) + " " + ds("soworkstate")
    End If
    If Not IsNull(ds("soworkzipcode")) Then
        itmx.SubItems(3) = itmx.SubItems(3) + " " + ds("soworkzipcode")
    End If
    ds.MoveNext
Wend
likeframe.Left = 400
likeframe.Top = 1100
likeframe.Visible = True
Screen.MousePointer = 0
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


tab1:
TP = "magistrate"
Return
tab2:
TP = "writother"
Return
tab3:
TP = "familycourt"
Return
tab4:
TP = "executions"
Return
End Sub

Private Sub papertype_Click()
infoframe.Refresh

End Sub



Private Sub professional_Click()
infoframe.Refresh

End Sub

Private Sub savebutton_Click()
If UCase(frmLogin.txtUserName) = "DEMO" And UCase(frmLogin.txtPassword) = "DEMO" Then
    msg = MsgBox("Not available in DEMO version.", 48, "Genesis Information Log")
    Screen.MousePointer = 0
    Exit Sub
End If
Dim SAVETYPES, SAVETYPEP, SAVETYPEPR, SAVETYPED, SAVETYPEPL, SAVETYPEDE As String
Screen.MousePointer = 11
SAVETYPES = ""
SAVETYPEP = ""
SAVETYPEPR = ""
SAVETYPED = ""
SAVETYPEPL = ""
SAVETYPEDE = ""
If Val(frmLogin.CEDIT(maintab.Tab)) = 0 And Val(frmLogin.CSUPERVISOR(maintab.Tab)) = 0 Then
    msg = MsgBox("Your USER ID does not have sufficient access to perform this task.", 48, "Genesis Information Log")
    Screen.MousePointer = 0
    SAVEERR = 1
    GoTo ExitPoint
End If
Dim dp As String
If (served.Value = 1 And (Not IsDate(servicedate) Or servicetime = "" Or personserved = "" Or locationserved = "")) Then
    msg = MsgBox("Served has been clicked, but either service date, time, person, or location served is blank.", 48, "Genesis Error Log")
    Screen.MousePointer = 0
    SAVEERR = 1
    GoTo ExitPoint
End If
If nonservice.Value = 1 And (nsreason = "" Or Not IsDate(servicedate)) Then
    msg = MsgBox("Non-Service has been clicked, but either service date, or Non-Service Reason is blank.", 48, "Genesis Error Log")
    Screen.MousePointer = 0
    SAVEERR = 1
    GoTo ExitPoint
End If
If assignedon > "" And Not IsDate(assignedon) Then
    msg = MsgBox("ASSIGNED ON date invalid.", 48, "Genesis Error Log")
    assignedon.SetFocus
    Screen.MousePointer = 0
    SAVEERR = 1
    GoTo ExitPoint
End If
If maintab.Tab = 3 Then
    If apptdate <> "" And Not IsDate(apptdate) Then
        msg = MsgBox("APPT. DATE date invalid.", 48, "Genesis Error Log")
        apptdate.SetFocus
        Screen.MousePointer = 0
        SAVEERR = 1
        GoTo ExitPoint
    End If
    If estpayoffdate <> "" And Not IsDate(estpayoffdate) Then
        msg = MsgBox("EST. PAYOFF DATE date invalid.", 48, "Genesis Error Log")
        estpayoffdate.SetFocus
        Screen.MousePointer = 0
        SAVEERR = 1
        GoTo ExitPoint
    End If
    If judgementdate <> "" And Not IsDate(judgementdate) Then
        msg = MsgBox("JUDGEMENT DATE date invalid.", 48, "Genesis Error Log")
        judgementdate.SetFocus
        Screen.MousePointer = 0
        SAVEERR = 1
        GoTo ExitPoint
    End If
    If datesatisfied <> "" And Not IsDate(datesatisfied) Then
        msg = MsgBox("DATE SATISFIED date invalid.", 48, "Genesis Error Log")
        datesatisfied.SetFocus
        Screen.MousePointer = 0
        SAVEERR = 1
        GoTo ExitPoint
    End If
    If intrate > "" And Not IsNumeric(intrate) Then
        msg = MsgBox("INT. RATE is not a valid numeric.", 48, "Genesis Error Log")
        intrate.SetFocus
        Screen.MousePointer = 0
        SAVEERR = 1
        GoTo ExitPoint
    End If
End If
If servicedate <> "" And Not IsDate(servicedate) Then
    msg = MsgBox("SERVICE DATE date invalid.", 48, "Genesis Error Log")
    servicedate.SetFocus
    Screen.MousePointer = 0
    SAVEERR = 1
    GoTo ExitPoint
End If
If served.Value = 1 Then
    If locationserved = "" Then
        msg = MsgBox("A valid Location Served must be entered.", 48, "Genesis Error Log")
        locationserved.SetFocus
        Screen.MousePointer = 0
        SAVEERR = 1
        GoTo ExitPoint
    End If
    If Not IsDate(servicedate) Then
        msg = MsgBox("A valid date must be entered for service.", 48, "Genesis Error Log")
        servicedate.SetFocus
        Screen.MousePointer = 0
        SAVEERR = 1
        GoTo ExitPoint
    End If
    If servicetime = "" And maintab.Tab <> 3 Then
        msg = MsgBox("A valid time must be entered for service.", 48, "Genesis Error Log")
        servicetime.SetFocus
        Screen.MousePointer = 0
        SAVEERR = 1
        GoTo ExitPoint
    End If
End If
If nonservice.Value = 1 Then
    If Not IsDate(servicedate) Then
        msg = MsgBox("A valid date must be entered for non-service.", 48, "Genesis Error Log")
        servicedate.SetFocus
        Screen.MousePointer = 0
        SAVEERR = 1
        GoTo ExitPoint
    End If
End If
If assignedto = "" Then
    assignedto = "UNASSIGNED"
    If Not IsDate(assignedon) Then
        assignedon = Format$(Date$, "mm/dd/yyyy")
    End If
End If
If serviceof = "" Then
    msg = MsgBox("Invalid entry in SERVICE OF field.", 48, "Genesis Error Log")
    serviceof.SetFocus
    Screen.MousePointer = 0
    SAVEERR = 1
    GoTo ExitPoint
End If
If serviceofsort = "" Then
    msg = MsgBox("Invalid entry in SERVICE OF SORT NAME field.", 48, "Genesis Error Log")
    serviceofsort.SetFocus
    Screen.MousePointer = 0
    SAVEERR = 1
    GoTo ExitPoint
End If
If Not IsDate(datereceived) Then
    msg = MsgBox("Invalid entry in DATE RECEIVED field.", 48, "Genesis Error Log")
    datereceived.SetFocus
    Screen.MousePointer = 0
    SAVEERR = 1
    GoTo ExitPoint
End If
If Val(iteration) = 0 Then
    msg = MsgBox("Invalid entry in ITERATION field.", 48, "Genesis Error Log")
    iteration.SetFocus
    Screen.MousePointer = 0
    SAVEERR = 1
    GoTo ExitPoint
End If
If casenumber = "" Then
    msg = MsgBox("CASE NUMBER must be entered.", 48, "Genesis Error Log")
    casenumber.SetFocus
    Screen.MousePointer = 0
    SAVEERR = 1
    GoTo ExitPoint
End If
If courtdate <> "" And Not IsDate(courtdate) Then
    msg = MsgBox("Invalid entry in COURT DATE field.", 48, "Genesis Error Log")
    courtdate.SetFocus
    Screen.MousePointer = 0
    SAVEERR = 1
    GoTo ExitPoint
End If
    
If papertype = "" Then
    msg = MsgBox("PAPER TYPE must be entered.", 48, "Genesis Error Log")
    papertype.SetFocus
    Screen.MousePointer = 0
    SAVEERR = 1
    GoTo ExitPoint
End If
If defendant = "" Then
    msg = MsgBox("DEFENDANT must be entered.", 48, "Genesis Error Log")
    defendant.SetFocus
    Screen.MousePointer = 0
    SAVEERR = 1
    GoTo ExitPoint
End If
If defendantsort = "" Then
    msg = MsgBox("DEFENDANT SORT NAME must be entered.", 48, "Genesis Error Log")
    defendantsort.SetFocus
    Screen.MousePointer = 0
    SAVEERR = 1
    GoTo ExitPoint
End If
If plaintiff = "" Then
    msg = MsgBox("PLAINTIFF must be entered.", 48, "Genesis Error Log")
    plaintiff.SetFocus
    Screen.MousePointer = 0
    SAVEERR = 1
    GoTo ExitPoint
End If
If plaintiffsort = "" Then
    msg = MsgBox("PLAINTIFF SORT must be entered.", 48, "Genesis Error Log")
    plaintiffsort.SetFocus
    Screen.MousePointer = 0
    SAVEERR = 1
    GoTo ExitPoint
End If
If maintab.Tab = 3 Then
    Call commissandint
End If
If maintab.Tab > 3 Then
    GoSub tab1
End If
Styp$ = "C"
If lastserviceof > "" And lastserviceof <> serviceof Then
    msg = MsgBox("The SERVICE OF field has changed.  Do you wish to replace this record, rather than making a copy.", 4, "Genesis Information Log")
    If msg = 6 Then
        Styp$ = "R"
    End If
End If
On maintab.Tab + 1 GoSub tab1, tab2, tab3, tab4
Dim db As Database, ds, rs As Recordset
On Error GoTo oderror1
od1:
Set db = OpenDatabase(nwc + dbname)
If Styp$ = "R" Then
    Set ds = db.OpenRecordset("select * from " + TP + " where serviceof = " + Chr$(34) + lastserviceof + Chr$(34) + " and datereceived = #" + datereceived + "# and iteration = " + Chr$(34) + iteration + Chr$(34))
    If Not ds.EOF Then
        ds.Delete
    End If
End If
Set ds = db.OpenRecordset("select * from " + TP + " where serviceof = " + Chr$(34) + serviceof + Chr$(34) + " and datereceived = #" + datereceived + "# and iteration = " + Chr$(34) + iteration + Chr$(34))
thisisnew% = 0
If ds.EOF Then
    ds.AddNew
    thisisnew% = 1
Else
    ds.MoveFirst
    ds.Edit
End If
If maintab.Tab = 2 Then
    ds("osce") = Left$(servicefee, 10)
    ds("ivd") = ivd.Value
    ds("fl1") = custodian
End If
If maintab.Tab <> 2 Then
    If Val(receiptd) > 0 Then
        ds("receiptnum") = Val(receiptd)
    Else
        ds("receiptnum") = Null
    End If
    If Val(checkd) > 0 Then
        ds("checknum") = Val(checkd)
    Else
        ds("checknum") = Null
    End If
End If
On Error GoTo oderror2
od2:
Set rs = db.OpenRecordset("SELECT STARTRECEIPT FROM SYSTEM")
If Not rs.EOF Then
    rs.MoveFirst
    rs.Edit
    If rs("STARTRECEIPT") <= Val(receiptd) Then
        rs("STARTRECEIPT") = Val(receiptd) + 1
        nextreceipt = Val(receiptd) + 1
        rs.Update
    End If
End If
On Error GoTo oderror1
ds("serviceof") = Left$(serviceof, 60)
ds("datereceived") = datereceived
ds("iteration") = iteration
ds("serviceofsort") = serviceofsort
ds("casenumber") = casenumber
ds("fs2") = armedforces.Value
ds("fs1") = corporate.Value
ds("fl1") = title
ds("papertype") = papertype
If Not IsDate(courtdate) Then
    ds("courtdate") = Null
Else
    ds("courtdate") = courtdate
End If
ds("courttime") = courttime
ds("defendant") = Left$(defendant, 60)
ds("defendantsort") = defendantsort
ds("sohomeaddress") = sohomeaddress
ds("sohomeaddress2") = sohomeaddress2
ds("sohomestate") = sohomestate
ds("sohomezipcode") = sohomezipcode
ds("sohomephone") = sohomephone
ds("soworkaddress") = soworkaddress
ds("soworkaddress2") = soworkaddress2
ds("soworkstate") = soworkstate
ds("soworkzipcode") = soworkzipcode
ds("soworkphone") = soworkphone
ds("daystorespond") = daystorespond
If maintab.Tab <> 2 Then
    ds("servicefee") = Val(servicefee)
    ds("bill") = bill
    If Not IsDate(feedate) Then
        feedate = datereceived
    End If
    ds("feedate") = feedate
Else
    ds("osce") = Left$(servicefee, 10)
End If
ds("defendant") = Left$(defendant, 60)
ds("defendantsort") = defendantsort
ds("dhomeaddress") = dhomeaddress
ds("dhomeaddress2") = dhomeaddress2
ds("dhomestate") = dhomestate
ds("dhomezipcode") = dhomezipcode
ds("dhomephone") = dhomephone
ds("dworkaddress") = dworkaddress
ds("dworkaddress2") = dworkaddress2
ds("dworkstate") = dworkstate
ds("dworkzipcode") = dworkzipcode
ds("dworkphone") = dworkphone
ds("plaintiff") = Left$(plaintiff, 60)
ds("plaintiffsort") = plaintiffsort
ds("phomeaddress") = phomeaddress
ds("phomeaddress2") = phomeaddress2
ds("phomestate") = phomestate
ds("phomezipcode") = phomezipcode
ds("phomephone") = phomephone
ds("pworkaddress") = pworkaddress
ds("pworkaddress2") = pworkaddress2
ds("pworkstate") = pworkstate
ds("pworkzipcode") = pworkzipcode
ds("pworkphone") = pworkphone
ds("assignedto") = Left$(assignedto, 50)
If Not IsDate(assignedon) Then
   ds("assignedon") = Null
Else
   ds("assignedon") = assignedon
End If
ds("served") = served.Value
ds("nonservice") = nonservice.Value
ds("nsreason") = nsreason
ds("premarks") = premarks
If maintab.Tab = 3 Then
    If levy.Text > "" Then
        ds("levy") = levy.Text
    Else
        ds("LEVY") = " "
    End If
End If
ds("wremarks") = wremarks
If Not IsDate(servicedate) Then
    ds("servicedate") = Null
Else
    ds("servicedate") = servicedate
End If
ds("servicetime") = servicetime
ds("personserved") = personserved
ds("locationserved") = locationserved
ds("relationship") = relationship
ds("professional") = Left$(professional, 50)
If maintab.Tab = 3 Then
    Call commissandint
    Screen.MousePointer = 11
        If Not IsDate(apptdate) Then
            ds("apptdate") = Null
        Else
            ds("apptdate") = apptdate
        End If
        If Val(intrate) = 0 Then
            intrate = exintrate
        End If
        ds("intrate") = intrate
        If Not IsDate(datesatisfied) Then
            ds("datesatisfied") = Null
        Else
            ds("datesatisfied") = datesatisfied
        End If
        If Not IsDate(judgementdate) Then
           ds("judgementdate") = Null
        Else
           ds("judgementdate") = judgementdate
        End If
        ds("judgementamount") = Val(judgementamount)
        If Not IsDate(estpayoffdate) Then
           ds("estpayoffdate") = Null
        Else
           ds("estpayoffdate") = estpayoffdate
        End If
        ds("nulla") = nulla.Value
        ds("COMMISSION") = Val(commission)
        ds("INTERest") = Val(INTEREST)
        ds("PERDAY") = Val(perday)
        ds("BALANCE") = Val(balance)
        ds("totalinterest") = totalinterest
        ds("totalcommission") = totalcommission
        ds("totalpayments") = totalpayments
End If
GoSub lines
'CES Code
ds("userfullname") = frmLogin.userfullname
ds("userid") = frmLogin.userid
ds("ORINUMBER") = frmLogin.orinumber
ds("udate") = Format$(Now, "mm/dd/yyyy")
ds("utime") = Format$(Now, "hh:mm:ss")
'********
ds.Update
If Val(receiptd) > 0 Then
    On Error GoTo oderror3
od3:
    Set rs = db.OpenRecordset("select * from receipt WHERE ITERATION = " + Chr$(34) + iteration + Chr$(34) + " AND SERVICEOF = " + Chr$(34) + serviceof + Chr$(34) + " AND DATERECEIVED = #" + Format$(datereceived, "mm/dd/yyyy") + "# AND RECEIPTNUM = " + receiptd)
    If rs.EOF Then
        rs.AddNew
    Else
        rs.MoveFirst
        rs.Edit
    End If
    rs("iteration") = iteration
    rs("serviceof") = Left$(serviceof, 60)
    rs("datereceived") = feedate
    rs("casenumber") = casenumber
    rs("papertype") = papertype
    rs("servicefee") = Val(servicefee)
    rs("receiptnumber") = receiptd
    If checkd = "" Then
        rs("checknum") = " "
    Else
        rs("checknum") = checkd
    End If
    rs("receiptnum") = Val(receiptd)
    rs("datereceipt") = Format$(Date$, "mm/dd/yyyy")
    rs("defendant") = Left$(defendant, 60)
    rs("plaintiff") = Left$(plaintiff, 60)
    rs.Update
    lastserviceof = ""
End If
If maintab.Tab = 3 Then
        On Error GoTo oderror4
od4:
        Set ds = db.OpenRecordset("select * from executionspay where serviceof = " + Chr$(34) + serviceof + Chr$(34) + " and datereceived = #" + datereceived + "# and iteration = " + Chr$(34) + iteration + Chr$(34))
        If Not ds.EOF Then
           ds.MoveFirst
           While Not ds.EOF
               ds.Delete
               ds.MoveNext
           Wend
        End If
        For t% = 0 To expaygrid.Rows
            On Error GoTo ER
            expaygrid.Row = t%
            expaygrid.Col = 0
            dp$ = expaygrid.Text
            expaygrid.Col = 1
            a$ = expaygrid.Text
            expaygrid.Col = 7
            pr$ = expaygrid.Text
            pr$ = Left$(pr$, 10)
            If dp$ > "" And a$ > "" Then
                Set ds = db.OpenRecordset("select * from executionspay where serviceof = " + Chr$(34) + serviceof + Chr$(34) + " and datereceived = #" + datereceived + "# and iteration = " + Chr$(34) + iteration + Chr$(34) + " and datepaid = #" + dp$ + "# and amount = " + a$ + " and payrem = " + Chr$(34) + pr$ + Chr$(34))
                ds.AddNew
                ds("serviceof") = Left$(serviceof, 60)
                ds("datereceived") = datereceived
                ds("iteration") = iteration
                expaygrid.Col = 0
                ds("datepaid") = expaygrid.Text
                dp = expaygrid.Text
                expaygrid.Col = 1
                ds("amount") = expaygrid.Text
                amt = expaygrid.Text
                expaygrid.Col = 2
                ds("receipt") = expaygrid.Text
                rcnum = expaygrid.Text
                ds("RECEIPTNUM") = Val(ds("RECEIPT"))
                expaygrid.Col = 7
                ds("servicefee") = Val(expaygrid.Text)
                On Error GoTo oderror5
od5:
                Set rs = db.OpenRecordset("SELECT STARTRECEIPT FROM SYSTEM")
                If Not rs.EOF Then
                    rs.MoveFirst
                    rs.Edit
                    If rs("STARTRECEIPT") <= Val(ds("receipt")) Then
                        rs("STARTRECEIPT") = Val(ds("receipt")) + 1
                        nextreceipt = Val(receiptd) + 1
                        rs.Update
                    End If
                End If
                On Error GoTo oderror4
                expaygrid.Col = 3
                ds("check") = Val(expaygrid.Text)
                cknum = expaygrid.Text
                expaygrid.Col = 4
                ds("principal") = Val(expaygrid.Text)
                expaygrid.Col = 5
                ds("commiss") = Val(expaygrid.Text)
                expaygrid.Col = 6
                ds("inter") = Val(expaygrid.Text)
                expaygrid.Col = 7
                ds("payrem") = Val(expaygrid.Text)
                expaygrid.Col = 8
                ds("payrem") = Left$(expaygrid.Text, 10)
                'CES Code
                ds("userfullname") = frmLogin.userfullname
                ds("userid") = frmLogin.userid
                ds("ORINUMBER") = frmLogin.orinumber
                ds("udate") = Format$(Now, "mm/dd/yyyy")
                ds("utime") = Format$(Now, "hh:mm:ss")
                ds.Update
                
                If Val(rcnum) > 0 Then
                    Set rs = db.OpenRecordset("select * from receipt WHERE ITERATION = " + Chr$(34) + iteration + Chr$(34) + " AND SERVICEOF = " + Chr$(34) + serviceof + Chr$(34) + " AND DATERECEIVED = #" + Format$(dp, "mm/dd/yyyy") + "# AND RECEIPTNUM = " + rcnum)
                    If rs.EOF Then
                        rs.AddNew
                    Else
                        rs.MoveFirst
                        rs.Edit
                    End If
                    rs("iteration") = iteration
                    rs("serviceof") = Left$(serviceof, 60)
                    rs("datereceived") = dp
                    rs("casenumber") = casenumber
                    rs("papertype") = papertype
                    rs("servicefee") = Val(amt)
                    rs("receiptnumber") = rcnum
                    If cknum = "" Then
                        rs("checknum") = " "
                    Else
                        rs("checknum") = cknum
                    End If
                    rs("receiptnum") = Val(rcnum)
                    rs("datereceipt") = Format$(dp, "mm/dd/yyyy")
                    rs("defendant") = Left$(defendant, 60)
                    rs("plaintiff") = Left$(plaintiff, 60)
                    rs.Update
                End If
                    
            End If
LOOPIT:
        Next t%
End If

  
On Error GoTo oderror6
od6:
Set ds = db.OpenRecordset("select papertype from papers where papertype = " + Chr$(34) + papertype + Chr$(34))
If ds.EOF Then
    SAVETYPEP = "N"
    ds.AddNew
    ds("papertype") = papertype
    ds.Update
End If
db.Close
On Error GoTo oderror7
od7:
Set db = OpenDatabase(nwl + "LAWSUITE.MDB")
If professional > "" Then
    On Error GoTo oderror8
od8:
    Set ds = db.OpenRecordset("select profname,type from professionals where profname = " + Chr$(34) + professional + Chr$(34) + " and type = " + Chr$(34) + typ$ + Chr$(34))
    If ds.EOF Then
        SAVETYPEPR = "N"
        ds.AddNew
        ds("profname") = Left$(professional, 50)
        ds("type") = typ$
        ds.Update
    End If
End If
On Error GoTo oderror9
od9:
Set ds = db.OpenRecordset("select * from PEOPLE where dpname = " + Chr$(34) + serviceof + Chr$(34))
If ds.EOF Then
    SAVETYPES = "N"
   ds.AddNew
Else
    SAVETYPES = "O"
    ds.MoveFirst
    ds.Edit
End If
ds("dpname") = Left$(serviceof, 60)
ds("dpsort") = serviceofsort
ds("dphaddress") = sohomeaddress
ds("dphaddress2") = sohomeaddress2
ds("hstate") = sohomestate
ds("hzipcode") = sohomezipcode
ds("dphphone") = sohomephone
ds("dpwaddress") = soworkaddress
ds("dpwaddress2") = soworkaddress2
ds("wstate") = soworkstate
ds("wzipcode") = soworkzipcode

ds("dpwphone") = soworkphone
LF$ = Left$(serviceof, 60)
GoSub SETLF
ds("DPNAMELF") = Left$(LF$, 60)
ds.Update
On Error GoTo oderror10
od10:
Set ds = db.OpenRecordset("select * from PEOPLE where dpname = " + Chr$(34) + plaintiff + Chr$(34))
If ds.EOF Then
    SAVETYPEPL = "N"
   ds.AddNew
Else
    SAVETYPEPL = "O"
    ds.MoveFirst
    ds.Edit
End If
ds("dpname") = Left$(plaintiff, 60)
ds("dpsort") = plaintiffsort
ds("dphaddress") = phomeaddress
ds("dphaddress2") = phomeaddress2
ds("hstate") = phomestate
ds("hzipcode") = phomezipcode
ds("dphphone") = phomephone
ds("dpwaddress") = pworkaddress
ds("dpwaddress2") = pworkaddress2
ds("wstate") = pworkstate
ds("wzipcode") = pworkzipcode
ds("dpwphone") = pworkphone
LF$ = Left$(plaintiff, 60)
GoSub SETLF
ds("DPNAMELF") = Left$(LF$, 60)
ds.Update
On Error GoTo oderror11
od11:
Set ds = db.OpenRecordset("select * from PEOPLE where dpname = " + Chr$(34) + defendant + Chr$(34))
If ds.EOF Then
    SAVETYPEDE = "N"
   ds.AddNew
Else
    SAVETYPEDE = "O"
    ds.MoveFirst
    ds.Edit
End If
ds("dpname") = Left$(defendant, 60)
ds("dpsort") = defendantsort
ds("dphaddress") = dhomeaddress
ds("dphaddress2") = dhomeaddress2
ds("hstate") = dhomestate
ds("hzipcode") = dhomezipcode
ds("dphphone") = dhomephone
ds("dpwaddress") = dworkaddress
ds("dpwaddress2") = dworkaddress2
ds("wstate") = dworkstate
ds("wzipcode") = dworkzipcode
ds("dpwphone") = dworkphone
LF$ = Left$(defendant, 60)
GoSub SETLF
ds("DPNAMELF") = Left$(LF$, 60)
ds.Update
If assignedto > "" Then
    On Error GoTo oderror12
od12:
    Set ds = db.OpenRecordset("select profname,type from professionals where profname = " + Chr$(34) + assignedto + Chr$(34) + " and type = 'D'")
    If ds.EOF Then
        SAVETYPEDE = "N"
        ds.AddNew
        ds("profname") = Left$(assignedto, 50)
        ds("type") = "D"
        ds.Update
    End If
End If
On Error GoTo 0
Call clearit
If SAVETYPEP = "N" Then
    papertype.AddItem papertype
End If
If SAVETYPEPR = "N" Then
    professional.AddItem Left$(professional, 50)
End If
If SAVETYPEDE = "N" Then
    assignedto.AddItem Left$(assignedto, 50)
End If
Data1.Refresh
serviceof.Refresh
defendant.Refresh
plaintiff.Refresh
If autoprint.Value = 1 And thisisnew% = 1 And FROMG = 0 Then
    FROMP = 0
    Y$ = Right$(Format$(datereceived, "mmddyyyy"), 4)
    m$ = Left$(Format$(datereceived, "mmddyyyy"), 2)
    d$ = Mid$(Format$(datereceived, "mmddyyyy"), 3, 2)
    If maintab.Tab = 0 Then
        report.SelectionFormula = "{magistrate.serviceof} = " + Chr$(34) + serviceof + Chr$(34) + " and {magistrate.datereceived} = date(" + Y$ + "," + m$ + "," + d$ + ") AND {MAGISTRATE.ITERATION} = " + Chr$(34) + iteration + Chr$(34)
    End If
    If maintab.Tab = 1 Then
        report.SelectionFormula = "{writother.serviceof} = " + Chr$(34) + serviceof + Chr$(34) + " and {writother.datereceived} = date(" + Y$ + "," + m$ + "," + d$ + ") AND {WRITOTHER.ITERATION} = " + Chr$(34) + iteration + Chr$(34)
    End If
    If maintab.Tab = 2 Then
        report.SelectionFormula = "{familycourt.serviceof} = " + Chr$(34) + serviceof + Chr$(34) + " and {familycourt.datereceived} = date(" + Y$ + "," + m$ + "," + d$ + ") AND {FAMILYCOURT.ITERATION} = " + Chr$(34) + iteration + Chr$(34)
    End If
    If maintab.Tab = 3 Then
        report.SelectionFormula = "{executions.serviceof} = " + Chr$(34) + serviceof + Chr$(34) + " and {executions.datereceived} = date(" + Y$ + "," + m$ + "," + d$ + ") AND {EXECUTIONS.ITERATION} = " + Chr$(34) + iteration + Chr$(34)
    End If
    If List1.ListCount > 1 And List1.ListIndex > -1 Then
        Call defaultprinter(List1.List(List1.ListIndex))
    End If
    If List1.ListIndex > -1 And List2.ListIndex > -1 Then
        If List1.List(List1.ListIndex) <> List2.List(List2.ListIndex) And prepareprinter = 1 Then
            msg = MsgBox("Prepare for Worksheet/Affidavit/Letter/Report Printing.", 48, "Genesis Error Log")
        End If
    End If
    If Val(frmLogin.CPRINT(maintab.Tab)) = 1 Or Val(frmLogin.CSUPERVISOR(maintab.Tab)) = 1 Then
    If maintab.Tab = 0 Then
        report.ReportFileName = nwc + "mworksht.rpt"
        report.Destination = crptToPrinter
        report.CopiesToPrinter = 1
        report.Action = 1
    End If
    If maintab.Tab = 1 Then
        report.Destination = crptToPrinter
        report.CopiesToPrinter = 1
        report.ReportFileName = nwc + "wworksht.rpt"
        report.Action = 1
    End If
    If maintab.Tab = 2 Then
        report.Destination = crptToPrinter
        report.CopiesToPrinter = 1
        report.ReportFileName = nwc + "fworksht.rpt"
        report.Action = 1
    End If
    If maintab.Tab = 3 Then
        report.Destination = crptToPrinter
        report.CopiesToPrinter = 1
        report.ReportFileName = nwc + "eworksht.rpt"
        report.Action = 1
    End If
    End If
End If
If autoprint.Value = 1 And thisisnew% = 1 And maintab.Tab <> 2 And Val(servicefee) > 0 And FROMG = 0 Then
    If Val(frmLogin.CPRINT(maintab.Tab)) = 1 Or Val(frmLogin.CSUPERVISOR(maintab.Tab)) = 1 Then
    fee = Val(servicefee)
    If List2.ListCount > 1 And List2.ListIndex > -1 Then
        Call defaultprinter(List2.List(List2.ListIndex))
    End If
    If List1.ListIndex > -1 And List2.ListIndex > -1 Then
        If List1.List(List1.ListIndex) <> List2.List(List2.ListIndex) And prepareprinter = 1 Then
            msg = MsgBox("Prepare for Receipt/Check Printing.", 48, "Genesis Error Log")
        End If
    End If
    rnumber = receiptd
    cnumber = checkd
    receiptframe.Left = 1000
    receiptframe.Top = 2000
    Call LOADOTHER
    If fromdefendant = 0 And fromplaintiff = 0 And othername = "" Then
        fromdefendant = 1
    End If
    receiptframe.Visible = True
    Screen.MousePointer = 0
    othername.SetFocus
    End If
End If
'RLB
If blnSavePressed Then
    MsgBox "'" & RemoveSpecChars(maintab.TabCaption(maintab.Tab)) & "': Paper Saved Successfully", vbOKOnly, "Genesis Information Log"
End If
'*****
If autoprint = 0 Or (autoprint = 1 And thisisnew% = 0) Or Val(servicefee) = 0 Or maintab.Tab = 2 Then
    If FROMG = 0 Then
        Call clearbutton_Click
    End If
End If
Screen.MousePointer = 0
db.Close
If FROMG = 0 Then
    If SEARCHTYPE = 1 Then
        SEARCHTYPE = 0
        maintab.Tab = 5
    Else
        SEARCHTYPE = 0
    End If
End If
ExitPoint:
blnSavePressed = False
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
oderror3:
a = Error$(Err)
If Err > 3200 Then
    Resume od3
Else
    Resume Next
End If
oderror4:
If Err > 3200 Then
    Resume od4
Else
    Resume Next
    Resume
End If
oderror5:
If Err > 3200 Then
    Resume od5
Else
    Resume Next
End If
oderror6:
If Err > 3200 Then
    Resume od6
Else
    Resume Next
End If
oderror7:
If Err > 3200 Then
    Resume od7
Else
    Resume Next
End If
oderror8:
If Err > 3200 Then
    Resume od8
Else
    Resume Next
End If
oderror9:
If Err > 3200 Then
    Resume od9
Else
    Resume Next
End If
oderror10:
If Err > 3200 Then
    Resume od10
Else
    Resume Next
End If
oderror11:
If Err > 3200 Then
    Resume od11
Else
    Resume Next
End If
oderror12:
If Err > 3200 Then
    Resume od12
Else
    Resume Next
End If
ER:
Resume LOOPIT
peopleerr:
msg = MsgBox("Unable to add People information because PEOPLE table has reached its maximum size.", 48, "Genesis Error Log")
Resume Next
tab1:
TP = "magistrate"
typ$ = "M"
Return
tab2:
TP = "writother"
typ$ = "A"
Return
tab3:
TP = "familycourt"
typ$ = "C"
Return
tab4:
TP = "executions"
typ$ = "A"
Call commissandint
Return
lines:
Printer.FontBold = False
Printer.FontUnderline = False
Printer.FontItalic = False
Printer.FontName = "Times New Roman"
Printer.FontSize = 10
holdwidth = Printer.Width - 1000
If (nonservice.Value = 0 And served.Value = 0) Or (served.Value = 1 And (Not IsDate(servicedate) Or servicetime = "" Or personserved = "" Or locationserved = "")) Then
    ds("line1") = ""
    ds("line2") = ""
    ds("line3") = ""
    ds("line4") = ""
    ds("line5") = ""
    ds("line6") = ""
    ds("line7") = ""
    Return
End If
If nonservice.Value = 1 And (nsreason = "" Or Not IsDate(servicedate)) Then
    ds("line1") = ""
    ds("line2") = ""
    ds("line3") = ""
    ds("line4") = ""
    ds("line5") = ""
    ds("line6") = ""
    ds("line7") = ""
    Return
End If
If Day(servicedate) = 1 Or Day(servicedate) = 21 Or Day(servicedate) = 31 Then
    dp = Mid$(Str$(Day(servicedate)), 2) + "st"
Else
If Day(servicedate) = 2 Or Day(servicedate) = 22 Then
    dp = Mid$(Str$(Day(servicedate)), 2) + "nd"
Else
If Day(servicedate) = 3 Or Day(servicedate) = 23 Then
    dp = Mid$(Str$(Day(servicedate)), 2) + "rd"
Else
    dp = Mid$(Str$(Day(servicedate)), 2) + "th"
End If
End If
End If
If served.Value = 1 Then
    If relationship > "" Then
        temp$ = "PERSONALLY comes the undersigned, who says on oath that on the " + dp + " day of " + Format$(servicedate, "mmmm") + ", " + Right$(Format$(servicedate, "mmddyyyy"), 4) + ", at " + servicetime + ", he/she served the " + papertype + " on " + serviceof + " by delivering unto " + personserved + ", " + relationship + ",  at " + locationserved + " personally copy(ies) thereof.  Service of process was made in accordance with applicable statutes and the Rules of Civil Procedures in effect at the time of service.  " + premarks
    Else
        temp$ = "PERSONALLY comes the undersigned, who says on oath that on the " + dp + " day of " + Format$(servicedate, "mmmm") + ", " + Right$(Format$(servicedate, "mmddyyyy"), 4) + ", at " + servicetime + ", he/she served the " + papertype + " on " + serviceof + " by delivering unto " + personserved + "  at " + locationserved + " personally copy(ies) thereof.  Service of process was made in accordance with applicable statutes and the Rules of Civil Procedures in effect at the time of service.  " + premarks
    End If
Else
    temp$ = "PERSONALLY comes the undersigned, who says on oath that as of the " + dp + " day of " + Format$(servicedate, "mmmm") + ", " + Right$(Format$(servicedate, "mmddyyyy"), 4) + ", after several attempts to serve the above " + papertype + " on " + serviceof + ", he/she was unable to complete service in accordance with applicable statutes and the Rules of Civil Procedures in effect.  Service could not be completed for the following reasons: " + nsreason
End If
TEMP2$ = temp$
Printer.FontBold = False
Printer.FontUnderline = False
Printer.FontItalic = False
Printer.FontName = "Times New Roman"
Printer.FontSize = 10
If holdwidth < Printer.TextWidth(TEMP2$) Then
    TEMP2$ = ""
    While Printer.TextWidth(TEMP2$) <= holdwidth
        FOUNDSPACE% = 0
        For t% = Len(TEMP2$) + 1 To Len(temp$)
            If Mid$(temp$, t%, 1) = " " Then
                FOUNDSPACE% = t%
                t% = Len(temp$)
            End If
        Next t%
        lasttemp2$ = TEMP2$
        If FOUNDSPACE% = 0 Then
            FOUNDSPACE% = Len(temp$)
        End If
        TEMP2$ = TEMP2$ + Mid$(temp$, Len(TEMP2$) + 1, FOUNDSPACE% - Len(TEMP2$))
        Printer.FontBold = False
        Printer.FontUnderline = False
        Printer.FontItalic = False
        Printer.FontName = "Times New Roman"
        Printer.FontSize = 8
        Printer.FontSize = 10
    Wend
    TEMP2$ = lasttemp2$
End If
temp$ = Mid$(temp$, Len(TEMP2$) + 1)
If Left$(TEMP2$, 1) = " " Then
    TEMP2$ = Mid$(TEMP2$, 2)
End If
ds("line1") = TEMP2$
TEMP2$ = temp$
If holdwidth < Printer.TextWidth(TEMP2$) Then
    TEMP2$ = ""
    While Printer.TextWidth(TEMP2$) <= holdwidth
        FOUNDSPACE% = 0
        For t% = Len(TEMP2$) + 1 To Len(temp$)
            If Mid$(temp$, t%, 1) = " " Then
                FOUNDSPACE% = t%
                t% = Len(temp$)
            End If
        Next t%
        If FOUNDSPACE% = 0 Then
            FOUNDSPACE% = Len(temp$)
        End If
        lasttemp2$ = TEMP2$
        TEMP2$ = TEMP2$ + Mid$(temp$, Len(TEMP2$) + 1, FOUNDSPACE% - Len(TEMP2$))
    Wend
    TEMP2$ = lasttemp2$
End If
temp$ = Mid$(temp$, Len(TEMP2$) + 1)
If Left$(TEMP2$, 1) = " " Then
    TEMP2$ = Mid$(TEMP2$, 2)
End If
ds("line2") = TEMP2$
TEMP2$ = temp$
If holdwidth < Printer.TextWidth(TEMP2$) Then
    TEMP2$ = ""
    While Printer.TextWidth(TEMP2$) <= holdwidth
        FOUNDSPACE% = 0
        For t% = Len(TEMP2$) + 1 To Len(temp$)
            If Mid$(temp$, t%, 1) = " " Then
                FOUNDSPACE% = t%
                t% = Len(temp$)
            End If
        Next t%
        If FOUNDSPACE% = 0 Then
            FOUNDSPACE% = Len(temp$)
        End If
        lasttemp2$ = TEMP2$
        TEMP2$ = TEMP2$ + Mid$(temp$, Len(TEMP2$) + 1, FOUNDSPACE% - Len(TEMP2$))
    Wend
    TEMP2$ = lasttemp2$
End If
temp$ = Mid$(temp$, Len(TEMP2$) + 1)
ds("line3") = TEMP2$
TEMP2$ = temp$
If holdwidth < Printer.TextWidth(TEMP2$) Then
    TEMP2$ = ""
    While Printer.TextWidth(TEMP2$) <= holdwidth
        FOUNDSPACE% = 0
        For t% = Len(TEMP2$) + 1 To Len(temp$)
            If Mid$(temp$, t%, 1) = " " Then
                FOUNDSPACE% = t%
                t% = Len(temp$)
            End If
        Next t%
        If FOUNDSPACE% = 0 Then
            FOUNDSPACE% = Len(temp$)
        End If
        lasttemp2$ = TEMP2$
        TEMP2$ = TEMP2$ + Mid$(temp$, Len(TEMP2$) + 1, FOUNDSPACE% - Len(TEMP2$))
    Wend
    TEMP2$ = lasttemp2$
End If
temp$ = Mid$(temp$, Len(TEMP2$) + 1)
ds("line4") = TEMP2$
TEMP2$ = temp$
If holdwidth < Printer.TextWidth(TEMP2$) Then
    TEMP2$ = ""
    While Printer.TextWidth(TEMP2$) <= holdwidth
        FOUNDSPACE% = 0
        For t% = Len(TEMP2$) + 1 To Len(temp$)
            If Mid$(temp$, t%, 1) = " " Then
                FOUNDSPACE% = t%
                t% = Len(temp$)
            End If
        Next t%
        If FOUNDSPACE% = 0 Then
            FOUNDSPACE% = Len(temp$)
        End If
        lasttemp2$ = TEMP2$
        TEMP2$ = TEMP2$ + Mid$(temp$, Len(TEMP2$) + 1, FOUNDSPACE% - Len(TEMP2$))
    Wend
    TEMP2$ = lasttemp2$
End If
temp$ = Mid$(temp$, Len(TEMP2$) + 1)
ds("line5") = TEMP2$
TEMP2$ = temp$
If holdwidth < Printer.TextWidth(TEMP2$) Then
    TEMP2$ = ""
    While Printer.TextWidth(TEMP2$) <= holdwidth
        FOUNDSPACE% = 0
        For t% = Len(TEMP2$) + 1 To Len(temp$)
            If Mid$(temp$, t%, 1) = " " Then
                FOUNDSPACE% = t%
                t% = Len(temp$)
            End If
        Next t%
        If FOUNDSPACE% = 0 Then
            FOUNDSPACE% = Len(temp$)
        End If
        lasttemp2$ = TEMP2$
        TEMP2$ = TEMP2$ + Mid$(temp$, Len(TEMP2$) + 1, FOUNDSPACE% - Len(TEMP2$))
    Wend
    TEMP2$ = lasttemp2$
End If
temp$ = Mid$(temp$, Len(TEMP2$) + 1)
ds("line6") = TEMP2$
TEMP2$ = temp$
If holdwidth < Printer.TextWidth(TEMP2$) Then
    TEMP2$ = ""
    While Printer.TextWidth(TEMP2$) <= holdwidth
        FOUNDSPACE% = 0
        For t% = Len(TEMP2$) + 1 To Len(temp$)
            If Mid$(temp$, t%, 1) = " " Then
                FOUNDSPACE% = t%
                t% = Len(temp$)
            End If
        Next t%
        If FOUNDSPACE% = 0 Then
            FOUNDSPACE% = Len(temp$)
        End If
        lasttemp2$ = TEMP2$
        TEMP2$ = TEMP2$ + Mid$(temp$, Len(TEMP2$) + 1, FOUNDSPACE% - Len(TEMP2$))
    Wend
    TEMP2$ = lasttemp2$
End If
temp$ = Mid$(temp$, Len(TEMP2$) + 1)
ds("line7") = TEMP2$
Return
SETLF:
hoLdname = LF$
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
If Right$(tso$, 1) = " " Then
    tso$ = Left$(tso$, Len(tso$) - 1)
End If
firstspace% = 0
While Right$(tso$, 1) = " " And Len(tso$) > 1
    tso$ = Left$(tso$, Len(tso$) - 1)
Wend
For tt% = Len(tso$) To 1 Step -1
    If Mid$(tso$, tt%, 1) = " " Then
        If Mid$(tso$, tt% + 1, 3) = "JR." Or Mid$(tso$, tt% + 1, 3) = "SR." Or Mid$(tso$, tt% + 1, 3) = "III" Or Mid$(tso$, tt% + 1, 2) = "IV" Then
            aa = 1
        Else
            firstspace% = tt%
            tt% = 1
        End If
    End If
Next tt%
If firstspace% = 0 Then
    If osort1$ = "" Then
        osort1$ = tso$
    End If
    GoTo GGO
End If
tempsort$ = Mid$(tso$, firstspace% + 1)
If Left$(tempsort$, 1) = " " Then
    tempsort$ = Mid$(tempsort$, 2)
End If
tso$ = Left$(tso$, firstspace% - 1)
If Right$(tso$, 1) = " " Then
    tso$ = Left$(tso$, Len(tso$) - 1)
End If
tempsort$ = tempsort$ + ", " + tso$
If osort1$ = "" Then
    osort1$ = tempsort$
End If
'If InStr(osort1$, "JR.") Then
'    If Mid$(osort1$, InStr(osort1$, "JR.") + 3, 1) = " " Then
'        osort1$ = Left$(osort1$, InStr(osort1$, "JR.") - 1) + Mid$(osort1$, InStr(osort1$, "JR.") + 4) + ", JR."
'    Else
'        osort1$ = Left$(osort1$, InStr(osort1$, "JR.") - 1) + Mid$(osort1$, InStr(osort1$, "JR.") + 3) + ", JR."
'End If
'End If
'If InStr(osort1$, "SR.") Then
'    If Mid$(osort1$, InStr(osort1$, "SR.") + 3, 1) = " " Then
'        osort1$ = Left$(osort1$, InStr(osort1$, "SR.") - 1) + Mid$(osort1$, InStr(osort1$, "SR.") + 4) + ", SR."
'    Else
'        osort1$ = Left$(osort1$, InStr(osort1$, "SR.") - 1) + Mid$(osort1$, InStr(osort1$, "SR.") + 3) + ", SR."
'End If
'End If
'If InStr(osort1$, "III") Then
'    If Mid$(osort1$, InStr(osort1$, "III") + 3, 1) = " " Then
'        osort1$ = Left$(osort1$, InStr(osort1$, "III") - 1) + Mid$(osort1$, InStr(osort1$, "III") + 4) + ", III"
'    Else
'        osort1$ = Left$(osort1$, InStr(osort1$, "III") - 1) + Mid$(osort1$, InStr(osort1$, "III") + 3) + ", III"
'    End If
'End If
'If InStr(osort1$, "IV") Then
'    If Mid$(osort1$, InStr(osort1$, "IV") + 2, 1) = " " Then
'        osort1$ = Left$(osort1$, InStr(osort1$, "IV") - 1) + Mid$(osort1$, InStr(osort1$, "IV") + 3) + ", III"
'    Else
'        osort1$ = Left$(osort1$, InStr(osort1$, "IV") - 1) + Mid$(osort1$, InStr(osort1$, "IV") + 2) + ", III"
'    End If
'End If
If Left$(osort1$, 1) = " " Then
    osort1$ = Mid$(osort1$, 2)
End If
GGO:
LF$ = osort1$
Return
End Sub


Private Sub serviceof_KeyPress(KeyAscii As Integer)
CSERVICEOF = 1
If KeyAscii = 34 Then
    msg = MsgBox("The " + Chr$(34) + " character is not allowed.  Use the ' character instead.", 48, "Genesis Error Log")
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Or KeyAscii = 9 Then
    If serviceof > "" Then
        Call serviceof_Click(0)
    End If
End If
If KeyAscii = 13 Then
    datereceived.SetFocus
End If
End Sub

Private Sub professional_LostFocus()
If maintab.Tab > 3 Then
    GoSub tab1
End If
On maintab.Tab + 1 GoSub tab1, tab2, tab3, tab4
If Len(professional) > 50 Then
    msg = MsgBox("Maximum length of 50 has been exceeded for " + pr$ + " entry.  This entry will be truncated.", 48, "Genesis Error Log")
    professional = Left$(professional, 50)
End If
Exit Sub
tab1:
pr$ = "MAGISTRATE"
Return
tab2:
pr$ = "ATTORNEY"
Return
tab3:
pr$ = "MAGISTRATE"
Return
tab4:
pr$ = "MAGISTRATE"
Return
End Sub

Private Sub clearbutton_Click()
mugshot.Picture = LoadPicture()
serviceof = ""
datereceived = ""
daterlist.clear
iteration = ""
serviceofsort = ""
casenumber = ""
corporate.Value = 0
armedforces.Value = 0
title = ""
papertype = ""
courtdate = ""
courttime = ""
defendant = ""
defendantsort = ""
sohomeaddress = ""
sohomeaddress2 = ""
sohomestate = ""
sohomezipcode = ""
sohomephone = ""
soworkaddress = ""
soworkaddress2 = ""
soworkstate = ""
soworkzipcode = ""
soworkphone = ""
daystorespond = ""
servicefee = ""
bill = 0
feedate = ""
receiptd = ""
checkd = ""
defendant = ""
defendantsort = ""
dhomeaddress = ""
dhomeaddress2 = ""
dhomestate = ""
dhomezipcode = ""
dhomephone = ""
dworkaddress = ""
dworkaddress2 = ""
dworkstate = ""
dworkzipcode = ""
dworkphone = ""
plaintiff = ""
plaintiffsort = ""
phomeaddress = ""
phomeaddress2 = ""
phomestate = ""
phomezipcode = ""
phomephone = ""
pworkaddress = ""
pworkaddress2 = ""
pworkstate = ""
pworkzipcode = ""
pworkphone = ""
assignedto = ""
assignedon = ""
served.Value = 0
nonservice.Value = 0
nsreason = ""
premarks = ""
levy.Text = ""
wremarks = ""
servicedate = ""
servicetime = ""
personserved = ""
locationserved = ""
relationship = ""
professional = ""
ivd.Value = 0
custodian = ""
apptdate = ""
intrate = ""
datesatisfied = ""
judgementdate = ""
judgementamount = ""
estpayoffdate = ""
nulla.Value = 0
commission = ""
perday = ""
balance = ""
expaygrid.Rows = 1
expaygrid.Row = 0
For t% = 0 To 8
    expaygrid.Col = t%
    expaygrid.Text = ""
Next t%
DATEPAID = ""
amount = ""
receipt = ""
check = ""
remarks = ""
eservicefee = ""
INTEREST = ""
commiss = ""
total = ""
commission = ""
principal = ""
lastserviceof = ""
serviceof.SetFocus
End Sub

Private Sub holdlast(PreviousTab As Integer)
If PreviousTab = 0 Then
        Open "holdm" For Output As #1
End If
If PreviousTab = 1 Then
        Open "holdw" For Output As #1
End If
If PreviousTab = 2 Then
        Open "holdf" For Output As #1
End If
If PreviousTab = 3 Then
        Open "holde" For Output As #1
End If
If PreviousTab > 3 Then
    Exit Sub
End If
Print #1, serviceof
Print #1, datereceived
Print #1, iteration
Print #1, serviceofsort
Print #1, casenumber
Print #1, armedforces.Value
Print #1, corporate.Value
Print #1, title
Print #1, papertype
Print #1, courtdate
Print #1, courttime
Print #1, defendant
Print #1, defendantsort
Print #1, sohomeaddress
Print #1, sohomeaddress2
Print #1, sohomestate
Print #1, sohomezipcode
Print #1, sohomephone
Print #1, soworkaddress
Print #1, soworkaddress2
Print #1, soworkstate
Print #1, soworkzipcode
Print #1, soworkphone
Print #1, daystorespond
Print #1, servicefee
Print #1, bill
Print #1, feedate
Print #1, receiptd
Print #1, checkd
Print #1, defendant
Print #1, defendantsort
Print #1, dhomeaddress
Print #1, dhomeaddress2
Print #1, dhomestate
Print #1, dhomezipcode
Print #1, dhomephone
Print #1, dworkaddress
Print #1, dworkaddress2
Print #1, dworkstate
Print #1, dworkzipcode
Print #1, dworkphone
Print #1, plaintiff
Print #1, plaintiffsort
Print #1, phomeaddress
Print #1, phomeaddress2
Print #1, phomestate
Print #1, phomezipcode
Print #1, phomephone
Print #1, pworkaddress
Print #1, pworkaddress2
Print #1, pworkstate
Print #1, pworkzipcode
Print #1, pworkphone
Print #1, assignedto
Print #1, assignedon
Print #1, served.Value
Print #1, nonservice.Value
Print #1, nsreason
Print #1, premarks
Print #1, levy.Text
Print #1, wremarks
Print #1, servicedate
Print #1, servicetime
Print #1, personserved
Print #1, locationserved
Print #1, relationship
Print #1, professional
If PreviousTab = 2 Then
        Print #1, ivd.Value
        Print #1, custodian
End If
If PreviousTab = 3 Then
        Print #1, apptdate
        Print #1, intrate
        Print #1, datesatisfied
        Print #1, judgementdate
        Print #1, judgementamount
        Print #1, estpayoffdate
        Print #1, nulla.Value
        Print #1, INTEREST
        Print #1, commission
        Print #1, perday
        Print #1, balance
       
        For t% = 1 To expaygrid.Rows
            expaygrid.Row = t% - 1
            expaygrid.Col = 0
            If expaygrid.Text > "" Then
                For tt% = 0 To 8
                    expaygrid.Col = tt%
                    Print #1, expaygrid.Text
                Next tt%
            End If
        Next t%
End If
Close #1
End Sub

Private Sub floodlast(thistab As Integer)
On Error GoTo SKIPIT
If thistab = 0 Then
        Open "holdm" For Input As #1
End If
If thistab = 1 Then
        Open "holdw" For Input As #1
End If
If thistab = 2 Then
        Open "holdf" For Input As #1
End If
If thistab = 3 Then
        Open "holde" For Input As #1
End If
If thistab > 3 Then
    Exit Sub
End If
Line Input #1, a$
serviceof = a$
Line Input #1, a$
datereceived = a$
Line Input #1, a$
iteration = a$
Line Input #1, a$
serviceofsort = a$
Line Input #1, a$
casenumber = a$
Line Input #1, a$
armedforces.Value = Val(a$)
Line Input #1, a$
corporate.Value = Val(a$)
Line Input #1, a$
title = a$
Line Input #1, a$
papertype = a$
Line Input #1, a$
courtdate = a$
Line Input #1, a$
courttime = a$
Line Input #1, a$
defendant = a$
Line Input #1, a$
defendantsort = a$
Line Input #1, a$
sohomeaddress = a$
Line Input #1, a$
sohomeaddress2 = a$
Line Input #1, a$
sohomestate = a$
Line Input #1, a$
sohomezipcode = a$
Line Input #1, a$
sohomephone = a$
Line Input #1, a$
soworkaddress = a$
Line Input #1, a$
soworkaddress2 = a$
Line Input #1, a$
soworkstate = a$
Line Input #1, a$
soworkzipcode = a$
Line Input #1, a$
soworkphone = a$
Line Input #1, a$
daystorespond = a$
Line Input #1, a$
servicefee = a$
Line Input #1, a$
bill = Val(a$)
Line Input #1, a$
feedate = a$
Line Input #1, a$
receiptd = a$
Line Input #1, a$
checkd = a$
Line Input #1, a$
defendant = a$
Line Input #1, a$
defendantsort = a$
Line Input #1, a$
dhomeaddress = a$
Line Input #1, a$
dhomeaddress2 = a$
Line Input #1, a$
dhomestate = a$
Line Input #1, a$
dhomezipcode = a$
Line Input #1, a$
dhomephone = a$
Line Input #1, a$
dworkaddress = a$
Line Input #1, a$
dworkaddress2 = a$
Line Input #1, a$
dworkstate = a$
Line Input #1, a$
dworkzipcode = a$
Line Input #1, a$
dworkphone = a$
Line Input #1, a$
plaintiff = a$
Line Input #1, a$
plaintiffsort = a$
Line Input #1, a$
phomeaddress = a$
Line Input #1, a$
phomeaddress2 = a$
Line Input #1, a$
phomestate = a$
Line Input #1, a$
phomezipcode = a$
Line Input #1, a$
phomephone = a$
Line Input #1, a$
pworkaddress = a$
Line Input #1, a$
pworkaddress2 = a$
Line Input #1, a$
pworkstate = a$
Line Input #1, a$
pworkzipcode = a$
Line Input #1, a$
pworkphone = a$
Line Input #1, a$
assignedto = a$
Line Input #1, a$
assignedon = a$
Line Input #1, a$
served.Value = a$
Line Input #1, a$
nonservice.Value = a$
Line Input #1, a$
nsreason = a$
Line Input #1, a$
premarks = a$
Line Input #1, a$
levy.Text = a$
Line Input #1, a$
wremarks = a$
Line Input #1, a$
servicedate = a$
Line Input #1, a$
servicetime = a$
Line Input #1, a$
personserved = a$
Line Input #1, a$
locationserved = a$
Line Input #1, a$
relationship = a$
Line Input #1, a$
professional = a$
If thistab = 2 Then
        Line Input #1, a$
        ivd.Value = a$
        Line Input #1, a$
        custodian = a$
End If
If thistab = 3 Then
        Line Input #1, a$
        apptdate = a$
        Line Input #1, a$
        intrate = a$
        Line Input #1, a$
        datesatisfied = a$
        Line Input #1, a$
        judgementdate = a$
        Line Input #1, a$
        judgementamount = a$
        Line Input #1, a$
        estpayoffdate = a$
        Line Input #1, a$
        nulla.Value = Val(a$)
        Line Input #1, a$
        INTEREST = Val(a$)
        Line Input #1, a$
        commission = a$
        Line Input #1, a$
        perday = a$
        Line Input #1, a$
        balance = a$
        expaygrid.Rows = 1
        expaygrid.Row = 0
        For t% = 0 To 8
            expaygrid.Col = t%
            expaygrid.Text = ""
        Next t%
        expaygrid.Row = 0
        While Not EOF(1)
            Line Input #1, a$
            Line Input #1, b$
            Line Input #1, c$
            Line Input #1, d$
            Line Input #1, e$
            Line Input #1, f$
            Line Input #1, g$
            Line Input #1, h$
            Line Input #1, i$
            expaygrid.AddItem a$ + Chr$(9) + b$ + Chr$(9) + c$ + Chr$(9) + d$ + Chr$(9) + e$ + Chr$(9) + f$ + Chr$(9) + g$ + Chr$(9) + h$ + Chr$(9) + i$
        Wend
End If
Close #1
going:
Exit Sub
SKIPIT:
Close #1
Resume going
End Sub

Private Sub sheriff_KeyPress(KeyAscii As Integer)
If KeyAscii = 34 Then
    msg = MsgBox("The " + Chr$(34) + " character is not allowed.  Use the ' character instead.", 48, "Genesis Error Log")
    KeyAscii = 0
    Exit Sub
End If

End Sub

Private Sub sheriffaddress_KeyPress(KeyAscii As Integer)
If KeyAscii = 34 Then
    msg = MsgBox("The " + Chr$(34) + " character is not allowed.  Use the ' character instead.", 48, "Genesis Error Log")
    KeyAscii = 0
    Exit Sub
End If

End Sub

Private Sub sheriffaddress2_KeyPress(KeyAscii As Integer)
If KeyAscii = 34 Then
    msg = MsgBox("The " + Chr$(34) + " character is not allowed.  Use the ' character instead.", 48, "Genesis Error Log")
    KeyAscii = 0
    Exit Sub
End If

End Sub

Private Sub sohomeaddress_GotFocus()
If sohomeaddress = "" And sohomeaddress2 = "" And shomestate = "" And sohomezipcode = "" And sohomephone = "" Then
    On Error Resume Next
    If serviceof > "" And serviceofsort > "" Then
       Dim db As Database, ds As Recordset
       On Error GoTo oderror
od:
       Set db = OpenDatabase(nwl + "LAWSUITE.MDB")
       Set ds = db.OpenRecordset("select * from PEOPLE where dpname = " + Chr$(34) + serviceof + Chr$(34) + " and dpsort = " + Chr$(34) + serviceofsort + Chr$(34))
       If Not ds.EOF Then
          ds.MoveFirst
          sohomeaddress = ds("dphaddress")
          sohomeaddress2 = ds("dphaddress2")
          sohomestate = ds("hstate")
          sohomezipcode = ds("hzipcode")
          sohomephone = ds("dphphone")
        Else
            Set ds = db.OpenRecordset("select * from PEOPLE where dpname = " + Chr$(34) + serviceof + Chr$(34))
            If Not ds.EOF Then
              ds.MoveFirst
              sohomeaddress = ds("dphaddress")
              sohomeaddress2 = ds("dphaddress2")
                sohomestate = ds("hstate")
                sohomezipcode = ds("hzipcode")
              sohomephone = ds("dphphone")
            End If
       End If
       db.Close
    End If
End If
On Error GoTo 0
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If


End Sub


Private Sub sohomeaddress_KeyPress(KeyAscii As Integer)
If KeyAscii = 34 Then
    msg = MsgBox("The " + Chr$(34) + " character is not allowed.  Use the ' character instead.", 48, "Genesis Error Log")
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
    sohomeaddress2.SetFocus
End If

End Sub

Private Sub sohomeaddress2_KeyPress(KeyAscii As Integer)
If KeyAscii = 34 Then
    msg = MsgBox("The " + Chr$(34) + " character is not allowed.  Use the ' character instead.", 48, "Genesis Error Log")
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
    sohomestate.SetFocus
End If

End Sub

Private Sub sohomephone_GotFocus()
If Len(sohomephone) = 0 Then
    Dim db As Database, rs As Recordset
    On Error GoTo oderror
od:
    Set db = OpenDatabase(nwl + "LAWSUITE.MDB")
    Set rs = db.OpenRecordset("SELECT AREACODE FROM DEFAULTS")
    If rs.EOF Then
        db.Close
        Exit Sub
    End If
    rs.MoveFirst
    If IsNull(rs("AREACODE")) Then
        db.Close
        Exit Sub
    End If
    If Len(rs("AREACODE")) <> 3 Then
        db.Close
        Exit Sub
    End If
    Call sendopenpara
    Call SENDCHAR(Left$(rs("AREACODE"), 1))
    Call SENDCHAR(Mid$(rs("AREACODE"), 2, 1))
    Call SENDCHAR(Mid$(rs("AREACODE"), 3, 1))
    Call SENDEND
    db.Close
End If
On Error GoTo 0
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If


End Sub
Private Sub sohomephone_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
If Len(sohomephone) = 3 Then
    Call sendclosepara
End If
If Len(sohomephone) = 4 Then
    Call sendspace
End If
If Len(sohomephone) = 8 Then
    Call senddash
End If
If Len(sohomephone) = 13 Then
    Call sendspace
End If
End If
If KeyAscii = 13 Then
    soworkaddress.SetFocus
End If

End Sub
Private Sub SENDEND()
SendKeys "{END}"
SendKeys "{NUMLOCK}"
End Sub



Private Sub sohomephone_LostFocus()
If Len(sohomephone) = 5 Or Len(sohomephone) = 6 Then
    sohomephone = ""
End If
End Sub

Private Sub sohomestate_KeyPress(KeyAscii As Integer)
If KeyAscii = 34 Then
    msg = MsgBox("The " + Chr$(34) + " character is not allowed.  Use the ' character instead.", 48, "Genesis Error Log")
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
    sohomezipcode.SetFocus
End If
End Sub

Private Sub sohomezipcode_KeyPress(KeyAscii As Integer)
If KeyAscii = 34 Then
    msg = MsgBox("The " + Chr$(34) + " character is not allowed.  Use the ' character instead.", 48, "Genesis Error Log")
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
    sohomephone.SetFocus
End If
End Sub

Private Sub soworkaddress_GotFocus()
    If soworkaddress = "" And soworkaddress2 = "" And soworkstate = "" And soworkzipcode = "" And soworkphone = "" Then
    On Error Resume Next
    If serviceof > "" And serviceofsort > "" Then
       Dim db As Database, ds As Recordset
       On Error GoTo oderror
od:
       Set db = OpenDatabase(nwl + "LAWSUITE.MDB")
       Set ds = db.OpenRecordset("select * from PEOPLE where dpname = " + Chr$(34) + serviceof + Chr$(34) + " and dpsort = " + Chr$(34) + serviceofsort + Chr$(34))
       If Not ds.EOF Then
          ds.MoveFirst
          soworkaddress = ds("dpwaddress")
          soworkaddress2 = ds("dpwaddress2")
          soworkstate = ds("wstate")
          soworkzipcode = ds("wzipcode")
          soworkphone = ds("dPwphone")
       Else
           Set ds = db.OpenRecordset("select * from PEOPLE where dpname = " + Chr$(34) + serviceof + Chr$(34))
           If Not ds.EOF Then
              ds.MoveFirst
              soworkaddress = ds("dpwaddress")
              soworkaddress2 = ds("dpwaddress2")
                soworkstate = ds("wstate")
                soworkzipcode = ds("wzipcode")
              soworkphone = ds("dPwphone")
            End If
        End If
        db.Close
    End If
End If
On Error GoTo 0
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If



End Sub

Private Sub VScroll1_Change()

End Sub

Private Sub soworkaddress_KeyPress(KeyAscii As Integer)
If KeyAscii = 34 Then
    msg = MsgBox("The " + Chr$(34) + " character is not allowed.  Use the ' character instead.", 48, "Genesis Error Log")
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
    soworkaddress2.SetFocus
End If

End Sub

Private Sub soworkaddress2_KeyPress(KeyAscii As Integer)
If KeyAscii = 34 Then
    msg = MsgBox("The " + Chr$(34) + " character is not allowed.  Use the ' character instead.", 48, "Genesis Error Log")
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
    soworkstate.SetFocus
End If

End Sub

Private Sub soworkphone_GotFocus()
If Len(soworkphone) = 0 Then
    Dim db As Database, rs As Recordset
    On Error GoTo oderror
od:
    Set db = OpenDatabase(nwl + "LAWSUITE.MDB")
    Set rs = db.OpenRecordset("SELECT AREACODE FROM DEFAULTS")
    If rs.EOF Then
        db.Close
        Exit Sub
    End If
    rs.MoveFirst
    If IsNull(rs("AREACODE")) Then
        db.Close
        Exit Sub
    End If
    If Len(rs("AREACODE")) <> 3 Then
        db.Close
        Exit Sub
    End If
    Call sendopenpara
    Call SENDCHAR(Left$(rs("AREACODE"), 1))
    Call SENDCHAR(Mid$(rs("AREACODE"), 2, 1))
    Call SENDCHAR(Mid$(rs("AREACODE"), 3, 1))
    Call SENDEND
    db.Close
End If
On Error GoTo 0
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If


End Sub

Private Sub soworkphone_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
If Len(soworkphone) = 3 Then
    Call sendclosepara
End If
If Len(soworkphone) = 4 Then
    Call sendspace
End If
If Len(soworkphone) = 8 Then
    Call senddash
End If
If Len(soworkphone) = 13 Then
    Call sendspace
End If
End If
If KeyAscii = 13 Then
    casenumber.SetFocus
End If

End Sub


Private Sub soworkphone_LostFocus()
If Len(soworkphone) = 5 Or Len(soworkphone) = 6 Then
    soworkphone = ""
End If

End Sub

Private Sub soworkstate_KeyPress(KeyAscii As Integer)
If KeyAscii = 34 Then
    msg = MsgBox("The " + Chr$(34) + " character is not allowed.  Use the ' character instead.", 48, "Genesis Error Log")
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
    soworkzipcode.SetFocus
End If
End Sub

Private Sub soworkzipcode_KeyPress(KeyAscii As Integer)
If KeyAscii = 34 Then
    msg = MsgBox("The " + Chr$(34) + " character is not allowed.  Use the ' character instead.", 48, "Genesis Error Log")
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
    soworkphone.SetFocus
End If
End Sub

Private Sub todate_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
If Len(todate) = 1 Or Len(todate) = 4 Then
    Call sendslash
End If
End If


End Sub

Private Sub treasurer_KeyPress(KeyAscii As Integer)
If KeyAscii = 34 Then
    msg = MsgBox("The " + Chr$(34) + " character is not allowed.  Use the ' character instead.", 48, "Genesis Error Log")
    KeyAscii = 0
    Exit Sub
End If

End Sub

Private Sub treasureraddress1_KeyPress(KeyAscii As Integer)
If KeyAscii = 34 Then
    msg = MsgBox("The " + Chr$(34) + " character is not allowed.  Use the ' character instead.", 48, "Genesis Error Log")
    KeyAscii = 0
    Exit Sub
End If

End Sub

Private Sub treasureraddress2_KeyPress(KeyAscii As Integer)
If KeyAscii = 34 Then
    msg = MsgBox("The " + Chr$(34) + " character is not allowed.  Use the ' character instead.", 48, "Genesis Error Log")
    KeyAscii = 0
    Exit Sub
End If

End Sub

Private Sub UPDSHERIFF_Click()
Screen.MousePointer = 11
If Val(frmLogin.SUPERVISOR) = 1 Then
    Dim db As Database, ds As Recordset
    On Error GoTo oderror
od:
    Set db = OpenDatabase(nwc + dbname)
    Set ds = db.OpenRecordset("select * from system")
    If ds.EOF Then
        ds.AddNew
    Else
        ds.MoveFirst
        ds.Edit
    End If
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
On Error GoTo 0
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

Private Sub wlbdr_Click()
If wlbdr.Value = True Then
    fromdate.SetFocus
End If
End Sub

Private Sub clearit()
On Error Resume Next
If maintab.Tab = 0 Then
   Kill "holdm"
End If
If maintab.Tab = 1 Then
   Kill "holdw"
End If
If maintab.Tab = 2 Then
   Kill "holdf"
End If
If maintab.Tab = 3 Then
   Kill "holde"
End If
End Sub
Private Sub backdate()
Dim tempdate As String, totsi As Integer
Dim TUN, TOV, TINT1, TINT2 As Single
If CVDate(DATEPAID) > CVDate(estpayoffdate) Then
    estpayoffdate = DATEPAID
End If
If Val(judgementamount) = 0 Then
    Exit Sub
End If
Screen.MousePointer = 11
Call setintervals(totsi, DATEPAID)
BINTer = 0
tb = 0
bcommiss = 0
bprincip = Val(judgementamount)
T2 = Val(judgementamount) * (Val(intrate) / 365)
T2 = Val(Format$(T2, "######0.00"))
T2 = T2 * DateDiff("d", judgementdate, paydate(1))
T2 = Val(Format$(T2, "######0.00"))
If Val(judgementamount) + T2 > Val(exonfirst) Then
    bcommiss = (Val(exonfirst) * Val(excommrate1))
    bcommiss = bcommiss + ((bprincip + T2 - Val(exonfirst)) * Val(excommrate2))
Else
    bcommiss = (bprincip + T2) * Val(excommrate1)
End If
bcommiss = Val(Format$(bcommiss, "####0.00"))
For t% = 1 To totsi
    tempdate = Format$(paydate(t%), "mm/dd/yyyy")
    If CVDate(paydate(t%)) <= CVDate(Format$(DATEPAID, "yyyy/mm/dd")) Then
        T2 = bprincip * (Val(intrate) / 365)
        T2 = Val(Format$(T2, "######0.00"))
        T2 = T2 * DateDiff("d", (paydate(t% - 1)), tempdate)
        T2 = Val(Format$(T2, "######0.00"))
        BINTer = Val(BINTer) + T2 - payi(t% - 1)
        TUN = 0
        TOV = 0
        TINT1 = 0
        TINT2 = 0
        If Val(judgementamount) < Val(exonfirst) Then
            If Val(judgementamount) + tb + T2 < Val(exonfirst) Then
                TUN = T2
                TOV = 0
                TINT1 = Val(excommrate1)
                TINT2 = 0
            Else
            If Val(judgementamount) + tb < Val(exonfirst) Then
                TUN = (Val(exonfirst) - Val(judgementamount) - tb)
                TOV = T2 - TUN
                TINT1 = Val(excommrate1)
                TINT2 = Val(excommrate2)
            Else
                TUN = 0
                TOV = T2
                TINT1 = 0
                TINT2 = Val(excommrate2)
            End If
            End If
        Else
            TUN = 0
            TOV = T2
            TINT1 = 0
            TINT2 = Val(excommrate2)
        End If
        tb = tb + T2
        bcommiss = bcommiss + (TUN * TINT1) + (TOV * TINT2)
        bcommiss = bcommiss - payc(t% - 1)
        bprincip = bprincip - payamount(t%)
    Else
        t% = totsi
    End If
Next t%
If payi(totsi) > 0 Or payc(totsi) > 0 Then
    BINTer = Val(BINTer) - payi(totsi)
    bcommiss = bcommiss - payc(totsi)
End If
bcommiss = Format(bcommiss, "####0.00")
BINTer = Format(BINTer, "####0.00")
bprincip = Format$(bprincip, "#######0.00")
Screen.MousePointer = 0
Exit Sub

End Sub
Private Sub printreceipt(fee As Single, rnumber, cnumber, rname, raddress1, raddress2, rstate, rzipcode, dt As String)
Dim db As Database, rs As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nwc + dbname)
Set rs = db.OpenRecordset("select office, sheriffaddress,sheriffaddress2,sheriffphone from system")
If rs.EOF Then
    msg = MsgBox("Sheriff information is missing.  Enter applicable information on the SYSTEM tab.", 48, "Genesis Error Log")
    db.Close
    Exit Sub
Else
    rs.MoveFirst
End If
Printer.FontSize = 12
Printer.Print Tab(5); rs("office")
Printer.Print Tab(5); rs("sheriffaddress"); Tab(70); "Receipt#:  "; Tab(82); rnumber
Printer.Print Tab(5); rs("sheriffaddress2"); Tab(70); "Date:"; Tab(82); Format$(Date$, "mm/dd/yyyy")
Printer.Print Tab(5); rs("sheriffphone"); Tab(70); "Amount:"; Tab(82); Format$(fee, "$#########0.00")
Printer.Print Tab(70); "Check#:"; Tab(82); cnumber
Printer.Print Tab(70); "Received:"; Tab(82); Format$(dt, "mm/dd/yyyy")
Printer.Print
Printer.Print
Printer.Print Tab(10); "FROM:"; Tab(30); rname
Printer.Print Tab(30); raddress1
Printer.Print Tab(30); raddress2 + " " + rstate + " " + rzipcode
Printer.Print
Printer.Print Tab(10); "RE:"; Tab(30); "CASE #:"; Tab(40); casenumber
Printer.Print
Printer.Print Tab(30); plaintiff
Printer.Print Tab(30); "vs"
Printer.Print Tab(30); defendant
Printer.Print
Printer.Print Tab(10); "SERVICE OF: "; Tab(30); serviceof
Printer.Print Tab(10); "PAPER TYPE:"; Tab(30); papertype
Printer.EndDoc
If Val(rnumber) > 0 Then
    Set rs = db.OpenRecordset("select * from receipt WHERE ITERATION = " + Chr$(34) + iteration + Chr$(34) + " AND SERVICEOF = " + Chr$(34) + serviceof + Chr$(34) + " AND DATERECEIVED = #" + Format$(datereceived, "mm/dd/yyyy") + "# AND RECEIPTNUM = " + rnumber)
    If rs.EOF Then
        rs.AddNew
    Else
        rs.MoveFirst
        rs.Edit
    End If
    On Error Resume Next
    rs("iteration") = iteration
    rs("serviceof") = serviceof
    rs("datereceived") = feedate
    rs("casenumber") = casenumber
    rs("papertype") = papertype
    rs("servicefee") = fee
    rs("receiptnum") = Val(rnumber)
    rs("checknum") = cnumber
    rs("datereceipt") = Format$(Date$, "mm/dd/yyyy")
    rs("defendant") = defendant
    rs("plaintiff") = plaintiff
    rs("from") = rname
    rs("fromaddress1") = raddress1
    rs("fromaddress2") = raddress2
    rs.Update
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

Private Sub wremarks_KeyPress(KeyAscii As Integer)
If KeyAscii = 34 Then
    msg = MsgBox("The " + Chr$(34) + " character is not allowed.  Use the ' character instead.", 48, "Genesis Error Log")
    KeyAscii = 0
    Exit Sub
End If

End Sub
Private Sub setsort(ff, LF As Integer, hoLdname, holdsort As String)
osort1$ = ""
If ff = 1 Then
    GoSub setfirst
Else
    GoSub setlast
End If
holdsort = osort1$
Exit Sub
setfirst:
If Left$(hoLdname, 1) = " " Then
   hoLdname = Mid$(hoLdname, 2)
   osort1$ = Left$(hoLdname, 15)
End If
If InStr(hoLdname, " CORP") > 0 Or InStr(hoLdname, ",INC") > 0 Or InStr(hoLdname, "COMPANY") > 0 Or InStr(hoLdname, "INC.") > 0 Then
   osort1$ = Left$(hoLdname, 15)
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
If InStr(tso$, ",") > 0 Then
    tso$ = Left$(tso$, InStr(tso$, ",") - 1)
End If
firstspace% = 0
While Right$(tso$, 1) = " " And Len(tso$) > 1
   tso$ = Left$(tso$, Len(tso$) - 1)
Wend
For t% = Len(tso$) To 1 Step -1
    If Mid$(tso$, t%, 1) = " " Then
       firstspace% = t%
       t% = 1
    End If
Next t%
If firstspace% = 0 Then
    If osort1$ = "" Then
        osort1$ = Left$(tso$, 15)
    End If
    Return
End If
tempsort$ = Mid$(tso$, firstspace% + 1)
If Len(tempsort$) >= 10 Then
    tempsort$ = Left$(tempsort$, 10)
Else
    tempsort$ = tempsort$ + Space$(10 - Len(tempsort$))
End If
tso$ = Left$(tso$, firstspace% - 1)
If Len(tso$) >= 5 Then
    tso$ = Left$(tso$, 5)
Else
    tso$ = tso$ + Space$(5 - Len(tso$))
End If
tempsort$ = tempsort$ + tso$
If osort1$ = "" Then
   osort1$ = tempsort$
End If
Return
setlast:
If Left$(hoLdname, 1) = " " Then
   hoLdname = Mid$(hoLdname, 2)
   osort1$ = Left$(hoLdname, 15)
End If
If InStr(hoLdname, " CORP") > 0 Or InStr(hoLdname, ",INC") > 0 Or InStr(hoLdname, "COMPANY") > 0 Or InStr(hoLdname, "INC.") > 0 Then
   osort1$ = Left$(hoLdname, 15)
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
For t% = 1 To Len(tso$)
    If Mid$(tso$, t%, 1) = " " Then
       firstspace% = t%
       t% = Len(tso$)
    End If
Next t%
If firstspace% = 0 Then
    If osort1$ = "" Then
        osort1$ = Left$(tso$, 15)
    End If
    Return
End If
tempsort$ = Mid$(tso$, firstspace% + 1)
If Len(tempsort$) >= 5 Then
    tempsort$ = Left$(tempsort$, 5)
Else
    tempsort$ = tempsort$ + Space$(5 - Len(tempsort$))
End If
tso$ = Left$(tso$, firstspace% - 1)
If Right$(tso$, 1) = "," Then
    tso$ = Left$(tso$, Len(tso$) - 1)
End If
If Len(tso$) >= 10 Then
    tso$ = Left$(tso$, 10)
Else
    tso$ = Left$(tso$, firstspace% - 1)
    tso$ = tso$ + Space$(10 - Len(tso$))
End If
tempsort$ = tso$ + tempsort$
If osort1$ = "" Then
   osort1$ = tempsort$
End If
Return
End Sub
Private Sub defaultprinter(pname As String)
Dim osinfo As OSVERSIONINFO
Dim retvalue As Integer

osinfo.dwOSVersionInfoSize = 148
osinfo.szCSDVersion = Space$(128)
retvalue = GetVersionExA(osinfo)
Call Win95SetDefaultPrinter(pname)

End Sub
Private Sub sendcolon()
SendKeys ":"
End Sub
Private Sub LOADOTHER()
On Error GoTo oderror
Dim db As Database, rs As Recordset
od:
Set db = OpenDatabase(nwl + "LAWSUITE.MDB")
othername.clear
Set rs = db.OpenRecordset("select * from professionals where type = 'A'")
If Not rs.EOF Then
    rs.MoveFirst
End If
While Not rs.EOF
    othername.AddItem rs("PROFNAME")
    rs.MoveNext
Wend
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
Private Sub cp(plin As String)
AutoRedraw = -1
HalfWidth = Printer.TextWidth(plin) / 2
Printer.CurrentX = Printer.ScaleWidth / 2 - HalfWidth
Printer.Print plin
End Sub
Private Sub setfnln()
Screen.MousePointer = 11
If Val(frmLogin.SUPERVISOR) = 1 Then
    Dim db As Database, ds As Recordset
    On Error GoTo oderror
od:
    Set db = OpenDatabase(nwc + dbname)
    Set ds = db.OpenRecordset("select fnf,lnf from system")
    If ds.EOF Then
        ds.AddNew
    Else
        ds.MoveFirst
        ds.Edit
    End If
    ds("fnf") = fnf
    ds("lnf") = lnf
    ds.Update
    db.Close
End If
Screen.MousePointer = 0
On Error GoTo 0
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If


End Sub
Private Sub printroutines(pname As String, inp As String)
Dim db As Database, ds As Recordset, rs1, ds2 As Recordset, ds3 As Recordset, ds4 As Recordset
Dim ps, pns, rs, rns, s, ns, pn, rn, n As Integer, cty As String
Dim tps, tpns, trs, trns, ts, tns, tpn, trn, tn As Integer
On Error GoTo oderror1
od1:
Set db = OpenDatabase(nwc + dbname)
Set ds = db.OpenRecordset("select office from system")
If ds.EOF Then
    msg = MsgBox("Incomplete Sheriff information.", 48, "Genesis Error Log")
    Screen.MousePointer = 0
    db.Close
    Exit Sub
End If
ds.MoveFirst
If Not IsNull(ds("office")) Then
    cty = ds("office")
Else
    msg = MsgBox("Incomplete Sheriff information.", 48, "Genesis Error Log")
    Screen.MousePointer = 0
    db.Close
    Exit Sub
End If

Select Case pname
Case "ofsrbdrrtn"
Set rs1 = db.OpenRecordset("select * from ofsrbdr")
If Not rs1.EOF Then
    rs1.MoveFirst
End If
While Not rs1.EOF
    rs1.Delete
    rs1.MoveNext
Wend
Set ds = db.OpenRecordset("select serviceof,served,nonservice,assignedto,servicedate from magistrate where (SERVICEDATE >= #" + fromdate + "# and SERVICEDATE <= #" + todate + "#) or assignedon between #" + fromdate + "# and #" + todate + "# " + _
"UNION ALL select serviceof,served,nonservice,assignedto,servicedate from familycourt where SERVICEDATE >= #" + fromdate + "# and (SERVICEDATE <= #" + todate + "#) or assignedon between #" + fromdate + "# and #" + todate + "#  " + _
"UNION ALL select serviceof,served,nonservice,assignedto,servicedate from writother where (SERVICEDATE >= #" + fromdate + "# and SERVICEDATE <= #" + todate + "#) or assignedon between #" + fromdate + "# and #" + todate + "#  " + _
"UNION ALL select serviceof,served,nonservice,assignedto,servicedate from executions where (SERVICEDATE >= #" + fromdate + "# and SERVICEDATE <= #" + todate + "#) or assignedon between #" + fromdate + "# and #" + todate + "#  ORDER BY assignedto,servicedate,served,nonservice")
lastdep$ = ""
lastdt$ = ""
s = 0
ns = 0
n = 0
If Not ds.EOF Then
    ds.MoveFirst
    lastdep$ = ds("assignedto")
Else
    NOREPORT = 1
    db.Close
    Exit Sub
End If
While Not ds.EOF
    If lastdep$ <> ds("assignedto") Then
        n = 0
        Set ds2 = db.OpenRecordset("select count(*) as ctrow from magistrate where assignedon between #" + fromdate + "# and #" + todate + "# and served = '0' and nonservice = '0' and assignedto = " + Chr$(34) + lastdep$ + Chr$(34))
        If Not ds2.EOF Then
            ds2.MoveFirst
            n = n + ds2("ctrow")
        End If
        Set ds2 = db.OpenRecordset("select count(*) as ctrow from familycourt where assignedon between #" + fromdate + "# and #" + todate + "# and served = '0' and nonservice = '0' and assignedto = " + Chr$(34) + lastdep$ + Chr$(34))
        If Not ds2.EOF Then
            ds2.MoveFirst
            n = n + ds2("ctrow")
        End If
        Set ds2 = db.OpenRecordset("select count(*) as ctrow from executions where assignedon between #" + fromdate + "# and #" + todate + "# and served = '0' and nonservice = '0' and assignedto = " + Chr$(34) + lastdep$ + Chr$(34))
        If Not ds2.EOF Then
            ds2.MoveFirst
            n = n + ds2("ctrow")
        End If
        Set ds2 = db.OpenRecordset("select count(*) as ctrow from writother where assignedon between #" + fromdate + "# and #" + todate + "# and served = '0' and nonservice = '0' and assignedto = " + Chr$(34) + lastdep$ + Chr$(34))
        If Not ds2.EOF Then
            ds2.MoveFirst
            n = n + ds2("ctrow")
        End If
        Set rs1 = db.OpenRecordset("select * from ofsrbdr")
        rs1.AddNew
        rs1("officersname") = lastdep$
        rs1("served") = s
        rs1("nonservice") = ns
        rs1("notserved") = n
        rs1.Update
        s = 0
        ns = 0
        n = 0
        lastdep$ = ds("assignedto")
    End If
    a = ds("serviceof")
    If ds("served") = "1" And ds("servicedate") >= CDate(fromdate) And ds("servicedate") <= CDate(todate) Then
        s = s + 1
    End If
    If ds("nonservice") = "1" And ds("servicedate") >= CDate(fromdate) And ds("servicedate") <= CDate(todate) Then
        ns = ns + 1
    End If
    If ds("served") <> "1" And ds("nonservice") <> "1" Then
        n = n + 1
    End If
    ds.MoveNext
Wend
If lastdep$ > "" Then
    Set rs1 = db.OpenRecordset("select * from ofsrbdr")
    rs1.AddNew
    rs1("officersname") = lastdep$
    rs1("served") = s
    rs1("nonservice") = ns
    rs1("notserved") = n
    rs1.Update
End If
db.Close
Exit Sub
Case "sdrbdrrtn"
Set rs1 = db.OpenRecordset("select * from sdrbdr")
If Not rs1.EOF Then
    rs1.MoveFirst
End If
While Not rs1.EOF
    rs1.Delete
    rs1.MoveNext
Wend
If inp = "S" Then
    Set ds = db.OpenRecordset("select served,nonservice,assignedto,servicedate,servicetime,locationserved,serviceof from magistrate where SERVICEDATE >= #" + fromdate + "# and SERVICEDATE <= #" + todate + "# and served = '1' UNION ALL select served,nonservice,assignedto,servicedate,servicetime,locationserved,serviceof from familycourt where SERVICEDATE >= #" + fromdate + "# and SERVICEDATE <= #" + todate + "# and served = '1' UNION ALL select served,nonservice,assignedto,servicedate,servicetime,locationserved,serviceof from writother where SERVICEDATE >= #" + fromdate + "# and SERVICEDATE <= #" + todate + "# and served = '1' UNION ALL select served,nonservice,assignedto,servicedate,servicetime,locationserved,serviceof from executions where SERVICEDATE >= #" + fromdate + "# and SERVICEDATE <= #" + todate + "# and served = '1' ORDER BY assignedto,servicedate,servicetime")
Else
    Set ds = db.OpenRecordset("select served,nonservice,assignedto,servicedate,servicetime,locationserved,serviceof from magistrate where SERVICEDATE >= #" + fromdate + "# and SERVICEDATE <= #" + todate + "# and nonservice = '1' UNION ALL select served,nonservice,assignedto,servicedate,servicetime,locationserved,serviceof from familycourt where SERVICEDATE >= #" + fromdate + "# and SERVICEDATE <= #" + todate + "# and nonservice = '1' UNION ALL select served,nonservice,assignedto,servicedate,servicetime,locationserved,serviceof from writother where SERVICEDATE >= #" + fromdate + "# and SERVICEDATE <= #" + todate + "# and nonservice = '1' UNION ALL select served,nonservice,assignedto,servicedate,servicetime,locationserved,serviceof from executions where SERVICEDATE >= #" + fromdate + "# and SERVICEDATE <= #" + todate + "# and nonservice = '1' ORDER BY assignedto,servicedate,servicetime")
End If
If Not ds.EOF Then
    ds.MoveFirst
    lastdep$ = ds("assignedto")
Else
    NOREPORT = 1
    db.Close
    Exit Sub
End If
While Not ds.EOF
    Set rs1 = db.OpenRecordset("select * from sdrbdr")
    rs1.AddNew
    rs1("Officer's Name") = ds("assignedto")
    rs1("Service Date") = ds("servicedate")
    rs1("Service Time") = Left$(ds("servicetime"), 10)
    rs1("Service Of") = ds("serviceof")
    rs1("Location") = ds("locationserved")
    rs1.Update
    ds.MoveNext
Wend
db.Close
Exit Sub
Case "wlbdrrtn"
Set rs1 = db.OpenRecordset("select * from wlbdr")
If Not rs1.EOF Then
    rs1.MoveFirst
End If
While Not rs1.EOF
    rs1.Delete
    rs1.MoveNext
Wend
Set ds = db.OpenRecordset("select datereceived,iteration,serviceof,plaintiff,papertype,served,nonservice,servicedate from writother where datereceived >= #" + fromdate + "# and datereceived <= #" + todate + "# order by datereceived,serviceof,iteration")
If Not ds.EOF Then
    ds.MoveFirst
Else
    NOREPORT = 1
    db.Close
    Exit Sub
End If
While Not ds.EOF
    If ds("served") = 1 Then
        st$ = "Served"
    End If
    If ds("nonservice") = 1 Then
        st$ = "Non-Service"
    End If
    If ds("served") <> 1 And ds("nonservice") <> 1 Then
        st$ = "Outstanding"
    End If
    Set rs1 = db.OpenRecordset("select * from wlbdr")
    rs1.AddNew
    rs1("Date Received") = ds("datereceived")
    rs1("Service Of") = ds("serviceof")
    rs1("Iteration") = ds("iteration")
    rs1("Plaintiff") = ds("plaintiff")
    rs1("Paper Type") = ds("paperType")
    rs1("Status") = st$
    rs1("Service/Non-Service Date") = ds("servicedate")
    rs1.Update
    ds.MoveNext
Wend
db.Close
Exit Sub
Case "aprbdrrtn"
Set rs1 = db.OpenRecordset("select * from wlbdr")
If Not rs1.EOF Then
    rs1.MoveFirst
End If
While Not rs1.EOF
    rs1.Delete
    rs1.MoveNext
Wend
Set ds = db.OpenRecordset("select datereceived,iteration,serviceof,plaintiff,papertype,served,nonservice,servicedate from writother where datereceived >= #" + fromdate + "# and datereceived <= #" + todate + "# UNION select datereceived,iteration,serviceof,plaintiff,papertype,served,nonservice,servicedate from familycourt where datereceived >= #" + fromdate + "# and datereceived <= #" + todate + "# UNION select datereceived,iteration,serviceof,plaintiff,papertype,served,nonservice,servicedate from magistrate where datereceived >= #" + fromdate + "# and datereceived <= #" + todate + "# UNION select datereceived,iteration,serviceof,plaintiff,papertype,served,nonservice,servicedate from executions where datereceived >= #" + fromdate + "# and datereceived <= #" + todate + "# order by datereceived,serviceof,iteration")
If Not ds.EOF Then
    ds.MoveFirst
Else
    NOREPORT = 1
    db.Close
    Exit Sub
End If
While Not ds.EOF
    If ds("served") = 1 Then
        st$ = "Served"
    End If
    If ds("nonservice") = 1 Then
        st$ = "Non-Service"
    End If
    If ds("served") <> 1 And ds("nonservice") <> 1 Then
        st$ = "Outstanding"
    End If
    Set rs1 = db.OpenRecordset("select * from wlbdr")
    rs1.AddNew
    rs1("Date Received") = ds("datereceived")
    rs1("Service Of") = ds("serviceof")
    rs1("Iteration") = ds("iteration")
    rs1("Plaintiff") = ds("plaintiff")
    rs1("Paper Type") = ds("paperType")
    rs1("Status") = st$
    rs1("Service/Non-Service Date") = ds("servicedate")
    rs1.Update
    ds.MoveNext
Wend
db.Close
Exit Sub
Case "walbdrrtn"
Set rs1 = db.OpenRecordset("select * from albdr")
If Not rs1.EOF Then
    rs1.MoveFirst
End If
While Not rs1.EOF
    rs1.Delete
    rs1.MoveNext
Wend
Set ds = db.OpenRecordset("select datereceived,iteration,serviceof,plaintiff,papertype,courtdate,assignedto from writother where datereceived >= #" + fromdate + "# and datereceived <= #" + todate + "# order by datereceived,serviceof,iteration")
If Not ds.EOF Then
    ds.MoveFirst
Else
    NOREPORT = 1
    db.Close
    Exit Sub
End If
While Not ds.EOF
    Set rs1 = db.OpenRecordset("select * from albdr")
    rs1.AddNew
    rs1("Date Received") = ds("datereceived")
    rs1("Service Of") = ds("serviceof")
    rs1("Iteration") = ds("iteration")
    rs1("Plaintiff") = ds("plaintiff")
    rs1("Paper Type") = ds("paperType")
    rs1("Court Date") = ds("courtdate")
    rs1("Assigned To") = ds("assignedto")
    rs1.Update
    ds.MoveNext
Wend
db.Close
Exit Sub
Case "malbdrrtn"
Set rs1 = db.OpenRecordset("select * from albdr")
If Not rs1.EOF Then
    rs1.MoveFirst
End If
While Not rs1.EOF
    rs1.Delete
    rs1.MoveNext
Wend
Set ds = db.OpenRecordset("select datereceived,iteration,serviceof,plaintiff,papertype,courtdate,assignedto from magistrate where datereceived >= #" + fromdate + "# and datereceived <= #" + todate + "# order by datereceived,serviceof,iteration")
If Not ds.EOF Then
    ds.MoveFirst
Else
    NOREPORT = 1
    db.Close
    Exit Sub
End If
While Not ds.EOF
    Set rs1 = db.OpenRecordset("select * from albdr")
    rs1.AddNew
    rs1("Date Received") = ds("datereceived")
    rs1("Service Of") = ds("serviceof")
    rs1("Iteration") = ds("iteration")
    rs1("Plaintiff") = ds("plaintiff")
    rs1("Paper Type") = ds("paperType")
    rs1("Court Date") = ds("courtdate")
    rs1("Assigned To") = ds("assignedto")
    rs1.Update
    ds.MoveNext
Wend
db.Close
Exit Sub
Case "falbdrrtn"
Set rs1 = db.OpenRecordset("select * from albdr")
If Not rs1.EOF Then
    rs1.MoveFirst
End If
While Not rs1.EOF
    rs1.Delete
    rs1.MoveNext
Wend
Set ds = db.OpenRecordset("select datereceived,iteration,serviceof,plaintiff,papertype,courtdate,assignedto from familycourt where datereceived >= #" + fromdate + "# and datereceived <= #" + todate + "# order by datereceived,serviceof,iteration")
If Not ds.EOF Then
    ds.MoveFirst
Else
    NOREPORT = 1
    db.Close
    Exit Sub
End If
While Not ds.EOF
    Set rs1 = db.OpenRecordset("select * from albdr")
    rs1.AddNew
    rs1("Date Received") = ds("datereceived")
    rs1("Service Of") = ds("serviceof")
    rs1("Iteration") = ds("iteration")
    rs1("Plaintiff") = ds("plaintiff")
    rs1("Paper Type") = ds("paperType")
    rs1("Court Date") = ds("courtdate")
    rs1("Assigned To") = ds("assignedto")
    rs1.Update
    ds.MoveNext
Wend
db.Close
Exit Sub
Case "ealbdrrtn"
Set rs1 = db.OpenRecordset("select * from albdr")
If Not rs1.EOF Then
    rs1.MoveFirst
End If
While Not rs1.EOF
    rs1.Delete
    rs1.MoveNext
Wend
Set ds = db.OpenRecordset("select datereceived,iteration,serviceof,plaintiff,papertype,courtdate,assignedto from executions where datereceived >= #" + fromdate + "# and datereceived <= #" + todate + "# order by datereceived,serviceof,iteration")
If Not ds.EOF Then
    ds.MoveFirst
Else
    NOREPORT = 1
    db.Close
    Exit Sub
End If
While Not ds.EOF
    Set rs1 = db.OpenRecordset("select * from albdr")
    rs1.AddNew
    rs1("Date Received") = ds("datereceived")
    rs1("Service Of") = ds("serviceof")
    rs1("Iteration") = ds("iteration")
    rs1("Plaintiff") = ds("plaintiff")
    rs1("Paper Type") = ds("paperType")
    rs1("Court Date") = ds("courtdate")
    rs1("Assigned To") = ds("assignedto")
    rs1.Update
    ds.MoveNext
Wend
db.Close
Exit Sub
Case "fclbdrrtn"
Set rs1 = db.OpenRecordset("select * from wlbdr")
If Not rs1.EOF Then
    rs1.MoveFirst
End If
While Not rs1.EOF
    rs1.Delete
    rs1.MoveNext
Wend
Set ds = db.OpenRecordset("select datereceived,iteration,serviceof,plaintiff,papertype,served,nonservice,servicedate from familycourt where datereceived >= #" + fromdate + "# and datereceived <= #" + todate + "# order by datereceived,serviceof,iteration")
If Not ds.EOF Then
    ds.MoveFirst
Else
    NOREPORT = 1
    db.Close
    Exit Sub
End If
While Not ds.EOF
    If ds("served") = 1 Then
        st$ = "Served"
    End If
    If ds("nonservice") = 1 Then
        st$ = "Non-Service"
    End If
    If ds("served") <> 1 And ds("nonservice") <> 1 Then
        st$ = "Outstanding"
    End If
    Set rs1 = db.OpenRecordset("select * from wlbdr")
    rs1.AddNew
    rs1("Date Received") = ds("datereceived")
    rs1("Service Of") = ds("serviceof")
    rs1("Iteration") = ds("iteration")
    rs1("Plaintiff") = ds("plaintiff")
    rs1("Paper Type") = ds("paperType")
    rs1("Status") = st$
    rs1("Service/Non-Service Date") = ds("servicedate")
    rs1.Update
    ds.MoveNext
Wend
db.Close
Exit Sub
Case "mlbdrrtn"
Set rs1 = db.OpenRecordset("select * from wlbdr")
If Not rs1.EOF Then
    rs1.MoveFirst
End If
While Not rs1.EOF
    rs1.Delete
    rs1.MoveNext
Wend
Set ds = db.OpenRecordset("select datereceived,iteration,serviceof,plaintiff,papertype,served,nonservice,servicedate from magistrate where datereceived >= #" + fromdate + "# and datereceived <= #" + todate + "# order by datereceived,serviceof,iteration")
If Not ds.EOF Then
    ds.MoveFirst
Else
    NOREPORT = 1
    db.Close
    Exit Sub
End If
While Not ds.EOF
    If ds("served") = 1 Then
        st$ = "Served"
    End If
    If ds("nonservice") = 1 Then
        st$ = "Non-Service"
    End If
    If ds("served") <> 1 And ds("nonservice") <> 1 Then
        st$ = "Outstanding"
    End If
    Set rs1 = db.OpenRecordset("select * from wlbdr")
    rs1.AddNew
    rs1("Date Received") = ds("datereceived")
    rs1("Service Of") = ds("serviceof")
    rs1("Iteration") = ds("iteration")
    rs1("Plaintiff") = ds("plaintiff")
    rs1("Paper Type") = ds("paperType")
    rs1("Status") = st$
    rs1("Service/Non-Service Date") = ds("servicedate")
    rs1.Update
    ds.MoveNext
Wend
db.Close
Exit Sub
Case "elbdrrtn"
Set rs1 = db.OpenRecordset("select * from wlbdr")
If Not rs1.EOF Then
    rs1.MoveFirst
End If
While Not rs1.EOF
    rs1.Delete
    rs1.MoveNext
Wend
Set ds = db.OpenRecordset("select datereceived,iteration,serviceof,plaintiff,papertype,served,nonservice,servicedate from executions where datereceived >= #" + fromdate + "# and datereceived <= #" + todate + "# order by datereceived,serviceof,iteration")
If Not ds.EOF Then
    ds.MoveFirst
Else
    NOREPORT = 1
    db.Close
    Exit Sub
End If
While Not ds.EOF
    If ds("served") = 1 Then
        st$ = "Served"
    End If
    If ds("nonservice") = 1 Then
        st$ = "Non-Service"
    End If
    If ds("served") <> 1 And ds("nonservice") <> 1 Then
        st$ = "Outstanding"
    End If
    Set rs1 = db.OpenRecordset("select * from wlbdr")
    rs1.AddNew
    rs1("Date Received") = ds("datereceived")
    rs1("Service Of") = ds("serviceof")
    rs1("Iteration") = ds("iteration")
    rs1("Plaintiff") = ds("plaintiff")
    rs1("Paper Type") = ds("paperType")
    rs1("Status") = st$
    rs1("Service/Non-Service Date") = ds("servicedate")
    rs1.Update
    ds.MoveNext
Wend
db.Close
Case "opsrrtn"
Set rs1 = db.OpenRecordset("select * from wlbdr")
If Not rs1.EOF Then
    rs1.MoveFirst
End If
While Not rs1.EOF
    rs1.Delete
    rs1.MoveNext
Wend
If inp = "A" Then
    Set ds = db.OpenRecordset("select served,nonservice,datereceived,serviceof,iteration,assignedto,papertype,servicedate, assignedon from magistrate where assignedon between #" + fromdate + "# and #" + todate + "# or servicedate between #" + fromdate + "# and #" + todate + "#" + _
    "UNION select served,nonservice,datereceived,serviceof,iteration,assignedto,papertype,servicedate, assignedon from writother where assignedon between #" + fromdate + "# and #" + todate + "# or servicedate between #" + fromdate + "# and #" + todate + "#  " + _
    "UNION select served,nonservice,datereceived,serviceof,iteration,assignedto,papertype,servicedate, assignedon from familycourt where assignedon between #" + fromdate + "# and #" + todate + "# or servicedate between #" + fromdate + "# and #" + todate + "# " + _
    "UNION select served,nonservice,datereceived,serviceof,iteration,assignedto,papertype,servicedate, assignedon from executions where assignedon between #" + fromdate + "# and #" + todate + "# or servicedate between #" + fromdate + "# and #" + todate + "# order by assignedto,datereceived,iteration")
Else
    Set ds = db.OpenRecordset("select served,nonservice,datereceived,serviceof,iteration,assignedto,papertype,servicedate, assignedon from magistrate where (assignedon between #" + fromdate + "# and #" + todate + "# or servicedate between #" + fromdate + "# and #" + todate + "#) and assignedto like '*" + inp + "*' " + _
    "UNION select served,nonservice,datereceived,serviceof,iteration,assignedto,papertype,servicedate, assignedon from writother where (assignedon between #" + fromdate + "# and #" + todate + "# or servicedate between #" + fromdate + "# and #" + todate + "#) and assignedto like '*" + inp + "*'  " + _
    "UNION select served,nonservice,datereceived,serviceof,iteration,assignedto,papertype,servicedate, assignedon from familycourt where (assignedon between #" + fromdate + "# and #" + todate + "# or servicedate between #" + fromdate + "# and #" + todate + "#) and assignedto like '*" + inp + "*' " + _
    "UNION select served,nonservice,datereceived,serviceof,iteration,assignedto,papertype,servicedate, assignedon from executions where (assignedon between #" + fromdate + "# and #" + todate + "# or servicedate between #" + fromdate + "# and #" + todate + "#) and assignedto like '*" + inp + "*' order by assignedto,datereceived,iteration")
End If
lastdep$ = ""
ps = 0
If Not ds.EOF Then
    ds.MoveFirst
    lastdep$ = ds("assignedto")
Else
    NOREPORT = 1
    db.Close
    Exit Sub
End If
While Not ds.EOF
    If ds("served") = 1 And ds("servicedate") <= CDate(todate) And ds("servicedate") >= CDate(fromdate) Then
        st$ = "Served"
    End If
    If ds("nonservice") = 1 And ds("servicedate") <= CDate(todate) And ds("servicedate") >= CDate(fromdate) Then
        st$ = "Non-Service"
    End If
    If ds("served") <> 1 And ds("nonservice") <> 1 And ds("assignedon") >= CDate(fromdate) And ds("assignedon") <= CDate(todate) Then
        st$ = "Outstanding"
    End If
    Set rs1 = db.OpenRecordset("select * from wlbdr")
    rs1.AddNew
    rs1("Date Received") = ds("datereceived")
    rs1("Service Of") = ds("serviceof")
    rs1("Iteration") = ds("iteration")
    rs1("Plaintiff") = ds("assignedto")
    rs1("Paper Type") = ds("paperType")
    rs1("Status") = st$
    rs1("Service/Non-Service Date") = ds("servicedate")
    rs1.Update
    lastdep$ = ds("assignedto")
    ds.MoveNext
Wend
ds.Close
db.Close
Exit Sub
Case "opborrtn"
Set rs1 = db.OpenRecordset("select * from opbor")
If Not rs1.EOF Then
    rs1.MoveFirst
End If
While Not rs1.EOF
    rs1.Delete
    rs1.MoveNext
Wend
If inp = "A" Then
    Set ds = db.OpenRecordset("select datereceived,iteration,casenumber,serviceof,plaintiff,papertype,courtdate,served,nonservice,assignedto from magistrate where assignedon <= #" + Date$ + "# and served <> '1' and nonservice <> '1' UNION select datereceived,iteration,casenumber,serviceof,plaintiff,papertype,courtdate,served,nonservice,assignedto from writother where assignedon <= #" + Date$ + "# and served <> '1' and nonservice <>'1' UNION select datereceived,iteration,casenumber,serviceof,plaintiff,papertype,courtdate,served,nonservice,assignedto from familycourt where assignedon <= #" + Date$ + "# and served <> '1' and nonservice <> '1' order by assignedto,datereceived,iteration")
Else
    Set ds = db.OpenRecordset("select datereceived,iteration,casenumber,serviceof,plaintiff,papertype,courtdate,served,nonservice,assignedto from magistrate where assignedon <= #" + Date$ + "# and served <> '1' and nonservice <> '1' AND ASSIGNEDTO LIKE '*" + inp + "*' UNION select datereceived,iteration,casenumber,serviceof,plaintiff,papertype,courtdate,served,nonservice,assignedto from writother where assignedon <= #" + Date$ + "# and served <> '1' and nonservice <>'1'  AND ASSIGNEDTO LIKE '*" + inp + "*' UNION select datereceived,iteration,casenumber,serviceof,plaintiff,papertype,courtdate,served,nonservice,assignedto from familycourt where assignedon <= #" + Date$ + "# and served <> '1' and nonservice <> '1' AND ASSIGNEDTO LIKE '*" + inp + "*' order by assignedto,datereceived,iteration")
End If
lastdep$ = ""
ps = 0
If Not ds.EOF Then
    ds.MoveFirst
    lastdep$ = ds("assignedto")
Else
    NOREPORT = 1
    db.Close
    Exit Sub
End If
While Not ds.EOF
        Set rs1 = db.OpenRecordset("select * from opbor")
        rs1.AddNew
        rs1("Officer") = ds("assignedto")
        rs1("Date Received") = ds("datereceived")
        rs1("Service Of") = ds("serviceof")
        rs1("Iteration") = ds("iteration")
        rs1("Plaintiff") = ds("plaintiff")
        rs1("Paper Type") = ds("paperType")
        rs1("Court Date") = ds("courtdate")
        rs1.Update
        lastdep$ = ds("assignedto")
    ds.MoveNext
Wend
db.Close
Exit Sub
Case "ompborrtn"
Set rs1 = db.OpenRecordset("select * from opbor")
If Not rs1.EOF Then
    rs1.MoveFirst
End If
While Not rs1.EOF
    rs1.Delete
    rs1.MoveNext
Wend
If inp = "A" Then
    Set ds = db.OpenRecordset("select datereceived,iteration,casenumber,serviceof,plaintiff,papertype,courtdate,served,nonservice,assignedto from magistrate where assignedon <= #" + Date$ + "# and served <> '1' and nonservice <> '1' order by assignedto,datereceived,iteration")
Else
    Set ds = db.OpenRecordset("select datereceived,iteration,casenumber,serviceof,plaintiff,papertype,courtdate,served,nonservice,assignedto from magistrate where assignedon <= #" + Date$ + "# and served <> '1' and nonservice <> '1' AND ASSIGNEDTO LIKE '*" + inp + "*' order by assignedto,datereceived,iteration")
End If
lastdep$ = ""
ps = 0
If Not ds.EOF Then
    ds.MoveFirst
    lastdep$ = ds("assignedto")
Else
    NOREPORT = 1
    db.Close
    Exit Sub
End If
While Not ds.EOF
        Set rs1 = db.OpenRecordset("select * from opbor")
        rs1.AddNew
        rs1("Officer") = ds("assignedto")
        rs1("Date Received") = ds("datereceived")
        rs1("Service Of") = ds("serviceof")
        rs1("Iteration") = ds("iteration")
        rs1("Plaintiff") = ds("plaintiff")
        rs1("Paper Type") = ds("paperType")
        rs1("Court Date") = ds("courtdate")
        rs1.Update
        lastdep$ = ds("assignedto")
    ds.MoveNext
Wend
db.Close
Exit Sub
Case "oepborrtn"
Set rs1 = db.OpenRecordset("select * from opbor")
If Not rs1.EOF Then
    rs1.MoveFirst
End If
While Not rs1.EOF
    rs1.Delete
    rs1.MoveNext
Wend
If inp = "A" Then
    Set ds = db.OpenRecordset("select datereceived,iteration,casenumber,serviceof,plaintiff,papertype,courtdate,served,nonservice,assignedto from executions where assignedon <= #" + Date$ + "# and served <> '1' and nonservice <> '1' order by assignedto,datereceived,iteration")
Else
    Set ds = db.OpenRecordset("select datereceived,iteration,casenumber,serviceof,plaintiff,papertype,courtdate,served,nonservice,assignedto from executions where assignedon <= #" + Date$ + "# and served <> '1' and nonservice <> '1' AND ASSIGNEDTO LIKE '*" + inp + "*' order by assignedto,datereceived,iteration")
End If
lastdep$ = ""
ps = 0
If Not ds.EOF Then
    ds.MoveFirst
    lastdep$ = ds("assignedto")
Else
    NOREPORT = 1
    db.Close
    Exit Sub
End If
While Not ds.EOF
        Set rs1 = db.OpenRecordset("select * from opbor")
        rs1.AddNew
        rs1("Officer") = ds("assignedto")
        rs1("Date Received") = ds("datereceived")
        rs1("Service Of") = ds("serviceof")
        rs1("Iteration") = ds("iteration")
        rs1("Plaintiff") = ds("plaintiff")
        rs1("Paper Type") = ds("paperType")
        rs1("Court Date") = ds("courtdate")
        rs1.Update
        lastdep$ = ds("assignedto")
    ds.MoveNext
Wend
db.Close
Exit Sub
Case "ofcpborrtn"
Set rs1 = db.OpenRecordset("select * from opbor")
If Not rs1.EOF Then
    rs1.MoveFirst
End If
While Not rs1.EOF
    rs1.Delete
    rs1.MoveNext
Wend
If inp = "A" Then
    Set ds = db.OpenRecordset("select datereceived,iteration,casenumber,serviceof,plaintiff,papertype,courtdate,served,nonservice,assignedto from familycourt where assignedon <= #" + Date$ + "# and served <> '1' and nonservice <> '1' order by assignedto,datereceived,iteration")
Else
    Set ds = db.OpenRecordset("select datereceived,iteration,casenumber,serviceof,plaintiff,papertype,courtdate,served,nonservice,assignedto from familycourt where assignedon <= #" + Date$ + "# and served <> '1' and nonservice <> '1' AND ASSIGNEDTO LIKE '*" + inp + "*' order by assignedto,datereceived,iteration")
End If
lastdep$ = ""
ps = 0
If Not ds.EOF Then
    ds.MoveFirst
    lastdep$ = ds("assignedto")
Else
    NOREPORT = 1
    db.Close
    Exit Sub
End If
While Not ds.EOF
        Set rs1 = db.OpenRecordset("select * from opbor")
        rs1.AddNew
        rs1("Officer") = ds("assignedto")
        rs1("Date Received") = ds("datereceived")
        rs1("Service Of") = ds("serviceof")
        rs1("Iteration") = ds("iteration")
        rs1("Plaintiff") = ds("plaintiff")
        rs1("Paper Type") = ds("paperType")
        rs1("Court Date") = ds("courtdate")
        rs1.Update
        lastdep$ = ds("assignedto")
    ds.MoveNext
Wend
db.Close
Exit Sub
Case "owpborrtn"
Set rs1 = db.OpenRecordset("select * from opbor")
If Not rs1.EOF Then
    rs1.MoveFirst
End If
While Not rs1.EOF
    rs1.Delete
    rs1.MoveNext
Wend
If inp = "A" Then
    Set ds = db.OpenRecordset("select datereceived,iteration,casenumber,serviceof,plaintiff,papertype,courtdate,served,nonservice,assignedto from writother where assignedon <= #" + Date$ + "# and served <> '1' and nonservice <> '1' order by assignedto,datereceived,iteration")
Else
    Set ds = db.OpenRecordset("select datereceived,iteration,casenumber,serviceof,plaintiff,papertype,courtdate,served,nonservice,assignedto from writother where assignedon <= #" + Date$ + "# and served <> '1' and nonservice <> '1' AND ASSIGNEDTO LIKE '*" + inp + "*'  order by assignedto,datereceived,iteration")
End If
    
lastdep$ = ""
ps = 0
If Not ds.EOF Then
    ds.MoveFirst
    lastdep$ = ds("assignedto")
Else
    NOREPORT = 1
    db.Close
    Exit Sub
End If
While Not ds.EOF
        Set rs1 = db.OpenRecordset("select * from opbor")
        rs1.AddNew
        rs1("Officer") = ds("assignedto")
        rs1("Date Received") = ds("datereceived")
        rs1("Service Of") = ds("serviceof")
        rs1("Iteration") = ds("iteration")
        rs1("Plaintiff") = ds("plaintiff")
        rs1("Paper Type") = ds("paperType")
        rs1("Court Date") = ds("courtdate")
        rs1.Update
        lastdep$ = ds("assignedto")
    ds.MoveNext
Wend
db.Close
Exit Sub
End Select
Exit Sub
centerprint:
AutoRedraw = -1
HalfWidth = Printer.TextWidth(lin$) / 2
Printer.CurrentX = Printer.ScaleWidth / 2 - HalfWidth
Printer.Print lin$
Return
rightprint:
AutoRedraw = -1
allWidth = Printer.TextWidth(lin$)
Printer.CurrentX = Printer.ScaleWidth - allWidth - 1000
Printer.Print lin$
Return
oderror1:
If Err > 3200 Then
    Resume od1
Else
    Resume Next
End If

End Sub

Private Sub checkcivil()
Dim db As Database, rs As Recordset
On Error GoTo ldberror
Set db = OpenDatabase(nwl + "lawsuite.mdb")
db.Close
donext:
On Error GoTo cdberror
Set db = OpenDatabase(nwc + "civil.mdb")
db.Close
GETOUT:
Exit Sub
ldberror:
On Error GoTo ldberror2
RepairDatabase (nwl + "lawsuite.mdb")
Resume donext
ldberror2:
msg = MsgBox("Database error occurred.  Exit sotfware and run DBREPAIR utility.", 48, "Genesis Error Log")
End
cdberror:
On Error GoTo cdberror2
RepairDatabase (nwc + "civil.mdb")
Resume GETOUT
cdberror2:
msg = MsgBox("Database error occurred.  Exit sotfware and run DBREPAIR utility.", 48, "Genesis Error Log")
End
End Sub

