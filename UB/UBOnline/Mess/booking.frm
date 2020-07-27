VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form booking 
   BackColor       =   &H00808000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Booking Report"
   ClientHeight    =   7500
   ClientLeft      =   750
   ClientTop       =   405
   ClientWidth     =   10920
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   10920
   WindowState     =   1  'Minimized
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   9000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   600
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   68
      Top             =   0
      Width           =   10920
      _ExtentX        =   19262
      _ExtentY        =   1111
      ButtonWidth     =   1005
      ButtonHeight    =   953
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Clear"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Print"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            ImageIndex      =   3
         EndProperty
      EndProperty
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   9240
         Top             =   480
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   9
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "booking.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "booking.frx":0454
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "booking.frx":08A8
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "booking.frx":0CFC
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "booking.frx":1150
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "booking.frx":15A4
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "booking.frx":18C0
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "booking.frx":1D14
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "booking.frx":2168
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   6825
      Left            =   10575
      TabIndex        =   66
      Top             =   720
      Width           =   250
   End
   Begin VB.PictureBox Picture1 
      Height          =   6840
      Left            =   15
      ScaleHeight     =   6780
      ScaleWidth      =   10515
      TabIndex        =   65
      Top             =   630
      Width           =   10575
      Begin VB.PictureBox PICTURE2 
         Height          =   13000
         Left            =   0
         Picture         =   "booking.frx":25BC
         ScaleHeight     =   12945
         ScaleWidth      =   10395
         TabIndex        =   64
         Top             =   600
         Width           =   10455
         Begin VB.TextBox szipcode 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   7185
            TabIndex        =   21
            Top             =   3360
            Width           =   645
         End
         Begin VB.TextBox sstate 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   6555
            MaxLength       =   2
            TabIndex        =   20
            Top             =   3360
            Width           =   525
         End
         Begin VB.Frame indexframe 
            Caption         =   "Index Frame"
            Height          =   5415
            Left            =   10320
            TabIndex        =   76
            Top             =   9060
            Visible         =   0   'False
            Width           =   8950
            Begin VB.CommandButton Command3 
               Caption         =   "C     L     O     S     E"
               Height          =   255
               Left            =   120
               TabIndex        =   78
               Top             =   5040
               Width           =   8655
            End
            Begin MSComctlLib.ListView indexlist 
               Height          =   4695
               Left            =   120
               TabIndex        =   77
               Top             =   240
               Width           =   8695
               _ExtentX        =   15346
               _ExtentY        =   8281
               View            =   3
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   5
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Incident Number"
                  Object.Width           =   2293
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Defendant Name"
                  Object.Width           =   5292
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "Arrest Number"
                  Object.Width           =   2117
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   3
                  Text            =   "Charge A"
                  Object.Width           =   5292
               EndProperty
               BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   4
                  Text            =   "Subject#"
                  Object.Width           =   0
               EndProperty
            End
         End
         Begin VB.ComboBox ARRESTINGOFFICER 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   4590
            TabIndex        =   33
            Top             =   5280
            Width           =   2505
         End
         Begin VB.ComboBox BOOKINGOFFICER 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   350
            TabIndex        =   31
            Top             =   5280
            Width           =   3345
         End
         Begin VB.ListBox SETHNICITY 
            BackColor       =   &H00808000&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   1020
            TabIndex        =   67
            Top             =   2640
            Width           =   615
         End
         Begin VB.TextBox SHT 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1680
            TabIndex        =   10
            Top             =   2760
            Width           =   855
         End
         Begin VB.ListBox SRACE 
            BackColor       =   &H00808000&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   420
            Left            =   6600
            TabIndex        =   5
            Top             =   2025
            Width           =   855
         End
         Begin VB.ListBox SSEX 
            BackColor       =   &H00808000&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   420
            Left            =   7470
            TabIndex        =   6
            Top             =   2025
            Width           =   870
         End
         Begin VB.TextBox SBIRTHDATE 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   8460
            TabIndex        =   7
            Top             =   2070
            Width           =   960
         End
         Begin VB.CommandButton Command2 
            Caption         =   "I  N  D  E  X"
            Height          =   375
            Left            =   2880
            TabIndex        =   75
            Top             =   1350
            Width           =   1035
         End
         Begin MSDBCtls.DBCombo sname 
            Bindings        =   "booking.frx":5715
            DataSource      =   "Data1"
            Height          =   315
            Left            =   350
            TabIndex        =   4
            Top             =   2040
            Width           =   6135
            _ExtentX        =   10821
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   8421376
            ForeColor       =   16777215
            ListField       =   "dpname"
            Text            =   ""
         End
         Begin VB.ComboBox incidentnumber 
            Height          =   315
            Left            =   3780
            TabIndex        =   2
            Top             =   480
            Width           =   3135
         End
         Begin VB.TextBox timeofarrest 
            BackColor       =   &H00808000&
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
            Height          =   255
            Left            =   2520
            TabIndex        =   1
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox dateofarrest 
            BackColor       =   &H00808000&
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
            Height          =   255
            Left            =   1680
            TabIndex        =   0
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox othercases 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   9
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   56
            Top             =   7320
            Width           =   945
         End
         Begin VB.TextBox othercases 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   8
            Left            =   480
            MaxLength       =   50
            TabIndex        =   55
            Top             =   7320
            Width           =   945
         End
         Begin VB.TextBox othercases 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   7
            Left            =   9000
            MaxLength       =   50
            TabIndex        =   54
            Top             =   6960
            Width           =   945
         End
         Begin VB.TextBox othercases 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   6
            Left            =   8040
            MaxLength       =   50
            TabIndex        =   53
            Top             =   6960
            Width           =   945
         End
         Begin VB.TextBox othercases 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   5
            Left            =   7080
            MaxLength       =   50
            TabIndex        =   52
            Top             =   6960
            Width           =   945
         End
         Begin VB.TextBox othercases 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   4
            Left            =   6120
            MaxLength       =   50
            TabIndex        =   51
            Top             =   6960
            Width           =   945
         End
         Begin VB.TextBox othercases 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   3
            Left            =   5160
            MaxLength       =   50
            TabIndex        =   50
            Top             =   6960
            Width           =   945
         End
         Begin VB.TextBox othercases 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   2
            Left            =   4200
            MaxLength       =   50
            TabIndex        =   49
            Top             =   6960
            Width           =   945
         End
         Begin VB.TextBox othercases 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   1
            Left            =   3240
            MaxLength       =   50
            TabIndex        =   48
            Top             =   6960
            Width           =   945
         End
         Begin VB.TextBox othercases 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   0
            Left            =   2280
            MaxLength       =   50
            TabIndex        =   47
            Top             =   6960
            Width           =   945
         End
         Begin VB.ListBox ucrlist 
            BackColor       =   &H00808000&
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
            Height          =   540
            Index           =   2
            Left            =   7560
            Sorted          =   -1  'True
            Style           =   1  'Checkbox
            TabIndex        =   59
            Top             =   7800
            Width           =   2575
         End
         Begin VB.ListBox ucrlist 
            BackColor       =   &H00808000&
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
            Height          =   540
            Index           =   1
            Left            =   4680
            Sorted          =   -1  'True
            Style           =   1  'Checkbox
            TabIndex        =   58
            Top             =   7800
            Width           =   2575
         End
         Begin VB.ListBox ucrlist 
            BackColor       =   &H00808000&
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
            Height          =   540
            Index           =   0
            Left            =   1965
            Sorted          =   -1  'True
            Style           =   1  'Checkbox
            TabIndex        =   57
            Top             =   7800
            Width           =   2580
         End
         Begin VB.CheckBox armedwithsemiautomatic 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   150
            Index           =   1
            Left            =   5800
            TabIndex        =   40
            Top             =   6000
            Width           =   175
         End
         Begin VB.CheckBox armedwithsemiautomatic 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   150
            Index           =   0
            Left            =   5800
            TabIndex        =   37
            Top             =   5685
            Width           =   175
         End
         Begin VB.ListBox at 
            BackColor       =   &H00808000&
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
            Height          =   480
            Left            =   7455
            Sorted          =   -1  'True
            TabIndex        =   72
            Top             =   1230
            Visible         =   0   'False
            Width           =   2955
         End
         Begin VB.ListBox dt 
            BackColor       =   &H00808000&
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
            Height          =   480
            ItemData        =   "booking.frx":5729
            Left            =   7455
            List            =   "booking.frx":572B
            Sorted          =   -1  'True
            TabIndex        =   70
            Top             =   585
            Visible         =   0   'False
            Width           =   2955
         End
         Begin VB.TextBox arrestnumber 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   4125
            MaxLength       =   20
            TabIndex        =   3
            Top             =   1425
            Width           =   2115
         End
         Begin VB.TextBox agency 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   7200
            MaxLength       =   50
            TabIndex        =   34
            Top             =   5280
            Width           =   1425
         End
         Begin VB.TextBox arrestingunit 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   8760
            MaxLength       =   50
            TabIndex        =   35
            Top             =   5280
            Width           =   1545
         End
         Begin VB.ListBox sresident 
            BackColor       =   &H00808000&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   420
            Left            =   7890
            TabIndex        =   22
            Top             =   3300
            Width           =   1140
         End
         Begin VB.TextBox statutec 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   7560
            MaxLength       =   50
            TabIndex        =   62
            Top             =   8340
            Width           =   2575
         End
         Begin VB.TextBox statuteb 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   4680
            MaxLength       =   50
            TabIndex        =   61
            Top             =   8340
            Width           =   2575
         End
         Begin VB.TextBox statutea 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1965
            MaxLength       =   50
            TabIndex        =   60
            Top             =   8340
            Width           =   2575
         End
         Begin VB.CheckBox within 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Under 18 - Within Dept"
            ForeColor       =   &H00808000&
            Height          =   200
            Left            =   1680
            TabIndex        =   45
            Top             =   6600
            Width           =   2055
         End
         Begin VB.CheckBox referred 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Under 18 - Referred"
            ForeColor       =   &H00808000&
            Height          =   200
            Left            =   3720
            TabIndex        =   46
            Top             =   6600
            Width           =   2175
         End
         Begin VB.CheckBox taken 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Custody"
            ForeColor       =   &H00808000&
            Height          =   200
            Left            =   9480
            TabIndex        =   44
            Top             =   6000
            Width           =   875
         End
         Begin VB.CheckBox summoned 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Summoned"
            ForeColor       =   &H00808000&
            Height          =   200
            Left            =   8280
            TabIndex        =   43
            Top             =   6000
            Width           =   1215
         End
         Begin VB.CheckBox onviewarrest 
            BackColor       =   &H00FFFFFF&
            Caption         =   "On View Arrest"
            ForeColor       =   &H00808000&
            Height          =   200
            Left            =   6840
            TabIndex        =   42
            Top             =   6000
            Width           =   1455
         End
         Begin VB.TextBox bookingofficerunit 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   3840
            MaxLength       =   50
            TabIndex        =   32
            Top             =   5280
            Width           =   705
         End
         Begin VB.TextBox employer 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   350
            MaxLength       =   50
            TabIndex        =   28
            Top             =   4680
            Width           =   3105
         End
         Begin VB.TextBox nextofkinaddress 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   6240
            MaxLength       =   60
            TabIndex        =   30
            Top             =   4680
            Width           =   4095
         End
         Begin VB.TextBox driverslicensestate 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   9480
            MaxLength       =   2
            TabIndex        =   27
            Top             =   3960
            Width           =   795
         End
         Begin VB.TextBox driverslicense 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   5640
            MaxLength       =   20
            TabIndex        =   26
            Top             =   3960
            Width           =   3555
         End
         Begin VB.TextBox nextofkin 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   3600
            MaxLength       =   60
            TabIndex        =   29
            Top             =   4680
            Width           =   2535
         End
         Begin VB.TextBox birthplace 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   3960
            MaxLength       =   20
            TabIndex        =   25
            Top             =   3960
            Width           =   1450
         End
         Begin VB.TextBox alias 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   350
            MaxLength       =   60
            TabIndex        =   24
            Top             =   3960
            Width           =   3495
         End
         Begin VB.TextBox phone 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   9090
            MaxLength       =   20
            TabIndex        =   23
            Top             =   3360
            Width           =   1250
         End
         Begin VB.TextBox idnumber 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   9400
            MaxLength       =   20
            TabIndex        =   17
            Top             =   2760
            Width           =   915
         End
         Begin VB.TextBox ncic 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   8760
            MaxLength       =   20
            TabIndex        =   16
            Top             =   2760
            Width           =   555
         End
         Begin VB.TextBox ssn 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   5520
            MaxLength       =   12
            TabIndex        =   14
            Top             =   2760
            Width           =   1595
         End
         Begin VB.TextBox docketnumber 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   9570
            MaxLength       =   20
            TabIndex        =   8
            Top             =   2040
            Width           =   750
         End
         Begin VB.CheckBox armedwithautomatic 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   150
            Index           =   0
            Left            =   5800
            TabIndex        =   38
            Top             =   5835
            Width           =   175
         End
         Begin VB.CheckBox armedwithautomatic 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Check1"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   150
            Index           =   1
            Left            =   5800
            TabIndex        =   41
            Top             =   6150
            Width           =   175
         End
         Begin VB.ListBox armedlist 
            BackColor       =   &H00808000&
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
            Height          =   270
            Index           =   1
            Left            =   3120
            Sorted          =   -1  'True
            TabIndex        =   39
            Top             =   6000
            Width           =   2595
         End
         Begin VB.ListBox armedlist 
            BackColor       =   &H00808000&
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
            Height          =   270
            Index           =   0
            Left            =   3120
            Sorted          =   -1  'True
            TabIndex        =   36
            Top             =   5730
            Width           =   2595
         End
         Begin RichTextLib.RichTextBox REMARKS 
            Height          =   3510
            Left            =   360
            TabIndex        =   63
            Top             =   8880
            Width           =   9855
            _ExtentX        =   17383
            _ExtentY        =   6191
            _Version        =   393217
            TextRTF         =   $"booking.frx":572D
         End
         Begin VB.TextBox scity 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   4950
            TabIndex        =   19
            Top             =   3360
            Width           =   1530
         End
         Begin VB.TextBox saddress 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   350
            TabIndex        =   18
            Top             =   3360
            Width           =   4455
         End
         Begin VB.TextBox speculiarities 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   7155
            TabIndex        =   15
            Top             =   2760
            Width           =   1545
         End
         Begin VB.TextBox SEYES 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   4560
            TabIndex        =   13
            Top             =   2760
            Width           =   855
         End
         Begin VB.TextBox SHAIR 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   3600
            TabIndex        =   12
            Top             =   2760
            Width           =   855
         End
         Begin VB.TextBox SWEIGHT 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   2640
            TabIndex        =   11
            Top             =   2760
            Width           =   855
         End
         Begin VB.TextBox sage 
            BackColor       =   &H00808000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   350
            TabIndex        =   9
            Top             =   2760
            Width           =   615
         End
         Begin VB.CommandButton SpellCk 
            Caption         =   "Spelling"
            Height          =   255
            Left            =   8745
            TabIndex        =   79
            Top             =   8640
            Width           =   1470
         End
         Begin VB.Image mugshot 
            BorderStyle     =   1  'Fixed Single
            Height          =   1770
            Left            =   15
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1605
         End
         Begin VB.Label WEAPONS 
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   2040
            TabIndex        =   74
            Top             =   8640
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label atlabel 
            BackStyle       =   0  'Transparent
            Caption         =   "Activity Connected to Drug Type"
            ForeColor       =   &H00808000&
            Height          =   255
            Left            =   7455
            TabIndex        =   73
            Top             =   1035
            Visible         =   0   'False
            Width           =   2775
         End
         Begin VB.Label dtlabel 
            BackStyle       =   0  'Transparent
            Caption         =   "Most Serious Drug Type"
            ForeColor       =   &H00808000&
            Height          =   255
            Left            =   7455
            TabIndex        =   71
            Top             =   390
            Visible         =   0   'False
            Width           =   2295
         End
      End
   End
   Begin Crystal.CrystalReport REPORT 
      Left            =   0
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      Destination     =   1
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label subjectnumber 
      Caption         =   "Label1"
      Height          =   135
      Left            =   840
      TabIndex        =   69
      Top             =   1080
      Width           =   615
   End
End
Attribute VB_Name = "booking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PICFILE, schanged As Integer, itmx As ListItem, pics(5000), incs(5000), nams(5000) As String, inct As Integer, painted As Integer, nametype As Integer
'Dim fromIncident As Boolean
Private Sub Clearbooking_Click()
End Sub

Private Sub arrestingunit_GotFocus()
If arrestingunit > "" Or ARRESTINGOFFICER = "" Then
    Exit Sub
End If
Dim db As DAO.Database, rs As DAO.Recordset
Set db = DAO.OpenDatabase(nwl + "LAWSUITE.MDB")
Set rs = db.OpenRecordset("SELECT PROFIDNUM FROM PROFESSIONALS WHERE PROFNAME = '" + ARRESTINGOFFICER + "' AND TYPE = 'D'")
If Not rs.EOF Then
    rs.MoveFirst
    If Not IsNull(rs("PROFIDNUM")) Then
        arrestingunit = rs("PROFIDNUM")
    End If
End If
db.Close
End Sub

Private Sub bookingofficerunit_GotFocus()
If bookingofficerunit > "" Or BOOKINGOFFICER = "" Then
    Exit Sub
End If
Dim db As DAO.Database, rs As DAO.Recordset
Set db = DAO.OpenDatabase(nwl + "LAWSUITE.MDB")
Set rs = db.OpenRecordset("SELECT PROFIDNUM FROM PROFESSIONALS WHERE PROFNAME = '" + BOOKINGOFFICER + "' AND TYPE = 'D'")
If Not rs.EOF Then
    rs.MoveFirst
    If Not IsNull(rs("PROFIDNUM")) Then
        bookingofficerunit = rs("PROFIDNUM")
    End If
End If
db.Close
End Sub

Private Sub Command3_Click()
indexframe.Visible = False
Command2.SetFocus
End Sub


Private Sub Command1_Click()
paframe.Visible = False
End Sub

Private Sub Command2_Click()
Dim db As Database, rs As Recordset
Set db = OpenDatabase(nwb + "booking.mdb")
Set rs = db.OpenRecordset("select incidentnumber, arrestnumber, sname, number, chargea from booking")
indexlist.ListItems.clear
If Not rs.EOF Then
    rs.MoveFirst
    While Not rs.EOF
        Set itmx = indexlist.ListItems.add(, , rs("incidentnumber"))
        itmx.SubItems(1) = rs("sname")
        itmx.SubItems(2) = rs("arrestnumber")
        itmx.SubItems(3) = rs("chargea")
        If Not IsNull(rs("number")) Then
            itmx.SubItems(4) = rs("number")
        End If
        rs.MoveNext
    Wend
End If
db.Close
indexframe.Top = 100
indexframe.Left = 200
indexframe.Visible = True
indexlist.SetFocus

End Sub

Private Sub Command5_Click()
loframe.Visible = False
End Sub




Private Sub DATEOFARREST_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
    If Len(dateofarrest) = 1 Or Len(dateofarrest) = 4 Then
        Call sendslash
    End If
End If

End Sub
Private Sub sendslash()
SendKeys "/"
End Sub


Private Sub ethnicity_ItemClick(ByVal Item As MSComctlLib.ListItem)
End Sub

Private Sub eyes_ItemClick(ByVal Item As MSComctlLib.ListItem)
End Sub


Private Sub Form_Load()
fromexport = False
For t% = 0 To Forms.Count - 1
    If Forms(t%).Name = "iexport" And Forms(t%).Visible = True Then
        fromexport = True
        t% = Forms.Count - 1
    End If
Next t%

painted = 0
nametype = 1
Picture1.Height = PICTURE2.Height
VScroll1.Max = Picture1.Height
'VScroll1.Max = Picture2.Height - Picture1.Height
VScroll1.LargeChange = VScroll1.Max / 10
VScroll1.SmallChange = VScroll1.Max / 100
PICFILE = ""
On Error Resume Next
On Error GoTo 0
Dim db As Database, rs As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nwb + "booking.mdb")
Set rs = db.OpenRecordset("select office, orinumber from system")
On Error Resume Next
If Not rs.EOF Then
    rs.MoveFirst
    orinumber = rs("orinumber")
    office = rs("office")
End If
db.Close
Call defaultcodes
Call loadcodes
Call loadlist
Call LoadOfficers
If Not fromexport Then
    Me.Height = 7900
    Me.Width = 11040
    Me.Top = 0
    Me.Left = 0
    Me.WindowState = vbNormal
End If
Screen.MousePointer = 0
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If
Exit Sub
End Sub

Private Sub hair_ItemClick(ByVal Item As MSComctlLib.ListItem)

End Sub

Private Sub Form_Paint()
painted = painted + 1
If painted = 1 Then
    If Dir("C:\TOBOOKING") > "" Then
        
        Open "C:\TOBOOKING" For Input As #1
        Line Input #1, a$
        booking.incidentnumber = a$
        Line Input #1, a$
        subjectnumber = a$
        txtSubject = subjectnumber
        Line Input #1, a$
        dateofarrest = a$
        Line Input #1, a$
        timeofarrest = a$
        Line Input #1, a$
        sname = a$
        Line Input #1, a$
        For t% = 0 To SRACE.ListCount - 1
            If Left$(SRACE.List(t%), Len(a$)) = a$ Then
                SRACE.ListIndex = t%
                t% = SRACE.ListCount
            End If
        Next t%
        Line Input #1, a$
        For t% = 0 To SSEX.ListCount - 1
            If Left$(SSEX.List(t%), Len(a$)) = a$ Then
                SSEX.ListIndex = t%
                t% = SSEX.ListCount
            End If
        Next t%
        Line Input #1, a$
        SBIRTHDATE = a$
        Line Input #1, a$
        sage = a$
        Line Input #1, a$
        For t% = 0 To SETHNICITY.ListCount - 1
            If Left$(SETHNICITY.List(t%), Len(a$)) = a$ Then
                SETHNICITY.ListIndex = t%
                t% = SETHNICITY.ListCount
            End If
        Next t%
        Line Input #1, a$
        SHT = a$
        Line Input #1, a$
        SWEIGHT = a$
        Line Input #1, a$
        SHAIR = a$
        Line Input #1, a$
        SEYES = a$
        Line Input #1, a$
        speculiarities = a$
        Line Input #1, a$
        saddress = a$
        Line Input #1, a$
        scity = a$
        Line Input #1, a$
        sstate = a$
        Line Input #1, a$
        szipcode = a$
        Close #1
        Kill "C:\TOBOOKING"
    Else
        txtSubject = ""
    End If
    
    If incidentnumber > "" Then
        Call findincident
        arrestnumber.SetFocus
    End If

End If
fromIncident = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set booking = Nothing
End Sub

Private Sub incidentnumber_Click()
Call findincident

End Sub

Private Sub indexlist_ItemClick(ByVal Item As MSComctlLib.ListItem)
Set itmx = indexlist.ListItems(indexlist.SelectedItem.Index)
incidentnumber = itmx
arrestnumber = itmx.SubItems(2)
sname = itmx.SubItems(1)
subjectnumber = itmx.SubItems(4)
Call findincident
Call Command3_Click
End Sub

Private Sub race_ItemClick(ByVal Item As MSComctlLib.ListItem)

End Sub


Private Sub othercases_LostFocus(Index As Integer)
If Index = 0 And othercases(Index) = "" Then
    ucrlist(0).SetFocus
End If
End Sub


Private Sub remarks_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Shift = vbCtrlMask) And (KeyCode = vbKeyF2) Then
        Call SpellCk_Click
    End If
End Sub



Private Sub sbirthdate_Change()
If IsDate(SBIRTHDATE) Then
    sage = DateDiff("yyyy", CDate(SBIRTHDATE), CDate(Date$))
End If
End Sub

Private Sub sbirthdate_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
    If Len(SBIRTHDATE) = 1 Or Len(SBIRTHDATE) = 4 Then
        SendKeys "/"
    End If
End If

End Sub

Private Sub SETHNICITY_GotFocus()
SETHNICITY.Height = 765
SETHNICITY.Width = 2565

End Sub

Private Sub SETHNICITY_LostFocus()
SETHNICITY.Height = 255
SETHNICITY.Width = 615

End Sub

Private Sub sex_ItemClick(ByVal Item As MSComctlLib.ListItem)

End Sub

Private Sub SHT_Change()

End Sub

Private Sub SHT_GotFocus()
End Sub

Private Sub sname_Click(Area As Integer)
Call findincident
Call setpopup(sname, "L")
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

Private Sub SpellCk_Click()
BeginSpellCheck REMARKS.Text, REMARKS
End Sub

Private Sub SRACE_GotFocus()
SRACE.Height = 765
SRACE.Width = 2565

End Sub

Private Sub SRACE_LostFocus()
SRACE.Height = 450
SRACE.Width = 855
End Sub

Private Sub SSEX_GotFocus()
SSEX.Height = 765
SSEX.Width = 2565

End Sub

Private Sub SSEX_LostFocus()
SSEX.Height = 450
SSEX.Width = 855

End Sub


Private Sub TIMEOFARREST_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
If Len(timeofarrest) = 1 Then
    SendKeys ":"
End If
End If

End Sub

Private Sub nextofkin_LostFocus()
If nextofkin > "" And InStr(nextofkin, ",") = 0 Then
    msg = MsgBox("All names in the Booking Report system should be entered in the format last name + comma + firstname.", 48, "Invalid Data Format")
'    nextofkin.SetFocus
End If
End Sub

Private Sub othercases_GotFocus(Index As Integer)
If Me.ActiveControl.Top > (-1 * PICTURE2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * PICTURE2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    VScroll1 = Me.ActiveControl.Top - 500
Else
    VScroll1 = 0
End If
End If


End Sub

Private Sub remarks_GotFocus()
If Me.ActiveControl.Top > (-1 * PICTURE2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * PICTURE2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    VScroll1 = Me.ActiveControl.Top - 500
Else
    VScroll1 = 0
End If
End If

End Sub

Private Sub statutea_GotFocus()
If Me.ActiveControl.Top > (-1 * PICTURE2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * PICTURE2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    VScroll1 = Me.ActiveControl.Top - 500
Else
    VScroll1 = 0
End If
End If

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Screen.MousePointer = 11
Select Case Button
    Case "Clear"
        Call clearroutine
        
    
    Case "Delete"
        If UCase(frmLogin.txtUserName) = "DEMO" And UCase(frmLogin.txtPassword) = "DEMO" Then
            msg = MsgBox("Not available in DEMO version.", 48, "Genesis Information Log")
            Screen.MousePointer = 0
            Exit Sub
        End If
        Call deleteroutine

    Case "Print"
        editerr% = 0
        POPMSG$ = ""
        Call editroutine(editerr%, POPMSG$)
        If editerr% = 0 Then
            Call saveroutine
            REPORT.ReportFileName = nwb + "BOOKING.RPT"
            REPORT.SelectionFormula = "{booking.INCIDENTNUMBER} = '" + incidentnumber + "' and {booking.sname} = " + Chr$(34) + sname + Chr$(34)
            REPORT.Action = 1
        End If

    Case "Exit"
        Open "pp.tag" For Output As #1
        Print #1, incidentnumber
        Close #1
        Unload booking
        incident.WindowState = vbMaximized
        incident.Show
        
    Case "Save"
        If UCase(frmLogin.txtUserName) = "DEMO" And UCase(frmLogin.txtPassword) = "DEMO" Then
            msg = MsgBox("Not available in DEMO version.", 48, "Genesis Information Log")
            Screen.MousePointer = 0
            Exit Sub
        End If
        editerr% = 0
        POPMSG$ = ""
        Call editroutine(editerr%, POPMSG$)
        If editerr% = 0 Then
            Call saveroutine
        Else
            MsgBox POPMSG$, 48, "Genesis Error Log"
        End If
        
End Select
Screen.MousePointer = 0
End Sub

Private Sub ucrlist_Click(Index As Integer)
found35a = False
For a% = 0 To 2
    For t% = 0 To ucrlist(a%).ListCount - 1
        If ucrlist(a%).Selected(t%) Then
            If InStr(ucrlist(a%).List(t%), "35A") > 0 Then
                found35a = True
                t% = ucrlist(a%).ListCount - 1
            End If
        End If
    Next t%
Next a%
If found35a Then
    dt.Visible = True
    at.Visible = True
Else
    dt.Visible = False
    at.Visible = False
End If
End Sub

Private Sub ucrlist_GotFocus(Index As Integer)
If Me.ActiveControl.Top > (-1 * PICTURE2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * PICTURE2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    If VScroll1.Max > Me.ActiveControl.Top - 500 Then
        VScroll1 = Me.ActiveControl.Top - 500
    Else
        VScroll1 = VScroll1.Max
    End If
Else
    VScroll1 = 0
End If
End If

End Sub

Private Sub VScroll1_Change()
PICTURE2.Top = -VScroll1.Value
End Sub
Friend Sub saveroutine()
Dim db, db2 As Database, rs, rs2 As Recordset, ab(2) As String
Set db = OpenDatabase(nwb + "booking.MDB")
Set db2 = OpenDatabase(nwi + "incident.mdb")
schanged = 0
On Error GoTo oderror
od:
Set rs = db.OpenRecordset("select * from booking where incidentnumber = " + Chr$(34) + incidentnumber + Chr$(34) + " and sname = " + Chr$(34) + sname + Chr$(34))
If rs.EOF Then
    rs.AddNew
    rs("original") = Date$
    schanged = 1
Else
    rs.MoveFirst
    rs.Edit
End If
If rs("schanged") = 1 Then
    schanged = 1
End If
On Error Resume Next
rs("lastupdate") = Date$
rs("INCIDENTNUMBER") = incidentnumber
If rs("arrestnumber") <> arrestnumber Then
    schanged = 1
End If
rs("arrestnumber") = arrestnumber
rs("sname") = sname
If rs("srace") <> Left$(SRACE, 1) Then
    schanged = 1
End If
If rs("ssex") <> Left$(SSEX, 1) Then
    schanged = 1
End If
rs("srace") = Left$(SRACE, 1)
rs("ssex") = Left$(SSEX, 1)
If IsDate(SBIRTHDATE) Then
    rs("sbirthdate") = SBIRTHDATE
Else
    rs("SBIRTHDATE") = Null
End If
If rs("sage") <> sage Then
    schanged = 1
End If
If rs("sethnicity") <> Left$(SETHNICITY, 1) Then
    schanged = 1
End If
rs("number") = Val(subjectnumber)
rs("sage") = sage
rs("sethnicity") = Left$(SETHNICITY, 1)
rs("sheight") = SHT
rs("sWeight") = SWEIGHT
rs("shair") = SHAIR
rs("seyes") = SEYES
rs("speculiarities") = speculiarities
rs("saddress") = saddress
rs("scity") = scity
rs("sstate") = sstate
rs("szipcode") = szipcode
rs("docketnumber") = docketnumber
rs("ssn") = ssn
rs("ncic") = ncic
rs("idnumber") = idnumber
If rs("dateofarrest") <> dateofarrest Then
    schanged = 1
End If
If rs("timeofarrest") <> timeofarrest Then
    schanged = 1
End If
rs("dateofarrest") = dateofarrest
If rs("sresident") <> Left$(sresident.List(sresident.ListIndex), 1) Then
    schanged = 1
End If
If sresident.ListIndex > -1 Then
    rs("sresident") = Left$(sresident.List(sresident.ListIndex), 1)
Else
    rs("sresident") = ""
End If
rs("phone") = phone
rs("alias") = alias
rs("birthplace") = birthplace
rs("driverslicense") = driverslicense
rs("driverslicensestate") = driverslicensestate
rs("employer") = employer
rs("nextofkin") = nextofkin
rs("nextofkinaddress") = nextofkinaddress
rs("bookingofficer") = BOOKINGOFFICER
rs("bookingofficerUNIT") = bookingofficerunit
rs("arrestingofficer") = ARRESTINGOFFICER
rs("arrestingunit") = arrestingunit
rs("agency") = agency
If rs("ARMEDWITH1") <> Mid$(armedlist(0).List(armedlist(0).ListIndex), InStr(armedlist(0).List(armedlist(0).ListIndex), "(") + 1, 2) Then
    schanged = 1
End If
If armedlist(0).ListIndex > -1 Then
    rs("ARMEDWITH1") = Mid$(armedlist(0).List(armedlist(0).ListIndex), InStr(armedlist(0).List(armedlist(0).ListIndex), "(") + 1, 2)
Else
    rs("ARMEDWITH1") = ""
End If
If rs("ARMEDWITH2") <> Mid$(armedlist(1).List(armedlist(1).ListIndex), InStr(armedlist(1).List(armedlist(1).ListIndex), "(") + 1, 2) Then
    schanged = 1
End If
If armedlist(1).ListIndex > -1 Then
    rs("ARMEDWITH2") = Mid$(armedlist(1).List(armedlist(1).ListIndex), InStr(armedlist(1).List(armedlist(1).ListIndex), "(") + 1, 2)
Else
    rs("ARMEDWITH2") = ""
End If
If rs("ARMEDWITHAUTOMATIC1") <> armedwithautomatic(0) Then
    schanged = 1
End If
If rs("ARMEDWITHAUTOMATIC2") <> armedwithautomatic(1) Then
    schanged = 1
End If
If rs("ARMEDWITHSEMIAUTOMATIC1") <> armedwithsemiautomatic(0) Then
    schanged = 1
End If
If rs("ARMEDWITHSEMIAUTOMATIC2") <> armedwithsemiautomatic(1) Then
    schanged = 1
End If
rs("ARMEDWITHAUTOMATIC1") = armedwithautomatic(0)
rs("ARMEDWITHAUTOMATIC2") = armedwithautomatic(1)
rs("ARMEDWITHsemiAUTOMATIC1") = armedwithsemiautomatic(0)
rs("ARMEDWITHsemiAUTOMATIC2") = armedwithsemiautomatic(1)
If rs("ONVIEW") <> onviewarrest.Value Then
    schanged = 1
End If
If rs("SUMMONED") <> summoned.Value Then
    schanged = 1
End If
If rs("TAKEN") <> taken.Value Then
    schanged = 1
End If
If rs("WITHIN") <> within Then
    schanged = 1
End If
If rs("REFERRED") <> referred Then
    schanged = 1
End If
rs("ONVIEW") = onviewarrest.Value
rs("SUMMONED") = summoned.Value
rs("TAKEN") = taken.Value
rs("WITHIN") = within
rs("REFERRED") = referred
For p% = 1 To 10
    If rs("OTHERCASES" + Mid$(Str$(p%), 2)) <> othercases(p% - 1) Then
        schanged = 1
    End If
    rs("OTHERCASES" + Mid$(Str$(p%), 2)) = othercases(p% - 1)
Next p%
If rs("CHARGEA") <> ucrlist(0).List(ucrlist(0).ListIndex) Then
    schanged = 1
End If
rs("CHARGEA") = ucrlist(0).List(ucrlist(0).ListIndex)
If ucrlist(1).ListIndex > -1 Then
    If rs("CHARGEb") <> ucrlist(1).List(ucrlist(1).ListIndex) Then
        schanged = 1
    End If
    rs("CHARGEb") = ucrlist(1).List(ucrlist(1).ListIndex)
Else
    rs("CHARGEb") = Null
End If
If ucrlist(2).ListIndex > -1 Then
    If rs("CHARGEc") <> ucrlist(2).List(ucrlist(2).ListIndex) Then
        schanged = 1
    End If
    rs("CHARGEc") = ucrlist(2).List(ucrlist(2).ListIndex)
Else
    rs("CHARGEc") = Null
End If
rs("STATUTEA") = statutea
rs("STATUTEB") = statuteb
rs("STATUTEC") = statutec
rs("REMARKS") = REMARKS.Text
If dt.Visible = True Then
    If rs("dt") <> Left$(dt.List(dt.ListIndex), 1) Then
        schanged = 1
    End If
    rs("dt") = Left$(dt.List(dt.ListIndex), 1)
Else
    rs("dt") = Null
End If
If at.Visible = True Then
    If rs("at") <> Mid$(at.List(at.ListIndex), InStr(at.List(at.ListIndex), "(") + 1, 1) Then
        schanged = 1
    End If
    rs("at") = Mid$(at.List(at.ListIndex), InStr(at.List(at.ListIndex), "(") + 1, 1)
Else
    rs("at") = Null
End If
rs("schanged") = schanged
On Error GoTo 0
Set rs2 = db2.OpenRecordset("select abgroup from ucr where code = " + Chr$(34) + ucrlist(0).List(ucrlist(0).ListIndex) + Chr$(34))
If Not rs2.EOF Then
    rs2.MoveFirst
    rs("bgroup") = rs2("abgroup")
Else
    rs("bgroup") = Null
End If
rs.Update
Set db = OpenDatabase(nwl + "lawsuite.mdb")

'----- OFFICERS
If BOOKINGOFFICER > "" Then
    Set rs = db.OpenRecordset("select profidnum,profname, type from professionals where profname =" + Chr$(34) + BOOKINGOFFICER + Chr$(34))
    If rs.EOF Then
        rs.AddNew
    Else
        rs.MoveFirst
        rs.Edit
    End If
    rs("profname") = BOOKINGOFFICER
    rs("profidnum") = bookingofficerunit
    If rs.EOF Then
        BOOKINGOFFICER.AddItem BOOKINGOFFICER
        ARRESTINGOFFICER.AddItem BOOKINGOFFICER
    End If
    rs("type") = "D"
    rs.Update
End If
If ARRESTINGOFFICER > "" Then
    Set rs = db.OpenRecordset("select profidnum,profname, type from professionals where profname =" + Chr$(34) + ARRESTINGOFFICER + Chr$(34))
    If rs.EOF Then
        rs.AddNew
    Else
        rs.MoveFirst
        rs.Edit
    End If
    rs("profname") = ARRESTINGOFFICER
    rs("profidnum") = arrestingunit
    If rs.EOF Then
'        reportingofficer.AddItem ARRESTINGOFFICER
        ARRESTINGOFFICER.AddItem ARRESTINGOFFICER
    End If
    rs("type") = "D"
    rs.Update
End If

'-----PEOPLE
Set rs = db.OpenRecordset("select * from people where dpnamelf =" + Chr$(34) + sname + Chr$(34))
If rs.EOF Then
    rs.AddNew
Else
    rs.MoveFirst
    rs.Edit
End If
rs("dpnamelf") = sname
rs("dphaddress") = saddress
rs("dphaddress2") = scity + ", " + sstate + " " + szipcode
rs("dpsort") = Left$(sname, 15)
If phone > "" Then
    rs("dphphone") = phone
    rs("resident") = Left$(sresident.List(sresident.ListIndex), 1)
End If
rs("HEIGHT") = SHT
rs("WEIGHT") = SWT
rs("HAIR") = SHAIR
rs("EYES") = SEYES
rs("PECULIARITIES") = speculiarities
rs("race") = Left$(SRACE.List(SRACE.ListIndex), 1)
rs("sex") = Left$(SSEX.List(SSEX.ListIndex), 1)
rs("age") = sage
rs("ethnicity") = Left$(SETHNICITY.List(SETHNICITY.ListIndex), 1)
rs("ssn") = ssn
rs("idnumber") = idnumber
rs("alias") = alias
rs("dl") = driverlicense
rs("dlstate") = driverlicensestate
If IsDate(SBIRTHDATE) Then
    rs("birthdate") = SBIRTHDATE
End If
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
db.Close
If othercases(0) = "" And othercases(1) = "" And othercases(2) = "" And othercases(3) = "" And othercases(4) = "" And othercases(5) = "" And othercases(6) = "" And othercases(7) = "" And othercases(8) = "" And othercases(9) = "" Then
    Exit Sub
End If

Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume Next
End If
End Sub
Private Sub loadcodes()
Dim db As Database, rs As Recordset, itmx As ListItem
On Error GoTo oderror
od:
Set db = OpenDatabase(nwb + "booking.mdb")
Set rs = db.OpenRecordset("select * from codes where type = 'ucr' or type = 'armedwith' OR TYPE = 'drugtype' or type = 'activity' ORDER BY CODE")
On Error Resume Next
armedlist(0).clear
armedlist(1).clear
dt.clear
at.clear
If rs.EOF Then
    db.Close
    Exit Sub
End If
rs.MoveFirst
While Not rs.EOF
    Select Case rs("type")
        Case "drugtype"
            dt.AddItem rs("code")
        Case "activity"
            at.AddItem rs("code")
        Case "ucr"
            For t% = 0 To 2
                ucrlist(t%).AddItem rs("code")
            Next t%
        Case "armedwith"
            armedlist(0).AddItem rs("code")
            If rs("DEFAULT") = "Y" Then
                armedlist(0).ListIndex = armedlist(0).ListCount - 1
            End If
            armedlist(1).AddItem rs("code")
    End Select
    rs.MoveNext
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
Private Sub deleteroutine()
msg = MsgBox("Are you sure?", 4, "Genesis Information Log")
If msg <> 6 Then
    Exit Sub
End If
Dim db As Database, rs, rs2 As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nwb + "booking.MDB")
Set rs = db.OpenRecordset("select * from booking where incidentnumber = " + Chr$(34) + incidentnumber + Chr$(34) + " and sname = " + Chr$(34) + sname + Chr$(34))
On Error Resume Next
If Not rs.EOF Then
    rs.MoveFirst
    rs.Delete
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
Private Sub defaultcodes()
Dim db As Database, rs As Recordset, itmx As ListItem
On Error GoTo oderror
od:
Set db = OpenDatabase(nwb + "booking.mdb")
Set rs = db.OpenRecordset("select * from codes")
If rs.EOF Then
    db.Close
    On Error Resume Next
    Exit Sub
End If
rs.MoveFirst
On Error Resume Next
SSEX.clear
SRACE.clear
SETHNICITY.clear
sresident.clear
weapontype.ListItems.clear
widx% = 0
While Not rs.EOF
    Select Case rs("type")
        Case "sex"
            SSEX.AddItem rs("code")
            If UCase(rs("default")) = "Y" Then
                SSEX.ListIndex = SSEX.ListCount - 1
            End If
        Case "race"
            SRACE.AddItem rs("code")
            If UCase(rs("default")) = "Y" Then
                SRACE.ListIndex = SRACE.ListCount - 1
            End If
        Case "ethnicity"
            SETHNICITY.AddItem rs("code")
            If UCase(rs("default")) = "Y" Then
                SETHNICITY.ListIndex = SETHNICITY.ListCount - 1
            End If
        Case "resident"
            sresident.AddItem rs("code")
            If UCase(rs("default")) = "Y" Then
                sresident.ListIndex = sresident.ListCount - 1
            End If
    End Select
    rs.MoveNext
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

Private Sub loadlist()
Data1.DatabaseName = nwl + "lawsuite.mdb"
Data1.Refresh
Data1.RecordSource = "select dpname from PEOPLE order by dpsort"
Data1.Refresh
sname.DataField = dpname
sname.Refresh
On Error Resume Next
incidentnumber.clear
Dim db As Database, rs As Recordset
Set db = OpenDatabase(nwb + "booking.mdb")
Set rs = db.OpenRecordset("select incidentnumber from booking order by incidentnumber desc")
If Not rs.EOF Then
    rs.MoveFirst
    HI = Right$(Date$, 2) + "-" + Format$(Val(Mid$(rs("INCIDENTNUMBER"), 4)) + 1, "00000")
    While Not rs.EOF
        incidentnumber.AddItem rs("incidentnumber")
        rs.MoveNext
    Wend
End If
If HI = "" Then
    HI = Right$(Date$, 2) + "-000001"
End If
incidentnumber = HI
db.Close
End Sub
Friend Sub findincident()
If incidentnumber = "" Then
    Exit Sub
End If
Dim db As Database, rs As Recordset
On Error GoTo oderror
od:

If subjectnumber = "" Then
    SSS$ = ""
Else
    SSS$ = " and number = " + subjectnumber
End If

Set db = OpenDatabase(nwb + "booking.mdb")
Set rs = db.OpenRecordset("select * from booking where incidentnumber = " + Chr$(34) + incidentnumber + Chr$(34) + SSS$)
If rs.EOF Then
    On Error Resume Next
    db.Close
    Exit Sub
End If
rs.MoveFirst
incidentnumber = rs("Incidentnumber")
arrestnumber = rs("arrestnumber")
dateofarrest = rs("dateofarrest")
If Not IsNull(rs("timeofarrest")) Then
    timeofarrest = rs("timeofarrest")
End If
sname = rs("sname")
If Not IsNull(rs("srace")) Then
    For uu% = 0 To SRACE.ListCount - 1
        If Left$(SRACE.List(uu%), 1) = rs("srace") Then
            SRACE.ListIndex = uu%
            uu% = SRACE.ListCount - 1
        End If
    Next uu%
End If
If Not IsNull(rs("ssex")) Then
    For uu% = 0 To SSEX.ListCount - 1
        If Left$(SSEX.List(uu%), 1) = rs("ssex") Then
            SSEX.ListIndex = uu%
            uu% = SSEX.ListCount - 1
        End If
    Next uu%
End If
If Not IsNull(rs("sbirthdate")) Then
    SBIRTHDATE = rs("sbirthdate")
End If
sage = rs("Sage")
If Not IsNull(rs("sethnicity")) Then
    For uu% = 0 To SETHNICITY.ListCount - 1
        If Left$(SETHNICITY.List(uu%), 1) = rs("sethnicity") Then
            SETHNICITY.ListIndex = uu%
            uu% = SETHNICITY.ListCount - 1
        End If
    Next uu%
End If
SHT = rs("sheight")
SWEIGHT = rs("sweight")
SHAIR = rs("shair")
SEYES = rs("seyes")
If Not IsNull(rs("speculiarities")) Then
    speculiarities = rs("speculiarities")
End If
saddress = rs("saddress")
scity = rs("scity")
sstate = rs("sstate")
szipcode = rs("szipcode")
If Not IsNull(rs("docketnumber")) Then
    docketnumber = rs("docketnumber")
End If
If Not IsNull(rs("ssn")) Then
    ssn = rs("ssn")
End If
If Not IsNull(rs("ncic")) Then
    ncic = rs("ncic")
End If
If Not IsNull(rs("idnumber")) Then
    idnumber = rs("idnumber")
End If
If Not IsNull(rs("agency")) Then
    agency = rs("agency")
End If
If Not IsNull(rs("sresident")) Then
For t% = 0 To sresident.ListCount - 1
    If rs("sresident") = Left$(sresident.List(t%), 1) Then
        sresident.ListIndex = t%
        t% = sresident.ListCount - 1
    End If
Next t%
End If
If Not IsNull(rs("ARMEDWITH1")) Then
For t% = 0 To armedlist(0).ListCount - 1
    If rs("ARMEDWITH1") = Mid$(armedlist(0).List(t%), InStr(armedlist(0).List(t%), "(") + 1, 2) Then
        armedlist(0).ListIndex = t%
        t% = armedlist(0).ListCount - 1
    End If
Next t%
End If
If Not IsNull(rs("ARMEDWITH2")) Then
For t% = 0 To armedlist(1).ListCount - 1
    If rs("ARMEDWITH2") = Mid$(armedlist(1).List(t%), InStr(armedlist(1).List(t%), "(") + 1, 2) Then
        armedlist(1).ListIndex = t%
        t% = armedlist(1).ListCount - 1
    End If
Next t%
End If
If Not IsNull(rs("phone")) Then
    phone = rs("phone")
End If
If Not IsNull(rs("alias")) Then
    alias = rs("alias")
End If
If Not IsNull(rs("birthplace")) Then
    birthplace = rs("birthplace")
End If
If Not IsNull(rs("driverslicense")) Then
    driverslicense = rs("driverslicense")
End If
If Not IsNull(rs("driverslicensestate")) Then
    driverslicensestate = rs("driverslicensestate")
End If
If Not IsNull(rs("employer")) Then
    employer = rs("employer")
End If
If Not IsNull(rs("nextofkin")) Then
    nextofkin = rs("nextofkin")
End If
If Not IsNull(rs("nextofkinaddress")) Then
    nextofkinaddress = rs("nextofkinaddress")
End If
If Not IsNull(rs("bookingofficer")) Then
    BOOKINGOFFICER = rs("bookingofficer")
End If
If Not IsNull(rs("bookingofficerUNIT")) Then
    bookingofficerunit = rs("bookingofficerUNIT")
End If
If Not IsNull(rs("arrestingofficer")) Then
    ARRESTINGOFFICER = rs("arrestingofficer")
End If
If Not IsNull(rs("arrestingUNIT")) Then
    arrestingunit = rs("arrestingUNIT")
End If
If Not IsNull(rs("ARMEDWITHAUTOMATIC1")) Then
    armedwithautomatic(0) = rs("ARMEDWITHAUTOMATIC1")
End If
If Not IsNull(rs("ARMEDWITHAUTOMATIC2")) Then
    armedwithautomatic(1) = rs("ARMEDWITHAUTOMATIC2")
End If
If Not IsNull(rs("ARMEDWITHsemiAUTOMATIC1")) Then
    armedwithsemiautomatic(0) = rs("ARMEDWITHsemiAUTOMATIC1")
End If
If Not IsNull(rs("ARMEDWITHsemiAUTOMATIC2")) Then
    armedwithsemiautomatic(1) = rs("ARMEDWITHsemiAUTOMATIC2")
End If
If Not IsNull(rs("ONVIEW")) Then
    onviewarrest.Value = rs("ONVIEW")
End If
If Not IsNull(rs("SUMMONED")) Then
    summoned.Value = rs("SUMMONED")
End If
If Not IsNull(rs("TAKEN")) Then
    taken.Value = rs("TAKEN")
End If
If Not IsNull(rs("WITHIN")) Then
    within = rs("WITHIN")
End If
If Not IsNull(rs("REFERRED")) Then
    referred = rs("REFERRED")
End If
For p% = 1 To 10
    If Not IsNull(rs("OTHERCASES" + Mid$(Str$(p%), 2))) Then
        othercases(p% - 1) = rs("OTHERCASES" + Mid$(Str$(p%), 2))
    End If
Next p%
If Not IsNull(rs("CHARGEA")) Then
    For uu% = 0 To ucrlist(0).ListCount - 1
        If ucrlist(0).List(uu%) = rs("chargea") Then
            ucrlist(0).ListIndex = uu%
            ucrlist(0).Selected(uu%) = True
            uu% = ucrlist(0).ListCount - 1
        End If
    Next uu%
End If
If Not IsNull(rs("CHARGEB")) Then
    For uu% = 0 To ucrlist(1).ListCount - 1
        If ucrlist(1).List(uu%) = rs("chargeb") Then
            ucrlist(1).ListIndex = uu%
            ucrlist(1).Selected(uu%) = True
            uu% = ucrlist(1).ListCount - 1
        End If
    Next uu%
End If
If Not IsNull(rs("CHARGEC")) Then
    For uu% = 0 To ucrlist(2).ListCount - 1
        If ucrlist(2).List(uu%) = rs("chargec") Then
            ucrlist(2).ListIndex = uu%
            ucrlist(2).Selected(uu%) = True
            uu% = ucrlist(2).ListCount - 1
        End If
    Next uu%
End If
If Not IsNull(rs("STATUTEA")) Then
    statutea = rs("STATUTEA")
End If
If Not IsNull(rs("STATUTEB")) Then
    statuteb = rs("STATUTEB")
End If
If Not IsNull(rs("STATUTEC")) Then
    statutec = rs("STATUTEC")
End If
If Not IsNull(rs("REMARKS")) Then
    REMARKS.Text = rs("REMARKS")
End If
rs.Edit
On Error Resume Next
If dt.ListIndex = -1 And at.ListIndex = -1 Then
    If Not IsNull(rs("dt")) Then
        dt.Visible = True
        at.Visible = True
        For rr% = 0 To dt.ListCount - 1
            If Left$(dt.List(rr%), 1) = rs("dt") Then
                dt.ListIndex = rr%
                rr% = dt.ListCount - 1
            End If
        Next rr%
        For rr% = 0 To at.ListCount - 1
            If Mid$(at.List(rr%), InStr(at.List(rr%), "(") + 1, 1) = rs("at") Then
                at.ListIndex = rr%
                rr% = at.ListCount - 1
            End If
        Next rr%
    End If
End If
ssql = "select mugshot from people where dpnamelf = '" + sname + "'"
If ssn > "" Then
    ssql = ssql + " and ssn = '" + ssn + "'"
End If
If idnumber > "" Then
    ssql = ssql + " and idnumber = '" + idnumber + "'"
End If
If IsDate(SBIRTHDATE) Then
    ssql = ssql + " and birthdate = #" + SBIRTHDATE + "#"
End If
Set db = OpenDatabase(nwl + "lawsuite.mdb")
Set rs = db.OpenRecordset(ssql)
If Not rs.EOF Then
    rs.MoveFirst
    If Not IsNull(rs("mugshot")) Then
        mugshot.Picture = LoadPicture(rs("mugshot"))
    End If
End If
found35a = False
For t% = 0 To ucrlist(Index).ListCount - 1
    If ucrlist(Index).Selected(t%) Then
        If InStr(ucrlist(Index).List(t%), "35A") > 0 Then
            found35a = True
            t% = ucrlist(Index).ListCount - 1
        End If
    End If
Next t%
If found35a Then
    dt.Visible = True
    at.Visible = True
Else
    dt.Visible = False
    at.Visible = False
End If
db.Close
Exit Sub
oderror:
If Err > 3200 Then
    Resume od
Else
    Resume
End If


End Sub

Private Sub LoadOfficers()
Dim db As Database, rs As Recordset
Set db = OpenDatabase(nwl + "LAWSUITE.mdb")
Set rs = db.OpenRecordset("select PROFNAME from PROFESSIONALS WHERE TYPE = 'D'")
BOOKINGOFFICER.clear
ARRESTINGOFFICER.clear
If Not rs.EOF Then
    rs.MoveFirst
    While Not rs.EOF
        BOOKINGOFFICER.AddItem rs("PROFNAME")
        ARRESTINGOFFICER.AddItem rs("PROFNAME")
        rs.MoveNext
    Wend
End If
db.Close
End Sub

Private Sub within_GotFocus()
If Me.ActiveControl.Top > (-1 * PICTURE2.Top) And Me.ActiveControl.Top + Me.ActiveControl.Height < (-1 * PICTURE2.Top) + VScroll1.Height Then
Else
If Me.ActiveControl.Top > 500 Then
    VScroll1 = Me.ActiveControl.Top - 500
Else
    VScroll1 = 0
End If
End If


End Sub
Friend Sub editroutine(editerr As Integer, msg As String)
Dim db, db2 As Database, rs, rs2 As Recordset, ab(2) As String
Set db = OpenDatabase(nwb + "booking.MDB")
Set db2 = OpenDatabase(nwi + "incident.mdb")
editerr = 1
Set rs = db2.OpenRecordset("select dateofoffense2 from incidentreportc where incidentnumber = '" + incidentnumber + "'")
If Not rs.EOF Then
    rs.MoveFirst
    If Not IsNull(rs("dateofoffense2")) Then
        If IsDate(rs("dateofoffense2")) Then
            offensedate = rs("dateofoffense2")
        Else
            msg = "Unable to locate offense date."
            Exit Sub
        End If
    Else
        msg = "Unable to locate offense date."
        Exit Sub
    End If
Else
    msg = "Unable to locate offense date."
    Exit Sub
End If
Set rs = db2.OpenRecordset("select excleardate from incidentreporto where incidentnumber = '" + incidentnumber + "' and excleardate is not null")
If Not rs.EOF Then
    msg = "Bookings are not allowed with an Exceptional Clearance."
    Exit Sub
End If
found35a = False
For t% = 0 To 2
    ab(t%) = ""
    ucrlist(t%).ListIndex = -1
    For Y% = 0 To ucrlist(t%).ListCount - 1
        If ucrlist(t%).Selected(Y%) Then
            If InStr(ucrlist(t%).List(Y%), "35A") > 0 Then
                found35a = True
            End If
            Set rs2 = db2.OpenRecordset("select abgroup from ucr where code = '" + ucrlist(t%).List(Y%) + "'")
            If Not rs2.EOF Then
                rs2.MoveFirst
                ab(t%) = rs2("abgroup")
            End If
            ucrlist(t%).ListIndex = Y%
            Y% = ucrlist(t%).ListCount - 1
        End If
    Next Y%
Next t%
If ab(0) = "B" And (ab(1) = "A" Or ab(2) = "A") Or _
   ab(0) = "B" And ab(1) = "B" And ab(2) = "A" Then
    msg = "Arrest charges must be entered in the order of most serious first."
    Exit Sub
End If
If ucrlist(0).ListIndex = -1 Then
    msg = "A valid UCR must be selected."
    Exit Sub
End If
If found35a Then
    If dt.ListIndex = -1 Or at.ListIndex = -1 Then
        msg = "A drug and activity type must be selected for UCR 35A."
        Exit Sub
    End If
End If
'==== Mandatories E - 8 = GIVEN
'==== Mandatories E - 36
'==== Mandatories E - 40
'===== Error 652,752
tage = sage
If Val(sage) = 0 Then
    msg = "A valid age must be entered."
    Exit Sub
End If
If ((Len(tage) = 4 And Val(Right$(tage, 2)) < 18) Or (Len(tage) = 2 And Val(tage) < 18)) Then
    If within <> 1 And referred <> 1 Then
        msg = "If arrestee is under 18, Handled Within Department or Referred to Other Authority must be selected."
        Exit Sub
    End If
End If
If ((Len(tage) = 4 And Val(Right$(tage, 2)) > 17) Or (Len(tage) = 2 And Val(tage) > 17)) Then
    If within = 1 Or referred = 1 Then
        msg = "If arrestee is over 17, Handled Within Department and Referred to Other Authority cannot be selected."
        Exit Sub
    End If
End If
'===== Error 601,701
If onviewarrest.Value = 0 And taken.Value = 0 And summoned.Value = 0 Then
    msg = "A type of arrest must be selected."
    Exit Sub
End If
'==== Resident Status
If sresident.ListIndex = -1 Then
    msg = "Residency data has not been entered."
    Exit Sub
End If
If SRACE.ListIndex = -1 Then
    msg = "Race data has not been entered."
    Exit Sub
End If
If SSEX.ListIndex = -1 Then
    msg = "Sex data has not been entered."
    Exit Sub
End If
If SETHNICITY.ListIndex = -1 Then
    msg = "Ethnicity data has not been entered."
    Exit Sub
End If
'===== Error 601,701
If ucrlist(0).ListIndex = -1 Then
    msg = "An arrest UCR code must be selected."
    Exit Sub
End If
'===== Error 670
If InStr(ucrlist(0).List(ucrlist(0).ListIndex), "09C") > 0 Then
    msg = "Bookings are not allowed for Justifiable Homocide."
    Exit Sub
End If
For rr% = 1 To 2
    If ucrlist(rr%).ListIndex > -1 Then
        If InStr(ucrlist(rr%).List(ucrlist(rr%).ListIndex), "09C") > 0 Then
            msg = "Bookings are not allowed for Justifiable Homocide."
            Exit Sub
        End If
    End If
Next rr%
'===== Error 601,701,655,755
If armedlist(0).ListIndex = -1 Then
    msg = "A selection must be made from ARMED WITH."
    Exit Sub
Else
    If armedwithautomatic(0) Or armedwithsemiautomatic(0) Then
        Select Case Mid$(armedlist(0).List(armedlist(0).ListIndex), InStr(armedlist(0).List(armedlist(0).ListIndex), "(") + 1, 2)
            Case "11", "12", "13", "14", "15"
            Case Else
                msg = "Automatic Weapon indicator not allowed with selected weapon type."
                Exit Sub
        End Select
    End If
End If

'===== Error 641,741
'If Val(sage) = 99 Then
'    msg = "An arrestee age of 99 has been entered.  Is this correct?", 4, "Genesis Information Log")
'    If msg = "7 Then
'        Exit Sub
'    End If
'End If
'===== Error 606,706,607,707,655,755
If armedlist(1).ListIndex > -1 Then
    If InStr(armedlist(0).List(armedlist(0).ListIndex), "(01)") = 0 Or InStr(armedlist(1).List(armedlist(1).ListIndex), "(01)") = 0 Then
        If armedlist(0).ListIndex = armedlist(1).ListIndex Then
            msg = "Duplicate values for Armed With are not allowed."
            Exit Sub
        End If
        If InStr(armedlist(0).List(armedlist(0).ListIndex), "(01)") > 0 Or InStr(armedlist(1).List(armedlist(1).ListIndex), "(01)") > 0 Then
            msg = "Other values for Armed WIth are not allowed in combination with Unarmed."
            Exit Sub
        End If
    End If
    If armedwithautomatic(1) Or armedwithsemiautomatic(1) Then
        Select Case Mid$(armedlist(1).List(armedlist(1).ListIndex), InStr(armedlist(1).List(armedlist(1).ListIndex), "(") + 1, 2)
            Case "11", "12", "13", "14", "15"
            Case Else
                msg = "Automatic Weapon indicator not allowed with selected weapon type."
                Exit Sub
        End Select
    End If
End If
incidentnumber = UCase(incidentnumber)
'===== Error 617, 717
For t% = 1 To Len(incidentnumber)
    If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789- ", Mid$(incidentnumber, t%, 1)) = 0 Then
        msg = "An invalid character has been found in the INCIDENT Number field.  Valid characters are A-Z, 0-9, and Hyphen.  Do not enter any Blanks because these are computer generated."
        t% = Len(incidentnumber)
        Exit Sub
    End If
Next t%
'===== Error 615,715
incidentnumber = incidentnumber + Space$(12 - Len(incidentnumber))
'===== Mandatories 47,52
If Val(sage) < 18 And sage <> "00" Then
    If within Or referred Then
    Else
        msg = "For arrestees under 18, a disposition must be selected."
        Exit Sub
    End If
End If
'===== Error 601,701
If armedlist(0).ListIndex > -1 Or armedlist(1).ListIndex > -1 Or arrestnumber > "" Or ucrlist(0).ListIndex > -1 Or onviewarrest.Value = 1 Or summoned.Value = 1 Or taken.Value = 1 Or within Or referred Then
    If armedlist(0).ListIndex = -1 Or arrestnumber = "" Or ucrlist(0).ListIndex = -1 Or (onviewarrest.Value = 0 And summoned.Value = 0 And taken.Value = 0) Or Not IsDate(dateofarrest) Or (Val(sage) = 0 And sage <> "00") Then
        msg = "All arrestee information must be entered."
        Exit Sub
    End If
End If
If CDate(dateofarrest) < offensedate Then
    msg = "Date of arrest cannot be prior to offense date."
    Exit Sub
End If
'===== SCEdit 4/21/92 P28
If dt.Visible = True Then
    If dt.ListIndex = -1 Or at.ListIndex = -1 Then
        msg = "For drug-related arrests, a drug type/activity combination must be selected for each arrestee."
        Exit Sub
    End If
End If
founda = False
matcha = False
Set rs = db2.OpenRecordset("select * from incidentsupport where incidentnumber = " + Chr$(34) + incidentnumber + Chr$(34))
If Not rs.EOF Then
    rs.MoveFirst
    foundmatch = False
    For u% = 1 To 10
        If Not IsNull(rs("ucr" + CStr(u%))) Then
            Set rs3 = db2.OpenRecordset("select abgroup from ucr where abbrev = '" + rs("ucr" + CStr(u%)) + "'")
            rs3.MoveFirst
            For t% = 0 To ucrlist(0).ListCount - 1
                If InStr(ucrlist(0).List(t%), "(" + rs("ucr" + CStr(u%)) + ")") > 0 Then
                    If rs3("abgroup") = "A" Then
                        founda = True
                        t% = ucrlist(0).ListCount - 1
                    End If
                End If
            Next t%
            For t% = 0 To 2
                For tt% = 0 To ucrlist(t%).ListCount - 1
                    If ucrlist(t%).Selected(tt%) Then
                        If InStr(ucrlist(t%).List(tt%), "(" + rs("ucr" + CStr(u%)) + ")") > 0 Then
                            foundmatch = True
                            If rs3("abgroup") = "A" Then
                                matcha = True
                                tt% = ucrlist(t%).ListCount - 1
                                t% = 2
                            End If
                        End If
                    End If
                Next tt%
            Next t%
        End If
    Next u%
    If Not foundmatch Then
        msg = "At least one of the UCR codes selected on the incident report must be selected in CHARGEA."
        Exit Sub
    Else
    If founda And Not matcha Then
        msg = "If a group A offense exists on the incident, the booking must associate to at least one of the group A codes."
        Exit Sub
    End If
    End If
End If
editerr = 0
End Sub
Friend Sub clearroutine()
arrestnumber = ""
incidentnumber = ""
sname = ""
SRACE.ListIndex = -1
SSEX.ListIndex = -1
SBIRTHDATE = ""
sage = ""
SETHNICITY.ListIndex = -1
SHT = ""
SWEIGHT = ""
SHAIR = ""
SEYES = ""
speculiarities = ""
saddress = ""
scity = ""
sstate = ""
szipcode = ""
PICFILE = ""
sresident.ListIndex = -1
armedlist(0).ListIndex = -1
armedwithautomatic(0) = 0
armedlist(1).ListIndex = -1
armedwithautomatic(1) = 0
armedwithsemiautomatic(0) = 0
armedwithsemiautomatic(1) = 0
onviewarrest = 0
summoned = 0
taken = 0
within = 0
referred = 0
docketnumber = ""
ssn = ""
ncic = ""
agency = ""
idnumber = ""
phone = ""
birthplace = ""
alias = ""
nextofkin = ""
nextofkinaddress = ""
employer = ""
BOOKINGOFFICER = ""
bookingofficerunit = ""
ARRESTINGOFFICER = ""
arrestingunit = ""
driverslicense = ""
driverslicensestate = ""
For p% = 0 To 9
    othercases(p%) = ""
Next p%
For pp% = 0 To 2
    For p% = 0 To ucrlist(pp%).ListCount - 1
        ucrlist(pp%).Selected(p%) = False
    Next p%
Next pp%
ucrlist(0).ListIndex = -1
ucrlist(1).ListIndex = -1
ucrlist(2).ListIndex = -1
statutea = ""
statuteb = ""
statutec = ""
REMARKS.Text = ""
dateofarrest = ""
timeofarrest = ""
mugshot.Picture = LoadPicture()
VScroll1 = 0
fromexport = False
For t% = 0 To Forms.Count - 1
    If LCase(Forms(t%).Name) = "iexport" Then
        fromexport = True
        t% = Forms.Count - 1
    End If
Next t%
If Not fromexport Then
    Call loadlist
End If

End Sub
