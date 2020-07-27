VERSION 5.00
Begin VB.Form SECURITY 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Security"
   ClientHeight    =   8145
   ClientLeft      =   90
   ClientTop       =   600
   ClientWidth     =   11685
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   11685
   Begin VB.CheckBox rmsbrowse 
      Caption         =   "abrowse"
      DataField       =   "bbrowse"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   5025
      TabIndex        =   149
      Top             =   6120
      Width           =   300
   End
   Begin VB.CheckBox rmsdelete 
      Caption         =   "adelete"
      DataField       =   "bdelete"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   5025
      TabIndex        =   148
      Top             =   6435
      Width           =   300
   End
   Begin VB.CheckBox rmsedit 
      Caption         =   "aedit"
      DataField       =   "bedit"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   5025
      TabIndex        =   147
      Top             =   6765
      Width           =   300
   End
   Begin VB.CheckBox rmssupervisor 
      Caption         =   "asupervisor"
      DataField       =   "bsupervisor"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   7380
      TabIndex        =   146
      Top             =   6765
      Width           =   300
   End
   Begin VB.CheckBox rmsreport 
      Caption         =   "areport"
      DataField       =   "breport"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   7380
      TabIndex        =   145
      Top             =   6435
      Width           =   300
   End
   Begin VB.CheckBox rmsprint 
      Caption         =   "aprint"
      DataField       =   "bprint"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   7380
      TabIndex        =   144
      Top             =   6120
      Width           =   300
   End
   Begin VB.CheckBox abrowse 
      Caption         =   "abrowse"
      DataField       =   "bbrowse"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   5040
      TabIndex        =   41
      Top             =   4560
      Width           =   300
   End
   Begin VB.CheckBox adelete 
      Caption         =   "adelete"
      DataField       =   "bdelete"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   5025
      TabIndex        =   42
      Top             =   4875
      Width           =   300
   End
   Begin VB.CheckBox aedit 
      Caption         =   "aedit"
      DataField       =   "bedit"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   5040
      TabIndex        =   43
      Top             =   5205
      Width           =   300
   End
   Begin VB.CheckBox asupervisor 
      Caption         =   "asupervisor"
      DataField       =   "bsupervisor"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   7380
      TabIndex        =   46
      Top             =   5205
      Width           =   300
   End
   Begin VB.CheckBox areport 
      Caption         =   "areport"
      DataField       =   "breport"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   7380
      TabIndex        =   45
      Top             =   4875
      Width           =   300
   End
   Begin VB.CheckBox aprint 
      Caption         =   "aprint"
      DataField       =   "bprint"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   7380
      TabIndex        =   44
      Top             =   4560
      Width           =   300
   End
   Begin VB.TextBox orinumber 
      DataField       =   "password"
      DataSource      =   "datPrimaryRS"
      Height          =   315
      Left            =   6570
      MaxLength       =   9
      TabIndex        =   4
      Text            =   " "
      Top             =   540
      Width           =   2415
   End
   Begin VB.CheckBox jprint 
      DataField       =   "bprint"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   11250
      TabIndex        =   50
      Top             =   1440
      Width           =   300
   End
   Begin VB.CheckBox jreport 
      DataField       =   "breport"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   11250
      TabIndex        =   51
      Top             =   1755
      Width           =   300
   End
   Begin VB.CheckBox jsupervisor 
      DataField       =   "bsupervisor"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   11250
      TabIndex        =   52
      Top             =   2085
      Width           =   300
   End
   Begin VB.CheckBox jbrowse 
      DataField       =   "bbrowse"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   8880
      TabIndex        =   47
      Top             =   1440
      Width           =   300
   End
   Begin VB.CheckBox jdelete 
      DataField       =   "bdelete"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   8880
      TabIndex        =   48
      Top             =   1755
      Width           =   300
   End
   Begin VB.CheckBox jedit 
      DataField       =   "bedit"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   8895
      TabIndex        =   49
      Top             =   2085
      Width           =   300
   End
   Begin VB.CheckBox BPRINT 
      DataField       =   "bprint"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   7380
      TabIndex        =   38
      Top             =   2940
      Width           =   300
   End
   Begin VB.CheckBox BREPORT 
      DataField       =   "breport"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   7380
      TabIndex        =   39
      Top             =   3255
      Width           =   300
   End
   Begin VB.CheckBox BSUPERVISOR 
      DataField       =   "bsupervisor"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   7380
      TabIndex        =   40
      Top             =   3585
      Width           =   300
   End
   Begin VB.CheckBox IPRINT 
      DataField       =   "iprint"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   7380
      TabIndex        =   32
      Top             =   1380
      Width           =   300
   End
   Begin VB.CheckBox IREPORT 
      DataField       =   "ireport"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   7380
      TabIndex        =   33
      Top             =   1695
      Width           =   300
   End
   Begin VB.CheckBox ISUPERVISOR 
      DataField       =   "isupervisor"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   7380
      TabIndex        =   34
      Top             =   2025
      Width           =   300
   End
   Begin VB.CheckBox RPRINT 
      DataField       =   "rprint"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   11340
      TabIndex        =   56
      Top             =   2940
      Width           =   300
   End
   Begin VB.CheckBox RREPORT 
      DataField       =   "rreport"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   11340
      TabIndex        =   57
      Top             =   3270
      Width           =   300
   End
   Begin VB.CheckBox RSUPERVISOR 
      DataField       =   "rsupervisor"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   11340
      TabIndex        =   58
      Top             =   3585
      Width           =   300
   End
   Begin VB.CheckBox WPRINT 
      DataField       =   "wprint"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   11340
      TabIndex        =   62
      Top             =   4605
      Width           =   300
   End
   Begin VB.CheckBox WREPORT 
      DataField       =   "wreport"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   11340
      TabIndex        =   63
      Top             =   4905
      Width           =   300
   End
   Begin VB.CheckBox WSUPERVISOR 
      DataField       =   "wsupervisor"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   11340
      TabIndex        =   64
      Top             =   5205
      Width           =   300
   End
   Begin VB.TextBox userfullname 
      Height          =   330
      Left            =   1035
      MaxLength       =   100
      TabIndex        =   3
      Top             =   555
      Width           =   3990
   End
   Begin VB.CheckBox CDELETE 
      DataField       =   "cdelete"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   1
      Left            =   1095
      TabIndex        =   12
      Top             =   3525
      Width           =   300
   End
   Begin VB.CheckBox CSUPERVISOR 
      DataField       =   "csupervisor"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   1
      Left            =   3435
      TabIndex        =   16
      Top             =   3810
      Width           =   300
   End
   Begin VB.CheckBox CREPORT 
      DataField       =   "creport"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   1
      Left            =   3435
      TabIndex        =   15
      Top             =   3525
      Width           =   300
   End
   Begin VB.CheckBox CPRINT 
      DataField       =   "cprint"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   1
      Left            =   3435
      TabIndex        =   14
      Top             =   3210
      Width           =   300
   End
   Begin VB.CheckBox CEDIT 
      DataField       =   "cedit"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   1
      Left            =   1095
      TabIndex        =   13
      Top             =   3810
      Width           =   300
   End
   Begin VB.CheckBox CBROWSE 
      DataField       =   "cbrowse"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   1
      Left            =   1095
      TabIndex        =   11
      Top             =   3210
      Width           =   300
   End
   Begin VB.CheckBox CBROWSE 
      DataField       =   "cbrowse"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   3
      Left            =   1095
      TabIndex        =   23
      Top             =   6090
      Width           =   300
   End
   Begin VB.CheckBox CEDIT 
      DataField       =   "cedit"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   3
      Left            =   1095
      TabIndex        =   25
      Top             =   6735
      Width           =   300
   End
   Begin VB.CheckBox CPRINT 
      DataField       =   "cprint"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   3
      Left            =   3435
      TabIndex        =   26
      Top             =   6090
      Width           =   300
   End
   Begin VB.CheckBox CREPORT 
      DataField       =   "creport"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   3
      Left            =   3435
      TabIndex        =   27
      Top             =   6405
      Width           =   300
   End
   Begin VB.CheckBox CSUPERVISOR 
      DataField       =   "csupervisor"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   3
      Left            =   3435
      TabIndex        =   28
      Top             =   6735
      Width           =   300
   End
   Begin VB.CheckBox CDELETE 
      DataField       =   "cdelete"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   3
      Left            =   1095
      TabIndex        =   24
      Top             =   6405
      Width           =   300
   End
   Begin VB.CheckBox CBROWSE 
      DataField       =   "cbrowse"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   2
      Left            =   1095
      TabIndex        =   17
      Top             =   4650
      Width           =   300
   End
   Begin VB.CheckBox CEDIT 
      DataField       =   "cedit"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   2
      Left            =   1095
      TabIndex        =   19
      Top             =   5295
      Width           =   300
   End
   Begin VB.CheckBox CPRINT 
      DataField       =   "cprint"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   2
      Left            =   3435
      TabIndex        =   20
      Top             =   4650
      Width           =   300
   End
   Begin VB.CheckBox CREPORT 
      DataField       =   "creport"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   2
      Left            =   3435
      TabIndex        =   21
      Top             =   4965
      Width           =   300
   End
   Begin VB.CheckBox CSUPERVISOR 
      DataField       =   "csupervisor"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   2
      Left            =   3435
      TabIndex        =   22
      Top             =   5295
      Width           =   300
   End
   Begin VB.CheckBox CDELETE 
      DataField       =   "cdelete"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   2
      Left            =   1095
      TabIndex        =   18
      Top             =   4965
      Width           =   300
   End
   Begin VB.ComboBox userid 
      Height          =   315
      Left            =   1035
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   150
      Width           =   4005
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Save"
      Height          =   300
      Left            =   15
      TabIndex        =   102
      Top             =   7755
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   300
      Left            =   1455
      TabIndex        =   101
      Top             =   7755
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   300
      Left            =   2895
      TabIndex        =   100
      Top             =   7755
      Width           =   1095
   End
   Begin VB.CheckBox CDELETE 
      DataField       =   "cdelete"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   0
      Left            =   1095
      TabIndex        =   6
      Top             =   2085
      Width           =   300
   End
   Begin VB.TextBox password 
      DataField       =   "password"
      DataSource      =   "datPrimaryRS"
      Height          =   315
      Left            =   6570
      TabIndex        =   1
      Text            =   " "
      Top             =   135
      Width           =   2415
   End
   Begin VB.CheckBox WEDIT 
      DataField       =   "wedit"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   8955
      TabIndex        =   61
      Top             =   5205
      Width           =   285
   End
   Begin VB.CheckBox WDELETE 
      DataField       =   "wdelete"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   8955
      TabIndex        =   60
      Top             =   4905
      Width           =   165
   End
   Begin VB.CheckBox WBROWSE 
      DataField       =   "wbrowse"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   8955
      TabIndex        =   59
      Top             =   4605
      Width           =   165
   End
   Begin VB.CheckBox SUPERVISOR 
      DataField       =   "supervisor"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   11130
      TabIndex        =   2
      Top             =   285
      Width           =   1000
   End
   Begin VB.CheckBox REDIT 
      DataField       =   "redit"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   8955
      TabIndex        =   55
      Top             =   3585
      Width           =   300
   End
   Begin VB.CheckBox RDELETE 
      DataField       =   "rdelete"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   8955
      TabIndex        =   54
      Top             =   3270
      Width           =   300
   End
   Begin VB.CheckBox RBROWSE 
      DataField       =   "rbrowse"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   8955
      TabIndex        =   53
      Top             =   2925
      Width           =   300
   End
   Begin VB.CheckBox IEDIT 
      DataField       =   "iedit"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   5025
      TabIndex        =   31
      Top             =   2025
      Width           =   300
   End
   Begin VB.CheckBox IDELETE 
      DataField       =   "idelete"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   5025
      TabIndex        =   30
      Top             =   1695
      Width           =   300
   End
   Begin VB.CheckBox IBROWSE 
      DataField       =   "ibrowse"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   5025
      TabIndex        =   29
      Top             =   1380
      Width           =   300
   End
   Begin VB.CheckBox CSUPERVISOR 
      DataField       =   "csupervisor"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   0
      Left            =   3435
      TabIndex        =   10
      Top             =   2415
      Width           =   300
   End
   Begin VB.CheckBox CREPORT 
      DataField       =   "creport"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   0
      Left            =   3435
      TabIndex        =   9
      Top             =   2085
      Width           =   300
   End
   Begin VB.CheckBox CPRINT 
      DataField       =   "cprint"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   0
      Left            =   3435
      TabIndex        =   8
      Top             =   1770
      Width           =   300
   End
   Begin VB.CheckBox CEDIT 
      DataField       =   "cedit"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   0
      Left            =   1095
      TabIndex        =   7
      Top             =   2415
      Width           =   300
   End
   Begin VB.CheckBox CBROWSE 
      DataField       =   "cbrowse"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   0
      Left            =   1095
      TabIndex        =   5
      Top             =   1770
      Width           =   300
   End
   Begin VB.CheckBox BEDIT 
      DataField       =   "bedit"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   5025
      TabIndex        =   37
      Top             =   3585
      Width           =   300
   End
   Begin VB.CheckBox BDELETE 
      DataField       =   "bdelete"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   5025
      TabIndex        =   36
      Top             =   3255
      Width           =   300
   End
   Begin VB.CheckBox BBROWSE 
      DataField       =   "bbrowse"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Left            =   5025
      TabIndex        =   35
      Top             =   2940
      Width           =   300
   End
   Begin VB.Label lblLabels 
      Caption         =   "DELETE"
      Height          =   255
      Index           =   65
      Left            =   4185
      TabIndex        =   156
      Top             =   6435
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   "EDIT"
      Height          =   255
      Index           =   66
      Left            =   4185
      TabIndex        =   155
      Top             =   6765
      Width           =   735
   End
   Begin VB.Label lblLabels 
      Caption         =   "PRINT"
      Height          =   255
      Index           =   67
      Left            =   5385
      TabIndex        =   154
      Top             =   6120
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "REPORT"
      Height          =   255
      Index           =   68
      Left            =   5385
      TabIndex        =   153
      Top             =   6435
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "PRODUCT SUPERVISOR"
      Height          =   255
      Index           =   69
      Left            =   5385
      TabIndex        =   152
      Top             =   6765
      Width           =   2175
   End
   Begin VB.Label lblLabels 
      Caption         =   "BROWSE"
      Height          =   255
      Index           =   70
      Left            =   4185
      TabIndex        =   151
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label Label9 
      Caption         =   "RMS RECORDS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   3870
      TabIndex        =   150
      Top             =   5640
      Width           =   3495
   End
   Begin VB.Label lblLabels 
      Caption         =   "DELETE"
      Height          =   255
      Index           =   64
      Left            =   4185
      TabIndex        =   143
      Top             =   4875
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   "EDIT"
      Height          =   255
      Index           =   63
      Left            =   4185
      TabIndex        =   142
      Top             =   5205
      Width           =   735
   End
   Begin VB.Label lblLabels 
      Caption         =   "PRINT"
      Height          =   255
      Index           =   62
      Left            =   5385
      TabIndex        =   141
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "REPORT"
      Height          =   255
      Index           =   61
      Left            =   5385
      TabIndex        =   140
      Top             =   4875
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "PRODUCT SUPERVISOR"
      Height          =   255
      Index           =   60
      Left            =   5385
      TabIndex        =   139
      Top             =   5205
      Width           =   2175
   End
   Begin VB.Label lblLabels 
      Caption         =   "BROWSE"
      Height          =   255
      Index           =   59
      Left            =   4185
      TabIndex        =   138
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "VICTIM'S ADVOCATE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   3870
      TabIndex        =   137
      Top             =   4080
      Width           =   3495
   End
   Begin VB.Label lblLabels 
      Caption         =   "ORI/AGENCY #:"
      Height          =   255
      Index           =   58
      Left            =   5280
      TabIndex        =   136
      Top             =   540
      Width           =   1815
   End
   Begin VB.Shape Shape4 
      Height          =   1470
      Left            =   7825
      Top             =   960
      Width           =   3840
   End
   Begin VB.Shape Shape3 
      Height          =   3105
      Left            =   7825
      Top             =   2445
      Width           =   3840
   End
   Begin VB.Label lblLabels 
      Caption         =   "DELETE"
      Height          =   255
      Index           =   57
      Left            =   8040
      TabIndex        =   135
      Top             =   1755
      Width           =   750
   End
   Begin VB.Label lblLabels 
      Caption         =   "EDIT"
      Height          =   255
      Index           =   56
      Left            =   8040
      TabIndex        =   134
      Top             =   2085
      Width           =   765
   End
   Begin VB.Label lblLabels 
      Caption         =   "PRINT"
      Height          =   255
      Index           =   55
      Left            =   9240
      TabIndex        =   133
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "REPORT"
      Height          =   255
      Index           =   54
      Left            =   9240
      TabIndex        =   132
      Top             =   1755
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "PRODUCT SUPERVISOR"
      Height          =   255
      Index           =   53
      Left            =   9240
      TabIndex        =   131
      Top             =   2085
      Width           =   2175
   End
   Begin VB.Label lblLabels 
      Caption         =   "BROWSE"
      Height          =   255
      Index           =   52
      Left            =   8040
      TabIndex        =   130
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "DETENTION CENTER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   7845
      TabIndex        =   129
      Top             =   960
      Width           =   2760
   End
   Begin VB.Shape Shape2 
      Height          =   6135
      Left            =   3840
      Top             =   960
      Width           =   3960
   End
   Begin VB.Shape Shape1 
      Height          =   6150
      Left            =   90
      Top             =   960
      Width           =   3675
   End
   Begin VB.Label lblLabels 
      Caption         =   "FULL NAME"
      Height          =   255
      Index           =   51
      Left            =   15
      TabIndex        =   128
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "PRODUCT SUPERVISOR"
      Height          =   255
      Index           =   45
      Left            =   1455
      TabIndex        =   127
      Top             =   3855
      Width           =   2175
   End
   Begin VB.Label lblLabels 
      Caption         =   "REPORT"
      Height          =   255
      Index           =   46
      Left            =   1455
      TabIndex        =   126
      Top             =   3525
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "PRINT"
      Height          =   255
      Index           =   47
      Left            =   1455
      TabIndex        =   125
      Top             =   3210
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "EDIT"
      Height          =   255
      Index           =   48
      Left            =   255
      TabIndex        =   124
      Top             =   3855
      Width           =   1005
   End
   Begin VB.Label lblLabels 
      Caption         =   "DELETE"
      Height          =   255
      Index           =   49
      Left            =   255
      TabIndex        =   123
      Top             =   3525
      Width           =   1005
   End
   Begin VB.Label lblLabels 
      Caption         =   "BROWSE"
      Height          =   255
      Index           =   50
      Left            =   255
      TabIndex        =   122
      Top             =   3210
      Width           =   1005
   End
   Begin VB.Label Label6 
      Caption         =   "Writs/Other Papers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Index           =   3
      Left            =   255
      TabIndex        =   121
      Top             =   2850
      Width           =   3765
   End
   Begin VB.Label Label6 
      Caption         =   "Executions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Index           =   2
      Left            =   210
      TabIndex        =   120
      Top             =   5595
      Width           =   3450
   End
   Begin VB.Label lblLabels 
      Caption         =   "BROWSE"
      Height          =   255
      Index           =   44
      Left            =   255
      TabIndex        =   119
      Top             =   6090
      Width           =   1005
   End
   Begin VB.Label lblLabels 
      Caption         =   "DELETE"
      Height          =   255
      Index           =   43
      Left            =   255
      TabIndex        =   118
      Top             =   6405
      Width           =   1005
   End
   Begin VB.Label lblLabels 
      Caption         =   "EDIT"
      Height          =   255
      Index           =   42
      Left            =   255
      TabIndex        =   117
      Top             =   6735
      Width           =   1005
   End
   Begin VB.Label lblLabels 
      Caption         =   "PRINT"
      Height          =   255
      Index           =   41
      Left            =   1455
      TabIndex        =   116
      Top             =   6090
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "REPORT"
      Height          =   255
      Index           =   40
      Left            =   1455
      TabIndex        =   115
      Top             =   6405
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "PRODUCT SUPERVISOR"
      Height          =   255
      Index           =   39
      Left            =   1455
      TabIndex        =   114
      Top             =   6735
      Width           =   2175
   End
   Begin VB.Label Label6 
      Caption         =   "Family Court Papers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   113
      Top             =   4290
      Width           =   3870
   End
   Begin VB.Label lblLabels 
      Caption         =   "BROWSE"
      Height          =   255
      Index           =   32
      Left            =   255
      TabIndex        =   112
      Top             =   4650
      Width           =   1005
   End
   Begin VB.Label lblLabels 
      Caption         =   "DELETE"
      Height          =   255
      Index           =   31
      Left            =   255
      TabIndex        =   111
      Top             =   4965
      Width           =   1005
   End
   Begin VB.Label lblLabels 
      Caption         =   "EDIT"
      Height          =   255
      Index           =   30
      Left            =   255
      TabIndex        =   110
      Top             =   5295
      Width           =   1005
   End
   Begin VB.Label lblLabels 
      Caption         =   "PRINT"
      Height          =   255
      Index           =   29
      Left            =   1455
      TabIndex        =   109
      Top             =   4650
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "REPORT"
      Height          =   255
      Index           =   28
      Left            =   1455
      TabIndex        =   108
      Top             =   4965
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "PRODUCT SUPERVISOR"
      Height          =   255
      Index           =   27
      Left            =   1455
      TabIndex        =   107
      Top             =   5295
      Width           =   2175
   End
   Begin VB.Label Label6 
      Caption         =   "Magistrate Papers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Index           =   0
      Left            =   255
      TabIndex        =   106
      Top             =   1410
      Width           =   3930
   End
   Begin VB.Label Label2 
      Caption         =   "BOOKING MANAGER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   3875
      TabIndex        =   105
      Top             =   2460
      Width           =   3495
   End
   Begin VB.Label lblLabels 
      Caption         =   "BROWSE"
      Height          =   255
      Index           =   0
      Left            =   4185
      TabIndex        =   104
      Top             =   2940
      Width           =   1095
   End
   Begin VB.Label lblLabels 
      Caption         =   "USER ID:"
      Height          =   255
      Index           =   26
      Left            =   0
      TabIndex        =   103
      Top             =   105
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "WARRANT MANAGER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   7845
      TabIndex        =   99
      Top             =   4110
      Width           =   4215
   End
   Begin VB.Label lblLabels 
      Caption         =   "PRODUCT SUPERVISOR"
      Height          =   255
      Index           =   24
      Left            =   9315
      TabIndex        =   98
      Top             =   5205
      Width           =   2175
   End
   Begin VB.Label lblLabels 
      Caption         =   "REPORT"
      Height          =   255
      Index           =   23
      Left            =   9315
      TabIndex        =   97
      Top             =   4905
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "PRINT"
      Height          =   255
      Index           =   22
      Left            =   9315
      TabIndex        =   96
      Top             =   4605
      Width           =   735
   End
   Begin VB.Label lblLabels 
      Caption         =   "EDIT"
      Height          =   255
      Index           =   21
      Left            =   8115
      TabIndex        =   95
      Top             =   5205
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "DELETE"
      Height          =   255
      Index           =   20
      Left            =   8115
      TabIndex        =   94
      Top             =   4905
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "BROWSE"
      Height          =   255
      Index           =   19
      Left            =   8115
      TabIndex        =   93
      Top             =   4605
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "RESTRAINING ORDER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   7845
      TabIndex        =   92
      Top             =   2460
      Width           =   4215
   End
   Begin VB.Label lblLabels 
      Caption         =   "PRODUCT SUPERVISOR"
      Height          =   255
      Index           =   17
      Left            =   9315
      TabIndex        =   91
      Top             =   3585
      Width           =   2175
   End
   Begin VB.Label lblLabels 
      Caption         =   "REPORT"
      Height          =   255
      Index           =   16
      Left            =   9315
      TabIndex        =   90
      Top             =   3255
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "PRINT"
      Height          =   255
      Index           =   15
      Left            =   9315
      TabIndex        =   89
      Top             =   2940
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "EDIT"
      Height          =   255
      Index           =   14
      Left            =   8115
      TabIndex        =   88
      Top             =   3585
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "DELETE"
      Height          =   255
      Index           =   13
      Left            =   8115
      TabIndex        =   87
      Top             =   3255
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "BROWSE"
      Height          =   255
      Index           =   12
      Left            =   8115
      TabIndex        =   86
      Top             =   2940
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "INCIDENT REPORT/OPTIONS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   3875
      TabIndex        =   85
      Top             =   960
      Width           =   4530
   End
   Begin VB.Label lblLabels 
      Caption         =   "PRODUCT SUPERVISOR"
      Height          =   255
      Index           =   38
      Left            =   5385
      TabIndex        =   84
      Top             =   2025
      Width           =   2175
   End
   Begin VB.Label lblLabels 
      Caption         =   "REPORT"
      Height          =   255
      Index           =   37
      Left            =   5385
      TabIndex        =   83
      Top             =   1695
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "PRINT"
      Height          =   255
      Index           =   36
      Left            =   5385
      TabIndex        =   82
      Top             =   1380
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "EDIT"
      Height          =   255
      Index           =   35
      Left            =   4185
      TabIndex        =   81
      Top             =   2025
      Width           =   1005
   End
   Begin VB.Label lblLabels 
      Caption         =   "DELETE"
      Height          =   255
      Index           =   34
      Left            =   4185
      TabIndex        =   80
      Top             =   1695
      Width           =   1005
   End
   Begin VB.Label lblLabels 
      Caption         =   "BROWSE"
      Height          =   255
      Index           =   33
      Left            =   4185
      TabIndex        =   79
      Top             =   1380
      Width           =   1005
   End
   Begin VB.Label lblLabels 
      Caption         =   "PRODUCT SUPERVISOR"
      Height          =   255
      Index           =   5
      Left            =   5385
      TabIndex        =   78
      Top             =   3585
      Width           =   2175
   End
   Begin VB.Label lblLabels 
      Caption         =   "REPORT"
      Height          =   255
      Index           =   4
      Left            =   5385
      TabIndex        =   77
      Top             =   3255
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "PRINT"
      Height          =   255
      Index           =   3
      Left            =   5385
      TabIndex        =   76
      Top             =   2940
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "EDIT"
      Height          =   255
      Index           =   2
      Left            =   4185
      TabIndex        =   75
      Top             =   3585
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "DELETE"
      Height          =   255
      Index           =   1
      Left            =   4185
      TabIndex        =   74
      Top             =   3255
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "CIVIL PROCESS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   120
      TabIndex        =   73
      Top             =   1035
      Width           =   4215
   End
   Begin VB.Label lblLabels 
      Caption         =   "LAW ENFORCEMENT SUITE SUPERVISOR LEVEL:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Index           =   25
      Left            =   9015
      TabIndex        =   72
      Top             =   105
      Width           =   2895
   End
   Begin VB.Label lblLabels 
      Caption         =   "PASSWORD:"
      Height          =   255
      Index           =   18
      Left            =   5280
      TabIndex        =   71
      Top             =   135
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "PRODUCT SUPERVISOR"
      Height          =   255
      Index           =   11
      Left            =   1455
      TabIndex        =   70
      Top             =   2415
      Width           =   2175
   End
   Begin VB.Label lblLabels 
      Caption         =   "REPORT"
      Height          =   255
      Index           =   10
      Left            =   1455
      TabIndex        =   69
      Top             =   2085
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "PRINT"
      Height          =   255
      Index           =   9
      Left            =   1455
      TabIndex        =   68
      Top             =   1770
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "EDIT"
      Height          =   255
      Index           =   8
      Left            =   255
      TabIndex        =   67
      Top             =   2415
      Width           =   1005
   End
   Begin VB.Label lblLabels 
      Caption         =   "DELETE"
      Height          =   255
      Index           =   7
      Left            =   255
      TabIndex        =   66
      Top             =   2085
      Width           =   1005
   End
   Begin VB.Label lblLabels 
      Caption         =   "BROWSE"
      Height          =   255
      Index           =   6
      Left            =   255
      TabIndex        =   65
      Top             =   1770
      Width           =   1005
   End
End
Attribute VB_Name = "SECURITY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
'Me.Height = 8895
On Error Resume Next
Dim db As Database, rs As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nwl + "lawsuite.mdb")
Set rs = db.OpenRecordset("select userid from security order by userid")
userid.clear
If Not rs.EOF Then
    rs.MoveFirst
    While Not rs.EOF
        userid.AddItem rs("userid")
        rs.MoveNext
    Wend
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

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub
Private Sub cmdAdd_Click()
If abrowse = 1 Or adelete = 1 Or aedit = 1 Or aprint = 1 Or areport = 1 Or asupervisor = 1 Then
    If IBROWSE = 0 And IDELETE = 0 And IEDIT = 0 And IPRINT = 0 And IREPORT = 0 And ISUPERVISOR = 0 Then
        MsgBox "Victim's Advocate logins must have some form of Incident Report authority.", 48, "Genesis Error Log"
        Exit Sub
    End If
End If
If BBROWSE = 1 Or BDELETE = 1 Or BEDIT = 1 Or BPRINT = 1 Or BREPORT = 1 Or BSUPERVISOR = 1 Then
    If IBROWSE = 0 And IDELETE = 0 And IEDIT = 0 And IPRINT = 0 And IREPORT = 0 And ISUPERVISOR = 0 Then
        MsgBox "Booking Report logins must have some form of Incident Report authority.", 48, "Genesis Error Log"
        Exit Sub
    End If
End If
If Len(userid) > 10 Or Len(password) > 10 Then
    msg = MsgBox("User Id and Password can be no more than 8 characters long.", 48, "Genesis Error Log")
    Exit Sub
End If
'RLB Code
    If Trim(userfullname.Text) = "" Then
        MsgBox "A user's Full Name must be entered.", 48, "Genesis Error Log"
        Exit Sub
    End If
'********
If Trim(orinumber) = "" Then
    MsgBox "An ORI Number or other agency identifying number must be entered.", 48, "Genesis Error Log"
    Exit Sub
End If
        
        
Screen.MousePointer = 11
Dim db As Database, rs As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nwl + "lawsuite.mdb")
userid = UCase(userid)
password = UCase(password)
    
If Right$(password, 1) = " " Then
    password = Left$(password, Len(password) - 1)
End If
Set rs = db.OpenRecordset("select * from security where userid = '" + userid + "'")
If rs.EOF Then
    userid.AddItem userid
    rs.AddNew
Else
    rs.MoveFirst
    rs.Edit
End If
rs("userid") = userid
tmp = ""
For t% = 1 To Len(password)
    tmp = tmp + Chr$(Asc(Mid$(password, t%, 1)) + 10)
Next t%
'RLB Code
If Trim(userfullname) <> "" Then
    rs("userfullname") = userfullname
End If
'********
'On Error Resume Next
On Error GoTo 0
rs("orinumber") = orinumber
rs("password") = tmp
rs("cbrowse") = CBROWSE(0)
rs("cdelete") = CDELETE(0)
rs("cedit") = CEDIT(0)
rs("cprint") = CPRINT(0)
rs("creport") = CREPORT(0)
rs("csupervisor") = CSUPERVISOR(0)
rs("cbrowsew") = CBROWSE(1)
rs("cdeletew") = CDELETE(1)
rs("ceditw") = CEDIT(1)
rs("cprintw") = CPRINT(1)
rs("creportw") = CREPORT(1)
rs("csupervisorw") = CSUPERVISOR(1)
rs("cbrowsef") = CBROWSE(2)
rs("cdeletef") = CDELETE(2)
rs("ceditf") = CEDIT(2)
rs("cprintf") = CPRINT(2)
rs("creportf") = CREPORT(2)
rs("csupervisorf") = CSUPERVISOR(2)
rs("cbrowsee") = CBROWSE(3)
rs("cdeletee") = CDELETE(3)
rs("cedite") = CEDIT(3)
rs("cprinte") = CPRINT(3)
rs("creporte") = CREPORT(3)
rs("csupervisore") = CSUPERVISOR(3)
rs("ibrowse") = IBROWSE
rs("idelete") = IDELETE
rs("iedit") = IEDIT
rs("iprint") = IPRINT
rs("ireport") = IREPORT
rs("isupervisor") = ISUPERVISOR
rs("abrowse") = abrowse
rs("adelete") = adelete
rs("aedit") = aedit
rs("aprint") = aprint
rs("areport") = areport
rs("asupervisor") = asupervisor
rs("rmsbrowse") = rmsbrowse
rs("rmsdelete") = rmsdelete
rs("rmsedit") = rmsedit
rs("rmsprint") = rmsprint
rs("rmsreport") = rmsreport
rs("rmssupervisor") = rmssupervisor
rs("jbrowse") = jbrowse
rs("jdelete") = jdelete
rs("jedit") = jedit
rs("jprint") = jprint
rs("jreport") = jreport
rs("jsupervisor") = jsupervisor
'RLB code - Bonnie suggest make service rights = incident rights
rs("sedit") = IEDIT
rs("sprint") = IPRINT
rs("sreport") = IREPORT
rs("ssupervisor") = ISUPERVISOR
'***
rs("bbrowse") = BBROWSE
rs("bdelete") = BDELETE
rs("bedit") = BEDIT
rs("bprint") = BPRINT
rs("breport") = BREPORT
rs("bsupervisor") = BSUPERVISOR
rs("rbrowse") = RBROWSE
rs("rdelete") = RDELETE
rs("redit") = REDIT
rs("rprint") = RPRINT
rs("rreport") = RREPORT
rs("rsupervisor") = RSUPERVISOR
rs("wbrowse") = WBROWSE
rs("wdelete") = WDELETE
rs("wedit") = WEDIT
rs("wprint") = WPRINT
rs("wreport") = WREPORT
rs("wsupervisor") = WSUPERVISOR
rs("supervisor") = SUPERVISOR
rs.Update
db.Close
On Error GoTo 0
Call nullfields
Call nullkeyfields
Screen.MousePointer = 0
On Error GoTo 0
Exit Sub
oderror:
Resume od
End Sub

Private Sub cmdDelete_Click()
Screen.MousePointer = 11
Dim db As Database, rs As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nwl + "lawsuite.mdb")
Set rs = db.OpenRecordset("select * from security where userid = '" + userid + "'")
If Not rs.EOF Then
    rs.MoveFirst
    rs.Delete
End If
db.Close
Call nullfields
Call nullkeyfields
Screen.MousePointer = 0
On Error GoTo 0
Exit Sub
oderror:
Resume od
End Sub

Private Sub cmdRefresh_Click()
  'This is only needed for multi user apps
  On Error GoTo RefreshErr
  datPrimaryRS.Refresh
  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub

Private Sub cmdUpdate_Click()
  On Error GoTo UpdateErr

  datPrimaryRS.Recordset.UpdateBatch adAffectAll
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub cmdClose_Click()
Unload SECURITY
End Sub

Private Sub txtFields_Change(index As Integer)
End Sub


Private Sub userid_Click()
If userid = "" Then
    Exit Sub
End If
Dim db As Database, rs As Recordset
On Error GoTo oderror
od:
Set db = OpenDatabase(nwl + "lawsuite.mdb")
huser = userid
Set rs = db.OpenRecordset("select * from security where userid = '" + userid + "'")
If rs.EOF Then
    Call nullfields
    db.Close
    Exit Sub
End If
rs.MoveFirst
Call nullfields
On Error Resume Next
'RLB Code
If Not IsNull(rs("userfullname")) Then
    userfullname.Text = rs("userfullname")
Else
    userfullname.Text = ""
End If
'********
If Not IsNull(rs("orinumber")) Then
    orinumber = rs("orinumber")
Else
    orinumber = ""
End If
CBROWSE(0) = rs("cbrowse")
CDELETE(0) = rs("cdelete")
CEDIT(0) = rs("cedit")
CPRINT(0) = rs("cprint")
CREPORT(0) = rs("creport")
CSUPERVISOR(0) = rs("csupervisor")
CBROWSE(1) = rs("cbrowsew")
CDELETE(1) = rs("cdeletew")
CEDIT(1) = rs("ceditw")
CPRINT(1) = rs("cprintw")
CREPORT(1) = rs("creportw")
CSUPERVISOR(1) = rs("csupervisorw")
CBROWSE(2) = rs("cbrowsef")
CDELETE(2) = rs("cdeletef")
CEDIT(2) = rs("ceditf")
CPRINT(2) = rs("cprintf")
CREPORT(2) = rs("creportf")
CSUPERVISOR(2) = rs("csupervisorf")
CBROWSE(3) = rs("cbrowsee")
CDELETE(3) = rs("cdeletee")
CEDIT(3) = rs("cedite")
CPRINT(3) = rs("cprinte")
CREPORT(3) = rs("creporte")
CSUPERVISOR(3) = rs("csupervisore")
IBROWSE = rs("ibrowse")
IDELETE = rs("idelete")
IEDIT = rs("iedit")
IPRINT = rs("iprint")
IREPORT = rs("ireport")
ISUPERVISOR = rs("isupervisor")
abrowse = rs("abrowse")
adelete = rs("adelete")
aedit = rs("aedit")
aprint = rs("aprint")
areport = rs("areport")
asupervisor = rs("asupervisor")
rmsbrowse = rs("rmsbrowse")
rmsdelete = rs("rmsdelete")
rmsedit = rs("rmsedit")
rmsprint = rs("rmsprint")
rmsreport = rs("rmsreport")
rmssupervisor = rs("rmssupervisor")
jbrowse = rs("jbrowse")
jdelete = rs("jdelete")
jedit = rs("jedit")
jprint = rs("jprint")
jreport = rs("jreport")
jsupervisor = rs("jsupervisor")
BBROWSE = rs("bbrowse")
BDELETE = rs("bdelete")
BEDIT = rs("bedit")
BPRINT = rs("bprint")
BREPORT = rs("breport")
BSUPERVISOR = rs("bsupervisor")
WBROWSE = rs("wbrowse")
WDELETE = rs("wdelete")
WEDIT = rs("wedit")
WPRINT = rs("wprint")
WREPORT = rs("wreport")
WSUPERVISOR = rs("wsupervisor")
RBROWSE = rs("rbrowse")
RDELETE = rs("rdelete")
REDIT = rs("redit")
RPRINT = rs("rprint")
RREPORT = rs("rreport")
RSUPERVISOR = rs("rsupervisor")
SUPERVISOR = rs("supervisor")
password = ""
For t% = 1 To Len(rs("password"))
    password = password + Chr$(Asc(Mid$(rs("password"), t%, 1)) - 10)
Next t%
db.Close
On Error GoTo 0
Exit Sub
oderror:
Resume od

End Sub
Private Sub nullfields()
For t% = 0 To 3
    CEDIT(t%) = 0
    CBROWSE(t%) = 0
    CDELETE(t%) = 0
    CPRINT(t%) = 0
    CREPORT(t%) = 0
    CSUPERVISOR(t%) = 0
Next t%
IBROWSE = 0
IDELETE = 0
IEDIT = 0
IPRINT = 0
IREPORT = 0
ISUPERVISOR = 0
abrowse = 0
adelete = 0
aedit = 0
aprint = 0
areport = 0
asupervisor = 0
rmsbrowse = 0
rmsdelete = 0
rmsedit = 0
rmsprint = 0
rmsreport = 0
rmssupervisor = 0
jbrowse = 0
jdelete = 0
jedit = 0
jprint = 0
jreport = 0
jsupervisor = 0
BBROWSE = 0
BDELETE = 0
BEDIT = 0
BPRINT = 0
BREPORT = 0
BSUPERVISOR = 0
WBROWSE = 0
WDELETE = 0
WEDIT = 0
WPRINT = 0
WREPORT = 0
WSUPERVISOR = 0
RBROWSE = 0
RDELETE = 0
REDIT = 0
RPRINT = 0
RREPORT = 0
RSUPERVISOR = 0
SUPERVISOR = 0
'********
End Sub
Private Sub nullkeyfields()
password = ""
userid = ""
userfullname = ""
orinumber = ""

End Sub
