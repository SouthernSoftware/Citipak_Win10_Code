VERSION 5.00
Begin VB.Form lineup 
   Caption         =   "Line Up"
   ClientHeight    =   7350
   ClientLeft      =   210
   ClientTop       =   795
   ClientWidth     =   9975
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7350
   ScaleWidth      =   9975
   Begin VB.Frame lineupframe 
      BackColor       =   &H00404000&
      Caption         =   "LINE UP"
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
      Height          =   7300
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   9795
      Begin VB.Frame bigpic 
         Height          =   6000
         Left            =   9720
         TabIndex        =   28
         Top             =   2640
         Visible         =   0   'False
         Width           =   6255
         Begin VB.CommandButton Command1 
            Caption         =   "CLOSE"
            Height          =   315
            Left            =   5160
            TabIndex        =   29
            Top             =   5520
            Width           =   990
         End
         Begin VB.Image bigimage 
            Height          =   5175
            Left            =   120
            Stretch         =   -1  'True
            Top             =   240
            Width           =   6015
         End
      End
      Begin VB.CommandButton cmdPrintLineup 
         Caption         =   "PRINT"
         Height          =   315
         Left            =   5085
         TabIndex        =   32
         Top             =   6720
         Width           =   975
      End
      Begin VB.CheckBox chkLineUpReport 
         BackColor       =   &H00404000&
         Caption         =   "Lineup Report"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   3000
         TabIndex        =   31
         Top             =   6840
         Width           =   1770
      End
      Begin VB.CheckBox chkIndPictures 
         BackColor       =   &H00404000&
         Caption         =   "Individual Mugshot"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   3000
         TabIndex        =   30
         Top             =   6540
         Width           =   1890
      End
      Begin VB.CommandButton priorscreen 
         Caption         =   "< < < < <"
         Height          =   345
         Left            =   120
         TabIndex        =   2
         Top             =   6720
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.CommandButton nextscreen 
         Caption         =   "> > > > >"
         Height          =   345
         Left            =   8280
         TabIndex        =   1
         Top             =   6720
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.Label tinc4 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   9
         Left            =   7920
         TabIndex        =   72
         Top             =   5640
         Width           =   900
      End
      Begin VB.Label tinc3 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   9
         Left            =   8880
         TabIndex        =   71
         Top             =   5400
         Width           =   900
      End
      Begin VB.Label tinc2 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   9
         Left            =   7920
         TabIndex        =   70
         Top             =   5400
         Width           =   900
      End
      Begin VB.Label tinc 
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
         Index           =   9
         Left            =   7920
         TabIndex        =   69
         Top             =   5160
         Width           =   1800
      End
      Begin VB.Image thumbnail 
         Height          =   1575
         Index           =   9
         Left            =   7920
         Stretch         =   -1  'True
         ToolTipText     =   "Click to Select for Lineup"
         Top             =   3360
         Width           =   1545
      End
      Begin VB.Shape selectbox 
         BorderColor     =   &H00FFFFFF&
         Height          =   1815
         Index           =   9
         Left            =   7800
         Top             =   3240
         Visible         =   0   'False
         Width           =   1770
      End
      Begin VB.Label tinc4 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   8
         Left            =   6000
         TabIndex        =   68
         Top             =   5640
         Width           =   900
      End
      Begin VB.Label tinc3 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   8
         Left            =   6960
         TabIndex        =   67
         Top             =   5400
         Width           =   900
      End
      Begin VB.Label tinc2 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   8
         Left            =   6000
         TabIndex        =   66
         Top             =   5400
         Width           =   900
      End
      Begin VB.Label tinc 
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
         Index           =   8
         Left            =   6000
         TabIndex        =   65
         Top             =   5160
         Width           =   1800
      End
      Begin VB.Image thumbnail 
         Height          =   1575
         Index           =   8
         Left            =   6000
         Stretch         =   -1  'True
         ToolTipText     =   "Click to Select for Lineup"
         Top             =   3360
         Width           =   1545
      End
      Begin VB.Shape selectbox 
         BorderColor     =   &H00FFFFFF&
         Height          =   1815
         Index           =   8
         Left            =   5880
         Top             =   3240
         Visible         =   0   'False
         Width           =   1770
      End
      Begin VB.Label tinc4 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   7
         Left            =   4080
         TabIndex        =   64
         Top             =   5640
         Width           =   900
      End
      Begin VB.Label tinc3 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   7
         Left            =   5040
         TabIndex        =   63
         Top             =   5400
         Width           =   900
      End
      Begin VB.Label tinc2 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   7
         Left            =   4080
         TabIndex        =   62
         Top             =   5400
         Width           =   900
      End
      Begin VB.Label tinc 
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
         Index           =   7
         Left            =   4080
         TabIndex        =   61
         Top             =   5160
         Width           =   1800
      End
      Begin VB.Image thumbnail 
         Height          =   1575
         Index           =   7
         Left            =   4080
         Stretch         =   -1  'True
         ToolTipText     =   "Click to Select for Lineup"
         Top             =   3360
         Width           =   1545
      End
      Begin VB.Shape selectbox 
         BorderColor     =   &H00FFFFFF&
         Height          =   1815
         Index           =   7
         Left            =   3960
         Top             =   3240
         Visible         =   0   'False
         Width           =   1770
      End
      Begin VB.Label tinc4 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   6
         Left            =   2160
         TabIndex        =   60
         Top             =   5640
         Width           =   900
      End
      Begin VB.Label tinc3 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   6
         Left            =   3120
         TabIndex        =   59
         Top             =   5400
         Width           =   900
      End
      Begin VB.Label tinc2 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   6
         Left            =   2160
         TabIndex        =   58
         Top             =   5400
         Width           =   900
      End
      Begin VB.Label tinc 
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
         Index           =   6
         Left            =   2160
         TabIndex        =   57
         Top             =   5160
         Width           =   1800
      End
      Begin VB.Image thumbnail 
         Height          =   1575
         Index           =   6
         Left            =   2160
         Stretch         =   -1  'True
         ToolTipText     =   "Click to Select for Lineup"
         Top             =   3360
         Width           =   1545
      End
      Begin VB.Shape selectbox 
         BorderColor     =   &H00FFFFFF&
         Height          =   1815
         Index           =   6
         Left            =   2040
         Top             =   3240
         Visible         =   0   'False
         Width           =   1770
      End
      Begin VB.Label tinc4 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   56
         Top             =   5640
         Width           =   900
      End
      Begin VB.Label tinc3 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   1200
         TabIndex        =   55
         Top             =   5400
         Width           =   900
      End
      Begin VB.Label tinc2 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   54
         Top             =   5400
         Width           =   900
      End
      Begin VB.Label tinc 
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
         Index           =   5
         Left            =   240
         TabIndex        =   53
         Top             =   5160
         Width           =   1800
      End
      Begin VB.Image thumbnail 
         Height          =   1575
         Index           =   5
         Left            =   240
         Stretch         =   -1  'True
         ToolTipText     =   "Click to Select for Lineup"
         Top             =   3360
         Width           =   1545
      End
      Begin VB.Shape selectbox 
         BorderColor     =   &H00FFFFFF&
         Height          =   1815
         Index           =   5
         Left            =   120
         Top             =   3240
         Visible         =   0   'False
         Width           =   1770
      End
      Begin VB.Label tinc4 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   7920
         TabIndex        =   52
         Top             =   2640
         Width           =   900
      End
      Begin VB.Label tinc3 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   8880
         TabIndex        =   51
         Top             =   2400
         Width           =   900
      End
      Begin VB.Label tinc2 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   7920
         TabIndex        =   50
         Top             =   2400
         Width           =   900
      End
      Begin VB.Label tinc 
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
         Index           =   4
         Left            =   7920
         TabIndex        =   49
         Top             =   2160
         Width           =   1800
      End
      Begin VB.Image thumbnail 
         Height          =   1575
         Index           =   4
         Left            =   7920
         Stretch         =   -1  'True
         ToolTipText     =   "Click to Select for Lineup"
         Top             =   360
         Width           =   1545
      End
      Begin VB.Shape selectbox 
         BorderColor     =   &H00FFFFFF&
         Height          =   1815
         Index           =   4
         Left            =   7800
         Top             =   240
         Visible         =   0   'False
         Width           =   1770
      End
      Begin VB.Label tinc4 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   6000
         TabIndex        =   48
         Top             =   2640
         Width           =   900
      End
      Begin VB.Label tinc3 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   6960
         TabIndex        =   47
         Top             =   2400
         Width           =   900
      End
      Begin VB.Label tinc2 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   6000
         TabIndex        =   46
         Top             =   2400
         Width           =   900
      End
      Begin VB.Label tinc 
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
         Index           =   3
         Left            =   6000
         TabIndex        =   45
         Top             =   2160
         Width           =   1800
      End
      Begin VB.Image thumbnail 
         Height          =   1575
         Index           =   3
         Left            =   6000
         Stretch         =   -1  'True
         ToolTipText     =   "Click to Select for Lineup"
         Top             =   360
         Width           =   1545
      End
      Begin VB.Shape selectbox 
         BorderColor     =   &H00FFFFFF&
         Height          =   1815
         Index           =   3
         Left            =   5880
         Top             =   240
         Visible         =   0   'False
         Width           =   1770
      End
      Begin VB.Label tinc4 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   4080
         TabIndex        =   44
         Top             =   2640
         Width           =   900
      End
      Begin VB.Label tinc3 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   5040
         TabIndex        =   43
         Top             =   2400
         Width           =   900
      End
      Begin VB.Label tinc2 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   4080
         TabIndex        =   42
         Top             =   2400
         Width           =   900
      End
      Begin VB.Label tinc 
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
         Index           =   2
         Left            =   4080
         TabIndex        =   41
         Top             =   2160
         Width           =   1800
      End
      Begin VB.Image thumbnail 
         Height          =   1575
         Index           =   2
         Left            =   4080
         Stretch         =   -1  'True
         ToolTipText     =   "Click to Select for Lineup"
         Top             =   360
         Width           =   1545
      End
      Begin VB.Shape selectbox 
         BorderColor     =   &H00FFFFFF&
         Height          =   1815
         Index           =   2
         Left            =   3960
         Top             =   240
         Visible         =   0   'False
         Width           =   1770
      End
      Begin VB.Label tinc4 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   2160
         TabIndex        =   40
         Top             =   2640
         Width           =   900
      End
      Begin VB.Label tinc3 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   3120
         TabIndex        =   39
         Top             =   2400
         Width           =   900
      End
      Begin VB.Label tinc2 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   2160
         TabIndex        =   38
         Top             =   2400
         Width           =   900
      End
      Begin VB.Label tinc 
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
         Index           =   1
         Left            =   2160
         TabIndex        =   37
         Top             =   2160
         Width           =   1800
      End
      Begin VB.Image thumbnail 
         Height          =   1575
         Index           =   1
         Left            =   2160
         Stretch         =   -1  'True
         ToolTipText     =   "Click to Select for Lineup"
         Top             =   360
         Width           =   1545
      End
      Begin VB.Shape selectbox 
         BorderColor     =   &H00FFFFFF&
         Height          =   1815
         Index           =   1
         Left            =   2040
         Top             =   240
         Visible         =   0   'False
         Width           =   1770
      End
      Begin VB.Label tinc4 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   36
         Top             =   2640
         Width           =   900
      End
      Begin VB.Label tinc3 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   35
         Top             =   2400
         Width           =   900
      End
      Begin VB.Label tinc2 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   34
         Top             =   2400
         Width           =   900
      End
      Begin VB.Label tinc 
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
         Index           =   0
         Left            =   240
         TabIndex        =   33
         Top             =   2160
         Width           =   1800
      End
      Begin VB.Shape selectbox 
         BorderColor     =   &H00FFFFFF&
         Height          =   1815
         Index           =   0
         Left            =   120
         Top             =   240
         Visible         =   0   'False
         Width           =   1770
      End
      Begin VB.Image thumbnail 
         Height          =   1575
         Index           =   0
         Left            =   240
         Stretch         =   -1  'True
         ToolTipText     =   "Click to Select for Lineup"
         Top             =   360
         Width           =   1545
      End
      Begin VB.Label lineuppage 
         BackColor       =   &H00404000&
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   9000
         TabIndex        =   3
         Top             =   45
         Width           =   255
      End
   End
   Begin VB.Frame loframe 
      BackColor       =   &H00404000&
      Caption         =   "LINE UP OPTIONS FRAME"
      ForeColor       =   &H0000FFFF&
      Height          =   3495
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   8055
      Begin VB.TextBox fromage 
         Height          =   285
         Left            =   5400
         TabIndex        =   18
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox toage 
         Height          =   285
         Left            =   6720
         TabIndex        =   17
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox fromht 
         Height          =   285
         Left            =   5400
         TabIndex        =   16
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox toht 
         Height          =   285
         Left            =   6720
         TabIndex        =   15
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox fromwt 
         Height          =   285
         Left            =   5400
         TabIndex        =   14
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox towt 
         Height          =   285
         Left            =   6720
         TabIndex        =   13
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox numsuspects 
         Height          =   285
         Left            =   3120
         TabIndex        =   12
         Top             =   2880
         Width           =   975
      End
      Begin VB.CommandButton cmdCreateLineup 
         Caption         =   "Create LineUp"
         Height          =   615
         Left            =   5415
         TabIndex        =   11
         Top             =   2760
         Width           =   1095
      End
      Begin VB.CommandButton cmdCloseLineup 
         Caption         =   "Close"
         Height          =   615
         Left            =   6720
         TabIndex        =   10
         Top             =   2760
         Width           =   1095
      End
      Begin VB.ListBox race 
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
         ForeColor       =   &H00000000&
         Height          =   780
         Left            =   720
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   9
         Top             =   360
         Width           =   1740
      End
      Begin VB.ListBox sex 
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
         ForeColor       =   &H00000000&
         Height          =   780
         Left            =   735
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   8
         Top             =   1335
         Width           =   1740
      End
      Begin VB.ListBox ethnicity 
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
         ForeColor       =   &H00000000&
         Height          =   780
         Left            =   720
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   7
         Top             =   2280
         Width           =   1740
      End
      Begin VB.ListBox hair 
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
         ForeColor       =   &H00000000&
         Height          =   780
         Left            =   3000
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   6
         Top             =   360
         Width           =   1740
      End
      Begin VB.ListBox eyes 
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
         ForeColor       =   &H00000000&
         Height          =   780
         Left            =   3000
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   5
         Top             =   1335
         Width           =   1740
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Race"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Sex"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Age Range"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   5400
         TabIndex        =   25
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Ethnicity"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   75
         TabIndex        =   24
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Height Range"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   5400
         TabIndex        =   23
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Weight Range"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   5400
         TabIndex        =   22
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Hair"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   2520
         TabIndex        =   21
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Eyes"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   2520
         TabIndex        =   20
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Number of Suspects"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   3000
         TabIndex        =   19
         Top             =   2520
         Width           =   1455
      End
   End
End
Attribute VB_Name = "lineup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PICFILE, schanged As Integer, pics(5000), incs2(5000), incs3(5000), incs4(5000), nams(5000) As String, inct As Integer
Private Sub cmdCloseLineup_Click()
loframe.Visible = False
If lineupframe.Visible = False Then
    Unload Me
End If
End Sub

Private Sub cmdCloseLineupPics_Click()
lineupframe.Visible = False
If loframe.Visible = False Then
    Unload Me
End If
End Sub

Private Sub cmdCreateLineup_Click()
If Val(numsuspects) = 0 Then
    numsuspects = 999
End If
Dim qra, qse, qet, qha, qey, qag, qhe, qwe As String
qra = ""
qse = ""
qet = ""
qha = ""
qey = ""
qag = ""
qhe = ""
qwe = ""
race.ListIndex = -1
For Y% = 0 To race.ListCount - 1
    If race.Selected(Y%) Then
        If qra = "" Then
            qra = "(race = '" + race.List(Y%) + "' or race = '" + Left(race.List(Y%), 1) + "'"
        Else
            qra = qra + " or race = '" + race.List(Y%) + "' or race = '" + Left(race.List(Y%), 1) + "'"
        End If
    End If
Next Y%
sex.ListIndex = -1
For Y% = 0 To sex.ListCount - 1
    If sex.Selected(Y%) Then
        If qse = "" Then
            qse = "(sex = '" + sex.List(Y%) + "' or sex = '" + Left(sex.List(Y%), 1) + "'"
        Else
            qse = qse + " or sex = '" + sex.List(Y%) + "' or sex = '" + Left(sex.List(Y%), 1) + "'"
        End If
    End If
Next Y%
ethnicity.ListIndex = -1
For Y% = 0 To ethnicity.ListCount - 1
    If ethnicity.Selected(Y%) Then
        If qet = "" Then
            qet = "(ethnicity = '" + ethnicity.List(Y%) + "' or ethnicity = '" + Left(ethnicity.List(Y%), 1) + "'"
        Else
            qet = qet + " or ethnicity = '" + ethnicity.List(Y%) + "' or ethnicity = '" + Left(ethnicity.List(Y%), 1) + "'"
        End If
    End If
Next Y%
hair.ListIndex = -1
For Y% = 0 To hair.ListCount - 1
    If hair.Selected(Y%) Then
        If qha = "" Then
            qha = "(hair = '" + hair.List(Y%) + "'"
        Else
            qha = qha + " or hair = '" + hair.List(Y%) + "'"
        End If
    End If
Next Y%
eyes.ListIndex = -1
For Y% = 0 To eyes.ListCount - 1
    If eyes.Selected(Y%) Then
        If qey = "" Then
            qey = "(eyes = '" + eyes.List(Y%) + "'"
        Else
            qey = qey + " or eyes = '" + eyes.List(Y%) + "'"
        End If
    End If
Next Y%
If Val(toage) > 0 And Val(fromage) > 0 Then
    If qag = "" Then
        qag = "(age BETWEEN " + Format$(Val(fromage), "00") + " AND " + Format$(Val(toage), "00") + ")" 'Str$(t%), 2) + "',"
    Else
        qag = qag + " AND " + "(age BETWEEN (" + Format$(Val(toage), "00") + " AND " + Format$(Val(fromage), "00") + ")"
        End If
End If
If Val(toht) > 0 And Val(fromht) > 0 Then
    qhe = "(height between " + Chr$(34) + fromht + Chr$(34) + " and " + Chr$(34) + toht + Chr$(34) + ")"
End If
If Val(towt) > 0 And Val(fromwt) > 0 Then
    qwe = "(weight between '" + fromwt + "' and '" + towt + "')"
End If
If qra > "" Then
    qra = qra + ")"
End If
If qse > "" Then
    qse = qse + ")"
End If
If qet > "" Then
    qet = qet + ")"
End If
If qha > "" Then
    qha = qha + ")"
End If
If qey > "" Then
    qey = qey + ")"
End If
Dim buildsql As String
buildsql = ""
If qra > "" Then
    buildsql = qra + " and "
End If
If qse > "" Then
    buildsql = buildsql + qse + " and "
End If
If qet > "" Then
    buildsql = buildsql + qet + " and "
End If
If qha > "" Then
    buildsql = buildsql + qha + " and "
End If
If qey > "" Then
    buildsql = buildsql + qey + " and "
End If
If qag > "" Then
    buildsql = buildsql + qag + " and "
End If
If qhe > "" Then
    buildsql = buildsql + qhe + " and "
End If
If qwe > "" Then
    buildsql = buildsql + qwe + " and "
End If
Dim db As DAO.Database, rs As DAO.Recordset
Set db = DAO.OpenDatabase(nwl + "lawsuite.mdb")
Set rs = db.OpenRecordset("select * from people where " + buildsql + " mugshot is not null")
If Not rs.EOF Then
    rs.MoveFirst
Else
    db.Close
    msg = MsgBox("No suspects found matching the criteria entered.", 48, "Genesis Error Log")
    Exit Sub
End If
inct = 0
While Not rs.EOF
    inct = inct + 1
    If Not IsNull(rs("birthdate")) Then
        bd = rs("birthdate")
    Else
        bd = ""
    End If
    If Not IsNull(rs("idnumber")) Then
        idn = rs("idnumber")
    Else
        idn = ""
    End If
    If Not IsNull(rs("ssn")) Then
        ss = rs("ssn")
    Else
        ss = ""
    End If
    incs2(inct) = CStr(idn)
    incs3(inct) = ss
    incs4(inct) = (bd)
    nams(inct) = rs("dpnamelf")
    pics(inct) = rs("mugshot")
    If inct = Val(numsuspects) Then
        rs.MoveLast
    End If
    rs.MoveNext
Wend
db.Close
lineuppage = "1"
Call lineup
End Sub
Private Sub lineup()
lineupframe.Left = 0
lineupframe.Top = 0
vscrScreenScroller = 0
lineupframe.Visible = True
StartT% = ((Val(lineuppage) - 1) * 10) + 1
If inct >= StartT% + 9 Then
    stopt% = StartT% + 9
Else
    stopt% = inct
End If
tn% = 0
For t% = StartT% To stopt%
    Set thumbnail(tn%) = LoadPicture(pics(t%))
    thumbnail(tn%).Refresh
    tinc(tn%) = nams(t%)
    tinc2(tn%) = incs2(t%)
    tinc3(tn%) = incs4(t%)
    tinc4(tn%) = incs3(t%)
    tn% = tn% + 1
Next t%
For t% = tn% To 9
    Set thumbnail(t%) = LoadPicture()
    tinc(t%) = ""
    tinc2(t%) = ""
    tinc3(t%) = ""
    tinc4(t%) = ""
Next t%
If stopt% < inct Then
    nextscreen.Visible = True
Else
    nextscreen.Visible = False
End If
If StartT% > 1 Then
    priorscreen.Visible = True
Else
    priorscreen.Visible = False
End If


End Sub
Private Sub cmdPrintLineup_Click()
If chkIndPictures.Value = 0 And chkLineUpReport.Value = 0 Then
    msg = MsgBox("Please select a report type.", vbOKOnly, "Genesis Error Log")
    Exit Sub
End If
Screen.MousePointer = 11
Dim db As Database, rs As Recordset
Dim novisible As Boolean
novisible = True
For t% = 0 To 9
    If selectbox(t%).Visible = True Then
        'frmBookingReport.Visible = False
        
        If chkIndPictures.Value = 1 Then
            novisible = False
            idnumber = tinc2(t%)
            ssn = tinc4(t%)
            BIRTHDATE = tinc3(t%)
            
            Printer.FontBold = True
            Printer.FontSize = 24
            Printer.CurrentX = 3000
            Printer.Print "Mugshot Print"
            Printer.Print
            Printer.Print
            Printer.FontSize = 14
            Printer.Print Tab(15); "DEFENDANT NAME:"; Tab(50); nams(t% + 1)
            Printer.Print Tab(15); "ID NUMBER:"; Tab(50); idnumber
            Printer.Print Tab(15); "SOCIAL SECURITY NUMBER:"; Tab(50); ssn
            Printer.Print Tab(15); "BIRTHDATE:"; Tab(50); BIRTHDATE
            Printer.Print
            Printer.Print
            Printer.Print
            Printer.CurrentX = 2400
            Printer.PaintPicture thumbnail(t%).Picture, Printer.CurrentX, Printer.CurrentY
            Printer.EndDoc
            
        End If
    End If
Next t%

If chkLineUpReport.Value = 1 Then

    Printer.FontName = "Times New Roman"
    Printer.FontSize = 8
      
    lineuppage = 0
    

    For t% = 1 To inct Step 10
    
        Printer.Orientation = 2
        lineuppage = lineuppage + 1
        Call lineup
        Printer.FontSize = 24
        Printer.FontBold = True
        Printer.Print Tab(29); "Mugshot Portfolio"
        Printer.FontSize = 8
        Printer.FontBold = False
        PICS10 = 0
        CURRX = 500
        CURRY = 1500
        numpics = 0
        
        If t% + 9 > inct Then
            stoptt% = inct
        Else
            stoptt% = t% + 9
        End If
        
        For tt% = t% To stoptt%
        
            If numpics = 5 Then
                CURRX = 500
                CURRY = CURRY + 2800
                GoSub addinfotoprint
                CURRY = CURRY + 500
                numpics = 0
            End If
            
            Printer.PaintPicture thumbnail(PICS10).Picture, CURRX, CURRY, 2250, 2750
            PICS10 = PICS10 + 1
            numpics = numpics + 1
            CURRX = CURRX + 3000
            
        Next tt%
            
        If numpics > 0 Then
            CURRX = 500
            CURRY = CURRY + 2800
            tt% = stoptt%
            GoSub addinfotoprint
        End If
    
        If t% + 10 <= inct Then
            Printer.NewPage
        End If
    
    Next t%
    
End If


Printer.EndDoc
Screen.MousePointer = 0
If lineuppage > 1 Then
    lineuppage = 1
    Call lineup
End If

Exit Sub

addinfotoprint:
Printer.CurrentX = CURRX
Printer.CurrentY = CURRY
PICCT% = 0
For c% = tt% - numpics To tt% - 1
    PICCT% = PICCT% + 1
    Printer.Print Left$(incs2(c%), 20); Tab(10 + (46 * PICCT%));
Next c%
Printer.Print
Printer.CurrentX = CURRX
PICCT% = 0
For c% = tt% - numpics To tt% - 1
    PICCT% = PICCT% + 1
    Printer.Print Left$(nams(c%), 20); Tab(10 + (46 * PICCT%));
Next c%
Printer.Print
Return
End Sub

Private Sub Command1_Click()
bigpic.Visible = False
End Sub

Private Sub Form_Load()
Me.Height = 7750
Me.Width = 10095
Me.Top = 0
Me.Left = 0
Dim db As DAO.Database, rs As DAO.Recordset
Set db = DAO.OpenDatabase(nwl + "lawsuite.mdb")
race.clear
Set rs = db.OpenRecordset("select distinct race from people where race is not null")
If Not rs.EOF Then
    rs.MoveFirst
    While Not rs.EOF
        If Len(rs("race")) > 1 Then
            race.AddItem rs("race")
        End If
        rs.MoveNext
    Wend
End If
sex.clear
Set rs = db.OpenRecordset("select distinct sex from people where sex is not null")
If Not rs.EOF Then
    rs.MoveFirst
    While Not rs.EOF
        If Len(rs("sex")) > 1 Then
            sex.AddItem rs("sex")
        End If
        rs.MoveNext
    Wend
End If
ethnicity.clear
Set rs = db.OpenRecordset("select distinct ethnicity from people where ethnicity is not null")
If Not rs.EOF Then
    rs.MoveFirst
    While Not rs.EOF
        If Len(rs("ethnicity")) > 1 Then
            ethnicity.AddItem rs("ethnicity")
        End If
        rs.MoveNext
    Wend
End If
Set rs = db.OpenRecordset("select distinct hair from people where hair is not null")
hair.clear
If Not rs.EOF Then
    rs.MoveFirst
    While Not rs.EOF
        If rs("hair") > "" Then
            hair.AddItem rs("hair")
        End If
        rs.MoveNext
    Wend
End If
Set rs = db.OpenRecordset("select distinct eyes from people where eyes is not null")
eyes.clear
If Not rs.EOF Then
    rs.MoveFirst
    While Not rs.EOF
        If rs("eyes") > "" Then
            eyes.AddItem rs("eyes")
        End If
        rs.MoveNext
    Wend
End If
numsuspects = ""
fromage = ""
toage = ""
fromht = ""
toht = ""
fromwt = ""
towt = ""

End Sub

Private Sub Label1_Click()

End Sub

Private Sub nextscreen_Click()
lineuppage = Val(lineuppage) + 1
Call lineup
End Sub

Private Sub priorscreen_Click()
lineuppage = Val(lineuppage) - 1
Call lineup
End Sub

Private Sub thumbnail_Click(index As Integer)
If selectbox(index).Visible = True Then
    selectbox(index).Visible = False
Else
    selectbox(index).Visible = True
End If
End Sub

Private Sub thumbnail_DblClick(index As Integer)
bigimage.Picture = LoadPicture(pics(index + 1))
bigpic.Left = 1000
bigpic.Top = 1000
bigpic.Visible = True

End Sub

