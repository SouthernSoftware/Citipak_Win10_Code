VERSION 5.00
Begin VB.Form frmGLSetup 
   BackColor       =   &H008F8265&
   Caption         =   "frmG/L Setup Maintnenace"
   ClientHeight    =   2508
   ClientLeft      =   48
   ClientTop       =   324
   ClientWidth     =   3744
   LinkTopic       =   "Form1"
   ScaleHeight     =   8844
   ScaleWidth      =   12192
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdExitGLMaint 
      Caption         =   "E&xit G/L Maintenance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   2520
      TabIndex        =   13
      Top             =   6600
      Width           =   3612
   End
   Begin VB.CommandButton cmdGLClosingOp 
      Caption         =   "G/L &Closing Operations"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   2520
      TabIndex        =   12
      Top             =   5400
      Width           =   3612
   End
   Begin VB.CommandButton cmdGLSysConfigUtil 
      Caption         =   "System Configuration and &Utilities"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   2520
      TabIndex        =   11
      Top             =   6000
      Width           =   3612
   End
   Begin VB.CommandButton cmdDeptMaint 
      Caption         =   "&Department Maintenance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   2520
      TabIndex        =   10
      Top             =   3000
      Width           =   3612
   End
   Begin VB.CommandButton cmdSetPostDates 
      Caption         =   "&Set Allowable Posting Dates"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   2520
      TabIndex        =   9
      Top             =   4800
      Width           =   3612
   End
   Begin VB.CommandButton cmdPostXTrans 
      Caption         =   "&Post External Transactions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   2520
      TabIndex        =   8
      Top             =   4200
      Width           =   3612
   End
   Begin VB.CommandButton cmdBankMaint 
      Caption         =   "&Bank Maintenance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   2520
      TabIndex        =   7
      Top             =   3600
      Width           =   3612
   End
   Begin VB.CommandButton cmdChartofAccts 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Chart of &Accounts"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   2520
      TabIndex        =   6
      Top             =   2400
      Width           =   3612
   End
   Begin VB.CommandButton cmdFundMaint 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Fund Maintenance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   2520
      MaskColor       =   &H8000000F&
      TabIndex        =   5
      Top             =   1800
      Width           =   3612
   End
   Begin VB.Frame frmcolmn1 
      BackColor       =   &H00C0C0C0&
      Height          =   252
      Index           =   0
      Left            =   600
      TabIndex        =   4
      Top             =   1200
      Width           =   972
   End
   Begin VB.Frame frmcolmn2 
      BackColor       =   &H00C0C0C0&
      Height          =   252
      Index           =   1
      Left            =   7080
      TabIndex        =   3
      Top             =   1200
      Width           =   972
   End
   Begin VB.TextBox txtcolmn1 
      BackColor       =   &H8000000B&
      Enabled         =   0   'False
      Height          =   5892
      Index           =   0
      Left            =   720
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1440
      Width           =   732
   End
   Begin VB.TextBox txtcolmn2 
      BackColor       =   &H8000000B&
      Enabled         =   0   'False
      Height          =   5892
      Index           =   1
      Left            =   7200
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1440
      Width           =   732
   End
   Begin VB.Frame frmtop 
      Height          =   1212
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8652
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "GENERAL LEDGER MAINTENANCE MENU"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         Left            =   1440
         TabIndex        =   14
         Top             =   480
         Width           =   5772
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   360
      X2              =   8520
      Y1              =   600
      Y2              =   600
   End
End
Attribute VB_Name = "frmGLSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

