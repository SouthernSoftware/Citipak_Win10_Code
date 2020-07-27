VERSION 5.00
Begin VB.Form frmGLMain 
   BackColor       =   &H008F8265&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Citipak General Ledger"
   ClientHeight    =   7368
   ClientLeft      =   3924
   ClientTop       =   1884
   ClientWidth     =   11340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7368
   ScaleWidth      =   11340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame frmtop 
      Height          =   1212
      Left            =   1800
      TabIndex        =   13
      Top             =   960
      Width           =   8652
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "GENERAL LEDGER MAIN MENU"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   1560
         TabIndex        =   14
         Top             =   480
         Width           =   5532
      End
   End
   Begin VB.TextBox txtcolmn2 
      BackColor       =   &H8000000B&
      Enabled         =   0   'False
      Height          =   5892
      Index           =   1
      Left            =   9000
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2400
      Width           =   732
   End
   Begin VB.TextBox txtcolmn1 
      BackColor       =   &H8000000B&
      Enabled         =   0   'False
      Height          =   5892
      Index           =   0
      Left            =   2520
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2400
      Width           =   732
   End
   Begin VB.Frame frmcolmn2 
      BackColor       =   &H00C0C0C0&
      Height          =   252
      Index           =   1
      Left            =   8880
      TabIndex        =   10
      Top             =   2160
      Width           =   972
   End
   Begin VB.Frame frmcolmn1 
      BackColor       =   &H00C0C0C0&
      Height          =   252
      Index           =   0
      Left            =   2400
      TabIndex        =   9
      Top             =   2160
      Width           =   972
   End
   Begin VB.CommandButton cmdGLSetup 
      BackColor       =   &H008F8265&
      Caption         =   "G/L  &Setup and Maintenance"
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
      Left            =   4320
      MaskColor       =   &H8000000F&
      TabIndex        =   0
      Top             =   2640
      Width           =   3612
   End
   Begin VB.CommandButton cmdBudgetMaint 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Budget Maintenance"
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
      Left            =   4320
      TabIndex        =   1
      Top             =   3240
      Width           =   3612
   End
   Begin VB.CommandButton cmdGenJournal 
      Caption         =   "&General Journal"
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
      Left            =   4320
      TabIndex        =   3
      Top             =   4440
      Width           =   3612
   End
   Begin VB.CommandButton cmdCashReceipts 
      Caption         =   "&Cash Receipts"
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
      Left            =   4320
      TabIndex        =   4
      Top             =   5040
      Width           =   3612
   End
   Begin VB.CommandButton cmdCashDisbursements 
      Caption         =   "Cash &Disbursements"
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
      Left            =   4320
      TabIndex        =   5
      Top             =   5640
      Width           =   3612
   End
   Begin VB.CommandButton dmdGetDistributions 
      Caption         =   "Get Distributions for &Interfacing"
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
      Left            =   4320
      TabIndex        =   2
      Top             =   3840
      Width           =   3612
   End
   Begin VB.CommandButton cmdGLReports 
      Caption         =   "G/L &Reports"
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
      Left            =   4320
      TabIndex        =   7
      Top             =   6840
      Width           =   3612
   End
   Begin VB.CommandButton cmdBankRecon 
      Caption         =   "Ban&k Reconcilliation"
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
      Left            =   4320
      TabIndex        =   6
      Top             =   6240
      Width           =   3612
   End
   Begin VB.CommandButton cmdExitGL 
      Caption         =   "E&xit G/L Menu"
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
      Left            =   4320
      TabIndex        =   8
      Top             =   7440
      Width           =   3612
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   2160
      X2              =   10320
      Y1              =   1560
      Y2              =   1560
   End
End
Attribute VB_Name = "frmGLMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



