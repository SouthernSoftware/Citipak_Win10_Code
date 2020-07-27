VERSION 5.00
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "IMP32X30.OCX"
Begin VB.Form frmBLCnvtHelpCustCat 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Business License Conversion: Help For Customer Conversion Version #1."
   ClientHeight    =   7440
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   10932
   Icon            =   "frmBLCnvtHelpCustCat.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   10932
   StartUpPosition =   1  'CenterOwner
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   7116
      Left            =   192
      TabIndex        =   0
      Top             =   180
      Width           =   10476
      _Version        =   196609
      _ExtentX        =   18478
      _ExtentY        =   12552
      _StockProps     =   70
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   12648447
      Caption         =   ""
      Picture         =   "frmBLCnvtHelpCustCat.frx":08CA
      Begin VB.CommandButton cmdExit 
         Caption         =   "ESC Exit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         Left            =   4080
         TabIndex        =   1
         Top             =   6240
         Width           =   2364
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "If category codes are not being converted then all category data for customers will be deleted."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   588
         Left            =   336
         TabIndex        =   9
         Top             =   1584
         Width           =   9756
      End
      Begin VB.Label Label5 
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "5. Version #1 customer code conversion zeroes out all customer balances."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   336
         TabIndex        =   8
         Top             =   5232
         Width           =   10044
      End
      Begin VB.Label Label4 
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "4. Version #1 customer code conversion indexes customer billing names, license numbers, sort names, customer numbers."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   540
         Left            =   336
         TabIndex        =   7
         Top             =   4608
         Width           =   10044
      End
      Begin VB.Label Label7 
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmBLCnvtHelpCustCat.frx":08E6
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   540
         Left            =   336
         TabIndex        =   6
         Top             =   3648
         Width           =   10044
      End
      Begin VB.Label Label10 
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmBLCnvtHelpCustCat.frx":0998
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   540
         Left            =   336
         TabIndex        =   5
         Top             =   3024
         Width           =   10044
      End
      Begin VB.Label Label11 
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "1. Version #1 customer code conversion looks for and reports blank license numbers."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   336
         TabIndex        =   4
         Top             =   2640
         Width           =   9804
      End
      Begin VB.Label Label9 
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmBLCnvtHelpCustCat.frx":0A1F
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   828
         Left            =   336
         TabIndex        =   3
         Top             =   624
         Width           =   9708
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "VERSION #1 CUSTOMER CODE CONVERSION"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   288
         TabIndex        =   2
         Top             =   288
         Width           =   9804
      End
   End
End
Attribute VB_Name = "frmBLCnvtHelpCustCat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
  Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      Call cmdExit_Click
      KeyCode = 0
  End Select

End Sub


