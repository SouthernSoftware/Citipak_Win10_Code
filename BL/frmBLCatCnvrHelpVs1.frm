VERSION 5.00
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "IMP32X30.OCX"
Begin VB.Form frmBLCatCnvrHelpVs1 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Business License Conversion: Help For Category Code Conversion Version #1"
   ClientHeight    =   4740
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   10932
   Icon            =   "frmBLCatCnvrHelpVs1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   10932
   StartUpPosition =   2  'CenterScreen
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   4332
      Left            =   240
      TabIndex        =   0
      Top             =   168
      Width           =   10476
      _Version        =   196609
      _ExtentX        =   18478
      _ExtentY        =   7641
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
      Picture         =   "frmBLCatCnvrHelpVs1.frx":08CA
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
         Left            =   4032
         TabIndex        =   1
         Top             =   3312
         Width           =   2364
      End
      Begin VB.Label Label2 
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmBLCatCnvrHelpVs1.frx":08E6
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
         Height          =   1404
         Left            =   336
         TabIndex        =   4
         Top             =   1776
         Width           =   10044
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "VERSION #1 CATEORY CODE CONVERSION"
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
         TabIndex        =   3
         Top             =   144
         Width           =   9804
      End
      Begin VB.Label Label4 
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "1. Version #1 category code conversion indexes category code numbers."
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
         Left            =   384
         TabIndex        =   2
         Top             =   1152
         Width           =   10044
      End
   End
End
Attribute VB_Name = "frmBLCatCnvrHelpVs1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
  Unload frmBLCatCnvrHelpVs1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      cmdExit_Click
      KeyCode = 0
  End Select

End Sub


