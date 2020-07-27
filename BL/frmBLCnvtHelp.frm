VERSION 5.00
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "IMP32X30.OCX"
Begin VB.Form frmBLCnvtHelp 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Business License Conversion: Help For Full Conversion"
   ClientHeight    =   5532
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   10932
   Icon            =   "frmBLCnvtHelp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5532
   ScaleWidth      =   10932
   StartUpPosition =   1  'CenterOwner
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   5148
      Left            =   240
      TabIndex        =   0
      Top             =   168
      Width           =   10476
      _Version        =   196609
      _ExtentX        =   18478
      _ExtentY        =   9080
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
      Picture         =   "frmBLCnvtHelp.frx":08CA
      Begin VB.CommandButton cmdExit 
         Caption         =   "ESC E&xit"
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
         Left            =   4320
         TabIndex        =   1
         Top             =   4368
         Width           =   2364
      End
      Begin VB.Label Label2 
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmBLCnvtHelp.frx":08E6
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
         Height          =   828
         Left            =   288
         TabIndex        =   5
         Top             =   1536
         Width           =   10044
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "TRANSACTION CONVERSION"
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
         Left            =   1584
         TabIndex        =   4
         Top             =   96
         Width           =   7356
      End
      Begin VB.Label Label3 
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmBLCnvtHelp.frx":09AF
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
         Height          =   1356
         Left            =   288
         TabIndex        =   3
         Top             =   2496
         Width           =   10044
      End
      Begin VB.Label Label9 
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmBLCnvtHelp.frx":0B92
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
         TabIndex        =   2
         Top             =   432
         Width           =   9708
      End
   End
End
Attribute VB_Name = "frmBLCnvtHelp"
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
      SendKeys "%X"
      KeyCode = 0
  End Select

End Sub

