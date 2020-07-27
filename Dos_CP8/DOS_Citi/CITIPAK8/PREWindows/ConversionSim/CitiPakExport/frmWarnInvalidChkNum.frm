VERSION 5.00
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "IMP32X30.OCX"
Begin VB.Form frmWarnInvalidChkNum 
   BorderStyle     =   0  'None
   Caption         =   "Invalid Check Number"
   ClientHeight    =   3480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6960
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3480
   ScaleWidth      =   6960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   3516
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6972
      _Version        =   196609
      _ExtentX        =   12298
      _ExtentY        =   6202
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
      BackColor       =   192
      Caption         =   ""
      FrameColor      =   192
      FrameThreeDHighlightColor=   8454143
      FrameThreeDShadowColor=   8454143
      FrameThreeDStyle=   2
      FrameThreeDWidth=   4
      FrameWidth      =   8
      Picture         =   "frmWarnInvalidChkNum.frx":0000
      Begin VB.CommandButton cmdOK 
         Caption         =   "F10 &OK"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   444
         Left            =   2754
         TabIndex        =   1
         Top             =   2352
         Width           =   1548
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Invalid Starting Check Number."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   13.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   348
         Left            =   1362
         TabIndex        =   4
         Top             =   1056
         Width           =   4284
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Please Enter a NEW Check Number."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   13.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   396
         Left            =   594
         TabIndex        =   3
         Top             =   1632
         Width           =   5868
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ERROR!"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   13.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   300
         Left            =   2592
         TabIndex        =   2
         Top             =   480
         Width           =   1980
      End
   End
End
Attribute VB_Name = "frmWarnInvalidChkNum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
  Unload frmWarnInvalidChkNum
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF10 Then Call cmdOK_Click
End Sub

Private Sub Form_Load()
  MainLog ("Invalid check number warning issued.")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    Call UnloadAllFormsAndOpn
    KillFile "prrun.opn"
    MainLog ("Payroll.exe terminated via menu bar on frmWarnInvalidChkNum.")
    End
  End If
End Sub

