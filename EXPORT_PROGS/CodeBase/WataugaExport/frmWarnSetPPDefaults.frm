VERSION 5.00
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "IMP32X30.OCX"
Begin VB.Form frmWarnSetPPDefaults 
   BorderStyle     =   0  'None
   ClientHeight    =   3588
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7056
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3588
   ScaleWidth      =   7056
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
      BackColor       =   192
      Caption         =   ""
      FrameColor      =   192
      FrameThreeDHighlightColor=   8454143
      FrameThreeDShadowColor=   8454143
      FrameThreeDStyle=   2
      FrameThreeDWidth=   4
      FrameWidth      =   8
      Picture         =   "frmWarnSetPPDefaults.frx":0000
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
         Left            =   2562
         TabIndex        =   4
         Top             =   480
         Width           =   1980
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Select ""Set Pay Period Defaults"""
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
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "NO Pay Period Defaults!"
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
         TabIndex        =   2
         Top             =   1056
         Width           =   4284
      End
   End
End
Attribute VB_Name = "frmWarnSetPPDefaults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class

Private Sub cmdOK_Click()
   Unload frmWarnSetPPDefaults
   DoEvents
   MainLog ("No payroll defaults saved warning issued.")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%O"
      KeyCode = 0
    Case Else:
  End Select

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    Call UnloadAllFormsAndOpn
    KillFile "prrun.opn"
    MainLog ("Payroll.exe terminated via menu bar on frmWarnSetPPDefaults.")
    End
  End If
End Sub

