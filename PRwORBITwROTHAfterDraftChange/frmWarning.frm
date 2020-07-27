VERSION 5.00
Begin VB.Form frmWarning 
   BackColor       =   &H000000C0&
   BorderStyle     =   0  'None
   Caption         =   "Escape Warning"
   ClientHeight    =   5640
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8148
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5640
   ScaleWidth      =   8148
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAbandon 
      Caption         =   "E&xit  ABANDON CHANGES"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   2256
      TabIndex        =   7
      Top             =   4854
      Width           =   3420
   End
   Begin VB.CommandButton cmdReview 
      Caption         =   "F11 &REVIEW CHANGES"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   2256
      TabIndex        =   6
      Top             =   3414
      Width           =   3420
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "F10  &SAVE CHANGES"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   2304
      TabIndex        =   5
      Top             =   1974
      Width           =   3372
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "Select to Abandon Changes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   372
      Left            =   2040
      TabIndex        =   4
      ToolTipText     =   "Reloads form with no changes saved"
      Top             =   4374
      Width           =   3972
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "Select to Review Changes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   372
      Left            =   2040
      TabIndex        =   3
      ToolTipText     =   "Goes back to form to allow editing"
      Top             =   2934
      Width           =   3972
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "Select to Save Changes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   372
      Left            =   2040
      TabIndex        =   2
      ToolTipText     =   "Saves form as is and exits to previous menu"
      Top             =   1494
      Width           =   3972
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Abandon Changes?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   2544
      TabIndex        =   1
      Top             =   780
      Width           =   2892
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "WARNING: DATA HAS BEEN CHANGED!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   492
      Left            =   1440
      TabIndex        =   0
      Top             =   294
      Width           =   5412
   End
End
Attribute VB_Name = "frmWarning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Enum SaveChangeOptions11
  scoInvalidOption = 0
  scoSaveChanges
  scoAbandonChanges
  scoReviewChanges
End Enum

Private m_scoOption As SaveChangeOptions11

'// Create a property to get the Selection value.
'   NOTE: A Read-Only property has a Property Get but
'         no Property Let or Property Set
Property Get Selection() As SaveChangeOptions11
  Selection = m_scoOption
End Property

Private Sub cmdAbandon_Click()
'  On Error Resume Next
  m_scoOption = scoAbandonChanges
  Unload frmWarning
  MainLog ("Exit warning issued...abandon option chosen.")
End Sub

Private Sub cmdReview_Click()
'  On Error Resume Next
  m_scoOption = scoReviewChanges
  Unload frmWarning
  MainLog ("Exit warning issued...review option chosen.")

End Sub

Private Sub cmdSave_Click()
'  On Error Resume Next
  m_scoOption = scoSaveChanges
  Unload frmWarning
  MainLog ("Exit warning issued...save option chosen.")

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyF11:
      SendKeys "%R"
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%S"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%X"
      KeyCode = 0
    Case Else:
  End Select

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    Call UnloadAllFormsAndOpn
    KillFile "prrun.opn"
    MainLog ("Payroll.exe terminated via menu bar on frmWarning.")
    End
  End If
End Sub

