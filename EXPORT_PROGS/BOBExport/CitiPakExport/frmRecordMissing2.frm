VERSION 5.00
Begin VB.Form frmRecordMissing2 
   Caption         =   "Missing Record Advisory"
   ClientHeight    =   3336
   ClientLeft      =   48
   ClientTop       =   324
   ClientWidth     =   6576
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3336
   ScaleMode       =   0  'User
   ScaleWidth      =   7135.937
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEscape 
      Caption         =   "ESC &Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   2274
      TabIndex        =   0
      Top             =   2196
      Width           =   1980
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   " for assistance"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   300
      Left            =   1698
      TabIndex        =   3
      Top             =   1620
      Width           =   3228
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   2796
      Left            =   546
      Top             =   270
      Width           =   5484
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "Call Southern Software"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   300
      Left            =   1698
      TabIndex        =   2
      Top             =   1140
      Width           =   3228
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "File ""PRSTADEF.DAT"" cannot be found"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   1170
      TabIndex        =   1
      Top             =   516
      Width           =   4332
   End
End
Attribute VB_Name = "frmRecordMissing2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class

Public Enum MissingRecord2
  mrInvalidOption = 0
  mrEscape
End Enum

Private m_mrOption As MissingRecord2

'// Create a property to get the Selection value.
'   NOTE: A Read-Only property has a Property Get but
'         no Property Let or Property Set
Property Get Selection() As MissingRecord2
  Selection = m_mrOption
End Property

Private Sub cmdEscape_Click()
  On Error Resume Next
  m_mrOption = mrEscape
  Unload frmRecordMissing2
'  frmRecordMissing2.Hide

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%C"
      KeyCode = 0
    Case Else:
  End Select

End Sub

