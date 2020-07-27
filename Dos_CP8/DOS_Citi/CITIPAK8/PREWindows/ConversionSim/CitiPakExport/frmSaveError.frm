VERSION 5.00
Begin VB.Form frmSaveError 
   BackColor       =   &H000000C0&
   Caption         =   "Save Error"
   ClientHeight    =   4680
   ClientLeft      =   48
   ClientTop       =   324
   ClientWidth     =   7848
   LinkTopic       =   "Form1"
   ScaleHeight     =   4680
   ScaleWidth      =   7848
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton OKCmd 
      BackColor       =   &H00C00000&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   2880
      MaskColor       =   &H0080FFFF&
      TabIndex        =   9
      Top             =   3840
      Width           =   1932
   End
   Begin VB.OptionButton optAbandon 
      BackColor       =   &H000000C0&
      Caption         =   "ABANDON"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   5280
      TabIndex        =   4
      ToolTipText     =   "Reloads form without any changes saved"
      Top             =   3000
      Width           =   1572
   End
   Begin VB.OptionButton optSave 
      BackColor       =   &H000000C0&
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   5280
      TabIndex        =   2
      ToolTipText     =   "Saves all changes and exits to previous menu"
      Top             =   1560
      Width           =   1572
   End
   Begin VB.OptionButton optReview 
      BackColor       =   &H000000C0&
      Caption         =   "REVIEW"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   5280
      TabIndex        =   3
      ToolTipText     =   "Goes back to form to allow editing"
      Top             =   2280
      Width           =   1572
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000000C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2292
      Left            =   4920
      TabIndex        =   5
      Top             =   1320
      Width           =   2292
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
      Left            =   600
      TabIndex        =   8
      ToolTipText     =   "Reloads form with no changes saved"
      Top             =   3120
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
      Left            =   600
      TabIndex        =   7
      ToolTipText     =   "Goes back to form to allow editing"
      Top             =   2400
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
      Left            =   600
      TabIndex        =   6
      ToolTipText     =   "Saves form as is and exits to previous menu"
      Top             =   1680
      Width           =   3972
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
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
      Left            =   2400
      TabIndex        =   1
      Top             =   720
      Width           =   2892
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "WARNING DATA HAS BEEN CHANGED!"
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
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   5412
   End
End
Attribute VB_Name = "frmSaveError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Enum SaveChangeOptions
  scoInvalidOption = 0
  scoSaveChanges
  scoAbandonChanges
  scoReviewChanges
End Enum

Private m_scoOption As SaveChangeOptions

'// Create a property to get the Selection value.
'   NOTE: A Read-Only property has a Property Get but
'         no Property Let or Property Set
Property Get Selection() As SaveChangeOptions
  Selection = m_scoOption
End Property

Private Sub Form_Load()
  On Error Resume Next
  
  '// Clear all options on form load
  '   (so none are selected)
  optSave = False
  optReview = False
  optAbandon = False
End Sub

Private Sub OKCmd_Click()
   frmSaveError.Hide
End Sub

Private Sub optAbandon_Click()
  On Error Resume Next
  
  m_scoOption = scoAbandonChanges
End Sub

Private Sub optReview_Click()
  On Error Resume Next
  
  m_scoOption = scoReviewChanges
End Sub

Private Sub optSave_Click()
  On Error Resume Next
  
  m_scoOption = scoSaveChanges
End Sub
