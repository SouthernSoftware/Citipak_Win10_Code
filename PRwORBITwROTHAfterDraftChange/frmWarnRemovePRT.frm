VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Begin VB.Form frmWarnRemovePRT 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6810
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3750
   ScaleWidth      =   6810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   3744
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6816
      _Version        =   196609
      _ExtentX        =   12023
      _ExtentY        =   6604
      _StockProps     =   70
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.24
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
      FrameThreeDWidth=   4
      FrameWidth      =   8
      Picture         =   "frmWarnRemovePRT.frx":0000
      Begin VB.Timer Timer1 
         Interval        =   355
         Left            =   6336
         Top             =   96
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdDelete 
         Height          =   540
         Left            =   1182
         TabIndex        =   5
         Top             =   2688
         Width           =   2052
         _Version        =   131072
         _ExtentX        =   3619
         _ExtentY        =   952
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         GrayAreaColor   =   12632256
         BorderShowDefault=   -1  'True
         ButtonType      =   0
         NoPointerFocus  =   0   'False
         Value           =   0   'False
         GroupID         =   0
         GroupSelect     =   0
         DrawFocusRect   =   2
         DrawFocusRectCell=   -1
         GrayAreaPictureStyle=   0
         Static          =   0   'False
         BackStyle       =   1
         AutoSize        =   0
         AutoSizeOffsetTop=   0
         AutoSizeOffsetBottom=   0
         AutoSizeOffsetLeft=   0
         AutoSizeOffsetRight=   0
         DropShadowOffsetX=   3
         DropShadowOffsetY=   3
         DropShadowType  =   0
         DropShadowColor =   0
         Redraw          =   -1  'True
         ButtonDesigner  =   "frmWarnRemovePRT.frx":001C
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdESC 
         Height          =   540
         Left            =   3582
         TabIndex        =   6
         Top             =   2688
         Width           =   2052
         _Version        =   131072
         _ExtentX        =   3619
         _ExtentY        =   952
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         GrayAreaColor   =   12632256
         BorderShowDefault=   -1  'True
         ButtonType      =   0
         NoPointerFocus  =   0   'False
         Value           =   0   'False
         GroupID         =   0
         GroupSelect     =   0
         DrawFocusRect   =   2
         DrawFocusRectCell=   -1
         GrayAreaPictureStyle=   0
         Static          =   0   'False
         BackStyle       =   1
         AutoSize        =   0
         AutoSizeOffsetTop=   0
         AutoSizeOffsetBottom=   0
         AutoSizeOffsetLeft=   0
         AutoSizeOffsetRight=   0
         DropShadowOffsetX=   3
         DropShadowOffsetY=   3
         DropShadowType  =   0
         DropShadowColor =   0
         Redraw          =   -1  'True
         ButtonDesigner  =   "frmWarnRemovePRT.frx":0231
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Height          =   348
         Left            =   312
         TabIndex        =   7
         Top             =   1056
         Width           =   6204
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Remove this payroll transaction?"
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
         Height          =   348
         Left            =   1272
         TabIndex        =   4
         Top             =   624
         Width           =   4284
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Press ""F3"" to DELETE."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   300
         Left            =   474
         TabIndex        =   3
         Top             =   1584
         Width           =   5868
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Press ""ESC"" to Cancel."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   300
         Left            =   1794
         TabIndex        =   2
         Top             =   2016
         Width           =   3228
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "WARNING!"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   300
         Left            =   2424
         TabIndex        =   1
         Top             =   240
         Width           =   1980
      End
   End
End
Attribute VB_Name = "frmWarnRemovePRT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Public Enum PRTRemove
  prtrInvalidOption = 0
  prtrEscape
  prtrDelete
End Enum

Private m_prtrOption As PRTRemove

'// Create a property to get the Selection value.
'   NOTE: A Read-Only property has a Property Get but
'         no Property Let or Property Set
Property Get Selection() As PRTRemove
  Selection = m_prtrOption
End Property

Private Sub cmdEscape_Click()

End Sub

Private Sub cmdDelete_Click()
'  On Error Resume Next
  m_prtrOption = prtrDelete
  Unload frmWarnRemovePRT
  MainLog ("Employee deleted from edit transaction process.")
End Sub

Private Sub cmdESC_Click()
'  On Error Resume Next
  m_prtrOption = prtrEscape
  Unload frmWarnRemovePRT

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyF3:
      Call cmdDelete_Click
      KeyCode = 0
    Case vbKeyEscape:
      Call cmdESC_Click
      KeyCode = 0
    Case Else:
  End Select

End Sub

Private Sub Form_Load()
  If Exist("TEMPIF.dat") Then '2/3/04
    Label5.Caption = "Deleting will void current registers."
  End If

End Sub

Private Sub Timer1_Timer()
  Static tog As Boolean
  tog = Not tog
  If tog Then
    vaImprint1.BackColor = 210
  Else
    vaImprint1.BackColor = 192
  End If
End Sub

