VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Begin VB.Form frmWarnAccrueNow 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6810
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3732
   ScaleMode       =   0  'User
   ScaleWidth      =   6667.138
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   3744
      Left            =   -6
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
      FrameWidth      =   0
      Picture         =   "frmWarnAccrueNow.frx":0000
      Begin VB.Timer Timer1 
         Interval        =   355
         Left            =   6336
         Top             =   96
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdCont 
         Height          =   540
         Left            =   1350
         TabIndex        =   5
         Top             =   2685
         Width           =   1560
         _Version        =   131072
         _ExtentX        =   2752
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
         ButtonDesigner  =   "frmWarnAccrueNow.frx":001C
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdESC 
         Height          =   540
         Left            =   3888
         TabIndex        =   6
         Top             =   2688
         Width           =   1572
         _Version        =   131072
         _ExtentX        =   2773
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
         ButtonDesigner  =   "frmWarnAccrueNow.frx":0233
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
         Left            =   2412
         TabIndex        =   4
         Top             =   288
         Width           =   1980
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Press ""ESC"" to Cancel."
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
         Height          =   300
         Left            =   1716
         TabIndex        =   3
         Top             =   2016
         Width           =   3372
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Press ""F3"" to Continue."
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
         Height          =   300
         Left            =   468
         TabIndex        =   2
         Top             =   1632
         Width           =   5868
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Are you sure you want to post employee benefits?"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   732
         Left            =   1260
         TabIndex        =   1
         Top             =   768
         Width           =   4284
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "frmWarnAccrueNow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Public Enum PRTAccrue
  prtaInvalidOption = 0
  prtaEscape
  prtaContinue
End Enum

Private m_prtaOption As PRTAccrue

'// Create a property to get the Selection value.
'   NOTE: A Read-Only property has a Property Get but
'         no Property Let or Property Set
Property Get Selection() As PRTAccrue
  Selection = m_prtaOption
End Property

Private Sub cmdEscape_Click()

End Sub

Private Sub cmdCont_Click()
'  On Error Resume Next
  m_prtaOption = prtaContinue
  Unload frmWarnAccrueNow
  MainLog ("Accrual processing begun after warning.")
End Sub

Private Sub cmdESC_Click()
'  On Error Resume Next
  m_prtaOption = prtaEscape
  Unload frmWarnAccrueNow

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyF3:
      Call cmdCont_Click
      KeyCode = 0
    Case vbKeyEscape:
      Call cmdESC_Click
      KeyCode = 0
    Case Else:
  End Select

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

