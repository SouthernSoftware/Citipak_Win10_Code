VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Begin VB.Form frmWarnBadGLNumVer2 
   BackColor       =   &H000000C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6840
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4290
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   4320
      Left            =   0
      TabIndex        =   0
      Top             =   -48
      Width           =   6816
      _Version        =   196609
      _ExtentX        =   12023
      _ExtentY        =   7620
      _StockProps     =   70
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      Picture         =   "WarnBadGLNumVer2.frx":0000
      Begin VB.Timer Timer1 
         Interval        =   355
         Left            =   6336
         Top             =   96
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdReturn 
         Height          =   732
         Left            =   3600
         TabIndex        =   1
         Top             =   2880
         Width           =   2748
         _Version        =   131072
         _ExtentX        =   4847
         _ExtentY        =   1291
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
         ButtonDesigner  =   "WarnBadGLNumVer2.frx":001C
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdSave 
         Height          =   732
         Left            =   528
         TabIndex        =   2
         Top             =   2880
         Width           =   2748
         _Version        =   131072
         _ExtentX        =   4847
         _ExtentY        =   1291
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
         ButtonDesigner  =   "WarnBadGLNumVer2.frx":0251
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ERROR!"
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
         Left            =   2400
         TabIndex        =   4
         Top             =   432
         Width           =   1980
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   $"WarnBadGLNumVer2.frx":0485
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1356
         Left            =   576
         TabIndex        =   3
         Top             =   1104
         Width           =   5676
      End
   End
End
Attribute VB_Name = "frmWarnBadGLNumVer2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Enum BadGLNUM2Option
  badgl2InvalidOption = 0
  badgl2Return
  badgl2Save
'  badgl2GLList
End Enum

Private m_badgl2Option As BadGLNUMOption

'// Create a property to get the Selection value.
'   NOTE: A Read-Only property has a Property Get but
'         no Property Let or Property Set
Property Get Selection() As BadGLNUM2Option
  Selection = m_badgl2Option
End Property

'Private Sub cmdGLList_Click()
'  m_badgl2Option = badgl2GLList
'  Unload frmWarnBadGLNumVer2
'  DoEvents
'  MainLog ("GL List option activated on frmWarnBadGLNumVer2.")
'
'End Sub

Private Sub cmdReturn_Click()
'  On Error Resume Next
  m_badgl2Option = badgl2Return
  Unload frmWarnBadGLNumVer2
  DoEvents
  MainLog ("Return option activated on frmWarnBadGLNumVer2.")

End Sub

Private Sub cmdSave_Click()
'  On Error Resume Next
  m_badgl2Option = badgl2Save
  Unload frmWarnBadGLNumVer2
  DoEvents
  MainLog ("Save option activated on frmWarnBadGLNumVer2.")
  

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyF10:
      Call cmdReturn_Click
      KeyCode = 0
    Case vbKeyF9:
      Call cmdSave_Click
      KeyCode = 0
'    Case vbKeyF12:
'      Call cmdReturn_Click
'      KeyCode = 0
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

