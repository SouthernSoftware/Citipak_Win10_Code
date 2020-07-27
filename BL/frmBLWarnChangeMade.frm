VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmBLWarnChangeMade 
   BackColor       =   &H000000C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4395
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7410
   Icon            =   "frmBLWarnChangeMade.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4395
   ScaleWidth      =   7410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   355
      Left            =   6816
      Top             =   3696
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSave 
      Height          =   540
      Left            =   4608
      TabIndex        =   0
      Top             =   3288
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
      ButtonDesigner  =   "frmBLWarnChangeMade.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdReview 
      Height          =   540
      Left            =   1236
      TabIndex        =   1
      Top             =   3288
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
      ButtonDesigner  =   "frmBLWarnChangeMade.frx":0AA5
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   540
      Left            =   2922
      TabIndex        =   2
      Top             =   3288
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
      ButtonDesigner  =   "frmBLWarnChangeMade.frx":0C82
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(ESC)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   372
      Left            =   2424
      TabIndex        =   12
      Top             =   2328
      Width           =   696
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Press"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   1680
      TabIndex        =   11
      Top             =   2328
      Width           =   720
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "to REVIEW Changes."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   396
      Left            =   3216
      TabIndex        =   10
      Top             =   2328
      Width           =   3336
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(X)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   372
      Left            =   2448
      TabIndex        =   9
      Top             =   1920
      Width           =   528
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "to ABANDON Changes."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   3060
      TabIndex        =   8
      Top             =   1908
      Width           =   3312
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "to SAVE Changes."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   396
      Left            =   3264
      TabIndex        =   7
      Top             =   1512
      Width           =   2184
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(F10)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   372
      Left            =   2424
      TabIndex        =   6
      Top             =   1512
      Width           =   696
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Warning! data has been changed."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   16.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   564
      Left            =   168
      TabIndex        =   5
      Top             =   480
      Width           =   7080
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Press"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   1656
      TabIndex        =   4
      Top             =   1512
      Width           =   696
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Press"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   1680
      TabIndex        =   3
      Top             =   1920
      Width           =   720
   End
End
Attribute VB_Name = "frmBLWarnChangeMade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Enum SaveChangeOptions1
'begin by creating a new type = SaveChangeOptions1
  scoInvalidOption = 0 'enum member
  scoSaveChanges 'enum member
  scoAbandonChanges 'enum member
  scoReviewChanges 'enum member
End Enum

Private m_scoOption As SaveChangeOptions1
Property Get Selection() As SaveChangeOptions1
  'create a new property called Selection
  Selection = m_scoOption
End Property

Private Sub cmdReview_Click()
  'if review is chosen then the selection is scoReviewChanges
  m_scoOption = scoReviewChanges
  Unload frmBLWarnChangeMade
  MainLog ("Exit warning issued...review option chosen.")
End Sub

Private Sub cmdExit_Click()
  'if exit is chosen then the selection is scoAbandonChanges
  m_scoOption = scoAbandonChanges
  Unload frmBLWarnChangeMade
  MainLog ("Exit warning issued...abandon option chosen.")
End Sub

Private Sub cmdSave_Click()
  'if save is chosen then the selection is scoSaveChanges
  m_scoOption = scoSaveChanges
  Unload frmBLWarnChangeMade
  MainLog ("Exit warning issued...save option chosen.")
End Sub

Private Sub Timer1_Timer()
'the timer is set to 355 which means that everytime
'355 is reached this sub starts over...since tog
'is static it is remembered even though the sub closes
  Static tog As Boolean
  tog = Not tog
  If tog Then
    Me.BackColor = 210
  Else
    Me.BackColor = &HC0&
  End If
End Sub

Private Sub Form_Load()
  Dim RetVal As Long, winhand As Long
  MakeWindowTopMost Me.hwnd, True
  DoEvents
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      Call cmdReview_Click
      KeyCode = 0
    Case vbKeyX:
      Call cmdExit_Click
      KeyCode = 0
    Case vbKeyF10:
      Call cmdSave_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub



