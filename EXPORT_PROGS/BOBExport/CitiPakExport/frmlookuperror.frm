VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "BTN32A20.OCX"
Begin VB.Form frmLookupError 
   BackColor       =   &H000000C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ERROR !!!"
   ClientHeight    =   3396
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   6744
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   10.8
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   3396
   ScaleWidth      =   6744
   StartUpPosition =   2  'CenterScreen
   Begin fpBtnAtlLibCtl.fpBtn cmdOK 
      Height          =   540
      Left            =   2592
      TabIndex        =   0
      Top             =   2448
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
      ButtonDesigner  =   "frmlookuperror.frx":0000
   End
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   6288
      Top             =   72
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Please reenter a new search and try again."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   420
      Left            =   264
      TabIndex        =   2
      Top             =   1632
      Width           =   6216
   End
   Begin VB.Label Label 
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
      ForeColor       =   &H0000FFFF&
      Height          =   564
      Left            =   264
      TabIndex        =   1
      Top             =   624
      Width           =   6216
   End
End
Attribute VB_Name = "frmLookupError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
  Call UnLoadWarning
End Sub

Private Sub Timer1_Timer()
  Dim BkColor As Integer
  Static tog As Boolean
  tog = Not tog
  If tog Then
    Me.BackColor = 210
  Else
    Me.BackColor = 192
  End If
End Sub

Private Sub Form_Load()
  Dim RetVal As Long, winhand As Long
  MakeWindowTopMost Me.hwnd, True
  DoEvents
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape, vbKeyReturn
      KeyCode = 0
      Call cmdOk_Click
    Case Else:
  End Select
End Sub
Private Sub UnLoadWarning()
  Unload frmLookupError
  DoEvents
End Sub
