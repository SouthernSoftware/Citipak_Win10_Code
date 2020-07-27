VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "BTN32A20.OCX"
Begin VB.Form frmBLMessageBoxJrWOpts 
   BackColor       =   &H000000C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Business License Message"
   ClientHeight    =   3168
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   6996
   Icon            =   "frmBLMessageBoxJrWOpts.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3168
   ScaleWidth      =   6996
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox fptxtChoice 
      Height          =   288
      Left            =   144
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   2256
      Width           =   492
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   684
      Left            =   1200
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2280
      Width           =   1884
      _Version        =   131072
      _ExtentX        =   3323
      _ExtentY        =   1206
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
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
      ButtonDesigner  =   "frmBLMessageBoxJrWOpts.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdCont 
      Height          =   684
      Left            =   3984
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2280
      Width           =   1884
      _Version        =   131072
      _ExtentX        =   3323
      _ExtentY        =   1206
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
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
      ButtonDesigner  =   "frmBLMessageBoxJrWOpts.frx":0AA5
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1164
      Left            =   408
      TabIndex        =   1
      Top             =   522
      Width           =   6156
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   1836
      Left            =   180
      Top             =   210
      Width           =   6636
   End
End
Attribute VB_Name = "frmBLMessageBoxJrWOpts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
'  Unload Me
  Me.Hide
  fptxtChoice = "abort"
End Sub

Private Sub cmdCont_Click()
  Me.Hide
  fptxtChoice = "continue"
End Sub

Private Sub Form_Load()
  fptxtChoice.Visible = False
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
  Case vbKeyEscape:
    Call cmdExit_Click
    KeyCode = 0
  Case vbKeyF10:
    Call cmdCont_Click
    KeyCode = 0
  Case Else:
  End Select
End Sub




