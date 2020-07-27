VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmTaxMsgW4Opts 
   BackColor       =   &H000000C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Message"
   ClientHeight    =   4488
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   7008
   Icon            =   "frmTaxMsgW4Opts.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4488
   ScaleWidth      =   7008
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   240
      Top             =   2640
   End
   Begin VB.TextBox fptxtChoice 
      Height          =   288
      Left            =   6240
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   3000
      Width           =   492
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdDontSave 
      Height          =   564
      Left            =   972
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2628
      Width           =   2340
      _Version        =   131072
      _ExtentX        =   4128
      _ExtentY        =   995
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
      ButtonDesigner  =   "frmTaxMsgW4Opts.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSave 
      Height          =   564
      Left            =   972
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3468
      Width           =   2340
      _Version        =   131072
      _ExtentX        =   4128
      _ExtentY        =   995
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
      ButtonDesigner  =   "frmTaxMsgW4Opts.frx":0AAA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdReview 
      Height          =   564
      Left            =   3672
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2628
      Width           =   2340
      _Version        =   131072
      _ExtentX        =   4128
      _ExtentY        =   995
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
      ButtonDesigner  =   "frmTaxMsgW4Opts.frx":0C85
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdAbandonAll 
      Height          =   564
      Left            =   3672
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3468
      Width           =   2340
      _Version        =   131072
      _ExtentX        =   4128
      _ExtentY        =   995
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
      ButtonDesigner  =   "frmTaxMsgW4Opts.frx":0E61
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1650
      Left            =   405
      TabIndex        =   4
      Top             =   375
      UseMnemonic     =   0   'False
      Width           =   6150
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   2070
      Left            =   180
      Top             =   180
      Width           =   6630
   End
End
Attribute VB_Name = "frmTaxMsgW4Opts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAbandonAll_Click()
  Me.Hide
  fptxtChoice.Text = "abandon"
End Sub

Private Sub cmdDontSave_Click()
  Me.Hide
  fptxtChoice = "dontsave"
End Sub

Private Sub cmdSave_Click()
  Me.Hide
  fptxtChoice = "save"
End Sub

Private Sub cmdReview_Click()
  Me.Hide
  fptxtChoice.Text = "review"
End Sub

Private Sub Form_Load()
  fptxtChoice.Visible = False
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyF3:
      Call cmdDontSave_Click
      KeyCode = 0
    Case vbKeyF10:
      Call cmdSave_Click
      KeyCode = 0
    Case vbKeyF5:
      Call cmdReview_Click
      KeyCode = 0
    Case vbKeyEscape:
      Call cmdAbandonAll_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub
Private Sub Timer1_Timer()
  Static tog As Boolean
  tog = Not tog
  
  If tog Then
    Me.BackColor = &HFF&
  Else
    Me.BackColor = &HC0&
  End If
End Sub
