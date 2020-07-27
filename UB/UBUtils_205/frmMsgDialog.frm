VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmMsgDialog 
   BackColor       =   &H000000C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4128
   ClientLeft      =   6420
   ClientTop       =   1200
   ClientWidth     =   6588
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   10.2
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4128
   ScaleMode       =   0  'User
   ScaleWidth      =   5588
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   350
      Left            =   6096
      Top             =   3672
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdOK 
      Height          =   468
      Left            =   1608
      TabIndex        =   0
      Top             =   3240
      Width           =   1284
      _Version        =   131072
      _ExtentX        =   2265
      _ExtentY        =   825
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
      ButtonDesigner  =   "frmMsgDialog.frx":0000
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdCancel 
      Height          =   468
      Left            =   3648
      TabIndex        =   1
      Top             =   3240
      Width           =   1284
      _Version        =   131072
      _ExtentX        =   2265
      _ExtentY        =   825
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
      ButtonDesigner  =   "frmMsgDialog.frx":01D5
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdOKOnly 
      Height          =   468
      Left            =   2616
      TabIndex        =   3
      Top             =   3240
      Width           =   1284
      _Version        =   131072
      _ExtentX        =   2265
      _ExtentY        =   825
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
      ButtonDesigner  =   "frmMsgDialog.frx":03AE
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   396
      Index           =   0
      Left            =   480
      TabIndex        =   8
      Top             =   336
      Width           =   5652
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   396
      Index           =   1
      Left            =   480
      TabIndex        =   7
      Top             =   888
      Width           =   5652
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   396
      Index           =   2
      Left            =   480
      TabIndex        =   6
      Top             =   1440
      Width           =   5652
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   396
      Index           =   3
      Left            =   480
      TabIndex        =   5
      Top             =   2016
      Width           =   5652
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   396
      Index           =   4
      Left            =   480
      TabIndex        =   4
      Top             =   2568
      Width           =   5652
   End
   Begin VB.Label RetLabel 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   168
      TabIndex        =   2
      Top             =   3408
      Visible         =   0   'False
      Width           =   1140
   End
End
Attribute VB_Name = "frmMsgDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
  frmMsgDialog.RetLabel = "0"
  Call WeAreOutOfHere
End Sub

Private Sub cmdOk_Click()
  frmMsgDialog.RetLabel = "-1"
  Call WeAreOutOfHere
End Sub

Private Sub WeAreOutOfHere()
  DoEvents
  frmMsgDialog.Hide
End Sub

Private Sub cmdOkOnly_Click()
  Call WeAreOutOfHere
End Sub

Private Sub Form_Activate()
'  DoEvents
  If frmMsgDialog.RetLabel = "-2" Then
    Me.cmdOk.Visible = False
    Me.cmdCancel.Visible = False
  Else
    Me.cmdOKOnly.Visible = False
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'  Stop
  Select Case KeyCode
  Case vbKeyEscape
    KeyCode = 0
    Call cmdCancel_Click
  Case Else
    KeyCode = 0
  End Select
End Sub

Private Sub Timer1_Timer()
  Dim BkColor As Integer
  Static tog As Boolean
  tog = Not tog
  If tog Then
    Me.BackColor = 230
  Else
    Me.BackColor = 192
  End If
End Sub

