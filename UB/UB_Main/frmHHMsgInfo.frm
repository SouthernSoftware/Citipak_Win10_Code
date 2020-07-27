VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "BTN32A20.OCX"
Begin VB.Form frmHHMsgInfo 
   BackColor       =   &H000000C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3228
   ClientLeft      =   6792
   ClientTop       =   2628
   ClientWidth     =   4920
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3228
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   Begin fpBtnAtlLibCtl.fpBtn cmdOK 
      Height          =   396
      Left            =   1920
      TabIndex        =   0
      Top             =   2472
      Width           =   1020
      _Version        =   131072
      _ExtentX        =   1799
      _ExtentY        =   698
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
      ButtonDesigner  =   "frmHHMsgInfo.frx":0000
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   384
      Index           =   4
      Left            =   96
      TabIndex        =   5
      Top             =   1728
      Width           =   4716
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   384
      Index           =   3
      Left            =   96
      TabIndex        =   4
      Top             =   1248
      Width           =   4716
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   384
      Index           =   2
      Left            =   96
      TabIndex        =   3
      Top             =   768
      Width           =   4716
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   384
      Index           =   1
      Left            =   96
      TabIndex        =   2
      Top             =   288
      Width           =   4716
   End
   Begin VB.Label RetLabel 
      BackStyle       =   0  'Transparent
      Height          =   228
      Left            =   3936
      TabIndex        =   1
      Top             =   2592
      Visible         =   0   'False
      Width           =   756
   End
End
Attribute VB_Name = "frmHHMsgInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
  frmHHMsgInfo.RetLabel = "0"
  Call WeAreOutOfHere
End Sub

Private Sub cmdOk_Click()
  frmHHMsgInfo.RetLabel = "-1"
  Call WeAreOutOfHere
End Sub

Private Sub WeAreOutOfHere()
'Can't unload until calling procedure has a chance to get ret value.
  DoEvents
  frmHHMsgInfo.Hide
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
'  Case vbKeyEscape
'    KeyCode = 0
'    Call cmdCancel_Click
  Case vbKeyReturn
    KeyCode = 0
    Call cmdOk_Click
  Case Else
    KeyCode = 0
  End Select
End Sub

