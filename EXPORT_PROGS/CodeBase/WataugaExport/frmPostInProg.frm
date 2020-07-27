VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "BTN32A20.OCX"
Begin VB.Form frmPostInProg 
   BackColor       =   &H000040C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2028
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4344
   LinkTopic       =   "Form1"
   ScaleHeight     =   2028
   ScaleWidth      =   4344
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin fpBtnAtlLibCtl.fpBtn fpBtn1 
      Height          =   828
      Left            =   624
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   576
      Width           =   3036
      _Version        =   131072
      _ExtentX        =   5355
      _ExtentY        =   1460
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
      DropShadowType  =   2
      DropShadowColor =   0
      Redraw          =   -1  'True
      ButtonDesigner  =   "frmPostInProg.frx":0000
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      Height          =   1500
      Left            =   246
      Top             =   264
      Width           =   3852
   End
End
Attribute VB_Name = "frmPostInProg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    MainLog ("Payroll.exe terminated via menu bar on frmPostinProg.")
    Call Terminate
    End
  End If
End Sub

