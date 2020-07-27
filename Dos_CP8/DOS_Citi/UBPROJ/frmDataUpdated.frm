VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmDataUpdated 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " "
   ClientHeight    =   2580
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   4704
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawWidth       =   4
   Icon            =   "frmDataUpdated.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   4704
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   75
      Left            =   72
      Top             =   48
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdOK 
      Height          =   492
      Left            =   1704
      TabIndex        =   0
      Top             =   1584
      Width           =   1284
      _Version        =   131072
      _ExtentX        =   2265
      _ExtentY        =   868
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
      DrawFocusRect   =   1
      DrawFocusRectCell=   -1
      GrayAreaPictureStyle=   0
      Static          =   0   'False
      BackStyle       =   1
      AutoSize        =   0
      AutoSizeOffsetTop=   0
      AutoSizeOffsetBottom=   0
      AutoSizeOffsetLeft=   0
      AutoSizeOffsetRight=   0
      DropShadowOffsetX=   1
      DropShadowOffsetY=   1
      DropShadowType  =   1
      DropShadowColor =   0
      Redraw          =   -1  'True
      ButtonDesigner  =   "frmDataUpdated.frx":000C
   End
   Begin fpBtnAtlLibCtl.fpBtn fpBtn1 
      Height          =   636
      Left            =   576
      TabIndex        =   1
      Top             =   504
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   1122
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
      Static          =   -1  'True
      BackStyle       =   1
      AutoSize        =   0
      AutoSizeOffsetTop=   0
      AutoSizeOffsetBottom=   0
      AutoSizeOffsetLeft=   0
      AutoSizeOffsetRight=   0
      DropShadowOffsetX=   1
      DropShadowOffsetY=   1
      DropShadowType  =   1
      DropShadowColor =   0
      Redraw          =   -1  'True
      ButtonDesigner  =   "frmDataUpdated.frx":01E5
   End
End
Attribute VB_Name = "frmDataUpdated"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOK_Click()
  Unload frmDataUpdated
End Sub

Private Sub Timer1_Timer()
  Static tog As Byte
  tog = tog + 1
  Select Case tog
  Case 1
    cmdOK.ForeColor = &H80000012
  Case 2
    cmdOK.ForeColor = &H80000011
  Case 3
    cmdOK.ForeColor = &H80000010
  Case 4
    cmdOK.ForeColor = &H8000000F
  Case 5
    cmdOK.ForeColor = &H8000000E
  End Select
  If tog >= 5 Then
    tog = 0
  End If
  DoEvents
End Sub

'Private Sub Timer1_Timer()
'  Static tog As Boolean
'  tog = Not tog
'  If tog Then
'    cmdOK.ForeColor = &H8000000E
'  Else
'    cmdOK.ForeColor = &H80000012
'  End If
'End Sub
