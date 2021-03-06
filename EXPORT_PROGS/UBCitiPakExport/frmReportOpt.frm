VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmReportOpt 
   BackColor       =   &H008A775B&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report Option"
   ClientHeight    =   2550
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   6540
   Icon            =   "frmReportOpt.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   6540
   StartUpPosition =   2  'CenterScreen
   Begin fpBtnAtlLibCtl.fpBtn cmdText 
      Height          =   510
      Left            =   3540
      TabIndex        =   2
      Top             =   1080
      Width           =   1260
      _Version        =   131072
      _ExtentX        =   2222
      _ExtentY        =   900
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
      ButtonDesigner  =   "frmReportOpt.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdGraphic 
      Height          =   510
      Left            =   1740
      TabIndex        =   1
      Top             =   1080
      Width           =   1260
      _Version        =   131072
      _ExtentX        =   2222
      _ExtentY        =   900
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
      ButtonDesigner  =   "frmReportOpt.frx":0ADA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   510
      Left            =   2640
      TabIndex        =   3
      Top             =   1800
      Width           =   1260
      _Version        =   131072
      _ExtentX        =   2222
      _ExtentY        =   900
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
      ButtonDesigner  =   "frmReportOpt.frx":0CEE
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Select Report Format  - Graphics or Text  "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   345
      Left            =   405
      TabIndex        =   0
      Top             =   480
      Width           =   5745
   End
End
Attribute VB_Name = "frmReportOpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim HasFocus As Integer

Private Sub cmdExit_Click()
  HasFocus = 3
  Unload frmReportOpt
End Sub

Private Sub cmdGraphic_GotFocus()
  HasFocus = 1
End Sub

Private Sub cmdText_Click()
  Unload frmReportOpt
  RptOpt = 2
End Sub

Private Sub cmdText_GotFocus()
  HasFocus = 2
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    If HasFocus = 1 Then
      Call cmdGraphic_Click
    ElseIf HasFocus = 2 Then
      Call cmdText_Click
    ElseIf HasFocus = 3 Then
      Call cmdExit_Click
    End If
  End If
  
  Select Case KeyCode
    Case vbKeyDown:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyUp:
      SendKeys "+{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "x"
      cmdExit_Click
      KeyCode = 0
    Case Else:
  End Select

End Sub

Private Sub cmdGraphic_Click()
  Unload frmReportOpt
  RptOpt = 1
End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  RptOpt = 0
  HasFocus = 0
'  Set Temp_Class = New Resize_Class
'  Temp_Class.InitResizeClass Me
End Sub




