VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmTCShowPctComp 
   BackColor       =   &H00D0D0D0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2796
   ClientLeft      =   36
   ClientTop       =   420
   ClientWidth     =   5592
   Icon            =   "frmTaxShowPctComp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2796
   ScaleWidth      =   5592
   StartUpPosition =   2  'CenterScreen
   Begin fpBtnAtlLibCtl.fpBtn CmdCancel 
      Height          =   480
      Left            =   2208
      TabIndex        =   0
      Top             =   2040
      Width           =   1164
      _Version        =   131072
      _ExtentX        =   2053
      _ExtentY        =   847
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
      ButtonDesigner  =   "frmTaxShowPctComp.frx":08CA
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   396
      Left            =   859
      TabIndex        =   1
      Top             =   1350
      Width           =   3852
      _ExtentX        =   6795
      _ExtentY        =   699
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   " 00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2251
      TabIndex        =   5
      Top             =   966
      Width           =   732
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "% Complete."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3118
      TabIndex        =   4
      Top             =   966
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Processing:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   907
      TabIndex        =   3
      Top             =   966
      Width           =   1596
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   636
      Left            =   607
      TabIndex        =   2
      Top             =   294
      Width           =   4380
   End
End
Attribute VB_Name = "frmTCShowPctComp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim vWidth%, vHeight%, vTop%, vLeft%
Public Out As Boolean
Private Sub Form_Initialize()
  vLeft = (Screen.Width * 0.5)  ' Set width of form.
  vTop = (Screen.Height * 0.5) ' Set height of form.
  vWidth = 525 '(Screen.Width - vWidth) \ 2   ' Center form horizontally.
  vHeight = 280 '((Screen.Height - vHeight) \ 2) + 10  ' Center form vertically.
Out = False
End Sub

Private Sub cmdCancel_Click()
  If MsgBox("Are You Sure You Want To Cancel?", vbYesNo + vbSystemModal, "Cancel Processing") = vbYes Then
    frmTCShowPctComp.Out = True
  Else
    MakeWindowTopMost Me.hwnd, True
  End If
End Sub

Private Sub Form_Load()
Dim RetVal As Long, winhand As Long
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  MakeWindowTopMost Me.hwnd, True
  ProgressBar1.Value = 0
End Sub

Public Sub ShowPctComp(ByVal cnt As Long, ByVal TotalCnt As Long)
  Dim PctComp As Long
  PctComp = Int((cnt / TotalCnt) * 100)
  frmTCShowPctComp.Label3 = PctComp
  ProgressBar1.Value = PctComp
  If PctComp = 100 Then
    MakeWindowTopMost Me.hwnd, False
    Unload frmTCShowPctComp
    DoEvents
  Else
    DoEvents
  End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    Unload frmTCShowPctComp
    DoEvents
  End If
End Sub




