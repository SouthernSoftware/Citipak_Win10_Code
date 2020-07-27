VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBLShowPctComp 
   BackColor       =   &H00D0D0D0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2700
   ClientLeft      =   30
   ClientTop       =   30
   ClientWidth     =   5505
   ControlBox      =   0   'False
   Icon            =   "frmBLShowPctComp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   5505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin fpBtnAtlLibCtl.fpBtn CmdCancel 
      Height          =   480
      Left            =   2160
      TabIndex        =   5
      Top             =   2070
      Width           =   1170
      _Version        =   131072
      _ExtentX        =   2064
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
      ButtonDesigner  =   "frmBLShowPctComp.frx":000C
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   396
      Left            =   816
      TabIndex        =   0
      Top             =   1404
      Width           =   3852
      _ExtentX        =   6800
      _ExtentY        =   688
      _Version        =   393216
      Appearance      =   1
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
      Left            =   564
      TabIndex        =   4
      Top             =   348
      Width           =   4380
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
      Left            =   864
      TabIndex        =   3
      Top             =   1020
      Width           =   1596
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
      Left            =   3075
      TabIndex        =   2
      Top             =   1020
      Width           =   1815
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
      Left            =   2208
      TabIndex        =   1
      Top             =   1020
      Width           =   732
   End
End
Attribute VB_Name = "frmBLShowPctComp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsBLTextBoxOverrider
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
    frmBLShowPctComp.Out = True
  Else
    MakeWindowTopMost Me.hwnd, True
  End If
End Sub

Private Sub Form_Load()
Dim RetVal As Long, winhand As Long
  Set Over = New clsBLTextBoxOverrider
  Over.OverRide Me
  MakeWindowTopMost Me.hwnd, True
  ProgressBar1.Value = 0
End Sub

Public Sub ShowPctComp(ByVal cnt As Long, ByVal TotalCnt As Long)
  Dim PctComp As Long
  PctComp = Int((cnt / TotalCnt) * 100)
  frmBLShowPctComp.Label3 = PctComp
  ProgressBar1.Value = PctComp
  If PctComp = 100 Then
    MakeWindowTopMost Me.hwnd, False
    Unload frmBLShowPctComp
    DoEvents
  Else
    DoEvents
  End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    Unload frmBLShowPctComp
    DoEvents
  End If
End Sub



