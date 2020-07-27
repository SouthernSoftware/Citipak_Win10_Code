VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "BTN32A20.OCX"
Begin VB.Form frmBLPrintAppsRenwlsMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Business License Print Applications/Renewals Menu"
   ClientHeight    =   8868
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   11652
   Icon            =   "frmBLPrintAppsRenwlsMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   11652
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Turn Menu Help On"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   4032
      TabIndex        =   4
      Top             =   3216
      Width           =   3612
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit to Issue Apps/Lic Menu"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   4005
      TabIndex        =   2
      ToolTipText     =   "Click this button to exit business license and return to the main Citipak menu."
      Top             =   5376
      Width           =   3612
   End
   Begin VB.CommandButton cmdReprints 
      BackColor       =   &H008F8265&
      Caption         =   "&Reprint Applications/Renewals"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   4032
      MaskColor       =   &H8000000F&
      TabIndex        =   1
      ToolTipText     =   "Click this button to make changes to existing category codes"
      Top             =   4656
      Width           =   3612
   End
   Begin VB.CommandButton cmdPrintAppsRens 
      BackColor       =   &H008F8265&
      Caption         =   "Print &Applications/Renewals"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   4005
      MaskColor       =   &H8000000F&
      TabIndex        =   0
      ToolTipText     =   "Click this button to add a brand new category code."
      Top             =   3936
      Width           =   3612
   End
   Begin fpBtnAtlLibCtl.fpBln btnHelp 
      Height          =   300
      Left            =   5568
      TabIndex        =   5
      Top             =   6048
      Width           =   684
      _Version        =   131072
      _ExtentX        =   1206
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   9405029
      ForeColor       =   8454143
      Text            =   ""
      Shape           =   0
      ShapeRoundWidth =   180
      ShapeRoundHeight=   180
      BorderWidth     =   -1
      BorderColor     =   -2147483630
      ThreeDWidth     =   -1
      ThreeDShadowColor=   -2147483632
      ThreeDHighlightColor=   16777215
      ThreeDText      =   0
      ThreeDTextHighlightColor=   16777215
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignPictureH   =   0
      AlignPictureV   =   0
      PictureStyle    =   0
      WordWrap        =   -1  'True
      ScaleMode       =   1
      ThreeDStyle     =   2
      Position        =   0
      PosBaseX        =   0
      PosBaseY        =   0
      PosOffsetX      =   -100
      PosOffsetY      =   300
      MaxWidth        =   5000
      CloudInset      =   100
      CloudMinWidth   =   600
      TailShape       =   2
      TailType        =   2
      TailBaseOffsetOutside=   300
      TailBaseOffsetInside=   100
      TailBaseAxisOutside=   0
      TailBaseAxisInside=   0
      TailBubbleCount =   3
      AlignTextH      =   1
      AlignTextV      =   1
      DropShadowOffsetX=   3
      DropShadowOffsetY=   3
      DropShadowType  =   0
      DropShadowColor =   0
      AutoScan        =   0
      ScanAllDescendants=   -1  'True
      Interval        =   500
      IntervalNext    =   200
      AutoSize        =   -1  'True
      UseTagProp      =   -1  'True
      HideOnInactiveApp=   0   'False
      HideOnMouseDown =   2
      HideOnKeyDown   =   2
      HideOnFocus     =   0   'False
      ScanDisabledControls=   -1  'True
      ThreeDAppearance=   0
      FollowFocus     =   0   'False
      TemplateName    =   ""
   End
   Begin VB.Label lblBalloon 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "HELP BALLOONS ON"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   252
      Left            =   4800
      TabIndex        =   6
      Top             =   2928
      Width           =   2100
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   1092
      Index           =   1
      Left            =   1500
      Top             =   886
      Width           =   8652
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   8590
      X2              =   9540
      Y1              =   2083
      Y2              =   2083
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   2100
      X2              =   3060
      Y1              =   2083
      Y2              =   2083
   End
   Begin VB.Line Line14 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   8703
      X2              =   9403
      Y1              =   8079
      Y2              =   8079
   End
   Begin VB.Line Line13 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2219
      X2              =   2929
      Y1              =   8090
      Y2              =   8090
   End
   Begin VB.Line Line12 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2220
      X2              =   2220
      Y1              =   2203
      Y2              =   8075
   End
   Begin VB.Line Line7 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2100
      X2              =   3060
      Y1              =   2203
      Y2              =   2203
   End
   Begin VB.Line Line10 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      X1              =   8592
      X2              =   9542
      Y1              =   2203
      Y2              =   2203
   End
   Begin VB.Line Line11 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      X1              =   8700
      X2              =   8700
      Y1              =   2206
      Y2              =   8078
   End
   Begin VB.Line Line9 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      X1              =   9550
      X2              =   9550
      Y1              =   2197
      Y2              =   2082
   End
   Begin VB.Line Line8 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      X1              =   8590
      X2              =   8590
      Y1              =   2100
      Y2              =   2205
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      X1              =   2100
      X2              =   2100
      Y1              =   2083
      Y2              =   2203
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      X1              =   3060
      X2              =   3060
      Y1              =   2203
      Y2              =   2083
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000014&
      X1              =   8580
      X2              =   9540
      Y1              =   1963
      Y2              =   1963
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000014&
      X1              =   2100
      X2              =   3060
      Y1              =   1963
      Y2              =   1963
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PRINT APPLICATIONS/RENEWALS MENU"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2820
      TabIndex        =   3
      Top             =   1246
      Width           =   6012
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   1212
      Left            =   1500
      Top             =   766
      Width           =   8652
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   0
      Left            =   2100
      Top             =   1966
      Width           =   972
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H8000000B&
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   0
      Left            =   2220
      Top             =   2196
      Width           =   732
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   2
      Left            =   8592
      Top             =   1966
      Width           =   972
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H8000000B&
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   1
      Left            =   8700
      Top             =   2194
      Width           =   732
   End
End
Attribute VB_Name = "frmBLPrintAppsRenwlsMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsBLTextBoxOverrider
  Private Temp_Class As Resize_Class

Private Sub cmdExit_Click()
  frmBLIssueAppsLics.Show
  DoEvents
  Unload frmBLPrintAppsRenwlsMenu
End Sub

Private Sub cmdPrintAppsRens_Click()
  frmBLPrintAppsRenwls.Show
  DoEvents
  Unload frmBLPrintAppsRenwlsMenu
End Sub

Private Sub cmdReprints_Click()
  If Not Exist("artmpcus.dat") Then
    frmBLMessageBoxJr.Label1.Caption = "No renewal applications have been created in the text version so no reprints are possible."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Close
    Exit Sub
  End If
  
  frmBLReprintAppsRnwls.Show
  DoEvents
  Unload frmBLPrintAppsRenwlsMenu
End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsBLTextBoxOverrider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
    DoEvents
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%X"
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      ClearInUse PWcnt
      MainLog ("BusinessLicense.exe terminated via menu bar on frmBLPrintAppsRenwlsMenu.")
      Call Terminate
      End
    End If
  End If
End Sub

