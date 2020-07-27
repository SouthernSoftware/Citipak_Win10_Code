VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmBLMainMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "v 2.05 Business License Main Menu"
   ClientHeight    =   8730
   ClientLeft      =   30
   ClientTop       =   600
   ClientWidth     =   11565
   Icon            =   "frmBLMainMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11565
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Interval        =   2000
      Left            =   360
      Top             =   1080
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   360
      Top             =   360
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdAppsAdvLtr 
      Height          =   444
      Left            =   3960
      TabIndex        =   4
      Tag             =   "Press to bring up a menu with links to all application and advance letter printing options as well as mailing label printing."
      Top             =   3420
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   783
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
      ButtonDesigner  =   "frmBLMainMenu.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdCustReports 
      Height          =   450
      Left            =   3960
      TabIndex        =   3
      Tag             =   "Press to bring up a menu with links to several business license reporting options."
      Top             =   2892
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   794
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
      ButtonDesigner  =   "frmBLMainMenu.frx":0AAA
   End
   Begin fpBtnAtlLibCtl.fpBln btnHelp 
      Height          =   300
      Left            =   3150
      TabIndex        =   1
      Top             =   7320
      Width           =   690
      _Version        =   131072
      _ExtentX        =   1217
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
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
   Begin fpBtnAtlLibCtl.fpBtn cmdCustMaint 
      Height          =   444
      Left            =   3960
      TabIndex        =   2
      Tag             =   "Press to bring up a menu with links to various customer maintenance options."
      Top             =   2376
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   783
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
      ButtonDesigner  =   "frmBLMainMenu.frx":0C96
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdIssueLics 
      Height          =   450
      Left            =   3960
      TabIndex        =   5
      Tag             =   "Press to bring up a menu of links used in setting and processing business licenses."
      Top             =   3936
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   794
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
      ButtonDesigner  =   "frmBLMainMenu.frx":0E7E
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPenalties 
      Height          =   435
      Left            =   3960
      TabIndex        =   6
      Tag             =   $"frmBLMainMenu.frx":1064
      Top             =   4470
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   767
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
      ButtonDesigner  =   "frmBLMainMenu.frx":10EF
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdEnterPayments 
      Height          =   450
      Left            =   3960
      TabIndex        =   7
      Tag             =   "Press to bring up a menu with links to all payment transaction features."
      Top             =   4980
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   794
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
      ButtonDesigner  =   "frmBLMainMenu.frx":12D5
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdTown 
      Height          =   444
      Left            =   3960
      TabIndex        =   8
      Tag             =   "Press to bring up the Town Setup screen with links to selecting application and delinquent notice forms."
      Top             =   5508
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   783
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
      ButtonDesigner  =   "frmBLMainMenu.frx":14B6
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdCategoryMaint 
      Height          =   444
      Left            =   3960
      TabIndex        =   9
      Tag             =   "Press to bring up a menu with links to category maintenance features including adding new categories and editing existing ones."
      Top             =   6036
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   783
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
      ButtonDesigner  =   "frmBLMainMenu.frx":16A2
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdHelp 
      Height          =   450
      Left            =   3960
      TabIndex        =   10
      Top             =   6555
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   794
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
      ButtonDesigner  =   "frmBLMainMenu.frx":188A
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   450
      Left            =   3960
      TabIndex        =   11
      Tag             =   "Press to exit business license and return to the Citipak main menu."
      Top             =   7080
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   794
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
      ButtonDesigner  =   "frmBLMainMenu.frx":1A6F
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSoSoft 
      Height          =   444
      Left            =   3960
      TabIndex        =   12
      Top             =   7608
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   783
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
      ButtonDesigner  =   "frmBLMainMenu.frx":1C58
   End
   Begin VB.Line Line12 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2086
      X2              =   2086
      Y1              =   2130
      Y2              =   8002
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   150
      Index           =   4
      Left            =   1970
      Top             =   2000
      Width           =   990
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   155
      Index           =   3
      Left            =   8550
      Top             =   1995
      Width           =   985
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   1095
      Index           =   1
      Left            =   1455
      Top             =   820
      Width           =   8655
   End
   Begin VB.Line Line14 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   8666
      X2              =   9366
      Y1              =   8020
      Y2              =   8020
   End
   Begin VB.Line Line13 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2086
      X2              =   2795
      Y1              =   8025
      Y2              =   8025
   End
   Begin VB.Line Line11 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      X1              =   8666
      X2              =   8666
      Y1              =   2136
      Y2              =   8008
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "BUSINESS LICENSE MAIN MENU"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2775
      TabIndex        =   0
      Top             =   1170
      Width           =   6015
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   1215
      Left            =   1455
      Top             =   690
      Width           =   8655
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   0
      Left            =   1966
      Top             =   1890
      Width           =   975
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H8000000B&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5895
      Index           =   0
      Left            =   2086
      Top             =   2130
      Width           =   735
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H8000000B&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5895
      Index           =   1
      Left            =   8655
      Top             =   2130
      Width           =   735
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   2
      Left            =   8550
      Top             =   1890
      Width           =   975
   End
   Begin VB.Menu mnuSoSoftOptions 
      Caption         =   "SoSoft Options"
      Begin VB.Menu mnuFixIndianTrails 
         Caption         =   "Fix Indian Trails"
      End
   End
End
Attribute VB_Name = "frmBLMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsBLTextBoxOverrider
  Private Temp_Class As Resize_Class
Private Sub cmdHelp_Click()
  If InStr(cmdHelp.Text, "On") Then
    cmdHelp.Text = "Turn Menu &Help Off"
    btnHelp.AutoScan = fpAutoScanPopupOnly
  ElseIf InStr(cmdHelp.Text, "Off") Then
    cmdHelp.Text = "Turn Menu &Help On"
    btnHelp.AutoScan = fpAutoScanOff
  End If
End Sub

Private Sub cmdAppsAdvLtr_Click()
  On Error Resume Next
  If Not Exist("arcust.dat") Then
    frmBLMessageBoxJr.Label1.Caption = "Error: No customers have been saved. Form load aborted."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  
  If Not Exist("artownsu.dat") Then
    frmBLMessageBoxJr.Label1.Caption = "Please complete the Town Setup data before continuing."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If

  If LevelPass = 1 Then
    frmBLIssueAppsLics.Show
    DoEvents
    Unload frmBLMainMenu
  Else
    If LevelPass = 2 Then
      MsgBox "Reports Only Password.", vbOKOnly, "Access Denied"
    ElseIf LevelPass = 3 Then
      MsgBox "Payments Only Password.", vbOKOnly, "Access Denied"
    Else
      MsgBox "Invalid Password.", vbOKOnly, "Access Denied"
    End If
  End If

End Sub

Private Sub cmdCategoryMaint_Click()
  If Not Exist("artownsu.dat") Then
    frmBLMessageBoxJrWOpts.Label1.Caption = "Please save data on the Town Setup screen before continuing. Press F10 to jump to the Town Setup screen otherwise press ESC to return to the screen."
    frmBLMessageBoxJrWOpts.Label1.Top = 700
    frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Jump"
    frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC Cancel"
    frmBLMessageBoxJrWOpts.Show vbModal
    If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "continue" Then
      Unload frmBLMessageBoxJrWOpts
      Close
      frmBLTownSetup.Show
      DoEvents
      Unload Me
      Exit Sub
    Else
      Unload frmBLMessageBoxJrWOpts
      Close
      Exit Sub
    End If
  End If

  If LevelPass = 1 Then
    frmBLCategoryMaintMenu.Show
    DoEvents
    Unload frmBLMainMenu
  Else
    If LevelPass = 2 Then
      MsgBox "Reports Only Password.", vbOKOnly, "Access Denied"
    ElseIf LevelPass = 3 Then
      MsgBox "Payments Only Password.", vbOKOnly, "Access Denied"
    Else
      MsgBox "Invalid Password.", vbOKOnly, "Access Denied"
    End If
  End If
End Sub

Private Sub cmdExit_Click()
  Close
  MainLog ("BusinessLicense.exe terminated via normal exit in Business License Main Menu.")
  
  Call Ready4others(PWcnt)
  If Exist(QPTrim$(StartPath) + "\" + "Citipak.exe") Then
    Shell QPTrim$(StartPath) + "\" + "Citipak.exe", vbMaximizedFocus
  End If
  DoEvents
  
  Timer1.Enabled = True
'  Call ClearInUse(PWcnt)
'  Call Terminate
'  DoEvents
'  End

End Sub

Private Sub cmdCustMaint_Click()
  If Not Exist("artownsu.dat") Then
    frmBLMessageBoxJr.Label1.Caption = "Please complete the Town Setup data before continuing."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If

  If LevelPass = 1 Then
    frmBLCustMaintMenu.Show
    DoEvents
    Unload frmBLMainMenu
  Else
    If LevelPass = 2 Then
      MsgBox "Reports Only Password.", vbOKOnly, "Access Denied"
    ElseIf LevelPass = 3 Then
      MsgBox "Payments Only Password.", vbOKOnly, "Access Denied"
    Else
      MsgBox "Invalid Password.", vbOKOnly, "Access Denied"
    End If
  End If
End Sub

Private Sub cmdCustReports_Click()
  
  If Not Exist("artownsu.dat") Then
    frmBLMessageBoxJr.Label1.Caption = "Please complete the Town Setup data before continuing."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  
'  If LevelPass = 3 Then
'    MsgBox "Payments Only Password.", vbOKOnly, "Access Denied"
'    Exit Sub
'  End If
  
  frmBLCustReportsMenu.Show
  DoEvents
  Unload frmBLMainMenu
End Sub

Private Sub cmdIssueLics_Click()
  On Error Resume Next
  
  If Not Exist("arcust.dat") Then
    frmBLMessageBoxJr.Label1.Caption = "Error: No customers have been saved. Form load aborted."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  
  If Not Exist("artownsu.dat") Then
    frmBLMessageBoxJr.Label1.Caption = "Please complete the Town Setup data before continuing."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  
  If LevelPass = 1 Then
    frmBLPrintLicMenu.Show
    DoEvents
    Unload frmBLMainMenu
  Else
    If LevelPass = 2 Then
      MsgBox "Reports Only Password.", vbOKOnly, "Access Denied"
    ElseIf LevelPass = 3 Then
      MsgBox "Payments Only Password.", vbOKOnly, "Access Denied"
    Else
      MsgBox "Invalid Password.", vbOKOnly, "Access Denied"
    End If
  End If
  
End Sub

Private Sub cmdEnterPayments_Click()
  If Not Exist("arcust.dat") Then
    frmBLMessageBoxJr.Label1.Caption = "Error: No customers have been saved. Form load aborted."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  
  If Not Exist("artownsu.dat") Then
    frmBLMessageBoxJr.Label1.Caption = "Please complete the Town Setup data before continuing."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If

  If LevelPass = 1 Or LevelPass = 3 Then
    frmPaymentDate.Show
    DoEvents
    Unload frmBLMainMenu
  Else
    MsgBox "Reports Only Password.", vbOKOnly, "Access Denied"
  End If
End Sub

Private Sub cmdPenalties_Click()
  On Error GoTo ERRORSTUFF
  
  If Not Exist("arcust.dat") Then
    frmBLMessageBoxJr.Label1.Caption = "Error: No customers have been saved. Form load aborted."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  
  If Not Exist("artownsu.dat") Then
    frmBLMessageBoxJr.Label1.Caption = "Please save Town Setup data before continuing."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  
  If LevelPass = 1 Then
    frmBLPenProcMenu.Show
    DoEvents
    Unload frmBLMainMenu
  Else
    If LevelPass = 2 Then
      MsgBox "Reports Only Password.", vbOKOnly, "Access Denied"
    ElseIf LevelPass = 3 Then
      MsgBox "Payments Only Password.", vbOKOnly, "Access Denied"
    Else
      MsgBox "Invalid Password.", vbOKOnly, "Access Denied"
    End If
  End If
  
  Exit Sub

ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLIssueAppsLics", "cmdPenalty_Click", Erl)
     Case emrExitProc:
       Resume Proc_Exit
     Case emrResume:
       Resume
     Case emrResumeNext:
       Resume Next
     Case Else
      '--- Technically, this should never happen.
       Resume Proc_Exit
   End Select
  
Proc_Exit:
  '--- Cleanup code goes here...
    Close
    ClearInUse PWcnt
    Terminate

End Sub

Private Sub cmdSoSoft_Click()
  frmSoSoftMenu.Show
  DoEvents
  Unload frmBLMainMenu
End Sub

Private Sub cmdTown_Click()
  If LevelPass = 1 Then
    frmBLTownSetup.Show
    DoEvents
    Unload frmBLMainMenu
  Else
    If LevelPass = 2 Then
      MsgBox "Reports Only Password.", vbOKOnly, "Access Denied"
    ElseIf LevelPass = 3 Then
      MsgBox "Payments Only Password.", vbOKOnly, "Access Denied"
    Else
      MsgBox "Invalid Password.", vbOKOnly, "Access Denied"
    End If
  End If
End Sub

Private Sub Form_Load()
  Dim FirstThru As Boolean
  Dim cnt&, dl&
  Dim ThisDir$
  
'  If App.PrevInstance Then
'    ActivatePrevInstance 'don't want two payroll
'    'programs open at once
'  End If
'
'  'the next series of code is used to get the
'  'identity of the current clerk using payroll
'  'and recorded anytime MainLog is accessed
'  cnt& = 199
'  ComputerName$ = String$(200, 0)
'  dl& = GetUserName(ComputerName$, cnt)
'  ComputerName$ = QPTrim$(ComputerName$)
  If FromBL = False Then
    FromBL = True
    cmdExit.Enabled = False
  End If
  
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsBLTextBoxOverrider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  
  'this saves the current path
  StartPath = App.Path
  If Right$(StartPath, 1) = "\" Then
    StartPath = Mid$(StartPath, 1, Len(StartPath) - 1)
  End If
  
  ThisDir = StartPath + "\BLRPTS"
  
  If Not DirExists(ThisDir) Then
    frmBLMessageBoxJrWOpts.Label1.Caption = "The directory 'BLRPTS' could not be located in the Citipak directory. Without the 'BLRPTS' directory graphics report printing is not possible. If you wish to create the 'BLRPTS' directory then press F10. Otherwise press ESC and call Southern Software @ 1-800-842-8190 for support."
    frmBLMessageBoxJrWOpts.Label1.Top = 500
    frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Make BLRPTS"
    frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC Escape"
    frmBLMessageBoxJrWOpts.Show vbModal
    If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "continue" Then
      Unload frmBLMessageBoxJrWOpts
      MkDir StartPath + "\BLRPTS"
    Else
      Unload frmBLMessageBoxJrWOpts
    End If
  End If
    
  DidPrint = 1
  
  KillFile "custlistopen.dat"
  KillFile "catlistopen.dat"
  KillFile "categoryedit.dat"
  KillFile "customeredit.dat"
  KillFile "adjustbalance.dat"
  KillFile "custbalList.dat"
  KillFile "custlistRpt.dat"
  KillFile "custlicList.dat"
  KillFile "custXlicList.dat"
  KillFile "custappList.dat"
  KillFile "custquickList.dat"
  KillFile "custappsRenews.dat"
  KillFile "custappIssue.dat"
  KillFile "transentry.dat"
  KillFile "pencalc.dat"
  KillFile "pencalcscr.dat"
  KillFile "dlnqnotice.dat"
  KillFile "dlnqmllbls.dat"
  KillFile "setstatus.dat"
  KillFile "mllbls.dat"
  KillFile "changeaccmeth.dat"
  KillFile "XlistInactiveY.dat"
  KillFile "inoutrpt.dat"
  KillFile "custinfomodal.dat"
  KillFile "transhistjr.dat"
  KillFile "custlookup.dat"
  KillFile "custByCat.dat"
  KillFile "custrptsmenu.dat"
  
  EditFlag = False 'used in entering/editing transactions
'  OPERNUM = 0
  GCatNum = 0
  GCustNum = 0
  GPayNum = 0
  cmdSoSoft.Visible = False
  ItemChangeFlag = False
  mnuSoSoftOptions.Visible = False
  
  If PWUser = "Sosoft Support" Then
    cmdSoSoft.Visible = True
    mnuSoSoftOptions.Visible = True
  End If
  If Exist("arcode.dat") Then
    If Not Exist("arcatcodeidx.dat") Then
      Call CreateCatCodeIdx
    End If
  End If
  If Exist("arcust.dat") Then
    If Not Exist("arcustnameidx.dat") Then
      Call CreateCustNameIdx
    End If
    If Not Exist("arcustnumidx.dat") Then
      Call CreateCustNumIdx
    End If
    If Not Exist("arlicnumidx.dat") Then
      Call CreateLicNumIdx
    End If
    If Not Exist("arsrhidx.dat") Then
      Call CreateCustSearchNameIdx
    End If
  End If
'  LevelPass = 1
'  PWcnt = 6
    
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    ''Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      If cmdExit.Enabled = True Then
        SendKeys "%s"
        Call cmdExit_Click
        KeyCode = 0
      End If
    Case Else:
  End Select
  PrintSign = False 'global used in printing laser licenses
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      ClearInUse PWcnt
      MainLog ("BusinessLicense.exe terminated via menu bar on frmBLMainMenu.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub mnuFixIndianTrails_Click()
  Dim TransRec As ARTransRecType
  Dim TransHandle As Integer
  Dim NumOfTransRecs As Double
  Dim TransCnt As Integer
  Dim TransRecd&
  Dim CustRec As ARCustRecType
  Dim CustHandle As Integer
  Dim ThisCustXNum As Integer
  Dim TDate As Integer
  'fix for #85 09/01/2009
  TDate = Date2Num("08/26/2009")
  OpenCustFile CustHandle
  Get CustHandle, 85, CustRec
  CustRec.AcctBal = 0
  Put CustHandle, 85, CustRec
  Close CustHandle
  TransRecd& = CustRec.FirstTrans
  OpenTransFile TransHandle
  NumOfTransRecs = LOF(TransHandle) / Len(TransRec)
  Do While TransRecd& > 0
    Get TransHandle, TransRecd&, TransRec
    If TransRec.TransDate > TDate Then
      TransRec.BalanceAfterTrans = 0
      TransRec.CashAmount = 0
      TransRec.CatCodeRec1 = 0
      TransRec.CatCodeRec2 = 0
      TransRec.CatCodeRec3 = 0
      TransRec.CatCodeRec4 = 0
      TransRec.CatCodeRec5 = 0
      TransRec.CatLicAmt1 = 0
      TransRec.CatLicAmt2 = 0
      TransRec.CatLicAmt3 = 0
      TransRec.CatLicAmt4 = 0
      TransRec.CatLicAmt5 = 0
      TransRec.CatLicBal1 = 0
      TransRec.CatLicBal2 = 0
      TransRec.CatLicBal3 = 0
      TransRec.CatLicBal4 = 0
      TransRec.CatLicBal5 = 0
      TransRec.ChkAmount = 0
      TransRec.FeeAmt = 0
      TransRec.IssAmt = 0
      TransRec.IssBal = 0
      TransRec.LicAmt = 0
      TransRec.LicBal = 0
      TransRec.PenAmt = 0
      TransRec.PenBal = 0
      TransRec.TransAmount = 0
      Put TransHandle, TransRecd&, TransRec
    End If
    TransRecd& = TransRec.NextTrans
  Loop
  Close
  MsgBox ("Done")

End Sub

Private Sub Timer1_Timer()
  Call Terminate2Shell
  DoEvents
  End
End Sub

Private Sub Timer2_Timer()
  cmdExit.Enabled = True
End Sub
