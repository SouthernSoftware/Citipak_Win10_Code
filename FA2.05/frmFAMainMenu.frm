VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmFAMainMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "v 2.05 Fixed Assets Main Menu"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmFAMainMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleMode       =   0  'User
   ScaleWidth      =   11652
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin fpBtnAtlLibCtl.fpBtn cmdClearSoSoftFlags 
      Height          =   495
      Left            =   4005
      TabIndex        =   8
      ToolTipText     =   "Use this option to clear all depreciation reversal flags allowing you to  continue with a new reversal."
      Top             =   7800
      Visible         =   0   'False
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   873
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   13684944
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
      ButtonDesigner  =   "frmFAMainMenu.frx":08CA
   End
   Begin VB.Timer Timer2 
      Interval        =   2000
      Left            =   360
      Top             =   840
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   360
      Top             =   360
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdItemMaint 
      Height          =   495
      Left            =   4005
      TabIndex        =   1
      ToolTipText     =   "Click this button to bring up the screen choices needed for adding or editing fixed assets."
      Top             =   2832
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   873
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   13684944
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
      ButtonDesigner  =   "frmFAMainMenu.frx":0AB7
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdReportsMenu 
      Height          =   495
      Left            =   4005
      TabIndex        =   2
      ToolTipText     =   "Click this button to bring up a menu of all available reports for fixed assets."
      Top             =   3552
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   873
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   13684944
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
      ButtonDesigner  =   "frmFAMainMenu.frx":0C9B
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdYearEndProc 
      Height          =   495
      Left            =   4005
      TabIndex        =   3
      ToolTipText     =   "Click this button to bring up a menu of all year end depreciation processes."
      Top             =   4272
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   873
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   13684944
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
      ButtonDesigner  =   "frmFAMainMenu.frx":0E7B
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdDisposal 
      Height          =   495
      Left            =   4005
      TabIndex        =   4
      ToolTipText     =   "Click this button to bring up a menu of all fixed asset disposal routines."
      Top             =   4992
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   873
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   13684944
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
      ButtonDesigner  =   "frmFAMainMenu.frx":1062
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdFAMaintMenu 
      Height          =   495
      Left            =   4005
      TabIndex        =   5
      ToolTipText     =   $"frmFAMainMenu.frx":124E
      Top             =   5712
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   873
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   13684944
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
      ButtonDesigner  =   "frmFAMainMenu.frx":12F7
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   495
      Left            =   4005
      TabIndex        =   6
      ToolTipText     =   "Click this button to exit fixed assets and return to the main Citipak menu."
      Top             =   6432
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   873
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   13684944
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
      ButtonDesigner  =   "frmFAMainMenu.frx":14DE
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSOSoftOnly 
      Height          =   495
      Left            =   4005
      TabIndex        =   7
      ToolTipText     =   "Use this utility to clear the most recently posted depreciation."
      Top             =   7080
      Visible         =   0   'False
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   873
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   13684944
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
      ButtonDesigner  =   "frmFAMainMenu.frx":16C3
   End
   Begin VB.Line Line11 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      X1              =   8699.76
      X2              =   8699.76
      Y1              =   2149.036
      Y2              =   7870.051
   End
   Begin VB.Line Line12 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2220.428
      X2              =   2220.428
      Y1              =   2146.112
      Y2              =   7876.874
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   135
      Index           =   4
      Left            =   8610
      Top             =   2091
      Width           =   960
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   135
      Index           =   3
      Left            =   2110
      Top             =   2091
      Width           =   955
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "FIXED ASSETS MAIN MENU"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2820
      TabIndex        =   0
      Top             =   1246
      Width           =   6012
   End
   Begin VB.Line Line13 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2199.434
      X2              =   2929.246
      Y1              =   7881.747
      Y2              =   7881.747
   End
   Begin VB.Line Line14 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   8682.765
      X2              =   9402.579
      Y1              =   7871.026
      Y2              =   7871.026
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   1095
      Index           =   1
      Left            =   1500
      Top             =   900
      Width           =   8655
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   0
      Left            =   2100
      Top             =   1966
      Width           =   972
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H8000000B&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   0
      Left            =   2220
      Top             =   2196
      Width           =   732
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   1212
      Left            =   1500
      Top             =   766
      Width           =   8652
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   2
      Left            =   8592
      Top             =   1966
      Width           =   972
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H8000000B&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   1
      Left            =   8700
      Top             =   2194
      Width           =   732
   End
End
Attribute VB_Name = "frmFAMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsFATextBoxOverRider
  Private Temp_Class As Resize_Class

Private Sub cmdClearSoSoftFlags_Click()
  frmFAClearRevFlags.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdDisposal_Click()
  If LevelPass = 1 Then
    frmFADisposalMenu.Show
    DoEvents
    Unload frmFAMainMenu
  Else
    MsgBox "Reports Only Password.", vbOKOnly, "Access Denied"
  End If
End Sub

Private Sub cmdExit_Click()
  Close
  MainLog ("FixedAssets.exe terminated via normal exit in Fixed Assets Main Menu.")
  
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

Private Sub cmdFAMaintMenu_Click()
  If LevelPass = 1 Then
    frmFAMaintMenu.Show
    DoEvents
    Unload frmFAMainMenu
  Else
    MsgBox "Reports Only Password.", vbOKOnly, "Access Denied"
  End If
End Sub

Private Sub cmdItemMaint_Click()
  If LevelPass = 1 Then
    frmFAItemMaintMenu.Show
    DoEvents
    Unload frmFAMainMenu
  Else
    MsgBox "Reports Only Password.", vbOKOnly, "Access Denied"
  End If
  
End Sub

Private Sub cmdReportsMenu_Click()
  frmFAReportMenu.Show
  DoEvents
  Unload frmFAMainMenu
End Sub

Private Sub cmdSOSoftOnly_Click()
  frmFAInternalOnly.Show
  DoEvents
  Unload frmFAMainMenu
End Sub

Private Sub cmdYearEndProc_Click()
  If LevelPass = 1 Then
    frmFAYearEndMenu.Show
    DoEvents
    Unload frmFAMainMenu
  Else
    MsgBox "Reports Only Password.", vbOKOnly, "Access Denied"
  End If
End Sub

Private Sub Form_Load()
  Dim FirstThru As Boolean
  Dim cnt&, dl&
  Dim ThisDir$
  
  On Error Resume Next
  Clipboard.Clear
'  If App.PrevInstance Then
'    ActivatePrevInstance 'don't want two payroll
'    'programs open at once
'  End If
  
  'the next series of code is used to get the
  'identity of the current clerk using payroll
  'and recorded anytime MainLog is accessed
'  cnt& = 199
'  ComputerName$ = String$(200, 0)
'  dl& = GetUserName(ComputerName$, cnt)
'  ComputerName$ = QPTrim$(ComputerName$)
  If FromFA = False Then
    cmdExit.Enabled = False
    FromFA = True
  End If
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsFATextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  
  'this saves the current path
'  StartPath = App.Path
'  If Right$(StartPath, 1) = "\" Then
'    StartPath = Mid$(StartPath, 1, Len(StartPath) - 1)
'  End If
  ThisDir = StartPath + "\FARPTS"
  
  If Not DirExists(ThisDir) Then
    frmFAEditItemMess.Label1.Caption = "The directory 'FARPTS' could not be located in the Citipak directory. Without the 'FARPTS' directory graphics report printing is not possible. If you wish to create the 'FARPTS' directory then press F10. Otherwise press ESC and call Southern Software @ 1-800-842-8190 for support."
    frmFAEditItemMess.Label1.Top = 900
    frmFAEditItemMess.cmdCont.Text = "F10 Make FARPTS"
    frmFAEditItemMess.cmdExit.Text = "ESC Escape"
    frmFAEditItemMess.Show vbModal
    If frmFAEditItemMess.fptxtChoice.Text = "continue" Then
      Unload frmFAEditItemMess
      MkDir StartPath + "\FARPTS"
    Else
      Unload frmFAEditItemMess
    End If
  End If
  
  KillFile ("dprhistbyitemrpt.dat")
  KillFile ("dprhistrpt.dat")
  KillFile ("valrpt.dat")
  KillFile ("itemchecklist.dat")
  KillFile (TempDprFileName)
  KillFile ("masteritemlistopen.dat")
  KillFile ("editdeptopen.dat")
  KillFile ("edititemopen.dat")
  KillFile ("Wrntyrpt.dat")
  KillFile ("taglistopen.dat")
  KillFile ("assetbycoderpt.dat")
  GRecNum = 0
  GCodeNum = 0
  GDeptNum = 0
  ItemChangeFlag = False
  AddItemFlag = False
  
'  LevelPass = 1 'use only for working in the environment
'  PWcnt = -3 'use only for working in the environment
  
  If PWcnt = -3 Then
    cmdSOSoftOnly.Visible = True
    cmdClearSoSoftFlags.Visible = True
  End If
  'if a depreciation reversal is needed then the sosoft
  'password will bring up a command button that allows for that

End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    'Me.Visible = False
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
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      ClearInUse PWcnt
      MainLog ("FixedAssets.exe terminated via menu bar on frmFAMainMenu.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub Timer1_Timer()
'  Call ClearInUse(PWcnt)
  Call Terminate2Shell
  Close
  End
End Sub

Private Sub Timer2_Timer()
  cmdExit.Enabled = True
End Sub
