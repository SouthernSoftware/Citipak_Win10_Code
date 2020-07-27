VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "BTN32A20.OCX"
Begin VB.Form frmGLMainMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "V2.05 General Ledger Main Menu"
   ClientHeight    =   8865
   ClientLeft      =   3930
   ClientTop       =   1890
   ClientWidth     =   12210
   Icon            =   "frmGLMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   12210
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   624
      Top             =   3144
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   624
      Top             =   2712
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdGenJournalMenu 
      Height          =   492
      Left            =   4296
      TabIndex        =   0
      Top             =   2352
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
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
      ButtonDesigner  =   "frmGLMain.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdCashReceiptsMenu 
      Height          =   492
      Left            =   4296
      TabIndex        =   1
      Top             =   2944
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
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
      ButtonDesigner  =   "frmGLMain.frx":0AB1
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdCashDisbMenu 
      Height          =   492
      Left            =   4296
      TabIndex        =   2
      Top             =   3536
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
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
      ButtonDesigner  =   "frmGLMain.frx":0C96
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdGLReportsMenu 
      Height          =   492
      Left            =   4296
      TabIndex        =   3
      Top             =   4128
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
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
      ButtonDesigner  =   "frmGLMain.frx":0E80
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdBudgetMaintMenu 
      Height          =   480
      Left            =   4290
      TabIndex        =   4
      Top             =   4725
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
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
      ButtonDesigner  =   "frmGLMain.frx":1063
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdBankReconMenu 
      Height          =   492
      Left            =   4296
      TabIndex        =   5
      Top             =   5312
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
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
      ButtonDesigner  =   "frmGLMain.frx":124D
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdGetDistMenu 
      Height          =   480
      Left            =   4290
      TabIndex        =   6
      Top             =   5910
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
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
      ButtonDesigner  =   "frmGLMain.frx":1438
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdGLSetupMenu 
      Height          =   492
      Left            =   4296
      TabIndex        =   7
      Top             =   6496
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
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
      ButtonDesigner  =   "frmGLMain.frx":1631
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPriorYrRpts 
      Height          =   480
      Left            =   4290
      TabIndex        =   8
      Top             =   7095
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
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
      ButtonDesigner  =   "frmGLMain.frx":1823
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExitGLMainMenu 
      Height          =   492
      Left            =   4296
      TabIndex        =   9
      Top             =   7680
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   868
      Enabled         =   0   'False
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
      ButtonDesigner  =   "frmGLMain.frx":1A12
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00FFFFFF&
      Height          =   132
      Left            =   8880
      Top             =   2280
      Width           =   972
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00FFFFFF&
      Height          =   132
      Left            =   2400
      Top             =   2280
      Width           =   972
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "GENERAL LEDGER MAIN MENU"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   372
      Left            =   3360
      TabIndex        =   10
      Top             =   1512
      Width           =   5292
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      Height          =   1092
      Left            =   1800
      Top             =   1080
      Width           =   8652
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00D0D0D0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   1212
      Left            =   1800
      Top             =   960
      Width           =   8652
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00D0D0D0&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   1
      Left            =   8880
      Top             =   2160
      Width           =   972
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   1
      X1              =   9000
      X2              =   9720
      Y1              =   8280
      Y2              =   8280
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   1
      X1              =   9000
      X2              =   9000
      Y1              =   2400
      Y2              =   8280
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D0D0D0&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   1
      Left            =   9000
      Top             =   2400
      Width           =   732
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00D0D0D0&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   0
      Left            =   2400
      Top             =   2160
      Width           =   972
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   2520
      X2              =   3240
      Y1              =   8280
      Y2              =   8280
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   2520
      X2              =   2520
      Y1              =   2400
      Y2              =   8280
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D0D0D0&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   0
      Left            =   2520
      Top             =   2400
      Width           =   732
   End
End
Attribute VB_Name = "frmGLMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
'Level of Password assigned during password sign-in
Private Sub cmdPriorYrRpts_Click()
  frmSelectPriorYr.Show
  Unload frmGLMainMenu
End Sub

Private Sub Form_Load()
  If App.PrevInstance Then
     ActivatePrevInstance
  End If
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Twiddle = "||//--\\"
  Me.HelpContextID = hlpGeneralLedger
  If DelayExit = True Then
    DelayExit = False
    Timer2.Enabled = True
  Else
    cmdExitGLMainMenu.Enabled = True
  End If
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub
Private Sub cmdBankReconMenu_Click()
  If Exist("GLACCT.DAT") Then
    If LevelPass = 1 Then
      frmBankReconMenu.Show
      Unload frmGLMainMenu
    Else
      MsgBox "Reports Only Password.", vbOKOnly, "Access Denied"
    End If
  Else
    MsgBox "You Should First Set Up GL Accounts.", vbOKOnly, "Missing Information"
  End If
End Sub
Private Sub cmdBudgetMaintMenu_Click()
  If Exist("GLACCT.DAT") Then
    If LevelPass = 1 Then
      frmBudgetMaintMenu.Show
      Unload frmGLMainMenu
    Else
      MsgBox "Reports Only Password.", vbOKOnly, "Access Denied"
    End If
  Else
    MsgBox "You Should First Set Up GL Accounts.", vbOKOnly, "Missing Information"
  End If
End Sub
Private Sub cmdCashDisbMenu_Click()
  If Exist("GLACCT.DAT") Then
    If LevelPass = 1 Then
      frmCashDisbMenu.Show
      Unload frmGLMainMenu
     Else
      MsgBox "Reports Only Password.", vbOKOnly, "Access Denied"
    End If
  Else
    MsgBox "You Should First Set Up GL Accounts.", vbOKOnly, "Missing Information"
  End If
End Sub
Private Sub cmdCashReceiptsMenu_Click()
  If Exist("GLACCT.DAT") Then
    If LevelPass = 1 Then
      frmCashReceiptsMenu.Show
      Unload frmGLMainMenu
    Else
      MsgBox "Reports Only Password.", vbOKOnly, "Access Denied"
    End If
  Else
    MsgBox "You Should First Set Up GL Accounts.", vbOKOnly, "Missing Information"
  End If
End Sub
Private Sub cmdExitGLMainMenu_Click()
  Call MainLog("EXIT GL Main Menu.")
  Ready4others PWcnt
  Shell "citipak.exe", vbMaximizedFocus
  frmGLMainMenu.Enabled = False
  Timer1.Enabled = True
End Sub
Private Sub cmdGenJournalMenu_Click()
  If Exist("GLACCT.DAT") Then
    If LevelPass = 1 Then
      frmGenJournalMenu.Show
      Unload frmGLMainMenu
    Else
      MsgBox "Reports Only Password.", vbOKOnly, "Access Denied"
    End If
  Else
    MsgBox "You Should First Set Up GL Accounts.", vbOKOnly, "Missing Information"
  End If
End Sub
Private Sub cmdGetDistMenu_Click()
  If Exist("GLACCT.DAT") Then
    If LevelPass = 1 Then
      frmGetDistMenu.Show
      Unload frmGLMainMenu
    Else
      MsgBox "Reports Only Password.", vbOKOnly, "Access Denied"
    End If
  Else
    MsgBox "You Should First Set Up GL Accounts.", vbOKOnly, "Missing Information"
  End If
End Sub
Private Sub cmdGLReportsMenu_Click()
  If Exist("GLACCT.DAT") Then
    frmGLReportsMenu.Show
    Unload frmGLMainMenu
  Else
    MsgBox "You Should First Set Up GL Accounts.", vbOKOnly, "Missing Information"
  End If
End Sub
Private Sub cmdGLSetupMenu_Click()
  If LevelPass = 1 Then
    frmGLSetupMenu.Show
    Unload frmGLMainMenu
   Else
      MsgBox "Reports Only Password.", vbOKOnly, "Access Denied"
    End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      If cmdExitGLMainMenu.Enabled = True Then
        cmdExitGLMainMenu_Click
      End If
      KeyCode = 0
      DoEvents
    Case Else:
  End Select
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      ClearInUse PWcnt
      Call MainLog("Close via GL MainMenu.")
    End If
  End If
End Sub

Private Sub Timer1_Timer()
  Unload Me
End Sub

Private Sub Timer2_Timer()
  cmdExitGLMainMenu.Enabled = True
End Sub
