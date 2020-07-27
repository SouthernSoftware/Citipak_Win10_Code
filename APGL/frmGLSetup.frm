VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmGLSetupMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "General Ledger Setup Maintenance"
   ClientHeight    =   8865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12225
   Icon            =   "frmGLSetup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   12225
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin fpBtnAtlLibCtl.fpBtn cmdFundMaintMenu 
      Height          =   492
      Left            =   4308
      TabIndex        =   0
      Top             =   2568
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
      ButtonDesigner  =   "frmGLSetup.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdChartofAcctsMenu 
      Height          =   492
      Left            =   4308
      TabIndex        =   1
      Top             =   3180
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
      ButtonDesigner  =   "frmGLSetup.frx":0AB2
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdDeptMaintMenu 
      Height          =   492
      Left            =   4308
      TabIndex        =   2
      Top             =   3792
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
      ButtonDesigner  =   "frmGLSetup.frx":0C9B
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdBankMaintMenu 
      Height          =   492
      Left            =   4308
      TabIndex        =   3
      Top             =   4392
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
      ButtonDesigner  =   "frmGLSetup.frx":0E89
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdFunctionMenu 
      Height          =   480
      Left            =   4305
      TabIndex        =   4
      Top             =   5010
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
      ButtonDesigner  =   "frmGLSetup.frx":1071
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSetPostDates 
      Height          =   492
      Left            =   4308
      TabIndex        =   5
      Top             =   5616
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
      ButtonDesigner  =   "frmGLSetup.frx":125D
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdGLClosingOpMenu 
      Height          =   492
      Left            =   4320
      TabIndex        =   6
      Top             =   6228
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
      ButtonDesigner  =   "frmGLSetup.frx":1450
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdGLSysConfigUtilMenu 
      Height          =   492
      Left            =   4308
      TabIndex        =   7
      Top             =   6828
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
      ButtonDesigner  =   "frmGLSetup.frx":163E
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExitGLSetupMenu 
      Height          =   492
      Left            =   4308
      TabIndex        =   8
      Top             =   7440
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
      ButtonDesigner  =   "frmGLSetup.frx":182B
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "GENERAL LEDGER SETUP AND MAINTENANCE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   1
      Left            =   2280
      TabIndex        =   9
      Top             =   1440
      Width           =   7692
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      Height          =   1092
      Left            =   1800
      Top             =   1080
      Width           =   8652
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00D0D0D0&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   1212
      Left            =   1800
      Top             =   960
      Width           =   8652
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   9840
      X2              =   9840
      Y1              =   2280
      Y2              =   2400
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   8880
      X2              =   8880
      Y1              =   2280
      Y2              =   2400
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      Index           =   3
      X1              =   8880
      X2              =   9840
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      Index           =   2
      X1              =   8880
      X2              =   9840
      Y1              =   2400
      Y2              =   2400
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
   Begin VB.Line Line6 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   3360
      X2              =   3360
      Y1              =   2280
      Y2              =   2400
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   2400
      X2              =   2400
      Y1              =   2280
      Y2              =   2400
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   2400
      X2              =   3360
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   2424
      X2              =   3384
      Y1              =   2400
      Y2              =   2400
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
Attribute VB_Name = "frmGLSetupMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class


Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Me.HelpContextID = hlpGLSetupAnd
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
    Case vbKeyEscape:
      cmdExitGLSetupMenu_Click
      KeyCode = 0
      DoEvents
    Case Else:
  End Select
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExitGLSetupMenu.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        ClearInUse PWcnt
      End If
    End If
  End If
End Sub

Private Sub cmdSetPostDates_Click()
  If Exist("GLSETUP.DAT") Then
    frmSetPostDates.Show
    Unload frmGLSetupMenu
  Else
    MsgBox "The Main Setup Information Should Be Completed First.", vbOKOnly, "Incomplete Setup Info."
  End If
End Sub

Private Sub cmdBankMaintMenu_Click()
  If Exist("GLACCT.DAT") Then
    frmBankMaintMenu.Show
    Unload frmGLSetupMenu
  Else
    MsgBox "The GL Accounts Should Be Entered First.", vbOKOnly, "Incomplete Setup Info."
  End If
End Sub
Private Sub cmdFunctionMenu_Click()
  If Exist("GLSETUP.DAT") Then
    frmFunctionMenu.Show
    Unload frmGLSetupMenu
  Else
    MsgBox "The Main Setup Information Should Be Completed First.", vbOKOnly, "Incomplete Setup Info."
  End If
End Sub

Private Sub cmdChartofAcctsMenu_Click()
  If Exist("GLSETUP.DAT") And Exist("GLFund.dat") Then
    frmChartAcctMaintMenu.Show
    Unload frmGLSetupMenu
  Else
    MsgBox "Setup Information And Funds Should Be Entered Prior To Account Entry.", vbOKOnly, "Incomplete Setup Info."
  End If
End Sub

Private Sub cmdDeptMaintMenu_Click()
  If Exist("GLACCT.DAT") Then
    frmDeptMaintMenu.Show
    Unload frmGLSetupMenu
  Else
    MsgBox "You Must First Enter The GL Accounts.", vbOKOnly, "Incomplete Setup Info."
  End If
End Sub

Private Sub cmdExitGLSetupMenu_Click()
  frmGLMainMenu.Show
  Unload frmGLSetupMenu
End Sub

Private Sub cmdFundMaintMenu_Click()
  If Exist("GLSETUP.DAT") Then
    frmFundMaintMenu.Show
    Unload frmGLSetupMenu
  Else
    MsgBox "The Main Setup Information Should Be Completed First.", vbOKOnly, "Incomplete Setup Info."
  End If
End Sub

Private Sub cmdGLClosingOpMenu_Click()
Dim FileHandle As Integer, WhosOnFirst As String
  If Exist("GLSETUP.DAT") And Exist("GLACCT.DAT") Then
    If CloseAccess = True Then
      If Exist("FClose.opn") Then
          FileHandle = FreeFile
          Open "FClose.opn" For Input As FileHandle
          Line Input #FileHandle, WhosOnFirst$
          Close FileHandle
          MsgBox "The Close Out Menu Has Been Opened By: " + WhosOnFirst$, vbOKOnly, "Menu Not Accessible"
          Call MainLog("Close Year, Access Denied.")
        Else
          FileHandle = FreeFile
          Open "FClose.opn" For Output As FileHandle
          Print #FileHandle, ComputerName$
          Close FileHandle
          Call MainLog("Opened Close Year Menu.")
          frmGLClosingOpMenu.Show
          Unload frmGLSetupMenu
        End If
    Else
      MsgBox "Your Password does not allow access to closing operations.", vbOKOnly, "Access Denied"
    End If
'    frmPassWord.Caption = "GL Closing"
'    frmPassWord.Callingfrm = 2
'    frmPassWord.Show 1
  Else
    MsgBox "The Setup Information Should Be Completed First.", vbOKOnly, "Incomplete Setup Info."
  End If
End Sub

Private Sub cmdGLSysConfigUtilMenu_Click()
  frmGLConfigUtilMenu.Show
  Unload frmGLSetupMenu
End Sub


