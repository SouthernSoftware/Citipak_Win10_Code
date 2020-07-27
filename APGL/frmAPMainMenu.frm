VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmAPMainMenu 
   AutoRedraw      =   -1  'True
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "v2.05  Accounts Payable Main Menu"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   12225
   Icon            =   "frmAPMainMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   12225
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   600
      Top             =   2088
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   624
      Top             =   1608
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdVendMaintMenu 
      Height          =   492
      Left            =   4302
      TabIndex        =   0
      Top             =   2808
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
      ButtonDesigner  =   "frmAPMainMenu.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdInvProcMenu 
      Height          =   492
      Left            =   4302
      TabIndex        =   2
      Top             =   4382
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
      ButtonDesigner  =   "frmAPMainMenu.frx":0AB4
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdAPChkMenu 
      Height          =   480
      Left            =   4305
      TabIndex        =   3
      Top             =   5175
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
      ButtonDesigner  =   "frmAPMainMenu.frx":0C9E
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdAPReportMenu 
      Height          =   492
      Left            =   4302
      TabIndex        =   4
      Top             =   5956
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
      ButtonDesigner  =   "frmAPMainMenu.frx":0E86
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExitAPMainMenu 
      Height          =   480
      Left            =   4305
      TabIndex        =   5
      Top             =   6750
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   847
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
      ButtonDesigner  =   "frmAPMainMenu.frx":1065
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPOMenu 
      Height          =   480
      Left            =   4305
      TabIndex        =   1
      Top             =   3600
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
      ButtonDesigner  =   "frmAPMainMenu.frx":124A
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   9000
      X2              =   9696
      Y1              =   8280
      Y2              =   8280
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   0
      X1              =   2520
      X2              =   2520
      Y1              =   2424
      Y2              =   8280
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   2520
      X2              =   3216
      Y1              =   8280
      Y2              =   8280
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   3360
      X2              =   3360
      Y1              =   2280
      Y2              =   2400
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   2400
      X2              =   2400
      Y1              =   2280
      Y2              =   2400
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   1
      X1              =   2400
      X2              =   3360
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   2400
      X2              =   3360
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ACCOUNTS PAYABLE MAIN MENU"
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
      Left            =   3462
      TabIndex        =   6
      Top             =   1440
      Width           =   5292
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000009&
      Height          =   1092
      Left            =   1800
      Top             =   1080
      Width           =   8652
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000009&
      Index           =   2
      X1              =   9840
      X2              =   9840
      Y1              =   2280
      Y2              =   2400
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000009&
      Index           =   2
      X1              =   8880
      X2              =   8880
      Y1              =   2280
      Y2              =   2400
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   5
      X1              =   8880
      X2              =   9840
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   4
      X1              =   8880
      X2              =   9840
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   2
      X1              =   9000
      X2              =   9000
      Y1              =   2424
      Y2              =   8280
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D0D0D0&
      BorderColor     =   &H00000000&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   2
      Left            =   9000
      Top             =   2400
      Width           =   732
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00D0D0D0&
      BorderColor     =   &H00D0D0D0&
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
      BorderColor     =   &H00000000&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   2
      Left            =   8880
      Top             =   2160
      Width           =   972
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000000&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   0
      Left            =   2400
      Top             =   2160
      Width           =   972
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D0D0D0&
      BorderColor     =   &H00000000&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   0
      Left            =   2520
      Top             =   2400
      Width           =   732
   End
End
Attribute VB_Name = "frmAPMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class

Private Sub cmdAPChkMenu_Click()
  Dim FileHandle As Integer, WhosOnFirst As String
  If Exist("APVENDOR.DAT") Then
    If LevelPass = 1 Then
      If Exist("APChk.opn") Then
        FileHandle = FreeFile
        Open "APChk.opn" For Input As FileHandle
        Line Input #FileHandle, WhosOnFirst$
        Close FileHandle
        MsgBox "AP Check Processing Has Been Opened By: " + WhosOnFirst$, vbOKOnly, "Not Accessible"
      Else
        FileHandle = FreeFile
        Open "APChk.opn" For Output As FileHandle
        Print #FileHandle, ComputerName$
        Close FileHandle
        Load frmAPChkProcessMenu
        DoEvents
        Call MainLog("Open AP Chk Menu.")
        frmAPChkProcessMenu.Show
        Unload frmAPMainMenu
      End If
    Else
      MsgBox "Your Password Does Not Allow Access To Check Processing.", vbOKOnly, "Access Denied"
    End If
  Else
    MsgBox "Vendor Information Must Be Completed First.", vbOKOnly, "Missing Information"
  End If
End Sub

Private Sub cmdAPReportMenu_Click()
  If Exist("APVENDOR.DAT") Then
    If LevelPass > 0 Then
      frmAPReportsMenu.Show
      Unload frmAPMainMenu
    Else
      MsgBox "Password Does Not Allow Access to PO's.", vbOKOnly, "Access Denied"
    End If
  Else
    MsgBox "Vendor Information Must Be Completed First.", vbOKOnly, "Missing Information"
  End If
End Sub

Private Sub cmdExitAPMainMenu_Click()
  Call MainLog("EXIT AP Main Menu.")
  Ready4others PWcnt
  Shell "citipak.exe", vbMaximizedFocus
  Timer1.Enabled = True
End Sub

Private Sub cmdInvProcMenu_Click()
  If Exist("APVENDOR.DAT") Then
    If LevelPass = 1 Then
      frmInvProcessMenu.Show
      Unload frmAPMainMenu
    Else
      MsgBox "Your Password Does Not Allow Access To Invoice Processing.", vbOKOnly, "Access Denied"
    End If
  Else
    MsgBox "Vendor Information Must Be Completed First.", vbOKOnly, "Missing Information"
  End If
End Sub

Private Sub cmdPOMenu_Click()
  If Exist("APVENDOR.DAT") Then
    If LevelPass = 1 Or OKtoPO Then
      frmPOProcessMenu.Show
      Unload frmAPMainMenu
    Else
      MsgBox "Password Does Not Allow Access to PO's.", vbOKOnly, "Access Denied"
    End If
  Else
    MsgBox "Vendor Information Must Be Completed First.", vbOKOnly, "Missing Information"
  End If
End Sub

Private Sub cmdVendMaintMenu_Click()
  If LevelPass = 1 Then
    frmAPVendMaintMenu.Show
    Unload frmAPMainMenu
    Else
      MsgBox "Your Password Does Not Allow Access to Vendor Maintenance.", vbOKOnly, "Access Denied"
    End If
End Sub

Private Sub Form_Load()
  Dim cnt&, dl&
  If App.PrevInstance Then
     ActivatePrevInstance
  End If
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  screenW = (Screen.Width / Screen.TwipsPerPixelX)
  cnt& = 199
  ComputerName$ = String$(200, 0)
  dl& = GetUserName(ComputerName$, cnt)
  ComputerName$ = QPTrim$(ComputerName$)
  Me.HelpContextID = hlpAccountsPayable
  If DelayExit = True Then
    DelayExit = False
    Timer2.Enabled = True
  Else
    cmdExitAPMainMenu.Enabled = True
  End If
End Sub
Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
'   ' Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
'    Me.SetFocus
  End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExitAPMainMenu.Enabled = True Then
      If MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        MainLog "Close AP"
        ClearInUse PWcnt
      End If
    Else
      Cancel = True
    End If
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      If cmdExitAPMainMenu.Enabled = True Then
        cmdExitAPMainMenu_Click
      End If
      KeyCode = 0
      DoEvents
    Case Else:
  End Select
End Sub

Private Sub Timer1_Timer()
  Unload frmAPMainMenu
End Sub

Private Sub Timer2_Timer()
  cmdExitAPMainMenu.Enabled = True
End Sub
