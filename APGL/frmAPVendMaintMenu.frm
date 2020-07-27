VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmAPVendMaintMenu 
   AutoRedraw      =   -1  'True
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "A/P Vendor Maintenance"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   12225
   Icon            =   "frmAPVendMaintMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   12225
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin fpBtnAtlLibCtl.fpBtn fpBtn1 
      Height          =   405
      Left            =   270
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2805
      Visible         =   0   'False
      Width           =   465
      _Version        =   131072
      _ExtentX        =   820
      _ExtentY        =   714
      Enabled         =   0   'False
      MousePointer    =   0
      Object.TabStop         =   0   'False
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
      ButtonDesigner  =   "frmAPVendMaintMenu.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdAddVendor 
      Height          =   492
      Left            =   4308
      TabIndex        =   0
      Top             =   2472
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
      ButtonDesigner  =   "frmAPVendMaintMenu.frx":0AA3
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdEditVendor 
      Height          =   492
      Left            =   4308
      TabIndex        =   1
      Top             =   3096
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
      ButtonDesigner  =   "frmAPVendMaintMenu.frx":0C8B
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSetDefDist 
      Height          =   492
      Left            =   4308
      TabIndex        =   2
      Top             =   3732
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
      ButtonDesigner  =   "frmAPVendMaintMenu.frx":0E70
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPrintVenList 
      Height          =   492
      Left            =   4308
      TabIndex        =   3
      Top             =   4356
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
      ButtonDesigner  =   "frmAPVendMaintMenu.frx":1060
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPrintVenFile 
      Height          =   492
      Left            =   4308
      TabIndex        =   4
      Top             =   4980
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
      ButtonDesigner  =   "frmAPVendMaintMenu.frx":1249
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPrintVenLabl 
      Height          =   480
      Left            =   4305
      TabIndex        =   5
      Top             =   5610
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
      ButtonDesigner  =   "frmAPVendMaintMenu.frx":1432
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExitVendMaintMenu 
      Height          =   492
      Left            =   4308
      TabIndex        =   9
      Top             =   7488
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
      ButtonDesigner  =   "frmAPVendMaintMenu.frx":161D
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdVendorUtil 
      Height          =   480
      Left            =   4305
      TabIndex        =   7
      Top             =   6870
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
      ButtonDesigner  =   "frmAPVendMaintMenu.frx":1811
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPrintVenDist 
      Height          =   492
      Left            =   4308
      TabIndex        =   6
      Top             =   6240
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
      ButtonDesigner  =   "frmAPVendMaintMenu.frx":19F7
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   4
      X1              =   8880
      X2              =   9840
      Y1              =   2424
      Y2              =   2424
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   2400
      X2              =   3360
      Y1              =   2424
      Y2              =   2424
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   2
      X1              =   9000
      X2              =   9000
      Y1              =   2448
      Y2              =   8256
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   5
      X1              =   8880
      X2              =   9840
      Y1              =   2304
      Y2              =   2304
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000009&
      Index           =   2
      X1              =   8880
      X2              =   8880
      Y1              =   2304
      Y2              =   2424
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000009&
      Index           =   2
      X1              =   9840
      X2              =   9840
      Y1              =   2304
      Y2              =   2424
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000009&
      Height          =   1092
      Left            =   1800
      Top             =   1080
      Width           =   8652
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A/P VENDOR MAINTENANCE MENU"
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
      Left            =   3516
      TabIndex        =   10
      Top             =   1440
      Width           =   5292
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   1
      X1              =   2400
      X2              =   3360
      Y1              =   2304
      Y2              =   2304
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   2400
      X2              =   2400
      Y1              =   2304
      Y2              =   2424
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   3360
      X2              =   3360
      Y1              =   2304
      Y2              =   2424
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   2520
      X2              =   3216
      Y1              =   8280
      Y2              =   8280
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   0
      X1              =   2520
      X2              =   2520
      Y1              =   2448
      Y2              =   8280
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   9000
      X2              =   9696
      Y1              =   8280
      Y2              =   8280
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00D0D0D0&
      BorderColor     =   &H00000000&
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
      Height          =   300
      Index           =   2
      Left            =   8880
      Top             =   2136
      Width           =   972
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D0D0D0&
      BorderColor     =   &H00000000&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5940
      Index           =   2
      Left            =   9000
      Top             =   2352
      Width           =   732
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00D0D0D0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillColor       =   &H00D0D0D0&
      Height          =   300
      Left            =   2400
      Top             =   2136
      Width           =   972
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D0D0D0&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5940
      Index           =   0
      Left            =   2520
      Top             =   2352
      Width           =   732
   End
End
Attribute VB_Name = "frmAPVendMaintMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Dim VendorIdx As VendorIdxRecType
Dim Vendor As VendorRecType
Private Temp_Class As Resize_Class

Private Sub cmdEditVendor_Click()
  If Exist("APVENDOR.DAT") Then
    frmLoadingRpt.Show
    DoEvents
    Load frmEditVendor
    frmEditVendor.Show
    DoEvents
    Unload frmLoadingRpt
    Unload frmAPVendMaintMenu
  Else
    MsgBox "Vendor Information Must Be Completed First.", vbOKOnly, "Missing Information"
  End If
End Sub

Private Sub cmdPrintVenDist_Click()
  If Exist("APVENDOR.DAT") Then
    frmPrnVendDefDist.Show
    Unload Me
  Else
    MsgBox "Vendor Information Must Be Completed First.", vbOKOnly, "Missing Information"
  End If
  'cmdPrintVenDist.SetFocus
End Sub

Private Sub cmdPrintVenFile_Click()
  If Exist("APVENDOR.DAT") Then
    frmPrnVendFile.Show
    Unload Me
  Else
    MsgBox "Vendor Information Must Be Completed First.", vbOKOnly, "Missing Information"
  End If
  'cmdPrintVenFile.SetFocus
End Sub

Private Sub cmdPrintVenLabl_Click()
  If Exist("APVENDOR.DAT") Then
    frmPrnVendLabl.Show
    Unload Me
  Else
    MsgBox "Vendor Information Must Be Completed First.", vbOKOnly, "Missing Information"
  End If
 ' cmdPrintVenLabl.SetFocus
End Sub

Private Sub cmdPrintVenList_Click()
  If Exist("APVENDOR.DAT") Then
    frmPrnVendList.Show
    Unload frmAPVendMaintMenu
  Else
    MsgBox "Vendor Information Must Be Completed First.", vbOKOnly, "Missing Information"
  End If
 ' cmdPrintVenList.SetFocus
End Sub

Private Sub cmdSetDefDist_Click()
  If Exist("APVENDOR.DAT") Then
    frmVenSetDefDist.Show
    Unload frmAPVendMaintMenu
  Else
    MsgBox "Vendor Information Must Be Completed First.", vbOKOnly, "Missing Information"
  End If
End Sub

Private Sub cmdVendorUtil_Click()
  If Exist("APVENDOR.DAT") Then
    If MsgBox("Continue with re-index vendors?", vbYesNo, "Continue?") = vbYes Then
      DeActivateControls frmAPVendMaintMenu
      IndexVendorFile frmAPVendMaintMenu
      ActivateControls frmAPVendMaintMenu
    End If
  Else
    MsgBox "Vendor Information Must Be Completed First.", vbOKOnly, "Missing Information"
  End If
End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  GetAcctStruct GLUserName$, GLFundLen, GLAcctLen, GLDetLen
  Me.HelpContextID = hlpAPVendMaint
  cmdVendorUtil.HelpContextID = hlpVendUtil
End Sub
Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
'   ' Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
 '   Me.SetFocus
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      cmdExitVendMaintMenu_Click
      KeyCode = 0
      DoEvents
    Case Else:
  End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExitVendMaintMenu.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        MainLog "Close AP"
        ClearInUse PWcnt
      End If
    End If
  End If
End Sub

Private Sub cmdAddVendor_Click()
  frmLoadingRpt.Show
  DoEvents
  Load frmAddVendor
  frmAddVendor.Show
  DoEvents
  Unload frmLoadingRpt
  Unload frmAPVendMaintMenu
End Sub

Private Sub cmdExitVendMaintMenu_Click()
  frmAPMainMenu.Show
  Unload frmAPVendMaintMenu
End Sub

