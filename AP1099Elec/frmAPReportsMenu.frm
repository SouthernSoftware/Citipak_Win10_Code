VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "BTN32A20.OCX"
Begin VB.Form frmAPReportsMenu 
   AutoRedraw      =   -1  'True
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "A/P Reports"
   ClientHeight    =   8868
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   12216
   Icon            =   "frmAPReportsMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   12216
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin fpBtnAtlLibCtl.fpBtn cmdVendHist 
      Height          =   420
      Left            =   4278
      TabIndex        =   0
      Top             =   2304
      Width           =   3660
      _Version        =   131072
      _ExtentX        =   6456
      _ExtentY        =   741
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
      ButtonDesigner  =   "frmAPReportsMenu.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdOpenPay 
      Height          =   420
      Left            =   4278
      TabIndex        =   1
      Top             =   2808
      Width           =   3660
      _Version        =   131072
      _ExtentX        =   6456
      _ExtentY        =   741
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
      ButtonDesigner  =   "frmAPReportsMenu.frx":0AE8
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdOpenPayDate 
      Height          =   420
      Left            =   4278
      TabIndex        =   2
      Top             =   3312
      Width           =   3660
      _Version        =   131072
      _ExtentX        =   6456
      _ExtentY        =   741
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
      ButtonDesigner  =   "frmAPReportsMenu.frx":0D05
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdOpenPOs 
      Height          =   420
      Left            =   4278
      TabIndex        =   3
      Top             =   3804
      Width           =   3660
      _Version        =   131072
      _ExtentX        =   6456
      _ExtentY        =   741
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
      ButtonDesigner  =   "frmAPReportsMenu.frx":0F2A
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPOHistGL 
      Height          =   420
      Left            =   4278
      TabIndex        =   5
      Top             =   4812
      Width           =   3660
      _Version        =   131072
      _ExtentX        =   6456
      _ExtentY        =   741
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
      ButtonDesigner  =   "frmAPReportsMenu.frx":114E
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdAPChkList 
      Height          =   420
      Left            =   4278
      TabIndex        =   6
      Top             =   5316
      Width           =   3660
      _Version        =   131072
      _ExtentX        =   6456
      _ExtentY        =   741
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
      ButtonDesigner  =   "frmAPReportsMenu.frx":133D
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPaidSupLst 
      Height          =   420
      Left            =   4278
      TabIndex        =   7
      Top             =   5808
      Width           =   3660
      _Version        =   131072
      _ExtentX        =   6456
      _ExtentY        =   741
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
      ButtonDesigner  =   "frmAPReportsMenu.frx":1526
   End
   Begin fpBtnAtlLibCtl.fpBtn cmd1099Forms 
      Height          =   420
      Left            =   4278
      TabIndex        =   8
      Top             =   6312
      Width           =   3660
      _Version        =   131072
      _ExtentX        =   6456
      _ExtentY        =   741
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
      ButtonDesigner  =   "frmAPReportsMenu.frx":1711
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSalesTaxRpt 
      Height          =   420
      Left            =   4278
      TabIndex        =   9
      Top             =   6816
      Width           =   3660
      _Version        =   131072
      _ExtentX        =   6456
      _ExtentY        =   741
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
      ButtonDesigner  =   "frmAPReportsMenu.frx":18FE
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSalesTaxCounty 
      Height          =   420
      Left            =   4278
      TabIndex        =   10
      Top             =   7320
      Width           =   3660
      _Version        =   131072
      _ExtentX        =   6456
      _ExtentY        =   741
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
      ButtonDesigner  =   "frmAPReportsMenu.frx":1AE6
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExitAPRptMenu 
      Height          =   420
      Left            =   4278
      TabIndex        =   11
      Top             =   7824
      Width           =   3660
      _Version        =   131072
      _ExtentX        =   6456
      _ExtentY        =   741
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
      ButtonDesigner  =   "frmAPReportsMenu.frx":1CD5
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdClosedPOs 
      Height          =   420
      Left            =   4278
      TabIndex        =   4
      Top             =   4308
      Width           =   3660
      _Version        =   131072
      _ExtentX        =   6456
      _ExtentY        =   741
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
      ButtonDesigner  =   "frmAPReportsMenu.frx":1EC7
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
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   4
      X1              =   8880
      X2              =   9840
      Y1              =   2400
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
   Begin VB.Line Line5 
      BorderColor     =   &H80000009&
      Index           =   2
      X1              =   8880
      X2              =   8880
      Y1              =   2280
      Y2              =   2400
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000009&
      Index           =   2
      X1              =   9840
      X2              =   9840
      Y1              =   2280
      Y2              =   2400
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
      Caption         =   "A/P REPORT PROCESSING MENU"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3132
      TabIndex        =   12
      Top             =   1440
      Width           =   5988
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   2400
      X2              =   3360
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   1
      X1              =   2424
      X2              =   3360
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   2400
      X2              =   2400
      Y1              =   2280
      Y2              =   2400
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   3360
      X2              =   3360
      Y1              =   2280
      Y2              =   2400
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   2544
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
      Y1              =   2424
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
      Height          =   5940
      Index           =   0
      Left            =   2520
      Top             =   2352
      Width           =   732
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000000&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   2
      Left            =   8880
      Top             =   2160
      Width           =   972
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
End
Attribute VB_Name = "frmAPReportsMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class

Private Sub cmd1099Forms_Click()
  frm1099ProcessMenu.Show
  Unload frmAPReportsMenu
End Sub

Private Sub cmdAPChkList_Click()
  frmPrnAPChkList.Show
  Unload frmAPReportsMenu
End Sub

Private Sub cmdClosedPOs_Click()
  frmPrnClosedPOs.Show
  Unload frmAPReportsMenu
End Sub

Private Sub cmdOpenPay_Click()
  frmPrnOpenPays.Show
  Unload frmAPReportsMenu
End Sub

Private Sub cmdOpenPayDate_Click()
  frmPrnOpenPayDate.Show
  Unload frmAPReportsMenu
End Sub

Private Sub cmdOpenPOs_Click()
  frmPrnOpenPOs.Show
  Unload frmAPReportsMenu
End Sub

Private Sub cmdPaidSupLst_Click()
  frmPrnPaidSupList.Show
  Unload frmAPReportsMenu
End Sub

Private Sub cmdPOHistGL_Click()
  frmPrnPOHistGL.Show
  Unload frmAPReportsMenu
End Sub

Private Sub cmdSalesTaxCounty_Click()
  frmPrnSalesTaxRptCounty.Show
  Unload Me
End Sub

Private Sub cmdSalesTaxRpt_Click()
  frmPrnSalesTaxRpt.Show
  Unload frmAPReportsMenu
End Sub

Private Sub cmdVendHist_Click()
  frmPrnVendHist.Show
  Unload frmAPReportsMenu
End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Me.HelpContextID = hlpAPReports
End Sub
Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
'    Me.SetFocus
  End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
      If MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        MainLog "Close AP"
        ClearInUse PWcnt
      End If
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      cmdExitAPRptMenu_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub


Private Sub cmdExitAPRptMenu_Click()
  frmAPMainMenu.Show
  Unload frmAPReportsMenu
End Sub
