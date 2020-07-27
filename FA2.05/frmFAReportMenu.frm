VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmFAReportMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fixed Assets Report Menu"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmFAReportMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleMode       =   0  'User
   ScaleWidth      =   11652
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin fpBtnAtlLibCtl.fpBtn cmdAssetsByCode 
      Height          =   495
      Left            =   3960
      TabIndex        =   7
      ToolTipText     =   "Click this button to begin retrieving data to create a detailed report of fixed asset attributes."
      Top             =   5655
      Width           =   3750
      _Version        =   131072
      _ExtentX        =   6615
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
      ButtonDesigner  =   "frmFAReportMenu.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdNewDeletedItem 
      Height          =   495
      Left            =   3960
      TabIndex        =   6
      ToolTipText     =   $"frmFAReportMenu.frx":0AB2
      Top             =   5070
      Width           =   3750
      _Version        =   131072
      _ExtentX        =   6615
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
      ButtonDesigner  =   "frmFAReportMenu.frx":0B3C
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdWarrantyRpt 
      Height          =   495
      Left            =   3960
      TabIndex        =   2
      ToolTipText     =   "Click this button to begin retrieving warranty information for fixed assets."
      Top             =   2790
      Width           =   3750
      _Version        =   131072
      _ExtentX        =   6615
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
      ButtonDesigner  =   "frmFAReportMenu.frx":0D2A
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdMasterItemList 
      Height          =   492
      Left            =   3960
      TabIndex        =   1
      ToolTipText     =   "Click this button to begin retrieving data to create a detailed report of fixed asset attributes."
      Top             =   2208
      Width           =   3756
      _Version        =   131072
      _ExtentX        =   6625
      _ExtentY        =   868
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
      ButtonDesigner  =   "frmFAReportMenu.frx":0F0D
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdDprHistRpt 
      Height          =   495
      Left            =   3960
      TabIndex        =   3
      ToolTipText     =   "Click this button to begin retrieving data pertaining to fixed asset depreciation for a given year.."
      Top             =   3360
      Width           =   3750
      _Version        =   131072
      _ExtentX        =   6615
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
      ButtonDesigner  =   "frmFAReportMenu.frx":10F6
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdDeprHistByItem 
      Height          =   495
      Left            =   3960
      TabIndex        =   4
      ToolTipText     =   "Click this button to begin retrieving depreciation information for selected fixed assets."
      Top             =   3924
      Width           =   3750
      _Version        =   131072
      _ExtentX        =   6615
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
      ButtonDesigner  =   "frmFAReportMenu.frx":12E6
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdItemCheckList 
      Height          =   495
      Left            =   3960
      TabIndex        =   5
      ToolTipText     =   "Click this button to begin retrieving data for a simple list of fixed assets."
      Top             =   4500
      Width           =   3750
      _Version        =   131072
      _ExtentX        =   6615
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
      ButtonDesigner  =   "frmFAReportMenu.frx":14D6
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdFund 
      Height          =   495
      Left            =   3960
      TabIndex        =   8
      ToolTipText     =   "Click this button to begin retrieving data to create a detailed report of fixed asset attributes."
      Top             =   6216
      Width           =   3750
      _Version        =   131072
      _ExtentX        =   6615
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
      ButtonDesigner  =   "frmFAReportMenu.frx":16C0
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdItemValRpt 
      Height          =   495
      Left            =   3960
      TabIndex        =   9
      ToolTipText     =   "Click this button to begin retrieving information regarding the worth of a fixed asset."
      Top             =   6792
      Width           =   3750
      _Version        =   131072
      _ExtentX        =   6615
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
      ButtonDesigner  =   "frmFAReportMenu.frx":18A7
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   495
      Left            =   3960
      TabIndex        =   10
      ToolTipText     =   "Click this button to begin retrieving information regarding the worth of a fixed asset."
      Top             =   7368
      Width           =   3750
      _Version        =   131072
      _ExtentX        =   6615
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
      ButtonDesigner  =   "frmFAReportMenu.frx":1A8C
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   135
      Index           =   3
      Left            =   2110
      Top             =   2091
      Width           =   960
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
   Begin VB.Line Line14 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   8692.762
      X2              =   9402.579
      Y1              =   7881.747
      Y2              =   7881.747
   End
   Begin VB.Line Line13 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2219.428
      X2              =   2929.246
      Y1              =   7881.747
      Y2              =   7881.747
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   1095
      Index           =   1
      Left            =   1500
      Top             =   895
      Width           =   8655
   End
   Begin VB.Line Line12 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2220.428
      X2              =   2220.428
      Y1              =   2146.112
      Y2              =   7876.874
   End
   Begin VB.Line Line11 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      X1              =   8699.76
      X2              =   8699.76
      Y1              =   2149.036
      Y2              =   7879.797
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "FIXED ASSETS REPORT MENU"
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
Attribute VB_Name = "frmFAReportMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsFATextBoxOverRider
  Private Temp_Class As Resize_Class

Private Sub cmdAssetsByCode_Click()
  Dim TagIdx As TagNumbSortIdxType
  Dim TagIdxHandle As Integer
  Dim TagIdxCnt As Long
  Dim CodeRec As FAAssetCodeRecType
  Dim CodeHandle As Integer
  Dim NumOfCodeRecs As Integer
  
  OpenTagIdxFile TagIdxHandle
  TagIdxCnt = LOF(TagIdxHandle) / Len(TagIdx)
  Close TagIdxHandle
  If TagIdxCnt = 0 Then
    MsgBox "No fixed assets have been saved."
    Close
    Exit Sub
  End If
  
  OpenFACodeNameFile CodeHandle
  NumOfCodeRecs = LOF(CodeHandle) / Len(CodeRec)
  Close CodeHandle
  If NumOfCodeRecs = 0 Then
    MsgBox "No asset code records saved."
    Close
    Exit Sub
  End If
  
  frmFAAssByCodeRpt.Show
  DoEvents
  Unload frmFAReportMenu
End Sub

Private Sub cmdDeprHistByItem_Click()
  Dim TagIdx As TagNumbSortIdxType
  Dim TagIdxHandle As Integer
  Dim TagIdxCnt As Long

  OpenTagIdxFile TagIdxHandle
  TagIdxCnt = LOF(TagIdxHandle) / Len(TagIdx)
  Close TagIdxHandle
  If TagIdxCnt = 0 Then
    MsgBox "No fixed assets have been saved."
    Close
    Exit Sub
  End If
  
  frmFADprHistByItem.Show
  DoEvents
  Unload frmFAReportMenu

End Sub

Private Sub cmdDprHistRpt_Click()
  Dim TagIdx As TagNumbSortIdxType
  Dim TagIdxHandle As Integer
  Dim TagIdxCnt As Long

  OpenTagIdxFile TagIdxHandle
  TagIdxCnt = LOF(TagIdxHandle) / Len(TagIdx)
  Close TagIdxHandle
  If TagIdxCnt = 0 Then
    MsgBox "No fixed assets have been saved."
    Close
    Exit Sub
  End If
  
  frmFADprHistRpt.Show
  DoEvents
  Unload frmFAReportMenu
End Sub

Private Sub cmdExit_Click()
  frmFAMainMenu.Show
  Close
  DoEvents
  Unload frmFAReportMenu
End Sub

Private Sub cmdFund_Click()
  Dim TagIdx As TagNumbSortIdxType
  Dim TagIdxHandle As Integer
  Dim TagIdxCnt As Long
  Dim FundRec As FAFundCodeType
  Dim FundHandle As Integer
  Dim NumOfFundRecs As Integer
  
  OpenTagIdxFile TagIdxHandle
  TagIdxCnt = LOF(TagIdxHandle) / Len(TagIdx)
  Close TagIdxHandle
  If TagIdxCnt = 0 Then
    MsgBox "No fixed assets have been saved."
    Close
    Exit Sub
  End If
  
  OpenFAFundCodeFile FundHandle
  NumOfFundRecs = LOF(FundHandle) / Len(FundRec)
  Close FundHandle
  If NumOfFundRecs = 0 Then
    MsgBox "No fund code records saved."
    Close
    Exit Sub
  End If
  
  frmFAAssByFundRpt.Show
  DoEvents
  Unload frmFAReportMenu
End Sub

Private Sub cmdItemCheckList_Click()
  Dim TagIdx As TagNumbSortIdxType
  Dim TagIdxHandle As Integer
  Dim TagIdxCnt As Long

  OpenTagIdxFile TagIdxHandle
  TagIdxCnt = LOF(TagIdxHandle) / Len(TagIdx)
  Close TagIdxHandle
  If TagIdxCnt = 0 Then
    MsgBox "No fixed assets have been saved."
    Close
    Exit Sub
  End If
  
  frmFAItemCheckList.Show
  DoEvents
  Unload frmFAReportMenu
End Sub

Private Sub cmdItemValRpt_Click()
  Dim TagIdx As TagNumbSortIdxType
  Dim TagIdxHandle As Integer
  Dim TagIdxCnt As Long

  OpenTagIdxFile TagIdxHandle
  TagIdxCnt = LOF(TagIdxHandle) / Len(TagIdx)
  Close TagIdxHandle
  If TagIdxCnt = 0 Then
    MsgBox "No fixed assets have been saved."
    Close
    Exit Sub
  End If
  
  frmFAValueRange.Show
  DoEvents
  Unload frmFAReportMenu

End Sub

Private Sub cmdMasterItemList_Click()
  Dim TagIdx As TagNumbSortIdxType
  Dim TagIdxHandle As Integer
  Dim TagIdxCnt As Long

  OpenTagIdxFile TagIdxHandle
  TagIdxCnt = LOF(TagIdxHandle) / Len(TagIdx)
  Close TagIdxHandle
  If TagIdxCnt = 0 Then
    MsgBox "No fixed assets have been saved."
    Close
    Exit Sub
  End If
  
  frmFAMasterItemListing.Show
  DoEvents
  Unload frmFAReportMenu
End Sub

Private Sub cmdNewDeletedItem_Click()
  Dim TagIdx As TagNumbSortIdxType
  Dim TagIdxHandle As Integer
  Dim TagIdxCnt As Long

  OpenTagIdxFile TagIdxHandle
  TagIdxCnt = LOF(TagIdxHandle) / Len(TagIdx)
  Close TagIdxHandle
  If TagIdxCnt = 0 Then
    MsgBox "No fixed assets have been saved."
    Close
    Exit Sub
  End If
  
  frmFAItemsAddDelOptRpt.Show
  DoEvents
  Unload frmFAReportMenu
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdWarrantyRpt_Click()
  Dim TagIdx As TagNumbSortIdxType
  Dim TagIdxHandle As Integer
  Dim TagIdxCnt As Long

  OpenTagIdxFile TagIdxHandle
  TagIdxCnt = LOF(TagIdxHandle) / Len(TagIdx)
  Close TagIdxHandle
  If TagIdxCnt = 0 Then
    MsgBox "No fixed assets have been saved."
    Close
    Exit Sub
  End If
  
  frmWarrantyRpt.Show
  DoEvents
  Unload frmFAReportMenu

End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsFATextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)

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
      SendKeys "%X"
      Call cmdExit_Click
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
      MainLog ("FixedAssets.exe terminated via menu bar on frmFAReportMenu.")
      Call Terminate
      End
    End If
  End If
End Sub



