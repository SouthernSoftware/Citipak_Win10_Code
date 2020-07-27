VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmBLOmitList 
   BackColor       =   &H000000C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Omission List"
   ClientHeight    =   6585
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   6930
   Icon            =   "frmBLOmitList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   6930
   StartUpPosition =   2  'CenterScreen
   Begin LpLib.fpList fpList1 
      Height          =   2100
      Left            =   390
      TabIndex        =   1
      Top             =   2550
      Width           =   6195
      _Version        =   196608
      _ExtentX        =   10927
      _ExtentY        =   3704
      TextAlias       =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Columns         =   0
      Sorted          =   0
      LineWidth       =   1
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   -1
      ColumnWidthScale=   2
      RowHeight       =   -1
      MultiSelect     =   0
      WrapList        =   0   'False
      WrapWidth       =   0
      SelMax          =   -1
      AutoSearch      =   1
      SearchMethod    =   0
      VirtualMode     =   0   'False
      VRowCount       =   0
      DataSync        =   3
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483627
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   -2147483642
      BorderWidth     =   1
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ScrollHScale    =   2
      ScrollHInc      =   0
      ColsFrozen      =   0
      ScrollBarV      =   1
      NoIntegralHeight=   0   'False
      HighestPrecedence=   0
      AllowColResize  =   0
      AllowColDragDrop=   0
      ReadOnly        =   0   'False
      VScrollSpecial  =   0   'False
      VScrollSpecialType=   0
      EnableKeyEvents =   -1  'True
      EnableTopChangeEvent=   -1  'True
      DataAutoHeadings=   -1  'True
      DataAutoSizeCols=   2
      SearchIgnoreCase=   -1  'True
      ScrollBarH      =   1
      VirtualPageSize =   0
      VirtualPagesAhead=   0
      ExtendCol       =   0
      ColumnLevels    =   1
      ListGrayAreaColor=   -2147483637
      GroupHeaderHeight=   -1
      GroupHeaderShow =   0   'False
      AllowGrpResize  =   0
      AllowGrpDragDrop=   0
      MergeAdjustView =   0   'False
      ColumnHeaderShow=   0   'False
      ColumnHeaderHeight=   -1
      GrpsFrozen      =   0
      BorderGrayAreaColor=   -2147483637
      ExtendRow       =   0
      DataField       =   ""
      OLEDragMode     =   0
      OLEDropMode     =   0
      Redraw          =   -1  'True
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      ColDesigner     =   "frmBLOmitList.frx":08CA
   End
   Begin VB.TextBox fptxtChoice 
      Height          =   288
      Left            =   6384
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   5856
      Width           =   492
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   540
      Left            =   810
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5085
      Width           =   1800
      _Version        =   131072
      _ExtentX        =   3175
      _ExtentY        =   952
      Enabled         =   -1  'True
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
      ButtonDesigner  =   "frmBLOmitList.frx":0C3A
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPrint 
      Height          =   540
      Left            =   4410
      TabIndex        =   3
      TabStop         =   0   'False
      Tag             =   $"frmBLOmitList.frx":0E17
      Top             =   5760
      Width           =   1800
      _Version        =   131072
      _ExtentX        =   3175
      _ExtentY        =   952
      Enabled         =   -1  'True
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
      ButtonDesigner  =   "frmBLOmitList.frx":0EC2
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdDontDelete 
      Height          =   540
      Left            =   2880
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5088
      Width           =   3324
      _Version        =   131072
      _ExtentX        =   5863
      _ExtentY        =   952
      Enabled         =   -1  'True
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
      ButtonDesigner  =   "frmBLOmitList.frx":109E
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdDelete 
      Height          =   540
      Left            =   816
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5760
      Width           =   3324
      _Version        =   131072
      _ExtentX        =   5863
      _ExtentY        =   952
      Enabled         =   -1  'True
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
      ButtonDesigner  =   "frmBLOmitList.frx":1286
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   2124
      Left            =   384
      Top             =   240
      Width           =   6204
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1836
      Left            =   528
      TabIndex        =   0
      Top             =   384
      Width           =   5916
   End
End
Attribute VB_Name = "frmBLOmitList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsBLTextBoxOverrider
  Private Temp_Class As Resize_Class

Private Sub cmdDelete_Click()
  fptxtChoice.Text = "delete"
  Me.Hide
End Sub

Private Sub cmdDontDelete_Click()
  fptxtChoice.Text = "continue"
  Me.Hide
End Sub

Private Sub cmdExit_Click()
  fptxtChoice.Text = "abort"
  Me.Hide
End Sub

Private Sub cmdPrint_Click()
  Dim RptHandle As Integer
  Dim ReportFile$
  Dim MaxLines As Integer
  Dim LineCnt As Integer
  Dim x As Integer
  Dim CustRec As ARCustRecType
  Dim CHandle As Integer
  Dim FF$
  
  FF$ = Chr$(12)

  MaxLines = 50
  If OmitCnt > 0 Then
    OpenCustFile CHandle
    ReportFile$ = "AROmitLt.PRN"
    RptHandle = FreeFile
    Open ReportFile$ For Output As #RptHandle
    GoSub PrintHeader
    For x = 1 To OmitCnt
      If LineCnt > MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintHeader
      End If
      Get CHandle, OmitList(x), CustRec
      LineCnt = LineCnt + 1
      Print #RptHandle, Tab(5); QPTrim$(CustRec.CustName); Tab(45); QPTrim$(CustRec.CustNumb)
    Next x
    Print #RptHandle, FF$
    Close CHandle
    Close RptHandle
    ViewPrint ReportFile$, "Customers Omitted Listing", True
    KillFile ReportFile$
  
  End If
  MainLog ("User printed a list of all customers excluded from license processing because they were already involved in an unposted penalty file.")
  
  Exit Sub
  
PrintHeader:
  Print #RptHandle, "List Of Customers Included In Unposted Penalty Process"
  Print #RptHandle, Date
  Print #RptHandle, ""
  Print #RptHandle, Tab(5); "Customer Name"; Tab(40); "Customer Number"
  Print #RptHandle, String$(55, "=")
  LineCnt = LineCnt + 5
  Return
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%A"
      Call cmdExit_Click
      KeyCode = 0
    Case vbKeyF3:
      SendKeys "%P"
      Call cmdPrint_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%C"
      Call cmdDontDelete_Click
      KeyCode = 0
    Case vbKeyF5:
      SendKeys "%D"
      Call cmdDelete_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Dim x As Integer
  Dim CustRec As ARCustRecType
  Dim CHandle As Integer
  
  fptxtChoice.Visible = False
  OpenCustFile CHandle

  For x = 1 To OmitCnt
    Get CHandle, OmitList(x), CustRec
    fpList1.AddItem QPTrim(CustRec.CustName) + "   #" + QPTrim$(CustRec.CustNumb)
  Next x
  Close CHandle
End Sub

Private Sub fpBtn1_Click()

End Sub
