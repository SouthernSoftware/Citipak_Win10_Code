VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmVATaxPayEditList 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Edit Transaction List"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmVATaxPayEditList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpList fpListPPay 
      Height          =   1680
      Left            =   720
      TabIndex        =   6
      Top             =   4800
      Width           =   10095
      _Version        =   196608
      _ExtentX        =   17806
      _ExtentY        =   2963
      TextAlias       =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Columns         =   4
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
      ColumnHeaderShow=   -1  'True
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
      ColDesigner     =   "frmVATaxPayEditList.frx":08CA
   End
   Begin LpLib.fpList fpListRPay 
      Height          =   1680
      Left            =   720
      TabIndex        =   5
      Top             =   2160
      Width           =   10095
      _Version        =   196608
      _ExtentX        =   17806
      _ExtentY        =   2963
      TextAlias       =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Columns         =   4
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
      ColumnHeaderShow=   -1  'True
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
      ColDesigner     =   "frmVATaxPayEditList.frx":0C3E
   End
   Begin EditLib.fpText fptxtOperator 
      Height          =   372
      Left            =   4320
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1140
      Width           =   2892
      _Version        =   196608
      _ExtentX        =   5106
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   -2147483642
      BorderWidth     =   1
      ButtonDisable   =   0   'False
      ButtonHide      =   0   'False
      ButtonIncrement =   1
      ButtonMin       =   0
      ButtonMax       =   100
      ButtonStyle     =   0
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   1
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
      AutoBeep        =   0   'False
      AutoCase        =   0
      CaretInsert     =   0
      CaretOverWrite  =   3
      UserEntry       =   0
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483637
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   0
      ControlType     =   1
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   255
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   540
      Left            =   3660
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   7080
      Width           =   2040
      _Version        =   131072
      _ExtentX        =   3598
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
      ButtonDesigner  =   "frmVATaxPayEditList.frx":0FB2
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdEdit 
      Height          =   540
      Left            =   5925
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   7080
      Width           =   2055
      _Version        =   131072
      _ExtentX        =   3625
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
      ButtonDesigner  =   "frmVATaxPayEditList.frx":1190
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPrint 
      Height          =   540
      Left            =   4140
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   7785
      Width           =   3375
      _Version        =   131072
      _ExtentX        =   5953
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
      ButtonDesigner  =   "frmVATaxPayEditList.frx":136C
   End
   Begin VB.Label Label3 
      BackColor       =   &H008F8265&
      Caption         =   "Personal Transactions"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   252
      Left            =   600
      TabIndex        =   8
      Top             =   4320
      Width           =   2772
   End
   Begin VB.Label Label1 
      BackColor       =   &H008F8265&
      Caption         =   "Real Transactions"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   252
      Left            =   600
      TabIndex        =   7
      Top             =   1680
      Width           =   2772
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   2052
      Left            =   600
      Top             =   4680
      Width           =   10452
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   2052
      Left            =   600
      Top             =   2040
      Width           =   10452
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Payment List"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Left            =   3132
      TabIndex        =   1
      Top             =   528
      Width           =   5292
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   660
      Index           =   1
      Left            =   1488
      Top             =   372
      Width           =   8652
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   780
      Left            =   1488
      Top             =   252
      Width           =   8652
   End
End
Attribute VB_Name = "frmVATaxPayEditList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class
  Dim PrintType As String

Private Sub cmdEdit_Click()
  If fpListRPay.SelCount > 0 Then
    Call RealEdit
  ElseIf fpListPPay.SelCount > 0 Then
    Call PersEdit
  Else
    Call TaxMsg(900, "Please make a selection from one of the two lists.")
  End If
End Sub

Private Sub cmdExit_Click()
  KillFile "C:\CPWork\editpyment.dat"
  GCustNum = 0
  Unload Me
  DoEvents
  frmVATaxPayMenu.Show
End Sub

Private Sub cmdPrint_Click()
  frmVATaxRptOptForPayEdit.Show vbModal
  If frmVATaxRptOptForPayEdit.fptxtPrintType.Text = "Graphical Name" Then
    Unload frmVATaxRptOptForPayEdit
    PrintType = "N"
    Call PrintGraphics
  ElseIf frmVATaxRptOptForPayEdit.fptxtPrintType.Text = "Graphical Entry" Then
    Unload frmVATaxRptOptForPayEdit
    PrintType = "E"
    Call PrintGraphics
  ElseIf frmVATaxRptOptForPayEdit.fptxtPrintType.Text = "Text Name" Then
    PrintType = "N"
    frmVATaxMsg.Label1.Caption = "Pitch 17 is recommended for this report."
    frmVATaxMsg.Label1.Top = 900
    frmVATaxMsg.Show vbModal
    Unload frmVATaxRptOptForPayEdit
    Call PrintText
  ElseIf frmVATaxRptOptForPayEdit.fptxtPrintType.Text = "Text Entry" Then
    PrintType = "E"
    frmVATaxMsg.Label1.Caption = "Pitch 17 is recommended for this report."
    frmVATaxMsg.Label1.Top = 900
    frmVATaxMsg.Show vbModal
    Unload frmVATaxRptOptForPayEdit
    Call PrintText
  End If
End Sub
Private Sub PrintGraphics()
  Dim RPayRec As TaxPaymentRecType
  Dim RPayHandle As Integer
  Dim NumOfRRecs As Integer
  Dim PPayRec As TaxPaymentRecType
  Dim PPayHandle As Integer
  Dim NumOfPRecs As Integer
  Dim x As Integer
  Dim OHandle As Integer
  Dim NumOfOperRecs As Integer
  Dim OperRec As CitiPassType
  Dim NumOfOperRPayRecs As Integer
  Dim dlm$
  Dim TaxMasterRec As TaxMasterType
  Dim MHandle As Integer
  Dim Town$
  Dim RptFile$
  Dim RptHandle As Integer
  Dim SubRptFile1$
  Dim SubRptHandle1 As Integer
  Dim SubRptFile2$
  Dim SubRptHandle2 As Integer
  Dim SubRptFile3$
  Dim SubRptHandle3 As Integer
  Dim RTempRPayRec As RealPayListType
  Dim RTHandle As Integer
  Dim NumOfRTRecs As Integer
  Dim PTempRPayRec As PersPayListType
  Dim PTHandle As Integer
  Dim NumOfPTRecs As Integer
  Dim GPrinc As Double
  Dim RGInt As Double
  Dim GAdvCol As Double
  Dim GLateList As Double
  Dim GRPenalty As Double
  Dim GRev1 As Double
  Dim GRev2 As Double
  Dim GRev3 As Double
  Dim RGTot As Double
  Dim GPers As Double
  Dim GMachTools As Double
  Dim GMerchCap As Double
  Dim GFarmEq As Double
  Dim GMobHomes As Double
  Dim GPersonal As Double
  Dim GPInterest As Double
  Dim GPPenalty As Double
  Dim GPRev1 As Double
  Dim GPRev2 As Double
  Dim GPRev3 As Double
  Dim PGInt As Double
  Dim GOverPay As Double
  Dim RGOverPay As Double
  Dim PGOverPay As Double
  Dim PGTot As Double
  Dim Operator$
  Dim TotalPaid#
  Dim RTotalPaid#
  Dim PTotalPaid#
  Dim GDisc As Double
  Dim RGDisc As Double
  Dim PGDisc As Double
  Dim RYearCnt As Integer
  Dim PYearCnt As Integer
  Dim GYearCnt As Integer
  Dim y As Integer
  Dim ThisYear As Integer
  Dim PrincByYrTot As Double
  Dim RIntByYrTot As Double
  Dim AdvColByYrTot As Double
  Dim LateListByYrTot As Double
  Dim RPenByYrTot As Double
  Dim Rev1ByYrTot As Double
  Dim Rev2ByYrTot As Double
  Dim Rev3ByYrTot As Double
  Dim PersByYrTot As Double
  Dim PRev1ByYrTot As Double
  Dim PRev2ByYrTot As Double
  Dim PRev3ByYrTot As Double
  Dim MTByYrTot As Double
  Dim MCByYrTot As Double
  Dim FEByYrTot As Double
  Dim MHByYrTot As Double
  Dim PIntByYrTot As Double
  Dim PPenByYrTot As Double
  Dim RDiscByYrTot As Double
  Dim PDiscByYrTot As Double
  Dim GDiscByYrTot As Double
  Dim RTotPaidByYrTot As Double
  Dim PTotPaidByYrTot As Double
  Dim ROverPayByYrTot As Double
  Dim POverPayByYrTot As Double
  Dim CheckCnt As Integer
  Dim RCheckCnt As Integer
  Dim PCheckCnt As Integer
  Dim HoldPrincByYr As Double
  Dim HoldRIntByYr As Double
  Dim HoldAdvColByYr As Double
  Dim HoldLateListByYr As Double
  Dim HoldRPenByYr As Double
  Dim HoldRev1ByYr As Double
  Dim HoldRev2ByYr As Double
  Dim HoldRev3ByYr As Double
  Dim HoldRDiscByYr As Double
  Dim HoldPersByYr As Double
  Dim HoldMTByYr As Double
  Dim HoldMCByYr As Double
  Dim HoldFEByYr As Double
  Dim HoldMHByYr As Double
  Dim HoldPIntByYr As Double
  Dim HoldPPenByYr As Double
  Dim HoldPRev1ByYr As Double
  Dim HoldPRev2ByYr As Double
  Dim HoldPRev3ByYr As Double
  Dim HoldROverPayByYr As Double
  Dim HoldRTotPaidByYr As Double
  Dim HoldPOverPayByYr As Double
  Dim HoldPTotPaidByYr As Double
  Dim HoldPDiscByYr As Double
  Dim Thisx As Integer
  Dim LilYear As Integer
  Dim Nextx As Integer
  Dim HoldYears As Integer
  Dim Done As Boolean
  Dim GOPAmt As Double
  Dim ROPAmt As Double
  Dim POPAmt As Double
  Dim BigName$
  Dim LilName$
  Dim HoldName$
  Dim HoldRec As Integer
  Dim NextOne As Integer
  Dim GTotCash As Double
  Dim GTotCheck As Double
  Dim GTotCharge As Double
  Dim GTotChange As Double
  Dim GTotCount As Integer
  Dim GTotDisc As Double
  Dim RTotPaid As Double
  Dim RTotCash As Double
  Dim RTotChk As Double
  Dim RTotChrg As Double
  Dim RTotChng As Double
  Dim PTotPaid As Double
  Dim PTotCash As Double
  Dim PTotChk As Double
  Dim PTotChrg As Double
  Dim PTotChng As Double
  Dim GTotChkCnt As Single
  Dim OptDesc1$
  Dim OptDesc2$
  Dim OptDesc3$
  Dim POpt1ByYrTot As Double
  Dim POpt2ByYrTot As Double
  Dim POpt3ByYrTot As Double
  Dim Dif As Double
  Dim Disc As Double
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  
  On Error GoTo ERRORSTUFF
  
  SubRptFile1$ = "TAXRPTS\SubEdPay1.RPT"
  SubRptHandle1 = FreeFile
  Open SubRptFile1$ For Output As #SubRptHandle1
  
  OpenTaxSetUpFile MHandle
  Get MHandle, 1, TaxMasterRec
  Close MHandle
  Town = QPTrim$(TaxMasterRec.Name)
  OptDesc1$ = QPTrim$(TaxMasterRec.POptRev1)
  OptDesc2$ = QPTrim$(TaxMasterRec.POptRev2)
  OptDesc3$ = QPTrim$(TaxMasterRec.POptRev3)
  
  Operator$ = "Operator # " + CStr(OperNum) + " " + PWUser
  fptxtOperator.Text = "Operator # " + CStr(OperNum)
  OpenTempRealPayFile RPayHandle, OperNum
  NumOfRRecs = LOF(RPayHandle) / Len(RPayRec)
  If NumOfRRecs > 0 Then
    ReDim RCustNArr(1 To NumOfRRecs) As String
    ReDim RCustRArr(1 To NumOfRRecs) As Integer
    If PrintType = "N" Then
      GoSub SortRCustomers
    End If
  End If
  
  OpenTempPersPayFile PPayHandle, OperNum
  NumOfPRecs = LOF(PPayHandle) / Len(PPayRec)
  If NumOfPRecs > 0 Then
    ReDim PCustNArr(1 To NumOfPRecs) As String
    ReDim PCustRArr(1 To NumOfPRecs) As Integer
    If PrintType = "N" Then
      GoSub SortPCustomers
    End If
  End If
  
  dlm = "~"
  
  RptFile$ = "TAXRPTS\TaxEdPay.RPT"
  RptHandle = FreeFile
  RCheckCnt = 0
  PCheckCnt = 0
  ROPAmt = 0
  POPAmt = 0
  TotalPaid = 0
  RTotCash = 0
  RTotChk = 0
  RTotChrg = 0
  RTotChng = 0
  Open RptFile$ For Output As #RptHandle
  
  For x = 1 To NumOfRRecs
    If PrintType = "N" Then
      Get RPayHandle, RCustRArr(x), RPayRec
    Else
      Get RPayHandle, x, RPayRec
    End If
    If RPayRec.LastPayRec = 0 Then GoTo MoveOnR
    If RPayRec.ChkAmt > 0 Then
      RCheckCnt = RCheckCnt + 1
    End If
    TotalPaid = OldRound(RPayRec.CashAmt + RPayRec.ChkAmt + RPayRec.ChrgAmt + RPayRec.DiscAmt - RPayRec.Change)
    GOPAmt = OldRound(GOPAmt + RPayRec.PrePayAmt)
    RTotalPaid = OldRound(RTotalPaid + TotalPaid)
    ROPAmt = OldRound(ROPAmt + RPayRec.PrePayAmt)
    GTotCash = OldRound(GTotCash + RPayRec.CashAmt)
    RTotCash = OldRound(RTotCash + RPayRec.CashAmt)
    GTotCharge = OldRound(GTotCharge + RPayRec.ChrgAmt)
    RTotChrg = OldRound(RTotChrg + RPayRec.ChrgAmt)
    GTotCheck = OldRound(GTotCheck + RPayRec.ChkAmt)
    If RPayRec.ChkAmt > 0 Then GTotChkCnt = GTotChkCnt + 1
    RTotChk = OldRound(RTotChk + RPayRec.ChkAmt)
    GTotChange = OldRound(GTotChange + RPayRec.Change)
    RTotChng = OldRound(RTotChng + RPayRec.Change)
    GTotDisc = OldRound(GTotDisc + RPayRec.DiscAmt)
    GTotCount = GTotCount + 1
    '                   0                    1                               2
    Print #RptHandle, Town; dlm; MakeRegDate(RPayRec.PayDate); dlm; CStr(RPayRec.CustAcct); dlm;
    '                           3                          4                    5                     6
    Print #RptHandle, QPTrim$(RPayRec.CustName); dlm; RPayRec.CashAmt; dlm; RPayRec.ChkAmt; dlm; RPayRec.ChrgAmt; dlm;
    '                       7                   8                 9                  10               11                 12
    Print #RptHandle, RPayRec.DiscAmt; dlm; TotalPaid; dlm; RPayRec.Change; dlm; Operator; dlm; RCheckCnt; dlm; ROPAmt; dlm;
    '                 13           14               15             16            17             18
    Print #RptHandle, "R"; dlm; NumOfRRecs; dlm; RTotCash; dlm; RTotChrg; dlm; RTotChk; dlm; RTotChng; dlm;
    '                    19                20             21             22
    Print #RptHandle, GTotChkCnt; dlm; GTotCount; dlm; GTotDisc; dlm; RTotalPaid
MoveOnR:
  Next x
  
'  Close RPayHandle
  
  For x = 1 To NumOfPRecs
    If PrintType = "N" Then
      Get PPayHandle, PCustRArr(x), PPayRec
    Else
      Get PPayHandle, x, PPayRec
    End If
    If PPayRec.LastPayRec = 0 Then GoTo MoveOnP
    If PPayRec.ChkAmt > 0 Then
      PCheckCnt = PCheckCnt + 1
    End If
    TotalPaid = OldRound(PPayRec.CashAmt + PPayRec.ChkAmt + PPayRec.ChrgAmt + PPayRec.DiscAmt - PPayRec.Change)
    GOPAmt = OldRound(GOPAmt + PPayRec.PrePayAmt)
    PTotalPaid = OldRound(PTotalPaid + TotalPaid)
    POPAmt = OldRound(POPAmt + PPayRec.PrePayAmt)
    GTotCash = OldRound(GTotCash + PPayRec.CashAmt)
    PTotCash = OldRound(PTotCash + PPayRec.CashAmt)
    GTotCharge = OldRound(GTotCharge + PPayRec.ChrgAmt)
    PTotChrg = OldRound(PTotChrg + PPayRec.ChrgAmt)
    GTotCheck = OldRound(GTotCheck + PPayRec.ChkAmt)
    PTotChk = OldRound(PTotChk + PPayRec.ChkAmt)
    GTotChange = OldRound(GTotChange + PPayRec.Change)
    If PPayRec.ChkAmt > 0 Then GTotChkCnt = GTotChkCnt + 1
    PTotChng = OldRound(PTotChng + PPayRec.Change)
    GTotDisc = OldRound(GTotDisc + PPayRec.DiscAmt)
    GTotCount = GTotCount + 1
    '                   0                    1                               2
    Print #RptHandle, Town; dlm; MakeRegDate(PPayRec.PayDate); dlm; CStr(PPayRec.CustAcct); dlm;
    '                           3                          4                    5                   6
    Print #RptHandle, QPTrim$(PPayRec.CustName); dlm; PPayRec.CashAmt; dlm; PPayRec.ChkAmt; dlm; PPayRec.ChrgAmt; dlm;
    '                       7                  8                  9                10               11                    12
    Print #RptHandle, PPayRec.DiscAmt; dlm; TotalPaid; dlm; PPayRec.Change; dlm; Operator; dlm; PCheckCnt; dlm; POPAmt; dlm;
    '                  13           14              15             16            17             18
    Print #RptHandle, "P"; dlm; NumOfPRecs; dlm; PTotCash; dlm; PTotChrg; dlm; PTotChk; dlm; PTotChng; dlm;
    '                    19                20             21             22
    Print #RptHandle, GTotChkCnt; dlm; GTotCount; dlm; GTotDisc; dlm; PTotalPaid
MoveOnP:
  Next x
  
  Close RptHandle
'  Close RPayHandle
  
  RYearCnt = 0
  ReDim RYears(1 To 1) As Integer
  PYearCnt = 0
  ReDim PYears(1 To 1) As Integer
  
  OpenRealPayListFile RTHandle, OperNum
  NumOfRTRecs = LOF(RTHandle) / Len(RTempRPayRec)
  RYearCnt = 0
  For x = 1 To NumOfRTRecs
    Get RTHandle, x, RTempRPayRec
    If RTempRPayRec.PrevListRec < 0 Then GoTo MoveOnRL
    ThisYear = RTempRPayRec.TaxYear
    For y = 1 To RYearCnt
      If y <> x Then
        If RTempRPayRec.TaxYear = RYears(y) Then
          Exit For
        End If
      End If
    Next y
    If y > RYearCnt Then
      RYearCnt = RYearCnt + 1
      ReDim Preserve RYears(1 To RYearCnt) As Integer
      RYears(RYearCnt) = ThisYear
    End If
MoveOnRL:
  Next x
  
  OpenPersPayListFile PTHandle, OperNum
  NumOfPTRecs = LOF(PTHandle) / Len(PTempRPayRec)
  PYearCnt = 0
  For x = 1 To NumOfPTRecs
    Get PTHandle, x, PTempRPayRec
    If PTempRPayRec.PrevListRec < 0 Then GoTo MoveOnPL:
    ThisYear = PTempRPayRec.TaxYear
    For y = 1 To PYearCnt
      If y <> x Then
        If PTempRPayRec.TaxYear = PYears(y) Then
          Exit For
        End If
      End If
    Next y
    If y > PYearCnt Then
      PYearCnt = PYearCnt + 1
      ReDim Preserve PYears(1 To PYearCnt) As Integer
      PYears(PYearCnt) = ThisYear
    End If
MoveOnPL:
  Next x
  
  If RYearCnt = 0 Then
    GoTo NoRs
  End If
  
  ReDim PrincByYr(1 To RYearCnt) As Double
  ReDim RIntByYr(1 To RYearCnt) As Double
  ReDim AdvColByYr(1 To RYearCnt) As Double
  ReDim LateListByYr(1 To RYearCnt) As Double
  ReDim RPenByYr(1 To RYearCnt) As Double
  ReDim Rev1ByYr(1 To RYearCnt) As Double
  ReDim Rev2ByYr(1 To RYearCnt) As Double
  ReDim Rev3ByYr(1 To RYearCnt) As Double
  ReDim RDiscByYr(1 To RYearCnt) As Double
  ReDim RTotPaidByYr(1 To RYearCnt) As Double
  ReDim ROverPayByYr(1 To RYearCnt) As Double
  
  GoSub SortRYears
NoRs:
  If RYearCnt >= PYearCnt Then
    GYearCnt = RYearCnt
  Else
    GYearCnt = PYearCnt
  End If
  
  ReDim GDiscByYr(1 To GYearCnt) As Double
  ReDim GTotPaidByYr(1 To GYearCnt) As Double
  ReDim GOverPayByYr(1 To GYearCnt) As Double
  
  SubRptFile2$ = "TAXRPTS\SubEdPay2.RPT"
  SubRptHandle2 = FreeFile
  Open SubRptFile2$ For Output As #SubRptHandle2
  
  SubRptFile3$ = "TAXRPTS\SubEdPay3.RPT"
  SubRptHandle3 = FreeFile
  Open SubRptFile3$ For Output As #SubRptHandle3
  
  '************************************
  If GTotCount = 0 Then GoSub PersPrint
  '************************************
  
  If ROPAmt > 0 Then
    For x = 1 To NumOfRRecs
      If PrintType = "N" Then
        Get RPayHandle, RCustRArr(x), RPayRec
      Else
        Get RPayHandle, x, RPayRec
      End If
'      Get RPayHandle, RCustRArr(x), RPayRec
      If RPayRec.LastPayRec = 0 Then GoTo Deleted
      If RPayRec.PrePayAmt > 0 Then
        Print #SubRptHandle3, RPayRec.CustAcct; dlm; QPTrim$(RPayRec.CustName); dlm; RPayRec.TotPaid; dlm; RPayRec.PrePayAmt
      End If
Deleted:
    Next x
  End If
  
  GPrinc = 0
  RGInt = 0
  GAdvCol = 0
  GLateList = 0
  GRPenalty = 0
  GRev1 = 0
  GRev2 = 0
  GRev3 = 0
  RGTot = 0
  RGDisc = 0
  RGOverPay = 0
  
  For x = 1 To NumOfRTRecs
    Get RTHandle, x, RTempRPayRec
       If RTempRPayRec.PrevListRec < 0 Then GoTo NextRL
       RTempRPayRec.CustRec = RTempRPayRec.CustRec
       GPrinc = OldRound(GPrinc + RTempRPayRec.Principle1)
       RGTot = OldRound(RGTot + RTempRPayRec.Principle1)
       RGInt = OldRound(RGInt + RTempRPayRec.Interest1)
       RGTot = OldRound(RGTot + RTempRPayRec.Interest1)
       GAdvCol = OldRound(GAdvCol + RTempRPayRec.Collection)
       RGTot = OldRound(RGTot + RTempRPayRec.Collection)
       GLateList = OldRound(GLateList + RTempRPayRec.LateList)
       RGTot = OldRound(RGTot + RTempRPayRec.LateList)
       GRPenalty = OldRound(GRPenalty + RTempRPayRec.Penalty)
       RGTot = OldRound(RGTot + RTempRPayRec.Penalty)
       GRev1 = OldRound(GRev1 + RTempRPayRec.OptRev1)
       RGTot = OldRound(RGTot + RTempRPayRec.OptRev1)
       GRev2 = OldRound(GRev2 + RTempRPayRec.OptRev2)
       RGTot = OldRound(RGTot + RTempRPayRec.OptRev2)
       GRev3 = OldRound(GRev3 + RTempRPayRec.OptRev3)
       RGTot = OldRound(RGTot + RTempRPayRec.OptRev3)
       RGDisc = OldRound(RGDisc + RTempRPayRec.DiscAmt)
       RGTot = OldRound(RGTot + RTempRPayRec.DiscAmt)
       RGOverPay = OldRound(RGOverPay + RTempRPayRec.PrePayAmt)
       RGTot = OldRound(RGTot + RTempRPayRec.PrePayAmt)
       For y = 1 To RYearCnt
         If RYears(y) = RTempRPayRec.TaxYear Then
           PrincByYr(y) = OldRound(PrincByYr(y) + RTempRPayRec.Principle1)
           RIntByYr(y) = OldRound(RIntByYr(y) + RTempRPayRec.Interest1)
           AdvColByYr(y) = OldRound(AdvColByYr(y) + RTempRPayRec.Collection)
           LateListByYr(y) = OldRound(LateListByYr(y) + RTempRPayRec.LateList)
           RPenByYr(y) = OldRound(RPenByYr(y) + RTempRPayRec.Penalty)
           Rev1ByYr(y) = OldRound(Rev1ByYr(y) + RTempRPayRec.OptRev1)
           Rev2ByYr(y) = OldRound(Rev2ByYr(y) + RTempRPayRec.OptRev2)
           Rev3ByYr(y) = OldRound(Rev3ByYr(y) + RTempRPayRec.OptRev3)
           RDiscByYr(y) = OldRound(RDiscByYr(y) + RTempRPayRec.DiscAmt)
           RTotPaidByYr(y) = OldRound(RTotPaidByYr(y) + RTempRPayRec.TotPaid)
           ROverPayByYr(y) = OldRound(ROverPayByYr(y) + RTempRPayRec.PrePayAmt)
           GDiscByYr(y) = OldRound(GDiscByYr(y) + RDiscByYr(y))
           GTotPaidByYr(y) = OldRound(GTotPaidByYr(y) + RTotPaidByYr(y))
           GOverPayByYr(y) = OldRound(GOverPayByYr(y) + ROverPayByYr(y))
           Exit For
         End If
       Next y
NextRL:
  Next x

  '                        0                             1             2              3
  Print #SubRptHandle1, OldRound(GPrinc + RGDisc); dlm; RGInt; dlm; GAdvCol; dlm; GLateList; dlm;
  '                       4           5           6           7            8           9             10
  Print #SubRptHandle1, GRev1; dlm; GRev2; dlm; GRev3; dlm; RGTot; dlm; GPrinc; dlm; RGDisc; dlm; RGOverPay; dlm;
  '                                11                                   12                                    13
  Print #SubRptHandle1, QPTrim$(TaxMasterRec.OptRev1); dlm; QPTrim$(TaxMasterRec.OptRev2); dlm; QPTrim$(TaxMasterRec.OptRev3); dlm;
  '                      14          15
  Print #SubRptHandle1, "R"; dlm; GRPenalty; dlm; ""; dlm; ""; dlm; ""
  Close RTHandle
  
  Done = False
  For x = 1 To RYearCnt
    If x = RYearCnt Then Done = True
    PrincByYrTot = OldRound(PrincByYrTot + PrincByYr(x) + RDiscByYr(x))
    RIntByYrTot = OldRound(RIntByYrTot + RIntByYr(x))
    AdvColByYrTot = OldRound(AdvColByYrTot + AdvColByYr(x))
    LateListByYrTot = OldRound(LateListByYrTot + LateListByYr(x))
    RPenByYrTot = OldRound(RPenByYrTot + RPenByYr(x))
    Rev1ByYrTot = OldRound(Rev1ByYrTot + Rev1ByYr(x))
    Rev2ByYrTot = OldRound(Rev2ByYrTot + Rev2ByYr(x))
    Rev3ByYrTot = OldRound(Rev3ByYrTot + Rev3ByYr(x))
    RDiscByYrTot = OldRound(RDiscByYrTot + RDiscByYr(x))
    ROverPayByYrTot = OldRound(ROverPayByYrTot + ROverPayByYr(x))
    RTotPaidByYrTot = OldRound(RTotPaidByYrTot + RTotPaidByYr(x))
    '                        0                                 1                       2                 3
    Print #SubRptHandle2, RYears(x); dlm; OldRound(PrincByYr(x) + RDiscByYr(x)); dlm; RIntByYr(x); dlm; AdvColByYr(x); dlm;
    '                           4                   5                 6                 7
    Print #SubRptHandle2, LateListByYr(x); dlm; Rev1ByYr(x); dlm; Rev2ByYr(x); dlm; Rev3ByYr(x); dlm;
    '                          8                  9                10                 11
    Print #SubRptHandle2, RDiscByYr(x); dlm; PrincByYrTot; dlm; RIntByYrTot; dlm; AdvColByYrTot; dlm;
    '                           12                  13                14                 15
    Print #SubRptHandle2, LateListByYrTot; dlm; Rev1ByYrTot; dlm; Rev2ByYrTot; dlm; Rev3ByYrTot; dlm;
    '                         16                        17                               18                                    19
    Print #SubRptHandle2, RDiscByYrTot; dlm; QPTrim$(TaxMasterRec.OptRev1); dlm; QPTrim$(TaxMasterRec.OptRev2); dlm; QPTrim$(TaxMasterRec.OptRev3); dlm;
    '                      20             21                    22                23           24                 25
    Print #SubRptHandle2, Done; dlm; ROverPayByYr(x); dlm; ROverPayByYrTot; dlm; "R"; dlm; RPenByYr(x); dlm; RPenByYrTot; dlm;
    '                     26       27       28       29       30       31       32       33       34
    Print #SubRptHandle2, ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""
  Next x
  
  '************************************
  If PTotalPaid > 0 Then GoSub PersPrint
  '************************************
  
  Close
  
  arVATaxPayTransRpt.Show
  
  Exit Sub
  
PersPrint:
  
  ReDim PersByYr(1 To PYearCnt) As Double
  ReDim PIntByYr(1 To PYearCnt) As Double
  ReDim MTByYr(1 To PYearCnt) As Double
  ReDim MCByYr(1 To PYearCnt) As Double
  ReDim FEByYr(1 To PYearCnt) As Double
  ReDim MHByYr(1 To PYearCnt) As Double
  ReDim PenByYr(1 To PYearCnt) As Double
  ReDim OPt1ByYr(1 To PYearCnt) As Double
  ReDim OPt2ByYr(1 To PYearCnt) As Double
  ReDim OPt3ByYr(1 To PYearCnt) As Double
  ReDim PDiscByYr(1 To PYearCnt) As Double
  ReDim PTotPaidByYr(1 To PYearCnt) As Double
  ReDim POverPayByYr(1 To PYearCnt) As Double
  GoSub SortPYears
  If POPAmt > 0 Then
    For x = 1 To NumOfPRecs
      If PrintType = "N" Then
        Get PPayHandle, PCustRArr(x), PPayRec
      Else
        Get PPayHandle, x, PPayRec
      End If
'      Get PPayHandle, PCustRArr(x), PPayRec 'temp list
      If PPayRec.LastPayRec = 0 Then GoTo NextPers
      If PPayRec.PrePayAmt > 0 Then
        Print #SubRptHandle3, PPayRec.CustAcct; dlm; QPTrim$(PPayRec.CustName); dlm; PPayRec.TotPaid; dlm; PPayRec.PrePayAmt
      End If
NextPers:
    Next x
  End If
  Close SubRptHandle3
  GPers = 0
  GMachTools = 0
  GMerchCap = 0
  GFarmEq = 0
  GMobHomes = 0
  GPInterest = 0
  GPPenalty = 0
  GPRev1 = 0
  GPRev2 = 0
  GPRev3 = 0
  PGTot = 0
  For x = 1 To NumOfPTRecs
    Get PTHandle, x, PTempRPayRec
      If PTempRPayRec.PrevListRec < 0 Then GoTo NextPersL
      GPers = OldRound(GPers + PTempRPayRec.Personal)
      PGTot = OldRound(PGTot + PTempRPayRec.Personal) '
      PGInt = OldRound(PGInt + PTempRPayRec.Interest)
      GMachTools = OldRound(GMachTools + PTempRPayRec.MachTools)
      PGTot = OldRound(PGTot + PTempRPayRec.MachTools) '
      GMerchCap = OldRound(GMerchCap + PTempRPayRec.MerchCap)
      PGTot = OldRound(PGTot + PTempRPayRec.MerchCap) '
      GFarmEq = OldRound(GFarmEq + PTempRPayRec.FarmEquip)
      PGTot = OldRound(PGTot + PTempRPayRec.FarmEquip) '
      GMobHomes = OldRound(GMobHomes + PTempRPayRec.MobHomes)
      PGTot = OldRound(PGTot + PTempRPayRec.MobHomes) '
      GPInterest = OldRound(GPInterest + PTempRPayRec.Interest)
      PGTot = OldRound(PGTot + PTempRPayRec.Interest) '
      GPPenalty = OldRound(GPPenalty + PTempRPayRec.Penalty)
      PGTot = OldRound(PGTot + PTempRPayRec.Penalty) '
      GPRev1 = OldRound(GPRev1 + PTempRPayRec.Opt1)
      PGTot = OldRound(PGTot + PTempRPayRec.Opt1) '
      GPRev2 = OldRound(GPRev2 + PTempRPayRec.Opt2)
      PGTot = OldRound(PGTot + PTempRPayRec.Opt2) '
      GPRev3 = OldRound(GPRev3 + PTempRPayRec.Opt3) '
      PGTot = OldRound(PGTot + PTempRPayRec.Opt3)
      PGDisc = OldRound(PGDisc + PTempRPayRec.DiscAmt)
      PGTot = OldRound(PGTot + PTempRPayRec.DiscAmt) '
      PGOverPay = OldRound(PGOverPay + PTempRPayRec.PrePayAmt)
      PGTot = OldRound(PGTot + PTempRPayRec.PrePayAmt) '
      For y = 1 To PYearCnt
        If PYears(y) = PTempRPayRec.TaxYear Then
          PTempRPayRec.DiscAmt = PTempRPayRec.DiscAmt
          PersByYr(y) = OldRound(PersByYr(y) + PTempRPayRec.Personal)
          PIntByYr(y) = OldRound(PIntByYr(y) + PTempRPayRec.Interest)
          MTByYr(y) = OldRound(MTByYr(y) + PTempRPayRec.MachTools)
          MCByYr(y) = OldRound(MCByYr(y) + PTempRPayRec.MerchCap)
          FEByYr(y) = OldRound(FEByYr(y) + PTempRPayRec.FarmEquip)
          MHByYr(y) = OldRound(MHByYr(y) + PTempRPayRec.MobHomes)
          PenByYr(y) = OldRound(PenByYr(y) + PTempRPayRec.Penalty)
          OPt1ByYr(y) = OldRound(OPt1ByYr(y) + PTempRPayRec.Opt1)
          OPt2ByYr(y) = OldRound(OPt2ByYr(y) + PTempRPayRec.Opt2)
          OPt3ByYr(y) = OldRound(OPt3ByYr(y) + PTempRPayRec.Opt3)
          PDiscByYr(y) = OldRound(PDiscByYr(y) + PTempRPayRec.DiscAmt)
          PTotPaidByYr(y) = OldRound(PTotPaidByYr(y) + PTempRPayRec.TotPaid)
          POverPayByYr(y) = OldRound(POverPayByYr(y) + PTempRPayRec.PrePayAmt)
          GDiscByYr(y) = OldRound(GDiscByYr(y) + PDiscByYr(y))
          GTotPaidByYr(y) = OldRound(GTotPaidByYr(y) + PTotPaidByYr(y))
          GOverPayByYr(y) = OldRound(GOverPayByYr(y) + POverPayByYr(y))
          Exit For
        End If
     Next y
NextPersL:
  Next x

  '                       0             1               2              3
  Print #SubRptHandle1, GPers; dlm; GMachTools; dlm; GMerchCap; dlm; GFarmEq; dlm;
  '                        4              5              6            7           8           9             10
  Print #SubRptHandle1, GMobHomes; dlm; PGInt; dlm; GPPenalty; dlm; PGTot; dlm; GPers; dlm; PGDisc; dlm; PGOverPay; dlm;
  '                        11             12             13          14           15            16           17           18
  Print #SubRptHandle1, OptDesc1; dlm; OptDesc2; dlm; OptDesc3; dlm; "P"; dlm; GPPenalty; dlm; GPRev1; dlm; GPRev2; dlm; GPRev3
  
  
  Done = False
  'now personal
  For x = 1 To PYearCnt
    If x = PYearCnt Then Done = True
    PersByYrTot = OldRound(PersByYrTot + PersByYr(x) + PDiscByYr(x))
    PIntByYrTot = OldRound(PIntByYrTot + PIntByYr(x))
    MTByYrTot = OldRound(MTByYrTot + MTByYr(x))
    MCByYrTot = OldRound(MCByYrTot + MCByYr(x))
    FEByYrTot = OldRound(FEByYrTot + FEByYr(x))
    MHByYrTot = OldRound(MHByYrTot + MHByYr(x))
    PPenByYrTot = OldRound(PPenByYrTot + PenByYr(x))
    POpt1ByYrTot = OldRound(POpt1ByYrTot + OPt1ByYr(x))
    POpt2ByYrTot = OldRound(POpt2ByYrTot + OPt2ByYr(x))
    POpt3ByYrTot = OldRound(POpt3ByYrTot + OPt3ByYr(x))
    PDiscByYrTot = OldRound(PDiscByYrTot + PDiscByYr(x))
    POverPayByYrTot = OldRound(POverPayByYrTot + POverPayByYr(x))
    PTotPaidByYrTot = OldRound(PTotPaidByYrTot + PTotPaidByYr(x))
    '                        0                1                 2               3
    Print #SubRptHandle2, PYears(x); dlm; OldRound(PersByYr(x) + PDiscByYr(x)); dlm; MTByYr(x); dlm; MCByYr(x); dlm;
    '                        4               5               6              7
    Print #SubRptHandle2, FEByYr(x); dlm; MHByYr(x); dlm; PIntByYr(x); dlm; 0; dlm;
    '                          8                  9                10             11
    Print #SubRptHandle2, PDiscByYr(x); dlm; PersByYrTot; dlm; MTByYrTot; dlm; MCByYrTot; dlm;
    '                        12              13              14             15
    Print #SubRptHandle2, FEByYrTot; dlm; MHByYrTot; dlm; PIntByYrTot; dlm; 0; dlm;
    '                         16                 17                18              19
    Print #SubRptHandle2, PDiscByYrTot; dlm; "Mob Homes"; dlm; "Interest"; dlm; ""; dlm;
    '                      20             21                    22                23           24                25
    Print #SubRptHandle2, Done; dlm; POverPayByYr(x); dlm; POverPayByYrTot; dlm; "P"; dlm; PenByYr(x); dlm; PPenByYrTot; dlm;
    '                         26                27                28                 29
    Print #SubRptHandle2, OPt1ByYr(x); dlm; OPt2ByYr(x); dlm; OPt3ByYr(x); dlm; POpt1ByYrTot; dlm;
    '                          30                 31                32              33              34
    Print #SubRptHandle2, POpt2ByYrTot; dlm; POpt3ByYrTot; dlm; OptDesc1$; dlm; OptDesc2$; dlm; OptDesc3$
  
  Next x
  
  Return
  
SortRYears:
  
  LilYear = 1900
  Nextx = 1
  Do
    For x = Nextx To RYearCnt
      If RYears(x) > LilYear Then
        LilYear = RYears(x)
        Thisx = x
      End If
    Next x
    HoldYears = RYears(Nextx)
    HoldPrincByYr = PrincByYr(Nextx)
    HoldRIntByYr = RIntByYr(Nextx)
    HoldAdvColByYr = AdvColByYr(Nextx)
    HoldLateListByYr = LateListByYr(Nextx)
    HoldRPenByYr = RPenByYr(Nextx)
    HoldRev1ByYr = Rev1ByYr(Nextx)
    HoldRev2ByYr = Rev2ByYr(Nextx)
    HoldRev3ByYr = Rev3ByYr(Nextx)
    HoldRDiscByYr = RDiscByYr(Nextx)
    HoldRTotPaidByYr = RTotPaidByYr(Nextx)
    RYears(Nextx) = RYears(Thisx)
    PrincByYr(Nextx) = PrincByYr(Thisx)
    RIntByYr(Nextx) = RIntByYr(Thisx)
    AdvColByYr(Nextx) = AdvColByYr(Thisx)
    LateListByYr(Nextx) = LateListByYr(Thisx)
    RPenByYr(Nextx) = RPenByYr(Thisx)
    Rev1ByYr(Nextx) = Rev1ByYr(Thisx)
    Rev2ByYr(Nextx) = Rev2ByYr(Thisx)
    Rev3ByYr(Nextx) = Rev3ByYr(Thisx)
    RDiscByYr(Nextx) = RDiscByYr(Thisx)
    RTotPaidByYr(Nextx) = RTotPaidByYr(Thisx)
    RYears(Thisx) = HoldYears
    PrincByYr(Thisx) = HoldPrincByYr
    RIntByYr(Thisx) = HoldRIntByYr
    AdvColByYr(Thisx) = HoldAdvColByYr
    LateListByYr(Thisx) = HoldLateListByYr
    RPenByYr(Thisx) = HoldRPenByYr
    Rev1ByYr(Thisx) = HoldRev1ByYr
    Rev2ByYr(Thisx) = HoldRev2ByYr
    Rev3ByYr(Thisx) = HoldRev3ByYr
    RDiscByYr(Thisx) = HoldRDiscByYr
    RTotPaidByYr(Thisx) = HoldRTotPaidByYr
    LilYear = 1900
    Nextx = Nextx + 1
    If Nextx > RYearCnt Then Exit Do
  Loop
  
  Return
  
SortPYears:
  'now personal
  LilYear = 1900
  Nextx = 1
  Do
    For x = Nextx To PYearCnt
      If PYears(x) > LilYear Then
        LilYear = PYears(x)
        Thisx = x
      End If
    Next x
    HoldYears = PYears(Nextx)
    HoldPersByYr = PersByYr(Nextx)
    HoldMTByYr = MTByYr(Nextx)
    HoldMCByYr = MCByYr(Nextx)
    HoldFEByYr = FEByYr(Nextx)
    HoldMHByYr = MHByYr(Nextx)
    HoldPIntByYr = PIntByYr(Nextx)
    HoldPPenByYr = PenByYr(Nextx)
    HoldPDiscByYr = PDiscByYr(Nextx)
    HoldPTotPaidByYr = PTotPaidByYr(Nextx)
    PYears(Nextx) = PYears(Thisx)
    PersByYr(Nextx) = PersByYr(Thisx)
    MTByYr(Nextx) = MTByYr(Thisx)
    MCByYr(Nextx) = MCByYr(Thisx)
    FEByYr(Nextx) = FEByYr(Thisx)
    MHByYr(Nextx) = MHByYr(Thisx)
    PIntByYr(Nextx) = PIntByYr(Thisx)
    PenByYr(Nextx) = PenByYr(Thisx)
    PDiscByYr(Nextx) = PDiscByYr(Thisx)
    PTotPaidByYr(Nextx) = PTotPaidByYr(Thisx)
    PYears(Thisx) = HoldYears
    PersByYr(Thisx) = HoldPersByYr
    MTByYr(Thisx) = HoldMTByYr
    MCByYr(Thisx) = HoldMCByYr
    FEByYr(Thisx) = HoldFEByYr
    MHByYr(Thisx) = HoldMHByYr
    PIntByYr(Thisx) = HoldPIntByYr
    PenByYr(Thisx) = HoldPPenByYr
    PDiscByYr(Thisx) = HoldPDiscByYr
    PTotPaidByYr(Thisx) = HoldPTotPaidByYr
    LilYear = 1900
    Nextx = Nextx + 1
    If Nextx > PYearCnt Then Exit Do
  Loop

  Return
  
SortRCustomers:
  OpenTaxCustFile TCHandle, NumOfTCRecs
  Nextx = 0
  For x = 1 To NumOfRRecs
    Get RPayHandle, x, RPayRec
    If RPayRec.LastPayRec = 0 Then GoTo MoveALongR
    Get TCHandle, RPayRec.CustAcct, TaxCust
    Nextx = Nextx + 1
'    RCustNArr(Nextx) = QPTrim$(RPayRec.CustName)
    RCustNArr(Nextx) = QPTrim$(TaxCust.SName)
    RCustRArr(Nextx) = x
MoveALongR:
  Next x
  Close TCHandle
  NumOfRRecs = Nextx
  If NumOfRRecs = 1 Then Return
  
  BigName$ = ""
  For x = 1 To NumOfRRecs
    If RCustNArr(x) > BigName Then
      BigName = RCustNArr(x)
    End If
  Next x
  
  LilName = BigName + "z"
  NextOne = 1
  
  Do
    For x = NextOne To NumOfRRecs
      If RCustNArr(x) < LilName Then
        LilName = RCustNArr(x)
        Thisx = x
      End If
    Next x
    HoldName = RCustNArr(NextOne)
    HoldRec = RCustRArr(NextOne)
    RCustNArr(NextOne) = RCustNArr(Thisx)
    RCustRArr(NextOne) = RCustRArr(Thisx)
    RCustNArr(Thisx) = HoldName
    RCustRArr(Thisx) = HoldRec
    NextOne = NextOne + 1
    LilName = BigName + "z"
    If NextOne > NumOfRRecs Then Exit Do
  Loop
  
  Return

SortPCustomers:
  OpenTaxCustFile TCHandle, NumOfTCRecs
  Nextx = 0
  For x = 1 To NumOfPRecs
    Get PPayHandle, x, PPayRec
    If PPayRec.LastPayRec = 0 Then GoTo MoveALongP
    Get TCHandle, PPayRec.CustAcct, TaxCust
    Nextx = Nextx + 1
'    PCustNArr(Nextx) = QPTrim$(PPayRec.CustName)
    PCustNArr(Nextx) = QPTrim$(TaxCust.SName)
    PCustRArr(Nextx) = x
MoveALongP:
  Next x
  Close TCHandle
  NumOfPRecs = Nextx
  If NumOfPRecs = 1 Then Return
  
  BigName$ = ""
  For x = 1 To NumOfPRecs
    If PCustNArr(x) > BigName Then
      BigName = PCustNArr(x)
    End If
  Next x
  
  LilName = BigName + "z"
  NextOne = 1
  
  Do
    For x = NextOne To NumOfPRecs
      If PCustNArr(x) < LilName Then
        LilName = PCustNArr(x)
        Thisx = x
      End If
    Next x
    HoldName = PCustNArr(NextOne)
    HoldRec = PCustRArr(NextOne)
    PCustNArr(NextOne) = PCustNArr(Thisx)
    PCustRArr(NextOne) = PCustRArr(Thisx)
    PCustNArr(Thisx) = HoldName
    PCustRArr(Thisx) = HoldRec
    NextOne = NextOne + 1
    LilName = BigName + "z"
    If NextOne > NumOfPRecs Then Exit Do
  Loop
  Return
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPayEditList", "PrintGraphics", Erl)
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%E"
      Call cmdExit_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%d"
      Call cmdEdit_Click
      KeyCode = 0
    Case vbKeyF5:
      SendKeys "%P"
      Call cmdPrint_Click
      KeyCode = 0
    Case vbKeyReturn:
      If fpListRPay.SelCount > 0 Then
        Call RealEdit
      ElseIf fpListPPay.SelCount > 0 Then
        Call PersEdit
      End If
      KeyCode = 0
    Case Else:
  End Select

End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Me.HelpContextID = hlpPrintTransaction
  Call LoadMe
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      KillFile "C:\CPWork\editpyment.dat"
      MainLog ("CitiTaxes.exe terminated via menu bar on frmVATaxPayEditList.")
      Call Terminate
      End
    End If
  End If

End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    'Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
    DoEvents
  End If
End Sub

Private Sub LoadMe()
  Dim PayRec As TaxPaymentRecType
  Dim RPayHandle As Integer
  Dim NumOfRRecs As Integer
  Dim PPayHandle As Integer
  Dim NumOfPRecs As Integer
  Dim x As Integer
  Dim OHandle As Integer
  Dim NumOfOperRecs As Integer
  Dim OperRec As CitiPassType
  Dim NumOfOperPayRecs As Integer
  Dim One As Integer
  Dim ThisFile As Integer
  Dim FileName$
  
  On Error GoTo ERRORSTUFF
  
  FileName = "C:\CPWork\editpyment.dat" 'used when using the transaction history report
  ThisFile = FreeFile
  Open FileName For Output As ThisFile
  One = 1
  Print #ThisFile, One
  Close ThisFile
  
  fptxtOperator.Text = "Operator # " + CStr(OperNum)
  If Exist("TAXRCPR" + CStr(OperNum) + ".DAT") Then
    OpenTempRealPayFile RPayHandle, OperNum
    NumOfRRecs = LOF(RPayHandle) / Len(PayRec)
    If NumOfRRecs = 0 Then
      Label1.Caption = "No Real Transactions"
    Else
      fpListRPay.ListIndex = 0
    End If
  End If
  If Exist("TAXPCPR" + CStr(OperNum) + ".DAT") Then
    OpenTempPersPayFile PPayHandle, OperNum
    NumOfPRecs = LOF(PPayHandle) / Len(PayRec)
    If NumOfPRecs = 0 Then
      Label1.Caption = "No Personal Transactions"
    ElseIf fpListRPay.SelCount = 0 Then
      fpListPPay.ListIndex = 0
    End If
  End If
  For x = 1 To NumOfRRecs
    If RPayHandle > 0 Then
      Get RPayHandle, x, PayRec
      If PayRec.LastPayRec = 0 Then GoTo NextR
      fpListRPay.InsertRow = CStr(PayRec.CustAcct) + Chr(9) + QPTrim$(PayRec.CustName) + Chr(9) + QPTrim$(Using$("$###,##0.00", PayRec.AmtPaid)) + Chr(9) + QPTrim$(Using$("$###,##0.00", OldRound(PayRec.AmtOwed)))
      DoEvents
    End If
NextR:
  Next x
  For x = 1 To NumOfPRecs
    If PPayHandle > 0 Then
      Get PPayHandle, x, PayRec
      If PayRec.LastPayRec = 0 Then GoTo NextP
      fpListPPay.InsertRow = CStr(PayRec.CustAcct) + Chr(9) + QPTrim$(PayRec.CustName) + Chr(9) + QPTrim$(Using$("$###,##0.00", PayRec.AmtPaid)) + Chr(9) + QPTrim$(Using$("$###,##0.00", OldRound(PayRec.AmtOwed)))
      DoEvents
    End If
NextP:
  Next x
  Close
  If fpListRPay.ListCount > 0 Then
    fpListRPay.ListIndex = 0
  ElseIf fpListPPay.ListCount > 0 Then
    fpListPPay.ListIndex = 0
  End If
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPayEditList", "LoadMe", Erl)
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

Private Sub RealEdit()
  Dim PayRec As TaxPaymentRecType
  Dim PayHandle As Integer
  Dim NumOfPRecs As Integer
  Dim x As Integer
  Dim Operator$
  Dim AcctNum As Long
  
  On Error GoTo ERRORSTUFF
  
  If fpListRPay.ListIndex = -1 Then
    Call TaxMsg(900, "Please make a selection from the list.")
    Exit Sub
  End If
  
  fpListRPay.Col = 0
  fpListRPay.Row = fpListRPay.ListIndex
  AcctNum = CLng(QPTrim$(fpListRPay.ColText))
  
  Operator = CStr(OperNum)
  GCustNum = 0
  GPayNum = 0
  
  OpenTempRealPayFile PayHandle, OperNum
  NumOfPRecs = LOF(PayHandle) / Len(PayRec)
  For x = 1 To NumOfPRecs
    Get PayHandle, x, PayRec
    If PayRec.CustAcct = AcctNum Then
      GCustNum = AcctNum
      GPayNum = x
      Exit For
    End If
  Next x
  Close PayHandle
  
  If x > NumOfPRecs Then
    Call TaxMsg(900, "The transaction selected could not be found.")
    Exit Sub
  End If
  
  frmVATaxPaymentEntry.fpLongAcctNum = GCustNum
  frmVATaxPaymentEntry.Show
  DoEvents
  Me.Hide
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPayEditList", "Edit", Erl)
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

Private Sub fpListPPay_Click()
  fpListRPay.Action = ActionDeselectAll
End Sub

Private Sub fpListRPay_Click()
  fpListPPay.Action = ActionDeselectAll
End Sub

Private Sub fpListRPay_DblClick()
  Call RealEdit
End Sub

Private Sub PersEdit()
  Dim PayRec As TaxPaymentRecType
  Dim PayHandle As Integer
  Dim NumOfPRecs As Integer
  Dim x As Integer
  Dim Operator$
  Dim AcctNum As Long
  
  On Error GoTo ERRORSTUFF
  
  If fpListPPay.ListIndex = -1 Then
    Call TaxMsg(900, "Please make a selection from the list.")
    Exit Sub
  End If
  
  fpListPPay.Col = 0
  fpListPPay.Row = fpListPPay.ListIndex
  AcctNum = CLng(QPTrim$(fpListPPay.ColText))
  
  Operator = CStr(OperNum)
  GCustNum = 0
  GPayNum = 0
  OpenTempPersPayFile PayHandle, OperNum
  NumOfPRecs = LOF(PayHandle) / Len(PayRec)
  For x = 1 To NumOfPRecs
    Get PayHandle, x, PayRec
    If PayRec.CustAcct = AcctNum Then
      GCustNum = AcctNum
      GPayNum = x
      Exit For
    End If
  Next x
  Close PayHandle
  
  If x > NumOfPRecs Then
    Call TaxMsg(900, "The transaction selected could not be found.")
    Exit Sub
  End If
  
  frmVATaxPersPaymentEntry.fpLongAcctNum = GCustNum
  frmVATaxPersPaymentEntry.Show
  DoEvents
  Me.Hide
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPayEditList", "Edit", Erl)
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

Private Sub fpListPPay_DblClick()
  Call PersEdit
End Sub

Private Sub PrintText()
  Dim RPayRec As TaxPaymentRecType
  Dim RPayHandle As Integer
  Dim NumOfRRecs As Integer
  Dim PPayRec As TaxPaymentRecType
  Dim PPayHandle As Integer
  Dim NumOfPRecs As Integer
  Dim x As Integer
  Dim OHandle As Integer
  Dim NumOfOperRecs As Integer
  Dim OperRec As CitiPassType
  Dim NumOfOperRPayRecs As Integer
  Dim TaxMasterRec As TaxMasterType
  Dim MHandle As Integer
  Dim Town$
  Dim RptFile$
  Dim RptHandle As Integer
  Dim SubRptFile1$
  Dim SubRptHandle1 As Integer
  Dim SubRptFile2$
  Dim SubRptHandle2 As Integer
  Dim SubRptFile3$
  Dim SubRptHandle3 As Integer
  Dim RTempRPayRec As RealPayListType
  Dim RTHandle As Integer
  Dim NumOfRTRecs As Integer
  Dim PTempRPayRec As PersPayListType
  Dim PTHandle As Integer
  Dim NumOfPTRecs As Integer
  Dim GPrinc As Double
  Dim RGInt As Double
  Dim GAdvCol As Double
  Dim GLateList As Double
  Dim GRPenalty As Double
  Dim GRev1 As Double
  Dim GRev2 As Double
  Dim GRev3 As Double
  Dim RGTot As Double
  Dim GPers As Double
  Dim GMachTools As Double
  Dim GMerchCap As Double
  Dim GFarmEq As Double
  Dim GMobHomes As Double
  Dim GPersonal As Double
  Dim GPPenalty As Double
  Dim PGInt As Double
  Dim GPOptRev1 As Double
  Dim GPOptRev2 As Double
  Dim GPOptRev3 As Double
  Dim GOverPay As Double
  Dim RGOverPay As Double
  Dim PGOverPay As Double
  Dim PGTot As Double
  Dim Operator$
  Dim TotalPaid#
  Dim RTotalPaid#
  Dim PTotalPaid#
  Dim GDisc As Double
  Dim RGDisc As Double
  Dim PGDisc As Double
  Dim RYearCnt As Integer
  Dim PYearCnt As Integer
  Dim GYearCnt As Integer
  Dim y As Integer
  Dim ThisYear As Integer
  Dim PrincByYrTot As Double
  Dim RIntByYrTot As Double
  Dim AdvColByYrTot As Double
  Dim LateListByYrTot As Double
  Dim RPenByYrTot As Double
  Dim Rev1ByYrTot As Double
  Dim Rev2ByYrTot As Double
  Dim Rev3ByYrTot As Double
  Dim PersByYrTot As Double
  Dim MTByYrTot As Double
  Dim MCByYrTot As Double
  Dim FEByYrTot As Double
  Dim MHByYrTot As Double
  Dim PIntByYrTot As Double
  Dim PPenByYrTot As Double
  Dim POpt1ByYrTot As Double
  Dim POpt2ByYrTot As Double
  Dim POpt3ByYrTot As Double
  Dim RDiscByYrTot As Double
  Dim PDiscByYrTot As Double
  Dim GDiscByYrTot As Double
  Dim RTotPaidByYrTot As Double
  Dim PTotPaidByYrTot As Double
  Dim ROverPayByYrTot As Double
  Dim POverPayByYrTot As Double
  Dim CheckCnt As Integer
  Dim RCheckCnt As Integer
  Dim PCheckCnt As Integer
  Dim HoldPrincByYr As Double
  Dim HoldRIntByYr As Double
  Dim HoldAdvColByYr As Double
  Dim HoldLateListByYr As Double
  Dim HoldRPenByYr As Double
  Dim HoldRev1ByYr As Double
  Dim HoldRev2ByYr As Double
  Dim HoldRev3ByYr As Double
  Dim HoldRDiscByYr As Double
  Dim HoldPersByYr As Double
  Dim HoldMTByYr As Double
  Dim HoldMCByYr As Double
  Dim HoldFEByYr As Double
  Dim HoldMHByYr As Double
  Dim HoldPIntByYr As Double
  Dim HoldPenByYr As Double
  Dim HoldROverPayByYr As Double
  Dim HoldRTotPaidByYr As Double
  Dim HoldPOverPayByYr As Double
  Dim HoldPTotPaidByYr As Double
  Dim HoldPDiscByYr As Double
  Dim Thisx As Integer
  Dim LilYear As Integer
  Dim Nextx As Integer
  Dim HoldYears As Integer
  Dim Done As Boolean
  Dim GOPAmt As Double
  Dim ROPAmt As Double
  Dim POPAmt As Double
  Dim BigName$
  Dim LilName$
  Dim HoldName$
  Dim HoldRec As Integer
  Dim NextOne As Integer
  Dim GTotCash As Double
  Dim GTotCheck As Double
  Dim GTotCharge As Double
  Dim GTotChange As Double
  Dim GTotCount As Integer
  Dim GTotDisc As Double
  Dim RTotCash As Double
  Dim RTotCheck As Double
  Dim RTotCharge As Double
  Dim RTotChange As Double
  Dim RTotCount As Integer
  Dim RTotDisc As Double
  Dim PTotCash As Double
  Dim PTotCheck As Double
  Dim PTotCharge As Double
  Dim PTotChange As Double
  Dim PTotCount As Integer
  Dim PTotDisc As Double
  Dim GTotalPaid As Double
  Dim NoD$, UseThis$, DLine$, sLine$
  Dim FF$, MaxLines As Integer, LineCnt As Integer
  Dim RevTrunc1 As String * 12
  Dim RevTrunc2 As String * 12
  Dim RevTrunc3 As String * 12
  Dim GTotCredit As Double
  Dim RTotCredit As Double
  Dim PTotCredit As Double
  Dim OPToOwed As Double
  Dim OPPaid As Double
  Dim OPCnt As Integer
  Dim YTotal As Double
  Dim OperLen As Integer
  Dim ThisTab As Integer
  Dim Page As Integer
  Dim Opt1Desc$
  Dim Opt2Desc$
  Dim Opt3Desc$
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim RReceiptCnt As Integer
  Dim PReceiptCnt As Integer
  
  On Error GoTo ERRORSTUFF
  
  NoD = "###,##0.00"
  UseThis$ = "$###,##0.00"
  DLine$ = String(109, "=")
  sLine$ = String(109, "-")
  FF$ = Chr(12)
  MaxLines = 56
  LineCnt = 0
  
  OpenTaxSetUpFile MHandle
  Get MHandle, 1, TaxMasterRec
  Close MHandle
  Opt1Desc = QPTrim$(TaxMasterRec.POptRev1)
  Opt2Desc = QPTrim$(TaxMasterRec.POptRev2)
  Opt3Desc = QPTrim$(TaxMasterRec.POptRev3)
  Town = QPTrim$(TaxMasterRec.Name)
  
  Operator$ = "Operator # " + CStr(OperNum) + " " + PWUser
  fptxtOperator.Text = "Operator # " + CStr(OperNum)
  OpenTempRealPayFile RPayHandle, OperNum
  NumOfRRecs = LOF(RPayHandle) / Len(RPayRec)
  If NumOfRRecs > 0 Then
    ReDim RCustNArr(1 To NumOfRRecs) As String
    ReDim RCustRArr(1 To NumOfRRecs) As Integer
    If PrintType = "N" Then
      GoSub SortRCustomers
    End If
  End If
  
  OpenTempPersPayFile PPayHandle, OperNum
  NumOfPRecs = LOF(PPayHandle) / Len(PPayRec)
  If NumOfPRecs > 0 Then
    ReDim PCustNArr(1 To NumOfPRecs) As String
    ReDim PCustRArr(1 To NumOfPRecs) As Integer
    If PrintType = "N" Then
      GoSub SortPCustomers
    End If
  End If
  
  RptFile$ = "TAXRPTS\TaxEdPay.RPT"
  RptHandle = FreeFile
  RCheckCnt = 0
  PCheckCnt = 0
  ROPAmt = 0
  POPAmt = 0
  TotalPaid = 0
  RTotCash = 0
  RTotCharge = 0
  RTotCheck = 0
  RTotChange = 0
  RTotDisc = 0
  RTotCount = 0
  RTotCredit = 0
  GTotCredit = 0
  Open RptFile$ For Output As #RptHandle
  GoSub PrintHeader
  
  If NumOfRRecs > 0 Then
    Print #RptHandle, "REAL TRANSACTIONS"
  End If
  RReceiptCnt = 0
  For x = 1 To NumOfRRecs
    If PrintType = "N" Then
      Get RPayHandle, RCustRArr(x), RPayRec
    Else
      Get RPayHandle, x, RPayRec
    End If
    If RPayRec.LastPayRec = 0 Then GoTo MoveOnR1
    If RPayRec.ChkAmt > 0 Then
      RCheckCnt = RCheckCnt + 1
    End If
    TotalPaid = OldRound(RPayRec.CashAmt + RPayRec.ChkAmt + RPayRec.ChrgAmt + RPayRec.DiscAmt - RPayRec.Change)
    GOPAmt = OldRound(GOPAmt + RPayRec.PrePayAmt)
    RTotalPaid = OldRound(RTotalPaid + TotalPaid)
    ROPAmt = OldRound(ROPAmt + RPayRec.PrePayAmt)
    GTotalPaid = OldRound(GTotalPaid + TotalPaid)
    GTotCash = OldRound(GTotCash + RPayRec.CashAmt)
    RTotCash = OldRound(RTotCash + RPayRec.CashAmt)
    GTotCharge = OldRound(GTotCharge + RPayRec.ChrgAmt)
    RTotCharge = OldRound(RTotCharge + RPayRec.ChrgAmt)
    GTotCheck = OldRound(GTotCheck + RPayRec.ChkAmt)
    RTotCheck = OldRound(RTotCheck + RPayRec.ChkAmt)
    GTotChange = OldRound(GTotChange + RPayRec.Change)
    RTotChange = OldRound(RTotChange + RPayRec.Change)
    GTotDisc = OldRound(GTotDisc + RPayRec.DiscAmt)
    RTotDisc = OldRound(RTotDisc + RPayRec.DiscAmt)
    GTotCount = GTotCount + 1
    RTotCount = RTotCount + 1
    Print #RptHandle, MakeRegDate(RPayRec.PayDate); Tab(16); CStr(RPayRec.CustAcct); Tab(26);
    Print #RptHandle, QPTrim$(RPayRec.CustName); Tab(50); Using(NoD, RPayRec.CashAmt); Tab(60); Using(NoD, RPayRec.ChkAmt); Tab(70); Using(NoD, RPayRec.ChrgAmt); Tab(80);
    Print #RptHandle, Using(NoD, RPayRec.DiscAmt); Tab(90); Using(NoD, TotalPaid); Tab(100); Using(NoD, RPayRec.Change)
    RReceiptCnt = RReceiptCnt + 1
    LineCnt = LineCnt + 1
    If x >= NumOfRRecs - 6 Then
      If LineCnt > MaxLines - 6 Then
        Print #RptHandle, FF$
        GoSub PrintHeader
      End If
    Else
      If LineCnt > MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintHeader
      End If
    End If
MoveOnR1:
  Next x
  Print #RptHandle, sLine
  Print #RptHandle, Tab(21); "Real Totals"; Tab(50); Using(NoD, RTotCash); Tab(60); Using(NoD, RTotCheck); Tab(70); Using(NoD, RTotCharge); Tab(80); Using(NoD, RTotDisc); Tab(90); Using(NoD, RTotalPaid); Tab(100); Using(NoD, RTotChange)
  Print #RptHandle, "Total Number of Real Receipts: " + CStr(RReceiptCnt)
  Print #RptHandle, "Total Number of Real Checks: " + CStr(RCheckCnt)
  Print #RptHandle,
  LineCnt = LineCnt + 5
  '---------------------------------------------------------------------------------------
  If ROPAmt > 0 Then
    If LineCnt >= MaxLines - 7 Then
      Print #RptHandle, FF$
      GoSub PrintOPHeader
      GoTo NewPage
    End If
'    Print #RptHandle,
    Print #RptHandle,
    Print #RptHandle, String(89, "-")
    Print #RptHandle, "Real Over Payment Summary"
    Print #RptHandle, Tab(38); "Payment Applied To"; Tab(60); "Over Payment"; Tab(78); "Total Amount"
    Print #RptHandle, "Cust Num"; Tab(12); "Customer"; Tab(45); "Amount Owed"; Tab(66); "Amount"; Tab(86); "Paid"
    Print #RptHandle, String(89, "-")
    LineCnt = LineCnt + 6
NewPage:
    OPToOwed = 0
    OPPaid = 0
    OPCnt = 0
    For x = 1 To NumOfRRecs
      If PrintType = "N" Then
        Get RPayHandle, RCustRArr(x), RPayRec
      Else
        Get RPayHandle, x, RPayRec
      End If
'      Get RPayHandle, RCustRArr(x), RPayRec
      If RPayRec.LastPayRec = 0 Then GoTo MoveOnR2
      If RPayRec.PrePayAmt > 0 Then
        OPToOwed = OldRound(OPToOwed + RPayRec.TotPaid)
        OPPaid = OldRound(OPPaid + RPayRec.PrePayAmt)
        OPCnt = OPCnt + 1
        Print #RptHandle, RPayRec.CustAcct; Tab(12); QPTrim$(RPayRec.CustName); Tab(45); Using(UseThis, RPayRec.TotPaid); Tab(61); Using(UseThis, RPayRec.PrePayAmt); Tab(79); Using(UseThis, OldRound(RPayRec.TotPaid + RPayRec.PrePayAmt))
        LineCnt = LineCnt + 1
      End If
      If LineCnt >= MaxLines - 7 Then
        Print #RptHandle, FF$
        GoSub PrintOPHeader
      End If
MoveOnR2:
    Next x
    If LineCnt >= MaxLines - 7 Then
      Print #RptHandle, FF$
      GoSub PrintOPHeader
    End If
    If LineCnt >= MaxLines - 2 Then
      Print #RptHandle, FF$
      GoSub PrintOPHeader
    End If
    Print #RptHandle, String(89, "-")
    Print #RptHandle, "Totals"; Tab(12); "# Over Payments: " + CStr(OPCnt); Tab(45); Using(UseThis, OPToOwed); Tab(61); Using(UseThis, OPPaid); Tab(79); Using(UseThis, OldRound(OPToOwed + OPPaid))
    Print #RptHandle,
    Print #RptHandle,
    LineCnt = LineCnt + 4
  Else
    Print #RptHandle, "No Real Over Payment Activity"
    Print #RptHandle,
    LineCnt = LineCnt + 2
  End If
  
  '---------------------------------------------------------------------------------------
  Close RPayHandle
  PTotCash = 0
  PTotCharge = 0
  PTotCheck = 0
  PTotChange = 0
  PTotDisc = 0
  PTotCount = 0
  
  If NumOfPRecs > 0 Then
    Print #RptHandle, "PERSONAL TRANSACTIONS"
  End If
  PReceiptCnt = 0
  For x = 1 To NumOfPRecs
    If PrintType = "N" Then
      Get PPayHandle, PCustRArr(x), PPayRec
    Else
      Get PPayHandle, x, PPayRec
    End If
    If PPayRec.LastPayRec = 0 Then GoTo MoveOnP1
    If PPayRec.ChkAmt > 0 Then
      PCheckCnt = PCheckCnt + 1
    End If
    TotalPaid = OldRound(PPayRec.CashAmt + PPayRec.ChkAmt + PPayRec.ChrgAmt + PPayRec.DiscAmt - PPayRec.Change)
    GOPAmt = OldRound(GOPAmt + PPayRec.PrePayAmt)
    GTotalPaid = OldRound(GTotalPaid + TotalPaid)
    PTotalPaid = OldRound(PTotalPaid + TotalPaid)
    POPAmt = OldRound(POPAmt + PPayRec.PrePayAmt)
    GTotCash = OldRound(GTotCash + PPayRec.CashAmt)
    PTotCash = OldRound(PTotCash + PPayRec.CashAmt)
    GTotCharge = OldRound(GTotCharge + PPayRec.ChrgAmt)
    PTotCharge = OldRound(PTotCharge + PPayRec.ChrgAmt)
    GTotCheck = OldRound(GTotCheck + PPayRec.ChkAmt)
    PTotCheck = OldRound(PTotCheck + PPayRec.ChkAmt)
    GTotChange = OldRound(GTotChange + PPayRec.Change)
    PTotChange = OldRound(PTotChange + PPayRec.Change)
    GTotDisc = OldRound(GTotDisc + PPayRec.DiscAmt)
    PTotDisc = OldRound(PTotDisc + PPayRec.DiscAmt)
    GTotCount = GTotCount + 1
    PTotCount = PTotCount + 1
    Print #RptHandle, MakeRegDate(PPayRec.PayDate); Tab(16); CStr(PPayRec.CustAcct); Tab(26);
    Print #RptHandle, QPTrim$(PPayRec.CustName); Tab(50); Using(NoD, PPayRec.CashAmt); Tab(60); Using(NoD, PPayRec.ChkAmt); Tab(70); Using(NoD, PPayRec.ChrgAmt); Tab(80);
    Print #RptHandle, Using(NoD, PPayRec.DiscAmt); Tab(90); Using(NoD, TotalPaid); Tab(100); Using(NoD, PPayRec.Change)
    PReceiptCnt = PReceiptCnt + 1
    LineCnt = LineCnt + 1
    If x >= NumOfRRecs - 6 Then
      If LineCnt > MaxLines - 6 Then
        Print #RptHandle, FF$
        GoSub PrintHeader
      End If
    Else
      If LineCnt > MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintHeader
      End If
    End If
MoveOnP1:
  Next x
  
  Print #RptHandle, sLine
  Print #RptHandle, Tab(21); "Personal Totals"; Tab(50); Using(NoD, PTotCash); Tab(60); Using(NoD, PTotCheck); Tab(70); Using(NoD, PTotCharge); Tab(80); Using(NoD, PTotDisc); Tab(90); Using(NoD, PTotalPaid); Tab(100); Using(NoD, PTotChange)
  Print #RptHandle, "Total Number of Personal Receipts: " + CStr(PReceiptCnt)
  Print #RptHandle, "Total Number of Personal Checks: " + CStr(PCheckCnt)
  LineCnt = LineCnt + 4
  
  '---------------------------------------------------------------------------------------
  If POPAmt > 0 Then
     If LineCnt >= MaxLines - 7 Then
       Print #RptHandle, FF$
       GoSub PrintOPHeader
       GoTo NewPageP
     End If
     Print #RptHandle,
     Print #RptHandle,
     Print #RptHandle, String(89, "-")
     Print #RptHandle, "Personal Over Payment Summary"
     Print #RptHandle, Tab(38); "Payment Applied To"; Tab(60); "Over Payment"; Tab(78); "Total Amount"
     Print #RptHandle, "Cust Num"; Tab(12); "Customer"; Tab(45); "Amount Owed"; Tab(66); "Amount"; Tab(86); "Paid"
     Print #RptHandle, String(89, "-")
     LineCnt = LineCnt + 7
NewPageP:
     OPToOwed = 0
     OPPaid = 0
     OPCnt = 0
     For x = 1 To NumOfPRecs
       If PrintType = "N" Then
         Get PPayHandle, PCustRArr(x), PPayRec
       Else
         Get PPayHandle, x, PPayRec
       End If
'       Get PPayHandle, PCustRArr(x), PPayRec
       If PPayRec.LastPayRec = 0 Then GoTo MoveOnP2
       If PPayRec.PrePayAmt > 0 Then
         OPToOwed = OldRound(OPToOwed + PPayRec.TotPaid)
         OPPaid = OldRound(OPPaid + PPayRec.PrePayAmt)
         OPCnt = OPCnt + 1
         Print #RptHandle, PPayRec.CustAcct; Tab(12); QPTrim$(PPayRec.CustName); Tab(45); Using(UseThis, PPayRec.TotPaid); Tab(61); Using(UseThis, PPayRec.PrePayAmt); Tab(79); Using(UseThis, OldRound(PPayRec.TotPaid + PPayRec.PrePayAmt))
         LineCnt = LineCnt + 1
       End If
       If LineCnt >= MaxLines - 7 Then
         Print #RptHandle, FF$
         GoSub PrintOPHeader
       End If
MoveOnP2:
     Next x
     If LineCnt >= MaxLines - 7 Then
       Print #RptHandle, FF$
       GoSub PrintOPHeader
     End If
     If LineCnt >= MaxLines - 2 Then
       Print #RptHandle, FF$
       GoSub PrintOPHeader
     End If
     Print #RptHandle, String(89, "-")
     Print #RptHandle, "Totals"; Tab(12); "# Over Payments: " + CStr(OPCnt); Tab(45); Using(UseThis, OPToOwed); Tab(61); Using(UseThis, OPPaid); Tab(79); Using(UseThis, OldRound(OPToOwed + OPPaid))
  Else
    Print #RptHandle,
    Print #RptHandle, "No Personal Over Payment Activity"
    LineCnt = LineCnt + 2
  End If
  
  '---------------------------------------------------------------------------------------
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle, sLine
  Print #RptHandle, Tab(21); "Grand Totals"; Tab(50); Using(NoD, OldRound(PTotCash + RTotCash)); Tab(60); Using(NoD, OldRound(PTotCheck + RTotCheck)); Tab(70); Using(NoD, OldRound(PTotCharge + RTotCharge)); Tab(80); Using(NoD, OldRound(PTotDisc + RTotDisc)); Tab(90); Using(NoD, OldRound(PTotalPaid + RTotalPaid)); Tab(100); Using(NoD, OldRound(PTotChange + RTotChange))
  Print #RptHandle, "Total Number of Receipts: " + CStr(OldRound(NumOfPRecs + NumOfRRecs))
  Print #RptHandle, "Total Number of Checks: " + CStr(OldRound(PCheckCnt + RCheckCnt))
  LineCnt = LineCnt + 6
  Close RPayHandle
  
  RYearCnt = 0
  ReDim RYears(1 To 1) As Integer
  PYearCnt = 0
  ReDim PYears(1 To 1) As Integer
  
  OpenRealPayListFile RTHandle, OperNum
  NumOfRTRecs = LOF(RTHandle) / Len(RTempRPayRec)
  RYearCnt = 0
  For x = 1 To NumOfRTRecs
    Get RTHandle, x, RTempRPayRec
    If RTempRPayRec.PrevListRec < 0 Then GoTo MoveOnR3
    ThisYear = RTempRPayRec.TaxYear
    For y = 1 To RYearCnt
      If y <> x Then
        If RTempRPayRec.TaxYear = RYears(y) Then
          Exit For
        End If
      End If
    Next y
    If y > RYearCnt Then
      RYearCnt = RYearCnt + 1
      ReDim Preserve RYears(1 To RYearCnt) As Integer
      RYears(RYearCnt) = ThisYear
    End If
MoveOnR3:
  Next x
  
  OpenPersPayListFile PTHandle, OperNum
  NumOfPTRecs = LOF(PTHandle) / Len(PTempRPayRec)
  PYearCnt = 0
  For x = 1 To NumOfPTRecs
    Get PTHandle, x, PTempRPayRec
    If PTempRPayRec.PrevListRec < 0 Then GoTo MoveOnP3
    ThisYear = PTempRPayRec.TaxYear
    For y = 1 To PYearCnt
      If y <> x Then
        If PTempRPayRec.TaxYear = PYears(y) Then
          Exit For
        End If
      End If
    Next y
    If y > PYearCnt Then
      PYearCnt = PYearCnt + 1
      ReDim Preserve PYears(1 To PYearCnt) As Integer
      PYears(PYearCnt) = ThisYear
    End If
MoveOnP3:
  Next x
  
  If RYearCnt = 0 Then
    GoTo NoRs
  End If
  
  ReDim PrincByYr(1 To RYearCnt) As Double
  ReDim RIntByYr(1 To RYearCnt) As Double
  ReDim AdvColByYr(1 To RYearCnt) As Double
  ReDim LateListByYr(1 To RYearCnt) As Double
  ReDim RPenByYr(1 To RYearCnt) As Double
  ReDim Rev1ByYr(1 To RYearCnt) As Double
  ReDim Rev2ByYr(1 To RYearCnt) As Double
  ReDim Rev3ByYr(1 To RYearCnt) As Double
  ReDim RDiscByYr(1 To RYearCnt) As Double
  ReDim RTotPaidByYr(1 To RYearCnt) As Double
  ReDim ROverPayByYr(1 To RYearCnt) As Double
  GoSub SortRYears
  
NoRs:
  If RYearCnt >= PYearCnt Then
    GYearCnt = RYearCnt
  Else
    GYearCnt = PYearCnt
  End If
  
  ReDim GDiscByYr(1 To GYearCnt) As Double
  ReDim GTotPaidByYr(1 To GYearCnt) As Double
  ReDim GOverPayByYr(1 To GYearCnt) As Double
  
  '************************************
  If RTotCount = 0 Then GoSub PersPrint
  '************************************
  

  If NumOfRTRecs = 0 Then GoSub PersPrint
  
  GPrinc = 0
  RGInt = 0
  GAdvCol = 0
  GLateList = 0
  GRev1 = 0
  GRev2 = 0
  GRev3 = 0
  RGTot = 0
  RGDisc = 0
  RGOverPay = 0
  
  For x = 1 To NumOfRTRecs
    Get RTHandle, x, RTempRPayRec
       If RTempRPayRec.PrevListRec < 0 Then GoTo MoveOnR4
       GPrinc = OldRound(GPrinc + RTempRPayRec.Principle1)
       RGTot = OldRound(RGTot + RTempRPayRec.Principle1)
       RGInt = OldRound(RGInt + RTempRPayRec.Interest1)
       RGTot = OldRound(RGTot + RTempRPayRec.Interest1)
       GAdvCol = OldRound(GAdvCol + RTempRPayRec.Collection)
       RGTot = OldRound(RGTot + RTempRPayRec.Collection)
       GLateList = OldRound(GLateList + RTempRPayRec.LateList)
       RGTot = OldRound(RGTot + RTempRPayRec.LateList)
       GRPenalty = OldRound(GRPenalty + RTempRPayRec.Penalty)
       RGTot = OldRound(RGTot + RTempRPayRec.Penalty)
       GRev1 = OldRound(GRev1 + RTempRPayRec.OptRev1)
       RGTot = OldRound(RGTot + RTempRPayRec.OptRev1)
       GRev2 = OldRound(GRev2 + RTempRPayRec.OptRev2)
       RGTot = OldRound(RGTot + RTempRPayRec.OptRev2)
       GRev3 = OldRound(GRev3 + RTempRPayRec.OptRev3)
       RGTot = OldRound(RGTot + RTempRPayRec.OptRev3)
       RGDisc = OldRound(RGDisc + RTempRPayRec.DiscAmt)
       RGTot = OldRound(RGTot + RTempRPayRec.DiscAmt)
       RGOverPay = OldRound(RGOverPay + RTempRPayRec.PrePayAmt)
       RGTot = OldRound(RGTot + RTempRPayRec.PrePayAmt)
       For y = 1 To RYearCnt
         If RYears(y) = RTempRPayRec.TaxYear Then
           PrincByYr(y) = OldRound(PrincByYr(y) + RTempRPayRec.Principle1)
           RIntByYr(y) = OldRound(RIntByYr(y) + RTempRPayRec.Interest1)
           AdvColByYr(y) = OldRound(AdvColByYr(y) + RTempRPayRec.Collection)
           LateListByYr(y) = OldRound(LateListByYr(y) + RTempRPayRec.LateList)
           RPenByYr(y) = OldRound(RPenByYr(y) + RTempRPayRec.Penalty)
           Rev1ByYr(y) = OldRound(Rev1ByYr(y) + RTempRPayRec.OptRev1)
           Rev2ByYr(y) = OldRound(Rev2ByYr(y) + RTempRPayRec.OptRev2)
           Rev3ByYr(y) = OldRound(Rev3ByYr(y) + RTempRPayRec.OptRev3)
           RDiscByYr(y) = OldRound(RDiscByYr(y) + RTempRPayRec.DiscAmt)
           RTotPaidByYr(y) = OldRound(RTotPaidByYr(y) + RTempRPayRec.TotPaid)
           ROverPayByYr(y) = OldRound(ROverPayByYr(y) + RTempRPayRec.PrePayAmt)
           GDiscByYr(y) = OldRound(GDiscByYr(y) + RDiscByYr(y))
           GTotPaidByYr(y) = OldRound(GTotPaidByYr(y) + RTotPaidByYr(y))
           GOverPayByYr(y) = OldRound(GOverPayByYr(y) + ROverPayByYr(y))
           Exit For
         End If
       Next y
MoveOnR4:
  Next x
  
  Print #RptHandle, FF$
  GoSub PrintSubHeader
  
  Print #RptHandle, "Real Source Summary"
  Print #RptHandle, Tab(20); "Revenue Source"; Tab(55); "Amount"
  Print #RptHandle, Tab(20); String(41, "-")
  Print #RptHandle, Tab(20); "Tax Principle + Discount:"; Tab(50); Using(UseThis, OldRound(GPrinc + GDisc))
  Print #RptHandle, Tab(20); String(41, ".")
  Print #RptHandle, Tab(20); "Tax Principle:"; Tab(50); Using(UseThis, GPrinc)
  Print #RptHandle, Tab(20); "Discount:"; Tab(50); Using(UseThis, RGDisc)
  Print #RptHandle, Tab(20); "Interest:"; Tab(50); Using(UseThis, RGInt)
  Print #RptHandle, Tab(20); "Advertising/Collections"; Tab(50); Using(UseThis, GAdvCol)
  Print #RptHandle, Tab(20); "Late Listing:"; Tab(50); Using(UseThis, GLateList)
  Print #RptHandle, Tab(20); "Penalty:"; Tab(50); Using(UseThis, GRPenalty)
  Print #RptHandle, Tab(20); QPTrim$(TaxMasterRec.OptRev1); Tab(50); Using(UseThis, GRev1)
  Print #RptHandle, Tab(20); QPTrim$(TaxMasterRec.OptRev2); Tab(50); Using(UseThis, GRev2)
  Print #RptHandle, Tab(20); QPTrim$(TaxMasterRec.OptRev3); Tab(50); Using(UseThis, GRev3)
  Print #RptHandle, Tab(20); "Over Payment:"; Tab(50); Using(UseThis, RGOverPay)
  Print #RptHandle, Tab(20); Tab(20); String(41, "-")
  Print #RptHandle, Tab(20); "Real Totals:"; Tab(50); Using(UseThis, RGTot)
  
  Close RTHandle
  
  Done = False
  RevTrunc1 = QPTrim$(TaxMasterRec.OptRev1)
  RevTrunc2 = QPTrim$(TaxMasterRec.OptRev2)
  RevTrunc3 = QPTrim$(TaxMasterRec.OptRev3)
  RSet RevTrunc1 = QPTrim$(RevTrunc1)
  RSet RevTrunc2 = QPTrim$(RevTrunc2)
  RSet RevTrunc3 = QPTrim$(RevTrunc3)
  
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle, "Real Breakdown By Year"
  Print #RptHandle, Tab(2); "Tax Year"; Tab(14); "Principle Paid"; Tab(34); "Adv/Col Paid"; Tab(52); RevTrunc1; Tab(72); RevTrunc3; Tab(90); "Discounts Allowed"
  Print #RptHandle, Tab(15); "Interest Paid"; Tab(32); "Late List Paid"; Tab(57); "Penalty"; Tab(72); RevTrunc2; Tab(95); "Over Payment"; Tab(110); "Payments - Disc"
  Print #RptHandle, String(124, "-")
  
  For x = 1 To RYearCnt
    If x = RYearCnt Then Done = True
    PrincByYrTot = OldRound(PrincByYrTot + PrincByYr(x)) ' + DiscByYr(x))
    RIntByYrTot = OldRound(RIntByYrTot + RIntByYr(x))
    AdvColByYrTot = OldRound(AdvColByYrTot + AdvColByYr(x))
    LateListByYrTot = OldRound(LateListByYrTot + LateListByYr(x))
    Rev1ByYrTot = OldRound(Rev1ByYrTot + Rev1ByYr(x))
    Rev2ByYrTot = OldRound(Rev2ByYrTot + Rev2ByYr(x))
    Rev3ByYrTot = OldRound(Rev3ByYrTot + Rev3ByYr(x))
    RDiscByYrTot = OldRound(RDiscByYrTot + RDiscByYr(x))
    ROverPayByYrTot = OldRound(ROverPayByYrTot + ROverPayByYr(x))
    RTotPaidByYrTot = OldRound(RTotPaidByYrTot + RTotPaidByYr(x))
    YTotal = OldRound(PrincByYr(x) + RIntByYr(x) + AdvColByYr(x) + LateListByYr(x) + RPenByYr(x) + Rev1ByYr(x) + Rev2ByYr(x) + Rev3ByYr(x) + ROverPayByYr(x))
    Print #RptHandle, Tab(4); CStr(RYears(x)); Tab(17); Using(UseThis, OldRound(PrincByYr(x) + RDiscByYr(x))); Tab(35); Using(UseThis, AdvColByYr(x));
    Print #RptHandle, Tab(53); Using(UseThis, Rev1ByYr(x)); Tab(73); Using(UseThis, Rev3ByYr(x)); Tab(96); Using(UseThis, RDiscByYr(x))
    LineCnt = LineCnt + 1
    If x >= LineCnt - 4 Then
      If LineCnt > MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintSubHeader
        Print #RptHandle, "Real BreakDown By Year"
      End If
    Else
      If LineCnt > MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintSubHeader
        Print #RptHandle, "Real BreakDown By Year"
      End If
    End If
    Print #RptHandle, Tab(17); Using(UseThis, RIntByYr(x)); Tab(35); Using(UseThis, LateListByYr(x)); Tab(53); Using(UseThis, RPenByYr(x));
    Print #RptHandle, Tab(73); Using(UseThis, Rev2ByYr(x)); Tab(96); Using(UseThis, ROverPayByYr(x)); Tab(114); Using(UseThis, YTotal)
    LineCnt = LineCnt + 1
    If x >= LineCnt - 4 Then
      If LineCnt > MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintSubHeader
        Print #RptHandle, "Real BreakDown By Year"
      End If
    Else
      If LineCnt > MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintSubHeader
        Print #RptHandle, "Real BreakDown By Year"
      End If
    End If
    Print #RptHandle, String(124, "-")
    LineCnt = LineCnt + 1
    If x >= LineCnt - 4 Then
      If LineCnt > MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintSubHeader
        Print #RptHandle, "Real BreakDown By Year"
      End If
    Else
      If LineCnt > MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintSubHeader
        Print #RptHandle, "Real BreakDown By Year"
      End If
    End If
  Next x
  
  '************************************
  If PTotalPaid > 0 Then GoSub PersPrint
  '************************************
  Print #RptHandle, FF$
  Close
  
  ViewPrint RptFile, "Printing Payment Edit Report", True
  
  Exit Sub
  
PersPrint:
  ReDim PersByYr(1 To PYearCnt) As Double
  ReDim PIntByYr(1 To PYearCnt) As Double
  ReDim MTByYr(1 To PYearCnt) As Double
  ReDim MCByYr(1 To PYearCnt) As Double
  ReDim FEByYr(1 To PYearCnt) As Double
  ReDim MHByYr(1 To PYearCnt) As Double
  ReDim PenByYr(1 To PYearCnt) As Double
  ReDim OPt1ByYr(1 To PYearCnt) As Double
  ReDim OPt2ByYr(1 To PYearCnt) As Double
  ReDim OPt3ByYr(1 To PYearCnt) As Double
  ReDim PDiscByYr(1 To PYearCnt) As Double
  ReDim PTotPaidByYr(1 To PYearCnt) As Double
  ReDim POverPayByYr(1 To PYearCnt) As Double
  GoSub SortPYears
  
  GPers = 0
  GMachTools = 0
  GMerchCap = 0
  GFarmEq = 0
  GMobHomes = 0
  GPPenalty = 0
  PGInt = 0
  PGTot = 0
  
  For x = 1 To NumOfPTRecs
    Get PTHandle, x, PTempRPayRec
      If PTempRPayRec.PrevListRec < 0 Then GoTo MoveOnP4
      GPers = OldRound(GPers + PTempRPayRec.Personal)
      PGTot = OldRound(PGTot + PTempRPayRec.Personal)
      PGInt = OldRound(PGInt + PTempRPayRec.Interest)
      PGTot = OldRound(PGTot + PTempRPayRec.Interest)
      GMachTools = OldRound(GMachTools + PTempRPayRec.MachTools)
      PGTot = OldRound(PGTot + PTempRPayRec.MachTools)
      GMerchCap = OldRound(GMerchCap + PTempRPayRec.MerchCap)
      PGTot = OldRound(PGTot + PTempRPayRec.MerchCap)
      GFarmEq = OldRound(GFarmEq + PTempRPayRec.FarmEquip)
      PGTot = OldRound(PGTot + PTempRPayRec.FarmEquip)
      GMobHomes = OldRound(GMobHomes + PTempRPayRec.MobHomes)
      PGTot = OldRound(PGTot + PTempRPayRec.MobHomes)
      GPPenalty = OldRound(GPPenalty + PTempRPayRec.Penalty)
      PGTot = OldRound(PGTot + PTempRPayRec.Penalty)
      GPOptRev1 = OldRound(GPOptRev1 + PTempRPayRec.Opt1)
      PGTot = OldRound(PGTot + PTempRPayRec.Opt1)
      GPOptRev2 = OldRound(GPOptRev2 + PTempRPayRec.Opt2)
      PGTot = OldRound(PGTot + PTempRPayRec.Opt2)
      GPOptRev3 = OldRound(GPOptRev3 + PTempRPayRec.Opt3)
      PGTot = OldRound(PGTot + PTempRPayRec.Opt3)
'      PGDisc = OldRound(PGDisc + PTempRPayRec.DiscAmt)
      PGDisc = PTempRPayRec.DiscAmt
      PGTot = OldRound(PGTot + PTempRPayRec.DiscAmt)
      PGOverPay = OldRound(PGOverPay + PTempRPayRec.PrePayAmt)
      PGTot = OldRound(PGTot + PTempRPayRec.PrePayAmt)
      For y = 1 To PYearCnt
        If PYears(y) = PTempRPayRec.TaxYear Then
          PersByYr(y) = OldRound(PersByYr(y) + PTempRPayRec.Personal)
          PIntByYr(y) = OldRound(PIntByYr(y) + PTempRPayRec.Interest)
          MTByYr(y) = OldRound(MTByYr(y) + PTempRPayRec.MachTools)
          MCByYr(y) = OldRound(MCByYr(y) + PTempRPayRec.MerchCap)
          FEByYr(y) = OldRound(FEByYr(y) + PTempRPayRec.FarmEquip)
          MHByYr(y) = OldRound(MHByYr(y) + PTempRPayRec.MobHomes)
          PenByYr(y) = OldRound(PenByYr(y) + PTempRPayRec.Penalty)
          OPt1ByYr(y) = OldRound(OPt1ByYr(y) + PTempRPayRec.Opt1)
          OPt2ByYr(y) = OldRound(OPt2ByYr(y) + PTempRPayRec.Opt2)
          OPt3ByYr(y) = OldRound(OPt3ByYr(y) + PTempRPayRec.Opt3)
          PDiscByYr(y) = OldRound(PDiscByYr(y) + PTempRPayRec.DiscAmt)
          PTotPaidByYr(y) = OldRound(PTotPaidByYr(y) + PTempRPayRec.TotPaid)
          POverPayByYr(y) = OldRound(POverPayByYr(y) + PTempRPayRec.PrePayAmt)
          GDiscByYr(y) = OldRound(GDiscByYr(y) + PDiscByYr(y))
          GTotPaidByYr(y) = OldRound(GTotPaidByYr(y) + PTotPaidByYr(y))
          GOverPayByYr(y) = OldRound(GOverPayByYr(y) + POverPayByYr(y))
          Exit For
        End If
     Next y
MoveOnP4:
  Next x

  Print #RptHandle, FF$
  GoSub PrintSubHeader
  
  Print #RptHandle, "Personal Source Summary"
  Print #RptHandle, Tab(20); "Revenue Source"; Tab(55); "Amount"
  Print #RptHandle, Tab(20); String(41, "-")
  Print #RptHandle, Tab(20); "Tax Personal:"; Tab(50); Using(UseThis, GPers)
  Print #RptHandle, Tab(20); "Discount:"; Tab(50); Using(UseThis, PGDisc)
  Print #RptHandle, Tab(20); "Mach Tools:"; Tab(50); Using(UseThis, GMachTools)
  Print #RptHandle, Tab(20); "Merch Cap:"; Tab(50); Using(UseThis, GMerchCap)
  Print #RptHandle, Tab(20); "Farm Equip"; Tab(50); Using(UseThis, GFarmEq)
  Print #RptHandle, Tab(20); "Mob Homes:"; Tab(50); Using(UseThis, GMobHomes)
  Print #RptHandle, Tab(20); "Interest"; Tab(50); Using(UseThis, PGInt)
  Print #RptHandle, Tab(20); "Penalty"; Tab(50); Using(UseThis, GPPenalty)
  If Opt1Desc <> "" Then
    Print #RptHandle, Tab(20); Opt1Desc; Tab(50); Using(UseThis, GPOptRev1)
  End If
  If Opt2Desc <> "" Then
    Print #RptHandle, Tab(20); Opt2Desc; Tab(50); Using(UseThis, GPOptRev2)
  End If
  If Opt3Desc <> "" Then
    Print #RptHandle, Tab(20); Opt3Desc; Tab(50); Using(UseThis, GPOptRev3)
  End If
  Print #RptHandle, Tab(20); "Over Payment:"; Tab(50); Using(UseThis, PGOverPay)
  Print #RptHandle, Tab(20); Tab(20); String(41, "-")
  Print #RptHandle, Tab(20); "Personal Totals:"; Tab(50); Using(UseThis, PGTot)
  
  Close RTHandle
  Dim ThisOpt1 As String * 13
  Dim ThisOpt2 As String * 14
  Dim ThisOpt3 As String * 12
  
  Done = False
  'now personal
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle, "Personal Breakdown By Year"
  Print #RptHandle, Tab(2); "Tax Year"; Tab(15); "Personal Paid"; Tab(32); "Merch Cap Paid"; Tab(50); "Mob Homes Paid"; Tab(72); "Penalty Paid"; Tab(90); "Discounts Allowed"
  Print #RptHandle, Tab(13); "Mach Tools Paid"; Tab(31); "Farm Equip Paid"; Tab(51); "Interest Paid"; Tab(72); "Over Payment"; Tab(92); "Payments - Disc"
  
  If Opt1Desc <> "" And Opt2Desc <> "" And Opt3Desc <> "" Then
    RSet ThisOpt1 = Opt1Desc
    RSet ThisOpt2 = Opt2Desc
    RSet ThisOpt3 = Opt3Desc
    Print #RptHandle, Tab(15); ThisOpt1; Tab(32); ThisOpt2; Tab(52); ThisOpt3
  ElseIf Opt1Desc <> "" And Opt2Desc = "" And Opt3Desc = "" Then
    RSet ThisOpt1 = Opt1Desc
    Print #RptHandle, Tab(15); ThisOpt1
  ElseIf Opt1Desc = "" And Opt2Desc <> "" And Opt3Desc = "" Then
    RSet ThisOpt2 = Opt2Desc
    Print #RptHandle, Tab(15); ThisOpt2
  ElseIf Opt1Desc = "" And Opt2Desc = "" And Opt3Desc <> "" Then
    RSet ThisOpt3 = Opt3Desc
    Print #RptHandle, Tab(15); ThisOpt3
  ElseIf Opt1Desc <> "" And Opt2Desc <> "" And Opt3Desc = "" Then
    RSet ThisOpt1 = Opt1Desc
    RSet ThisOpt2 = Opt2Desc
    Print #RptHandle, Tab(15); ThisOpt1; Tab(32); ThisOpt2
  ElseIf Opt1Desc <> "" And Opt2Desc = "" And Opt3Desc <> "" Then
    RSet ThisOpt1 = Opt1Desc
    RSet ThisOpt3 = Opt3Desc
    Print #RptHandle, Tab(15); ThisOpt1; Tab(32); ThisOpt3
  ElseIf Opt1Desc = "" And Opt2Desc <> "" And Opt3Desc <> "" Then
    RSet ThisOpt2 = Opt2Desc
    RSet ThisOpt3 = Opt3Desc
    Print #RptHandle, Tab(15); ThisOpt2; Tab(32); ThisOpt3
  End If
  
  Print #RptHandle, String(106, "-")
  
  For x = 1 To PYearCnt
    If x = PYearCnt Then Done = True
    PersByYrTot = OldRound(PersByYrTot + PersByYr(x) + PDiscByYr(x))
    PIntByYrTot = OldRound(PIntByYrTot + PIntByYr(x))
    MTByYrTot = OldRound(MTByYrTot + MTByYr(x))
    MCByYrTot = OldRound(MCByYrTot + MCByYr(x))
    FEByYrTot = OldRound(FEByYrTot + FEByYr(x))
    MHByYrTot = OldRound(MHByYrTot + MHByYr(x))
    PPenByYrTot = OldRound(PPenByYrTot + PenByYr(x))
    POpt1ByYrTot = OldRound(POpt1ByYrTot + OPt1ByYr(x))
    POpt2ByYrTot = OldRound(POpt2ByYrTot + OPt2ByYr(x))
    POpt3ByYrTot = OldRound(POpt3ByYrTot + OPt3ByYr(x))
    PDiscByYrTot = OldRound(PDiscByYrTot + PDiscByYr(x))
    POverPayByYrTot = OldRound(POverPayByYrTot + POverPayByYr(x))
    PTotPaidByYrTot = OldRound(PTotPaidByYrTot + PTotPaidByYr(x))
    YTotal = OldRound(OPt1ByYr(x) + OPt2ByYr(x) + OPt3ByYr(x) + PersByYr(x) + PIntByYr(x) + MTByYr(x) + MCByYr(x) + FEByYr(x) + MHByYr(x) + PenByYr(x) + POverPayByYr(x))
    Print #RptHandle, Tab(4); CStr(PYears(x)); Tab(17); Using(UseThis, OldRound(PersByYr(x) + PDiscByYr(x))); Tab(35); Using(UseThis, MCByYr(x));
    Print #RptHandle, Tab(53); Using(UseThis, MHByYr(x)); Tab(73); Using(UseThis, PenByYr(x)); Tab(96); Using(UseThis, PDiscByYr(x))
    LineCnt = LineCnt + 1
    If x >= LineCnt - 5 Then
      If LineCnt > MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintSubHeader
        Print #RptHandle, "Personal BreakDown By Year"
      End If
    Else
      If LineCnt > MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintSubHeader
        Print #RptHandle, "Personal BreakDown By Year"
      End If
    End If
    Print #RptHandle, Tab(17); Using(UseThis, MTByYr(x)); Tab(35); Using(UseThis, FEByYr(x)); Tab(53); Using(UseThis, PIntByYr(x));
    Print #RptHandle, Tab(73); Using(UseThis, POverPayByYr(x)); Tab(96); Using(UseThis, YTotal)
    LineCnt = LineCnt + 1
    
    If Opt1Desc <> "" And Opt2Desc <> "" And Opt3Desc <> "" Then
      Print #RptHandle, Tab(17); Using(UseThis, OPt1ByYr(x)); Tab(35); Using(UseThis, OPt2ByYr(x)); Tab(53); Using(UseThis, OPt3ByYr(x))
    ElseIf Opt1Desc <> "" And Opt2Desc = "" And Opt3Desc = "" Then
      Print #RptHandle, Tab(17); Using(UseThis, OPt1ByYr(x))
    ElseIf Opt1Desc = "" And Opt2Desc <> "" And Opt3Desc = "" Then
      Print #RptHandle, Tab(17); Using(UseThis, OPt2ByYr(x))
    ElseIf Opt1Desc = "" And Opt2Desc = "" And Opt3Desc <> "" Then
      Print #RptHandle, Tab(17); Using(UseThis, OPt3ByYr(x))
    ElseIf Opt1Desc <> "" And Opt2Desc <> "" And Opt3Desc = "" Then
      Print #RptHandle, Tab(17); Using(UseThis, OPt1ByYr(x)); Tab(35); Using(UseThis, OPt2ByYr(x))
    ElseIf Opt1Desc <> "" And Opt2Desc = "" And Opt3Desc <> "" Then
      Print #RptHandle, Tab(17); Using(UseThis, OPt1ByYr(x)); Tab(35); Using(UseThis, OPt3ByYr(x))
    ElseIf Opt1Desc = "" And Opt2Desc <> "" And Opt3Desc <> "" Then
      Print #RptHandle, Tab(17); Using(UseThis, OPt2ByYr(x)); Tab(35); Using(UseThis, OPt3ByYr(x))
    End If
    
    LineCnt = LineCnt + 1
    
    If x >= LineCnt - 5 Then
      If LineCnt > MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintSubHeader
        Print #RptHandle, "Personal BreakDown By Year"
      End If
    Else
      If LineCnt > MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintSubHeader
        Print #RptHandle, "Personal BreakDown By Year"
      End If
    End If
    Print #RptHandle, String(106, "-")
    LineCnt = LineCnt + 1
    If x >= LineCnt - 5 Then
      If LineCnt > MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintSubHeader
        Print #RptHandle, "Personal BreakDown By Year"
      End If
    Else
      If LineCnt > MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintSubHeader
        Print #RptHandle, "Personal BreakDown By Year"
      End If
    End If
  Next x
  
  Return
  
SortRYears:
  
  LilYear = 1900
  Nextx = 1
  Do
    For x = Nextx To RYearCnt
      If RYears(x) > LilYear Then
        LilYear = RYears(x)
        Thisx = x
      End If
    Next x
    HoldYears = RYears(Nextx)
    HoldPrincByYr = PrincByYr(Nextx)
    HoldRIntByYr = RIntByYr(Nextx)
    HoldAdvColByYr = AdvColByYr(Nextx)
    HoldLateListByYr = LateListByYr(Nextx)
    HoldRev1ByYr = Rev1ByYr(Nextx)
    HoldRev2ByYr = Rev2ByYr(Nextx)
    HoldRev3ByYr = Rev3ByYr(Nextx)
    HoldRDiscByYr = RDiscByYr(Nextx)
    HoldRTotPaidByYr = RTotPaidByYr(Nextx)
    RYears(Nextx) = RYears(Thisx)
    PrincByYr(Nextx) = PrincByYr(Thisx)
    RIntByYr(Nextx) = RIntByYr(Thisx)
    AdvColByYr(Nextx) = AdvColByYr(Thisx)
    LateListByYr(Nextx) = LateListByYr(Thisx)
    Rev1ByYr(Nextx) = Rev1ByYr(Thisx)
    Rev2ByYr(Nextx) = Rev2ByYr(Thisx)
    Rev3ByYr(Nextx) = Rev3ByYr(Thisx)
    RDiscByYr(Nextx) = RDiscByYr(Thisx)
    RTotPaidByYr(Nextx) = RTotPaidByYr(Thisx)
    RYears(Thisx) = HoldYears
    PrincByYr(Thisx) = HoldPrincByYr
    RIntByYr(Thisx) = HoldRIntByYr
    AdvColByYr(Thisx) = HoldAdvColByYr
    LateListByYr(Thisx) = HoldLateListByYr
    Rev1ByYr(Thisx) = HoldRev1ByYr
    Rev2ByYr(Thisx) = HoldRev2ByYr
    Rev3ByYr(Thisx) = HoldRev3ByYr
    RDiscByYr(Thisx) = HoldRDiscByYr
    RTotPaidByYr(Thisx) = HoldRTotPaidByYr
    LilYear = 1900
    Nextx = Nextx + 1
    If Nextx > RYearCnt Then Exit Do
  Loop
  
  Return
  
SortPYears:
  'now personal
  LilYear = 1900
  Nextx = 1
  Do
    For x = Nextx To PYearCnt
      If PYears(x) > LilYear Then
        LilYear = PYears(x)
        Thisx = x
      End If
    Next x
    HoldYears = PYears(Nextx)
    HoldPersByYr = PersByYr(Nextx)
    HoldMTByYr = MTByYr(Nextx)
    HoldMCByYr = MCByYr(Nextx)
    HoldFEByYr = FEByYr(Nextx)
    HoldMHByYr = MHByYr(Nextx)
    HoldPIntByYr = PIntByYr(Nextx)
    HoldPenByYr = PenByYr(Nextx)
    HoldPDiscByYr = PDiscByYr(Nextx)
    HoldPTotPaidByYr = PTotPaidByYr(Nextx)
    PYears(Nextx) = PYears(Thisx)
    PersByYr(Nextx) = PersByYr(Thisx)
    MTByYr(Nextx) = MTByYr(Thisx)
    MCByYr(Nextx) = MCByYr(Thisx)
    FEByYr(Nextx) = FEByYr(Thisx)
    MHByYr(Nextx) = MHByYr(Thisx)
    PIntByYr(Nextx) = PIntByYr(Thisx)
    PenByYr(Nextx) = PenByYr(Thisx)
    PDiscByYr(Nextx) = PDiscByYr(Thisx)
    PTotPaidByYr(Nextx) = PTotPaidByYr(Thisx)
    PYears(Thisx) = HoldYears
    PersByYr(Thisx) = HoldPersByYr
    MTByYr(Thisx) = HoldMTByYr
    MCByYr(Thisx) = HoldMCByYr
    FEByYr(Thisx) = HoldFEByYr
    MHByYr(Thisx) = HoldMHByYr
    PIntByYr(Thisx) = HoldPIntByYr
    PenByYr(Thisx) = HoldPenByYr
    PDiscByYr(Thisx) = HoldPDiscByYr
    PTotPaidByYr(Thisx) = HoldPTotPaidByYr
    LilYear = 1900
    Nextx = Nextx + 1
    If Nextx > PYearCnt Then Exit Do
  Loop

  Return
  
SortRCustomers:
  OpenTaxCustFile TCHandle, NumOfTCRecs
  Nextx = 0
  For x = 1 To NumOfRRecs
    Get RPayHandle, x, RPayRec
    If RPayRec.LastPayRec = 0 Then GoTo MoveALongR
    Get TCHandle, RPayRec.CustAcct, TaxCust
    Nextx = Nextx + 1
'    RCustNArr(Nextx) = QPTrim$(RPayRec.CustName)
    RCustNArr(Nextx) = QPTrim$(TaxCust.SName)
    RCustRArr(Nextx) = x
MoveALongR:
  Next x
  Close TCHandle
  NumOfRRecs = Nextx
  BigName$ = ""
  For x = 1 To NumOfRRecs
    If RCustNArr(x) > BigName Then
      BigName = RCustNArr(x)
    End If
  Next x
  
  LilName = BigName + "z"
  NextOne = 1
  
  Do
    For x = NextOne To NumOfRRecs
      If RCustNArr(x) < LilName Then
        LilName = RCustNArr(x)
        Thisx = x
      End If
    Next x
    HoldName = RCustNArr(NextOne)
    HoldRec = RCustRArr(NextOne)
    RCustNArr(NextOne) = RCustNArr(Thisx)
    RCustRArr(NextOne) = RCustRArr(Thisx)
    RCustNArr(Thisx) = HoldName
    RCustRArr(Thisx) = HoldRec
    NextOne = NextOne + 1
    LilName = BigName + "z"
    If NextOne > NumOfRRecs Then Exit Do
  Loop
  
  Return

SortPCustomers:
  OpenTaxCustFile TCHandle, NumOfTCRecs
  Nextx = 0
  For x = 1 To NumOfPRecs
    Get PPayHandle, x, PPayRec
    If PPayRec.LastPayRec = 0 Then GoTo MoveALongP
    Get TCHandle, PPayRec.CustAcct, TaxCust
    Nextx = Nextx + 1
'    PCustNArr(Nextx) = QPTrim$(PPayRec.CustName)
    PCustNArr(Nextx) = QPTrim$(TaxCust.SName)
    PCustRArr(Nextx) = x
MoveALongP:
  Next x
  Close TCHandle
  NumOfPRecs = Nextx
  BigName$ = ""
  For x = 1 To NumOfPRecs
    If PCustNArr(x) > BigName Then
      BigName = PCustNArr(x)
    End If
  Next x
  
  LilName = BigName + "z"
  NextOne = 1
  
  Do
    For x = NextOne To NumOfPRecs
      If PCustNArr(x) < LilName Then
        LilName = PCustNArr(x)
        Thisx = x
      End If
    Next x
    HoldName = PCustNArr(NextOne)
    HoldRec = PCustRArr(NextOne)
    PCustNArr(NextOne) = PCustNArr(Thisx)
    PCustRArr(NextOne) = PCustRArr(Thisx)
    PCustNArr(Thisx) = HoldName
    PCustRArr(Thisx) = HoldRec
    NextOne = NextOne + 1
    LilName = BigName + "z"
    If NextOne > NumOfPRecs Then Exit Do
  Loop
  Return
  
PrintHeader:
  OperLen = Len("Operator # " + CStr(OperNum) + " " + PWUser)
  ThisTab = OperLen / 2
  ThisTab = ThisTab + 45
  Page = Page + 1
  Print #RptHandle, Tab(45); "Tax Payment Transaction Journal"
  Print #RptHandle, Tab(ThisTab); "Operator # " + CStr(OperNum) + " " + PWUser
  Print #RptHandle, Town; Tab(100); "Page #" + CStr(Page)
  Print #RptHandle, "Report Date: " + CStr(Date)
  Print #RptHandle, Tab(95); "Total"
  Print #RptHandle, "Date"; Tab(13); "Cust Num"; Tab(26); "Customer"; Tab(56); "Cash"; Tab(65); "Check"; Tab(74); "Charge"; Tab(82); "Discount"; Tab(92); "Credited"; Tab(104); "Change"
  Print #RptHandle, sLine
  LineCnt = 7
  
  Return
 
PrintSubHeader:
  OperLen = Len("Operator # " + CStr(OperNum) + " " + PWUser)
  ThisTab = OperLen / 2
  ThisTab = ThisTab + 45
  
  Page = Page + 1
  Print #RptHandle, Tab(45); "Tax Payment Transaction Journal"
  Print #RptHandle, Tab(ThisTab); "Operator # " + CStr(OperNum) + " " + PWUser
  Print #RptHandle, Town; Tab(100); "Page #" + CStr(Page)
  Print #RptHandle, "Report Date: " + CStr(Date)
  Print #RptHandle, sLine
  LineCnt = 6
  
  Return
  
PrintOPHeader:
  OperLen = Len("Operator # " + CStr(OperNum) + " " + PWUser)
  ThisTab = OperLen / 2
  ThisTab = ThisTab + 45
  
  Page = Page + 1
  Print #RptHandle, Tab(45); "Tax Payment Transaction Journal"
  Print #RptHandle, Tab(ThisTab); "Operator # " + CStr(OperNum) + " " + PWUser
  Print #RptHandle, Town; Tab(100); "Page #" + CStr(Page)
  Print #RptHandle, "Report Date: " + CStr(Date)
  Print #RptHandle, String(89, "-")
  Print #RptHandle, "Over Payment Summary"
  Print #RptHandle, Tab(38); "Payment Applied To"; Tab(60); "Over Payment"; Tab(78); "Total Amount"
  Print #RptHandle, "Cust Num"; Tab(12); "Customer"; Tab(45); "Amount Owed"; Tab(66); "Amount"; Tab(86); "Paid"
  Print #RptHandle, String(89, "-")
  LineCnt = 6
  
  Return
  
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPayEditList", "PrintGraphics", Erl)
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

