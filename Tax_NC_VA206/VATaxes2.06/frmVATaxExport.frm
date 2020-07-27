VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmVATaxExport 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Export Function"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmVATaxExport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcmbIncAbsYN 
      Height          =   384
      Left            =   6960
      TabIndex        =   2
      Top             =   4200
      Width           =   852
      _Version        =   196608
      _ExtentX        =   1503
      _ExtentY        =   677
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
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
      Text            =   ""
      Columns         =   0
      Sorted          =   0
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   -1
      ColumnWidthScale=   2
      RowHeight       =   -1
      WrapList        =   0   'False
      WrapWidth       =   0
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
      DataFieldList   =   ""
      ColumnEdit      =   -1
      ColumnBound     =   -1
      Style           =   2
      MaxDrop         =   8
      ListWidth       =   -1
      EditHeight      =   -1
      GrayAreaColor   =   -2147483633
      ListLeftOffset  =   0
      ComboGap        =   -2
      MaxEditLen      =   150
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
      ListPosition    =   0
      ButtonThreeDAppearance=   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      Redraw          =   -1  'True
      AutoSearchFill  =   0   'False
      AutoSearchFillDelay=   500
      EditMarginLeft  =   1
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   1
      EditAlignV      =   0
      ColDesigner     =   "frmVATaxExport.frx":08CA
   End
   Begin LpLib.fpCombo fpcmbExOrder 
      Height          =   384
      Left            =   5280
      TabIndex        =   0
      Top             =   2640
      Width           =   3012
      _Version        =   196608
      _ExtentX        =   5313
      _ExtentY        =   677
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
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
      Text            =   ""
      Columns         =   0
      Sorted          =   0
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   -1
      ColumnWidthScale=   2
      RowHeight       =   -1
      WrapList        =   0   'False
      WrapWidth       =   0
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
      DataFieldList   =   ""
      ColumnEdit      =   -1
      ColumnBound     =   -1
      Style           =   2
      MaxDrop         =   8
      ListWidth       =   -1
      EditHeight      =   -1
      GrayAreaColor   =   -2147483633
      ListLeftOffset  =   0
      ComboGap        =   -2
      MaxEditLen      =   150
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
      ListPosition    =   0
      ButtonThreeDAppearance=   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      Redraw          =   -1  'True
      AutoSearchFill  =   0   'False
      AutoSearchFillDelay=   500
      EditMarginLeft  =   1
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   1
      EditAlignV      =   0
      ColDesigner     =   "frmVATaxExport.frx":0BC1
   End
   Begin LpLib.fpCombo fpcmbPrintOpt 
      Height          =   384
      Left            =   5640
      TabIndex        =   1
      Top             =   3432
      Width           =   3012
      _Version        =   196608
      _ExtentX        =   5313
      _ExtentY        =   677
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
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
      Text            =   ""
      Columns         =   0
      Sorted          =   0
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   -1
      ColumnWidthScale=   2
      RowHeight       =   -1
      WrapList        =   0   'False
      WrapWidth       =   0
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
      DataFieldList   =   ""
      ColumnEdit      =   -1
      ColumnBound     =   -1
      Style           =   2
      MaxDrop         =   8
      ListWidth       =   -1
      EditHeight      =   -1
      GrayAreaColor   =   -2147483633
      ListLeftOffset  =   0
      ComboGap        =   -2
      MaxEditLen      =   150
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
      ListPosition    =   0
      ButtonThreeDAppearance=   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      Redraw          =   -1  'True
      AutoSearchFill  =   0   'False
      AutoSearchFillDelay=   500
      EditMarginLeft  =   1
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   1
      EditAlignV      =   0
      ColDesigner     =   "frmVATaxExport.frx":0EB8
   End
   Begin EditLib.fpText fptxtFile 
      Height          =   375
      Left            =   5400
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5040
      Width           =   2055
      _Version        =   196608
      _ExtentX        =   3625
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
      ControlType     =   0
      Text            =   "TAXDATA1"
      CharValidationText=   ""
      MaxLength       =   50
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
      Height          =   492
      Left            =   3480
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   6600
      Width           =   2052
      _Version        =   131072
      _ExtentX        =   3619
      _ExtentY        =   868
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
      ButtonDesigner  =   "frmVATaxExport.frx":11AF
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
      Height          =   492
      Left            =   6120
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   6600
      Width           =   2052
      _Version        =   131072
      _ExtentX        =   3619
      _ExtentY        =   868
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
      ButtonDesigner  =   "frmVATaxExport.frx":138B
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   ".TXT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   7560
      TabIndex        =   9
      Top             =   5085
      Width           =   615
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Overview Printout:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   3105
      TabIndex        =   8
      Top             =   3555
      Width           =   2295
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Export To:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   3840
      TabIndex        =   7
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Include Abstracts (Y/N):"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   3840
      TabIndex        =   6
      Top             =   4320
      Width           =   2895
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Export Orders:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   3495
      Left            =   2400
      Top             =   2280
      Width           =   7095
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   750
      Index           =   1
      Left            =   1500
      Top             =   660
      Width           =   8655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Export Function"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2940
      TabIndex        =   3
      Top             =   795
      Width           =   6015
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   840
      Left            =   1500
      Top             =   600
      Width           =   8655
   End
End
Attribute VB_Name = "frmVATaxExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class

Private Sub cmdExit_Click()
  frmVATaxCustMaintMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub PrintText()
  Dim TaxSetupRec As TaxMasterType
  Dim TaxCustRec As TaxCustType
  Dim PropertyRec As PropertyRecType
  Dim PersRec As PersonalRecType
  Dim ReportFile$
  Dim FF$, x As Long
  Dim LineCnt As Integer
  Dim CustCnt As Integer
  Dim RptHandle As Integer
  Dim IdxRec As CustNameIdxType
  Dim IdxHandle As Integer
  Dim NumOfIdx As Long
  Dim DetailFlag As Boolean
  Dim RptOut As Integer
  Dim RptOutFile$
  Dim NumOfPersRecs As Long
  Dim PersHandle As Integer
  Dim RealRec As PropertyRecType
  Dim NumOfRealRecs As Long
  Dim RealHandle As Integer
  Dim TaxRec As TaxCustType
  Dim NumOfTaxRecs As Long
  Dim TaxHandle As Integer
  Dim UseNameIdx As Integer
  Dim PropertyRecord As Long
  Dim FileName$, CustRecNo As Long
  Dim SearchRec As SrchNameIdxType
  Dim SHandle As Integer
  Dim NumOfSrchRecs As Long
  Dim NumOfRecs As Long
  
  On Error GoTo ERRORSTUFF
  
  UseNameIdx = 0
  ReportFile$ = "TaxCust.PRN"   'Report File Name
  FF$ = Chr$(12)
  DetailFlag = False
  LineCnt = 0
  CustCnt = 0
  
  OpenTaxCustFile TaxHandle, NumOfTaxRecs
  If QPTrim$(fpcmbExOrder.Text) = "Name Order" Then
    OpenNameIdxFile IdxHandle, NumOfIdx
    If NumOfIdx = 0 Then
      frmVATaxMsg.Label1.Caption = "There are no customers saved."
      frmVATaxMsg.Label1.Top = 900
      frmVATaxMsg.Show
      fpcmbExOrder.SetFocus
      Close IdxHandle
      Exit Sub
    End If
    ReDim IdxArray(1 To NumOfIdx) As Long
    UseNameIdx = 1
    For x = 1 To NumOfIdx
      Get IdxHandle, x, IdxRec
      IdxArray(x) = IdxRec.CustRec
    Next x
    Close IdxHandle
    NumOfRecs = NumOfIdx
  ElseIf QPTrim$(fpcmbExOrder.Text) = "Search Name" Then
    OpenSrchNameIdxFile SHandle, NumOfSrchRecs
    If NumOfSrchRecs = 0 Then
      frmVATaxMsg.Label1.Caption = "There are no search names indexed."
      frmVATaxMsg.Label1.Top = 900
      frmVATaxMsg.Show
      fpcmbExOrder.SetFocus
      Close SHandle
      Exit Sub
    End If
    ReDim IdxArray(1 To NumOfSrchRecs) As Long
    UseNameIdx = 2
    For x = 1 To NumOfSrchRecs
      Get SHandle, x, SearchRec
      IdxArray(x) = SearchRec.CustRec
    Next x
    Close SHandle
    NumOfRecs = NumOfSrchRecs
  Else
    NumOfRecs = NumOfTaxRecs
  End If
  
  If QPTrim$(fpcmbIncAbsYN.Text) = "Y" Then
    DetailFlag = True
  End If

  RptHandle = FreeFile

  Open ReportFile$ For Output As #RptHandle

  FileName$ = QPTrim$(fptxtFile.Text) + ".TXT"
  
  RptOut = FreeFile
  Open FileName$ For Output As #RptOut

  OpenPersPropFile PersHandle, NumOfPersRecs
  OpenRealPropFile RealHandle, NumOfRealRecs
  
  frmVATaxShowPctComp.Label1 = "Gathering Export Data"
  frmVATaxShowPctComp.Show , Me
  EnableCloseButton Me.hwnd, False
  cmdProcess.Enabled = False
  cmdExit.Enabled = False
  
  For x = 1 To NumOfRecs ' NumOfTaxRecs
    If UseNameIdx > 0 Then
      CustRecNo = IdxArray(x)
    Else
      CustRecNo = x
    End If

    Get TaxHandle, CustRecNo, TaxRec

    If Not TaxRec.Deleted Then
        CustCnt = CustCnt + 1
        Print #RptOut, Using("#####", CustRecNo);
        Print #RptOut, "]";
        Print #RptOut, TaxRec.CustName; "]";
        Print #RptOut, (TaxRec.Addr1); "]";
        Print #RptOut, (TaxRec.Addr2); "]";
        Print #RptOut, (TaxRec.City); "]";
        Print #RptOut, (TaxRec.State); "]";
        Print #RptOut, (TaxRec.Zip); "]";
        Print #RptOut, (TaxRec.CSSN);
        If DetailFlag = False Then
          Print #RptOut, ""
          GoTo SkipEm
          Else
          Print #RptOut, "]";
        End If
        'Now Show Property Records Next

        If TaxRec.FirstPropRec > 0 Then
          PropertyRecord = TaxRec.FirstPropRec
          Do While PropertyRecord <> 0
            Get #RealHandle, PropertyRecord, RealRec
            If PropertyRecord <> TaxRec.FirstPropRec Then
                Print #RptOut, Using("#####", CustRecNo);
                Print #RptOut, "]";
                Print #RptOut, TaxRec.CustName; "]";
                Print #RptOut, (TaxRec.Addr1); "]";
                Print #RptOut, (TaxRec.Addr2); "]";
                Print #RptOut, (TaxRec.City); "]";
                Print #RptOut, (TaxRec.State); "]";
                Print #RptOut, (TaxRec.Zip); "]";
                Print #RptOut, (TaxRec.CSSN); "]";
            End If

            Print #RptOut, RealRec.RealPin; "]";
            Print #RptOut, RealRec.PROPNOT1; "]";
            Print #RptOut, RealRec.PROPNOT2; "]";
            Print #RptOut, RealRec.PROPNOT3; "]";
            Print #RptOut, Using("##########", RealRec.PROPVALU);
            Print #RptOut, "]";
'            Print #RptOut, Using("##########", RealRec.EXMPSENI);
            Print #RptOut, Using("##########", RealRec.EXMPOTHR); '6/14/06
            Print #RptOut, "]";
            Print #RptOut, RealRec.MORTCODE; "]";
            Print #RptOut, RealRec.Map; "]";
            Print #RptOut, RealRec.BLOCK; "]";
            Print #RptOut, RealRec.LOTNUMB; "]";
            PropertyRecord = RealRec.NextRec
            Loop
        Else
          Print #RptOut, "]"; "]"; "]"; "]"; "]"; "]"; "]"; "]"; "]"; "]";
        End If

        'NOW CHECK PERSONAL PROPERTY

        If TaxRec.FirstPersRec > 0 Then 'closed
          PropertyRecord = TaxRec.FirstPersRec
           Do While PropertyRecord <> 0
           Get #PersHandle, PropertyRecord, PersRec
            Print #RptOut, PersRec.PropPin; "]";
            Print #RptOut, Using("########", 0); '6/14/06 no more pers exemptions PersRec.PersVal);
            Print #RptOut, "]";
            Print #RptOut, Using("########", 0); '6/14/06 no more pers exemptions PersRec.EXMPSENI);
            Print #RptOut, "]";
            Print #RptOut, PersRec.DESC1; "]";
            Print #RptOut, PersRec.DESC2; "]";
            Print #RptOut, PersRec.DESC3; "]";
            Print #RptOut, PersRec.Desc4; "]";
            Print #RptOut, PersRec.Desc5; "]"
            PropertyRecord = PersRec.NextRec
           Loop
        Else
          Print #RptOut, " ]"; " ]"; " ]"; " ]"; " ]"; " ]"; " ]"; " ]"
        End If
     End If
SkipEm:
     frmVATaxShowPctComp.ShowPctComp x, NumOfTaxRecs
     If frmVATaxShowPctComp.Out = True Then
       Close
       frmVATaxShowPctComp.Out = False
       Unload frmVATaxShowPctComp
       EnableCloseButton Me.hwnd, True
       cmdProcess.Enabled = True
       cmdExit.Enabled = True
       Exit Sub
     End If
  Next
  
  Unload frmVATaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdProcess.Enabled = True
  cmdExit.Enabled = True
  
  If CustCnt = 0 Then
    frmVATaxMsg.Label1.Caption = "There are no export records."
    frmVATaxMsg.Label1.Top = 900
    frmVATaxMsg.Show vbModal
    Close
    Exit Sub
  End If
  
  GoSub PrintLayOutTable
  
  Print #RptHandle, FF$

  Close         'Close all open files now

  ViewPrint ReportFile$, "Tax Customer Export"
  Kill ReportFile$

  Exit Sub

PrintLayOutTable:
  Print #RptHandle, "Data File Export Facility : Taxes"
  Print #RptHandle, "            Date of Export: "; Date$
  Print #RptHandle, "  FileName Containing Data: "; FileName$
  Print #RptHandle, "           Field Delimiter:  ]"
  Print #RptHandle,
  Print #RptHandle, " 1 Customer Acct Number"
  Print #RptHandle, " 2 Customer Name"
  Print #RptHandle, " 3 Address Line 1"
  Print #RptHandle, " 4 Address Line 2"
  Print #RptHandle, " 5 City"
  Print #RptHandle, " 6 State"
  Print #RptHandle, " 7 Zip Code"
  Print #RptHandle, " 8 Customer Soc Sec #"
  If DetailFlag Then
   Print #RptHandle, " 9 Real Property PIN"
   Print #RptHandle, "10 Real Prop Desc Line 1"
   Print #RptHandle, "11 Real Prop Desc Line 2"
   Print #RptHandle, "12 Real Prop Desc Line 3"
   Print #RptHandle, "13 Real Prop Value"
   Print #RptHandle, "14 Real Prop Senior Citizen Exemption"
   Print #RptHandle, "15 Mortage Code"
   Print #RptHandle, "16 Map "
   Print #RptHandle, "17 Block"
   Print #RptHandle, "18 Lot"
   Print #RptHandle, "19 Personal Property Pin #"
   Print #RptHandle, "20 Personal Value"
   Print #RptHandle, "21 Exemption Value"
   Print #RptHandle, "22 Desc Line 1"
   Print #RptHandle, "23 Desc Line 2"
   Print #RptHandle, "24 Desc Line 3"
   Print #RptHandle, "25 Desc Line 4"
   Print #RptHandle, "26 Desc Line 5"
  End If

  Print #RptHandle, FF$
  Return

ERRORSTUFF:
   Unload frmVATaxShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxExport", "PrintText", Erl)
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

Private Sub cmdProcess_Click()
  If fpcmbPrintOpt.Text = "Graphical" Then
    Call PrintGraphics
  Else
    frmVATaxMsg.Label1.Caption = "Pitch 10 is recommended for this printout."
    frmVATaxMsg.Label1.Top = 900
    frmVATaxMsg.Show vbModal
    Call PrintText
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%E"
      Call cmdExit_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%P"
      Call cmdProcess_Click
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
  Me.HelpContextID = hlpExportCustomer
  Call LoadMe

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("CitiTaxes.exe terminated via menu bar on frmVATaxExport.")
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
  fpcmbIncAbsYN.Text = "N"
  fpcmbIncAbsYN.AddItem "N"
  fpcmbIncAbsYN.AddItem "Y"
  fpcmbExOrder.Text = "Name Order"
  fpcmbExOrder.AddItem "Name Order"
  fpcmbExOrder.AddItem "Account Number"
  fpcmbExOrder.AddItem "Search Name"
  fpcmbPrintOpt.Text = "Graphical"
  fpcmbPrintOpt.AddItem "Text"
  fpcmbPrintOpt.AddItem "Graphical"
End Sub

Private Sub fpcmbExOrder_Change()
  If QPTrim$(fpcmbExOrder.Text) = "" Then
    fpcmbExOrder.Text = "Name Order"
  End If
End Sub

Private Sub fpcmbExOrder_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbExOrder.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbExOrder.ListIndex = -1
  End If
  If fpcmbExOrder.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbPrintOpt.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbIncAbsYN_Change()
  If QPTrim$(fpcmbIncAbsYN.Text) = "" Then
    fpcmbIncAbsYN.Text = "N"
  End If
End Sub

Private Sub fpcmbIncAbsYN_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbIncAbsYN.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbIncAbsYN.ListIndex = -1
  End If
  If fpcmbIncAbsYN.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbExOrder.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If


End Sub

Private Sub fpcmbPrintOpt_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbPrintOpt.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbPrintOpt.ListIndex = -1
  End If
  If fpcmbPrintOpt.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbIncAbsYN.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub PrintGraphics()
  Dim TaxSetupRec As TaxMasterType
  Dim TaxCustRec As TaxCustType
  Dim PropertyRec As PropertyRecType
  Dim PersRec As PersonalRecType
  Dim ReportFile$
  Dim x As Long
  Dim CustCnt As Integer
  Dim RptHandle As Integer
  Dim IdxRec As CustNameIdxType
  Dim IdxHandle As Integer
  Dim NumOfIdx As Long
  Dim DetailFlag As Boolean
  Dim RptOut As Integer
  Dim RptOutFile$
  Dim NumOfPersRecs As Long
  Dim PersHandle As Integer
  Dim RealRec As PropertyRecType
  Dim NumOfRealRecs As Long
  Dim RealHandle As Integer
  Dim TaxRec As TaxCustType
  Dim NumOfTaxRecs As Long
  Dim TaxHandle As Integer
  Dim UseNameIdx As Integer
  Dim PropertyRecord As Long
  Dim FileName$, CustRecNo As Long
  Dim SearchRec As SrchNameIdxType
  Dim SHandle As Integer
  Dim NumOfSrchRecs As Long
  Dim NumOfRecs As Long
  Dim dlm$
  
  On Error GoTo ERRORSTUFF
  dlm$ = "~"
  UseNameIdx = 0
  ReportFile$ = "TAXRPTS\TaxExport.RPT"   'Report File Name
  DetailFlag = False
  CustCnt = 0
  OpenTaxCustFile TaxHandle, NumOfTaxRecs
  
  If QPTrim$(fpcmbExOrder.Text) = "Name Order" Then
    OpenNameIdxFile IdxHandle, NumOfIdx
    If NumOfIdx = 0 Then
      frmVATaxMsg.Label1.Caption = "There are no customers saved."
      frmVATaxMsg.Label1.Top = 900
      frmVATaxMsg.Show
      fpcmbExOrder.SetFocus
      Close IdxHandle
      Exit Sub
    End If
    ReDim IdxArray(1 To NumOfIdx) As Long
    UseNameIdx = 1
    For x = 1 To NumOfIdx
      Get IdxHandle, x, IdxRec
      IdxArray(x) = IdxRec.CustRec
    Next x
    Close IdxHandle
    NumOfRecs = NumOfIdx
  ElseIf QPTrim$(fpcmbExOrder.Text) = "Search Name" Then
    OpenSrchNameIdxFile SHandle, NumOfSrchRecs
    If NumOfSrchRecs = 0 Then
      frmVATaxMsg.Label1.Caption = "There are no search names indexed."
      frmVATaxMsg.Label1.Top = 900
      frmVATaxMsg.Show
      fpcmbExOrder.SetFocus
      Close SHandle
      Exit Sub
    End If
    ReDim IdxArray(1 To NumOfSrchRecs) As Long
    UseNameIdx = 2
    For x = 1 To NumOfSrchRecs
      Get SHandle, x, SearchRec
      IdxArray(x) = SearchRec.CustRec
    Next x
    Close SHandle
    NumOfRecs = NumOfSrchRecs
  Else
    NumOfRecs = NumOfTaxRecs
  End If
  
  If QPTrim$(fpcmbIncAbsYN.Text) = "Y" Then
    DetailFlag = True
  End If

  RptHandle = FreeFile

  Open ReportFile$ For Output As #RptHandle

  FileName$ = QPTrim$(fptxtFile.Text) + ".TXT"
  
  RptOut = FreeFile
  Open FileName$ For Output As #RptOut

  OpenPersPropFile PersHandle, NumOfPersRecs
  OpenRealPropFile RealHandle, NumOfRealRecs
  
  frmVATaxShowPctComp.Label1 = "Gathering Export Data"
  frmVATaxShowPctComp.Show , Me
  EnableCloseButton Me.hwnd, False
  cmdProcess.Enabled = False
  cmdExit.Enabled = False
  
  For x = 1 To NumOfRecs 'NumOfTaxRecs
    If UseNameIdx > 0 Then
      CustRecNo = IdxArray(x)
    Else
      CustRecNo = x
    End If

    Get TaxHandle, CustRecNo, TaxRec

    If Not TaxRec.Deleted Then
        Print #RptOut, Using("#####", CustRecNo);
        Print #RptOut, "]";
        Print #RptOut, TaxRec.CustName; "]";
        Print #RptOut, (TaxRec.Addr1); "]";
        Print #RptOut, (TaxRec.Addr2); "]";
        Print #RptOut, (TaxRec.City); "]";
        Print #RptOut, (TaxRec.State); "]";
        Print #RptOut, (TaxRec.Zip); "]";
        Print #RptOut, (TaxRec.CSSN);
        If DetailFlag = False Then 'closed
          Print #RptOut, ""
          GoTo SkipEm
          Else
          Print #RptOut, "]";
        End If
        'Now Show Property Records Next

        If TaxRec.FirstPropRec > 0 Then 'closed
          PropertyRecord = TaxRec.FirstPropRec
          Do While PropertyRecord <> 0
            Get #RealHandle, PropertyRecord, RealRec
            If PropertyRecord <> TaxRec.FirstPropRec Then 'closed
                Print #RptOut, Using("#####", CustRecNo);
                Print #RptOut, "]";
'                Print #RptOut, (TaxRec.FName); "]";
'                Print #RptOut, (TaxRec.LName); "]";
                Print #RptOut, TaxRec.CustName; "]";
                Print #RptOut, (TaxRec.Addr1); "]";
                Print #RptOut, (TaxRec.Addr2); "]";
                Print #RptOut, (TaxRec.City); "]";
                Print #RptOut, (TaxRec.State); "]";
                Print #RptOut, (TaxRec.Zip); "]";
                Print #RptOut, (TaxRec.CSSN); "]";
'             ElseIf PropertyRecord = TaxRec.FirstPropRec Then
            End If

            Print #RptOut, RealRec.RealPin; "]";
            Print #RptOut, RealRec.PROPNOT1; "]";
            Print #RptOut, RealRec.PROPNOT2; "]";
            Print #RptOut, RealRec.PROPNOT3; "]";
            Print #RptOut, Using("##########", RealRec.PROPVALU);
            Print #RptOut, "]";
'            Print #RptOut, Using("##########", RealRec.EXMPSENI);
            Print #RptOut, Using("##########", RealRec.EXMPOTHR); '6/14/06
            Print #RptOut, "]";
            Print #RptOut, RealRec.MORTCODE; "]";
            Print #RptOut, RealRec.Map; "]";
            Print #RptOut, RealRec.BLOCK; "]";
            Print #RptOut, RealRec.LOTNUMB; "]";
            PropertyRecord = RealRec.NextRec
            Loop
        Else
          Print #RptOut, "]"; "]"; "]"; "]"; "]"; "]"; "]"; "]"; "]"; "]";
        End If


        'NOW CHECK PERSONAL PROPERTY

        If TaxRec.FirstPersRec > 0 Then 'closed
          PropertyRecord = TaxRec.FirstPersRec
           Do While PropertyRecord <> 0
           Get #PersHandle, PropertyRecord, PersRec
            Print #RptOut, PersRec.PropPin; "]";
            Print #RptOut, Using("########", 0); '6/14/06 no more pers exemptions PersRec.PersVal);
            Print #RptOut, "]";
            Print #RptOut, Using("########", 0); '6/14/06 no more pers exemptions PersRec.EXMPSENI);
            Print #RptOut, "]";
            Print #RptOut, PersRec.DESC1; "]";
            Print #RptOut, PersRec.DESC2; "]";
            Print #RptOut, PersRec.DESC3; "]";
            Print #RptOut, PersRec.Desc4; "]";
            Print #RptOut, PersRec.Desc5; "]"
            PropertyRecord = PersRec.NextRec
           Loop
'            If OldRecord! = PropertyRecord! Then PropertyRecord! = 0
          Else
           Print #RptOut, " ]"; " ]"; " ]"; " ]"; " ]"; " ]"; " ]"; " ]"
          End If
     End If
SkipEm:
     CustCnt = CustCnt + 1
     frmVATaxShowPctComp.ShowPctComp x, NumOfTaxRecs
     If frmVATaxShowPctComp.Out = True Then
       Close
       frmVATaxShowPctComp.Out = False
       Unload frmVATaxShowPctComp
       EnableCloseButton Me.hwnd, True
       cmdProcess.Enabled = True
       cmdExit.Enabled = True
       Exit Sub
     End If
  Next
  Unload frmVATaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdProcess.Enabled = True
  cmdExit.Enabled = True
  
  If CustCnt = 0 Then
    frmVATaxMsg.Label1.Caption = "There are no export files to print."
    frmVATaxMsg.Label1.Top = 900
    frmVATaxMsg.Show vbModal
    Close
    Exit Sub
  End If
  
  Print #RptHandle, DetailFlag
  Close

  arVATaxExportRpt.Show
  
  Exit Sub
  
ERRORSTUFF:
   Unload frmVATaxShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxExport", "PrintGraphics", Erl)
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

Private Sub fptxtFile_Change()
  If InStr(fptxtFile.Text, ".") Then
    Call TaxMsg(900, "This file will be saved as a .txt file by default. Please do not use a '.' in the name.")
    fptxtFile.Text = ReplaceString(fptxtFile.Text, ".", "")
    Exit Sub
  End If
End Sub
