VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmFAPrePostPrint 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fixed Assets Pre-Depreciation Post Reports"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmFAPrePostPrint.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   5850
      Left            =   1928
      TabIndex        =   0
      Top             =   1433
      Width           =   7785
      _Version        =   196609
      _ExtentX        =   13732
      _ExtentY        =   10319
      _StockProps     =   70
      BackColor       =   13684944
      Caption         =   ""
      FrameColor      =   -2147483630
      FrameThreeDStyle=   1
      FrameWidth      =   2
      Picture         =   "frmFAPrePostPrint.frx":08CA
      Begin LpLib.fpCombo fpcomboPrintOpt 
         Height          =   405
         Left            =   3600
         TabIndex        =   2
         ToolTipText     =   "Select Graphic for a robust report that takes more time to process. Select Text for a faster report."
         Top             =   2355
         Width           =   2370
         _Version        =   196608
         _ExtentX        =   4180
         _ExtentY        =   714
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
         AutoSearch      =   2
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
         AutoSearchFill  =   -1  'True
         AutoSearchFillDelay=   200
         EditMarginLeft  =   1
         EditMarginTop   =   1
         EditMarginRight =   0
         EditMarginBottom=   3
         ResizeRowToFont =   0   'False
         TextTipMultiLine=   0
         AutoMenu        =   -1  'True
         EditAlignH      =   1
         EditAlignV      =   0
         ColDesigner     =   "frmFAPrePostPrint.frx":08E6
      End
      Begin LpLib.fpCombo fpcmbType 
         Height          =   405
         Left            =   3090
         TabIndex        =   1
         ToolTipText     =   "Select the order this report will display data."
         Top             =   1530
         Width           =   3240
         _Version        =   196608
         _ExtentX        =   5715
         _ExtentY        =   714
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
         Object.TabStop         =   0   'False
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
         AutoSearch      =   2
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
         MaxEditLen      =   5
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
         AutoSearchFill  =   -1  'True
         AutoSearchFillDelay=   200
         EditMarginLeft  =   1
         EditMarginTop   =   1
         EditMarginRight =   0
         EditMarginBottom=   3
         ResizeRowToFont =   0   'False
         TextTipMultiLine=   0
         AutoMenu        =   -1  'True
         EditAlignH      =   1
         EditAlignV      =   0
         ColDesigner     =   "frmFAPrePostPrint.frx":0BDD
      End
      Begin EditLib.fpCurrency fpcurrLeast 
         Height          =   390
         Left            =   2880
         TabIndex        =   3
         ToolTipText     =   $"frmFAPrePostPrint.frx":0ED4
         Top             =   3840
         Width           =   2295
         _Version        =   196608
         _ExtentX        =   4048
         _ExtentY        =   688
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
         AlignTextH      =   2
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   -1  'True
         AutoBeep        =   0   'False
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
         Text            =   "$0.00"
         CurrencyDecimalPlaces=   -1
         CurrencyNegFormat=   0
         CurrencyPlacement=   0
         CurrencySymbol  =   ""
         DecimalPoint    =   ""
         FixedPoint      =   -1  'True
         LeadZero        =   0
         MaxValue        =   "9000000000"
         MinValue        =   "-9000000000"
         NegToggle       =   0   'False
         Separator       =   ""
         UseSeparator    =   0   'False
         IncInt          =   1
         IncDec          =   1
         BorderGrayAreaColor=   -2147483637
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
         Height          =   690
         Left            =   1590
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Click this button to create the desired report."
         Top             =   4680
         Width           =   1875
         _Version        =   131072
         _ExtentX        =   3307
         _ExtentY        =   1217
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   0   'False
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
         ButtonDesigner  =   "frmFAPrePostPrint.frx":0F66
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
         Height          =   684
         Left            =   4560
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Click this button to create the report based on the parameters entered above."
         Top             =   4680
         Width           =   1884
         _Version        =   131072
         _ExtentX        =   3323
         _ExtentY        =   1206
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   0   'False
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
         ButtonDesigner  =   "frmFAPrePostPrint.frx":1142
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
         Caption         =   "Print Option:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1920
         TabIndex        =   7
         Top             =   2445
         Width           =   1500
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Pre-Depreciation Post Reports"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   375
         Left            =   1590
         TabIndex        =   6
         Top             =   570
         Width           =   4815
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   684
         Left            =   1536
         Top             =   432
         Width           =   4908
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Enter the minimum original value to apply to fixed assets with depreciation flags set to 'No'"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   840
         TabIndex        =   5
         Top             =   3120
         Width           =   6345
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Report Type:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1440
         TabIndex        =   4
         Top             =   1590
         Width           =   1470
      End
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   6075
      Left            =   1830
      Top             =   1328
      Width           =   7980
   End
End
Attribute VB_Name = "frmFAPrePostPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsFATextBoxOverRider
  Private Temp_Class As Resize_Class

Private Sub cmdExit_Click()
  frmFAYearEndMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdProcess_Click()
  If fpcmbType.Text = "Department" Then
    If fpcomboPrintOpt.Text = "Graphical" Then
      Call PrintGraphicsByDept
    ElseIf fpcomboPrintOpt.Text = "Text" Then
      MsgBox "Pitch 17 is recommended for this report."
      Call PrintTextByDept
    Else
      Exit Sub
    End If
  ElseIf fpcmbType.Text = "Fund/Asset" Then
    If fpcomboPrintOpt.Text = "Graphical" Then
      Call PrintGraphicsByFundAsset
    ElseIf fpcomboPrintOpt.Text = "Text" Then
      MsgBox "Pitch 17 is recommended for this report."
      Call PrintTextByFundAsset
    Else
      Exit Sub
    End If
  End If
  
End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsFATextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Call LoadMe
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
      SendKeys "%C"
      Call cmdExit_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%P"
      Call cmdProcess_Click
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
      MainLog ("FixedAssets.exe terminated via menu bar on frmFAPrePostPrint.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub LoadMe()
  fpcmbType.Text = "Department"
  fpcmbType.AddItem "Department"
  fpcmbType.AddItem "Fund/Asset"
  fpcomboPrintOpt.AddItem "Graphical"
  fpcomboPrintOpt.AddItem "Text"
  fpcomboPrintOpt.Text = "Graphical"
  
End Sub

Private Sub fpcmbType_KeyDown(KeyCode As Integer, Shift As Integer)
  'this prevents the user from inadvertently changing data in the combo box when
  'tabbing through the fields
  If KeyCode = vbKeySpace Then
    fpcmbType.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbType.ListIndex = -1
  End If
  If fpcmbType.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcomboPrintOpt.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcomboPrintOpt_KeyDown(KeyCode As Integer, Shift As Integer)
  'this prevents the user from inadvertently changing data in the combo box when
  'tabbing through the fields
  If KeyCode = vbKeySpace Then
    fpcomboPrintOpt.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcomboPrintOpt.ListIndex = -1
  End If
  If fpcomboPrintOpt.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcurrLeast.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If
End Sub

Private Sub PrintGraphicsByDept()
  Dim DOrigCost#, DBookTotal#, DCDep#, DYDep#, OrigCost#, BookTotal#, CDep#, YDep#
  Dim YrFile As Integer
  Dim FAYear(1) As FAYearEndType
  Dim YearRecNum As Integer
  Dim LastYr$
  Dim ReportFile$
  Dim CurDep#
  Dim ItemCnt&
  Dim RptHandle As Integer
  Dim FAFile As Integer
  Dim FAItemRec As FAItemRecType
  Dim DepFile As Integer
  Dim FADep(1) As FADepFileType
  Dim NumOfDepRecs As Integer
  Dim cnt&
  Dim ItemRecNo As Long
  Dim DeptNumber As Integer
  Dim DCurDep#
  Dim YTDDep#
  Dim NumOfFARecs As Integer
  Dim dlm$, x As Integer
  Dim Employer$
  Dim FASetUpRec As FASetupRecType
  Dim FASHandle As Integer
  Dim DeptIdx As DeptNumbSortIdxType
  Dim DIdxHandle As Integer
  Dim DIdxRecNums As Integer
  Dim DeptDesc$, Dpr4Year$
  Dim AccuDpr As Double '9/21/2004
  
  On Error GoTo ERRORSTUFF
  
  OpenDeptIdxFile DIdxHandle
  DIdxRecNums = LOF(DIdxHandle) / Len(DeptIdx)
  If DIdxRecNums > 0 Then
    ReDim DeptIndx(1 To DIdxRecNums) As String
    ReDim DeptNum(1 To DIdxRecNums) As Integer
    For x = 1 To DIdxRecNums
      Get DIdxHandle, x, DeptIdx
      DeptIndx(x) = QPTrim$(DeptIdx.DeptIdxDesc)
      DeptNum(x) = DeptIdx.DeptNumb
    Next x
    Close DIdxHandle
  End If
  
  dlm$ = "~"
  OpenFASetUpFile FASHandle
  Get FASHandle, 1, FASetUpRec
  Close FASHandle
  
  Employer$ = FASetUpRec.TownName
  OpenYearFile YrFile
  YearRecNum = LOF(YrFile) / Len(FAYear(1))
  If YearRecNum = 0 Then
    LastYr$ = "N/A"
  Else
    Get YrFile, 1, FAYear(1)
    LastYr$ = FAYear(1).CurYear
  End If
  Close YrFile
  
  ReportFile$ = "FARPTS\FADEPEDT.RPT"  'Report File Name
  ItemCnt& = 0
  
  RptHandle = FreeFile
  
  Open ReportFile$ For Output As #RptHandle
  
  OpenFAItemFile FAFile
  NumOfFARecs = LOF(FAFile) / Len(FAItemRec)
  
  'Open Deprec Edit File
  OpenDeprEditFile DepFile
  NumOfDepRecs = LOF(DepFile) / Len(FADep(1))
  Get DepFile, 1, FADep(1)
  Dpr4Year$ = QPTrim$(FADep(1).CurrYear)
  
  For cnt& = 1 To NumOfDepRecs
    Get DepFile, cnt&, FADep(1)
    ItemRecNo = FADep(1).AssetRecord
    Get FAFile, ItemRecNo, FAItemRec
    If FAItemRec.ORGCOST < fpcurrLeast And FAItemRec.DEPYN = "N" Then GoTo SkipEm3
    If cnt& = 1 Then
      DeptNumber = FAItemRec.IDEPT
    End If
    
    If DIdxRecNums > 0 Then
      For x = 1 To DIdxRecNums
        If DeptNum(x) = FAItemRec.IDEPT Then
          DeptDesc = QPTrim$(DeptIndx(x))
          Exit For
        End If
      Next x
    End If
    
    If DeptNumber <> FAItemRec.IDEPT Then 'reached the point where
'    'dept totals can be printed
      DeptNumber = FAItemRec.IDEPT
      DOrigCost# = 0
      DCurDep# = 0
      DYDep# = 0
    End If
    'Figure Values
    'Calc Depreciation for This Period
    YTDDep# = FAItemRec.DEP2DATE
    AccuDpr = 0
    AccuDpr = OldRound(FADep(1).CurYrDep + FAItemRec.DEP2DATE)
    '                     0                   1
    Print #RptHandle, Employer; dlm; FAItemRec.ItemTag; dlm;
    '                         2                     3
    Print #RptHandle, FAItemRec.IDESC1; dlm; FAItemRec.IDEPT; dlm;
    '                         4                     5                       6
    Print #RptHandle, FAItemRec.ILIFE; dlm; FAItemRec.ORGCOST; dlm; FAItemRec.DEP2DATE; dlm;
    '                         7                    8                  9                    10
    Print #RptHandle, FADep(1).CurYrDep; dlm; DeptNumber; dlm; FADep(1).CurrYear; dlm; DeptDesc; dlm;
    If FADep(1).PctFlag Then
      '                  11
      Print #RptHandle, "*"; dlm;
    Else
      '                  11
      Print #RptHandle, " "; dlm;
    End If
    '                    12             13
    Print #RptHandle, Dpr4Year$; dlm; AccuDpr
    'SubTotal Here
    ItemCnt& = ItemCnt& + 1
    'Grand Totals Here
    OrigCost# = OrigCost# + FAItemRec.ORGCOST
    CurDep# = CurDep# + FADep(1).CurYrDep
    YDep# = YDep# + YTDDep#
    'Dept Totals Here
    DOrigCost# = DOrigCost# + FAItemRec.ORGCOST
    DCurDep# = DCurDep# + FADep(1).CurYrDep
    DYDep# = DYDep# + YTDDep#
    
    
SkipEm3:
  Next cnt&
  Close         'Close all open files now
  
  arFADprBeforePostRpt.Show
  frmFALoadReport.Show
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFAYearEndMenu", "PrintGraphics", Erl)
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
    ClearInUse (PWcnt)
    Terminate
    Unload Me
End Sub

Private Sub PrintTextByFundAsset()
  Dim DOrigCost#, DBookTotal#, DCDep#, DYDep#, OrigCost#, BookTotal#, CDep#, YDep#, TAccuDpr#
  Dim YrFile As Integer
  Dim FAYear(1) As FAYearEndType
  Dim YearRecNum As Integer
  Dim LastYr$
  Dim ReportFile$
  Dim Dash80$
  Dim FF$, CurDep#
  Dim MaxLines As Integer
  Dim LineCnt&, ItemCnt&
  Dim RptHandle As Integer
  Dim FAFile As Integer
  Dim FAItemRec As FAItemRecType
  Dim DepFile As Integer
  Dim FADep(1) As FADepFileType
  Dim NumOfDepRecs As Integer
  Dim cnt&, Page As Integer
  Dim ItemRecNo As Long
  Dim FndAssNumber$ ' As Integer
  Dim DCurDep#
  Dim YTDDep#
  Dim NumOfFARecs As Integer
  Dim x As Integer
  Dim DItemCnt&
  Dim DAccuDpr As Double
  Dim Employer$
  Dim FASetUpRec As FASetupRecType
  Dim FASHandle As Integer
  Dim AccuDpr As Double '9/21/2004
  Dim FundAssetSort$
  Dim Big$
  Dim ThisBig$
  Dim HoldFundAsset$
  Dim HoldRec As Integer
  Dim Nextx As Integer
  Dim ThisCnt As Integer
  Dim ThisFundAsset$
  Dim FHandle As Integer
  Dim FundRec As FAFundCodeType
  Dim NumOfFundRecs As Integer
  Dim CHandle As Integer
  Dim CodeRec As FAAssetCodeRecType
  Dim NumOfCodeRecs As Integer
  Dim ThisFndAssDesc$
  Dim MaxFndAss As Integer
  Dim HoldFndAssDesc$
  Dim HoldDesc$
  Dim HoldDpr As FADepFileType
  Dim ThisFADesc As Integer
  
  On Error GoTo ERRORSTUFF
  
  OpenFACodeNameFile CHandle
  NumOfCodeRecs = LOF(CHandle) / Len(CodeRec)
  If NumOfCodeRecs = 0 Then
    MsgBox "No asset codes can be found. Report printing aborted."
    Close
    Exit Sub
  End If
  
  ReDim CodeNum(1 To NumOfCodeRecs) As String
  ReDim CodeDesc(1 To NumOfCodeRecs) As String
  
  For x = 1 To NumOfCodeRecs
    Get CHandle, x, CodeRec
    CodeNum(x) = QPTrim$(CodeRec.ASSETCODE)
    CodeDesc(x) = QPTrim$(CodeRec.AssetDesc)
  Next x
  Close CHandle
  
  OpenFAFundCodeFile FHandle
  NumOfFundRecs = LOF(FHandle) / Len(FundRec)
  If NumOfFundRecs = 0 Then
    MsgBox "No fund codes can be found. Report printing aborted."
    Close
    Exit Sub
  End If
  
  ReDim FundNum(1 To NumOfFundRecs) As String
  ReDim FundDesc(1 To NumOfFundRecs) As String
  
  For x = 1 To NumOfFundRecs
    Get FHandle, x, FundRec
    FundNum(x) = CStr(FundRec.FundNum)
    FundDesc(x) = QPTrim$(FundRec.FundDesc)
  Next x
  Close FHandle
  
  OpenYearFile YrFile
  YearRecNum = LOF(YrFile) / Len(FAYear(1))
  If YearRecNum = 0 Then
    LastYr$ = "N/A"
  Else
    Get YrFile, 1, FAYear(1)
    LastYr$ = FAYear(1).CurYear
  End If
  Close YrFile
  
  ReportFile$ = "FADEPEDTFUNDASSET.PRN"  'Report File Name
  Dash80$ = String$(80, "=")
  FF$ = Chr$(12)
  
  OpenFASetUpFile FASHandle
  Get FASHandle, 1, FASetUpRec
  Close FASHandle
  
  Employer$ = FASetUpRec.TownName
  
  MaxLines = 53
  LineCnt& = 0
  ItemCnt& = 0
  DItemCnt& = 0
  
  RptHandle = FreeFile
  
  Open ReportFile$ For Output As #RptHandle
  
  OpenFAItemFile FAFile
  NumOfFARecs = LOF(FAFile) / Len(FAItemRec)
  
  'Open Deprec Edit File
  OpenDeprEditFile DepFile
  NumOfDepRecs = LOF(DepFile) / Len(FADep(1))
  Get DepFile, 1, FADep(1)
  GoSub PrintMasterHeader3
  If NumOfDepRecs = 0 Then
    Close
    MsgBox "No temporary depreciation records have been saved. Use the build depreciation feature to create these records."
    Exit Sub
  End If
  
  ReDim FundAsset(1 To NumOfDepRecs) As String 'make arrays that will
  'be used to sort fixed assets by fund and asset code
  ReDim FndAssdesc(1 To NumOfDepRecs) As String 'this array coincides
  'with each fundasset
  
  Nextx = 1
  ReDim SwapDepSort(1 To NumOfDepRecs) As FADepFileType
  For cnt& = 1 To NumOfDepRecs
    Get DepFile, cnt&, FADep(1) 'gather an array that holds only
    'each unique fund/asset number
    SwapDepSort(cnt) = FADep(1)
    ItemRecNo = FADep(1).AssetRecord
    Get FAFile, ItemRecNo, FAItemRec
    ThisFundAsset$ = CStr(FAItemRec.FundNum) + QPTrim$(FAItemRec.ASSETCODE)
    FundAsset(cnt) = ThisFundAsset
    For x = 1 To NumOfFundRecs
      If CStr(FAItemRec.FundNum) = FundNum(x) Then
        FndAssdesc(Nextx) = FundDesc(x)
        Exit For
      End If
    Next x
    For x = 1 To NumOfCodeRecs
      If QPTrim$(FAItemRec.ASSETCODE) = CodeNum(x) Then
        FndAssdesc(Nextx) = FndAssdesc(cnt) + "/" + CodeDesc(x)
        Exit For
      End If
    Next x
    Nextx = Nextx + 1
  Next cnt
  
  Big = ""
  For x = 1 To NumOfDepRecs
    If FundAsset(x) > Big Then
      Big = FundAsset(x)
    End If
  Next x
  
  Big = Big + "z"
  ThisBig = Big
  Nextx = 1
  Do
    For x = Nextx To NumOfDepRecs
      If FundAsset(x) < Big Then
        Big = FundAsset(x)
        ThisCnt = x
      End If
    Next x
    HoldDesc = FndAssdesc(Nextx)
    FndAssdesc(Nextx) = FndAssdesc(ThisCnt)
    FndAssdesc(ThisCnt) = HoldDesc
    HoldFundAsset = FundAsset(Nextx)
    FundAsset(Nextx) = FundAsset(ThisCnt)
    FundAsset(ThisCnt) = HoldFundAsset
    HoldDpr = SwapDepSort(ThisCnt)
    SwapDepSort(ThisCnt) = SwapDepSort(Nextx)
    SwapDepSort(Nextx) = HoldDpr
    Nextx = Nextx + 1
    If Nextx = NumOfDepRecs + 1 Then Exit Do
    Big = ThisBig
  Loop
  
  Nextx = 1
  For cnt& = 1 To NumOfDepRecs
    ItemRecNo = SwapDepSort(cnt).AssetRecord
    
    Get FAFile, ItemRecNo, FAItemRec
    If FAItemRec.ORGCOST < fpcurrLeast And FAItemRec.DEPYN = "N" Then GoTo SkipEm3
    If cnt& = 1 Then
      FndAssNumber = CStr(FAItemRec.FundNum) + QPTrim$(FAItemRec.ASSETCODE)
    End If
    
    If LineCnt& >= MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintMasterHeader3
    End If
    If FndAssNumber <> CStr(FAItemRec.FundNum) + QPTrim$(FAItemRec.ASSETCODE) Then 'data is being read in dept order
      'Print Subtotals and Clear
      Print #RptHandle, String$(122, "-")
      Print #RptHandle, "Totals for: "; FndAssNumber; ; "  "; FndAssdesc(cnt - 1); "  "; "#Items:"; DItemCnt;
      Print #RptHandle, Tab(64); Using("###,###,##0.00", DOrigCost#);
      Print #RptHandle, Tab(79); Using("###,###,##0.00", DYDep#);
      Print #RptHandle, Tab(93); Using("###,###,##0.00", DCurDep#);
      Print #RptHandle, Tab(109); Using("###,###,##0.00", DAccuDpr#)
      LineCnt& = LineCnt& + 2
      
      Print #RptHandle, "": LineCnt& = LineCnt& + 1
      Print #RptHandle, "": LineCnt& = LineCnt& + 1
      
      FndAssNumber = CStr(FAItemRec.FundNum) + QPTrim$(FAItemRec.ASSETCODE) 'FAItemRec.IDEPT
      DOrigCost# = 0
      DCurDep# = 0
      DYDep# = 0
      DItemCnt& = 0
      DAccuDpr = 0
    End If
    
    'Figure Values
    'Calc Depreciation for This Period
    YTDDep# = FAItemRec.DEP2DATE
    AccuDpr = 0
    AccuDpr = OldRound(SwapDepSort(cnt).CurYrDep + YTDDep#)
    Print #RptHandle, FAItemRec.ItemTag; Tab(22); Left$(FAItemRec.IDESC1, 28);
    Print #RptHandle, Tab(51); CStr(FAItemRec.FundNum) + "/" + QPTrim$(FAItemRec.ASSETCODE);
    Print #RptHandle, Tab(58); Using("###", FAItemRec.ILIFE);
    Print #RptHandle, Tab(64); Using("###,###,##0.00", FAItemRec.ORGCOST);
    Print #RptHandle, Tab(79); Using("###,###,##0.00", YTDDep#);
    Print #RptHandle, Tab(93); Using("###,###,##0.00", SwapDepSort(cnt).CurYrDep);
    If FADep(1).PctFlag Then
      Print #RptHandle, "*";
    End If
    Print #RptHandle, Tab(108); Using("###,###,##0.00#", AccuDpr#)
    'SubTotal Here
    LineCnt& = LineCnt& + 1
    ItemCnt& = ItemCnt& + 1
    DItemCnt& = DItemCnt& + 1
    'Grand Totals Here
    OrigCost# = OrigCost# + FAItemRec.ORGCOST
    CurDep# = CurDep# + SwapDepSort(cnt).CurYrDep
    YDep# = YDep# + YTDDep#
    TAccuDpr = TAccuDpr + AccuDpr
    'Fund/Asset Totals Here
    DOrigCost# = DOrigCost# + FAItemRec.ORGCOST
    DCurDep# = DCurDep# + SwapDepSort(cnt).CurYrDep
    DYDep# = DYDep# + YTDDep#
    DAccuDpr# = DAccuDpr# + AccuDpr#
    
SkipEm3:
  Next cnt&
  'First Print Subtotals
  
  Print #RptHandle, String$(122, "-")
  Print #RptHandle, "Totals for: "; FndAssNumber; ; "  "; FndAssdesc(cnt - 1); "  "; "#Items:"; DItemCnt;
  Print #RptHandle, Tab(64); Using("###,###,##0.00", DOrigCost#);
  Print #RptHandle, Tab(79); Using("###,###,##0.00", DYDep#);
  Print #RptHandle, Tab(93); Using("###,###,##0.00", DCurDep#);
  Print #RptHandle, Tab(109); Using("###,###,##0.00", DAccuDpr#)
  LineCnt& = LineCnt& + 2
  
  Print #RptHandle, "": LineCnt& = LineCnt& + 1
  Print #RptHandle, "": LineCnt& = LineCnt& + 1
  
  GoSub PrintDepRepEnding1
  Print #RptHandle, Chr$(18);   ' oki 320 10 cpi
  Close         'Close all open files now
  
  ViewPrint ReportFile$, "Current Depreciation Report", True
  KillFile (ReportFile$)
  
  Exit Sub
  
PrintMasterHeader3:
  Page = Page + 1
  Print #RptHandle, Tab(20); "Master Asset Listing : Depreciation Edit Report For "; FADep(1).CurrYear
  Print #RptHandle, Employer
  Print #RptHandle, "Report Date: "; Date$; Tab(68); "Page #"; Page
  Print #RptHandle, "Asset Number"; Tab(22); "Description"; Tab(51); "Fd/Ast"; Tab(58); "Life"; Tab(65); "Original Cost"; Tab(81); "Dprc To Date"; Tab(94); "Cur Yr Deprec"; Tab(113); "Accum Dprc"
  Print #RptHandle, String$(122, "=")
  LineCnt& = 5
  Return
  
PrintDepRepEnding1:
  Print #RptHandle, String$(122, "-")
  Print #RptHandle, "Grand Totals: "; Tab(15); "# Items: "; Tab(26); Using("######0", ItemCnt);
  Print #RptHandle, Tab(64); Using("###,###,##0.00", OrigCost#);
  Print #RptHandle, Tab(79); Using("###,###,##0.00", YDep#);
  Print #RptHandle, Tab(93); Using("###,###,##0.00", CurDep#);
  Print #RptHandle, Tab(109); Using("###,###,##0.00", TAccuDpr#)
  Print #RptHandle, FF$
  Return
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFAYearEndMenu", "PrintText", Erl)
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
    ClearInUse (PWcnt)
    Terminate
    Unload Me

End Sub

Private Sub PrintGraphicsByFundAsset()
  Dim DOrigCost#, DBookTotal#, DCDep#, DYDep#, OrigCost#, BookTotal#, CDep#, YDep#
  Dim YrFile As Integer
  Dim FAYear(1) As FAYearEndType
  Dim YearRecNum As Integer
  Dim LastYr$
  Dim ReportFile$
  Dim CurDep#
  Dim ItemCnt&
  Dim RptHandle As Integer
  Dim FAFile As Integer
  Dim FAItemRec As FAItemRecType
  Dim DepFile As Integer
  Dim FADep(1) As FADepFileType
  Dim NumOfDepRecs As Integer
  Dim cnt&
  Dim ItemRecNo As Long
  Dim FndAssNumber$ ' As Integer
  Dim DCurDep#
  Dim YTDDep#
  Dim NumOfFARecs As Integer
  Dim dlm$, x As Integer
  Dim Employer$
  Dim FASetUpRec As FASetupRecType
  Dim FASHandle As Integer
  Dim Dpr4Year$
  Dim AccuDpr As Double '9/21/2004
  Dim FundAssetSort$
  Dim Big$
  Dim ThisBig$
  Dim HoldFundAsset$
  Dim HoldRec As Integer
  Dim Nextx As Integer
  Dim ThisCnt As Integer
  Dim ThisFundAsset$
  Dim FHandle As Integer
  Dim FundRec As FAFundCodeType
  Dim NumOfFundRecs As Integer
  Dim CHandle As Integer
  Dim CodeRec As FAAssetCodeRecType
  Dim NumOfCodeRecs As Integer
'  Dim MaxFndAss As Integer
'  Dim NewRec As Integer
  Dim HoldFndAssDesc$
  Dim HoldDesc$
  Dim HoldDpr As FADepFileType
  
  On Error GoTo ERRORSTUFF
  
  OpenFACodeNameFile CHandle 'gather code data for later use when
  'needing a code description
  NumOfCodeRecs = LOF(CHandle) / Len(CodeRec)
  If NumOfCodeRecs = 0 Then
    MsgBox "No asset codes can be found. Report printing aborted."
    Close
    Exit Sub
  End If
  
  ReDim CodeNum(1 To NumOfCodeRecs) As String
  ReDim CodeDesc(1 To NumOfCodeRecs) As String
  
  For x = 1 To NumOfCodeRecs
    Get CHandle, x, CodeRec
    CodeNum(x) = QPTrim$(CodeRec.ASSETCODE)
    CodeDesc(x) = QPTrim$(CodeRec.AssetDesc)
  Next x
  Close CHandle
  
  'gather fund data for later use when needing a fund description
  OpenFAFundCodeFile FHandle
  NumOfFundRecs = LOF(FHandle) / Len(FundRec)
  If NumOfFundRecs = 0 Then
    MsgBox "No fund codes can be found. Report printing aborted."
    Close
    Exit Sub
  End If
    
  ReDim FundNum(1 To NumOfFundRecs) As String
  ReDim FundDesc(1 To NumOfFundRecs) As String
  
  For x = 1 To NumOfFundRecs
    Get FHandle, x, FundRec
    FundNum(x) = CStr(FundRec.FundNum)
    FundDesc(x) = QPTrim$(FundRec.FundDesc)
  Next x
  Close FHandle
  
  dlm$ = "~"
  OpenFASetUpFile FASHandle
  Get FASHandle, 1, FASetUpRec
  Close FASHandle
  
  Employer$ = FASetUpRec.TownName
  OpenYearFile YrFile
  YearRecNum = LOF(YrFile) / Len(FAYear(1))
  If YearRecNum = 0 Then
    LastYr$ = "N/A"
  Else
    Get YrFile, 1, FAYear(1)
    LastYr$ = FAYear(1).CurYear
  End If
  Close YrFile
  
  ReportFile$ = "FARPTS\FADEPEDTFNDASS.RPT"  'Report File Name
  ItemCnt& = 0
  
  RptHandle = FreeFile
  
  Open ReportFile$ For Output As #RptHandle
  
  OpenFAItemFile FAFile
  NumOfFARecs = LOF(FAFile) / Len(FAItemRec)
  
  'Open Deprec Edit File
  OpenDeprEditFile DepFile
  NumOfDepRecs = LOF(DepFile) / Len(FADep(1))
  Get DepFile, 1, FADep(1)
  Dpr4Year$ = QPTrim$(FADep(1).CurrYear)
  
  ReDim FundAsset(1 To NumOfDepRecs) As String 'make arrays that will
  'be used to sort fixed assets by fund and asset code
  ReDim FndAssdesc(1 To NumOfDepRecs) As String 'this array coincides
  'with each fundasset
  
  Nextx = 1
  ReDim SwapDepSort(1 To NumOfDepRecs) As FADepFileType
  For cnt& = 1 To NumOfDepRecs
    Get DepFile, cnt&, FADep(1) 'gather an array that holds only
    'each unique fund/asset number
    SwapDepSort(cnt) = FADep(1)
    ItemRecNo = FADep(1).AssetRecord
    Get FAFile, ItemRecNo, FAItemRec
    ThisFundAsset$ = CStr(FAItemRec.FundNum) + QPTrim$(FAItemRec.ASSETCODE)
    FundAsset(cnt) = ThisFundAsset
    For x = 1 To NumOfFundRecs
      If CStr(FAItemRec.FundNum) = FundNum(x) Then
        FndAssdesc(Nextx) = FundDesc(x)
        Exit For
      End If
    Next x
    For x = 1 To NumOfCodeRecs
      If QPTrim$(FAItemRec.ASSETCODE) = CodeNum(x) Then
        FndAssdesc(Nextx) = FndAssdesc(cnt) + "/" + CodeDesc(x)
        Exit For
      End If
    Next x
    Nextx = Nextx + 1
  Next cnt
  
  Big = ""
  For x = 1 To NumOfDepRecs
    If FundAsset(x) > Big Then
      Big = FundAsset(x)
    End If
  Next x
  
  Big = Big + "z"
  ThisBig = Big
  Nextx = 1
  Do
    For x = Nextx To NumOfDepRecs
      If FundAsset(x) < Big Then
        Big = FundAsset(x)
        ThisCnt = x
      End If
    Next x
    HoldDesc = FndAssdesc(Nextx)
    FndAssdesc(Nextx) = FndAssdesc(ThisCnt)
    FndAssdesc(ThisCnt) = HoldDesc
    HoldFundAsset = FundAsset(Nextx)
    FundAsset(Nextx) = FundAsset(ThisCnt)
    FundAsset(ThisCnt) = HoldFundAsset
    HoldDpr = SwapDepSort(ThisCnt)
    SwapDepSort(ThisCnt) = SwapDepSort(Nextx)
    SwapDepSort(Nextx) = HoldDpr
    Nextx = Nextx + 1
    If Nextx = NumOfDepRecs + 1 Then Exit Do
    Big = ThisBig
  Loop
  
  Nextx = 1
  For cnt& = 1 To NumOfDepRecs
    ItemRecNo = SwapDepSort(cnt).AssetRecord
    Get FAFile, ItemRecNo, FAItemRec
    If FAItemRec.ORGCOST < fpcurrLeast And FAItemRec.DEPYN = "N" Then GoTo SkipEm3
    If cnt& = 1 Then
      FndAssNumber = CStr(FAItemRec.FundNum) + QPTrim$(FAItemRec.ASSETCODE)
    End If
    
    If FndAssNumber <> CStr(FAItemRec.FundNum) + QPTrim$(FAItemRec.ASSETCODE) Then 'reached the point where
'    'dept totals can be printed
      FndAssNumber = CStr(FAItemRec.FundNum) + QPTrim$(FAItemRec.ASSETCODE)
      DOrigCost# = 0
      DCurDep# = 0
      DYDep# = 0
      Nextx = Nextx + 1
'      DAccuDpr = 0
    End If
    'Figure Values
    'Calc Depreciation for This Period
    YTDDep# = FAItemRec.DEP2DATE
    AccuDpr = 0
    AccuDpr = OldRound(SwapDepSort(cnt).CurYrDep + FAItemRec.DEP2DATE)
    '                     0                   1
    Print #RptHandle, Employer; dlm; FAItemRec.ItemTag; dlm;
    '                         2                     3
    Print #RptHandle, FAItemRec.IDESC1; dlm; CStr(FAItemRec.FundNum) + QPTrim$(FAItemRec.ASSETCODE); dlm;
    '                         4                     5                       6
    Print #RptHandle, FAItemRec.ILIFE; dlm; FAItemRec.ORGCOST; dlm; FAItemRec.DEP2DATE; dlm;
    '                         7                    8                  9                    10
    Print #RptHandle, SwapDepSort(cnt).CurYrDep; dlm; FndAssNumber; dlm; SwapDepSort(cnt).CurrYear; dlm; FndAssdesc(cnt); dlm;
    If FADep(1).PctFlag Then
      '                  11
      Print #RptHandle, "*"; dlm;
    Else
      '                  11
      Print #RptHandle, " "; dlm;
    End If
    '                    12             13
    Print #RptHandle, Dpr4Year$; dlm; AccuDpr
    'SubTotal Here
    ItemCnt& = ItemCnt& + 1
    'Grand Totals Here
    OrigCost# = OrigCost# + FAItemRec.ORGCOST
    CurDep# = CurDep# + SwapDepSort(cnt).CurYrDep
    YDep# = YDep# + YTDDep#
    'Dept Totals Here
    DOrigCost# = DOrigCost# + FAItemRec.ORGCOST
    DCurDep# = DCurDep# + SwapDepSort(cnt).CurYrDep
    DYDep# = DYDep# + YTDDep#
    
    
SkipEm3:
  Next cnt&
  Close         'Close all open files now
  
  arFADprBeforePostRptFndAss.Show
  frmFALoadReport.Show
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFAYearEndMenu", "PrintGraphics", Erl)
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
    ClearInUse (PWcnt)
    Terminate
    Unload Me
End Sub

Private Sub PrintTextByDept()
  Dim DOrigCost#, DBookTotal#, DCDep#, DYDep#, OrigCost#, BookTotal#, CDep#, YDep#, TAccuDpr#
  Dim YrFile As Integer
  Dim FAYear(1) As FAYearEndType
  Dim YearRecNum As Integer
  Dim LastYr$
  Dim ReportFile$
  Dim Dash80$
  Dim FF$, CurDep#
  Dim MaxLines As Integer
  Dim LineCnt&, ItemCnt&
  Dim RptHandle As Integer
  Dim FAFile As Integer
  Dim FAItemRec As FAItemRecType
  Dim DepFile As Integer
  Dim FADep(1) As FADepFileType
  Dim NumOfDepRecs As Integer
  Dim cnt&, Page As Integer
  Dim ItemRecNo As Long
  Dim DeptNumber As Integer
  Dim DCurDep#
  Dim YTDDep#
  Dim NumOfFARecs As Integer
  Dim DeptIdx As DeptNumbSortIdxType
  Dim DIdxHandle As Integer
  Dim DIdxRecNums As Integer
  Dim DeptDesc$, x As Integer
  Dim DItemCnt&
  Dim AccuDpr As Double
  Dim DAccuDpr As Double
  Dim Employer$
  Dim FASetUpRec As FASetupRecType
  Dim FASHandle As Integer
  
  On Error GoTo ERRORSTUFF
  
  OpenDeptIdxFile DIdxHandle
  DIdxRecNums = LOF(DIdxHandle) / Len(DeptIdx)
  If DIdxRecNums > 0 Then
    ReDim DeptIndx(1 To DIdxRecNums) As String
    ReDim DeptNum(1 To DIdxRecNums) As Integer
    For x = 1 To DIdxRecNums
      Get DIdxHandle, x, DeptIdx
      DeptIndx(x) = QPTrim$(DeptIdx.DeptIdxDesc)
      DeptNum(x) = DeptIdx.DeptNumb
    Next x
    Close DIdxHandle
  End If
  
  OpenYearFile YrFile
  YearRecNum = LOF(YrFile) / Len(FAYear(1))
  If YearRecNum = 0 Then
    LastYr$ = "N/A"
  Else
    Get YrFile, 1, FAYear(1)
    LastYr$ = FAYear(1).CurYear
  End If
  Close YrFile
  
  ReportFile$ = "FADEPEDT.PRN"  'Report File Name
  Dash80$ = String$(80, "=")
  FF$ = Chr$(12)
  
  OpenFASetUpFile FASHandle
  Get FASHandle, 1, FASetUpRec
  Close FASHandle
  
  Employer$ = FASetUpRec.TownName
  
  MaxLines = 53
  LineCnt& = 0
  ItemCnt& = 0
  DItemCnt& = 0
  
  RptHandle = FreeFile
  
  Open ReportFile$ For Output As #RptHandle
  
  OpenFAItemFile FAFile
  NumOfFARecs = LOF(FAFile) / Len(FAItemRec)
  
  'Open Deprec Edit File
  OpenDeprEditFile DepFile
  NumOfDepRecs = LOF(DepFile) / Len(FADep(1))
  Get DepFile, 1, FADep(1)
  GoSub PrintMasterHeader3
  If NumOfDepRecs = 0 Then
    Close
    MsgBox "No temporary depreciation records have been saved. Use the build depreciation feature to create these records."
    Exit Sub
  End If
  
  For cnt& = 1 To NumOfDepRecs
    Get DepFile, cnt&, FADep(1)
    ItemRecNo = FADep(1).AssetRecord
    Get FAFile, ItemRecNo, FAItemRec
    If FAItemRec.ORGCOST < fpcurrLeast And FAItemRec.DEPYN = "N" Then GoTo SkipEm3
    If cnt& = 1 Then
      DeptNumber = FAItemRec.IDEPT
    End If
    
    If DIdxRecNums > 0 Then
      For x = 1 To DIdxRecNums
        If DeptNum(x) = FAItemRec.IDEPT Then
          DeptDesc = QPTrim$(DeptIndx(x))
          Exit For
        End If
      Next x
    End If
    
    If LineCnt& >= MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintMasterHeader3
    End If
    If DeptNumber <> FAItemRec.IDEPT Then 'data is being read in dept order
      'Print Subtotals and Clear
      Print #RptHandle, String$(122, "-")
      Print #RptHandle, "Totals for Dept Number: "; DeptNumber; ; "  "; DeptDesc; "  "; "#Items:"; DItemCnt;
      Print #RptHandle, Tab(64); Using("###,###,##0.00", DOrigCost#);
      Print #RptHandle, Tab(79); Using("###,###,##0.00", DYDep#);
      Print #RptHandle, Tab(93); Using("###,###,##0.00", DCurDep#);
      Print #RptHandle, Tab(109); Using("###,###,##0.00", DAccuDpr#)
      LineCnt& = LineCnt& + 2
      
      Print #RptHandle, "": LineCnt& = LineCnt& + 1
      Print #RptHandle, "": LineCnt& = LineCnt& + 1
      
      DeptNumber = FAItemRec.IDEPT
      DOrigCost# = 0
      DCurDep# = 0
      DYDep# = 0
      DItemCnt& = 0
      DAccuDpr = 0
    End If
    
    'Figure Values
    'Calc Depreciation for This Period
'SkipThisDeptTotal:
    YTDDep# = FAItemRec.DEP2DATE
    AccuDpr = 0
    AccuDpr = OldRound(FADep(1).CurYrDep + YTDDep#)
    Print #RptHandle, FAItemRec.ItemTag; Tab(22); Left$(FAItemRec.IDESC1, 28);
    Print #RptHandle, Tab(51); FAItemRec.IDEPT;
    Print #RptHandle, Tab(58); Using("###", FAItemRec.ILIFE);
    Print #RptHandle, Tab(64); Using("###,###,##0.00", FAItemRec.ORGCOST);
    Print #RptHandle, Tab(79); Using("###,###,##0.00", YTDDep#);
    Print #RptHandle, Tab(93); Using("###,###,##0.00", FADep(1).CurYrDep);
    If FADep(1).PctFlag Then
      Print #RptHandle, "*";
    End If
    Print #RptHandle, Tab(108); Using("###,###,##0.00#", AccuDpr#)
    'SubTotal Here
    LineCnt& = LineCnt& + 1
    ItemCnt& = ItemCnt& + 1
    DItemCnt& = DItemCnt& + 1
    'Grand Totals Here
    OrigCost# = OrigCost# + FAItemRec.ORGCOST
    CurDep# = CurDep# + FADep(1).CurYrDep
    YDep# = YDep# + YTDDep#
    TAccuDpr = TAccuDpr + AccuDpr
    'Dept Totals Here
    DOrigCost# = DOrigCost# + FAItemRec.ORGCOST
    DCurDep# = DCurDep# + FADep(1).CurYrDep
    DYDep# = DYDep# + YTDDep#
    DAccuDpr# = DAccuDpr# + AccuDpr#
    
SkipEm3:
  Next cnt&
  'First Print Subtotals
  
'  Print #RptHandle, String$(105, "-")
  Print #RptHandle, String$(122, "-")
  Print #RptHandle, "Totals for Dept Number: "; DeptNumber; ; "  "; DeptDesc; "  "; "#Items:"; DItemCnt;
  Print #RptHandle, Tab(64); Using("###,###,##0.00", DOrigCost#);
  Print #RptHandle, Tab(79); Using("###,###,##0.00", DYDep#);
  Print #RptHandle, Tab(93); Using("###,###,##0.00", DCurDep#);
  Print #RptHandle, Tab(109); Using("###,###,##0.00", DAccuDpr#)
  LineCnt& = LineCnt& + 2
  
  Print #RptHandle, "": LineCnt& = LineCnt& + 1
  Print #RptHandle, "": LineCnt& = LineCnt& + 1
  
  GoSub PrintDepRepEnding1
  Print #RptHandle, Chr$(18);   ' oki 320 10 cpi
  Close         'Close all open files now
  
  ViewPrint ReportFile$, "Current Depreciation Report", True
  KillFile (ReportFile$)
  
  Exit Sub
  
PrintMasterHeader3:
  Page = Page + 1
  Print #RptHandle, Tab(20); "Master Asset Listing : Depreciation Edit Report For "; FADep(1).CurrYear
  Print #RptHandle, Employer
  Print #RptHandle, "Report Date: "; Date$; Tab(68); "Page #"; Page
  Print #RptHandle, "Asset Number"; Tab(22); "Description"; Tab(51); "Dept"; Tab(58); "Life"; Tab(65); "Original Cost"; Tab(81); "Dprc To Date"; Tab(94); "Cur Yr Deprec"; Tab(113); "Accum Dprc"
'  Print #RptHandle, String$(105, "=")
  Print #RptHandle, String$(122, "=")
  LineCnt& = 5
  Return
  
PrintDepRepEnding1:
'  Print #RptHandle, String$(105, "-")
  Print #RptHandle, String$(122, "-")
  Print #RptHandle, "Grand Totals: "; Tab(15); "# Items: "; Tab(26); Using("######0", ItemCnt);
  Print #RptHandle, Tab(64); Using("###,###,##0.00", OrigCost#);
  Print #RptHandle, Tab(79); Using("###,###,##0.00", YDep#);
  Print #RptHandle, Tab(93); Using("###,###,##0.00", CurDep#);
  Print #RptHandle, Tab(109); Using("###,###,##0.00", TAccuDpr#)
  Print #RptHandle, FF$
  Return
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFAYearEndMenu", "PrintText", Erl)
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
    ClearInUse (PWcnt)
    Terminate
    Unload Me

End Sub

