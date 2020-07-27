VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "BTN32A20.OCX"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "IMP32X30.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "EDT32X30.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Begin VB.Form frmBLTransJrnlByCat 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Business License Transactions by Category"
   ClientHeight    =   8868
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   11652
   Icon            =   "frmBLTransJrnlByCust.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   11652
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   6564
      Left            =   1920
      TabIndex        =   6
      Top             =   1152
      Width           =   7788
      _Version        =   196609
      _ExtentX        =   13737
      _ExtentY        =   11578
      _StockProps     =   70
      Caption         =   ""
      FrameColor      =   -2147483630
      FrameThreeDStyle=   1
      FrameWidth      =   2
      Picture         =   "frmBLTransJrnlByCust.frx":08CA
      Begin LpLib.fpCombo fpcmbCatCode 
         Height          =   384
         Left            =   2256
         TabIndex        =   0
         ToolTipText     =   "Choose one of the transaction types on which to report."
         Top             =   2016
         Width           =   4812
         _Version        =   196608
         _ExtentX        =   8488
         _ExtentY        =   677
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
         ColDesigner     =   "frmBLTransJrnlByCust.frx":08E6
      End
      Begin LpLib.fpCombo fpcmbPrintOpt 
         Height          =   384
         Left            =   2976
         TabIndex        =   3
         ToolTipText     =   "Select graphical to print on a laser printer or choose text to print on a dot matrix printer."
         Top             =   4080
         Width           =   3564
         _Version        =   196608
         _ExtentX        =   6286
         _ExtentY        =   677
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
         ColDesigner     =   "frmBLTransJrnlByCust.frx":0BA5
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdExit 
         Height          =   684
         Left            =   1920
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Press to exit this screen."
         Top             =   5376
         Width           =   1884
         _Version        =   131072
         _ExtentX        =   3323
         _ExtentY        =   1206
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
         ButtonDesigner  =   "frmBLTransJrnlByCust.frx":0E64
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
         Height          =   684
         Left            =   4176
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Press to activate the report."
         Top             =   5376
         Width           =   1884
         _Version        =   131072
         _ExtentX        =   3323
         _ExtentY        =   1206
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
         ButtonDesigner  =   "frmBLTransJrnlByCust.frx":1042
      End
      Begin EditLib.fpDateTime fptxtBDate 
         Height          =   348
         Left            =   4128
         TabIndex        =   1
         ToolTipText     =   "Enter the date for which the report will begin it's report."
         Top             =   2784
         Width           =   1740
         _Version        =   196608
         _ExtentX        =   3069
         _ExtentY        =   614
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
         BackColor       =   16777215
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
         ButtonStyle     =   2
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483633
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   0
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   -1  'True
         AutoBeep        =   0   'False
         CaretInsert     =   0
         CaretOverWrite  =   3
         UserEntry       =   0
         HideSelection   =   -1  'True
         InvalidColor    =   12648447
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
         Text            =   "11/20/2002"
         DateCalcMethod  =   0
         DateTimeFormat  =   5
         UserDefinedFormat=   "mm/dd/yyyy"
         DateMax         =   "00000000"
         DateMin         =   "00000000"
         TimeMax         =   "000000"
         TimeMin         =   "000000"
         TimeString1159  =   ""
         TimeString2359  =   ""
         DateDefault     =   "00000000"
         TimeDefault     =   "000000"
         TimeStyle       =   0
         BorderGrayAreaColor=   -2147483637
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483633
         Appearance      =   2
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         PopUpType       =   1
         DateCalcY2KSplit=   60
         CaretPosition   =   0
         IncYear         =   1
         IncMonth        =   1
         IncDay          =   1
         IncHour         =   1
         IncMinute       =   1
         IncSecond       =   1
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         StartMonth      =   4
         ButtonAlign     =   0
         BoundDataType   =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpDateTime fptxtEDate 
         Height          =   348
         Left            =   4128
         TabIndex        =   2
         ToolTipText     =   "This date, themost current depreciation date, is automatically calculated and cannot be edited."
         Top             =   3360
         Width           =   1740
         _Version        =   196608
         _ExtentX        =   3069
         _ExtentY        =   614
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
         ButtonStyle     =   2
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483633
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   0
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   -1  'True
         AutoBeep        =   0   'False
         CaretInsert     =   0
         CaretOverWrite  =   3
         UserEntry       =   0
         HideSelection   =   -1  'True
         InvalidColor    =   12648447
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
         Text            =   "11/20/2002"
         DateCalcMethod  =   0
         DateTimeFormat  =   5
         UserDefinedFormat=   "mm/dd/yyyy"
         DateMax         =   "00000000"
         DateMin         =   "00000000"
         TimeMax         =   "000000"
         TimeMin         =   "000000"
         TimeString1159  =   ""
         TimeString2359  =   ""
         DateDefault     =   "00000000"
         TimeDefault     =   "000000"
         TimeStyle       =   0
         BorderGrayAreaColor=   -2147483637
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483633
         Appearance      =   2
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         PopUpType       =   1
         DateCalcY2KSplit=   60
         CaretPosition   =   0
         IncYear         =   1
         IncMonth        =   1
         IncDay          =   1
         IncHour         =   1
         IncMinute       =   1
         IncSecond       =   1
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         StartMonth      =   4
         ButtonAlign     =   0
         BoundDataType   =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Category:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   912
         TabIndex        =   11
         Top             =   2112
         Width           =   1116
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Print Order:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   1392
         TabIndex        =   10
         Top             =   4176
         Width           =   1308
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   3516
         Left            =   384
         Top             =   1488
         Width           =   7068
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
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Transactions by Category"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   13.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   396
         Left            =   1776
         TabIndex        =   9
         Top             =   576
         Width           =   4332
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Beginning Date:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2160
         TabIndex        =   8
         Top             =   2832
         Width           =   1740
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ending Date:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2352
         TabIndex        =   7
         Top             =   3456
         Width           =   1548
      End
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   6828
      Left            =   1800
      Top             =   1020
      Width           =   8052
   End
End
Attribute VB_Name = "frmBLTransJrnlByCat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsBLTextBoxOverrider
  Private Temp_Class As Resize_Class

Private Sub cmdCodeList_Click()
  frmBLCategoryList.Show vbModal
  DoEvents
End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsBLTextBoxOverrider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Call LoadMe
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
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
      Call cmdExit_Click
      SendKeys "%C"
      KeyCode = 0
    Case vbKeyF10:
      Call cmdProcess_Click
      SendKeys "%P"
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
      MainLog ("BusinessLicense.exe terminated via menu bar on frmBLTarnsJrnlByCat.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub LoadMe()
  Dim CodeRec As ARNewCatCodeRecType
  Dim CodeHandle As Integer
  Dim NumOfCodeRecs As Integer
  Dim x As Integer
  Dim CodeIdxRec As CatCodeIdxType
  Dim IdxHandle As Integer
  Dim IdxCnt As Integer
  
  On Error GoTo ERRORSTUFF
  
  OpenCatCodeIdxFile IdxHandle
  IdxCnt = LOF(IdxHandle) / Len(CodeIdxRec)
  If IdxCnt = 0 Then
    Call CreateCatCodeIdx
    IdxCnt = LOF(IdxHandle) / Len(CodeIdxRec)
    If IdxCnt = 0 Then
      MsgBox "No category codes can be found."
      Close
      Exit Sub
    End If
  End If
  
  ReDim IdxRec(1 To IdxCnt) As Integer
  
  For x = 1 To IdxCnt
    Get IdxHandle, x, CodeIdxRec
    IdxRec(x) = CodeIdxRec.CatCodeRec
  Next x
  Close IdxHandle
  
  OpenCatCodeFile CodeHandle
  NumOfCodeRecs = LOF(CodeHandle) / Len(CodeRec)
  
  If NumOfCodeRecs <> IdxCnt Then
    frmBLMessageBoxJr.Label1.Caption = "Error: The number of category codes saved and the number of categories indexed are not the same. Please reindex the category codes on the Category Maintenance Menu."
    frmBLMessageBoxJr.Label1.Top = 700
    frmBLMessageBoxJr.Show vbModal
'    MsgBox "Error: The number of category codes saved and the number of categories indexed are not the same. Please reindex the category codes on the Category Maintenance Menu."
    Close
    Exit Sub
  End If
  
  fpcmbCatCode.Text = "ALL"
  fpcmbCatCode.AddItem "ALL"
  For x = 1 To IdxCnt
    Get CodeHandle, IdxRec(x), CodeRec
    fpcmbCatCode.AddItem QPTrim$(CodeRec.CatCode) + "   " + QPTrim$(CodeRec.CODEDESC)
  Next x
  Close CodeHandle
    
  fptxtBDate = "01/01/" + Mid(Date, 7, 4)
  fptxtEDate = Date
  fpcmbPrintOpt.AddItem "Graphical"
  fpcmbPrintOpt.AddItem "Text"
  fpcmbPrintOpt.Text = "Graphical"
  Exit Sub
  
ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLTransJrnlByCat", "LoadMe", Erl)
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

Private Sub fpcmbPrintOpt_Change()
  If QPTrim$(fpcmbPrintOpt.Text) = "" Then
    fpcmbPrintOpt.Text = "Graphical"
  End If
End Sub

Private Sub fpcmbPrintOpt_KeyDown(KeyCode As Integer, Shift As Integer)
  'this keeps the user from inadvertently changing data on this
  'combo box if they are scrolling through the form
  If KeyCode = vbKeySpace Then
    fpcmbPrintOpt.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbPrintOpt.ListIndex = -1
  End If
  If fpcmbPrintOpt.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbCatCode.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If
End Sub

Private Sub cmdExit_Click()
  frmBLTransJrnlMenu.Show
  DoEvents
  Unload frmBLTransJrnlByCat
End Sub

Private Sub cmdProcess_Click()
  If fpcmbPrintOpt.Text = "Graphical" Then
'    Call PrintGraphics
    Call PrintNewText
  ElseIf fpcmbPrintOpt.Text = "Text" Then
'    MsgBox "Pitch 10 is recommended for this report."
    Call PrintNewText
  Else
    Exit Sub
  End If
End Sub

Private Sub PrintText()
  Dim BegDate$
  Dim BegDateNum As Integer
  Dim EndDate$
  Dim EndDateNum As Integer
  Dim ReportFile$
  Dim RptHandle As Integer
  Dim TransCnt As Double
  Dim TransRec As ARTransRecType
  Dim THandle As Integer
  Dim LineCnt As Integer
  Dim MaxLines As Integer
  Dim FF$, cnt As Double
  Dim TotalTrans As Double
  Dim TotalAmt As Double
  Dim TotalPaid As Double
  Dim FeePd As Double
  Dim Category$
  Dim LeftOver As Double
  Dim CustRec As ARCustRecType
  Dim CHandle As Integer
  Dim BILLCAT1$
  Dim BILLCAT2$
  Dim BILLCAT3$
  Dim BILLCAT4$
  Dim BILLCAT5$
  Dim Fee1#
  Dim Fee2#
  Dim Fee3#
  Dim Fee4#
  Dim Fee5#
  Dim CatCnt!, CatFnd!
  Dim ll As Double
  Dim CategoryDesc$
  Dim CodeRec As ARNewCatCodeRecType
  Dim NumOFARCatRecs As Integer
  Dim COHandle As Integer
  Dim LCnt As Integer
  Dim Page As Integer
  Dim TRNumRecs As Double
  Dim CountNum As Double
  Dim BigNum$
  Dim HoldThis As TransIdxType
  Dim ThisRec As Double
  Dim SmallNum$
  Dim Nextx As Double
  Dim x As Double
  
  On Error GoTo ERRORSTUFF
  
  ReportFile$ = "TRANCUST.PRN"
  FF$ = Chr$(12)
  MaxLines = 53
  LineCnt = 0
  ReDim Cat$(300), CatAmt#(300), GTotalAmt#(103)
  
  FF$ = Chr(12)
  
  BegDate = fptxtBDate.Text
  BegDateNum = Date2Num(fptxtBDate.Text)
  EndDate = fptxtEDate.Text
  EndDateNum = Date2Num(fptxtEDate.Text)
  If EndDateNum < BegDateNum Then
    fptxtEDate.BackColor = &HFFFF&
    fptxtBDate.BackColor = &HFFFF&
    frmBLMessageBoxJr.Label1.Caption = "The ending date comes before the beginning date. Please reenter these values."
    frmBLMessageBoxJr.Label1.Top = 700
    frmBLMessageBoxJr.Show vbModal
'    MsgBox "The ending date comes before the beginning date. Please reenter these values."
    Close
    fptxtEDate.BackColor = &HFFFFFF
    fptxtBDate.BackColor = &HFFFFFF
    fptxtBDate.SetFocus
    Exit Sub
  End If
  
  OpenTransFile THandle 'used also in GetReportInformation2

  ReDim TransIdx(1 To 1) As TransIdxType

  GoSub GetReportInformation2
  
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  
  GoSub PrintRptHeader2
  OpenCustFile CHandle
  TransCnt = LOF(THandle) / Len(TransRec)
  frmBLShowPctComp.Label1 = "Loading Transaction List"
  frmBLShowPctComp.Show
  EnableCloseButton Me.hwnd, False
  cmdExit.Enabled = False
  cmdProcess.Enabled = False
  For cnt = 1 To CountNum
    Get THandle, TransIdx(cnt).TransRecNum, TransRec
    If Val(TransRec.CustomerNumber) = 0 Then
      GoTo BadCustSkip
    End If
    
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintRptHeader2
    End If
    
    'Get Customer
    Get CHandle, Val(TransRec.CustomerNumber), CustRec
    
    Print #RptHandle, MakeRegDate(TransRec.TransDate);
    Print #RptHandle, Tab(13); Left$(CustRec.CustName, 25);
    Print #RptHandle, Tab(40); "";
    
    Select Case TransRec.TransType
    Case 1
      Print #RptHandle, "Charge";
    Case 2
      Print #RptHandle, "Payment";
    Case 6
      Print #RptHandle, "Penalty";
    Case 9
      Print #RptHandle, "Beg Bal";
    Case 100
      Print #RptHandle, "DN Pen Adj.";
    Case 101
      Print #RptHandle, "UP Pen Adj.";
    Case 102
      Print #RptHandle, "UP Lic Adj.";
    Case 103
      Print #RptHandle, "DN Lic Adj.";
      
    End Select
    
    'print
    
    Print #RptHandle, Tab(50); Left$(TransRec.TransDesc, 18);
    Print #RptHandle, Tab(69); Using("$###,##0.00", TransRec.TransAmount)
  
    TotalTrans = TotalTrans + 1
    LineCnt = LineCnt + 1
    TotalAmt# = TotalAmt# + TransRec.TransAmount
    Rem total by category
TotalUp:
    GTotalAmt#(TransRec.TransType) = GTotalAmt#(TransRec.TransType) + TransRec.TransAmount
    
    If TransRec.TransType = 2 Then
      BILLCAT1$ = QPTrim$(CustRec.BILLCAT1)
      Fee1# = CustRec.Fee1
      BILLCAT2$ = QPTrim$(CustRec.BILLCAT2)
      Fee2# = CustRec.Fee2
      BILLCAT3$ = QPTrim$(CustRec.BILLCAT3)
      Fee3# = CustRec.Fee3
      BILLCAT4$ = QPTrim$(CustRec.BILLCAT4)
      Fee4# = CustRec.Fee4
      BILLCAT5$ = QPTrim$(CustRec.BILLCAT5)
      Fee5# = CustRec.Fee5
      
      TotalPaid# = TransRec.TransAmount
  
      If TotalPaid# > Fee1# Then
        FeePd# = Fee1#
        Category$ = BILLCAT1$
        GoSub UpdateSummary
        LeftOver# = TotalPaid# - Fee1#
        LeftOver# = Int((LeftOver# * 100) + 0.5) / 100
      Else
        FeePd# = TotalPaid#
        Category$ = BILLCAT1$
        GoSub UpdateSummary
        GoTo FinishSummary
      End If
      
      If LeftOver# > Fee2# Then
        FeePd# = Fee2#
        Category$ = BILLCAT2$
        GoSub UpdateSummary
        LeftOver# = LeftOver# - Fee2#
        LeftOver# = Int((LeftOver# * 100) + 0.5) / 100
      Else
        FeePd# = LeftOver#
        Category$ = BILLCAT2$
        GoSub UpdateSummary
        GoTo FinishSummary
      End If
      If LeftOver# > Fee3# Then
        FeePd# = Fee3#
        Category$ = BILLCAT3$
        GoSub UpdateSummary
        LeftOver# = LeftOver# - Fee3#
        LeftOver# = Int((LeftOver# * 100) + 0.5) / 100
      Else
        FeePd# = LeftOver#
        Category$ = BILLCAT3$
        GoSub UpdateSummary
        GoTo FinishSummary
      End If
      If LeftOver# > Fee4# Then
        FeePd# = Fee4#
        Category$ = BILLCAT4$
        GoSub UpdateSummary
        LeftOver# = LeftOver# - Fee4#
        LeftOver# = Int((LeftOver# * 100) + 0.5) / 100
      Else
        FeePd# = LeftOver#
        Category$ = BILLCAT4$
        GoSub UpdateSummary
        GoTo FinishSummary
      End If
      LeftOver# = Int((LeftOver# * 100) + 0.5) / 100
      FeePd# = LeftOver#
      Category$ = BILLCAT5$
      GoSub UpdateSummary
      
FinishSummary:
      
    End If
    
BadCustSkip:
    frmBLShowPctComp.ShowPctComp cnt, CountNum
    If frmBLShowPctComp.Out = True Then
      Close
      frmBLShowPctComp.Out = False
      Unload frmBLShowPctComp
      EnableCloseButton Me.hwnd, True
      cmdExit.Enabled = True
      cmdProcess.Enabled = True
      Exit Sub
    End If
    
  Next cnt
  Unload frmBLShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  cmdProcess.Enabled = True
  
  GoSub PrintRptEnding2
  Print #RptHandle, Chr$(18);   ' oki 320 10 cpi
  Close         'Close all open files now
  
  ViewPrint ReportFile, "Transaction Journal", True
  Kill ReportFile$
  
  Exit Sub
  
UpdateSummary:
  If FeePd# <= 0 Then Return
  If CatCnt! = 0 Then
    CatCnt! = 1
    Cat$(1) = Category$
    CatAmt#(1) = FeePd#
    Return
  End If
  For CatFnd! = 1 To CatCnt!
    If Cat$(CatFnd!) = Category$ Then
      CatAmt#(CatFnd!) = CatAmt#(CatFnd!) + FeePd#
      Return
    End If
  Next CatFnd!
  CatCnt! = CatCnt! + 1
  Cat$(CatCnt!) = Category$
  CatAmt#(CatCnt!) = FeePd#
  Return
  
PrintRptHeader2:
  Page = Page + 1
  Print #RptHandle, Tab(22); "Business License : Transactions Journal"
  Print #RptHandle, ""
  Print #RptHandle, "               Report Date: "; Date$; Tab(65); "Page #"; Page
  Print #RptHandle, "Beginning Transaction Date: "; BegDate$
  Print #RptHandle, "   Ending Transaction Date: "; EndDate$
  Print #RptHandle, ""
  Print #RptHandle, "  Date"; Tab(13); "Customer Name"; Tab(40); "Type"; Tab(50); "Description"; Tab(70); "Amount"
  Print #RptHandle, String$(80, "=")
  LineCnt = 5
  Return
  
PrintRptEnding2:
  Print #RptHandle, String$(80, "-")
  Print #RptHandle, Tab(69); Using("$###,##0.00", TotalAmt#)
  For cnt = 1 To 101
    If GTotalAmt#(cnt) <> 0 Then
      Print #RptHandle, "Trans Type : ";
      Select Case cnt
      Case 1
        Print #RptHandle, "Charge";
      Case 2
        Print #RptHandle, "Payment";
      Case 6
        Print #RptHandle, "Penalty";
      Case 9
        Print #RptHandle, "Beg Bal";
      Case 100
        Print #RptHandle, "DN Pen Adj.";
      Case 101
        Print #RptHandle, "UP Pen Adj.";
      Case 102
        Print #RptHandle, "UP Lic Adj.";
      Case 103
        Print #RptHandle, "DN Lic Adj.";
      
    End Select
      Print #RptHandle, "     Total Amount: "; Using("$#,###,##0.00", GTotalAmt#(cnt))
    End If
  Next cnt
  Print #RptHandle, FF$
  
  If CatCnt! > 0 Then
    Page = Page + 1
    Print #RptHandle, Tab(20); "Business License : Transactions Journal "
    Print #RptHandle, "Report Date: "; Date$; Tab(65); "Page #"; Page
    Print #RptHandle, "Total Payments by Category"
    LCnt = 1

    For ll = 1 To CatCnt!
      If LCnt > 55 Then
        Page = Page + 1
        Print #RptHandle, Tab(20); "Business License : Transactions Journal "
        Print #RptHandle, "Report Date: "; Date$; Tab(65); "Page #"; Page
        Print #RptHandle, "Total Payments by Category"
        LCnt = 3
      End If
      
      'Get Catagory Desc First
      CategoryDesc$ = ""
      OpenCatCodeFile COHandle
      NumOFARCatRecs = LOF(COHandle) \ Len(CodeRec)
      For cnt = 1 To NumOFARCatRecs
        Get COHandle, cnt, CodeRec
        If Cat$(ll) = RTrim$(CodeRec.CatCode) Then
          CategoryDesc$ = CodeRec.CODEDESC
          Exit For
        End If
      Next cnt
      
      Close COHandle
      
      Print #RptHandle, Cat$(ll); Tab(10); CategoryDesc$; Tab(50); Using("$###,###,##0.00", CatAmt#(ll))
      LCnt = LCnt + 1
    Next ll
    Print #RptHandle, FF$
  End If
  Return

GetReportInformation2:
  TRNumRecs = LOF(THandle) / Len(TransRec)
  frmBLShowPctComp.Label1 = "Building Index"
  frmBLShowPctComp.Show
  EnableCloseButton Me.hwnd, False
  cmdExit.Enabled = False
  cmdProcess.Enabled = False
  For cnt = 1 To TRNumRecs
    Get THandle, cnt, TransRec
    If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
      CountNum = CountNum + 1
      ReDim Preserve TransIdx(1 To CountNum)
      TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
      TransIdx(CountNum).TransRecNum = cnt
    End If
    frmBLShowPctComp.ShowPctComp cnt, TRNumRecs
    If frmBLShowPctComp.Out = True Then
      Close
      frmBLShowPctComp.Out = False
      Unload frmBLShowPctComp
      EnableCloseButton Me.hwnd, True
      cmdExit.Enabled = True
      cmdProcess.Enabled = True
      Exit Sub
    End If
  Next cnt
  
  If CountNum = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "There are no transactions saved between " + fptxtBDate.Text + " and " + fptxtEDate.Text + "."
    frmBLMessageBoxJr.Label1.Top = 700
    frmBLMessageBoxJr.Show vbModal
'    MsgBox "There are no transactions saved between " + fptxtBDate.Text + " and " + fptxtEDate.Text + "."
    Close
    EnableCloseButton Me.hwnd, True
    cmdExit.Enabled = True
    cmdProcess.Enabled = True
    Exit Sub
  End If
  
  BigNum = "A"
  For x = 1 To CountNum
    If QPTrim$(TransIdx(x).TransWho) > BigNum Then
      BigNum = QPTrim$(TransIdx(x).TransWho)
    End If
  Next x
  
  SmallNum = BigNum + "A"
  Nextx = 1
  frmBLShowPctComp.Label1 = "Sorting Transactions"
  frmBLShowPctComp.Show
  EnableCloseButton Me.hwnd, False
  cmdExit.Enabled = False
  cmdProcess.Enabled = False
  Do
    For x = Nextx To CountNum
      If QPTrim$(TransIdx(x).TransWho) < SmallNum Then
        SmallNum = QPTrim$(TransIdx(x).TransWho)
        ThisRec = x
      End If
    Next x
    HoldThis = TransIdx(Nextx)
    TransIdx(Nextx) = TransIdx(ThisRec)
    TransIdx(ThisRec) = HoldThis
    If Nextx = CountNum Then Exit Do
    SmallNum = BigNum + "A"
    Nextx = Nextx + 1
    frmBLShowPctComp.ShowPctComp Nextx, CountNum
    If frmBLShowPctComp.Out = True Then
      Close
      frmBLShowPctComp.Out = False
      Unload frmBLShowPctComp
      EnableCloseButton Me.hwnd, True
      cmdExit.Enabled = True
      cmdProcess.Enabled = True
      Exit Sub
    End If
  Loop
  
  Return
  
ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLTransJrnlByCat", "PrintText", Erl)
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


Private Sub PrintNewText()
  Dim BegDate$
  Dim BegDateNum As Integer
  Dim EndDate$
  Dim EndDateNum As Integer
  Dim ReportFile$
  Dim RptHandle As Integer
  Dim TransCnt As Double
  Dim TransRec As ARTransRecType
  Dim THandle As Integer
  Dim LineCnt As Integer
  Dim MaxLines As Integer
  Dim FF$, cnt As Double
  Dim TotalTrans As Double
  Dim TotalAmt As Double
  Dim TotalPaid As Double
  Dim FeePd As Double
  Dim Category$
  Dim LeftOver As Double
  Dim CustRec As ARCustRecType
  Dim CustHandle As Integer
  Dim BILLCAT1$
  Dim BILLCAT2$
  Dim BILLCAT3$
  Dim BILLCAT4$
  Dim BILLCAT5$
  Dim Fee1#
  Dim Fee2#
  Dim Fee3#
  Dim Fee4#
  Dim Fee5#
  Dim CatCnt!, CatFnd!
  Dim ll As Double
  Dim CategoryDesc$
  Dim CodeRec As ARNewCatCodeRecType
  Dim NumOFARCatRecs As Integer
  Dim COHandle As Integer
  Dim LCnt As Integer
  Dim Page As Integer
  Dim TRNumRecs As Double
  Dim CountNum As Double
  Dim BigNum$
  Dim HoldThis As TransIdxType
  Dim ThisRec As Double
  Dim SmallNum$
  Dim Nextx As Double
  Dim x As Double
  Dim CustIdxRec As CustNameIdxType
  Dim IdxHandle As Integer
  Dim CustIdxCnt As Integer
  Dim NumOfCustRecs As Integer
  Dim NextT As Double
  Dim CatCodeIdxRec As CatCodeIdxType
  Dim CIdxHandle As Integer
  Dim NumOfCatIdx As Integer
  Dim Code$
  Dim CodeCnt As Integer 'how many codes this customer has
  Dim FeeAmt As Double
  Dim LicAmt As Double
  Dim PenAmt As Double
  
  On Error GoTo ERRORSTUFF
  
  fpcmbCatCode.Row = -1
  Code$ = Mid(fpcmbCatCode.ColText, 1, 5)
  
  OpenCatCodeIdxFile CIdxHandle
  NumOfCatIdx = LOF(CIdxHandle) / Len(CatCodeIdxRec)
  If NumOfCatIdx = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "No category codes saved."
    frmBLMessageBoxJr.Label1.Top = 700
    frmBLMessageBoxJr.Show vbModal
    MsgBox "No category codes saved."
    Close
    Exit Sub
  End If
  
  ReDim CatIdxRec(1 To NumOfCatIdx) As Integer
  For x = 1 To NumOfCatIdx
    Get CIdxHandle, x, CatCodeIdxRec
    CatIdxRec(x) = CatCodeIdxRec.CatCodeRec
  Next x
  Close CIdxHandle
  
  ReportFile$ = "CDTRNRPT.PRN"
  FF$ = Chr$(12)
  MaxLines = 53
  LineCnt = 0
  ReDim Cat$(300), CatAmt#(300), GTotalAmt#(103)
  
  BegDate = fptxtBDate.Text
  BegDateNum = Date2Num(fptxtBDate.Text)
  EndDate = fptxtEDate.Text
  EndDateNum = Date2Num(fptxtEDate.Text)
  If EndDateNum < BegDateNum Then
    fptxtEDate.BackColor = &HFFFF&
    fptxtBDate.BackColor = &HFFFF&
    frmBLMessageBoxJr.Label1.Caption = "The ending date comes before the beginning date. Please reenter these values."
    frmBLMessageBoxJr.Label1.Top = 700
    frmBLMessageBoxJr.Show vbModal
'    MsgBox "The ending date comes before the beginning date. Please reenter these values."
    Close
    fptxtEDate.BackColor = &HFFFFFF
    fptxtBDate.BackColor = &HFFFFFF
    fptxtBDate.SetFocus
    Exit Sub
  End If
  
  OpenCustNameIdxFile IdxHandle
  CustIdxCnt = LOF(IdxHandle) / Len(CustIdxRec)
  
  If CustIdxCnt = 0 Then
    Call CreateCustNameIdx
    CustIdxCnt = LOF(IdxHandle) / Len(CustIdxRec)
  End If
  
  If CustIdxCnt = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "No customer records have been saved."
    frmBLMessageBoxJr.Label1.Top = 700
    frmBLMessageBoxJr.Show vbModal
'    MsgBox "No customer records have been saved."
    Close
    Exit Sub
  Else
    ReDim IdxRec(1 To CustIdxCnt) As Integer
    For x = 1 To CustIdxCnt
      Get IdxHandle, x, CustIdxRec
      IdxRec(x) = CustIdxRec.CustRec
    Next x
  End If
  Close IdxHandle
  
  OpenCustFile CustHandle
  NumOfCustRecs = LOF(CustHandle) / Len(CustRec)
  If NumOfCustRecs <> CustIdxCnt Then
    frmBLMessageBoxJr.Label1.Caption = "Error: The number of customers on file and the number of customers indexed are not the same. Please call Southern Software at 1-800-842-8190 for assistance."
    frmBLMessageBoxJr.Label1.Top = 700
    frmBLMessageBoxJr.Show vbModal
'    MsgBox "Error: The number of customers on file and the number of customers indexed are not the same. Please re-sort customer indices under the Customer Maintenance Menu."
    Close
    Exit Sub
  End If
  
  OpenCatCodeFile COHandle
  Nextx = 23
  Get COHandle, CatIdxRec(Nextx), CodeRec
  
  OpenTransFile THandle 'used also in GetReportInformation2
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  For x = 1 To NumOfCustRecs
    Get CustHandle, IdxRec(x), CustRec
      If CustRec.FirstTrans <> 0 Then
        NextT = CustRec.FirstTrans
      Else
        GoTo BadCust
      End If
      
      Do
        Get THandle, NextT, TransRec
          If TransRec.TransDate < BegDateNum Or TransRec.TransDate > EndDateNum Then
            GoTo OutOfDate
          Else
            Print #RptHandle, MakeRegDate(TransRec.TransDate); Tab(12); QPTrim$(CustRec.BillName); Tab(50); QPTrim$(TransRec.TransDesc); Tab(65); Using$("$###0.00", TransRec.TransAmount)
          End If
OutOfDate:
          If TransRec.NextTrans <> 0 Then
            NextT = TransRec.NextTrans
          Else
            Exit Do
          End If
      Loop
BadCust:
  Next x

  Print #RptHandle, Chr$(18);   ' oki 320 10 cpi
  Close         'Close all open files now
  
  ViewPrint ReportFile, "Transaction Journal", True
  
  Exit Sub

ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLTransJrnlByCat", "PrintNewText", Erl)
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

Private Sub fpcmbCatCode_Change()
  If QPTrim$(fpcmbCatCode.Text) = "" Then
    fpcmbCatCode.Text = "ALL"
  End If
End Sub

Private Sub fpcmbCatCode_KeyDown(KeyCode As Integer, Shift As Integer)
  'this keeps the user from inadvertently changing data on this
  'combo box if they are scrolling through the form
  If KeyCode = vbKeySpace Then
    fpcmbCatCode.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbCatCode.ListIndex = -1
  End If
  If fpcmbCatCode.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fptxtBDate.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If
End Sub

