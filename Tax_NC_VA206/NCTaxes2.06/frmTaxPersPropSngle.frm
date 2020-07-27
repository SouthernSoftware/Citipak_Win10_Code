VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmTaxPersProp 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Personal Property Information"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11610
   Icon            =   "frmTaxPersPropSngle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11610
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcmbLateListYN 
      Height          =   390
      Left            =   9195
      TabIndex        =   10
      Top             =   4800
      Width           =   900
      _Version        =   196608
      _ExtentX        =   1587
      _ExtentY        =   688
      Text            =   ""
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
      ColDesigner     =   "frmTaxPersPropSngle.frx":08CA
   End
   Begin LpLib.fpCombo fpcmbDiscoveryYN 
      Height          =   390
      Left            =   9195
      TabIndex        =   9
      Top             =   4320
      Width           =   900
      _Version        =   196608
      _ExtentX        =   1587
      _ExtentY        =   688
      Text            =   ""
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
      ColDesigner     =   "frmTaxPersPropSngle.frx":0BC1
   End
   Begin EditLib.fpCurrency fpCurrPersVal 
      Height          =   375
      Left            =   3765
      TabIndex        =   2
      Top             =   2880
      Width           =   2775
      _Version        =   196608
      _ExtentX        =   4895
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
   Begin EditLib.fpText fptxtThisCust 
      Height          =   390
      Left            =   2858
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   1130
      Width           =   6015
      _Version        =   196608
      _ExtentX        =   10610
      _ExtentY        =   688
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
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
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      AutoCase        =   1
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
   Begin EditLib.fpDateTime fptxtDate 
      Height          =   375
      Left            =   7013
      TabIndex        =   1
      Top             =   2040
      Width           =   1740
      _Version        =   196608
      _ExtentX        =   3069
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
      Text            =   "02/24/2005"
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
      ButtonColor     =   13684944
      AutoMenu        =   0   'False
      StartMonth      =   4
      ButtonAlign     =   0
      BoundDataType   =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSave 
      Height          =   420
      Left            =   4957
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   7920
      Width           =   1695
      _Version        =   131072
      _ExtentX        =   2990
      _ExtentY        =   741
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
      ButtonDesigner  =   "frmTaxPersPropSngle.frx":0EB8
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   420
      Left            =   7005
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   7920
      Width           =   1695
      _Version        =   131072
      _ExtentX        =   2990
      _ExtentY        =   741
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
      ButtonDesigner  =   "frmTaxPersPropSngle.frx":1094
   End
   Begin EditLib.fpText fptxtRecord 
      Height          =   390
      Left            =   315
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   8235
      Visible         =   0   'False
      Width           =   2295
      _Version        =   196608
      _ExtentX        =   4048
      _ExtentY        =   688
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
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
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      AutoCase        =   1
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
      MaxLength       =   25
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
   Begin fpBtnAtlLibCtl.fpBtn cmdDelete 
      Height          =   420
      Left            =   2910
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   7920
      Width           =   1695
      _Version        =   131072
      _ExtentX        =   2990
      _ExtentY        =   741
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
      ButtonDesigner  =   "frmTaxPersPropSngle.frx":1270
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdAdd1 
      Height          =   420
      Left            =   315
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   8055
      Visible         =   0   'False
      Width           =   1695
      _Version        =   131072
      _ExtentX        =   2990
      _ExtentY        =   741
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
      ButtonDesigner  =   "frmTaxPersPropSngle.frx":144D
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPageDown1 
      Height          =   420
      Left            =   1035
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   8145
      Visible         =   0   'False
      Width           =   1695
      _Version        =   131072
      _ExtentX        =   2990
      _ExtentY        =   741
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
      ButtonDesigner  =   "frmTaxPersPropSngle.frx":1627
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPageUp1 
      Height          =   420
      Left            =   765
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   7920
      Visible         =   0   'False
      Width           =   1695
      _Version        =   131072
      _ExtentX        =   2990
      _ExtentY        =   741
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
      ButtonDesigner  =   "frmTaxPersPropSngle.frx":1803
   End
   Begin EditLib.fpText fptxtPropPin 
      Height          =   390
      Left            =   4298
      TabIndex        =   0
      Top             =   2040
      Width           =   1695
      _Version        =   196608
      _ExtentX        =   2990
      _ExtentY        =   688
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
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
      AutoCase        =   1
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
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   20
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
   Begin EditLib.fpCurrency fpCurrMobHome 
      Height          =   375
      Left            =   3765
      TabIndex        =   3
      Top             =   3360
      Width           =   2775
      _Version        =   196608
      _ExtentX        =   4895
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
   Begin EditLib.fpCurrency fpCurrMerchCap 
      Height          =   375
      Left            =   3765
      TabIndex        =   4
      Top             =   3840
      Width           =   2775
      _Version        =   196608
      _ExtentX        =   4895
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
   Begin EditLib.fpCurrency fpCurrFarmEq 
      Height          =   375
      Left            =   3765
      TabIndex        =   5
      Top             =   4320
      Width           =   2775
      _Version        =   196608
      _ExtentX        =   4895
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
   Begin EditLib.fpCurrency fpCurrMachTools 
      Height          =   375
      Left            =   3765
      TabIndex        =   6
      Top             =   4800
      Width           =   2775
      _Version        =   196608
      _ExtentX        =   4895
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
   Begin EditLib.fpCurrency fpCurrSnCitizen 
      Height          =   375
      Left            =   8685
      TabIndex        =   7
      Top             =   3000
      Width           =   1695
      _Version        =   196608
      _ExtentX        =   2990
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
   Begin EditLib.fpCurrency fpCurrOther 
      Height          =   375
      Left            =   8685
      TabIndex        =   8
      Top             =   3480
      Width           =   1695
      _Version        =   196608
      _ExtentX        =   2990
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
   Begin EditLib.fpText fptxtDesc 
      Height          =   390
      Index           =   0
      Left            =   3600
      TabIndex        =   11
      Top             =   5520
      Width           =   6375
      _Version        =   196608
      _ExtentX        =   11245
      _ExtentY        =   688
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
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
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      AutoCase        =   1
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
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   30
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
   Begin EditLib.fpText fptxtDesc 
      Height          =   390
      Index           =   1
      Left            =   3600
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   5900
      Width           =   6375
      _Version        =   196608
      _ExtentX        =   11245
      _ExtentY        =   688
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
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
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      AutoCase        =   1
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
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   30
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   -1  'True
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
   Begin EditLib.fpText fptxtDesc 
      Height          =   390
      Index           =   2
      Left            =   3600
      TabIndex        =   13
      Top             =   6270
      Width           =   6375
      _Version        =   196608
      _ExtentX        =   11245
      _ExtentY        =   688
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
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
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      AutoCase        =   1
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
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   30
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
   Begin EditLib.fpText fptxtDesc 
      Height          =   390
      Index           =   3
      Left            =   3600
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   6650
      Width           =   6375
      _Version        =   196608
      _ExtentX        =   11245
      _ExtentY        =   688
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
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
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      AutoCase        =   1
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
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   30
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
   Begin EditLib.fpText fptxtDesc 
      Height          =   390
      Index           =   4
      Left            =   3600
      TabIndex        =   15
      Top             =   7020
      Width           =   6375
      _Version        =   196608
      _ExtentX        =   11245
      _ExtentY        =   688
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
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
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
      AutoBeep        =   0   'False
      AutoCase        =   1
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
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   30
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
   Begin VB.Shape Shape4 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   1695
      Left            =   6765
      Top             =   2520
      Width           =   4215
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   2895
      Left            =   645
      Top             =   2520
      Width           =   6135
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Required Fields = *"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   270
      Left            =   720
      TabIndex        =   41
      Top             =   1800
      Width           =   1740
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   2175
      Left            =   645
      Top             =   5400
      Width           =   10335
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   1215
      Left            =   6765
      Top             =   4200
      Width           =   4215
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Descriptions - Notes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   645
      TabIndex        =   39
      Top             =   5400
      Width           =   2580
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Other:"
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
      Height          =   270
      Left            =   7005
      TabIndex        =   38
      Top             =   3600
      Width           =   1500
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Senior Citizen:"
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
      Height          =   270
      Left            =   7005
      TabIndex        =   37
      Top             =   3120
      Width           =   1500
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Exemptions"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   6765
      TabIndex        =   36
      Top             =   2520
      Width           =   2100
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Late List Y/N?:"
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
      Height          =   270
      Left            =   7605
      TabIndex        =   35
      Top             =   4920
      Width           =   1380
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Machine/Tools Value:"
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
      Height          =   270
      Left            =   1245
      TabIndex        =   34
      Top             =   4920
      Width           =   2340
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Farm Equipment Value:"
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
      Height          =   270
      Left            =   1245
      TabIndex        =   33
      Top             =   4440
      Width           =   2340
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Merchant Capital Value:"
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
      Height          =   270
      Left            =   1245
      TabIndex        =   32
      Top             =   3960
      Width           =   2340
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile Home Value:"
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
      Height          =   270
      Left            =   1245
      TabIndex        =   31
      Top             =   3480
      Width           =   2340
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Personal Value:"
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
      Height          =   270
      Left            =   1245
      TabIndex        =   30
      Top             =   3000
      Width           =   2340
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Valuations"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   645
      TabIndex        =   29
      Top             =   2520
      Width           =   2100
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Record Sequence:"
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
      Height          =   270
      Left            =   765
      TabIndex        =   24
      Top             =   7605
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Label lblMode 
      BackStyle       =   0  'Transparent
      Caption         =   "Mode:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4778
      TabIndex        =   23
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Discovery Y/N?:"
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
      Height          =   270
      Left            =   7365
      TabIndex        =   19
      Top             =   4440
      Width           =   1620
   End
   Begin VB.Label Label72 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "*Pin Number:"
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
      Height          =   270
      Left            =   2858
      TabIndex        =   18
      Top             =   2145
      Width           =   1260
   End
   Begin VB.Label Label71 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
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
      Height          =   270
      Left            =   6098
      TabIndex        =   17
      Top             =   2160
      Width           =   780
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personal Property Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2911
      TabIndex        =   16
      Top             =   360
      Width           =   6015
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   1380
      Index           =   1
      Left            =   1478
      Top             =   300
      Width           =   8655
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   1440
      Left            =   1478
      Top             =   240
      Width           =   8655
   End
End
Attribute VB_Name = "frmTaxPersProp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class
  Dim CustName$
  Dim WhichRec As Integer
  Dim PersRecs() As Long
  Dim NumOfCustPPRecs As Integer
  Dim TempPROPPIN$
  Dim TempPROPDATE As Integer
  Dim TempPersVal#
  Dim TempMHVALUE#
  Dim TempMCVALUE#
  Dim TempCVALUE#
  Dim TempMTVALUE#
  Dim TempEXMPSENI#
  Dim TempEXMPOTHR#
  Dim TempDISCOV$
  Dim TempLateList$
  Dim TempDESC1$
  Dim TempDESC2$
  Dim TempDESC3$
  Dim TempDesc4$
  Dim TempDesc5$
  Dim DontExit As Boolean
  
'Private Sub cmdAdd_Click()
'  If Check4Changes(WhichRec) = True Then
'    Exit Sub
'  End If
'
'  If NumOfCustPPRecs = 0 Then
'    WhichRec = 0
'  Else
'    WhichRec = NumOfCustPPRecs + 1
'  End If
'
'  Call LoadAdd(WhichRec)
'
'  cmdAdd.Enabled = False
'  cmdPageDown.Enabled = False
'  cmdPageUp.Enabled = False
'  cmdDelete.Enabled = False
'End Sub

Private Sub cmdDelete_Click()
  Dim CustName$
  Dim ThisPin$
  Dim PersVal$
  Dim MobVal$
  Dim MerchVal$
  Dim FarmVal$
  Dim MachVal$
  
  frmTaxMsgWOpts.Label1.Caption = "Are you sure you wish to delete this record? Press F10 to continue the deletion. Otherwise, press ESC to abort the deletion."
  frmTaxMsgWOpts.Label1.Top = 900
  frmTaxMsgWOpts.cmdExit.Text = "ESC Abort"
  frmTaxMsgWOpts.cmdCont.Text = "F10 Delete OK"
  Me.ZOrder 0
  frmTaxCustAddEdit.Visible = False
  If EditCust = True Then
    frmTaxCustLookup.Visible = False
  End If
  If AddCust = True Then
    frmTaxCustMaintMenu.Visible = False
  End If
  frmTaxMsgWOpts.Show vbModal
  If EditCust = True Then
    frmTaxCustLookup.Visible = True
  End If
  If AddCust = True Then
    frmTaxCustMaintMenu.Visible = True
  End If
  frmTaxCustAddEdit.Visible = True
  Me.Show
  If frmTaxMsgWOpts.fptxtChoice.Text = "abort" Then
    Unload frmTaxMsgWOpts
    fptxtPropPin.SetFocus
    Exit Sub
  Else
    Unload frmTaxMsgWOpts
  End If
  CustName$ = QPTrim$(fptxtThisCust.Text)
  ThisPin$ = QPTrim$(fptxtPropPin.Text)
  PersVal$ = QPTrim$(fpCurrPersVal.Text)
  MobVal$ = QPTrim$(fpCurrMobHome.Text)
  MerchVal$ = QPTrim$(fpCurrMerchCap.Text)
  FarmVal$ = QPTrim$(fpCurrFarmEq.Text)
  MachVal$ = QPTrim$(fpCurrMachTools.Text)
  
  Call DelPersAbstract(PersRecs(), WhichRec, GCustNum)
  Call ClearAfterDelete
  Call TaxMsg(900, "The personal property was deleted successfully.")
  If PersRecs(0) = 0 Then
    frmTaxCustAddEdit.Show
    DoEvents
    Unload Me
    Exit Sub
  End If
  Call GetPersRecList(PersRecs(), GCustNum, CustName)
  NumOfCustPPRecs = PersRecs(0)
  MainLog ("PERSONAL PROPERTY DELETION: User deleted the following personal property for : " + CustName + " - Pin # " + ThisPin + " - Personal Value: " + PersVal + " - Mobile Value: " + MobVal + " - Merchant Value: " + MerchVal + " - Farm Value: " + FarmVal + " - Machine Value: " + MachVal + ".")
  If PersRecs(0) = 0 Then
    WhichRec = 0
    Call Loadme
  Else
    WhichRec = 1
    Call LoadAgain(WhichRec)
  End If
  frmTaxMsg.Label1.Caption = "The personal property was deleted successfully."
  frmTaxMsg.Label1.Top = 900
  Me.ZOrder 0
  frmTaxCustAddEdit.Visible = False
  If EditCust = True Then
    frmTaxCustLookup.Visible = False
  End If
  If AddCust = True Then
    frmTaxCustMaintMenu.Visible = False
  End If
  frmTaxMsg.Show vbModal
  If EditCust = True Then
    frmTaxCustLookup.Visible = True
  End If
  If AddCust = True Then
    frmTaxCustMaintMenu.Visible = True
  End If
  frmTaxCustAddEdit.Visible = True
  Me.Show
End Sub

Private Sub cmdExit_Click()
  
'  If cmdAdd.Enabled = False Then
'    frmTaxMsgWOpts.Label1.Caption = "Do you wish to exit without saving any changes? Press F10 to save. Press ESC to exit without saving."
'    frmTaxMsgWOpts.Label1.Top = 900
'    frmTaxMsgWOpts.cmdCont.Text = "F10 Save Changes"
'    frmTaxMsgWOpts.cmdExit.Text = "ESC OK to Exit"
'    Me.ZOrder 0
'    frmTaxCustAddEdit.Visible = False
'    If EditCust = True Then
'      frmTaxCustLookup.Visible = False
'    End If
'    If AddCust = True Then
'      frmTaxCustMaintMenu.Visible = False
'    End If
'    frmTaxMsgWOpts.Show vbModal
'    If EditCust = True Then
'      frmTaxCustLookup.Visible = True
'    End If
'    If AddCust = True Then
'      frmTaxCustMaintMenu.Visible = True
'    End If
'    frmTaxCustAddEdit.Visible = True
'    Me.Show
'    frmTaxCustLookup.Show
'    frmTaxCustAddEdit.Show
'    If frmTaxMsgWOpts.fptxtChoice.Text = "abort" Then
'      Unload frmTaxMsgWOpts
'      Unload Me
'      Exit Sub
'    Else
'      Unload frmTaxMsgWOpts
'      Call cmdSave_Click
'      If DontExit = True Then
'        DontExit = False
'        Exit Sub
'      Else
'        Unload Me
'        Exit Sub
'      End If
'    End If
'  End If
  
'  frmTaxCustLookup.Show
'  frmTaxCustAddEdit.Show
  If Check4Changes(WhichRec) = True Then
    Exit Sub
  End If
  
  If DontExit = False Then
    Unload Me
  Else
    DontExit = False
  End If


End Sub

'Private Sub cmdPageUp_Click()
'  If Check4Changes(WhichRec) = True Then
'    Exit Sub
'  End If
'
'  If WhichRec = NumOfCustPPRecs Then
'    frmTaxMsg.Label1.Caption = "Upper limit reached."
'    frmTaxMsg.Label1.Top = 900
'    Me.ZOrder 0
'    frmTaxCustAddEdit.Visible = False
'    If EditCust = True Then
'      frmTaxCustLookup.Visible = False
'    End If
'    If AddCust = True Then
'      frmTaxCustMaintMenu.Visible = False
'    End If
'    frmTaxMsg.Show vbModal
'    If EditCust = True Then
'      frmTaxCustLookup.Visible = True
'    End If
'    If AddCust = True Then
'      frmTaxCustMaintMenu.Visible = True
'    End If
'    frmTaxCustAddEdit.Visible = True
'    Me.ZOrder 0
'    frmTaxCustAddEdit.ZOrder 1
'    Exit Sub
'  End If
'
'  WhichRec = WhichRec + 1
'  Call LoadAgain(WhichRec)
'
'End Sub

'Private Sub cmdPageDown_Click()
'  If Check4Changes(WhichRec) = True Then
'    Exit Sub
'  End If
'
'  If WhichRec = 0 Or WhichRec = 1 Then
'    frmTaxMsg.Label1.Caption = "Lower limit reached."
'    frmTaxMsg.Label1.Top = 900
'    Me.ZOrder 0
'    frmTaxCustAddEdit.Visible = False
'    If EditCust = True Then
'      frmTaxCustLookup.Visible = False
'    End If
'    If AddCust = True Then
'      frmTaxCustMaintMenu.Visible = False
'    End If
'    frmTaxMsg.Show vbModal
'    If EditCust = True Then
'      frmTaxCustLookup.Visible = True
'    End If
'    If AddCust = True Then
'      frmTaxCustMaintMenu.Visible = True
'    End If
'    frmTaxCustAddEdit.Visible = True
'    Me.ZOrder 0
'    frmTaxCustAddEdit.ZOrder 1
'    Exit Sub
'  End If
'
'  WhichRec = WhichRec - 1
'  Call LoadAgain(WhichRec)
'
'End Sub

Private Sub cmdSave_Click()
  Dim PersPropRec As PersonalRecType
  Dim PHandle As Integer
  Dim NumOfPersRecs As Long
  Dim CustPin&
  Dim TaxRec As TaxCustType
  Dim THandle As Integer
  Dim NumOfCustRecs As Long
  Dim WhatPers&
  Dim LastPers&
  Dim CustPinRec As PINRecType
  Dim CPHandle As Integer
  Dim NumOfCPRecs As Long
  Dim IntPinRec As InternalPinType
  Dim IHandle As Integer
  Dim NumOfIntPins As Long
  Dim NextIntPin As Long
  Dim NextRec As Long
  Dim CustName$
  
  If QPTrim$(fptxtPropPin.Text) = "" Then
    frmTaxMsg.Label1.Caption = "The 'Pin Number' field is a requirement. Please enter a 'Pin Number' value."
    frmTaxMsg.Label1.Top = 900
    Me.ZOrder 0
    frmTaxCustAddEdit.Visible = False
    If EditCust = True Then
      frmTaxCustLookup.Visible = False
    End If
    If AddCust = True Then
      frmTaxCustMaintMenu.Visible = False
    End If
    frmTaxMsg.Show vbModal
    If EditCust = True Then
      frmTaxCustLookup.Visible = True
    End If
    If AddCust = True Then
      frmTaxCustMaintMenu.Visible = True
    End If
    frmTaxCustAddEdit.Visible = True
    Me.Show
    fptxtPropPin.SetFocus
    DontExit = True
    Exit Sub
  End If
  
  If fpCurrPersVal.Value = 0 And fpCurrMobHome.Value = 0 And fpCurrMerchCap.Value = 0 And fpCurrFarmEq.Value = 0 And fpCurrMachTools.Value = 0 Then
    frmTaxMsgWOpts.Label1.Caption = "No property values have been entered. Press F10 to save anyway. Otherwise, press ESC to abort the save procedure."
    frmTaxMsgWOpts.Label1.Top = 800
    frmTaxMsgWOpts.cmdCont.Text = "F10 Save Anyway"
    frmTaxMsgWOpts.cmdExit.Text = "ESC Abort Save"
    Me.ZOrder 0
    frmTaxCustAddEdit.Visible = False
    If EditCust = True Then
      frmTaxCustLookup.Visible = False
    End If
    If AddCust = True Then
      frmTaxCustMaintMenu.Visible = False
    End If
    frmTaxMsgWOpts.Show vbModal
    If EditCust = True Then
      frmTaxCustLookup.Visible = True
    End If
    If AddCust = True Then
      frmTaxCustMaintMenu.Visible = True
    End If
    frmTaxCustAddEdit.Visible = True
    Me.Show
    If frmTaxMsgWOpts.fptxtChoice.Text = "abort" Then
      Unload frmTaxMsgWOpts
      fpCurrPersVal.SetFocus
      DontExit = True
      Exit Sub
    Else
      Unload frmTaxMsgWOpts
    End If
  End If
  
  OpenTaxCustFile THandle, NumOfCustRecs
  Get THandle, GCustNum, TaxRec
  CustPin& = TaxRec.PIN
  
  PersPropRec.PropPin = QPTrim$(fptxtPropPin.Text)
  PersPropRec.PROPDATE = Date2Num(fptxtDate.Text)
  PersPropRec.PersVal = fpCurrPersVal
  PersPropRec.MHVALUE = fpCurrMobHome
  PersPropRec.MCVALUE = fpCurrMerchCap
  PersPropRec.CVALUE = fpCurrFarmEq
  PersPropRec.MTVALUE = fpCurrMachTools
  PersPropRec.EXMPSENI = fpCurrSnCitizen
  PersPropRec.EXMPOTHR = fpCurrOther
  PersPropRec.DISCOV = fpcmbDiscoveryYN.Text
  PersPropRec.LateList = fpcmbLateListYN.Text
  PersPropRec.DESC1 = fptxtDesc(0).Text
  PersPropRec.DESC2 = fptxtDesc(1).Text
  PersPropRec.DESC3 = fptxtDesc(2).Text
  PersPropRec.Desc4 = fptxtDesc(3).Text
  PersPropRec.Desc5 = fptxtDesc(4).Text
  PersPropRec.Deleted = 0
  PersPropRec.CustPin = CustPin&
  
  OpenPersPropFile PHandle, NumOfPersRecs
  
  WhatPers = NumOfPersRecs + 1

  If WhichRec = 0 Then 'first pers prop record for this customer
    PersPropRec.LastYrPrinted = 0
    PersPropRec.VehTaxYear = 0
    PersPropRec.DMVSubmitted = "N"
    PersPropRec.Blank = ""
    TaxRec.FirstPersRec = WhatPers&
    Put THandle, GCustNum, TaxRec
    Close THandle
    ReDim Preserve PersRecs(0 To 1) As Long
    PersRecs(0) = 1 '# of props for this customer
    PersRecs(1) = WhatPers 'record # for this prop
    NumOfCustPPRecs = 1
    PersPropRec.NextRec = 0
    Put PHandle, WhatPers, PersPropRec
    fptxtRecord.Text = CStr(NumOfCustPPRecs) + " of " + CStr(NumOfCustPPRecs)
  ElseIf WhichRec > NumOfCustPPRecs Then 'adding to existing pers prop
    NumOfCustPPRecs = NumOfCustPPRecs + 1
    ReDim Preserve PersRecs(0 To WhichRec) As Long
    PersRecs(0) = PersRecs(0) + 1
    PersRecs(WhichRec) = WhatPers
    PersPropRec.NextRec = 0
    Put PHandle, WhatPers, PersPropRec
    Get PHandle, PersRecs(NumOfCustPPRecs - 1), PersPropRec
    PersPropRec.NextRec = WhatPers
    Put PHandle, PersRecs(NumOfCustPPRecs - 1), PersPropRec
    fptxtRecord.Text = CStr(NumOfCustPPRecs) + " of " + CStr(NumOfCustPPRecs)
  Else 'editing existing data
    Get PHandle, PersRecs(WhichRec), PersPropRec
    PersPropRec.PropPin = QPTrim$(fptxtPropPin.Text)
    PersPropRec.PROPDATE = Date2Num(fptxtDate.Text)
    PersPropRec.PersVal = fpCurrPersVal
    PersPropRec.MHVALUE = fpCurrMobHome
    PersPropRec.MCVALUE = fpCurrMerchCap
    PersPropRec.CVALUE = fpCurrFarmEq
    PersPropRec.MTVALUE = fpCurrMachTools
    PersPropRec.EXMPSENI = fpCurrSnCitizen
    PersPropRec.EXMPOTHR = fpCurrOther
    PersPropRec.DISCOV = fpcmbDiscoveryYN.Text
    PersPropRec.LateList = fpcmbLateListYN.Text
    PersPropRec.DESC1 = fptxtDesc(0).Text
    PersPropRec.DESC2 = fptxtDesc(1).Text
    PersPropRec.DESC3 = fptxtDesc(2).Text
    PersPropRec.Desc4 = fptxtDesc(3).Text
    PersPropRec.Desc5 = fptxtDesc(4).Text
    Put PHandle, PersRecs(WhichRec), PersPropRec
    Call LogSaves
  End If
  
  Close PHandle
  Close THandle
  
  ReDim PersRecs(0 To 0) As Long
  Call GetPersRecList(PersRecs(), GCustNum, CustName)
  
  Call MakePersPINFile
  
  cmdDelete.Enabled = True
  Call AssignTemps
  
  Me.ZOrder 0
  frmTaxCustAddEdit.Visible = False
  If EditCust = True Then
    frmTaxCustLookup.Visible = False
  End If
  If AddCust = True Then
    frmTaxCustMaintMenu.Visible = False
  End If
  Call Savemsg(900, "Your personal property data has been saved successfully.")
'  frmTaxMsg.Show vbModal
  If EditCust = True Then
    frmTaxCustLookup.Visible = True
  End If
  If AddCust = True Then
    frmTaxCustMaintMenu.Visible = True
  End If
  frmTaxCustAddEdit.Visible = True
  Unload Me 'added after removing multi prop feature
'  Me.Show 'remarked after removing multi prop feature
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
      SendKeys "%S"
      Call cmdSave_Click
      KeyCode = 0
    Case vbKeyF3:
      SendKeys "%D"
      Call cmdDelete_Click
      KeyCode = 0
'    Case vbKeyF8:
'      SendKeys "%A"
'      Call cmdAdd_Click
'      KeyCode = 0
'    Case vbKeyPageUp:
'      Call cmdPageUp_Click
'      KeyCode = 0
'    Case vbKeyPageDown:
'      Call cmdPageDown_Click
'      KeyCode = 0
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
  DontExit = False
  Call Loadme
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("CitiTaxes.exe terminated via menu bar on frmTaxPersProp.")
      Call Terminate
      End
    End If
  End If
End Sub
Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
    DoEvents
  End If
End Sub

Private Sub Loadme()
  Dim PersPropRec As PersonalRecType
  Dim PHandle As Integer
  Dim NumOfPersRecs As Long
  Dim x As Integer
  
  fpcmbDiscoveryYN.AddItem "Y"
  fpcmbDiscoveryYN.AddItem "N"
  fpcmbLateListYN.AddItem "Y"
  fpcmbLateListYN.AddItem "N"
  
  ReDim PersRecs(0 To 0) As Long
  Call GetPersRecList(PersRecs(), GCustNum, CustName)
  fptxtThisCust.Text = CustName
  NumOfCustPPRecs = PersRecs(0)
  
  If NumOfCustPPRecs = 0 Then
    WhichRec = 0
    fptxtPropPin.Text = ""
    fptxtRecord.Text = "None Saved"
    lblMode.Caption = "Mode: Adding"
    fptxtDate.Text = Date
    fpcmbDiscoveryYN.Text = "N"
    fpcmbLateListYN.Text = "N"
    fpCurrPersVal = 0
    fpCurrMobHome = 0
    fpCurrMerchCap = 0
    fpCurrFarmEq = 0
    fpCurrMachTools = 0
    fpCurrSnCitizen = 0
    fpCurrOther = 0
    fptxtDesc(0).Text = ""
    fptxtDesc(1).Text = ""
    fptxtDesc(2).Text = ""
    fptxtDesc(3).Text = ""
    fptxtDesc(4).Text = ""
  Else
    OpenPersPropFile PHandle, NumOfPersRecs
    Get PHandle, PersRecs(1), PersPropRec
'    PersPropRec.CustPin = PersPropRec.CustPin
    Close PHandle
    WhichRec = 1
    fptxtRecord.Text = "1 of " + CStr(NumOfCustPPRecs)
    lblMode.Caption = "Mode: Editing"
    fptxtDate.Text = MakeRegDate(PersPropRec.PROPDATE)
    fptxtPropPin.Text = QPTrim$(PersPropRec.PropPin)
    If PersPropRec.DISCOV <> "Y" Then
      fpcmbDiscoveryYN.Text = "N"
    Else
      fpcmbDiscoveryYN.Text = "Y"
    End If
    If PersPropRec.LateList <> "Y" Then
      fpcmbLateListYN.Text = "N"
    Else
      fpcmbLateListYN.Text = "Y"
    End If
    fpCurrPersVal = PersPropRec.PersVal
    fpCurrMobHome = PersPropRec.MHVALUE
    fpCurrMerchCap = PersPropRec.MCVALUE
    fpCurrFarmEq = PersPropRec.CVALUE
    fpCurrMachTools = PersPropRec.MTVALUE
    fpCurrSnCitizen = PersPropRec.EXMPSENI
    fpCurrOther = PersPropRec.EXMPOTHR
    fptxtDesc(0).Text = PersPropRec.DESC1
    fptxtDesc(1).Text = PersPropRec.DESC2
    fptxtDesc(2).Text = PersPropRec.DESC3
    fptxtDesc(3).Text = PersPropRec.Desc4
    fptxtDesc(4).Text = PersPropRec.Desc5
    Call AssignTemps
  End If
  
End Sub

Private Sub LoadAgain(WhichRec)
  Dim PersPropRec As PersonalRecType
  Dim PHandle As Integer
  Dim NumOfPersRecs As Long
  
  OpenPersPropFile PHandle, NumOfPersRecs
  Get PHandle, PersRecs(WhichRec), PersPropRec
  Close PHandle
  fptxtRecord.Text = CStr(WhichRec) + " of " + CStr(NumOfCustPPRecs)
  lblMode.Caption = "Mode: Editing"
  fptxtDate.Text = MakeRegDate(PersPropRec.PROPDATE)
  fptxtPropPin.Text = QPTrim$(PersPropRec.PropPin)
  If PersPropRec.DISCOV <> "Y" Then
    fpcmbDiscoveryYN.Text = "N"
  Else
    fpcmbDiscoveryYN.Text = "Y"
  End If
  If PersPropRec.LateList <> "Y" Then
    fpcmbLateListYN.Text = "N"
  Else
    fpcmbLateListYN.Text = "Y"
  End If
  fpCurrPersVal = PersPropRec.PersVal
  fpCurrMobHome = PersPropRec.MHVALUE
  fpCurrMerchCap = PersPropRec.MCVALUE
  fpCurrFarmEq = PersPropRec.CVALUE
  fpCurrMachTools = PersPropRec.MTVALUE
  fpCurrSnCitizen = PersPropRec.EXMPSENI
  fpCurrOther = PersPropRec.EXMPOTHR
  fptxtDesc(0).Text = PersPropRec.DESC1
  fptxtDesc(1).Text = PersPropRec.DESC2
  fptxtDesc(2).Text = PersPropRec.DESC3
  fptxtDesc(3).Text = PersPropRec.Desc4
  fptxtDesc(4).Text = PersPropRec.Desc5
  Call AssignTemps
   
End Sub
Private Sub LoadAdd(WhichRec)
  If NumOfCustPPRecs > 0 Then
    fptxtRecord.Text = "Adding Record # " + CStr(WhichRec)
  Else
    fptxtRecord.Text = "Adding 1st Record"
  End If
  lblMode.Caption = "Mode: Adding"
  fptxtDate.Text = Date
  fptxtPropPin.Text = ""
  fpcmbDiscoveryYN.Text = "N"
  fpcmbLateListYN.Text = "N"
  fpCurrPersVal = 0
  fpCurrMobHome = 0
  fpCurrMerchCap = 0
  fpCurrFarmEq = 0
  fpCurrMachTools = 0
  fpCurrSnCitizen = 0
  fpCurrOther = 0
  fptxtDesc(0).Text = ""
  fptxtDesc(1).Text = ""
  fptxtDesc(2).Text = ""
  fptxtDesc(3).Text = ""
  fptxtDesc(4).Text = ""
   
End Sub

Private Sub fpcmbDiscoveryYN_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbDiscoveryYN.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbDiscoveryYN.ListIndex = -1
  End If
  If fpcmbDiscoveryYN.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbLateListYN.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbLateListYN_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbLateListYN.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbLateListYN.ListIndex = -1
  End If
  If fpcmbLateListYN.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fptxtDesc(0).SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fptxtDesc_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If Index = 4 Then
    If KeyCode = vbKeyDown Then
      fptxtPropPin.SetFocus
    ElseIf KeyCode = vbKeyUp Then
      fptxtDesc(3).SetFocus
    End If
  End If
End Sub

Private Sub fptxtPropPin_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    fptxtDate.SetFocus
  ElseIf KeyCode = vbKeyUp Then
    fptxtDesc(4).SetFocus
  End If
End Sub

Private Function Check4Changes(WhichRec) As Boolean
  Dim ThisControl As Control
  Dim ThisDesc$
  Dim ThatDesc$
  Dim PersRec As PersonalRecType
  Dim PHandle As Integer
  Dim NumOfPersRecs As Long
  Dim ThisDbl#
  Dim ThatDbl#
  Dim choice$
  Dim NoEntry As Boolean
  
  On Error GoTo ERRORSTUFF
  Check4Changes = False
  NoEntry = True
  If PersRecs(WhichRec) > 0 Then
    OpenPersPropFile PHandle, NumOfPersRecs
    Get PHandle, PersRecs(WhichRec), PersRec
  Else
    GoSub EntryCheck
    If NoEntry = True Then Exit Function
    frmTaxMsgWOpts.Label1.Caption = "Do you wish to exit without saving any changes? Press F10 to save. Press ESC to exit without saving."
    frmTaxMsgWOpts.Label1.Top = 900
    frmTaxMsgWOpts.cmdCont.Text = "F10 Save Changes"
    frmTaxMsgWOpts.cmdExit.Text = "ESC OK to Exit"
    Me.ZOrder 0
    frmTaxCustAddEdit.Visible = False
    If EditCust = True Then
      frmTaxCustLookup.Visible = False
    End If
    If AddCust = True Then
      frmTaxCustMaintMenu.Visible = False
    End If
    frmTaxMsgWOpts.Show vbModal
    If EditCust = True Then
      frmTaxCustLookup.Visible = True
    End If
    If AddCust = True Then
      frmTaxCustMaintMenu.Visible = True
    End If
    frmTaxCustAddEdit.Visible = True
    Me.Show
    If frmTaxMsgWOpts.fptxtChoice.Text = "abort" Then
      Unload frmTaxMsgWOpts
      Exit Function
    Else
      Unload frmTaxMsgWOpts
      Call cmdSave_Click
      Exit Function
    End If
  End If
    
  Set ThisControl = fptxtPropPin
  ThisDesc = QPTrim$(fptxtPropPin.Text)
  ThatDesc = TempPROPPIN$
  If ThatDesc <> ThisDesc Then
    frmTaxMsgW4Opts.Label1.Caption = "The 'Pin Number' field has been changed from " + ThatDesc + " to " + ThisDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmTaxMsgW4Opts.Label1.Top = 575
    Me.ZOrder 0
    frmTaxCustAddEdit.Visible = False
    If EditCust = True Then
      frmTaxCustLookup.Visible = False
    End If
    If AddCust = True Then
      frmTaxCustMaintMenu.Visible = False
    End If
    frmTaxMsgW4Opts.Show vbModal
    If EditCust = True Then
      frmTaxCustLookup.Visible = True
    End If
    If AddCust = True Then
      frmTaxCustMaintMenu.Visible = True
    End If
    frmTaxCustAddEdit.Visible = True
    Me.Show
    choice = frmTaxMsgW4Opts.fptxtChoice.Text
    Unload frmTaxMsgW4Opts
    If choice = "save" Then
      PersRec.PropPin = QPTrim$(ThisControl.Text)
      Put PHandle, PersRecs(WhichRec), PersRec
      Call Savemsg(900, "The Pin Number has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
     
  Set ThisControl = fptxtDate
  ThisDesc = fptxtDate.Text
  ThatDesc = MakeRegDate(TempPROPDATE)
  If ThatDesc <> ThisDesc Then
    frmTaxMsgW4Opts.Label1.Caption = "The 'Date' field has been changed from " + ThatDesc + " to " + ThisDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmTaxMsgW4Opts.Label1.Top = 575
    Me.ZOrder 0
    frmTaxCustAddEdit.Visible = False
    If EditCust = True Then
      frmTaxCustLookup.Visible = False
    End If
    If AddCust = True Then
      frmTaxCustMaintMenu.Visible = False
    End If
    frmTaxMsgW4Opts.Show vbModal
    If EditCust = True Then
      frmTaxCustLookup.Visible = True
    End If
    If AddCust = True Then
      frmTaxCustMaintMenu.Visible = True
    End If
    frmTaxCustAddEdit.Visible = True
    Me.Show
    choice = frmTaxMsgW4Opts.fptxtChoice.Text
    Unload frmTaxMsgW4Opts
    If choice = "save" Then
      PersRec.PROPDATE = Date2Num(ThisDesc)
      Put PHandle, PersRecs(WhichRec), PersRec
      Call Savemsg(900, "The Date has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
     
  Set ThisControl = fpCurrPersVal
  ThisDbl = fpCurrPersVal.Value
  ThatDbl = TempPersVal#
  If ThatDbl <> ThisDbl Then
    frmTaxMsgW4Opts.Label1.Caption = "The 'Personal Value' field has been changed from " + QPTrim$(Using$("$###,##0.00", ThatDbl)) + " to " + QPTrim$(Using$("$###,##0.00", ThisDbl)) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmTaxMsgW4Opts.Label1.Top = 575
    Me.ZOrder 0
    frmTaxCustAddEdit.Visible = False
    If EditCust = True Then
      frmTaxCustLookup.Visible = False
    End If
    If AddCust = True Then
      frmTaxCustMaintMenu.Visible = False
    End If
    frmTaxMsgW4Opts.Show vbModal
    If EditCust = True Then
      frmTaxCustLookup.Visible = True
    End If
    If AddCust = True Then
      frmTaxCustMaintMenu.Visible = True
    End If
    frmTaxCustAddEdit.Visible = True
    Me.Show
    choice = frmTaxMsgW4Opts.fptxtChoice.Text
    Unload frmTaxMsgW4Opts
    If choice = "save" Then
      PersRec.PersVal = ThisDbl
      Put PHandle, PersRecs(WhichRec), PersRec
      Call Savemsg(900, "Personal Value has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
     
  Set ThisControl = fpCurrMobHome
  ThisDbl = fpCurrMobHome.Value
  ThatDbl = TempMHVALUE#
  If ThatDbl <> ThisDbl Then
    frmTaxMsgW4Opts.Label1.Caption = "The 'Mobile Home Value' field has been changed from " + QPTrim$(Using$("$###,##0.00", ThatDbl)) + " to " + QPTrim$(Using$("$###,##0.00", ThisDbl)) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmTaxMsgW4Opts.Label1.Top = 575
    Me.ZOrder 0
    frmTaxCustAddEdit.Visible = False
    If EditCust = True Then
      frmTaxCustLookup.Visible = False
    End If
    If AddCust = True Then
      frmTaxCustMaintMenu.Visible = False
    End If
    frmTaxMsgW4Opts.Show vbModal
    If EditCust = True Then
      frmTaxCustLookup.Visible = True
    End If
    If AddCust = True Then
      frmTaxCustMaintMenu.Visible = True
    End If
    frmTaxCustAddEdit.Visible = True
    Me.Show
    choice = frmTaxMsgW4Opts.fptxtChoice.Text
    Unload frmTaxMsgW4Opts
    If choice = "save" Then
      PersRec.MHVALUE = ThisDbl
      Put PHandle, PersRecs(WhichRec), PersRec
      Call Savemsg(900, "Mobile Home Value has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
     
  Set ThisControl = fpCurrMerchCap
  ThisDbl = fpCurrMerchCap.Value
  ThatDbl = TempMCVALUE#
  If ThatDbl <> ThisDbl Then
    frmTaxMsgW4Opts.Label1.Caption = "The 'Merchant Capital Value' field has been changed from " + QPTrim$(Using$("$###,##0.00", ThatDbl)) + " to " + QPTrim$(Using$("$###,##0.00", ThisDbl)) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmTaxMsgW4Opts.Label1.Top = 575
    Me.ZOrder 0
    frmTaxCustAddEdit.Visible = False
    If EditCust = True Then
      frmTaxCustLookup.Visible = False
    End If
    If AddCust = True Then
      frmTaxCustMaintMenu.Visible = False
    End If
    frmTaxMsgW4Opts.Show vbModal
    If EditCust = True Then
      frmTaxCustLookup.Visible = True
    End If
    If AddCust = True Then
      frmTaxCustMaintMenu.Visible = True
    End If
    frmTaxCustAddEdit.Visible = True
    Me.Show
    choice = frmTaxMsgW4Opts.fptxtChoice.Text
    Unload frmTaxMsgW4Opts
    If choice = "save" Then
      PersRec.MCVALUE = ThisDbl
      Put PHandle, PersRecs(WhichRec), PersRec
      Call Savemsg(900, "Merchant Capital Value has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
     
  Set ThisControl = fpCurrFarmEq
  ThisDbl = fpCurrFarmEq.Value
  ThatDbl = TempCVALUE#
  If ThatDbl <> ThisDbl Then
    frmTaxMsgW4Opts.Label1.Caption = "The 'Farm Equipment Value' field has been changed from " + QPTrim$(Using$("$###,##0.00", ThatDbl)) + " to " + QPTrim$(Using$("$###,##0.00", ThisDbl)) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmTaxMsgW4Opts.Label1.Top = 575
    Me.ZOrder 0
    frmTaxCustAddEdit.Visible = False
    If EditCust = True Then
      frmTaxCustLookup.Visible = False
    End If
    If AddCust = True Then
      frmTaxCustMaintMenu.Visible = False
    End If
    frmTaxMsgW4Opts.Show vbModal
    If EditCust = True Then
      frmTaxCustLookup.Visible = True
    End If
    If AddCust = True Then
      frmTaxCustMaintMenu.Visible = True
    End If
    frmTaxCustAddEdit.Visible = True
    Me.Show
    choice = frmTaxMsgW4Opts.fptxtChoice.Text
    Unload frmTaxMsgW4Opts
    If choice = "save" Then
      PersRec.CVALUE = ThisDbl
      Put PHandle, PersRecs(WhichRec), PersRec
      Call Savemsg(900, "Farm Equipment Value has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
     
  Set ThisControl = fpCurrMachTools
  ThisDbl = fpCurrMachTools.Value
  ThatDbl = TempMTVALUE#
  If ThatDbl <> ThisDbl Then
    frmTaxMsgW4Opts.Label1.Caption = "The 'Machine Tools Value' field has been changed from " + QPTrim$(Using$("$###,##0.00", ThatDbl)) + " to " + QPTrim$(Using$("$###,##0.00", ThisDbl)) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmTaxMsgW4Opts.Label1.Top = 575
    Me.ZOrder 0
    frmTaxCustAddEdit.Visible = False
    If EditCust = True Then
      frmTaxCustLookup.Visible = False
    End If
    If AddCust = True Then
      frmTaxCustMaintMenu.Visible = False
    End If
    frmTaxMsgW4Opts.Show vbModal
    If EditCust = True Then
      frmTaxCustLookup.Visible = True
    End If
    If AddCust = True Then
      frmTaxCustMaintMenu.Visible = True
    End If
    frmTaxCustAddEdit.Visible = True
    Me.Show
    choice = frmTaxMsgW4Opts.fptxtChoice.Text
    Unload frmTaxMsgW4Opts
    If choice = "save" Then
      PersRec.MTVALUE = ThisDbl
      Put PHandle, PersRecs(WhichRec), PersRec
      Call Savemsg(900, "Machine Tools Value has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
     
  Set ThisControl = fpCurrSnCitizen
  ThisDbl = fpCurrSnCitizen.Value
  ThatDbl = TempEXMPSENI#
  If ThatDbl <> ThisDbl Then
    frmTaxMsgW4Opts.Label1.Caption = "The 'Senior Citizen' field has been changed from " + QPTrim$(Using$("$###,##0.00", ThatDbl)) + " to " + QPTrim$(Using$("$###,##0.00", ThisDbl)) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmTaxMsgW4Opts.Label1.Top = 575
    Me.ZOrder 0
    frmTaxCustAddEdit.Visible = False
    If EditCust = True Then
      frmTaxCustLookup.Visible = False
    End If
    If AddCust = True Then
      frmTaxCustMaintMenu.Visible = False
    End If
    frmTaxMsgW4Opts.Show vbModal
    If EditCust = True Then
      frmTaxCustLookup.Visible = True
    End If
    If AddCust = True Then
      frmTaxCustMaintMenu.Visible = True
    End If
    frmTaxCustAddEdit.Visible = True
    Me.Show
    choice = frmTaxMsgW4Opts.fptxtChoice.Text
    Unload frmTaxMsgW4Opts
    If choice = "save" Then
      PersRec.EXMPSENI = ThisDbl
      Put PHandle, PersRecs(WhichRec), PersRec
      Call Savemsg(900, "Senior Citizen has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
     
  Set ThisControl = fpCurrOther
  ThisDbl = fpCurrOther.Value
  ThatDbl = TempEXMPOTHR#
  If ThatDbl <> ThisDbl Then
    frmTaxMsgW4Opts.Label1.Caption = "The 'Other' field has been changed from " + QPTrim$(Using$("$###,##0.00", ThatDbl)) + " to " + QPTrim$(Using$("$###,##0.00", ThisDbl)) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmTaxMsgW4Opts.Label1.Top = 575
    Me.ZOrder 0
    frmTaxCustAddEdit.Visible = False
    If EditCust = True Then
      frmTaxCustLookup.Visible = False
    End If
    If AddCust = True Then
      frmTaxCustMaintMenu.Visible = False
    End If
    frmTaxMsgW4Opts.Show vbModal
    If EditCust = True Then
      frmTaxCustLookup.Visible = True
    End If
    If AddCust = True Then
      frmTaxCustMaintMenu.Visible = True
    End If
    frmTaxCustAddEdit.Visible = True
    Me.Show
    choice = frmTaxMsgW4Opts.fptxtChoice.Text
    Unload frmTaxMsgW4Opts
    If choice = "save" Then
      PersRec.EXMPOTHR = ThisDbl
      Put PHandle, PersRecs(WhichRec), PersRec
      Call Savemsg(900, "Other has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
     
  Set ThisControl = fpcmbDiscoveryYN
  ThisDesc = QPTrim$(fpcmbDiscoveryYN.Text)
  ThatDesc = TempDISCOV$
  If ThatDesc <> ThisDesc Then
    frmTaxMsgW4Opts.Label1.Caption = "The 'Discovery (Y/N)' field has been changed from " + ThatDesc + " to " + ThisDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmTaxMsgW4Opts.Label1.Top = 575
    Me.ZOrder 0
    frmTaxCustAddEdit.Visible = False
    If EditCust = True Then
      frmTaxCustLookup.Visible = False
    End If
    If AddCust = True Then
      frmTaxCustMaintMenu.Visible = False
    End If
    frmTaxMsgW4Opts.Show vbModal
    If EditCust = True Then
      frmTaxCustLookup.Visible = True
    End If
    If AddCust = True Then
      frmTaxCustMaintMenu.Visible = True
    End If
    frmTaxCustAddEdit.Visible = True
    Me.Show
    choice = frmTaxMsgW4Opts.fptxtChoice.Text
    Unload frmTaxMsgW4Opts
    If choice = "save" Then
      PersRec.DISCOV = QPTrim$(ThisControl.Text)
      Put PHandle, PersRecs(WhichRec), PersRec
      Call Savemsg(900, "Discovery (Y/N) has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
     
  Set ThisControl = fpcmbLateListYN
  ThisDesc = QPTrim$(fpcmbLateListYN.Text)
  ThatDesc = TempLateList$
  If ThatDesc <> ThisDesc Then
    frmTaxMsgW4Opts.Label1.Caption = "The 'Late List (Y/N)' field has been changed from " + ThatDesc + " to " + ThisDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmTaxMsgW4Opts.Label1.Top = 575
    Me.ZOrder 0
    frmTaxCustAddEdit.Visible = False
    If EditCust = True Then
      frmTaxCustLookup.Visible = False
    End If
    If AddCust = True Then
      frmTaxCustMaintMenu.Visible = False
    End If
    frmTaxMsgW4Opts.Show vbModal
    If EditCust = True Then
      frmTaxCustLookup.Visible = True
    End If
    If AddCust = True Then
      frmTaxCustMaintMenu.Visible = True
    End If
    frmTaxCustAddEdit.Visible = True
    Me.Show
    choice = frmTaxMsgW4Opts.fptxtChoice.Text
    Unload frmTaxMsgW4Opts
    If choice = "save" Then
      PersRec.LateList = QPTrim$(ThisControl.Text)
      Put PHandle, PersRecs(WhichRec), PersRec
      Call Savemsg(900, "Late List (Y/N) has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
     
  Set ThisControl = fptxtDesc(0)
  ThisDesc = fptxtDesc(0).Text
  ThatDesc = TempDESC1$
  If ThatDesc <> ThisDesc Then
    frmTaxMsgW4Opts.Label1.Caption = "The 'Description Line #1' field has been changed from " + ThatDesc + " to " + ThisDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmTaxMsgW4Opts.Label1.Top = 575
    Me.ZOrder 0
    frmTaxCustAddEdit.Visible = False
    If EditCust = True Then
      frmTaxCustLookup.Visible = False
    End If
    If AddCust = True Then
      frmTaxCustMaintMenu.Visible = False
    End If
    frmTaxMsgW4Opts.Show vbModal
    If EditCust = True Then
      frmTaxCustLookup.Visible = True
    End If
    If AddCust = True Then
      frmTaxCustMaintMenu.Visible = True
    End If
    frmTaxCustAddEdit.Visible = True
    Me.Show
    choice = frmTaxMsgW4Opts.fptxtChoice.Text
    Unload frmTaxMsgW4Opts
    If choice = "save" Then
      PersRec.DESC1 = QPTrim$(ThisControl.Text)
      Put PHandle, PersRecs(WhichRec), PersRec
      Call Savemsg(900, "Description Line #1 has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
     
  Set ThisControl = fptxtDesc(1)
  ThisDesc = fptxtDesc(1).Text
  ThatDesc = TempDESC2$
  If ThatDesc <> ThisDesc Then
    frmTaxMsgW4Opts.Label1.Caption = "The 'Description Line #2' field has been changed from " + ThatDesc + " to " + ThisDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmTaxMsgW4Opts.Label1.Top = 575
    Me.ZOrder 0
    frmTaxCustAddEdit.Visible = False
    If EditCust = True Then
      frmTaxCustLookup.Visible = False
    End If
    If AddCust = True Then
      frmTaxCustMaintMenu.Visible = False
    End If
    frmTaxMsgW4Opts.Show vbModal
    If EditCust = True Then
      frmTaxCustLookup.Visible = True
    End If
    If AddCust = True Then
      frmTaxCustMaintMenu.Visible = True
    End If
    frmTaxCustAddEdit.Visible = True
    Me.Show
    choice = frmTaxMsgW4Opts.fptxtChoice.Text
    Unload frmTaxMsgW4Opts
    If choice = "save" Then
      PersRec.DESC2 = QPTrim$(ThisControl.Text)
      Put PHandle, PersRecs(WhichRec), PersRec
      Call Savemsg(900, "Description Line #2 has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
     
  Set ThisControl = fptxtDesc(2)
  ThisDesc = fptxtDesc(2).Text
  ThatDesc = TempDESC3$
  If ThatDesc <> ThisDesc Then
    frmTaxMsgW4Opts.Label1.Caption = "The 'Description Line #3' field has been changed from " + ThatDesc + " to " + ThisDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmTaxMsgW4Opts.Label1.Top = 575
    Me.ZOrder 0
    frmTaxCustAddEdit.Visible = False
    If EditCust = True Then
      frmTaxCustLookup.Visible = False
    End If
    If AddCust = True Then
      frmTaxCustMaintMenu.Visible = False
    End If
    frmTaxMsgW4Opts.Show vbModal
    If EditCust = True Then
      frmTaxCustLookup.Visible = True
    End If
    If AddCust = True Then
      frmTaxCustMaintMenu.Visible = True
    End If
    frmTaxCustAddEdit.Visible = True
    Me.Show
    choice = frmTaxMsgW4Opts.fptxtChoice.Text
    Unload frmTaxMsgW4Opts
    If choice = "save" Then
      PersRec.DESC3 = QPTrim$(ThisControl.Text)
      Put PHandle, PersRecs(WhichRec), PersRec
      Call Savemsg(900, "Description Line #3 has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
     
  Set ThisControl = fptxtDesc(3)
  ThisDesc = fptxtDesc(3).Text
  ThatDesc = TempDesc4$
  If ThatDesc <> ThisDesc Then
    frmTaxMsgW4Opts.Label1.Caption = "The 'Description Line #4' field has been changed from " + ThatDesc + " to " + ThisDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmTaxMsgW4Opts.Label1.Top = 575
    Me.ZOrder 0
    frmTaxCustAddEdit.Visible = False
    If EditCust = True Then
      frmTaxCustLookup.Visible = False
    End If
    If AddCust = True Then
      frmTaxCustMaintMenu.Visible = False
    End If
    frmTaxMsgW4Opts.Show vbModal
    If EditCust = True Then
      frmTaxCustLookup.Visible = True
    End If
    If AddCust = True Then
      frmTaxCustMaintMenu.Visible = True
    End If
    frmTaxCustAddEdit.Visible = True
    Me.Show
    choice = frmTaxMsgW4Opts.fptxtChoice.Text
    Unload frmTaxMsgW4Opts
    If choice = "save" Then
      PersRec.Desc4 = QPTrim$(ThisControl.Text)
      Put PHandle, PersRecs(WhichRec), PersRec
      Call Savemsg(900, "Description Line #4 has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
     
  Set ThisControl = fptxtDesc(4)
  ThisDesc = fptxtDesc(4).Text
  ThatDesc = TempDesc5$
  If ThatDesc <> ThisDesc Then
    frmTaxMsgW4Opts.Label1.Caption = "The 'Description Line #5' field has been changed from " + ThatDesc + " to " + ThisDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmTaxMsgW4Opts.Label1.Top = 575
    Me.ZOrder 0
    frmTaxCustAddEdit.Visible = False
    If EditCust = True Then
      frmTaxCustLookup.Visible = False
    End If
    If AddCust = True Then
      frmTaxCustMaintMenu.Visible = False
    End If
    frmTaxMsgW4Opts.Show vbModal
    If EditCust = True Then
      frmTaxCustLookup.Visible = True
    End If
    If AddCust = True Then
      frmTaxCustMaintMenu.Visible = True
    End If
    frmTaxCustAddEdit.Visible = True
    Me.Show
    choice = frmTaxMsgW4Opts.fptxtChoice.Text
    Unload frmTaxMsgW4Opts
    If choice = "save" Then
      PersRec.Desc5 = QPTrim$(ThisControl.Text)
      Put PHandle, PersRecs(WhichRec), PersRec
      Call Savemsg(900, "Description Line #5 has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
  
  Close PHandle
  
  Exit Function
  
EntryCheck:
  If QPTrim$(fptxtPropPin.Text) <> "" Then
    NoEntry = False
    Return
  ElseIf fpCurrPersVal.Value <> 0 Then
    NoEntry = False
    Return
  ElseIf fpCurrMobHome.Value <> 0 Then
    NoEntry = False
    Return
  ElseIf fpCurrMerchCap.Value <> 0 Then
    NoEntry = False
    Return
  ElseIf fpCurrFarmEq.Value <> 0 Then
    NoEntry = False
    Return
  ElseIf fpCurrMachTools.Value <> 0 Then
    NoEntry = False
    Return
  ElseIf fpCurrSnCitizen.Value <> 0 Then
    NoEntry = False
    Return
  ElseIf fpCurrOther.Value <> 0 Then
    NoEntry = False
    Return
  ElseIf QPTrim$(fptxtDesc(0).Text) <> "" Then
    NoEntry = False
    Return
  ElseIf QPTrim$(fptxtDesc(1).Text) <> "" Then
    NoEntry = False
    Return
  ElseIf QPTrim$(fptxtDesc(2).Text) <> "" Then
    NoEntry = False
    Return
  ElseIf QPTrim$(fptxtDesc(3).Text) <> "" Then
    NoEntry = False
    Return
  ElseIf QPTrim$(fptxtDesc(4).Text) <> "" Then
    NoEntry = False
    Return
  End If

  Return
  
HandleChoice:
    Select Case choice
      Case "abandon"
        Close PHandle
        Unload Me
        Exit Function
      Case "dontsave"
      Case "review"
        ThisControl.SetFocus
        Close PHandle
        Check4Changes = True
        Exit Function
      Case Else
    End Select
      
  Return
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxPersProp", "Check4Changes", Erl)
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
  
End Function

Private Sub AssignTemps()
  TempPROPPIN$ = QPTrim$(fptxtPropPin.Text)
  TempPROPDATE% = Date2Num(fptxtDate.Text)
  TempPersVal# = fpCurrPersVal.Value
  TempMHVALUE# = fpCurrMobHome.Value
  TempMCVALUE# = fpCurrMerchCap.Value
  TempCVALUE# = fpCurrFarmEq.Value
  TempMTVALUE# = fpCurrMachTools.Value
  TempEXMPSENI# = fpCurrSnCitizen.Value
  TempEXMPOTHR# = fpCurrOther.Value
  TempDISCOV$ = fpcmbDiscoveryYN.Text
  TempLateList$ = fpcmbLateListYN.Text
  TempDESC1$ = fptxtDesc(0).Text
  TempDESC2$ = fptxtDesc(1).Text
  TempDESC3$ = fptxtDesc(2).Text
  TempDesc4$ = fptxtDesc(3).Text
  TempDesc5$ = fptxtDesc(4).Text

End Sub

Private Sub LogSaves()
  Dim PersRec As PersonalRecType
  Dim PHandle As Integer
  Dim NumOfPersRecs As Long

  OpenPersPropFile PHandle, NumOfPersRecs
  Get PHandle, PersRecs(WhichRec), PersRec
  Close PHandle
  
  If QPTrim$(TempPROPPIN$) <> QPTrim$(PersRec.PropPin) Then
    MainLog ("For " + QPTrim$(CustName$) + " Personal Property #" + CStr(WhichRec) + ": the Property PIN# was changed from " + QPTrim$(TempPROPPIN$) + " to " + QPTrim$(PersRec.PropPin) + " and saved.")
  End If
  
  If TempPROPDATE% <> PersRec.PROPDATE Then
    MainLog ("For " + QPTrim$(CustName$) + " Personal Property #" + CStr(WhichRec) + ": the date was changed from " + MakeRegDate(TempPROPDATE) + " to " + MakeRegDate(PersRec.PROPDATE) + " and saved.")
  End If
  
  If TempPersVal# <> PersRec.PersVal Then
    MainLog ("For " + QPTrim$(CustName$) + " Personal Property #" + CStr(WhichRec) + ": the Personal Value was changed from " + QPTrim$(Using("$###,###,##0.00", TempPersVal)) + " to " + QPTrim$(Using("$###,###,##0.00", PersRec.PersVal)) + " and saved.")
  End If
  
  If TempMHVALUE# <> PersRec.MHVALUE Then
    MainLog ("For " + QPTrim$(CustName$) + " Personal Property #" + CStr(WhichRec) + ": the Mobile Home Value was changed from " + QPTrim$(Using("$###,###,##0.00", TempMHVALUE)) + " to " + QPTrim$(Using("$###,###,##0.00", PersRec.MHVALUE)) + " and saved.")
  End If
  
  If TempMCVALUE# <> PersRec.MCVALUE Then
    MainLog ("For " + QPTrim$(CustName$) + " Personal Property #" + CStr(WhichRec) + ": the Merchant Capital Value was changed from " + QPTrim$(Using("$###,###,##0.00", TempMCVALUE)) + " to " + QPTrim$(Using("$###,###,##0.00", PersRec.MCVALUE)) + " and saved.")
  End If
  
  If TempCVALUE# <> PersRec.CVALUE Then
    MainLog ("For " + QPTrim$(CustName$) + " Personal Property #" + CStr(WhichRec) + ": the Farm Equipment Value was changed from " + QPTrim$(Using("$###,###,##0.00", TempCVALUE)) + " to " + QPTrim$(Using("$###,###,##0.00", PersRec.CVALUE)) + " and saved.")
  End If
  
  If TempMTVALUE# <> PersRec.MTVALUE Then
    MainLog ("For " + QPTrim$(CustName$) + " Personal Property #" + CStr(WhichRec) + ": the Machine Tools Value was changed from " + QPTrim$(Using("$###,###,##0.00", TempMTVALUE)) + " to " + QPTrim$(Using("$###,###,##0.00", PersRec.MTVALUE)) + " and saved.")
  End If
  
  If TempEXMPSENI# <> PersRec.EXMPSENI Then
    MainLog ("For " + QPTrim$(CustName$) + " Personal Property #" + CStr(WhichRec) + ": the Senior Exemptions Value was changed from " + QPTrim$(Using("$###,###,##0.00", TempEXMPSENI)) + " to " + QPTrim$(Using("$###,###,##0.00", PersRec.EXMPSENI)) + " and saved.")
  End If
  
  If TempEXMPOTHR# <> PersRec.EXMPOTHR Then
    MainLog ("For " + QPTrim$(CustName$) + " Personal Property #" + CStr(WhichRec) + ": the Other Exemptions Value was changed from " + QPTrim$(Using("$###,###,##0.00", TempEXMPOTHR)) + " to " + QPTrim$(Using("$###,###,##0.00", PersRec.EXMPOTHR)) + " and saved.")
  End If
  
  If QPTrim$(TempDISCOV$) <> QPTrim$(PersRec.DISCOV) Then
    MainLog ("For " + QPTrim$(CustName$) + " Personal Property #" + CStr(WhichRec) + ": the Discovery Y/N? was changed from " + QPTrim$(TempDISCOV$) + " to " + QPTrim$(PersRec.DISCOV) + " and saved.")
  End If
  
  If QPTrim$(TempLateList$) <> QPTrim$(PersRec.LateList) Then
    MainLog ("For " + QPTrim$(CustName$) + " Personal Property #" + CStr(WhichRec) + ": the Late List Y/N? was changed from " + QPTrim$(TempLateList$) + " to " + QPTrim$(PersRec.LateList) + " and saved.")
  End If
  
  If QPTrim$(TempDESC1$) <> QPTrim$(PersRec.DESC1) Then
    MainLog ("For " + QPTrim$(CustName$) + " Personal Property #" + CStr(WhichRec) + ": the Notes Line #1 was changed from " + QPTrim$(TempDESC1$) + " to " + QPTrim$(PersRec.DESC1) + " and saved.")
  End If
  
  If QPTrim$(TempDESC2$) <> QPTrim$(PersRec.DESC2) Then
    MainLog ("For " + QPTrim$(CustName$) + " Personal Property #" + CStr(WhichRec) + ": the Notes Line #2 was changed from " + QPTrim$(TempDESC2$) + " to " + QPTrim$(PersRec.DESC2) + " and saved.")
  End If
  
  If QPTrim$(TempDESC3$) <> QPTrim$(PersRec.DESC3) Then
    MainLog ("For " + QPTrim$(CustName$) + " Personal Property #" + CStr(WhichRec) + ": the Notes Line #3 was changed from " + QPTrim$(TempDESC3$) + " to " + QPTrim$(PersRec.DESC3) + " and saved.")
  End If
  
  If QPTrim$(TempDesc4$) <> QPTrim$(PersRec.Desc4) Then
    MainLog ("For " + QPTrim$(CustName$) + " Personal Property #" + CStr(WhichRec) + ": the Notes Line #4 was changed from " + QPTrim$(TempDesc4$) + " to " + QPTrim$(PersRec.Desc4) + " and saved.")
  End If
  
  If QPTrim$(TempDesc5$) <> QPTrim$(PersRec.Desc5) Then
    MainLog ("For " + QPTrim$(CustName$) + " Personal Property #" + CStr(WhichRec) + ": the Notes Line #5 was changed from " + QPTrim$(TempDesc5$) + " to " + QPTrim$(PersRec.Desc5) + " and saved.")
  End If
  
End Sub

Private Sub ClearAfterDelete()
  fptxtPropPin.Text = ""
  fptxtDate = Date
  fpCurrPersVal = 0
  fpCurrMobHome = 0
  fpCurrMerchCap = 0
  fpCurrFarmEq = 0
  fpCurrMachTools = 0
  fpCurrSnCitizen = 0
  fpCurrOther = 0
  fpcmbDiscoveryYN.Text = "N"
  fpcmbLateListYN.Text = "N"
  fptxtDesc(0).Text = ""
  fptxtDesc(1).Text = ""
  fptxtDesc(2).Text = ""
  fptxtDesc(3).Text = ""
  fptxtDesc(4).Text = ""
  
End Sub
