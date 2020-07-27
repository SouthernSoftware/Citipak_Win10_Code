VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmVATaxPManualBillEntry 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manual Bill Entry"
   ClientHeight    =   8760
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   11655
   Icon            =   "frmVATaxPManualBillEntry.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpList fpListInEdit 
      Height          =   732
      Left            =   2724
      TabIndex        =   7
      Top             =   4080
      Width           =   8172
      _Version        =   196608
      _ExtentX        =   14414
      _ExtentY        =   1291
      TextAlias       =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
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
      Columns         =   5
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
      ColDesigner     =   "frmVATaxPManualBillEntry.frx":08CA
   End
   Begin LpLib.fpList fpPropList 
      Height          =   732
      Left            =   2724
      TabIndex        =   6
      Top             =   3120
      Width           =   8172
      _Version        =   196608
      _ExtentX        =   14414
      _ExtentY        =   1291
      TextAlias       =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
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
      ColDesigner     =   "frmVATaxPManualBillEntry.frx":0C6A
   End
   Begin LpLib.fpCombo fpcmbBillType 
      Height          =   372
      Left            =   7524
      TabIndex        =   1
      Top             =   1320
      Width           =   3420
      _Version        =   196608
      _ExtentX        =   6032
      _ExtentY        =   656
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
      ColDesigner     =   "frmVATaxPManualBillEntry.frx":0FDE
   End
   Begin EditLib.fpCurrency fpCurrInt 
      Height          =   324
      Left            =   3564
      TabIndex        =   14
      Top             =   7320
      Width           =   1572
      _Version        =   196608
      _ExtentX        =   2773
      _ExtentY        =   572
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
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
      MinValue        =   "0"
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
      IncDec          =   1
      BorderGrayAreaColor=   -2147483637
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin VB.Timer MsgAlertTimer 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   -36
      Top             =   0
   End
   Begin EditLib.fpLongInteger fptxtAcctNum 
      Height          =   390
      Left            =   2724
      TabIndex        =   0
      Top             =   1320
      Width           =   1575
      _Version        =   196608
      _ExtentX        =   2778
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
      AlignTextH      =   1
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
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
      Text            =   "0"
      MaxValue        =   "2147483647"
      MinValue        =   "-2147483648"
      NegFormat       =   1
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
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
   Begin EditLib.fpCurrency fpCurrPers 
      Height          =   324
      Left            =   3564
      TabIndex        =   9
      Top             =   5520
      Width           =   1572
      _Version        =   196608
      _ExtentX        =   2773
      _ExtentY        =   572
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
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
      MinValue        =   "0"
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
      IncDec          =   1
      BorderGrayAreaColor=   -2147483637
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fptxtName 
      Height          =   390
      Left            =   2724
      TabIndex        =   2
      Top             =   1800
      Width           =   7095
      _Version        =   196608
      _ExtentX        =   12515
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
   Begin EditLib.fpDoubleSingle fpDblSngBillNum 
      Height          =   375
      Left            =   6084
      TabIndex        =   4
      Top             =   2280
      Width           =   1575
      _Version        =   196608
      _ExtentX        =   2778
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
      Text            =   "0"
      DecimalPlaces   =   0
      DecimalPoint    =   ""
      FixedPoint      =   0   'False
      LeadZero        =   0
      MaxValue        =   "999999999"
      MinValue        =   "0"
      NegFormat       =   1
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
   Begin EditLib.fpLongInteger fpLongTaxYear 
      Height          =   375
      Left            =   9324
      TabIndex        =   5
      Top             =   2280
      Width           =   1095
      _Version        =   196608
      _ExtentX        =   1931
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
      Text            =   "0"
      MaxValue        =   "2147483647"
      MinValue        =   "-2147483648"
      NegFormat       =   1
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
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
   Begin EditLib.fpDateTime fptxtBillDate 
      Height          =   375
      Left            =   2724
      TabIndex        =   3
      Top             =   2280
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
   Begin EditLib.fpCurrency fpCurrMT 
      Height          =   324
      Left            =   3564
      TabIndex        =   10
      Top             =   5880
      Width           =   1572
      _Version        =   196608
      _ExtentX        =   2773
      _ExtentY        =   572
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
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
      MinValue        =   "0"
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
      IncDec          =   1
      BorderGrayAreaColor=   -2147483637
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpCurrency fpCurrMC 
      Height          =   324
      Left            =   3564
      TabIndex        =   11
      Top             =   6240
      Width           =   1572
      _Version        =   196608
      _ExtentX        =   2773
      _ExtentY        =   572
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
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
      MinValue        =   "0"
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
      IncDec          =   1
      BorderGrayAreaColor=   -2147483637
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpCurrency fpCurrFE 
      Height          =   324
      Left            =   3564
      TabIndex        =   12
      Top             =   6600
      Width           =   1572
      _Version        =   196608
      _ExtentX        =   2773
      _ExtentY        =   572
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
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
      MinValue        =   "0"
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
      IncDec          =   1
      BorderGrayAreaColor=   -2147483637
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpCurrency fpCurrOpt1 
      Height          =   324
      Left            =   8724
      TabIndex        =   16
      Top             =   5880
      Width           =   1572
      _Version        =   196608
      _ExtentX        =   2773
      _ExtentY        =   572
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
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
      MinValue        =   "0"
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
      IncDec          =   1
      BorderGrayAreaColor=   -2147483637
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpCurrency fpCurrOpt2 
      Height          =   324
      Left            =   8724
      TabIndex        =   17
      Top             =   6240
      Width           =   1572
      _Version        =   196608
      _ExtentX        =   2773
      _ExtentY        =   572
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
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
      MinValue        =   "0"
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
      IncDec          =   1
      BorderGrayAreaColor=   -2147483637
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpCurrency fpCurrOpt3 
      Height          =   324
      Left            =   8724
      TabIndex        =   18
      Top             =   6600
      Width           =   1572
      _Version        =   196608
      _ExtentX        =   2773
      _ExtentY        =   572
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
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
      MinValue        =   "0"
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
      IncDec          =   1
      BorderGrayAreaColor=   -2147483637
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fptxtActiveProp 
      Height          =   396
      Left            =   3804
      TabIndex        =   8
      Top             =   5040
      Width           =   7092
      _Version        =   196608
      _ExtentX        =   12515
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
   Begin EditLib.fpCurrency fpCurrCredit 
      Height          =   324
      Left            =   8724
      TabIndex        =   19
      Top             =   6960
      Visible         =   0   'False
      Width           =   1572
      _Version        =   196608
      _ExtentX        =   2773
      _ExtentY        =   572
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
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
      ControlType     =   1
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   -1
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   ""
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "-900000000"
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
      IncDec          =   1
      BorderGrayAreaColor=   -2147483637
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdLookup 
      Height          =   375
      Left            =   4410
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1320
      Width           =   1800
      _Version        =   131072
      _ExtentX        =   3175
      _ExtentY        =   661
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
      ButtonDesigner  =   "frmVATaxPManualBillEntry.frx":130D
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPropDet 
      Height          =   360
      Left            =   570
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   3810
      Width           =   2040
      _Version        =   131072
      _ExtentX        =   3598
      _ExtentY        =   635
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
      ButtonDesigner  =   "frmVATaxPManualBillEntry.frx":14EF
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSave 
      Height          =   492
      Left            =   7848
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   8040
      Width           =   1572
      _Version        =   131072
      _ExtentX        =   2773
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
      ButtonDesigner  =   "frmVATaxPManualBillEntry.frx":16D5
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   492
      Left            =   2520
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   8040
      Width           =   1572
      _Version        =   131072
      _ExtentX        =   2773
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
      ButtonDesigner  =   "frmVATaxPManualBillEntry.frx":18B1
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdMessage 
      Height          =   495
      Left            =   4290
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   8040
      Width           =   1575
      _Version        =   131072
      _ExtentX        =   2778
      _ExtentY        =   873
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
      ButtonDesigner  =   "frmVATaxPManualBillEntry.frx":1A8D
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdDelete 
      Height          =   492
      Left            =   6060
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   8040
      Width           =   1572
      _Version        =   131072
      _ExtentX        =   2773
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
      ButtonDesigner  =   "frmVATaxPManualBillEntry.frx":1C6B
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdCancel 
      Height          =   495
      Left            =   9990
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   8040
      Visible         =   0   'False
      Width           =   1575
      _Version        =   131072
      _ExtentX        =   2778
      _ExtentY        =   873
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
      ButtonDesigner  =   "frmVATaxPManualBillEntry.frx":1E48
   End
   Begin EditLib.fpCurrency fpCurrMH 
      Height          =   324
      Left            =   3564
      TabIndex        =   13
      Top             =   6960
      Width           =   1572
      _Version        =   196608
      _ExtentX        =   2773
      _ExtentY        =   572
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
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
      MinValue        =   "0"
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
      IncDec          =   1
      BorderGrayAreaColor=   -2147483637
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpCurrency fpCurrPen 
      Height          =   324
      Left            =   8724
      TabIndex        =   15
      Top             =   5520
      Width           =   1572
      _Version        =   196608
      _ExtentX        =   2773
      _ExtentY        =   572
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
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
      MinValue        =   "0"
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
      IncDec          =   1
      BorderGrayAreaColor=   -2147483637
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin VB.Label Label24 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Penalty Amount:"
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
      Height          =   252
      Left            =   6564
      TabIndex        =   53
      Top             =   5604
      Width           =   2052
   End
   Begin VB.Label Label23 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Interest Amount:"
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
      Height          =   252
      Left            =   1404
      TabIndex        =   52
      Top             =   7380
      Width           =   1932
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Mob Homes Amount:"
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
      Height          =   252
      Left            =   1404
      TabIndex        =   51
      Top             =   7032
      Width           =   1932
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   540
      Index           =   1
      Left            =   1584
      Top             =   225
      Width           =   8655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Manual Personal Tax Bill Entry"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3234
      TabIndex        =   50
      Top             =   270
      Width           =   5295
   End
   Begin VB.Label Label69 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name:"
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
      Height          =   270
      Left            =   684
      TabIndex        =   49
      Top             =   1890
      Width           =   1860
   End
   Begin VB.Label Label72 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Property Listings:"
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
      Height          =   396
      Left            =   564
      TabIndex        =   48
      Top             =   3348
      Width           =   1980
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Bill #:"
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
      Height          =   270
      Left            =   4644
      TabIndex        =   47
      Top             =   2370
      Width           =   1260
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Year:"
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
      Height          =   255
      Left            =   8004
      TabIndex        =   46
      Top             =   2385
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Billing Date:"
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
      Left            =   1209
      TabIndex        =   45
      Top             =   2370
      Width           =   1335
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Personal Amount:"
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
      Height          =   252
      Left            =   1284
      TabIndex        =   44
      Top             =   5604
      Width           =   2052
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Mach Tools Amount:"
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
      Height          =   216
      Left            =   1404
      TabIndex        =   43
      Top             =   5964
      Width           =   1932
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Merch Cap Amount:"
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
      Height          =   336
      Left            =   1524
      TabIndex        =   42
      Top             =   6324
      Width           =   1812
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Farm Equip Amount:"
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
      Height          =   252
      Left            =   1404
      TabIndex        =   41
      Top             =   6660
      Width           =   1932
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Opt 1 Amount:"
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
      Height          =   252
      Left            =   5364
      TabIndex        =   40
      Top             =   5964
      Width           =   3252
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Opt 2 Amount:"
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
      Height          =   240
      Left            =   5364
      TabIndex        =   39
      Top             =   6324
      Width           =   3252
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Opt 3 Amount:"
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
      Height          =   240
      Left            =   5364
      TabIndex        =   38
      Top             =   6660
      Width           =   3252
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   2892
      Left            =   444
      Top             =   4920
      Width           =   10812
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Type:"
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
      Left            =   6324
      TabIndex        =   37
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Customer And Transaction Data"
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
      Left            =   444
      TabIndex        =   36
      Top             =   960
      Width           =   3780
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Tax Amounts"
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
      Height          =   276
      Left            =   444
      TabIndex        =   35
      Top             =   4920
      Width           =   1548
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Property Data"
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
      Left            =   444
      TabIndex        =   34
      Top             =   2760
      Width           =   1785
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Account #:"
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
      Height          =   255
      Left            =   729
      TabIndex        =   33
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   444
      X2              =   11244
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "In Edit Listings:"
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
      Height          =   396
      Left            =   564
      TabIndex        =   32
      Top             =   4400
      Width           =   1980
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PIN #"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   3444
      TabIndex        =   31
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TYPE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   4764
      TabIndex        =   30
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ADDRESS OR PERSONAL CATEGORIES"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   6204
      TabIndex        =   29
      Top             =   2880
      Width           =   3735
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   2604
      X2              =   10884
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   3972
      Left            =   444
      Top             =   960
      Width           =   10812
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Active Property:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   372
      Left            =   1644
      TabIndex        =   28
      Top             =   5160
      Width           =   2052
   End
   Begin VB.Label lblCredit 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Credit Being  Applied:"
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
      Height          =   252
      Left            =   6120
      TabIndex        =   27
      Top             =   7032
      Visible         =   0   'False
      Width           =   2496
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   660
      Left            =   1584
      Top             =   120
      Width           =   8655
   End
End
Attribute VB_Name = "frmVATaxPManualBillEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  'Private Temp_Class As Resize_Class
  Public RealRec As Long
  Public PersRec As Long
  Dim BtnFnt As Double
  Dim EditCnt As Integer
  Dim InEdit() As Long
  Dim InEditM() As Long
  Dim TempAcctNum As Long
  Dim TempTransDate As Integer
  Dim TempTaxYear As Integer
  Dim TempPers As Double
  Dim TempIntAmount As Double
  Dim TempMachTools As Double
  Dim TempMerchCap As Double
  Dim TempFarmEquip As Double
  Dim TempMobHomes As Double
  Dim TempPenalty As Double
  Dim TempOptRev1 As Double
  Dim TempOptRev2 As Double
  Dim TempOptRev3 As Double
  Dim TempBillType As String * 1   'R=REAL P=PERS C=COMB
  Dim TempBillNum As Long
  Dim TempSName As String * 50
  Dim ThisTaxYear As Integer
  Dim DontExit As Boolean
  Dim EditMode As Boolean
'  Dim TempRealRec As Long
  Dim TempPersRec As Long
  Dim InListActive As Boolean
  Dim LookUpMode As Boolean
  Public PostSaveLoad As Boolean
  Dim ExitOK As Boolean 'designed to keep fptxtAcctNum.LostFocus from activating
  
Private Sub cmdCancel_Click()
  Dim TaxMRec As TaxMTransactionType
  Dim TMHandle As Integer
  Dim NumOfTMRecs As Integer
  
  If fpCurrCredit.Value <> 0 Then
    fpCurrCredit.Value = 0
    OpenTaxManualBillFile TMHandle, NumOfTMRecs
    Get TMHandle, ThisMRec, TaxMRec
    TaxMRec.OverPayUsed = 0
    Put TMHandle, ThisMRec, TaxMRec
    Close TMHandle
    Call TaxMsg(900, "The credit amount has been cancelled.")
    lblCredit.Visible = False
    cmdCancel.Visible = False
    fpCurrCredit.Visible = False
  End If
End Sub

Private Sub cmdDelete_Click()
  Dim TaxMRec As TaxMTransactionType
  Dim TMHandle As Integer
  Dim NumOfTMRecs As Integer
  Dim ThisRec As Integer
  Dim x As Integer
  Dim ThisPin$
  
  On Error GoTo ERRORSTUFF
  
  ThisRec = fpListInEdit.ListIndex
  If ThisRec = -1 Then
    Call TaxMsg(900, "No edited property has been selected.")
    Exit Sub
  End If
  
  If TaxMsgWOpts(900, "Are you sure you want to delete this bill? Press F10 to delete. Otherwise, press ESC to abort.", "F10 Delete", "ESC Abort") = "abort" Then
    Unload frmVATaxMsgWOpts
    fptxtAcctNum.SetFocus
    Exit Sub
  Else
    Unload frmVATaxMsgWOpts
  End If
    
  fpListInEdit.Row = fpListInEdit.ListIndex
  fpListInEdit.Col = 3
  
  ThisRec = CLng(fpListInEdit.ColText)
  fpListInEdit.Col = 0
  ThisPin = QPTrim$(fpListInEdit.ColText)
  If ThisPin = "" Then ThisPin = "N/A"
  OpenTaxManualBillFile TMHandle, NumOfTMRecs
  Get TMHandle, ThisRec, TaxMRec
  TaxMRec.Deleted = True
  Put TMHandle, ThisRec, TaxMRec
  Close TMHandle
  
  If NumOfTMRecs = 1 Then
    KillFile TaxManualBill '5.16.07
  End If
  
  Call TaxMsg(900, "The selected manual personal bill was deleted successfully.")
  PostSaveLoad = True
  Call EnterEditCheck
  PostSaveLoad = False
  lblCredit.Visible = False
  cmdCancel.Visible = False
  fpCurrCredit.Visible = False
  MainLog ("Manual personal bill with property pin # " + ThisPin + " deleted.")
  Call ClearBillFields
  Call AssignTemps
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPManualBillEntry", "cmdDelete_Click", Erl)
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

Private Sub cmdExit_Click()
  Dim ThisTot As Double
  
  If InListActive = True Then
    If Check4Changes = True Then
      Exit Sub
    End If
  Else
    ThisTot = OldRound(CDbl(fpCurrPers.Value) + CDbl(fpCurrInt.Value) + CDbl(fpCurrMT.Value))
    ThisTot = OldRound(ThisTot + CDbl(fpCurrMC.Value) + CDbl(fpCurrFE.Value) + CDbl(fpCurrMH.Value) + CDbl(fpCurrOpt1.Value))
    ThisTot = OldRound(ThisTot + CDbl(fpCurrOpt2.Value) + CDbl(fpCurrOpt3.Value) + CDbl(fpCurrPen.Value))
    If ThisTot > 0 Then
      If Check4Changes = True Then
        Exit Sub
      End If
    End If
  End If
  
  KillFile "C:\CPWork\pmanualbill.dat"
  TempAcctNum = 0
  ExitOK = True
  If Exist("C:\CPWork\manualedit.dat") Then
    Call frmVATaxManualBillEdit.ClearAndUpdateList
    Unload Me
    DoEvents
    Exit Sub
  Else
    frmVATaxManualBillMenu.Show
    DoEvents
    Unload Me
  End If
End Sub

Private Sub cmdLookup_Click()
  LookUpMode = True
  If Check4Changes = True Then
    Exit Sub
  End If
  LookUpMode = False
  frmVATaxCustLookup.Show
  DoEvents
End Sub

Private Sub cmdMessage_Click()
   If GCustNum > 0 Then
    frmVATaxMessage.Show vbModal
  End If
End Sub

Private Sub cmdPropDet_Click()
  Dim ThisClass$
  
  On Error GoTo ERRORSTUFF
  
  If fpListInEdit.ListIndex <> -1 Then
    fpListInEdit.Row = fpListInEdit.ListIndex
    fpListInEdit.Col = 1
    ThisClass = QPTrim$(fpListInEdit.ColText)
    fpListInEdit.Col = 4
    If ThisClass = "PERSONAL" Then
      PersRec = CLng(fpListInEdit.ColText)
      frmVATaxPersDetail.Show vbModal
      Exit Sub
'    ElseIf ThisClass = "REAL" Then
'      RealRec = CLng(fpListInEdit.ColText)
'      frmVATaxRealDetail.Show vbModal
'      Exit Sub
'    ElseIf ThisClass = "MOCK REAL" Then
'      Call TaxMsg(900, "The MOCK REAL classification has no detail data.")
'      Exit Sub
    Else
      Call TaxMsg(800, "The classification for the selected property cannot be determined. Detail data cannot be loaded.")
      Exit Sub
    End If
  ElseIf fpPropList.ListIndex <> -1 Then
    fpPropList.Row = fpPropList.ListIndex
    fpPropList.Col = 1
    ThisClass = QPTrim$(fpPropList.ColText)
    fpPropList.Col = 3
    If ThisClass = "PERSONAL" Then
      PersRec = CLng(fpPropList.ColText)
      frmVATaxPersDetail.Show vbModal
      Exit Sub
'    ElseIf ThisClass = "REAL" Then
'      RealRec = CLng(fpPropList.ColText)
'      frmVATaxRealDetail.Show vbModal
'      Exit Sub
'    ElseIf ThisClass = "MOCK REAL" Then
'      Call TaxMsg(900, "The MOCK REAL classification has no detail data.")
'      Exit Sub
    Else
      Call TaxMsg(800, "The classification for the selected property cannot be determined. Detail data cannot be loaded.")
      Exit Sub
    End If
  Else
    If CDbl(fptxtAcctNum.Value) = 0 Then
      Call TaxMsg(900, "Please enter a valid customer number.")
      fptxtAcctNum.SetFocus
    Else
      Call TaxMsg(900, "Please make a property selection.")
    End If
  End If
   
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPManualBillEntry", "cmdPropDet_Click", Erl)
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

Private Sub cmdSave_Click()
  Dim TaxMRec As TaxMTransactionType
  Dim TMHandle As Integer
  Dim NumOfTMRecs As Integer
  Dim ThisYear As Integer
  Dim ThatYear As Integer
  Dim TaxMasterRec As TaxMasterType
  Dim TXHandle As Integer
  Dim TotalBill As Double
  Dim SaveHere As Long
  Dim ThisRec As Long
  Dim ThatRec As Long
  Dim RealPropsOK As Boolean
  Dim PersPropsOK As Boolean
  Dim x As Long
  Dim ThisClass$, SaveRec As Long
  Dim OverAmt As Double
  
  On Error GoTo ERRORSTUFF
  
  If fptxtAcctNum.Value = 0 Then
    Call TaxMsg(900, "Please enter a valid customer account number.")
    fptxtAcctNum.SetFocus
    Exit Sub
  End If
  
  If fptxtActiveProp.Text = "NOTHING SELECTED" Then
    Call TaxMsg(900, "No property selected for manual personal billing.")
    Exit Sub
  End If
  
  OpenTaxSetUpFile TXHandle
  Get TXHandle, 1, TaxMasterRec
  Close TXHandle
  ThisYear = TaxMasterRec.PTaxYear
  ThatYear = fpLongTaxYear
  If Abs(ThisYear - ThatYear) > 2 Then
    If TaxMsgWOpts(900, "The tax year entered is " + Using("###0", ThatYear) + ". If this is correct press F10 to continue. Otherwise, press ESC to edit.", "F10 Continue", "ESC Edit") = "abort" Then
      Unload frmVATaxMsgWOpts
      Close
      fpLongTaxYear.SetFocus
      Exit Sub
    Else
      Unload frmVATaxMsgWOpts
      MainLog ("User warned that the year entered " + Using("###0", ThatYear) + " might be suspect and elected to continue processing anyway.")
    End If
  End If
  
  If RevsAndGLsOK(Me, ThatYear, "P") = False Then
    Exit Sub
  End If

  TotalBill = 0
  TotalBill = OldRound(CDbl(fpCurrPers.Value) + CDbl(fpCurrInt.Value) + CDbl(fpCurrMT.Value) + CDbl(fpCurrMC.Value))
  TotalBill = OldRound(TotalBill + CDbl(fpCurrFE.Value) + CDbl(fpCurrMH.Value) + (fpCurrOpt1.Value))
  TotalBill = OldRound(TotalBill + CDbl(fpCurrOpt2.Value) + CDbl(fpCurrOpt3.Value) + CDbl(fpCurrPen.Value))
  
  If TotalBill = 0 Then
    Call TaxMsg(900, "The total equals zero. Save aborted.")
    Close
    fptxtAcctNum.SetFocus
    Exit Sub
  End If
  
  If fpDblSngBillNum.Value = 0 Then
    Call TaxMsg(700, "The tax bill number entered is zero. All bills are required to have a bill number greater than zero for the program to recognize this bill in other program functions.")
    MainLog ("User warned that the manual tax bill number entered is zero thereby aborting the save process.")
    Close
    fpDblSngBillNum.SetFocus
    Exit Sub
  End If
  
  If fpCurrCredit.Value = 0 Then
    If Look4TempCreditUsed = False Then 'can only use once
      OverAmt = GetCustPersBalance(GCustNum, -1)
      If OverAmt < 0 Then
        If TaxMsgWOpts(700, "This customer has a personal credit balance of " + QPTrim$(Using$("$##,##0.00", OverAmt)) + ". If you wish to apply the credit to this bill then press F10. Otherwise, press ESC to ignore the credit amount.", "F10 Apply", "ESC Ignore") = "abort" Then
          Unload frmVATaxMsgWOpts
          OverAmt = 0
          MainLog ("A personal credit balance of " + QPTrim$(Using$("$##,##0.00", OverAmt)) + " was not applied to this personal manual bill.")
        Else
          Unload frmVATaxMsgWOpts
          Call TaxMsg(900, "A personal credit of " + QPTrim$(Using$("$##,##0.00", OverAmt)) + " will be applied at posting for this personal manual bill.")
          lblCredit.Visible = True
          fpCurrCredit.Visible = True
          fpCurrCredit = OverAmt 'TaxMRec.OverPayUsed
          cmdCancel.Visible = True
        End If
      Else
        OverAmt = 0
      End If
    End If
  Else
    OverAmt = CDbl(fpCurrCredit.Value)
  End If
  
  OpenTaxManualBillFile TMHandle, NumOfTMRecs
  SaveHere = 0
  SaveRec = 0
  If EditMode = True Then
    Get TMHandle, ThisMRec, TaxMRec
    SaveHere = ThisMRec
    fpListInEdit.Row = fpListInEdit.ListIndex
    fpListInEdit.Col = 1
    ThisClass = fpListInEdit.ColText
    fpListInEdit.Col = 4
    SaveRec = CLng(fpListInEdit.ColText)
  Else
    SaveHere = NumOfTMRecs + 1
    fpPropList.Row = fpPropList.ListIndex
    fpPropList.Col = 1
    ThisClass = fpPropList.ColText
    fpPropList.Col = 3
    SaveRec = CLng(fpPropList.ColText)
  End If
  TaxMRec.Account = fptxtAcctNum
  TaxMRec.TransDate = Date2Num(fptxtBillDate.Text)
  TaxMRec.TaxYear = fpLongTaxYear
  TaxMRec.Desc = "M Tax Bill #" + CStr(fpDblSngBillNum)
  TaxMRec.Personal = fpCurrPers
  TaxMRec.IntAmount = fpCurrInt
  TaxMRec.MachTools = fpCurrMT
  TaxMRec.MerchCap = fpCurrMC
  TaxMRec.FarmEquip = fpCurrFE
  TaxMRec.MobHomes = fpCurrMH
  TaxMRec.Penalty = fpCurrPen
  TaxMRec.OptRev1 = fpCurrOpt1
  TaxMRec.OptRev2 = fpCurrOpt2
  TaxMRec.OptRev3 = fpCurrOpt3
  TaxMRec.TaxAmount = 0
  TaxMRec.AdColAmount = 0
  TaxMRec.LateList = 0
  TaxMRec.OverPayUsed = OverAmt
  TaxMRec.BillType = Mid(fpcmbBillType.Text, 1, 1)
  TaxMRec.BillNum = fpDblSngBillNum.Value
  TaxMRec.SName = QPTrim$(fptxtName.Text)
  TaxMRec.TName = QPTrim$(fptxtName.Text)
  TaxMRec.Deleted = 0
  TaxMRec.Class = ThisClass
  TaxMRec.PersRec = SaveRec
  TaxMRec.RealRec = 0
  TaxMRec.Padding = ""
  Put TMHandle, SaveHere, TaxMRec
  Close TMHandle
  
  Call Savemsg(900, "Manual personal bill data has been saved successfully.")
  
  If LookUpMode = True Then
    Call frmVATaxManualBillEdit.ClearAndUpdateList
    Exit Sub
  End If
    
  If TaxMsgWOpts(800, "Press F10 to enter another manual personal tax bill. Otherwise, press ESC to return to the Manual Tax Billing Menu.", "F10 Enter New", "ESC Exit") = "abort" Then
    Unload frmVATaxMsgWOpts
    If LookUpMode = True Then
      Unload frmVATaxManualBillEdit
    End If
    KillFile "C:\CPWork\pmanualbill.dat"
    KillFile "C:\CPWork\manualedit.dat"
    frmVATaxManualBillMenu.Show
    Call ClearBillFields
    TempAcctNum = 0
    DoEvents
    Unload Me
    Exit Sub
  Else
    Unload frmVATaxMsgWOpts
    Call ClearBillFields
    Call AssignTemps
    fpListInEdit.ListIndex = -1
    fpPropList.ListIndex = -1
    fptxtAcctNum.SetFocus
    PostSaveLoad = True
    Call EnterEditCheck
    PostSaveLoad = False
  End If
     
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPManualBillEntry", "cmdSave_Click", Erl)
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
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%E"
      Call cmdExit_Click
      KeyCode = 0
    Case vbKeyF7:
      SendKeys "%L"
      Call cmdLookup_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%S"
      Call cmdSave_Click
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
      KillFile "C:\CPWork\pmanualbill.dat"
      MainLog ("CitiTaxes.exe terminated via menu bar on frmVATaxPManualBillEntry.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    'Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
    DoEvents
  End If
End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  'Set Temp_Class = New Resize_Class
  'Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  MainLog ("User opened frmVATaxPManualBillEntry.")
  PersRec = 0
  RealRec = 0
  ExitOK = False
  InListActive = False
  LookUpMode = False
  PostSaveLoad = False
  Me.HelpContextID = hlpEnterTaxBill
  Call LoadMe
End Sub
Public Sub LoadMeWOEdit()
  Dim TaxCustRec As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim ThisOpt$
  Dim x As Long, GotIt As Boolean
  Dim PersPropRec As PersonalRecType
  Dim PHandle As Integer
  Dim NumOfPersRecs As Long
  Dim PersList$
  
  On Error GoTo ERRORSTUFF
  
  If GCustNum > 0 Then
    OpenTaxCustFile TCHandle, NumOfTCRecs
    Get TCHandle, GCustNum, TaxCustRec
    Close TCHandle
    fptxtName.Text = QPTrim$(TaxCustRec.CustName)
    fptxtAcctNum.Text = CStr(TaxCustRec.Acct)
    If CustHasMsg(GCustNum) Then
      MsgAlertTimer.Enabled = True
    Else
      MsgAlertTimer.Enabled = False
      cmdMessage.ForeColor = &H80000012
    End If
    Call AssignTemps
  End If
  
  fpLongTaxYear = ThisTaxYear
  fptxtBillDate = Date
  
  fpListInEdit.Clear
  fpPropList.Clear
  
  GotIt = False
  PersList = ""
  OpenPersPropFile PHandle, NumOfPersRecs
  For x = 1 To NumOfPersRecs
    Get PHandle, x, PersPropRec
    If TaxCustRec.PIN = PersPropRec.CustPin Then
      If PersPropRec.CVALUE > 0 Then
        PersList = "FARM EQ" '7
      End If
      If PersPropRec.MCValue > 0 Then
        If QPTrim$(PersList) = "" Then
          PersList = "MER CAP"
        Else
          PersList = PersList + ", MER CAP" '16
        End If
      End If
      If PersPropRec.MHValue > 0 Then
        If QPTrim$(PersList) = "" Then
          PersList = "MOB HM"
        Else
          PersList = PersList + ", MOB HM" '24
        End If
      End If
      If PersPropRec.MTValue > 0 Then
        If QPTrim$(PersList) = "" Then
          PersList = "MCH TLS"
        Else
          PersList = PersList + ", MCH TLS" '33
        End If
      End If
      If PersPropRec.PersVal > 0 Then
        If QPTrim$(PersList) = "" Then
          PersList = "PERSONAL"
        Else
          PersList = PersList + ", PERSONAL" '41
        End If
      End If
      fpPropList.InsertRow = QPTrim$(PersPropRec.PropPin) + Chr(9) + "PERSONAL" + Chr(9) + PersList + Chr(9) + CStr(x)
    End If
    PersList = ""
  Next x
    
  Close PHandle
  
  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPManualBillEntry", "LoadMeWOEdit", Erl)
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
Private Sub LoadMe()
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim ThisOpt$
  
  On Error GoTo ERRORSTUFF
  
  fptxtActiveProp.Text = "NOTHING SELECTED"
  EditMode = False
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  ThisOpt = QPTrim$(TaxMasterRec.POptRev1)
  If ThisOpt <> "" Then
    Label9.Caption = ThisOpt + ":"
  Else
    Label9.Caption = "NOT IN USE:"
    fpCurrOpt1.Enabled = False
  End If
  
  ThisOpt = QPTrim$(TaxMasterRec.POptRev2)
  If ThisOpt <> "" Then
    Label10.Caption = ThisOpt + ":"
  Else
    Label10.Caption = "NOT IN USE:"
    fpCurrOpt2.Enabled = False
  End If
  
  ThisOpt = QPTrim$(TaxMasterRec.POptRev3)
  If ThisOpt <> "" Then
    Label11.Caption = ThisOpt + ":"
  Else
    fpCurrOpt3.Enabled = False
    Label11.Caption = "NOT IN USE:"
  End If
  ThisTaxYear = TaxMasterRec.PTaxYear
  
  fpcmbBillType.Text = "PERSONAL ONLY"
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPManualBillEntry", "LoadMe", Erl)
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

Private Sub fpcmbBillType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbBillType.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbBillType.ListIndex = -1
  End If
  If fpcmbBillType.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fptxtName.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

'Private Sub fpcmbPersList_KeyDown(KeyCode As Integer, Shift As Integer)
'  If KeyCode = vbKeySpace Then
'    fpcmbPersList.ListDown = True
'  End If
'  If KeyCode = vbKeyDelete Then
'    fpcmbPersList.ListIndex = -1
'  End If
'  If fpcmbPersList.ListDown <> True Then
'    If KeyCode = vbKeyDown Then
'      fptxtAcctNum.SetFocus
'      KeyCode = 0
'    Else
'      If KeyCode = vbKeyUp Then
'        SendKeys "+{Tab}"
'        KeyCode = 0
'      End If
'    End If
'  End If
'
'End Sub

'Private Sub fpcmbPropList_KeyDown(KeyCode As Integer, Shift As Integer)
'  If KeyCode = vbKeySpace Then
'    fpcmbPropList.ListDown = True
'  End If
'  If KeyCode = vbKeyDelete Then
'    fpcmbPropList.ListIndex = -1
'  End If
'  If fpcmbPropList.ListDown <> True Then
'    If KeyCode = vbKeyDown Then
'      fpCurrPropVal.SetFocus
'      KeyCode = 0
'    Else
'      If KeyCode = vbKeyUp Then
'        SendKeys "+{Tab}"
'        KeyCode = 0
'      End If
'    End If
'  End If
'
'End Sub

Private Function Check4ValidCustNum(ThisCust As Long) As Boolean
  Dim TaxRec As TaxCustType
  Dim CHandle As Integer
  Dim NumOfCRecs As Long
  Dim x As Long
  Dim Number$
  Dim Name$
  Dim Found As Boolean

  On Error GoTo ERRORSTUFF
  
  Check4ValidCustNum = True
  
  If fptxtAcctNum.Value = 0 Then
    Check4ValidCustNum = False
    Exit Function
  End If
  
  OpenTaxCustFile CHandle, NumOfCRecs
  
  If NumOfCRecs = 0 Then
    frmVATaxMsg.Label1.Caption = "There are no tax customers saved."
    frmVATaxMsg.Label1.Top = 900
    frmVATaxMsg.Show vbModal
    Close CHandle
    Exit Function
  End If
  
  For x = 1 To NumOfCRecs
    Get CHandle, x, TaxRec
    If ThisCust = TaxRec.Acct Then
      If TaxRec.Deleted <> 0 Then
        Check4ValidCustNum = False
      End If
      Exit For
    End If
  Next x

  Close CHandle

  If x > NumOfCRecs Then
    Call Clearscreen
    Check4ValidCustNum = False
  End If
    
  Exit Function
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPManualBillEntry", "Check4ValidCustNum", Erl)
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
  
End Function

Public Sub Clearscreen()
  fpListInEdit.Clear
  InListActive = False
  fpPropList.Clear
  ThisMRec = 0
  fptxtAcctNum.Text = "0"
  fpDblSngBillNum.Text = "0"
  fptxtName.Text = ""
  fpLongTaxYear = ThisTaxYear
  fptxtBillDate = Date
  fpCurrPers = 0
  fpCurrInt = 0
  fpCurrMT = 0
  fpCurrMC = 0
  fpCurrFE = 0
  fpCurrMH = 0
  fpCurrPen = 0
  fpCurrOpt1 = 0
  fpCurrOpt2 = 0
  fpCurrOpt3 = 0
  fpcmbBillType.Text = "PERSONAL ONLY" '"COMBINED"
  TempAcctNum = 0
  TempTransDate = Date2Num(Date)
  TempTaxYear = ThisTaxYear
  TempPers = 0
  TempIntAmount = 0
  TempMachTools = 0
  TempMerchCap = 0
  TempFarmEquip = 0
  TempMobHomes = 0
  TempPenalty = 0
  TempOptRev1 = 0
  TempOptRev2 = 0
  TempOptRev3 = 0
  TempBillType = "P" 'Mid(fpcmbBillType.Text, 1, 1)
  TempSName = ""
End Sub

Private Sub AssignTemps()
  TempAcctNum = fptxtAcctNum
  TempTransDate = Date2Num(fptxtBillDate.Text)
  TempTaxYear = fpLongTaxYear.Value
  TempPers = fpCurrPers
  TempIntAmount = fpCurrInt
  TempMachTools = fpCurrMT
  TempMerchCap = fpCurrMC
  TempFarmEquip = fpCurrFE
  TempMobHomes = fpCurrMH
  TempPenalty = fpCurrPen
  TempOptRev1 = fpCurrOpt1
  TempOptRev2 = fpCurrOpt2
  TempOptRev3 = fpCurrOpt3
  TempBillType = Mid(fpcmbBillType.Text, 1, 1)
  TempBillNum = fpDblSngBillNum.Value
  TempSName = QPTrim$(fptxtName.Text)
End Sub

Public Sub LoadMeEdit()
  Dim TaxCustRec As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim TaxMRec As TaxMTransactionType
  Dim TMHandle As Integer
  Dim NumOfTMRecs As Integer
  Dim ThisOpt$
  Dim x As Long, GotIt As Boolean
  Dim PersPropRec As PersonalRecType
  Dim PHandle As Integer
  Dim NumOfPersRecs As Long
  Dim y As Integer, TextLine$
  Dim PersList$, ThisClass$
  Dim EditLoadCnt As Integer
  Dim HighLight As Boolean
  
  On Error GoTo ERRORSTUFF
  
  HighLight = False
  EditLoadCnt = 0
  OpenTaxManualBillFile TMHandle, NumOfTMRecs
  If ThisMRec <> 0 Then
    Get TMHandle, ThisMRec, TaxMRec
    fptxtName.Text = QPTrim$(TaxMRec.SName)
    fptxtAcctNum.Text = CStr(TaxMRec.Account)
    fpLongTaxYear = TaxMRec.TaxYear
    fptxtBillDate = MakeRegDate(TaxMRec.TransDate)
    fpCurrPers = TaxMRec.Personal
    fpCurrInt = TaxMRec.IntAmount
    fpCurrMT = TaxMRec.MachTools
    fpCurrMC = TaxMRec.MerchCap
    fpDblSngBillNum = TaxMRec.BillNum
    fpCurrFE = TaxMRec.FarmEquip
    fpCurrMH = TaxMRec.MobHomes
    fpCurrPen = TaxMRec.Penalty
    fpCurrOpt1 = TaxMRec.OptRev1
    fpCurrOpt2 = TaxMRec.OptRev2
    fpCurrOpt3 = TaxMRec.OptRev3
    fpcmbBillType.Text = "PERSONAL ONLY"
    If TaxMRec.OverPayUsed <> 0 Then
      lblCredit.Visible = True
      fpCurrCredit.Visible = True
      fpCurrCredit = TaxMRec.OverPayUsed
      cmdCancel.Visible = True
    Else
      lblCredit.Visible = False
      fpCurrCredit.Visible = False
      cmdCancel.Visible = False
    End If
    Call AssignTemps
  End If
  
  If TempAcctNum = fptxtAcctNum And PostSaveLoad = False Then Exit Sub
  
  fpLongTaxYear = ThisTaxYear
  fptxtBillDate = Date
  If GCustNum > 0 Then
    OpenTaxCustFile TCHandle, NumOfTCRecs
    Get TCHandle, GCustNum, TaxCustRec
    Close TCHandle
    fptxtName.Text = QPTrim$(TaxCustRec.CustName)
    fptxtAcctNum.Text = CStr(TaxCustRec.Acct)
    If CustHasMsg(GCustNum) Then
      MsgAlertTimer.Enabled = True
    Else
      MsgAlertTimer.Enabled = False
      cmdMessage.ForeColor = &H80000012
    End If
    Call AssignTemps
  End If
  
  fpListInEdit.Clear
  fpPropList.Clear
  
  OpenPersPropFile PHandle, NumOfPersRecs
  For x = 1 To NumOfPersRecs
    Get PHandle, x, PersPropRec
    If TaxCustRec.PIN = PersPropRec.CustPin Then
      If PersPropRec.CVALUE > 0 Then
        PersList = "FARM EQ" '7
      End If
      If PersPropRec.MCValue > 0 Then
        If QPTrim$(PersList) = "" Then
          PersList = "MER CAP"
        Else
          PersList = PersList + ", MER CAP" '16
        End If
      End If
      If PersPropRec.MHValue > 0 Then
        If QPTrim$(PersList) = "" Then
          PersList = "MOB HM"
        Else
          PersList = PersList + ", MOB HM" '24
        End If
      End If
      If PersPropRec.MTValue > 0 Then
        If QPTrim$(PersList) = "" Then
          PersList = "MCH TLS"
        Else
          PersList = PersList + ", MCH TLS" '33
        End If
      End If
      If PersPropRec.PersVal > 0 Then
        If QPTrim$(PersList) = "" Then
          PersList = "PERSONAL"
        Else
          PersList = PersList + ", PERSONAL" '41
        End If
      End If
      For y = 1 To EditCnt
        Get TMHandle, InEditM(y), TaxMRec
        If TaxMRec.Deleted = True Then GoTo NotThisOne
'        If TaxMRec.PersRec = InEdit(y) Then
        If x = InEdit(y) Then
          Select Case TaxMRec.Class
            Case "P"
              ThisClass = "PERSONAL"
            Case Else
              ThisClass = "UNKNOWN"
          End Select
          fpListInEdit.AddItem QPTrim$(PersPropRec.PropPin) + Chr(9) + ThisClass + Chr(9) + PersList + Chr(9) + CStr(InEditM(y)) + Chr(9) + CStr(x)
          EditLoadCnt = EditLoadCnt + 1
          If InEditM(y) = ThisMRec Then GoSub HighlightPers
          Exit For
        End If
NotThisOne:
      Next y
      If y > EditCnt Then
        fpPropList.AddItem QPTrim$(PersPropRec.PropPin) + Chr(9) + "PERSONAL" + Chr(9) + PersList + Chr(9) + CStr(x)
      End If
    End If
    PersList = ""
  Next x
    
  Close PHandle
  Close TMHandle
  If EditLoadCnt > 0 And HighLight = False Then
    fpListInEdit.ListIndex = 0
  End If
  
  Exit Sub
  
HighlightPers:
  If PostSaveLoad = True And ThisMRec > 0 Then
    fptxtActiveProp.Text = QPTrim$(PersPropRec.PropPin) + "  " + ThisClass + "  " + PersList
    fpListInEdit.ListIndex = EditLoadCnt - 1
    HighLight = True
  End If
  
  Return
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPManualBillEntry", "LoadMeEdit", Erl)
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

Public Sub EnterEditCheck()
  Dim TaxMRec As TaxMTransactionType
  Dim TMHandle As Integer
  Dim NumOfTMRecs As Integer
  Dim x As Long
  Dim OpNum$
  Dim BillType$
  
  On Error GoTo ERRORSTUFF
  
  OpNum = CStr(OperNum)
  
  If Check4CustInPayBatch(GCustNum, OpNum, BillType$) = True Then
    If BillType = "R" Then GoTo GoAhead
    Call TaxMsg(700, "This customer is included in a personal payment file for operator #" + OpNum + " that has not been posted. Please either post this payment or delete this customer from the payment file.")
    fptxtAcctNum.Text = "0"
    fptxtAcctNum.SetFocus
    Exit Sub
  End If
  
GoAhead:
  ReDim InEdit(1 To 1) As Long
  ReDim InEditM(1 To 1) As Long
  EditCnt = 0
  If PostSaveLoad = False Then
    ThisMRec = 0
  End If
  If GCustNum > 0 Then
    OpenTaxManualBillFile TMHandle, NumOfTMRecs
    For x = 1 To NumOfTMRecs
      Get TMHandle, x, TaxMRec
      If TaxMRec.Deleted = True Then GoTo Deleted
      If TaxMRec.Account = GCustNum Then
        If TaxMRec.Class = "R" Then GoTo Deleted
        EditCnt = EditCnt + 1
        ReDim Preserve InEdit(1 To EditCnt) As Long
        If TaxMRec.Class = "P" Then
          InEdit(EditCnt) = TaxMRec.PersRec
        End If
        ReDim Preserve InEditM(1 To EditCnt) As Long
        InEditM(EditCnt) = x
      End If
Deleted:
    Next x
  End If
  
  Close TMHandle
  
  If EditCnt > 0 Then
    EditMode = True
    Call ClearBillFields
    Call LoadMeEdit
    If Not Exist("C:\CPWork\manualedit.dat") Then
      Call TaxMsg(900, "This customer is currently in personal edit mode.")
    End If
    Exit Sub
  Else
    EditMode = False
    Call Clearscreen
    Call ClearBillFields
    Call LoadMeWOEdit
  End If
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPManualBillEntry", "EnterEditCheck", Erl)
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

Private Sub fpListInEdit_Click()
  Dim TextLine$
  
  On Error GoTo ERRORSTUFF
  
  If PostSaveLoad = True Then Exit Sub
  fpPropList.Action = ActionDeselectAll
  If fpListInEdit.ListIndex = -1 Then
    fptxtActiveProp.Text = "NOTHING SELECTED"
    Exit Sub
  Else
    EditMode = True
  End If
  fpListInEdit.Row = fpListInEdit.ListIndex
  fpListInEdit.Col = 3
  If QPTrim$(fpListInEdit.ColText) = "" Then
    ThisMRec = 0
    InListActive = False
    Exit Sub
  End If
  ThisMRec = CLng(fpListInEdit.ColText)
  fpListInEdit.Col = 0
  TextLine = QPTrim$(fpListInEdit.ColText)
  fpListInEdit.Col = 1
  TextLine = TextLine + "  " + QPTrim$(fpListInEdit.ColText)
  fpListInEdit.Col = 2
  TextLine = TextLine + "  " + QPTrim$(fpListInEdit.ColText)
  fptxtActiveProp.Text = TextLine
  Call LoadMeEdit
  
  fpListInEdit.ListIndex = fpListInEdit.Row
  
  InListActive = True
  Call AssignTemps
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPManualBillEntry", "fpListInEdit_Click", Erl)
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

Private Sub fpPropList_Click()
  Dim TextLine$
  Dim ThisBill$
  fpListInEdit.Action = ActionDeselectAll
  If fpPropList.ListIndex = -1 Then
    fptxtActiveProp.Text = "NOTHING SELECTED"
    Exit Sub
  End If
  
  EditMode = False
  ThisBill = CStr(fpDblSngBillNum)
  Call ClearBillFields
  fpDblSngBillNum = CDbl(ThisBill)
  fpDblSngBillNum.SetFocus
  fpPropList.Col = 0
  TextLine = QPTrim$(fpPropList.ColText)
  fpPropList.Col = 1
  TextLine = TextLine + "  " + QPTrim$(fpPropList.ColText)
  fpPropList.Col = 2
  TextLine = TextLine + "  " + QPTrim$(fpPropList.ColText)
  fptxtActiveProp.Text = TextLine
  Call AssignTemps
End Sub

Private Sub fptxtAcctNum_LostFocus()

  On Error GoTo ERRORSTUFF
  
  If ExitOK = True Then Exit Sub
  If CLng(fptxtAcctNum.Text) = 0 Then Exit Sub
  If TempAcctNum = CLng(fptxtAcctNum.Value) Then Exit Sub
  
  If TempAcctNum <> 0 Then
    LookUpMode = True
    If Check4Changes = True Then
      Exit Sub
    End If
  End If
  LookUpMode = False
  
  If Check4ValidCustNum(fptxtAcctNum.Value) = False Then
    frmVATaxMsg.Label1.Caption = "The customer number is not valid. Please enter a valid customer number."
    frmVATaxMsg.Label1.Top = 800
    frmVATaxMsg.Show vbModal
    Call Clearscreen
    fptxtAcctNum.SetFocus
    Exit Sub
  End If
  GCustNum = fptxtAcctNum.Value
  Call EnterEditCheck
  
  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPManualBillEntry", "fptxtAcctNum_LostFocus", Erl)
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

Private Function Check4Changes() As Boolean
  Dim choice As String
  Dim ThisControl As Control
  Dim ThisDesc As String
  Dim ThatDesc As String
  Dim ThisText As String
  Dim ThisDbl As Double
  Dim ThatDbl As Double
  Dim ThisInt As Integer
  Dim ThatInt As Integer
  Dim ThisLong As Long
  Dim ThatLong As Long
  
  On Error GoTo ERRORSTUFF
  Check4Changes = False
  If fptxtAcctNum.Value = 0 Then Exit Function
  
  
  Set ThisControl = fpDblSngBillNum
  ThisLong = fpDblSngBillNum.Value
  ThatLong = TempBillNum
  If ThisLong <> ThatLong Then
    frmVATaxMsgW3Opts.Label1.Caption = "Changes have been made. Do you wish to save these changes. Press F10 to save, press F5 to review or press ESC to abandon all changes."
    frmVATaxMsgW3Opts.Label1.Top = 800
    frmVATaxMsgW3Opts.cmdCont.Text = "F10 Save"
    frmVATaxMsgW3Opts.cmdExit.Text = "ESC Don't Save"
    frmVATaxMsgW3Opts.cmdOption.Text = "F5 Review"
    frmVATaxMsgW3Opts.Show vbModal
    choice = frmVATaxMsgW3Opts.fptxtChoice.Text
    Unload frmVATaxMsgW3Opts
    If choice = "continue" Then
      DontExit = False
      Call cmdSave_Click
      Exit Function
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fpLongTaxYear
  ThisLong = fpLongTaxYear.Value
  ThatLong = TempTaxYear
  If ThisLong <> ThatLong Then
    frmVATaxMsgW3Opts.Label1.Caption = "Changes have been made. Do you wish to save these changes. Press F10 to save, press F5 to review or press ESC to abandon all changes."
    frmVATaxMsgW3Opts.Label1.Top = 800
    frmVATaxMsgW3Opts.cmdCont.Text = "F10 Save"
    frmVATaxMsgW3Opts.cmdExit.Text = "ESC Don't Save"
    frmVATaxMsgW3Opts.cmdOption.Text = "F5 Review"
    frmVATaxMsgW3Opts.Show vbModal
    choice = frmVATaxMsgW3Opts.fptxtChoice.Text
    Unload frmVATaxMsgW3Opts
    If choice = "continue" Then
      DontExit = False
      Call cmdSave_Click
      Exit Function
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fpCurrPers
  ThisDbl = fpCurrPers.Value
  ThatDbl = TempPers
  If ThisDbl <> ThatDbl Then
    frmVATaxMsgW3Opts.Label1.Caption = "Changes have been made. Do you wish to save these changes. Press F10 to save, press F5 to review or press ESC to abandon all changes."
    frmVATaxMsgW3Opts.Label1.Top = 800
    frmVATaxMsgW3Opts.cmdCont.Text = "F10 Save"
    frmVATaxMsgW3Opts.cmdExit.Text = "ESC Don't Save"
    frmVATaxMsgW3Opts.cmdOption.Text = "F5 Review"
    frmVATaxMsgW3Opts.Show vbModal
    choice = frmVATaxMsgW3Opts.fptxtChoice.Text
    Unload frmVATaxMsgW3Opts
    If choice = "continue" Then
      DontExit = False
      Call cmdSave_Click
      Exit Function
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fpCurrInt
  ThisDbl = fpCurrInt.Value
  ThatDbl = TempIntAmount
  If ThisDbl <> ThatDbl Then
    frmVATaxMsgW3Opts.Label1.Caption = "Changes have been made. Do you wish to save these changes. Press F10 to save, press F5 to review or press ESC to abandon all changes."
    frmVATaxMsgW3Opts.Label1.Top = 800
    frmVATaxMsgW3Opts.cmdCont.Text = "F10 Save"
    frmVATaxMsgW3Opts.cmdExit.Text = "ESC Don't Save"
    frmVATaxMsgW3Opts.cmdOption.Text = "F5 Review"
    frmVATaxMsgW3Opts.Show vbModal
    choice = frmVATaxMsgW3Opts.fptxtChoice.Text
    Unload frmVATaxMsgW3Opts
    If choice = "continue" Then
      DontExit = False
      Call cmdSave_Click
      Exit Function
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fpCurrMT
  ThisDbl = fpCurrMT.Value
  ThatDbl = TempMachTools
  If ThisDbl <> ThatDbl Then
    frmVATaxMsgW3Opts.Label1.Caption = "Changes have been made. Do you wish to save these changes. Press F10 to save, press F5 to review or press ESC to abandon all changes."
    frmVATaxMsgW3Opts.Label1.Top = 800
    frmVATaxMsgW3Opts.cmdCont.Text = "F10 Save"
    frmVATaxMsgW3Opts.cmdExit.Text = "ESC Don't Save"
    frmVATaxMsgW3Opts.cmdOption.Text = "F5 Review"
    frmVATaxMsgW3Opts.Show vbModal
    choice = frmVATaxMsgW3Opts.fptxtChoice.Text
    Unload frmVATaxMsgW3Opts
    If choice = "continue" Then
      DontExit = False
      Call cmdSave_Click
      Exit Function
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fpCurrMC
  ThisDbl = fpCurrMC.Value
  ThatDbl = TempMerchCap
  If ThisDbl <> ThatDbl Then
    frmVATaxMsgW3Opts.Label1.Caption = "Changes have been made. Do you wish to save these changes. Press F10 to save, press F5 to review or press ESC to abandon all changes."
    frmVATaxMsgW3Opts.Label1.Top = 800
    frmVATaxMsgW3Opts.cmdCont.Text = "F10 Save"
    frmVATaxMsgW3Opts.cmdExit.Text = "ESC Don't Save"
    frmVATaxMsgW3Opts.cmdOption.Text = "F5 Review"
    frmVATaxMsgW3Opts.Show vbModal
    choice = frmVATaxMsgW3Opts.fptxtChoice.Text
    Unload frmVATaxMsgW3Opts
    If choice = "continue" Then
      DontExit = False
      Call cmdSave_Click
      Exit Function
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fpCurrFE
  ThisDbl = fpCurrFE.Value
  ThatDbl = TempFarmEquip
  If ThisDbl <> ThatDbl Then
    frmVATaxMsgW3Opts.Label1.Caption = "Changes have been made. Do you wish to save these changes. Press F10 to save, press F5 to review or press ESC to abandon all changes."
    frmVATaxMsgW3Opts.Label1.Top = 800
    frmVATaxMsgW3Opts.cmdCont.Text = "F10 Save"
    frmVATaxMsgW3Opts.cmdExit.Text = "ESC Don't Save"
    frmVATaxMsgW3Opts.cmdOption.Text = "F5 Review"
    frmVATaxMsgW3Opts.Show vbModal
    choice = frmVATaxMsgW3Opts.fptxtChoice.Text
    Unload frmVATaxMsgW3Opts
    If choice = "continue" Then
      DontExit = False
      Call cmdSave_Click
      Exit Function
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fpCurrMH
  ThisDbl = fpCurrMH.Value
  ThatDbl = TempMobHomes
  If ThisDbl <> ThatDbl Then
    frmVATaxMsgW3Opts.Label1.Caption = "Changes have been made. Do you wish to save these changes. Press F10 to save, press F5 to review or press ESC to abandon all changes."
    frmVATaxMsgW3Opts.Label1.Top = 800
    frmVATaxMsgW3Opts.cmdCont.Text = "F10 Save"
    frmVATaxMsgW3Opts.cmdExit.Text = "ESC Don't Save"
    frmVATaxMsgW3Opts.cmdOption.Text = "F5 Review"
    frmVATaxMsgW3Opts.Show vbModal
    choice = frmVATaxMsgW3Opts.fptxtChoice.Text
    Unload frmVATaxMsgW3Opts
    If choice = "continue" Then
      DontExit = False
      Call cmdSave_Click
      Exit Function
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fpCurrPen
  ThisDbl = fpCurrPen.Value
  ThatDbl = TempPenalty
  If ThisDbl <> ThatDbl Then
    frmVATaxMsgW3Opts.Label1.Caption = "Changes have been made. Do you wish to save these changes. Press F10 to save, press F5 to review or press ESC to abandon all changes."
    frmVATaxMsgW3Opts.Label1.Top = 800
    frmVATaxMsgW3Opts.cmdCont.Text = "F10 Save"
    frmVATaxMsgW3Opts.cmdExit.Text = "ESC Don't Save"
    frmVATaxMsgW3Opts.cmdOption.Text = "F5 Review"
    frmVATaxMsgW3Opts.Show vbModal
    choice = frmVATaxMsgW3Opts.fptxtChoice.Text
    Unload frmVATaxMsgW3Opts
    If choice = "continue" Then
      DontExit = False
      Call cmdSave_Click
      Exit Function
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fpCurrOpt1
  ThisDbl = fpCurrOpt1.Value
  ThatDbl = TempOptRev1
  If ThisDbl <> ThatDbl Then
    frmVATaxMsgW3Opts.Label1.Caption = "Changes have been made. Do you wish to save these changes. Press F10 to save, press F5 to review or press ESC to abandon all changes."
    frmVATaxMsgW3Opts.Label1.Top = 800
    frmVATaxMsgW3Opts.cmdCont.Text = "F10 Save"
    frmVATaxMsgW3Opts.cmdExit.Text = "ESC Don't Save"
    frmVATaxMsgW3Opts.cmdOption.Text = "F5 Review"
    frmVATaxMsgW3Opts.Show vbModal
    choice = frmVATaxMsgW3Opts.fptxtChoice.Text
    Unload frmVATaxMsgW3Opts
    If choice = "continue" Then
      DontExit = False
      Call cmdSave_Click
      Exit Function
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fpCurrOpt2
  ThisDbl = fpCurrOpt2.Value
  ThatDbl = TempOptRev2
  If ThisDbl <> ThatDbl Then
    frmVATaxMsgW3Opts.Label1.Caption = "Changes have been made. Do you wish to save these changes. Press F10 to save, press F5 to review or press ESC to abandon all changes."
    frmVATaxMsgW3Opts.Label1.Top = 800
    frmVATaxMsgW3Opts.cmdCont.Text = "F10 Save"
    frmVATaxMsgW3Opts.cmdExit.Text = "ESC Don't Save"
    frmVATaxMsgW3Opts.cmdOption.Text = "F5 Review"
    frmVATaxMsgW3Opts.Show vbModal
    choice = frmVATaxMsgW3Opts.fptxtChoice.Text
    Unload frmVATaxMsgW3Opts
    If choice = "continue" Then
      DontExit = False
      Call cmdSave_Click
      Exit Function
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fpCurrOpt3
  ThisDbl = fpCurrOpt3.Value
  ThatDbl = TempOptRev3
  If ThisDbl <> ThatDbl Then
    frmVATaxMsgW3Opts.Label1.Caption = "Changes have been made. Do you wish to save these changes. Press F10 to save, press F5 to review or press ESC to abandon all changes."
    frmVATaxMsgW3Opts.Label1.Top = 800
    frmVATaxMsgW3Opts.cmdCont.Text = "F10 Save"
    frmVATaxMsgW3Opts.cmdExit.Text = "ESC Don't Save"
    frmVATaxMsgW3Opts.cmdOption.Text = "F5 Review"
    frmVATaxMsgW3Opts.Show vbModal
    choice = frmVATaxMsgW3Opts.fptxtChoice.Text
    Unload frmVATaxMsgW3Opts
    If choice = "continue" Then
      DontExit = False
      Call cmdSave_Click
      Exit Function
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fpcmbBillType
  ThisDesc = Mid(fpcmbBillType.Text, 1, 1)
  ThatDesc = TempBillType
  If ThisDesc <> ThatDesc Then
    frmVATaxMsgW3Opts.Label1.Caption = "Changes have been made. Do you wish to save these changes. Press F10 to save, press F5 to review or press ESC to abandon all changes."
    frmVATaxMsgW3Opts.Label1.Top = 800
    frmVATaxMsgW3Opts.cmdCont.Text = "F10 Save"
    frmVATaxMsgW3Opts.cmdExit.Text = "ESC Don't Save"
    frmVATaxMsgW3Opts.cmdOption.Text = "F5 Review"
    frmVATaxMsgW3Opts.Show vbModal
    choice = frmVATaxMsgW3Opts.fptxtChoice.Text
    Unload frmVATaxMsgW3Opts
    If choice = "continue" Then
      DontExit = False
      Call cmdSave_Click
      Exit Function
    Else
      GoSub HandleChoice
    End If
  End If
  
  Exit Function
  
HandleChoice:
  Select Case choice
    Case "abort"
      If LookUpMode = False Then
        frmVATaxManualBillMenu.Show
        DoEvents
        KillFile "C:\CPWork\pmanualbill.dat"
        Unload Me
      End If
      Exit Function
    Case "option"
      fptxtAcctNum = TempAcctNum
      If ThisControl.Enabled = True Then
        ThisControl.SetFocus
      End If
      Check4Changes = True
      Exit Function
    Case Else
  End Select
      
  Return
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPManualBillEntry", "Check4Changes", Erl)
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
  
Public Sub ClearBillFields()
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
'  fptxtAcctNum.Text = "0"
  fpCurrPers.Value = 0
  fpCurrInt.Value = 0
  fpCurrMT.Value = 0
  fpCurrMC.Value = 0
  fpCurrFE.Value = 0
  fpCurrMH.Value = 0
  fpCurrPen.Value = 0
  fpCurrOpt1.Value = 0
  fpCurrOpt2.Value = 0
  fpCurrOpt3.Value = 0
  fpDblSngBillNum.Value = 0
  fpcmbBillType.Text = "PERSONAL ONLY"
  fptxtBillDate = Date
  fpLongTaxYear = TaxMasterRec.PTaxYear
  fptxtActiveProp.Text = "NOTHING SELECTED"
End Sub

Public Sub MsgAlertTimer_Timer()
  Static tog As Double
  Static TogState As Boolean
  If Me.Visible Then
    If BtnFnt# = 0 Then
      BtnFnt# = cmdMessage.FontSize
    End If
    If TogState Then
      tog = tog + 1
    Else
      tog = tog - 1
    End If
    Select Case tog
    Case 1
      cmdMessage.ForeColor = &H80000012
      cmdMessage.FontSize = BtnFnt
    Case 2
      cmdMessage.ForeColor = &H80000011
      cmdMessage.FontSize = BtnFnt - 0.7
    Case 3
      cmdMessage.ForeColor = &H80000011
      cmdMessage.FontSize = BtnFnt - 1.4
    Case 4
      cmdMessage.ForeColor = &H80000010
      cmdMessage.FontSize = BtnFnt - 2.1
    Case 5
      cmdMessage.ForeColor = &H80000010
      cmdMessage.FontSize = BtnFnt - 2.8
    Case 6
      cmdMessage.ForeColor = &H8000000F
      cmdMessage.FontSize = BtnFnt - 3.5
    Case 7
      cmdMessage.ForeColor = &H8000000F
      cmdMessage.FontSize = BtnFnt - 4.2
    Case 8
      cmdMessage.ForeColor = &H8000000E
      cmdMessage.FontSize = BtnFnt - 4.9
    Case 9
      cmdMessage.ForeColor = &H8000000E
      cmdMessage.FontSize = BtnFnt - 5.6
    End Select
    Select Case tog
    Case Is < 0, Is > 9
      TogState = Not TogState
    End Select
  End If

End Sub

Private Function Look4TempCreditUsed() As Boolean
  Dim TaxMRec As TaxMTransactionType
  Dim TMHandle As Integer
  Dim NumOfTMRecs As Integer
  Dim ThisRec As Integer
  Dim x As Integer
  
  Look4TempCreditUsed = False
  OpenTaxManualBillFile TMHandle, NumOfTMRecs
  For x = 1 To NumOfTMRecs
    Get TMHandle, x, TaxMRec
    If TaxMRec.Deleted = True Then GoTo SkipIt
    If TaxMRec.Account = GCustNum Then
      If TaxMRec.OverPayUsed <> 0 Then
        Look4TempCreditUsed = True
        Exit For
      End If
    End If
SkipIt:
  Next x
  Close TMHandle
End Function

