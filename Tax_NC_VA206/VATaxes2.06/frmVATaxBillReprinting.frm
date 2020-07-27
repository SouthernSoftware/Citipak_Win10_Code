VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmVATaxBillReprinting 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Bill Reprinting"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmVATaxBillReprinting.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcmbType 
      Height          =   405
      Left            =   5370
      TabIndex        =   0
      Top             =   960
      Width           =   2700
      _Version        =   196608
      _ExtentX        =   4762
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
      ColDesigner     =   "frmVATaxBillReprinting.frx":08CA
   End
   Begin EditLib.fpDoubleSingle fpDblSnglPersLastBill 
      Height          =   372
      Left            =   9000
      TabIndex        =   2
      Top             =   2160
      Width           =   1572
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
      DecimalPlaces   =   -1
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
   Begin EditLib.fpText fptxtRealOrder 
      Height          =   372
      Left            =   1200
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   6480
      Width           =   2652
      _Version        =   196608
      _ExtentX        =   4683
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
   Begin EditLib.fpDoubleSingle fpDblSnglRealStartBill 
      Height          =   372
      Left            =   2640
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   4188
      Width           =   1572
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
      NoSpecialKeys   =   2
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
      ControlType     =   1
      Text            =   "0"
      DecimalPlaces   =   -1
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
   Begin EditLib.fpLongInteger fpLongRealTaxYear 
      Height          =   372
      Left            =   2880
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3708
      Width           =   1092
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
      NoSpecialKeys   =   2
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
      ControlType     =   1
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
   Begin EditLib.fpDoubleSingle fpDblSnglRealRate 
      Height          =   372
      Left            =   2640
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5028
      Width           =   1332
      _Version        =   196608
      _ExtentX        =   2355
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
      NoSpecialKeys   =   2
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
      ControlType     =   1
      Text            =   "0.0000"
      DecimalPlaces   =   4
      DecimalPoint    =   "."
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "-9000000000"
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
   Begin EditLib.fpDoubleSingle fpDblSnglRealLateList 
      Height          =   372
      Left            =   2640
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   5520
      Width           =   1332
      _Version        =   196608
      _ExtentX        =   2350
      _ExtentY        =   656
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
      NoSpecialKeys   =   2
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
      ControlType     =   1
      Text            =   "0.0000"
      DecimalPlaces   =   4
      DecimalPoint    =   "."
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "-9000000000"
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
   Begin EditLib.fpDoubleSingle fpDblSnglRealLastBill 
      Height          =   372
      Left            =   2280
      TabIndex        =   4
      Top             =   2640
      Width           =   1572
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
      DecimalPlaces   =   -1
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
   Begin EditLib.fpDoubleSingle fpDblSnglRealFirstBill 
      Height          =   372
      Left            =   2280
      TabIndex        =   3
      Top             =   2160
      Width           =   1572
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
      DecimalPlaces   =   -1
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
   Begin fpBtnAtlLibCtl.fpBtn cmdRList 
      Height          =   372
      Left            =   1320
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   3240
      Width           =   2532
      _Version        =   131072
      _ExtentX        =   4466
      _ExtentY        =   656
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
      ButtonDesigner  =   "frmVATaxBillReprinting.frx":0BC1
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
      Height          =   495
      Left            =   7185
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   8040
      Width           =   2055
      _Version        =   131072
      _ExtentX        =   3625
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
      ButtonDesigner  =   "frmVATaxBillReprinting.frx":0DAB
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   495
      Left            =   2400
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   8040
      Width           =   2055
      _Version        =   131072
      _ExtentX        =   3625
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
      ButtonDesigner  =   "frmVATaxBillReprinting.frx":0F8A
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdAlign 
      Height          =   495
      Left            =   4800
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   8040
      Width           =   2040
      _Version        =   131072
      _ExtentX        =   3598
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
      ButtonDesigner  =   "frmVATaxBillReprinting.frx":1166
   End
   Begin EditLib.fpDateTime fptxtRealDueDate 
      Height          =   372
      Left            =   1680
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   7320
      Width           =   1740
      _Version        =   196608
      _ExtentX        =   3069
      _ExtentY        =   656
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
      AlignTextH      =   1
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   2
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
      ControlType     =   1
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
   Begin EditLib.fpDoubleSingle fpDblSnglMHRate 
      Height          =   372
      Left            =   9336
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   4536
      Width           =   1332
      _Version        =   196608
      _ExtentX        =   2350
      _ExtentY        =   656
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
      NoSpecialKeys   =   2
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
      ControlType     =   1
      Text            =   "0.0000"
      DecimalPlaces   =   4
      DecimalPoint    =   "."
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "-9000000000"
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
   Begin EditLib.fpDateTime fptxtPersDueDate 
      Height          =   372
      Left            =   6960
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   7320
      Width           =   1740
      _Version        =   196608
      _ExtentX        =   3069
      _ExtentY        =   656
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
      AlignTextH      =   1
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   2
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
   Begin EditLib.fpDoubleSingle fpDblSnglStartPersBill 
      Height          =   372
      Left            =   8040
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   3708
      Width           =   1212
      _Version        =   196608
      _ExtentX        =   2138
      _ExtentY        =   656
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
      NoSpecialKeys   =   2
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
      DecimalPlaces   =   -1
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
   Begin EditLib.fpDoubleSingle fpDblSnglPersRate 
      Height          =   372
      Left            =   6480
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   4536
      Width           =   1332
      _Version        =   196608
      _ExtentX        =   2350
      _ExtentY        =   656
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
      NoSpecialKeys   =   2
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
      ControlType     =   1
      Text            =   "0.0000"
      DecimalPlaces   =   4
      DecimalPoint    =   "."
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "-9000000000"
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
   Begin EditLib.fpDoubleSingle fpDblSnglPersLateList 
      Height          =   372
      Left            =   9336
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   5448
      Width           =   1332
      _Version        =   196608
      _ExtentX        =   2350
      _ExtentY        =   656
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
      NoSpecialKeys   =   2
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
      ControlType     =   1
      Text            =   "0.0000"
      DecimalPlaces   =   4
      DecimalPoint    =   "."
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "-9000000000"
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
   Begin EditLib.fpLongInteger fpLongPersTaxYear 
      Height          =   372
      Left            =   8040
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   3192
      Width           =   1092
      _Version        =   196608
      _ExtentX        =   1926
      _ExtentY        =   656
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
      NoSpecialKeys   =   2
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
      ControlType     =   1
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
   Begin EditLib.fpText fptxtPersOrder 
      Height          =   372
      Left            =   6360
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   6432
      Width           =   2772
      _Version        =   196608
      _ExtentX        =   4890
      _ExtentY        =   656
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
      NoSpecialKeys   =   2
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
   Begin EditLib.fpDoubleSingle fpDblSnglMTRate 
      Height          =   372
      Left            =   6480
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   4992
      Width           =   1332
      _Version        =   196608
      _ExtentX        =   2350
      _ExtentY        =   656
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
      NoSpecialKeys   =   2
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
      ControlType     =   1
      Text            =   "0.0000"
      DecimalPlaces   =   4
      DecimalPoint    =   "."
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "-9000000000"
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
   Begin EditLib.fpDoubleSingle fpDblSnglFERate 
      Height          =   372
      Left            =   6480
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   5448
      Width           =   1332
      _Version        =   196608
      _ExtentX        =   2350
      _ExtentY        =   656
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
      NoSpecialKeys   =   2
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
      ControlType     =   1
      Text            =   "0.0000"
      DecimalPlaces   =   4
      DecimalPoint    =   "."
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "-9000000000"
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
   Begin EditLib.fpDoubleSingle fpDblSnglMCRate 
      Height          =   372
      Left            =   9336
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   4992
      Width           =   1332
      _Version        =   196608
      _ExtentX        =   2350
      _ExtentY        =   656
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
      NoSpecialKeys   =   2
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
      ControlType     =   1
      Text            =   "0.0000"
      DecimalPlaces   =   4
      DecimalPoint    =   "."
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "-9000000000"
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
   Begin EditLib.fpDoubleSingle fpDblSnglPersFirstBill 
      Height          =   372
      Left            =   6000
      TabIndex        =   1
      Top             =   2160
      Width           =   1572
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
      DecimalPlaces   =   -1
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
   Begin fpBtnAtlLibCtl.fpBtn cmdPList 
      Height          =   372
      Left            =   6480
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   2760
      Width           =   2532
      _Version        =   131072
      _ExtentX        =   4466
      _ExtentY        =   656
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
      ButtonDesigner  =   "frmVATaxBillReprinting.frx":1347
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   456
      Left            =   4560
      Top             =   1452
      Width           =   6468
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   456
      Left            =   648
      Top             =   1452
      Width           =   3852
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Select Bill Type:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   372
      Left            =   3336
      TabIndex        =   48
      Top             =   1044
      Width           =   1932
   End
   Begin VB.Shape Shape10 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   612
      Left            =   4800
      Top             =   2040
      Width           =   6012
   End
   Begin VB.Label Label27 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Last Bill:"
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
      Height          =   372
      Left            =   7920
      TabIndex        =   51
      Top             =   2232
      Width           =   972
   End
   Begin VB.Label Label26 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "First Bill:"
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
      Height          =   372
      Left            =   4920
      TabIndex        =   50
      Top             =   2232
      Width           =   972
   End
   Begin VB.Label Label25 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Pers Prop:"
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
      Height          =   372
      Left            =   5136
      TabIndex        =   49
      Top             =   4632
      Width           =   1212
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Starting Bill No:"
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
      Height          =   372
      Left            =   6120
      TabIndex        =   47
      Top             =   3780
      Width           =   1812
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Late List:"
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
      Height          =   252
      Left            =   8136
      TabIndex        =   46
      Top             =   5508
      Width           =   1092
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Pers Tax Year:"
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
      Height          =   252
      Left            =   6240
      TabIndex        =   45
      Top             =   3276
      Width           =   1692
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Pers Printing Order:"
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
      Height          =   372
      Left            =   6576
      TabIndex        =   44
      Top             =   6072
      Width           =   2292
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pers Due Date:"
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
      Height          =   252
      Left            =   6960
      TabIndex        =   43
      Top             =   6960
      Width           =   1812
   End
   Begin VB.Shape Shape9 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   1776
      Left            =   4800
      Top             =   4212
      Width           =   5988
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Rate"
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
      Height          =   252
      Left            =   6600
      TabIndex        =   42
      Top             =   4272
      Width           =   1092
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   6000
      Left            =   4560
      Top             =   1920
      Width           =   6468
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PERSONAL PROPERTY"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   372
      Left            =   6240
      TabIndex        =   41
      Top             =   1524
      Width           =   3012
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Mach/Tools:"
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
      Height          =   372
      Left            =   4896
      TabIndex        =   40
      Top             =   5076
      Width           =   1452
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Farm Equip:"
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
      Height          =   372
      Left            =   4896
      TabIndex        =   39
      Top             =   5508
      Width           =   1452
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Mbl Homes:"
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
      Height          =   372
      Left            =   7896
      TabIndex        =   38
      Top             =   4632
      Width           =   1332
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Rate"
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
      Height          =   252
      Left            =   9456
      TabIndex        =   37
      Top             =   4272
      Width           =   1092
   End
   Begin VB.Label Label24 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Merch Cap:"
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
      Height          =   372
      Left            =   7896
      TabIndex        =   36
      Top             =   5076
      Width           =   1332
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "REAL PROPERTY"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   372
      Left            =   1080
      TabIndex        =   25
      Top             =   1524
      Width           =   3012
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Real Due Date:"
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
      Height          =   252
      Left            =   1560
      TabIndex        =   24
      Top             =   6960
      Width           =   1812
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   1092
      Left            =   1080
      Top             =   2040
      Width           =   3012
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "First Bill:"
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
      Height          =   372
      Left            =   1200
      TabIndex        =   18
      Top             =   2232
      Width           =   972
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Last Bill:"
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
      Height          =   372
      Left            =   1200
      TabIndex        =   17
      Top             =   2712
      Width           =   972
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Rate"
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
      Height          =   252
      Left            =   2760
      TabIndex        =   16
      Top             =   4740
      Width           =   1092
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Late List:"
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
      Height          =   252
      Left            =   960
      TabIndex        =   15
      Top             =   5616
      Width           =   1452
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Real Estate:"
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
      Height          =   252
      Left            =   1080
      TabIndex        =   14
      Top             =   5088
      Width           =   1332
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Starting Bill No:"
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
      Height          =   372
      Left            =   840
      TabIndex        =   13
      Top             =   4260
      Width           =   1692
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Real Tax Year:"
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
      Height          =   252
      Left            =   1080
      TabIndex        =   12
      Top             =   3780
      Width           =   1692
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   660
      Index           =   1
      Left            =   1488
      Top             =   204
      Width           =   8652
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Bill Reprinting"
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
      Left            =   3108
      TabIndex        =   11
      Top             =   372
      Width           =   5292
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   6012
      Left            =   648
      Top             =   1896
      Width           =   3852
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   1320
      Left            =   828
      Top             =   4668
      Width           =   3492
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Real Printing Order:"
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
      Height          =   372
      Left            =   1440
      TabIndex        =   10
      Top             =   6120
      Width           =   2292
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   780
      Left            =   1488
      Top             =   96
      Width           =   8652
   End
End
Attribute VB_Name = "frmVATaxBillReprinting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class
  Dim BillFormat$
  Dim RealFirstNum As Long
  Dim RealSecondNum As Long
  Dim RealBillCnt As Long
  Dim PersFirstNum As Long
  Dim PersSecondNum As Long
  Dim PersBillCnt As Long
  Dim PersYear As Integer
  Dim MTTaxRate#, FETaxRate#, MHTaxRate#
  Dim MCTaxRate#, PersTaxRate#, PLateRate#
  Dim RLateRate#, RealRate#
  Dim GPPTRADisc#
  Dim RealYear As Integer
  Dim NoAlign As Boolean
  Dim TownName$, Add1$, Add2$, Add3$
  Dim GMaxVehTaxVal#, GMultiYear As Integer
  Public Real As Boolean, GMinVehTaxVal#

Private Sub cmdAlign_Click()
  Dim Handle As Integer
  Dim TempHandle As Integer
  Dim cnt As Integer
  Dim TextLine$
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  
'  On Error GoTo ERRORSTUFF
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close
  
  If fpcmbType.Text = "REAL" Then
    Select Case TaxMasterRec.TaxForm
      Case 30000
        Call TaxMsg(900, "Pitch 10 is recommended for this bill.")
        If Exist("VASTANDRMSK.TXT") Then
          alnRpt = "VASTANDRMSK.TXT"
        Else
          Call TaxMsg(900, "Please install the file 'VASTANDRMSK.TXT'.")
          Exit Sub
        End If
      Case 20003
        Call TaxMsg(900, "Pitch 10 is recommended for this bill.")
        If Exist("MdltwnRMask.TXT") Then
          alnRpt = "MdltwnRMask.TXT"
        Else
          Call TaxMsg(900, "Please install the file 'MdltwnRMask.TXT'.")
          Exit Sub
        End If
      Case 20004
        Call TaxMsg(900, "Pitch 10 is recommended for this bill.")
        If Exist("CdrBluffRMask.TXT") Then
          alnRpt = "CdrBluffRMask.TXT"
        Else
          Call TaxMsg(900, "Please install the file 'CdrBluffRMask.TXT'.")
          Exit Sub
        End If
      Case Else
        Call TaxMsg(900, "No mask is available.")
        Exit Sub
    End Select
  End If
    
  If fpcmbType.Text = "PERSONAL" Then
    Select Case TaxMasterRec.TaxForm
      Case 30000
        Call TaxMsg(900, "Pitch 10 is recommended for this bill.")
        If Exist("VASTANDPMSK.TXT") Then
          alnRpt = "VASTANDPMSK.TXT"
        Else
          Call TaxMsg(900, "Please install the file 'VASTANDPMSK.TXT'.")
          Exit Sub
        End If
      Case 20003
        Call TaxMsg(900, "Pitch 10 is recommended for this bill.")
        If Exist("MdltwnPMask.TXT") Then
          alnRpt = "MdltwnPMask.TXT"
        Else
          Call TaxMsg(900, "Please install the file 'MdltwnPMask.TXT'.")
          Exit Sub
        End If
      Case 20004
        Call TaxMsg(900, "Pitch 10 is recommended for this bill.")
        If Exist("CdrBluffPMask.TXT") Then
          alnRpt = "CdrBluffPMask.TXT"
        Else
          Call TaxMsg(900, "Please install the file 'CdrBluffPMask.TXT'.")
          Exit Sub
        End If
      Case Else
        Call TaxMsg(900, "No mask is available.")
        Exit Sub
    End Select
  End If
  
'  If fpcmbType.Text = "REAL" Then
'    alnRpt = "TAXREMSK.DAT"
'  Else
'    alnRpt = "TAXPPMSK.DAT"
'  End If
  Handle = FreeFile
  Open alnRpt For Input As #Handle
  TempHandle = FreeFile
  Open "TAXALIGN.MSK" For Output As #TempHandle
  Do While Not eof(Handle)
    Line Input #Handle, TextLine   ' Read line into variable.
    Print #TempHandle, TextLine
  Loop
  Close
  alnRpt = "TAXALIGN.MSK"
  doAlign = True
  frmVATaxPrint.Show vbModal
  alnRpt = ""
  doAlign = False
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxBillReprinting", "cmdAlign_Click", Erl)
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
 
End Sub

Private Sub cmdExit_Click()
  frmVATaxBillPrintMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdRList_Click()
  Real = True
  frmVATaxPrintedBillsList.Show vbModal
  DoEvents
End Sub

Private Sub cmdProcess_Click()
  Dim PTaxBill As VAPPTaxBillType
  Dim RTaxBill As VARETaxBillType
  Dim RptHandle As Integer
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim x As Long, BillNo&
  Dim WhatRec&, PrnCnt&
  Dim PINTemp As String * 18
  Dim CustName$
  Dim RptFile$, FBill&, LBill&
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim PersRec As PersonalRecType
  Dim PHandle As Integer
  Dim NumOfPRecs As Long
  Dim RZipRec As BillPrintRZipIdxType
  Dim PZipRec As BillPrintPZipIdxType
  Dim ZHandle As Integer
  Dim NumOfZRecs As Long
  Dim MortRec As BillPrintMortIdxType
  Dim MRHandle As Integer
  Dim NumOfMRRecs As Long
  Dim PrintIt As Boolean
  
  On Error GoTo ERRORSTUFF
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  If TaxMasterRec.TaxForm = 16716 Or TaxMasterRec.TaxForm = 20002 Then
    If fpcmbType.Text = "REAL" Then
      FBill = fpDblSnglRealFirstBill.Value
      If FBill < RealFirstNum Then
        Call TaxMsg(900, "The first bill cannot be less than " + CStr(RealFirstNum) + ". Please re-enter and try again.")
        fpDblSnglRealFirstBill.SetFocus
        Exit Sub
      End If
  
      LBill = fpDblSnglRealLastBill.Value
      If LBill > RealSecondNum Then
        Call TaxMsg(900, "The last bill cannot be greater than " + CStr(RealSecondNum) + ". Please re-enter and try again.")
        fpDblSnglRealLastBill.SetFocus
        Exit Sub
      End If
    
      If FBill > LBill Then
        Call TaxMsg(900, "The first bill number must be smaller than the last.")
        fpDblSnglRealFirstBill.SetFocus
        Exit Sub
      End If
      Call PrintLaserReal1
      Exit Sub
    ElseIf fpcmbType.Text = "PERSONAL" Then
      FBill = fpDblSnglPersFirstBill.Value
      If FBill < PersFirstNum Then
        Call TaxMsg(900, "The first bill cannot be less than " + CStr(PersFirstNum) + ". Please re-enter and try again.")
        fpDblSnglPersFirstBill.SetFocus
        Exit Sub
      End If
  
      LBill = fpDblSnglPersLastBill.Value
      If LBill > PersSecondNum Then
        Call TaxMsg(900, "The last bill cannot be greater than " + CStr(PersSecondNum) + ". Please re-enter and try again.")
        fpDblSnglPersLastBill.SetFocus
        Exit Sub
      End If
  
      If FBill > LBill Then
        Call TaxMsg(900, "The first bill number must be smaller than the last.")
        fpDblSnglPersFirstBill.SetFocus
        Exit Sub
      End If
      
      If TaxMasterRec.TaxForm = 16716 Then
        Call PrintLaserPers1
        Exit Sub
      Else
        Call PrintPersLaserItemized
        Exit Sub
      End If
    End If
  End If

  If fpcmbType.Text = "REAL" Then
    FBill = fpDblSnglRealFirstBill.Value
    If FBill < RealFirstNum Then
      Call TaxMsg(900, "The first bill cannot be less than " + CStr(RealFirstNum) + ". Please re-enter and try again.")
      fpDblSnglRealFirstBill.SetFocus
      Exit Sub
    End If
  
    LBill = fpDblSnglRealLastBill.Value
    If LBill > RealSecondNum Then
      Call TaxMsg(900, "The last bill cannot be greater than " + CStr(RealSecondNum) + ". Please re-enter and try again.")
      fpDblSnglRealLastBill.SetFocus
      Exit Sub
    End If
    
    If FBill > LBill Then
      Call TaxMsg(900, "The first bill number must be smaller than the last.")
      fpDblSnglRealFirstBill.SetFocus
      Exit Sub
    End If
    
    If TaxMasterRec.TaxForm = 20003 Then
      Call PrintMdltwnReal
      Exit Sub
    End If
    
    If TaxMasterRec.TaxForm = 20004 Then
      Call TaxMsg(900, "Pitch 12 is recommended for this bill.")
      Call PrintCdrBluffReal
      Exit Sub
    End If
    
    RptHandle = FreeFile
    RptFile$ = "TAXREALBILRE.PRN"
    Open RptFile For Output As RptHandle
  
    OpenRealTaxBillFile TBHandle, NumOfTBRecs
    OpenTaxCustFile TCHandle, NumOfTCRecs
    Call TaxMsg(900, "Pitch 10 is recommended for this bill.")

    PrnCnt = fpDblSnglRealFirstBill
    
    NumOfMRRecs = 0
    NumOfZRecs = 0
    If Exist("MORTIDX.DAT") Then '12/6/06
      OpenMortIdxFile MRHandle, NumOfMRRecs
      NumOfTBRecs = NumOfMRRecs
    ElseIf Exist("RZIPIDX.DAT") Then '12/6/06
      OpenRZipIdxFile ZHandle, NumOfZRecs
      NumOfTBRecs = NumOfZRecs
    End If
    
    For x = 1 To NumOfTBRecs
      If NumOfMRRecs > 0 Then '12/6/06
        Get MRHandle, x, MortRec
        WhatRec& = MortRec.TaxBillRec
      ElseIf NumOfZRecs > 0 Then '12/6/06
        Get ZHandle, x, RZipRec
        WhatRec& = RZipRec.TaxBillRec
      Else
        WhatRec& = x '12/6/06
      End If
      Get TBHandle, WhatRec&, RTaxBill
      If RTaxBill.BillPrinted = False Then GoTo NextReal
      Get TCHandle, RTaxBill.CustRec, TaxCust
      If RTaxBill.BillNumber >= FBill And RTaxBill.BillNumber <= LBill Then
        If InStr(TaxMasterRec.Name, "HALIFAX") Then
          Call PrintHalifaxStandardReal(RptHandle, TBHandle, RTaxBill, PrnCnt) 'TCHandle, TaxCust, PrnCnt, RealRec, RHandle)
        Else
          Call PrintRealVAStandard(RptHandle, TBHandle, RTaxBill, TCHandle, TaxCust, PrnCnt)
        End If
        PrnCnt = PrnCnt + 1
      End If
NextReal:
    Next x
    Close
    ViewPrint RptFile$, "Real Tax Bill Reprinting", True
  ElseIf fpcmbType.Text = "PERSONAL" Then
    OpenPersPropFile PHandle, NumOfPRecs
    OpenTaxCustFile TCHandle, NumOfTCRecs
    FBill = fpDblSnglPersFirstBill.Value
    If FBill < PersFirstNum Then
      Call TaxMsg(900, "The first bill cannot be less than " + CStr(PersFirstNum) + ". Please re-enter and try again.")
      fpDblSnglPersFirstBill.SetFocus
      Exit Sub
    End If
  
    LBill = fpDblSnglPersLastBill.Value
    If LBill > PersSecondNum Then
      Call TaxMsg(900, "The last bill cannot be greater than " + CStr(PersSecondNum) + ". Please re-enter and try again.")
      fpDblSnglPersLastBill.SetFocus
      Exit Sub
    End If
  
    If FBill > LBill Then
      Call TaxMsg(900, "The first bill number must be smaller than the last.")
      fpDblSnglPersFirstBill.SetFocus
      Exit Sub
    End If
    
    If TaxMasterRec.TaxForm = 20003 Then
      Call PrintMdltwnPers
      Exit Sub
    End If
    
    If TaxMasterRec.TaxForm = 20004 Then
      Call TaxMsg(900, "Pitch 10 is recommended for this bill.")
      Call PrintCdrBluffPers
      Exit Sub
    End If
    
    RptHandle = FreeFile
    RptFile$ = "TAXPERSBILRE.PRN"
  
    Open RptFile For Output As RptHandle
  
    OpenPersTaxBillFile TBHandle, NumOfTBRecs
    PrnCnt = fpDblSnglPersFirstBill
    Call TaxMsg(900, "Pitch 10 is recommended for this bill.")
    
    If Exist("PZIPIDX.DAT") Then '12/6/06
      OpenPZipIdxFile ZHandle, NumOfZRecs
      NumOfTBRecs = NumOfZRecs
    End If
    
    For x = 1 To NumOfTBRecs
      If NumOfZRecs > 0 Then '12/6/06
        Get ZHandle, x, PZipRec
        WhatRec& = PZipRec.TaxBillRec
      Else
        WhatRec& = x '12/6/06
      End If
      Get TBHandle, WhatRec&, PTaxBill
      If PTaxBill.BillPrinted = False Then GoTo NextPers
      Get TCHandle, PTaxBill.CustRec, TaxCust
      If PTaxBill.BillNumber >= FBill And PTaxBill.BillNumber <= LBill Then
        If InStr(TaxMasterRec.Name, "HALIFAX") Then
          Call PrintHalifaxStandardPersonal(RptHandle, TBHandle, PTaxBill, TCHandle, TaxCust, PrnCnt, PersRec, PHandle)
        Else
          Call PrintPersVAStandard(RptHandle, TBHandle, PTaxBill, TCHandle, TaxCust, PrnCnt, PersRec, PHandle)
        End If
        PrnCnt = PrnCnt + 1
      End If
NextPers:
    Next x
    Close
    ViewPrint RptFile$, "Personal Tax Bill Reprinting", True
  End If
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxBillReprinting", "cmdProcess_Click", Erl)
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
      MainLog ("CitiTaxes.exe terminated via menu bar on frmVATaxBillReprinting.")
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

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  MainLog ("User opened frmVATaxBillPrinting.")
  Me.HelpContextID = hlpReprintTax
  Call LoadMe
End Sub

Private Sub LoadMe()
  Dim PTaxBill As VAPPTaxBillType
  Dim RTaxBill As VARETaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim RealBillInfo As VARETaxBillInfoType
  Dim BIHandle As Integer
  Dim PersBillInfo As VAPPTaxBillInfoType
  Dim IdxType As Integer
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim MortCodeRec As MortCodeRecType
  Dim MCHandle As Integer
  Dim NumMortCodes As Integer
  Dim x As Integer
  Dim WhatRec As Long
  Dim RZipRec As BillPrintRZipIdxType
  Dim PZipRec As BillPrintPZipIdxType
  Dim ZHandle As Integer
  Dim NumOfZRecs As Long
  Dim MortRec As BillPrintMortIdxType
  Dim MRHandle As Integer
  Dim NumOfMRRecs As Long
  Dim RZipYN As Boolean
  Dim PZipYN As Boolean
  Dim MortYN As Boolean
  
  On Error GoTo ERRORSTUFF
  
  RZipYN = False
  PZipYN = False
  MortYN = False
  doAlign = False
  cmdPList.Enabled = False
  cmdRList.Enabled = False
  If Exist(RealTaxBillInfoFile) Then
    cmdRList.Enabled = True
    OpenRealBillInfoFile BIHandle
    Get BIHandle, 1, RealBillInfo
    Close BIHandle
    OpenRealTaxBillFile TBHandle, NumOfTBRecs
    If Exist("MORTIDX.DAT") Then '12/6/06
      OpenMortIdxFile MRHandle, NumOfMRRecs
      NumOfTBRecs = NumOfMRRecs
      MortYN = True
    ElseIf Exist("RZIPIDX.DAT") Then '12/6/06
      OpenRZipIdxFile ZHandle, NumOfZRecs
      NumOfTBRecs = NumOfZRecs
      RZipYN = True
    End If

    For x = NumOfTBRecs To 1 Step -1
      If MortYN = True Then '12/6/06
        Get MRHandle, x, MortRec
        Get TBHandle, MortRec.TaxBillRec, RTaxBill
        If RTaxBill.BillNumber > 0 Then
          fpDblSnglRealLastBill = RTaxBill.BillNumber
          Exit For
        End If
      ElseIf RZipYN = True Then '12/6/06
        Get ZHandle, x, RZipRec
        Get TBHandle, RZipRec.TaxBillRec, RTaxBill
        If RTaxBill.BillNumber > 0 Then
          fpDblSnglRealLastBill = RTaxBill.BillNumber
          Exit For
        End If
      Else
        Get TBHandle, x, RTaxBill
        If RTaxBill.BillNumber > 0 Then
          fpDblSnglRealLastBill = RTaxBill.BillNumber
          Exit For
        End If
      End If
    Next x
    
    For x = 1 To NumOfTBRecs
      If MortYN = True Then '12/6/06
        Get MRHandle, x, MortRec
        Get TBHandle, MortRec.TaxBillRec, RTaxBill
        If RTaxBill.BillNumber > 0 Then
          If RealFirstNum = 0 Then
            RealFirstNum = RTaxBill.BillNumber
            BillCnt = NumOfMRRecs
            Exit For
          End If
        End If
      ElseIf RZipYN = True Then '12/6/06
        Get ZHandle, x, RZipRec
        Get TBHandle, RZipRec.TaxBillRec, RTaxBill
        If RTaxBill.BillNumber > 0 Then
          If RealFirstNum = 0 Then
            RealFirstNum = RTaxBill.BillNumber
            BillCnt = NumOfZRecs
            Exit For
          End If
        End If
      Else
        Get TBHandle, x, RTaxBill
        If RTaxBill.BillNumber > 0 Then
          If RealFirstNum = 0 Then
            RealFirstNum = RTaxBill.BillNumber
            BillCnt = BillCnt + 1
          End If
        End If
      End If
    Next x
    
    Close TBHandle
    RealFirstNum = RealBillInfo.BillNum
    RealSecondNum = fpDblSnglRealLastBill
    fpDblSnglRealFirstBill = RealFirstNum
    fpDblSnglRealRate = RealBillInfo.RealRate
    RealRate = RealBillInfo.RealRate
    fpDblSnglRealLateList = RealBillInfo.LATEPCT
    RLateRate = RealBillInfo.LATEPCT
    fpLongRealTaxYear = RealBillInfo.TaxYear
    RealYear = RealBillInfo.TaxYear
    fptxtRealDueDate = MakeRegDate(RealBillInfo.DueDate)
    If fptxtRealDueDate.Text = "12/31/1979" Then fptxtRealDueDate.Text = "N/A"
    Select Case QPTrim$(RealBillInfo.PRNORDER)
      Case "1"
        fptxtRealOrder.Text = "Account Number Order"
      Case "2"
        fptxtRealOrder.Text = "Customer Name Order"
      Case "3"
        fptxtRealOrder.Text = "Search Name Order"
      Case "4"
        fptxtRealOrder.Text = "Social Security Order"
      Case Else
        fptxtRealOrder.Text = "Unknown"
     End Select
    If RealBillInfo.BillNum > 0 Then
      fpDblSnglRealStartBill = RealBillInfo.BillNum
    Else
      fpDblSnglRealStartBill = 0
    End If
  Else
    fpDblSnglRealRate = 0
    fpDblSnglRealLateList = 0
    fpLongRealTaxYear = 0
    fptxtRealDueDate = "N/A"
    fpDblSnglRealStartBill = 0
    fptxtRealOrder.Text = "N/A"
  End If
   
  If Exist(PersTaxBillInfoFile) Then
    cmdPList.Enabled = True
    OpenPersBillInfoFile BIHandle
    Get BIHandle, 1, PersBillInfo
    Close BIHandle
    OpenPersTaxBillFile TBHandle, NumOfTBRecs
    
    If Exist("PZIPIDX.DAT") Then '12/6/06
      OpenPZipIdxFile ZHandle, NumOfZRecs
      NumOfTBRecs = NumOfZRecs
      PZipYN = True
    End If
    
    For x = NumOfTBRecs To 1 Step -1
      If PZipYN = True Then '12/6/06
        Get ZHandle, x, PZipRec
        Get TBHandle, PZipRec.TaxBillRec, PTaxBill
        If PTaxBill.BillNumber > 0 Then
          fpDblSnglPersLastBill = PTaxBill.BillNumber
          Exit For
        End If
      Else
        Get TBHandle, x, PTaxBill
        If PTaxBill.BillNumber > 0 Then
          fpDblSnglPersLastBill = PTaxBill.BillNumber
          Exit For
        End If
      End If
    Next x
    PersFirstNum = 0
    For x = 1 To NumOfTBRecs
      If PZipYN = True Then '12/6/06
        Get ZHandle, x, PZipRec
        Get TBHandle, PZipRec.TaxBillRec, PTaxBill
        If PTaxBill.BillNumber > 0 Then
          If PersFirstNum = 0 Then
            PersFirstNum = 1 'PTaxBill.BillNumber
            BillCnt = NumOfZRecs 'BillCnt + 1
            Exit For
          End If
        End If
      Else
        Get TBHandle, x, PTaxBill
        If PTaxBill.BillNumber > 0 Then
          If PersFirstNum = 0 Then
            PersFirstNum = PTaxBill.BillNumber
            BillCnt = BillCnt + 1
          End If
        End If
      End If
    Next x
    Close TBHandle
    fpDblSnglPersFirstBill = PersFirstNum
    PersSecondNum = fpDblSnglPersLastBill.Value
    fpDblSnglPersRate = PersBillInfo.PERSRATE
    PersTaxRate# = PersBillInfo.PERSRATE
    fpDblSnglMCRate = PersBillInfo.MCRate
    MCTaxRate# = PersBillInfo.MCRate
    fpDblSnglFERate = PersBillInfo.FERate
    FETaxRate# = PersBillInfo.FERate
    fpDblSnglMTRate = PersBillInfo.MTRate
    MTTaxRate# = PersBillInfo.MTRate
    fpDblSnglMHRate = PersBillInfo.MHRate
    MHTaxRate# = PersBillInfo.MHRate
    fpDblSnglPersLateList = PersBillInfo.LATEPCT
    PLateRate# = PersBillInfo.LATEPCT
    fpLongPersTaxYear = PersBillInfo.TaxYear
    PersYear = PersBillInfo.TaxYear
    fptxtPersDueDate = MakeRegDate(PersBillInfo.DueDate)
    Select Case QPTrim$(PersBillInfo.PRNORDER)
      Case "1"
        fptxtPersOrder.Text = "Account Number Order"
      Case "2"
        fptxtPersOrder.Text = "Customer Name Order"
      Case "3"
        fptxtPersOrder.Text = "Search Name Order"
      Case "4"
        fptxtPersOrder.Text = "Social Security Order"
      Case Else
        fptxtPersOrder.Text = "Unknown"
    End Select
    If PersBillInfo.BillNum > 0 Then
      fpDblSnglStartPersBill = PersBillInfo.BillNum
    Else
      fpDblSnglStartPersBill = 0
    End If
  Else
    fpDblSnglPersRate = 0
    fpDblSnglMCRate = 0
    fpDblSnglFERate = 0
    fpDblSnglMTRate = 0
    fpDblSnglMHRate = 0
    fpLongPersTaxYear = 0
    fptxtPersDueDate = "N/A"
    fpDblSnglStartPersBill = 0
    fptxtPersOrder.Text = "N/A"
  End If
  
  If RealBillInfo.RealRate > 0 And PersBillInfo.PERSRATE > 0 Then
    fpcmbType.AddItem "REAL"
    fpcmbType.AddItem "PERSONAL"
  End If
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  TownName$ = QPTrim$(TaxMasterRec.Name)
  Add1$ = QPTrim$(TaxMasterRec.Add1)
  Add2$ = QPTrim$(TaxMasterRec.Add2)
  Add3$ = QPTrim$(TaxMasterRec.City) + ", " + QPTrim$(TaxMasterRec.TownState) + "  " + QPTrim$(TaxMasterRec.Zip)
  GPPTRADisc# = TaxMasterRec.PPTRADisc
  GMaxVehTaxVal = TaxMasterRec.MaxVehTaxVal
  GMinVehTaxVal = TaxMasterRec.MinVehTaxVal
  GMultiYear = TaxMasterRec.MultiYear
  
  BillFormat$ = Left$(TaxMasterRec.TaxForm, 1)
  If TaxMasterRec.TaxForm = 16716 Or TaxMasterRec.TaxForm = 20002 Then
    cmdAlign.Enabled = False
  End If
  fptxtPersDueDate.Text = Date
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxBillPrinting", "LoadMe", Erl)
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
 
End Sub
Private Sub PrintPersVAStandard(RptFile As Integer, TBHandle As Integer, PTaxBill As VAPPTaxBillType, TCHandle As Integer, TaxCust As TaxCustType, PrnCnt As Long, PersRec As PersonalRecType, PHandle As Integer)
  'checked OK against mask (taxppmsk.dat) on 10/21/2005
  
  Dim x As Long, PYearStr$
  Dim File$, LC As Integer, CustName$
  Dim WhatYear As Integer, WhatPers&
  Dim CarCount As Integer
  Dim PPTRAVal#
  Dim NumOfPRecs As Long, PPTRADiscount#
  Dim VehDesc$, PERC!
  Dim TaxAmt#, LCnt As Integer
  Dim PYear As Integer
  Dim GMinVehVal As Double
  Dim TotOth As Double
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim Zip$
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  WhatYear = PersYear
  Zip = InsertZipDash(PTaxBill.CustZip)
'  QPTrim$ (PTaxBill.CustZip)
  If WhatYear = 1999 Then PERC! = 27.5
  If WhatYear = 2000 Then PERC! = 47.5
  If WhatYear >= 2001 Then PERC! = GPPTRADisc
  Print #RptFile, "~"
  Print #RptFile, Tab(63); "TAX YEAR: "; WhatYear
  Print #RptFile, Tab(75); Using$("#####", PTaxBill.BillNumber)
  Print #RptFile, " "
  Print #RptFile, " "
  Print #RptFile, " "
  Print #RptFile, Tab(5); TownName$
  Print #RptFile, Tab(5); Add1$
  Print #RptFile, Tab(5); Add2$
  Print #RptFile, Tab(5); Add3$
  Print #RptFile, " "
  Print #RptFile, " "
  Print #RptFile, " "
'  Print #RptFile, " "
'  Print #RptFile, " " 'added
  Print #RptFile, Tab(5); "Acct # "; Using$("#####0", PTaxBill.CustRec)
  Print #RptFile, Tab(5); QPTrim$(PTaxBill.CustName)
  Print #RptFile, Tab(5); QPTrim$(PTaxBill.CustAdd1)
  Print #RptFile, Tab(5); QPTrim$(PTaxBill.CustAdd2)
  Print #RptFile, Tab(5); QPTrim$(PTaxBill.CustAdd3) + " " + Zip
  For LC = 18 To 21
'  For LC = 19 To 21 'added
   Print #RptFile, " "
  Next LC
  Print #RptFile, Tab(37); "PROP VALUE"; Tab(50); "TAX AMOUNT"; Tab(62); "PPTRA DIS"; Tab(72); "TOTAL DUE"
  Print #RptFile, " "
 'Line 24 Starts Here
  Print #RptFile, "Personal Property"; Tab(32); Using$("#.00", PersTaxRate#);
   Print #RptFile, Tab(37); Using$("#####0.00", PTaxBill.PersValue);
'   If InStr(TaxMasterRec.Name, "CHILHOWIE") Then
'     Print #RptFile, Tab(51); Using$("#####0.00", PTaxBill.PersTaxDue + (TaxMasterRec.MinBill - PTaxBill.PersTaxDue)); ' - TaxBill.OverPayAmt);
'   Else
     Print #RptFile, Tab(51); Using$("#####0.00", PTaxBill.PersTaxDue); ' - PTaxBill.OverPayAmt);
'   End If
   Print #RptFile, Tab(63); Using$("####0.00", PTaxBill.PPTRADiscnt);
'   Print #RptFile, Tab(72); Using$("#####0.00", OldRound(PTaxBill.PersTaxDue - PTaxBill.PPTRADiscnt)) ' - PTaxBill.OverPayAmt))
'   If InStr(TaxMasterRec.Name, "CHILHOWIE") Then '11/06/06
'     Print #RptFile, Tab(72); Using$("#####0.00", OldRound(PTaxBill.PersTaxNet + PTaxBill.ChillHowieFudge)) ' - TaxBill.OverPayAmt))
'   Else
     Print #RptFile, Tab(72); Using$("#####0.00", OldRound(PTaxBill.PersTaxDue - PTaxBill.PPTRADiscnt)) ' - TaxBill.OverPayAmt))
'   End If
   
  Print #RptFile, "Machinery/Tools"; Tab(32); Using$("#.00", MTTaxRate#);
   Print #RptFile, Tab(37); Using$("#####0.00", PTaxBill.MTValue);
   Print #RptFile, Tab(51); Using$("#####0.00", PTaxBill.MTTaxDue);
   Print #RptFile, Tab(72); Using$("#####0.00", PTaxBill.MTTaxDue)
  Print #RptFile, "Farm Equipment";
   Print #RptFile, Tab(32); Using("#.00", FETaxRate#);
   Print #RptFile, Tab(37); Using$("#####0.00", PTaxBill.FEValue);
   Print #RptFile, Tab(51); Using$("#####0.00", PTaxBill.FETaxDue);
   Print #RptFile, Tab(72); Using$("#####0.00", PTaxBill.FETaxDue)
  Print #RptFile, "Mobile Homes";
   Print #RptFile, Tab(32); Using$("#.00", MHTaxRate#);
   Print #RptFile, Tab(37); Using$("#####0.00", PTaxBill.MHValue);
   Print #RptFile, Tab(51); Using$("#####0.00", PTaxBill.MHTaxDue);
   Print #RptFile, Tab(72); Using$("#####0.00", PTaxBill.MHTaxDue)
  Print #RptFile, "Merchant Capital";
   Print #RptFile, Tab(32); Using$("#.00", MCTaxRate#);
   Print #RptFile, Tab(37); Using$("#####0.00", PTaxBill.MCValue);
   Print #RptFile, Tab(51); Using$("#####0.00", PTaxBill.MCTaxDue);
   Print #RptFile, Tab(72); Using$("#####0.00", PTaxBill.MCTaxDue)
'   TotOth = OldRound(PTaxBill.OptRevTax1 + PTaxBill.OptRevTax2 + PTaxBill.OptRevTax3)
'  If InStr(TaxMasterRec.Name, "CHILHOWIE") Then
'    TotOth = OldRound(PTaxBill.OptRevTax2 + PTaxBill.OptRevTax3)
'  Else
    TotOth = OldRound(PTaxBill.OptRevTax1 + PTaxBill.OptRevTax2 + PTaxBill.OptRevTax3)
'  End If
  If PTaxBill.OverPayAmt > 0 And TotOth = 0 Then '6/22/06
    Print #RptFile, " PPTRA Vehicle Information"; Tab(43); "** Applied Credit Of: " + QPTrim$(Using$("$##,##0.00", PTaxBill.OverPayAmt)) + " **"
  ElseIf PTaxBill.OverPayAmt > 0 And TotOth > 0 Then '6/22/06
     Print #RptFile, " PPTRA Vehicle Information"; Tab(30); "* Applied Credit Of: " + QPTrim$(Using$("$##,##0.00", PTaxBill.OverPayAmt)) + " *"; Tab(60); "Other Taxes: " + Using$("#,##0.00", TotOth)
  ElseIf PTaxBill.OverPayAmt = 0 And TotOth > 0 Then '6/22/06
     Print #RptFile, " PPTRA Vehicle Information"; Tab(60); "Other Taxes: " + Using$("#,##0.00", TotOth)
  Else
    Print #RptFile, " PPTRA Vehicle Information"
  End If
 'Line 30 to 35 here to print vehicles
  CarCount = 0
  WhatPers& = TaxCust.FirstPersRec
  Do
    Get PHandle, WhatPers&, PersRec
    PYearStr$ = CStr(PersRec.TaxBillYear)
    PYear = Val(PYearStr$)
    If PYear > 0 And PYear <> WhatYear Then
      Return
        'Do Not Process This Record
    End If
    If PersRec.PPTRAYN = "Y" Then
      If OldRound#(PersRec.PersVal) > GMaxVehTaxVal Then
        If GMultiYear <> 0 Then
          PersRec.PersVal = OldRound(PersRec.PersVal)
        End If
        PPTRAVal# = GMaxVehTaxVal
      Else
        PPTRAVal# = OldRound#(PersRec.PersVal)
      End If

      If PPTRAVal# <= GMinVehTaxVal Then
        PPTRADiscount# = OldRound#((OldRound#(PPTRAVal# / 100) * PersTaxRate#))
      Else
        PPTRADiscount# = OldRound#((OldRound#((PPTRAVal# / 100) * (PERC! / 100)) * PersTaxRate#))
      End If

      VehDesc$ = " VIN# " + QPTrim$(PersRec.Vin)
      VehDesc$ = QPTrim$(VehDesc$)
      TaxAmt# = OldRound((PersTaxRate# / 100) * PersRec.PersVal)
      PTaxBill.PersTaxDue = PTaxBill.PersTaxDue
      PTaxBill.PPTRADiscnt = PTaxBill.PPTRADiscnt
      Print #RptFile, "*" + VehDesc$;
      Print #RptFile, Tab(37); Using$("#####0.00", PersRec.PersVal) ';
      CarCount = CarCount + 1
    End If
    
    If CarCount >= 6 Then
      Print #RptFile, ""
      Print #RptFile, Tab(48); "Total Tax Due ";
      Print #RptFile, Using$("$#######0.00", OldRound(PTaxBill.TotalBillDue - PTaxBill.OverPayAmt))
      Print #RptFile, Tab(48); "Tax Due Date: " + fptxtPersDueDate.Text 'DueDate$
      Print #RptFile, ""
      Print #RptFile,
      Print #RptFile,
      Print #RptFile,
      Print #RptFile, "BN"; Using("####0", PrnCnt) 'x
      Print #RptFile, "~"

      Print #RptFile, "~"
      Print #RptFile, Tab(62); "TAX YEAR: "; WhatYear
      Print #RptFile, Tab(75); Using$("####0", PTaxBill.BillNumber)
      Print #RptFile, " "
      Print #RptFile, " "
      Print #RptFile, Tab(5); TownName$
      Print #RptFile, Tab(5); Add1$
      Print #RptFile, Tab(5); Add2$
      Print #RptFile, Tab(5); Add3$
      Print #RptFile, " "
      Print #RptFile, " "
      Print #RptFile, " "
      Print #RptFile, Tab(5); "Acct # " + Using$("####0", PTaxBill.CustRec) + " Vehicle Listing Cont'd"
      Print #RptFile, Tab(5); CustName$
      Print #RptFile, Tab(5); QPTrim$(PTaxBill.CustAdd1)
      Print #RptFile, Tab(5); QPTrim$(PTaxBill.CustAdd2)
      Print #RptFile, Tab(5); QPTrim$(PTaxBill.CustAdd3) + " " + Zip
      For LC = 18 To 21
       Print #RptFile, " "
      Next LC
      Print #RptFile, Tab(37); "PROP VALUE"; Tab(50); "TAX AMOUNT"; Tab(62); "PPTRA DIS."; Tab(72); "TOTAL DUE"
      Print #RptFile, " "
      Print #RptFile, "Vehicle Listing Continued ..."
      Print #RptFile, ""
      Print #RptFile, ""
      Print #RptFile, " PPTRA Vehicle Information"
      Print #RptFile, ""
      Print #RptFile, ""
      CarCount = 0
    End If
    WhatPers& = PersRec.NextRec
  Loop While WhatPers& > 0

   ' Finish the bill up here
  If CarCount < 6 Then
    For LCnt = CarCount To 6: Print #RptFile, "": Next LCnt
  End If
'  Print #RptFile, '10/24
  If InStr(PTaxBill.CommentPlace, "LEFT") Then
    Print #RptFile, PTaxBill.Comment; Tab(48); "Total Tax Due ";
    Print #RptFile, Using$("$#######0.00", OldRound(PTaxBill.TotalBillDue - PTaxBill.OverPayAmt))
    Print #RptFile, PTaxBill.Comment2; Tab(48); "Tax Due Date: " + fptxtPersDueDate.Text
    Print #RptFile,
    Print #RptFile,
  ElseIf InStr(PTaxBill.CommentPlace, "RIGHT") Then
    Print #RptFile, Tab(48); "Total Tax Due ";
    Print #RptFile, Using$("$#######0.00", OldRound(PTaxBill.TotalBillDue - PTaxBill.OverPayAmt))
    Print #RptFile, Tab(48); "Tax Due Date: " + fptxtPersDueDate.Text
    Print #RptFile, Tab(48); PTaxBill.Comment
    Print #RptFile, Tab(48); PTaxBill.Comment2
  Else
    Print #RptFile, Tab(48); "Total Tax Due ";
    Print #RptFile, Using$("$#######0.00", OldRound(PTaxBill.TotalBillDue - PTaxBill.OverPayAmt))
    Print #RptFile, Tab(48); "Tax Due Date: " + fptxtPersDueDate.Text
    Print #RptFile, ""
    Print #RptFile, ""
  End If
  Print #RptFile, ""
  Print #RptFile, "BN"; Using$("####0", PrnCnt) 'x
  Print #RptFile, "~"
  
End Sub

Private Sub cmdPList_Click()
  Real = False
  frmVATaxPrintedBillsList.Show vbModal
  DoEvents

End Sub
Private Sub PrintRealVAStandard(RptFile As Integer, TBHandle As Integer, RTaxBill As VARETaxBillType, TCHandle As Integer, TaxCust As TaxCustType, PrnCnt As Long)
 'checked OK against mask (TAXREMSK.DAT) on 10/21/2005
 'STANDARD REAL ESTATE BILL FORMAT AS SOLD BY SOUTHERN SOFTWARE
 'TAXRESTD.BI
  Dim LC As Long, RealTaxRate#
  Dim CustName As String * 45, WhatYear As Integer
  Dim TaxAmt#, LCnt As Integer
  Dim ThisDesc As String * 28
  Dim TotOth As Double
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  RealTaxRate# = fpDblSnglRealRate
  WhatYear = RealYear

  CustName$ = QPTrim$(TaxCust.CustName)
  Print #RptFile, "~"
  Print #RptFile, Tab(64); "TAX YEAR: "; WhatYear
  Print #RptFile, Tab(75); Using$("#####", RTaxBill.BillNumber)
  Print #RptFile, " "
  Print #RptFile, " "
  Print #RptFile, " " 'added
  Print #RptFile, Tab(5); TownName$
  Print #RptFile, Tab(5); Add1$
  Print #RptFile, Tab(5); Add2$
  Print #RptFile, Tab(5); Add3$
  Print #RptFile, " "
  Print #RptFile, " "
  Print #RptFile, " " 'added
  Print #RptFile, " "
  Print #RptFile, Tab(5); "PIN:  " + QPTrim$(RTaxBill.RealPin)
  Print #RptFile, Tab(5); "ACCT: " + Using$("#####", RTaxBill.CustRec)
  Print #RptFile, Tab(5); QPTrim$(RTaxBill.CustName)
  Print #RptFile, Tab(5); Left$(RTaxBill.CustAdd1, 35)
  Print #RptFile, Tab(5); Left$(RTaxBill.CustAdd2, 35)
  Print #RptFile, Tab(5); QPTrim$(RTaxBill.CustAdd3) + " " + InsertZipDash(RTaxBill.CustZip)

  For LC = 19 To 20 'made 18 = 19
    Print #RptFile, " "
  Next LC
  Print #RptFile, Tab(41); "LAND"; Tab(51); "BUILDING"; Tab(61); "NET TOTAL"; Tab(72); "TOTAL DUE"
  Print #RptFile, " "
 'Line 23 Starts Here
  ThisDesc = QPTrim$(RTaxBill.RDesc1)
  Print #RptFile, ThisDesc; 'QPTrim$(RTaxBill.RDesc1);
  Print #RptFile, Tab(30); Using("#0.00", RealTaxRate#);
  If RTaxBill.RealValue > RTaxBill.ExptValue Then
    Print #RptFile, Tab(37); Using("######0.00", (RTaxBill.RealValue - RTaxBill.ExptValue)); ' - RTaxBill.PersValue));
    Print #RptFile, Tab(50); Using("#####0.00", RTaxBill.BldgValue);
  ElseIf RTaxBill.BldgValue > RTaxBill.ExptValue Then
    Print #RptFile, Tab(37); Using("######0.00", RTaxBill.RealValue); ' - RTaxBill.PersValue));
    Print #RptFile, Tab(50); Using("#####0.00", (RTaxBill.BldgValue - RTaxBill.ExptValue));
  ElseIf RTaxBill.RealValue + RTaxBill.BldgValue > RTaxBill.ExptValue Then
        Print #RptFile, Tab(37); Using("######0.00", RTaxBill.RealValue - (RTaxBill.ExptValue * (RTaxBill.RealValue / (RTaxBill.RealValue + RTaxBill.BldgValue)))); ' - RTaxBill.PersValue));
        Print #RptFile, Tab(50); Using("#####0.00", RTaxBill.BldgValue - (RTaxBill.ExptValue * (RTaxBill.BldgValue / (RTaxBill.RealValue + RTaxBill.BldgValue)))); ' - RTaxBill.PersValue));
  Else
    Print #RptFile, Tab(37); Using("######0.00", RTaxBill.RealValue); ' - RTaxBill.PersValue));
    Print #RptFile, Tab(50); Using("#####0.00", RTaxBill.BldgValue);
  End If
  Print #RptFile, Tab(61); Using("#####0.00", OldRound(RTaxBill.RealValue + RTaxBill.BldgValue - RTaxBill.ExptValue));
  Print #RptFile, Tab(71); Using("######0.00", OldRound(RTaxBill.TotalBillDue)) ' - RTaxBill.OverPayAmt))
  Print #RptFile, QPTrim$(RTaxBill.RDesc2)
  TotOth = OldRound(RTaxBill.OptRevTax1 + RTaxBill.OptRevTax2 + RTaxBill.OptRevTax3 + RTaxBill.LateTaxDue)
  If RTaxBill.OverPayAmt > 0 And TotOth > 0 Then
    Print #RptFile, "** Applied Credit Of " + QPTrim$(Using$("$##,##0.00", RTaxBill.OverPayAmt)) + " **"; Tab(40); "Other Taxes Included: " + Using$("$#,##0.00", TotOth)
    For LCnt = 26 To 36: Print #RptFile, "": Next LCnt
  ElseIf RTaxBill.OverPayAmt > 0 And TotOth = 0 Then
    Print #RptFile, "** Applied Credit Of " + QPTrim$(Using$("$##,##0.00", RTaxBill.OverPayAmt)) + " **"
    For LCnt = 26 To 36: Print #RptFile, "": Next LCnt
  ElseIf RTaxBill.OverPayAmt = 0 And TotOth > 0 Then
    Print #RptFile, Tab(40); "Other Taxes Included: " + Using$("$#,##0.00", TotOth)
    For LCnt = 26 To 36: Print #RptFile, "": Next LCnt
  Else
    For LCnt = 25 To 36: Print #RptFile, "": Next LCnt
  End If
 'Lines 25 to 36 are blank
'Line 37 for Totals
'       Print #RptFile, ""
'  Print #RptFile,
'  Print #RptFile,
'  Print #RptFile,
'  Print #RptFile, RTaxBill.Comment; Tab(48); "Total Tax Due ... ";
'  Print #RptFile, Using$("$######0.00", OldRound(RTaxBill.TotalBillDue - RTaxBill.OverPayAmt))
'  Print #RptFile, RTaxBill.Comment2; Tab(48); "Tax Due Date: " + fptxtRealDueDate.Text ' DueDate$
  If InStr(RTaxBill.CommentPlace, "LEFT") Then
    Print #RptFile, RTaxBill.Comment; Tab(48); "Total Tax Due ";
    Print #RptFile, Using$("$#######0.00", OldRound(RTaxBill.TotalBillDue - RTaxBill.OverPayAmt))
    Print #RptFile, RTaxBill.Comment2; Tab(48); "Tax Due Date: " + fptxtRealDueDate.Text
    Print #RptFile,
    Print #RptFile,
  ElseIf InStr(RTaxBill.CommentPlace, "RIGHT") Then
    Print #RptFile, Tab(48); "Total Tax Due ";
    Print #RptFile, Using$("$#######0.00", OldRound(RTaxBill.TotalBillDue - RTaxBill.OverPayAmt))
    Print #RptFile, Tab(48); "Tax Due Date: " + fptxtRealDueDate.Text
    Print #RptFile, Tab(48); RTaxBill.Comment
    Print #RptFile, Tab(48); RTaxBill.Comment2
  Else
    Print #RptFile, Tab(48); "Total Tax Due ";
    Print #RptFile, Using$("$#######0.00", OldRound(RTaxBill.TotalBillDue - RTaxBill.OverPayAmt))
    Print #RptFile, Tab(48); "Tax Due Date: " + fptxtRealDueDate.Text
    Print #RptFile,
    Print #RptFile,
  End If
  Print #RptFile, "BN"; Using$("#####", PrnCnt)
  Print #RptFile, "~"
  
End Sub

Private Sub fpcmbType_Change()
  If fpcmbType.Text = "REAL" Then
    fpDblSnglRealFirstBill.Enabled = True
    fpDblSnglRealLastBill.Enabled = True
    fpDblSnglPersFirstBill.Enabled = False
    fpDblSnglPersLastBill.Enabled = False
    cmdPList.Enabled = False
    cmdRList.Enabled = True
  ElseIf fpcmbType.Text = "PERSONAL" Then
    fpDblSnglRealFirstBill.Enabled = False
    fpDblSnglRealLastBill.Enabled = False
    fpDblSnglPersFirstBill.Enabled = True
    fpDblSnglPersLastBill.Enabled = True
    cmdRList.Enabled = False
    cmdPList.Enabled = True
  End If
    
End Sub

Private Sub PrintLaserPers1()
  Dim ToPrint As String
  Dim TaxRptT As Integer
  Dim ReportFile As String
  Dim RptHandle As Integer
  Dim TBRec As VAPPTaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim x As Long
  Dim dlm$, BillNo&, PrnCnt As Long
  Dim TBDRec As TxBillLaser1DefaultsType
  Dim TBDHandle As Integer
  Dim ThisRate As Double
  Dim TotValue As Double
  Dim FBill&
  Dim LBill&
  Dim ThisOpt1Desc As String * 15
  Dim ThisOpt2Desc As String * 15
  Dim ThisOpt3Desc As String * 15
  Dim BZip As String
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim ZipRec As BillPrintPZipIdxType, CustArr As Long
  Dim ZHandle As Integer
  Dim NumOfZRecs As Long
  
  On Error GoTo ERRORSTUFF
  
  dlm$ = "~"
  ReportFile$ = StartPath$ + "/TaxPBil1.PRN"
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  FBill = fpDblSnglPersFirstBill.Value
  LBill = fpDblSnglPersLastBill.Value
  
  frmVATaxShowPctComp.Label1 = "Printing Personal Tax Bills"
  frmVATaxShowPctComp.cmdCancel.Visible = False
  frmVATaxShowPctComp.Show , Me
  cmdProcess.Enabled = False
  cmdExit.Enabled = False
  EnableCloseButton Me.hwnd, False
  
  OpenTxBillPersFile TBDHandle
  Get #TBDHandle, 1, TBDRec
  Close TBDHandle
  
  arVATaxBillPersLaser.Head1 = QPTrim(TBDRec.TxtHead1)
  arVATaxBillPersLaser.Head2 = QPTrim(TBDRec.TxtHead2)
  arVATaxBillPersLaser.LblOpt1 = QPTrim(TBDRec.txtOpt1)
  arVATaxBillPersLaser.LblOpt2 = QPTrim(TBDRec.TxtOpt2)
  arVATaxBillPersLaser.LblOpt3 = QPTrim(TBDRec.TxtOpt3)
  arVATaxBillPersLaser.LblOpt4 = QPTrim(TBDRec.TxtOpt4)
  arVATaxBillPersLaser.LblPgph1 = QPTrim(TBDRec.txtPgph0)
  arVATaxBillPersLaser.LblPgph2 = QPTrim(TBDRec.txtPgph1)
  arVATaxBillPersLaser.LblPgph3 = QPTrim(TBDRec.txtPgph2)
  arVATaxBillPersLaser.LblPgph4 = QPTrim(TBDRec.txtPgph3)
  arVATaxBillPersLaser.LblPgph5 = QPTrim(TBDRec.txtPgph4)
  arVATaxBillPersLaser.LblPgph6 = QPTrim(TBDRec.txtPgph5)
  arVATaxBillPersLaser.LblPgph7 = QPTrim(TBDRec.txtPgph6)
  arVATaxBillPersLaser.LblPgph8 = QPTrim(TBDRec.txtPgph7)
  arVATaxBillPersLaser.LblOpt5 = QPTrim(TBDRec.TxtOpt5)
  arVATaxBillPersLaser.LblHead4 = QPTrim(TBDRec.txtHead4)
  arVATaxBillPersLaser.LblHead5 = QPTrim(TBDRec.txtHead5)
  arVATaxBillPersLaser.LblHead6 = QPTrim(TBDRec.txtHead6)
  arVATaxBillPersLaser.LblOpt6 = QPTrim(TBDRec.TxtOpt6)
  arVATaxBillPersLaser.LblOpt7 = QPTrim(TBDRec.TxtOpt7)
  If TBDRec.dologo = 1 Then
    If Exist("towntaxlogo.bmp") Then
      arVATaxBillPersLaser.Image1.Picture = LoadPicture("towntaxlogo.bmp")
      arVATaxBillPersLaser.Image1.Visible = True
    End If
  End If
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenPersTaxBillFile TBHandle, NumOfTBRecs
  If Exist("PZipIdx.Dat") Then '12/6/06
    OpenPZipIdxFile ZHandle, NumOfZRecs
    NumOfTBRecs = NumOfZRecs
  End If
  
  For x = 1 To NumOfTBRecs
    If NumOfZRecs > 0 Then '12/6/06
      Get ZHandle, x, ZipRec
      CustArr = ZipRec.TaxBillRec
    Else
      CustArr = x '12/6/06
    End If
    Get TBHandle, CustArr, TBRec
      If TBRec.BillPrinted = False Then GoTo SkipIt
      If TBRec.BillNumber >= FBill And TBRec.BillNumber <= LBill Then
        If TBRec.PPTRAValue > 0 Then
          TotValue = OldRound(TBRec.PPTRAValue + TBRec.FEValue + TBRec.MCValue + TBRec.MHValue + TBRec.MTValue - (TBRec.ExptValue)) ' + TBRec.PPTRAValue))
        Else
          TotValue = OldRound(TBRec.PersValue + TBRec.FEValue + TBRec.MCValue + TBRec.MHValue + TBRec.MTValue - (TBRec.ExptValue)) ' + TBRec.PPTRAValue))
        End If
        GoSub GetBarCodeData
        '                         0                         1
        Print #RptHandle, TBRec.BillNumber; dlm; QPTrim$(TBRec.CustName); dlm;
        '                           2                           3
        Print #RptHandle, QPTrim$(TBRec.CustAdd1); dlm; QPTrim$(TBRec.CustAdd2); dlm;
        '                             4                      5
        Print #RptHandle, QPTrim$(TBRec.CustAdd3); dlm; TBRec.CustPin; dlm;
        '                   6                 7
        Print #RptHandle, TotValue; dlm; TBRec.RDesc1; dlm;
        '                        8                   9                    10
        Print #RptHandle, TBRec.PersValue; dlm; TBRec.FEValue; dlm; TBRec.ExptValue; dlm;
        '                    11                         12                                          13
        Print #RptHandle, TBRec.PPTRAValue; dlm; TBRec.PPTRADiscnt; dlm; OldRound(TBRec.TotalBillDue - TBRec.OverPayAmt); dlm;
        '                       14                 15                   16
        Print #RptHandle, TBDRec.dologo; dlm; TBRec.MHValue; dlm; TBRec.MCValue; dlm;
        '                      17                     18                             19
        Print #RptHandle, TBRec.MTValue; dlm; OldRound(TBRec.PersTaxDue); dlm; TBRec.PersTaxNet; dlm;
        '                         20                    21                    22
        Print #RptHandle, TBRec.PersTaxRate; dlm; TBRec.FETaxDue; dlm; TBRec.FETaxRate; dlm;
        '                       23                    24                    25
        Print #RptHandle, TBRec.MCTaxDue; dlm; TBRec.MCTaxRate; dlm; TBRec.MHTaxDue; dlm;
        '                       26                    27                     28
        Print #RptHandle, TBRec.MHTaxRate; dlm; TBRec.MTTaxDue; dlm; TBRec.MTTaxRate; dlm;
        '                        29                     30                     31
        Print #RptHandle, TBRec.OptRevTax1; dlm; TBRec.OptRevTax2; dlm; TBRec.OptRevTax3; dlm;
        ThisOpt1Desc = QPTrim$(TBRec.OptRevDesc1)
        ThisOpt2Desc = QPTrim$(TBRec.OptRevDesc2)
        ThisOpt3Desc = QPTrim$(TBRec.OptRevDesc3)
        '                        32              33                 34              35             36                  37                    38
        Print #RptHandle, ThisOpt1Desc; dlm; ThisOpt2Desc; dlm; ThisOpt3Desc; dlm; BZip; dlm; TBRec.CustZip; dlm; TBDRec.dologo; dlm; TBRec.OverPayAmt; dlm;
        '                           39                      40
        Print #RptHandle, TBRec.PriorYrBalance; dlm; TBRec.PrintPrior
      End If
SkipIt:
    frmVATaxShowPctComp.ShowPctComp x, NumOfTBRecs
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      cmdProcess.Enabled = True
      cmdExit.Enabled = True
      EnableCloseButton Me.hwnd, True
      Exit Sub
    End If
  Next x
  Unload frmVATaxShowPctComp
  cmdProcess.Enabled = True
  cmdExit.Enabled = True
  EnableCloseButton Me.hwnd, True
  
  Close
  arVATaxBillPersLaser.GetName ReportFile$
  arVATaxBillPersLaser.Show
  
  Exit Sub
  
GetBarCodeData:
  If TBDRec.UseBarCode = False Then
    BZip = ""
    Return
  ElseIf TBDRec.UseBarCode = True Then
    Get TCHandle, TBRec.CustPin, TaxCust
    If Len(QPTrim$(TaxCust.Zip)) < 10 Or Len(QPTrim$(TaxCust.DeliveryPt)) <> 2 Then
      BZip = ""
    Else
      BZip = QPTrim$(TaxCust.Zip) + QPTrim$(TaxCust.DeliveryPt)
    End If
  End If
  
  Return

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxBillPrinting", "PrintLaser1", Erl)
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
 

End Sub

Private Sub PrintLaserReal1()
  Dim ToPrint As String
  Dim TaxRptT As Integer
  Dim ReportFile As String
  Dim RptHandle As Integer
  Dim TBRec As VARETaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim x As Long
  Dim dlm$
  Dim TBDRec As TxBillLaser1DefaultsType
  Dim TBDHandle As Integer
  Dim ThisRate As Double
  Dim FBill&
  Dim LBill&
  Dim BZip$
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim CustArr As Long '12/6/06
  Dim ZipRec As BillPrintPZipIdxType
  Dim ZHandle As Integer
  Dim NumOfZRecs As Long
  Dim MortRec As BillPrintMortIdxType
  Dim MRHandle As Integer
  Dim NumOfMRRecs As Long
  
  On Error GoTo ERRORSTUFF
  
  dlm$ = "~"
  ReportFile$ = StartPath$ + "/TaxRBil1.PRN"
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  FBill = fpDblSnglRealFirstBill.Value
  LBill = fpDblSnglRealLastBill.Value
  
  frmVATaxShowPctComp.Label1 = "Printing Real Tax Bills"
  frmVATaxShowPctComp.cmdCancel.Visible = False
  frmVATaxShowPctComp.Show , Me
  cmdProcess.Enabled = False
  cmdExit.Enabled = False
  EnableCloseButton Me.hwnd, False
  
  OpenTxBillRealFile TBDHandle
  Get #TBDHandle, 1, TBDRec
  Close TBDHandle
  
  ARptVATempTaxBill.Head1 = QPTrim(TBDRec.TxtHead1)
  ARptVATempTaxBill.Head2 = QPTrim(TBDRec.TxtHead2)
  ARptVATempTaxBill.LblOpt1 = QPTrim(TBDRec.txtOpt1)
  ARptVATempTaxBill.LblOpt2 = QPTrim(TBDRec.TxtOpt2)
  ARptVATempTaxBill.LblOpt3 = QPTrim(TBDRec.TxtOpt3)
  ARptVATempTaxBill.LblOpt4 = QPTrim(TBDRec.TxtOpt4)
  ARptVATempTaxBill.LblPgph1 = QPTrim(TBDRec.txtPgph0)
  ARptVATempTaxBill.LblPgph2 = QPTrim(TBDRec.txtPgph1)
  ARptVATempTaxBill.LblPgph3 = QPTrim(TBDRec.txtPgph2)
  ARptVATempTaxBill.LblPgph4 = QPTrim(TBDRec.txtPgph3)
  ARptVATempTaxBill.LblPgph5 = QPTrim(TBDRec.txtPgph4)
  ARptVATempTaxBill.LblPgph6 = QPTrim(TBDRec.txtPgph5)
  ARptVATempTaxBill.LblPgph7 = QPTrim(TBDRec.txtPgph6)
  ARptVATempTaxBill.LblPgph8 = QPTrim(TBDRec.txtPgph7)
  ARptVATempTaxBill.LblOpt5 = QPTrim(TBDRec.TxtOpt5)
  ARptVATempTaxBill.LblHead4 = QPTrim(TBDRec.txtHead4)
  ARptVATempTaxBill.LblHead5 = QPTrim(TBDRec.txtHead5)
  ARptVATempTaxBill.LblHead6 = QPTrim(TBDRec.txtHead6)
  ARptVATempTaxBill.LblOpt6 = QPTrim(TBDRec.TxtOpt6)
  ARptVATempTaxBill.LblOpt7 = QPTrim(TBDRec.TxtOpt7)
  If TBDRec.dologo = 1 Then
    If Exist("towntaxlogo.bmp") Then
      ARptVATempTaxBill.Image1.Picture = LoadPicture("towntaxlogo.bmp")
      ARptVATempTaxBill.Image1.Visible = True
    End If
  End If
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenRealTaxBillFile TBHandle, NumOfTBRecs
  If Exist("MORTIDX.DAT") Then '12/6/06
    OpenMortIdxFile MRHandle, NumOfMRRecs
    NumOfTBRecs = NumOfMRRecs
  ElseIf Exist("RZIPIDX.DAT") Then '12/6/06
    OpenRZipIdxFile ZHandle, NumOfZRecs
    NumOfTBRecs = NumOfZRecs
  End If
  
  For x = 1 To NumOfTBRecs
    If NumOfMRRecs > 0 Then '12/6/06
      Get MRHandle, x, MortRec
      CustArr = MortRec.TaxBillRec
    ElseIf NumOfZRecs > 0 Then '12/6/06
      Get ZHandle, x, ZipRec
      CustArr = ZipRec.TaxBillRec
    Else
      CustArr = x '12/6/06
    End If
    Get TBHandle, CustArr, TBRec
    If TBRec.BillPrinted = False Then GoTo SkipIt
    If TBRec.BillNumber > 0 Then
      If TBRec.BillNumber >= FBill And TBRec.BillNumber <= LBill Then
'        Put TBHandle, x, TBRec
        GoSub GetBarCodeData
        '                         0                         1
        Print #RptHandle, TBRec.BillNumber; dlm; QPTrim$(TBRec.CustName); dlm;
        '                           2                           3
        Print #RptHandle, QPTrim$(TBRec.CustAdd1); dlm; QPTrim$(TBRec.CustAdd2); dlm;
        '                             4                          5
        Print #RptHandle, QPTrim$(TBRec.CustAdd3); dlm; TBRec.CustPin; dlm;
        '                            6                            7
        Print #RptHandle, QPTrim$(TBRec.RDesc2); dlm; QPTrim$(TBRec.RDesc1); dlm;
        '                        8                     9                    10
        Print #RptHandle, TBRec.RealValue; dlm; TBRec.BldgValue; dlm; TBRec.ExptValue; dlm;
        If OldRound(TBRec.RealTaxDue - TBRec.OverPayAmt) > 0 Then
          ThisRate = TBRec.RealTaxRate
        Else
          ThisRate = 0
        End If
        '                                         11                              12
        Print #RptHandle, OldRound(TBRec.RealValue + TBRec.BldgValue - TBRec.ExptValue); dlm; ThisRate; dlm;
        '                                     13                                 14                15                       16
        Print #RptHandle, OldRound(TBRec.TotalBillDue - TBRec.OverPayAmt); dlm; BZip; dlm; QPTrim$(TBRec.CustZip); dlm; TBDRec.dologo; dlm;
        '                        17                     18                     19
        Print #RptHandle, TBRec.OptRevTax1; dlm; TBRec.OptRevTax2; dlm; TBRec.OptRevTax3; dlm;
        '                       20                   21                  22                    23                     24                       25                        26
        Print #RptHandle, TBRec.Opt1Desc; dlm; TBRec.Opt2Desc; dlm; TBRec.Opt3Desc; dlm; TBRec.OverPayAmt; dlm; TBRec.LateTaxDue; dlm; TBRec.PriorYrBalance; dlm; TBRec.PrintPrior
      End If
    End If
SkipIt:
    frmVATaxShowPctComp.ShowPctComp x, NumOfTBRecs
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      cmdProcess.Enabled = True
      cmdExit.Enabled = True
      EnableCloseButton Me.hwnd, True
      Exit Sub
    End If
  Next x
  Unload frmVATaxShowPctComp
  cmdProcess.Enabled = True
  cmdExit.Enabled = True
  EnableCloseButton Me.hwnd, True
  
  Close
  
  ARptVATempTaxBill.GetName ReportFile$
  
  ARptVATempTaxBill.Show
  
  Exit Sub
  
GetBarCodeData:
  If TBDRec.UseBarCode = False Then
    BZip = ""
    Return
  ElseIf TBDRec.UseBarCode = True Then
    Get TCHandle, TBRec.CustPin, TaxCust
    If Len(QPTrim$(TaxCust.Zip)) < 10 Or Len(QPTrim$(TaxCust.DeliveryPt)) <> 2 Then
      BZip = ""
    Else
      BZip = QPTrim$(TaxCust.Zip) + QPTrim$(TaxCust.DeliveryPt)
    End If
  End If
  
  Return

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxBillPrinting", "PrintLaser1", Erl)
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
 

End Sub

Private Sub PrintPersLaserItemized()
  Dim ToPrint As String
  Dim TaxRptT As Integer
  Dim ReportFile As String
  Dim RptHandle As Integer
  Dim TBRec As VAPPTaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim x As Long
  Dim dlm$, BillNo&, PrnCnt As Long
  Dim TBDRec As TxBillLaserItemized
  Dim TBDHandle As Integer
  Dim ThisRate As Double
  Dim TotValue As Double
  Dim ThisOpt1Desc As String * 15
  Dim ThisOpt2Desc As String * 15
  Dim ThisOpt3Desc As String * 15
  Dim BZip As String
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim PersRec As PersonalRecType
  Dim PHandle As Integer
  Dim NumOfPRecs As Long
  Dim NumOfPers As Integer
  Dim NextRec As Long
  Dim thisVin As String
  Dim y As Integer
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim NumOfOpts As Integer
  Dim FBill&
  Dim LBill&
  Dim MinTaxedAmt As Double
  Dim PocahFlag As Boolean
  Dim ZipRec As BillPrintPZipIdxType, CustArr As Long
  Dim ZHandle As Integer
  Dim NumOfZRecs As Long
  
  On Error GoTo ERRORSTUFF
  
  FBill = fpDblSnglPersFirstBill.Value
  LBill = fpDblSnglPersLastBill.Value
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  NumOfOpts = 0
  If QPTrim$(TaxMasterRec.POptRev1) <> "" Then
    NumOfOpts = NumOfOpts + 1
  End If
  If QPTrim$(TaxMasterRec.POptRev2) <> "" Then
    NumOfOpts = NumOfOpts + 1
  End If
  If QPTrim$(TaxMasterRec.POptRev3) <> "" Then
    NumOfOpts = NumOfOpts + 1
  End If
  If InStr(TaxMasterRec.Name, "POCAHONTAS") > 0 Then
    PocahFlag = True
  End If
  MinTaxedAmt = TaxMasterRec.MinVehTaxVal
  
  dlm$ = "~"
  ReportFile$ = StartPath$ + "/TaxPLsrItem.RPT"
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  BillNo& = fpDblSnglStartPersBill.Value
  
  frmVATaxShowPctComp.Label1 = "Printing Personal Tax Bills"
  frmVATaxShowPctComp.cmdCancel.Visible = False
  frmVATaxShowPctComp.Show , Me
  cmdProcess.Enabled = False
  cmdExit.Enabled = False
  EnableCloseButton Me.hwnd, False
  
  OpenLaserPersItemized TBDHandle
  Get #TBDHandle, 1, TBDRec
  Close TBDHandle
  
  If TBDRec.dologo = 1 Then
    If Exist("towntaxlogo.bmp") Then
      arVATaxLaserPersItemized.Image1.Picture = LoadPicture("towntaxlogo.bmp")
      arVATaxLaserPersItemized.Image1.Visible = True
    End If
  End If
  OpenTaxCustFile TCHandle, NumOfTCRecs
  ReDim VinArray(1 To 1) As String
  OpenPersPropFile PHandle, NumOfPRecs
  OpenPersTaxBillFile TBHandle, NumOfTBRecs
  If Exist("PZipIdx.Dat") Then '12/6/06
    OpenPZipIdxFile ZHandle, NumOfZRecs
    NumOfTBRecs = NumOfZRecs
  End If
  
  For x = 1 To NumOfTBRecs
    NumOfPers = 0
    If NumOfZRecs > 0 Then '12/6/06
      Get ZHandle, x, ZipRec
      CustArr = ZipRec.TaxBillRec
    Else
      CustArr = x '12/6/06
    End If
    Get TBHandle, CustArr, TBRec
    Get TCHandle, TBRec.CustPin, TaxCust
    If TBRec.BillNumber > 0 And TBRec.BillPrinted = True Then
      If TBRec.BillNumber >= FBill And TBRec.BillNumber <= LBill Then
        If TBRec.PPTRAValue > 0 Then
          TotValue = OldRound(TBRec.PPTRAValue + TBRec.FEValue + TBRec.MCValue + TBRec.MHValue + TBRec.MTValue - (TBRec.ExptValue)) ' + TBRec.PPTRAValue))
        Else
          TotValue = OldRound(TBRec.PersValue + TBRec.FEValue + TBRec.MCValue + TBRec.MHValue + TBRec.MTValue - (TBRec.ExptValue)) ' + TBRec.PPTRAValue))
        End If
        GoSub GetBarCodeData
        '                         0                         1
        Print #RptHandle, TBRec.BillNumber; dlm; QPTrim$(TBRec.CustName); dlm;
        '                           2                           3
        Print #RptHandle, QPTrim$(TBRec.CustAdd1); dlm; QPTrim$(TBRec.CustAdd2); dlm;
        '                             4                      5
        Print #RptHandle, QPTrim$(TBRec.CustAdd3); dlm; TBRec.CustPin; dlm;
        '                     6               7
        Print #RptHandle, TotValue; dlm; TBRec.RDesc1; dlm;
        '                        8                   9                    10
        Print #RptHandle, TBRec.PersValue; dlm; TBRec.FEValue; dlm; TBRec.ExptValue; dlm;
        '                    11                      12                        13
        Print #RptHandle, TBRec.PPTRAValue; dlm; TBRec.PPTRADiscnt; dlm; OldRound(TBRec.TotalBillDue - TBRec.OverPayAmt); dlm;
        '                       14                 15                   16
        Print #RptHandle, TBDRec.dologo; dlm; TBRec.MHValue; dlm; TBRec.MCValue; dlm;
        '                      17                            18                        19
        Print #RptHandle, TBRec.MTValue; dlm; OldRound(TBRec.PersTaxDue); dlm; TBRec.PersTaxNet; dlm;
        '                         20                    21                    22
        Print #RptHandle, TBRec.PersTaxRate; dlm; TBRec.FETaxDue; dlm; TBRec.FETaxRate; dlm;
        '                       23                    24                    25
        Print #RptHandle, TBRec.MCTaxDue; dlm; TBRec.MCTaxRate; dlm; TBRec.MHTaxDue; dlm;
        '                       26                    27                     28
        Print #RptHandle, TBRec.MHTaxRate; dlm; TBRec.MTTaxDue; dlm; TBRec.MTTaxRate; dlm;
        '                        29                     30                     31
        Print #RptHandle, TBRec.OptRevTax1; dlm; TBRec.OptRevTax2; dlm; TBRec.OptRevTax3; dlm;
        ThisOpt1Desc = QPTrim$(TBRec.OptRevDesc1)
        ThisOpt2Desc = QPTrim$(TBRec.OptRevDesc2)
        ThisOpt3Desc = QPTrim$(TBRec.OptRevDesc3)
        '                        32              33                 34              35             36                  37
        Print #RptHandle, ThisOpt1Desc; dlm; ThisOpt2Desc; dlm; ThisOpt3Desc; dlm; BZip; dlm; TBRec.CustZip; dlm; TBDRec.dologo; dlm;
        '                           38                           39                            40
        Print #RptHandle, QPTrim(TBDRec.TxtHead1); dlm; QPTrim(TBDRec.TxtHead2); dlm; QPTrim(TBDRec.txtOpt1); dlm;
        '                           41                           42                 43
        Print #RptHandle, QPTrim(TBDRec.TxtOpt2); dlm; QPTrim(TBDRec.TxtOpt3); dlm; ""; dlm;
        '                           44                           45                             46
        Print #RptHandle, QPTrim(TBDRec.txtPgph0); dlm; QPTrim(TBDRec.txtPgph1); dlm; QPTrim(TBDRec.txtPgph2); dlm;
        '                           47                           48                            49
        Print #RptHandle, QPTrim(TBDRec.txtPgph3); dlm; QPTrim(TBDRec.txtPgph4); dlm; "                     "; dlm;
        '                           50               51       52
        Print #RptHandle, "                  "; dlm; ""; dlm; ""; dlm;
        '                           53                           54                            55
        Print #RptHandle, QPTrim(TBDRec.txtHead3); dlm; QPTrim(TBDRec.txtHead4); dlm; QPTrim(TBDRec.txtHead5); dlm;
        '                 56       57
        Print #RptHandle, ""; dlm; ""; dlm;
        NextRec = TaxCust.FirstPersRec
        If NextRec > 0 Then
          Get PHandle, NextRec, PersRec
          NumOfPers = NumOfPers + 1
        
          '                           58                     59                    60
          Print #RptHandle, QPTrim$(PersRec.Vin); dlm; QPTrim$(PersRec.MakeMod); dlm; PersRec.PersVal; dlm;
          '                       61                     62                   63                  64                  65
          Print #RptHandle, PersRec.MTValue; dlm; PersRec.MCValue; dlm; PersRec.CVALUE; dlm; PersRec.MHValue; dlm; NumOfPers; dlm; TBRec.OverPayAmt; dlm; TBRec.TotalBillDue
          NextRec = PersRec.NextRec
        Else
          '                 58       59       60       61       62       63       64          65
          Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; NumOfPers; dlm; TBRec.OverPayAmt; dlm; TBRec.TotalBillDue
        End If
'        Do While NextRec > 0
        
        Do While NextRec > 0
          Get PHandle, NextRec, PersRec
          If PocahFlag Then  '092414
            If PersRec.PersVal < MinTaxedAmt Then
              GoTo SkipThisCar
            End If
          End If

          NumOfPers = NumOfPers + 1
          '                         0                         1                           2                                3
          Print #RptHandle, TBRec.BillNumber; dlm; QPTrim$(TBRec.CustName); dlm; QPTrim$(TBRec.CustAdd1); dlm; QPTrim$(TBRec.CustAdd2); dlm;
          '                             4                        5               6            7        8        9        10
          Print #RptHandle, QPTrim$(TBRec.CustAdd3); dlm; TBRec.CustPin; dlm; TotValue; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm;
          '                 11       12                          13
          Print #RptHandle, ""; dlm; ""; dlm; OldRound(TBRec.TotalBillDue - TBRec.OverPayAmt); dlm;
          '                 14       15       16       17       18
          Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm;
          '                 19       20       21       22       23       24       25       26
          Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm;
          '                 27       28       29       30       31           32                  33                34
          Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ThisOpt1Desc; dlm; ThisOpt2Desc; dlm; ThisOpt3Desc; dlm;
          '                 35       36       37                38                            39                            40
          Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; QPTrim(TBDRec.TxtHead1); dlm; QPTrim(TBDRec.TxtHead2); dlm; QPTrim(TBDRec.txtOpt1); dlm;
          '                            41                           42
          Print #RptHandle, QPTrim(TBDRec.TxtOpt2); dlm; QPTrim(TBDRec.TxtOpt3); dlm;
          '                 43                 44                           45                            46
          Print #RptHandle, ""; dlm; QPTrim(TBDRec.txtPgph0); dlm; QPTrim(TBDRec.txtPgph1); dlm; QPTrim(TBDRec.txtPgph2); dlm;
          '                            47                 48       49       50
          Print #RptHandle, QPTrim(TBDRec.txtPgph3); dlm; ""; dlm; ""; dlm; ""; dlm;
          '                 51       52                53                               54                        55
          Print #RptHandle, ""; dlm; ""; dlm; QPTrim(TBDRec.txtHead3); dlm; QPTrim(TBDRec.txtHead4); dlm; QPTrim(TBDRec.txtHead5); dlm;
          '                 56       57              58
          Print #RptHandle, ""; dlm; ""; dlm; QPTrim$(PersRec.Vin); dlm;
          '                         59                           60                   61                       62
          Print #RptHandle, QPTrim$(PersRec.MakeMod); dlm; PersRec.PersVal; dlm; PersRec.MTValue; dlm; PersRec.MCValue; dlm;
          '                        63                   64                65
          Print #RptHandle, PersRec.CVALUE; dlm; PersRec.MHValue; dlm; NumOfPers; dlm; TBRec.OverPayAmt; dlm; TBRec.TotalBillDue
          
          If NumOfOpts = 1 Then
            If NumOfPers = 13 Then NumOfPers = 0
          ElseIf NumOfOpts = 2 Then
            If NumOfPers = 12 Then NumOfPers = 0
          ElseIf NumOfOpts = 3 Then
            If NumOfPers = 11 Then NumOfPers = 0
          Else
            If NumOfPers = 14 Then NumOfPers = 0
          End If
SkipThisCar:
          NextRec = PersRec.NextRec
        Loop
        BillNo& = BillNo& + 1
        PrnCnt = PrnCnt + 1
      End If
    End If
    frmVATaxShowPctComp.ShowPctComp x, NumOfTBRecs
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      cmdProcess.Enabled = True
      cmdExit.Enabled = True
      EnableCloseButton Me.hwnd, True
      Exit Sub
    End If
  Next x
  Unload frmVATaxShowPctComp
  cmdProcess.Enabled = True
  cmdExit.Enabled = True
  EnableCloseButton Me.hwnd, True
  
  Close
  arVATaxLaserPersItemized.Show
  
  Exit Sub
  
GetBarCodeData:
  If TBDRec.UseBarCode = False Then
    BZip = ""
    Return
  ElseIf TBDRec.UseBarCode = True Then
    Get TCHandle, TBRec.CustPin, TaxCust
    If Len(QPTrim$(TaxCust.Zip)) < 10 Or Len(QPTrim$(TaxCust.DeliveryPt)) <> 2 Then
      BZip = ""
    Else
      BZip = QPTrim$(TaxCust.Zip) + QPTrim$(TaxCust.DeliveryPt)
    End If
  End If
  
  Return
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxBillRePrinting", "PrintPersLaserItemized", Erl)
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
 
End Sub

Private Sub PrintMdltwnReal()
  Dim TaxBill As VARETaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim BillInfo As VARETaxBillInfoType
  Dim BIHandle As Integer
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long, RealTaxRate#
  Dim File$, WordLen As Integer
  Dim CustName As String * 45
  Dim RptFile#, ch$, y As Integer
  Dim CarCount As Integer
  Dim TownName$, Add1$, Add2$, Add3$
  Dim RHandle As Integer
  Dim NumOfRRecs As Long
  Dim RealRec As PropertyRecType
  Dim TaxAmt#, Tab1 As Integer, Tab2 As Integer, Tab3 As Integer
  Dim DueDate$, WorkName$
  Dim FBill&, TotOpt As Double
  Dim LBill&
  Dim ZipRec As BillPrintRZipIdxType, CustArr As Long
  Dim ZHandle As Integer
  Dim NumOfZRecs As Long
  Dim MortRec As BillPrintMortIdxType
  Dim MRHandle As Integer
  Dim NumOfMRRecs As Long
  
  FBill = fpDblSnglRealFirstBill.Value
  LBill = fpDblSnglRealLastBill.Value
  DueDate$ = fptxtRealDueDate
  RealTaxRate# = fpDblSnglRealRate
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  TownName = QPTrim$(TaxMasterRec.Name)
  Add1$ = QPTrim$(TaxMasterRec.Add1)
  Add2$ = QPTrim$(TaxMasterRec.Add2)
  Add3$ = QPTrim$(TaxMasterRec.City) + ", " + QPTrim$(TaxMasterRec.TownState) + " " + QPTrim$(TaxMasterRec.Zip)
  
  OpenRealPropFile RHandle, NumOfRRecs
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenRealTaxBillFile TBHandle, NumOfTBRecs
  OpenRealBillInfoFile BIHandle
  Get BIHandle, 1, BillInfo
  Close BIHandle
  File$ = StartPath$ + "/TxBMdltwnRE.PRN"
  RptFile# = FreeFile
  Open File$ For Output As #RptFile
  
  'Must Calc Late Fee Here
  frmVATaxShowPctComp.Label1 = "Creating Real Bills"
  frmVATaxShowPctComp.Show , Me
  frmVATaxShowPctComp.cmdCancel.Visible = False
  EnableCloseButton Me.hwnd, False
  cmdExit.Enabled = False
  cmdProcess.Enabled = False
  GoSub LoadHeaders
  If Exist("MORTIDX.DAT") Then '12/6/06
    OpenMortIdxFile MRHandle, NumOfMRRecs
    NumOfTBRecs = NumOfMRRecs
  ElseIf Exist("RZIPIDX.DAT") Then '12/6/06
    OpenRZipIdxFile ZHandle, NumOfZRecs
    NumOfTBRecs = NumOfZRecs
  End If
  
  Tab1 = 44 - Len(TownName) / 2
  Tab2 = 44 - Len(Add1) / 2
  Tab3 = 44 - Len(Add3) / 2
  For x = 1 To NumOfTBRecs
    If NumOfMRRecs > 0 Then '12/6/06
      Get MRHandle, x, MortRec
      CustArr = MortRec.TaxBillRec
    ElseIf NumOfZRecs > 0 Then '12/6/06
      Get ZHandle, x, ZipRec
      CustArr = ZipRec.TaxBillRec
    Else
      CustArr = x '12/6/06
    End If
    Get TBHandle, CustArr, TaxBill
    If TaxBill.BillPrinted = False Then GoTo NotThisOne
    If TaxBill.BillNumber > 0 Then
      If TaxBill.BillNumber >= FBill And TaxBill.BillNumber <= LBill Then
        Get TCHandle, TaxBill.CustRec, TaxCust
        CustName$ = QPTrim$(TaxCust.CustName)
          'Must Calc Late Fee Here
        TotOpt = OldRound(TaxBill.OptRevTax1 + TaxBill.OptRevTax2 + TaxBill.OptRevTax3)
          
          Print #RptFile, "                                R E A L   E S T A T E"
          Print #RptFile, "                                 T A X   N O T I C E"
          Print #RptFile, Tab(Tab1); TownName
          Print #RptFile, Tab(Tab2); Add1
          Print #RptFile, Tab(Tab3); Add3
          Print #RptFile,
          Print #RptFile, "            VALUATION AMOUNT: "; Using$("$##,###,###.00", OldRound#(TaxBill.RealValue + TaxBill.BldgValue));
          Print #RptFile, Tab(50); "ACCT. #: "; CStr(TaxBill.CustRec)
          Print #RptFile, "                   EXEMPTION: "; Using$("$##,###,###.00", TaxBill.ExptValue);
          Print #RptFile, Tab(50); "PIN. #: "; QPTrim$(TaxBill.RealPin)
          Print #RptFile, "         LATE PENALTY AMOUNT: "; Using$("$##,###,###.00", TaxBill.LateTaxDue);
          Print #RptFile, Tab(50); "RECPT #: "; Using$("#####0", TaxBill.BillNumber)
'          Print #RptFile, "              TAX AMOUNT DUE: "; Using$("$##,###,###.00", TaxBill.TotalBillDue);
          If TotOpt > 0 Then
            Print #RptFile, "              TAX AMOUNT DUE: "; Using$("$##,###,###.00", TaxBill.TotalBillDue) + "*";
          Else
            Print #RptFile, "              TAX AMOUNT DUE: "; Using$("$##,###,###.00", TaxBill.TotalBillDue);
          End If
          Print #RptFile, Tab(50); "TAX RATE %: "; Using$("#0.0000", RealTaxRate#)
          Print #RptFile, Tab(50); "TAX YEAR: "; CStr(TaxBill.TaxYear)
          Print #RptFile, Tab(50); "DUE DATE: "; DueDate$
'          Print #RptFile, Tab(11); Left$(QPTrim$(CustName$), 45)
          If TotOpt > 0 Then
            Print #RptFile, Tab(11); Left$(QPTrim$(CustName$), 45); Tab(50); "*PLUS " + QPTrim$(Using$("$##,##0.00", TotOpt)) + " IN ADDED TAXES."
          Else
            Print #RptFile, Tab(11); Left$(QPTrim$(CustName$), 45)
          End If
          Print #RptFile, Tab(11); Left$(QPTrim$(TaxBill.CustAdd1), 35)
          Print #RptFile, Tab(11); Left$(QPTrim$(TaxBill.CustAdd2), 35)
          Print #RptFile, Tab(11); QPTrim$(TaxBill.CustAdd3); " "; QPTrim$(TaxBill.CustZip)
          Print #RptFile,
          Print #RptFile,
          Print #RptFile, Tab(31); "T H A N K   Y O U ! ! !"
          Print #RptFile,
          Print #RptFile, "~"
        End If
      End If
NotThisOne:
    frmVATaxShowPctComp.ShowPctComp x, NumOfTBRecs
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      EnableCloseButton Me.hwnd, True
      cmdExit.Enabled = True
      cmdProcess.Enabled = True
      Exit Sub
    End If
  Next x
  Unload frmVATaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  cmdProcess.Enabled = True
  
  Close
  ViewPrint File$, "Real Property Tax Bills", True
  Exit Sub
  
LoadHeaders:
  WorkName = ""
  WordLen = Len(TownName)
  For y = 1 To WordLen
    ch = Mid(TownName, y, 1)
    WorkName = WorkName + ch + " "
  Next y
  TownName = WorkName
  
  WorkName = ""
  WordLen = Len(Add1)
  For y = 1 To WordLen
    ch = Mid(Add1, y, 1)
    WorkName = WorkName + ch + " "
  Next y
  Add1 = WorkName
  
  WorkName = ""
  WordLen = Len(Add2)
  For y = 1 To WordLen
    ch = Mid(Add2, y, 1)
    WorkName = WorkName + ch + " "
  Next y
  Add2 = WorkName
  
  WorkName = ""
  WordLen = Len(Add3)
  For y = 1 To WordLen
    ch = Mid(Add3, y, 1)
    WorkName = WorkName + ch + " "
  Next y
  Add3 = WorkName
  
  Return
  
End Sub
Private Sub PrintMdltwnPers()
  Dim TaxBill As VAPPTaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim BillInfo As VAPPTaxBillInfoType
  Dim BIHandle As Integer
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long, PersTaxRate#
  Dim PYear As Integer, PYearStr$
  Dim File$, LC As Integer
  Dim CustName$, WhatYear As Integer
  Dim RptFile#, WhatPers&
  Dim CarCount As Integer
  Dim TownName$, Add1$, Add2$, Add3$
  Dim PHandle As Integer, PPTRAVal#
  Dim NumOfPRecs As Long, PPTRADiscount#
  Dim PersRec As PersonalRecType
  Dim VehDesc$, PrnCnt As Integer
  Dim TaxAmt#, LCnt As Integer
  Dim TotOpt As Double
  Dim Tab1 As Integer, Tab2 As Integer, Tab3 As Integer, Tab4 As Integer
  Dim DueDate$, WorkName$
  Dim FBill&
  Dim LBill&
  Dim ZipRec As BillPrintPZipIdxType, CustArr As Long
  Dim ZHandle As Integer
  Dim NumOfZRecs As Long
  
  FBill = fpDblSnglPersFirstBill.Value
  LBill = fpDblSnglPersLastBill.Value
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  TaxMasterRec.MaxVehTaxVal = OldRound(TaxMasterRec.MaxVehTaxVal)
  TownName = QPTrim$(TaxMasterRec.Name)
  Add1$ = QPTrim$(TaxMasterRec.Add1)
  Add2$ = QPTrim$(TaxMasterRec.Add2)
  Add3$ = QPTrim$(TaxMasterRec.City) + ", " + QPTrim$(TaxMasterRec.TownState) + " " + QPTrim$(TaxMasterRec.Zip)
  
  OpenPersPropFile PHandle, NumOfPRecs
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenPersTaxBillFile TBHandle, NumOfTBRecs
  OpenPersBillInfoFile BIHandle
  Get BIHandle, 1, BillInfo
  Close BIHandle
  File$ = StartPath$ + "/TxBStandPP.PRN"
  RptFile# = FreeFile
  Open File$ For Output As #RptFile
  
  frmVATaxShowPctComp.Label1 = "Creating Personal Bills"
  frmVATaxShowPctComp.Show , Me
  frmVATaxShowPctComp.cmdCancel.Visible = False
  EnableCloseButton Me.hwnd, False
  cmdExit.Enabled = False
  cmdProcess.Enabled = False
  PersTaxRate# = fpDblSnglPersRate
  Tab1 = 40 - Len(TownName) / 2
  Tab2 = 40 - Len(Add1) / 2
  Tab3 = 40 - Len(Add3) / 2
  
  If Exist("PZipIdx.Dat") Then '12/6/06
    OpenPZipIdxFile ZHandle, NumOfZRecs
    NumOfTBRecs = NumOfZRecs
  End If
  
  For x = 1 To NumOfTBRecs
    If NumOfZRecs > 0 Then '12/6/06
      Get ZHandle, x, ZipRec
      CustArr = ZipRec.TaxBillRec
    Else
      CustArr = x '12/6/06
    End If
    Get TBHandle, CustArr, TaxBill
    If TaxBill.BillPrinted = False Then GoTo Natta
    If TaxBill.BillNumber > 0 Then
      If TaxBill.BillNumber >= FBill And TaxBill.BillNumber <= LBill Then
        Get TCHandle, TaxBill.CustRec, TaxCust
        WhatYear = TaxBill.TaxYear
        DueDate$ = MakeRegDate(TaxBill.DueDate)
        GoSub PrintIt
      End If
    End If
Natta:
    frmVATaxShowPctComp.ShowPctComp x, NumOfTBRecs
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      EnableCloseButton Me.hwnd, True
      cmdExit.Enabled = True
      cmdProcess.Enabled = True
      Exit Sub
    End If
  Next x
  Unload frmVATaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  cmdProcess.Enabled = True
  
  Close
  
  ViewPrint File$, "Personal Property Tax Bills", True
  Exit Sub

PrintIt:
  CustName$ = RTrim$(TaxCust.CustName)
  CustName$ = LTrim$(CustName$)
  TotOpt = OldRound(TaxBill.OptRevTax1 + TaxBill.OptRevTax2 + TaxBill.OptRevTax3)
  
  'Must Calc Late Fee Here
  
  Print #RptFile,
  Print #RptFile, Tab(Tab1); TownName
  Print #RptFile, Tab(Tab2); Add1
  Print #RptFile, Tab(Tab3); Add3
  Print #RptFile, Tab(27); "PERSONAL PROPERTY TAX BILL"
  Print #RptFile, Tab(30);
  For LC = 6 To 8
    Print #RptFile, " "
  Next
  Print #RptFile, Tab(10); "ACCT # "; Using$("######0", TaxBill.CustRec);
  Print #RptFile, Tab(63); "BILL # "; Using$("######0", TaxBill.BillNumber)
  Print #RptFile, Tab(10); Left$(CustName$, 25);
  Print #RptFile, Tab(63); "TAX YEAR: "; CStr(WhatYear)
  Print #RptFile, Tab(10); Left$(QPTrim$(TaxBill.CustAdd1), 25);
  Print #RptFile, Tab(63); "TAX RATE: "; Using("#0.##0", PersTaxRate#) + "%"
  Print #RptFile, Tab(10); Left$(QPTrim$(TaxBill.CustAdd2), 25)
  Print #RptFile, Tab(10); QPTrim$(TaxBill.CustAdd3); " "; QPTrim(TaxBill.CustZip)
  For LC = 14 To 17
    Print #RptFile, " "
  Next
  Print #RptFile, Tab(37); "PROPERTY"; Tab(51); "   TAX"; Tab(61); "   PPTRA"
  Print #RptFile, Tab(38); "  VALUE"; Tab(51); "AMOUNT"; Tab(61); "DISCOUNT"; Tab(71); "TOTAL DUE"
  'Line 23 Starts Here
  Print #RptFile, Tab(2); "Personal Property";
  Print #RptFile, Tab(38); Using$("###,##0", TaxBill.PersValue);
  Print #RptFile, Tab(47); Using$("###,##0.00", TaxBill.PersTaxDue);
  Print #RptFile, Tab(59); Using("###,##0.00", TaxBill.PPTRADiscnt);
  If TotOpt > 0 Then
    Print #RptFile, Tab(70); Using("###,##0.00", TaxBill.TotalBillDue) + "*"
  Else
    Print #RptFile, Tab(70); Using("###,##0.00", TaxBill.TotalBillDue)
  End If

  CarCount = 0
  Print #RptFile,
  Print #RptFile, " PPTRA Information"

  WhatPers& = TaxCust.FirstPersRec
  Do
    Get PHandle, WhatPers&, PersRec
    If PersRec.PPTRAYN = "Y" Then
      If OldRound#(PersRec.PersVal) > TaxMasterRec.MaxVehTaxVal Then
        PPTRAVal# = TaxMasterRec.MaxVehTaxVal
      Else
        PPTRAVal# = OldRound#(PersRec.PersVal)
      End If
      If PPTRAVal# <= TaxMasterRec.MinVehTaxVal Then
        PPTRADiscount# = OldRound#((OldRound#(PPTRAVal# / 100) * PersTaxRate#))
        PPTRADiscount# = OldRound(PPTRADiscount# / TaxMasterRec.MultiYear)
      Else
        PPTRADiscount# = OldRound#((OldRound#((PPTRAVal# / 100) * (TaxMasterRec.PPTRADisc / 100)) * PersTaxRate#))
        PPTRADiscount# = OldRound(PPTRADiscount# / TaxMasterRec.MultiYear)
      End If

      VehDesc$ = " VIN# " + QPTrim$(PersRec.Vin)
      VehDesc$ = QPTrim$(VehDesc$)
      TaxAmt# = (PersTaxRate# / 100) * PersRec.PersVal
      TaxAmt# = OldRound(TaxAmt# / TaxMasterRec.MultiYear)
      
      Print #RptFile, Tab(2); "*" + VehDesc$;
      Print #RptFile, Tab(38); Using("###,##0", PersRec.PersVal);
      Print #RptFile, Tab(47); Using("###,##0.00", TaxAmt#);
      Print #RptFile, Tab(59); Using("###,##0.00", PPTRADiscount#)
      CarCount = CarCount + 1
    End If

    WhatPers& = PersRec.NextRec
  Loop While WhatPers& > 0

  ' Finish the bill up here
  Print #RptFile, ""
  Print #RptFile, ""
  If TotOpt > 0 Then
    Print #RptFile, "*Includes " + QPTrim$(Using("$##,##0.00", TotOpt)) + " in added taxes."; Tab(40); "Total Tax Due by "; DueDate$;
    Print #RptFile, Tab(69); Using("$###,##0.00", TaxBill.TotalBillDue) + "*"
  Else
    Print #RptFile, Tab(40); "Total Tax Due by "; DueDate$;
    Print #RptFile, Tab(69); Using("$###,##0.00", TaxBill.TotalBillDue)
  End If
  Print #RptFile,
  Print #RptFile,
  If CarCount > 0 Then
    Print #RptFile, " The Personal Property Relief Act provides that the tax on the first"
    Print #RptFile, " " + Using$("$##,##0.00", TaxMasterRec.MaxVehTaxVal) + " of value of your personal car, motorcycle and pickup  "
    Print #RptFile, " or panel truck under 7,501 pounds, which is a qualifying vehicle, has been"
    Print #RptFile, " reduced by " + Using$("#0.00", TaxMasterRec.PPTRADisc) + "% this year. If your qualifying vehicle's value is"
    Print #RptFile, " " + Using$("$#,##0.00", TaxMasterRec.MinVehTaxVal) + " or less, your tax has been eliminated. These reductions are"
    Print #RptFile, " based on the local tax rates in effect on July 1 or August 1, 1997,"
    Print #RptFile, " whichever was higher. Please contact the Town Office with any questions."
    Print #RptFile, ""
    Print #RptFile, ""
  End If

  Print #RptFile,
  Print #RptFile,
  Print #RptFile,

  Return

End Sub

Private Sub PrintCdrBluffPers()
  Dim TaxBill As VAPPTaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim BillInfo As VAPPTaxBillInfoType
  Dim BIHandle As Integer
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long, PersTaxRate#
  Dim PYear As Integer, PYearStr$
  Dim File$, LC As Integer
  Dim CustName$, WhatYear As Integer
  Dim RptFile#, WhatPers&
  Dim CarCount As Integer
  Dim TownName$, Add1$, Add2$, Add3$
  Dim PHandle As Integer, PPTRAVal#
  Dim NumOfPRecs As Long, PPTRADiscount#
  Dim PersRec As PersonalRecType
  Dim VehDesc As String * 30, PrnCnt As Integer
  Dim TaxAmt#, LCnt As Integer
  Dim TotOth As Double, OptTot As Double
  Dim DueDate$
  Dim FBill&
  Dim LBill&
  Dim ZipRec As BillPrintPZipIdxType, CustArr As Long
  Dim ZHandle As Integer
  Dim NumOfZRecs As Long
  
  FBill = fpDblSnglPersFirstBill.Value
  LBill = fpDblSnglPersLastBill.Value
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  TaxMasterRec.MaxVehTaxVal = OldRound(TaxMasterRec.MaxVehTaxVal)
  TownName = QPTrim$(TaxMasterRec.Name)
  Add1$ = QPTrim$(TaxMasterRec.Add1)
  Add2$ = QPTrim$(TaxMasterRec.Add2)
  Add3$ = QPTrim$(TaxMasterRec.City) + ", " + QPTrim$(TaxMasterRec.TownState) + " " + QPTrim$(TaxMasterRec.Zip)
  
  OpenPersPropFile PHandle, NumOfPRecs
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenPersTaxBillFile TBHandle, NumOfTBRecs
  OpenPersBillInfoFile BIHandle
  Get BIHandle, 1, BillInfo
  Close BIHandle
  File$ = StartPath$ + "/TxCdrBluffPP.PRN"
  RptFile# = FreeFile
  Open File$ For Output As #RptFile
  
  frmVATaxShowPctComp.Label1 = "Creating Personal Bills"
  frmVATaxShowPctComp.Show , Me
  frmVATaxShowPctComp.cmdCancel.Visible = False
  EnableCloseButton Me.hwnd, False
  cmdExit.Enabled = False
  cmdProcess.Enabled = False
  PersTaxRate# = fpDblSnglPersRate
  
  If Exist("PZipIdx.Dat") Then '12/6/06
    OpenPZipIdxFile ZHandle, NumOfZRecs
    NumOfTBRecs = NumOfZRecs
  End If
  
  PrnCnt = fpDblSnglPersFirstBill.Value
  For x = 1 To NumOfTBRecs
    If NumOfZRecs > 0 Then '12/6/06
      Get ZHandle, x, ZipRec
      CustArr = ZipRec.TaxBillRec
    Else
      CustArr = x '12/6/06
    End If
    Get TBHandle, CustArr, TaxBill
    If TaxBill.BillNumber < 0 Then GoTo Natta
    If TaxBill.BillNumber >= FBill And TaxBill.BillNumber <= LBill Then
      Get TCHandle, TaxBill.CustRec, TaxCust
      WhatYear = TaxBill.TaxYear
      DueDate$ = MakeRegDate(TaxBill.DueDate)
      GoSub PrintIt
      PrnCnt = PrnCnt + 1
    End If
Natta:
    frmVATaxShowPctComp.ShowPctComp x, NumOfTBRecs
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      EnableCloseButton Me.hwnd, True
      cmdExit.Enabled = True
      cmdProcess.Enabled = True
      Exit Sub
    End If
  Next x
  Unload frmVATaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  cmdProcess.Enabled = True
  
  Close
  
  ViewPrint File$, "Personal Property Tax Bills", True
  Exit Sub

PrintIt:
   CustName$ = QPTrim$(TaxBill.CustName)
   Print #RptFile, "~"; Tab(34); Str$(WhatYear); " PERSONAL PROPERTY"
   Print #RptFile, Tab(5); TownName$
   Print #RptFile, Tab(5); Add1$
   Print #RptFile, Tab(5); Add3$
   Print #RptFile, Tab(5); "   "
   Call InsertSSNDashes(TaxCust.CSSN)
   Print #RptFile, Tab(10); TaxCust.CSSN; Tab(65); "PP"; Using("##.###", PersTaxRate#)
   Print #RptFile, " "
   Print #RptFile, " "

  'Line 30 to 35 here to print vehicles
    CarCount = 0
    WhatPers& = TaxCust.FirstPersRec
    OptTot = OldRound(TaxBill.OptRevTax1 + TaxBill.OptRevTax2 + TaxBill.OptRevTax3)
    If OptTot > 0 Then
      Print #RptFile, "Other Taxes:"; Tab(64); Using("##,###.##", OptTot#)
    End If
    Do
      Get PHandle, WhatPers&, PersRec
'      PYear$ = Right$(PersRec.Desc5, 4)
      PYear = PersRec.TaxBillYear
      If Left$(PersRec.Desc5, 1) = "Y" Or PersRec.PPTRAYN = "Y" Then 'added PersRec.PPTRAYN on 11/26/07
        If OldRound#(PersRec.PersVal) > TaxMasterRec.MaxVehTaxVal Then
          PPTRAVal# = TaxMasterRec.MaxVehTaxVal
        Else
          PPTRAVal# = OldRound#(PersRec.PersVal)
        End If
        If PPTRAVal# <= TaxMasterRec.MinVehTaxVal Then
          PPTRADiscount# = OldRound#((OldRound#(PPTRAVal# / 100) * PersTaxRate#))
        Else
          PPTRADiscount# = OldRound#((OldRound#((PPTRAVal# / 100) * (TaxMasterRec.PPTRADisc / 100)) * PersTaxRate#))
        End If
      Else
        PPTRADiscount# = 0
      End If
'      VehDesc$ = QPTrim$(PersRec.Desc4) + " " + Left$(PersRec.DESC2, 22) + "   " + Left$(PersRec.Desc5, 1)
'      VehDesc$ = QPTrim$(VehDesc$)
      VehDesc$ = CStr(PersRec.ModYear) + " " + QPTrim$(PersRec.MakeMod) + " " + QPTrim$(PersRec.DESC1)
      VehDesc = Left(VehDesc, 28) + " " + PersRec.PPTRAYN 'changed from above at Cedar Bluff's request 11/26/07
      CarCount = CarCount + 1
      If PersRec.PersVal <= 0 Then
        TaxAmt# = (MHTaxRate# / 100) * PersRec.MHValue
        Print #RptFile, VehDesc$;
        Print #RptFile, Tab(33); Using("##,###,###", PersRec.MHValue);
        Print #RptFile, Tab(44); Using("###,###.##", TaxAmt#);
        Print #RptFile, Tab(54); Using("##,###.##", PPTRADiscount#);
        Print #RptFile, Tab(64); Using("##,###.##", TaxAmt#)
      Else
        TaxAmt# = (PersTaxRate# / 100) * PersRec.PersVal
        Print #RptFile, VehDesc$;
        Print #RptFile, Tab(33); Using("##,###,###", PersRec.PersVal);
        Print #RptFile, Tab(44); Using("###,###.##", TaxAmt#);
        Print #RptFile, Tab(54); Using("#####.##", PPTRADiscount#);
        Print #RptFile, Tab(64); Using("##,###.##", OldRound#(TaxAmt# - PPTRADiscount#))
      End If
      If OptTot > 0 Then
        If CarCount = 4 Then CarCount = 5
      End If
      If (CarCount >= 5) And (PersRec.NextRec > 0) Then
        Print #RptFile, " "
        Print #RptFile, ""
        Print #RptFile, Tab(10); Using("#####", PrnCnt);
        Print #RptFile, Tab(36); DueDate$; Tab(66); "CONT'D"                 '; USING "$$#####,#.##"; TaxBill.TotalBillDue
        Print #RptFile,
        Print #RptFile, Tab(9); CustName$
        Print #RptFile, Tab(9); QPTrim$(TaxBill.CustAdd1)
        Print #RptFile, Tab(9); QPTrim$(TaxBill.CustAdd2)
        Print #RptFile, Tab(9); QPTrim$(TaxBill.CustAdd3); " "; TaxBill.CustZip
        Print #RptFile,
        Print #RptFile,
        Print #RptFile, "~"
        Print #RptFile, "~"
        Print #RptFile, Tab(5); TownName$
        Print #RptFile, Tab(5); Add1$
        Print #RptFile, Tab(5); Add2$
        Print #RptFile, Tab(5); Add3$;
        Print #RptFile, Tab(10); QPTrim$(TaxCust.CSSN)
        Print #RptFile, " "
        Print #RptFile, '"Vehicle Listing Continued ..."
        CarCount = 0
      ElseIf (CarCount >= 5) And (PersRec.NextRec <= 0) Then
        Print #RptFile, ""
      '  CarCount = 0
      End If
      WhatPers& = PersRec.NextRec
    Loop While WhatPers& > 0

  ' Finish the bill up here
    If CarCount < 5 Then
      For LCnt = CarCount To 5
        Print #RptFile, ""
      Next
    End If
'
    Print #RptFile, ""
    Print #RptFile, Tab(10); Using("#####", PrnCnt);
    Print #RptFile, Tab(36); DueDate$; Tab(62); Using("$###,###.##", TaxBill.TotalBillDue)
    Print #RptFile,
    Print #RptFile, Tab(9); CustName$
    Print #RptFile, Tab(9); TaxBill.CustAdd1
    Print #RptFile, Tab(9); TaxBill.CustAdd2
    Print #RptFile, Tab(9); QPTrim$(TaxBill.CustAdd3); " "; QPTrim$(TaxBill.CustZip)
    Print #RptFile,
    Print #RptFile,
    Print #RptFile, "~"
  Return

End Sub
Private Sub PrintCdrBluffReal()
  Dim TaxBill As VARETaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim BillInfo As VARETaxBillInfoType
  Dim BIHandle As Integer
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long, RealTaxRate#
  Dim File$, WordLen As Integer
  Dim CustName As String * 45
  Dim RptFile#, ch$, y As Integer
  Dim TownName$, Add1$, Add2$, Add3$
  Dim RHandle As Integer, PrnCnt As Integer
  Dim NumOfRRecs As Long, WhatYear As Integer
  Dim RealRec As PropertyRecType
  Dim TaxAmt#, Tab1 As Integer, Tab2 As Integer, Tab3 As Integer
  Dim DueDate$, WorkName$
  Dim FBill&
  Dim LBill&, OptTot As Double
  Dim ZipRec As BillPrintRZipIdxType, CustArr As Long
  Dim ZHandle As Integer
  Dim NumOfZRecs As Long
  Dim MortRec As BillPrintMortIdxType
  Dim MRHandle As Integer
  Dim NumOfMRRecs As Long
  
  FBill = fpDblSnglRealFirstBill.Value
  LBill = fpDblSnglRealLastBill.Value
  
  DueDate$ = fptxtRealDueDate
  RealTaxRate# = fpDblSnglRealRate
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  TownName = QPTrim$(TaxMasterRec.Name)
  Add1$ = QPTrim$(TaxMasterRec.Add1)
  Add2$ = QPTrim$(TaxMasterRec.Add2)
  Add3$ = QPTrim$(TaxMasterRec.City) + ", " + QPTrim$(TaxMasterRec.TownState) + " " + QPTrim$(TaxMasterRec.Zip)

  OpenRealPropFile RHandle, NumOfRRecs
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenRealTaxBillFile TBHandle, NumOfTBRecs
  OpenRealBillInfoFile BIHandle
  Get BIHandle, 1, BillInfo
  Close BIHandle
  File$ = StartPath$ + "/TxBCdrBluffRE.PRN"
  RptFile# = FreeFile
  Open File$ For Output As #RptFile

'  OpenTaxCustFile TCHandle, NumOfTCRecs
  'Must Calc Late Fee Here
  frmVATaxShowPctComp.Label1 = "Creating Real Bills"
  frmVATaxShowPctComp.Show , Me
  frmVATaxShowPctComp.cmdCancel.Visible = False
  EnableCloseButton Me.hwnd, False
  cmdExit.Enabled = False
  cmdProcess.Enabled = False

  If Exist("MORTIDX.DAT") Then '12/6/06
    OpenMortIdxFile MRHandle, NumOfMRRecs
    NumOfTBRecs = NumOfMRRecs
  ElseIf Exist("RZIPIDX.DAT") Then '12/6/06
    OpenRZipIdxFile ZHandle, NumOfZRecs
    NumOfTBRecs = NumOfZRecs
  End If
  
  For x = 1 To NumOfTBRecs
    If NumOfMRRecs > 0 Then '12/6/06
      Get MRHandle, x, MortRec
      CustArr = MortRec.TaxBillRec
    ElseIf NumOfZRecs > 0 Then '12/6/06
      Get ZHandle, x, ZipRec
      CustArr = ZipRec.TaxBillRec
    Else
      CustArr = x '12/6/06
    End If
    Get TBHandle, CustArr, TaxBill
    Get TCHandle, TaxBill.CustRec, TaxCust
    If TaxBill.BillNumber < 0 Then GoTo NotThisOne
    If TaxBill.BillNumber >= FBill And TaxBill.BillNumber <= LBill Then
      CustName$ = QPTrim$(TaxBill.CustName)
      WhatYear = TaxBill.TaxYear
  'Must Calc Late Fee Here
      OptTot = OldRound(TaxBill.OptRevTax1 + TaxBill.OptRevTax2 + TaxBill.OptRevTax3)
      Print #RptFile, "~"
      Print #RptFile, Tab(50); CStr(WhatYear); Tab(78); Using("########", TaxBill.BillNumber)
      Print #RptFile,
      Print #RptFile, " "
      Print #RptFile, " "
      Print #RptFile, " "
      Print #RptFile, Tab(28); Using("#.##", RealTaxRate#);
      Print #RptFile, Tab(36); Using("###,###,###", TaxBill.RealValue);
      Print #RptFile, Tab(48); Using("##,###,###", TaxBill.BldgValue);
      Print #RptFile, Tab(61); Using("##,###,###", TaxBill.RealValue + TaxBill.BldgValue);
      Print #RptFile, Tab(75); Using("###,###.##", TaxBill.RealTaxDue);
      Print #RptFile, Tab(90); QPTrim$(TaxBill.Comment) + "%"
      Print #RptFile, " "
      Print #RptFile, Tab(68); DueDate$
      Print #RptFile, QPTrim$(TaxBill.RDesc1)
      Print #RptFile, QPTrim$(TaxBill.RDesc2)
      Print #RptFile, ""
      Print #RptFile, ""
      Print #RptFile, ""
      Print #RptFile, Tab(7); "ACCT # "; CStr(TaxBill.CustRec)
      Print #RptFile, Tab(7); Left$(CustName$, 45)
      Print #RptFile, Tab(7); Left$(TaxBill.CustAdd1, 35)
      Print #RptFile, Tab(7); Left$(TaxBill.CustAdd2, 35)
      Print #RptFile, Tab(7); QPTrim$(TaxBill.CustAdd3); " "; TaxBill.CustZip
      Print #RptFile,
'      Print #RptFile, "BN"; Using("#####", PrnCnt)
      If OptTot = 0 Then
        Print #RptFile, "BN"; Using("#####", PrnCnt)
      Else
        Print #RptFile, "BN"; Using("#####", PrnCnt); Tab(54); "Tax Due includes " + QPTrim$(Using$("$##,##0.00", OptTot)) + " in other taxes."
      End If
        Print #RptFile, "~"
      End If
NotThisOne:
    frmVATaxShowPctComp.ShowPctComp x, NumOfTBRecs
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      EnableCloseButton Me.hwnd, True
      cmdExit.Enabled = True
      cmdProcess.Enabled = True
      Exit Sub
    End If
  Next x
  Unload frmVATaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  cmdProcess.Enabled = True

  Close
  
  ViewPrint File$, "Real Property Tax Bills", True

End Sub

Private Sub PrintHalifaxStandardPersonal(RptFile As Integer, TBHandle As Integer, PTaxBill As VAPPTaxBillType, TCHandle As Integer, TaxCust As TaxCustType, PrnCnt As Long, PersRec As PersonalRecType, PHandle As Integer)
 'TAXPPSTD.BI
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim PYear As Integer, PYearStr$
  Dim File$, LC As Integer
  Dim CustName$, WhatYear As Integer
  Dim WhatPers&
  Dim CarCount As Integer
  Dim TownName$, Add1$, Add2$, Add3$
  Dim PPTRAVal#
  Dim PPTRADiscount#
  Dim VehDesc$, PERC!
  Dim TaxAmt#, LCnt As Integer
  Dim MultiYear As Integer
  Dim TotOth As Double
  Dim PrintComments As String
  
  WhatYear = PersYear
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  MultiYear = TaxMasterRec.MultiYear
  TaxMasterRec.MaxVehTaxVal = OldRound(TaxMasterRec.MaxVehTaxVal)
  TownName = QPTrim$(TaxMasterRec.Name)
  Add1$ = QPTrim$(TaxMasterRec.Add1)
  Add2$ = QPTrim$(TaxMasterRec.Add2)
  Add3$ = QPTrim$(TaxMasterRec.City) + ", " + QPTrim$(TaxMasterRec.TownState) + " " + QPTrim$(TaxMasterRec.Zip)
  
  If WhatYear = 1999 Then PERC! = 27.5
  If WhatYear = 2000 Then PERC! = 47.5
  If WhatYear >= 2001 Then PERC! = TaxMasterRec.PPTRADisc
  
  CustName$ = QPTrim$(TaxCust.CustName)
  Print #RptFile, "~"
  Print #RptFile, Tab(63); "TAX YEAR: "; WhatYear
  Print #RptFile, Tab(75); Using$("#####", PTaxBill.BillNumber)
  Print #RptFile, " "
  Print #RptFile, " "
  Print #RptFile, Tab(5); " "
  Print #RptFile, Tab(5); " "
  Print #RptFile, Tab(5); " "
  Print #RptFile, Tab(5); " "
  Print #RptFile, " "
  Print #RptFile, " "
  Print #RptFile, " "
  Print #RptFile, " "
  Print #RptFile, Tab(5); "Acct # "; Using$("#####0", PTaxBill.CustRec)
  Print #RptFile, Tab(5); CustName$
  Print #RptFile, Tab(5); QPTrim$(PTaxBill.CustAdd1)
  Print #RptFile, Tab(5); QPTrim$(PTaxBill.CustAdd2)
  Print #RptFile, Tab(5); QPTrim$(PTaxBill.CustAdd3) + " " + QPTrim$(PTaxBill.CustZip)
'  For LC = 18 To 21
  For LC = 19 To 21 'added
   Print #RptFile, " "
  Next LC
  Print #RptFile, Tab(37); "PROP VALUE"; Tab(50); "TAX AMOUNT"; Tab(62); "PPTRA DIS"; Tab(72); "TOTAL DUE"
  Print #RptFile, " "
 'Line 24 Starts Here
  Print #RptFile, "Personal Property"; Tab(32); Using$("#.00", PersTaxRate#);
   Print #RptFile, Tab(37); Using$("#####0.00", PTaxBill.PersValue);
   Print #RptFile, Tab(51); Using$("#####0.00", PTaxBill.PersTaxDue); ' - PTaxBill.OverPayAmt);
   Print #RptFile, Tab(63); Using$("####0.00", PTaxBill.PPTRADiscnt);
   Print #RptFile, Tab(72); Using$("#####0.00", OldRound(PTaxBill.PersTaxDue - PTaxBill.PPTRADiscnt)) ' - PTaxBill.OverPayAmt))
   
  Print #RptFile, "Machinery/Tools"; Tab(32); Using$("#.00", MTTaxRate#);
   Print #RptFile, Tab(37); Using$("#####0.00", PTaxBill.MTValue);
   Print #RptFile, Tab(51); Using$("#####0.00", PTaxBill.MTTaxDue);
   Print #RptFile, Tab(72); Using$("#####0.00", PTaxBill.MTTaxDue)
  Print #RptFile, "Farm Equipment";
   Print #RptFile, Tab(32); Using("#.00", FETaxRate#);
   Print #RptFile, Tab(37); Using$("#####0.00", PTaxBill.FEValue);
   Print #RptFile, Tab(51); Using$("#####0.00", PTaxBill.FETaxDue);
   Print #RptFile, Tab(72); Using$("#####0.00", PTaxBill.FETaxDue)
  Print #RptFile, "Mobile Homes";
   Print #RptFile, Tab(32); Using$("#.00", MHTaxRate#);
   Print #RptFile, Tab(37); Using$("#####0.00", PTaxBill.MHValue);
   Print #RptFile, Tab(51); Using$("#####0.00", PTaxBill.MHTaxDue);
   Print #RptFile, Tab(72); Using$("#####0.00", PTaxBill.MHTaxDue)
  Print #RptFile, "Merchant Capital";
   Print #RptFile, Tab(32); Using$("#.00", MCTaxRate#);
   Print #RptFile, Tab(37); Using$("#####0.00", PTaxBill.MCValue);
   Print #RptFile, Tab(51); Using$("#####0.00", PTaxBill.MCTaxDue);
   Print #RptFile, Tab(72); Using$("#####0.00", PTaxBill.MCTaxDue)
   TotOth = OldRound(PTaxBill.OptRevTax1 + PTaxBill.OptRevTax2 + PTaxBill.OptRevTax3)
   If PTaxBill.OverPayAmt > 0 And TotOth = 0 Then '6/22/06
     Print #RptFile, " PPTRA Vehicle Information"; Tab(43); "** Applied Credit Of: " + QPTrim$(Using$("$##,##0.00", PTaxBill.OverPayAmt)) + " **"
   ElseIf PTaxBill.OverPayAmt > 0 And TotOth > 0 Then '6/22/06
     Print #RptFile, " PPTRA Vehicle Information"; Tab(30); "* Applied Credit Of: " + QPTrim$(Using$("$##,##0.00", PTaxBill.OverPayAmt)) + " *"; Tab(60); "Other Taxes: " + Using$("#,##0.00", TotOth)
  ElseIf PTaxBill.OverPayAmt = 0 And TotOth > 0 Then '6/22/06
     Print #RptFile, " PPTRA Vehicle Information"; Tab(60); "Other Taxes: " + Using$("#,##0.00", TotOth)
   Else
     Print #RptFile, " PPTRA Vehicle Information"
   End If
 'Line 30 to 35 here to print vehicles
  CarCount = 0
  WhatPers& = TaxCust.FirstPersRec
  Do
    Get PHandle, WhatPers&, PersRec
    PYearStr$ = CStr(PersRec.TaxBillYear)
    PYear = Val(PYearStr$)
    If PYear > 0 And PYear <> WhatYear Then
      Return
    End If
    If PersRec.PPTRAYN = "Y" Then
      If OldRound#(PersRec.PersVal) > TaxMasterRec.MaxVehTaxVal Then
        If TaxMasterRec.MultiYear <> 0 Then
          PersRec.PersVal = OldRound(PersRec.PersVal)
        End If
        PPTRAVal# = TaxMasterRec.MaxVehTaxVal
      Else
        PPTRAVal# = OldRound#(PersRec.PersVal)
      End If

      If PPTRAVal# <= TaxMasterRec.MinVehTaxVal Then
        PPTRADiscount# = OldRound#((OldRound#(PPTRAVal# / 100) * PersTaxRate#))
      Else
        PPTRADiscount# = OldRound#((OldRound#((PPTRAVal# / 100) * (PERC! / 100)) * PersTaxRate#))
      End If
      VehDesc$ = " VIN# " + QPTrim$(PersRec.Vin)
      VehDesc$ = QPTrim$(VehDesc$)
      TaxAmt# = OldRound((PersTaxRate# / 100) * PersRec.PersVal)
      PTaxBill.PersTaxDue = PTaxBill.PersTaxDue
      PTaxBill.PPTRADiscnt = PTaxBill.PPTRADiscnt
      Print #RptFile, "*" + VehDesc$;
      Print #RptFile, Tab(37); Using$("#####0.00", PersRec.PersVal) ';
      CarCount = CarCount + 1
    End If
    
    If CarCount >= 6 Then
      Print #RptFile, ""
      Print #RptFile, Tab(48); "Total Tax Due ";
      Print #RptFile, Using$("$#######0.00", OldRound(PTaxBill.TotalBillDue - PTaxBill.OverPayAmt))
      Print #RptFile, Tab(48); "Tax Due Date: " + fptxtPersDueDate.Text  'DueDate$
      Print #RptFile, ""
      Print #RptFile,
      Print #RptFile,
      Print #RptFile,
      Print #RptFile, "BN"; Using("####0", PrnCnt) 'x
      Print #RptFile, "~"

      Print #RptFile, "~"
      Print #RptFile, Tab(62); "TAX YEAR: "; WhatYear
      Print #RptFile, Tab(75); Using$("####0", PTaxBill.BillNumber)
      Print #RptFile, " "
      Print #RptFile, " "
      Print #RptFile, Tab(5); TownName$
      Print #RptFile, Tab(5); Add1$
      Print #RptFile, Tab(5); Add2$
      Print #RptFile, Tab(5); Add3$
      Print #RptFile, " "
      Print #RptFile, " "
      Print #RptFile, " "
      Print #RptFile, Tab(5); "Acct # " + Using$("####0", PTaxBill.CustRec) + " Vehicle Listing Cont'd"
      Print #RptFile, Tab(5); CustName$
      Print #RptFile, Tab(5); QPTrim$(PTaxBill.CustAdd1)
      Print #RptFile, Tab(5); QPTrim$(PTaxBill.CustAdd2)
      Print #RptFile, Tab(5); QPTrim$(PTaxBill.CustAdd3) + " " + QPTrim$(PTaxBill.CustZip)
      For LC = 18 To 21
       Print #RptFile, " "
      Next LC
      Print #RptFile, Tab(37); "PROP VALUE"; Tab(50); "TAX AMOUNT"; Tab(62); "PPTRA DIS."; Tab(73); "TOTAL DUE"
      Print #RptFile, " "
      Print #RptFile, "Vehicle Listing Continued ..."
      Print #RptFile, ""
      Print #RptFile, ""
      Print #RptFile, " PPTRA Vehicle Information"
      Print #RptFile, ""
      Print #RptFile, ""
      CarCount = 0
    End If
    WhatPers& = PersRec.NextRec
  Loop While WhatPers& > 0

   ' Finish the bill up here
  If CarCount < 6 Then
    For LCnt = CarCount To 6: Print #RptFile, "": Next LCnt
  End If
  If InStr(PTaxBill.CommentPlace, "LEFT") Then
    Print #RptFile,
    Print #RptFile,
    Print #RptFile, PTaxBill.Comment; Tab(48); "Total Tax Due ";
    Print #RptFile, Using$("$#######0.00", OldRound(PTaxBill.TotalBillDue - PTaxBill.OverPayAmt))
    Print #RptFile, PTaxBill.Comment2; Tab(48); "Tax Due Date: " + fptxtPersDueDate.Text
  ElseIf InStr(PTaxBill.CommentPlace, "RIGHT") Then
    Print #RptFile, Tab(48); "Total Tax Due ";
    Print #RptFile, Using$("$#######0.00", OldRound(PTaxBill.TotalBillDue - PTaxBill.OverPayAmt))
    Print #RptFile, Tab(48); "Tax Due Date: " + fptxtPersDueDate.Text
    Print #RptFile, Tab(48); PTaxBill.Comment
    Print #RptFile, Tab(48); PTaxBill.Comment2
  Else
    Print #RptFile,
    Print #RptFile,
    Print #RptFile, Tab(48); "Total Tax Due ";
    Print #RptFile, Using$("$#######0.00", OldRound(PTaxBill.TotalBillDue - PTaxBill.OverPayAmt))
    Print #RptFile, Tab(48); "Tax Due Date: " + fptxtPersDueDate.Text
  End If
  Print #RptFile, " " 'added
  Print #RptFile, " " 'added
  Print #RptFile, "BN"; Using$("####0", PrnCnt) 'x
  Print #RptFile, "~"
  
  Exit Sub
  

End Sub

Private Sub PrintHalifaxStandardReal(RptFile As Integer, TBHandle As Integer, RTaxBill As VARETaxBillType, PrnCnt As Long) 'TCHandle As Integer, TaxCust As TaxCustType, PrnCnt As Long, PersRec As PersonalRecType, PHandle As Integer)
 'checked OK against mask (TAXREMSK.DAT) on 10/21/2005
 'STANDARD REAL ESTATE BILL FORMAT AS SOLD BY SOUTHERN SOFTWARE
 'TAXRESTD.BI
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
'  Dim TaxCust As TaxCustType
'  Dim TCHandle As Integer
'  Dim NumOfTCRecs As Long
  Dim RealTaxRate#
  Dim RYear As Integer, RYearStr$
  Dim LC As Integer
  Dim CustName As String * 45, WhatYear As Integer
  Dim WhatReal&
  Dim TownName$, Add1$, Add2$, Add3$
  Dim TaxAmt#, LCnt As Integer
  Dim ThisDesc As String * 28
  Dim TotOth As Double
  
  RealTaxRate# = fpDblSnglRealRate
  WhatYear = RealYear
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  'Must Calc Late Fee Here
  Print #RptFile, "~"
  Print #RptFile, Tab(64); "TAX YEAR: "; WhatYear
  Print #RptFile, Tab(75); Using$("#####", RTaxBill.BillNumber)
  Print #RptFile, " "
  Print #RptFile, " "
  Print #RptFile, " "
  Print #RptFile, " "
  Print #RptFile, " "
  Print #RptFile, " "
  Print #RptFile, " "
  Print #RptFile, " "
  Print #RptFile, " "
  Print #RptFile, " "
  Print #RptFile, Tab(5); "PIN:  " + QPTrim$(RTaxBill.RealPin)
  Print #RptFile, Tab(5); "ACCT: " + Using$("#####", RTaxBill.CustRec)
  Print #RptFile, Tab(5); QPTrim$(RTaxBill.CustName)
  Print #RptFile, Tab(5); Left$(RTaxBill.CustAdd1, 35)
  Print #RptFile, Tab(5); Left$(RTaxBill.CustAdd2, 35)
  Print #RptFile, Tab(5); QPTrim$(RTaxBill.CustAdd3) + " " + RTaxBill.CustZip

  For LC = 19 To 20 'made 18 = 19
    Print #RptFile, " "
  Next LC
  Print #RptFile, Tab(41); "LAND"; Tab(51); "BUILDING"; Tab(61); "NET TOTAL"; Tab(72); "TOTAL DUE"
  Print #RptFile, " "
  'Line 23 Starts Here
  ThisDesc = QPTrim$(RTaxBill.RDesc1)
  Print #RptFile, ThisDesc; 'QPTrim$(RTaxBill.RDesc1);
  Print #RptFile, Tab(30); Using("#0.00", RealTaxRate#);
  If RTaxBill.RealValue > RTaxBill.ExptValue Then
    Print #RptFile, Tab(37); Using("######0.00", (RTaxBill.RealValue - RTaxBill.ExptValue)); ' - RRTaxBill.PersValue));
    Print #RptFile, Tab(50); Using("#####0.00", RTaxBill.BldgValue);
  ElseIf RTaxBill.BldgValue > RTaxBill.ExptValue Then
    Print #RptFile, Tab(37); Using("######0.00", RTaxBill.RealValue); ' - RRTaxBill.PersValue));
    Print #RptFile, Tab(50); Using("#####0.00", (RTaxBill.BldgValue - RTaxBill.ExptValue));
  ElseIf RTaxBill.RealValue + RTaxBill.BldgValue > RTaxBill.ExptValue Then
    Print #RptFile, Tab(37); Using("######0.00", RTaxBill.RealValue - (RTaxBill.ExptValue * (RTaxBill.RealValue / (RTaxBill.RealValue + RTaxBill.BldgValue)))); ' - RRTaxBill.PersValue));
    Print #RptFile, Tab(50); Using("#####0.00", RTaxBill.BldgValue - (RTaxBill.ExptValue * (RTaxBill.BldgValue / (RTaxBill.RealValue + RTaxBill.BldgValue)))); ' - RRTaxBill.PersValue));
  Else
    Print #RptFile, Tab(37); Using("######0.00", RTaxBill.RealValue); ' - RRTaxBill.PersValue));
    Print #RptFile, Tab(50); Using("#####0.00", RTaxBill.BldgValue);
  End If
  Print #RptFile, Tab(61); Using("#####0.00", OldRound(RTaxBill.RealValue + RTaxBill.BldgValue - RTaxBill.ExptValue));
  Print #RptFile, Tab(71); Using("######0.00", OldRound(RTaxBill.TotalBillDue)) ' - RTaxBill.OverPayAmt))
  Print #RptFile, QPTrim$(RTaxBill.RDesc2)
  TotOth = OldRound(RTaxBill.OptRevTax1 + RTaxBill.OptRevTax2 + RTaxBill.OptRevTax3 + RTaxBill.LateTaxDue)
  If RTaxBill.OverPayAmt > 0 And TotOth > 0 Then
    Print #RptFile, "** Applied Credit Of " + QPTrim$(Using$("$##,##0.00", RTaxBill.OverPayAmt)) + " **"; Tab(40); "Other Taxes Included: " + Using$("$#,##0.00", TotOth)
    For LCnt = 26 To 36: Print #RptFile, "": Next LCnt
  ElseIf RTaxBill.OverPayAmt > 0 And TotOth = 0 Then
    Print #RptFile, "** Applied Credit Of " + QPTrim$(Using$("$##,##0.00", RTaxBill.OverPayAmt)) + " **"
    For LCnt = 26 To 36: Print #RptFile, "": Next LCnt
  ElseIf RTaxBill.OverPayAmt = 0 And TotOth > 0 Then
    Print #RptFile, Tab(40); "Other Taxes Included: " + Using$("$#,##0.00", TotOth)
    For LCnt = 26 To 36: Print #RptFile, "": Next LCnt
  Else
    For LCnt = 25 To 36: Print #RptFile, "": Next LCnt
  End If
 'Lines 25 to 36 are blank
'Line 37 for Totals
  Print #RptFile, Tab(48); "Total Tax Due ";
  Print #RptFile, Using$("$#######0.00", OldRound(RTaxBill.TotalBillDue - RTaxBill.OverPayAmt))
  Print #RptFile, Tab(48); "Tax Due Date: " + MakeRegDate(RTaxBill.DueDate)
  Print #RptFile,
  Print #RptFile,
  Print #RptFile,
  Print #RptFile, "BN"; Using$("#####", PrnCnt)
  Print #RptFile, "~"
  
End Sub

