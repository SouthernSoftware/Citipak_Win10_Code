VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmTaxBillReprinting 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Bill Reprinting"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmTaxBillReprinting.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcmbPrintPriorYN 
      Height          =   390
      Left            =   8520
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   5400
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
      ColDesigner     =   "frmTaxBillReprinting.frx":08CA
   End
   Begin EditLib.fpText fptxtOrder 
      Height          =   372
      Left            =   5520
      TabIndex        =   2
      Top             =   6832
      Width           =   2652
      _Version        =   196608
      _ExtentX        =   4678
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
   Begin EditLib.fpDoubleSingle fpDblSnglStartBill 
      Height          =   372
      Left            =   5880
      TabIndex        =   3
      Top             =   3832
      Width           =   1572
      _Version        =   196608
      _ExtentX        =   2773
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
   Begin EditLib.fpLongInteger fpLongTaxYear 
      Height          =   372
      Left            =   5880
      TabIndex        =   4
      Top             =   3232
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
      Left            =   6360
      TabIndex        =   5
      Top             =   4792
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
   Begin EditLib.fpDoubleSingle fpDblSnglPersRate 
      Height          =   372
      Left            =   6360
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5392
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
   Begin EditLib.fpDoubleSingle fpDblSnglLateList 
      Height          =   372
      Left            =   6000
      TabIndex        =   7
      Top             =   6232
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
   Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
      Height          =   492
      Left            =   7188
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   7776
      Width           =   2064
      _Version        =   131072
      _ExtentX        =   3641
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
      ButtonDesigner  =   "frmTaxBillReprinting.frx":0CA5
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   492
      Left            =   2388
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   7788
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
      ButtonDesigner  =   "frmTaxBillReprinting.frx":0E84
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdAlign 
      Height          =   492
      Left            =   4800
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   7788
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
      ButtonDesigner  =   "frmTaxBillReprinting.frx":1060
   End
   Begin EditLib.fpDoubleSingle fpDblSnglLastBill 
      Height          =   372
      Left            =   5520
      TabIndex        =   1
      Top             =   2524
      Width           =   1572
      _Version        =   196608
      _ExtentX        =   2773
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
   Begin EditLib.fpDoubleSingle fpDblSnglFirstBill 
      Height          =   372
      Left            =   5520
      TabIndex        =   0
      Top             =   2044
      Width           =   1572
      _Version        =   196608
      _ExtentX        =   2773
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
   Begin fpBtnAtlLibCtl.fpBtn cmdList 
      Height          =   372
      Left            =   7560
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   2284
      Width           =   2052
      _Version        =   131072
      _ExtentX        =   3619
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
      ButtonDesigner  =   "frmTaxBillReprinting.frx":1241
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Print Prior Year Balance (Y/N?):"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1020
      Left            =   8160
      TabIndex        =   23
      Top             =   4440
      Width           =   1665
      WordWrap        =   -1  'True
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   1092
      Left            =   4320
      Top             =   1924
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
      Left            =   4440
      TabIndex        =   20
      Top             =   2116
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
      Left            =   4440
      TabIndex        =   19
      Top             =   2596
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
      Left            =   6480
      TabIndex        =   18
      Top             =   4504
      Width           =   1092
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Late List Pct:"
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
      Left            =   4440
      TabIndex        =   17
      Top             =   6304
      Width           =   1452
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Personal Property:"
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
      Left            =   3960
      TabIndex        =   16
      Top             =   5512
      Width           =   2172
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
      Left            =   4320
      TabIndex        =   15
      Top             =   4912
      Width           =   1692
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
      Left            =   4080
      TabIndex        =   14
      Top             =   3904
      Width           =   1692
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
      Height          =   252
      Left            =   4680
      TabIndex        =   13
      Top             =   3304
      Width           =   1092
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   660
      Index           =   1
      Left            =   1488
      Top             =   568
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
      TabIndex        =   12
      Top             =   736
      Width           =   5292
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   5775
      Left            =   1605
      Top             =   1665
      Width           =   8415
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   1572
      Left            =   3588
      Top             =   4432
      Width           =   4452
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Printing Order:"
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
      Left            =   3600
      TabIndex        =   11
      Top             =   6952
      Width           =   1692
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   780
      Left            =   1488
      Top             =   460
      Width           =   8652
   End
End
Attribute VB_Name = "frmTaxBillReprinting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class
  Dim BillFormat As Long
  Dim FirstNum As Long
  Dim SecondNum As Long
  Dim BillCnt As Long
  Dim Opt1Desc$
  Dim Opt2Desc$
  Dim Opt3Desc$
  Dim TownName As String

Private Sub cmdAlign_Click()
  Dim Handle As Integer
  Dim TempHandle As Integer
  Dim Cnt As Integer
  Dim TextLine$
  
  'on error goto ERRORSTUFF
  
  If BillFormat = 20304 Then
    If Exist("TAXMSKPC1.DAT") Then
      alnRpt = "TAXMSKPC1.DAT"
    Else
      Call TaxMsg(900, "The mask for this bill format, " + "TAXMSKPC1.DAT," + " could not be found.")
      Close
      Exit Sub
    End If
  ElseIf BillFormat = 20002 Then
    If Exist("TXMSKHMLT24TF.DAT") Then
      Call TaxMsg(900, "12 Pitch is recommended for this form. Each mask prints 2 bills to match the way the bill forms have been printed.")
      alnRpt = "TXMSKHMLT24TF.DAT"
    Else
      Call TaxMsg(900, "The mask for this bill format, " + "TXMSKHMLT24TF.DAT," + " could not be found.")
      Close
      Exit Sub
    End If
  ElseIf BillFormat = 20003 Then
    If Exist("TXMSKPH24TF.DAT") Then
      Call TaxMsg(900, "10 Pitch is recommended for this form.")
      alnRpt = "TXMSKPH24TF.DAT"
    Else
      Call TaxMsg(900, "The mask for this bill format, " + "TXMSKPH24TF.DAT," + " could not be found.")
      Close
      Exit Sub
    End If
  ElseIf BillFormat = 20004 Then
    If Exist("TXMSKSYL23TF.DAT") Then
      Call TaxMsg(900, "12 Pitch is recommended for this form.")
      alnRpt = "TXMSKSYL23TF.DAT"
    Else
      Call TaxMsg(900, "The mask for this bill format, " + "TXMSKSYL23TF.DAT," + " could not be found.")
      Close
      Exit Sub
    End If
  ElseIf BillFormat = 20005 Then
    If Exist("TXMSKBSC32TF.DAT") Then
      Call TaxMsg(900, "10 Pitch is recommended for this form.")
      alnRpt = "TXMSKBSC32TF.DAT"
    Else
      Call TaxMsg(900, "The mask for this bill format, " + "TXMSKBSC32TF.DAT," + " could not be found.")
      Close
      Exit Sub
    End If
  ElseIf BillFormat = 20006 Then
    If Exist("TXMSKLLN21TF.DAT") Then
      Call TaxMsg(900, "10 Pitch is recommended for this form.")
      alnRpt = "TXMSKLLN21TF.DAT"
    Else
      Call TaxMsg(900, "The mask for this bill format, " + "TXMSKLLN21TF.DAT," + " could not be found.")
      Close
      Exit Sub
    End If
  ElseIf BillFormat = 21837 Then
    If Exist("TAXBLMSK.DAT") Then
      alnRpt = "TAXBLMSK.DAT"
    Else
      Call TaxMsg(900, "The mask for this bill format, " + "TAXBLMSK.DAT," + " could not be found.")
      Close
      Exit Sub
    End If
  ElseIf BillFormat = 20007 Then
    Call TaxMsg(900, "No alignment is necessary for the legal size laser format.")
    Close
    Exit Sub
  Else
    Call TaxMsg(900, "The mask for this bill format could not be found.")
    Close
    Exit Sub
  End If
  
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
  frmTaxPrint.Show vbModal
  alnRpt = ""
  doAlign = False
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxBillReprinting", "cmdAlign_Click", Erl)
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
  frmTaxBillPrintMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdList_Click()
  frmTaxPrintedBillsList.Show vbModal
  DoEvents
End Sub

Private Sub cmdProcess_Click()
  Dim RptHandle As Integer
  Dim TaxBill As TaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim x As Long, BillNo&
  Dim WhatRec&, PrnCnt&
  Dim PINTemp As String * 18
  Dim CustName$
  Dim RptFile$, FBill&, LBill&
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim FF$
  Dim ZipRec As BillPrintZipIdxType
  Dim ZHandle As Integer
  Dim NumOfZRecs As Long
  Dim MortRec As BillPrintMortIdxType
  Dim MRHandle As Integer
  Dim NumOfMRRecs As Long
  
  'on error goto ERRORSTUFF
  
  FF$ = Chr(12)
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  FBill = fpDblSnglFirstBill.Value
  If FBill < FirstNum Then
    Call TaxMsg(900, "The first bill cannot be less than " + CStr(FirstNum) + ". Please re-enter and try again.")
    fpDblSnglFirstBill.SetFocus
    Exit Sub
  End If
  
  LBill = fpDblSnglLastBill.Value
  If LBill > SecondNum Then
    Call TaxMsg(900, "The last bill cannot be greater than " + CStr(SecondNum) + ". Please re-enter and try again.")
    fpDblSnglLastBill.SetFocus
    Exit Sub
  End If
  
  If FBill > LBill Then
    Call TaxMsg(900, "The first bill number must be smaller than the last.")
    fpDblSnglFirstBill.SetFocus
    Exit Sub
  End If
  
  If TaxMasterRec.TaxForm = 16716 Then
    Call PrintLaser1
    Exit Sub
  End If
  
  If TaxMasterRec.TaxForm = 20007 Then
    Call PrintLSRLEGAL
    Exit Sub
  End If
  
  If TaxMasterRec.TaxForm = 20008 Then
    Call PrintLSRLEGALHP
    Exit Sub
  End If
  
  RptHandle = FreeFile
  
  If TaxMasterRec.TaxForm = 20304 Then
    RptFile$ = "TAXBILLPC1.PRN"
  ElseIf TaxMasterRec.TaxForm = 20002 Then
    RptFile$ = "TAXBILLHAMLET.PRN"
  ElseIf TaxMasterRec.TaxForm = 20003 Then
    RptFile$ = "TAXBILLPH24TF.PRN"
  ElseIf TaxMasterRec.TaxForm = 20004 Then
    RptFile$ = "TAXBILLSYL23TF.PRN"
  ElseIf TaxMasterRec.TaxForm = 20005 Then
    RptFile$ = "TAXBILLBSC32TF.PRN"
  ElseIf TaxMasterRec.TaxForm = 21837 Then
    RptFile$ = "TAXBILLSTANDARD.PRN"
  ElseIf BillFormat = 20006 Then
    RptFile$ = "TAXBILLLLN21TF.PRN"
  End If
  
  Open RptFile For Output As RptHandle
  
  OpenTaxBillFile TBHandle, NumOfTBRecs
  
  If TaxMasterRec.TaxForm = 20002 Then
    Call TaxMsg(800, "12 Pitch is recommended for this form. Please be sure to always start on bill form #1 and not on bill form #2.")
  ElseIf TaxMasterRec.TaxForm = 20003 Then
    Call TaxMsg(900, "10 Pitch is recommended for this form.")
  ElseIf TaxMasterRec.TaxForm = 20004 Then
    Call TaxMsg(900, "12 Pitch is recommended for this form.")
  ElseIf TaxMasterRec.TaxForm = 20005 Then
    Call TaxMsg(900, "12 Pitch is recommended for this form.")
  ElseIf TaxMasterRec.TaxForm = 20006 Then
    Call TaxMsg(900, "10 Pitch is recommended for this form.")
  ElseIf TaxMasterRec.TaxForm = 20304 Then
    Call TaxMsg(900, "12 Pitch is recommended for this form.")
  End If
  
  NumOfMRRecs = 0
  NumOfZRecs = 0
  If Exist("MORTIDX.DAT") Then '12/6/06
    OpenMortIdxFile MRHandle, NumOfMRRecs
    NumOfTBRecs = NumOfMRRecs
  ElseIf Exist("ZIPIDX.DAT") Then '12/6/06
    OpenZipIdxFile ZHandle, NumOfZRecs
    NumOfTBRecs = NumOfZRecs
  End If
  
  For x = 1 To NumOfTBRecs
    If NumOfMRRecs > 0 Then '12/6/06
      Get MRHandle, x, MortRec
      WhatRec& = MortRec.TaxBillRec
    ElseIf NumOfZRecs > 0 Then '12/6/06
      Get ZHandle, x, ZipRec
      WhatRec& = ZipRec.TaxBillRec
    Else
      WhatRec& = x '12/6/06
    End If
    Get TBHandle, WhatRec&, TaxBill
    If TaxBill.BillPrinted Then
      If TaxBill.BillNumber >= FBill And TaxBill.BillNumber <= LBill Then
        RSet PINTemp = QPTrim$(TaxBill.RealPin)
        PrnCnt = PrnCnt + 1
        CustName$ = QPTrim$(TaxBill.CustName)
        If TaxMasterRec.TaxForm = 20304 Then
          Call PrintPostCard1(RptHandle, TBHandle, TaxBill, CustName, PINTemp, PrnCnt)
        ElseIf TaxMasterRec.TaxForm = 20002 Then
          Call PrintHMLT24TF(RptHandle, TBHandle, TaxBill, CustName, PINTemp, PrnCnt)
        ElseIf TaxMasterRec.TaxForm = 20003 Then
          Call PrintPH24TF(RptHandle, TBHandle, TaxBill, CustName, PINTemp, PrnCnt)
        ElseIf TaxMasterRec.TaxForm = 20004 Then
          Call PrintSYL23TF(RptHandle, TBHandle, TaxBill, CustName, PINTemp, PrnCnt)
        ElseIf TaxMasterRec.TaxForm = 20005 Then
          Call PrintBSC32TF(RptHandle, TBHandle, TaxBill, CustName, PINTemp, PrnCnt)
        ElseIf TaxMasterRec.TaxForm = 20006 Then
          Call PrintLLN21TF(RptHandle, TBHandle, TaxBill, CustName, PINTemp, PrnCnt)
        ElseIf TaxMasterRec.TaxForm = 21837 Then 'multi form
          Call PrintStandard(RptHandle, TBHandle, TaxBill, CustName, PINTemp, PrnCnt)
        Else
          Call TaxMsg(900, "The bill format cannot be recognized. Reprints aborted.")
          Close
          Exit Sub
        End If
      End If
    End If
  Next x
  
  Close TBHandle
  Print #RptHandle, FF$
  Close RptHandle
  
  ViewPrint RptFile$, "Tax Bill Reprinting", True
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxBillReprinting", "cmdProcess_Click", Erl)
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
      MainLog ("CitiTaxes.exe terminated via menu bar on frmTaxBillReprinting.")
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
  MainLog ("User opened frmTaxBillPrinting.")
  Me.HelpContextID = hlpReprintTax
  Call LoadMe
End Sub

Private Sub LoadMe()
  Dim BillInfo As TaxBillInfoType
  Dim BIHandle As Integer
  Dim IdxType As Integer
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim MortCodeRec As MortCodeRecType
  Dim MCHandle As Integer
  Dim NumMortCodes As Integer
  Dim x As Integer
  Dim TaxBill As TaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim WhatRec As Long
  Dim ZipRec As BillPrintZipIdxType
  Dim ZHandle As Integer
  Dim NumOfZRecs As Long
  Dim MortRec As BillPrintMortIdxType
  Dim MRHandle As Integer
  Dim NumOfMRRecs As Long
  Dim MortYN As Boolean
  Dim ZipYN As Boolean
  
  'on error goto ERRORSTUFF
  
  doAlign = False
  OpenTaxBillFile TBHandle, NumOfTBRecs
  MortYN = False
  ZipYN = False
  If Exist("MORTIDX.DAT") Then '12/6/06
    OpenMortIdxFile MRHandle, NumOfMRRecs
    NumOfTBRecs = NumOfMRRecs
    MortYN = True
  ElseIf Exist("ZIPIDX.DAT") Then '12/6/06
    OpenZipIdxFile ZHandle, NumOfZRecs
    NumOfTBRecs = NumOfZRecs
    ZipYN = True
  End If

  For x = NumOfTBRecs To 1 Step -1
    If MortYN = True Then '12/6/06
      Get MRHandle, x, MortRec
      Get TBHandle, MortRec.TaxBillRec, TaxBill
      If TaxBill.BillNumber > 0 Then
        fpDblSnglLastBill = TaxBill.BillNumber
        Exit For
      End If
    ElseIf ZipYN = True Then '12/6/06
      Get ZHandle, x, ZipRec
      Get TBHandle, ZipRec.TaxBillRec, TaxBill
      If TaxBill.BillNumber > 0 Then
        fpDblSnglLastBill = TaxBill.BillNumber
        Exit For
      End If
    Else
      Get TBHandle, x, TaxBill
      If TaxBill.BillNumber > 0 Then
        fpDblSnglLastBill = TaxBill.BillNumber
        Exit For
      End If
    End If
  Next x
  
  Close TBHandle
  SecondNum = fpDblSnglLastBill
  OpenBillInfoFile BIHandle
  Get BIHandle, 1, BillInfo
  Close BIHandle
  fpLongTaxYear = BillInfo.TaxYear
  fpDblSnglFirstBill = BillInfo.BillNum
  FirstNum = BillInfo.BillNum
  fpDblSnglStartBill = BillInfo.BillNum
  fpDblSnglRealRate = BillInfo.REALRATE
  fpDblSnglPersRate = BillInfo.PERSRATE
  fpDblSnglLateList = BillInfo.LATEPCT
  Select Case QPTrim$(BillInfo.PRNORDER)
    Case "1"
      fptxtOrder = "Account Number Order"
    Case "2"
      fptxtOrder = "Customer Name Order"
    Case "3"
      fptxtOrder = "Search Name Order"
    Case "4"
      fptxtOrder = "Social Security Order"
    Case Else
      fptxtOrder = "Unknown"
  End Select
    
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  TownName = QPTrim$(TaxMasterRec.City)
  Opt1Desc = QPTrim$(TaxMasterRec.OptRev1)
  Opt2Desc = QPTrim$(TaxMasterRec.OptRev2)
  Opt3Desc = QPTrim$(TaxMasterRec.OptRev3)
  If TaxMasterRec.TaxForm = 16716 Or TaxMasterRec.TaxForm = 20007 Then
    cmdAlign.Enabled = False
  End If
  
  BillFormat = TaxMasterRec.TaxForm
  fpcmbPrintPriorYN.Visible = False
  Label11.Visible = False
  If TaxMasterRec.TaxForm = 16716 Then
    fpcmbPrintPriorYN.Text = "N"
    fpcmbPrintPriorYN.AddItem "Y"
    fpcmbPrintPriorYN.AddItem "N"
    fpcmbPrintPriorYN.Visible = True
    Label11.Visible = True
  End If
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxBillReprinting", "LoadMe", Erl)
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

Private Sub PrintBills(ByVal RptHandle As Integer, TBHandle As Integer, TaxBill As TaxBillType, CustName$, PINTemp$, PrnCnt As Long)
  Dim DueDate$
  Dim TAXRATE#
  Dim NetTaxVal#
  Dim LC As Integer
  
  'on error goto ERRORSTUFF
  
  DueDate$ = "12-31-" + QPTrim$(Str$(TaxBill.TaxYear))
  TAXRATE# = TaxBill.RealTaxRate
  If TAXRATE# = 0 Then
    TAXRATE# = TaxBill.PersTaxRate
  End If

  NetTaxVal# = OldRound#(TaxBill.RealValue + TaxBill.PersValue)
  NetTaxVal# = OldRound#(NetTaxVal# - TaxBill.ExptValue)

  Print #RptHandle,
  Print #RptHandle, Tab(29); "TOWN OF MAGGIE VALLEY"
  Print #RptHandle, Tab(29); "    3987 SOCO RD."
  Print #RptHandle, Tab(29); "MAGGIE VALLEY NC 28751"
  Print #RptHandle, Tab(29); "  PROPERTY TAX BILL"

  For LC = 1 To 3
    Print #RptHandle, " "
  Next

  Print #RptHandle, Tab(12); "ACCT # "; TaxBill.CustRec;
  Print #RptHandle, Tab(65); "BILL #"; Using("#####0", TaxBill.BillNumber)
  Print #RptHandle, Tab(12); Left$(CustName$, 25);
  Print #RptHandle, Tab(63); "TAX YEAR "; TaxBill.TaxYear
  Print #RptHandle, Tab(12); Left$(TaxBill.CustAdd1, 25);
  Print #RptHandle, Tab(63); "TAX RATE "; Using("##0.00", TAXRATE#)
  Print #RptHandle, Tab(12); Left$(TaxBill.CustAdd2, 25)
  Print #RptHandle, Tab(12); QPTrim$(TaxBill.CustAdd3); " "; Left$(TaxBill.CustZip, 5) + "-" + Mid$(TaxBill.CustZip, 6, 4)
  For LC = 1 To 4
    Print #RptHandle, " "
  Next
  Print #RptHandle, Tab(39); "[--------- VALUATIONS --------]"
  Print #RptHandle, Tab(2); "PROPERTY DESCRIPTION"; Tab(30); "RATE"; Tab(40); "REAL"; Tab(48); "PERSONAL"; Tab(61); "EXEMPT"; Tab(72); "TOTAL"
  Print #RptHandle, " "
  'Line 23 Starts Here
  Print #RptHandle, Tab(30); Using(".##", TAXRATE#);
  Print #RptHandle, Tab(35); Using("###,###,##0", TaxBill.RealValue);
  Print #RptHandle, Tab(47); Using("###,###,##0", TaxBill.PersValue);
  Print #RptHandle, Tab(59); Using("###,##0", TaxBill.ExptValue);
  Print #RptHandle, Tab(68); Using("###,###,##0", (TaxBill.PersValue + TaxBill.RealValue))
  Print #RptHandle, Tab(2); QPTrim$(TaxBill.RDesc1)
  Print #RptHandle, Tab(2); QPTrim$(PINTemp)


   Print #RptHandle, ""
   Print #RptHandle, ""
   Print #RptHandle, ""
   Print #RptHandle, Tab(2); "NOTE: "
   Print #RptHandle, Tab(2); "      A 2% PENALTY WILL BE ADDED AFTER DUE DATE."
   Print #RptHandle, Tab(2); "      .75% ADDED ON FIRST OF EACH MONTH THEREAFTER."
   Print #RptHandle, ""
   Print #RptHandle, Tab(2); "      PLEASE SUBMIT YOUR ACCOUNT# OR, PARCEL ID# ON CHECK"
   Print #RptHandle, Tab(2); "      TO PROCESS YOUR PAYMENT. CONTACT THE TAX COLLECTOR"
   Print #RptHandle, Tab(2); "      IF YOU HAVE IF ANY QUESTIONS."
   Print #RptHandle, Tab(2); "      (PHONE: 828-926-0866  EXT. 101)"
   Print #RptHandle,
   Print #RptHandle,
   Print #RptHandle,
   Print #RptHandle,
   Print #RptHandle, Tab(39); " TAX DUE DATE: "; DueDate$
   Print #RptHandle,
   Print #RptHandle, Tab(39); "TOTAL TAX DUE: "; Using("###,###,##0", TaxBill.TotalBillDue)
   Print #RptHandle,
   Print #RptHandle,
   Print #RptHandle, "BN"; Using("#####0", PrnCnt)
   Print #RptHandle, Chr$(12);
   
   Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxBillReprinting", "PrintBills", Erl)
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

Private Sub PrintStandard(ByVal RptHandle As Integer, TBHandle As Integer, TaxBill As TaxBillType, CustName$, PINTemp$, PrnCnt As Long)
  Dim TAXRATE#
  Dim NetTaxVal#
  
  'on error goto ERRORSTUFF
  
  TAXRATE# = TaxBill.RealTaxRate
  If TAXRATE# = 0 Then
    TAXRATE# = TaxBill.PersTaxRate
  End If

  Print #RptHandle, "~"; Tab(50); Using("###0", PrnCnt)
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle, TaxBill.TaxYear;
  Print #RptHandle, Tab(7); Using("######", TaxBill.BillNumber);
  Print #RptHandle, Tab(16); Using("#####", TaxBill.CustRec);
  Print #RptHandle, Tab(23); QPTrim$(PINTemp); Tab(37); Using("####", TaxBill.TaxYear);

  Print #RptHandle, Tab(42); Using("########", TaxBill.CustRec);
  'PRINT #RptHandle, TAB(49); USING "######,#.##"; TaxBill.TotalBillDue + TaxBill.PriorYrBalance
  Print #RptHandle, Tab(51); Using("######", TaxBill.CustRec)
  Print #RptHandle,
  Print #RptHandle, Tab(11); Left$(QPTrim$(CustName$), 21)
  Print #RptHandle, Tab(11); Left$(QPTrim$(TaxBill.RDesc1), 21)
  Print #RptHandle, Tab(11); Left$(QPTrim$(TaxBill.RDesc2), 21)
  'v line 12
  Print #RptHandle,
  
  Print #RptHandle, Using("###,###,##0", TaxBill.RealValue); Tab(15); TaxBill.PersValue;
  NetTaxVal# = OldRound#(TaxBill.RealValue + TaxBill.PersValue)
  NetTaxVal# = OldRound#(NetTaxVal# - TaxBill.ExptValue)
  Print #RptHandle, Tab(25); Using("#,###,##0", TaxBill.ExptValue);
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle, Tab(36); Left$(QPTrim$(CustName$), 25)
  Print #RptHandle, Using("###,###,##0", NetTaxVal#);
  Print #RptHandle, Tab(16); Using("#.##", TAXRATE#);
  Print #RptHandle, Tab(21); Using("##,###,##0.00", OldRound#(TaxBill.TotalBillDue - TaxBill.LateTaxDue));
  Print #RptHandle, Tab(36); Left$(QPTrim$(TaxBill.CustAdd1), 25)
  Print #RptHandle, Tab(36); Left$(QPTrim$(TaxBill.CustAdd2), 25)
  Print #RptHandle, Tab(36); QPTrim$(TaxBill.CustAdd3); " "; QPTrim$(TaxBill.CustZip)
  Print #RptHandle, Tab(21); Using("##,###,##0.00", TaxBill.LateTaxDue)
  Print #RptHandle,
  Print #RptHandle, Tab(2); Using("###,##0.0", TaxBill.PriorYrBalance);
  If TaxBill.OverPayAmt > 0 Then
    Print #RptHandle, Tab(21); Using("#,###,##0.00", TaxBill.TotalBillDue + TaxBill.PriorYrBalance); Tab(47); Using("#,###,##0.00", TaxBill.TotalBillDue + TaxBill.PriorYrBalance - TaxBill.OverPayAmt)
  Else
    Print #RptHandle, Tab(21); Using("#,###,##0.00", TaxBill.TotalBillDue + TaxBill.PriorYrBalance); Tab(47); Using("#,###,##0.00", TaxBill.TotalBillDue + TaxBill.PriorYrBalance)
  End If
  If TaxBill.OverPayAmt > 0 Then
    Print #RptHandle, "Credit of " + QPTrim$(Using$("$###,##0.00", TaxBill.OverPayAmt)) + " has been applied."
  Else
    Print #RptHandle,
  End If
  Print #RptHandle, "~"

  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxBillReprinting", "PrintStandard", Erl)
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

Private Sub PrintLaser1()
  Dim ToPrint As String
  Dim TaxRptT As Integer
  Dim ReportFile As String
  Dim RptHandle As Integer
  Dim TBRec As TaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim x As Long
  Dim dlm$
  Dim TBDRec As TxBill1DefaultsType
  Dim TBDHandle As Integer
  Dim FBill&, PrnCnt&
  Dim LBill&
  Dim NCnt As Integer
  Dim ThisRate As Double
  Dim BZip As String
  Dim TaxCust As TaxCustType
  Dim NumOfTCRecs As Long
  Dim TCHandle As Integer
  Dim PrintPriorYN As Boolean
  Dim ZipRec As BillPrintZipIdxType
  Dim ZHandle As Integer
  Dim NumOfZRecs As Long
  Dim MortRec As BillPrintMortIdxType
  Dim MRHandle As Integer
  Dim NumOfMRRecs As Long
  Dim CustArr As Long
  
  'on error goto ERRORSTUFF
  
  If fpcmbPrintPriorYN.Text = "N" Then
    PrintPriorYN = False
  Else
    PrintPriorYN = True
  End If
  dlm$ = "~"
  FBill = fpDblSnglFirstBill.Value
  LBill = fpDblSnglLastBill.Value
  
  ReportFile$ = StartPath$ + "/TaxBil1.PRN"
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  OpenTxBill1File TBDHandle
  Get #TBDHandle, 1, TBDRec
  Close TBDHandle
  ARptTempTaxBill.Head1 = QPTrim(TBDRec.TxtHead1)
  ARptTempTaxBill.Head2 = QPTrim(TBDRec.TxtHead2)
  ARptTempTaxBill.LblOpt1 = QPTrim(TBDRec.txtOpt1)
  ARptTempTaxBill.LblOpt2 = QPTrim(TBDRec.TxtOpt2)
  ARptTempTaxBill.LblOpt3 = QPTrim(TBDRec.TxtOpt3)
  ARptTempTaxBill.LblOpt4 = QPTrim(TBDRec.TxtOpt4)
  ARptTempTaxBill.LblPgph1 = QPTrim(TBDRec.txtPgph0)
  ARptTempTaxBill.LblPgph2 = QPTrim(TBDRec.txtPgph1)
  ARptTempTaxBill.LblPgph3 = QPTrim(TBDRec.txtPgph2)
  ARptTempTaxBill.LblPgph4 = QPTrim(TBDRec.txtPgph3)
  ARptTempTaxBill.LblPgph5 = QPTrim(TBDRec.txtPgph4)
  ARptTempTaxBill.LblPgph6 = QPTrim(TBDRec.txtPgph5)
  ARptTempTaxBill.LblPgph7 = QPTrim(TBDRec.txtPgph6)
  ARptTempTaxBill.LblPgph8 = QPTrim(TBDRec.txtPgph7)
  ARptTempTaxBill.LblOpt5 = QPTrim(TBDRec.TxtOpt5)
  ARptTempTaxBill.LblHead4 = QPTrim(TBDRec.txtHead4)
  ARptTempTaxBill.LblHead5 = QPTrim(TBDRec.txtHead5)
  ARptTempTaxBill.LblHead6 = QPTrim(TBDRec.txtHead6)
  ARptTempTaxBill.LblOpt6 = QPTrim(TBDRec.TxtOpt6)
  ARptTempTaxBill.LblOpt7 = QPTrim(TBDRec.TxtOpt7)
  If TBDRec.dologo = 1 Then
    If Exist("towntaxlogo.bmp") Then
      ARptTempTaxBill.Image1.Picture = LoadPicture("towntaxlogo.bmp")
      ARptTempTaxBill.Image1.Visible = True
    End If
  End If
  
  frmTaxShowPctComp.Label1 = "Printing Tax Bills"
  frmTaxShowPctComp.cmdCancel.Visible = False
  frmTaxShowPctComp.Show , Me
  cmdProcess.Enabled = False
  cmdExit.Enabled = False
  EnableCloseButton Me.hwnd, False
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenTaxBillFile TBHandle, NumOfTBRecs
  
  If Exist("MORTIDX.DAT") Then '12/6/06
    OpenMortIdxFile MRHandle, NumOfMRRecs
    NumOfTBRecs = NumOfMRRecs
  ElseIf Exist("ZIPIDX.DAT") Then '12/6/06
    OpenZipIdxFile ZHandle, NumOfZRecs
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
      If TBRec.BillNumber >= FBill And TBRec.BillNumber <= LBill Then
        If TBRec.TotalBillDue > 0 Then
          GoSub GetBarCodeData
          Print #RptHandle, TBRec.BillNumber; dlm; QPTrim$(TBRec.CustName); dlm;
          Print #RptHandle, QPTrim$(TBRec.CustAdd1); dlm; QPTrim$(TBRec.CustAdd2); dlm;
          Print #RptHandle, QPTrim$(TBRec.CustAdd3); dlm; TBRec.CustPin; dlm;
          Print #RptHandle, QPTrim$(TBRec.RealPin); dlm; QPTrim$(TBRec.RDesc1); dlm;
          Print #RptHandle, TBRec.RealValue; dlm; TBRec.PersValue; dlm; TBRec.ExptValue; dlm;
          If TBRec.RealTaxDue > 0 And TBRec.PersTaxDue > 0 Then
            ThisRate = TBRec.RealTaxRate
          ElseIf TBRec.RealTaxDue <= 0 And TBRec.PersTaxDue > 0 Then
            ThisRate = TBRec.PersTaxRate
          ElseIf TBRec.RealTaxDue > 0 And TBRec.PersTaxDue <= 0 Then
            ThisRate = TBRec.RealTaxRate
          Else
            ThisRate = 0
          End If
          Print #RptHandle, OldRound(TBRec.RealValue + TBRec.PersValue); dlm; ThisRate; dlm;
          If TBRec.OverPayAmt > 0 Then
            '                                        13                              14            15
            Print #RptHandle, OldRound(TBRec.TotalBillDue - TBRec.OverPayAmt); dlm; BZip; dlm; TBDRec.dologo; dlm;
            '                         16                   17                     18
            Print #RptHandle, TBRec.OptRevTax1; dlm; TBRec.OptRevTax2; dlm; TBRec.OptRevTax3; dlm;
            '                    19             20             21                    22                       23                     24
            Print #RptHandle, Opt1Desc; dlm; Opt2Desc; dlm; Opt3Desc; dlm; QPTrim$(TBRec.CustZip); dlm; TBRec.LateTaxDue; dlm; TBRec.PriorYrBalance; dlm; PrintPriorYN '  TBRec.OverPayAmt
          Else
            '                        13                 14             15
            Print #RptHandle, TBRec.TotalBillDue; dlm; BZip; dlm; TBDRec.dologo; dlm;
            '                         16                   17                     18
            Print #RptHandle, TBRec.OptRevTax1; dlm; TBRec.OptRevTax2; dlm; TBRec.OptRevTax3; dlm;
          '                    19             20             21                 22                          23                      24
          Print #RptHandle, Opt1Desc; dlm; Opt2Desc; dlm; Opt3Desc; dlm; QPTrim$(TBRec.CustZip); dlm; TBRec.LateTaxDue; dlm; TBRec.PriorYrBalance; dlm; PrintPriorYN 'TBRec.OverPayAmt
          End If
        End If
      End If
    frmTaxShowPctComp.ShowPctComp x, NumOfTBRecs
    If frmTaxShowPctComp.Out = True Then
      Close
      frmTaxShowPctComp.Out = False
      Unload frmTaxShowPctComp
      cmdProcess.Enabled = True
      cmdExit.Enabled = True
      EnableCloseButton Me.hwnd, True
      Exit Sub
    End If
  Next x
  Close TBHandle
  
  Unload frmTaxShowPctComp
  cmdProcess.Enabled = True
  cmdExit.Enabled = True
  EnableCloseButton Me.hwnd, True
  
  Close
  ARptTempTaxBill.GetName ReportFile$
  ARptTempTaxBill.Show

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
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxBillReprinting", "PrintLaser1", Erl)
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

Private Sub PrintPostCard1(ByVal RptHandle As Integer, TBHandle As Integer, TaxBill As TaxBillType, CustName$, PINTemp$, PrnCnt As Long)
  'good for Troy
  Dim x As Long, BillNo&
  Dim PersRec As PersonalRecType
  Dim PHandle As Integer
  Dim NumOfPRecs As Long
  Dim RealRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRRecs As Long
  Dim ThisRec As Long
  Dim ThisDesc As String * 20
  Dim LotsAcres As String * 20
  Dim ThisCName As String * 22
  Dim ThisName As String * 28
  Dim ThisAdd1 As String * 28
  Dim ThisAdd2 As String * 28
  Dim ThisAdd3 As String * 28
  Dim FF$
  Dim PastDue#
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  
  'on error goto ERRORSTUFF
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  FF$ = Chr(12)
  PastDue = GetCustBalance(TaxBill.CustRec, -1) 'added 8/14/06
  ThisName = QPTrim$(TaxBill.CustName)
  ThisCName = QPTrim$(TaxBill.CustName)
  ThisAdd1 = QPTrim$(TaxBill.CustAdd1)
  ThisAdd2 = QPTrim$(TaxBill.CustAdd2)
  ThisAdd3 = QPTrim$(TaxBill.CustAdd3) + " " + QPTrim$(TaxBill.CustZip)
  OpenPersPropFile PHandle, NumOfPRecs
  OpenRealPropFile RHandle, NumOfRRecs
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  '-----------------------------------------
  Print #RptHandle, Tab(1); Using$("###0", TaxBill.TaxYear); Tab(7); Using$("####0", TaxBill.BillNumber); Tab(14); Using$("######0", TaxBill.CustPin);
'  Print #RptHandle, Tab(22); ""; Tab(33); Using$("###0", TaxBill.TaxYear); Tab(40); Using$("####0", BillNo); Tab(50); Using$("######0", TaxBill.CustPin)
  If InStr(TaxMasterRec.Name, "SEVEN DEVILS") Then
    Print #RptHandle, Tab(22); ""; Tab(40); Using$("###0", TaxBill.TaxYear); Tab(50); ""; Tab(60); Using$("######0", TaxBill.CustPin)
  ElseIf InStr(TaxMasterRec.Name, "ANDREWS") Then
    Print #RptHandle, Tab(22); ""; Tab(40); Using$("###0", TaxBill.TaxYear); Tab(50); Using$("######0", TaxBill.CustPin); Tab(60); Using$("$##,##0.00", TaxBill.TotalBillDue)
  Else
    Print #RptHandle, Tab(26); Right(QPTrim$(TaxBill.RealPin), 11); Tab(40); Using$("###0", TaxBill.TaxYear); Tab(50); Using$("####0", TaxBill.BillNumber); Tab(60); Using$("######0", TaxBill.CustPin)
  End If
  '---end of line 1-------------------------
  Print #RptHandle, Tab(10); ThisCName 'end of line 2
  If TaxBill.RealPropRecord > 0 And TaxBill.PersPropRecord > 0 Then
    ThisDesc = "Real And Personal"
    Get RHandle, TaxBill.RealPropRecord, RealRec
    LotsAcres = QPTrim$(RealRec.LOTACRE) + "/" + QPTrim$(RealRec.LOTNUMB)
  ElseIf TaxBill.RealPropRecord > 0 Then
    Get RHandle, TaxBill.RealPropRecord, RealRec
    ThisDesc = QPTrim$(RealRec.PROPNOT1)
    LotsAcres = QPTrim$(RealRec.LOTACRE) + "/" + QPTrim$(RealRec.LOTNUMB)
  ElseIf TaxBill.PersPropRecord > 0 Then
    Get PHandle, TaxBill.PersPropRecord, PersRec
    ThisDesc = QPTrim$(PersRec.DESC1)
    LotsAcres = ""
  Else
    ThisDesc = ""
    LotsAcres = ""
  End If
  If TaxBill.OverPayAmt > 0 Then 'late tax should not come into play because of the overpay amt
    TaxBill.TotalBillDue = OldRound(TaxBill.TotalBillDue - TaxBill.OverPayAmt)
  End If
  Print #RptHandle, Tab(10); ThisDesc 'end of line 3
  Print #RptHandle, Tab(10); LotsAcres 'end of line 4
  Print #RptHandle, 'space
  Print #RptHandle, Tab(2); Using$("$###,###.00", TaxBill.RealValue); Tab(15); Using$("$###,##0.00", TaxBill.PersValue); Tab(26); Using$("$###,##0.00", TaxBill.ExptValue);
  Print #RptHandle, 'space
  Print #RptHandle, 'space
  Print #RptHandle,
  If TaxBill.RealTaxRate > 0 Then
    Print #RptHandle, Tab(2); Using$("$###,##0.00", OldRound(TaxBill.RealValue + TaxBill.PersValue)); Tab(16); Using$("##0.00", TaxBill.RealTaxRate);
  ElseIf TaxBill.PersTaxRate > 0 Then
    Print #RptHandle, Tab(2); Using$("$###,##0.00", OldRound(TaxBill.RealValue + TaxBill.PersValue)); Tab(16); Using$("##0.00", TaxBill.PersTaxRate);
  Else
    Print #RptHandle, Tab(2); Using$("$###,##0.00", OldRound(TaxBill.RealValue + TaxBill.PersValue)); Tab(16); Using$("##0.00", 0);
  End If
  Print #RptHandle, Tab(26); Using$("$###,##0.00", OldRound(TaxBill.OverPayAmt + TaxBill.TotalBillDue - TaxBill.LateTaxDue)); Tab(42); ThisName '8/16/06 added OverPayment
  Print #RptHandle, Tab(42); ThisAdd1
  Print #RptHandle, Tab(42); ThisAdd2
  Print #RptHandle, Tab(28); Using$("$#,##0.00", TaxBill.LateTaxDue); Tab(42); ThisAdd3
  Print #RptHandle,
  Print #RptHandle, Using$("$###,##0.00", PastDue); Tab(26); Using$("$###,##0.00", TaxBill.TotalBillDue)
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  
  Close RHandle
  Close PHandle
  
  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxBillReprinting", "PrintPostCard1", Erl)
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

Private Sub fpDblSnglFirstBill_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyUp Then
    fpDblSnglLastBill.SetFocus
  ElseIf KeyCode = vbKeyDown Then
    fpDblSnglLastBill.SetFocus
  End If

End Sub

Private Sub fpDblSnglLastBill_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    fpDblSnglFirstBill.SetFocus
  ElseIf KeyCode = vbKeyUp Then
    fpDblSnglFirstBill.SetFocus
  End If
End Sub

Private Sub PrintHMLT24TF(ByVal RptHandle As Integer, TBHandle As Integer, TaxBill As TaxBillType, CustName$, PINTemp$, PrnCnt As Long)
  Dim NetTaxVal#
  Static Cnt As Integer
  
'  'on error goto ERRORSTUFF
  Print #RptHandle, "~"
  Print #RptHandle, Using("#####", PrnCnt)
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle, Tab(16); Using("######", TaxBill.BillNumber);
  If TaxBill.RealValue < -1000 Then
    TaxBill.RealValue = 0
  End If
  If TaxBill.PersValue < -1000 Then
    TaxBill.PersValue = 0
  End If
  Print #RptHandle, Tab(48); Using("###,###,###", TaxBill.RealValue)
  Print #RptHandle, Tab(16); CStr(TaxBill.TaxYear);
  Print #RptHandle, Tab(48); Using("###,###,###", TaxBill.PersValue)
  If TaxBill.RealTaxRate# = 0 Then
    TaxBill.RealTaxRate# = TaxBill.PersTaxRate#
  End If
  Print #RptHandle, Tab(16); Using("#.0##", TaxBill.RealTaxRate#);
  NetTaxVal# = OldRound#(TaxBill.RealValue + TaxBill.PersValue)
  Print #RptHandle, Tab(48); Using("###,###,###", NetTaxVal#);
  Print #RptHandle, Tab(84); Using("$#,###,##0.00", TaxBill.TotalBillDue - TaxBill.LateTaxDue)
  Print #RptHandle, Tab(20); Using("#####0", TaxBill.CustRec);
  Print #RptHandle, Tab(55); Using("#,###,###", TaxBill.ExptValue);
  Print #RptHandle, Tab(81); Using("$#,###,##0.00", TaxBill.LateTaxDue)
  Print #RptHandle, Tab(14); QPTrim$(TaxBill.TownShip);
  NetTaxVal# = OldRound#(NetTaxVal# - TaxBill.ExptValue)
  Print #RptHandle, Tab(48); Using("###,###,###", NetTaxVal#)
  Print #RptHandle, Tab(18); QPTrim$(TaxBill.LotOrAcre);
  Print #RptHandle, " "; QPTrim$(TaxBill.LASize)
  If QPTrim$(TaxBill.RealPin) <> "" Then
    Print #RptHandle, Tab(18); QPTrim$(TaxBill.RealPin);
  Else
    Print #RptHandle, Tab(18); "";
  End If
  Print #RptHandle, Tab(86); Using("$##,###,##0.00", TaxBill.TotalBillDue - TaxBill.OverPayAmt) 'added OverPayAmt
  If QPTrim$(TaxBill.RDesc1) <> "" Then
    Print #RptHandle, Tab(18); QPTrim$(TaxBill.RDesc1)
  Else
    Print #RptHandle, ""
  End If
  If QPTrim$(TaxBill.RDesc2) <> "" Then
    Print #RptHandle, Tab(18); QPTrim$(TaxBill.RDesc2)
  Else
    Print #RptHandle, ""
  End If
  Print #RptHandle, ""
  Print #RptHandle, ""
  Print #RptHandle, Tab(8); Left$(QPTrim$(CustName$), 35)
  Print #RptHandle, Tab(8); Left$(QPTrim$(TaxBill.CustAdd1), 35)
  Print #RptHandle, Tab(8); Left$(QPTrim$(TaxBill.CustAdd2), 35)
  Print #RptHandle, Tab(8); QPTrim$(TaxBill.CustAdd3); " "; QPTrim$(TaxBill.CustZip)
  Print #RptHandle, ""
  Print #RptHandle, ""
  Print #RptHandle, ""
  Cnt = Cnt + 1
  If Cnt <> 1 Then
    Cnt = 0
    Print #RptHandle, "" '8/16/06
  End If
  
End Sub

Private Sub PrintPH24TF(ByVal RptHandle As Integer, TBHandle As Integer, TaxBill As TaxBillType, CustName$, PINTemp$, PrnCnt As Long)
  Dim TAXRATE#
  Dim NetTaxVal#
  
  If InStr(TownName, "WHITAKERS") Then
   If QPTrim$(TaxBill.RealPin) = "" Then GoTo EmptyPin
   RSet PINTemp = Mid(TaxBill.RealPin, Len(QPTrim$(TaxBill.RealPin)) - 3, 4)
  End If
  
EmptyPin:
  
  TAXRATE# = CDbl(fpDblSnglRealRate.Value)
  
  Print #RptHandle, "~"; Tab(40); Using("###0", PrnCnt)
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle, Using("###0", TaxBill.TaxYear);
  Print #RptHandle, Tab(5); Using("#####0", TaxBill.BillNumber);
  Print #RptHandle, Tab(15); Using("####0", TaxBill.CustRec);
  Print #RptHandle, Tab(23); QPTrim$(PINTemp); Tab(34); Using("###0", TaxBill.TaxYear);

  Print #RptHandle, Tab(38); Using("#######0", TaxBill.CustRec);
  Print #RptHandle, Tab(46); Using("#,###,##0.00", OldRound#(TaxBill.TotalBillDue + TaxBill.PriorYrBalance))
  Print #RptHandle, Tab(11); QPTrim$(Left$(CustName$, 21))
  Print #RptHandle, Tab(11); QPTrim$(Left$(TaxBill.RDesc1, 21))
  Print #RptHandle, Tab(11); QPTrim$(Left$(TaxBill.RDesc2, 21))
  'v line 12
  Print #RptHandle,
  
  Print #RptHandle, Using("###,###,##0", TaxBill.RealValue); Tab(12); Using("###,###,##0", TaxBill.PersValue);
  NetTaxVal# = OldRound#(TaxBill.RealValue + TaxBill.PersValue)
  NetTaxVal# = OldRound#(NetTaxVal# - TaxBill.ExptValue)

  Print #RptHandle, Tab(23); Using("#,###,##0", TaxBill.ExptValue);
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle, Tab(36); Left$(CustName$, 25)
  Print #RptHandle, Using("###,###,##0", NetTaxVal#);
  Print #RptHandle, Tab(14); Using("#.000", TAXRATE#);
  Print #RptHandle, Tab(19); Using("##,###,##0.00", OldRound#(TaxBill.TotalBillDue - TaxBill.LateTaxDue));
  Print #RptHandle, Tab(36); Left$(QPTrim$(TaxBill.CustAdd1), 25)
  Print #RptHandle, Tab(36); Left$(QPTrim$(TaxBill.CustAdd2), 25)
  Print #RptHandle, Tab(36); QPTrim$(TaxBill.CustAdd3); " "; QPTrim$(TaxBill.CustZip)
  Print #RptHandle, Tab(21); Using("#######0.00", TaxBill.LateTaxDue)
  Print #RptHandle,
  Print #RptHandle, Tab(2); Using("###,##0.00", TaxBill.PriorYrBalance);
  Print #RptHandle, Tab(19); Using("##,###,##0.00", OldRound(TaxBill.TotalBillDue + TaxBill.PriorYrBalance)) '; Tab(47); Using("##,###,##0.00", OldRound(TaxBill.TotalBillDue + TaxBill.PriorYrBalance))
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle, "~"

End Sub

Private Sub PrintSYL23TF(ByVal RptHandle As Integer, TBHandle As Integer, TaxBill As TaxBillType, CustName$, PINTemp$, PrnCnt As Long)
  Dim NetTaxVal#
  Dim TAXRATE#
  
  TAXRATE# = CDbl(fpDblSnglRealRate.Value)
  
  Print #RptHandle, Chr$(27); Chr$(58); "~"
  Print #RptHandle, 'added 6.23.06
  Print #RptHandle, Tab(32); Using$("###0", TaxBill.TaxYear);
  Print #RptHandle, Tab(44); Using("#.00", TAXRATE#);
  Print #RptHandle, Tab(78); Using("#####0", TaxBill.CustRec);
  Print #RptHandle, Tab(90); Using("#####0", TaxBill.BillNumber)
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  NetTaxVal# = OldRound#(TaxBill.RealValue + TaxBill.PersValue)
  Print #RptHandle, Using("#,###,##0", TaxBill.PersValue);

  Print #RptHandle, Tab(13); Using("##,###,##0", TaxBill.RealValue);

  Print #RptHandle, Tab(26); Using("##,###,##0", NetTaxVal#);
  NetTaxVal# = OldRound#(NetTaxVal# - TaxBill.ExptValue)
  Print #RptHandle, Tab(37); Using("#####0.00", OldRound(TaxBill.TotalBillDue - TaxBill.LateTaxDue));
  Print #RptHandle, Tab(49); Using("##,###,##0", TaxBill.ExptValue);
  Print #RptHandle, Tab(82); Using("###0.00", TaxBill.LateTaxDue);
  Print #RptHandle, Tab(89); Using("####0.00", TaxBill.TotalBillDue - TaxBill.OverPayAmt)
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle, Tab(8); '"IF TAXES ARE ESCROWED SEND BILL TO"
  Print #RptHandle, Tab(8); '"MORTGAGE COMPANY."
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle, 'added 6.23.06
  Print #RptHandle, Tab(8); Left$(QPTrim$(CustName$), 25)
  Print #RptHandle, Tab(8); Left$(QPTrim$(TaxBill.CustAdd1), 25)
  Print #RptHandle, Tab(8); Left$(QPTrim$(TaxBill.CustAdd2), 25)
  Print #RptHandle, Tab(8); QPTrim$(TaxBill.CustAdd3); " "; QPTrim$(TaxBill.CustZip)
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle, "~"

End Sub

Private Sub PrintBSC32TF(ByVal RptHandle As Integer, TBHandle As Integer, TaxBill As TaxBillType, CustName$, PINTemp$, PrnCnt As Long)
  Dim NetTaxVal#
  Dim TAXRATE#

  TAXRATE# = TaxBill.RealTaxRate
  If TAXRATE# = 0 Then
    TAXRATE# = TaxBill.PersTaxRate
  End If

  Print #RptHandle, Chr$(27); Chr$(48); "~" '; Tab(50); Using("###0", PrnCnt)
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle, Using$("###0", TaxBill.TaxYear);
  Print #RptHandle, Tab(10); Using("#####0", TaxBill.BillNumber);
  Print #RptHandle, Tab(19); Using("####0", TaxBill.CustRec);
  Print #RptHandle, Tab(34); Using("###0", TaxBill.TaxYear);
  Print #RptHandle, Tab(43); Using("#0.00", TAXRATE#);
  Print #RptHandle, Tab(48); Using("#####0.00", OldRound(TaxBill.TotalBillDue + TaxBill.PriorYrBalance))
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle, Tab(2); QPTrim$(Left$(TaxBill.RDesc1, 21))
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle, QPTrim$(PINTemp); Tab(20); Using("#######0", TaxBill.ExptValue)
  Print #RptHandle,
  Print #RptHandle, Tab(48); Using("###,##0.00", OldRound(TaxBill.TotalBillDue + TaxBill.PriorYrBalance))
  Print #RptHandle,
  NetTaxVal# = OldRound#(TaxBill.RealValue + TaxBill.PersValue)
  Print #RptHandle, Using("#######0", TaxBill.RealValue); Tab(11); Using("#######0", TaxBill.PersValue); Tab(22); Using("########0", NetTaxVal#)
  NetTaxVal# = OldRound#(NetTaxVal# - TaxBill.ExptValue)
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  
  Print #RptHandle, Tab(24); Using$("###0", TaxBill.TaxYear);
  Print #RptHandle, Tab(30); Using$("#####0", TaxBill.BillNumber);
  Print #RptHandle, Tab(39); Using$("####0", TaxBill.CustRec);
  Print #RptHandle, Tab(48); Using$("###,##0.00", OldRound#(TaxBill.TotalBillDue + TaxBill.PriorYrBalance))
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle, Tab(30); Left$(CustName$, 25)
  Print #RptHandle, Tab(30); Left$(QPTrim$(TaxBill.CustAdd1), 25)
  Print #RptHandle, Tab(30); Left$(QPTrim$(TaxBill.CustAdd2), 25)
  Print #RptHandle, Tab(30); QPTrim$(TaxBill.CustAdd3); " "; QPTrim$(TaxBill.CustZip)
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle, "~"

End Sub

Private Sub PrintLLN21TF(ByVal RptHandle As Integer, TBHandle As Integer, TaxBill As TaxBillType, CustName$, PINTemp$, PrnCnt As Long)
  Dim NetTaxVal#
  Dim TAXRATE#
  
  TAXRATE# = TaxBill.RealTaxRate
  If TAXRATE# = 0 Then
    TAXRATE# = TaxBill.PersTaxRate
  End If

  NetTaxVal# = OldRound#(TaxBill.RealValue + TaxBill.PersValue)
  NetTaxVal# = OldRound#(NetTaxVal# - TaxBill.ExptValue)

  Print #RptHandle, "~"; Tab(78); "~"
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle, Tab(50); TaxBill.TaxYear; Tab(59); Using("#####0", TaxBill.BillNumber);
  Print #RptHandle, Tab(69); QPTrim$(TaxBill.RealPin)
  Print #RptHandle,
  Print #RptHandle, Tab(50); Using("#####0", TaxBill.CustRec)
  Print #RptHandle,
  Print #RptHandle, Tab(51); QPTrim$(Left$(TaxBill.RDesc1, 21))
  Print #RptHandle,
  Print #RptHandle, Tab(51); Using("##,###,##0", TaxBill.RealValue); Tab(61); Using("##,###,##0", TaxBill.PersValue); Tab(71); Using("###,###,##0", NetTaxVal#)
'  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle, Tab(5); Left$(CustName$, 25);
  Print #RptHandle, Tab(66); Using("#.00", TAXRATE#);
  Print #RptHandle, Tab(71); Using("####0.00", OldRound#(TaxBill.TotalBillDue - TaxBill.LateTaxDue))
  Print #RptHandle, Tab(5); Left$(TaxBill.CustAdd1, 25)
  Print #RptHandle, Tab(5); Left$(TaxBill.CustAdd2, 25);
  Print #RptHandle, Tab(71); Using("####.00", TaxBill.LateTaxDue)
  Print #RptHandle, Tab(5); QPTrim$(TaxBill.CustAdd3); " "; QPTrim$(TaxBill.CustZip)
'  Print #RptHandle, Tab(51); "09-01-"; Right$(Date$, 4); Tab(71); Using; "#####.##"; TaxBill.TotalBillDue
  Print #RptHandle, Tab(71); Using("####0.00", TaxBill.TotalBillDue - TaxBill.OverPayAmt) 'added OverPayAmt 8/15/06
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle, "~"; Tab(78); "~"

End Sub

Private Sub PrintLSRLEGAL()
  Dim RptHandle As Integer
  Dim RptFile$
  Dim TaxBill As TaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim x As Long, BillNo&
  Dim dlm$
  Dim RealPropRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRealRecs As Long
  Dim PersPropRec As PersonalRecType
  Dim PHandle As Integer
  Dim NumOfPersRecs As Long
  Dim LA$
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim FBill As Long
  Dim LBill As Long
  Dim CustCSZ$
  Dim PinNum$
  Dim Desc As String * 29
  Dim Name As String * 28
  Dim ZipRec As BillPrintZipIdxType
  Dim ZHandle As Integer
  Dim NumOfZRecs As Long
  Dim MortRec As BillPrintMortIdxType
  Dim MRHandle As Integer
  Dim NumOfMRRecs As Long
  Dim CustArr As Long
  
  dlm$ = "~"
  FBill = CLng(fpDblSnglFirstBill.Value)
  LBill = CLng(fpDblSnglLastBill.Value)
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close
  
  RptFile$ = "TAXRPTS\TXLSRLEGAL.RPT"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle
  
  OpenTaxBillFile TBHandle, NumOfTBRecs
  OpenRealPropFile RHandle, NumOfRealRecs
  OpenPersPropFile PHandle, NumOfPersRecs
  
  If Exist("MORTIDX.DAT") Then '12/6/06
    OpenMortIdxFile MRHandle, NumOfMRRecs
    NumOfTBRecs = NumOfMRRecs
  ElseIf Exist("ZIPIDX.DAT") Then '12/6/06
    OpenZipIdxFile ZHandle, NumOfZRecs
    NumOfTBRecs = NumOfZRecs
  End If
  
  For x = 1 To NumOfTBRecs
    If NumOfMRRecs > 0 Then  '12/6/06
      Get MRHandle, x, MortRec
      CustArr = MortRec.TaxBillRec
    ElseIf NumOfZRecs > 0 Then '12/6/06
      Get ZHandle, x, ZipRec
      CustArr = ZipRec.TaxBillRec
    Else
      CustArr = x '12/6/06
    End If
    Get TBHandle, CustArr, TaxBill
      If TaxBill.BillNumber >= FBill And TaxBill.BillNumber <= LBill Then
        If TaxBill.LotOrAcre = "A" Then
          LA = "Acre"
        ElseIf TaxBill.LotOrAcre = "L" Then
          LA = "Lot"
        Else
          LA = "NA"
        End If
        If QPTrim$(TaxBill.LASize) <> "" Then
          LA = LA + "  " + "Parcel Size: " + CStr(TaxBill.LASize)
        End If
        TaxBill.CustZip = InsertZipDash(QPTrim$(TaxBill.CustZip))
        CustCSZ = QPTrim$(TaxBill.CustAdd3) + " " + QPTrim$(TaxBill.CustZip)
        If QPTrim$(TaxBill.RealPin) <> "" Then
          PinNum = QPTrim$(TaxBill.RealPin)
        ElseIf QPTrim$(TaxBill.PersPin) <> "" Then
          PinNum = QPTrim$(TaxBill.PersPin)
        Else
          PinNum = ""
        End If
        Desc = QPTrim$(QPTrim$(TaxBill.RDesc1))
        Name = QPTrim$(TaxBill.CustName)
        '                              0                      1                        2               3
        Print #RptHandle, CStr(TaxBill.TaxYear); dlm; TaxBill.BillNumber; dlm; TaxBill.CustRec; dlm; PinNum; dlm;
        '                   4          5        6               7                         8
        Print #RptHandle, Name; dlm; Desc; dlm; LA; dlm; TaxBill.RealValue; dlm; TaxBill.PersValue; dlm;
        '                        9                                       10                                     11
        Print #RptHandle, TaxBill.ExptValue; dlm; OldRound(TaxBill.PersValue + TaxBill.RealValue); dlm; TaxBill.RealTaxRate; dlm;
        '                                     12                                         13                        14
        Print #RptHandle, OldRound(TaxBill.RealTaxDue + TaxBill.PersTaxDue); dlm; TaxBill.LateTaxDue; dlm; TaxBill.TotalBillDue - TaxBill.OverPayAmt; dlm; 'added OverPayAmt 8/15/06
        '                         15                           16                      17
'        Print #RptHandle, QPTrim$(TaxMasterRec.Name); dlm; QPTrim$(TaxMasterRec.Add1); dlm; QPTrim$(TaxMasterRec.City) + ", " + QPTrim$(TaxMasterRec.TownState) + " " + QPTrim$(TaxMasterRec.Zip); dlm;
        Print #RptHandle, "                       "; dlm; "                  "; dlm; "                     "; dlm;
        '                  18                                     19                              20                     21
        Print #RptHandle, Name; dlm; QPTrim$(TaxBill.CustAdd1); dlm; QPTrim$(TaxBill.CustAdd2); dlm; CustCSZ
     End If
  Next x
  
  Close
  
  arTaxLsrLegal.Show
  
End Sub

Private Sub PrintLSRLEGALHP()
  Dim RptHandle As Integer
  Dim RptFile$
  Dim TaxBill As TaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim x As Long, BillNo&
  Dim dlm$
  Dim RealPropRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRealRecs As Long
  Dim PersPropRec As PersonalRecType
  Dim PHandle As Integer
  Dim NumOfPersRecs As Long
  Dim LA$
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim FBill As Long
  Dim LBill As Long
  Dim Desc As String * 29
  Dim Name As String * 28
  Dim CustCSZ$
  Dim PinNum$
  Dim ZipRec As BillPrintZipIdxType
  Dim ZHandle As Integer
  Dim NumOfZRecs As Long
  Dim MortRec As BillPrintMortIdxType
  Dim MRHandle As Integer
  Dim NumOfMRRecs As Long
  Dim CustArr As Long '12/6/06
  
  dlm$ = "~"
  FBill = CLng(fpDblSnglFirstBill.Value)
  LBill = CLng(fpDblSnglLastBill.Value)
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close
  
  RptFile$ = "TAXRPTS\TXLSRLEGALHP.RPT"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle
  
  OpenTaxBillFile TBHandle, NumOfTBRecs
  If Exist("MORTIDX.DAT") Then '12/6/06
    OpenMortIdxFile MRHandle, NumOfMRRecs
    NumOfTBRecs = NumOfMRRecs
  ElseIf Exist("ZIPIDX.DAT") Then '12/6/06
    OpenZipIdxFile ZHandle, NumOfZRecs
    NumOfTBRecs = NumOfZRecs
  End If
  
  OpenRealPropFile RHandle, NumOfRealRecs
  OpenPersPropFile PHandle, NumOfPersRecs
  
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
      If TaxBill.BillNumber >= FBill And TaxBill.BillNumber <= LBill Then
        If TaxBill.LotOrAcre = "A" Then
          LA = "Acre"
        ElseIf TaxBill.LotOrAcre = "L" Then
          LA = "Lot"
        Else
          LA = "NA"
        End If
        If QPTrim$(TaxBill.LASize) <> "" Then
          LA = LA + "  " + "Parcel Size: " + CStr(TaxBill.LASize)
        End If
        Desc = QPTrim$(TaxBill.RDesc1)
        Name = QPTrim$(TaxBill.CustName)
        TaxBill.CustZip = InsertZipDash(QPTrim$(TaxBill.CustZip))
        CustCSZ = QPTrim$(TaxBill.CustAdd3) + " " + QPTrim$(TaxBill.CustZip)
        If QPTrim$(TaxBill.RealPin) <> "" Then
          PinNum = QPTrim$(TaxBill.RealPin)
        ElseIf QPTrim$(TaxBill.PersPin) <> "" Then
          PinNum = QPTrim$(TaxBill.PersPin)
        Else
          PinNum = ""
        End If
        '                              0                      1                        2                3
        Print #RptHandle, CStr(TaxBill.TaxYear); dlm; TaxBill.BillNumber; dlm; TaxBill.CustRec; dlm; PinNum; dlm;
        '                   4          5        6                 7                         8
        Print #RptHandle, Name; dlm; Desc; dlm; LA; dlm; TaxBill.RealValue; dlm; TaxBill.PersValue; dlm;
        '                        9                                       10                                     11
        Print #RptHandle, TaxBill.ExptValue; dlm; OldRound(TaxBill.PersValue + TaxBill.RealValue); dlm; TaxBill.RealTaxRate; dlm;
        '                                     12                                         13                                  14
        Print #RptHandle, OldRound(TaxBill.RealTaxDue + TaxBill.PersTaxDue); dlm; TaxBill.LateTaxDue; dlm; TaxBill.TotalBillDue - TaxBill.OverPayAmt; dlm; 'added OverPayAmt 8/15/06
        '                         15                           16                      17
        Print #RptHandle, "                       "; dlm; "                  "; dlm; "                     "; dlm;
        '                   18                     19                                20                  21
        Print #RptHandle, Name; dlm; QPTrim$(TaxBill.CustAdd1); dlm; QPTrim$(TaxBill.CustAdd2); dlm; CustCSZ
     End If
  Next x
  
  Close
  
  arTaxLsrLegalHP.Show
  
End Sub

