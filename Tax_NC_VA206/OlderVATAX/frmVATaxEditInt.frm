VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmVATaxEditInt 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Editing Interest Transaction"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmVATaxEditInt.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11640
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcmbType 
      Height          =   405
      Left            =   4140
      TabIndex        =   23
      Top             =   1800
      Width           =   3360
      _Version        =   196608
      _ExtentX        =   5927
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
      ColDesigner     =   "frmVATaxEditInt.frx":08CA
   End
   Begin EditLib.fpCurrency fpCurrInt 
      Height          =   372
      Left            =   5748
      TabIndex        =   11
      Top             =   6096
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
   Begin EditLib.fpLongInteger fpLongAcctNum 
      Height          =   396
      Left            =   5340
      TabIndex        =   0
      Top             =   3456
      Width           =   1572
      _Version        =   196608
      _ExtentX        =   2773
      _ExtentY        =   698
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
      AlignTextH      =   1
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
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   "0"
      MaxValue        =   "2147483647"
      MinValue        =   "0"
      NegFormat       =   1
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
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
      Height          =   372
      Left            =   4284
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4056
      Width           =   4092
      _Version        =   196608
      _ExtentX        =   7223
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
      AlignTextH      =   1
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   2
      AutoAdvance     =   -1  'True
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
   Begin EditLib.fpText fptxtRecord 
      Height          =   396
      Left            =   5700
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2520
      Width           =   2172
      _Version        =   196608
      _ExtentX        =   3836
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
   Begin EditLib.fpDoubleSingle fpDblSnglStartBill 
      Height          =   372
      Left            =   5640
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5268
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
      AlignTextH      =   1
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
   Begin EditLib.fpText fptxtRTaxYear 
      Height          =   396
      Left            =   3960
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   4680
      Width           =   1212
      _Version        =   196608
      _ExtentX        =   2143
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
      CharValidationText=   "1 2 3 4 5 6 7 8 9 0"
      MaxLength       =   4
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
   Begin fpBtnAtlLibCtl.fpBtn cmdLookup 
      Height          =   372
      Left            =   7080
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3480
      Width           =   1812
      _Version        =   131072
      _ExtentX        =   3196
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
      ButtonDesigner  =   "frmVATaxEditInt.frx":0BC1
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSave 
      Height          =   495
      Left            =   8250
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   7080
      Width           =   1680
      _Version        =   131072
      _ExtentX        =   2963
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
      ButtonDesigner  =   "frmVATaxEditInt.frx":0DA3
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   1092
      Left            =   1800
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   7080
      Width           =   1692
      _Version        =   131072
      _ExtentX        =   2984
      _ExtentY        =   1926
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
      ButtonDesigner  =   "frmVATaxEditInt.frx":0F7F
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdDelete 
      Height          =   495
      Left            =   8250
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   7680
      Width           =   1680
      _Version        =   131072
      _ExtentX        =   2963
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
      ButtonDesigner  =   "frmVATaxEditInt.frx":115B
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPrev 
      Height          =   495
      Left            =   6090
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   7080
      Width           =   1680
      _Version        =   131072
      _ExtentX        =   2963
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
      ButtonDesigner  =   "frmVATaxEditInt.frx":1338
      Begin fpBtnAtlLibCtl.fpBtn fpBtn1 
         Height          =   495
         Left            =   360
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   840
         Visible         =   0   'False
         Width           =   9990
         _Version        =   131072
         _ExtentX        =   17621
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
         ButtonDesigner  =   "frmVATaxEditInt.frx":1513
      End
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdNext 
      Height          =   495
      Left            =   6090
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   7680
      Width           =   1680
      _Version        =   131072
      _ExtentX        =   2963
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
      ButtonDesigner  =   "frmVATaxEditInt.frx":16EF
      Begin fpBtnAtlLibCtl.fpBtn fpBtn3 
         Height          =   9996
         Left            =   3996
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   2004
         Visible         =   0   'False
         Width           =   2064
         _Version        =   131072
         _ExtentX        =   3641
         _ExtentY        =   17632
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
         ButtonDesigner  =   "frmVATaxEditInt.frx":18C6
      End
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdHome 
      Height          =   495
      Left            =   3930
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   7080
      Width           =   1680
      _Version        =   131072
      _ExtentX        =   2963
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
      ButtonDesigner  =   "frmVATaxEditInt.frx":1AA2
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdEnd 
      Height          =   495
      Left            =   3930
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   7680
      Width           =   1680
      _Version        =   131072
      _ExtentX        =   2963
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
      ButtonDesigner  =   "frmVATaxEditInt.frx":1C79
   End
   Begin EditLib.fpText fptxtPTaxYear 
      Height          =   396
      Left            =   8280
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   4680
      Width           =   1212
      _Version        =   196608
      _ExtentX        =   2143
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
      CharValidationText=   "1 2 3 4 5 6 7 8 9 0"
      MaxLength       =   4
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
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Personal Tax Year:"
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
      Left            =   5640
      TabIndex        =   26
      Top             =   4728
      Width           =   2532
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   1560
      X2              =   10200
      Y1              =   5856
      Y2              =   5856
   End
   Begin VB.Label Label27 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Select A Billing Type"
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
      Height          =   276
      Left            =   4854
      TabIndex        =   24
      Top             =   1440
      Width           =   1980
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   3492
      Left            =   1560
      Top             =   3216
      Width           =   8652
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Interest:"
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
      Left            =   4560
      TabIndex        =   10
      Top             =   6192
      Width           =   972
   End
   Begin VB.Label Label6 
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
      Left            =   2040
      TabIndex        =   9
      Top             =   4728
      Width           =   1812
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Bill No:"
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
      Left            =   4560
      TabIndex        =   8
      Top             =   5340
      Width           =   972
   End
   Begin VB.Label Label4 
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
      Height          =   276
      Left            =   3660
      TabIndex        =   6
      Top             =   2628
      Width           =   1860
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Acct Number:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   2808
      TabIndex        =   4
      Top             =   3576
      Width           =   2412
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   3264
      TabIndex        =   3
      Top             =   4188
      Width           =   852
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   660
      Index           =   1
      Left            =   1500
      Top             =   576
      Width           =   8652
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Edit Interest Transaction"
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
      Left            =   3120
      TabIndex        =   1
      Top             =   732
      Width           =   5292
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   780
      Left            =   1500
      Top             =   468
      Width           =   8652
   End
End
Attribute VB_Name = "frmVATaxEditInt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Public GIntRec As Long
  Dim ThisManyRRecs As Long
  Dim ThisManyPRecs As Long
  'Private Temp_Class As Resize_Class
  Dim ThisInt As Double
  Dim PrevOK As Boolean
  Dim NextOK As Boolean
  Dim ExitOK As Boolean
  Dim FromTypeChng As Boolean

Private Sub cmdDelete_Click()
  Dim IntRec As InterestRecType
  Dim NumOfIRRecs As Long
  Dim IRHandle As Integer
  
  If GIntRec = 0 Then Exit Sub
  If fpcmbType.Text = "REAL" Then
    OpenRInterestRecFile IRHandle, NumOfIRRecs
  ElseIf fpcmbType.Text = "PERSONAL" Then
    OpenPInterestRecFile IRHandle, NumOfIRRecs
  ElseIf fpcmbType.Text = "NO PERSONAL" Then
    Call TaxMsg(900, "No personal interest calculations have been saved.")
  ElseIf fpcmbType.Text = "NO REAL" Then
    Call TaxMsg(900, "No real interest calculations have been saved.")
  End If
  Get IRHandle, GIntRec, IntRec
  IntRec.DelFlag = True
  Put IRHandle, GIntRec, IntRec
  Close IRHandle
  ThisManyRRecs = ThisManyRRecs - 1
  
  Call TaxMsg(900, "This record has been deleted successfully.")
  
  Call Clearscreen
  
End Sub

Private Sub cmdEnd_Click()
  Dim IntRec As InterestRecType
  Dim NumOfIRRecs As Long
  Dim IRHandle As Integer
  
  If ThisInt <> CDbl(fpCurrInt.Value) Then
    If TaxMsgWOpts(900, "Do you wish to save your data?", "F10 Save", "ESC Don't Save") = "abort" Then
      Unload frmVATaxMsgWOpts
    Else
      Unload frmVATaxMsgWOpts
      NextOK = True
      Call cmdSave_Click
    End If
  End If
  
  If GIntRec = ThisManyRRecs Then Exit Sub
  
  If fpcmbType.Text = "REAL" Then
    GIntRec = ThisManyRRecs
    OpenRInterestRecFile IRHandle, NumOfIRRecs
  ElseIf fpcmbType.Text = "PERSONAL" Then
    GIntRec = ThisManyPRecs
    OpenPInterestRecFile IRHandle, NumOfIRRecs
  ElseIf fpcmbType.Text = "NO PERSONAL" Then
    Call TaxMsg(900, "No personal interest calculations have been saved.")
  ElseIf fpcmbType.Text = "NO REAL" Then
    Call TaxMsg(900, "No real interest calculations have been saved.")
  End If
  Get IRHandle, GIntRec, IntRec
  GCustNum = IntRec.CustRec
  Close IRHandle
  
  Call LoadMeEdit
  
End Sub

Private Sub cmdExit_Click()
  Dim IntRec As InterestRecType
  Dim NumOfIRRecs As Long
  Dim IRHandle As Integer
  
  If ThisInt <> CDbl(fpCurrInt.Value) Then
    If TaxMsgWOpts(900, "Do you wish to save your data?", "F10 Save", "ESC Don't Save") = "abort" Then
      Unload frmVATaxMsgWOpts
    Else
      Unload frmVATaxMsgWOpts
      ExitOK = True
      Call cmdSave_Click
    End If
  End If
  
  ExitOK = True
  GCustNum = 0
  GIntRec = 0
  ThisInt = 0
  frmVATaxInterestMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdHome_Click()
  Dim IntRec As InterestRecType
  Dim NumOfIRRecs As Long
  Dim IRHandle As Integer
  
  If ThisInt <> CDbl(fpCurrInt.Value) Then
    If TaxMsgWOpts(900, "Do you wish to save your data?", "F10 Save", "ESC Don't Save") = "abort" Then
      Unload frmVATaxMsgWOpts
    Else
      Unload frmVATaxMsgWOpts
      NextOK = True
      Call cmdSave_Click
    End If
  End If
  
  If GIntRec = 1 Then Exit Sub
  GIntRec = 1
  If fpcmbType.Text = "REAL" Then
    OpenRInterestRecFile IRHandle, NumOfIRRecs
  ElseIf fpcmbType.Text = "PERSONAL" Then
    OpenPInterestRecFile IRHandle, NumOfIRRecs
  ElseIf fpcmbType.Text = "NO PERSONAL" Then
    Call TaxMsg(900, "No personal interest calculations have been saved.")
  ElseIf fpcmbType.Text = "NO REAL" Then
    Call TaxMsg(900, "No real interest calculations have been saved.")
  End If
  Get IRHandle, GIntRec, IntRec
  GCustNum = IntRec.CustRec
  Close IRHandle
  
  Call LoadMeEdit
End Sub

Private Sub cmdLookup_Click()
  If ThisInt <> CDbl(fpCurrInt.Value) Then
    If TaxMsgWOpts(900, "Do you wish to save your data?", "F10 Save", "ESC Don't Save") = "abort" Then
      Unload frmVATaxMsgWOpts
    Else
      Unload frmVATaxMsgWOpts
      Call cmdSave_Click
    End If
  End If
  If fpcmbType.Text = "REAL" Then
    frmVATaxCustLookupIntOnly.Show
  ElseIf fpcmbType.Text = "PERSONAL" Then
    frmVATaxCustLookUpPIntOnly.Show
  ElseIf fpcmbType.Text = "NO PERSONAL" Then
    Call TaxMsg(900, "No personal interest calculations have been saved.")
  ElseIf fpcmbType.Text = "NO REAL" Then
    Call TaxMsg(900, "No real interest calculations have been saved.")
  End If
  DoEvents
End Sub

Private Sub cmdNext_Click()
  Dim IntRec As InterestRecType
  Dim NumOfIRRecs As Long
  Dim IRHandle As Integer
  
  If ThisInt <> CDbl(fpCurrInt.Value) Then
    If TaxMsgWOpts(900, "Do you wish to save your data?", "F10 Save", "ESC Don't Save") = "abort" Then
      Unload frmVATaxMsgWOpts
    Else
      Unload frmVATaxMsgWOpts
      NextOK = True
      Call cmdSave_Click
    End If
  End If
  
  If GIntRec = ThisManyRRecs Then Exit Sub
  GIntRec = GIntRec + 1
  If fpcmbType.Text = "REAL" Then
    OpenRInterestRecFile IRHandle, NumOfIRRecs
    Get IRHandle, GIntRec, IntRec
    GCustNum = IntRec.CustRec
    Close IRHandle
  ElseIf fpcmbType.Text = "PERSONAL" Then
    OpenPInterestRecFile IRHandle, NumOfIRRecs
    Get IRHandle, GIntRec, IntRec
    GCustNum = IntRec.CustRec
    Close IRHandle
  ElseIf fpcmbType.Text = "NO PERSONAL" Then
    Call TaxMsg(900, "No personal interest calculations have been saved.")
  ElseIf fpcmbType.Text = "NO REAL" Then
    Call TaxMsg(900, "No real interest calculations have been saved.")
  End If
  Call LoadMeEdit

End Sub

Private Sub cmdPrev_Click()
  Dim IntRec As InterestRecType
  Dim NumOfIRRecs As Long
  Dim IRHandle As Integer
  
  If ThisInt <> CDbl(fpCurrInt.Value) Then
    If TaxMsgWOpts(900, "Do you wish to save your data?", "F10 Save", "ESC Don't Save") = "abort" Then
      Unload frmVATaxMsgWOpts
    Else
      Unload frmVATaxMsgWOpts
      PrevOK = True
      Call cmdSave_Click
    End If
  End If
  
  If GIntRec <= 1 Then Exit Sub
  GIntRec = GIntRec - 1
  If fpcmbType.Text = "REAL" Then
    OpenRInterestRecFile IRHandle, NumOfIRRecs
  ElseIf fpcmbType.Text = "PERSONAL" Then
    OpenPInterestRecFile IRHandle, NumOfIRRecs
  ElseIf fpcmbType.Text = "NO PERSONAL" Then
    Call TaxMsg(900, "No personal interest calculations have been saved.")
  ElseIf fpcmbType.Text = "NO REAL" Then
    Call TaxMsg(900, "No real interest calculations have been saved.")
  End If
  Get IRHandle, GIntRec, IntRec
  GCustNum = IntRec.CustRec
  Close IRHandle
  
  Call LoadMeEdit
End Sub

Private Sub cmdSave_Click()
  Dim IntRec As InterestRecType
  Dim NumOfIRRecs As Long
  Dim IRHandle As Integer
  
  If GIntRec = 0 Then Exit Sub
  If FromTypeChng = True Then
    If fpcmbType.Text = "REAL" Then
      OpenPInterestRecFile IRHandle, NumOfIRRecs
      Get IRHandle, GIntRec, IntRec
      IntRec.Amount = CDbl(fpCurrInt.Value)
      Put IRHandle, GIntRec, IntRec
    ElseIf fpcmbType.Text = "PERSONAL" Then
      OpenRInterestRecFile IRHandle, NumOfIRRecs
      Get IRHandle, GIntRec, IntRec
      IntRec.Amount = CDbl(fpCurrInt.Value)
      Put IRHandle, GIntRec, IntRec
    End If
  Else
    If fpcmbType.Text = "REAL" Then
      OpenRInterestRecFile IRHandle, NumOfIRRecs
      Get IRHandle, GIntRec, IntRec
      IntRec.Amount = CDbl(fpCurrInt.Value)
      Put IRHandle, GIntRec, IntRec
    ElseIf fpcmbType.Text = "PERSONAL" Then
      OpenPInterestRecFile IRHandle, NumOfIRRecs
      Get IRHandle, GIntRec, IntRec
      IntRec.Amount = CDbl(fpCurrInt.Value)
      Put IRHandle, GIntRec, IntRec
    End If
  End If
  Close IRHandle
  
  Call Savemsg(900, "Your data has been saved successfully.")
  If PrevOK = True Then
    PrevOK = False
    Exit Sub
  ElseIf NextOK = True Then
    NextOK = False
    Exit Sub
  End If
  
  Call Clearscreen
  
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
      SendKeys "%S"
      Call cmdSave_Click
      KeyCode = 0
    Case vbKeyF7:
      SendKeys "%L"
      Call cmdLookup_Click
      KeyCode = 0
    Case vbKeyF3:
      SendKeys "%D"
      Call cmdDelete_Click
      KeyCode = 0
    Case vbKeyHome:
      Call cmdHome_Click
      KeyCode = 0
    Case vbKeyEnd:
      Call cmdEnd_Click
      KeyCode = 0
    Case vbKeyPageUp:
      Call cmdNext_Click
      KeyCode = 0
    Case vbKeyPageDown:
      Call cmdPrev_Click
      KeyCode = 0
    Case Else:
  End Select

End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  'Set Temp_Class = New Resize_Class
  'Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  GIntRec = 0
  ExitOK = False
  PrevOK = False
  NextOK = False
  Me.HelpContextID = hlpEditInterest
  Call LoadMe
  FromTypeChng = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("CitiTaxes.exe terminated via menu bar on frmVATaxEditInt.")
      Call Terminate
      End
    End If
  End If

End Sub
'Private Sub Form_Resize()
'  If Me.WindowState <> vbMinimized Then
'    Me.Visible = False
'    'Temp_Class.ResizeControls Me
'    Me.Visible = True
'    Me.SetFocus
'    DoEvents
'  End If
'End Sub

Private Sub LoadMe()
  Dim IntRec As InterestRecType
  Dim NumOfIRRecs As Long
  Dim IRHandle As Integer
  Dim IntCnt As Long
  Dim PIntRec As InterestRecType
  Dim NumOfPIRRecs As Long
  Dim PIRHandle As Integer
  Dim PIntCnt As Long
  Dim x As Integer
  
  If Exist(TaxRIntFile) Then
    OpenRInterestRecFile IRHandle, NumOfIRRecs
    For x = 1 To NumOfIRRecs
      Get IRHandle, x, NumOfIRRecs
      If IntRec.DelFlag = True Then
        GoTo SkipItR
      Else
        IntCnt = IntCnt + 1
      End If
SkipItR:
    Next x
    Close IRHandle
  End If
  
  If Exist(TaxPIntFile) Then
    OpenPInterestRecFile PIRHandle, NumOfPIRRecs
    For x = 1 To NumOfPIRRecs
      Get PIRHandle, x, NumOfPIRRecs
      If PIntRec.DelFlag = True Then
        GoTo SkipItP
      Else
        PIntCnt = PIntCnt + 1
      End If
SkipItP:
    Next x
    Close PIRHandle
  End If
  
  If IntCnt > 0 And PIntCnt > 0 Then
    fpcmbType.Text = "REAL"
    fpcmbType.AddItem "REAL"
    fpcmbType.AddItem "PERSONAL"
  ElseIf IntCnt > 0 And PIntCnt = 0 Then
    fpcmbType.Text = "REAL"
    fpcmbType.AddItem "REAL"
    fpcmbType.AddItem "NO PERSONAL"
  ElseIf IntCnt = 0 And PIntCnt > 0 Then
    fpcmbType.Text = "PERSONAL"
'    fpcmbType.AddItem "REAL" '4/9/07
    fpcmbType.AddItem "NO REAL"
  End If
  
  ThisManyRRecs = IntCnt
  ThisManyPRecs = PIntCnt
  If IntCnt > 0 Then
    fptxtRecord.Text = "0 of " + CStr(ThisManyRRecs)
  ElseIf PIntCnt > 0 Then
    fptxtRecord.Text = "0 of " + CStr(ThisManyPRecs)
  End If
  fpLongAcctNum = 0
  fptxtName.Text = ""
  fptxtRTaxYear = 0
  fptxtPTaxYear = 0
  fpDblSnglStartBill = 0
  fpCurrInt = 0
  ThisInt = 0
  
End Sub

Public Sub LoadMeEdit()
  Dim IntRec As InterestRecType
  Dim NumOfIRRecs As Long
  Dim IRHandle As Integer
  Dim x As Long
  
  If GIntRec = 0 Then
    Call TaxMsg(900, "ERROR: There is a problem with the internal number assignment for this customer. Please try again.")
    fpLongAcctNum.SetFocus
    Exit Sub
  End If
  If fpcmbType.Text = "REAL" Then
    OpenRInterestRecFile IRHandle, NumOfIRRecs
    Get IRHandle, GIntRec, IntRec
    Close IRHandle
    fptxtRTaxYear = CStr(IntRec.TaxYear)
    fptxtRecord.Text = CStr(GIntRec) + " of " + CStr(ThisManyRRecs)
  ElseIf fpcmbType.Text = "PERSONAL" Then
    OpenPInterestRecFile IRHandle, NumOfIRRecs
    Get IRHandle, GIntRec, IntRec
    Close IRHandle
    fptxtPTaxYear = CStr(IntRec.TaxYear)
    fptxtRecord.Text = CStr(GIntRec) + " of " + CStr(ThisManyPRecs)
  End If
  fpLongAcctNum = IntRec.CustRec
  fptxtName.Text = QPTrim$(IntRec.CustName)
  fpDblSnglStartBill = IntRec.BillNumber
  fpCurrInt = IntRec.Amount
  ThisInt = IntRec.Amount

End Sub

Private Sub Clearscreen()
  GCustNum = 0
  GIntRec = 0
  fptxtRecord.Text = "0 of " + CStr(ThisManyRRecs)
  fpLongAcctNum = 0
  fptxtName.Text = ""
  fptxtRTaxYear = 0
  fptxtPTaxYear = 0
  fpDblSnglStartBill = 0
  fpCurrInt = 0
  ThisInt = 0

End Sub
Private Sub fpcmbType_Change()
  If ThisInt <> CDbl(fpCurrInt.Value) Then
    If TaxMsgWOpts(1000, "Do you wish to save your data?", "F10 Save", "ESC Don't Save") = "abort" Then
      Unload frmVATaxMsgWOpts
    Else
      Unload frmVATaxMsgWOpts
      FromTypeChng = True
      Call cmdSave_Click
      FromTypeChng = False
    End If
  End If
  Clearscreen
  Call GetThisManyRecs
  If fpcmbType.Text = "PERSONAL" Then
    fptxtRecord.Text = "0 of " + CStr(ThisManyPRecs)
  Else
    fptxtRecord.Text = "0 of " + CStr(ThisManyRRecs)
  End If
End Sub

Private Sub fpcmbType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbType.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbType.ListIndex = -1
  End If
  If fpcmbType.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fptxtRecord.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpCurrInt_Change()
  If fpLongAcctNum.Value = 0 Then
    fpCurrInt.Value = 0
  End If
End Sub

Private Sub fpCurrInt_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyHome Then
    Call cmdHome_Click
  ElseIf KeyCode = vbKeyEnd Then
    Call cmdEnd_Click
  End If
End Sub

Private Sub fpLongAcctNum_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyHome Then
    Call cmdHome_Click
  ElseIf KeyCode = vbKeyEnd Then
    Call cmdEnd_Click
  End If

End Sub

Private Sub fpLongAcctNum_LostFocus()
  Dim ThisRec As Long
  Dim IntRec As InterestRecType
  Dim NumOfIRRecs As Long
  Dim IRHandle As Integer
  Dim x As Long
  
  On Error GoTo ERRORSTUFF
  
  If ExitOK = True Then Exit Sub
  ThisRec = CLng(fpLongAcctNum.Text)
  If ThisRec = 0 Then
    Exit Sub
  End If
  
  If ThisRec = GCustNum Then
    Exit Sub
  End If
  
  If ThisInt <> CDbl(fpCurrInt.Value) Then
    If TaxMsgWOpts(900, "Do you wish to save your data?", "F10 Save", "ESC Don't Save") = "abort" Then
      Unload frmVATaxMsgWOpts
    Else
      Unload frmVATaxMsgWOpts
      Call cmdSave_Click
    End If
  End If
  
  If fpcmbType.Text = "REAL" Then
    OpenRInterestRecFile IRHandle, NumOfIRRecs
    For x = 1 To NumOfIRRecs
      Get IRHandle, x, IntRec
      If IntRec.DelFlag = True Then GoTo SkipItR
      If IntRec.CustRec = CLng(fpLongAcctNum.Text) Then
        Exit For
      End If
SkipItR:
    Next x
  ElseIf fpcmbType.Text = "PERSONAL" Then
    OpenPInterestRecFile IRHandle, NumOfIRRecs
    For x = 1 To NumOfIRRecs
      Get IRHandle, x, IntRec
      If IntRec.DelFlag = True Then GoTo SkipItP
      If IntRec.CustRec = CLng(fpLongAcctNum.Text) Then
        Exit For
      End If
SkipItP:
    Next x
  End If
  Close IRHandle
  
  If x > NumOfIRRecs Then
    Call TaxMsg(800, "The customer number entered could not be found in the interest calculation records. Please try another number.")
    Call Clearscreen
    fpLongAcctNum.Text = ThisRec
    fpLongAcctNum.SetFocus
    Exit Sub
  Else
    GCustNum = IntRec.CustRec
    GIntRec = x
    Call LoadMeEdit
  End If
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxEditInt", "fpLongAcctNum_LostFocus", Erl)
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

Private Sub GetThisManyRecs()
  Dim IntRec As InterestRecType
  Dim NumOfIRRecs As Long
  Dim IRHandle As Integer
  
  If fpcmbType.Text = "REAL" Then
    OpenRInterestRecFile IRHandle, NumOfIRRecs
    ThisManyRRecs = NumOfIRRecs
  ElseIf fpcmbType.Text = "PERSONAL" Then
    OpenPInterestRecFile IRHandle, NumOfIRRecs
    ThisManyRRecs = NumOfIRRecs
  End If
  
  Close IRHandle
End Sub
