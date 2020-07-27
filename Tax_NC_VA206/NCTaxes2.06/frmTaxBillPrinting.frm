VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmTaxBillPrinting 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Bill Printing Information"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmTaxBillPrinting.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcmbPrintPriorYN 
      Height          =   390
      Left            =   9840
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   4680
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
      ColDesigner     =   "frmTaxBillPrinting.frx":08CA
   End
   Begin LpLib.fpCombo fpcmbBarcode 
      Height          =   390
      Left            =   6690
      TabIndex        =   0
      ToolTipText     =   "For a bar code to appear on a bill the tax customer must have a valid 9 digit zip code and a delivery point value."
      Top             =   1920
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
      ColDesigner     =   "frmTaxBillPrinting.frx":0CA5
   End
   Begin VB.OptionButton OptMort 
      BackColor       =   &H008F8265&
      Caption         =   "Print By Mortgage Code"
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
      Height          =   255
      Left            =   4440
      TabIndex        =   27
      Top             =   6600
      Width           =   2415
   End
   Begin VB.OptionButton optZip 
      BackColor       =   &H008F8265&
      Caption         =   "Print By Zip Code"
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
      Height          =   255
      Left            =   7200
      TabIndex        =   26
      Top             =   6600
      Width           =   1815
   End
   Begin VB.OptionButton optLeave 
      BackColor       =   &H008F8265&
      Caption         =   "Print Unsorted"
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
      Height          =   255
      Left            =   2640
      TabIndex        =   25
      Top             =   6600
      Width           =   1695
   End
   Begin EditLib.fpText fptxtOrder 
      Height          =   375
      Left            =   5400
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5985
      Width           =   2655
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
   Begin EditLib.fpDoubleSingle fpDblSnglStartBill 
      Height          =   375
      Left            =   6000
      TabIndex        =   1
      Top             =   3105
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
   Begin EditLib.fpLongInteger fpLongTaxYear 
      Height          =   375
      Left            =   5880
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2625
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
      Height          =   375
      Left            =   6360
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4065
      Width           =   1335
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
      Height          =   375
      Left            =   6360
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   4665
      Width           =   1335
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
      Height          =   375
      Left            =   6000
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5505
      Width           =   1215
      _Version        =   196608
      _ExtentX        =   2143
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
      Text            =   "0"
      DecimalPlaces   =   -1
      DecimalPoint    =   ""
      FixedPoint      =   0   'False
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
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   7800
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
      ButtonDesigner  =   "frmTaxBillPrinting.frx":1080
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   492
      Left            =   2400
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   7800
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
      ButtonDesigner  =   "frmTaxBillPrinting.frx":125F
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdAlign 
      Height          =   492
      Left            =   4800
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   7800
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
      ButtonDesigner  =   "frmTaxBillPrinting.frx":143B
   End
   Begin EditLib.fpText fptxtCurrForm 
      Height          =   396
      Left            =   5532
      TabIndex        =   18
      TabStop         =   0   'False
      ToolTipText     =   "Late notices are selected on the System Setup screen."
      Top             =   1440
      Width           =   2856
      _Version        =   196608
      _ExtentX        =   5038
      _ExtentY        =   698
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
   Begin EditLib.fpDateTime fptxtDueDate 
      Height          =   375
      Left            =   5640
      TabIndex        =   21
      TabStop         =   0   'False
      Tag             =   $"frmTaxBillPrinting.frx":161C
      Top             =   7200
      Visible         =   0   'False
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
      AutoAdvance     =   0   'False
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
   Begin VB.Line Line1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   2400
      X2              =   9240
      Y1              =   6480
      Y2              =   6480
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
      Left            =   9480
      TabIndex        =   24
      Top             =   3720
      Width           =   1665
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Use Bar Code (Y/N?):"
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
      Height          =   300
      Left            =   4050
      TabIndex        =   22
      Top             =   2040
      Width           =   2388
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Due Date:"
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
      Left            =   3840
      TabIndex        =   20
      Top             =   7260
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Format In Use:"
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
      Height          =   300
      Left            =   3252
      TabIndex        =   19
      Top             =   1500
      Width           =   2028
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
      Height          =   375
      Left            =   3480
      TabIndex        =   16
      Top             =   6060
      Width           =   1695
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   1575
      Left            =   3600
      Top             =   3705
      Width           =   4455
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   4440
      Left            =   2400
      Top             =   2520
      Width           =   6855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Bill Printing Information"
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
      Left            =   3120
      TabIndex        =   13
      Top             =   750
      Width           =   5295
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   660
      Index           =   1
      Left            =   1500
      Top             =   585
      Width           =   8655
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
      Left            =   4680
      TabIndex        =   12
      Top             =   2700
      Width           =   1095
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
      Height          =   375
      Left            =   4080
      TabIndex        =   11
      Top             =   3180
      Width           =   1815
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
      Height          =   255
      Left            =   4320
      TabIndex        =   10
      Top             =   4110
      Width           =   1695
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
      Height          =   375
      Left            =   3960
      TabIndex        =   9
      Top             =   4710
      Width           =   2175
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
      Height          =   255
      Left            =   4440
      TabIndex        =   8
      Top             =   5580
      Width           =   1455
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
      Height          =   255
      Left            =   6480
      TabIndex        =   7
      Top             =   3780
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   780
      Left            =   1500
      Top             =   480
      Width           =   8655
   End
End
Attribute VB_Name = "frmTaxBillPrinting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class
  Dim BillFormat As Long
  Dim Opt1Desc$
  Dim Opt2Desc$
  Dim Opt3Desc$

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
  ElseIf BillFormat = 21837 Then 'standard
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
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxBillPrinting", "cmdAlign_Click", Erl)
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
  Close
  frmTaxBillPrintMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdProcess_Click()
  Dim BillInfo As TaxBillInfoType
  Dim BIHandle As Integer
  
  If fpDblSnglStartBill <= 0 Then
    Call TaxMsg(900, "Please enter a valid bill number.")
    fpDblSnglStartBill.SetFocus
    Exit Sub
  End If
  
  If optLeave.Value = True Then
    KillFile "ZIPIDX.DAT"
    KillFile "MORTIDX.DAT"
  End If
  
  If OptMort.Value = True Then
    If Not Exist("TAXMORT.DAT") Then
      Call TaxMsg(900, "No mortgage codes are saved. Attempt to index by mortgage codes is aborted.")
      OptMort.Value = False
      optLeave.Value = True
      GoTo NoMortCodes
    End If
    If Not Exist("MORTIDX.DAT") Then
      Call SortByMortCode
    End If
  End If
  
  If optZip.Value = True Then
    If Not Exist("ZIPIDX.DAT") Then
      Call SortByZipCode
    End If
  End If
  
NoMortCodes:
  Select Case fptxtCurrForm.Text
    Case "POSTCARD"
      Call TaxMsg(900, "12 Pitch is recommended for this form.")
      Call PrintPostCard1
    Case "LASER"
      Call PrintLaser1
    Case "MULTI-PART" 'standard
      Call PrintMulti
    Case "HMLT24TF"
      Call TaxMsg(800, "12 Pitch is recommended for this form. Please be sure to always start on bill form #1 and not on bill form #2.")
      Call PrintHMLT24TF
    Case "PH24TF"
      Call TaxMsg(900, "10 Pitch is recommended for this form.")
      Call PrintPH24TF
    Case "SYL23TF"
      Call TaxMsg(900, "12 Pitch is recommended for this form.")
      Call PrintSYL23TF
    Case "BSC32TF"
      Call TaxMsg(900, "12 Pitch is recommended for this form.")
      Call PrintBSC32TF
    Case "LLN21TF"
      Call TaxMsg(900, "12 Pitch is recommended for this form.")
      Call PrintLLN21TF
    Case "LASER LEGAL"
      Call PrintLaserLegal
    Case "LASER LEGAL HP"
      Call PrintLaserLegalHP
    Case "EXPORT REAL"
      Call PrintExpReal
    Case "EXPORT PERSONAL"
      Call PrintExpPers
    Case Else
      Call TaxMsg(700, "The current bill format, " + QPTrim$(fptxtCurrForm.Text) + ", is not set up for bill printing at this time. Please select a different format from the Tax System Setup screen.")
      Close
      Exit Sub
  End Select
  
  OpenBillInfoFile BIHandle
  Get BIHandle, 1, BillInfo
  BillInfo.BillNum = fpDblSnglStartBill
  Put BIHandle, 1, BillInfo
  Close BIHandle
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
      MainLog ("CitiTaxes.exe terminated via menu bar on frmTaxBillPrinting.")
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
  Me.HelpContextID = hlpPrintTaxBills
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
  
  'on error goto ERRORSTUFF
  
  If fptxtCurrForm.Text = "EXPORT REAL" Or fptxtCurrForm.Text = "EXPORT PERSONAL" Then
    OptMort.Enabled = False
    optZip.Enabled = False
    optLeave.Enabled = False
    GoTo KeepGoing
  End If
  
  If Exist("MORTIDX.DAT") Then
    OptMort.Value = True
  ElseIf Exist("ZIPIDX.DAT") Then
    optZip.Value = True
  Else
    optLeave.Value = True
  End If
  
KeepGoing:

  doAlign = False
  OpenBillInfoFile BIHandle
  Get BIHandle, 1, BillInfo
  Close BIHandle
  fpLongTaxYear = BillInfo.TaxYear
  If BillInfo.BillNum > 0 Then
    fpDblSnglStartBill = BillInfo.BillNum
  Else
    fpDblSnglStartBill = 0
  End If
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
  BillFormat = TaxMasterRec.TaxForm
  Opt1Desc = QPTrim$(TaxMasterRec.OptRev1)
  Opt2Desc = QPTrim$(TaxMasterRec.OptRev2)
  Opt3Desc = QPTrim$(TaxMasterRec.OptRev3)
  Label10.Visible = False
  fpcmbBarcode.Visible = False
  fpcmbBarcode.Text = "No"
  fpcmbBarcode.AddItem "No"
  fpcmbBarcode.AddItem "Yes"
'  BillFormat$ = Left$(TaxMasterRec.TaxForm, 1)
  If TaxMasterRec.TaxForm = 16716 Or TaxMasterRec.TaxForm = 20007 Then
    cmdAlign.Enabled = False
  End If
  
  fpcmbPrintPriorYN.Visible = False
  Label11.Visible = False
  
  Select Case TaxMasterRec.TaxForm
    Case 21837
      fptxtCurrForm.Text = "MULTI-PART" 'standard
    Case 20304
      fptxtCurrForm.Text = "POSTCARD"
    Case 16716
      fptxtCurrForm.Text = "LASER"
      Label10.Visible = True
      fpcmbBarcode.Visible = True
      fpcmbPrintPriorYN.Text = "N"
      fpcmbPrintPriorYN.AddItem "Y"
      fpcmbPrintPriorYN.AddItem "N"
      fpcmbPrintPriorYN.Visible = True
      Label11.Visible = True
    Case 29999
      fptxtCurrForm.Text = "EXPORT COMBINED"
      OptMort.Enabled = False
      optLeave.Enabled = False
      optZip.Enabled = False
    Case 20000
      fptxtCurrForm.Text = "EXPORT REAL"
      OptMort.Enabled = False
      optLeave.Enabled = False
      optZip.Enabled = False
    Case 20001
      fptxtCurrForm.Text = "EXPORT PERSONAL"
      fptxtDueDate.Visible = True
      fptxtDueDate.Text = Date
      Label9.Visible = True
      OptMort.Enabled = False
      optLeave.Enabled = False
      optZip.Enabled = False
    Case 20002
      fptxtCurrForm.Text = "HMLT24TF"
    Case 20003
      fptxtCurrForm.Text = "PH24TF"
    Case 20004
      fptxtCurrForm.Text = "SYL23TF"
    Case 20005
      fptxtCurrForm.Text = "BSC32TF"
    Case 20006
      fptxtCurrForm.Text = "LLN21TF"
    Case 20007
      fptxtCurrForm.Text = "LASER LEGAL"
    Case 20008
      fptxtCurrForm.Text = "LASER LEGAL HP"
    Case Else
      fptxtCurrForm.Text = "UNKNOWN"
  End Select
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxBillPrinting", "LoadMe", Erl)
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
  Print #RptHandle, Tab(11); Left$(CustName$, 21)
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
  Print #RptHandle, Tab(36); Left$(CustName$, 25)
  Print #RptHandle, Using("###,###,##0", NetTaxVal#);
  Print #RptHandle, Tab(16); Using("#.##", TAXRATE#);
  Print #RptHandle, Tab(21); Using("##,###,##0.00", OldRound#(TaxBill.TotalBillDue - TaxBill.LateTaxDue));
  Print #RptHandle, Tab(36); Left$(QPTrim$(TaxBill.CustAdd1), 25)
  Print #RptHandle, Tab(36); Left$(QPTrim$(TaxBill.CustAdd2), 25)
  Print #RptHandle, Tab(36); QPTrim$(TaxBill.CustAdd3); " "; QPTrim$(TaxBill.CustZip)
  Print #RptHandle, Tab(21); Using("##,###,##0.00", TaxBill.LateTaxDue)
  Print #RptHandle,
  Print #RptHandle, Tab(2); Using("###,##0.00", TaxBill.PriorYrBalance);
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
  If PrnCnt > 0 Then
    Call MakeFile 'Makefile creates the .dat identifying that prebilling took place
  End If

End Sub

Private Sub MakeFile()
  Dim One As Integer
  Dim AHandle As Integer
  
  KillFile "txblsprn.dat"
  One = 1
  AHandle = FreeFile
  Open "txblsprn.dat" For Output As AHandle
  Print #AHandle, One
  Close AHandle

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
  Dim dlm$, BillNo&, PrnCnt As Long
  Dim TBDRec As TxBill1DefaultsType
  Dim TBDHandle As Integer
  Dim ThisRate As Double
  Dim BZip As String
  Dim TaxCust As TaxCustType
  Dim NumOfTCRecs As Long
  Dim TCHandle As Integer
  Dim PriorBal# '8/15/06
  Dim PrintPriorYN As Boolean '8/15/06
  Dim CustArr As Long '12/6/06
  Dim ZipRec As BillPrintZipIdxType
  Dim ZHandle As Integer
  Dim NumOfZRecs As Long
  Dim MortRec As BillPrintMortIdxType
  Dim MRHandle As Integer
  Dim NumOfMRRecs As Long
  
  'on error goto ERRORSTUFF
  
  If fpcmbPrintPriorYN.Text = "N" Then
    PrintPriorYN = False
  Else
    PrintPriorYN = True
  End If
  dlm$ = "~"
  ReportFile$ = StartPath$ + "/TaxBil1.PRN"
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  BillNo& = fpDblSnglStartBill.Value
  
  frmTaxShowPctComp.Label1 = "Printing Tax Bills"
  frmTaxShowPctComp.cmdCancel.Visible = False
  frmTaxShowPctComp.Show , Me
  cmdProcess.Enabled = False
  cmdExit.Enabled = False
  EnableCloseButton Me.hwnd, False
  
  OpenTxBill1File TBDHandle
  Get #TBDHandle, 1, TBDRec
  If fpcmbBarcode.Text = "No" Then
    TBDRec.UseBarCode = False
    Put #TBDHandle, 1, TBDRec
  ElseIf fpcmbBarcode.Text = "Yes" Then
    TBDRec.UseBarCode = True
    Put #TBDHandle, 1, TBDRec
  End If
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
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  
  OpenTaxBillFile TBHandle, NumOfTBRecs
  
  If OptMort.Value = True Then '12/6/06
    OpenMortIdxFile MRHandle, NumOfMRRecs
    NumOfTBRecs = NumOfMRRecs
  ElseIf optZip.Value = True Then '12/6/06
    OpenZipIdxFile ZHandle, NumOfZRecs
    NumOfTBRecs = NumOfZRecs
  End If
  
  For x = 1 To NumOfTBRecs
    If OptMort.Value = True Then '12/6/06
      Get MRHandle, x, MortRec
      CustArr = MortRec.TaxBillRec
    ElseIf optZip.Value = True Then '12/6/06
      Get ZHandle, x, ZipRec
      CustArr = ZipRec.TaxBillRec
    Else
      CustArr = x '12/6/06
    End If
    Get TBHandle, CustArr, TBRec
    If TBRec.BillNumber > 0 Then
      If TBRec.TotalBillDue > 0 Then
        TBRec.BillNumber = BillNo&
        TBRec.BillPrinted = True
        Put TBHandle, CustArr, TBRec '12/6/06
        GoSub GetBarCodeData
        '                         0                         1
        Print #RptHandle, TBRec.BillNumber; dlm; QPTrim$(TBRec.CustName); dlm;
        '                           2                           3
        Print #RptHandle, QPTrim$(TBRec.CustAdd1); dlm; QPTrim$(TBRec.CustAdd2); dlm;
        '                             4                          5
        Print #RptHandle, QPTrim$(TBRec.CustAdd3); dlm; TBRec.CustPin; dlm;
        '                            6                            7
        Print #RptHandle, QPTrim$(TBRec.RealPin); dlm; QPTrim$(TBRec.RDesc1); dlm;
        '                        8                     9                   10
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
        '                                          11                                            12
        Print #RptHandle, OldRound(TBRec.RealValue + TBRec.PersValue - TBRec.ExptValue); dlm; ThisRate; dlm;
        If TBRec.OverPayAmt > 0 Then
          '                                        13                              14            15
          Print #RptHandle, OldRound(TBRec.TotalBillDue - TBRec.OverPayAmt); dlm; BZip; dlm; TBDRec.dologo; dlm;
          '                         16                   17                     18
          Print #RptHandle, TBRec.OptRevTax1; dlm; TBRec.OptRevTax2; dlm; TBRec.OptRevTax3; dlm;
          '                    19             20             21                   22                        23                       24                      25
          Print #RptHandle, Opt1Desc; dlm; Opt2Desc; dlm; Opt3Desc; dlm; QPTrim$(TBRec.CustZip); dlm; TBRec.LateTaxDue; dlm; TBRec.PriorYrBalance; dlm; PrintPriorYN 'TBRec.OverPayAmt
        Else
          '                        13                 14             15
          Print #RptHandle, TBRec.TotalBillDue; dlm; BZip; dlm; TBDRec.dologo; dlm;
          '                         16                   17                     18
          Print #RptHandle, TBRec.OptRevTax1; dlm; TBRec.OptRevTax2; dlm; TBRec.OptRevTax3; dlm;
          '                    19             20             21                 22                         23                     24                         25
          Print #RptHandle, Opt1Desc; dlm; Opt2Desc; dlm; Opt3Desc; dlm; QPTrim$(TBRec.CustZip); dlm; TBRec.LateTaxDue; dlm; TBRec.PriorYrBalance; dlm; PrintPriorYN ' TBRec.OverPayAmt
        End If
        BillNo& = BillNo& + 1
        PrnCnt = PrnCnt + 1
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
  Unload frmTaxShowPctComp
  cmdProcess.Enabled = True
  cmdExit.Enabled = True
  EnableCloseButton Me.hwnd, True
  
  Close
  
  ARptTempTaxBill.GetName ReportFile$
  ARptTempTaxBill.Show
  If PrnCnt > 0 Then
    Call MakeFile 'Makefile creates the .dat identifying that bill printing took place
  End If
  
  Exit Sub
  
GetBarCodeData:
  If fpcmbBarcode.Text = "No" Then
    BZip = ""
    Return
  End If
  Get TCHandle, TBRec.CustPin, TaxCust
  If Len(QPTrim$(TaxCust.Zip)) < 10 Or Len(QPTrim$(TaxCust.DeliveryPt)) <> 2 Then
    BZip = ""
  Else
    BZip = QPTrim$(TaxCust.Zip) + QPTrim$(TaxCust.DeliveryPt)
  End If
  
  Return
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxBillPrinting", "PrintLaser1", Erl)
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

Private Sub PrintMulti()
  Dim RptHandle As Integer
  Dim RptFile$
  Dim TaxBill As TaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim x As Long, BillNo&
  Dim PrnCnt&
  Dim PINTemp As String * 18
  Dim CustName$, PCnt As Long
  Dim CustArr As Long
  Dim ZipRec As BillPrintZipIdxType
  Dim ZHandle As Integer
  Dim NumOfZRecs As Long
  Dim MortRec As BillPrintMortIdxType
  Dim MRHandle As Integer
  Dim NumOfMRRecs As Long
  
  'on error goto ERRORSTUFF
  
  BillNo& = fpDblSnglStartBill.Value
  RptHandle = FreeFile
  RptFile$ = "TAXBIL.PRN"

  Open RptFile For Output As RptHandle
  
  OpenTaxBillFile TBHandle, NumOfTBRecs
  frmTaxShowPctComp.Label1 = "Printing Tax Bills"
  frmTaxShowPctComp.cmdCancel.Visible = False
  frmTaxShowPctComp.Show , Me
  cmdProcess.Enabled = False
  cmdExit.Enabled = False
  cmdAlign.Enabled = False
  EnableCloseButton Me.hwnd, False
  
  If OptMort.Value = True Then '12/6/06
    OpenMortIdxFile MRHandle, NumOfMRRecs
    NumOfTBRecs = NumOfMRRecs
  ElseIf optZip.Value = True Then '12/6/06
    OpenZipIdxFile ZHandle, NumOfZRecs
    NumOfTBRecs = NumOfZRecs
  End If
  
  For x = 1 To NumOfTBRecs
    If OptMort.Value = True Then '12/6/06
      Get MRHandle, x, MortRec
      CustArr = MortRec.TaxBillRec
    ElseIf optZip.Value = True Then '12/6/06
      Get ZHandle, x, ZipRec
      CustArr = ZipRec.TaxBillRec
    Else
      CustArr = x '12/6/06
    End If
    Get TBHandle, CustArr, TaxBill
    If TaxBill.BillNumber > 0 Then
      If TaxBill.TotalBillDue > 0 Then
        RSet PINTemp = TaxBill.RealPin
        CustName$ = QPTrim$(TaxBill.CustName)
        TaxBill.BillNumber = BillNo&
        TaxBill.BillPrinted = True
        Put TBHandle, CustArr, TaxBill
        Call PrintStandard(RptHandle, TBHandle, TaxBill, CustName, PINTemp, PrnCnt)
        BillNo& = BillNo& + 1
        PrnCnt = PrnCnt + 1
      End If
    End If
    frmTaxShowPctComp.ShowPctComp x, NumOfTBRecs
    If frmTaxShowPctComp.Out = True Then
      Close
      frmTaxShowPctComp.Out = False
      Unload frmTaxShowPctComp
      cmdProcess.Enabled = True
      cmdExit.Enabled = True
      cmdAlign.Enabled = True
      EnableCloseButton Me.hwnd, True
      Exit Sub
    End If
  Next x
  Unload frmTaxShowPctComp
  cmdProcess.Enabled = True
  cmdExit.Enabled = True
  cmdAlign.Enabled = True
  EnableCloseButton Me.hwnd, True
  Close TBHandle
  Close RptHandle
  ViewPrint RptFile$, "Tax Bill Printing", True
  If PrnCnt > 0 Then
    Call MakeFile 'Makefile creates the .dat identifying that bill printing took place
  End If
  Close
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxBillPrinting", "PrintMulti", Erl)
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

Private Sub PrintExpReal()
  Dim x As Long, y As Integer
  Dim TaxXRec As TaxBillExportRealType
  Dim TXHandle As Integer
  Dim TBRec As TaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim ReportFile As String
  Dim RptHandle As Integer
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim TaxYear$
  Dim FF10$
  Dim Map$
  Dim RealRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRRecs As Long
  Dim BillNo&
  Dim PrnCnt As Long
  Dim PastDue#
  
  'on error goto ERRORSTUFF
  PrnCnt = 0
  FF10$ = "#######.#0"
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  TaxYear$ = CStr(TaxMasterRec.TaxYear)
  ReportFile$ = "LCRE" + TaxYear + ".TXT"
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  
  BillNo& = fpDblSnglStartBill.Value
  
  OpenTaxBillFile TBHandle, NumOfTBRecs
  OpenRealPropFile RHandle, NumOfRRecs
  frmTaxShowPctComp.Label1 = "Creating Real Export Tax Bills"
  frmTaxShowPctComp.cmdCancel.Visible = False
  frmTaxShowPctComp.Show , Me
  cmdProcess.Enabled = False
  cmdExit.Enabled = False
  EnableCloseButton Me.hwnd, False
  ReDim CustDone(1 To 1) As Long
  Dim CDCnt As Integer
  
  For x = 1 To NumOfTBRecs
    Get TBHandle, x, TBRec
   ' If TBRec.CustRec = 756 Then Stop
    If TBRec.BillNumber > 0 Then
      If TBRec.TotalBillDue > 0 Then
        TBRec.BillNumber = BillNo&
        TBRec.BillPrinted = True
        Put TBHandle, x, TBRec
        BillNo& = BillNo& + 1
      End If
'      If TBRec.OverPayAmt > 0 Then
'        TBRec.TotalBillDue = OldRound(TBRec.TotalBillDue - TBRec.OverPayAmt)
'      End If
'      PastDue = GetCustBalance(TBRec.CustPin, -1)
      For y = 1 To CDCnt
        If TBRec.CustPin = CustDone(y) Then
          PastDue = 0
        End If
      Next y
      RSet TaxXRec.TaxBillNum = Using$("######0", TBRec.BillNumber)
      RSet TaxXRec.CustName = QPTrim$(TBRec.CustName)
      RSet TaxXRec.Add1 = QPTrim$(TBRec.CustAdd1)
      RSet TaxXRec.Add2 = QPTrim$(TBRec.CustAdd2)
      RSet TaxXRec.Add3 = QPTrim$(TBRec.CustAdd3) + " " + TBRec.CustZip
      RSet TaxXRec.TaxYear = Using$("###0", TBRec.TaxYear)
      RSet TaxXRec.CustAcct = Using$("####0", TBRec.CustPin)
      If TBRec.RealPropRecord > 0 Then
        Get RHandle, TBRec.RealPropRecord, RealRec
        RSet TaxXRec.MapNum = QPTrim$(RealRec.Map)
      Else
        RSet TaxXRec.MapNum = "NA"
      End If
      RSet TaxXRec.PropDesc1 = QPTrim$(TBRec.RDesc1)
      RSet TaxXRec.TAXRATE = Using$("#.#0", TBRec.RealTaxRate)
      RSet TaxXRec.LandVal = Using$(FF10, 0)
      RSet TaxXRec.BldgVal = Using$(FF10, 0)
      RSet TaxXRec.RealVal = Using$(FF10$, TBRec.RealValue)
      RSet TaxXRec.CurrTaxAmt = Using$(FF10$, TBRec.RealTaxDue)
      RSet TaxXRec.PropDesc2 = QPTrim$(TBRec.RDesc2)
      RSet TaxXRec.PropDesc3 = QPTrim$(TBRec.RDesc3)
      RSet TaxXRec.TotTaxAmt = Using$(FF10$, TBRec.TotalBillDue) ' - TBRec.OverPayAmt)
      RSet TaxXRec.LateListAmt = Using$(FF10$, TBRec.LateTaxDue)
      RSet TaxXRec.ExemptAmt = Using$(FF10$, TBRec.ExptValue)
      PrnCnt = PrnCnt + 1
'      If QPTrim(TaxXRec.TaxBillNum) = "199" Then Stop
      Print #RptHandle, TaxXRec.TaxBillNum; TaxXRec.CustName;
      Print #RptHandle, TaxXRec.Add1; TaxXRec.Add2; TaxXRec.Add3;
      Print #RptHandle, TaxXRec.TaxYear; TaxXRec.CustAcct;
      Print #RptHandle, TaxXRec.MapNum; TaxXRec.PropDesc1;
      Print #RptHandle, TaxXRec.TAXRATE; TaxXRec.LandVal;
      Print #RptHandle, TaxXRec.BldgVal; TaxXRec.RealVal;
      Print #RptHandle, TaxXRec.CurrTaxAmt; TaxXRec.PropDesc2;
      PastDue = GetCustBalance(CLng(TaxXRec.CustAcct), -1)
      Print #RptHandle, TaxXRec.PropDesc3; TaxXRec.TotTaxAmt; Using$(FF10$, PastDue); 'added last ";" on 11/20/06
      If TBRec.RealPropRecord > 0 Then 'added 11/20/06 at request of Boiling Springs
        Get RHandle, TBRec.RealPropRecord, RealRec
        RSet TaxXRec.RealPin = QPTrim$(RealRec.RealPin)
        Print #RptHandle, TaxXRec.RealPin; '(RealRec.RealPin);
      Else
        Print #RptHandle, "                  NA";
      End If
      Print #RptHandle, TaxXRec.LateListAmt; '9/1/2010 added semi colon
      Print #RptHandle, TaxXRec.ExemptAmt 'added 9/1/2010 requested by Beech Mt.
      For y = 1 To CDCnt
        If TBRec.CustPin = CustDone(y) Then
            GoTo AlreadyInArr
        End If
      Next y

      CDCnt = CDCnt + 1
      ReDim Preserve CustDone(1 To CDCnt) As Long
      CustDone(CDCnt) = TBRec.CustPin

AlreadyInArr:
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
Skip:
  Next x
  Unload frmTaxShowPctComp
  cmdProcess.Enabled = True
  cmdExit.Enabled = True
  EnableCloseButton Me.hwnd, True
  Close
  
  If PrnCnt > 0 Then
    Call MakeFile 'Makefile creates the .dat identifying that bill printing took place
  End If
  
  Call TaxMsg(800, "The tax billing records have been successfully saved in the file named 'LCRE" + TaxYear + ".TXT' located in the Citipak folder.")
        
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxBillPrinting", "PrintExpReal", Erl)
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

Private Sub PrintExpPers()
  Dim x As Long, y As Integer
  Dim TaxXRec As TaxBillExportPersType
  Dim TXHandle As Integer
  Dim TBRec As TaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim ReportFile As String
  Dim RptHandle As Integer
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim TaxYear$
  Dim FF11$
  Dim Map$
  Dim PersRec As PersonalRecType
  Dim PHandle As Integer
  Dim NumOfPRecs As Long
  Dim BillNo&
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim ThisSSN As String * 11
  Dim PersExemp As Double
  Dim NextRec As Long
  Dim ExempAmt As Double
  Dim PersDue As Double
  Dim TotPersVal As Double
  Dim FF9$
  Dim PrnCnt As Long
  Dim PastDue#
  
  'on error goto ERRORSTUFF
  
  PrnCnt = 0
  FF11$ = "########.#0"
  FF9$ = "######.#0"
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  TaxYear$ = CStr(TaxMasterRec.TaxYear)
  ReportFile$ = "LCPP" + TaxYear + ".TXT"
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  
  BillNo& = fpDblSnglStartBill.Value
  
  OpenTaxBillFile TBHandle, NumOfTBRecs
  OpenPersPropFile PHandle, NumOfPRecs
  OpenTaxCustFile TCHandle, NumOfTCRecs
  
  frmTaxShowPctComp.Label1 = "Creating Personal Export Tax Bills"
  frmTaxShowPctComp.cmdCancel.Visible = False
  frmTaxShowPctComp.Show , Me
  cmdProcess.Enabled = False
  cmdExit.Enabled = False
  EnableCloseButton Me.hwnd, False
  
  For x = 1 To NumOfTBRecs
    Get TBHandle, x, TBRec
    If TBRec.BillNumber > 0 Then
      If TBRec.TotalBillDue > 0 Then 'assign and save bill number
        TBRec.BillNumber = BillNo&
        TBRec.BillPrinted = True
        Put TBHandle, x, TBRec
        BillNo& = BillNo& + 1
      End If
      If TBRec.OverPayAmt > 0 Then
        TBRec.TotalBillDue = OldRound(TBRec.TotalBillDue - TBRec.OverPayAmt)
      End If
      RSet TaxXRec.CustName = QPTrim$(TBRec.CustName)
      RSet TaxXRec.Add1 = QPTrim$(TBRec.CustAdd1)
      RSet TaxXRec.Add2 = QPTrim$(TBRec.CustAdd2)
      Get TCHandle, TBRec.CustRec, TaxCust
      RSet TaxXRec.City = QPTrim$(TaxCust.City)
      RSet TaxXRec.State = QPTrim$(TaxCust.State)
      RSet TaxXRec.Zip = QPTrim$(TaxCust.Zip)
      RSet TaxXRec.CustAcct = Using$("#####0", TBRec.CustRec)
      ThisSSN = QPTrim$(TaxCust.CSSN)
      ThisSSN = ReplaceString(ThisSSN, "-", "")
      Call InsertSSNDashes(ThisSSN)
      RSet TaxXRec.SSN1 = QPTrim$(ThisSSN)
      ThisSSN = QPTrim$(TaxCust.OSSN)
      ThisSSN = ReplaceString(ThisSSN, "-", "")
      Call InsertSSNDashes(ThisSSN)
      RSet TaxXRec.SSN2 = QPTrim$(ThisSSN)
      RSet TaxXRec.DueDate = fptxtDueDate.Text
      NextRec = TaxCust.FirstPersRec
      
      PersExemp = 0
      If NextRec = 0 Then
        RSet TaxXRec.LessRelief = Using$(FF11$, 0)
        RSet TaxXRec.NetDue = Using$(FF11$, 0)
        RSet TaxXRec.TotDue = Using$(FF11$, 0)
        RSet TaxXRec.RepeatDesc = ""
        RSet TaxXRec.RepeatID = ""
        RSet TaxXRec.RepeatValue = Using$(FF11$, 0)
        RSet TaxXRec.RepeatTaxRate = Using$("#.#0", 0)
        RSet TaxXRec.RepeatTotTax = Using$(FF9$, 0)
        RSet TaxXRec.RepeatTaxRelief = Using$(FF9, 0)
        RSet TaxXRec.RepeatTotTax = Using$(FF9$, 0)
      Else
        Get PHandle, NextRec, PersRec
        PersExemp = OldRound(PersExemp + PersRec.EXMPOTHR + PersRec.EXMPSENI)
        ExempAmt = OldRound(TBRec.PersTaxRate * PersExemp)
        PersDue = OldRound(ExempAmt + TBRec.PersTaxDue)
        RSet TaxXRec.LessRelief = Using$(FF11$, ExempAmt)
        RSet TaxXRec.NetDue = Using$(FF11$, TBRec.PersTaxDue)
        RSet TaxXRec.TotDue = Using$(FF11$, OldRound(TBRec.TotalBillDue - TBRec.OverPayAmt))
        RSet TaxXRec.RepeatDesc = QPTrim$(PersRec.DESC1)
        RSet TaxXRec.RepeatID = QPTrim$(PersRec.PropPin)
        TotPersVal = OldRound(PersRec.CVALUE + PersRec.PersVal + PersRec.MCVALUE + PersRec.MHVALUE + PersRec.MTVALUE)
        RSet TaxXRec.RepeatValue = Using$(FF11$, TotPersVal)
        RSet TaxXRec.RepeatTaxRate = Using$("#.#0", TBRec.PersTaxRate)
        RSet TaxXRec.RepeatTotTax = Using$(FF9$, PersDue)
        RSet TaxXRec.RepeatTaxRelief = Using$(FF9, ExempAmt)
        RSet TaxXRec.RepeatTotTax = Using$(FF9$, TBRec.PersTaxDue)
      End If
      PrnCnt = PrnCnt + 1
      Print #RptHandle, TaxXRec.CustName; TaxXRec.Add1; TaxXRec.Add2;
      Print #RptHandle, TaxXRec.City; TaxXRec.State; TaxXRec.Zip;
      Print #RptHandle, TaxXRec.CustAcct; TaxXRec.SSN1; TaxXRec.SSN2;
      Print #RptHandle, TaxXRec.DueDate; TaxXRec.TotDue; TaxXRec.LessRelief;
      Print #RptHandle, TaxXRec.NetDue;
'      For y = 1 To 75
'        Print #RptHandle, TaxXRec.RepeatDesc; TaxXRec.RepeatID;
'        Print #RptHandle, TaxXRec.RepeatValue; TaxXRec.RepeatTaxRate;
'        Print #RptHandle, TaxXRec.RepeatTotTax; TaxXRec.RepeatTaxRelief;
'        Print #RptHandle, TaxXRec.RepeatNetTax;
'      Next y
      Print #RptHandle, TaxXRec.RepeatDesc; TaxXRec.RepeatID;
      Print #RptHandle, TaxXRec.RepeatValue; TaxXRec.RepeatTaxRate;
      Print #RptHandle, TaxXRec.RepeatTotTax; TaxXRec.RepeatTaxRelief;
      PastDue = GetCustBalance(CLng(TaxXRec.CustAcct), -1)
      Print #RptHandle, Using$(FF9$, (TaxXRec.RepeatTotTax - TaxXRec.RepeatTaxRelief)); Using$(FF11$, PastDue);
      Print #RptHandle, Using$(FF11, TBRec.PersValue) 'added 11/20/06 at request of Boiling Springs
    
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
  
  Unload frmTaxShowPctComp
  cmdProcess.Enabled = True
  cmdExit.Enabled = True
  EnableCloseButton Me.hwnd, True
  
  Close
  
  If PrnCnt > 0 Then
    Call MakeFile 'Makefile creates the .dat identifying that bill printing took place
  End If
  
  Call TaxMsg(800, "The tax billing records have been successfully saved in the file named 'LCPP" + TaxYear + ".TXT' located in the Citipak folder.")
        
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxBillPrinting", "PrintExpPers", Erl)
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

Private Sub PrintPostCard1()
  'good for Troy
  Dim RptHandle As Integer
  Dim RptFile$
  Dim TaxBill As TaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim x As Long, BillNo&
  Dim PrnCnt&
  Dim PINTemp As String * 18
  Dim CustName$, PCnt As Long
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
  Dim CustArr As Long
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim ZipRec As BillPrintZipIdxType
  Dim ZHandle As Integer
  Dim NumOfZRecs As Long
  Dim MortRec As BillPrintMortIdxType
  Dim MRHandle As Integer
  Dim NumOfMRRecs As Long
  
  'on error goto ERRORSTUFF
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  FF$ = Chr(12)
  BillNo& = fpDblSnglStartBill.Value
  
  RptHandle = FreeFile
  RptFile$ = "TAXBILLPC1.PRN"
  Open RptFile For Output As RptHandle
  
  OpenTaxBillFile TBHandle, NumOfTBRecs
  OpenPersPropFile PHandle, NumOfPRecs
  OpenRealPropFile RHandle, NumOfRRecs
  
  frmTaxShowPctComp.Label1 = "Printing Tax Bills"
  frmTaxShowPctComp.cmdCancel.Visible = False
  frmTaxShowPctComp.Show , Me
  cmdProcess.Enabled = False
  cmdExit.Enabled = False
  cmdAlign.Enabled = False
  EnableCloseButton Me.hwnd, False

  If OptMort.Value = True Then '12/6/06
    OpenMortIdxFile MRHandle, NumOfMRRecs
    NumOfTBRecs = NumOfMRRecs
  ElseIf optZip.Value = True Then '12/6/06
    OpenZipIdxFile ZHandle, NumOfZRecs
    NumOfTBRecs = NumOfZRecs
  End If
  
  For x = 1 To NumOfTBRecs
    If OptMort.Value = True Then '12/6/06
      Get MRHandle, x, MortRec
      CustArr = MortRec.TaxBillRec
    ElseIf optZip.Value = True Then '12/6/06
      Get ZHandle, x, ZipRec
      CustArr = ZipRec.TaxBillRec
    Else
      CustArr = x '12/6/06
    End If
    Get TBHandle, CustArr, TaxBill
    If TaxBill.BillNumber > 0 Then
      If TaxBill.TotalBillDue > 0 Then
        RSet PINTemp = TaxBill.RealPin
        PastDue = GetCustBalance(TaxBill.CustRec, -1)
        CustName$ = QPTrim$(TaxBill.CustName)
        TaxBill.BillNumber = BillNo&
        TaxBill.BillPrinted = True
        Put TBHandle, CustArr, TaxBill '12/6/06
        GoSub PrintPC1
        BillNo& = BillNo& + 1
        PrnCnt = PrnCnt + 1
      End If
    End If
    frmTaxShowPctComp.ShowPctComp x, NumOfTBRecs
    If frmTaxShowPctComp.Out = True Then
      Close
      frmTaxShowPctComp.Out = False
      Unload frmTaxShowPctComp
      cmdProcess.Enabled = True
      cmdExit.Enabled = True
      cmdAlign.Enabled = True
      EnableCloseButton Me.hwnd, True
      Exit Sub
    End If
  Next x
  Unload frmTaxShowPctComp
  cmdProcess.Enabled = True
  cmdExit.Enabled = True
  cmdAlign.Enabled = True
  EnableCloseButton Me.hwnd, True
  Print #RptHandle, FF$
  
  Close TBHandle
  Close RptHandle
  ViewPrint RptFile$, "Tax Bill Printing", True
  If PrnCnt > 0 Then
    Call MakeFile 'Makefile creates the .dat identifying that bill printing took place
  End If
  Close
  Exit Sub
  
PrintPC1:
  ThisName = QPTrim$(TaxBill.CustName)
  ThisCName = QPTrim$(TaxBill.CustName)
  ThisAdd1 = QPTrim$(TaxBill.CustAdd1)
  ThisAdd2 = QPTrim$(TaxBill.CustAdd2)
  ThisAdd3 = QPTrim$(TaxBill.CustAdd3) + " " + InsertZipDash(TaxBill.CustZip)
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  '-----------------------------------------
  Print #RptHandle, Tab(1); Using$("###0", TaxBill.TaxYear); Tab(7); Using$("####0", BillNo); Tab(14); Using$("######0", TaxBill.CustPin);
'  Print #RptHandle, Tab(22); ""; Tab(33); Using$("###0", TaxBill.TaxYear); Tab(40); Using$("####0", BillNo); Tab(50); Using$("######0", TaxBill.CustPin)
  If InStr(TaxMasterRec.Name, "SEVEN DEVILS") Then
    Print #RptHandle, Tab(22); ""; Tab(40); Using$("###0", TaxBill.TaxYear); Tab(50); ""; Tab(60); Using$("######0", TaxBill.CustPin)
  ElseIf InStr(TaxMasterRec.Name, "ANDREWS") Then
    Print #RptHandle, Tab(22); ""; Tab(40); Using$("###0", TaxBill.TaxYear); Tab(50); Using$("######0", TaxBill.CustPin); Tab(60); Using$("$##,##0.00", TaxBill.TotalBillDue)
  Else
    Print #RptHandle, Tab(26); Right(QPTrim$(TaxBill.RealPin), 11); Tab(40); Using$("###0", TaxBill.TaxYear); Tab(50); Using$("####0", BillNo); Tab(60); Using$("######0", TaxBill.CustPin)
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
  Print #RptHandle, Tab(10); ThisDesc
  Print #RptHandle, Tab(10); LotsAcres
  Print #RptHandle,
  Print #RptHandle, Tab(2); Using$("$###,###.00", TaxBill.RealValue); Tab(15); Using$("$###,##0.00", TaxBill.PersValue); Tab(26); Using$("$###,##0.00", TaxBill.ExptValue);
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  If TaxBill.RealTaxRate > 0 Then
    Print #RptHandle, Tab(2); Using$("$###,##0.00", OldRound(TaxBill.RealValue + TaxBill.PersValue)); Tab(16); Using$("##0.00", TaxBill.RealTaxRate);
  ElseIf TaxBill.PersTaxRate > 0 Then
    Print #RptHandle, Tab(2); Using$("$###,##0.00", OldRound(TaxBill.RealValue + TaxBill.PersValue)); Tab(16); Using$("##0.00", TaxBill.PersTaxRate);
  Else
    Print #RptHandle, Tab(2); Using$("$###,##0.00", OldRound(TaxBill.RealValue + TaxBill.PersValue)); Tab(16); Using$("##0.00", 0);
  End If
  Print #RptHandle, Tab(26); Using$("$###,##0.00", OldRound(TaxBill.OverPayAmt + TaxBill.TotalBillDue - TaxBill.LateTaxDue)); Tab(42); ThisName 'added OverPayment on 8/16/06
  Print #RptHandle, Tab(42); ThisAdd1
  Print #RptHandle, Tab(42); ThisAdd2
  Print #RptHandle, Tab(28); Using$("$#,##0.00", TaxBill.LateTaxDue); Tab(42); ThisAdd3
  Print #RptHandle,
  Print #RptHandle, Using$("$###,##0.00", PastDue); Tab(26); Using$("$###,##0.00", TaxBill.TotalBillDue)
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  
  Return

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxBillPrinting", "PrintPostCard1", Erl)
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

Private Sub fpcmbBarcode_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbBarcode.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbBarcode.ListIndex = -1
  End If
  If fpcmbBarcode.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpDblSnglStartBill.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpDblSnglStartBill.SetFocus
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpDblSnglStartBill_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
    If fpcmbBarcode.Visible = True Then
      fpcmbBarcode.SetFocus
    End If
  End If
End Sub

Private Sub PrintHMLT24TF()
  Dim RptHandle As Integer
  Dim RptFile$
  Dim TaxBill As TaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim x As Long, BillNo&
  Dim PrnCnt&
  Dim PINTemp As String * 18
  Dim CustName$
  Dim ThisRec As Long
  Dim ThisDesc As String * 20
  Dim LotsAcres As String * 20
  Dim FF$
  Dim NetTaxVal#
  Dim Cnt As Integer
  Dim CustArr As Long
  Dim ZipRec As BillPrintZipIdxType
  Dim ZHandle As Integer
  Dim NumOfZRecs As Long
  Dim MortRec As BillPrintMortIdxType
  Dim MRHandle As Integer
  Dim NumOfMRRecs As Long
  
'  'on error goto ERRORSTUFF
  
  FF$ = Chr(12)
  BillNo& = fpDblSnglStartBill.Value
  
  RptHandle = FreeFile
  RptFile$ = "TAXBILLHAMLET.PRN"
  Open RptFile For Output As RptHandle
  
  OpenTaxBillFile TBHandle, NumOfTBRecs
  
  frmTaxShowPctComp.Label1 = "Printing Tax Bills"
  frmTaxShowPctComp.cmdCancel.Visible = False
  frmTaxShowPctComp.Show , Me
  cmdProcess.Enabled = False
  cmdExit.Enabled = False
  cmdAlign.Enabled = False
  EnableCloseButton Me.hwnd, False
  BillNo& = fpDblSnglStartBill
  
  If OptMort.Value = True Then '12/6/06
    OpenMortIdxFile MRHandle, NumOfMRRecs
    NumOfTBRecs = NumOfMRRecs
  ElseIf optZip.Value = True Then '12/6/06
    OpenZipIdxFile ZHandle, NumOfZRecs
    NumOfTBRecs = NumOfZRecs
  End If
  
  For x = 1 To NumOfTBRecs
    If OptMort.Value = True Then '12/6/06
      Get MRHandle, x, MortRec
      CustArr = MortRec.TaxBillRec
    ElseIf optZip.Value = True Then '12/6/06
      Get ZHandle, x, ZipRec
      CustArr = ZipRec.TaxBillRec
    Else
      CustArr = x '12/6/06
    End If
    Get TBHandle, CustArr, TaxBill
    If TaxBill.BillNumber > 0 Then
      If TaxBill.TotalBillDue > 0 Then
        RSet PINTemp = TaxBill.RealPin
        CustName$ = QPTrim$(TaxBill.CustName)
        PrnCnt = PrnCnt + 1
        TaxBill.BillNumber = BillNo&
        GoSub Hamlet
        TaxBill.BillPrinted = True
        Put TBHandle, CustArr, TaxBill '12/6/06
        BillNo& = BillNo& + 1
      End If
    End If
    frmTaxShowPctComp.ShowPctComp x, NumOfTBRecs
    If frmTaxShowPctComp.Out = True Then
      Close
      frmTaxShowPctComp.Out = False
      Unload frmTaxShowPctComp
      cmdProcess.Enabled = True
      cmdExit.Enabled = True
      cmdAlign.Enabled = True
      EnableCloseButton Me.hwnd, True
      Exit Sub
    End If
  Next x
  Unload frmTaxShowPctComp
  cmdProcess.Enabled = True
  cmdExit.Enabled = True
  cmdAlign.Enabled = True
  EnableCloseButton Me.hwnd, True
  Print #RptHandle, FF$
  Close TBHandle
  Close RptHandle
  ViewPrint RptFile$, "Tax Bill Printing", True
  
  If PrnCnt > 0 Then
    Call MakeFile 'Makefile creates the .dat identifying that bill printing took place
  End If
  Close
  Exit Sub
  
Hamlet:
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
  
  Return
End Sub

Private Sub PrintPH24TF()
  Dim RptHandle As Integer
  Dim RptFile$
  Dim TaxBill As TaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim x As Long, BillNo&
  Dim PrnCnt&
  Dim PINTemp As String * 18
  Dim CustName$
  Dim ThisRec As Long
  Dim ThisDesc As String * 20
  Dim LotsAcres As String * 20
  Dim FF$, TAXRATE#
  Dim NetTaxVal#
  Dim CustArr As Long
  Dim ZipRec As BillPrintZipIdxType
  Dim ZHandle As Integer
  Dim NumOfZRecs As Long
  Dim MortRec As BillPrintMortIdxType
  Dim MRHandle As Integer
  Dim NumOfMRRecs As Long
  Dim WhitPINTemp As String * 4
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim TownName As String
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  TownName = QPTrim$(TaxMasterRec.City)
  
  FF$ = Chr(12)
  BillNo& = fpDblSnglStartBill.Value
  TAXRATE# = CDbl(fpDblSnglRealRate.Value)
  
  RptHandle = FreeFile
  RptFile$ = "TAXBILLPH24TF.PRN"
  Open RptFile For Output As RptHandle
  
  OpenTaxBillFile TBHandle, NumOfTBRecs
  
  frmTaxShowPctComp.Label1 = "Printing Tax Bills"
  frmTaxShowPctComp.cmdCancel.Visible = False
  frmTaxShowPctComp.Show , Me
  cmdProcess.Enabled = False
  cmdExit.Enabled = False
  cmdAlign.Enabled = False
  EnableCloseButton Me.hwnd, False
  
  PrnCnt = 1
  BillNo& = fpDblSnglStartBill
  
  If OptMort.Value = True Then '12/6/06
    OpenMortIdxFile MRHandle, NumOfMRRecs
    NumOfTBRecs = NumOfMRRecs
  ElseIf optZip.Value = True Then '12/6/06
    OpenZipIdxFile ZHandle, NumOfZRecs
    NumOfTBRecs = NumOfZRecs
  End If
  
  For x = 1 To NumOfTBRecs
    If OptMort.Value = True Then '12/6/06
      Get MRHandle, x, MortRec
      CustArr = MortRec.TaxBillRec
    ElseIf optZip.Value = True Then '12/6/06
      Get ZHandle, x, ZipRec
      CustArr = ZipRec.TaxBillRec
    Else
      CustArr = x '12/6/06
    End If
    Get TBHandle, CustArr, TaxBill
    If TaxBill.BillNumber > 0 Then
      If TaxBill.TotalBillDue > 0 Then
        RSet PINTemp = TaxBill.RealPin
        If InStr(TownName, "WHITAKERS") Then
         If QPTrim$(TaxBill.RealPin) = "" Then GoTo NoPin
         WhitPINTemp = Mid(TaxBill.RealPin, Len(QPTrim$(TaxBill.RealPin)) - 3, 4)
         RSet PINTemp = WhitPINTemp
        End If
NoPin:
        CustName$ = QPTrim$(TaxBill.CustName)
        PrnCnt = PrnCnt + 1
        TaxBill.BillNumber = BillNo&
        GoSub SubPrintPH24TF
        TaxBill.BillPrinted = True
        Put TBHandle, CustArr, TaxBill '12/6/06
        BillNo& = BillNo& + 1
      End If
    End If
    frmTaxShowPctComp.ShowPctComp x, NumOfTBRecs
    If frmTaxShowPctComp.Out = True Then
      Close
      frmTaxShowPctComp.Out = False
      Unload frmTaxShowPctComp
      cmdProcess.Enabled = True
      cmdExit.Enabled = True
      cmdAlign.Enabled = True
      EnableCloseButton Me.hwnd, True
      Exit Sub
    End If
  Next x
  
  Unload frmTaxShowPctComp
  cmdProcess.Enabled = True
  cmdExit.Enabled = True
  cmdAlign.Enabled = True
  EnableCloseButton Me.hwnd, True
  Print #RptHandle, FF$
  Close TBHandle
  
  Close RptHandle
  ViewPrint RptFile$, "Tax Bill Printing", True
  If PrnCnt > 0 Then
    Call MakeFile 'Makefile creates the .dat identifying that bill printing took place
  End If
  Close
  Exit Sub

SubPrintPH24TF:
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
  Return
End Sub

Private Sub PrintSYL23TF()
  Dim RptHandle As Integer
  Dim RptFile$
  Dim TaxBill As TaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim x As Long, BillNo&
  Dim PrnCnt&
  Dim PINTemp As String * 18
  Dim CustName$
  Dim ThisRec As Long
  Dim ThisDesc As String * 20
  Dim LotsAcres As String * 20
  Dim FF$
  Dim NetTaxVal#
  Dim TAXRATE#
  Dim CustArr As Long
  Dim ZipRec As BillPrintZipIdxType
  Dim ZHandle As Integer
  Dim NumOfZRecs As Long
  Dim MortRec As BillPrintMortIdxType
  Dim MRHandle As Integer
  Dim NumOfMRRecs As Long
  
  FF$ = Chr(12)
  BillNo& = fpDblSnglStartBill.Value
  
  RptHandle = FreeFile
  RptFile$ = "TAXBILLSYL23TF.PRN"
  Open RptFile For Output As RptHandle
  
  OpenTaxBillFile TBHandle, NumOfTBRecs
  
  frmTaxShowPctComp.Label1 = "Printing Tax Bills"
  frmTaxShowPctComp.cmdCancel.Visible = False
  frmTaxShowPctComp.Show , Me
  cmdProcess.Enabled = False
  cmdExit.Enabled = False
  cmdAlign.Enabled = False
  EnableCloseButton Me.hwnd, False
  
  PrnCnt = 1
  BillNo& = fpDblSnglStartBill
  
  If OptMort.Value = True Then '12/6/06
    OpenMortIdxFile MRHandle, NumOfMRRecs
    NumOfTBRecs = NumOfMRRecs
  ElseIf optZip.Value = True Then '12/6/06
    OpenZipIdxFile ZHandle, NumOfZRecs
    NumOfTBRecs = NumOfZRecs
  End If
  
  For x = 1 To NumOfTBRecs
    If OptMort.Value = True Then '12/6/06
      Get MRHandle, x, MortRec
      CustArr = MortRec.TaxBillRec
    ElseIf optZip.Value = True Then '12/6/06
      Get ZHandle, x, ZipRec
      CustArr = ZipRec.TaxBillRec
    Else
      CustArr = x '12/6/06
    End If
    Get TBHandle, CustArr, TaxBill
    If TaxBill.BillNumber > 0 Then
      If TaxBill.TotalBillDue > 0 Then
        RSet PINTemp = TaxBill.RealPin
        CustName$ = QPTrim$(TaxBill.CustName)
        PrnCnt = PrnCnt + 1
        TaxBill.BillNumber = BillNo&
        GoSub SubPrintSYL23TF
        TaxBill.BillPrinted = True
        Put TBHandle, CustArr, TaxBill '12/6/06
        BillNo& = BillNo& + 1
      End If
    End If
    frmTaxShowPctComp.ShowPctComp x, NumOfTBRecs
    If frmTaxShowPctComp.Out = True Then
      Close
      frmTaxShowPctComp.Out = False
      Unload frmTaxShowPctComp
      cmdProcess.Enabled = True
      cmdExit.Enabled = True
      cmdAlign.Enabled = True
      EnableCloseButton Me.hwnd, True
      Exit Sub
    End If
  Next x
  
  Unload frmTaxShowPctComp
  cmdProcess.Enabled = True
  cmdExit.Enabled = True
  cmdAlign.Enabled = True
  EnableCloseButton Me.hwnd, True
  Print #RptHandle, FF$
  Close TBHandle
  
  Close RptHandle
  ViewPrint RptFile$, "Tax Bill Printing", True
  Close
  If PrnCnt > 0 Then
    Call MakeFile 'Makefile creates the .dat identifying that bill printing took place
  End If
  
  Exit Sub

SubPrintSYL23TF:
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

  Return
End Sub

Private Sub PrintBSC32TF()
  Dim RptHandle As Integer
  Dim RptFile$
  Dim TaxBill As TaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim x As Long, BillNo&
  Dim PrnCnt&
  Dim PINTemp As String * 18
  Dim CustName$
  Dim ThisRec As Long
  Dim ThisDesc As String * 20
  Dim LotsAcres As String * 20
  Dim FF$
  Dim NetTaxVal#
  Dim TAXRATE#
  Dim CustArr As Long '12/6/06
  Dim ZipRec As BillPrintZipIdxType
  Dim ZHandle As Integer
  Dim NumOfZRecs As Long
  Dim MortRec As BillPrintMortIdxType
  Dim MRHandle As Integer
  Dim NumOfMRRecs As Long
  
  FF$ = Chr(12)
  BillNo& = fpDblSnglStartBill.Value
  RptHandle = FreeFile
  RptFile$ = "TAXBILLBSC32TF.PRN"
  Open RptFile For Output As RptHandle
  
  OpenTaxBillFile TBHandle, NumOfTBRecs
  
  frmTaxShowPctComp.Label1 = "Printing Tax Bills"
  frmTaxShowPctComp.cmdCancel.Visible = False
  frmTaxShowPctComp.Show , Me
  cmdProcess.Enabled = False
  cmdExit.Enabled = False
  cmdAlign.Enabled = False
  EnableCloseButton Me.hwnd, False
  
  PrnCnt = 1
  BillNo& = fpDblSnglStartBill
  
  If OptMort.Value = True Then '12/6/06
    OpenMortIdxFile MRHandle, NumOfMRRecs
    NumOfTBRecs = NumOfMRRecs
  ElseIf optZip.Value = True Then '12/6/06
    OpenZipIdxFile ZHandle, NumOfZRecs
    NumOfTBRecs = NumOfZRecs
  End If
  
  For x = 1 To NumOfTBRecs
    If OptMort.Value = True Then '12/6/06
      Get MRHandle, x, MortRec
      CustArr = MortRec.TaxBillRec
    ElseIf optZip.Value = True Then '12/6/06
      Get ZHandle, x, ZipRec
      CustArr = ZipRec.TaxBillRec
    Else
      CustArr = x '12/6/06
    End If
    Get TBHandle, CustArr, TaxBill
    If TaxBill.BillNumber > 0 Then
      If TaxBill.TotalBillDue > 0 Then
        RSet PINTemp = TaxBill.RealPin
        CustName$ = QPTrim$(TaxBill.CustName)
        PrnCnt = PrnCnt + 1
        TaxBill.BillNumber = BillNo&
        GoSub SubPrintBSC32TF
        TaxBill.BillPrinted = True
        Put TBHandle, CustArr, TaxBill '12/6/06
        BillNo& = BillNo& + 1
      End If
    End If
    frmTaxShowPctComp.ShowPctComp x, NumOfTBRecs
    If frmTaxShowPctComp.Out = True Then
      Close
      frmTaxShowPctComp.Out = False
      Unload frmTaxShowPctComp
      cmdProcess.Enabled = True
      cmdExit.Enabled = True
      cmdAlign.Enabled = True
      EnableCloseButton Me.hwnd, True
      Exit Sub
    End If
  Next x
  
  Unload frmTaxShowPctComp
  cmdProcess.Enabled = True
  cmdExit.Enabled = True
  cmdAlign.Enabled = True
  EnableCloseButton Me.hwnd, True
  Print #RptHandle, FF$
  Close TBHandle
  
  Close RptHandle
  ViewPrint RptFile$, "Tax Bill Printing", True
  
  If PrnCnt > 0 Then
    Call MakeFile 'Makefile creates the .dat identifying that bill printing took place
  End If
  
  Close
  Exit Sub

SubPrintBSC32TF:
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

  Return
  
End Sub

Private Sub PrintLLN21TF()
  Dim RptHandle As Integer
  Dim RptFile$
  Dim TaxBill As TaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim x As Long, BillNo&
  Dim PrnCnt&
  Dim PINTemp As String * 18
  Dim CustName$ ', PCnt As Long
  Dim ThisRec As Long
  Dim ThisDesc As String * 20
  Dim LotsAcres As String * 20
  Dim FF$
  Dim NetTaxVal#
  Dim TAXRATE#
  Dim CustArr As Long
  Dim ZipRec As BillPrintZipIdxType
  Dim ZHandle As Integer
  Dim NumOfZRecs As Long
  Dim MortRec As BillPrintMortIdxType
  Dim MRHandle As Integer
  Dim NumOfMRRecs As Long
  
  FF$ = Chr(12)
  BillNo& = fpDblSnglStartBill.Value
  
  RptHandle = FreeFile
  RptFile$ = "TAXBILLLLN21TF.PRN"
  Open RptFile For Output As RptHandle
  
  OpenTaxBillFile TBHandle, NumOfTBRecs
  
  frmTaxShowPctComp.Label1 = "Printing Tax Bills"
  frmTaxShowPctComp.cmdCancel.Visible = False
  frmTaxShowPctComp.Show , Me
  cmdProcess.Enabled = False
  cmdExit.Enabled = False
  cmdAlign.Enabled = False
  EnableCloseButton Me.hwnd, False
  
  PrnCnt = 1
  BillNo& = fpDblSnglStartBill
  
  If OptMort.Value = True Then '12/6/06
    OpenMortIdxFile MRHandle, NumOfMRRecs
    NumOfTBRecs = NumOfMRRecs
  ElseIf optZip.Value = True Then '12/6/06
    OpenZipIdxFile ZHandle, NumOfZRecs
    NumOfTBRecs = NumOfZRecs
  End If
  
  For x = 1 To NumOfTBRecs
    If OptMort.Value = True Then '12/6/06
      Get MRHandle, x, MortRec
      CustArr = MortRec.TaxBillRec
    ElseIf optZip.Value = True Then '12/6/06
      Get ZHandle, x, ZipRec
      CustArr = ZipRec.TaxBillRec
    Else
      CustArr = x '12/6/06
    End If
    Get TBHandle, CustArr, TaxBill
    If TaxBill.BillNumber > 0 Then
      If TaxBill.TotalBillDue > 0 Then
        RSet PINTemp = TaxBill.RealPin
        CustName$ = QPTrim$(TaxBill.CustName)
        PrnCnt = PrnCnt + 1
        TaxBill.BillNumber = BillNo&
        GoSub SubPrintLLN21TF
        TaxBill.BillPrinted = True
        Put TBHandle, CustArr, TaxBill '12/6/06
        BillNo& = BillNo& + 1
      End If
    End If
    frmTaxShowPctComp.ShowPctComp x, NumOfTBRecs
    If frmTaxShowPctComp.Out = True Then
      Close
      frmTaxShowPctComp.Out = False
      Unload frmTaxShowPctComp
      cmdProcess.Enabled = True
      cmdExit.Enabled = True
      cmdAlign.Enabled = True
      EnableCloseButton Me.hwnd, True
      Exit Sub
    End If
  Next x
  
  Unload frmTaxShowPctComp
  cmdProcess.Enabled = True
  cmdExit.Enabled = True
  cmdAlign.Enabled = True
  EnableCloseButton Me.hwnd, True
  Print #RptHandle, FF$
  Close TBHandle
  
  Close RptHandle
  ViewPrint RptFile$, "Tax Bill Printing", True
  
  If PrnCnt > 0 Then
    Call MakeFile 'Makefile creates the .dat identifying that bill printing took place
  End If
  Close
  Exit Sub

SubPrintLLN21TF:
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
  Print #RptHandle, Tab(71); Using("####0.00", TaxBill.TotalBillDue - TaxBill.OverPayAmt) 'added OverPayAmt 8/15/06
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle, "~"; Tab(78); "~"

  Return
End Sub

Private Sub PrintLaserLegal()
  Dim RptHandle As Integer
  Dim RptFile$
  Dim TaxBill As TaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim x As Long, BillNo&
  Dim dlm$
  Dim LA$
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim PrnCnt As Integer
  Dim CustCSZ$
  Dim PinNum$
  Dim Desc As String * 29
  Dim Name As String * 28
  Dim CustArr As Long '12/6/06
  Dim ZipRec As BillPrintZipIdxType
  Dim ZHandle As Integer
  Dim NumOfZRecs As Long
  Dim MortRec As BillPrintMortIdxType
  Dim MRHandle As Integer
  Dim NumOfMRRecs As Long
  
  dlm$ = "~"
  BillNo& = fpDblSnglStartBill.Value
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close
  
  RptFile$ = "TAXRPTS\TXLSRLEGAL.RPT"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle
  
  OpenTaxBillFile TBHandle, NumOfTBRecs
  If OptMort.Value = True Then '12/6/06
    OpenMortIdxFile MRHandle, NumOfMRRecs
    NumOfTBRecs = NumOfMRRecs
  ElseIf optZip.Value = True Then '12/6/06
    OpenZipIdxFile ZHandle, NumOfZRecs
    NumOfTBRecs = NumOfZRecs
  End If
  
  For x = 1 To NumOfTBRecs
    If OptMort.Value = True Then '12/6/06
      Get MRHandle, x, MortRec
      CustArr = MortRec.TaxBillRec
    ElseIf optZip.Value = True Then '12/6/06
      Get ZHandle, x, ZipRec
      CustArr = ZipRec.TaxBillRec
    Else
      CustArr = x '12/6/06
    End If
    Get TBHandle, CustArr, TaxBill
    If TaxBill.BillNumber > 0 Then
      TaxBill.BillNumber = BillNo&
      TaxBill.BillPrinted = True
      Put TBHandle, CustArr, TaxBill '12/6/06
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
      InsertZipDash (QPTrim$(TaxBill.CustZip))
      '                              0                      1                        2               3
      Print #RptHandle, CStr(TaxBill.TaxYear); dlm; TaxBill.BillNumber; dlm; TaxBill.CustRec; dlm; PinNum; dlm;
      '                  4           5        6               7                         8
      Print #RptHandle, Name; dlm; Desc; dlm; LA; dlm; TaxBill.RealValue; dlm; TaxBill.PersValue; dlm;
      '                        9                                       10                                     11
      Print #RptHandle, TaxBill.ExptValue; dlm; OldRound(TaxBill.PersValue + TaxBill.RealValue); dlm; TaxBill.RealTaxRate; dlm;
      '                                     12                                         13                                14
      Print #RptHandle, OldRound(TaxBill.RealTaxDue + TaxBill.PersTaxDue); dlm; TaxBill.LateTaxDue; dlm; TaxBill.TotalBillDue - TaxBill.OverPayAmt; dlm;  'added OverPayAmt 8/15/06
      '                         15                           16                      17
      Print #RptHandle, "                       "; dlm; "                  "; dlm; "                     "; dlm;
      '                  18                   19                                 20                     21
      Print #RptHandle, Name; dlm; QPTrim$(TaxBill.CustAdd1); dlm; QPTrim$(TaxBill.CustAdd2); dlm; CustCSZ
      PrnCnt = PrnCnt + 1
      BillNo& = BillNo& + 1
    End If
  Next x
  
  Close
  
  arTaxLsrLegal.Show
  If PrnCnt > 0 Then
    Call MakeFile 'Makefile creates the .dat identifying that bill printing took place
  End If
  
End Sub

Private Sub PrintLaserLegalHP()
  Dim RptHandle As Integer
  Dim RptFile$
  Dim TaxBill As TaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim x As Long, BillNo&
  Dim dlm$
  Dim LA$
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim PrnCnt As Integer
  Dim Desc As String * 29
  Dim Name As String * 28
  Dim CustCSZ$
  Dim PinNum$
  Dim CustArr As Long
  Dim ZipRec As BillPrintZipIdxType
  Dim ZHandle As Integer
  Dim NumOfZRecs As Long
  Dim MortRec As BillPrintMortIdxType
  Dim MRHandle As Integer
  Dim NumOfMRRecs As Long
  
  dlm$ = "~"
  
  BillNo& = fpDblSnglStartBill.Value
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close
  
  RptFile$ = "TAXRPTS\TXLSRLEGALHP.RPT"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle
  
  OpenTaxBillFile TBHandle, NumOfTBRecs
  If OptMort.Value = True Then '12/6/06
    OpenMortIdxFile MRHandle, NumOfMRRecs
    NumOfTBRecs = NumOfMRRecs
  ElseIf optZip.Value = True Then '12/6/06
    OpenZipIdxFile ZHandle, NumOfZRecs
    NumOfTBRecs = NumOfZRecs
  End If
  
  For x = 1 To NumOfTBRecs
    If OptMort.Value = True Then '12/6/06
      Get MRHandle, x, MortRec
      CustArr = MortRec.TaxBillRec
    ElseIf optZip.Value = True Then '12/6/06
      Get ZHandle, x, ZipRec
      CustArr = ZipRec.TaxBillRec
    Else
      CustArr = x '12/6/06
    End If
    Get TBHandle, CustArr, TaxBill
    If TaxBill.BillNumber > 0 Then
      TaxBill.BillNumber = BillNo&
      TaxBill.BillPrinted = True
      Put TBHandle, CustArr, TaxBill '12/6/06
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
      '                              0                      1                        2               3
      Print #RptHandle, CStr(TaxBill.TaxYear); dlm; TaxBill.BillNumber; dlm; TaxBill.CustRec; dlm; PinNum; dlm;
      '                  4           5        6               7                         8
      Print #RptHandle, Name; dlm; Desc; dlm; LA; dlm; TaxBill.RealValue; dlm; TaxBill.PersValue; dlm;
      '                        9                                       10                                     11
      Print #RptHandle, TaxBill.ExptValue; dlm; OldRound(TaxBill.PersValue + TaxBill.RealValue); dlm; TaxBill.RealTaxRate; dlm;
      '                                     12                                         13                                   14
      Print #RptHandle, OldRound(TaxBill.RealTaxDue + TaxBill.PersTaxDue); dlm; TaxBill.LateTaxDue; dlm; TaxBill.TotalBillDue - TaxBill.OverPayAmt; dlm; 'added OverPayAmt 8/15/06
      '                         15                           16                      17
'      Print #RptHandle, QPTrim$(TaxMasterRec.Name); dlm; QPTrim$(TaxMasterRec.Add1); dlm; QPTrim$(TaxMasterRec.City) + ", " + QPTrim$(TaxMasterRec.TownState) + " " + QPTrim$(TaxMasterRec.Zip); dlm;
      Print #RptHandle, "                       "; dlm; "                  "; dlm; "                     "; dlm;
      '                  18                    19                             20                     21
      Print #RptHandle, Name; dlm; QPTrim$(TaxBill.CustAdd1); dlm; QPTrim$(TaxBill.CustAdd2); dlm; CustCSZ
      PrnCnt = PrnCnt + 1
      BillNo& = BillNo& + 1
    End If
  Next x
  
  Close
  
  arTaxLsrLegalHP.Show
  If PrnCnt > 0 Then
    Call MakeFile 'Makefile creates the .dat identifying that bill printing took place
  End If
  
End Sub

Private Sub SortByMortCode()
  Dim x As Long, z As Integer
  Dim NextRec As Long
  Dim MortRec As MortCodeRecType
  Dim MHandle As Integer
  Dim NumOfMCodes As Integer
  Dim TBRec As TaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim NoBill As Long
  Dim MortCnt As Long
  Dim MRRec As BillPrintMortIdxType
  Dim MRHandle As Integer
  Dim NumOfMRRecs As Long
  Dim TBMortCnt As Long
  
  KillFile "ZIPIDX.DAT"
  KillFile "MORTIDX.DAT"
  
  OpenMortCodeFile MHandle, NumOfMCodes
  OpenTaxBillFile TBHandle, NumOfTBRecs
  
  MortCnt = 0
  ReDim MortArr(1 To 1) As String
  frmTaxShowPctComp.Label1 = "Sorting By Mortgage Code"
  frmTaxShowPctComp.Show , Me
  For x = 1 To NumOfTBRecs
    Get TBHandle, x, TBRec
    If TBRec.BillNumber > 0 And QPTrim$(TBRec.MORTCODE) <> "" Then
      For z = 1 To TBMortCnt
        If QPTrim$(MortArr(z)) = QPTrim$(TBRec.MORTCODE) Then
          Exit For
        End If
      Next z
      If z > TBMortCnt Then
        TBMortCnt = TBMortCnt + 1
        ReDim Preserve MortArr(1 To TBMortCnt) As String
        MortArr(TBMortCnt) = QPTrim$(TBRec.MORTCODE)
      End If
    End If
    frmTaxShowPctComp.ShowPctComp x, NumOfTBRecs
    If frmTaxShowPctComp.Out = True Then
      Close
      frmTaxShowPctComp.Out = False
      Unload frmTaxShowPctComp
      Exit Sub
    End If
  Next x
  
  Unload frmTaxShowPctComp
  
  frmTaxShowPctComp.Label1 = "Sorting By Mortgage Code"
  frmTaxShowPctComp.Show , Me
  
  OpenMortIdxFile MRHandle, NumOfMRRecs
  
  For z = 1 To TBMortCnt
    For x = 1 To NumOfTBRecs
      Get TBHandle, x, TBRec
      If TBRec.BillNumber > 0 Then
        If QPTrim$(TBRec.MORTCODE) = QPTrim$(MortArr(z)) Then
          MortCnt = MortCnt + 1
          MRRec.TaxBillRec = x
          Put MRHandle, MortCnt, MRRec
        End If
      End If
    Next x
    frmTaxShowPctComp.ShowPctComp z, TBMortCnt 'changed from NumOfMCodes to TBMortCnt on 7/18/07
    If frmTaxShowPctComp.Out = True Then
      Close
      frmTaxShowPctComp.Out = False
      Unload frmTaxShowPctComp
      Exit Sub
    End If
  Next z
  Unload frmTaxShowPctComp
  
  frmTaxShowPctComp.Label1 = "Sorting By Mortgage Code"
  frmTaxShowPctComp.Show , Me
  
  For x = 1 To NumOfTBRecs
    Get TBHandle, x, TBRec
    If TBRec.BillNumber > 0 And QPTrim$(TBRec.MORTCODE) = "" Then
      MortCnt = MortCnt + 1
      MRRec.TaxBillRec = x
      Put MRHandle, MortCnt, MRRec
    End If
    frmTaxShowPctComp.ShowPctComp x, NumOfTBRecs
    If frmTaxShowPctComp.Out = True Then
      Close
      frmTaxShowPctComp.Out = False
      Unload frmTaxShowPctComp
      Exit Sub
    End If
  Next x
  
  Unload frmTaxShowPctComp
            
  Close
  
End Sub

Private Sub SortByZipCode()
  Dim x As Long, z As Integer
  Dim NextRec As Long
  Dim TBRec As TaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim NoBill As Long
  Dim Nextx As Long
  Dim ThisZipCnt As Long
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim ThisCustRec As Long
  Dim ZipRec As BillPrintZipIdxType
  Dim ZHandle As Integer
  Dim NumOfZRecs As Long
  Dim ZipCnt As Long
  Dim ThisZip$
  Dim Big$, SaveBig$
  Dim Hold$, Nextz As Long
  Dim Thisx As Integer
  
  KillFile "ZIPIDX.DAT"
  KillFile "MORTIDX.DAT"
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenTaxBillFile TBHandle, NumOfTBRecs
  
  ReDim ThisZipArr(1 To 1) As String
  
  frmTaxShowPctComp.Label1 = "Sorting By Zip Code"
  frmTaxShowPctComp.Show , Me
  For x = 1 To NumOfTBRecs
    Get TBHandle, x, TBRec
    If TBRec.BillNumber <= 0 Then GoTo SkipIt
    ThisCustRec = TBRec.CustRec
    If ThisCustRec > 0 Then
      Get TCHandle, ThisCustRec, TaxCust
      ThisZip = QPTrim$(TaxCust.Zip)
      If Len(ThisZip) = 6 And Mid(ThisZip, 6, 1) = "-" Then
        ThisZip = Mid(ThisZip, 1, 5)
      End If
      For z = 1 To ThisZipCnt
        If ThisZip = ThisZipArr(z) Then Exit For
      Next z
      If z > ThisZipCnt Then
        ThisZipCnt = ThisZipCnt + 1
        ReDim Preserve ThisZipArr(1 To ThisZipCnt) As String
        ThisZipArr(ThisZipCnt) = ThisZip
      End If
    End If
SkipIt:
    frmTaxShowPctComp.ShowPctComp x, NumOfTBRecs
    If frmTaxShowPctComp.Out = True Then
      Close
      frmTaxShowPctComp.Out = False
      Unload frmTaxShowPctComp
      Exit Sub
    End If
  Next x
  
  Unload frmTaxShowPctComp
  GoSub SortLowToHigh
  frmTaxShowPctComp.Label1 = "Sorting By Zip Code"
  frmTaxShowPctComp.Show , Me
  
  ReDim ZipArray(1 To ThisZipCnt, 1 To NumOfTBRecs) As Long
  OpenZipIdxFile ZHandle, NumOfZRecs
  
  For x = 1 To NumOfTBRecs
    Get TBHandle, x, TBRec
    If TBRec.BillNumber > 0 Then
      ThisCustRec = TBRec.CustRec
      If ThisCustRec > 0 Then
        ThisZip = QPTrim$(TBRec.CustZip)
        If Len(ThisZip) = 6 And Mid(ThisZip, 6, 1) = "-" Then
          ThisZip = Mid(ThisZip, 1, 5)
        End If
        For z = 1 To ThisZipCnt
          If ThisZip = ThisZipArr(z) Then
            ZipArray(z, x) = x
            ZipCnt = ZipCnt + 1
            Exit For
          End If
        Next z
      End If
    End If
  Next x
  
  Nextz = 0
  For z = 1 To ThisZipCnt
    For x = 1 To NumOfTBRecs 'ZipCnt
      If ZipArray(z, x) > 0 Then
        ZipRec.TaxBillRec = x
        Nextz = Nextz + 1
        Put ZHandle, Nextz, ZipRec
      End If
    Next x
    frmTaxShowPctComp.ShowPctComp z, ThisZipCnt
    If frmTaxShowPctComp.Out = True Then
      Close
      frmTaxShowPctComp.Out = False
      Unload frmTaxShowPctComp
      Exit Sub
    End If
 Next z
  
  Unload frmTaxShowPctComp
  
  Close
  Exit Sub
  
SortLowToHigh:
  Big = "0"
  
  For x = 1 To ThisZipCnt
    If ThisZipArr(x) > Big Then
      Big = ThisZipArr(x)
    End If
  Next x
  
  Big = Big + "9"
  SaveBig = Big
  Nextx = 1
  Do While Nextx <= ThisZipCnt
    For x = Nextx To ThisZipCnt
      If ThisZipArr(x) < Big Then
        Big = ThisZipArr(x)
        Thisx = x
      End If
    Next x
    Hold = ThisZipArr(Nextx)
    ThisZipArr(Nextx) = ThisZipArr(Thisx)
    ThisZipArr(Thisx) = Hold
    Big = SaveBig
    Nextx = Nextx + 1
  Loop
  
  Return
    
End Sub


