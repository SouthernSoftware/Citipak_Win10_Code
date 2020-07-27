VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmVATaxEditPen 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Penalty Transactions Editing"
   ClientHeight    =   8760
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   11655
   Icon            =   "frmVATaxEditPen.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcmbType 
      Height          =   384
      Left            =   4920
      TabIndex        =   23
      Top             =   1920
      Width           =   3372
      _Version        =   196608
      _ExtentX        =   5948
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
      ColDesigner     =   "frmVATaxEditPen.frx":08CA
   End
   Begin EditLib.fpCurrency fpCurrPen 
      Height          =   372
      Left            =   5724
      TabIndex        =   0
      Top             =   5748
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
      Left            =   5316
      TabIndex        =   1
      Top             =   3468
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
      Left            =   4260
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4188
      Width           =   4092
      _Version        =   196608
      _ExtentX        =   7218
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
      Left            =   5760
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2532
      Width           =   2172
      _Version        =   196608
      _ExtentX        =   3831
      _ExtentY        =   698
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
      Left            =   7536
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   4920
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
   Begin EditLib.fpText fptxtTaxYear 
      Height          =   396
      Left            =   3816
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4932
      Width           =   1212
      _Version        =   196608
      _ExtentX        =   2138
      _ExtentY        =   698
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
      Left            =   7056
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3492
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
      ButtonDesigner  =   "frmVATaxEditPen.frx":0BC1
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSave 
      Height          =   480
      Left            =   8220
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   6870
      Width           =   1695
      _Version        =   131072
      _ExtentX        =   2990
      _ExtentY        =   847
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
      ButtonDesigner  =   "frmVATaxEditPen.frx":0DA3
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   1092
      Left            =   1800
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   6840
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
      ButtonDesigner  =   "frmVATaxEditPen.frx":0F7F
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdDelete 
      Height          =   492
      Left            =   8220
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   7452
      Width           =   1692
      _Version        =   131072
      _ExtentX        =   2984
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
      ButtonDesigner  =   "frmVATaxEditPen.frx":115B
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPrev 
      Height          =   492
      Left            =   6060
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   6852
      Width           =   1692
      _Version        =   131072
      _ExtentX        =   2984
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
      ButtonDesigner  =   "frmVATaxEditPen.frx":1338
      Begin fpBtnAtlLibCtl.fpBtn fpBtn1 
         Height          =   495
         Left            =   0
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   840
         Visible         =   0   'False
         Width           =   10000
         _Version        =   131072
         _ExtentX        =   17639
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
         ButtonDesigner  =   "frmVATaxEditPen.frx":1513
      End
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdNext 
      Height          =   492
      Left            =   6060
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   7452
      Width           =   1692
      _Version        =   131072
      _ExtentX        =   2984
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
      ButtonDesigner  =   "frmVATaxEditPen.frx":16EF
      Begin fpBtnAtlLibCtl.fpBtn fpBtn3 
         Height          =   495
         Left            =   0
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   2000
         Visible         =   0   'False
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
         ButtonDesigner  =   "frmVATaxEditPen.frx":18C6
      End
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdHome 
      Height          =   492
      Left            =   3900
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   6852
      Width           =   1692
      _Version        =   131072
      _ExtentX        =   2984
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
      ButtonDesigner  =   "frmVATaxEditPen.frx":1AA2
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdEnd 
      Height          =   492
      Left            =   3900
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   7452
      Width           =   1692
      _Version        =   131072
      _ExtentX        =   2984
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
      ButtonDesigner  =   "frmVATaxEditPen.frx":1C79
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Select A Type:"
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
      Left            =   3360
      TabIndex        =   24
      Top             =   2040
      Width           =   1380
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Edit Penalty Transaction"
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
      Left            =   3090
      TabIndex        =   22
      Top             =   1236
      Width           =   5292
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   660
      Index           =   1
      Left            =   1470
      Top             =   1068
      Width           =   8652
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
      Left            =   3240
      TabIndex        =   21
      Top             =   4320
      Width           =   852
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
      Left            =   2784
      TabIndex        =   20
      Top             =   3588
      Width           =   2412
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
      Left            =   3720
      TabIndex        =   19
      Top             =   2640
      Width           =   1860
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
      Left            =   6456
      TabIndex        =   18
      Top             =   4992
      Width           =   972
   End
   Begin VB.Label Label6 
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
      Left            =   2616
      TabIndex        =   17
      Top             =   4992
      Width           =   1092
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Penalty:"
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
      Left            =   4056
      TabIndex        =   16
      Top             =   5844
      Width           =   1452
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   3120
      Left            =   1536
      Top             =   3240
      Width           =   8652
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   1536
      X2              =   10176
      Y1              =   5508
      Y2              =   5508
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   780
      Left            =   1470
      Top             =   960
      Width           =   8652
   End
End
Attribute VB_Name = "frmVATaxEditPen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Public RGPenRec As Long
  Public PGPenRec As Long
  Dim ThisManyRRecs As Long
  Dim ThisManyPRecs As Long
  'Private Temp_Class As Resize_Class
  Dim ThisRPen As Double
  Dim ThisPPen As Double
  Dim PrevOK As Boolean
  Dim NextOK As Boolean
  Dim ExitOK As Boolean

Private Sub cmdDelete_Click()
  Dim PenRec As PenaltyRecType
  Dim NumOfPRRecs As Long
  Dim PRHandle As Integer
  
  If fpLongAcctNum.Value = 0 Then
    Call TaxMsg(900, "Please supply a valid customer number.")
    Close
    fpLongAcctNum.SetFocus
    Exit Sub
  End If
  
  If TaxMsgWOpts(900, "Are you sure you want to delete this penalty transaction?", "F10 Delete", "ESC Don't Delete") = "abort" Then
    Unload frmVATaxMsgWOpts
    Exit Sub
  Else
    Unload frmVATaxMsgWOpts
  End If
  
  If fpcmbType.Text = "REAL" Then
    If RGPenRec = 0 Then Exit Sub
    OpenRPenRecFile PRHandle, NumOfPRRecs
    Get PRHandle, RGPenRec, PenRec
    PenRec.DelFlag = True
    Put PRHandle, RGPenRec, PenRec
    Close PRHandle
    ThisManyRRecs = ThisManyRRecs - 1
    Call TaxMsg(900, "This real record has been deleted successfully.")
  ElseIf fpcmbType.Text = "PERSONAL" Then
    If PGPenRec = 0 Then Exit Sub
    OpenPPenRecFile PRHandle, NumOfPRRecs
    Get PRHandle, PGPenRec, PenRec
    PenRec.DelFlag = True
    Put PRHandle, PGPenRec, PenRec
    Close PRHandle
    ThisManyPRecs = ThisManyPRecs - 1
    Call TaxMsg(900, "This personal record has been deleted successfully.")
  End If
  Call Clearscreen
  
End Sub

Private Sub cmdEnd_Click()
  Dim PenRec As PenaltyRecType
  Dim NumOfPRRecs As Long
  Dim PRHandle As Integer
  
  If fpcmbType.Text = "REAL" Then
    If ThisRPen <> CDbl(fpCurrPen.Value) And fpLongAcctNum.Value > 0 Then
      If TaxMsgWOpts(900, "Do you wish to save your data?", "F10 Save", "ESC Don't Save") = "abort" Then
        Unload frmVATaxMsgWOpts
      Else
        Unload frmVATaxMsgWOpts
        NextOK = True
        Call cmdSave_Click
      End If
    End If
  
    If RGPenRec = ThisManyRRecs Then Exit Sub
    RGPenRec = ThisManyRRecs
    OpenRPenRecFile PRHandle, NumOfPRRecs
    Get PRHandle, RGPenRec, PenRec
    GCustNum = PenRec.CustRec
    Close PRHandle
  ElseIf fpcmbType.Text = "PERSONAL" Then
    If ThisPPen <> CDbl(fpCurrPen.Value) And fpLongAcctNum.Value > 0 Then
      If TaxMsgWOpts(900, "Do you wish to save your data?", "F10 Save", "ESC Don't Save") = "abort" Then
        Unload frmVATaxMsgWOpts
      Else
        Unload frmVATaxMsgWOpts
        NextOK = True
        Call cmdSave_Click
      End If
    End If
  
    If PGPenRec = ThisManyPRecs Then Exit Sub
    PGPenRec = ThisManyPRecs
    OpenPPenRecFile PRHandle, NumOfPRRecs
    Get PRHandle, PGPenRec, PenRec
    GCustNum = PenRec.CustRec
    Close PRHandle
  End If
  Call LoadMeEdit
  
End Sub

Private Sub cmdExit_Click()
  Dim PenRec As PenaltyRecType
  Dim NumOfPRRecs As Long
  Dim PRHandle As Integer
  
  If fpcmbType.Text = "REAL" Then
    If ThisRPen <> CDbl(fpCurrPen.Value) And fpLongAcctNum.Value > 0 Then
      If TaxMsgWOpts(900, "Do you wish to save your data?", "F10 Save", "ESC Don't Save") = "abort" Then
        Unload frmVATaxMsgWOpts
      Else
        Unload frmVATaxMsgWOpts
        ExitOK = True
        Call cmdSave_Click
      End If
    End If
  ElseIf fpcmbType.Text = "PERSONAL" Then
    If ThisPPen <> CDbl(fpCurrPen.Value) And fpLongAcctNum.Value > 0 Then
      If TaxMsgWOpts(900, "Do you wish to save your data?", "F10 Save", "ESC Don't Save") = "abort" Then
        Unload frmVATaxMsgWOpts
      Else
        Unload frmVATaxMsgWOpts
        ExitOK = True
        Call cmdSave_Click
      End If
    End If
  End If
  ExitOK = True
  GCustNum = 0
  RGPenRec = 0
  PGPenRec = 0
  frmVATaxPenaltyMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdHome_Click()
  Dim PenRec As PenaltyRecType
  Dim NumOfPRRecs As Long
  Dim PRHandle As Integer
  
  If fpcmbType.Text = "REAL" Then
    If ThisRPen <> CDbl(fpCurrPen.Value) And fpLongAcctNum.Value > 0 Then
      If TaxMsgWOpts(900, "Do you wish to save your data?", "F10 Save", "ESC Don't Save") = "abort" Then
        Unload frmVATaxMsgWOpts
      Else
        Unload frmVATaxMsgWOpts
        NextOK = True
        Call cmdSave_Click
      End If
    End If
  
    If RGPenRec = 1 Then Exit Sub
    RGPenRec = 1
    OpenRPenRecFile PRHandle, NumOfPRRecs
    Get PRHandle, RGPenRec, PenRec
    GCustNum = PenRec.CustRec
    Close PRHandle
  ElseIf fpcmbType.Text = "PERSONAL" Then
    If ThisPPen <> CDbl(fpCurrPen.Value) And fpLongAcctNum.Value > 0 Then
      If TaxMsgWOpts(900, "Do you wish to save your data?", "F10 Save", "ESC Don't Save") = "abort" Then
        Unload frmVATaxMsgWOpts
      Else
        Unload frmVATaxMsgWOpts
        NextOK = True
        Call cmdSave_Click
      End If
    End If
  
    If PGPenRec = 1 Then Exit Sub
    PGPenRec = 1
    OpenRPenRecFile PRHandle, NumOfPRRecs
    Get PRHandle, PGPenRec, PenRec
    GCustNum = PenRec.CustRec
    Close PRHandle
  End If
  
  Call LoadMeEdit
End Sub

Private Sub cmdLookup_Click()
  If fpcmbType.Text = "REAL" Then
    If ThisRPen <> CDbl(fpCurrPen.Value) And fpLongAcctNum.Value > 0 Then
      If TaxMsgWOpts(900, "Do you wish to save your data?", "F10 Save", "ESC Don't Save") = "abort" Then
        Unload frmVATaxMsgWOpts
      Else
        Unload frmVATaxMsgWOpts
        Call cmdSave_Click
      End If
    End If
    frmVATaxCustLookupRPenOnly.Show
    DoEvents
  ElseIf fpcmbType.Text = "PERSONAL" Then
    If ThisPPen <> CDbl(fpCurrPen.Value) And fpLongAcctNum.Value > 0 Then
      If TaxMsgWOpts(900, "Do you wish to save your data?", "F10 Save", "ESC Don't Save") = "abort" Then
        Unload frmVATaxMsgWOpts
      Else
        Unload frmVATaxMsgWOpts
        Call cmdSave_Click
      End If
    End If
    frmVATaxCustLookupPPenOnly.Show
    DoEvents
  End If
End Sub

Private Sub cmdNext_Click()
  Dim PenRec As PenaltyRecType
  Dim NumOfPRRecs As Long
  Dim PRHandle As Integer
  
  If fpcmbType.Text = "REAL" Then
    If ThisRPen <> CDbl(fpCurrPen.Value) And fpLongAcctNum.Value > 0 Then
      If TaxMsgWOpts(900, "Do you wish to save your data?", "F10 Save", "ESC Don't Save") = "abort" Then
        Unload frmVATaxMsgWOpts
      Else
        Unload frmVATaxMsgWOpts
        NextOK = True
        Call cmdSave_Click
      End If
    End If
  
    If RGPenRec = ThisManyRRecs Then Exit Sub
    RGPenRec = RGPenRec + 1
    OpenRPenRecFile PRHandle, NumOfPRRecs
    Get PRHandle, RGPenRec, PenRec
    GCustNum = PenRec.CustRec
    Close PRHandle
  ElseIf fpcmbType.Text = "PERSONAL" Then
    If ThisPPen <> CDbl(fpCurrPen.Value) And fpLongAcctNum.Value > 0 Then
      If TaxMsgWOpts(900, "Do you wish to save your data?", "F10 Save", "ESC Don't Save") = "abort" Then
        Unload frmVATaxMsgWOpts
      Else
        Unload frmVATaxMsgWOpts
        NextOK = True
        Call cmdSave_Click
      End If
    End If
  
    If PGPenRec = ThisManyPRecs Then Exit Sub
    PGPenRec = PGPenRec + 1
    OpenPPenRecFile PRHandle, NumOfPRRecs
    Get PRHandle, PGPenRec, PenRec
    GCustNum = PenRec.CustRec
    Close PRHandle
  End If
  
  Call LoadMeEdit

End Sub

Private Sub cmdPrev_Click()
  Dim PenRec As PenaltyRecType
  Dim NumOfPRRecs As Long
  Dim PRHandle As Integer
  
  If fpcmbType.Text = "REAL" Then
    If ThisRPen <> CDbl(fpCurrPen.Value) And fpLongAcctNum.Value > 0 Then
      If TaxMsgWOpts(900, "Do you wish to save your data?", "F10 Save", "ESC Don't Save") = "abort" Then
        Unload frmVATaxMsgWOpts
      Else
        Unload frmVATaxMsgWOpts
        PrevOK = True
        Call cmdSave_Click
      End If
    End If
    If RGPenRec <= 1 Then Exit Sub
    RGPenRec = RGPenRec - 1
    OpenRPenRecFile PRHandle, NumOfPRRecs
    Get PRHandle, RGPenRec, PenRec
    GCustNum = PenRec.CustRec
    Close PRHandle
  ElseIf fpcmbType.Text = "PERSONAL" Then
    If ThisPPen <> CDbl(fpCurrPen.Value) And fpLongAcctNum.Value > 0 Then
      If TaxMsgWOpts(900, "Do you wish to save your data?", "F10 Save", "ESC Don't Save") = "abort" Then
        Unload frmVATaxMsgWOpts
      Else
        Unload frmVATaxMsgWOpts
        PrevOK = True
        Call cmdSave_Click
      End If
    End If
    If PGPenRec <= 1 Then Exit Sub
    PGPenRec = PGPenRec - 1
    OpenPPenRecFile PRHandle, NumOfPRRecs
    Get PRHandle, PGPenRec, PenRec
    GCustNum = PenRec.CustRec
    Close PRHandle
  End If
  
  Call LoadMeEdit
End Sub

Private Sub cmdSave_Click()
  Dim PenRec As PenaltyRecType
  Dim NumOfPRRecs As Long
  Dim PRHandle As Integer
  
  If fpLongAcctNum.Value = 0 Then
    Call TaxMsg(900, "Please supply a valid customer number.")
    fpLongAcctNum.SetFocus
    Exit Sub
  End If
  
  If fpcmbType.Text = "REAL" Then
    If RGPenRec = 0 Then Exit Sub
    OpenRPenRecFile PRHandle, NumOfPRRecs
    Get PRHandle, RGPenRec, PenRec
    PenRec.CustPin = PenRec.CustPin
    PenRec.Amount = CDbl(fpCurrPen.Value)
    Put PRHandle, RGPenRec, PenRec
    Close PRHandle
  ElseIf fpcmbType.Text = "PERSONAL" Then
    If PGPenRec = 0 Then Exit Sub
    OpenPPenRecFile PRHandle, NumOfPRRecs
    Get PRHandle, PGPenRec, PenRec
    
    PenRec.Amount = CDbl(fpCurrPen.Value)
    Put PRHandle, PGPenRec, PenRec
    Close PRHandle
  End If
  
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
  RGPenRec = 0
  PGPenRec = 0
  ExitOK = False
  PrevOK = False
  NextOK = False
  Me.HelpContextID = hlpEditPenalty
  Call LoadMe
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("CitiTaxes.exe terminated via menu bar on frmVATaxEditPen.")
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

Private Sub LoadMe()
  Dim PenRec As InterestRecType
  Dim NumOPRRecs As Long
  Dim PRHandle As Integer
  Dim RPenCnt As Long, DumbCnt As Long
  Dim PPenCnt As Long
  Dim x As Integer
  
  ThisManyRRecs = 0
  ThisManyPRecs = 0
  OpenRPenRecFile PRHandle, NumOPRRecs
  If NumOPRRecs = 0 Then
    Close PRHandle
    GoTo NoReal
  Else
    ThisManyRRecs = NumOPRRecs
  End If
  For x = 1 To NumOPRRecs
    Get PRHandle, x, NumOPRRecs
    If PenRec.DelFlag = True Then
      GoTo SkipIt
    Else
      RPenCnt = RPenCnt + 1
    End If
SkipIt:
  Next x
  If RPenCnt > 0 Then
    fpcmbType.AddItem "REAL"
    fpcmbType.Text = "REAL"
  End If
  Close PRHandle
  
NoReal:
  OpenPPenRecFile PRHandle, NumOPRRecs
  If NumOPRRecs = 0 Then
    Close PRHandle
    GoTo NoPers
  Else
    ThisManyPRecs = NumOPRRecs
  End If
  
  For x = 1 To NumOPRRecs
    Get PRHandle, x, NumOPRRecs
    If PenRec.DelFlag = True Then
      GoTo SkipIt2
    Else
      PPenCnt = PPenCnt + 1
    End If
SkipIt2:
  Next x
  If PPenCnt > 0 Then
    fpcmbType.AddItem "PERSONAL"
    If fpcmbType.Text <> "REAL" Then
      fpcmbType.Text = "PERSONAL"
    End If
  End If
  Close PRHandle

NoPers:
  If fpcmbType.Text = "REAL" Then
    ThisManyRRecs = RPenCnt
    fptxtRecord.Text = "0 of " + CStr(ThisManyRRecs)
  ElseIf fpcmbType.Text = "PERSONAL" Then
    ThisManyPRecs = PPenCnt
    fptxtRecord.Text = "0 of " + CStr(ThisManyPRecs)
  End If
  fpLongAcctNum = 0
  fptxtName.Text = ""
  fptxtTaxYear = 0
  fpDblSnglStartBill = 0
  fpCurrPen = 0
  ThisRPen = 0
  ThisPPen = 0
  
End Sub

Public Sub LoadMeEdit()
  Dim PenRec As PenaltyRecType
  Dim NumOfPRRecs As Long
  Dim PRHandle As Integer
  Dim x As Long
  
  If fpcmbType.Text = "REAL" Then
    If RGPenRec = 0 Then
      Call TaxMsg(900, "ERROR: There is a problem with the internal number assignment for this customer. Please try again.")
      fpLongAcctNum.SetFocus
      Exit Sub
    End If
  
    OpenRPenRecFile PRHandle, NumOfPRRecs
    Get PRHandle, RGPenRec, PenRec
    Close PRHandle
    fptxtRecord.Text = CStr(RGPenRec) + " of " + CStr(ThisManyRRecs)
    fpLongAcctNum = PenRec.CustRec
    fptxtName.Text = QPTrim$(PenRec.CustName)
    fptxtTaxYear = CStr(PenRec.TaxYear)
    fpDblSnglStartBill = PenRec.BillNumber
    fpCurrPen = PenRec.Amount
    ThisRPen = PenRec.Amount
  ElseIf fpcmbType.Text = "PERSONAL" Then
    If PGPenRec = 0 Then
      Call TaxMsg(900, "ERROR: There is a problem with the internal number assignment for this customer. Please try again.")
      fpLongAcctNum.SetFocus
      Exit Sub
    End If
  
    OpenPPenRecFile PRHandle, NumOfPRRecs
    Get PRHandle, PGPenRec, PenRec
    Close PRHandle
    fptxtRecord.Text = CStr(PGPenRec) + " of " + CStr(ThisManyPRecs)
    fpLongAcctNum = PenRec.CustRec
    fptxtName.Text = QPTrim$(PenRec.CustName)
    fptxtTaxYear = CStr(PenRec.TaxYear)
    fpDblSnglStartBill = PenRec.BillNumber
    fpCurrPen = PenRec.Amount
    ThisPPen = PenRec.Amount
  End If
End Sub

Private Sub Clearscreen()
  GCustNum = 0
  RGPenRec = 0
  PGPenRec = 0
  If fpcmbType.Text = "REAL" Then
    fptxtRecord.Text = "0 of " + CStr(ThisManyRRecs)
  ElseIf fpcmbType.Text = "PERSONAL" Then
    fptxtRecord.Text = "0 of " + CStr(ThisManyPRecs)
  End If
  fpLongAcctNum = 0
  fptxtName.Text = ""
  fptxtTaxYear = 0
  fpDblSnglStartBill = 0
  fpCurrPen = 0
  ThisRPen = 0
  ThisPPen = 0
End Sub

Private Sub fpCurrPen_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyHome Then
    Call cmdHome_Click
  ElseIf KeyCode = vbKeyEnd Then
    Call cmdEnd_Click
  End If

End Sub

Private Sub fpcmbType_Click()
  Dim PenRec As PenaltyRecType
  Dim NumOfPRRecs As Long
  Dim PRHandle As Integer
  
  If fpcmbType.Text = "REAL" Then
    If ThisPPen <> fpCurrPen And PGPenRec > 0 Then
      If TaxMsgWOpts(900, "Do you wish to save your changes?", "F10 Save", "ESC Don't Save") = "abort" Then
        Unload frmVATaxMsgWOpts
        Call Clearscreen
        Exit Sub
      Else
        OpenPPenRecFile PRHandle, NumOfPRRecs
        Get PRHandle, PGPenRec, PenRec
        PenRec.Amount = CDbl(fpCurrPen.Value)
        Put PRHandle, PGPenRec, PenRec
        Close PRHandle
        Call Savemsg(900, "Your data has been saved successfully.")
        Call Clearscreen
        Exit Sub
      End If
    Else
      Call Clearscreen
    End If
  ElseIf fpcmbType.Text = "PERSONAL" Then
    If ThisRPen <> fpCurrPen And RGPenRec > 0 Then
      If TaxMsgWOpts(900, "Do you wish to save your changes?", "F10 Save", "ESC Don't Save") = "abort" Then
        Unload frmVATaxMsgWOpts
        Call Clearscreen
        Exit Sub
      Else
        OpenRPenRecFile PRHandle, NumOfPRRecs
        Get PRHandle, RGPenRec, PenRec
        PenRec.Amount = CDbl(fpCurrPen.Value)
        Put PRHandle, RGPenRec, PenRec
        Close PRHandle
        Call Savemsg(900, "Your data has been saved successfully.")
        Call Clearscreen
        Exit Sub
      End If
    Else
      Call Clearscreen
    End If
  End If
End Sub

Private Sub fpLongAcctNum_Change()
  Dim PenRec As PenaltyRecType
  Dim NumOfPRRecs As Long
  Dim PRHandle As Integer
  
'  If fpcmbType.Text = "REAL" Then
'    If ThisRPen <> fpCurrPen And RGPenRec > 0 Then
'      If TaxMsgWOpts(900, "Do you wish to save your changes?", "F10 Save", "ESC Don't Save") = "abort" Then
'        Unload frmVATaxMsgWOpts
'        Call Clearscreen
'        Exit Sub
'      Else
'        OpenRPenRecFile PRHandle, NumOfPRRecs
'        Get PRHandle, RGPenRec, PenRec
'        PenRec.Amount = CDbl(fpCurrPen.Value)
'        Put PRHandle, RGPenRec, PenRec
'        Close PRHandle
'        Call Savemsg(900, "Your data has been saved successfully.")
'        Call Clearscreen
'        Exit Sub
'      End If
'    End If
'  ElseIf fpcmbType.Text = "PERSONAL" Then
'    If ThisPPen <> fpCurrPen And PGPenRec > 0 Then
'      If TaxMsgWOpts(900, "Do you wish to save your changes?", "F10 Save", "ESC Don't Save") = "abort" Then
'        Unload frmVATaxMsgWOpts
'        Call Clearscreen
'        Exit Sub
'      Else
'        OpenPPenRecFile PRHandle, NumOfPRRecs
'        Get PRHandle, PGPenRec, PenRec
'        PenRec.Amount = CDbl(fpCurrPen.Value)
'        Put PRHandle, PGPenRec, PenRec
'        Close PRHandle
'        Call Savemsg(900, "Your data has been saved successfully.")
'        Call Clearscreen
'        Exit Sub
'      End If
'    End If
'  End If

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
  Dim PenRec As PenaltyRecType
  Dim NumOfPRRecs As Long
  Dim PRHandle As Integer
  Dim x As Long
  
  On Error GoTo ERRORSTUFF
  
  If ExitOK = True Then Exit Sub
  ThisRec = CLng(fpLongAcctNum.Text)
  If ThisRec = GCustNum Then
    Exit Sub
  End If
  
  If fpcmbType.Text = "REAL" Then
    If ThisRPen <> CDbl(fpCurrPen.Value) And ThisRec > 0 Then
      If TaxMsgWOpts(900, "Do you wish to save your data?", "F10 Save", "ESC Don't Save") = "abort" Then
        Unload frmVATaxMsgWOpts
      Else
        Unload frmVATaxMsgWOpts
        Call cmdSave_Click
      End If
    End If
  
    OpenRPenRecFile PRHandle, NumOfPRRecs
    For x = 1 To NumOfPRRecs
      Get PRHandle, x, PenRec
      If PenRec.DelFlag = True Then GoTo SkipIt
      If PenRec.CustRec = CLng(fpLongAcctNum.Text) Then
        Exit For
      End If
SkipIt:
    Next x
  
    Close PRHandle
  
    If x > NumOfPRRecs Then
      Call TaxMsg(800, "The customer number entered could not be found in the penalty calculation records. Please try another number.")
      Call Clearscreen
'      fpLongAcctNum.Text = 0 'ThisRec
      fpLongAcctNum.SetFocus
      Exit Sub
    Else
      GCustNum = PenRec.CustRec
      RGPenRec = x
      Call LoadMeEdit
    End If
  ElseIf fpcmbType.Text = "PERSONAL" Then
    If ThisPPen <> CDbl(fpCurrPen.Value) And ThisRec > 0 Then
      If TaxMsgWOpts(900, "Do you wish to save your data?", "F10 Save", "ESC Don't Save") = "abort" Then
        Unload frmVATaxMsgWOpts
      Else
        Unload frmVATaxMsgWOpts
        Call cmdSave_Click
      End If
    End If
  
    OpenPPenRecFile PRHandle, NumOfPRRecs
    For x = 1 To NumOfPRRecs
      Get PRHandle, x, PenRec
      If PenRec.DelFlag = True Then GoTo SkipIt2
      If PenRec.CustRec = CLng(fpLongAcctNum.Text) Then
        Exit For
      End If
SkipIt2:
    Next x
  
    Close PRHandle
  
    If x > NumOfPRRecs Then
      Call TaxMsg(800, "The customer number entered could not be found in the penalty calculation records. Please try another number.")
      Call Clearscreen
'      fpLongAcctNum.Text = 0 'ThisRec
      fpLongAcctNum.SetFocus
      Exit Sub
    Else
      GCustNum = PenRec.CustRec
      PGPenRec = x
      Call LoadMeEdit
    End If
  End If
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxEditPen", "fpLongAcctNum_LostFocus", Erl)
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

