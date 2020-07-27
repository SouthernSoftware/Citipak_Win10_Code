VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmVATaxCustInfoTHist 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Customer Information"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12360
   Icon            =   "frmVATaxCustInfoTHist.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   12360
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpList fpList1Pers 
      Height          =   1770
      Left            =   1560
      TabIndex        =   13
      Top             =   5160
      Width           =   9255
      _Version        =   196608
      _ExtentX        =   16325
      _ExtentY        =   3122
      TextAlias       =   ""
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
      ColumnHeaderShow=   -1  'True
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
      ColDesigner     =   "frmVATaxCustInfoTHist.frx":08CA
   End
   Begin LpLib.fpList fpList1Real 
      Height          =   1770
      Left            =   1545
      TabIndex        =   7
      Top             =   2280
      Width           =   9255
      _Version        =   196608
      _ExtentX        =   16325
      _ExtentY        =   3122
      TextAlias       =   ""
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
      ColumnHeaderShow=   -1  'True
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
      ColDesigner     =   "frmVATaxCustInfoTHist.frx":0C37
   End
   Begin EditLib.fpCurrency fpCurrRealBal 
      Height          =   372
      Left            =   6210
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1320
      Width           =   1932
      _Version        =   196608
      _ExtentX        =   3408
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
   Begin EditLib.fpCurrency fpCurrBalance 
      Height          =   372
      Left            =   9858
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1320
      Width           =   1932
      _Version        =   196608
      _ExtentX        =   3408
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
      Left            =   3233
      TabIndex        =   0
      TabStop         =   0   'False
      Tag             =   "Enter the official name of your town here. For example, 'Town Of Washington'."
      Top             =   720
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
      ControlType     =   0
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
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   540
      Left            =   3840
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   7800
      Width           =   1920
      _Version        =   131072
      _ExtentX        =   3387
      _ExtentY        =   952
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
      ButtonDesigner  =   "frmVATaxCustInfoTHist.frx":0FA4
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdDetail 
      Height          =   540
      Left            =   6600
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   7800
      Width           =   1920
      _Version        =   131072
      _ExtentX        =   3387
      _ExtentY        =   952
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
      ButtonDesigner  =   "frmVATaxCustInfoTHist.frx":1182
   End
   Begin EditLib.fpCurrency fpCurrPersBal 
      Height          =   372
      Left            =   2490
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1320
      Width           =   1932
      _Version        =   196608
      _ExtentX        =   3408
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
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Personal Transaction Count:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   252
      Left            =   1440
      TabIndex        =   15
      Top             =   7080
      Width           =   4452
   End
   Begin VB.Label Label7 
      BackColor       =   &H0080FFFF&
      Caption         =   "Personal Transactions:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   1452
      TabIndex        =   14
      Top             =   4800
      Width           =   2400
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   2292
      Left            =   1440
      Top             =   4800
      Width           =   9492
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080FFFF&
      Caption         =   "Real Transactions:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   1440
      TabIndex        =   12
      Top             =   1920
      Width           =   2052
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Personal Balance:"
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
      Left            =   570
      TabIndex        =   11
      Top             =   1440
      Width           =   1812
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Real Balance:"
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
      Left            =   4650
      TabIndex        =   9
      Top             =   1440
      Width           =   1452
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   2292
      Left            =   1428
      Top             =   1920
      Width           =   9492
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Real Transaction Count:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   252
      Left            =   1440
      TabIndex        =   4
      Top             =   4200
      Width           =   4212
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total Balance:"
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
      Left            =   8256
      TabIndex        =   3
      Top             =   1440
      Width           =   1572
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Customer Information"
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
      Left            =   3293
      TabIndex        =   1
      Top             =   360
      Width           =   6015
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   900
      Index           =   1
      Left            =   1853
      Top             =   300
      Width           =   8655
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   960
      Left            =   1853
      Top             =   240
      Width           =   8655
   End
End
Attribute VB_Name = "frmVATaxCustInfoTHist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class
  Public BillType As String

Private Sub cmdDetail_Click()
  Dim TRHandle As Integer
  Dim NumOfTaxTRecs As Long
  Dim TaxTRec As TaxTransactionType
  Dim ThisRec As Long
  
  If fpList1Pers.SelCount > 0 Then
    BillType = "Personal"
    If fpList1Pers.ListCount > 0 Then
      If fpList1Pers.SelCount = 0 Then
        Call TaxMsg(900, "Please make a selection from the personal transactions list.")
        Exit Sub
      End If
    Else
      Call TaxMsg(900, "No transactions available for personal detail report.")
      Exit Sub
    End If
    frmVATaxCustInfoTHist.fpList1Pers.Col = 3
    frmVATaxCustInfoTHist.fpList1Pers.Row = frmVATaxCustInfoTHist.fpList1Pers.ListIndex
    ThisRec = CLng(frmVATaxCustInfoTHist.fpList1Pers.ColText)
  ElseIf fpList1Real.SelCount > 0 Then
    BillType = "Real"
    If fpList1Real.ListCount > 0 Then
      If fpList1Real.SelCount = 0 Then
        Call TaxMsg(900, "Please make a selection from the real transactions list.")
        Exit Sub
      End If
    Else
      Call TaxMsg(900, "No transactions available for real detail report.")
      Exit Sub
    End If
    frmVATaxCustInfoTHist.fpList1Real.Col = 3
    frmVATaxCustInfoTHist.fpList1Real.Row = frmVATaxCustInfoTHist.fpList1Real.ListIndex
    ThisRec = CLng(frmVATaxCustInfoTHist.fpList1Real.ColText)
  Else
    Call TaxMsg(900, "Please make a selection from a transactions list.")
    Exit Sub
  End If
  
  If ThisRec > 0 Then
    OpenTaxTransFile TRHandle, NumOfTaxTRecs
    Get TRHandle, ThisRec, TaxTRec
    Close
    If TaxTRec.TranType = 1 Then
      If fpList1Pers.SelCount > 0 Then
        frmVATaxPersTransDetail.Show vbModal
      ElseIf fpList1Real.SelCount > 0 Then
        frmVATaxTransDetail.Show vbModal
      End If
    Else
      If fpList1Pers.SelCount > 0 Then
        frmVATaxPersTaxDetailNotBill.Show vbModal
      ElseIf fpList1Real.SelCount > 0 Then
        frmVATaxTransDetailNotBill.Show vbModal
      End If
    End If
  End If
End Sub

Private Sub cmdExit_Click()
'  If Exist("txpyment.dat") Then
'    frmVATaxPaymentEntry.Show
'    DoEvents
  Unload Me
  If Exist("C:\CPWork\txradjust.dat") Then
    frmVATaxAdjustments.Show
    DoEvents
  ElseIf Exist("C:\CPWork\txpadjust.dat") Then
    frmVATaxPAdjustments.Show
    DoEvents
  ElseIf Exist("C:\CPWork\custinq.dat") Then
    frmVATaxCustInq.Show
    DoEvents
  End If
  Unload Me
  frmVATaxCustAddEdit.Refresh '12/19/07
  
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If fpList1Real.ListIndex > -1 Then
    If KeyCode = vbKeyReturn Then
      Call cmdDetail_Click
      KeyCode = 0
    End If
  ElseIf fpList1Pers.ListIndex > -1 Then
    If KeyCode = vbKeyReturn Then
      Call cmdDetail_Click
      KeyCode = 0
    End If
  End If
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%C"
      Call cmdExit_Click
      KeyCode = 0
    Case vbKeyF5:
      SendKeys "%D"
      Call cmdDetail_Click
      KeyCode = 0
  End Select

End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Call LoadMe
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("CitiTaxes.exe terminated via menu bar on frmVATaxCustInfoTHist.")
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
  Dim TaxRec As TaxCustType
  Dim THandle As Integer
  Dim NumOfCustRecs As Long
  Dim x As Long
  Dim TaxTRec As TaxTransactionType
  Dim TRHandle As Integer
  Dim NumOfTaxTRecs As Long
  Dim PrevTranRec&
  Dim ThisDate$
  Dim ThisType$
  Dim ThisAmt#
  Dim ThisRec&
  Dim BillType$
  Dim TType$, Interest#, Penalty#
  Dim TaxYear$, BillNum$
  Dim RCnt As Integer
  Dim PCnt As Integer
  
  On Error GoTo ERRORSTUFF
  
  BillType = "NA"
  OpenTaxCustFile THandle, NumOfCustRecs
  Get THandle, GCustNum, TaxRec
  Close THandle
  
  fptxtThisCust.Text = QPTrim$(TaxRec.CustName)
  
  If TaxRec.LastTrans > 0 Then
    fpCurrBalance = GetCustBalance(GCustNum, -1)
    fpCurrRealBal = GetCustRealBalance(GCustNum, -1)
    fpCurrPersBal = GetCustPersBalance(GCustNum, -1)
  Else
    fpCurrBalance = 0
    fpCurrRealBal = 0
    fpCurrPersBal = 0
  End If
  
  OpenTaxTransFile TRHandle, NumOfTaxTRecs
  PrevTranRec& = TaxRec.LastTrans
  RCnt = 0
  PCnt = 0
  If PrevTranRec& > 0 Then
    Do While PrevTranRec& > 0
      Get TRHandle, PrevTranRec&, TaxTRec
      ThisDate = MakeRegDate(TaxTRec.TransDate)
      GoSub GetTransType
      ThisType = BillType
      If TaxTRec.TranType <> 1 Then
        ThisAmt = OldRound(TaxTRec.Amount + TaxTRec.DiscAmt)
      Else
        ThisAmt = TaxTRec.Amount
      End If
      
'      TaxTRec.CustPin = TaxTRec.CustPin
'      TaxTRec.CustomerRec = TaxTRec.CustomerRec
      If InStr(ThisType, "Credit Applied at Billing") Then
        ThisAmt = TaxTRec.Revenue.PrePaidUsed
      End If
      ThisRec = PrevTranRec&
      If TaxTRec.BillType = "R" Then
        fpList1Real.InsertRow = ThisDate & Chr(9) & ThisType & Chr(9) & Using$("$##,###,##0.00", ThisAmt) & Chr(9) & Using$("##########", ThisRec)
        RCnt = RCnt + 1
      Else
        fpList1Pers.InsertRow = ThisDate & Chr(9) & ThisType & Chr(9) & Using$("$##,###,##0.00", ThisAmt) & Chr(9) & Using$("##########", ThisRec)
        PCnt = PCnt + 1
      End If
      PrevTranRec& = TaxTRec.LastTrans
    Loop
  End If
  Close
  If fpList1Real.ListCount > 0 Then
    fpList1Real.ListIndex = 0
  ElseIf fpList1Pers.ListCount > 0 Then
    fpList1Pers.ListIndex = 0
  End If
  
  Label3.Caption = "Real Transaction Count: " + CStr(RCnt)
  Label8.Caption = "Personal Transaction Count: " + CStr(PCnt)
  Exit Sub
  
GetTransType:
  Select Case TaxTRec.TranType
  Case 1
    BillNum$ = ParseBillNum$(TaxTRec.Description)
    Select Case TaxTRec.BillType
    Case "R"
      BillType$ = "Real-Estate Bill: #" + BillNum
    Case "P"
      BillType$ = "Personal Property Bill: #" + BillNum
    Case "C"
      BillType$ = "Combined Bill: #" + BillNum
    Case "M"
      BillType$ = "Manual Bill: #" + BillNum
    Case Else
      BillType$ = "Bill: " + BillNum
    End Select
    TaxYear$ = QPTrim$(Str$(TaxTRec.TaxYear))
  Case 2
    BillNum$ = ParseBillNum$(TaxTRec.Description)
    If Len(BillNum$) = 0 Then
      If QPTrim$(TaxTRec.Description) = "Prepay" Then
        BillType = "Prepayment"
      Else
        BillType$ = "Payment ??? "
      End If
    Else
      If TaxTRec.Revenue.PrePaidAmt > 0 Then
        BillType = "Pre/Payment on: "
      Else
        BillType$ = "Payment on: "
      End If
    End If
    BillType$ = BillType$ + BillNum$
  Case 3
    BillNum$ = ParseBillNum$(TaxTRec.Description)
    BillType$ = "Release on: " + BillNum
  Case 4
    BillNum$ = ParseBillNum$(TaxTRec.Description)
    BillType$ = "Interest on: " + BillNum$
    Interest# = TaxTRec.Revenue.Interest#
  Case 5
    BillNum$ = ParseBillNum$(TaxTRec.Description)
    BillType$ = "Penalty on: " + BillNum$
    Penalty# = TaxTRec.Revenue.Penalty
  Case 6
    BillNum$ = ParseBillNum$(TaxTRec.Description)
    BillType$ = "Collection/Ad Cost on: " + BillNum$
  Case 7
    If TaxTRec.CustPin = 0 Then 'added 7/10/06
      If TaxTRec.Amount >= 0 Then
        BillType$ = "Adjust Bill Down"
      Else
        BillType$ = "Adjust Bill Up"
        TaxTRec.Amount = Abs(TaxTRec.Amount)
      End If
    Else
      BillNum$ = ParseBillNum$(TaxTRec.Description)
      BillType$ = "Adjust Paid Down on: " + BillNum$
    End If
  Case 9
    BillNum$ = ParseBillNum$(TaxTRec.Description)
    BillType$ = "Credit Applied at Billing on: #" + BillNum$
  Case 13
    BillNum$ = ParseBillNum$(TaxTRec.Description)
    BillType$ = "Adjust Bill Down on: " + BillNum$
  Case 14
    BillNum$ = ParseBillNum$(TaxTRec.Description)
    BillType$ = "Adjust Bill Up on: " + BillNum$
  Case 21
    BillNum$ = ParseBillNum$(TaxTRec.Description)
    BillType$ = "Paid Bill Plus Prepay on: " + BillNum$
  Case 22
    BillType$ = "Prepayment"
  Case 10
    BillNum$ = ParseBillNum$(TaxTRec.Description)
    BillType = "Adj Pay Down Affecting Credit on: " + BillNum$
  Case 11
    BillType = "Adjust Prepay Down"
  Case 12
    BillType = "Refund Prepay"
  Case 24
    BillNum$ = ParseBillNum$(TaxTRec.Description)
    BillType = "Adjust Bill Up Affecting Credit on: " + BillNum$
  Case 30
    BillNum$ = ParseBillNum$(TaxTRec.Description)
    BillType = "PPTRA Disc Removed From Bill#: " + BillNum$
  Case Else
    BillType$ = Str$(TaxTRec.TranType) + "??"
    
  End Select
  
  Return
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxCustInfoTHist", "LoadMe", Erl)
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

Private Sub fpList1Real_Click()
  fpList1Pers.Action = ActionDeselectAll
End Sub

Private Sub fpList1Real_DblClick()
  Dim TRHandle As Integer
  Dim NumOfTaxTRecs As Long
  Dim TaxTRec As TaxTransactionType
  Dim ThisRec As Long

  On Error GoTo ERRORSTUFF

  BillType = "Real"
  frmVATaxCustInfoTHist.fpList1Real.Col = 3
  frmVATaxCustInfoTHist.fpList1Real.Row = frmVATaxCustInfoTHist.fpList1Real.ListIndex
  ThisRec = CLng(frmVATaxCustInfoTHist.fpList1Real.ColText)

  If ThisRec > 0 Then
    OpenTaxTransFile TRHandle, NumOfTaxTRecs
    Get TRHandle, ThisRec, TaxTRec
    Close
    If TaxTRec.TranType = 1 Then
      frmVATaxTransDetail.Show vbModal
    Else
      frmVATaxTransDetailNotBill.Show vbModal
    End If
  Else
    frmVATaxTransDetailNotBill.Show vbModal
  End If

  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxCustInfoTHist", "fpList1Real_DblClick", Erl)
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

Private Sub fpList1Pers_Click()
  fpList1Real.Action = ActionDeselectAll
End Sub

Private Sub fpList1Pers_DblClick()
  Dim TRHandle As Integer
  Dim NumOfTaxTRecs As Long
  Dim TaxTRec As TaxTransactionType
  Dim ThisRec As Long

  On Error GoTo ERRORSTUFF
  
  BillType = "Personal"
  frmVATaxCustInfoTHist.fpList1Pers.Col = 3
  frmVATaxCustInfoTHist.fpList1Pers.Row = frmVATaxCustInfoTHist.fpList1Pers.ListIndex
  ThisRec = CLng(frmVATaxCustInfoTHist.fpList1Pers.ColText)

  If ThisRec > 0 Then
    OpenTaxTransFile TRHandle, NumOfTaxTRecs
    Get TRHandle, ThisRec, TaxTRec
    Close
    If TaxTRec.TranType = 1 Then
      frmVATaxPersTransDetail.Show vbModal
    Else
      frmVATaxPersTaxDetailNotBill.Show vbModal
    End If
  Else
    frmVATaxPersTaxDetailNotBill.Show vbModal
  End If

  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxCustInfoTHist", "fpList1Pers_DblClick", Erl)
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
