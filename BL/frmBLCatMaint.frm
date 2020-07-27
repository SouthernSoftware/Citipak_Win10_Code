VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmBLCatEdit 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Business License Category Maintenance"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmBLCatMaint.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   Tag             =   $"frmBLCatMaint.frx":08CA
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcmbCatType 
      Height          =   405
      Left            =   6870
      TabIndex        =   1
      ToolTipText     =   "Press F1 for help with this field."
      Top             =   1230
      Width           =   4035
      _Version        =   196608
      _ExtentX        =   7117
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
      ColDesigner     =   "frmBLCatMaint.frx":0A10
   End
   Begin EditLib.fpCurrency fpcurrRatePer 
      Height          =   390
      Left            =   9600
      TabIndex        =   4
      Top             =   6315
      Visible         =   0   'False
      Width           =   1260
      _Version        =   196608
      _ExtentX        =   2222
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
   Begin fpBtnAtlLibCtl.fpBtn cmdCodeList 
      Height          =   390
      Left            =   4650
      TabIndex        =   34
      ToolTipText     =   "Press F3 for help with this button."
      Top             =   1245
      Width           =   1365
      _Version        =   131072
      _ExtentX        =   2408
      _ExtentY        =   688
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
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
      ButtonDesigner  =   "frmBLCatMaint.frx":0D07
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSave 
      Height          =   648
      Left            =   8916
      TabIndex        =   33
      TabStop         =   0   'False
      ToolTipText     =   "Press F10 to commit the data on this screen to memory."
      Top             =   7848
      Width           =   1884
      _Version        =   131072
      _ExtentX        =   3323
      _ExtentY        =   1143
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
      ButtonDesigner  =   "frmBLCatMaint.frx":0EE7
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   645
      Left            =   6315
      TabIndex        =   32
      TabStop         =   0   'False
      ToolTipText     =   "Press ESC to exit this screen."
      Top             =   7845
      Width           =   1875
      _Version        =   131072
      _ExtentX        =   3307
      _ExtentY        =   1138
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
      ButtonDesigner  =   "frmBLCatMaint.frx":10C3
   End
   Begin EditLib.fpCurrency fpcurrRecUpTo 
      Height          =   345
      Index           =   0
      Left            =   3840
      TabIndex        =   6
      ToolTipText     =   "Press F1 for help with this field."
      Top             =   3720
      Width           =   2055
      _Version        =   196608
      _ExtentX        =   3625
      _ExtentY        =   609
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
   Begin EditLib.fpCurrency fpcurrFee 
      Height          =   390
      Left            =   6390
      TabIndex        =   3
      ToolTipText     =   "Press F1 for help with this field."
      Top             =   2520
      Width           =   1590
      _Version        =   196608
      _ExtentX        =   2815
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
   Begin EditLib.fpText fptxtCatCode 
      Height          =   390
      Left            =   2790
      TabIndex        =   0
      ToolTipText     =   "Press F1 for help with this field."
      Top             =   1230
      Width           =   1830
      _Version        =   196608
      _ExtentX        =   3238
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   -2147483637
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
      ThreeDTextHighlightColor=   -2147483637
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   1
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
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
      ControlType     =   0
      Text            =   ""
      CharValidationText=   "0 1 2 3 4 5 6 7 8 9 "
      MaxLength       =   5
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fptxtCatDesc 
      Height          =   390
      Left            =   4290
      TabIndex        =   2
      ToolTipText     =   "Press F1 for help with this field."
      Top             =   1710
      Width           =   5010
      _Version        =   196608
      _ExtentX        =   8826
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
      ThreeDInsideHighlightColor=   -2147483637
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
      ThreeDTextHighlightColor=   -2147483637
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   1
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
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
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   35
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fptxtRevGLAcctNum 
      Height          =   390
      Left            =   6240
      TabIndex        =   29
      ToolTipText     =   "Press F1 for help with this field."
      Top             =   6315
      Width           =   2985
      _Version        =   196608
      _ExtentX        =   5270
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
      ThreeDInsideHighlightColor=   -2147483637
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
      ThreeDTextHighlightColor=   -2147483637
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   1
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
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
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   150
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fptxtAcctsRec 
      Height          =   390
      Left            =   6240
      TabIndex        =   30
      ToolTipText     =   "Press F1 for help with this field."
      Top             =   6750
      Width           =   2985
      _Version        =   196608
      _ExtentX        =   5270
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
      ThreeDInsideHighlightColor=   -2147483637
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
      ThreeDTextHighlightColor=   -2147483637
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   1
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
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
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   150
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fptxtCashReceipt 
      Height          =   390
      Left            =   6240
      TabIndex        =   31
      ToolTipText     =   "Press F1 for help with this field."
      Top             =   7170
      Width           =   2985
      _Version        =   196608
      _ExtentX        =   5270
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   -2147483637
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
      ThreeDTextHighlightColor=   -2147483637
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   1
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
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
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   150
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fptxtPct 
      Height          =   345
      Index           =   0
      Left            =   6675
      TabIndex        =   7
      ToolTipText     =   "Press F1 for help with this field."
      Top             =   3720
      Width           =   1020
      _Version        =   196608
      _ExtentX        =   1799
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
      ThreeDInsideHighlightColor=   -2147483637
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
      ThreeDTextHighlightColor=   -2147483637
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   1
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
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
      ControlType     =   0
      Text            =   ""
      CharValidationText=   "1 2 3 4 5 6 7 8 9 0 . "
      MaxLength       =   7
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fptxtPct 
      Height          =   345
      Index           =   1
      Left            =   6675
      TabIndex        =   11
      ToolTipText     =   "Press F1 for help with this field."
      Top             =   4110
      Width           =   1020
      _Version        =   196608
      _ExtentX        =   1799
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
      ThreeDInsideHighlightColor=   -2147483637
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
      ThreeDTextHighlightColor=   -2147483637
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   1
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
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
      ControlType     =   0
      Text            =   ""
      CharValidationText=   "1 2 3 4 5 6 7 8 9 0 ."
      MaxLength       =   7
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fptxtPct 
      Height          =   345
      Index           =   2
      Left            =   6675
      TabIndex        =   15
      ToolTipText     =   "Press F1 for help with this field."
      Top             =   4485
      Width           =   1020
      _Version        =   196608
      _ExtentX        =   1799
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
      ThreeDInsideHighlightColor=   -2147483637
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
      ThreeDTextHighlightColor=   -2147483637
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   1
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
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
      ControlType     =   0
      Text            =   ""
      CharValidationText=   "1 2 3 4 5 6 7 8 9 0 ."
      MaxLength       =   7
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fptxtPct 
      Height          =   345
      Index           =   3
      Left            =   6675
      TabIndex        =   19
      ToolTipText     =   "Press F1 for help with this field."
      Top             =   4875
      Width           =   1020
      _Version        =   196608
      _ExtentX        =   1799
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
      ThreeDInsideHighlightColor=   -2147483637
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
      ThreeDTextHighlightColor=   -2147483637
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   1
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
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
      ControlType     =   0
      Text            =   ""
      CharValidationText=   "1 2 3 4 5 6 7 8 9 0 ."
      MaxLength       =   7
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fptxtPct 
      Height          =   345
      Index           =   4
      Left            =   6675
      TabIndex        =   23
      ToolTipText     =   "Press F1 for help with this field."
      Top             =   5250
      Width           =   1020
      _Version        =   196608
      _ExtentX        =   1799
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
      ThreeDInsideHighlightColor=   -2147483637
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
      ThreeDTextHighlightColor=   -2147483637
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   1
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
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
      ControlType     =   0
      Text            =   ""
      CharValidationText=   "1 2 3 4 5 6 7 8 9 0 ."
      MaxLength       =   7
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fptxtPct 
      Height          =   345
      Index           =   5
      Left            =   6675
      TabIndex        =   27
      ToolTipText     =   "Press F1 for help with this field."
      Top             =   5640
      Width           =   1020
      _Version        =   196608
      _ExtentX        =   1799
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
      ThreeDInsideHighlightColor=   -2147483637
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
      ThreeDTextHighlightColor=   -2147483637
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   1
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
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
      ControlType     =   0
      Text            =   ""
      CharValidationText=   "1 2 3 4 5 6 7 8 9 0 ."
      MaxLength       =   7
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpCurrency fpcurrRecUpTo 
      Height          =   345
      Index           =   1
      Left            =   3840
      TabIndex        =   10
      ToolTipText     =   "Press F1 for help with this field."
      Top             =   4110
      Width           =   2055
      _Version        =   196608
      _ExtentX        =   3625
      _ExtentY        =   609
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
      MaxValue        =   "999999999999"
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
   Begin EditLib.fpCurrency fpcurrRecUpTo 
      Height          =   345
      Index           =   2
      Left            =   3840
      TabIndex        =   14
      ToolTipText     =   "Press F1 for help with this field."
      Top             =   4485
      Width           =   2055
      _Version        =   196608
      _ExtentX        =   3625
      _ExtentY        =   609
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
   Begin EditLib.fpCurrency fpcurrRecUpTo 
      Height          =   345
      Index           =   3
      Left            =   3840
      TabIndex        =   18
      ToolTipText     =   "Press F1 for help with this field."
      Top             =   4875
      Width           =   2055
      _Version        =   196608
      _ExtentX        =   3625
      _ExtentY        =   609
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
   Begin EditLib.fpCurrency fpcurrRecUpTo 
      Height          =   345
      Index           =   4
      Left            =   3840
      TabIndex        =   22
      ToolTipText     =   "Press F1 for help with this field."
      Top             =   5250
      Width           =   2055
      _Version        =   196608
      _ExtentX        =   3625
      _ExtentY        =   609
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
   Begin EditLib.fpCurrency fpcurrRecUpTo 
      Height          =   345
      Index           =   5
      Left            =   3840
      TabIndex        =   26
      ToolTipText     =   "Press F1 for help with this field."
      Top             =   5640
      Width           =   2055
      _Version        =   196608
      _ExtentX        =   3625
      _ExtentY        =   609
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
   Begin EditLib.fpCurrency fpcurrBase 
      Height          =   345
      Index           =   0
      Left            =   1635
      TabIndex        =   5
      ToolTipText     =   "Press F1 for help with this field."
      Top             =   3720
      Width           =   1590
      _Version        =   196608
      _ExtentX        =   2815
      _ExtentY        =   614
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
   Begin EditLib.fpCurrency fpcurrBase 
      Height          =   345
      Index           =   1
      Left            =   1635
      TabIndex        =   9
      ToolTipText     =   "Press F1 for help with this field."
      Top             =   4110
      Width           =   1590
      _Version        =   196608
      _ExtentX        =   2815
      _ExtentY        =   614
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
   Begin EditLib.fpCurrency fpcurrBase 
      Height          =   345
      Index           =   2
      Left            =   1635
      TabIndex        =   13
      ToolTipText     =   "Press F1 for help with this field."
      Top             =   4485
      Width           =   1590
      _Version        =   196608
      _ExtentX        =   2815
      _ExtentY        =   614
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
   Begin EditLib.fpCurrency fpcurrBase 
      Height          =   345
      Index           =   3
      Left            =   1635
      TabIndex        =   17
      ToolTipText     =   "Press F1 for help with this field."
      Top             =   4875
      Width           =   1590
      _Version        =   196608
      _ExtentX        =   2815
      _ExtentY        =   614
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
   Begin EditLib.fpCurrency fpcurrBase 
      Height          =   345
      Index           =   4
      Left            =   1635
      TabIndex        =   21
      ToolTipText     =   "Press F1 for help with this field."
      Top             =   5250
      Width           =   1590
      _Version        =   196608
      _ExtentX        =   2815
      _ExtentY        =   614
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
   Begin EditLib.fpCurrency fpcurrBase 
      Height          =   345
      Index           =   5
      Left            =   1635
      TabIndex        =   25
      ToolTipText     =   "Press F1 for help with this field."
      Top             =   5640
      Width           =   1590
      _Version        =   196608
      _ExtentX        =   2815
      _ExtentY        =   614
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
   Begin EditLib.fpCurrency fpcurrOver 
      Height          =   345
      Index           =   0
      Left            =   8685
      TabIndex        =   8
      ToolTipText     =   "Press F1 for help with this field."
      Top             =   3720
      Width           =   1590
      _Version        =   196608
      _ExtentX        =   2815
      _ExtentY        =   614
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
   Begin EditLib.fpCurrency fpcurrOver 
      Height          =   345
      Index           =   1
      Left            =   8685
      TabIndex        =   12
      ToolTipText     =   "Press F1 for help with this field."
      Top             =   4110
      Width           =   1590
      _Version        =   196608
      _ExtentX        =   2815
      _ExtentY        =   614
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
   Begin EditLib.fpCurrency fpcurrOver 
      Height          =   345
      Index           =   2
      Left            =   8685
      TabIndex        =   16
      ToolTipText     =   "Press F1 for help with this field."
      Top             =   4485
      Width           =   1590
      _Version        =   196608
      _ExtentX        =   2815
      _ExtentY        =   614
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
   Begin EditLib.fpCurrency fpcurrOver 
      Height          =   345
      Index           =   3
      Left            =   8685
      TabIndex        =   20
      ToolTipText     =   "Press F1 for help with this field."
      Top             =   4875
      Width           =   1590
      _Version        =   196608
      _ExtentX        =   2815
      _ExtentY        =   614
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
   Begin EditLib.fpCurrency fpcurrOver 
      Height          =   345
      Index           =   4
      Left            =   8685
      TabIndex        =   24
      ToolTipText     =   "Press F1 for help with this field."
      Top             =   5250
      Width           =   1590
      _Version        =   196608
      _ExtentX        =   2815
      _ExtentY        =   614
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
   Begin EditLib.fpCurrency fpcurrOver 
      Height          =   345
      Index           =   5
      Left            =   8685
      TabIndex        =   28
      ToolTipText     =   "Press F1 for help with this field."
      Top             =   5640
      Width           =   1590
      _Version        =   196608
      _ExtentX        =   2815
      _ExtentY        =   614
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
   Begin fpBtnAtlLibCtl.fpBtn cmdGLList 
      Height          =   648
      Left            =   3708
      TabIndex        =   51
      TabStop         =   0   'False
      ToolTipText     =   "Press F5 to bring up a complete general ledger number listing."
      Top             =   7848
      Width           =   1884
      _Version        =   131072
      _ExtentX        =   3323
      _ExtentY        =   1143
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
      ButtonDesigner  =   "frmBLCatMaint.frx":12A1
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdHelp 
      Height          =   645
      Left            =   1110
      TabIndex        =   52
      TabStop         =   0   'False
      ToolTipText     =   "Press F1 to bring up a detailed help screen."
      Top             =   7845
      Width           =   1875
      _Version        =   131072
      _ExtentX        =   3307
      _ExtentY        =   1138
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
      ButtonDesigner  =   "frmBLCatMaint.frx":1480
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Category Maintenance"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2835
      TabIndex        =   53
      Top             =   555
      Width           =   6015
   End
   Begin VB.Line Line3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      X1              =   5565
      X2              =   11133
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line Line8 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      X1              =   630
      X2              =   1302
      Y1              =   2370
      Y2              =   2370
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "FLAT RATE TYPE/MULTIPLIER TYPE"
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
      Height          =   300
      Left            =   1440
      TabIndex        =   50
      Top             =   2235
      Width           =   3945
   End
   Begin VB.Line Line6 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      X1              =   5475
      X2              =   11139
      Y1              =   2370
      Y2              =   2370
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "RATE PER: (invisible for now)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   870
      Left            =   9555
      TabIndex        =   49
      Top             =   6690
      Visible         =   0   'False
      Width           =   1350
      WordWrap        =   -1  'True
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   6585
      Left            =   630
      Top             =   1125
      Width           =   10530
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ON AMOUNTS OVER"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   8250
      TabIndex        =   42
      Top             =   3390
      Width           =   2460
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PLUS %"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   6525
      TabIndex        =   43
      Top             =   3390
      Width           =   1305
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "FOR RECEIPTS UP TO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   3645
      TabIndex        =   44
      Top             =   3390
      Width           =   2460
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "BASE FEE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   1830
      TabIndex        =   45
      Top             =   3390
      Width           =   1260
   End
   Begin VB.Line Line4 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      X1              =   630
      X2              =   1830
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "GENERAL LEDGER ACCOUNTS:"
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
      Height          =   300
      Left            =   1920
      TabIndex        =   48
      Top             =   6075
      Width           =   3420
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      X1              =   630
      X2              =   1830
      Y1              =   3150
      Y2              =   3150
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      X1              =   9210
      X2              =   11130
      Y1              =   3150
      Y2              =   3150
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "IF THIS CODE IS BASED ON GROSS RECEIPTS (TYPE: STEP RATE)"
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
      Height          =   300
      Left            =   1965
      TabIndex        =   47
      Top             =   3000
      Width           =   7065
   End
   Begin VB.Label LabelFee 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "FLAT RATE AMOUNT:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   3795
      TabIndex        =   46
      Top             =   2610
      Width           =   2460
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cash Receipt G/L Account Number:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   2010
      TabIndex        =   41
      Top             =   7275
      Width           =   3900
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Accounts Receivable Number (If Any):"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   1395
      TabIndex        =   40
      Top             =   6840
      Width           =   4530
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Revenue G/L Account Number:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   2490
      TabIndex        =   39
      Top             =   6450
      Width           =   3420
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Category Desc:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   2370
      TabIndex        =   38
      Top             =   1800
      Width           =   1785
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Type:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   6045
      TabIndex        =   37
      Top             =   1320
      Width           =   690
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Category Code:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   870
      TabIndex        =   36
      Top             =   1320
      Width           =   1785
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Category Maintenance"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2820
      TabIndex        =   35
      Top             =   165
      Width           =   6015
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   750
      Index           =   1
      Left            =   1500
      Top             =   120
      Width           =   8655
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   840
      Left            =   1500
      Top             =   75
      Width           =   8655
   End
End
Attribute VB_Name = "frmBLCatEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsBLTextBoxOverrider
  Private Temp_Class As Resize_Class
  Dim AddFlag As Boolean
  Dim CatCodeNum$
  Dim TempRevNum$
  Dim TempAcctsRecNum$
  Dim TempCashNum$
  Dim AccMethodIsNone As Boolean
  Dim CatInUseFlag As Boolean
  Dim TempBaseRate(1 To 6) As Double
  Dim TempUpToAmt(1 To 6) As Double
  Dim TempPctAmt(1 To 6) As Double
  Dim TempAmtsOver(1 To 6) As Double
  Dim TempCode$
  Dim TempType$
  Dim TempDesc$
  Dim TempRate As Double
  Dim FirstSave As Boolean
  Dim FirstTimeThru As Boolean
  Dim ChangeFlag As Boolean
  
Private Sub cmdCodeList_Click()
  frmBLCategoryList.Show vbModal
  DoEvents
End Sub

Public Sub cmdExit_Click()
  Dim ChangeFlag As Boolean
  Dim CatFile As Integer
  Dim DoWhatFlag As SaveChangeOptions1
  Dim CatRec As ARNewCatCodeRecType
  
  On Error GoTo ERRORSTUFF
  
  If GCatNum = 0 Then GoTo CatNumIsZero 'user is exiting
  'without saving new record entries...also skips the change
  'check feature and if category list is open then the number
  'double clicked will be brought up to this screen
  
  OpenCatCodeFile CatFile
  Get CatFile, GCatNum, CatRec
  Close CatFile
  'this code is trapping for changes made but not saved...prevents the
  'user from exiting without saving changes when he thought he had already
  'saved them
  If QPTrim$(CatRec.CatCode) <> QPTrim$(fptxtCatCode.Text) Then
    ChangeFlag = True
    If fptxtCatCode.Enabled = True Then
      fptxtCatCode.SetFocus
    End If
    GoTo ChangeFound
  End If
  
  If QPTrim$(CatRec.CODEDESC) <> QPTrim$(fptxtCatDesc.Text) Then
    ChangeFlag = True
    fptxtCatDesc.SetFocus
    GoTo ChangeFound
  End If
    
  If CatRec.Fee <> fpcurrFee Then
    ChangeFlag = True
    fpcurrFee.SetFocus
    GoTo ChangeFound
  End If
    
  If QPTrim$(CatRec.CodeType) <> Mid(fpcmbCatType.Text, 1, 1) Then
    ChangeFlag = True
    If fpcmbCatType.Enabled = True Then
      fpcmbCatType.SetFocus
    End If
    GoTo ChangeFound
  End If
    
  If CatRec.BaseAmt1 <> fpcurrBase(0) Then
    ChangeFlag = True
    fpcurrBase(0).SetFocus
    GoTo ChangeFound
  End If
  
  If CatRec.Recpt1 <> fpcurrRecUpTo(0) Then
    ChangeFlag = True
    fpcurrRecUpTo(0).SetFocus
    GoTo ChangeFound
  End If
    
  If CatRec.Percent1 <> CSng(fptxtPct(0).Text) Then
    ChangeFlag = True
    fptxtPct(0).SetFocus
    GoTo ChangeFound
  End If
    
  If CatRec.Maximum1 <> fpcurrOver(0) Then
    ChangeFlag = True
    fpcurrOver(0).SetFocus
    GoTo ChangeFound
  End If
    
  If CatRec.BaseAmt2 <> fpcurrBase(1) Then
    ChangeFlag = True
    fpcurrBase(1).SetFocus
    GoTo ChangeFound
  End If
  
  If CatRec.Recpt2 <> fpcurrRecUpTo(1) Then
    ChangeFlag = True
    fpcurrRecUpTo(1).SetFocus
    GoTo ChangeFound
  End If
    
  If CatRec.Percent2 <> CSng(fptxtPct(1).Text) Then
    ChangeFlag = True
    fptxtPct(1).SetFocus
    GoTo ChangeFound
  End If
    
  If CatRec.Maximum2 <> fpcurrOver(1) Then
    ChangeFlag = True
    fpcurrOver(1).SetFocus
    GoTo ChangeFound
  End If
  
  If CatRec.BaseAmt3 <> fpcurrBase(2) Then
    ChangeFlag = True
    fpcurrBase(2).SetFocus
    GoTo ChangeFound
  End If
  
  If CatRec.Recpt3 <> fpcurrRecUpTo(2) Then
    ChangeFlag = True
    fpcurrRecUpTo(2).SetFocus
    GoTo ChangeFound
  End If
    
  If CatRec.Percent3 <> CSng(fptxtPct(2).Text) Then
    ChangeFlag = True
    fptxtPct(2).SetFocus
    GoTo ChangeFound
  End If
    
  If CatRec.Maximum3 <> fpcurrOver(2) Then
    ChangeFlag = True
    fpcurrOver(2).SetFocus
    GoTo ChangeFound
  End If
  
  If CatRec.BaseAmt4 <> fpcurrBase(3) Then
    ChangeFlag = True
    fpcurrBase(3).SetFocus
    GoTo ChangeFound
  End If
  
  If CatRec.Recpt4 <> fpcurrRecUpTo(3) Then
    ChangeFlag = True
    fpcurrRecUpTo(3).SetFocus
    GoTo ChangeFound
  End If
    
  If CatRec.Percent4 <> CSng(fptxtPct(3).Text) Then
    ChangeFlag = True
    fptxtPct(3).SetFocus
    GoTo ChangeFound
  End If
    
  If CatRec.Maximum4 <> fpcurrOver(3) Then
    ChangeFlag = True
    fpcurrOver(3).SetFocus
    GoTo ChangeFound
  End If
    
  If CatRec.BaseAmt5 <> fpcurrBase(4) Then
    ChangeFlag = True
    fpcurrBase(4).SetFocus
    GoTo ChangeFound
  End If
  
  If CatRec.Recpt5 <> fpcurrRecUpTo(4) Then
    ChangeFlag = True
    fpcurrRecUpTo(4).SetFocus
    GoTo ChangeFound
  End If
    
  If CatRec.Percent5 <> CSng(fptxtPct(4).Text) Then
    ChangeFlag = True
    fptxtPct(4).SetFocus
    GoTo ChangeFound
  End If
    
  If CatRec.Maximum5 <> fpcurrOver(4) Then
    ChangeFlag = True
    fpcurrOver(4).SetFocus
    GoTo ChangeFound
  End If
    
  If CatRec.BaseAmt6 <> fpcurrBase(5) Then
    ChangeFlag = True
    fpcurrBase(5).SetFocus
    GoTo ChangeFound
  End If
  
  If CatRec.Recpt6 <> fpcurrRecUpTo(5) Then
    ChangeFlag = True
    fpcurrRecUpTo(5).SetFocus
    GoTo ChangeFound
  End If
    
  If CatRec.Percent6 <> CSng(fptxtPct(5).Text) Then
    ChangeFlag = True
    fptxtPct(5).SetFocus
    GoTo ChangeFound
  End If
    
  If CatRec.Maximum6 <> fpcurrOver(5) Then
    ChangeFlag = True
    fpcurrOver(5).SetFocus
    GoTo ChangeFound
  End If
    
  '  CatRec.RateStep = Value#(Form$(29, 0), ecode)
  If GetGLNum(CatRec.REVGLNUM) <> QPTrim$(fptxtRevGLAcctNum.Text) Then
    ChangeFlag = True
    fptxtRevGLAcctNum.SetFocus
    GoTo ChangeFound
  End If
  
  If GetGLNum(CatRec.ARGLACCT) <> QPTrim$(fptxtAcctsRec.Text) Then
    ChangeFlag = True
    fptxtAcctsRec.SetFocus
    GoTo ChangeFound
  End If
    
  If GetGLNum(CatRec.CASHACCT) <> QPTrim$(fptxtCashReceipt.Text) Then
    ChangeFlag = True
    fptxtCashReceipt.SetFocus
    GoTo ChangeFound
  End If
    
ChangeFound:
  If ChangeFlag = True Then
    ChangeFlag = False
    ItemChangeFlag = True 'global
    DoWhatFlag = PromptSaveChanges(Me)
    Select Case DoWhatFlag
    Case SaveChangeOptions1.scoSaveChanges
      Call cmdSave_Click
      Exit Sub 'don't exit
    Case SaveChangeOptions1.scoReviewChanges 'review is just bringing back the current form
      If Exist("catlistopen.dat") Then
        Unload frmBLCategoryList
        KillFile "catlistopen.dat"
      End If
      Exit Sub
    Case SaveChangeOptions1.scoAbandonChanges 'abandon
      If Exist("catlistopen.dat") Then
        ItemChangeFlag = False 'this tells the cat list that it's OK to continue
        'with changing the data on this screen with the new tag number entered
        KillFile "catlistopen.dat"
        Exit Sub
      End If
      If frmBLCatCodeLookup.Visible = False Then
        frmBLCategoryMaintMenu.Show
      End If
      KillFile "categoryedit.dat"
      Call frmBLCatCodeLookup.RefreshSearchList
      Unload frmBLCatEdit
      Exit Sub
    Case Else:
    End Select
  End If
CatNumIsZero:

  If Exist("catlistopen.dat") Then
    KillFile ("catlistopen.dat")
    Exit Sub
  End If
  
  'this code sends the program back to the category
  'lookup screen if the user came from there (editing
  'an existing category) or to the category maintenance
  'menu (if this was a new addition).
  If frmBLCatCodeLookup.Visible = False Then
    frmBLCategoryMaintMenu.Show
  Else
    Call frmBLCatCodeLookup.RefreshSearchList
  End If
  KillFile "categoryedit.dat"
  
CatListOpen:
  DoEvents
  Unload frmBLCatEdit
  
  Exit Sub
  
ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLCatEdit", "cmdExit_Click", Erl)
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

Private Sub cmdPrint_Click()

End Sub

Private Sub cmdGLList_Click()
  If AccMethodIsNone = False Then
    frmBLGLList.Show vbModal
  Else
    frmBLMessageBoxJr.Label1.Caption = "Since the accounting method selected on the Town Setup screen is 'None' the GL List button is not needed and therefore is disabled."
    frmBLMessageBoxJr.Label1.Top = 700
    frmBLMessageBoxJr.Show vbModal
  End If
End Sub

Private Sub cmdHelp_Click()
  frmBLCatEditHelpS.Show vbModal
End Sub

Private Sub cmdSave_Click()
  Dim ARCatCodeRec As ARNewCatCodeRecType
  Dim CHandle As Integer
  Dim SaveHere As Integer
  Dim CatRecNums As Integer
  Dim CatFile As Integer
  Dim CatNumChangeFlag As Boolean
  Dim IdxFlag As Boolean
  Dim TownRec As TownSetUpType
  Dim TownHandle As Integer
  Dim x As Integer
  Dim Row$
  Dim MaxRev As Double
  Dim BaseFee As Double
  Dim ThisCnt As Integer
  
  On Error GoTo ERRORSTUFF
  
  ThisCnt = 0
  
  If Not Exist("artmppst.dat") Then GoTo DontWarn
  If Mid(fpcmbCatType.Text, 1, 1) <> QPTrim$(TempType$) Then GoTo Warn
  If fpcurrFee.DoubleValue <> TempRate Then GoTo Warn
  For x = 0 To 5
    If fpcurrBase(x).DoubleValue <> TempBaseRate(x + 1) Then GoTo Warn
    If fpcurrRecUpTo(x).DoubleValue <> TempUpToAmt(x + 1) Then GoTo Warn
    If Val(fptxtPct(x)) <> TempPctAmt(x + 1) Then GoTo Warn
    If fpcurrOver(x).DoubleValue <> TempAmtsOver(x + 1) Then GoTo Warn
  Next x
  
  GoTo DontWarn
  
Warn:
  If CatInTempFile(QPTrim(fptxtCatCode.Text), ThisCnt) = True Then
    If ThisCnt = 1 Then
      frmBLMessageBoxJrWOpts.Label1.Caption = "There is " + CStr(ThisCnt) + " customer using this category that is included in an unposted business license fee file. Saving changes for this category would render the unposted file inaccurate for this customer. If you wish to continue saving then the unposted file will be deleted and business license registers will have to be run again. Do you wish to continue anyway?"
    Else
      frmBLMessageBoxJrWOpts.Label1.Caption = "There are " + CStr(ThisCnt) + " customers using this category that are included in an unposted business license fee file. Saving changes for this category would render the unposted file inaccurate for these customers. If you wish to continue saving then the unposted file will be deleted and business license registers will have to be run again. Do you wish to continue anyway?"
    End If
    frmBLMessageBoxJrWOpts.Label1.Top = 400
    frmBLMessageBoxJrWOpts.Label1.Height = 1700
    frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Continue"
    frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC Abort"
    frmBLMessageBoxJrWOpts.Show vbModal
    If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "abort" Then
      Close
      Unload frmBLMessageBoxJrWOpts
      fpcmbCatType.SetFocus
      Exit Sub
    Else
      Unload frmBLMessageBoxJrWOpts
      KillFile "artmppst.dat"
      frmBLMessageBoxJr.Label1.Caption = "The temporary unposted business license file 'artmppst.dat' has been deleted. License registers will have to be reprocessed."
      frmBLMessageBoxJr.Label1.Top = 800
      frmBLMessageBoxJr.Show vbModal
      MainLog ("The user attempted to save category code # " + QPTrim$(fptxtCatCode.Text) + " and was alerted that this category has customers using it that are included in an unposted license file. They were warned that if they continued then the unposted file would be deleted. They elected to continue anyway.")
    End If
  End If
  
DontWarn:
  If GCatNum = 0 Then
    fptxtCatCode.BackColor = &H80FFFF
    frmBLMessageBoxJrWOpts.Label1.Caption = "Category codes can be edited as long as no customers are using them. Once customer data is saved using this category then category codes can no longer be edited. Press F10 to continue this save procedure. Otherwise press ESC to return to the category maintenance screen to review the category code entry."
    frmBLMessageBoxJrWOpts.Label1.Top = 400
    frmBLMessageBoxJrWOpts.Label1.Height = 1500
    frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Continue"
    frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC Escape"
    frmBLMessageBoxJrWOpts.Show vbModal
    If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "abort" Then
      fptxtCatCode.BackColor = &HFFFFFF
      Unload frmBLMessageBoxJrWOpts
      If fptxtCatCode.Enabled = True Then
        fptxtCatCode.SetFocus
      End If
      Exit Sub
    Else
      fptxtCatCode.BackColor = &HFFFFFF
      Unload frmBLMessageBoxJrWOpts
    End If
  End If
  
  'if a user is changing the type for a category that
  'is currently being used by a customer then that customer
  'will need to be edited to comply with the new type rate structure
  '...this code sends up a pop-up warning/list of all customers
  'who will need to have their data edited
  If GCatNum > 0 Then
    If CatInUseFlag = True Then
      If TempType$ <> Mid(fpcmbCatType.Text, 1, 1) Then
        frmBLTypeChngPrintOut.Show vbModal
        If frmBLTypeChngPrintOut.fptxtChoice.Text = "abort" Then
          Unload frmBLTypeChngPrintOut
          Close
          fpcmbCatType.SetFocus
          Exit Sub
        Else
          Unload frmBLTypeChngPrintOut
        End If
      End If
    End If
  End If
  
  OpenTownFile TownHandle
  Get TownHandle, 1, TownRec
  Close TownHandle
  
  If QPTrim$(fpcmbCatType.Text) = "Step Rate" Then
    For x = 1 To 5
      If fpcurrRecUpTo(x).DoubleValue > 0 Then
        If fpcurrRecUpTo(x - 1).DoubleValue = 0 Then
          fpcurrRecUpTo(x - 1).BackColor = &H8080FF
          fpcurrRecUpTo(x).BackColor = &H80FFFF
          frmBLSpecMsgBox.Label1.Caption = "Please do not leave a 'For Receipts Up To' field with a zero value (red) if the next row (yellow) is not a zero value."
          frmBLSpecMsgBox.Label1.Top = 700
          frmBLSpecMsgBox.Show vbModal
          fpcurrRecUpTo(x - 1).BackColor = &HFFFFFF
          fpcurrRecUpTo(x).BackColor = &HFFFFFF
          fpcurrRecUpTo(x - 1).SetFocus
          Close
          Exit Sub
        End If
      End If
    Next x
  End If
      
  If QPTrim$(fpcmbCatType.Text) = "Step Rate" Then
    For x = 0 To 5
      BaseFee = BaseFee + fpcurrBase(x).DoubleValue
    Next x
    If BaseFee = 0 Then
      frmBLMessageBoxJrWOpts.Label1.Caption = "You have selected 'Step Rate' for the category method but no base fees have been saved. Do you wish to continue saving anyway?"
      frmBLMessageBoxJrWOpts.Label1.Top = 700
      frmBLMessageBoxJrWOpts.Show vbModal
      If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "abort" Then
        Close
        Unload frmBLMessageBoxJrWOpts
        Exit Sub
      Else
        Unload frmBLMessageBoxJrWOpts
        MainLog ("User warned that they are saving a step rate category type but all base fees are zero. User elected to continue anyway.")
      End If
    End If
  End If
      
  If QPTrim$(fpcmbCatType.Text) = "Step Rate" Then
    For x = 1 To 5
      If fpcurrRecUpTo(x).DoubleValue = 0 Then Exit For
      If fpcurrRecUpTo(x).DoubleValue < fpcurrRecUpTo(x - 1).DoubleValue Then
        fpcurrRecUpTo(x - 1).BackColor = &H8080FF
        fpcurrRecUpTo(x).BackColor = &H80FFFF
        frmBLSpecMsgBox.Label1.Caption = "The amount in the 'For Receipts Up To' field on row # " + CStr(x) + " is more than the amount in the same field on row # " + CStr(x + 1) + ". Please make sure that these amounts increase with each valid row."
        frmBLSpecMsgBox.Label1.Top = 600
        frmBLSpecMsgBox.Show vbModal
        fpcurrRecUpTo(x - 1).BackColor = &HFFFFFF
        fpcurrRecUpTo(x).BackColor = &HFFFFFF
        fpcurrRecUpTo(x - 1).SetFocus
        Close
        Exit Sub
      End If
    Next x
  End If

  If GLNumsOK = False Then
    Exit Sub
  End If
  
  If GLNumsValid = False Then
    Exit Sub
  End If
  
  'Checks to make sure that the category number entered
  'does not exist in the category number list already
  If Check4ValidCat(QPTrim$(fptxtCatCode.Text), QPTrim$(fptxtCatDesc.Text), GCatNum) = False Then
    Exit Sub
  End If
  
  'this code verifies that a fee amount is saved if a category
  'is saved
  Select Case Mid(fpcmbCatType.Text, 1, 1)
    Case "M"
      If fpcurrFee.DoubleValue = 0 Then
        frmBLMessageBoxJrWOpts.Label1.Caption = "You have selected 'Multiplier' as the category code type but the 'RATE PER UNIT' field is 0.00. Do you wish to enter a new value in the 'Rate Per' field?"
        frmBLMessageBoxJrWOpts.Label1.Top = 700
        frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 New Value"
        frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC Cancel"
        frmBLMessageBoxJrWOpts.Show vbModal
        If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "continue" Then
          Unload frmBLMessageBoxJrWOpts
          If fpcurrRatePer.Enabled = True Then
            fpcurrFee.SetFocus
          End If
          Exit Sub
        Else
          Unload frmBLMessageBoxJrWOpts
        End If
      End If
    Case "F"
      If fpcurrFee.DoubleValue = 0 Then
        frmBLMessageBoxJrWOpts.Label1.Caption = "You have selected 'Flat Rate' as the category code type but the 'FLAT RATE AMOUNT' field is 0.00. Do you wish to enter a new value in the 'Fee Amount' field?"
        frmBLMessageBoxJrWOpts.Label1.Top = 700
        frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 New Value"
        frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC Save Zero"
        frmBLMessageBoxJrWOpts.Show vbModal
        If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "continue" Then
          Unload frmBLMessageBoxJrWOpts
          If fpcurrFee.Enabled = True Then
            fpcurrFee.SetFocus
          End If
          Exit Sub
        Else
          Unload frmBLMessageBoxJrWOpts
        End If
      End If
  End Select
  
  If QPTrim$(fptxtCatCode.Text) = "" Then
    frmBLMessageBoxJr.Label1.Caption = "Please enter a category code number."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    If fptxtCatCode.Enabled = True Then
      fptxtCatCode.SetFocus
    End If
    Exit Sub
  End If
  
  If QPTrim$(fptxtCatDesc.Text) = "" Then
    frmBLMessageBoxJr.Label1.Caption = "Please enter a category description."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    fptxtCatDesc.SetFocus
    Exit Sub
  End If
  
  CatNumChangeFlag = False
  IdxFlag = False
  
  'this code saves this data in the correct record
  OpenCatCodeFile CatFile
  CatRecNums = LOF(CatFile) / Len(ARCatCodeRec)
  If GCatNum = 0 Then
    AddFlag = True
    IdxFlag = True
    SaveHere = CatRecNums + 1
  Else
    Get CatFile, GCatNum, ARCatCodeRec
    SaveHere = GCatNum
  End If
  
  ARCatCodeRec.CatCode = QPTrim$(fptxtCatCode.Text)
  ARCatCodeRec.CODEDESC = QPTrim$(fptxtCatDesc.Text) 'actually only saves the first letter
  If fpcmbCatType.Text = "Step Rate" Then
    fpcurrFee = 0
  ElseIf fpcmbCatType.Text = "Multiplier" Or fpcmbCatType.Text = "Flat Rate" Then
    For x = 0 To 5
      fpcurrBase(x) = 0
      fpcurrRecUpTo(x) = 0
      fptxtPct(x).Text = 0
      fpcurrOver(x) = 0
    Next x
  End If
  ARCatCodeRec.Fee = fpcurrFee
  ARCatCodeRec.CodeType = QPTrim$(fpcmbCatType.Text)
  ARCatCodeRec.BaseAmt1 = fpcurrBase(0)
  ARCatCodeRec.Recpt1 = fpcurrRecUpTo(0)
  ARCatCodeRec.Percent1 = Val(fptxtPct(0).Text)
  ARCatCodeRec.Maximum1 = fpcurrOver(0)
  ARCatCodeRec.BaseAmt2 = fpcurrBase(1)
  ARCatCodeRec.Recpt2 = fpcurrRecUpTo(1)
  ARCatCodeRec.Percent2 = Val(fptxtPct(1).Text)
  ARCatCodeRec.Maximum2 = fpcurrOver(1)
  ARCatCodeRec.BaseAmt3 = fpcurrBase(2)
  ARCatCodeRec.Recpt3 = fpcurrRecUpTo(2)
  ARCatCodeRec.Percent3 = Val(fptxtPct(2).Text)
  ARCatCodeRec.Maximum3 = fpcurrOver(2)
  ARCatCodeRec.BaseAmt4 = fpcurrBase(3)
  ARCatCodeRec.Recpt4 = fpcurrRecUpTo(3)
  ARCatCodeRec.Percent4 = Val(fptxtPct(3).Text)
  ARCatCodeRec.Maximum4 = fpcurrOver(3)
  ARCatCodeRec.BaseAmt5 = fpcurrBase(4)
  ARCatCodeRec.Recpt5 = fpcurrRecUpTo(4)
  ARCatCodeRec.Percent5 = Val(fptxtPct(4).Text)
  ARCatCodeRec.Maximum5 = fpcurrOver(4)
  ARCatCodeRec.BaseAmt6 = fpcurrBase(5)
  ARCatCodeRec.Recpt6 = fpcurrRecUpTo(5)
  ARCatCodeRec.Percent6 = Val(fptxtPct(5).Text)
  ARCatCodeRec.Maximum6 = fpcurrOver(5)
  ARCatCodeRec.RateStep = fpcurrFee

  ARCatCodeRec.REVGLNUM = GetGLRecNum(fptxtRevGLAcctNum.Text)
  If ARCatCodeRec.REVGLNUM = 0 Then
    If QPTrim$(TownRec.AcctMeth) = "A" Or QPTrim$(TownRec.AcctMeth) = "C" Then
      fptxtRevGLAcctNum.SetFocus
        frmBLMessageBoxJr.Label1.Caption = "Could not verify the 'Revenue' number (" + QPTrim$(fptxtRevGLAcctNum.Text) + ") because General Ledger list is not available. 'Revenue' number not saved."
        frmBLMessageBoxJr.Label1.Top = 700
        frmBLMessageBoxJr.Show vbModal
        fptxtRevGLAcctNum.Text = ""
      If Not Exist("GLACCT.DAT") Or Not Exist("GLACCT.IDX") Then
        MainLog ("User warned that the GL revenue number entered, " + QPTrim$(fptxtRevGLAcctNum.Text) + ", is not being saved because no match in the GL list could be found. GL data is missing.")
      Else
        frmBLMessageBoxJr.Label1.Caption = "No match for the 'Revenue' number (" + QPTrim$(fptxtRevGLAcctNum.Text) + ") found in the General Ledger number list. 'Revenue' number not saved."
        frmBLMessageBoxJr.Label1.Top = 700
        frmBLMessageBoxJr.Show vbModal
        fptxtRevGLAcctNum.Text = ""
        MainLog ("User warned that the GL revenue number entered, " + QPTrim$(fptxtRevGLAcctNum.Text) + ", is not being saved because no match in the GL list could be found.")
      End If
    End If
  End If
  
  ARCatCodeRec.ARGLACCT = GetGLRecNum(fptxtAcctsRec.Text)
  If ARCatCodeRec.ARGLACCT = 0 Then
    If QPTrim$(TownRec.AcctMeth) = "A" Then
      fptxtAcctsRec.SetFocus
      If Not Exist("GLACCT.DAT") Or Not Exist("GLACCT.IDX") Then
        frmBLMessageBoxJr.Label1.Caption = "Could not verify the 'Accts Rec' number (" + QPTrim$(fptxtAcctsRec.Text) + ") because General Ledger list is not available. 'Accts Rec' number not saved."
        frmBLMessageBoxJr.Label1.Top = 700
        frmBLMessageBoxJr.Show vbModal
        fptxtAcctsRec.Text = ""
        MainLog ("User warned that the GL accounts receivable number entered, " + QPTrim$(fptxtAcctsRec.Text) + ", is not being saved because it could not be verified. GL data is missing.")
      Else
        frmBLMessageBoxJr.Label1.Caption = "No match for the 'Accts Rec' number (" + QPTrim$(fptxtAcctsRec) + ") found in the General Ledger number list. 'Accts Rec' number not saved."
        frmBLMessageBoxJr.Label1.Top = 700
        frmBLMessageBoxJr.Show vbModal
        fptxtAcctsRec.Text = ""
        MainLog ("User warned that the GL accounts receivable number entered, " + QPTrim$(fptxtAcctsRec.Text) + ", is not being saved because it could not be verified.")
      End If
    End If
  End If
  
  ARCatCodeRec.CASHACCT = GetGLRecNum(fptxtCashReceipt.Text)
  If ARCatCodeRec.CASHACCT = 0 Then
    If QPTrim$(TownRec.AcctMeth) = "A" Or QPTrim$(TownRec.AcctMeth) = "C" Then
      fptxtCashReceipt.SetFocus
      If Not Exist("GLACCT.DAT") Or Not Exist("GLACCT.IDX") Then
        frmBLMessageBoxJr.Label1.Caption = "Could not verify the 'Cash Receipt' number (" + QPTrim$(fptxtCashReceipt) + ") because General Ledger list is not available. 'Cash Receipt' number not saved."
        frmBLMessageBoxJr.Label1.Top = 700
        frmBLMessageBoxJr.Show vbModal
        fptxtCashReceipt.Text = ""
        MainLog ("User warned that the GL cash receipts number entered, " + QPTrim$(fptxtCashReceipt.Text) + ", is not being saved because it could not be verified. GL data is missing.")
      Else
        frmBLMessageBoxJr.Label1.Caption = "No match for the 'Cash Receipt' number (" + QPTrim$(fptxtCashReceipt) + ") found in the General Ledger number list. 'Cash Receipt' number not saved."
        frmBLMessageBoxJr.Label1.Top = 700
        frmBLMessageBoxJr.Show vbModal
        fptxtCashReceipt.Text = ""
        MainLog ("User warned that the GL cash receipts number entered, " + QPTrim$(fptxtCashReceipt.Text) + ", is not being saved because it could not be verified.")
      End If
    End If
  End If
  
  Put CatFile, SaveHere, ARCatCodeRec
  Close CatFile
  Call LogSave
  
  Call CreateCatCodeIdx

  If IdxFlag = True Or CatNumChangeFlag = True Then
    IdxFlag = False
    CatNumChangeFlag = False
  End If
  
  frmBLSucSave.Label1.Caption = "Your category data has been saved successfully."
  frmBLSucSave.Label1.Top = 700
  frmBLSucSave.Show vbModal
  
  If Exist("catlistopen.dat") Then '
    KillFile ("catlistopen.dat")
    ItemChangeFlag = False
    If fptxtCatCode.Enabled = True Then
      fptxtCatCode.SetFocus
    End If
    Exit Sub
  End If
  
  If AddFlag = True Then 'entering a list of several items is tedious
  'if after each save the program returns to the menu so this feature allows
  'the user to speed up the entry process
    AddFlag = False
    frmBLMessageBoxJrWOpts.Label1.Caption = "Do you wish to add another new category?"
    frmBLMessageBoxJrWOpts.Label1.Top = 900
    frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Add New"
    frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC Cancel"
    frmBLMessageBoxJrWOpts.Show vbModal
    If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "continue" Then
      Unload frmBLMessageBoxJrWOpts
      GCatNum = 0
      If fptxtCatCode.Enabled = True Then
        fptxtCatCode.SetFocus
      End If
      Call LoadMe
      Exit Sub
    Else
      Unload frmBLMessageBoxJrWOpts
      DoEvents
      Load frmBLCategoryMaintMenu
      frmBLCategoryMaintMenu.Show
      KillFile ("categoryedit.dat")
      DoEvents
      Unload frmBLCatEdit
      Exit Sub
    End If
  Else 'just editing an existing category...sends user back to lookup upon
  'completion
    'RefreshSearchList is a sub in the frmBLCatCodeLookup
    'that reloads the list with the latest saved data
    'and returns the user to the list with the most current category
    'highlighted
    Call frmBLCatCodeLookup.RefreshSearchList
    frmBLCatCodeLookup.Show
    DoEvents
    KillFile ("categoryedit.dat")
    Unload frmBLCatEdit
  End If
  
  Exit Sub
  
ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLCatEdit", "cmdSave_Click", Erl)
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

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsBLTextBoxOverrider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  FirstTimeThru = True
  Call LoadMe
  FirstTimeThru = False
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    ''Me.Visible = False
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
      SendKeys "%S"
      Call cmdSave_Click
      KeyCode = 0
    Case vbKeyF7:
      SendKeys "%L"
      Call cmdCodeList_Click
      KeyCode = 0
    Case vbKeyF12:
      SendKeys "%G"
      Call cmdGLList_Click
      KeyCode = 0
    Case vbKeyF1:
      SendKeys "%H"
      Call cmdHelp_Click
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
      KillFile "categoryedit.dat"
      ClearInUse PWcnt
      MainLog ("BusinessLicense.exe terminated via menu bar on frmBLCatEdit.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub fpcmbCatType_Change()
  Dim x As Integer
  
  If ChangeFlag = True Then
    ChangeFlag = False
    Exit Sub
  End If
  
'  If GCatNum > 0 Then
'    If FirstTimeThru = False Then
'      If ChangeFlag = False Then
'        ChangeFlag = True
'        If Exist("artmppst.dat") Then
'          frmBLMessageBoxJr.Label1.Caption = "There is a pending business license renewal file that has not been posted. Category types cannot be edited until this business license file is posted."
'          frmBLMessageBoxJr.Label1.Top = 600
'          frmBLMessageBoxJr.Show vbModal
'          If TempType$ = "S" Then
'            fpcmbCatType.Text = "Step Rate"
'          ElseIf TempType$ = "M" Then
'            fpcmbCatType.Text = "Multiplier"
'          ElseIf TempType$ = "F" Then
'            fpcmbCatType.Text = "Flat Rate"
'          End If
'          Close
'          Exit Sub
'        End If
'      End If
'    End If
'  End If
  
  If QPTrim$(fpcmbCatType.Text) = "" Then
    fpcmbCatType = "Flat Rate"
    LabelFee.Caption = "FLAT RATE AMOUNT"
    fpcurrFee.Enabled = True
    For x = 0 To 5
      fpcurrBase(x).Enabled = False
      fpcurrRecUpTo(x).Enabled = False
      fptxtPct(x).Enabled = False
      fpcurrOver(x).Enabled = False
    Next x
    Exit Sub
  End If
  
  If QPTrim$(fpcmbCatType.Text) = "Flat Rate" Then
    LabelFee.Caption = "FLAT RATE AMOUNT"
    fpcurrFee.Enabled = True
    For x = 0 To 5
      fpcurrBase(x).Enabled = False
      fpcurrRecUpTo(x).Enabled = False
      fptxtPct(x).Enabled = False
      fpcurrOver(x).Enabled = False
    Next x
  ElseIf QPTrim$(fpcmbCatType.Text) = "Multiplier" Then
    LabelFee.Caption = "RATE PER UNIT"
    fpcurrFee.Enabled = True
    For x = 0 To 5
      fpcurrBase(x).Enabled = False
      fpcurrRecUpTo(x).Enabled = False
      fptxtPct(x).Enabled = False
      fpcurrOver(x).Enabled = False
    Next x
  ElseIf QPTrim$(fpcmbCatType.Text) = "Step Rate" Then
    LabelFee.Caption = "NOT APPLICABLE"
    fpcurrFee.Enabled = False
    For x = 0 To 5
      fpcurrBase(x).Enabled = True
      fpcurrRecUpTo(x).Enabled = True
      fptxtPct(x).Enabled = True
      fpcurrOver(x).Enabled = True
    Next x
  End If
  
End Sub

Private Sub fpcmbCatType_KeyDown(KeyCode As Integer, Shift As Integer)
  'this keeps the user from inadvertently changing data on this
  'combo box if they are scrolling through the form
  If KeyCode = vbKeySpace Then
    fpcmbCatType.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbCatType.ListIndex = -1
  End If
  If fpcmbCatType.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      SendKeys "{Tab}"
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Public Sub LoadMe()
  Dim x As Integer
  Dim CatFile As Integer
  Dim NumOfCatRecs As Integer
  Dim RecNo As Integer
  Dim GLFundLen As Integer
  Dim GLAcctLen As Integer
  Dim GLDetLen As Integer
  Dim ValidCode As Boolean
  Dim cnt As Integer
  Dim CatCodeRecLen As Integer
  Dim CodeIdxHandle As Integer
  Dim CodeIdx As CatCodeIdxType
  Dim CodeIdxNum As Integer
  Dim One As Integer
  Dim DHandle As Integer
  Dim TownRec As TownSetUpType
  Dim TownHandle As Integer
  
  On Error GoTo ERRORSTUFF
  
  TempRevNum$ = ""
  TempAcctsRecNum$ = ""
  TempCashNum$ = ""
  For x = 1 To 6
    TempBaseRate(x) = 0
    TempUpToAmt(x) = 0
    TempPctAmt(x) = 0
    TempAmtsOver(x) = 0
  Next x
  TempCode$ = ""
  TempType$ = ""
  TempDesc$ = ""
  TempRate = 0
  FirstSave = False
  CatInUseFlag = False
  ChangeFlag = False
  AccMethodIsNone = False
  OpenTownFile TownHandle
  Get TownHandle, 1, TownRec
  Close TownHandle
  
  If QPTrim$(TownRec.AcctMeth) = "N" Then
    fptxtRevGLAcctNum.ControlType = ControlTypeReadOnly
    fptxtAcctsRec.ControlType = ControlTypeReadOnly
    fptxtCashReceipt.ControlType = ControlTypeReadOnly
    AccMethodIsNone = True
    Label8.Caption = "Accounting Method: 'None'"
  ElseIf QPTrim$(TownRec.AcctMeth) = "A" Then
    fptxtRevGLAcctNum.ControlType = ControlTypeNormal
    fptxtAcctsRec.ControlType = ControlTypeNormal
    fptxtCashReceipt.ControlType = ControlTypeNormal
    Label8.Caption = "Accounting Method: 'Accrual'"
  ElseIf QPTrim$(TownRec.AcctMeth) = "C" Then
    fptxtRevGLAcctNum.ControlType = ControlTypeNormal
    fptxtAcctsRec.ControlType = ControlTypeReadOnly
    fptxtCashReceipt.ControlType = ControlTypeNormal
    Label8.Caption = "Accounting Method: 'Cash'"
  End If
  
  One = 1
  DHandle = FreeFile
  Open "categoryedit.dat" For Output As DHandle Len = 2
  Print #DHandle, One
  Close DHandle
  AddFlag = False
  
  fpcmbCatType.Text = "Flat Rate"
  If fpcmbCatType.ListCount = 0 Then
    fpcmbCatType.AddItem "Flat Rate"
    fpcmbCatType.AddItem "Multiplier"
    fpcmbCatType.AddItem "Step Rate"
  End If
  
  GoSub LoadGLAcctInfo
  
  ReDim ARCatCodeRec(1) As ARNewCatCodeRecType
  OpenCatCodeFile CatFile
  NumOfCatRecs = LOF(CatFile) \ Len(ARCatCodeRec(1))
  
  If GCatNum > 0 Then
    Get CatFile, GCatNum, ARCatCodeRec(1)
    MainLog ("Category code # " + QPTrim$(ARCatCodeRec(1).CatCode) + "/" + QPTrim$(ARCatCodeRec(1).CODEDESC) + " edit screen opened.")
    If CustUsingCat(QPTrim$(ARCatCodeRec(1).CatCode)) = True Then
      fptxtCatCode.Enabled = False
      CatInUseFlag = True 'tells the program to alert the
      'user if he tries to save a new type for this category
      'that he will need to edit customers currently saved
      'with the old type or else license fees will be inaccurate
    Else
      fptxtCatCode.Enabled = True
    End If
    fptxtCatCode.Text = QPTrim$(ARCatCodeRec(1).CatCode)
    TempCode$ = QPTrim$(ARCatCodeRec(1).CatCode)
    CatCodeNum = QPTrim$(ARCatCodeRec(1).CatCode)
    fptxtCatDesc.Text = QPTrim$(ARCatCodeRec(1).CODEDESC)
    TempDesc$ = QPTrim$(ARCatCodeRec(1).CODEDESC)
    fpcurrFee = ARCatCodeRec(1).Fee
    TempRate = ARCatCodeRec(1).Fee
    
    If ARCatCodeRec(1).CodeType = "M" Then
      fpcmbCatType.Text = "Multiplier"
    ElseIf ARCatCodeRec(1).CodeType = "S" Then
      fpcmbCatType.Text = "Step Rate"
    ElseIf ARCatCodeRec(1).CodeType = "F" Then
      fpcmbCatType.Text = "Flat Rate"
    Else
      frmBLMessageBoxJr.Label1.Caption = "PROBLEM: This category does not have a category type saved. Please make sure to select one of the category types and re-save this category."
      frmBLMessageBoxJr.Label1.Top = 700
      frmBLMessageBoxJr.Show vbModal
    End If
    TempType$ = ARCatCodeRec(1).CodeType
    
    fpcurrBase(0) = ARCatCodeRec(1).BaseAmt1
    TempBaseRate(1) = ARCatCodeRec(1).BaseAmt1
    fpcurrRecUpTo(0) = ARCatCodeRec(1).Recpt1
    TempUpToAmt(1) = ARCatCodeRec(1).Recpt1
    fptxtPct(0).Text = CStr(ARCatCodeRec(1).Percent1)
    TempPctAmt(1) = ARCatCodeRec(1).Percent1
    TempPctAmt(1) = OldRound(TempPctAmt(1))
    fpcurrOver(0) = ARCatCodeRec(1).Maximum1
    TempAmtsOver(1) = ARCatCodeRec(1).Maximum1
    
    fpcurrBase(1) = ARCatCodeRec(1).BaseAmt2
    TempBaseRate(2) = ARCatCodeRec(1).BaseAmt2
    fpcurrRecUpTo(1) = ARCatCodeRec(1).Recpt2
    TempUpToAmt(2) = ARCatCodeRec(1).Recpt2
    fptxtPct(1).Text = CStr(ARCatCodeRec(1).Percent2)
    TempPctAmt(2) = ARCatCodeRec(1).Percent2
    TempPctAmt(2) = OldRound(TempPctAmt(2))
    fpcurrOver(1) = ARCatCodeRec(1).Maximum2
    TempAmtsOver(2) = ARCatCodeRec(1).Maximum2
    
    fpcurrBase(2) = ARCatCodeRec(1).BaseAmt3
    TempBaseRate(3) = ARCatCodeRec(1).BaseAmt3
    fpcurrRecUpTo(2) = ARCatCodeRec(1).Recpt3
    TempUpToAmt(3) = ARCatCodeRec(1).Recpt3
    fptxtPct(2).Text = CStr(ARCatCodeRec(1).Percent3)
    TempPctAmt(3) = ARCatCodeRec(1).Percent3
    TempPctAmt(3) = OldRound(TempPctAmt(3))
    fpcurrOver(2) = ARCatCodeRec(1).Maximum3
    TempAmtsOver(3) = ARCatCodeRec(1).Maximum3
    
    fpcurrBase(3) = ARCatCodeRec(1).BaseAmt4
    TempBaseRate(4) = ARCatCodeRec(1).BaseAmt4
    fpcurrRecUpTo(3) = ARCatCodeRec(1).Recpt4
    TempUpToAmt(4) = ARCatCodeRec(1).Recpt4
    fptxtPct(3).Text = CStr(ARCatCodeRec(1).Percent4)
    TempPctAmt(4) = ARCatCodeRec(1).Percent4
    TempPctAmt(4) = OldRound(TempPctAmt(4))
    fpcurrOver(3) = ARCatCodeRec(1).Maximum4
    TempAmtsOver(4) = ARCatCodeRec(1).Maximum4
    
    fpcurrBase(4) = ARCatCodeRec(1).BaseAmt5
    TempBaseRate(5) = ARCatCodeRec(1).BaseAmt5
    fpcurrRecUpTo(4) = ARCatCodeRec(1).Recpt5
    TempUpToAmt(5) = ARCatCodeRec(1).Recpt5
    fptxtPct(4).Text = CStr(ARCatCodeRec(1).Percent5)
    TempPctAmt(5) = ARCatCodeRec(1).Percent5
    TempPctAmt(5) = OldRound(TempPctAmt(5))
    fpcurrOver(4) = ARCatCodeRec(1).Maximum5
    TempAmtsOver(5) = ARCatCodeRec(1).Maximum5
    
    fpcurrBase(5) = ARCatCodeRec(1).BaseAmt6
    TempBaseRate(6) = ARCatCodeRec(1).BaseAmt6
    fpcurrRecUpTo(5) = ARCatCodeRec(1).Recpt6
    TempUpToAmt(6) = ARCatCodeRec(1).Recpt6
    fptxtPct(5).Text = CStr(ARCatCodeRec(1).Percent6)
    TempPctAmt(6) = ARCatCodeRec(1).Percent6
    TempPctAmt(6) = OldRound(TempPctAmt(6))
    fpcurrOver(5) = ARCatCodeRec(1).Maximum6
    TempAmtsOver(6) = ARCatCodeRec(1).Maximum6
    
    fptxtRevGLAcctNum.Text = GetGLNum(ARCatCodeRec(1).REVGLNUM)
    TempRevNum$ = GetGLNum(ARCatCodeRec(1).REVGLNUM)
    fptxtAcctsRec.Text = GetGLNum(ARCatCodeRec(1).ARGLACCT)
    TempAcctsRecNum$ = GetGLNum(ARCatCodeRec(1).ARGLACCT)
    fptxtCashReceipt.Text = GetGLNum(ARCatCodeRec(1).CASHACCT)
    TempCashNum$ = GetGLNum(ARCatCodeRec(1).CASHACCT)
  Else 'zero out
    FirstSave = True
    AddFlag = True 'tells program that this is a new addition
    'and the user might want to keep adding categories without
    'returning to the main menu everytime a save is made
    fptxtCatCode.Text = ""
    fptxtCatDesc.Text = ""
    fpcurrFee = 0
    fpcurrBase(0) = 0
    fpcurrRecUpTo(0) = 0
    fptxtPct(0).Text = "0"
    fpcurrOver(0) = 0
    fpcurrBase(1) = 0
    fpcurrRecUpTo(1) = 0
    fptxtPct(1).Text = "0"
    fpcurrOver(1) = 0
    fpcurrBase(2) = 0
    fpcurrRecUpTo(2) = 0
    fptxtPct(2).Text = "0"
    fpcurrOver(2) = 0
    fpcurrBase(3) = 0
    fpcurrRecUpTo(3) = 0
    fptxtPct(3).Text = "0"
    fpcurrOver(3) = 0
    fpcurrBase(4) = 0
    fpcurrRecUpTo(4) = 0
    fptxtPct(4).Text = "0"
    fpcurrOver(4) = 0
    fpcurrBase(5) = 0
    fpcurrRecUpTo(5) = 0
    fptxtPct(5).Text = "0"
    fpcurrOver(5) = 0
    fptxtRevGLAcctNum.Text = ""
    fptxtAcctsRec.Text = ""
    fptxtCashReceipt.Text = ""
    If TownRec.GL2Cats = "Y" Then
      fptxtRevGLAcctNum.Text = GetGLNum(TownRec.PENREVGLNUM)
      fptxtAcctsRec.Text = GetGLNum(TownRec.PENRECGLNUM)
      fptxtCashReceipt.Text = GetGLNum(TownRec.PENCASHACCT)
    End If
  End If
  
  Close CatFile

  Exit Sub

LoadGLAcctInfo:
  GetAcctStruct GLFundLen, GLAcctLen, GLDetLen
Return
  
ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLCatEdit", "LoadMe", Erl)
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

Private Sub Form_KeyPress(KeyAscii As Integer)
'  this causes all characters to be capitalized
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub

Private Function Check4ValidCat(CodeText$, CODEDESC$, CatNum As Integer) As Boolean
  Dim CodeRec As ARNewCatCodeRecType
  Dim CodeHandle As Integer
  Dim NumOfCodeRecs As Integer
  Dim x As Integer
  
  On Error GoTo ERRORSTUFF
  
  CODEDESC = QPTrim$(CODEDESC)
  Check4ValidCat = True
  CodeText = QPTrim$(CodeText)
  OpenCatCodeFile CodeHandle
  NumOfCodeRecs = LOF(CodeHandle) / Len(CodeRec)
  If NumOfCodeRecs = 0 Then Exit Function
  
  For x = 1 To NumOfCodeRecs
    If x = CatNum Then GoTo SameRec 'if this is a new
    'addition then CatNum will be zero so if a match occurs
    'then it will be invalid...if it's an edit then x will
    'skip over the catnum so if a mach is made it will mean
    'that the user has entered a code already in use
    Get CodeHandle, x, CodeRec
    If CodeText = QPTrim$(CodeRec.CatCode) Then
      Check4ValidCat = False
      Exit For
    End If
SameRec:
  Next x
  
  If Check4ValidCat = False Then
    If fptxtCatCode.Enabled = True Then
      fptxtCatCode.BackColor = 8454143
    End If
    frmBLMessageBoxJr.Label1.Caption = "The category code number entered is already in use. Please enter a different number."
    frmBLMessageBoxJr.Label1.Top = 800
    frmBLMessageBoxJr.Show vbModal
    If fptxtCatCode.Enabled = True Then
      fptxtCatCode.SetFocus
      fptxtCatCode.BackColor = 16777215
    End If
    Close CodeHandle
    Exit Function
  End If
  
  For x = 1 To NumOfCodeRecs
    Get CodeHandle, x, CodeRec
    If x = CatNum Then GoTo SameDesc
    If CODEDESC = QPTrim$(CodeRec.CODEDESC) Then
      Check4ValidCat = False
      Exit For
    End If
SameDesc:
  Next x

  If Check4ValidCat = False Then
    fptxtCatDesc.BackColor = 8454143
    frmBLMessageBoxJr.Label1.Caption = "The category description entered is already in use. Please enter a different description."
    frmBLMessageBoxJr.Label1.Top = 800
    frmBLMessageBoxJr.Show vbModal
    fptxtCatDesc.SetFocus
    fptxtCatDesc.BackColor = 16777215
  End If
    
  Close CodeHandle
  Exit Function
  
ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLCatEdit", "Check4ValidCat", Erl)
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

Private Sub fpcurrRatePer_KeyDown(KeyCode As Integer, Shift As Integer)
'  fpcurrRatePer.BackColor = 16777215
End Sub

Private Sub fptxtAcctsRec_Change()
  If fptxtAcctsRec.Text <> "" Then
    TempAcctsRecNum = QPTrim$(fptxtAcctsRec.Text) 'save the last number entered
  End If
End Sub

Private Sub fptxtCashReceipt_Change()
  If fptxtCashReceipt.Text <> "" Then
    TempCashNum$ = QPTrim$(fptxtCashReceipt.Text)
  End If
End Sub

Private Sub fptxtCatCode_KeyDown(KeyCode As Integer, Shift As Integer)
  fptxtCatCode.BackColor = 16777215
End Sub

Private Sub fptxtCatDesc_KeyDown(KeyCode As Integer, Shift As Integer)
  fptxtCatDesc.BackColor = 16777215
End Sub

Private Sub fptxtPct_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  fptxtPct(Index).BackColor = 16777215
End Sub

Private Function GLNumsOK() As Boolean
  Dim TownRec As TownSetUpType
  Dim TownHandle As Integer
  'looks to make sure the appropriate GL numbers are
  'populated based on the type of Accounting Method
  'selected in the Town Setup
  
  On Error GoTo ERRORSTUFF
  GLNumsOK = True
  
  OpenTownFile TownHandle
  Get TownHandle, 1, TownRec
  Close TownHandle
  Select Case QPTrim$(TownRec.AcctMeth)
    Case "A"
      If fptxtRevGLAcctNum.Text = "" Then
        GLNumsOK = False
        fptxtRevGLAcctNum.BackColor = &H80FFFF
        frmBLMessageBoxJr.Label1.Caption = "The 'Accrual' accounting method is the current method being used (set in Town Setup). This method requires the 'Revenue G/L Account Number' field to be filled in."
        frmBLMessageBoxJr.Label1.Top = 600
        frmBLMessageBoxJr.Show vbModal
        fptxtRevGLAcctNum.BackColor = &HFFFFFF
        fptxtRevGLAcctNum.SetFocus
        fptxtRevGLAcctNum.Text = TempRevNum$
      ElseIf fptxtAcctsRec.Text = "" Then
        GLNumsOK = False
        fptxtAcctsRec.BackColor = &H80FFFF
        frmBLMessageBoxJr.Label1.Caption = "The 'Accrual' accounting method is the current method being used (set in Town Setup). This method requires the 'Accounts Receivable Number' field to be filled in."
        frmBLMessageBoxJr.Label1.Top = 600
        frmBLMessageBoxJr.Show vbModal
        fptxtAcctsRec.BackColor = &HFFFFFF
        fptxtAcctsRec.SetFocus
        fptxtAcctsRec.Text = TempAcctsRecNum$
      ElseIf fptxtCashReceipt.Text = "" Then
        GLNumsOK = False
        fptxtCashReceipt.BackColor = &H80FFFF
        frmBLMessageBoxJr.Label1.Caption = "The 'Accrual' accounting method is the current method being used (set in Town Setup). This method requires the 'Cash Receipt G/L Account Number' field to be filled in."
        frmBLMessageBoxJr.Label1.Top = 600
        frmBLMessageBoxJr.Show vbModal
        fptxtCashReceipt.BackColor = &HFFFFFF
        fptxtCashReceipt.SetFocus
        fptxtCashReceipt.Text = TempCashNum$
      End If
    Case "C"
      If fptxtRevGLAcctNum.Text = "" Then
        GLNumsOK = False
        fptxtRevGLAcctNum.BackColor = &H80FFFF
        frmBLMessageBoxJr.Label1.Caption = "The 'Cash' accounting method is the current method being used. This method requires the 'Revenue G/L Account Number' field to be filled in."
        frmBLMessageBoxJr.Label1.Top = 600
        frmBLMessageBoxJr.Show vbModal
        fptxtRevGLAcctNum.BackColor = &HFFFFFF
        fptxtRevGLAcctNum.SetFocus
        fptxtRevGLAcctNum.Text = TempRevNum$
      ElseIf fptxtCashReceipt.Text = "" Then
        GLNumsOK = False
        fptxtCashReceipt.BackColor = &H80FFFF
        frmBLMessageBoxJr.Label1.Caption = "The 'Cash' accounting method is the current method being used. This method requires the 'Cash Receipt G/L Account Number' field to be filled in."
        frmBLMessageBoxJr.Label1.Top = 600
        frmBLMessageBoxJr.Show vbModal
        fptxtCashReceipt.BackColor = &HFFFFFF
        fptxtCashReceipt.SetFocus
        fptxtCashReceipt.Text = TempCashNum$
      ElseIf fptxtAcctsRec.Text <> "" Then
        GLNumsOK = False
        fptxtAcctsRec.BackColor = &H80FFFF
        frmBLMessageBoxJrWOpts.Label1.Caption = "The 'Cash' accounting method is the current method being used. This method requires the 'Accounts Receivable Number' field to be empty. Continuing will erase this number. OK to continue?"
        frmBLMessageBoxJrWOpts.Label1.Top = 700
        frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Continue"
        frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC Cancel"
        frmBLMessageBoxJrWOpts.Show vbModal
        If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "abort" Then
          Unload frmBLMessageBoxJrWOpts
          fptxtAcctsRec.BackColor = &HFFFFFF
          fptxtAcctsRec.SetFocus
        Else
          Unload frmBLMessageBoxJrWOpts
          fptxtAcctsRec.BackColor = &HFFFFFF
          fptxtAcctsRec.Text = ""
        End If
      End If
    Case "N"
      If fptxtRevGLAcctNum <> "" Then
        GLNumsOK = False
        fptxtRevGLAcctNum.BackColor = &H80FFFF
        frmBLMessageBoxJrWOpts.Label1.Caption = "No accounting method is being used. This requires the 'Revenue G/L Account Number' field to be empty. Continuing will erase this number. OK to continue?"
        frmBLMessageBoxJrWOpts.Label1.Top = 700
        frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Continue"
        frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC Cancel"
        frmBLMessageBoxJrWOpts.Show vbModal
        If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "abort" Then
          Unload frmBLMessageBoxJrWOpts
          fptxtRevGLAcctNum.BackColor = &HFFFFFF
          fptxtRevGLAcctNum.SetFocus
        Else
          Unload frmBLMessageBoxJrWOpts
          fptxtRevGLAcctNum.BackColor = &HFFFFFF
          fptxtRevGLAcctNum.Text = ""
        End If
      ElseIf fptxtAcctsRec.Text <> "" Then
        GLNumsOK = False
        fptxtAcctsRec.BackColor = &H80FFFF
        frmBLMessageBoxJrWOpts.Label1.Caption = "No accounting method is being used. This requires the 'Accounts Receivable Number' field to be empty. Continuing will erase this number. OK to continue?"
        frmBLMessageBoxJrWOpts.Label1.Top = 700
        frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Continue"
        frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC Cancel"
        frmBLMessageBoxJrWOpts.Show vbModal
        If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "abort" Then
          Unload frmBLMessageBoxJrWOpts
          fptxtAcctsRec.BackColor = &HFFFFFF
          fptxtAcctsRec.SetFocus
        Else
          Unload frmBLMessageBoxJrWOpts
          fptxtAcctsRec.BackColor = &HFFFFFF
          fptxtAcctsRec.Text = ""
        End If
      ElseIf fptxtCashReceipt.Text <> "" Then
        GLNumsOK = False
        fptxtCashReceipt.BackColor = &H80FFFF
        frmBLMessageBoxJrWOpts.Label1.Caption = "No accounting method is being used. This requires the 'Cash Receipt G/L Account Number' field to be empty. Continuing will erase this number. OK to continue?"
        frmBLMessageBoxJrWOpts.Label1.Top = 700
        frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Continue"
        frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC Cancel"
        frmBLMessageBoxJrWOpts.Show vbModal
        If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "abort" Then
          Unload frmBLMessageBoxJrWOpts
          fptxtCashReceipt.BackColor = &HFFFFFF
          fptxtCashReceipt.SetFocus
        Else
          Unload frmBLMessageBoxJrWOpts
          fptxtCashReceipt.BackColor = &HFFFFFF
          fptxtCashReceipt.Text = ""
        End If
      End If
    End Select
    
    Exit Function
    
ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLCatEdit", "GLNumsOK", Erl)
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

Private Sub fptxtRevGLAcctNum_Change()
  If fptxtRevGLAcctNum.Text <> "" Then
    TempRevNum$ = QPTrim$(fptxtRevGLAcctNum.Text)
  End If
End Sub

Private Function GLNumsValid() As Boolean
  Dim GLIdxRec As JGLAcctIdxType
  Dim IdxHandle As Integer
  Dim NumOfIdxRecs As Integer
  Dim x As Integer
  Dim GLAcctRec As GLAcctRecType
  Dim AcctHandle As Integer
  Dim RevNum$, Rev As Integer
  Dim AcctsRecNum$, Acct As Integer
  Dim CashRecNum$, Cash As Integer
  Dim ThisGLNum$
  Dim TownRec As TownSetUpType
  Dim TownHandle As Integer
  
  On Error GoTo ERRORSTUFF
  'looks to make sure that the GL numbers entered match
  'a GL number available on the GL list
  OpenTownFile TownHandle
  Get TownHandle, 1, TownRec
  Close TownHandle
  
  GLNumsValid = True
  
  If QPTrim$(TownRec.AcctMeth) <> "N" Then
    If Not Exist("GLACCT.IDX") Or Not Exist("GLACCT.DAT") Then
        If Not Exist("GLACCT.IDX") And Not Exist("GLACCT.DAT") Then
          frmBLMessageBoxJrWOpts.Label1.Caption = "The files 'GLACCT.IDX' and 'GLACCT.DAT' could not be found in the Citipak directory. These files are needed to verify the validity of General Ledger numbers. Numbers must be verified to be saved properly. Press F10 to continue saving anyway or ESC to abort."
          frmBLMessageBoxJrWOpts.Label1.Top = 500
          frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Continue"
          frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC Abort"
          frmBLMessageBoxJrWOpts.Show vbModal
          If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "abort" Then
            Unload frmBLMessageBoxJrWOpts
            GLNumsValid = False
            Exit Function
          Else
            MainLog ("User warned that 'GLACCT.DAT' and 'GLACCT.IDX' were missing from the Citipak directory and the user elected to continue saving the Town Setup data even though GL numbers could not be verified.")
            Unload frmBLMessageBoxJrWOpts
            GoTo CheckCompleted
          End If
        End If
        If Not Exist("GLACCT.IDX") Then
          frmBLMessageBoxJrWOpts.Label1.Caption = "The file 'GLACCT.IDX' could not be found in the Citipak directory. This file is required to verify the validity of General Ledger numbers. Numbers must be verified to be saved properly. Press F10 to continue saving anyway or ESC to abort."
          frmBLMessageBoxJrWOpts.Label1.Top = 500
          frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Continue"
          frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC Abort"
          frmBLMessageBoxJrWOpts.Show vbModal
          If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "abort" Then
            Unload frmBLMessageBoxJrWOpts
            GLNumsValid = False
            Exit Function
          Else
            MainLog ("User warned that 'GLACCT.IDX' was missing from the Citipak directory and the user elected to continue saving the Town Setup data even though GL numbers could not be verified.")
            Unload frmBLMessageBoxJrWOpts
            GoTo CheckCompleted
          End If
        End If
        If Not Exist("GLACCT.DAT") Then
          frmBLMessageBoxJrWOpts.Label1.Caption = "The file 'GLACCT.DAT' could not be found in the Citipak directory. This file is required to verify the validity of General Ledger numbers. Numbers must be verified to be saved properly. Press F10 to continue saving anyway or ESC to abort."
          frmBLMessageBoxJrWOpts.Label1.Top = 500
          frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Continue"
          frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC Abort"
          frmBLMessageBoxJrWOpts.Show vbModal
          If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "abort" Then
            Unload frmBLMessageBoxJrWOpts
            GLNumsValid = False
            Exit Function
          Else
            MainLog ("User warned that 'GLACCT.DAT' was missing from the Citipak directory and the user elected to continue saving the Town Setup data even though GL numbers could not be verified.")
            Unload frmBLMessageBoxJrWOpts
            GoTo CheckCompleted
          End If
        End If
  '    End If
      Exit Function
    End If
  End If
  
CheckCompleted:
  OpenGLIdxFile IdxHandle
  NumOfIdxRecs = LOF(IdxHandle) / Len(GLIdxRec)
  If NumOfIdxRecs = 0 Then
'    frmBLMessageBoxJr.Label1.Caption = "There are no General Ledger numbers indexed."
'    frmBLMessageBoxJr.Label1.Top = 900
'    frmBLMessageBoxJr.Show vbModal
    Close
    Exit Function
  End If
  ReDim IdxRec(1 To NumOfIdxRecs) As Integer
  
  For x = 1 To NumOfIdxRecs
    Get IdxHandle, x, GLIdxRec 'build GL number index
    IdxRec(x) = GLIdxRec.RecNo
  Next x
  Close IdxHandle
  
  'If the GL fields are populated then the accompanying
  'variable (ie RevNum) are assigned 0, otherwise they get 1
  RevNum$ = QPTrim$(fptxtRevGLAcctNum.Text)
  If RevNum <> "" Then
    Rev = 0
  Else
    Rev = 1
  End If
  
  AcctsRecNum$ = QPTrim$(fptxtAcctsRec.Text)
  If AcctsRecNum$ <> "" Then
    Acct = 0
  Else
    Acct = 1
  End If
  
  CashRecNum$ = QPTrim$(fptxtCashReceipt.Text)
  If CashRecNum$ <> "" Then
    Cash = 0
  Else
    Cash = 1
  End If
  
  'Now look at those that are populated and compare
  'them with existing GL numbers...when a match comes up
  'then make that GL variable equal 1 and don't check it anymore
  OpenGLAcctFile AcctHandle
  For x = 1 To NumOfIdxRecs
    Get AcctHandle, IdxRec(x), GLAcctRec
      If GLAcctRec.Deleted Then GoTo Notvalid
      ThisGLNum = QPTrim$(GLAcctRec.Num)
      If Rev = 1 Then GoTo RevIs1
      If ThisGLNum$ = RevNum$ Then
        Rev = 1
      End If
RevIs1:
      If Acct = 1 Then GoTo AcctIs1
      If ThisGLNum$ = AcctsRecNum$ Then
        Acct = 1
      End If
AcctIs1:
      If Cash = 1 Then GoTo CashIs1
      If ThisGLNum$ = CashRecNum$ Then
        Cash = 1
      End If
CashIs1:
      'each is either empty or has a match
      If Rev = 1 And Acct = 1 And Cash = 1 Then
        Exit For
      End If
Notvalid:
  Next x
  
  'at this point there is a 'no match' issue
  If x > NumOfIdxRecs Then
    If Rev = 0 Then 'we know if it's equal to 0 then it's populated
    'with a number that doesn't match
      fptxtRevGLAcctNum.BackColor = &H80FFFF
      frmBLMessageBoxJr.Label1.Caption = "The GL Number entered for 'Revenue G/L Account Number' does not match any GL Numbers on file."
      frmBLMessageBoxJr.Label1.Top = 800
      frmBLMessageBoxJr.Show vbModal
      fptxtRevGLAcctNum.BackColor = &HFFFFFF
      fptxtRevGLAcctNum.SetFocus
      GLNumsValid = False
      Exit Function
    ElseIf Acct = 0 Then
      fptxtAcctsRec.BackColor = &H80FFFF
      frmBLMessageBoxJr.Label1.Caption = "The GL Number entered for 'Accounts Receivable Number' does not match any GL Numbers on file."
      frmBLMessageBoxJr.Label1.Top = 800
      frmBLMessageBoxJr.Show vbModal
      fptxtAcctsRec.BackColor = &HFFFFFF
      fptxtAcctsRec.SetFocus
      GLNumsValid = False
      Exit Function
    ElseIf Cash = 0 Then
      fptxtCashReceipt.BackColor = &H80FFFF
      frmBLMessageBoxJr.Label1.Caption = "The GL Number entered for 'Cash Receipt G/L Account Number' does not match any GL Numbers on file."
      frmBLMessageBoxJr.Label1.Top = 800
      frmBLMessageBoxJr.Show vbModal
      fptxtCashReceipt.BackColor = &HFFFFFF
      fptxtCashReceipt.SetFocus
      GLNumsValid = False
      Exit Function
    End If
  End If
  
  Exit Function
  
  
ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLCatEdit", "GLNumsValid", Erl)
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
  
Private Function CustUsingCat(CatCode$) As Boolean
  Dim CustRec As ARCustRecType
  Dim CustHandle As Integer
  Dim x As Integer
  Dim NumOfCusts As Integer
  
  CustUsingCat = False
  OpenCustFile CustHandle
  NumOfCusts = LOF(CustHandle) / Len(CustRec)
  If NumOfCusts = 0 Then Exit Function
  
  For x = 1 To NumOfCusts
    Get CustHandle, x, CustRec
      If QPTrim$(CustRec.Deleted) = "Y" Or QPTrim$(CustRec.SortName) = "DELETED" Then GoTo NotThisOne
      Select Case CatCode$
        Case QPTrim$(CustRec.BILLCAT1), QPTrim$(CustRec.BILLCAT2), QPTrim$(CustRec.BILLCAT3), _
             QPTrim$(CustRec.BILLCAT4), QPTrim$(CustRec.BILLCAT5)
             CustUsingCat = True
             Close CustHandle
             Exit Function
        Case Else
      End Select
NotThisOne:
  Next x
  Close
End Function

Private Sub LogSave()
  Dim x As Integer
  
  If FirstSave = True Then
    MainLog ("Initial save for category " + QPTrim$(fptxtCatCode.Text) + ".")
    MainLog ("Initial description for category " + QPTrim$(fptxtCatCode.Text) + " is " + QPTrim$(fptxtCatDesc.Text) + ".")
    If TempType$ = "F" Then
      MainLog ("Initial 'type' save for " + QPTrim$(fptxtCatCode.Text) + " is Flat Rate.")
      MainLog ("Initial 'rate amount' save for category number " + QPTrim$(fptxtCatCode.Text) + " is " + fpcurrFee.Text + ".")
    ElseIf TempType$ = "M" Then
      MainLog ("Initial 'type' save for category " + QPTrim$(fptxtCatCode.Text) + " is Multiplier.")
      MainLog ("Initial 'rate amount' save for category number " + QPTrim$(fptxtCatCode.Text) + " is " + fpcurrFee.Text + ".")
    ElseIf TempType$ = "S" Then
      MainLog ("Initial 'type' save for category " + QPTrim$(fptxtCatCode.Text) + " is Step Rate.")
      For x = 0 To 5
        MainLog ("Initial 'base fee' save for category number " + QPTrim$(fptxtCatCode.Text) + " is " + fpcurrBase(x) + ".")
        MainLog ("Initial 'receipts up to' amount save for category number " + QPTrim$(fptxtCatCode.Text) + " is " + fpcurrRecUpTo(x) + ".")
        MainLog ("Initial 'plus %' amount for category number " + QPTrim$(fptxtCatCode.Text) + " is " + Using("##0.000", fptxtPct(x).Text) + ".")
        MainLog ("Initial 'on amounts over' amount for category number " + QPTrim$(fptxtCatCode.Text) + " is " + fpcurrOver(x) + ".")
      Next x
    End If
    MainLog ("Initial GL revenue number save for category " + QPTrim$(fptxtCatCode.Text) + " is " + fptxtRevGLAcctNum.Text + ".")
    MainLog ("Initial GL accounts receivable number save for category " + QPTrim$(fptxtCatCode.Text) + " is " + fptxtAcctsRec.Text + ".")
    MainLog ("Initial GL cash number save for category " + QPTrim$(fptxtCatCode.Text) + " is " + fptxtCashReceipt.Text + ".")
  Else
    If TempType = "S" Then TempType = "Step Rate"
    If TempType = "M" Then TempType = "Multiplier"
    If TempType = "F" Then TempType = "Flat Rate"
    MainLog ("Category code " + TempCode + " data saved " + ".")
    If TempDesc <> QPTrim$(fptxtCatDesc.Text) Then
      MainLog ("Description for " + TempDesc + " changed to " + QPTrim$(fptxtCatDesc.Text) + " for category " + QPTrim$(fptxtCatCode.Text) + ".")
    End If
    If TempType <> QPTrim$(fpcmbCatType.Text) Then
      MainLog ("The 'type' for category " + QPTrim$(fptxtCatCode.Text) + " changed from " + TempType + " to " + QPTrim$(fpcmbCatType.Text) + ".")
    End If
    If TempRate <> fpcurrFee.DoubleValue Then
      MainLog ("The 'rate amount' for category " + QPTrim$(fptxtCatCode.Text) + " changed from " + Using("$#,##0.00", TempRate) + " to " + Using("$#,##0.00", fpcurrFee.DoubleValue) + ".")
    End If
    
    For x = 1 To 5
      If TempBaseRate(x) <> fpcurrBase(x - 1).DoubleValue Then
        MainLog ("The 'base rate' amount for category " + QPTrim$(fptxtCatCode.Text) + " changed from " + Using$("$#,##0.00", TempBaseRate(x)) + " to " + Using("$#,##0.00", fpcurrBase(x - 1).DoubleValue) + ".")
      End If
      If TempUpToAmt(x) <> fpcurrRecUpTo(x - 1).DoubleValue Then
        MainLog ("The 'for receipts up to' amount for category " + QPTrim$(fptxtCatCode.Text) + " changed from " + Using$("$###,#,##0.00", TempUpToAmt(x)) + " to " + Using("$#,##0.00", fpcurrRecUpTo(x - 1).DoubleValue) + ".")
      End If
      If OldRound(TempPctAmt(x)) <> QPTrim$(fptxtPct(x - 1).Text) Then
        MainLog ("The 'plus %' amount for category " + QPTrim$(fptxtCatCode.Text) + " changed from " + CStr(OldRound(TempPctAmt(x))) + " to " + QPTrim$(fptxtPct(x - 1).Text) + ".")
      End If
    Next x
   
    If TempRevNum$ <> QPTrim$(fptxtRevGLAcctNum.Text) Then
      MainLog ("The GL revenue number for category " + QPTrim$(fptxtCatCode.Text) + " changed from " + TempRevNum$ + " to " + QPTrim$(fptxtRevGLAcctNum.Text) + ".")
    End If
   
    If TempAcctsRecNum$ <> QPTrim$(fptxtAcctsRec.Text) Then
      MainLog ("The GL revenue number for category " + QPTrim$(fptxtCatCode.Text) + " changed from " + TempAcctsRecNum$ + " to " + QPTrim$(fptxtAcctsRec.Text) + ".")
    End If
   
    If TempCashNum$ <> QPTrim$(fptxtCashReceipt.Text) Then
      MainLog ("The GL revenue number for category " + QPTrim$(fptxtCatCode.Text) + " changed from " + TempCashNum$ + " to " + QPTrim$(fptxtCashReceipt.Text) + ".")
    End If
  End If
End Sub
