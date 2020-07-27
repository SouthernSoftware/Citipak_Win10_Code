VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmFADsplSingle 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fixed Asset Single Item Disposal"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmFADsplSingle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpList fpListSearch 
      Height          =   2415
      Left            =   525
      TabIndex        =   6
      Top             =   5715
      Width           =   10635
      _Version        =   196608
      _ExtentX        =   18759
      _ExtentY        =   4260
      TextAlias       =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
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
      BorderStyle     =   1
      BorderColor     =   8454143
      BorderWidth     =   2
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ScrollHScale    =   2
      ScrollHInc      =   0
      ColsFrozen      =   0
      ScrollBarV      =   1
      NoIntegralHeight=   0   'False
      HighestPrecedence=   0
      AllowColResize  =   2
      AllowColDragDrop=   0
      ReadOnly        =   0   'False
      VScrollSpecial  =   0   'False
      VScrollSpecialType=   0
      EnableKeyEvents =   -1  'True
      EnableTopChangeEvent=   -1  'True
      DataAutoHeadings=   -1  'True
      DataAutoSizeCols=   3
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
      ColDesigner     =   "frmFADsplSingle.frx":08CA
   End
   Begin EditLib.fpCurrency fpcurrLow 
      Height          =   396
      Left            =   4272
      TabIndex        =   3
      ToolTipText     =   "Enter a starting purchase price for any fixed assets you wish to dispose of (optional)."
      Top             =   3312
      Width           =   1932
      _Version        =   196608
      _ExtentX        =   3408
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
   Begin EditLib.fpText fptxtTagNumber 
      Height          =   396
      Left            =   4272
      TabIndex        =   0
      ToolTipText     =   $"frmFADsplSingle.frx":0CC9
      Top             =   1824
      Width           =   4620
      _Version        =   196608
      _ExtentX        =   8149
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
   Begin EditLib.fpText fptxtDesc 
      Height          =   396
      Left            =   4272
      TabIndex        =   1
      ToolTipText     =   $"frmFADsplSingle.frx":0D73
      Top             =   2304
      Width           =   4620
      _Version        =   196608
      _ExtentX        =   8149
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
   Begin EditLib.fpText fpTxtSerialNum 
      Height          =   396
      Left            =   4272
      TabIndex        =   2
      ToolTipText     =   $"frmFADsplSingle.frx":0E23
      Top             =   2796
      Width           =   4620
      _Version        =   196608
      _ExtentX        =   8149
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
   Begin EditLib.fpCurrency fpcurrHigh 
      Height          =   396
      Left            =   6960
      TabIndex        =   4
      ToolTipText     =   "Enter the highest price for any fixed asset you wish to dispose of (optional)."
      Top             =   3312
      Width           =   1932
      _Version        =   196608
      _ExtentX        =   3408
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
   Begin EditLib.fpText fpDateAcquired 
      Height          =   396
      Left            =   4272
      TabIndex        =   5
      ToolTipText     =   "Enter the year of the fixed asset you wish to dispose of (optional)."
      Top             =   3840
      Width           =   1932
      _Version        =   196608
      _ExtentX        =   3408
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
      CharValidationText=   "0 , 1, 2, 3, 4, 5, 6, 7, 8, 9"
      MaxLength       =   4
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
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   690
      Left            =   3564
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   4524
      Width           =   1890
      _Version        =   131072
      _ExtentX        =   3334
      _ExtentY        =   1217
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      GrayAreaColor   =   13684944
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
      ButtonDesigner  =   "frmFADsplSingle.frx":0EC8
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSearch 
      Height          =   690
      Left            =   6204
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   $"frmFADsplSingle.frx":10A4
      Top             =   4524
      Width           =   1890
      _Version        =   131072
      _ExtentX        =   3334
      _ExtentY        =   1217
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      GrayAreaColor   =   13684944
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
      ButtonDesigner  =   "frmFADsplSingle.frx":114C
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "* = Already slated for disposal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   1056
      TabIndex        =   14
      Top             =   5232
      Width           =   2952
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "To:"
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
      Left            =   6336
      TabIndex        =   13
      Top             =   3408
      Width           =   456
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Fixed Assets Single Item Disposal Lookup"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2826
      TabIndex        =   12
      Top             =   633
      Width           =   6015
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   750
      Index           =   1
      Left            =   1386
      Top             =   498
      Width           =   8655
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   4140
      Left            =   900
      Top             =   1440
      Width           =   9816
   End
   Begin VB.Label lblDesc 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tag Number:"
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
      Left            =   2484
      TabIndex        =   11
      Top             =   1932
      Width           =   1548
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Description:"
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
      Left            =   1920
      TabIndex        =   10
      Top             =   2412
      Width           =   2136
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Serial Number:"
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
      Left            =   1920
      TabIndex        =   9
      Top             =   2892
      Width           =   2136
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Acquired Year:"
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
      Left            =   1872
      TabIndex        =   8
      Top             =   3888
      Width           =   2136
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Price From:"
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
      Left            =   1584
      TabIndex        =   7
      Top             =   3408
      Width           =   2472
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   840
      Left            =   1386
      Top             =   438
      Width           =   8655
   End
End
Attribute VB_Name = "frmFADsplSingle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsFATextBoxOverRider
  Private Temp_Class As Resize_Class

Private Sub cmdExit_Click()
  frmFADisposalMenu.Show
  Close
  DoEvents
  Unload frmFADsplSingle
End Sub

Private Sub cmdSearch_Click()
  Dim FAHandle As Integer
  Dim NumOfRecs As Integer
  Dim FAItemRec As FAItemRecType
  Dim x As Long
  Dim Found As Boolean
  Dim TagFlag As Boolean
  Dim DescFlag As Boolean
  Dim SerialFlag As Boolean
  Dim PriceFlag As Boolean
  Dim DateFlag As Boolean
  Dim TempTag$
  Dim TempDesc$
  Dim TempSerial$
  Dim TempPriceLow As Double
  Dim TempPriceHigh As Double
  Dim TempDateLow As Integer
  Dim TempDateHigh As Integer
  Dim FoundCnt As Integer
  Dim MatchCnt As Integer
  Dim PrintDesc$
  Dim OnlyOneFound$
  Dim TagIdx As TagNumbSortIdxType
  Dim TagIdxHandle As Integer
  Dim TempDsplFlag As Boolean
  Dim Message$
  
  On Error GoTo ERRORSTUFF
  Found = True
  TempDsplFlag = False
  fpListSearch.Clear
  TagFlag = False
  DescFlag = False
  SerialFlag = False
  PriceFlag = False
  DateFlag = False
  
  If QPTrim$(fptxtTagNumber.Text) <> "" Then 'user entered a
  'value to assist in the search
    TagFlag = True
    TempTag = QPTrim$(fptxtTagNumber.Text)
  End If
  
  If QPTrim$(fptxtDesc.Text) <> "" Then 'user entered a value to
  'assist in the search
    DescFlag = True
    TempDesc = QPTrim$(fptxtDesc.Text)
  End If
  
  If QPTrim$(fptxtSerialNum.Text) <> "" Then 'user entered a value
  'to assist in the search
    SerialFlag = True
    TempSerial = QPTrim$(fptxtSerialNum.Text)
  End If
  
  If fpcurrLow > fpcurrHigh Then
    MsgBox "Please enter an amount for the least Purchase Price that is less than the most Purchase Price."
    Close
    fpcurrLow.SetFocus
    Exit Sub
  End If
  
  TempPriceLow = fpcurrLow
  TempPriceHigh = fpcurrHigh
  If TempPriceHigh > 0 Then 'low price can be zero but if low
  'price is more than zero then high price must have a value greater
    PriceFlag = True
  Else
    PriceFlag = False
  End If
  
  If QPTrim$(fpDateAcquired) <> "" Then 'there is a value in this field
    If CInt(fpDateAcquired.Text) >= 1950 And CInt(fpDateAcquired.Text) <= 2100 Then
      DateFlag = True 'dates entered are acceptable
      TempDateLow = Date2Num("01/01/" + fpDateAcquired)
      TempDateHigh = Date2Num("12/31/" + fpDateAcquired)
    End If
  End If
  
  OpenTagIdxFile TagIdxHandle
  NumOfRecs = LOF(TagIdxHandle) \ Len(TagIdx)
  
  If NumOfRecs = 0 Then
    MsgBox "No records on file."
    Close TagIdxHandle
    Exit Sub
  End If
  
  ReDim TagIdxRecs(1 To NumOfRecs) As Integer
  For x = 1 To NumOfRecs
    Get TagIdxHandle, x, TagIdx
    TagIdxRecs(x) = TagIdx.DataRecNum 'load array with tag records
    'arranged in tag number numeric order
  Next x
  Close TagIdxHandle
  
  OpenFAItemFile FAHandle
  
  For x = 1 To NumOfRecs
    Get FAHandle, TagIdxRecs(x), FAItemRec
    If FAItemRec.DsplFlag = 2 Then GoTo NotAMatch 'don't need items that
    'are already disposed of
    
    If FAItemRec.DsplFlag = 1 Then TempDsplFlag = True 'this one is earmarked for
    'disposal but not yet disposed of
    Found = True
    'now start search
    If TagFlag = True Then 'user wants tag number used in search procedure
      If InStr(UCase$(FAItemRec.ItemTag), TempTag) > 0 Then
        Found = True 'tag number entered has matched
      Else
        Found = False 'move to next asset because lack of match here voids
        'this asset
        GoTo NotAMatch
      End If
    End If
    'all search related flags operate as does the TagFlag
    If DescFlag = True Then '
      If InStr(UCase$(FAItemRec.IDESC1), TempDesc) > 0 Or InStr(UCase$(FAItemRec.IDESC2), TempDesc) > 0 Then
        Found = True
      Else
        Found = False
        GoTo NotAMatch
      End If
    End If
    
    If SerialFlag = True Then
      If InStr(UCase$(FAItemRec.SERIALNO), TempSerial) > 0 Then
        Found = True
      Else
        Found = False
        GoTo NotAMatch
      End If
    End If
    
    If PriceFlag = True Then
      If FAItemRec.ORGCOST >= TempPriceLow And FAItemRec.ORGCOST <= TempPriceHigh Then
        Found = True
      Else
        Found = False
        GoTo NotAMatch
      End If
    End If
    
    If DateFlag = True Then
      If FAItemRec.AQURDATE >= TempDateLow And FAItemRec.AQURDATE <= TempDateHigh Then
        Found = True
      Else
        Found = False
        GoTo NotAMatch
      End If
    End If
    
    If Found Then 'an asset is found that fits the criteria...keep looking for more
    'that fits the criteria
      FoundCnt = FoundCnt + 1
      fpListSearch.Row = -1
      MatchCnt = MatchCnt + 1
      If QPTrim$(FAItemRec.IDESC1) <> "" Then 'get this asset's best description
        PrintDesc$ = QPTrim$(FAItemRec.IDESC1)
      Else
        PrintDesc$ = QPTrim$(FAItemRec.IDESC2)
      End If
      If TempDsplFlag = True Then 'if this asset is flagged for future disposal but
      'not yet disposed than indicate by adding an asterick to it's asset number
        fpListSearch.InsertRow = " * " & QPTrim$(FAItemRec.ItemTag) & Chr$(9) & " " & PrintDesc$ & " " & Chr$(9) & " " & QPTrim$(FAItemRec.SERIALNO) & Chr$(9) & "  " & Using$("$##,###,##0.00", FAItemRec.ORGCOST) & Chr$(9) & " " & MakeRegDate(FAItemRec.AQURDATE)
      Else
        fpListSearch.InsertRow = "   " & QPTrim$(FAItemRec.ItemTag) & Chr$(9) & " " & PrintDesc$ & " " & Chr$(9) & " " & QPTrim$(FAItemRec.SERIALNO) & Chr$(9) & "  " & Using$("$##,###,##0.00", FAItemRec.ORGCOST) & Chr$(9) & " " & MakeRegDate(FAItemRec.AQURDATE)
      End If
      'will continue to be assigned but is only used if no more than one found
      OnlyOneFound = QPTrim$(FAItemRec.ItemTag)
    End If
NotAMatch:
    TempDsplFlag = False 'reset this so the next asset won't be mistakenly flagged
  Next x
  
  If MatchCnt <= 0 Then 'unlikely scenario
    MsgBox "No match found"
    Exit Sub
    Close
  End If
    
  If FoundCnt = 1 Then 'OK we only found one valid asset
    For x = 1 To NumOfRecs
      Get FAHandle, x, FAItemRec
        If OnlyOneFound = QPTrim$(FAItemRec.ItemTag) Then
          If FAItemRec.DsplFlag = 1 Then
            Message = "This item has been selected for disposal on " + MakeRegDate(FAItemRec.DispDate) + ". Please access this item using the Edit Item Disposal List button on the Disposal Menu."
            MsgBox (Message)
            Close
            Exit Sub
          End If
          GRecNum = x 'assign this global value with this valid record number
          Exit For 'jump out of the for loop because we now have what we want
        Else
          Found = False
        End If
    Next x
    
    fptxtTagNumber.Text = ""
    fptxtDesc.Text = ""
    fptxtSerialNum.Text = ""
    fpListSearch.Clear
    FoundCnt = 0
    'not unloading this form because when the program moves to
    'frmFAPostSnglDspl the screen briefly shows the desktop plus
    'when frmFAPostSnglDspl closes it has no option but to return
    'to this form which when it exits unloads itself
    frmFAPostSnglDspl.Show
    DoEvents
  End If
  Close FAHandle
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFADsplSingle", "cmdSearch_Click", Erl)
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
    Unload Me
End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsFATextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
'    'Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    If fpListSearch.ListIndex <> -1 Then GoTo EmpAlreadySelected '8/6
    If Len(fptxtTagNumber.Text) > 0 Or Len(fptxtDesc.Text) > 0 Or Len(fptxtSerialNum.Text) > 0 Then
      Call cmdSearch_Click
      KeyCode = 0
      Exit Sub
    End If
EmpAlreadySelected:
    fpListSearch.Col = 1
    If QPTrim$(fpListSearch.ColText) = "" Then
      MsgBox "No item has been selected"
      Exit Sub
    Else
      Call fpListSearch_DblClick
      KeyCode = 0
      Exit Sub
    End If
  End If
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
      Call cmdSearch_Click
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
      MainLog ("FixedAssets.exe terminated via menu bar on frmFADsplSingle.")
      Call Terminate
      End
    End If
  End If
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
'  this causes all characters to be capitalized
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub


Private Sub fpListSearch_DblClick()
  Dim FAHandle As Integer
  Dim NumOfRecs As Integer
  Dim FAItemRec As FAItemRecType
  Dim x As Long
  Dim TagNum$
  Dim Desc$
  Dim SerialNum$
  Dim PriceNum As Double
  Dim DateNum As Integer
  Dim PrintDesc$
  Dim Found As Boolean
  Dim Message$
  Dim TempDsplFlag As Boolean
  
  On Error GoTo ERRORSTUFF
  TempDsplFlag = False
  fpListSearch.Col = 0
  'trap for double clicking on nothing
  If QPTrim$(fpListSearch.ColText) = "" Then
    MsgBox "No item has been selected"
    Exit Sub
  End If
  TagNum$ = QPTrim$(fpListSearch.ColText)
  
  If InStr(TagNum$, "*") > 0 Then 'looks for assets that
  'have already been earmarked for disposal...don't want
  'this asset processed here
    TagNum$ = Mid(QPTrim$(fpListSearch.ColText), 2) 'capture tag number
    TempDsplFlag = True
  End If
  
  fpListSearch.Col = 1 'capture tag description
  Desc$ = QPTrim$(fpListSearch.ColText)
  
  fpListSearch.Col = 2
  SerialNum$ = QPTrim$(fpListSearch.ColText) 'capture tag serial number
  
  fpListSearch.Col = 3
  PriceNum = CDbl(fpListSearch.ColText) 'capture the tag's price
  
  fpListSearch.Col = 4
  DateNum = Date2Num(fpListSearch.ColText) 'capture tag's acquisition date
  
  GRecNum = 0 'clear in case it holds a leftover value
  OpenFAItemFile FAHandle
  NumOfRecs = LOF(FAHandle) \ Len(FAItemRec)
  
  For x = 1 To NumOfRecs
    Get FAHandle, x, FAItemRec
    If QPTrim$(FAItemRec.IDESC1) <> "" Then
      PrintDesc$ = QPTrim$(FAItemRec.IDESC1) 'capture best tag description
    Else
      PrintDesc$ = QPTrim$(FAItemRec.IDESC2)
    End If
  
    If QPTrim$(FAItemRec.ItemTag) = QPTrim$(TagNum$) And InStr(UCase$(PrintDesc$), Desc$) > 0 And InStr(FAItemRec.SERIALNO, SerialNum$) >= 0 _
    And Len(QPTrim$(FAItemRec.ItemTag)) = Len(QPTrim$(TagNum$)) And FAItemRec.ORGCOST = PriceNum And FAItemRec.AQURDATE = DateNum Then '8/7 added Len = Len because
    'if two people had the same name and the emp number of one had a number that
    'included the other's (ie. 123 vs 1234) then then smaller number would not be accessed ever
      Found = True
      If TempDsplFlag = True Then 'user selected an asset already in disposal processing
        TempDsplFlag = False
        Message = "This item has been selected for disposal on " + MakeRegDate(FAItemRec.DispDate) + ". Please access this item using the Edit Item Disposal List button on the Disposal Menu."
        MsgBox (Message)
        Close
        Exit Sub
      End If
      fpListSearch.Row = -1
      fpListSearch.Col = 0
      GRecNum = x 'match is found so assign global variable
      Exit For 'got what we want so jump out of loop
    Else
      Found = False
    End If
      
  Next x
  
  Close FAHandle
  frmFAPostSnglDspl.Show
  DoEvents
  Unload frmFADsplSingle
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFADsplSingle", "fpListSearch_DblClick", Erl)
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
    Unload Me
  
End Sub

