VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "BTN32A20.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmCustVehicles 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vehicle Information"
   ClientHeight    =   6810
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   10545
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   10545
   StartUpPosition =   2  'CenterScreen
   Begin LpLib.fpCombo fpBusPers 
      Height          =   375
      Left            =   2880
      TabIndex        =   9
      Top             =   5115
      Width           =   615
      _Version        =   196608
      _ExtentX        =   1085
      _ExtentY        =   661
      Text            =   ""
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
      Text            =   ""
      Columns         =   1
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
      ThreeDOutsideStyle=   2
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   -2147483642
      BorderWidth     =   1
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
      ScrollBarH      =   3
      DataFieldList   =   ""
      ColumnEdit      =   0
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
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmCustVehicles.frx":0000
   End
   Begin LpLib.fpCombo fpDecalCat 
      Height          =   375
      Left            =   2880
      TabIndex        =   0
      Top             =   1530
      Width           =   4140
      _Version        =   196608
      _ExtentX        =   7302
      _ExtentY        =   661
      Text            =   ""
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
      Text            =   ""
      Columns         =   3
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
      ThreeDOutsideStyle=   2
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   -2147483642
      BorderWidth     =   1
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   1
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
      AutoSearchFillDelay=   500
      EditMarginLeft  =   2
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmCustVehicles.frx":032E
   End
   Begin EditLib.fpLongInteger fpVehRecNo 
      Height          =   300
      Left            =   1608
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   24
      Visible         =   0   'False
      Width           =   684
      _Version        =   196608
      _ExtentX        =   1206
      _ExtentY        =   529
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   0
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   0
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   1
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
      AllowNull       =   -1  'True
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
      Text            =   ""
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
      Appearance      =   1
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpLongInteger fpCustRecNo 
      Height          =   300
      Left            =   768
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   24
      Visible         =   0   'False
      Width           =   684
      _Version        =   196608
      _ExtentX        =   1206
      _ExtentY        =   529
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   0
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   0
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   1
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
      AllowNull       =   -1  'True
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
      Text            =   ""
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
      Appearance      =   1
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpBoolean fpValid 
      Height          =   324
      Left            =   2880
      TabIndex        =   4
      Top             =   2784
      Width           =   324
      _Version        =   196608
      _ExtentX        =   572
      _ExtentY        =   572
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      AutoToggle      =   -1  'True
      BooleanStyle    =   1
      ToggleFalse     =   "Nn"
      TextFalse       =   "N"
      BooleanPicture  =   0
      AlignPictureH   =   3
      AlignPictureV   =   1
      GroupId         =   0
      GroupTag        =   0
      GroupSelect     =   0
      MarginLeft      =   3
      MarginTop       =   0
      MarginRight     =   3
      MarginBottom    =   3
      MultiLine       =   0   'False
      AlignTextH      =   1
      AlignTextV      =   0
      ToggleTrue      =   "Yy"
      TextTrue        =   "Y"
      Value           =   0
      BooleanMode     =   0
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      BorderGrayAreaColor=   -2147483637
      ToggleGrayed    =   ""
      TextGrayed      =   ""
      AllowMnemonic   =   0   'False
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDOnFocusInvert=   0   'False
      Caption         =   "N"
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      BooleanDataType =   0
      OLEDropMode     =   0
   End
   Begin EditLib.fpText fpStateLic 
      CausesValidation=   0   'False
      Height          =   324
      Left            =   2880
      TabIndex        =   6
      Top             =   3888
      Width           =   5172
      _Version        =   196608
      _ExtentX        =   9123
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
      MarginTop       =   0
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
   Begin EditLib.fpDateTime fpExpireDate 
      Height          =   324
      Left            =   2880
      TabIndex        =   3
      Top             =   2376
      Width           =   1692
      _Version        =   196608
      _ExtentX        =   2984
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
      ButtonStyle     =   2
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483637
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   0
      HideSelection   =   0   'False
      InvalidColor    =   -2147483643
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483643
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   "11/06/2001"
      DateCalcMethod  =   0
      DateTimeFormat  =   5
      UserDefinedFormat=   "mm/dd/yyyy"
      DateMax         =   "20350101"
      DateMin         =   "19800101"
      TimeMax         =   "000000"
      TimeMin         =   "000000"
      TimeString1159  =   ""
      TimeString2359  =   ""
      DateDefault     =   "19800101"
      TimeDefault     =   "000000"
      TimeStyle       =   0
      BorderGrayAreaColor=   -2147483637
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
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
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      StartMonth      =   4
      ButtonAlign     =   0
      BoundDataType   =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpDoubleSingle fpFee 
      Height          =   348
      Left            =   7848
      TabIndex        =   1
      Top             =   1536
      Width           =   1380
      _Version        =   196608
      _ExtentX        =   2434
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
      AlignTextV      =   1
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   1
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
      Text            =   "0.00"
      DecimalPlaces   =   2
      DecimalPoint    =   ""
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
   Begin EditLib.fpText fpstickernum 
      CausesValidation=   0   'False
      Height          =   324
      Left            =   2880
      TabIndex        =   2
      Top             =   1968
      Width           =   2100
      _Version        =   196608
      _ExtentX        =   3704
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
      MarginTop       =   0
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
      MaxLength       =   12
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
   Begin EditLib.fpText fpMakeModl 
      CausesValidation=   0   'False
      Height          =   324
      Left            =   2880
      TabIndex        =   5
      Top             =   3480
      Width           =   4284
      _Version        =   196608
      _ExtentX        =   7556
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
      MarginTop       =   0
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
      MaxLength       =   25
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
   Begin EditLib.fpText fpVINDesc 
      CausesValidation=   0   'False
      Height          =   324
      Left            =   2880
      TabIndex        =   7
      Top             =   4296
      Width           =   6324
      _Version        =   196608
      _ExtentX        =   11155
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
      MarginTop       =   0
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
      MaxLength       =   40
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
   Begin EditLib.fpText fpNotes 
      CausesValidation=   0   'False
      Height          =   324
      Left            =   2880
      TabIndex        =   8
      Top             =   4704
      Width           =   6324
      _Version        =   196608
      _ExtentX        =   11155
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
      MarginTop       =   0
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
      MaxLength       =   40
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
   Begin fpBtnAtlLibCtl.fpBtn fpExit 
      Height          =   390
      Left            =   8070
      TabIndex        =   11
      Top             =   6240
      Width           =   1260
      _Version        =   131072
      _ExtentX        =   2222
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
      ButtonDesigner  =   "frmCustVehicles.frx":06A9
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdSave 
      Height          =   384
      Left            =   6354
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   6240
      Width           =   1248
      _Version        =   131072
      _ExtentX        =   2201
      _ExtentY        =   677
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   -1  'True
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
      ButtonDesigner  =   "frmCustVehicles.frx":0885
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdVehList 
      Height          =   384
      Left            =   4632
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   6240
      Width           =   1248
      _Version        =   131072
      _ExtentX        =   2201
      _ExtentY        =   677
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   -1  'True
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
      ButtonDesigner  =   "frmCustVehicles.frx":0A61
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdDelete 
      Height          =   390
      Left            =   2910
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   6240
      Width           =   1245
      _Version        =   131072
      _ExtentX        =   2201
      _ExtentY        =   677
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   -1  'True
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
      ButtonDesigner  =   "frmCustVehicles.frx":0C3C
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdAddNew 
      Height          =   390
      Left            =   1215
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   6240
      Width           =   1215
      _Version        =   131072
      _ExtentX        =   2143
      _ExtentY        =   688
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   -1  'True
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
      ButtonDesigner  =   "frmCustVehicles.frx":0E19
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000005&
      Height          =   468
      Left            =   2502
      Top             =   720
      Width           =   5532
   End
   Begin VB.Label lblHead 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Edit an Existing Vehicle"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   2610
      TabIndex        =   26
      Top             =   792
      Width           =   5412
   End
   Begin VB.Label lblmsg 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "MORE VEHICLES ON FILE, PRESS F10 FOR NEXT VEHICLE OR F5 FOR LIST"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   300
      Left            =   726
      TabIndex        =   24
      Top             =   5808
      Visible         =   0   'False
      Width           =   9084
   End
   Begin VB.Line Line1 
      X1              =   804
      X2              =   9000
      Y1              =   3288
      Y2              =   3288
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Business/Personal:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   384
      TabIndex        =   23
      Top             =   5088
      Width           =   2388
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "State License #:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   504
      TabIndex        =   22
      Top             =   3888
      Width           =   2268
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle Make/Model:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   240
      TabIndex        =   21
      Top             =   3492
      Width           =   2532
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fee:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   6408
      TabIndex        =   20
      Top             =   1584
      Width           =   1332
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Valid:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   1344
      TabIndex        =   19
      Top             =   2808
      Width           =   1428
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Decal Expires:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   600
      TabIndex        =   18
      Top             =   2412
      Width           =   2172
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Decal Sticker Number:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   192
      TabIndex        =   17
      Top             =   2004
      Width           =   2580
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Notes:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   1368
      TabIndex        =   15
      Top             =   4692
      Width           =   1404
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "VIN#/Description:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   408
      TabIndex        =   14
      Top             =   4296
      Width           =   2364
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Decal Category:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   504
      TabIndex        =   13
      Top             =   1608
      Width           =   2268
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   588
      Left            =   2502
      Top             =   600
      Width           =   5532
   End
End
Attribute VB_Name = "frmCustVehicles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim exitthisun As Boolean
Public Sub SetVehScreen()
  Dim NumOfVRecs As Long, DCvFile As Integer, DCVehReclen As Integer
  ReDim DCVRec(1) As DCVehType
startloop:
  If fpVehRecNo > 0 Then
    DCVehReclen = Len(DCVRec(1))
    DCvFile = FreeFile
    Open "DCVEH.DAT" For Random Access Read Write Shared As DCvFile Len = DCVehReclen
    NumOfVRecs = LOF(DCvFile) \ DCVehReclen
    Get DCvFile, fpVehRecNo, DCVRec(1)
    Close DCvFile
     ' AddFlag = 1
    If DCVRec(1).Active <> "Y" Then
      If DCVRec(1).NextRec <= 0 Then
        fpVehRecNo = 0
      Else
        fpVehRecNo = DCVRec(1).NextRec
      End If
      GoTo startloop
    End If
    lblHead.Caption = "Edit Existing Vehicle"
    If DCVRec(1).NextRec > 0 Then
      lblmsg.Visible = True
    Else
      lblmsg.Visible = False
    End If
    fpDecalCat.SearchText = QPTrim$(DCVRec(1).DecalCat)
    fpDecalCat.ColumnSearch = 1
    fpDecalCat.Action = ActionSearch
    If fpDecalCat.SearchIndex <> -1 Then
      fpDecalCat.ListIndex = fpDecalCat.SearchIndex
    End If
    fpFee = DCVRec(1).Fee
    fpMakeModl = DCVRec(1).makemodel
    fpStateLic = DCVRec(1).StateTag
    fpExpireDate = Num2Date$(DCVRec(1).ExpireDate)
    fpstickernum = DCVRec(1).Sticker
    Select Case DCVRec(1).Valid
    Case "N", " "
      fpValid.Value = ValueFalse
    Case "Y"
      fpValid.Value = ValueTrue
    End Select
    fpVinDesc = DCVRec(1).Desc
    fpNotes = DCVRec(1).Notes
    If DCVRec(1).PBFlag = "P" Then
      fpBusPers.ListIndex = 0
    ElseIf DCVRec(1).PBFlag = "B" Then
      fpBusPers.ListIndex = 1
    Else
      fpBusPers.ListIndex = -1
    End If
  Else
    lblHead.Caption = "Add New Vehicle"
    fpDecalCat.ListIndex = -1
    fpFee = 0
    fpMakeModl = ""
    fpStateLic = ""
    fpExpireDate = ""
    fpstickernum = ""
    fpValid.Value = ValueFalse
    fpVinDesc = ""
    fpNotes = ""
    fpBusPers.ListIndex = -1
  End If
End Sub

Private Sub Form_Activate()
  If exitthisun = False Then
    SetVehScreen
  End If
End Sub
Private Sub DelVehicle()
  Dim NumOfVRecs As Long, DCvFile As Integer, DCVehReclen As Integer
  Dim VehRecord As Long
  ReDim DCVRec(1) As DCVehType
  DCVehReclen = Len(DCVRec(1))
  DCvFile = FreeFile
  Open "DCVEH.DAT" For Random Access Read Write Shared As DCvFile Len = DCVehReclen
  NumOfVRecs = LOF(DCvFile) \ DCVehReclen

  VehRecord = fpVehRecNo
  Get DCvFile, VehRecord, DCVRec(1)
  'DCVRec(1).NextRec = 0
  DCVRec(1).Active = "N"
  'DCVRec(1).MasterRecord = -1
  Put DCvFile, VehRecord, DCVRec(1)
  Close
  DCLog PWUser$ + " Vehicle Delete " + Str(VehRecord)
  MsgBox "Vehicle Deleted", vbOKOnly, "Deleted"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      KeyCode = 0
      Call fpExit_Click
    Case vbKeyF10
      KeyCode = 0
      DoEvents
      fpCmdSave_Click
    Case vbKeyF2
      KeyCode = 0
      DoEvents
      fpCmdAddNew_Click
    Case vbKeyF3
      KeyCode = 0
      DoEvents
      fpCmdDelete_Click
    Case vbKeyF5
      KeyCode = 0
      DoEvents
      fpCmdVehList_Click
    Case Else:
  End Select
End Sub
Private Sub GetVehList()
  Dim NumOfDCRecs As Long, DCFile As Integer, Num1 As Long, Num2 As Long
  Dim cnt As Long, dcnt As Long, Cust As String
  Dim Build As String * 80

  ReDim DCCustRec(1) As DCCustRecType
  If fpCustRecNo > 0 Then
    OpenDCCustFile NumOfDCRecs, DCFile
    Get DCFile, fpCustRecNo, DCCustRec(1)
    Close DCFile
    If DCCustRec(1).FirstCar <= 0 Then
      MsgBox "No Vehicles to Display", vbOKOnly, "No Vehicles"
      Exit Sub
    Else
      Num1 = DCCustRec(1).FirstCar
      Cust$ = QPTrim$(DCCustRec(1).BILLNAME)
    End If
  
  Dim NumOfVRecs As Long, DCvFile As Integer, DCVehReclen As Integer
  ReDim DCVRec(1) As DCVehType
  If Num1 > 0 Then
    DCVehReclen = Len(DCVRec(1))
    DCvFile = FreeFile
    Open "DCVEH.DAT" For Random Access Read Write Shared As DCvFile Len = DCVehReclen
    NumOfVRecs = LOF(DCvFile) \ DCVehReclen
    cnt = Num1
    Do Until cnt = 0
    'For cnt = Num1 To Num2
    Get DCvFile, cnt, DCVRec(1)
      
    If DCVRec(1).Active = "Y" Then
      LSet Build$ = QPTrim$(DCVRec(1).makemodel)
      Mid$(Build$, 30) = QPTrim$(DCVRec(1).StateTag)
      Mid$(Build$, 55) = QPTrim$(DCVRec(1).Desc)
      Mid$(Build$, 75) = Chr9$ + Str$(cnt)
      frmVehDisplayList.fpList1.AddItem Build$
      dcnt = dcnt + 1
    End If
      cnt = DCVRec(1).NextRec
    Loop 'Next
    Close DCvFile
    If dcnt > 0 Then
      frmVehDisplayList.fpList1.ListIndex = 0
      frmVehDisplayList.Caption = "Vehicle List - " & Cust$
      frmVehDisplayList.Show 1
    End If
  End If
    If dcnt <= 0 Then
      MsgBox "No Vehicles to Display", vbOKOnly, "No Vehicles"
    End If
  End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  exitthisun = False
End Sub

'Private Sub fpBusPers_Change()
'  If fpBusPers.Text <> "B" Or fpBusPers.Text <> "P" Then
'    fpBusPers.Text = ""
'  End If
'End Sub

'Private Sub fpBusPers_ChangeMode(EditMode As Integer)
'  EditMode = True
'End Sub

'Private Sub fpBusPers_LostFocus()
'  If fpBusPers.Text <> "B" Or fpBusPers.Text <> "P" Then
'    fpBusPers.Text = ""
'  End If
'End Sub

'Private Sub Form_Load()
'  exitthisun = False
'End Sub

Private Sub fpCmdAddNew_Click()
    Select Case CheckSaveCustFile%
    Case True:  '-1 save chenges
        Call SaveVeh
        fpVehRecNo = 0
        SetVehScreen
 '     End If
      
 '   End If
    Case False:  '0= exit
      fpVehRecNo = 0
      SetVehScreen
    Case Else     '1 is review
      'stay right where you are
    End Select
End Sub

Private Sub fpCmdDelete_Click()
  If MsgBox("Are you sure you wish to delete this vehicle?", vbYesNo, "Delete Vehicle?") = vbYes Then
    DelVehicle
    SetVehScreen
  End If
End Sub

Private Sub fpCmdSave_Click()
  Dim NumOfVRecs As Long, DCvFile As Integer, DCVehReclen As Integer
  Dim NextRec As Long
  ReDim DCVRec(1) As DCVehType
  DCVehReclen = Len(DCVRec(1))
  If ChkVehInfoOK Then
    Call SaveVeh
    If fpVehRecNo > 0 Then
      DCvFile = FreeFile
      Open "DCVEH.DAT" For Random Access Read Write Shared As DCvFile Len = DCVehReclen
      NumOfVRecs = LOF(DCvFile) \ DCVehReclen
      Get DCvFile, fpVehRecNo, DCVRec(1)
      Close DCvFile
      If DCVRec(1).NextRec > 0 Then
        fpVehRecNo = DCVRec(1).NextRec
        SetVehScreen
      Else
        Unload Me
      End If
    Else
      Unload Me
    End If
  End If
End Sub

Private Sub fpCmdVehList_Click()
'need to check see if should save current vehicle first
    Select Case CheckSaveCustFile%
    Case True:  '-1 save chenges
    If ChkVehInfoOK Then
      Call SaveVeh
      GetVehList
    End If
    Case False:  '0= exit
       GetVehList
    Case Else     '1 is review
      'stay right where you are
    End Select
End Sub

'
'AddCars:
'  VehRecord! = DCCustRec(1).FirstCar
'  If VehRecord! <= 0 Then
'    AddFlag = 1
'    VehRecord! = NumOfVrecs + 1
'    GoTo Addfirst
'  End If
'
'MasterLoop:
'  Get DCVFile, VehRecord!, DCVRec(1)
'  If DCVRec(1).Active <> "Y" Then
'    VehRecord! = DCVRec(1).NextRec
'    If VehRecord! <= 0 Then Close: Exit Sub
'  End If
'  GoSub GetVehRecord
'
'      GoSub SaveVRecord
'      frm(1).FldNo = 3
'      If AddFlag = 0 Then
'        VehRecord! = DCVRec(1).NextRec
'        If VehRecord! > 0 Then
'          GoTo MasterLoop
'        End If
'      Else
'        For F = 1 To NumFlds
'          LSet Form$(F, 0) = ""
'        Next F
'        VehRecord! = NumOfVrecs + 1
'        Action = 1
'        AddFlag = 1
'      End If
'Return
Private Function CheckSaveCustFile%()
  Dim NumOfDCRecs As Long, DCFile As Integer, VehRecord As Long
  Dim NumOfVRecs As Long, DCvFile As Integer, DCVehReclen As Integer
  Dim PrevRec As Long, chgcnt As Integer
  ReDim DCVRec(1) As DCVehType
  DCVehReclen = Len(DCVRec(1))
  DCvFile = FreeFile
  chgcnt = 0
  If fpVehRecNo > 0 Then
    Open "DCVEH.DAT" For Random Access Read Write Shared As DCvFile Len = DCVehReclen
    NumOfVRecs = LOF(DCvFile) \ DCVehReclen
    Get DCvFile, fpVehRecNo, DCVRec(1)
    Close DCvFile
  '  If fpDecalCat.ListIndex = -1 Then Return
    fpDecalCat.col = 1
    If QPTrim$(DCVRec(1).DecalCat) <> QPTrim$(fpDecalCat.ColText) Then chgcnt = chgcnt + 1
    If DCVRec(1).Fee <> fpFee.Value Then chgcnt = chgcnt + 1
    If QPTrim$(DCVRec(1).makemodel) <> QPTrim$(fpMakeModl) Then chgcnt = chgcnt + 1
    If QPTrim$(DCVRec(1).StateTag) <> QPTrim$(fpStateLic) Then chgcnt = chgcnt + 1
    If DCVRec(1).ExpireDate <> Date2Num%(fpExpireDate) Then chgcnt = chgcnt + 1
    If QPTrim$(DCVRec(1).Sticker) <> QPTrim$(fpstickernum) Then chgcnt = chgcnt + 1
    If QPTrim$(DCVRec(1).Valid) <> QPTrim$(fpValid.Text) Then chgcnt = chgcnt + 1
    If QPTrim$(DCVRec(1).Desc) <> QPTrim$(fpVinDesc) Then chgcnt = chgcnt + 1
    If QPTrim$(DCVRec(1).Notes) <> QPTrim$(fpNotes) Then chgcnt = chgcnt + 1
    If QPTrim$(DCVRec(1).PBFlag) <> QPTrim$(fpBusPers.Text) Then chgcnt = chgcnt + 1
  Else
    If fpDecalCat.ListIndex <> -1 Then chgcnt = chgcnt + 1
    If fpFee.Value > 0 Then chgcnt = chgcnt + 1
    If QPTrim$(fpMakeModl) <> "" Then chgcnt = chgcnt + 1
    If QPTrim$(fpStateLic) <> "" Then chgcnt = chgcnt + 1
    If QPTrim$(fpstickernum) <> "" Then chgcnt = chgcnt + 1
    If QPTrim$(fpValid.Text) <> "N" Then chgcnt = chgcnt + 1
    If QPTrim$(fpVinDesc) <> "" Then chgcnt = chgcnt + 1
    If QPTrim$(fpNotes) <> "" Then chgcnt = chgcnt + 1
    If QPTrim$(fpBusPers.Text) <> "" Then chgcnt = chgcnt + 1
  End If
  If chgcnt > 0 Then
    frmChangedWarning.Show vbModal
    Select Case SaveFlag
    Case False
      CheckSaveCustFile% = False
    Case True
      CheckSaveCustFile% = True
    Case 1
      CheckSaveCustFile% = 1
    End Select
  Else
    CheckSaveCustFile% = False
  End If

End Function

Private Sub SaveVeh()
  Dim NumOfDCRecs As Long, DCFile As Integer, VehRecord As Long
  Dim NumOfVRecs As Long, DCvFile As Integer, DCVehReclen As Integer
  Dim PrevRec As Long
  ReDim DCVRec(1) As DCVehType
  DCVehReclen = Len(DCVRec(1))
  DCvFile = FreeFile
  Open "DCVEH.DAT" For Random Access Read Write Shared As DCvFile Len = DCVehReclen
  NumOfVRecs = LOF(DCvFile) \ DCVehReclen
  If fpVehRecNo > 0 Then
    Get DCvFile, fpVehRecNo, DCVRec(1)
  End If
  fpDecalCat.col = 1
  DCVRec(1).DecalCat = QPTrim$(fpDecalCat.ColText)
  DCVRec(1).Fee = fpFee.Value
  DCVRec(1).makemodel = QPTrim$(fpMakeModl)
  DCVRec(1).StateTag = QPTrim$(fpStateLic)
  DCVRec(1).ExpireDate = Date2Num%(fpExpireDate)
  DCVRec(1).Sticker = QPTrim$(fpstickernum)
  Select Case fpValid.Value
    Case ValueTrue
      DCVRec(1).Valid = "Y"
    Case ValueFalse
      DCVRec(1).Valid = "N"
  End Select
  DCVRec(1).Active = "Y"
  DCVRec(1).Desc = QPTrim$(fpVinDesc)
  DCVRec(1).Notes = QPTrim$(fpNotes)
  DCVRec(1).PBFlag = QPTrim$(fpBusPers.Text)
  DCVRec(1).MoreRoom = ""
  DCVRec(1).MasterRecord = fpCustRecNo
  If fpVehRecNo = 0 Then
    DCVRec(1).NextRec = 0
    VehRecord = NumOfVRecs + 1
  Else
    VehRecord = fpVehRecNo
  End If
  Put DCvFile, VehRecord, DCVRec(1)
  exitthisun = False
  DCLog PWUser$ + " Vehicle saved, " + Str(VehRecord)
  If fpVehRecNo = 0 Then
    GoSub UpdateVendorPointer
  End If
'Return

'  If fpVehRecNo = 0 Then
'    DCVRec(1).NextRec = 0
'    VehRecord = NumOfVRecs + 1
'    'fpVehRecNo = VehRecord
'  Else
'    VehRecord = fpVehRecNo
'  End If
'  Put DCvFile, VehRecord, DCVRec(1)
'  If fpVehRecNo > 0 Then
'    GoSub UpdateVendorPointer
'  End If
'  fpVehRecNo = VehRecord
Exit Sub
UpdateVendorPointer:
  ReDim DCCustRec(1) As DCCustRecType
  If fpCustRecNo > 0 Then
    OpenDCCustFile NumOfDCRecs, DCFile
    Get DCFile, fpCustRecNo, DCCustRec(1)
    If DCCustRec(1).FirstCar = 0 Then
      DCCustRec(1).FirstCar = VehRecord
      DCCustRec(1).LastCar = VehRecord
      Put DCFile, fpCustRecNo, DCCustRec(1)
    Else
      PrevRec = DCCustRec(1).LastCar
      DCCustRec(1).LastCar = VehRecord
      Put DCFile, fpCustRecNo, DCCustRec(1)
  
      Get DCvFile, PrevRec, DCVRec(1)
      DCVRec(1).NextRec = VehRecord
      Put DCvFile, PrevRec, DCVRec(1)
    End If
    Close DCvFile
  End If
  NumOfVRecs = NumOfVRecs + 1
Return
'UpdateVendorPointer:
'  If DCCustRec(1).FirstCar = 0 Then
'    DCCustRec(1).FirstCar = VehRecord!
'    DCCustRec(1).LastCar = VehRecord!
'    Put DCFile, AccountRecord, DCCustRec(1)
'  Else
'    PrevRec! = DCCustRec(1).LastCar
'    DCCustRec(1).LastCar = VehRecord!
'    Put DCFile, AccountRecord, DCCustRec(1)
'    Get DCvFile, PrevRec!, DCVRec(1)
'    DCVRec(1).NextRec = VehRecord!
'    Put DCvFile, PrevRec!, DCVRec(1)
'  End If
'  NumOfVRecs = NumOfVRecs + 1
'Return


End Sub

Private Sub fpDecalCat_Change()
  Dim lookrec As Integer
  Dim DCCatCodeRec As DCCatCodeRecType
  Dim DCCatCodeRecLen As Integer, ghandle As Integer
  Dim NumOFDCCatRecs As Integer
  DCCatCodeRecLen = Len(DCCatCodeRec)
  If fpDecalCat.ListIndex <> -1 Then
    fpDecalCat.col = 0
    lookrec = QPTrim$(fpDecalCat.ColText)
    ghandle = FreeFile
    Open "DCCODE.DAT" For Random Access Read Write Shared As ghandle Len = DCCatCodeRecLen
    NumOFDCCatRecs = LOF(ghandle) \ DCCatCodeRecLen
    Get #ghandle, lookrec, DCCatCodeRec
      fpFee = DCCatCodeRec.Fee
    Close ghandle
  End If
End Sub
Private Sub fpDecalCat_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpDecalCat.ListDown = True
  End If
  If fpDecalCat.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpFee.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpCmdSave.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub

Private Sub fpExit_Click()
    exitthisun = True
    Select Case CheckSaveCustFile%
    Case True:  '-1 save chenges
      
      If ChkVehInfoOK Then
        Call SaveVeh
        Unload Me
      End If
    Case False:  '0= exit
      Unload Me
    Case Else     '1 is review
      'stay right where you are
    End Select
  
End Sub
Private Function ChkVehInfoOK()
 Dim notEnoughtosave As Integer
  notEnoughtosave = True
  notEnoughtosave = 0
  ChkVehInfoOK = False   'assume the worst.

  If fpDecalCat.ListIndex = -1 Then notEnoughtosave = notEnoughtosave + 1
  If QPTrim$(fpMakeModl) = "" Then notEnoughtosave = notEnoughtosave + 1
  
DoneChk:
  DoEvents
  ReDim MsgText(0 To 5) As String
  Dim FntSize As Integer
  If notEnoughtosave > 0 Then
    frmMsgDialog.RetLabel = "-2"
    FntSize = frmMsgDialog.Label(2).FontSize
    frmMsgDialog.Label(2).FontSize = (FntSize + 2)
    frmMsgDialog.Label(1).FontSize = (FntSize + 2)
    frmMsgDialog.Label(4).FontSize = (FntSize + 2)
    MsgText(0) = "ERROR:"
    MsgText(1) = ""
    MsgText(2) = "The Code and Make/Model"
    MsgText(3) = "are Required Fields."
    MsgText(4) = ""
    MsgText(5) = "Please Enter This Information."
    GetOKorNot MsgText(), True
    ChkVehInfoOK = False
  Else
    ChkVehInfoOK = True
  End If

End Function

Private Sub fpFee_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpstickernum.SetFocus
  End If
End Sub
Private Sub fpstickernum_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpExpireDate.SetFocus
  End If
End Sub
Private Sub fpExpireDate_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpValid.SetFocus
  End If
End Sub
Private Sub fpValid_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpMakeModl.SetFocus
  End If
End Sub

Private Sub fpMakeModl_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpStateLic.SetFocus
  End If
End Sub

Private Sub fpStateLic_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpVinDesc.SetFocus
  End If
End Sub
Private Sub fpVINDesc_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpNotes.SetFocus
  End If
End Sub

Private Sub fpNotes_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpBusPers.SetFocus
  End If
End Sub

Private Sub fpBusPers_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpBusPers.ListDown = True
  End If
  If fpBusPers.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpCmdSave.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpNotes.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub

