VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmVATaxRefundOnPrepay 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Payment Refund On Prepaid Amounts"
   ClientHeight    =   8760
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   11655
   Icon            =   "frmVATaxRefundOnPrepay.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpList fpList1 
      Height          =   1800
      Left            =   1080
      TabIndex        =   23
      Top             =   1320
      Width           =   9372
      _Version        =   196608
      _ExtentX        =   16531
      _ExtentY        =   3175
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
      Columns         =   3
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
      ColDesigner     =   "frmVATaxRefundOnPrepay.frx":08CA
   End
   Begin VB.Timer MsgAlertTimer 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   0
      Top             =   0
   End
   Begin EditLib.fpCurrency fpCurrRefund 
      Height          =   372
      Left            =   5640
      TabIndex        =   1
      Top             =   6600
      Width           =   1860
      _Version        =   196608
      _ExtentX        =   3281
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
      MinValue        =   "0"
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
   Begin EditLib.fpDateTime fptxtDate 
      Height          =   408
      Left            =   5652
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   6120
      Width           =   1860
      _Version        =   196608
      _ExtentX        =   3281
      _ExtentY        =   720
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
      AutoAdvance     =   0   'False
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
      Text            =   "01/31/2005"
      DateCalcMethod  =   0
      DateTimeFormat  =   5
      UserDefinedFormat=   "mm/dd/yyyy"
      DateMax         =   "20350101"
      DateMin         =   "19800101"
      TimeMax         =   "000000"
      TimeMin         =   "000000"
      TimeString1159  =   ""
      TimeString2359  =   ""
      DateDefault     =   "20010101"
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
      ButtonColor     =   13684944
      AutoMenu        =   0   'False
      StartMonth      =   4
      ButtonAlign     =   0
      BoundDataType   =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fptxtState 
      Height          =   372
      Left            =   4200
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5352
      Width           =   612
      _Version        =   196608
      _ExtentX        =   1080
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
      MaxLength       =   2
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
   Begin EditLib.fpText fptxtName 
      Height          =   372
      Left            =   4200
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   4080
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
      AlignTextH      =   0
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
   Begin EditLib.fpText fptxtAddress 
      Height          =   372
      Left            =   4200
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4500
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
      AlignTextH      =   0
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
   Begin EditLib.fpText fptxtCity 
      Height          =   372
      Left            =   4200
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4920
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
      AlignTextH      =   0
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
   Begin EditLib.fpMask fptxtZip 
      Height          =   372
      Left            =   6840
      TabIndex        =   9
      TabStop         =   0   'False
      Tag             =   "This field contains the postal code for this business. This field cannot be edited."
      Top             =   5352
      Width           =   1452
      _Version        =   196608
      _ExtentX        =   2561
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
      AllowOverflow   =   0   'False
      BestFit         =   0   'False
      ClipMode        =   0
      DataFormatEx    =   0
      Mask            =   "#####-####"
      PromptChar      =   "_"
      PromptInclude   =   0   'False
      RequireFill     =   0   'False
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      AutoTab         =   0   'False
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fptxtNote 
      Height          =   348
      Left            =   4080
      TabIndex        =   2
      Top             =   7200
      Width           =   4140
      _Version        =   196608
      _ExtentX        =   7302
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
      ThreeDOutsideStyle=   2
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
      CaretInsert     =   2
      CaretOverWrite  =   2
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
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   20
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
   Begin fpBtnAtlLibCtl.fpBtn cmdListAlpha 
      Height          =   492
      Left            =   1980
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3360
      Width           =   2772
      _Version        =   131072
      _ExtentX        =   4890
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
      ButtonDesigner  =   "frmVATaxRefundOnPrepay.frx":0C10
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdListAcctNum 
      Height          =   492
      Left            =   6900
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   3360
      Width           =   2772
      _Version        =   131072
      _ExtentX        =   4890
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
      ButtonDesigner  =   "frmVATaxRefundOnPrepay.frx":0DF4
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPost 
      Height          =   495
      Left            =   7890
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   7800
      Width           =   1560
      _Version        =   131072
      _ExtentX        =   2752
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
      ButtonDesigner  =   "frmVATaxRefundOnPrepay.frx":0FE1
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   492
      Left            =   2196
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   7800
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
      ButtonDesigner  =   "frmVATaxRefundOnPrepay.frx":11BD
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdCustData 
      Height          =   492
      Left            =   6000
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   7800
      Width           =   1584
      _Version        =   131072
      _ExtentX        =   2794
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
      ButtonDesigner  =   "frmVATaxRefundOnPrepay.frx":1399
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdInfo 
      Height          =   492
      Left            =   4128
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   7800
      Width           =   1584
      _Version        =   131072
      _ExtentX        =   2794
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
      ButtonDesigner  =   "frmVATaxRefundOnPrepay.frx":1579
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Refund Amount:"
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
      Left            =   3600
      TabIndex        =   16
      Top             =   6720
      Width           =   1812
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   2052
      Left            =   960
      Top             =   1200
      Width           =   9612
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Note:"
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
      Height          =   336
      Left            =   3144
      TabIndex        =   15
      Top             =   7236
      Width           =   708
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Zip:"
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
      Left            =   5880
      TabIndex        =   14
      Top             =   5448
      Width           =   852
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "State:"
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
      TabIndex        =   13
      Top             =   5448
      Width           =   852
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "City:"
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
      TabIndex        =   12
      Top             =   5028
      Width           =   852
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
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
      Left            =   3120
      TabIndex        =   11
      Top             =   4608
      Width           =   972
   End
   Begin VB.Label Label2 
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
      TabIndex        =   10
      Top             =   4200
      Width           =   852
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000009&
      Height          =   600
      Left            =   2316
      Top             =   360
      Width           =   7008
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Refund On Prepayment Amounts"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2880
      TabIndex        =   4
      Top             =   480
      Width           =   5892
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Transaction Date:"
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
      Left            =   3612
      TabIndex        =   3
      Top             =   6240
      Width           =   1812
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   1932
      Left            =   3012
      Top             =   3960
      Width           =   5400
   End
   Begin VB.Shape Shape6 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   756
      Left            =   2316
      Top             =   240
      Width           =   7020
   End
End
Attribute VB_Name = "frmVATaxRefundOnPrepay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class
  Dim BtnFnt As Double
  Dim RefCustNum As Long

Private Sub cmdCustData_Click()
  If fpList1.ListIndex = -1 Then
    Call TaxMsg(900, "Please make a selection from the customer list.")
    Exit Sub
  End If
  
  If RefCustNum = 0 Then
    Exit Sub
  End If
  
  Call frmVATaxCustInq.LoadCust
  frmVATaxCustInq.Show vbModal
End Sub

Private Sub cmdExit_Click()
  GCustNum = 0
  KillFile "C:\CPWork\prepayrefund.dat"
  frmVATaxPayMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdInfo_Click()
  If fpList1.ListIndex = -1 Then
    Call TaxMsg(900, "Please make a selection from the customer list.")
    Exit Sub
  End If
  
  If RefCustNum > 0 Then
    frmVATaxMessage.Show vbModal
  End If
End Sub

Private Sub cmdListAcctNum_Click()
  Dim x As Long
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  
  On Error GoTo ERRORSTUFF
  
  fpList1.Clear
  OpenTaxCustFile TCHandle, NumOfTCRecs
  For x = 1 To RefNumCnt
    Get TCHandle, AcctNumList(x), TaxCust
    fpList1.InsertRow = CStr(TaxCust.Acct) + Chr(9) + QPTrim$(TaxCust.CustName) + Chr(9) + Using$("$###,##0.00", Abs(NumPreBal(x)))
  Next x
  Close TCHandle
  Call ClearData
  
  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxRefundOnPrepay", "cmdListAcctNum_Click", Erl)
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

Private Sub cmdListAlpha_Click()
  Dim x As Long
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  
  On Error GoTo ERRORSTUFF
  
  fpList1.Clear
  OpenTaxCustFile TCHandle, NumOfTCRecs
  For x = 1 To RefNameCnt
    Get TCHandle, AcctNameList(x), TaxCust
    fpList1.InsertRow = CStr(TaxCust.Acct) + Chr(9) + QPTrim$(TaxCust.CustName) + Chr(9) + Using$("$###,##0.00", Abs(NamePreBal(x)))
  Next x
  Close TCHandle
  Call ClearData
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxRefundOnPrepay", "cmdListAlpha_Click", Erl)
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

Private Sub cmdPost_Click()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim TransRec As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim ThisRec As Long
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim PersRefund As Double
  Dim RealRefund As Double
  Dim ThisBillType$
  Dim Split As Boolean
  
  On Error GoTo ERRORSTUFF
  
  Split = False
  If fpList1.ListIndex = -1 Then
    Call TaxMsg(900, "Please make a selection from the customer list.")
    Exit Sub
  End If
  If RefCustNum = 0 Then
    Call TaxMsg(800, "Please either double click your selection or press 'Enter' to load your selection.")
    Exit Sub
  End If
  If CDbl(fpCurrRefund.Value) = 0 Then
    Call TaxMsg(900, "The refund amount is zero. Post attempt aborted.")
    Exit Sub
  End If
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  If TaxMsgWOpts(900, "Are you sure you are ready to post? Press F10 to continue with posting. Otherwise, press ESC to abort this post.", "F10 Continue", "ESC Abort") = "abort" Then
    Unload frmVATaxMsgWOpts
    Exit Sub
  Else
    Unload frmVATaxMsgWOpts
    MainLog ("User given option of aborting refund posting for " + QPTrim$(fptxtName.Text) + " and they elected to contniue.")
  End If
  
  PersRefund = Abs(GetCustPersBalance(RefCustNum, -1))
  RealRefund = Abs(GetCustRealBalance(RefCustNum, -1))
  ThisBillType = ""
  If PersRefund = CDbl(fpCurrRefund.Value) Then
    ThisBillType = "P"
  ElseIf RealRefund = CDbl(fpCurrRefund.Value) Then
    ThisBillType = "R"
  ElseIf PersRefund + RealRefund = CDbl(fpCurrRefund.Value) Then
    Split = True
  End If
  
  OpenTaxTransFile TTHandle, NumOfTTRecs
  OpenTaxCustFile TCHandle, NumOfTCRecs
  Get TCHandle, RefCustNum, TaxCust
  If Split = True Then
    GoSub SplitIt
    GoTo Done
  End If
  ThisRec = NumOfTTRecs + 1
  TransRec.Amount = CDbl(fpCurrRefund.Value)
  TransRec.BelongTo = 0
  TransRec.BillType = ThisBillType
  TransRec.CustomerRec = RefCustNum
  TransRec.CustPin = TaxCust.PIN
  If Len(QPTrim$(fptxtNote.Text)) < 15 Then
    TransRec.Description = QPTrim$(fptxtNote.Text) + "/Prepay Refund"
  Else
    TransRec.Description = QPTrim$(fptxtNote.Text)
  End If
  TransRec.DiscAmt = 0
  TransRec.DiscXDate = 0
  TransRec.DMVBatch = 0
  TransRec.DMVSubmitted = ""
  TransRec.FromPrePay = "Y"
  TransRec.InternalPin = 0
  TransRec.LastTrans = TaxCust.LastTrans
  TaxCust.LastTrans = ThisRec
  TransRec.OperNum = OperNum
  TransRec.Padding = ""
  TransRec.PersPin = ""
  TransRec.Posted2GL = "N"
  TransRec.RealPin = ""
  If ThisBillType = "P" Then
    TransRec.TaxYear = TaxMasterRec.PTaxYear
  Else
    TransRec.TaxYear = TaxMasterRec.RTaxYear
  End If
  TransRec.TransDate = Date2Num(fptxtDate.Text)
  TransRec.TranType = 12 'refund on prepay
  TransRec.Revenue.Collection = 0
  TransRec.Revenue.CollectionPd = 0
  TransRec.Revenue.Future1 = 0
  TransRec.Revenue.Future1Pd = 0
  TransRec.Revenue.Future2 = 0
  TransRec.Revenue.Future2Pd = 0
  TransRec.Revenue.Interest = 0
  TransRec.Revenue.InterestPd = 0
  TransRec.Revenue.LateList = 0
  TransRec.Revenue.LateListPd = 0
  TransRec.Revenue.pad = ""
  TransRec.Revenue.Penalty = 0
  TransRec.Revenue.PenaltyPd = 0
  TransRec.Revenue.PrePaidAmt = 0
  TransRec.Revenue.PrePaidBal = 0
  TransRec.Revenue.PrePaidUsed = CDbl(fpCurrRefund.Value)
  TaxCust.PrePayBal = 0
  TaxCust.PrePayTrans = ThisRec
  TransRec.Revenue.Principle1 = 0
  TransRec.Revenue.Principle1Pd = 0
  TransRec.Revenue.Principle2 = 0
  TransRec.Revenue.Principle2Pd = 0
  TransRec.Revenue.Principle3 = 0
  TransRec.Revenue.Principle3Pd = 0
  TransRec.Revenue.Principle4 = 0
  TransRec.Revenue.Principle4Pd = 0
  TransRec.Revenue.Principle5 = 0
  TransRec.Revenue.Principle5Pd = 0
  TransRec.Revenue.RevOpt1 = 0
  TransRec.Revenue.RevOpt1Pd = 0
  TransRec.Revenue.RevOpt2 = 0
  TransRec.Revenue.RevOpt2Pd = 0
  TransRec.Revenue.RevOpt3 = 0
  TransRec.Revenue.RevOpt3Pd = 0
  Put TCHandle, RefCustNum, TaxCust
  Put TTHandle, ThisRec, TransRec
  Close
Done:
  Call Savemsg(900, "The refund transaction has been posted successfully.")
  Call cmdExit_Click
  
  Exit Sub
  
SplitIt:
  ThisRec = NumOfTTRecs + 1
  TransRec.Amount = PersRefund
  TransRec.BelongTo = 0
  TransRec.BillType = "P"
  TransRec.CustomerRec = RefCustNum
  TransRec.CustPin = TaxCust.PIN
  If Len(QPTrim$(fptxtNote.Text)) < 15 Then
    TransRec.Description = QPTrim$(fptxtNote.Text) + "/Prepay Refund"
  Else
    TransRec.Description = QPTrim$(fptxtNote.Text)
  End If
  TransRec.DiscAmt = 0
  TransRec.DiscXDate = 0
  TransRec.DMVBatch = 0
  TransRec.DMVSubmitted = ""
  TransRec.FromPrePay = "Y"
  TransRec.InternalPin = 0
  TransRec.LastTrans = TaxCust.LastTrans
  TaxCust.LastTrans = ThisRec
  TransRec.OperNum = OperNum
  TransRec.Padding = ""
  TransRec.PersPin = ""
  TransRec.Posted2GL = "N"
  TransRec.RealPin = ""
  TransRec.TaxYear = TaxMasterRec.PTaxYear
  TransRec.TransDate = Date2Num(fptxtDate.Text)
  TransRec.TranType = 12 'refund on prepay
  TransRec.Revenue.Collection = 0
  TransRec.Revenue.CollectionPd = 0
  TransRec.Revenue.Future1 = 0
  TransRec.Revenue.Future1Pd = 0
  TransRec.Revenue.Future2 = 0
  TransRec.Revenue.Future2Pd = 0
  TransRec.Revenue.Interest = 0
  TransRec.Revenue.InterestPd = 0
  TransRec.Revenue.LateList = 0
  TransRec.Revenue.LateListPd = 0
  TransRec.Revenue.pad = ""
  TransRec.Revenue.Penalty = 0
  TransRec.Revenue.PenaltyPd = 0
  TransRec.Revenue.PrePaidAmt = 0
  TransRec.Revenue.PrePaidBal = 0
  TransRec.Revenue.PrePaidUsed = PersRefund
  TaxCust.PrePayBal = RealRefund
  TaxCust.PrePayTrans = ThisRec
  TransRec.Revenue.Principle1 = 0
  TransRec.Revenue.Principle1Pd = 0
  TransRec.Revenue.Principle2 = 0
  TransRec.Revenue.Principle2Pd = 0
  TransRec.Revenue.Principle3 = 0
  TransRec.Revenue.Principle3Pd = 0
  TransRec.Revenue.Principle4 = 0
  TransRec.Revenue.Principle4Pd = 0
  TransRec.Revenue.Principle5 = 0
  TransRec.Revenue.Principle5Pd = 0
  TransRec.Revenue.RevOpt1 = 0
  TransRec.Revenue.RevOpt1Pd = 0
  TransRec.Revenue.RevOpt2 = 0
  TransRec.Revenue.RevOpt2Pd = 0
  TransRec.Revenue.RevOpt3 = 0
  TransRec.Revenue.RevOpt3Pd = 0
  Put TCHandle, RefCustNum, TaxCust
  Put TTHandle, ThisRec, TransRec
  
  'now real
  ThisRec = NumOfTTRecs + 2
  TransRec.Amount = RealRefund
  TransRec.BelongTo = 0
  TransRec.BillType = "R"
  TransRec.CustomerRec = RefCustNum
  TransRec.CustPin = TaxCust.PIN
  If Len(QPTrim$(fptxtNote.Text)) < 15 Then
    TransRec.Description = QPTrim$(fptxtNote.Text) + "/Prepay Refund"
  Else
    TransRec.Description = QPTrim$(fptxtNote.Text)
  End If
  TransRec.DiscAmt = 0
  TransRec.DiscXDate = 0
  TransRec.DMVBatch = 0
  TransRec.DMVSubmitted = ""
  TransRec.FromPrePay = "Y"
  TransRec.InternalPin = 0
  TransRec.LastTrans = TaxCust.LastTrans
  TaxCust.LastTrans = ThisRec
  TransRec.OperNum = OperNum
  TransRec.Padding = ""
  TransRec.PersPin = ""
  TransRec.Posted2GL = "N"
  TransRec.RealPin = ""
  TransRec.TaxYear = TaxMasterRec.RTaxYear
  TransRec.TransDate = Date2Num(fptxtDate.Text)
  TransRec.TranType = 12 'refund on prepay
  TransRec.Revenue.Collection = 0
  TransRec.Revenue.CollectionPd = 0
  TransRec.Revenue.Future1 = 0
  TransRec.Revenue.Future1Pd = 0
  TransRec.Revenue.Future2 = 0
  TransRec.Revenue.Future2Pd = 0
  TransRec.Revenue.Interest = 0
  TransRec.Revenue.InterestPd = 0
  TransRec.Revenue.LateList = 0
  TransRec.Revenue.LateListPd = 0
  TransRec.Revenue.pad = ""
  TransRec.Revenue.Penalty = 0
  TransRec.Revenue.PenaltyPd = 0
  TransRec.Revenue.PrePaidAmt = 0
  TransRec.Revenue.PrePaidBal = 0
  TransRec.Revenue.PrePaidUsed = RealRefund
  TaxCust.PrePayBal = 0
  TaxCust.PrePayTrans = ThisRec
  TransRec.Revenue.Principle1 = 0
  TransRec.Revenue.Principle1Pd = 0
  TransRec.Revenue.Principle2 = 0
  TransRec.Revenue.Principle2Pd = 0
  TransRec.Revenue.Principle3 = 0
  TransRec.Revenue.Principle3Pd = 0
  TransRec.Revenue.Principle4 = 0
  TransRec.Revenue.Principle4Pd = 0
  TransRec.Revenue.Principle5 = 0
  TransRec.Revenue.Principle5Pd = 0
  TransRec.Revenue.RevOpt1 = 0
  TransRec.Revenue.RevOpt1Pd = 0
  TransRec.Revenue.RevOpt2 = 0
  TransRec.Revenue.RevOpt2Pd = 0
  TransRec.Revenue.RevOpt3 = 0
  TransRec.Revenue.RevOpt3Pd = 0
  Put TCHandle, RefCustNum, TaxCust
  Put TTHandle, ThisRec, TransRec
  
  Close

  Return
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxRefundOnPrepay", "cmdPost_Click", Erl)
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
  If KeyCode = vbKeyReturn Then
    Call fpList1_Click
    Exit Sub
  End If
  
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
      Call cmdListAcctNum_Click
      KeyCode = 0
    Case vbKeyF6:
      SendKeys "%A"
      Call cmdListAlpha_Click
      KeyCode = 0
    Case vbKeyF4:
      SendKeys "%C"
      Call cmdCustData_Click
      KeyCode = 0
    Case vbKeyF2:
      SendKeys "%M"
      Call cmdInfo_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%P"
'      Call cmdPost_Click
      KeyCode = 0
    Case Else:
  End Select

End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Me.HelpContextID = hlpRefundFor
  Call LoadMe

End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      KillFile "C:\CPWork\prepayrefund.dat"
      GCustNum = 0
      MainLog ("CitiTaxes.exe terminated via menu bar on frmVATaxRefundOnPrepay.")
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
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim x As Long
  Dim NumOfTCRecs As Long
  Dim One As Integer
  Dim AHandle As Integer
  
  fptxtDate = Date
  fpList1.ListIndex = 0
  One = 1
  AHandle = FreeFile
  Open "C:\CPWork\prepayrefund.dat" For Output As AHandle
  Print #AHandle, One
  Close AHandle
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  For x = 1 To RefNumCnt
    Get TCHandle, AcctNumList(x), TaxCust
    fpList1.InsertRow = CStr(TaxCust.Acct) + Chr(9) + QPTrim$(TaxCust.CustName) + Chr(9) + Using$("$###,##0.00", Abs(NumPreBal(x)))
  Next x
  
  Close
End Sub

Private Sub fpList1_Click()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim PreBal As Double
  
  On Error GoTo ERRORSTUFF
  
  If fpList1.ListIndex = -1 Then
    Call TaxMsg(900, "Please make a selection from the list.")
    fpList1.ListIndex = 0
    Exit Sub
  End If
  
  fpList1.Col = 0
  fpList1.Row = fpList1.ListIndex
  RefCustNum = CLng(fpList1.ColText)
  GCustNum = RefCustNum
  If RefCustNum > 0 Then
    If CustHasMsg(RefCustNum) Then
      MsgAlertTimer.Enabled = True
    Else
      MsgAlertTimer.Enabled = False
      cmdInfo.FontSize = cmdCustData.FontSize
      cmdInfo.ForeColor = &H80000012
    End If
  End If
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  Get TCHandle, RefCustNum, TaxCust
  Close TCHandle
  
  fptxtName.Text = QPTrim$(TaxCust.CustName)
  fptxtAddress.Text = QPTrim$(TaxCust.Addr1)
  fptxtCity.Text = QPTrim$(TaxCust.City)
  fptxtState.Text = QPTrim$(TaxCust.State)
  fptxtZip.Text = QPTrim$(TaxCust.Zip)
  
  fpList1.Col = 2
  fpList1.Row = fpList1.ListIndex
  PreBal = CDbl(fpList1.ColText)
  fpCurrRefund = PreBal
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxRefundOnPrepay", "fpList1_Click", Erl)
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
Public Sub MsgAlertTimer_Timer()
  Static tog As Double
  Static TogState As Boolean
  If Me.Visible Then
    If BtnFnt# = 0 Then
      BtnFnt# = cmdInfo.FontSize
    End If
    If TogState Then
      tog = tog + 1
    Else
      tog = tog - 1
    End If
    Select Case tog
    Case 1
      cmdInfo.ForeColor = &H80000012
      cmdInfo.FontSize = BtnFnt
    Case 2
      cmdInfo.ForeColor = &H80000011
      cmdInfo.FontSize = BtnFnt - 0.7
    Case 3
      cmdInfo.ForeColor = &H80000011
      cmdInfo.FontSize = BtnFnt - 1.4
    Case 4
      cmdInfo.ForeColor = &H80000010
      cmdInfo.FontSize = BtnFnt - 2.1
    Case 5
      cmdInfo.ForeColor = &H80000010
      cmdInfo.FontSize = BtnFnt - 2.8
    Case 6
      cmdInfo.ForeColor = &H8000000F
      cmdInfo.FontSize = BtnFnt - 3.5
    Case 7
      cmdInfo.ForeColor = &H8000000F
      cmdInfo.FontSize = BtnFnt - 4.2
    Case 8
      cmdInfo.ForeColor = &H8000000E
      cmdInfo.FontSize = BtnFnt - 4.9
    Case 9
      cmdInfo.ForeColor = &H8000000E
      cmdInfo.FontSize = BtnFnt - 5.6
    End Select
    Select Case tog
    Case Is < 0, Is > 9
      TogState = Not TogState
    End Select
  End If
End Sub

Private Sub ClearData()
  MsgAlertTimer.Enabled = False
  cmdInfo.BackColor = &H8000000F
  cmdInfo.ForeColor = &H80000012
  cmdInfo.FontSize = cmdCustData.FontSize
  fptxtName.Text = ""
  fptxtAddress.Text = ""
  fptxtCity.Text = ""
  fptxtState.Text = ""
  fptxtZip.Text = ""
  fpCurrRefund = 0
  fptxtNote.Text = ""
  fptxtDate = Date
  GCustNum = 0
  RefCustNum = 0
End Sub
