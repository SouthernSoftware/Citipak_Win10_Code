VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmPaymentEntry 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payment Entry"
   ClientHeight    =   8865
   ClientLeft      =   3930
   ClientTop       =   2175
   ClientWidth     =   12210
   Icon            =   "frmPaymentEntry.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   12210
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboTenderType 
      Height          =   345
      Left            =   2910
      TabIndex        =   1
      Top             =   4125
      Width           =   2235
      _Version        =   196608
      _ExtentX        =   3942
      _ExtentY        =   609
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      Text            =   ""
      Columns         =   1
      Sorted          =   0
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   1
      ColumnWidthScale=   2
      RowHeight       =   -1
      WrapList        =   0   'False
      WrapWidth       =   0
      AutoSearch      =   2
      SearchMethod    =   2
      VirtualMode     =   0   'False
      VRowCount       =   0
      DataSync        =   3
      ThreeDInsideStyle=   0
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
      AutoSearchFillDelay=   100
      EditMarginLeft  =   1
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmPaymentEntry.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdCharge 
      Height          =   375
      Left            =   5460
      TabIndex        =   26
      Top             =   7635
      Width           =   1290
      _Version        =   131072
      _ExtentX        =   2275
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
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
      ButtonDesigner  =   "frmPaymentEntry.frx":0BF1
   End
   Begin fpBtnAtlLibCtl.fpBtn fpcmdCheck 
      Height          =   375
      Left            =   4110
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   7635
      Width           =   1260
      _Version        =   131072
      _ExtentX        =   2222
      _ExtentY        =   661
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
      ButtonDesigner  =   "frmPaymentEntry.frx":0DCE
   End
   Begin EditLib.fpLongInteger fpAcct 
      Height          =   330
      Left            =   1860
      TabIndex        =   0
      Top             =   1410
      Width           =   1875
      _Version        =   196608
      _ExtentX        =   3302
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
      AllowNull       =   -1  'True
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
      AutoBeep        =   0   'False
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
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   1
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
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   30
      Top             =   8508
      Width           =   12216
      _ExtentX        =   21537
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7144
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7144
            TextSave        =   "1:34 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7144
            TextSave        =   "10/28/2019"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin EditLib.fpDateTime txtPaymentDate 
      Height          =   330
      Left            =   10080
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   1440
      Width           =   1545
      _Version        =   196608
      _ExtentX        =   2730
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
      ControlType     =   1
      Text            =   "10/03/2001"
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
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      StartMonth      =   4
      ButtonAlign     =   0
      BoundDataType   =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fptxtDesc 
      Height          =   300
      Left            =   1464
      TabIndex        =   4
      Top             =   6816
      Width           =   3720
      _Version        =   196608
      _ExtentX        =   6562
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
      AutoCase        =   0
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
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   30
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
   Begin EditLib.fpCurrency fpTotOwed 
      Height          =   312
      Left            =   8208
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   7056
      Width           =   1836
      _Version        =   196608
      _ExtentX        =   3238
      _ExtentY        =   550
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
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   0
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
      AlignTextH      =   2
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
      ControlType     =   2
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   2
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   "$"
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "999999999.99"
      MinValue        =   "-999999999.99"
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
      IncDec          =   1
      BorderGrayAreaColor=   -2147483637
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpCurrency fpTotPaid 
      Height          =   312
      Left            =   10080
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   7056
      Width           =   1788
      _Version        =   196608
      _ExtentX        =   3154
      _ExtentY        =   550
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
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   0
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
      AlignTextH      =   2
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
      ControlType     =   2
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   2
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   "$"
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "999999999.99"
      MinValue        =   "-999999999.99"
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
      IncDec          =   1
      BorderGrayAreaColor=   -2147483637
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpDoubleSingle fpAmtOwed 
      Height          =   324
      Index           =   0
      Left            =   8208
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   2280
      Width           =   1836
      _Version        =   196608
      _ExtentX        =   3238
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
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
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   1
      ControlType     =   1
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
   Begin EditLib.fpDoubleSingle fpAmtPaid 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   324
      Index           =   0
      Left            =   10080
      TabIndex        =   5
      Top             =   2280
      Width           =   1788
      _Version        =   196608
      _ExtentX        =   3154
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   -1  'True
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   1
      HideSelection   =   0   'False
      InvalidColor    =   -2147483637
      InvalidOption   =   2
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
   Begin EditLib.fpDoubleSingle fpAmtPaid 
      Height          =   324
      Index           =   1
      Left            =   10080
      TabIndex        =   6
      Top             =   2592
      Width           =   1788
      _Version        =   196608
      _ExtentX        =   3154
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   -1  'True
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   1
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   2
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
   Begin EditLib.fpDoubleSingle fpAmtPaid 
      Height          =   324
      Index           =   2
      Left            =   10080
      TabIndex        =   7
      Top             =   2904
      Width           =   1788
      _Version        =   196608
      _ExtentX        =   3154
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   -1  'True
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   1
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   2
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
   Begin EditLib.fpDoubleSingle fpAmtPaid 
      Height          =   324
      Index           =   3
      Left            =   10080
      TabIndex        =   8
      Top             =   3216
      Width           =   1788
      _Version        =   196608
      _ExtentX        =   3154
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   -1  'True
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   1
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   2
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
   Begin EditLib.fpDoubleSingle fpAmtOwed 
      Height          =   324
      Index           =   1
      Left            =   8208
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   2592
      Width           =   1836
      _Version        =   196608
      _ExtentX        =   3238
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
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
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   1
      ControlType     =   1
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
   Begin EditLib.fpDoubleSingle fpAmtOwed 
      Height          =   324
      Index           =   2
      Left            =   8208
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   2904
      Width           =   1836
      _Version        =   196608
      _ExtentX        =   3238
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
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
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   1
      ControlType     =   1
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
   Begin EditLib.fpDoubleSingle fpAmtOwed 
      Height          =   324
      Index           =   3
      Left            =   8208
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   3216
      Width           =   1836
      _Version        =   196608
      _ExtentX        =   3238
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
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
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   1
      ControlType     =   1
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
   Begin EditLib.fpDoubleSingle fpAmtPaid 
      Height          =   324
      Index           =   4
      Left            =   10080
      TabIndex        =   9
      Top             =   3528
      Width           =   1788
      _Version        =   196608
      _ExtentX        =   3154
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   -1  'True
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   1
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   2
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
   Begin EditLib.fpDoubleSingle fpAmtPaid 
      Height          =   324
      Index           =   5
      Left            =   10080
      TabIndex        =   10
      Top             =   3840
      Width           =   1788
      _Version        =   196608
      _ExtentX        =   3154
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   -1  'True
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   1
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   2
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
   Begin EditLib.fpDoubleSingle fpAmtPaid 
      Height          =   324
      Index           =   6
      Left            =   10080
      TabIndex        =   11
      Top             =   4152
      Width           =   1788
      _Version        =   196608
      _ExtentX        =   3154
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   -1  'True
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   1
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   2
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
   Begin EditLib.fpDoubleSingle fpAmtPaid 
      Height          =   324
      Index           =   7
      Left            =   10080
      TabIndex        =   12
      Top             =   4464
      Width           =   1788
      _Version        =   196608
      _ExtentX        =   3154
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   -1  'True
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   1
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   2
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
   Begin EditLib.fpDoubleSingle fpAmtPaid 
      Height          =   324
      Index           =   8
      Left            =   10080
      TabIndex        =   13
      Top             =   4776
      Width           =   1788
      _Version        =   196608
      _ExtentX        =   3154
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   -1  'True
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   1
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   2
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
   Begin EditLib.fpDoubleSingle fpAmtPaid 
      Height          =   324
      Index           =   9
      Left            =   10080
      TabIndex        =   14
      Top             =   5088
      Width           =   1788
      _Version        =   196608
      _ExtentX        =   3154
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   -1  'True
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   1
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   2
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
   Begin EditLib.fpDoubleSingle fpAmtPaid 
      Height          =   324
      Index           =   10
      Left            =   10080
      TabIndex        =   15
      Top             =   5400
      Width           =   1788
      _Version        =   196608
      _ExtentX        =   3154
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   -1  'True
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   1
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   2
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
   Begin EditLib.fpDoubleSingle fpAmtPaid 
      Height          =   324
      Index           =   11
      Left            =   10080
      TabIndex        =   16
      Top             =   5712
      Width           =   1788
      _Version        =   196608
      _ExtentX        =   3154
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   -1  'True
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   1
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   2
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
   Begin EditLib.fpDoubleSingle fpAmtPaid 
      Height          =   324
      Index           =   12
      Left            =   10080
      TabIndex        =   18
      Top             =   6024
      Width           =   1788
      _Version        =   196608
      _ExtentX        =   3154
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   -1  'True
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   1
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   2
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
   Begin EditLib.fpDoubleSingle fpAmtPaid 
      Height          =   324
      Index           =   13
      Left            =   10080
      TabIndex        =   17
      Top             =   6336
      Width           =   1788
      _Version        =   196608
      _ExtentX        =   3154
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   -1  'True
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   1
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   2
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
   Begin EditLib.fpDoubleSingle fpAmtPaid 
      Height          =   324
      Index           =   14
      Left            =   10080
      TabIndex        =   19
      Top             =   6648
      Width           =   1788
      _Version        =   196608
      _ExtentX        =   3154
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   -1  'True
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   1
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   2
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
   Begin EditLib.fpDoubleSingle fpAmtOwed 
      Height          =   324
      Index           =   4
      Left            =   8208
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   3528
      Width           =   1836
      _Version        =   196608
      _ExtentX        =   3238
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
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
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   1
      ControlType     =   1
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
   Begin EditLib.fpDoubleSingle fpAmtOwed 
      Height          =   324
      Index           =   5
      Left            =   8208
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   3840
      Width           =   1836
      _Version        =   196608
      _ExtentX        =   3238
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
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
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   1
      ControlType     =   1
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
   Begin EditLib.fpDoubleSingle fpAmtOwed 
      Height          =   324
      Index           =   6
      Left            =   8208
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   4152
      Width           =   1836
      _Version        =   196608
      _ExtentX        =   3238
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
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
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   1
      ControlType     =   1
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
   Begin EditLib.fpDoubleSingle fpAmtOwed 
      Height          =   324
      Index           =   7
      Left            =   8208
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   4464
      Width           =   1836
      _Version        =   196608
      _ExtentX        =   3238
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
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
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   1
      ControlType     =   1
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
   Begin EditLib.fpDoubleSingle fpAmtOwed 
      Height          =   324
      Index           =   8
      Left            =   8208
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   4776
      Width           =   1836
      _Version        =   196608
      _ExtentX        =   3238
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
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
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   1
      ControlType     =   1
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
   Begin EditLib.fpDoubleSingle fpAmtOwed 
      Height          =   324
      Index           =   9
      Left            =   8208
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   5088
      Width           =   1836
      _Version        =   196608
      _ExtentX        =   3238
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
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
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   1
      ControlType     =   1
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
   Begin EditLib.fpDoubleSingle fpAmtOwed 
      Height          =   324
      Index           =   10
      Left            =   8208
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   5400
      Width           =   1836
      _Version        =   196608
      _ExtentX        =   3238
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
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
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   1
      ControlType     =   1
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
   Begin EditLib.fpDoubleSingle fpAmtOwed 
      Height          =   324
      Index           =   11
      Left            =   8208
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   5712
      Width           =   1836
      _Version        =   196608
      _ExtentX        =   3238
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
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
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   1
      ControlType     =   1
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
   Begin EditLib.fpDoubleSingle fpAmtOwed 
      Height          =   324
      Index           =   12
      Left            =   8208
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   6024
      Width           =   1836
      _Version        =   196608
      _ExtentX        =   3238
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
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
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   1
      ControlType     =   1
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
   Begin EditLib.fpDoubleSingle fpAmtOwed 
      Height          =   324
      Index           =   13
      Left            =   8208
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   6336
      Width           =   1836
      _Version        =   196608
      _ExtentX        =   3238
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
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
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   1
      ControlType     =   1
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
   Begin EditLib.fpDoubleSingle fpAmtOwed 
      Height          =   324
      Index           =   14
      Left            =   8208
      TabIndex        =   69
      TabStop         =   0   'False
      Top             =   6648
      Width           =   1836
      _Version        =   196608
      _ExtentX        =   3238
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
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
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   1
      ControlType     =   1
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
   Begin EditLib.fpCurrency fpChangeDue 
      Height          =   312
      Left            =   2904
      TabIndex        =   70
      TabStop         =   0   'False
      Top             =   5964
      Width           =   2244
      _Version        =   196608
      _ExtentX        =   3958
      _ExtentY        =   550
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
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   0
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
      AlignTextH      =   2
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
      ControlType     =   2
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   2
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   "$"
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "999999999.99"
      MinValue        =   "-999999999.99"
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
      IncDec          =   1
      BorderGrayAreaColor=   -2147483637
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpCurrency fpTotReceived 
      Height          =   312
      Left            =   2904
      TabIndex        =   71
      TabStop         =   0   'False
      Top             =   5388
      Width           =   2244
      _Version        =   196608
      _ExtentX        =   3958
      _ExtentY        =   550
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
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   0
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
      AlignTextH      =   2
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
      ControlType     =   2
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   2
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   "$"
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "999999999.99"
      MinValue        =   "-999999999.99"
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
      IncDec          =   1
      BorderGrayAreaColor=   -2147483637
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpCurrency fpTAmtOwed 
      Height          =   312
      Left            =   2904
      TabIndex        =   72
      TabStop         =   0   'False
      Top             =   3816
      Width           =   2244
      _Version        =   196608
      _ExtentX        =   3958
      _ExtentY        =   550
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
      ThreeDInsideStyle=   0
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   0
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
      ControlType     =   1
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   2
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   "$"
      DecimalPoint    =   "."
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
   Begin fpBtnAtlLibCtl.fpBtn fpCmdSave 
      Height          =   375
      Left            =   9390
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   7635
      Width           =   1230
      _Version        =   131072
      _ExtentX        =   2170
      _ExtentY        =   661
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
      ButtonDesigner  =   "frmPaymentEntry.frx":20A0
   End
   Begin fpBtnAtlLibCtl.fpBtn CmdExit 
      Height          =   375
      Left            =   10695
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   7635
      Width           =   1230
      _Version        =   131072
      _ExtentX        =   2170
      _ExtentY        =   661
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
      ButtonDesigner  =   "frmPaymentEntry.frx":227C
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdCash 
      Height          =   375
      Left            =   2880
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   7635
      Width           =   1140
      _Version        =   131072
      _ExtentX        =   2011
      _ExtentY        =   661
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
      ButtonDesigner  =   "frmPaymentEntry.frx":2458
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdInfo 
      Height          =   375
      Left            =   1605
      TabIndex        =   23
      Top             =   7635
      Width           =   1185
      _Version        =   131072
      _ExtentX        =   2090
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
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
      ButtonDesigner  =   "frmPaymentEntry.frx":3729
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdDist 
      Height          =   375
      Left            =   8100
      TabIndex        =   29
      Top             =   7635
      Width           =   1185
      _Version        =   131072
      _ExtentX        =   2090
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
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
      ButtonDesigner  =   "frmPaymentEntry.frx":3904
   End
   Begin EditLib.fpText fpCustRecNo 
      Height          =   324
      Left            =   504
      TabIndex        =   73
      TabStop         =   0   'False
      Top             =   648
      Visible         =   0   'False
      Width           =   1764
      _Version        =   196608
      _ExtentX        =   3111
      _ExtentY        =   572
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      NoSpecialKeys   =   3
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
      Text            =   "fpText1"
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
   Begin fpBtnAtlLibCtl.fpBtn fpcmdFind 
      Height          =   375
      Left            =   6840
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   7650
      Width           =   1185
      _Version        =   131072
      _ExtentX        =   2090
      _ExtentY        =   661
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
      ButtonDesigner  =   "frmPaymentEntry.frx":3ADF
   End
   Begin EditLib.fpText fpstatus 
      Height          =   300
      Left            =   432
      TabIndex        =   74
      Top             =   48
      Visible         =   0   'False
      Width           =   1140
      _Version        =   196608
      _ExtentX        =   2011
      _ExtentY        =   529
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
      ThreeDInsideStyle=   0
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
      AutoCase        =   0
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
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   1
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
   Begin EditLib.fpText fptaxexmpt 
      Height          =   300
      Left            =   432
      TabIndex        =   75
      Top             =   360
      Visible         =   0   'False
      Width           =   1884
      _Version        =   196608
      _ExtentX        =   3323
      _ExtentY        =   529
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
      ThreeDInsideStyle=   0
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
      AutoCase        =   0
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
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   1
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
   Begin EditLib.fpDoubleSingle fpChkAmt 
      Height          =   324
      Left            =   2904
      TabIndex        =   3
      Top             =   4800
      Width           =   2244
      _Version        =   196608
      _ExtentX        =   3958
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
      ThreeDInsideStyle=   0
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   0
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
      AutoBeep        =   -1  'True
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   1
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   2
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
   Begin EditLib.fpDoubleSingle fpCashAmt 
      Height          =   324
      Left            =   2904
      TabIndex        =   2
      Top             =   4464
      Width           =   2244
      _Version        =   196608
      _ExtentX        =   3958
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
      ThreeDInsideStyle=   0
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   0
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
      AutoBeep        =   -1  'True
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   1
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   2
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
   Begin fpBtnAtlLibCtl.fpBtn fpcmdDrawer 
      Height          =   375
      Left            =   285
      TabIndex        =   22
      Top             =   7635
      Width           =   1245
      _Version        =   131072
      _ExtentX        =   2196
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
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
      ButtonDesigner  =   "frmPaymentEntry.frx":4DB0
   End
   Begin VB.Label fpDeposit 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1368
      TabIndex        =   96
      Top             =   2952
      Width           =   1284
   End
   Begin VB.Label fptxtCity 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   1368
      TabIndex        =   95
      Top             =   2616
      Width           =   3924
   End
   Begin VB.Label fptxtAddress 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   1368
      TabIndex        =   94
      Top             =   2304
      Width           =   3924
   End
   Begin VB.Label fptxtName 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   1368
      TabIndex        =   93
      Top             =   1992
      Width           =   3924
   End
   Begin VB.Label fpRevSource 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Index           =   14
      Left            =   5472
      TabIndex        =   78
      Top             =   6648
      Width           =   2700
   End
   Begin VB.Label fpRevSource 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Index           =   13
      Left            =   5472
      TabIndex        =   79
      Top             =   6348
      Width           =   2700
   End
   Begin VB.Label fpRevSource 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Index           =   10
      Left            =   5472
      TabIndex        =   82
      Top             =   5400
      Width           =   2700
   End
   Begin VB.Label fpRevSource 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Index           =   9
      Left            =   5472
      TabIndex        =   83
      Top             =   5100
      Width           =   2700
   End
   Begin VB.Label fpRevSource 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Index           =   7
      Left            =   5472
      TabIndex        =   85
      Top             =   4476
      Width           =   2700
   End
   Begin VB.Label fpRevSource 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Index           =   5
      Left            =   5472
      TabIndex        =   87
      Top             =   3852
      Width           =   2700
   End
   Begin VB.Label fpRevSource 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Index           =   4
      Left            =   5472
      TabIndex        =   88
      Top             =   3552
      Width           =   2700
   End
   Begin VB.Label fpRevSource 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Index           =   1
      Left            =   5472
      TabIndex        =   91
      Top             =   2580
      Width           =   2700
   End
   Begin VB.Label fpRevSource 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Index           =   0
      Left            =   5472
      TabIndex        =   92
      Top             =   2280
      Width           =   2700
   End
   Begin VB.Label fpRevSource 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Index           =   2
      Left            =   5472
      TabIndex        =   90
      Top             =   2904
      Width           =   2700
   End
   Begin VB.Label fpRevSource 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Index           =   3
      Left            =   5472
      TabIndex        =   89
      Top             =   3228
      Width           =   2700
   End
   Begin VB.Label fpRevSource 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Index           =   6
      Left            =   5472
      TabIndex        =   86
      Top             =   4176
      Width           =   2700
   End
   Begin VB.Label fpRevSource 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Index           =   8
      Left            =   5472
      TabIndex        =   84
      Top             =   4800
      Width           =   2700
   End
   Begin VB.Label fpRevSource 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Index           =   11
      Left            =   5472
      TabIndex        =   81
      Top             =   5724
      Width           =   2700
   End
   Begin VB.Label fpRevSource 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Index           =   12
      Left            =   5472
      TabIndex        =   80
      Top             =   6048
      Width           =   2700
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Operator Name:"
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
      Height          =   312
      Left            =   4272
      TabIndex        =   77
      Top             =   1512
      Width           =   1824
   End
   Begin VB.Label lblOperName 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   6192
      TabIndex        =   76
      Top             =   1464
      Width           =   1860
   End
   Begin VB.Shape Shape3 
      Height          =   612
      Left            =   216
      Top             =   7464
      Width           =   11796
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H8000000E&
      BorderWidth     =   3
      Height          =   828
      Left            =   216
      Top             =   1032
      Width           =   11796
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000E&
      BorderWidth     =   3
      FillColor       =   &H8000000E&
      Height          =   5604
      Left            =   228
      Top             =   1848
      Width           =   11772
   End
   Begin VB.Line Line7 
      BorderWidth     =   3
      X1              =   5448
      X2              =   11880
      Y1              =   7008
      Y2              =   7008
   End
   Begin VB.Line Line6 
      X1              =   8184
      X2              =   8184
      Y1              =   1968
      Y2              =   7320
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "Revenue Desc"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5448
      TabIndex        =   54
      Top             =   1968
      Width           =   2724
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Entry"
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
      Left            =   4092
      TabIndex        =   53
      Top             =   516
      Width           =   4020
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000009&
      Height          =   456
      Left            =   2580
      Top             =   432
      Width           =   7020
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Operator Number:"
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
      Height          =   312
      Left            =   4272
      TabIndex        =   52
      Top             =   1176
      Width           =   1824
   End
   Begin VB.Label Label2b 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Account Number:"
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
      Height          =   372
      Index           =   1
      Left            =   360
      TabIndex        =   51
      Top             =   1128
      Width           =   2856
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
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
      Height          =   348
      Left            =   324
      TabIndex        =   50
      Top             =   1992
      Width           =   972
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Amount Owed:"
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
      Index           =   0
      Left            =   984
      TabIndex        =   49
      Top             =   3840
      Width           =   1728
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Date:"
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
      Height          =   372
      Index           =   1
      Left            =   8352
      TabIndex        =   48
      Top             =   1488
      Width           =   1584
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total Received:"
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
      Height          =   312
      Left            =   900
      TabIndex        =   47
      Top             =   5424
      Width           =   1812
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   216
      TabIndex        =   46
      Top             =   3372
      Width           =   5232
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cash Amount Paid:"
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
      Height          =   300
      Left            =   444
      TabIndex        =   45
      Top             =   4488
      Width           =   2268
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tender Type:"
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
      Height          =   300
      Left            =   1128
      TabIndex        =   44
      Top             =   4152
      Width           =   1584
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Dep Amt:"
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
      Height          =   348
      Left            =   96
      TabIndex        =   43
      Top             =   3000
      Width           =   1188
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
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
      Height          =   348
      Left            =   48
      TabIndex        =   42
      Top             =   2412
      Width           =   1248
   End
   Begin VB.Label lblchange 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Change Due:"
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
      Height          =   372
      Left            =   840
      TabIndex        =   41
      Top             =   6000
      Width           =   1872
   End
   Begin VB.Label Lbl11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Check/Charge Amt Paid:"
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
      Height          =   300
      Left            =   240
      TabIndex        =   40
      Top             =   4824
      Width           =   2472
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Description:"
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
      Height          =   372
      Left            =   168
      TabIndex        =   39
      Top             =   6840
      Width           =   1224
   End
   Begin VB.Line Line2 
      BorderWidth     =   3
      X1              =   2568
      X2              =   5268
      Y1              =   5232
      Y2              =   5232
   End
   Begin VB.Line Line3 
      X1              =   5436
      X2              =   5436
      Y1              =   1848
      Y2              =   7440
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Caption         =   "Amount Owed"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   312
      Left            =   8208
      TabIndex        =   38
      Top             =   1968
      Width           =   1836
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Caption         =   "Amount Paid"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   10080
      TabIndex        =   37
      Top             =   1968
      Width           =   1788
   End
   Begin VB.Line Line4 
      X1              =   10056
      X2              =   10056
      Y1              =   1944
      Y2              =   7344
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Source:"
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
      Height          =   312
      Left            =   8280
      TabIndex        =   36
      Top             =   1152
      Width           =   1656
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Totals:"
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
      Height          =   324
      Left            =   7104
      TabIndex        =   35
      Top             =   7104
      Width           =   900
   End
   Begin VB.Label lblOperator 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   6192
      TabIndex        =   34
      Top             =   1128
      Width           =   732
   End
   Begin VB.Label lblSource 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Left            =   10080
      TabIndex        =   33
      Top             =   1128
      Width           =   1560
   End
   Begin VB.Shape Shape6 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   600
      Left            =   2592
      Top             =   312
      Width           =   7020
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuPrnScn 
         Caption         =   "Prin&t Screen"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmPaymentEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim CashFlag As Boolean, uselook As Boolean, CustAcct As Long
Dim EditFlag As Boolean, TempAmtRecv As Double, Answer As Integer
Dim ChkOKFlag As Boolean, BeenDone As Boolean, PayListCnt As Long
Dim DistArray() As DistArrayType
Dim PayList() As PayListType
Dim fromform As Form, toform As Form, codeopt As Integer, noreset As Boolean
Dim Oper As String, PayListRec As Long, RecpPort As String, DefPayDate As String
Dim RevText$(1 To MaxRevsCnt)
Dim RctValidate As Boolean, Caldwell As Boolean, Biscoe As Boolean
Dim Lilesville As Boolean
Public Sub Wheretogo(xfrm As Form, tfrm As Form, Optional opt As Integer, Optional DDate As String)
  Set fromform = xfrm
  Set toform = tfrm
  If opt <> 0 Then
    codeopt = opt
  Else
    codeopt = 0
  End If
  uselook = True
  If DDate <> "" Then
    DefPayDate = DDate
  End If
End Sub
Private Sub Form_Activate()
  If Val(fpCustRecNo) > 0 And Not BeenDone Then
    BeenDone = True
    fpAcct = fpCustRecNo
    GetCustinfo
    DoEvents
  End If
End Sub

Private Sub cmdExit_Click()
  ChkEmptyAcct
  noreset = True
  Chk4Change
  If Answer = 1 Then
    Exit Sub
  ElseIf Answer = 2 Then
    CheckInfo
    If ChkOKFlag Then
      fpCmdSave_Click
    Else
      Exit Sub
    End If
  End If
  CustAcct = 0
  fpCustRecNo = 0
  BeenDone = False
  If codeopt = 1 Then
    ActivateControls frmCustEditLookUP
  ElseIf codeopt = 2 Then
    ActivateControls frmDisplayList
  End If
  If codeopt = 0 Then
    Load frmUBPaymentMenu
    DoEvents
    frmUBPaymentMenu.Show
  End If
  Erase PayList, RevText$
  UBLog "OUT: UTIL Payment" + " Oper:" + Oper$
  Unload Me
  DoEvents
End Sub

Private Sub fpAcct_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fpAcct_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyReturn, vbKeyDown, vbKeyUp, vbKeyTab, vbKeyC
      KeyCode = 0
    If Len(fpAcct) > 0 Then
      If fpcboTenderType.Enabled = True Then
        fpcboTenderType.SetFocus
      End If
    End If
  End Select
End Sub
Private Sub ChkEmptyAcct()
  If Len(fpAcct) <= 0 Then
    ClearScn
  End If
End Sub

Private Sub fpAmtPaid_LostFocus(Index As Integer)
  CalcBALFlds
End Sub

Private Sub fpAmtPaid_ChangeMode(Index As Integer, EditMode As Integer)
  EditMode = True
End Sub

Private Sub fpAmtPaid_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  Dim x As Integer
  If KeyCode = vbKeyReturn Or KeyCode = vbKeyRight Or KeyCode = vbKeyDown Then
    If Index < MaxRevsCnt Then
     For x = Index To (MaxRevsCnt - 1)
      If fpAmtPaid(x + 1).Enabled Then
        fpAmtPaid(x + 1).SetFocus
        Exit For
      Else
        fpCmdSave.SetFocus
        Exit For
      End If
     Next
    End If
  ElseIf KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Then
    If Index > 0 Then
     For x = Index To (MaxRevsCnt - 1)
      If fpAmtPaid(x - 1).Enabled Then
        fpAmtPaid(x - 1).SetFocus
        Exit For
      Else
        fptxtDesc.SetFocus
      End If
     Next
    End If
  End If

End Sub

Private Sub fpCashAmt_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fpCashAmt_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    If fpChkAmt.Enabled Then
      fpChkAmt.SetFocus
    Else
      fptxtDesc.SetFocus
    End If
  End If
End Sub


Private Sub fpChkAmt_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fpChkAmt_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fptxtDesc.SetFocus
  End If
End Sub

Private Sub fpCmdDist_Click()
If Len(fpAcct) > 0 Then
  If fpTotReceived > 0 Then
    TempAmtRecv = fpTotReceived
    Autodist
  End If
End If
End Sub
Private Sub Chk4Change()
  Dim cntout As Integer, cnt As Integer
  Dim NumofRevs As Integer, RevCnt As Integer, ListFile As Integer
  
  Dim PayFileName As String, UBPayRecLen As Integer
  ReDim UBPaymentRec(1) As UBPaymentRecType

  UBPayRecLen = Len(UBPaymentRec(1))
  If Len(fpAcct) > 0 Then
  NumofRevs = MaxRevsCnt
  cntout = 0
  Answer = 0
  If EditFlag = True Then
    PayFileName$ = UBPath$ + "UBPAY" + Oper$ + ".DAT"
    ListFile = FreeFile
    Open PayFileName$ For Random Shared As ListFile Len = UBPayRecLen
    Get ListFile, PayListRec&, UBPaymentRec(1)
    Close ListFile
    If fpTAmtOwed <> UBPaymentRec(1).AMTOWED Then cntout = cntout + 1
    If txtPaymentDate <> Num2Date(UBPaymentRec(1).PAYDATE) Then cntout = cntout + 1
    If QPTrim$(fptxtDesc) <> QPTrim$(UBPaymentRec(1).Desc) Then cntout = cntout + 1
    For cnt = 1 To NumofRevs
      If fpAmtOwed(cnt - 1) <> UBPaymentRec(1).PaidOwed(cnt).AMTOWE1 Then cntout = cntout + 1
      If fpAmtPaid(cnt - 1) <> UBPaymentRec(1).PaidOwed(cnt).AMTPD1 Then cntout = cntout + 1
    Next
    Select Case QPTrim(UBPaymentRec(1).TENDERTY)
      Case "Cash":
        If fpcboTenderType.ListIndex <> 0 Then cntout = cntout + 1
      Case "Check":
        If fpcboTenderType.ListIndex <> 1 Then cntout = cntout + 1
      Case "Cash & Check":
        If fpcboTenderType.ListIndex <> 2 Then cntout = cntout + 1
      Case "Charge":
        If fpcboTenderType.ListIndex <> 3 Then cntout = cntout + 1
    End Select
    If fpCashAmt <> UBPaymentRec(1).CASHAMT Then cntout = cntout + 1
    If fpChkAmt <> UBPaymentRec(1).CHKAMT Then cntout = cntout + 1
    If fpTotReceived <> UBPaymentRec(1).AMTRECD Then cntout = cntout + 1
    If fpChangeDue <> UBPaymentRec(1).Change Then cntout = cntout + 1
  Else
    
    If fpTotReceived <> 0 Or fpTotPaid <> 0 Then cntout = cntout + 1
  End If
  If cntout > 0 Then
    frmChangedWarning.Show vbModal, Me
    Select Case SaveFlag
    Case False
      Answer = 3
    Case True
      Answer = 2
    Case 1
      Answer = 1
    End Select
  Else
    Answer = 0
  End If
  End If
End Sub
Private Sub Chk4OKforNew()
  Dim FntSize As Integer
  Dim cntout As Integer, cnt As Integer
  Dim NumofRevs As Integer, RevCnt As Integer, ListFile As Integer
  
  Dim PayFileName As String, UBPayRecLen As Integer
  ReDim UBPaymentRec(1) As UBPaymentRecType

  UBPayRecLen = Len(UBPaymentRec(1))
  If Len(fpAcct) > 0 Then
  NumofRevs = MaxRevsCnt
  cntout = 0
  Answer = 0
  If EditFlag = True Then
    PayFileName$ = UBPath$ + "UBPAY" + Oper$ + ".DAT"
    ListFile = FreeFile
    Open PayFileName$ For Random Shared As ListFile Len = UBPayRecLen
    Get ListFile, PayListRec&, UBPaymentRec(1)
    Close ListFile
    If fpTAmtOwed <> UBPaymentRec(1).AMTOWED Then cntout = cntout + 1
    If txtPaymentDate <> Num2Date(UBPaymentRec(1).PAYDATE) Then cntout = cntout + 1
    If QPTrim$(fptxtDesc) <> QPTrim$(UBPaymentRec(1).Desc) Then cntout = cntout + 1
    For cnt = 1 To NumofRevs
      If fpAmtOwed(cnt - 1) <> UBPaymentRec(1).PaidOwed(cnt).AMTOWE1 Then cntout = cntout + 1
      If fpAmtPaid(cnt - 1) <> UBPaymentRec(1).PaidOwed(cnt).AMTPD1 Then cntout = cntout + 1
    Next
    Select Case QPTrim(UBPaymentRec(1).TENDERTY)
      Case "Cash":
        If fpcboTenderType.ListIndex <> 0 Then cntout = cntout + 1
      Case "Check":
        If fpcboTenderType.ListIndex <> 1 Then cntout = cntout + 1
      Case "Cash & Check":
        If fpcboTenderType.ListIndex <> 2 Then cntout = cntout + 1
      Case "Charge":
        If fpcboTenderType.ListIndex <> 3 Then cntout = cntout + 1
    End Select
    If fpCashAmt <> UBPaymentRec(1).CASHAMT Then cntout = cntout + 1
    If fpChkAmt <> UBPaymentRec(1).CHKAMT Then cntout = cntout + 1
    If fpTotReceived <> UBPaymentRec(1).AMTRECD Then cntout = cntout + 1
    If fpChangeDue <> UBPaymentRec(1).Change Then cntout = cntout + 1
  Else
    
    If fpTotReceived <> 0 Or fpTotPaid <> 0 Then cntout = cntout + 1
  End If
  If cntout > 0 Then
    ReDim MsgText(0 To 5) As String
    FntSize = frmMsgDialog.Label(1).FontSize
    frmMsgDialog.Label(1).FontSize = (FntSize + 2)
    frmMsgDialog.Label(2).FontSize = (FntSize + 2)
    frmMsgDialog.Label(3).FontSize = (FntSize + 2)
    MsgText(0) = "WARNING:Payment In Progress"
    MsgText(1) = ""
    MsgText(2) = "Do You Want to Abandon this Payment?"
    MsgText(3) = "Ok to Abandon,"
    MsgText(4) = "Cancel to Remain on Current Payment."
    MsgText(5) = ""
    If GetOKorNot(MsgText()) Then
     UBLog "USER WANTS TO Abandon"
     Answer = 2
    Else
     UBLog "USER Canceled"
     Answer = 1
    End If
  Else
    Answer = 0
  End If
  End If
End Sub

Private Sub fpcmdDrawer_Click()
  Dim Port As String, PortFile As Integer ', DPName As String, DefPrinter As String
  On Local Error Resume Next
  If RecpDef = 99 Then Exit Sub
  'RecPort = GetDEFPort%
  Port$ = QPTrim$(RecpPort)
   
   'DPName = QPTrim(Printers(2).DeviceName) 'RecpPort).DeviceName)
   ' DefPrinter = QPTrim(Printers(2).Port) 'RecpPort).Port)
  'Added this to allow for winxp network printer port names of ne00:, etc.
  'the device name worked so use that instead, but only for network printers.
'    If InStr(1, DPName, "\\", vbTextCompare) Then
'      Port$ = DPName
'      'frmViewPrint.PrintWSet DPName, Copies
'    Else
'      Port$ = DefPrinter
'      'frmViewPrint.PrintWSet DefPrinter, Copies
'    End If
'    If vbKeyDown = vbKeyEscape Then
'      Printer.KillDoc
'    End If

  
  UBLog "Oper: " + Oper$ + "Util Pay-Open Drawer"
  PortFile = FreeFile
  Open Port$ For Output As #PortFile
  If Biscoe = True Then
    Print #PortFile, Chr$(7)
'  ElseIf Lilesville = True Then
'    Print #PortFile, Chr$(27); "p"; Chr$(112); Chr$(0); Chr$(25); Chr$(250)
  Else
    Print #PortFile, Chr$(27); "p"; Chr$(0) '; Chr$(25) '; Chr$(250)
  End If
  Close PortFile
End Sub

Private Sub fpcmdFind_Click()
  Chk4OKforNew
  If Answer = 1 Then
    Exit Sub
  ElseIf Answer = 2 Then
    'continue on
  End If
  ClearScn
  frmCustEditLookUP.Caption = "Payment Customer Find"
  frmCustEditLookUP.Label1.Caption = "Payment Customer Find"
  frmCustEditLookUP.Wheretogo frmPaymentEntry, frmPaymentEntry
  Unload Me
  DoEvents
  frmCustEditLookUP.Show
  DoEvents
End Sub

Private Sub fpCmdInfo_Click()
If Len(fpAcct) > 0 Then
  If Len(fpCustRecNo) > 0 Then
    'DeActivateControls Me
    frmInfo.Label1 = "Loading. . ."
    frmInfo.Show
    DoEvents
    'here
    frmRptCustInq.fpCustRecNo = Me.fpCustRecNo
    'frmRptCustInq.Wheretogo frmPaymentEntry, frmRptCustInq, 0
    'Load frmRptCustInq
    frmRptCustInq.Show
    DoEvents
    Unload frmInfo
  End If
End If
End Sub
Private Sub fpCmdSave_Click()
  ChkEmptyAcct
  DoEvents
  If Len(fpAcct) <= 0 Then
    MsgBox "Invalid Account Information.", vbOKOnly, "Invalid Entry"
    Exit Sub
  End If
  CalcBALFlds
  CheckInfo
  If ChkOKFlag Then
    'DeActivateControls Me
    If fpcboTenderType.ListIndex = 1 Or fpcboTenderType.ListIndex = 2 Then
      frmPrintReceipt.setvallist = 1
    Else
      frmPrintReceipt.setvallist = 0
    End If

    frmPrintReceipt.Show 1
    If SavePay = True Then
      SaveTransaction
    
      If PrnRecp = True Or PrnVali = True Then
        PrintReceipt
      End If
    
      MsgBox "Transaction Complete.", vbOKOnly, "Complete"
      
      ClearScn
      
    End If
    
'      CustAcct = 0
'      fpCustRecNo = 0
'      BeenDone = False
'      If codeopt = 1 Then
'        ActivateControls frmCustEditLookUP
'      ElseIf codeopt = 2 Then
'        ActivateControls frmDisplayList
'      End If
'      If codeopt = 0 Then
'        Load frmUBPaymentMenu
'        DoEvents
'        frmUBPaymentMenu.Show
'      End If
'
'      UBLog "OUT: UTIL Payment" + " Oper:" + Oper$
'      Unload Me
'      DoEvents
'    ActivateControls Me
  End If

End Sub
'Private Sub txtDate_KeyDown(KeyCode As Integer, Shift As Integer)
'  If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Or KeyCode = vbKeyRight Then
'    KeyCode = 0
'    fpAmount(0).SetFocus
'  ElseIf KeyCode = vbKeyUp Or KeyCode = vbKeyLeft Then
'    fpCmdSave.SetFocus
'  End If
'End Sub

Private Sub fpcmdCash_Click()
If Len(fpAcct) > 0 Then
  fpcboTenderType.ListIndex = 0
  fpChkAmt.Enabled = False
  fpCashAmt.Enabled = True
  fpChkAmt = 0
  fpCashAmt = fpTAmtOwed.DoubleValue
  fpTotReceived = Round#(fpCashAmt.DoubleValue + fpChkAmt.DoubleValue)
  If fpTotReceived > 0 Then
    TempAmtRecv = fpTotReceived
    Autodist
  End If
  fptxtDesc.SetFocus
End If
End Sub

Private Sub fpcmdCheck_Click()
If Len(fpAcct) > 0 Then
  fpcboTenderType.ListIndex = 1
  fpCashAmt.Enabled = False
  fpChkAmt.Enabled = True
  fpCashAmt = 0
  fpChkAmt = fpTAmtOwed.DoubleValue
  fpTotReceived = Round#(fpCashAmt.DoubleValue + fpChkAmt.DoubleValue)
  If fpTotReceived > 0 Then
    TempAmtRecv = fpTotReceived
    Autodist
  End If
  fptxtDesc.SetFocus
End If
End Sub
Private Sub fpCmdCharge_Click()
If Len(fpAcct) > 0 Then
  fpcboTenderType.ListIndex = 3
  fpCashAmt.Enabled = False
  fpChkAmt.Enabled = True
  fpCashAmt = 0
  fpChkAmt = fpTAmtOwed.DoubleValue
  fpTotReceived = Round#(fpCashAmt.DoubleValue + fpChkAmt.DoubleValue)
  If fpTotReceived > 0 Then
    TempAmtRecv = fpTotReceived
    Autodist
  End If
  fpChangeDue.Enabled = False
  fptxtDesc.SetFocus
End If
End Sub

Private Sub fpCashAmt_LostFocus()
fpTotReceived = Round#(fpCashAmt.DoubleValue + fpChkAmt.DoubleValue)
If fpTotReceived > 0 Then
  If fpcboTenderType.ListIndex <> 3 Then
    fpChangeDue = Round#(fpTotReceived.DoubleValue - fpTotPaid.DoubleValue)
  End If
End If
End Sub

Private Sub fpChkAmt_LostFocus()
fpTotReceived = Round#(fpCashAmt.DoubleValue + fpChkAmt.DoubleValue)
If fpTotReceived.DoubleValue > 0 Then
  If fpcboTenderType.ListIndex <> 3 Then
    fpChangeDue = Round#(fpTotReceived.DoubleValue - fpTotPaid.DoubleValue)
  End If
End If
End Sub
Private Sub fpcboTenderType_DropDown()
  ClrAmts
End Sub

Private Sub fpcboTenderType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboTenderType.ListDown = True
    'ClrAmts
    KeyCode = 0
  End If
  If KeyCode = vbKeyDelete Then
    fpcboTenderType.ListIndex = -1
    fpcboTenderType.Action = ActionClearSearchBuffer
    'ClrAmts
  End If
  If fpcboTenderType.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      If fpCashAmt.Enabled = True Then
        fpCashAmt.SetFocus
      Else
        fpChkAmt.SetFocus
      End If
        KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpAcct.SetFocus
        KeyCode = 0
      End If
    End If
  End If
  
  DoEvents
End Sub
Private Sub ClrAmts()
  Dim cnt As Integer
  fpCashAmt = 0
  fpChkAmt = 0
  fpChangeDue.Enabled = True
  fpChangeDue = 0
  For cnt = 1 To 15
    fpAmtPaid(cnt - 1) = 0
  Next
  fpTotPaid = 0
  fpTotReceived = 0
End Sub
Private Sub fpcboTenderType_SelChange(ItemIndex As Long)
  If BeenDone Then
    fixamts
 End If
End Sub

Private Sub fixamts()
  
  fpcboTenderType.Action = ActionClearSearchBuffer
  If noreset = False Then
    If fpcboTenderType.ListIndex = 0 Then
      fpCashAmt.Enabled = True
      fpChkAmt = 0
      fpChkAmt.Enabled = False
      fpChangeDue.Enabled = True
      'ClrAmts
     ' fpCashAmt.SetFocus
    ElseIf fpcboTenderType.ListIndex = 1 Then
      fpCashAmt.Enabled = False
      fpCashAmt = 0
      fpChkAmt.Enabled = True
      fpChangeDue.Enabled = True
      'ClrAmts
     ' fpChkAmt.SetFocus
    ElseIf fpcboTenderType.ListIndex = 2 Then
      fpCashAmt.Enabled = True
      fpChkAmt.Enabled = True
      fpChangeDue.Enabled = True
     ' ClrAmts
     'fpCashAmt.SetFocus
    ElseIf fpcboTenderType.ListIndex = 3 Then
      fpCashAmt.Enabled = False
      fpCashAmt = 0
      fpChkAmt.Enabled = True
      fpChangeDue = 0
      fpChangeDue.Enabled = False
     ' ClrAmts
      'fpChkAmt.SetFocus
'    ElseIf fpcboTenderType.ListIndex = -1 Then
'      MsgBox "You Must Select A Tender Type.", vbOKOnly, "Invalid Selection"
'      fpcboTenderType.SetFocus
    End If
  End If
  fpTotReceived = Round#(fpCashAmt.DoubleValue + fpChkAmt.DoubleValue)
  If fpTotReceived > 0 Then
    If fpcboTenderType.ListIndex <> 3 Then
      fpChangeDue = Round#(fpTotReceived.DoubleValue - fpTotPaid.DoubleValue)
    End If
  End If
  DoEvents
  noreset = False
End Sub
Private Sub fpAcct_LostFocus()
'Dim Acct As Long
'    Acct = fpAcct
'    If Acct > 0 Then
'      If Acct > GetTaxCustCnt Then
'        MsgBox "Bad Account Number.", vbOKOnly, "Invalid Account"
'        fplngAcct.SetFocus
'        Exit Sub
'      ElseIf IsCustDeleted(Acct) Then
'        MsgBox "Deleted Account.", vbOKOnly, "Deleted Account"
'        fplngAcct.SetFocus
'        Exit Sub
'      Else
'       'If DoesCustOwe(Acct) Then
'          Cust2Screen (Acct)
'       ' Else
'       '   MsgBox "This Customer Does Not Owe A Balance.", vbOKOnly, "No Balance"
'      End If
'    Else
'      MsgBox "Bad Account Number.", vbOKOnly, "Invalid Account"
'      fplngAcct.SetFocus
'      Exit Sub
'    End If
  fpCustRecNo = fpAcct
  'If Val(fpCustRecNo) > 0 Then
    GetCustinfo
  'Else
    'MsgBox "NO", vbOKOnly
  '  fpAcct.SetFocus
  'End If
End Sub

Private Sub fptxtDesc_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fptxtDesc_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpAmtPaid(0).SetFocus
  End If
End Sub


Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        UBLog "Closed via PaymentEntry by " + PWUser$ + " operator-" + Oper$
        CitiTerminate
      End If
    End If
  End If
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

  Select Case KeyCode
'    Case vbKeyDown, vbKeyReturn:
'      SendKeys "{Tab}"
'      KeyCode = 0
'    Case vbKeyUp:
'      SendKeys "+{Tab}"
'      KeyCode = 0
    Case vbKeyEscape:
      KeyCode = 0
      DoEvents
      If cmdExit.Enabled Then
        Call cmdExit_Click
      End If
    Case vbKeyF2:
      KeyCode = 0
      DoEvents
      fpcmdDrawer_Click
    Case vbKeyF4:
      KeyCode = 0
      DoEvents
      If fpCmdInfo.Enabled Then
        Call fpCmdInfo_Click
      End If
    Case vbKeyF5:
      KeyCode = 0
      DoEvents
      If fpCmdCash.Enabled Then
        Call fpcmdCash_Click
      End If
    Case vbKeyF6:
      KeyCode = 0
      DoEvents
      If fpcmdCheck.Enabled Then
        Call fpcmdCheck_Click
      End If
    Case vbKeyF7:
      KeyCode = 0
      DoEvents
      If fpcmdFind.Enabled Then
        Call fpcmdFind_Click
      End If
    Case vbKeyF8:
      KeyCode = 0
      DoEvents
      If fpCmdCharge.Enabled Then
        Call fpCmdCharge_Click
      End If
    Case vbKeyF9:
      KeyCode = 0
      DoEvents
      If fpCmdDist.Enabled Then
        Call fpCmdDist_Click
      End If
    Case vbKeyF10:
      KeyCode = 0
      DoEvents
      If fpCmdSave.Enabled Then
        Call fpCmdSave_Click
      End If
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
  If InStr(TOWNNAME$, "CALDWELL") > 0 Then
    Caldwell = True
  Else
    Caldwell = False
  End If
  If InStr(TOWNNAME$, "BISCOE") > 0 Then
    Biscoe = False
  Else
    Biscoe = False
  End If
If InStr(TOWNNAME$, "LILESVILLE") > 0 Then
    Lilesville = True
  Else
    Lilesville = False
  End If
  txtPaymentDate.Text = DefPayDate
  Me.HelpContextID = hlpPaymentTransaction
  noreset = False
  fpcboTenderType.AddItem "Cash"
  fpcboTenderType.AddItem "Check"
  fpcboTenderType.AddItem "Cash & Check"
  fpcboTenderType.AddItem "Charge"
  LoadRevs
  lblOperator = OPERNUM
  lblOperName.Caption = PWUser
  lblSource.Caption = "Utility"
  Oper$ = QPTrim(lblOperator.Caption)
  UBLog " IN Oper " + Oper$ + ": UTIL Payment"
  LoadPayList
  GetRcpInfo
End Sub
Private Sub GetRcpInfo()
  Dim RP As Integer, lenRP As Integer, RP1 As Integer
  Dim RcptPrnFile As ReceiptPRNType
  RP1 = FreeFile
  lenRP = Len(RcptPrnFile)
  If Exist(RcptFileName$) Then
    Open RcptFileName$ For Random Shared As RP1 Len = lenRP
    Get RP1, 1, RcptPrnFile
    RecpPort = QPTrim(RcptPrnFile.RcpPort)
    If RcptPrnFile.PrnDefYN = 0 Then
      RecpDef = 0
    Else
      On Local Error GoTo nofound
      RP = FreeFile
      Open RecpPort For Output As RP
      Close RP
      RecpDef = 1
    End If
    If RcptPrnFile.CtlDefYN = 0 Then
      CntrlDef = 0
    Else
      CntrlDef = 1
    End If
    If RcptPrnFile.RValidate = 1 Then
      ValiDef = 1
      RctValidate = True
      GetUBBankINfo
    Else
      ValiDef = 0
      RctValidate = False
    End If
  Close RP1
  Else
    ValiDef = 0
    RecpDef = 99
  End If
Exit Sub
nofound:
  ValiDef = 0
  RecpDef = 99
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
  ' Me.Visible = False
    Temp_Class.ResizeControls Me
  '  Me.Visible = True
  '  Me.SetFocus
  End If
  DoEvents
'  If Me.Visible Then
'    Temp_Class.ResizeControls Me
'    DoEvents
'  End If
End Sub
Private Sub LoadRevs()
  Dim NumofRevs As Integer, UBSetupLen As Integer, RevCnt As Integer
  Dim InvRev As Integer, OutOfOrder As Boolean, x As Integer
  Dim tmp As DistArrayType
  NumofRevs = MaxRevsCnt
  ReDim UBSetUpRec(1) As UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen
  'ReDim RevText$(1 To MaxRevsCnt)
  ReDim Preserve DistArray(1 To NumofRevs) As DistArrayType
 ' RecpPort = Val(UBSetUpRec(1).RecpPort)
'  If RecpPort < 1 Or RecpPort > 2 Then
'    RecpPort = 1
'  End If

  For RevCnt = 1 To MaxRevsCnt
    RevText$(RevCnt) = Left$(QPTrim$(UBSetUpRec(1).Revenues(RevCnt).RevName), 14)
    DistArray(RevCnt).DistOrder = UBSetUpRec(1).Revenues(RevCnt).DistOr
    DistArray(RevCnt).DistCnt = RevCnt
    If Len(RevText$(RevCnt)) = 0 Then
      NumofRevs = RevCnt - 1
      Exit For
    End If
  Next

  Do
    OutOfOrder = False          'assume it's sorted
    For x = 1 To NumofRevs - 1
      If DistArray(x).DistOrder > DistArray(x + 1).DistOrder Then
        'SWAP DistArray(x), DistArray(x + 1)     'if we had to swap
        tmp = DistArray(x)
        DistArray(x) = DistArray(x + 1)
        DistArray(x + 1) = tmp
        OutOfOrder = True       'we're not done yet
      End If
    Next
  Loop While OutOfOrder

  For RevCnt = 1 To MaxRevsCnt
    RevText$(RevCnt) = Left$(QPTrim$(UBSetUpRec(1).Revenues(RevCnt).RevName), 14)
    If Len(RevText$(RevCnt)) = 0 Then
      NumofRevs = RevCnt - 1
      Exit For
    End If
  Next

'  If NumOfRevs < MaxRevsCnt Then
'    ReDim Preserve RevText$(1 To NumOfRevs)
'  End If

  For RevCnt = 1 To NumofRevs
    fpRevSource(RevCnt - 1).Caption = RevText$(RevCnt)
  Next
  For InvRev = NumofRevs To 14
    fpRevSource(InvRev).Enabled = False
    fpRevSource(InvRev).Visible = False
    fpAmtPaid(InvRev).Enabled = False
    fpAmtPaid(InvRev).Visible = False
    fpAmtOwed(InvRev).Enabled = False
    fpAmtOwed(InvRev).Visible = False
  Next

End Sub
Private Sub GetCustinfo()
  Dim UBCustRecLen As Integer, NumOfCustRecs As Long
  Dim CustFile As Integer, cnt As Integer, TotalBalance As Double
  Dim NumofRevs As Integer, RevCnt As Integer, ListFile As Integer
  Dim PayFileName As String, UBPayRecLen As Integer
  ReDim UBCustRec(1) As NewUBCustRecType
  ReDim UBPaymentRec(1) As UBPaymentRecType

  UBPayRecLen = Len(UBPaymentRec(1))

  UBCustRecLen = Len(UBCustRec(1))
  NumofRevs = MaxRevsCnt
  CashFlag = False
  If uselook = True Then
    Unload frmCustEditLookUP
    Unload frmDisplayList
    uselook = False
  End If
  If fpCustRecNo <> "" Then
    CustAcct = fpCustRecNo
  Else
    'MsgBox "You Must Enter An Account Number.", vbOKOnly, "Invalid Account"
    'tried this to fix hoglet's complaint 10-7-08
    'If Len(QPTrim(fpAcct.Text)) = 0 Then
      fpAcct.SetFocus
    'End If
    Exit Sub
  End If
  NumOfCustRecs& = FileSize(UBPath$ + "UBCUST.DAT") \ UBCustRecLen
  If CustAcct& > NumOfCustRecs& Or CustAcct& <= 0 Then
    UBLog "ERROR: Invalid Account:" + Str$(CustAcct&) + " Oper:" + Oper$
    CustAcct& = 0
    'LabelDel.Visible = True
    GoTo SkiptoHere
  End If
  
  If IsDeleted(CustAcct&) Then
    UBLog "ERROR: Deleted Account:" + Str$(CustAcct&) + " Oper:" + Oper$
    CustAcct& = 0
    'LabelDel.Caption = "Deleted Account!"
    'LabelDel.Visible = True
    GoTo SkiptoHere
  End If
  
  CheckPayList

 ' GoSub ClearForm
  CustFile = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As CustFile Len = UBCustRecLen
  UBLog PWUser + (" Open custfile in payentry getcustinfo - " + Str(CustAcct&) + " with - numcusts " + Str(NumOfCustRecs&))
  Get CustFile, CustAcct&, UBCustRec(1)
  'FOR Cnt = 1 TO 15
  '  UBCustRec(1).CurrRevAmts(Cnt) = 0
  'NEXT
  'PUT CustFile, CUSTACCT&, UBCustRec(1)
  Close CustFile
  UBLog PWUser + (" Closed custfile in payentry getcustinfo")
  UBLog "Oper:" + Oper$ + " Entering payment for Account:" + Str$(CustAcct&)
  If UBCustRec(1).CASHONLY = "Y" Then
    CashFlag = True
  End If
    If CashFlag Then
      fpcboTenderType.Clear
      fpcboTenderType.AddItem "Cash"
      fpcboTenderType.ListIndex = 0
      fpCmdCharge.Enabled = False
      fpcmdCheck.Enabled = False
    Else
      fpcboTenderType.Clear
      fpcboTenderType.AddItem "Cash"
      fpcboTenderType.AddItem "Check"
      fpcboTenderType.AddItem "Cash & Check"
      fpcboTenderType.AddItem "Charge"
      fpcboTenderType.ListIndex = -1
      fpCmdCharge.Enabled = True
      fpcmdCheck.Enabled = True
    End If

  If Not EditFlag Then
    For cnt = 1 To NumofRevs
      fpAmtOwed(cnt - 1) = Str$(UBCustRec(1).CurrRevAmts(cnt))
      fpAmtPaid(cnt - 1) = 0
    Next
    '''  If fpAmtOwed(cnt - 1) < 0 Then fpCmdDist.Enabled = False
    TotalBalance# = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
    'LSet Form$(CustAcctFld, 0) = Str$(CustAcct&)
    txtPaymentDate = DefPayDate
    fptxtName.Caption = UBCustRec(1).CustName
    fptxtAddress.Caption = UBCustRec(1).ADDR1
    If Len(QPTrim$(UBCustRec(1).PAYCMNT)) > 0 Then
      Label4.ForeColor = &HFFFF&
      Label4.Caption = UBCustRec(1).PAYCMNT
    Else
      Label4.Caption = ""
      Label4.ForeColor = &H80000012
    End If
    fptaxexmpt = UBCustRec(1).TAXEXPT
    fpTAmtOwed = TotalBalance#
    'LSet Form$(AmtOwedFld, 0) = Str$(TotalBalance#)
    fpStatus = UBCustRec(1).Status
    fpCashAmt = 0
    fpChkAmt = 0
    fpTotReceived = 0
    fpChangeDue = 0

  Else
    Oper$ = QPTrim$(lblOperator.Caption)
    UBLog "Oper:" + Oper$ + " Editing payment for Account:" + Str$(CustAcct&)
    PayFileName$ = UBPath$ + "UBPAY" + Oper$ + ".DAT"
    ListFile = FreeFile
    Open PayFileName$ For Random Shared As ListFile Len = UBPayRecLen
    Get ListFile, PayListRec&, UBPaymentRec(1)
    Close ListFile
    fptxtName.Caption = UBPaymentRec(1).CustName
    fptxtAddress.Caption = UBPaymentRec(1).CUSTADDR
    If Len(QPTrim$(UBPaymentRec(1).CUSTCMNT)) > 0 Then
      Label4.ForeColor = &HFFFF&
      Label4.Caption = UBPaymentRec(1).CUSTCMNT
    Else
      Label4.Caption = ""
      Label4.ForeColor = &H80000012
    End If
    'Label4.Caption = UBPaymentRec(1).CUSTCMNT
    fpTAmtOwed = UBPaymentRec(1).AMTOWED
    txtPaymentDate = Num2Date(UBPaymentRec(1).PAYDATE)
    fptxtDesc = QPTrim(UBPaymentRec(1).Desc)
    For cnt = 1 To NumofRevs
      fpAmtOwed(cnt - 1) = UBPaymentRec(1).PaidOwed(cnt).AMTOWE1
      fpAmtPaid(cnt - 1) = UBPaymentRec(1).PaidOwed(cnt).AMTPD1
      
    Next
  Select Case QPTrim(UBPaymentRec(1).TENDERTY)
    Case "Cash":
      fpcboTenderType.ListIndex = 0
    Case "Check":
      fpcboTenderType.ListIndex = 1
    Case "Cash & Check":
      fpcboTenderType.ListIndex = 2
    Case "Charge":
      fpcboTenderType.ListIndex = 3
    Case Else:
      fpcboTenderType.ListIndex = -1
  End Select
  fpCashAmt = UBPaymentRec(1).CASHAMT
  fpChkAmt = UBPaymentRec(1).CHKAMT
  fpTotReceived = UBPaymentRec(1).AMTRECD
  fpChangeDue = UBPaymentRec(1).Change

    'BCopy VARSEG(UBPaymentRec(1)), VarPtr(UBPaymentRec(1)), SSEG(Form$(0, 0)), SADD(Form$(0, 0)), UBPayRecLen, 0
    'UnPackBuffer 0, 0, Form$(), Fld()
  End If
  CustAcct& = Val(fpCustRecNo)
  fptxtCity.Caption = UBCustRec(1).CITY
  fpDeposit.Caption = Using$("$###,###.##", UBCustRec(1).DepositAmt)
  'LSet CITY$ = UBCustRec(1).CITY
  'fpcboTenderType.SetFocus
  BeenDone = True
  CalcBALFlds
  Exit Sub
SkiptoHere:
  BeenDone = True
  frmLookupError.Label.Caption = "Invalid Account Number"
  frmLookupError.Label1.Caption = "Please Enter A Valid Account Number."
  frmLookupError.Show 1
  ClearScn
  
End Sub
Private Sub ClearScn()
  Dim cnt As Integer
  BeenDone = False
  fpCustRecNo = 0
  fpAcct.Enabled = True
  fpAcct = ""
  'LabelDel.Visible = False
  'fpCmdTranHist.Enabled = False
  txtPaymentDate = DefPayDate
  fpStatus = ""
  fptaxexmpt = ""
  
  fptxtName.Caption = ""
  fptxtAddress.Caption = ""
  fptxtCity.Caption = ""
  fpDeposit.Caption = "$0.00"
  fptxtDesc = ""
  fpCustRecNo = 0
  fpcboTenderType.ListIndex = -1
  fpCashAmt = 0
  fpChkAmt = 0
  fpChangeDue = 0
  For cnt = 1 To 15
    fpAmtPaid(cnt - 1) = 0
    fpAmtOwed(cnt - 1) = 0
  Next
  fpTotOwed = 0
  fpTotPaid = 0
  fpTAmtOwed = 0
  fpTotReceived = 0
  fpAcct.SetFocus
End Sub

Private Sub CalcBALFlds()
  Dim TOwd As Double, cnt As Integer, TPay As Double
  TOwd# = 0
  TPay# = 0
  For cnt = 1 To MaxRevsCnt
    TOwd# = Round#(TOwd# + fpAmtOwed(cnt - 1).DoubleValue)
    'TCur# = Round#(TCur# + fpCurrent(cnt - 1).DoubleValue)
    'fpActual(cnt - 1) = Round#(fpCurrent(cnt - 1).DoubleValue - fpAmount(cnt - 1).DoubleValue)
    'fpActual(cnt - 1) = 0
    TPay# = Round#(TPay# + fpAmtPaid(cnt - 1).DoubleValue)
  Next
  fpTotOwed = TOwd#
  fpTotPaid = TPay#
  If fpTotReceived > 0 Then
    If fpcboTenderType.ListIndex <> 3 Then
      fpChangeDue = Round#(fpTotReceived.DoubleValue - fpTotPaid.DoubleValue)
    End If
  End If
End Sub
Private Sub Autodist()
  Dim cnt As Integer, ThisAmt As Double, UBTransRecLen As Integer
  Dim NumofRevs As Integer, WhatRev As Integer, UBTran As Integer
  Dim CustFile As Integer, UBCustRecLen As Integer, ThisTran As Long
  Dim DZCnt As Integer
  ReDim UBCustRec(1) As NewUBCustRecType

  NumofRevs = MaxRevsCnt
  For cnt = 1 To NumofRevs
    WhatRev = DistArray(cnt).DistCnt - 1
    If WhatRev >= 0 Then
    ThisAmt# = Val(fpAmtOwed(WhatRev))
    If ThisAmt# < 0 Then
      TempAmtRecv# = Round#(TempAmtRecv# - ThisAmt#)
    End If
    End If
  Next
  
  For cnt = 1 To NumofRevs
    WhatRev = DistArray(cnt).DistCnt - 1
    If WhatRev >= 0 Then
      ThisAmt# = fpAmtOwed(WhatRev)
      If ThisAmt# <> 0 Then
        If TempAmtRecv# >= ThisAmt# Then
          fpAmtPaid(WhatRev) = fpAmtOwed(WhatRev)
          TempAmtRecv# = Round#(TempAmtRecv# - ThisAmt#)
        Else
          ThisAmt# = TempAmtRecv#
          fpAmtPaid(WhatRev) = ThisAmt#
          TempAmtRecv# = 0
        End If
      ElseIf TempAmtRecv# = 0 Then
        fpAmtPaid(WhatRev) = 0
      ElseIf ThisAmt# = 0 Then
        fpAmtPaid(WhatRev) = 0
      End If
    End If
  Next

CalcBALFlds
 End Sub
    
Private Sub SaveTransaction()
  Dim NumofRevs As Integer, RevCnt As Integer, ListFile As Integer
  Dim PayFileName As String, UBPayRecLen As Integer
  Dim UBCustRecLen As Integer, NumOfCustRecs As Long, NumOfRecs As Long
  Dim CustFile As Integer, cnt As Integer
  
  ReDim UBCustRec(1) As NewUBCustRecType
  ReDim UBPaymentRec(1) As UBPaymentRecType
  Oper$ = QPTrim$(lblOperator.Caption)

  PayFileName$ = UBPath$ + "UBPAY" + Oper$ + ".DAT"

  UBPayRecLen = Len(UBPaymentRec(1))
  UBCustRecLen = Len(UBCustRec(1))
  NumofRevs = MaxRevsCnt
  For cnt = 1 To 15
    If fpAmtPaid(cnt - 1) < -100000# Then
      UBPaymentRec(1).PaidOwed(cnt).AMTPD1 = 0
    Else
      UBPaymentRec(1).PaidOwed(cnt).AMTPD1 = fpAmtPaid(cnt - 1)
    End If
    If fpAmtOwed(cnt - 1) < -100000# Then
      UBPaymentRec(1).PaidOwed(cnt).AMTOWE1 = 0
    Else
      UBPaymentRec(1).PaidOwed(cnt).AMTOWE1 = fpAmtOwed(cnt - 1)
    End If
  Next
  UBPaymentRec(1).OPERNUM = QPTrim(lblOperator.Caption)
  UBPaymentRec(1).PAYDATE = Date2Num(txtPaymentDate)
  UBPaymentRec(1).CustAcct = fpAcct
  UBPaymentRec(1).CustName = QPTrim(fptxtName.Caption)
  UBPaymentRec(1).CUSTADDR = QPTrim(fptxtAddress.Caption)
  UBPaymentRec(1).CUSTCMNT = QPTrim(Label4.Caption)
  UBPaymentRec(1).TaxExempt = QPTrim(fptaxexmpt)
  UBPaymentRec(1).AMTOWED = fpTAmtOwed
  Select Case fpcboTenderType.ListIndex
    Case 0:
      UBPaymentRec(1).TENDERTY = "Cash"
    Case 1:
      UBPaymentRec(1).TENDERTY = "Check"
    Case 2:
      UBPaymentRec(1).TENDERTY = "Cash & Check"
    Case 3:
      UBPaymentRec(1).TENDERTY = "Charge"
    Case Else:
      UBPaymentRec(1).TENDERTY = "Unknown"
  End Select
  UBPaymentRec(1).CASHAMT = fpCashAmt
  UBPaymentRec(1).CHKAMT = fpChkAmt
  UBPaymentRec(1).AMTRECD = fpTotReceived
  UBPaymentRec(1).Change = fpChangeDue
  UBPaymentRec(1).Desc = QPTrim(fptxtDesc)
  UBPaymentRec(1).TOTOWED = fpTotOwed
  UBPaymentRec(1).AMTPAID = fpTotPaid
  UBPaymentRec(1).Status = QPTrim(fpStatus)
  ListFile = FreeFile
  Open PayFileName$ For Random Shared As ListFile Len = UBPayRecLen
  If EditFlag Then
    Put #ListFile, PayListRec&, UBPaymentRec(1)
    EditFlag = False
  Else
    NumOfRecs& = (LOF(ListFile) \ UBPayRecLen) + 1
    PayListRec& = NumOfRecs&
    Put #ListFile, PayListRec&, UBPaymentRec(1)
  End If
  UBLog "Oper:" + Oper$ + " Updated Paylist for Account:" + Str$(UBPaymentRec(1).CustAcct)
  Close ListFile

  LoadPayList
  'ClearScn
End Sub

Private Sub CheckInfo()
  Dim TestDate As Integer, TestAmt As Double
  TestAmt = 0
  ChkOKFlag = True
  TestDate = Date2Num(txtPaymentDate)
  If TestDate < 0 Then
    ChkOKFlag = False
    MsgBox "Invalid Date.", vbOKOnly, "Request Canceled."
    GoTo BadDate
  End If
  If fpcboTenderType.ListIndex = -1 Then
    MsgBox "You Must Select A Tender Type.", vbOKOnly, "Invalid Selection"
    ChkOKFlag = False
    GoTo BadDate
  End If
  If fpcboTenderType.ListIndex = 0 And fpChkAmt.DoubleValue > 0 Then
    ChkOKFlag = False
    MsgBox "Invalid Tender Type. The Check/Charge Amount Should Be ZERO.", vbOKOnly, "Request Canceled."
    GoTo BadDate
  End If
  If (fpcboTenderType.ListIndex = 1 Or fpcboTenderType.ListIndex = 3) And fpCashAmt.DoubleValue > 0 Then
    ChkOKFlag = False
    MsgBox "Invalid Tender Type. The Cash Amount Should Be ZERO.", vbOKOnly, "Request Canceled."
    GoTo BadDate
  End If
  If fpcboTenderType.ListIndex = 2 And (fpChkAmt.DoubleValue <= 0 Or fpCashAmt.DoubleValue <= 0) Then
    ChkOKFlag = False
    MsgBox "Invalid Amounts. The Check and Cash Amount Should Be Greater than ZERO.", vbOKOnly, "Request Canceled."
    GoTo BadDate
  End If
  If fpTotReceived.DoubleValue <= 0 Or fpTotPaid.DoubleValue <= 0 Then
    ChkOKFlag = False
    MsgBox "Invalid Amount. The Total Received Should NOT Be ZERO.", vbOKOnly, "Request Canceled."
    GoTo BadDate
  End If
  If fpChangeDue.DoubleValue >= 0 Then
    TestAmt = Round#(fpTotReceived.DoubleValue - fpChangeDue.DoubleValue)
    If TestAmt <> fpTotPaid Then '.DoubleValue Then
      ChkOKFlag = False
      MsgBox "The Total Received does NOT equal the Total Paid.", vbOKOnly, "Request Canceled."
      GoTo BadDate
    End If
  Else
    ChkOKFlag = False
    MsgBox "The Amount Distributed May Not Be More Than Amount Received.", vbOKOnly, "Request Canceled."
    GoTo BadDate
  End If
  Exit Sub
BadDate:
  Exit Sub
End Sub
Private Sub LoadPayList()
  Dim cnt As Long, RevCnt As Integer, ListFile As Integer
  Dim PayFileName As String, UBPayRecLen As Integer, PayListRec As Long
  Dim PayRecpName As String, NumOfRecs As Long
  Dim PCustAcct As Long
  ReDim UBPaymentRec(1) As UBPaymentRecType

  UBPayRecLen = Len(UBPaymentRec(1))
  
  Oper$ = QPTrim$(lblOperator.Caption)

  PayFileName$ = UBPath$ + "UBPAY" + Oper$ + ".DAT"
  PayRecpName$ = UBPath$ + "UBRCP" + Oper$ + ".RPT"

  ListFile = FreeFile
  Open PayFileName$ For Random Shared As ListFile Len = UBPayRecLen
  NumOfRecs& = LOF(ListFile) \ UBPayRecLen
  If NumOfRecs& > 0 Then
    ReDim PayList(1 To NumOfRecs&) As PayListType
    For cnt& = 1 To NumOfRecs&
      Get #ListFile, cnt&, UBPaymentRec(1)
      PayList(cnt&).CustRec = UBPaymentRec(1).CustAcct
      PCustAcct = UBPaymentRec(1).CustAcct
      PayList(cnt&).Listrec = cnt&
    Next
  End If
  Close ListFile
  PayListCnt& = NumOfRecs&
End Sub

Private Sub CheckPayList()
  Dim cnt As Long ', PayListRec As Long, ListFile As Integer
'  Dim PayFileName As String, UBPayRecLen As Integer
'  Dim NumOfRecs As Long
'  ReDim UBPaymentRec(1) As UBPaymentRecType
'
'
'  UBPayRecLen = Len(UBPaymentRec(1))
'
'  Oper$ = QPTrim$(lblOperator.Caption)
'
'  PayFileName$ = "UBPAY" + Oper$ + ".DAT"
'
'  ListFile = FreeFile
'  Open PayFileName$ For Random Shared As ListFile Len = UBPayRecLen
'  NumOfRecs& = LOF(ListFile) \ UBPayRecLen
'  Close
'  PayListCnt& = NumOfRecs&
  EditFlag = False
  If PayListCnt& > 0 Then
    'ReDim Preserve PayList(1 To PayListCnt&) As PayListType
    For cnt = 1 To PayListCnt&
      If PayList(cnt).CustRec = CustAcct& Then
        PayListRec& = PayList(cnt).Listrec
        EditFlag = True
        Exit For
      End If
    Next
  End If
End Sub
Private Sub PrintReceipt()
  Dim ListFile As Integer, PayFileName As String, UBPayRecLen As Integer
  Dim RecptNum As Long, RHandle As Integer, PayRecpName As String
  Dim CutPaper As String, PostDate As String, RevCnt As Integer
  Dim NumofRevs As Integer, RecpRev As String, PCnt As Integer, zz As Integer
  Dim RHandle2 As Integer, PayRecpName2 As String, RptHandle2 As Integer
  ReDim UBPaymentRec(1) As UBPaymentRecType
'  ReDim Preserve RevText$(1 To MaxRevsCnt)
  RecpRev$ = Space$(15)
  CutPaper$ = Chr$(29) + Chr$(86) + Chr$(66) + Chr$(64)
  If InStr(TOWNNAME$, "DOBSON") > 0 Then CutPaper$ = Chr(27) & Chr(&H64) & Chr(&H30)

  UBPayRecLen = Len(UBPaymentRec(1))
  PayFileName$ = UBPath$ + "UBPAY" + Oper$ + ".DAT"
  PayRecpName$ = UBPath$ + "UBRCP" + Oper$ + ".RPT"
  PayRecpName2$ = UBPath$ + "UBVLD" + Oper$ + ".Rpt"
  PostDate$ = txtPaymentDate
  ListFile = FreeFile
  Open PayFileName$ For Random Shared As ListFile Len = UBPayRecLen
  RecptNum& = LOF(ListFile) / UBPayRecLen
  Get #ListFile, PayListRec&, UBPaymentRec(1)
  Close
  NumofRevs = MaxRevsCnt
  RHandle = FreeFile
  Open PayRecpName$ For Output As RHandle
  If Caldwell = True Then
  Print #RHandle, Chr$(27); "p"; Chr$(0); Chr$(25); Chr$(250); Chr$(7)
  Print #RHandle, Tab(7); TOWNNAME$; Tab(43); TOWNNAME$
  Print #RHandle, Tab(7); "UTILITY PAYMENT"; Tab(43); "UTILITY PAYMENT"
  Print #RHandle, Tab(7); "Date: "; PostDate$; Tab(43); "Date: "; PostDate$
  Print #RHandle,
  Print #RHandle, Tab(7); "CUSTOMER NAME & DESC."; Tab(43); "CUSTOMER NAME & DESC."
  Print #RHandle, Tab(7); UBPaymentRec(1).CustName; Tab(43); UBPaymentRec(1).CustName
  Print #RHandle, Tab(7); UBPaymentRec(1).CUSTADDR; Tab(43); UBPaymentRec(1).CUSTADDR
  Print #RHandle, Tab(7); UBPaymentRec(1).Desc; Tab(43); UBPaymentRec(1).Desc
  Print #RHandle, Tab(7); "Acct. No. "; UBPaymentRec(1).CustAcct; Tab(43); "Acct. No. "; UBPaymentRec(1).CustAcct
  Print #RHandle,
  Print #RHandle, Tab(7); "Total Owed:  "; Using("$##,###,###.##", UBPaymentRec(1).TOTOWED); Tab(43); "Total Owed:  "; Using("$##,###,###.##", UBPaymentRec(1).TOTOWED)
  Print #RHandle, Tab(7); "Total Paid:  "; Using("$##,###,###.##", UBPaymentRec(1).AMTPAID); Tab(43); "Total Paid:  "; Using("$##,###,###.##", UBPaymentRec(1).AMTPAID)
  Print #RHandle, Tab(7); "Change Due:  "; Using("$##,###,###.##", UBPaymentRec(1).Change); Tab(43); "Change Due:  "; Using("$##,###,###.##", UBPaymentRec(1).Change)
  Print #RHandle, Tab(7); "   Balance:  "; Using("$##,###,###.##", UBPaymentRec(1).TOTOWED - UBPaymentRec(1).AMTPAID); Tab(43); "   Balance:  "; Using("$##,###,###.##", UBPaymentRec(1).TOTOWED - UBPaymentRec(1).AMTPAID)
  Print #RHandle,
  '16 to here
  For RevCnt = 1 To NumofRevs
    If UBPaymentRec(1).PaidOwed(RevCnt).AMTPD1 <> 0 Or UBPaymentRec(1).PaidOwed(RevCnt).AMTOWE1 <> 0 Then
      LSet RecpRev$ = RevText$(RevCnt)
      Print #RHandle, Tab(7); RecpRev$; Using("$########.##", UBPaymentRec(1).PaidOwed(RevCnt).AMTPD1); Tab(43); RecpRev$; Using("$########.##", UBPaymentRec(1).PaidOwed(RevCnt).AMTPD1)
      PCnt = PCnt + 1
    End If
  Next
  If PCnt < 6 Then
    For zz = PCnt To 6
      Print #RHandle,
    Next
  End If

  Print #RHandle, Tab(7); "Operator: "; OPERNUM; Tab(43); "Operator: "; OPERNUM
   ' RecptNo& = FileSize(PayFileName$) \ UBPayRecLen
  Print #RHandle, Tab(7); "Receipt: "; Using("######", PayListRec&); Tab(43); "Receipt: "; Using("######", PayListRec&)
  Print #RHandle,
  Print #RHandle, Tab(7); " T H A N K   Y O U !"; Tab(43); " T H A N K   Y O U !"
  Print #RHandle,
  Close RHandle

  'Shell$ = "type " + PayRecpName$ + " > com2:"
  'SHELL Shell$

  'PrintRptFile Header$, PayRecpName$, RecpPort, RetCode%, 5


  Else  ' NOT CALDWELL    @@@@@@@@@@@@@@@@@@@@@@@@@@@@
  If CntrlDef = 1 Then
    If Biscoe = True Then
      Print #RHandle, Chr$(7)
    Else
    fpcmdDrawer_Click
    '  Print #RHandle, Chr$(27); "p"; Chr$(0) '; Chr$(25); Chr$(250)
    End If
  End If
  Print #RHandle, TOWNNAME$
  Print #RHandle, "UTILITY PAYMENT"
  Print #RHandle, "Date: "; PostDate$
  Print #RHandle, "Time: "; Time
  Print #RHandle,
  Print #RHandle, "CUSTOMER NAME & DESC. OF PAYMENT"
  Print #RHandle, UBPaymentRec(1).CustName
  Print #RHandle, UBPaymentRec(1).CUSTADDR
  Print #RHandle, UBPaymentRec(1).Desc
  Print #RHandle, "Acct. No. "; UBPaymentRec(1).CustAcct
  Print #RHandle,
  Print #RHandle, QPTrim(UBPaymentRec(1).TENDERTY)
  Print #RHandle,
  Print #RHandle, "       Cash: "; Using("$##,###,###.##", UBPaymentRec(1).CASHAMT)
  If QPTrim$(UBPaymentRec(1).TENDERTY) <> "Charge" Then
    Print #RHandle, "      Check: "; Using("$##,###,###.##", UBPaymentRec(1).CHKAMT)
    If Biscoe = False Then Print #RHandle, "     Charge: "; Using("$##,###,###.##", 0)
  Else
    Print #RHandle, "      Check: "; Using("$##,###,###.##", 0)
    If Biscoe = False Then Print #RHandle, "     Charge: "; Using("$##,###,###.##", UBPaymentRec(1).CHKAMT)
  End If
  Print #RHandle, " Total Owed: "; Using("$##,###,###.##", UBPaymentRec(1).TOTOWED)
  Print #RHandle, " Total Paid: "; Using("$##,###,###.##", UBPaymentRec(1).AMTRECD)
  Print #RHandle, " Change Due: "; Using("$##,###,###.##", UBPaymentRec(1).Change)
  Print #RHandle, "Amt Applied: "; Using("$##,###,###.##", UBPaymentRec(1).AMTPAID)
  Print #RHandle, "    Balance: "; Using("$##,###,###.##", UBPaymentRec(1).TOTOWED - UBPaymentRec(1).AMTPAID)
  Print #RHandle,
  For RevCnt = 1 To NumofRevs
    If UBPaymentRec(1).PaidOwed(RevCnt).AMTPD1 <> 0 Or UBPaymentRec(1).PaidOwed(RevCnt).AMTOWE1 <> 0 Then
      LSet RecpRev$ = RevText$(RevCnt)
      'PRINT #RHandle, RecpRev$; USING "$$####,#.##"; UBPaymentRec(1).PaidOwed(RevCnt).AmtOwe1; UBPaymentRec(1).PaidOwed(RevCnt).AmtPd1
      Print #RHandle, RecpRev$; Using("$#####.##", UBPaymentRec(1).PaidOwed(RevCnt).AMTPD1)
    End If
  Next
  Print #RHandle,
  Print #RHandle, "Operator: "; OPERNUM
  Print #RHandle, "Receipt#: "; Using("######", PayListRec&)
  Print #RHandle,
  Print #RHandle, "       T H A N K   Y O U !"
  Print #RHandle,
  Print #RHandle,
  Print #RHandle,
  Print #RHandle,
  'Print #RHandle,
  If CntrlDef = 1 Then
    If Biscoe = False Then Print #RHandle, CutPaper$
 
  Else
    Print #RHandle,
    Print #RHandle,
 '   Print #RHandle,
  End If
  Close RHandle

  'Shell$ = "type " + PayRecpName$ + " > com2:"
  'SHELL Shell$
'  If CntrlDef = 1 Then
'    fpcmdDrawer_Click
'  End If
  'PrintRptFile Header$, PayRecpName$, RecpPort, RetCode%, 5
 End If
  Dim RptHandle As Integer, LPTHandle As Integer
  Dim RptA As Integer, LPTA As Integer, ToPrintA As String
  Dim ToPrint As String, CopyLoop As Integer, DefPrinter As String
  On Error GoTo Cancel
  'Printer.Print
'''  to strReportFile DefPrinter'[ADDITIVE] | PortName]
10:
  DefPrinter = RecpPort '"LPT" + QPTrim$(Str$(RecpPort)) + ":"
20:
 ' MsgBox "Printer -" + DefPrinter, vbOKOnly
  
  For CopyLoop = 1 To 1 'Copies
    LPTHandle = FreeFile
    Open DefPrinter For Output As LPTHandle
    RptHandle = FreeFile
30:
    Open PayRecpName$ For Input As RptHandle
40:
    Do
      If frmPrint.cmdCancel = False Then
45:
        Line Input #RptHandle, ToPrint$
        
        ToPrint$ = RTrim$(ToPrint$)
        Print #LPTHandle, ToPrint$
      Else
50:
        Exit Do
        'Printer.EndDoc
      End If
    Loop Until eof(RptHandle)
60:
    Close RptHandle
62:
    Close LPTHandle
65:
    Next CopyLoop
68:
 Printer.EndDoc
69:
If Caldwell = False Then
  If QPTrim(UBPaymentRec(1).TENDERTY) = "Check" Or QPTrim(UBPaymentRec(1).TENDERTY) = "Cash & Check" Then
   If RctValidate And PrnVali = True Then
     RHandle2 = FreeFile
     Open PayRecpName2$ For Output As RHandle2
     Print #RHandle2, Chr$(27); Chr$(&H63); Chr$(&H30); Chr$(&H4)
     Print #RHandle2, Chr$(13); Chr$(10)
     Print #RHandle2, Tab(12); TOWNNAME$
     Print #RHandle2, Tab(12); "Bank- "; BnkAcctNum$
     Print #RHandle2, Tab(12); "FOR DEPOSIT ONLY"
     Print #RHandle2, Tab(12); "Acct. No. "; UBPaymentRec(1).CustAcct
     Print #RHandle2, Tab(12); "Date: "; PostDate$
     Print #RHandle2, Tab(12); "Time: "; Time
     Print #RHandle2,
     Print #RHandle2, Chr$(12)
     Close RHandle2
     LPTHandle = FreeFile
     Open DefPrinter For Output As LPTHandle
     RptHandle2 = FreeFile
     Open PayRecpName2$ For Input As RptHandle2
     Do
       If frmPrint.cmdCancel = False Then
         Line Input #RptHandle2, ToPrint$
         ToPrint$ = RTrim$(ToPrint$)
         Print #LPTHandle, ToPrint$
       Else
         Exit Do
       End If
     Loop Until eof(RptHandle2)
     Close RptHandle2
     Close LPTHandle
    Printer.EndDoc
    UBLog "Oper: " + Oper$ + " Print Validation Acct:" + Str(UBPaymentRec(1).CustAcct)
  End If
 End If
 End If
70:
 UBLog "Oper: " + Oper$ + " Print receipt Acct:" + Str(UBPaymentRec(1).CustAcct)
 KillFile PayRecpName$

80:
  Exit Sub
Cancel:
  If Err > 0 Then
    MsgBox "Error Code Was " + DefPrinter + Err.Description + Str$(Err) + " (PrintWSet - Line:" & Erl & ")"
  End If
  Close
  Exit Sub

  
End Sub
'Sub AddEditPayment(OPERNUM, PostDate$)
'  UBLog " IN: Bill Payments,  OPER:" + Str$(OPERNUM)
'
'
'
'  CITY$ = Space$(20)
'  Deps$ = Space$(12)
'  fmt$ = "$$####.##"
'  RecpRev$ = Space$(15)
'
'  'look into keeping date on payments edited on a different day
'  ReDim UBCustRec(1) As NewUBCustRecType
'  ReDim UBPaymentRec(1) As UBPaymentRecType
'  ReDim PayList(1 To 1) As PayListType
'
'  UBCustRecLen = Len(UBCustRec(1))
'  UBPayRecLen = Len(UBPaymentRec(1))
'  GoSub LoadPayList
'
'  NumOfRevs = MaxRevsCnt
'
'  ReDim RevText$(1 To MaxRevsCnt)
'  ReDim UBSetUpRec(1) As UBSetupRecType
'
'  ReDim DistArray(1 To MaxRevsCnt) As DistArrayType
'
'  UBSetupLen = Len(UBSetUpRec(1))
'  'FGetAH "UBSETUP.DAT", UBSetUpRec(1), UBSetupLen, 1            'load it
'  LoadUBSetUpFile UBSetUpRec(), UBSetupLen
'
'  RecpPort = Val(UBSetUpRec(1).RecpPort)
'  If RecpPort < 1 Or RecpPort > 2 Then
'    RecpPort = 1
'  End If
'
'  For RevCnt = 1 To MaxRevsCnt
'    RevText$(RevCnt) = left$(QPTrim$(UBSetUpRec(1).Revenues(RevCnt).REVNAME), 14)
'    DistArray(RevCnt).DistOrder = UBSetUpRec(1).Revenues(RevCnt).DistOr
'    DistArray(RevCnt).DistCnt = RevCnt
'    If Len(RevText$(RevCnt)) = 0 Then
'      NumOfRevs = RevCnt - 1
'      Exit For
'    End If
'  Next
'
'  ReDim Preserve DistArray(1 To NumOfRevs) As DistArrayType
'
'  Do
'    OutOfOrder = False          'assume it's sorted
'    For x = 1 To NumOfRevs - 1
'      If DistArray(x).DistOrder > DistArray(x + 1).DistOrder Then
'        SWAP DistArray(x), DistArray(x + 1)     'if we had to swap
'        OutOfOrder = True       'we're not done yet
'      End If
'    Next
'  Loop While OutOfOrder
'
'  TownName$ = UBSetUpRec(1).UTILNAME
'  If InStr(TownName$, "CALDWELL") > 0 Then
'    CaldFlag = True
'  End If
'
'  If NumOfRevs < MaxRevsCnt Then
'    ReDim Preserve RevText$(1 To NumOfRevs)
'  End If
'  'GoSub ClearForm
'
'  ReDim AmtOweFlds(1 To NumOfRevs)
'  'REDIM PrevAmtOwe(1 TO NumOfRevs) AS DOUBLE
'  ReDim AmtPadFlds(1 To NumOfRevs)
'
'  For cnt = 1 To NumOfRevs
'    AmtOweFlds(cnt) = FldNum%("AMTOWE" + LTrim$(Str$(cnt)), Fld())
'    AmtPadFlds(cnt) = FldNum%("AMTPD" + LTrim$(Str$(cnt)), Fld())
'  Next
'
'  AmtOwedFld = FldNum%("AMTOWED", Fld())        'these get field numbers
'  TenderFld = FldNum%("TENDERTY", Fld())        'so we can track what field
'  CashAmtFld = FldNum%("CASHAMT", Fld())        'the user is currently on
'  ChkAmtFld = FldNum%("CHKAMT", Fld())
'  AmtRecvFld = FldNum%("AMTRECD", Fld())
'  ChangeFld = FldNum%("CHANGE", Fld())
'  TotalFld = FldNum%("TOTOWED", Fld())
'  AmtPaidFld = FldNum%("AMTPAID", Fld())
'  DescFld = FldNum%("DESC", Fld())
'  CustAcctFld = FldNum%("CUSTACCT", Fld())
'  StatFld = FldNum%("STATUS", Fld())
'  TaxExemptFld = FldNum%("TAXEXPT", Fld())
'
'  '--define the multi-choice fields
'
'  ReDim Choice$(0 To 3, 0 To 0)
'
'  Choice$(0, 0) = QPTrim$(Str$(TenderFld))
'  'Choice$(1, 0) = "Cash"
'  'Choice$(2, 0) = "Check"
'  'Choice$(3, 0) = "Cash & Check"
'
'  Action = 1
'  FirstTime = True
'
'  DisplayUBScrn ScrnName$
'
'  Do
'
'    EditForm Form$(), Fld(), frm(1), Cnf, Action
'    If frm(1).Edited And frm(1).PrevFld <> frm(1).FldNo Then
'      BeenEditedFlag = True     'if the form has been edited
'    End If      'set the edited flag
'
'    If FirstTime Then
'      FirstTime = False         'if this is the first time
'      GoSub ShowRevSources      '
'      GoSub SetOperInfo
'      QPrintRC CITY$, 8, 15, -1
'      QPrintRC Deps$, 9, 15, -1
'    End If
'
'    If DistFlag And Not PrevFlag Then
'      TempAmtRecv# = Value#(Form$(AmtRecvFld, 0), ECode)
'      GoSub AutoDistribute:
'    End If
'    DistFlag = False
'
'    If frm(1).FldNo > CustAcctFld And frm(1).PrevFld = CustAcctFld Then
'      PrevFlag = False
'      CustAcct& = QPValL(Form$(CustAcctFld, 0))
'      GoSub CheckPayList
'      GoSub GetCustInfo
'
'    ElseIf frm(1).FldNo = CustAcctFld And frm(1).PrevFld <> CustAcctFld Then
'      MPaintBox 22, 37, 22, 41, 112
'      MPaintBox 22, 35, 22, 36, 126
'    End If
'
'    If frm(1).FldNo = TenderFld And frm(1).PrevFld <> TenderFld Then
'      MPaintBox 22, 18, 22, 22, 112             'this paints the cash and chec
'      MPaintBox 22, 28, 22, 31, 112             'buttons based on whether user
'      MPaintBox 22, 16, 22, 17, 126             'buttons based on whether user
'      MPaintBox 22, 26, 22, 27, 126             'buttons based on whether user
'      GoSub FixCashChkFlds
'    ElseIf frm(1).PrevFld = TenderFld And frm(1).FldNo <> TenderFld Then
'      MPaintBox 22, 16, 22, 22, 115             'is on tender type field or
'      MPaintBox 22, 26, 22, 31, 115             'on any another field
'      GoSub FixCashChkFlds
'    End If
'
'    '--Check for Key presses
'    Select Case frm(1).KeyCode
'    Case EscKey
'      If BeenEditedFlag Then
'        SaveFlag = PromptSaveData
'        Select Case SaveFlag
'        Case True               'user wants to save
'          StuffBuf Chr$(0) + Chr$(Abs(F10Key))
'        Case False              'user wants to abandon
'          ExitFlag = True
'        Case Else               'continue editing
'        End Select
'        Action = 1
'      Else
'        ExitFlag = True
'      End If
'
'    Case F4KEY  'Customer History
'      If CustAcct& > 0 Then
'        SaveScrn TempScrn()
'        CustomerInquiry CustAcct&
'        RestScrn TempScrn()
'        Action = 2
'      End If
'    Case -88    'Shift-F5 Previous cash payment
'      If frm(1).FldNo = TenderFld Then          'if user is on tender field
'        PrevFlag = True
'        GoSub DoCashPayment     'and F5key then go do the
'      End If    'cash payment routine
'      DistFlag = True
'    Case F5KEY
'      If frm(1).FldNo = TenderFld Then          'if user is on tender field
'        PrevFlag = False
'        GoSub DoCashPayment     'and F5key then go do the
'      End If    'cash payment routine
'      DistFlag = True
'    Case -89    'ShiftF6 Previous Check payment
'      If frm(1).FldNo = TenderFld Then          'if user is on tender field
'        PrevFlag = True
'        GoSub DoCheckPayment    'and F6key then go do the
'      End If    'cash payment routine
'      DistFlag = True
'    Case F6KEY  'Check Payment
'      If frm(1).FldNo = TenderFld Then          'if user is on tender field
'        GoSub DoCheckPayment    'and F6key then go do the
'      End If    'check payment routine
'      DistFlag = True
'
'    Case F7KEY  'Lookup Customer
'      If frm(1).FldNo = 3 Then  'if user is on the Customer field
'        SaveScrn TempScrn()     'and F7key then do lookup routine
'        MPaintBox 4, 5, 22, 75, 8
'        LookUp CustAcct&, "Payment", 2, False, False
'        RestScrn TempScrn()
'        If CustAcct& > 0 Then   'if this is a valid customer
'          GoSub CheckPayList
'          GoSub GetCustInfo     'go get customer info
'          frm(1).FldNo = 4
'          Action = 1
'        Else
'          GoSub ClearForm
'          frm(1).FldNo = 1
'          Action = 1
'        End If
'      End If
'
'      '    CASE -92    'Shift F9
'      '      TempAmtRecv# = Value#(Form$(AmtRecvFld, 0), ECode)
'      '      IF TempAmtRecv# > 0 THEN
'      '        GOSUB AutoDistributeOLD
'      '      END IF
'
'    Case F8KEY
'      OPENDrawer RecpPort
'

'
'Sub OPENDrawer(RecpPort)
'  On Local Error Resume Next
'
'  'RecPort = GetDEFPort%
'  Port$ = "LPT" + QPTrim$(Str$(RecpPort)) + ":"
'
'  PortFile = FreeFile
'  Open Port$ For Output As #PortFile
'  Print #PortFile, Chr$(27); "p"; Chr$(0); Chr$(25); Chr$(250)
'  Print #PortFile, Chr$(7)
'  Close PortFile
'
'End Sub
'    Case F9KEY
'
'      TempAmtRecv# = Value#(Form$(AmtRecvFld, 0), ECode)
'      If TempAmtRecv# > 0 Then
'        GoSub AutoDistribute
'      End If
'
'    Case F10Key 'Save
'      GoSub CheckPaymentInfo
'      If PaymentOKFlag Then
'        Select Case AskSavePayment(UBSetUpRec(1).RECPDEFT)
'        Case 1
'          SaveScrn TempScrn()
'          GoSub SaveTransaction 'do the save routine
'          If CaldFlag Then
'            GoSub CWPrintReceipt
'          Else
'            GoSub PrintReceipt
'          End If
'          RestScrn TempScrn()
'          DisplayUBScrn "PRESSKEY"
'          WaitForAction
'          RestScrn TempScrn()
'          GoSub ClearForm
'          GoSub ClearCity
'          frm(1).FldNo = 1
'          Action = 1
'        Case True
'          ReceiptFlag = False
'          SaveScrn TempScrn()
'          GoSub SaveTransaction 'do the save routine
'          DisplayUBScrn "PRESSKEY"
'          WaitForAction
'          RestScrn TempScrn()
'          GoSub ClearForm
'          GoSub ClearCity
'          frm(1).FldNo = 1
'          Action = 1
'        Case False
'          Action = 2
'        End Select
'      End If
'    Case Is <> 0
'      'STOP
'    End Select
'
'    '--check for mouse clicks on buttons not attached to the form
'    If frm(1).Presses Then
'      Select Case frm(1).MRow
'      Case 22   'Look for the f10 or esc button
'        Select Case frm(1).MCol
'        Case 5 To 13            'f4 History
'          PressButton F4KEY, 22, 5, 13
'        Case 15 To 23           'f5 cash
'          If frm(1).FldNo = TenderFld Then
'            PressButton F5KEY, 22, 15, 23
'          End If
'        Case 25 To 32           'f6 check
'          If frm(1).FldNo = TenderFld Then
'            PressButton F6KEY, 22, 25, 32
'          End If
'        Case 34 To 42           'f7 Look-Up
'          PressButton F7KEY, 22, 34, 42
'        Case 44 To 52           'f9 Save
'          PressButton F9KEY, 22, 44, 52
'        Case 54 To 63           'f10 Save
'          PressButton F10Key, 22, 54, 63
'        Case 65 To 75           '--cancel button
'          PressButton EscKey, 22, 65, 75
'        End Select
'      End Select                'row
'    End If
'  Loop Until ExitFlag
'
'  Erase TempScrn, UBCustRec, UBPaymentRec, RevText$, UBSetUpRec
'  UBLog "OUT: Bill Payments" + CrLf$
'  HideCursor
'
'  Exit Sub
'
'
'SetOperInfo:
'  LSet Form$(1, 0) = FUsing$(Str$(OPERNUM), "##")
'  LSet Form$(2, 0) = PostDate$
'  Action = 2
'Return
'
'ClearForm:
'  For F = 1 To NumFlds
'    LSet Form$(F, 0) = ""       '--Clear all fields
'  Next
'  BeenEditedFlag = False        'clear the edited flag
'Return
'
'DoCashPayment:
'  'IF PrevFlag THEN
'  '  ThisAmt# = 0
'  '  FOR Cnt = 1 TO NumOfRevs
'  '    ThisAmt# = Round#(ThisAmt# + PrevAmtOwe(Cnt))
'  '  NEXT
'  '  IF ThisAmt# > 0 THEN
'  '    LSET Form$(TenderFld, 0) = Choice$(1, 0)
'  '    LSET Form$(ChkAmtFld, 0) = "0"
'  '    LSET Form$(CashAmtFld, 0) = QPTrim$(STR$(ThisAmt#))
'  '    GOSUB PaymentCommon
'  '  END IF
'  'ELSE
'  LSet Form$(TenderFld, 0) = Choice$(1, 0)
'  LSet Form$(ChkAmtFld, 0) = "0"
'  LSet Form$(CashAmtFld, 0) = Form$(AmtOwedFld, 0)
'  GoSub PaymentCommon
'  'END IF
'Return
'
'DoCheckPayment:
'  '  IF PrevFlag THEN
'  '    ThisAmt# = 0
'  '    FOR Cnt = 1 TO NumOfRevs
'  '      ThisAmt# = Round#(ThisAmt# + PrevAmtOwe(Cnt))
'  '    NEXT
'  '    IF ThisAmt# > 0 THEN
'  '      LSET Form$(TenderFld, 0) = Choice$(2, 0)
'  '      LSET Form$(ChkAmtFld, 0) = QPTrim$(STR$(ThisAmt#))
'  '      LSET Form$(CashAmtFld, 0) = "0"
'  '      GOSUB PaymentCommon
'  '    END IF
'  '  ELSE
'  LSet Form$(TenderFld, 0) = Choice$(2, 0)
'  LSet Form$(ChkAmtFld, 0) = Form$(AmtOwedFld, 0)
'  LSet Form$(CashAmtFld, 0) = "0"
'  GoSub PaymentCommon
'  '  END IF
'Return
'
'PaymentCommon:
'  SaveField TenderFld, Form$(), Fld(), BadField
'  SaveField ChkAmtFld, Form$(), Fld(), BadField
'  SaveField CashAmtFld, Form$(), Fld(), BadField
'  LSet Form$(ChangeFld, 0) = "0"
'  SaveField ChangeFld, Form$(), Fld(), BadField
'
'  For cnt = 1 To NumOfRevs
'    'IF PrevFlag THEN
'    '  LSET Form$(AmtOweFlds(Cnt) + 1, 0) = QPTrim$(STR$(PrevAmtOwe(Cnt)))
'    'ELSE
'    LSet Form$(AmtOweFlds(cnt) + 1, 0) = Form$(AmtOweFlds(cnt), 0)
'    'END IF
'    SaveField AmtOweFlds(cnt) + 1, Form$(), Fld(), BadField
'  Next
'
'  MPaintBox 22, 16, 22, 22, 115 'is on tender type field or
'  MPaintBox 22, 26, 22, 31, 115 'on any another field
'
'  frm(1).FldNo = DescFld
'
'FixCashChkFlds:
'
'  Select Case QPTrim$(Form$(TenderFld, 0))
'  Case Choice$(1, 0)            'CASH               this sets the cash or chec
'    Fld(TenderFld + 1).Protected = False        'amount fields protected or
'    Fld(TenderFld + 2).Protected = True         'unprotected based of the
'    LSet Form$(ChkAmtFld, 0) = "0"
'  Case Choice$(2, 0)            'CHECK              tender type field selectio
'    Fld(TenderFld + 1).Protected = True
'    Fld(TenderFld + 2).Protected = False
'    LSet Form$(CashAmtFld, 0) = "0"
'  Case Else     'BOTH
'    Fld(TenderFld + 1).Protected = False
'    Fld(TenderFld + 2).Protected = False
'  End Select
'
'  SaveField ChkAmtFld, Form$(), Fld(), BadField
'  SaveField CashAmtFld, Form$(), Fld(), BadField
'
'  CalcFields 0, AmtPadFlds(1), Form$(), Fld()
'  CalcFields 0, CashAmtFld, Form$(), Fld()
'
'  PrintArray 1, NumFlds - 1, Form$(), Fld()
'
'  'GOSUB AutoDistribute:
'Return
'
'AutoDistribute:
'
'  For cnt = 1 To NumOfRevs
'    WhatRev = DistArray(cnt).DistCnt
'    ThisAmt# = Value(Form$(AmtOweFlds(WhatRev), 0), ECode)
'    If ThisAmt# < 0 Then
'      TempAmtRecv# = Round#(TempAmtRecv# - ThisAmt#)
'    End If
'  Next
'
'  For cnt = 1 To NumOfRevs
'    WhatRev = DistArray(cnt).DistCnt
'    ThisAmt# = Value(Form$(AmtOweFlds(WhatRev), 0), ECode)
'    If ThisAmt# <> 0 Then
'      If TempAmtRecv# >= ThisAmt# Then
'        LSet Form$(AmtOweFlds(WhatRev) + 1, 0) = QPTrim$(Form$(AmtOweFlds(WhatRev), 0))
'        TempAmtRecv# = Round#(TempAmtRecv# - ThisAmt#)
'      Else
'        ThisAmt# = TempAmtRecv#
'        LSet Form$(AmtOweFlds(WhatRev) + 1, 0) = Str$(ThisAmt#)
'        TempAmtRecv# = 0
'      End If
'    ElseIf TempAmtRecv# = 0 Then
'      LSet Form$(AmtOweFlds(WhatRev) + 1, 0) = Str$(0)
'    ElseIf ThisAmt# = 0 Then
'      LSet Form$(AmtOweFlds(WhatRev) + 1, 0) = Str$(0)
'    End If
'    SaveField AmtOweFlds(WhatRev) + 1, Form$(), Fld(), BadField
'  Next
'
'  SaveField ChkAmtFld, Form$(), Fld(), BadField
'  SaveField CashAmtFld, Form$(), Fld(), BadField
'
'  CalcFields 0, AmtPadFlds(1), Form$(), Fld()
'  CalcFields 0, CashAmtFld, Form$(), Fld()
'
'  PrintArray 1, NumFlds - 1, Form$(), Fld()
'Return
'
'SaveTransaction:
'  'DisplayUBScrn "UPDATDSK"
'  BCopy SSEG(Form$(0, 0)), SADD(Form$(0, 0)), VARSEG(UBPaymentRec(1)), VarPtr(UBPaymentRec(1)), UBPayRecLen, 0

'  FirstTime = True
'Return
'
'PrintReceipt:
'
'  'SaveScrn TempScrn()
'
'  ListFile = FreeFile
'  Open PayFileName$ For Random Shared As ListFile Len = UBPayRecLen
'  RecptNum& = LOF(ListFile) / UBPayRecLen
'  Get #ListFile, PayListRec&, UBPaymentRec(1)
'  Close
'
'  RHandle = FreeFile
'  Open PayRecpName$ For Output As RHandle
'  Print #RHandle, Chr$(27); "p"; Chr$(0); Chr$(25); Chr$(250)
'  Print #RHandle, Chr$(7)
'  Print #RHandle, TownName$
'  Print #RHandle, "UTILITY PAYMENT"
'  Print #RHandle, "Date: "; PostDate$
'  Print #RHandle,
'  Print #RHandle, "CUSTOMER NAME & DESC. OF PAYMENT"
'  Print #RHandle, UBPaymentRec(1).CustName
'  Print #RHandle, UBPaymentRec(1).CUSTADDR
'  Print #RHandle, UBPaymentRec(1).Desc
'  Print #RHandle, "Acct. No. "; UBPaymentRec(1).CustAcct
'  Print #RHandle,
'  Print #RHandle, "Total Owed: "; Using; "$$####,#.##"; UBPaymentRec(1).TOTOWED
'  Print #RHandle, "Total Paid: "; Using; "$$####,#.##"; UBPaymentRec(1).AMTPAID
'  Print #RHandle, "Change Due: "; Using; "$$####,#.##"; UBPaymentRec(1).CHANGE
'  Print #RHandle, "   Balance: "; Using; "$$####,#.##"; UBPaymentRec(1).TOTOWED - UBPaymentRec(1).AMTPAID
'  Print #RHandle,
'  For RevCnt = 1 To NumOfRevs
'    If UBPaymentRec(1).PaidOwed(RevCnt).AMTPD1 <> 0 Or UBPaymentRec(1).PaidOwed(RevCnt).AMTOWE1 <> 0 Then
'      LSet RecpRev$ = RevText$(RevCnt)
'      'PRINT #RHandle, RecpRev$; USING "$$####,#.##"; UBPaymentRec(1).PaidOwed(RevCnt).AmtOwe1; UBPaymentRec(1).PaidOwed(RevCnt).AmtPd1
'      Print #RHandle, RecpRev$; Using; "$$#####.##"; UBPaymentRec(1).PaidOwed(RevCnt).AMTPD1
'    End If
'  Next
'  Print #RHandle,
'  Print #RHandle, "Operator: "; OPERNUM
'  Print #RHandle, "Receipt#: "; Using; "######"; RecptNum&
'  Print #RHandle,
'  Print #RHandle, "       T H A N K   Y O U !"
'  Print #RHandle,
'  Print #RHandle,
'  Print #RHandle,
'  Print #RHandle,
'  Print #RHandle,
'  Print #RHandle, CutPaper$
'  Close RHandle
'
'  'Shell$ = "type " + PayRecpName$ + " > com2:"
'  'SHELL Shell$
'
'  PrintRptFile Header$, PayRecpName$, RecpPort, RetCode%, 5
'
'  KillFile PayRecpName$
'Return
'
'CWPrintReceipt:
'  PCnt = 0
'  ListFile = FreeFile
'  Open PayFileName$ For Random Shared As ListFile Len = UBPayRecLen
'  Get #ListFile, PayListRec&, UBPaymentRec(1)
'  Close
'
'  RHandle = FreeFile
'  Open PayRecpName$ For Output As RHandle
'  Print #RHandle, Chr$(27); "p"; Chr$(0); Chr$(25); Chr$(250); Chr$(7)
'  Print #RHandle, Tab(7); TownName$; Tab(43); TownName$
'  Print #RHandle, Tab(7); "UTILITY PAYMENT"; Tab(43); "UTILITY PAYMENT"
'  Print #RHandle, Tab(7); "Date: "; PostDate$; Tab(43); "Date: "; PostDate$
'  Print #RHandle,
'  Print #RHandle, Tab(7); "CUSTOMER NAME & DESC."; Tab(43); "CUSTOMER NAME & DESC."
'  Print #RHandle, Tab(7); UBPaymentRec(1).CustName; Tab(43); UBPaymentRec(1).CustName
'  Print #RHandle, Tab(7); UBPaymentRec(1).CUSTADDR; Tab(43); UBPaymentRec(1).CUSTADDR
'  Print #RHandle, Tab(7); UBPaymentRec(1).Desc; Tab(43); UBPaymentRec(1).Desc
'  Print #RHandle, Tab(7); "Acct. No. "; UBPaymentRec(1).CustAcct; Tab(43); "Acct. No. "; UBPaymentRec(1).CustAcct
'  Print #RHandle,
'  Print #RHandle, Using; "Total Owed:   $$####,#.##"; Tab(7); UBPaymentRec(1).TOTOWED; Tab(43); UBPaymentRec(1).TOTOWED
'  Print #RHandle, Using; "Total Paid:   $$####,#.##"; Tab(7); UBPaymentRec(1).AMTPAID; Tab(43); UBPaymentRec(1).AMTPAID
'  Print #RHandle, Using; "Change Due:   $$####,#.##"; Tab(7); UBPaymentRec(1).CHANGE; Tab(43); UBPaymentRec(1).CHANGE
'  Print #RHandle, Using; "   Balance:   $$####,#.##"; Tab(7); UBPaymentRec(1).TOTOWED - UBPaymentRec(1).AMTPAID; Tab(43); UBPaymentRec(1).TOTOWED - UBPaymentRec(1).AMTPAID
'  Print #RHandle,
'  '16 to here
'  For RevCnt = 1 To NumOfRevs
'    If UBPaymentRec(1).PaidOwed(RevCnt).AMTPD1 <> 0 Or UBPaymentRec(1).PaidOwed(RevCnt).AMTOWE1 <> 0 Then
'      LSet RecpRev$ = RevText$(RevCnt)
'      Print #RHandle, Using; RecpRev$ + "$$#####.##"; Tab(7); UBPaymentRec(1).PaidOwed(RevCnt).AMTPD1; Tab(43); UBPaymentRec(1).PaidOwed(RevCnt).AMTPD1
'      PCnt = PCnt + 1
'    End If
'  Next
'  If PCnt < 6 Then
'    For zz = PCnt To 6
'      Print #RHandle,
'    Next
'  End If
'
'  Print #RHandle, Tab(7); "Operator: "; OPERNUM; Tab(43); "Operator: "; OPERNUM
'  RecptNo& = FileSize(PayFileName$) \ UBPayRecLen
'  Print #RHandle, Tab(7); Using; "Receipt:  ######"; RecptNo&; Tab(43); RecptNo&
'  Print #RHandle,
'  Print #RHandle, Tab(7); " T H A N K   Y O U !"; Tab(43); " T H A N K   Y O U !"
'  Print #RHandle,
'  Close RHandle
'
'  'Shell$ = "type " + PayRecpName$ + " > com2:"
'  'SHELL Shell$
'
'  PrintRptFile Header$, PayRecpName$, RecpPort, RetCode%, 5
'
'  KillFile PayRecpName$
'Return
'
'ClearCity:
'  LSet CITY$ = ""
'  LSet Deps$ = ""
'  QPrintRC CITY$, 8, 15, -1
'  QPrintRC Deps$, 9, 15, -1
'Return
'
'CheckPaymentInfo:
'  PaymentOKFlag = True
'
'  TAmtRecv# = Value(Form$(AmtRecvFld, 0), ECode)
'  TAmtPaid# = Value(Form$(AmtPaidFld, 0), ECode)
'   ChangeAmt# = Value(Form$(ChangeFld, 0), ECode)
'
'  If TAmtPaid# = 0 Then
'    OK = MsgBox%("UB.QSL", "BADPYTOT")          'show bad scrn
'    Action = 2
'    PaymentOKFlag = False
'    frm(1).FldNo = frm(1).PrevFld
'    GoTo BadPayment
'  End If
'
'  If TAmtRecv# = Round#(TAmtPaid# + ChangeAmt#) And TAmtRecv# > 0 And ChangeAmt# >= 0 Then
'    PaymentOKFlag = True
'  Else
'    OK = MsgBox%("UB.QSL", "BADPYTOT")          'show bad scrn
'    Action = 2
'    PaymentOKFlag = False
'    frm(1).FldNo = frm(1).PrevFld
'    GoTo BadPayment
'  End If
'
'  TenderType$ = QPTrim$(Form$(TenderFld, 0))
'  If Len(TenderType$) = 0 Then
'    OK = MsgBox%("UB.QSL", "BADTENDR")
'    Action = 2
'    PaymentOKFlag = False
'    frm(1).FldNo = TenderFld
'    GoTo BadPayment
'  End If
'
'BadPayment:
'Return
'End Sub



'Private Sub txtPaymentDate_PopUpClose()
'fpcboTenderType.SetFocus
'End Sub
