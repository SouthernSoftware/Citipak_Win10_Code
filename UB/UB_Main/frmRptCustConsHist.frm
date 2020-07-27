VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "BTN32A20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmRptCustConsHist 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Consumption History"
   ClientHeight    =   8640
   ClientLeft      =   30
   ClientTop       =   540
   ClientWidth     =   12195
   Icon            =   "frmRptCustConsHist.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   12195
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboRptType 
      Height          =   375
      Left            =   5370
      TabIndex        =   3
      Top             =   4995
      Width           =   1920
      _Version        =   196608
      _ExtentX        =   3387
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
      ListWidth       =   3504
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
      ColDesigner     =   "frmRptCustConsHist.frx":08CA
   End
   Begin EditLib.fpText fpCustName 
      Height          =   348
      Left            =   5370
      TabIndex        =   10
      Top             =   3000
      Width           =   4212
      _Version        =   196608
      _ExtentX        =   7429
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
      AlignTextH      =   0
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
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   6
      Top             =   8280
      Width           =   12192
      _ExtentX        =   21511
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7117
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7117
            TextSave        =   "3:59 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7117
            TextSave        =   "5/14/2007"
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
   Begin fpBtnAtlLibCtl.fpBtn fpCmdPrint 
      Height          =   480
      Left            =   7788
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   6504
      Width           =   1332
      _Version        =   131072
      _ExtentX        =   2350
      _ExtentY        =   847
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
      ButtonDesigner  =   "frmRptCustConsHist.frx":0CA4
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdExit 
      Height          =   480
      Left            =   9408
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   6504
      Width           =   1332
      _Version        =   131072
      _ExtentX        =   2350
      _ExtentY        =   847
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
      ButtonDesigner  =   "frmRptCustConsHist.frx":0E81
   End
   Begin EditLib.fpLongInteger fpCustRecNo 
      Height          =   300
      Left            =   1416
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1608
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
   Begin EditLib.fpDateTime txtDate2 
      Height          =   348
      Left            =   5376
      TabIndex        =   2
      Top             =   4464
      Width           =   1692
      _Version        =   196608
      _ExtentX        =   2984
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
   Begin EditLib.fpDateTime txtDate1 
      Height          =   348
      Left            =   5370
      TabIndex        =   1
      Top             =   3960
      Width           =   1692
      _Version        =   196608
      _ExtentX        =   2984
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
   Begin EditLib.fpBoolean fpCompleteFlag 
      Height          =   348
      Left            =   5370
      TabIndex        =   0
      Top             =   3480
      Width           =   324
      _Version        =   196608
      _ExtentX        =   572
      _ExtentY        =   614
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
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      MultiLine       =   0   'False
      AlignTextH      =   1
      AlignTextV      =   0
      ToggleTrue      =   "Yy"
      TextTrue        =   "Y"
      Value           =   1
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
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      BooleanDataType =   0
      OLEDropMode     =   0
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Starting Date:"
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
      Height          =   324
      Left            =   3156
      TabIndex        =   14
      Top             =   3984
      Width           =   2028
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ending Date:"
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
      Height          =   324
      Index           =   0
      Left            =   3348
      TabIndex        =   13
      Top             =   4488
      Width           =   1836
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Print Complete History:"
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
      Height          =   324
      Left            =   2268
      TabIndex        =   12
      Top             =   3528
      Width           =   2916
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Select Report Type: "
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
      Height          =   324
      Index           =   1
      Left            =   2532
      TabIndex        =   9
      Top             =   4992
      Width           =   2724
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   3108
      Left            =   2262
      Top             =   2544
      Width           =   7668
   End
   Begin VB.Label PromptLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Customer:"
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
      Height          =   324
      Left            =   3756
      TabIndex        =   8
      Top             =   3048
      Width           =   1428
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3210
      Top             =   936
      Width           =   5772
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Print Customer Consumption History"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3414
      TabIndex        =   7
      Top             =   1176
      Width           =   5412
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   972
      Left            =   3210
      Top             =   816
      Width           =   5772
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuPrnScn 
         Caption         =   "Prin&t Screen"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmRptCustConsHist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim RecNo As Long, AcctNum As Long
Dim fromform As Form, toform As Form, codeopt As Integer
Public Sub Wheretogo(xfrm As Form, tfrm As Form, Optional opt As Integer)
  Set fromform = xfrm
  Set toform = tfrm
  If opt <> 0 Then
    codeopt = opt
  Else
    codeopt = 0
  End If
End Sub

Private Sub Form_Activate()
  GetName
End Sub
Private Function ValidDate()
  Dim TempDate1 As Integer, TempDate2 As Integer
  If CheckValDate(txtDate1) = False And CheckValDate(txtDate2) = False Then
    MsgBox "Date Is Not Valid. Please Correct.", vbOKOnly, "Invalid Date"
    ValidDate = False
  Else
    TempDate1 = DateDiff("d", "12/31/1979", txtDate1)
    TempDate2 = DateDiff("d", "12/31/1979", txtDate2)
    If TempDate1 > TempDate2 Then
      ValidDate = False
      MsgBox "The Starting And Ending Dates Must Be In Chronological Order Or Equal", vbOKOnly, "Invalid Date"
    Else
      ValidDate = True
    End If
  End If
End Function

Private Sub fpCompleteFlag_Change()
  If fpCompleteFlag.Value = ValueFalse Then
    txtDate1.Enabled = True
    txtDate2.Enabled = True
  Else
    txtDate1.Enabled = False
    txtDate2.Enabled = False
  End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If fpCmdExit.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        UBLog "Closed via RptCustConsHist by " + PWUser$
        CitiTerminate
      End If
    End If
  End If
End Sub

Private Sub fpCompleteFlag_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    If txtDate1.Enabled Then
      txtDate1.SetFocus
    Else
      fpcboRptType.SetFocus
    End If
  End If
End Sub

Private Sub txtDate1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    txtDate2.SetFocus
  End If
End Sub

Private Sub txtDate2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpcboRptType.SetFocus
  End If
End Sub

Private Sub fpcboRptType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboRptType.ListDown = True
  End If
  If fpcboRptType.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpcmdPrint.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        If txtDate2.Enabled = True Then
          txtDate2.SetFocus
        Else
          fpCompleteFlag.SetFocus
        End If
        KeyCode = 0
      End If
    End If
  End If
End Sub


Private Sub mnuExit_Click()
  fpCmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub

Private Sub fpCmdExit_Click()
 ' Load frmUBCustMenu
  DoEvents
  If codeopt = 1 Then
    ActivateControls frmCustEditLookUP
  ElseIf codeopt = 2 Then
    ActivateControls frmDisplayList
  End If

 ' frmUBCustMenu.Show
  Unload frmRptCustConsHist
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      KeyCode = 0
      Call fpCmdExit_Click
    Case vbKeyF10
      KeyCode = 0
      Call fpcmdPrint_Click
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
  txtDate1.Text = Format(Now, "mm/dd/yyyy")
  txtDate2.Text = Format(Now, "mm/dd/yyyy")
  fpCompleteFlag = ValueTrue
  txtDate1.Enabled = False
  txtDate2.Enabled = False
  fpcboRptType.InsertRow = "Graphics"
  fpcboRptType.InsertRow = "Text"
  fpcboRptType.ListIndex = 0
  Me.HelpContextID = hlpCustomerConsumpt
  'GetName
End Sub
Private Sub GetName()
  ReDim UBCustRec(1) As NewUBCustRecType
  Dim UBCustRecLen As Integer, UBSetupLen As Integer, UBCust As Integer
  RecNo& = fpCustRecNo
  UBCustRecLen = Len(UBCustRec(1))
  UBCust = FreeFile
  Open UBCustFile For Random Shared As UBCust Len = UBCustRecLen
  Get #UBCust, RecNo&, UBCustRec(1)
  Close UBCust
  fpCustName = UBCustRec(1).CustName
End Sub


Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    'Me.Visible = False
    Temp_Class.ResizeControls Me
    'Me.Visible = True
   ' Me.SetFocus
  End If
  DoEvents
End Sub
Private Sub fpcmdPrint_Click()
  If fpCompleteFlag.Value = ValueFalse Then
    If Not ValidDate Then
      Exit Sub
    End If
  End If
  DeActivateControls Me, True
  If fpcboRptType.ListIndex = 0 Then
    'do graphics
    CustConsumpHistRpt2
  ElseIf fpcboRptType.ListIndex = 1 Then
    'do text
    CustConsumpHistRpt
    ActivateControls Me, True
  Else
    ActivateControls Me, True
  End If
  'fpCmdExit_Click
End Sub


'***************************************
Private Sub CustConsumpHistRpt()
  Dim Dash80 As String, F As String
  ReDim UBCustRec(1) As NewUBCustRecType
  ReDim UBTranRec(1) As UBTransRecType
  ReDim UBSetUpRec(1) As UBSetupRecType
  Dim DidCnt As Long, CCnt As Integer, LineCnt As Integer
  Dim UBCustRecLen As Integer, UBTranRecLen As Integer
  Dim UBSetupLen As Integer, RevCnt As Integer, Rev2Flag As Integer
  Dim UBCust As Integer, UBRpt As Integer, UBTran As Integer
  Dim TroyFlag As Integer, AbortFlag As Integer
  Dim MCnt As Integer, ReportFile As String, MaxLines As Integer
  Dim ThisTrans As Long, MaxMeterAmt As Long, Completeflag As Boolean
  Dim MeterType As String, TotalConsp As Double, startdate As Integer
  Dim EstCnt As Integer, CubMeter As Integer, EndDate As Integer
  Dim MTRMulti As Double, MeterConsp As Double
  Dim FoundMtrTyp As Boolean
  Dim DidAMeter As Integer, EstFlag As Integer
  Dim MtrCnt As Integer
  FoundMtrTyp = False
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen
  Completeflag = frmRptCustConsHist.fpCompleteFlag.Text = "Y"
  If Not Completeflag Then
    startdate = Date2Num(txtDate1.Text)
    EndDate = Date2Num(txtDate2.Text)
  End If

  ReportFile$ = UBPath$ + "UBCONSMP.RPT"
  
'  If InStr(UBSetUpRec(1).UTILNAME, "TROY") > 0 Then
'    TroyFlag = True
'  End If
  MaxLines = 55
  If RecNo& = 0 Then
    GoTo ExitConsumpHist
  End If

  Dash80$ = String$(80, "-")
    
  UBTranRecLen = Len(UBTranRec(1))
  UBCustRecLen = Len(UBCustRec(1))
  
  UBCust = FreeFile
  Open UBCustFile For Random Shared As UBCust Len = UBCustRecLen
  Get #UBCust, RecNo&, UBCustRec(1)
  Close UBCust
  
  UBRpt = FreeFile
  Open ReportFile$ For Output As UBRpt

  UBTran = FreeFile
  Open UBPath$ + "UBTRANS.DAT" For Random Shared As UBTran Len = UBTranRecLen

  GoSub DoConsRptHeader

  ThisTrans& = UBCustRec(1).LastTrans

  Do While ThisTrans& > 0
    Get #UBTran, ThisTrans&, UBTranRec(1)
    If UBTranRec(1).TransType = TranUtilityBill Then
      If Not Completeflag Then
        If UBTranRec(1).TransDate >= startdate And UBTranRec(1).TransDate <= EndDate Then
          If LineCnt >= MaxLines Then
            GoSub DoConsRptHeader
          End If
          GoSub PrintConsDetail
          DidCnt = DidCnt + 1
'          If DidCnt = 12 Then
'            Exit Do
'          End If
        End If
      Else
          If LineCnt >= MaxLines Then
            GoSub DoConsRptHeader
          End If

        GoSub PrintConsDetail
        DidCnt = DidCnt + 1
        'If DidCnt = 12 Then
         ' Exit Do
        'End If
      End If
    End If
    ThisTrans& = UBTranRec(1).PrevTrans
  Loop
  GoSub DoConsFooter

  Close
  If DidCnt > 0 Then
  If Not AbortFlag Then
    ViewPrint ReportFile$, "Customer Consumption Report."
    'PrintRptFile "Customer Consumption Report.", "UBCONSMP.RPT", 1, RetCode, EntryPoint
  End If
  Else
    MsgBox "No Transactions To Report", vbOKOnly, "No Transactions"
  End If


ExitConsumpHist:
Exit Sub

PrintConsDetail:
  
  DidAMeter = False
  EstFlag = False
  For EstCnt = 1 To 7
    If UBTranRec(1).ESTREAD(EstCnt) = "Y" Then
      EstFlag = True
      Exit For
    End If
  Next
  For MtrCnt = 1 To 7
    MTRMulti# = 0
    FoundMtrTyp = False
'    For MCnt = 1 To 7
'    '06-09-04 changed to use first meter because of Troy old data
'      If UBTranRec(1).MtrTypes(MtrCnt) = GetCustMeterType%(UBCustRec(), MCnt) Then
'        MTRMulti# = UBCustRec(1).LocMeters(MCnt).MTRMulti
'        FoundMtrTyp = True
'        Exit For
'      End If
'    Next
'    If Not FoundMtrTyp Then
      MTRMulti# = UBCustRec(1).LocMeters(MtrCnt).MTRMulti
'    End If
    If MTRMulti# <= 0 Then
      'If TroyFlag Then
      '  MTRMulti# = 100
      'Else
        MTRMulti# = 1
      'End If
    End If

    If UBTranRec(1).MtrTypes(MtrCnt) <> 0 Then
      DidAMeter = True
      Select Case UBTranRec(1).MtrTypes(MtrCnt)
      Case MtrWaterOnly
        MeterType$ = "      Water"
        F$ = "W"
      Case MtrSewerOnly
        MeterType$ = "      Sewer"
        F$ = "S"
      Case MtrCombined
        MeterType$ = "Water/Sewer"
        F$ = "C"
      Case MtrElectric
        MeterType$ = "   Electric"
        F$ = "E"
      Case MtrDemand
        MeterType$ = " D Electric"
        F$ = "D"
      Case MtrIrrigation
        MeterType$ = " Irrigation"
        F$ = "I"
      Case MtrGas
        MeterType$ = "  Gas Meter"
        F$ = "G"
      Case MtrTouchRead
        MeterType$ = " Touch Read"
        F$ = "T"
      Case MtrLightsService
        MeterType$ = "  L Service"
      Case Else
        MeterType$ = "  ?????????"
      End Select
      For CCnt = 1 To 7
        If UBCustRec(1).LocMeters(CCnt).MTRType = F$ Then
          If UBCustRec(1).LocMeters(CCnt).MtrUnit = "C" Then
            CubMeter = True
          Else
            CubMeter = False
          End If
          Exit For
        End If
      Next
      GoSub PrintThisMeter
    End If
  Next
  If Not DidAMeter Then
    MeterType$ = "        "
    MtrCnt = 1
    GoSub PrintThisMeter
  End If

Return

PrintThisMeter:

  Print #UBRpt, Num2Date(UBTranRec(1).TransDate);
  If EstFlag Then
    Print #UBRpt, "*E";
  End If
  Print #UBRpt, Tab(19); MeterType$;
  Print #UBRpt, Tab(34); Using$("##########", UBTranRec(1).CurRead(MtrCnt));
  Print #UBRpt, Tab(46); Using$("##########", UBTranRec(1).PrevRead(MtrCnt));
  MeterConsp# = UBTranRec(1).CurRead(MtrCnt) - UBTranRec(1).PrevRead(MtrCnt)
  If MeterConsp# < 0 Then
    MaxMeterAmt& = 10& ^ (Len(Str$(UBTranRec(1).PrevRead(MtrCnt))) - 1)
    MeterConsp# = (MaxMeterAmt& - UBTranRec(1).PrevRead(MtrCnt)) + UBTranRec(1).CurRead(MtrCnt)
  End If
  MeterConsp# = MeterConsp# * MTRMulti#
  If CubMeter Then
    MeterConsp# = MeterConsp# * 7.481
  End If
  Print #UBRpt, Tab(56); Using$("##########", MeterConsp#);
  If UBTranRec(1).ReadDate <= 0 Then
    Print #UBRpt, "     ??-??-????"
  Else
    Print #UBRpt, "     "; Num2Date$(UBTranRec(1).ReadDate) '; "!"; UBTranRec(1).EstRead(MtrCnt); "!"
  End If
  LineCnt = LineCnt + 1
  TotalConsp# = TotalConsp# + MeterConsp#

Return

DoConsRptHeader:
  Print #UBRpt,
  Print #UBRpt, Tab(28); "Consumption History Report. "
  Print #UBRpt,
  If Not Completeflag Then
    Print #UBRpt, "Date Range: " + txtDate1.Text + " - " + txtDate2.Text
  Else
    Print #UBRpt, "Complete History"
  End If
  Print #UBRpt, " Account:"; RecNo&
  Print #UBRpt, "Ser Addr: "; UBCustRec(1).ServAddr
  Print #UBRpt, "Location: "; QPTrim$(UBCustRec(1).Book); "-"; QPTrim$(UBCustRec(1).SEQNUMB)
  Print #UBRpt, "Customer: "; UBCustRec(1).CustName; Tab(57); "Report Date: "; Date$
  Print #UBRpt,
  Print #UBRpt, "Transaction                         Current   Previous"
  Print #UBRpt, "   Date            Meter Type       Reading    Reading       Usage    ReadDate"
  Print #UBRpt, Dash80$
  LineCnt = 12
Return

DoConsFooter:
  If DidCnt > 0 Then
    Print #UBRpt, Dash80$
    Print #UBRpt, "Average Consumption: "; Using$("#########", TotalConsp# / DidCnt)
  Else
    Print #UBRpt, "NO TRANSACTIONS!!!"
    Print #UBRpt, Dash80$
  End If
Return
End Sub

Private Sub CustConsumpHistRpt2() 'Graphics report
  Dim F As String, ToPrint As String, ToPrintH As String
  ReDim UBCustRec(1) As NewUBCustRecType
  ReDim UBTranRec(1) As UBTransRecType
  ReDim UBSetUpRec(1) As UBSetupRecType
  Dim DidCnt As Long, CCnt As Integer
  Dim UBCustRecLen As Integer, UBTranRecLen As Integer
  Dim UBSetupLen As Integer, RevCnt As Integer, Rev2Flag As Integer
  Dim UBCust As Integer, UBRpt As Integer, UBTran As Integer
  Dim startdate As Integer, EndDate As Integer, Completeflag As Boolean
  Dim MCnt As Integer, PrnLoc As String, PrnSvcAddr As String
  Dim ThisTrans As Long, MaxMeterAmt As Long
  Dim MeterType As String, ReportFile As String
  Dim EstCnt As Integer, CubMeter As Integer
  Dim MTRMulti As Double, MeterConsp As Double, TotalConsp As Double
  Dim FoundMtrTyp As Boolean
  Dim DidAMeter As Integer, EstFlag As Integer
  Dim MtrCnt As Integer
  FoundMtrTyp = False
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen
  Completeflag = frmRptCustConsHist.fpCompleteFlag.Text = "Y"
  If Not Completeflag Then
    startdate = Date2Num(txtDate1.Text)
    EndDate = Date2Num(txtDate2.Text)
  End If

  ReportFile$ = UBPath$ + "UBCONSMP.RPT"
  

  If RecNo& = 0 Then
    GoTo ExitConsumpHist
  End If

    
  UBTranRecLen = Len(UBTranRec(1))
  UBCustRecLen = Len(UBCustRec(1))
  
  UBCust = FreeFile
  Open UBCustFile For Random Shared As UBCust Len = UBCustRecLen
  Get #UBCust, RecNo&, UBCustRec(1)
  Close UBCust
  
  UBRpt = FreeFile
  Open ReportFile$ For Output As UBRpt

  UBTran = FreeFile
  Open UBPath$ + "UBTRANS.DAT" For Random Shared As UBTran Len = UBTranRecLen

  GoSub DoConsRptHeader

  ThisTrans& = UBCustRec(1).LastTrans

  Do While ThisTrans& > 0
    Get #UBTran, ThisTrans&, UBTranRec(1)
    If UBTranRec(1).TransType = TranUtilityBill Then
      If Not Completeflag Then
        If UBTranRec(1).TransDate >= startdate And UBTranRec(1).TransDate <= EndDate Then
          GoSub PrintConsDetail
          DidCnt = DidCnt + 1
'          If DidCnt = 12 Then
'            Exit Do
'          End If
        End If
      Else
        GoSub PrintConsDetail
        DidCnt = DidCnt + 1
        'If DidCnt = 12 Then
          'Exit Do
        'End If
      End If
    End If
    ThisTrans& = UBTranRec(1).PrevTrans
  Loop
  GoSub DoConsFooter

  Close
  If DidCnt > 0 Then
    'ViewPrint ReportFile$, "Customer Consumption Report."
    'PrintRptFile "Customer Consumption Report.", "UBCONSMP.RPT", 1, RetCode, EntryPoint
  Load frmLoadingRpt
  frmLoadingRpt.setwherefrom frmRptCustConsHist
  ARptCustConsHist.txtDate = Now
  ARptCustConsHist.txtTown = TOWNNAME$
  If Completeflag Then
    ARptCustConsHist.txtComplete = "Complete History"
  Else
    ARptCustConsHist.txtComplete = "Date Range: " + txtDate1.Text + " - " + txtDate2.Text
  End If
  ARptCustConsHist.Title = "Customer Consumption History"
  ARptCustConsHist.totAvg = Using$("#########", TotalConsp# / DidCnt)
  ARptCustConsHist.txtLocation = PrnLoc$
  ARptCustConsHist.lblSerAddr = PrnSvcAddr$
  ARptCustConsHist.txtAcct = Str(RecNo&)
  ARptCustConsHist.GetName ReportFile$
  ARptCustConsHist.startrpt
  Else
    ActivateControls Me
    MsgBox "No Transactions To Report", vbOKOnly, "No Transactions"
  End If

ExitConsumpHist:
Exit Sub

PrintConsDetail:
  
  DidAMeter = False
  EstFlag = False
  For EstCnt = 1 To 7
    If UBTranRec(1).ESTREAD(EstCnt) = "Y" Then
      EstFlag = True
      Exit For
    End If
  Next
  For MtrCnt = 1 To 7
    MTRMulti# = 0
    FoundMtrTyp = False
'    For MCnt = 1 To 7
'    '06-09-04 Change to use first meter only because of Troy old history
'      If UBTranRec(1).MtrTypes(MtrCnt) = GetCustMeterType%(UBCustRec(), MCnt) Then
'        MTRMulti# = UBCustRec(1).LocMeters(MCnt).MTRMulti
'        FoundMtrTyp = True
'        Exit For
'      End If
'    Next
'    If Not FoundMtrTyp Then
      MTRMulti# = UBCustRec(1).LocMeters(MtrCnt).MTRMulti
'    End If
    If MTRMulti# <= 0 Then
        MTRMulti# = 1
    End If

    If UBTranRec(1).MtrTypes(MtrCnt) <> 0 Then
      DidAMeter = True
      Select Case UBTranRec(1).MtrTypes(MtrCnt)
      Case MtrWaterOnly
        MeterType$ = "      Water"
        F$ = "W"
      Case MtrSewerOnly
        MeterType$ = "      Sewer"
        F$ = "S"
      Case MtrCombined
        MeterType$ = "Water/Sewer"
        F$ = "C"
      Case MtrElectric
        MeterType$ = "   Electric"
        F$ = "E"
      Case MtrDemand
        MeterType$ = " D Electric"
        F$ = "D"
      Case MtrIrrigation
        MeterType$ = " Irrigation"
        F$ = "I"
      Case MtrGas
        MeterType$ = "  Gas Meter"
        F$ = "G"
      Case MtrTouchRead
        MeterType$ = " Touch Read"
        F$ = "T"
      Case MtrLightsService
        MeterType$ = "  L Service"
      Case Else
        MeterType$ = "  ?????????"
      End Select
      For CCnt = 1 To 7
        If UBCustRec(1).LocMeters(CCnt).MTRType = F$ Then
          If UBCustRec(1).LocMeters(CCnt).MtrUnit = "C" Then
            CubMeter = True
          Else
            CubMeter = False
          End If
          Exit For
        End If
      Next
      GoSub PrintThisMeter
    End If
  Next
  If Not DidAMeter Then
    MeterType$ = "        "
    MtrCnt = 1
    GoSub PrintThisMeter
  End If

Return

PrintThisMeter:

  ToPrint$ = Num2Date(UBTranRec(1).TransDate)
  If EstFlag Then
    ToPrint$ = ToPrint$ + " *E"
  End If
  ToPrint$ = ToPrint$ + "~" + MeterType$
  ToPrint$ = ToPrint$ + "~" + Using$("##########", UBTranRec(1).CurRead(MtrCnt))
  ToPrint$ = ToPrint$ + "~" + Using$("##########", UBTranRec(1).PrevRead(MtrCnt))
  MeterConsp# = UBTranRec(1).CurRead(MtrCnt) - UBTranRec(1).PrevRead(MtrCnt)
  If MeterConsp# < 0 Then
    MaxMeterAmt& = 10& ^ (Len(Str$(UBTranRec(1).PrevRead(MtrCnt))) - 1)
    MeterConsp# = (MaxMeterAmt& - UBTranRec(1).PrevRead(MtrCnt)) + UBTranRec(1).CurRead(MtrCnt)
  End If
  MeterConsp# = MeterConsp# * MTRMulti#
  If CubMeter Then
    MeterConsp# = MeterConsp# * 7.481
  End If
  ToPrint$ = ToPrint$ + "~" + Using$("##########", MeterConsp#)
  If UBTranRec(1).ReadDate <= 0 Then
    ToPrint$ = ToPrint$ + "~" + "     ??-??-????"
  Else
    ToPrint$ = ToPrint$ + "~" + "     " + Num2Date$(UBTranRec(1).ReadDate) '; "!"; UBTranRec(1).EstRead(MtrCnt); "!"
  End If

  TotalConsp# = TotalConsp# + MeterConsp#
  Print #UBRpt, ToPrintH$ + "~" + ToPrint$
  ToPrint$ = ""
Return

DoConsRptHeader:
  ToPrintH$ = QPTrim(UBCustRec(1).CustName)
  'ToPrintH$ = ToPrintH$ + "~" + Str(RecNo&)
  PrnSvcAddr$ = QPTrim(UBCustRec(1).ServAddr)
  PrnLoc$ = QPTrim$(UBCustRec(1).Book) + "-" + QPTrim$(UBCustRec(1).SEQNUMB)
 
Return

DoConsFooter:
'  If DidCnt > 0 Then
'    Print #UBRpt, Dash80$
'    Print #UBRpt, "Average Consumption: "; Using$("#########", TotalConsp# / DidCnt)
'  Else
'    Print #UBRpt, "NO TRANSACTIONS!!!"
'    Print #UBRpt, Dash80$
'  End If
Return
End Sub

