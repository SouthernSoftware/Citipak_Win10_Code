VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmPenaltyCalculation 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Apply Penalty/Late Fees"
   ClientHeight    =   8640
   ClientLeft      =   36
   ClientTop       =   540
   ClientWidth     =   12192
   Icon            =   "frmPenatlyCalculation.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   12192
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboRevenues 
      Height          =   348
      Left            =   5604
      TabIndex        =   2
      Top             =   2460
      Width           =   3612
      _Version        =   196608
      _ExtentX        =   6371
      _ExtentY        =   614
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.2
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
      EditMarginLeft  =   2
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmPenatlyCalculation.frx":08CA
   End
   Begin LpLib.fpCombo fpcboBalType 
      Height          =   348
      Left            =   5604
      TabIndex        =   3
      Top             =   2868
      Width           =   3540
      _Version        =   196608
      _ExtentX        =   6244
      _ExtentY        =   614
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.2
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
      AutoSearchFill  =   0   'False
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
      ColDesigner     =   "frmPenatlyCalculation.frx":0BED
   End
   Begin LpLib.fpCombo fpcboWhichever 
      Height          =   348
      Left            =   5604
      TabIndex        =   6
      Top             =   4116
      Width           =   1812
      _Version        =   196608
      _ExtentX        =   3196
      _ExtentY        =   614
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.2
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
      ColDesigner     =   "frmPenatlyCalculation.frx":0F10
   End
   Begin EditLib.fpDoubleSingle fpdblPercent 
      Height          =   348
      Left            =   5604
      TabIndex        =   4
      Top             =   3288
      Width           =   1452
      _Version        =   196608
      _ExtentX        =   2561
      _ExtentY        =   614
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.8
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
      HideSelection   =   0   'False
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
   Begin VB.TextBox txtDescription 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   5598
      TabIndex        =   1
      Top             =   2040
      Width           =   2652
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Esc E&xit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   10080
      TabIndex        =   13
      Top             =   7560
      Width           =   1332
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "F10 &Ok"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   8400
      TabIndex        =   12
      Top             =   7560
      Width           =   1332
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   14
      Top             =   8280
      Width           =   12192
      _ExtentX        =   21505
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7133
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7133
            TextSave        =   "11:12 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7133
            TextSave        =   "7/11/2003"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin EditLib.fpDateTime txtDate1 
      Height          =   348
      Left            =   5604
      TabIndex        =   0
      Top             =   1632
      Width           =   1692
      _Version        =   196608
      _ExtentX        =   2984
      _ExtentY        =   614
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.2
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
   Begin EditLib.fpText fptxtCycle2 
      Height          =   348
      Left            =   5592
      TabIndex        =   9
      Top             =   5496
      Width           =   636
      _Version        =   196608
      _ExtentX        =   1122
      _ExtentY        =   614
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.8
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
      AutoCase        =   0
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
      CharValidationText=   "0123456789"
      MaxLength       =   2
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
   Begin EditLib.fpText fptxtCycle1 
      Height          =   348
      Left            =   5592
      TabIndex        =   8
      Top             =   5076
      Width           =   636
      _Version        =   196608
      _ExtentX        =   1122
      _ExtentY        =   614
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.8
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
      AutoCase        =   0
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
      CharValidationText=   "0123456789"
      MaxLength       =   2
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
   Begin EditLib.fpCurrency fpMinBal 
      Height          =   348
      Left            =   5604
      TabIndex        =   7
      Top             =   4524
      Width           =   1476
      _Version        =   196608
      _ExtentX        =   2603
      _ExtentY        =   614
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.8
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
   Begin EditLib.fpText fptxtBook2 
      Height          =   348
      Left            =   5592
      TabIndex        =   11
      Top             =   6660
      Width           =   636
      _Version        =   196608
      _ExtentX        =   1122
      _ExtentY        =   614
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.8
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
      AutoCase        =   0
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
      CharValidationText=   "0123456789"
      MaxLength       =   2
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
   Begin EditLib.fpText fptxtBook1 
      Height          =   348
      Left            =   5592
      TabIndex        =   10
      Top             =   6240
      Width           =   636
      _Version        =   196608
      _ExtentX        =   1122
      _ExtentY        =   614
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.8
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
      AutoCase        =   0
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
      CharValidationText=   "0123456789"
      MaxLength       =   2
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
   Begin EditLib.fpCurrency fpAmount 
      Height          =   348
      Left            =   5604
      TabIndex        =   5
      Top             =   3696
      Width           =   1476
      _Version        =   196608
      _ExtentX        =   2603
      _ExtentY        =   614
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.8
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
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   5652
      Left            =   2478
      Top             =   1512
      Width           =   7236
   End
   Begin VB.Line Line1 
      X1              =   2508
      X2              =   9720
      Y1              =   4992
      Y2              =   4992
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Book Number to Start: "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   3240
      TabIndex        =   28
      Top             =   6228
      Width           =   2220
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Penalty Revenue Source:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Index           =   8
      Left            =   3060
      TabIndex        =   27
      Top             =   2520
      Width           =   2340
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cycle Number to End:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Index           =   7
      Left            =   3216
      TabIndex        =   26
      Top             =   5544
      Width           =   2172
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Book Number to End:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Index           =   6
      Left            =   3288
      TabIndex        =   25
      Top             =   6696
      Width           =   2100
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cycle Number to Start:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Index           =   5
      Left            =   3144
      TabIndex        =   24
      Top             =   5136
      Width           =   2244
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Minimum Balance:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Index           =   4
      Left            =   3564
      TabIndex        =   23
      Top             =   4620
      Width           =   1836
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Penalty Date:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   3732
      TabIndex        =   22
      Top             =   1704
      Width           =   1668
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Enter % to Charge:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Index           =   3
      Left            =   3492
      TabIndex        =   21
      Top             =   3360
      Width           =   1908
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Whichever Is:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Index           =   2
      Left            =   3900
      TabIndex        =   20
      Top             =   4200
      Width           =   1500
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Amount to Charge:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Index           =   0
      Left            =   2988
      TabIndex        =   19
      Top             =   3780
      Width           =   2412
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "---OR----"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Index           =   1
      Left            =   3936
      TabIndex        =   18
      Top             =   5904
      Width           =   1164
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Description of Charge:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Index           =   0
      Left            =   3276
      TabIndex        =   17
      Top             =   2100
      Width           =   2124
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3192
      Top             =   312
      Width           =   5772
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Apply Penaly/Late Fees"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3600
      TabIndex        =   16
      Top             =   552
      Width           =   5004
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Charge Penalty On:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Index           =   1
      Left            =   3516
      TabIndex        =   15
      Top             =   2940
      Width           =   1884
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   972
      Left            =   3192
      Top             =   192
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
Attribute VB_Name = "frmPenaltyCalculation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim CycleFlag As Boolean, OkFlag As Boolean, BadDate As Boolean
Private Sub cmdExit_Click()
  Load frmUBPenaltyMenu
  DoEvents
  frmUBPenaltyMenu.Show
  Unload frmPenaltyCalculation
  DoEvents
End Sub

Private Sub cmdOk_Click()
  CheckFields
  If OkFlag Then
    PenaltyProcess
    MsgBox "Penalty Calculation Is Complete.", vbOKOnly, "Procedure Complete"
    cmdExit_Click
  End If
End Sub

Private Sub txtDate1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    txtDescription.SetFocus
  End If
End Sub
Private Sub txtDescription_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
    fpcboRevenues.SetFocus
  ElseIf KeyCode = vbKeyUp Then
    txtDate1.SetFocus
  End If
End Sub
Private Sub fpcboRevenues_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboRevenues.ListDown = True
  End If
  If fpcboRevenues.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpcboBalType.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        txtDescription.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub
Private Sub fpcboBalType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboBalType.ListDown = True
  End If
  If fpcboBalType.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpdblPercent.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpcboRevenues.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub
Private Sub fpdblPercent_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpAmount.SetFocus
  End If
End Sub
Private Sub fpAmount_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpcboWhichever.SetFocus
  End If
End Sub
Private Sub fpcboWhichever_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboWhichever.ListDown = True
  End If
  If fpcboWhichever.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpMinBal.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpAmount.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub

Private Sub fpMinBal_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    If fptxtCycle1.Enabled = True Then
      fptxtCycle1.SetFocus
    Else
      fptxtBook1.SetFocus
    End If
  End If
End Sub
Private Sub fptxtCycle1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fptxtCycle2.SetFocus
  End If
End Sub
Private Sub fptxtCycle2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    cmdOk.SetFocus
  End If
End Sub

Private Sub fptxtbook1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fptxtBook2.SetFocus
  End If
End Sub
Private Sub fptxtbook2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    cmdOk.SetFocus
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
      If MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        'ClearInUse PWcnt
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
      Call cmdExit_Click
    Case vbKeyF10:
      KeyCode = 0
      DoEvents
      Call cmdOk_Click
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Dim UBSetupreclen As Integer
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TownName$
  ReDim UBSetUp(1) As UBSetupRecType
  UBSetupreclen = Len(UBSetUp(1))
  LoadUBSetUpFile UBSetUp(), UBSetupreclen

  If UBSetUp(1).BILLCYCL = "Y" Then
    CycleFlag = True
  End If
      If CycleFlag Then
        fptxtBook1 = "0"
        fptxtBook2 = "0"
        fptxtBook1.Enabled = False
        fptxtBook2.Enabled = False
      Else
        fptxtCycle1 = "0"
        fptxtCycle2 = "0"
        fptxtCycle1.Enabled = False
        fptxtCycle2.Enabled = False
      End If
  txtDate1.Text = Format(Now, "mm/dd/yyyy")
  fpcboBalType.AddItem "Current Balance"
  fpcboBalType.AddItem "Previous Balance"
  fpcboBalType.AddItem "Total Balance"
  fpcboBalType.ListIndex = 0
  fpcboWhichever.AddItem " "
  fpcboWhichever.AddItem "Less"
  fpcboWhichever.AddItem "Greater"
  fpcboWhichever.ListIndex = 0
  FillRevList fpcboRevenues
  fpcboRevenues.RemoveItem 0
  fpcboRevenues.ListIndex = -1
End Sub

Private Sub Form_Resize()
  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
  End If
End Sub
Private Sub CheckPostDate()
Dim PenDate As String
  PenDate$ = txtDate1.Text
  If Val(Left$(PenDate$, 2)) < 1 Or Val(Left$(PenDate$, 2)) > 12 Then
    If Val(Mid$(PenDate$, 4, 2)) < 1 Or Val(Mid$(PenDate$, 4, 2)) > 31 Then
      BadDate = True
    Else
      BadDate = False
    End If
  Else
    BadDate = False
  End If
End Sub
Private Function CheckFields()
  ReDim MsgText(0 To 5) As String
  Dim FntSize As Integer
  MsgText(2) = ""

  CheckPostDate
  If BadDate = True Then
    MsgText(3) = "You Have Not Entered a Proper"
    MsgText(4) = "Transaction Date!"
  ElseIf fpcboRevenues.ListIndex = -1 Then
    MsgText(3) = "No Revenue Source for applied penalty."
    MsgText(4) = "Correct and try again."
  ElseIf fpcboBalType.ListIndex = -1 Then
    MsgText(3) = "Must identify Balance source to apply."
    MsgText(4) = "penalty on."
  ElseIf fpdblPercent.DoubleValue <> 0 And fpAmount.DoubleValue <> 0 And fpcboWhichever.ListIndex < 1 Then
    MsgText(2) = "MAY NOT have Both a percentage"
    MsgText(3) = "and a fixed amount WITHOUT the"
    MsgText(4) = "'Whichever is:' parameter!"
  ElseIf fpcboWhichever.ListIndex > 0 And (fpdblPercent.DoubleValue = 0 Or fpAmount.DoubleValue = 0) Then
    MsgText(3) = "Invalid percentage, or Fixed amount"
    MsgText(4) = "with the 'Whichever is:' parameter!"
  ElseIf fpdblPercent.DoubleValue = 0 And fpAmount.DoubleValue = 0 Then
    MsgText(2) = "No Penalty Would Calculate Because"
    MsgText(3) = "BOTH the Percentage and the Amount are"
    MsgText(4) = "SET TO ZERO!"
  ElseIf Val(fptxtCycle1) = 0 And Val(fptxtCycle2) = 0 And Val(fptxtBook1) = 0 And Val(fptxtBook2) = 0 Or Val(fptxtCycle1) < Val(fptxtCycle2) Or Val(fptxtBook1) < Val(fptxtBook2) Then
    MsgText(3) = "Invalid Book or Cycle range!"
    MsgText(4) = "Correct and try again."
  Else
    OkFlag = True
  End If
  If Not OkFlag Then
    frmMsgDialog.RetLabel = "-2"
    FntSize = frmMsgDialog.Label(2).FontSize
    frmMsgDialog.Label(1).FontSize = (FntSize + 2)
    frmMsgDialog.Label(2).FontSize = (FntSize + 2)
    frmMsgDialog.Label(3).FontSize = (FntSize + 2)
    MsgText(0) = "ERROR:"
    MsgText(1) = ""
    MsgText(5) = ""
    GetOKorNot MsgText(), True
  End If

End Function

Private Sub PenaltyProcess()
  Dim UBCustRecLen As Integer, UBTranRecLen As Integer, UBSetupLen As Integer
  Dim Pct2ChgFld  As Integer, TennFlag As Boolean, PolockFlag As Boolean
  Dim HamFlag As Boolean, TuckFlag As Boolean, CashFlag As Boolean
  Dim SunSetFlag As Boolean, CycleFlag As Boolean, cnt As Integer
  Dim UseCycle As String, NumOfRevs As Integer, TempRev As String
  Dim PenFile As String, PHandle As Integer, CHandle As Integer
  Dim NumCustRecs As Long, CCnt As Long, MinBalance As Double
  Dim UsingBook As Boolean, ThisBook As Integer, FirstBook As Integer
  Dim LastBook As Integer, PenaltyDate As Integer, UsePrevFlag As Boolean
  Dim UseCurrFlag As Boolean, PctAmt As Double, FixAmt As Double
  Dim GreaterFlag As Boolean, UseBothFlag As Boolean, UsePctFlag As Boolean
  Dim RevSource As Integer, FirstCycle As Integer, LastCycle As Integer
  Dim UsingCycle As Boolean, TransDesc As String, PenBal As Double
  Dim CustPctPenalty As Double, CustFixPenalty As Double, PCnt As Integer
  Dim CustPenalty As Double, ThisCycle As Integer, TotalBalance As Double
  Dim thandle As Integer, PrevTranRec As Long, NOPenFlag As Boolean
  Dim hand2 As Integer, ExitFlag As Boolean
  UBLog " IN: Create Penalty File (CPF)"

  Pct2ChgFld = 5

  'SHARED Choice$()

  'ReDim TempScrn(0)
  ReDim Source$(15)

  ReDim PenaltyInfo(1) As PenaltyInfoType
  ReDim UBCustRec(1) As NewUBCustRecType
  ReDim UBSetUpRec(1) As UBSetupRecType
  ReDim UBTranRec(1 To 3) As UBTransRecType
  ReDim TaxAmt(1 To 15) As Double

  UBCustRecLen = Len(UBCustRec(1))
  UBTranRecLen = Len(UBTranRec(1))
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen
  TownName$ = UBSetUpRec(1).UTILNAME
  UseCycle$ = UBSetUpRec(1).BILLCYCL
  FrmShowPctComp.Label1 = "Calculating Penalties"
  FrmShowPctComp.Show , Me

  If OkFlag Then
    PenaltyDate = Date2Num%(txtDate1)
    MinBalance# = Val(fpMinBal)
    If MinBalance# < 0 Then
      MinBalance# = 0
    End If
    PenaltyInfo(1).PenDate = PenaltyDate
    PenaltyInfo(1).MinBalance = MinBalance#
    Select Case fpcboBalType.ListIndex
    Case 0   'Applying to Current
      UsePrevFlag = False
      UseCurrFlag = True
    Case 1    'Applying to Previous
      UsePrevFlag = True
      UseCurrFlag = False
    Case 2    'Applying to Both
      UsePrevFlag = True
      UseCurrFlag = True
    End Select
    PenaltyInfo(1).ChargeOn = QPTrim$(fpcboBalType.Text)
    'Get percent or fixed amount
    PctAmt# = fpdblPercent.DoubleValue
    FixAmt# = fpAmount.DoubleValue
    PenaltyInfo(1).PctCharge = PctAmt#
    PenaltyInfo(1).AmtCharge = FixAmt#

    If fpcboWhichever.ListIndex > 0 Then
      PenaltyInfo(1).GreatLess = Left$(fpcboWhichever.Text, 1)
      If PenaltyInfo(1).GreatLess = "G" Then
        GreaterFlag = True
      Else
        GreaterFlag = False
      End If
      PctAmt# = PctAmt# / 100
      UseBothFlag = True
    Else
      If PctAmt# > 0 Then
        PctAmt# = PctAmt# / 100
        FixAmt# = 0
        UsePctFlag = True
      Else
        PctAmt# = 0
        UsePctFlag = False
      End If
      UseBothFlag = False
    End If

    'Get the Rev source number
'    For cnt = 1 To NumOfRevs
'      ThisRev$ = QPTrim$(Form$(3, 0))
'      If InStr(Choice$(cnt, 0), ThisRev$) Then
'        RevSource = cnt
'        Exit For
'      End If
'    Next
    RevSource = fpcboRevenues.ListIndex + 1
    PenaltyInfo(1).RevSource = RevSource
    'Get Who/How to process
    '***********************
    If Val(fptxtCycle1) > 0 Or Val(fptxtCycle2) > 0 Then
      FirstCycle = Val(fptxtCycle1)
      LastCycle = Val(fptxtCycle2)
      UsingCycle = True
    Else
      FirstBook = Val(fptxtBook1)
      LastBook = Val(fptxtBook2)
      UsingBook = True
    End If
    TransDesc$ = QPTrim$(txtDescription)
    PenaltyInfo(1).PenDesc = TransDesc$
    PenaltyInfo(1).CycFirst = FirstCycle
    PenaltyInfo(1).CycLast = LastCycle
    PenaltyInfo(1).BookFirst = FirstBook
    PenaltyInfo(1).BookLast = LastBook
  Else          'there is an error. Scrn already displayed, wait for input
    'WaitForAction
  End If

'  Action = 1
'  RestScrn TempScrn()


  If InStr(TownName$, "TENN") > 0 And InStr(TownName$, "RIDGE") > 0 Then
    TennFlag = True
  End If

  If InStr(TownName$, "WARSAW") > 0 Then
    PolockFlag = True
  End If

  If InStr(TownName$, "HAMLET") > 0 Then
    HamFlag = True
  End If

  If InStr(TownName$, "TUCKASEIG") > 0 Then
    TuckFlag = True
  End If

  If InStr(TownName$, "CASHION") > 0 Then
    CashFlag = True
  End If
  If InStr(TownName$, "SUNSET") > 0 Then
    SunSetFlag = True
  End If

  If UBSetUpRec(1).BILLCYCL = "Y" Then
    CycleFlag = True
  End If

  For cnt = 1 To MaxRevsCnt
    Source$(cnt) = UBSetUpRec(1).Revenues(cnt).REVNAME
    TaxAmt(cnt) = UBSetUpRec(1).Revenues(cnt).TAXRATE
  Next

  'LibName$ = "UB"
  'ScrnName$ = "UBPENALT"

  '--define the multi-choice fields
  'NumFlds = -1
  'NumFlds = LibNumberOfFields(LibName$, ScrnName$) + 1
  '--define Quick Screen form editing arrays
  'ReDim frm(1) As FormInfo
  'ReDim Form$(NumFlds, 2)
 ' ReDim Fld(NumFlds) As FieldInfo

  '--for each screen, get first and last fields
 ' StartEl = 0
 ' LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode

  '--Clear all fields
 ' For F = 1 To NumFlds
 '   LSet Form$(F, 0) = ""
 ' Next

  '--Set choices
  NumOfRevs = 0
  For cnt = 1 To MaxRevsCnt
    TempRev$ = QPTrim$(Source$(cnt))
    If Len(TempRev$) = 0 Then
      NumOfRevs = cnt - 1
      Exit For
    End If
  Next

'  ReDim Choice$(15, 2)
'  Choice$(0, 0) = "3"
'  For TCnt = 1 To NumOfRevs
'    Choice$(TCnt, 0) = Source$(TCnt)
'  Next TCnt
'
'  Choice$(0, 1) = "4"
'  Choice$(1, 1) = "Current Balance"
'  Choice$(2, 1) = "Previous Balance"
'  Choice$(3, 1) = "Total Balance"
'
'  Choice$(0, 2) = "7"
'  Choice$(1, 2) = "LESS"
'  Choice$(2, 2) = "GREATER"

  PenFile$ = UBPath$ + "UBPENTRN.DAT"

  ' USE CYCLE CHECK
  '--Set screen number to one and display screen
'  DisplayUBScrn ScrnName$
'
'  Action = 1
'  FirstTime = True
'  Do
'
'    EditForm Form$(), Fld(), frm(1), Cnf, Action
'
'    If FirstTime Then
'      FirstTime = False
'      Action = 1
'      LSet Form$(1, 0) = Date$
'      LSet Form$(8, 0) = "0"
'      If CycleFlag Then
'        LSet Form$(11, 0) = "0"
'        LSet Form$(12, 0) = "0"
'        Fld(11).Protected = True
'        Fld(12).Protected = True
'      Else
'        LSet Form$(9, 0) = "0"
'        LSet Form$(10, 0) = "0"
'        Fld(9).Protected = True
'        Fld(10).Protected = True
'      End If
'      If SunSetFlag Then
'        QPrintRC " Sunset Beech Special.", 4, 50, -1
'      End If
'    End If
'
'    '--Check for Key presses
'    Select Case frm(1).KeyCode
'    Case F10Key
'      GoSub CheckPenaltyFlds
'      'If valid Data in Fields
'      'Then Process the Penalties
'    Case EscKey
'      ExitFlag = True
'    End Select
'
'  Loop Until ExitFlag Or OkFlag
'  If ExitFlag Then
'    UBLog " CPF: ABORTED Create Penalty File"
'    GoTo ExitPenalty
'  End If
'  BlockClear
'  ShowProcessingScrn "Calculating Penalty Charges"

  KillFile PenFile$
  PHandle = FreeFile
  Open PenFile$ For Random Shared As PHandle Len = UBTranRecLen

  CHandle = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As CHandle Len = UBCustRecLen

  NumCustRecs& = LOF(CHandle) / UBCustRecLen

  For CCnt& = 1 To NumCustRecs&
    FrmShowPctComp.ShowPctComp CCnt&, NumCustRecs&
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ExitFlag = True
      GoTo ExitPenalty
    End If

    Get CHandle, CCnt&, UBCustRec(1)
    If Not UBCustRec(1).DelFlag Then
      If UBCustRec(1).LATEFEE = "Y" Then
        LSet UBTranRec(1) = UBTranRec(2)        'Transrec 2 is blank
        'Make a clean transaction record
        If PolockFlag Then
          If UBCustRec(1).Status <> "A" And UBCustRec(1).Status <> "B" Then
            GoTo SkipEm
          End If
        ElseIf UBCustRec(1).Status <> "A" Then      'if they are not inactive
          GoTo SkipEm
        End If

        '05-01-97 fixed bug where CurrBalance+PrevBalance is <= 0
        If Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance) > 0 Then

          If UBCustRec(1).CurrBalance >= MinBalance# Or UBCustRec(1).PrevBalance > MinBalance# Then
            'if they have any balance
            If UsingBook Then   'if they want it by Book
              ThisBook = Val(UBCustRec(1).Book)
              If ThisBook >= FirstBook And ThisBook <= LastBook Then
                'if this is in the correct book
                If UseBothFlag Then             'both an amount and percent
                  If UsePrevFlag And Not UseCurrFlag Then       'use prev not curr
                    If UBCustRec(1).CurrBalance < 0 Then
                      PenBal# = UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance
                    Else
                      PenBal# = UBCustRec(1).PrevBalance
                    End If
                  ElseIf UseCurrFlag And Not UsePrevFlag Then   'use curr not
                    If UBCustRec(1).PrevBalance < 0 Then
                      PenBal# = UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance
                    Else
                      PenBal# = UBCustRec(1).CurrBalance
                    End If
                  ElseIf UsePrevFlag And UseCurrFlag Then       'use curr and
                    PenBal# = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
                  End If
                  CustPctPenalty# = Round#(PenBal# * PctAmt#)
                  CustFixPenalty# = FixAmt#
                  If PenBal# <= MinBalance# Then                'if cust had p
                    GoTo SkipEm
                  End If
                  If GreaterFlag Then
                    If CustPctPenalty# >= CustFixPenalty# Then
                      CustPenalty# = CustPctPenalty#
                    Else
                      CustPenalty# = CustFixPenalty#
                    End If
                  Else          'nope want whichever is less
                    If CustPctPenalty# >= CustFixPenalty# Then
                      CustPenalty# = CustFixPenalty#
                    Else
                      CustPenalty# = CustPctPenalty#
                    End If
                  End If
                  GoSub MakeTransaction
                ElseIf UsePctFlag Then          'if they want a percent penalty
                  If UsePrevFlag And Not UseCurrFlag Then       'using prev not curr
                    '030398 Modified to consider a credit in cur or prev balances
                    If UBCustRec(1).CurrBalance < 0 Then
                      PenBal# = UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance
                    Else
                      PenBal# = UBCustRec(1).PrevBalance
                    End If
                    If PenBal# <= MinBalance# Then    'if cust had prev bal
                      GoTo SkipEm
                    End If
                    CustPenalty# = Round#(PenBal# * PctAmt#)
                    GoSub MakeTransaction
                  ElseIf UseCurrFlag And Not UsePrevFlag Then   'using curr not prev
                    '030398 Modified to consider a credit in cur or prev balances
                    If UBCustRec(1).PrevBalance < 0 Then
                      PenBal# = UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance
                    Else
                      PenBal# = UBCustRec(1).CurrBalance
                    End If
                    'code added to exclude tax
                    '092898 Said they didn't take partial payments - Not!
                    If TennFlag Then            'AND UBCustRec(1).TaxExpt <> "Y" then
                      GoSub GetTennRidgeLastBill
                    End If
                    If CashFlag Then
                      GoSub GetCashionLastBill
                    End If

                    If PenBal# <= MinBalance# Then
                      GoTo SkipEm
                    End If
                    CustPenalty# = Round#(PenBal# * PctAmt#)
                    GoSub MakeTransaction
                  ElseIf UsePrevFlag And UseCurrFlag Then       'use curr and prev
                    PenBal# = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)

                    If SunSetFlag Then
                      GoSub CheckSunSet
                      'This adjusts PenBal# for sunsets calc
                    End If

                    If TuckFlag Then
                      GoSub CheckTucka
                      'This adjusts PenBal# for TUCKASEIGEE calc
                    End If

                    If PenBal# <= MinBalance# Then
                      GoTo SkipEm
                    End If
                    CustPenalty# = Round#(PenBal# * PctAmt#)
                    GoSub MakeTransaction
                  End If
                Else            'Using a FIXED penalty amount
                  If UsePrevFlag And Not UseCurrFlag Then
                    '030398 Modified to consider a credit in cur or prev balances
                    If UBCustRec(1).CurrBalance < 0 Then
                      PenBal# = UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance
                    Else
                      PenBal# = UBCustRec(1).PrevBalance
                    End If
                  ElseIf UseCurrFlag And Not UsePrevFlag Then
                    '030398 Modified to consider a credit in cur or prev balances
                    If UBCustRec(1).PrevBalance < 0 Then
                      PenBal# = UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance
                    Else
                      PenBal# = UBCustRec(1).CurrBalance
                    End If
                  ElseIf UsePrevFlag And UseCurrFlag Then
                    'do not need to check for prev >0 or curr>0 here!!
                    PenBal# = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
                  End If
                  If PenBal# <= MinBalance# Then
                    GoTo SkipEm
                  End If
                  CustPenalty# = FixAmt#
                  GoSub MakeTransaction
                End If
              End If
            ElseIf UsingCycle Then              'they are using cycles
              ThisCycle = UBCustRec(1).BILLCYCL
              If ThisCycle >= FirstCycle And ThisCycle <= LastCycle Then
                If UseBothFlag Then             'both an amount and percent
                  If UsePrevFlag And Not UseCurrFlag Then       'use prev not curr
                    If UBCustRec(1).CurrBalance < 0 Then
                      PenBal# = UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance
                    Else
                      PenBal# = UBCustRec(1).PrevBalance
                    End If
                  ElseIf UseCurrFlag And Not UsePrevFlag Then   'use curr not prev
                    If UBCustRec(1).PrevBalance < 0 Then
                      PenBal# = UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance
                    Else
                      PenBal# = UBCustRec(1).CurrBalance
                    End If
                  ElseIf UsePrevFlag And UseCurrFlag Then       'use curr and prev
                    PenBal# = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
                  End If
                  If PenBal# <= MinBalance# Then                'if cust had prev
                    GoTo SkipEm
                  End If
                  CustPctPenalty# = Round#(PenBal# * PctAmt#)
                  CustFixPenalty# = FixAmt#
                  If GreaterFlag Then
                    If CustPctPenalty# >= CustFixPenalty# Then
                      CustPenalty# = CustPctPenalty#
                    Else
                      CustPenalty# = CustFixPenalty#
                    End If
                  Else          'nope want whichever is less
                    If CustPctPenalty# >= CustFixPenalty# Then
                      CustPenalty# = CustFixPenalty#
                    Else
                      CustPenalty# = CustPctPenalty#
                    End If
                  End If
                  GoSub MakeTransaction

                ElseIf UsePctFlag Then   '*** Percentage method
                  If UsePrevFlag And Not UseCurrFlag Then       'use prev not curr
                    '030398 Modified to consider a credit in cur or prev balances
                    If UBCustRec(1).CurrBalance < 0 Then
                      PenBal# = UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance
                    Else
                      PenBal# = UBCustRec(1).PrevBalance
                    End If
                    If PenBal# <= MinBalance# Then              'if cust had prev bal
                      GoTo SkipEm
                    End If
                    CustPenalty# = Round#(PenBal# * PctAmt#)
                    GoSub MakeTransaction
                  ElseIf UseCurrFlag And Not UsePrevFlag Then   'use curr not prev
                    '030398 Modified to consider a credit in cur or prev balances
                    If UBCustRec(1).PrevBalance < 0 Then
                      PenBal# = UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance
                    Else
                      PenBal# = UBCustRec(1).CurrBalance
                    End If
                    If PenBal# <= MinBalance# Then
                      GoTo SkipEm
                    End If
                    CustPenalty# = Round#(PenBal# * PctAmt#)
                    GoSub MakeTransaction
                  ElseIf UsePrevFlag And UseCurrFlag Then       'use curr and prev
                    TotalBalance# = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
                    If TotalBalance# <= MinBalance# Then
                      GoTo SkipEm
                    End If
                    CustPenalty# = Round#(TotalBalance# * PctAmt#)
                    GoSub MakeTransaction
                  End If
                Else            'Using a FIXED penalty amount
                  If UsePrevFlag And Not UseCurrFlag Then
                    '030398 Modified to consider a credit in cur or prev balances
                    If UBCustRec(1).CurrBalance < 0 Then
                      PenBal# = UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance
                    Else
                      PenBal# = UBCustRec(1).PrevBalance
                    End If
                    If PenBal# <= MinBalance# Then              'if cust had prev bal
                      GoTo SkipEm
                    End If
                    CustPenalty# = FixAmt#
                    GoSub MakeTransaction
                  ElseIf UseCurrFlag And Not UsePrevFlag Then
                    '030398 Modified to consider a credit in cur or prev balances
                    If UBCustRec(1).PrevBalance < 0 Then
                      PenBal# = UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance
                    Else
                      PenBal# = UBCustRec(1).CurrBalance
                    End If
                    If PenBal# <= MinBalance# Then
                      GoTo SkipEm
                    End If
                    CustPenalty# = FixAmt#
                    GoSub MakeTransaction
                  ElseIf UsePrevFlag And UseCurrFlag Then
                    CustPenalty# = FixAmt#
                    TotalBalance# = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
                    If TotalBalance# > MinBalance# Then
                      GoSub MakeTransaction
                    End If
                  End If        'if current, previous or both
                End If          'if fixed amount or percent
              End If            'if this cycle
            End If              'if using cycle
          End If                'if balance >= minbalance
        End If                  'if balance >0
      End If                    'if allow late fee
    End If                      'if not deleted
SkipEm:
   ' ShowPctComp CCnt&, NumCustRecs&
  Next

  hand2 = FreeFile
  Open UBPath$ + "UBPENINF.DAT" For Random As hand2
  PenaltyInfo(1).PenCnt = PCnt
 ' FPutAH "UBPENINF.DAT", PenaltyInfo(1), Len(PenaltyInfo(1)), 1
  Put hand2, 1, PenaltyInfo(1)
  'BlockClear
 ' DisplayUBScrn "UPDATEOK"
 ' WaitForAction
Close
ExitPenalty:

  Erase Source$
  Erase UBCustRec, UBSetUpRec, UBTranRec
 ' Erase frm, Form$, Fld

  If Not ExitFlag Then
    UBLog " CPF: Created" + Str$(PenaltyInfo(1).PenCnt) + " work transactions."
  End If
  UBLog "OUT: Create Penalty File." + CrLf$

  Exit Sub

MakeTransaction:
  '011499 Corrected to check for a penalty amount of less than .01
  If Round#(CustPenalty#) > 0 Then
    PCnt = PCnt + 1
    UBTranRec(1).TransAmt = CustPenalty#
    UBTranRec(1).RevAmt(RevSource) = CustPenalty#
    UBTranRec(1).TransDate = PenaltyDate
    UBTranRec(1).TransType = TranPenaltyCharge
    UBTranRec(1).TransDesc = TransDesc$
    UBTranRec(1).CustAcctNo = CCnt&
    'UBTranRec(1).RunBalance = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
    UBTranRec(1).ActiveFlag = True
    Put PHandle, PCnt, UBTranRec(1)
  End If

  Return


GetTennRidgeLastBill:
  'FOpenS "UBTRANS.DAT", THandle
  thandle = FreeFile
  Open UBPath$ + "UBTRANS.DAT" For Random Shared As thandle Len = UBTranRecLen
  PrevTranRec& = UBCustRec(1).LastTrans

  If PrevTranRec& > 0 Then
    Do While PrevTranRec& > 0
      Get #thandle, PrevTranRec&, UBTranRec(3)
      If UBTranRec(3).TransType = TranUtilityBill Then
        PenBal# = Round#(UBTranRec(3).RevAmt(1) + UBTranRec(3).RevAmt(2))
        Exit Do
      End If
      PrevTranRec& = UBTranRec(3).PrevTrans
    Loop
  End If

  Close thandle
Return

GetCashionLastBill:
  thandle = FreeFile
  Open UBPath$ + "UBTRANS.DAT" For Random Shared As thandle Len = UBTranRecLen
  PrevTranRec& = UBCustRec(1).LastTrans
  If PrevTranRec& > 0 Then
    Do While PrevTranRec& > 0
      Get thandle, PrevTranRec&, UBTranRec(3)
      If UBTranRec(3).TransType = TranUtilityBill Then
        PenBal# = Round#(UBTranRec(3).RevAmt(1) + UBTranRec(3).RevAmt(2))
        PenBal# = Round#(PenBal# + UBTranRec(3).RevAmt(3) + UBTranRec(3).RevAmt(4))
        PenBal# = Round#(PenBal# + UBTranRec(3).RevAmt(5) + UBTranRec(3).RevAmt(6))
        Exit Do
      End If
      PrevTranRec& = UBTranRec(3).PrevTrans
    Loop
  End If

  Close thandle
Return

CheckSunSet:
  PenBal# = Round#(UBCustRec(1).CurrRevAmts(1) + UBCustRec(1).CurrRevAmts(5))
  If PenBal# < 0 Then
    PenBal# = -10000
  End If
Return

CheckTucka:
  If PenBal# > 0 Then
    If UBCustRec(1).CurrRevAmts(7) > 0 Then
      PenBal# = Round#(PenBal# - UBCustRec(1).CurrRevAmts(7))
    End If
  End If
  If PenBal# < 0 Then
    PenBal# = -10000
  End If

Return


ChkLastTrans:
  'IF UBCustRec(1).Status = "B" THEN
  thandle = FreeFile
  Open UBPath$ + "UBTRANS.DAT" For Random Shared As thandle Len = UBTranRecLen
  PrevTranRec& = UBCustRec(1).LastTrans
  If PrevTranRec& > 0 Then
    Get thandle, PrevTranRec&, UBTranRec(3)
    If UBTranRec(3).TransType = TranPenaltyCharge Then
      NOPenFlag = True
    End If
  End If
  Close thandle
Return

End Sub
