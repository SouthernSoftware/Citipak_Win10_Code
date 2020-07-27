VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#3.5#0"; "SPR32X35.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVenSetDefDist 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Set Vendor Default Distributions"
   ClientHeight    =   8640
   ClientLeft      =   36
   ClientTop       =   540
   ClientWidth     =   12192
   Icon            =   "frmVenSetDefDist.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   12192
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpList fplstVendor 
      Height          =   1200
      Left            =   5856
      TabIndex        =   16
      Top             =   1572
      Width           =   4356
      _Version        =   196608
      _ExtentX        =   7683
      _ExtentY        =   2117
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
      Columns         =   0
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
      ColumnHeaderShow=   0   'False
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
      ColDesigner     =   "frmVenSetDefDist.frx":08CA
   End
   Begin LpLib.fpCombo fpcboAcctNumNa 
      Height          =   384
      Left            =   2112
      TabIndex        =   1
      Top             =   3312
      Width           =   4944
      _Version        =   196608
      _ExtentX        =   8721
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
      Object.TabStop         =   -1  'True
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Text            =   ""
      Columns         =   4
      Sorted          =   0
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   3
      ColumnWidthScale=   2
      RowHeight       =   -1
      WrapList        =   0   'False
      WrapWidth       =   0
      AutoSearch      =   2
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
      ScrollBarH      =   3
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
      ColDesigner     =   "frmVenSetDefDist.frx":0B8E
   End
   Begin LpLib.fpCombo fpcboVendCode 
      Height          =   384
      Left            =   3192
      TabIndex        =   0
      Top             =   1620
      Width           =   2196
      _Version        =   196608
      _ExtentX        =   3873
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
      Object.TabStop         =   -1  'True
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Text            =   ""
      Columns         =   2
      Sorted          =   0
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   0
      ColumnWidthScale=   2
      RowHeight       =   -1
      WrapList        =   0   'False
      WrapWidth       =   0
      AutoSearch      =   2
      SearchMethod    =   1
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
      ScrollBarH      =   3
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
      EditMarginLeft  =   5
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmVenSetDefDist.frx":0F91
   End
   Begin VB.CommandButton cmdDelDist 
      BackColor       =   &H00D0D0D0&
      Caption         =   "F6 Del D&ist"
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
      Left            =   2742
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7560
      Width           =   1332
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00D0D0D0&
      Caption         =   "F10 &Save"
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
      Left            =   6326
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7560
      Width           =   1332
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00D0D0D0&
      Caption         =   "F3 &Delete"
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
      Left            =   4534
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7560
      Width           =   1332
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00D0D0D0&
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
      Left            =   8118
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7560
      Width           =   1332
   End
   Begin EditLib.fpDoubleSingle fpPercent 
      Height          =   396
      Left            =   7296
      TabIndex        =   2
      Top             =   3312
      Width           =   1356
      _Version        =   196608
      _ExtentX        =   2392
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
      ControlType     =   0
      Text            =   "0.00"
      DecimalPlaces   =   2
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
   Begin VB.CommandButton cmdAddDist 
      BackColor       =   &H00D0D0D0&
      Caption         =   "F9 &Add Distribution"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   8808
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3192
      Width           =   1572
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   2604
      Left            =   1992
      TabIndex        =   4
      Top             =   4188
      Width           =   8196
      _Version        =   196613
      _ExtentX        =   14457
      _ExtentY        =   4593
      _StockProps     =   64
      AutoSize        =   -1  'True
      BackColorStyle  =   1
      ButtonDrawMode  =   4
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   4
      MaxRows         =   8
      OperationMode   =   1
      ScrollBars      =   0
      SelectBlockOptions=   0
      ShadowColor     =   13684944
      ShadowDark      =   8421504
      SpreadDesigner  =   "frmVenSetDefDist.frx":1318
      VisibleCols     =   3
      VisibleRows     =   8
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   312
      Left            =   0
      TabIndex        =   9
      Top             =   8328
      Width           =   12192
      _ExtentX        =   21505
      _ExtentY        =   550
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
            TextSave        =   "2:07 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7133
            TextSave        =   "3/14/2005"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin EditLib.fpDoubleSingle fpTotPercent 
      Height          =   348
      Left            =   8688
      TabIndex        =   17
      Top             =   6936
      Width           =   1428
      _Version        =   196608
      _ExtentX        =   2519
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
      ControlType     =   0
      Text            =   "0.00"
      DecimalPlaces   =   2
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
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   612
      Left            =   3030
      Top             =   432
      Width           =   6132
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Set Default Distributions"
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
      Left            =   3870
      TabIndex        =   15
      Top             =   552
      Width           =   4452
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H80000016&
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   1524
      Index           =   0
      Left            =   1872
      Top             =   1356
      Width           =   8616
   End
   Begin VB.Label Label3b 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor:"
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
      Index           =   1
      Left            =   2115
      TabIndex        =   14
      Top             =   1650
      Width           =   900
   End
   Begin VB.Label lblCredits 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total Percent"
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
      Index           =   3
      Left            =   6888
      TabIndex        =   13
      Top             =   6984
      Width           =   1596
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000005&
      BorderWidth     =   3
      Height          =   420
      Left            =   1872
      Top             =   6912
      Width           =   8616
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "GL Account Number/Name"
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
      Index           =   0
      Left            =   2796
      TabIndex        =   12
      Top             =   2964
      Width           =   3132
   End
   Begin VB.Label Label3b 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Percent"
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
      Index           =   5
      Left            =   7512
      TabIndex        =   11
      Top             =   2988
      Width           =   936
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   4056
      Left            =   1872
      Top             =   2868
      Width           =   8616
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000005&
      BorderWidth     =   3
      Height          =   2796
      Left            =   1896
      Top             =   4092
      Width           =   8556
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Distributions :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   1968
      TabIndex        =   10
      Top             =   3756
      Width           =   1500
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D0D0D0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00D0D0D0&
      Height          =   732
      Left            =   3030
      Top             =   312
      Width           =   6132
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
Attribute VB_Name = "frmVenSetDefDist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Dim Vendor As VendorRecType
Dim VdefDist() As VendorDefDistRecType
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Private Temp_Class As Resize_Class
Dim VDefRec As Integer, VRecNum As Integer

Private Sub cmdAddDist_Click()
  If fpcboAcctNumNa.Text <> "" And fpPercent.DoubleValue > 0 Then
    vaSpread1.Row = vaSpread1.DataRowCnt + 1
    vaSpread1.col = 1
    vaSpread1.Text = Val(vaSpread1.Row)
    vaSpread1.col = 2
    fpcboAcctNumNa.col = 1
    vaSpread1.Text = fpcboAcctNumNa.ColText
    vaSpread1.col = 3
    fpcboAcctNumNa.col = 2
    vaSpread1.Text = fpcboAcctNumNa.ColText
    vaSpread1.col = 4
    fpTotPercent = (fpTotPercent.DoubleValue + fpPercent.DoubleValue)
    vaSpread1.Text = fpPercent
    fpcboAcctNumNa.ListIndex = -1
    fpPercent = 0
    fpcboAcctNumNa.SetFocus
    
  Else
    MsgBox "The Account and Percent Must Be Entered Before Adding To The Distribution List.", vbOKOnly, "Add Distribution Denied"
  End If

End Sub

Private Sub cmdDelDist_Click()
  If vaSpread1.ActiveRow > 0 Then
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.col = 4
    If vaSpread1.Text <> "" Then
      If MsgBox("You Wish to Delete this Distribution?", vbYesNo, "Delete Distribution") = vbYes Then
        vaSpread1.col = 4
        fpTotPercent = (fpTotPercent.DoubleValue - vaSpread1.Text)
        vaSpread1.DeleteRows vaSpread1.Row, 1
        fpcboAcctNumNa.SetFocus
      End If
    End If
  End If

End Sub

Private Sub cmdDelete_Click()
  If fpcboVendCode.ListIndex <> -1 Then
    If MsgBox("Are You Sure You Wish To Delete Distribution?", vbYesNo, "Delete?") = vbYes Then
      DelDefRec
      ClearScr
      VRecNum = 0
      VDefRec = 0
      fpcboVendCode.ListIndex = -1
    End If
  End If
End Sub

Private Sub cmdExit_Click()
  If Changed = True Then
    If MsgBox("Changes Have Been Made, Yes to Abandon or No to Complete Editing?", vbYesNo, "Exit ?") = vbYes Then
      frmAPVendMaintMenu.Show
      Unload frmVenSetDefDist
    Else
      fpcboAcctNumNa.SetFocus
    End If
  Else
    frmAPVendMaintMenu.Show
    Unload frmVenSetDefDist
  End If
End Sub

Private Sub cmdSave_Click()
  If Editing = False Then
    If vaSpread1.DataRowCnt > 0 Then
      If fpTotPercent.Value <> 100 Then
        MsgBox "The Total Percent Must Equal 100", vbOKOnly, "Incorrect Amount"
        vaSpread1.SetFocus
        Exit Sub
      End If
    End If
    SaveDefRec
    ClearScr
    VRecNum = 0
    VDefRec = 0
    fpcboVendCode.ListIndex = -1
    fpcboVendCode.SetFocus
  Else
    MsgBox "Please Complete Current Distribution or Clear It Before Continuing.", vbOKOnly, "Retry"
    cmdAddDist.SetFocus
  End If

End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  StatusBar1.Panels.Item(1).Text = GLUserName
  Me.HelpContextID = hlpDefDist
  VDefRec = 0
  VRecNum = 0
  VendCodeList fpcboVendCode
  Fixspread
  FillAcctNumName fpcboAcctNumNa

End Sub
Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
 '   Me.SetFocus
  End If
End Sub
Private Sub fpcboAcctNumNa_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboAcctNumNa.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcboAcctNumNa.ListIndex = -1
    fpcboAcctNumNa.Action = ActionClearSearchBuffer
  End If
  If fpcboAcctNumNa.ListDown <> True Then
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog "Close AP"
      ClearInUse PWcnt
    End If
  End If
End Sub

Private Sub fpcboAcctNumNa_LostFocus()
  fpcboAcctNumNa.Action = ActionClearSearchBuffer
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyUp:
      SendKeys "+{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      cmdExit_Click
      KeyCode = 0
    Case vbKeyF9:
      cmdAddDist_Click
      KeyCode = 0
    Case vbKeyF3:
      cmdDelete_Click
      KeyCode = 0
    Case vbKeyF6:
      cmdDelDist_Click
      KeyCode = 0
    Case vbKeyF10:
      cmdSave_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub
'This is to fix spreadsheet for various resolutions
Public Function Fixspread()
    Select Case screenW
      Case 1280
      If Screen.TwipsPerPixelX <> 12 Then
        coladj = 27
        vaSpread1.RowHeight(-1) = 23
        vaSpread1.RowHeight(0) = 23
      Else
        coladj = 19.8
        vaSpread1.RowHeight(-1) = 18
        vaSpread1.RowHeight(0) = 18
      End If
      Case 1152
      If Screen.TwipsPerPixelX <> 12 Then
        coladj = 23.5
        vaSpread1.RowHeight(0) = 19
        vaSpread1.RowHeight(-1) = 19
      Else
        coladj = 17
        vaSpread1.RowHeight(0) = 15
        vaSpread1.RowHeight(-1) = 15
      End If
      Case 1024
      If Screen.TwipsPerPixelX <> 12 Then
        coladj = 19.5
        vaSpread1.RowHeight(-1) = 17
        vaSpread1.RowHeight(0) = 17
      Else
        coladj = 14
      End If
      Case 800
        coladj = 13.5
        vaSpread1.Font.Size = 8
        vaSpread1.RowHeight(-1) = 13
      Case Else
        'don't worry be happpy
    End Select
    vaSpread1.ColWidth(-1) = vaSpread1.ColWidth(-1) + coladj
End Function
Private Sub ClearScr()
'Dont clear vendor code cause use for load up procedure.
  fplstVendor.Clear
  vaSpread1.ClearRange 1, 1, 4, 8, True
  fpTotPercent = 0
  fpcboAcctNumNa.ListIndex = -1
  fpPercent = 0
End Sub

Private Sub LoadUp()
  Dim APDefDistFile As Integer, NumDefRecs As Integer
  Dim VendorFile As Integer, NumVRecs As Integer, DefRecLen As Integer
  Dim Last As Integer, cnt As Integer, Dcnt As Integer, TmpAcct As Integer
  ReDim VdefDist(1) As VendorDefDistRecType
  VRecNum = 0
  VDefRec = 0
  DefRecLen = Len(VdefDist(1))
  fpcboVendCode.col = 1
  VRecNum = fpcboVendCode.ColText
  If VRecNum > 0 Then
    ClearScr
    OpenVendorFile VendorFile, NumVRecs
    Get VendorFile, VRecNum, Vendor
    fplstVendor.Row = -1
    fplstVendor.InsertRow = Vendor.VNAME
    fplstVendor.InsertRow = Vendor.Addr1
    fplstVendor.InsertRow = Vendor.Addr2
    fplstVendor.InsertRow = QPTrim$(Vendor.City) + " " + Vendor.State + " " + Vendor.Zip
    VDefRec = Vendor.DefDist
     If VDefRec > 0 Then
      fpTotPercent = 0
      OpenDefDistFile DefRecLen, APDefDistFile, NumDefRecs
      Get APDefDistFile, VDefRec, VdefDist(1)
      'If Not VdefDist(1).VRecNum = VRecNum Then 'Stop
      For cnt = 1 To 8 'Vendor.DefDist
        If QPTrim(VdefDist(1).DefDist(cnt).DefAcct) <> "" Then
        TmpAcct = AcctFind(QPTrim(VdefDist(1).DefDist(cnt).DefAcct))
          If TmpAcct > 0 Then
            vaSpread1.Row = vaSpread1.DataRowCnt + 1
            vaSpread1.col = 1
            vaSpread1.Text = VRecNum
            vaSpread1.col = 2
            vaSpread1.Text = QPTrim$(VdefDist(1).DefDist(cnt).DefAcct)
            vaSpread1.col = 3
            vaSpread1.Text = QPTrim$(VdefDist(1).DefDist(cnt).DefAcctName)
            vaSpread1.col = 4
            vaSpread1.Text = VdefDist(1).DefDist(cnt).DefPct
            fpTotPercent = (fpTotPercent.DoubleValue + VdefDist(1).DefDist(cnt).DefPct)
          Else
            MsgBox "Account could not be found and will not be loaded.", vbOKOnly, "Invalid Account"
           ' VDefRec = 0
          End If
       
        End If
      Next
     ' Else
      ' VDefRec = 0
     ' End If
    End If
'    VDefRec = Vendor.DefDist
'    If VDefRec > 0 Then
'      GoSub GetDefRec
    End If
 ' End If
 Close
End Sub


Private Sub SaveDefRec()
  Dim DefRecLen As Integer, APDefDistFile As Integer, NumDefRecs As Integer
  Dim cnt As Integer, DefRecNum As Integer
  Dim VendorFile As Integer, NumVRecs As Integer
  OpenDefDistFile DefRecLen, APDefDistFile, NumDefRecs
  VdefDist(1).VRecNum = VRecNum
  For cnt = 1 To vaSpread1.DataRowCnt
    vaSpread1.Row = cnt
    vaSpread1.col = 2
    If QPTrim(vaSpread1.Text) <> "" Then
      VdefDist(1).DefDist(cnt).DefAcct = vaSpread1.Text
      vaSpread1.col = 3
      VdefDist(1).DefDist(cnt).DefAcctName = vaSpread1.Text
      vaSpread1.col = 4
      VdefDist(1).DefDist(cnt).DefPct = vaSpread1.Text
    Else
      VdefDist(1).DefDist(cnt).DefAcct = ""
      VdefDist(1).DefDist(cnt).DefAcctName = ""
      VdefDist(1).DefDist(cnt).DefPct = 0
    End If
  Next
'     If vaSpread1.DataRowCnt > 0 Then
'      'If VDefRec Then
'      DefRecNum = vaSpread1.DataRowCnt
'
'      Else
'     DefRecNum = NumDefRecs + 1
'    End If
'
'    Put APDefDistFile, DefRecNum, VdefDist(1)
'    OpenVendorFile VendorFile, NumVRecs
'    Get VendorFile, VRecNum, Vendor
'    Vendor.DefDist = DefRecNum
'    Put VendorFile, VRecNum, Vendor
'
'  Close
'  cnt = 0
  If VDefRec Then
    DefRecNum = VDefRec
    Put APDefDistFile, DefRecNum, VdefDist(1)
  Else
'    If vaSpread1.DataRowCnt = 0 Then
'      DefRecNum = 0
'    Else
      DefRecNum = NumDefRecs + 1
      Put APDefDistFile, DefRecNum, VdefDist(1)
 '   End If
  End If
    OpenVendorFile VendorFile, NumVRecs
    Get VendorFile, VRecNum, Vendor
    If vaSpread1.DataRowCnt > 0 Then
      Vendor.DefDist = DefRecNum
    Else
      Vendor.DefDist = 0
    End If
    Put VendorFile, VRecNum, Vendor
  
  Close
  cnt = 0 ': Dcnt = 0

End Sub


Private Sub DelDefRec()
  Dim VendorFile As Integer, NumVRecs As Integer
  OpenVendorFile VendorFile, NumVRecs
  Get VendorFile, VRecNum, Vendor
  Vendor.DefDist = 0
  Put VendorFile, VRecNum, Vendor
End Sub


Private Sub fpcboVendCode_Click()
  If fpcboVendCode.ListIndex <> -1 Then
    LoadUp
    'fpcboAcctNumNa.SetFocus
  End If
End Sub


Private Sub fpcboVendCode_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboVendCode.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcboVendCode.ListIndex = -1
    fpcboVendCode.Action = ActionClearSearchBuffer
  End If
  If fpcboVendCode.ListDown <> True Then
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



Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub

Private Sub vaSpread1_DblClick(ByVal col As Long, ByVal Row As Long)
  Dim TempAcct As String
  Dim TempCol As Long, TempRow As Long
  TempRow = Row
  TempCol = col
  vaSpread1.Row = TempRow
  vaSpread1.col = 2
  TempAcct = QPTrim(vaSpread1.Text)

    If TempRow > 0 And TempAcct <> "" Then
      If fpcboAcctNumNa.ListIndex <> -1 Or fpPercent <> 0 Then
        If MsgBox("Complete Current Distribution, 'Yes' OR 'No', Clear Account and Percent ?" & Chr$(13) & "If You Select No The Current Distribution Will Not Be Saved.", vbYesNo, "Distribution") = vbYes Then
          Exit Sub
        Else
          fpcboAcctNumNa.ListIndex = -1
          fpPercent = 0
        End If
      End If

        fpcboAcctNumNa.SearchText = QPStrip(TempAcct)
        fpcboAcctNumNa.Action = 0
        If fpcboAcctNumNa.SearchIndex <> -1 Then
          fpcboAcctNumNa.ListIndex = fpcboAcctNumNa.SearchIndex
        End If
        vaSpread1.col = 1
        fpcboAcctNumNa.col = 0
        'fpcboAcctNumNa.ColText = vaSpread1.Text
        vaSpread1.col = 2
        fpcboAcctNumNa.col = 1
        fpcboAcctNumNa.ColText = vaSpread1.Text
        vaSpread1.col = 3
        fpcboAcctNumNa.col = 2
        fpcboAcctNumNa.ColText = vaSpread1.Text
        vaSpread1.col = 4
        fpPercent = vaSpread1.Text
        fpTotPercent = (fpTotPercent.DoubleValue - fpPercent.DoubleValue)
        'vaSpread1.ClearRange TempCol, TempRow, 4, TempRow, True
        vaSpread1.DeleteRows TempRow, 1
        fpcboAcctNumNa.SetFocus
    End If
  
End Sub
Private Function Editing()
  If fpcboVendCode.ListIndex = -1 Then
    Editing = False
  Else
    If fpcboAcctNumNa.ListIndex = -1 And fpPercent = 0 Then
      Editing = False
    Else
      Editing = True
    End If
  End If
End Function
Private Function Changed()
  Dim DefRecLen As Integer, APDefDistFile As Integer, NumDefRecs As Integer
  Dim cnt As Integer
  If fpcboVendCode.ListIndex <> -1 Then
    If Vendor.DefDist > 0 Then
      OpenDefDistFile DefRecLen, APDefDistFile, NumDefRecs
      Get APDefDistFile, VDefRec, VdefDist(1)
      If fpcboAcctNumNa.ListIndex <> -1 Or fpPercent <> 0 Then
        Changed = True
        Close
        Exit Function
      Else
        For cnt = 1 To 8 'Vendor.DefDist
          vaSpread1.Row = cnt
          vaSpread1.col = 2
          If vaSpread1.Text = QPTrim$(VdefDist(1).DefDist(cnt).DefAcct) Then
            vaSpread1.col = 3
            If vaSpread1.Text = QPTrim$(VdefDist(1).DefDist(cnt).DefAcctName) Then
              vaSpread1.col = 4
              If Val(vaSpread1.Text) = VdefDist(1).DefDist(cnt).DefPct Then
                Changed = False
              Else
                Changed = True
                Exit For
              End If
            Else
              Changed = True
              Exit For
            End If
        Else
          Changed = True
          Exit For
        End If
      Next
    End If
   Else
     If vaSpread1.DataRowCnt > 0 Then
        Changed = True
     Else
        Changed = False
     End If
   End If
  Else
    Changed = False
  End If
Close
               
End Function

