VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPrnSalesTaxRptCounty 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sales Tax Report By County"
   ClientHeight    =   8895
   ClientLeft      =   30
   ClientTop       =   495
   ClientWidth     =   12195
   Icon            =   "frmPrnSalesTaxRptCounty.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8895
   ScaleWidth      =   12195
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboRptType 
      Height          =   405
      Left            =   5115
      TabIndex        =   4
      Top             =   5355
      Width           =   1905
      _Version        =   196608
      _ExtentX        =   3360
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
      ColDesigner     =   "frmPrnSalesTaxRptCounty.frx":08CA
   End
   Begin LpLib.fpCombo fpcboStateCode 
      Height          =   405
      Left            =   5130
      TabIndex        =   1
      Top             =   3525
      Width           =   990
      _Version        =   196608
      _ExtentX        =   1746
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
      ColDesigner     =   "frmPrnSalesTaxRptCounty.frx":0C30
   End
   Begin LpLib.fpCombo fpcboCoAcct 
      Height          =   405
      Left            =   5130
      TabIndex        =   0
      Top             =   2925
      Width           =   4740
      _Version        =   196608
      _ExtentX        =   8361
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
      SearchMethod    =   2
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
      ColDesigner     =   "frmPrnSalesTaxRptCounty.frx":0F5F
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
      Left            =   8256
      TabIndex        =   5
      Top             =   7512
      Width           =   1332
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
      Left            =   10032
      TabIndex        =   6
      Top             =   7512
      Width           =   1332
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   7
      Top             =   8532
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
            TextSave        =   "9:34 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7117
            TextSave        =   "7/24/2009"
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
   Begin EditLib.fpDateTime fpDate1 
      Height          =   372
      Left            =   5136
      TabIndex        =   2
      Top             =   4128
      Width           =   1692
      _Version        =   196608
      _ExtentX        =   2984
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
   Begin EditLib.fpDateTime fpDate2 
      Height          =   372
      Left            =   5136
      TabIndex        =   3
      Top             =   4740
      Width           =   1692
      _Version        =   196608
      _ExtentX        =   2984
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
   Begin VB.Label Label10 
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
      Height          =   372
      Left            =   2664
      TabIndex        =   13
      Top             =   5400
      Width           =   2388
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "County Code:"
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
      Left            =   2784
      TabIndex        =   12
      Top             =   3576
      Width           =   2196
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "County Account:"
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
      Height          =   420
      Left            =   2784
      TabIndex        =   11
      Top             =   2964
      Width           =   2196
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   4212
      Left            =   1920
      Top             =   2232
      Width           =   8316
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Print Sales Tax Report By County"
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
      Left            =   3240
      TabIndex        =   10
      Top             =   1248
      Width           =   5700
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3216
      Top             =   1008
      Width           =   5772
   End
   Begin VB.Image Image1 
      Height          =   345
      Left            =   2280
      Picture         =   "frmPrnSalesTaxRptCounty.frx":1362
      Top             =   2370
      Width           =   360
   End
   Begin VB.Label Label3 
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
      Height          =   420
      Left            =   3408
      TabIndex        =   9
      Top             =   4788
      Width           =   1572
   End
   Begin VB.Label Label1 
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
      Height          =   420
      Left            =   3312
      TabIndex        =   8
      Top             =   4164
      Width           =   1668
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   972
      Left            =   3216
      Top             =   888
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
Attribute VB_Name = "frmPrnSalesTaxRptCounty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim GLSetup As GLSetupRecType
Dim Acct    As GLAcctRecType
Dim GLFundIdx As GLFundIndexType
Dim AcctIdx As GLAcctIndexType
Dim Vendor As VendorRecType
Dim VendorIdx As VendorIdxRecType
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer

Private Sub cmdExit_Click()
  frmAPReportsMenu.Show
  Unload Me
End Sub
Private Sub fpcboRptType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboRptType.ListDown = True
  End If
  If fpcboRptType.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      cmdOk.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpDate2.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub
Private Sub cmdOk_Click()


  If Oktogo = True Then
    If fpcboRptType.ListIndex = 0 Then
      rptopt = 1
    ElseIf fpcboRptType.ListIndex = 1 Then
      rptopt = 2
    End If
    If rptopt = 1 Then
      SalesTaxReportNOSTATE
    ElseIf rptopt = 2 Then
      SalesTaxReportCo2
    End If

  End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = True Then
      If MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        MainLog "Close AP"
        ClearInUse PWcnt
      End If
    Else
      Cancel = True
    End If
  End If
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
      SendKeys "%X"
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%O"
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  StatusBar1.Panels.Item(1).Text = GLUserName
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Me.HelpContextID = hlpCoSalesTx
  fpDate1.Text = Format(Now, "mm/dd/yyyy")
  fpDate2.Text = Format(Now, "mm/dd/yyyy")
  FillAcctNumName fpcboCoAcct
  VendcoCodeList fpcboStateCode
  fpcboStateCode.ListIndex = 0
  fpcboRptType.InsertRow = "Graphics"
  fpcboRptType.InsertRow = "Text"
  fpcboRptType.ListIndex = 0
End Sub
Private Sub fpcboCoAcct_LostFocus()
  fpcboCoAcct.Action = ActionClearSearchBuffer
End Sub

Private Sub fpcboStateCode_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboStateCode.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcboStateCode.ListIndex = -1
    fpcboStateCode.Action = ActionClearSearchBuffer
  End If
  If fpcboStateCode.ListDown <> True Then
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

Private Sub fpcboCoAcct_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboCoAcct.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcboCoAcct.ListIndex = -1
    fpcboCoAcct.Action = ActionClearSearchBuffer
  End If
  If fpcboCoAcct.ListDown <> True Then
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


Private Sub Form_Resize()
'  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
'  End If
End Sub
Private Function Oktogo()
Dim TempDate1 As Integer, TempDate2 As Integer
    If CheckValDate(fpDate1) = False And CheckValDate(fpDate2) = False Then
      MsgBox "Date Is Not Valid. Please Correct.", vbOKOnly, "Invalid Date"
      Oktogo = False
    Else
      TempDate1 = DateDiff("d", "12/31/1979", fpDate1)
      TempDate2 = DateDiff("d", "12/31/1979", fpDate2)
      If TempDate1 > TempDate2 Then
        Oktogo = False
        MsgBox "The Starting And Ending Dates Must Be In Chronological Order Or Equal", vbOKOnly, "Invalid Date"
      Else
        Oktogo = True
      End If
    End If
    If fpcboCoAcct.ListIndex <> -1 Then
      Oktogo = True
    Else
      MsgBox "Please Select County Account.", vbOKOnly, "Invalid Accounts"
      Oktogo = False
    End If
End Function

Private Sub SalesTaxReportCo2()
  Dim Combined As Boolean, StateTaxRecAcct As String, CountyTaxRecAcct As String
  Dim StateFactor As Double, BegDate As Integer, EndDate As Integer
  Dim State As Integer, PrintFlag As Integer, StTotTax As Double
  Dim TotState As Double, TotCounty As Double, TCnt As Integer
  Dim CoTotTax As Double, ll As Integer, VendorFile As Integer
  Dim NumVRecs As Integer, LdRecLen As Integer, APLedgerFile As Integer
  Dim NumTran As Long, APDRecLen As Integer, APDistFile As Integer
  Dim NumDistRecs As Long, cnt As Integer, NextTran As Long, SCnt As Integer
  Dim NextDist As Long, offset As Integer, Coffset As Integer
  Dim StTax As Double, ccnt As Integer, CoTax As Double, PRNFile As Integer
  Dim ReportFile As String, Header As String, TotList As Integer
  Dim StateAmt As Double, CountyAmt As Double, StateCd As String
  Dim all As Boolean, Newrp As String, TEST As String
  fpcboCoAcct.col = 1
  CountyTaxRecAcct$ = QPTrim$(fpcboCoAcct.ColText)
  BegDate = DateDiff("d", "12/31/1979", fpDate1)
  EndDate = DateDiff("d", "12/31/1979", fpDate2)
  PRNFile = FreeFile
  Newrp = "SalTax"
  GetRPTName Newrp
  ReportFile$ = Newrp
  Open ReportFile$ For Output As #PRNFile
  Header$ = "Sales Tax Report"
  CoTotTax# = 0
  If fpcboStateCode.ListIndex <= 0 Then
    StateCd = 0
    all = True
    TotList = fpcboStateCode.ListCount - 1
  Else
    StateCd = QPTrim(fpcboStateCode.Text)
    all = False
    TotList = 1
  End If
    FrmShowPctComp.Label1 = "Searching Codes For Sales Tax Report"
    FrmShowPctComp.Show , Me
    DoEvents
   DeActivateControls frmPrnSalesTaxRptCounty, True
  ReDim StSalesTaxPaid#(0 To 999)

    ReDim CoSalesTaxPaid#(0 To 999)
  GoSub Printhead
  For State = 1 To TotList
    PrintFlag = 0
    FrmShowPctComp.ShowPctComp State, TotList
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ActivateControls frmPrnSalesTaxRptCounty, True
      Unload FrmShowPctComp
      GoTo CancelExit
    End If
 
    If all Then
      fpcboStateCode.Row = State
      StateCd = QPTrim(fpcboStateCode.List)
    End If
   ' PrintFlag = 0
    StTotTax# = 0
    TotState# = 0
    TotCounty# = 0
    TCnt = 0
    '
    For ll = 0 To 999
      StSalesTaxPaid#(ll) = 0
      If Not Combined Then
        CoSalesTaxPaid#(ll) = 0
      End If
    Next ll

    Dim Vendor As VendorRecType
    Close VendorFile
    OpenVendorFile VendorFile, NumVRecs

    Dim ApLedger As APLedger81RecType
    LdRecLen = Len(ApLedger)
    Close APLedgerFile
    OpenAPLedgerFile APLedgerFile, NumTran&, LdRecLen
    Dim APDist As APDistRecType
    APDRecLen = Len(APDist)
    Close APDistFile
    OpenAPDistFile APDistFile, NumDistRecs&, APDRecLen

    For cnt = 1 To NumVRecs
      Get VendorFile, cnt, Vendor
      NextTran = Vendor.FrstTran
      If NextTran > 0 Then
        Do
          Get APLedgerFile, NextTran, ApLedger
          If ApLedger.TRCode = 1 Then
            '--check this out
            If ApLedger.GLDistDate >= BegDate And ApLedger.GLDistDate <= EndDate Then
              TCnt = TCnt + 1
              NextDist& = ApLedger.FrstDist
              Do
                Get APDistFile, NextDist&, APDist

'                If QPTrim$(APDist.DistAcctNum) = StateTaxRecAcct$ Then
'                  Get VendorFile, APLedger.VRecNum, Vendor
'                  'offset = State
'
'                  Coffset = Val(Vendor.CoCode)
'
'                  If StateCd = QPTrim(Vendor.StCode) Then
'                    SCnt = SCnt + 1
'                    StTax# = StTax# + APDist.DistAmt
'                    PrintFlag = 1
'                    'offset = State
'                    If all Then offset = 0
'                    StSalesTaxPaid#(Coffset) = StSalesTaxPaid#(Coffset) + APDist.DistAmt
'                    'LPRINT APDist.DistAmt
'                  End If
'                End If
'                If Not Combined Then
                  If QPTrim$(APDist.DistAcctNum) = CountyTaxRecAcct$ Then
                    Get VendorFile, ApLedger.VRecNum, Vendor
                    If StateCd = QPTrim(Vendor.CoCode) Then
                      ccnt = ccnt + 1
                      CoTax# = CoTax# + APDist.DistAmt
                      PrintFlag = 1
                      'TEST = Vendor.vnum
                      If all Then
                        offset = State
                      Else
                        offset = 1
                      End If
                      'If offset < 0 Or offset > 999 Then offset = 0
                      '  LPRINT Vendor.Vname, APDist.DistAmt
                      CoSalesTaxPaid#(Coffset) = CoSalesTaxPaid#(Coffset) + APDist.DistAmt
                    End If
                  End If
                'End If
                NextDist& = APDist.NextDist
              Loop Until NextDist& = 0

            End If

          End If
          NextTran = ApLedger.NextTrans
        Loop Until NextTran = 0
      End If

    Next cnt

    If PrintFlag = 1 Then
      GoSub PrintState
    End If
      '
    'End If
' "State Code Being Searched: ";: Color 15: Print State
  Next State
    Print #PRNFile, "--------------------------------------------"
    Print #PRNFile, "Account Totals"; Tab(34); Using("######.##", CoTotTax#)
    Print #PRNFile, Chr$(12)
  If TotList < 1 Then
    FrmShowPctComp.ShowPctComp 1, 1
    Print #PRNFile, "No Information to Display"
  End If
  Close
  Unload FrmShowPctComp
  ActivateControls frmPrnSalesTaxRptCounty, True
  ViewPrint ReportFile$, Header$
  KillFile ReportFile$
  
  Exit Sub
Printhead:
  Print #PRNFile, "Sales Tax Report"
  Print #PRNFile, "Account: "; CountyTaxRecAcct$
  Print #PRNFile,
  Print #PRNFile, "             County Code            County"
  Print #PRNFile, "--------------------------------------------"
Return
PrintState:


' ' If Combined Then
'    For cnt = 0 To 999
'      If StSalesTaxPaid#(cnt) > 0 Then
'        StateAmt# = (StSalesTaxPaid#(cnt) * StateFactor#)
'        CountyAmt# = (StSalesTaxPaid#(cnt) - StateAmt#)
'        StTotTax# = StTotTax# + StSalesTaxPaid#(cnt)
'        TotState# = TotState# + StateAmt#
'        TotCounty# = TotCounty# + CountyAmt#
'        Print #PRNFile, cnt; Tab(20); Using("######.##", StateAmt#); Tab(35); Using("######.##", CountyAmt#)
'      End If
'    Next
'    Print #PRNFile, "--------------------------------------------"
'    Print #PRNFile, "Totals"; Tab(20); Using("######.##", TotState#); Tab(35); Using("######.##", TotCounty#)
'  Else
    For cnt = 0 To 999
      If StSalesTaxPaid#(cnt) > 0 Or CoSalesTaxPaid#(cnt) > 0 Then
        Print #PRNFile, ; Tab(20); StateCd; Tab(35); Using("#####.##", CoSalesTaxPaid#(cnt))
        'StTotTax# = StTotTax# + StSalesTaxPaid#(cnt)
        CoTotTax# = CoTotTax# + CoSalesTaxPaid#(cnt)
      End If
    Next
'  End If

  
  Return
CancelExit:
  Exit Sub
End Sub

Private Function VendcoCodeList(x As fpCombo)
  Dim cnt As Integer, VendorFile As Integer, NumVRecs As Integer
  Dim Vendor As VendorRecType
    x.AddItem "All"
    OpenVendorFile VendorFile, NumVRecs
    For cnt = 1 To NumVRecs
      Get VendorFile, cnt, Vendor
      If Not Vendor.DelFlag Then
        x.SearchText = QPTrim(Vendor.CoCode)
        x.Action = 0
        If x.SearchIndex = -1 Then
          If Len(QPTrim(Vendor.CoCode)) > 0 Then
            x.AddItem Vendor.CoCode
          End If
        End If
      End If
    Next
   
Close VendorFile
End Function


Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub

Private Sub SalesTaxReportNOSTATE()
  Dim Combined As Boolean, StateTaxRecAcct As String, CountyTaxRecAcct As String
  Dim StateFactor As Double, BegDate As Integer, EndDate As Integer
  Dim State As Integer, PrintFlag As Integer, StTotTax As Double
  Dim TotState As Double, TotCounty As Double, TCnt As Integer
  Dim CoTotTax As Double, ll As Integer, VendorFile As Integer
  Dim NumVRecs As Integer, LdRecLen As Integer, APLedgerFile As Integer
  Dim NumTran As Long, APDRecLen As Integer, APDistFile As Integer
  Dim NumDistRecs As Long, cnt As Integer, NextTran As Long, SCnt As Integer
  Dim NextDist As Long, offset As Integer, Coffset As Integer
  Dim StTax As Double, ccnt As Integer, CoTax As Double, PRNFile As Integer
  Dim ReportFile As String, Header As String, TotList As Integer
  Dim StateAmt As Double, CountyAmt As Double, StateCd As String
  Dim all As Boolean, Newrp As String, ToPrint As String, User As String
  Dim ToPrint1 As String
  'If fpcboCombined.ListIndex = 1 Then Combined = True
 ' fpcboStAcct.col = 1
  fpcboCoAcct.col = 1
  CoTotTax# = 0
  User$ = QPTrim(GLUserName$)
 ' StateTaxRecAcct$ = QPTrim$(fpcboStAcct.ColText)
  CountyTaxRecAcct$ = QPTrim$(fpcboCoAcct.ColText)
 ' StateFactor# = Val(fpdblFactor)
  BegDate = DateDiff("d", "12/31/1979", fpDate1)
  EndDate = DateDiff("d", "12/31/1979", fpDate2)
  PRNFile = FreeFile
  Newrp = "SalTax.prn"
  'GetRPTName Newrp
  ReportFile$ = Newrp
  Open ReportFile$ For Output As #PRNFile
  Header$ = "Sales Tax Report"

  If fpcboStateCode.ListIndex <= 0 Then
    StateCd = 0
    all = True
    TotList = fpcboStateCode.ListCount - 1
  Else
    StateCd = QPTrim(fpcboStateCode.Text)
    all = False
    TotList = 1
  End If
    FrmShowPctComp.Label1 = "Searching Codes For Sales Tax Report"
    FrmShowPctComp.Show , Me
    DoEvents
   DeActivateControls frmPrnSalesTaxRptCounty, True
  ReDim StSalesTaxPaid#(0 To 999)
  'If Not Combined Then
    ReDim CoSalesTaxPaid#(0 To 999)
 ' End If
  For State = 1 To TotList
    PrintFlag = 0
    FrmShowPctComp.ShowPctComp State, TotList
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ActivateControls frmPrnSalesTaxRptCounty, True
      Unload FrmShowPctComp
      GoTo CancelExit
    End If
 
    If all Then
      fpcboStateCode.Row = State
      StateCd = QPTrim(fpcboStateCode.List)
    End If
   ' PrintFlag = 0
    StTotTax# = 0
    TotState# = 0
    TotCounty# = 0
    TCnt = 0
    '
    For ll = 0 To 999
      StSalesTaxPaid#(ll) = 0
      If Not Combined Then
        CoSalesTaxPaid#(ll) = 0
      End If
    Next ll

    Dim Vendor As VendorRecType
    Close VendorFile
    OpenVendorFile VendorFile, NumVRecs

    Dim ApLedger As APLedger81RecType
    LdRecLen = Len(ApLedger)
    Close APLedgerFile
    OpenAPLedgerFile APLedgerFile, NumTran&, LdRecLen
    Dim APDist As APDistRecType
    APDRecLen = Len(APDist)
    Close APDistFile
    OpenAPDistFile APDistFile, NumDistRecs&, APDRecLen

    For cnt = 1 To NumVRecs
      Get VendorFile, cnt, Vendor
      NextTran = Vendor.FrstTran
      If NextTran > 0 Then
        Do
          Get APLedgerFile, NextTran, ApLedger
          If ApLedger.TRCode = 1 Then
            '--check this out
            If ApLedger.GLDistDate >= BegDate And ApLedger.GLDistDate <= EndDate Then
              TCnt = TCnt + 1
              NextDist& = ApLedger.FrstDist
              Do
                Get APDistFile, NextDist&, APDist

'                If QPTrim$(APDist.DistAcctNum) = StateTaxRecAcct$ Then
'                  Get VendorFile, APLedger.VRecNum, Vendor
'                  'offset = State
'
'                  Coffset = Val(Vendor.CoCode)
'
'                  If StateCd = QPTrim(Vendor.StCode) Then
'                    SCnt = SCnt + 1
'                    StTax# = StTax# + APDist.DistAmt
'                    PrintFlag = 1
'                    'offset = State
'                    If all Then offset = 0
'                    StSalesTaxPaid#(Coffset) = StSalesTaxPaid#(Coffset) + APDist.DistAmt
'                    'LPRINT APDist.DistAmt
'                  End If
'                End If
                'If Not Combined Then
                  If QPTrim$(APDist.DistAcctNum) = CountyTaxRecAcct$ Then
                    Get VendorFile, ApLedger.VRecNum, Vendor
                    If StateCd = QPTrim(Vendor.CoCode) Then
                      ccnt = ccnt + 1
                      CoTax# = CoTax# + APDist.DistAmt
                      PrintFlag = 1
                      If all Then
                        offset = State
                      Else
                        offset = 1
                      End If
                      'If offset < 0 Or offset > 999 Then offset = 0
                      '  LPRINT Vendor.Vname, APDist.DistAmt
                      CoSalesTaxPaid#(Coffset) = CoSalesTaxPaid#(Coffset) + APDist.DistAmt
                    End If
                  End If
                'End If
                NextDist& = APDist.NextDist
              Loop Until NextDist& = 0

            End If

          End If
          NextTran = ApLedger.NextTrans
        Loop Until NextTran = 0
      End If

    Next cnt

    If PrintFlag = 1 Then
      GoSub PrintState
    End If
      '
    'End If
' "State Code Being Searched: ";: Color 15: Print State
  Next State
  If TotList < 1 Then
    FrmShowPctComp.ShowPctComp 1, 1
    MsgBox "No Information to Display", vbOKOnly, "No Information"
  End If
  Close
  Unload FrmShowPctComp
  Load frmLoadingRpt
  ActivateControls frmPrnSalesTaxRptCounty, True
  ARptSalesTaxNoState.GetName ReportFile$
  ARptSalesTaxNoState.Label1.Caption = "Sales Tax Report - " + CountyTaxRecAcct$
  ARptSalesTaxNoState.txtDate.Caption = Now
  ARptSalesTaxNoState.txtTown.Caption = User$
  ARptSalesTaxNoState.Label17.Caption = "Reporting : " + fpDate1.Text + " thru " + fpDate2.Text
  ARptSalesTaxNoState.startrpt

 ' ViewPrint ReportFile$, Header$
 ' KillFile ReportFile$
  
  Exit Sub

PrintState:
  ToPrint$ = Space(80)
  
  ToPrint$ = "1~"

    For cnt = 0 To 999
      If StSalesTaxPaid#(cnt) > 0 Or CoSalesTaxPaid#(cnt) > 0 Then
        'ToPrint1$ = Space(80)
        ToPrint1$ = ToPrint$ + StateCd + "~" + (Using("######.##", StSalesTaxPaid#(cnt))) + "~" + (Using("######.##", CoSalesTaxPaid#(cnt)))
        Print #PRNFile, ToPrint1$
        ToPrint$ = ""
        'StTotTax# = StTotTax# + StSalesTaxPaid#(cnt)
        CoTotTax# = CoTotTax# + CoSalesTaxPaid#(cnt)
      End If
    Next
  

  Return
CancelExit:
  Exit Sub
End Sub

