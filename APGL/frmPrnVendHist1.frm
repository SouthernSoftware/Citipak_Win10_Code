VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPrnVendHist1 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vendor History"
   ClientHeight    =   8616
   ClientLeft      =   36
   ClientTop       =   540
   ClientWidth     =   12192
   Icon            =   "frmPrnVendHist1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8616
   ScaleWidth      =   12192
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboRptType 
      Height          =   384
      Left            =   5112
      TabIndex        =   4
      Top             =   5328
      Width           =   1908
      _Version        =   196608
      _ExtentX        =   3365
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
      ColDesigner     =   "frmPrnVendHist1.frx":08CA
   End
   Begin LpLib.fpCombo fpcboVend1 
      Height          =   384
      Left            =   5136
      TabIndex        =   0
      Top             =   2856
      Width           =   4140
      _Version        =   196608
      _ExtentX        =   7302
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
      Columns         =   3
      Sorted          =   0
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   0
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
      EditMarginLeft  =   5
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmPrnVendHist1.frx":0C30
   End
   Begin LpLib.fpCombo fpcboVend2 
      Height          =   384
      Left            =   5136
      TabIndex        =   1
      Top             =   3480
      Width           =   4140
      _Version        =   196608
      _ExtentX        =   7302
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
      Columns         =   3
      Sorted          =   0
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   0
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
      EditMarginLeft  =   5
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmPrnVendHist1.frx":0FE3
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H008F8265&
      Caption         =   "Include Inactives:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   372
      Left            =   3072
      TabIndex        =   5
      Top             =   5928
      Width           =   2244
   End
   Begin VB.CommandButton cmdGo 
      BackColor       =   &H00D0D0D0&
      Caption         =   "F10 &Go"
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
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7440
      UseMaskColor    =   -1  'True
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
      Left            =   10032
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7440
      UseMaskColor    =   -1  'True
      Width           =   1332
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   8
      Top             =   8256
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
            TextSave        =   "11:18 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7133
            TextSave        =   "11/16/2004"
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
   Begin EditLib.fpDateTime fpDate1 
      Height          =   372
      Left            =   5136
      TabIndex        =   2
      Top             =   4092
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
      Top             =   4716
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
   Begin VB.Label Label4 
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
      TabIndex        =   14
      Top             =   5352
      Width           =   2388
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Starting Vendor:"
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
      Left            =   2976
      TabIndex        =   13
      Top             =   2928
      Width           =   2004
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   4428
      Left            =   1920
      Top             =   2352
      Width           =   8316
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Print Vendor History Report"
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
      Left            =   3984
      TabIndex        =   12
      Top             =   1176
      Width           =   4332
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3216
      Top             =   936
      Width           =   5772
   End
   Begin VB.Image Image1 
      Height          =   276
      Left            =   2496
      Picture         =   "frmPrnVendHist1.frx":1396
      Top             =   2736
      Width           =   288
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
      Height          =   420
      Left            =   3408
      TabIndex        =   11
      Top             =   4764
      Width           =   1572
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
      Height          =   420
      Left            =   3312
      TabIndex        =   10
      Top             =   4152
      Width           =   1668
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ending Vendor:"
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
      Left            =   3168
      TabIndex        =   9
      Top             =   3540
      Width           =   1812
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D0D0D0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00D0D0D0&
      Height          =   972
      Left            =   3216
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
Attribute VB_Name = "frmPrnVendHist1"
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
  Unload frmPrnVendHist
End Sub

Private Sub cmdGo_Click()
 If Oktogo = True Then
  If fpcboRptType.ListIndex = 0 Then
    rptopt = 1
  ElseIf fpcboRptType.ListIndex = 1 Then
    rptopt = 2
  End If
  If rptopt = 1 Then
    VendorHistory
  ElseIf rptopt = 2 Then
    VendorHistory2
  End If
 End If
End Sub

Private Sub fpcboRptType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboRptType.ListDown = True
  End If
  If fpcboRptType.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      cmdGo.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpDate2.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub

Private Sub fpcboVend1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboVend1.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcboVend1.ListIndex = -1
    fpcboVend1.Action = ActionClearSearchBuffer
  End If
  If fpcboVend1.ListDown <> True Then
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

Private Sub fpcboVend1_LostFocus()
  fpcboVend1.Action = ActionClearSearchBuffer
End Sub
Private Sub fpcboVend2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboVend2.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcboVend2.ListIndex = -1
    fpcboVend2.Action = ActionClearSearchBuffer
  End If
  If fpcboVend2.ListDown <> True Then
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

Private Sub fpcboVend2_LostFocus()
  fpcboVend2.Action = ActionClearSearchBuffer
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
      SendKeys "%G"
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
  VendCodeNameIA fpcboVend1
  VendCodeNameIA fpcboVend2
  fpcboVend1.ListIndex = 0
  fpcboVend2.ListIndex = fpcboVend2.ListCount - 1
  fpDate1.Text = Format(Now, "mm/dd/yyyy")
  fpDate2.Text = Format(Now, "mm/dd/yyyy")
  fpcboRptType.InsertRow = "Graphics"
  fpcboRptType.InsertRow = "Text"
  fpcboRptType.ListIndex = 0
End Sub
Private Function Oktogo()
Dim TempDate1 As Integer, TempDate2 As Integer
If fpcboVend1.ListIndex <> -1 And fpcboVend2.ListIndex <> -1 Then
  If fpcboVend1.ListIndex <= fpcboVend2.ListIndex Then
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
  Else
    MsgBox "Invalid Vendor Selection, The Vendor Selection Should Be Equal or in Ascending Order.", vbOKOnly, "Invalid Selection"
    Oktogo = False
  End If
 Else
   MsgBox "You Must Select A Vendor, Retry", vbOKOnly, "Invalid Selection"
   Oktogo = False
 End If
End Function
Private Sub Form_Resize()
'  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
'  End If
End Sub
Private Sub VendorHistory()
  Dim MaxLines As Integer, VendorIdxFile As Integer, NumActiveVendors As Integer
  Dim VendorFile As Integer, NumVRecs As Integer, Linecnt As Integer
  Dim PRNFile As Integer, cnt As Integer, HowMany As Integer, RptParm1 As String
  Dim ReportFile As String, ToPrint As String, Page As Integer
  Dim FF As String, Header As String, Dash As String, User As String
  Dim FromDate As Integer, thrudate As Integer, Vend1St As String
  Dim FundCode As String, VendLst As String, NumVendRecs As Integer
  Dim ChkFund As Boolean, LdRecLen As Integer, VCnt As Integer
  Dim VendIdxRecLen As Integer, APLedgerFile As Integer, Newrp As String
  Dim NumTran As Long, VCode As String, TransCnt As Long, VTBal As Double
  Dim VTDebit As Double, VTCredit As Double, POCredit As Double
  Dim PODebit As Double, POTBal As Double, TInvAmt As Double, TDate As String
  Dim NextTrans As Long, InRange As Boolean, Ten99 As String, RptParm2 As String
  Dim fmt As String, DistRecLEn As Integer, NumDistRecs As Long
  Dim APDistFile As Integer, nexdist As Long, POAmt As Double
  Dim Vactive As String, doone As Boolean
  Dim APDistRec(1) As APDistRecType
  DistRecLEn = Len(APDistRec(1))
  FrmShowPctComp.Label1 = "Creating Vendor History Report"
  FrmShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdExit.Enabled = False
  Me.cmdGo.Enabled = False
  mnuOptions.Enabled = False
  Dash$ = String$(78, "-")
  FF$ = Chr$(12)
  VendIdxRecLen = Len(VendorIdx)
  ReDim APLedgerRec(1) As APLedger81RecType
  LdRecLen = Len(APLedgerRec(1))
  Header$ = "Vendor History"
  MaxLines = 52
  Linecnt = 0
  User$ = QPTrim$(GLUserName$)
  Page = 0
  fmt$ = "$###,###.##"
  FromDate = DateDiff("d", "12/31/1979", fpDate1.Text)
  thrudate = DateDiff("d", "12/31/1979", fpDate2.Text)
  fpcboVend1.col = 0
  fpcboVend2.col = 0
  Vend1St$ = QPTrim$(fpcboVend1.ColText)
  VendLst$ = QPTrim$(fpcboVend2.ColText)

  NumVendRecs = (FileSize("apvendor.idx") \ VendIdxRecLen)
  ReDim VIndex(1 To NumVendRecs) As VendorIdxRecType
  OpenVendorIdx VendorIdxFile, NumActiveVendors
  For VCnt = 1 To NumVendRecs
    Get VendorIdxFile, VCnt, VendorIdx
    VIndex(VCnt).VendorCode = VendorIdx.VendorCode
    VIndex(VCnt).RecNum = VendorIdx.RecNum
  Next
  Close VendorIdxFile

  PRNFile = FreeFile
  'Newrp = "VHist"
  'GetRPTName Newrp
  ReportFile$ = "VendHist.PRN"
  Open ReportFile$ For Output As #PRNFile
'@@@@@@@@@@@@@@@@@@@@@
'what is this ?????
'  'dale
'  Open "DEADVEND.LST" For Output As #20
'  BDate = Date2Num("12/31/1997")
'  'dale
'@@@@@@@@@@@@@@@@@@@@@@@
  OpenVendorFile VendorFile, NumVRecs
  OpenAPLedgerFile APLedgerFile, NumTran&, LdRecLen
  GoSub VendHistHeader
  OpenAPDistFile APDistFile, NumDistRecs&, DistRecLEn

  For cnt = 1 To NumVendRecs
    FrmShowPctComp.ShowPctComp cnt, NumVendRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Me.cmdExit.Enabled = True
      Me.cmdGo.Enabled = True
      mnuOptions.Enabled = True
      Unload FrmShowPctComp
      GoTo CancelExit
    End If

    VCode$ = QPTrim$(VIndex(cnt).VendorCode)
    If VCode$ >= Vend1St$ And VCode$ <= VendLst$ Then
      TransCnt = 0
      VTBal# = 0
      VTDebit# = 0
      VTCredit# = 0
      POCredit# = 0
      PODebit# = 0
      POTBal# = 0
      Get VendorFile, VIndex(cnt).RecNum, Vendor
      TInvAmt# = 0
      If Vendor.DelFlag = 0 Then
        If Check1.Value = 1 Then
          doone = True
          If Vendor.ActiveFlag = 0 Then
            Vactive$ = "Active"
          Else
            Vactive$ = "Inactive"
          End If
        End If
        If Check1.Value = 0 Then
          If Vendor.ActiveFlag = 0 Then
            Vactive$ = "Active"
            doone = True
          Else
            doone = False
          End If
        End If
        If doone = True Then
'        If Linecnt >= MaxLines Then
'          Print #PRNFile, FF$
'          GoSub VendHistHeader
'        End If

'        Print #PRNFile, "Vendor: "; Vendor.VNAME
'        Linecnt = Linecnt + 1

        NextTrans& = Vendor.FrstTran

        Do Until NextTrans& = 0
          Get APLedgerFile, NextTrans&, APLedgerRec(1)
          InRange = False
          If APLedgerRec(1).TRDATE >= FromDate And APLedgerRec(1).TRDATE <= thrudate Then
            InRange = True
            TransCnt = TransCnt + 1
          End If
          'TransCnt = TransCnt + 1
'          If Linecnt >= MaxLines Then
'            Print #PRNFile, FF$
'            GoSub VendHistHeader
'            Print #PRNFile, "Vendor: "; Vendor.VNAME
'          End If
          If InRange Then
            TDate = Format(DateAdd("d", (APLedgerRec(1).TRDATE), "12-31-1979"), "mm/dd/yyyy")
            ToPrint$ = Vendor.VNAME + " - " + Vactive$ + "~" + TDate + "~"
          
          
          If APLedgerRec(1).TRCode = 1 Then
            VTCredit# = Round(VTCredit# + APLedgerRec(1).Amt)
            If APLedgerRec(1).PDCheckDate > 0 Then
              'PRINT #PRNFile, "Inv "; LEFT$(APLedgerRec(1).DOCNum, 20); APLed
              If APLedgerRec(1).Get1099 = "Y" Then
                Ten99$ = "1099: Y"
              ElseIf APLedgerRec(1).Get1099 = "N" Then
                Ten99$ = "1099: N"
              Else
                Ten99$ = "1099:  "
              End If
              If InRange Then
                ToPrint$ = ToPrint$ + Str(APLedgerRec(1).TRCode) + "~" + "Inv " + "~" + QPTrim$(APLedgerRec(1).DOCNum) + "/" + QPTrim(APLedgerRec(1).Comment) + "~" + Ten99$ + "~~" + Using(fmt$, Str$(APLedgerRec(1).Amt)) + "~" + "  Pd" + Str(APLedgerRec(1).PDCheckNum)
                Print #PRNFile, ToPrint$
                ' #PrnFile, "1099: "; APLedgerRec(1).Get1099
              End If
            Else
              If InRange Then
                ToPrint$ = ToPrint$ + Str(APLedgerRec(1).TRCode) + "~" + "Inv " + "~" + QPTrim$(APLedgerRec(1).DOCNum) + "/" + QPTrim(APLedgerRec(1).Comment) + "~" + Ten99$ + "~~" + Using(fmt$, Str$(APLedgerRec(1).Amt)) + "~" + "  Open"
                Print #PRNFile, ToPrint$
                ' "1099: "; APLedgerRec(1).Get1099
              End If
            End If
            If InRange Then
              TInvAmt# = Round#(TInvAmt# + APLedgerRec(1).Amt)
            End If

            '--02/27/97 to handle voided invoices
          ElseIf APLedgerRec(1).TRCode = -1 Then
              If APLedgerRec(1).Get1099 = "Y" Then
                Ten99$ = "1099: Y"
              ElseIf APLedgerRec(1).Get1099 = "N" Then
                Ten99$ = "1099: N"
              Else
                Ten99$ = "1099:  "
              End If
            If InRange Then
              ToPrint$ = ToPrint$ + Str(APLedgerRec(1).TRCode) + "~" + "Inv " + "~" + QPTrim$(APLedgerRec(1).DOCNum) + "/" + QPTrim(APLedgerRec(1).Comment) + "~" + Ten99$ + "~~" + Using(fmt$, Str$(APLedgerRec(1).Amt)) + "~" + "  Void"
              Print #PRNFile, ToPrint$
            End If
            '--02/27/97

          ElseIf APLedgerRec(1).TRCode = 4 Then
            POAmt# = 0
          'Have to calc the total of uncleared items, not just total of po
             nexdist& = APLedgerRec(1).FrstDist
             Do Until nexdist& = 0
               Get APDistFile, nexdist&, APDistRec(1)
               If APDistRec(1).DistStat <> "L" Then
                 POCredit# = Round(POCredit# + APDistRec(1).DistAmt)
                 POAmt# = Round(POAmt# + APDistRec(1).DistAmt)
               End If
               nexdist& = APDistRec(1).NextDist
              Loop
              'POCredit# = Round(POCredit# + APLedgerRec(1).Amt)
            If InRange Then
              If POAmt# <> APLedgerRec(1).Amt Then
                ToPrint$ = ToPrint$ + Str(APLedgerRec(1).TRCode) + "~" + "PO  " + "~" + Left$(APLedgerRec(1).DOCNum, 20) + "~~~" + Using(fmt$, Str$(POAmt#)) + "~" + "  Open Partial"
                Print #PRNFile, ToPrint$
              Else
                ToPrint$ = ToPrint$ + Str(APLedgerRec(1).TRCode) + "~" + "PO  " + "~" + Left$(APLedgerRec(1).DOCNum, 20) + "~~~" + Using(fmt$, Str$(APLedgerRec(1).Amt)) + "~" + "  Open"
                Print #PRNFile, ToPrint$
              End If
            End If
          ElseIf APLedgerRec(1).TRCode = -4 Then                'for cleared P
            PODebit# = Round(PODebit# + APLedgerRec(1).Amt)
            If InRange Then
              ToPrint$ = ToPrint$ + Str(APLedgerRec(1).TRCode) + "~" + "PO  " + "~" + Left$(APLedgerRec(1).DOCNum, 20) + "~~~" + Using(fmt$, Str$(APLedgerRec(1).Amt)) + "~" + "  Closed"
              Print #PRNFile, ToPrint$
            End If
          ElseIf APLedgerRec(1).TRCode = 3 Then
            VTDebit# = Round(VTDebit# + APLedgerRec(1).Amt)
            If InRange Then
              ToPrint$ = ToPrint$ + Str(APLedgerRec(1).TRCode) + "~" + "Chk " + "~" + Left$(APLedgerRec(1).DOCNum, 20) + "~~" + Using(fmt$, Str$(APLedgerRec(1).Amt)) + "~~~"
              Print #PRNFile, ToPrint$
            End If
          ElseIf APLedgerRec(1).TRCode = -3 Then
            If InRange Then
              ToPrint$ = ToPrint$ + Str(APLedgerRec(1).TRCode) + "~" + "Chk " + "~" + Left$(APLedgerRec(1).DOCNum, 20) + "~" + " VOID" + "~" + Using(fmt$, Str$(APLedgerRec(1).Amt)) + "~~~"
              Print #PRNFile, ToPrint$
            End If
          Else
            If InRange Then
              'Print #PRNFile, "Code is:"; APLedgerRec(1).TrCode
            End If
          End If
          If InRange Then
            'Linecnt = Linecnt + 1
            'PRINT #PrnFile, "PSL: !"; APLedgerRec(1).PSLFlag; "!"
          End If

          'PRINT #PrnFile, "1099? "; APLedgerRec(1).Get1099
          End If
          NextTrans& = APLedgerRec(1).NextTrans

        Loop

        'dale
'        If TransCnt = 0 Or APLedgerRec(1).TRDATE <= BDate Then
'          Print #20, "Vendor: "; Vendor.VNAME; Tab(40);
          If TransCnt = 0 Then
            ToPrint$ = Vendor.VNAME + " - " + Vactive$ + "~~~~~~~~~"
            Print #PRNFile, ToPrint$
          End If
'          Else
'            Print #20, "last: "; Format(DateAdd("d", (APLedgerRec(1).TRDATE), "12-31-1979"), "mm/dd/yyyy")
'          End If
'
'          DLCnt = DLCnt + 1
'          If DLCnt > 58 Then
'            Print #20, Chr$(12)
'            DLCnt = 0
'          End If
'        End If
'        'dale

'        If TransCnt = 0 Then
'          Print #PRNFile, "  NO TRANSACTIONS."
'          Print #PRNFile, Dash$
'          Linecnt = Linecnt + 2
'        Else
'          VTBal# = Round(VTCredit# - VTDebit#)
'          Print #PRNFile, "Current Balance :"; Using(fmt$, Str$(VTBal#));
'          Print #PRNFile, Tab(40); "Range Invoice Total :"; Using(fmt$, Str$(TInvAmt#))
'          Print #PRNFile, "Current On Order:"; Using(fmt$, Str$(POCredit#))
'
'          Print #PRNFile, Dash$
'          Linecnt = Linecnt + 2
'        End If
      End If
      End If
    End If
  Next

  'Print #PRNFile, FF$

'  Print #PRNFile, Tab(40 - (Int(Len(User$) / 2))); User$
'  Print #PRNFile, Tab(40 - (Int(Len(Header$) / 2))); Header$
'  Print #PRNFile, "Report Parameters: "
'  Print #PRNFile, "=============================================================================="

  'Print #PRNFile, "From Vendor: "; Vend1St$; "   Thru Vendor: "; VendLst$
  If FromDate = -32768 Then
    FromDate = 0
  End If

'  Print #PRNFile, "  From Date: "; fpDate1.Text; "     Thru Date: "; fpDate2.Text
'  Print #PRNFile, FF$
  If NumVRecs < 1 Then
    FrmShowPctComp.ShowPctComp 1, 1
    'Print #PRNFile, "No Information To Print"
  End If
  Close
  Erase VIndex
  Load frmLoadingRpt
  RptParm1$ = "From Vendor: " + Vend1St$ + "   Thru Vendor: " + VendLst$
  RptParm2$ = "From Date: " + fpDate1.Text + "   Thru Date: " + fpDate2.Text
  ARptVendHist.txtRptParm1.Caption = RptParm1$
  ARptVendHist.txtRptParm2.Caption = RptParm2$
  ARptVendHist.GetName ReportFile$
  ActivateControls frmPrnVendHist, True
  ARptVendHist.txtTown.Caption = GLUserName$
  ARptVendHist.txtDate.Caption = Now
  ARptVendHist.startrpt
  'ViewPrint ReportFile$, "Vendor History Report"
  'KillFile ReportFile$
  Me.cmdExit.Enabled = True
  Me.cmdGo.Enabled = True
  EnableCloseButton Me.hwnd, True
  mnuOptions.Enabled = True
  'If ListCnt = 0 Then Exit Sub
  'add a trap here to display an error scrn if no matching ledger recs
  'to pay

ExitSelPayables:

  Exit Sub

VendHistHeader:
'  Page = Page + 1
'  Print #PRNFile, Tab(40 - (Int(Len(User$) / 2))); User$
'  Print #PRNFile, Tab(40 - (Int(Len(Header$) / 2))); Header$
'  Print #PRNFile,
'  Print #PRNFile, "Report Date: "; Date$; Tab(67); "Page #"; Page
'  Print #PRNFile, "Inv Date    TrCode  Desc                        Debit      Credit  Status"
'  Print #PRNFile, "=============================================================================="
'  Linecnt = 6
Return
CancelExit:
  Exit Sub
End Sub
Private Sub VendorHistory2()
  Dim MaxLines As Integer, VendorIdxFile As Integer, NumActiveVendors As Integer
  Dim VendorFile As Integer, NumVRecs As Integer, Linecnt As Integer
  Dim PRNFile As Integer, cnt As Integer, HowMany As Integer
  Dim ReportFile As String, ToPrint As String, Page As Integer
  Dim FF As String, Header As String, Dash As String, User As String
  Dim FromDate As Integer, thrudate As Integer, Vend1St As String
  Dim FundCode As String, VendLst As String, NumVendRecs As Integer
  Dim ChkFund As Boolean, LdRecLen As Integer, VCnt As Integer
  Dim VendIdxRecLen As Integer, APLedgerFile As Integer, Newrp As String
  Dim NumTran As Long, VCode As String, TransCnt As Long, VTBal As Double
  Dim VTDebit As Double, VTCredit As Double, POCredit As Double
  Dim PODebit As Double, POTBal As Double, TInvAmt As Double
  Dim NextTrans As Long, InRange As Boolean, Ten99 As String
  Dim fmt As String, DistRecLEn As Integer, NumDistRecs As Long
  Dim APDistFile As Integer, nexdist As Long, POAmt As Double
  Dim Vactive As String, doone As Boolean
  Dim APDistRec(1) As APDistRecType
  DistRecLEn = Len(APDistRec(1))
  FrmShowPctComp.Label1 = "Creating Vendor History Report"
  FrmShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdExit.Enabled = False
  Me.cmdGo.Enabled = False
  mnuOptions.Enabled = False
  Dash$ = String$(78, "-")
  FF$ = Chr$(12)
  VendIdxRecLen = Len(VendorIdx)
  ReDim APLedgerRec(1) As APLedger81RecType
  LdRecLen = Len(APLedgerRec(1))
  Header$ = "Vendor History"
  MaxLines = 52
  Linecnt = 0
  User$ = QPTrim$(GLUserName$)
  Page = 0
  fmt$ = "$###,###.##"
  FromDate = DateDiff("d", "12/31/1979", fpDate1.Text)
  thrudate = DateDiff("d", "12/31/1979", fpDate2.Text)
  fpcboVend1.col = 0
  fpcboVend2.col = 0
  Vend1St$ = QPTrim$(fpcboVend1.ColText)
  VendLst$ = QPTrim$(fpcboVend2.ColText)

  NumVendRecs = (FileSize("apvendor.idx") \ VendIdxRecLen)
  ReDim VIndex(1 To NumVendRecs) As VendorIdxRecType
  OpenVendorIdx VendorIdxFile, NumActiveVendors
  For VCnt = 1 To NumVendRecs
    Get VendorIdxFile, VCnt, VendorIdx
    VIndex(VCnt).VendorCode = VendorIdx.VendorCode
    VIndex(VCnt).RecNum = VendorIdx.RecNum
  Next
  Close VendorIdxFile

  PRNFile = FreeFile
  Newrp = "VHist"
  GetRPTName Newrp
  ReportFile$ = Newrp
  Open ReportFile$ For Output As #PRNFile
'@@@@@@@@@@@@@@@@@@@@@
'what is this ?????
'  'dale
'  Open "DEADVEND.LST" For Output As #20
'  BDate = Date2Num("12/31/1997")
'  'dale
'@@@@@@@@@@@@@@@@@@@@@@@
  OpenVendorFile VendorFile, NumVRecs
  OpenAPLedgerFile APLedgerFile, NumTran&, LdRecLen
  GoSub VendHistHeader
  OpenAPDistFile APDistFile, NumDistRecs&, DistRecLEn

  For cnt = 1 To NumVendRecs
    FrmShowPctComp.ShowPctComp cnt, NumVendRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Me.cmdExit.Enabled = True
      Me.cmdGo.Enabled = True
      mnuOptions.Enabled = True
      Unload FrmShowPctComp
      GoTo CancelExit
    End If

    VCode$ = QPTrim$(VIndex(cnt).VendorCode)
    If VCode$ >= Vend1St$ And VCode$ <= VendLst$ Then
      TransCnt = 0
      VTBal# = 0
      VTDebit# = 0
      VTCredit# = 0
      POCredit# = 0
      PODebit# = 0
      POTBal# = 0
      Get VendorFile, VIndex(cnt).RecNum, Vendor
      TInvAmt# = 0
      If Vendor.DelFlag = 0 Then
        If Check1.Value = 1 Then
          doone = True
          If Vendor.ActiveFlag = 0 Then
            Vactive$ = "Active"
          Else
            Vactive$ = "Inactive"
          End If
        End If
        If Check1.Value = 0 Then
          If Vendor.ActiveFlag = 0 Then
            Vactive$ = "Active"
            doone = True
          Else
            doone = False
          End If
        End If
        If doone = True Then
        If Linecnt >= MaxLines Then
          Print #PRNFile, FF$
          GoSub VendHistHeader
        End If

        Print #PRNFile, "Vendor: "; Vendor.VNAME; " - "; Vactive$
        Linecnt = Linecnt + 1

        NextTrans& = Vendor.FrstTran

        Do Until NextTrans& = 0
          Get APLedgerFile, NextTrans&, APLedgerRec(1)
          InRange = False
          If APLedgerRec(1).TRDATE >= FromDate And APLedgerRec(1).TRDATE <= thrudate Then
            InRange = True
            TransCnt = TransCnt + 1
          End If
          'TransCnt = TransCnt + 1
          If Linecnt >= MaxLines Then
            Print #PRNFile, FF$
            GoSub VendHistHeader
            Print #PRNFile, "Vendor: "; Vendor.VNAME; " - " + Vactive$
          End If
          If InRange Then
            Print #PRNFile, Format(DateAdd("d", (APLedgerRec(1).TRDATE), "12-31-1979"), "mm/dd/yyyy"); "  ";
          End If
          
          If APLedgerRec(1).TRCode = 1 Then
            VTCredit# = Round(VTCredit# + APLedgerRec(1).Amt)
            If APLedgerRec(1).PDCheckDate > 0 Then
              'PRINT #PRNFile, "Inv "; LEFT$(APLedgerRec(1).DOCNum, 20); APLed
              If APLedgerRec(1).Get1099 = "Y" Then
                Ten99$ = "1099: Y"
              ElseIf APLedgerRec(1).Get1099 = "N" Then
                Ten99$ = "1099: N"
              Else
                Ten99$ = "1099:  "
              End If
              If InRange Then
                Print #PRNFile, "Inv "; Left$(APLedgerRec(1).DOCNum, 20); Ten99$
                Print #PRNFile, Tab(15); QPTrim(APLedgerRec(1).Comment); Tab(54); Using(fmt$, Str$(APLedgerRec(1).Amt)); "  Pd"; APLedgerRec(1).PDCheckNum

                'PRINT #PrnFile, "1099: "; APLedgerRec(1).Get1099
              End If
            Else
              If InRange Then
                Print #PRNFile, "Inv "; Left$(APLedgerRec(1).DOCNum, 20); Ten99$
                Print #PRNFile, Tab(15); QPTrim(APLedgerRec(1).Comment); Tab(54); Using(fmt$, Str$(APLedgerRec(1).Amt)); "  Open"

                'PRINT #PrnFile, "1099: "; APLedgerRec(1).Get1099
              End If
            End If
            If InRange Then
              TInvAmt# = Round#(TInvAmt# + APLedgerRec(1).Amt)
            End If

            '--02/27/97 to handle voided invoices
          ElseIf APLedgerRec(1).TRCode = -1 Then
            If APLedgerRec(1).Get1099 = "Y" Then
              Ten99$ = "1099: Y"
            ElseIf APLedgerRec(1).Get1099 = "N" Then
              Ten99$ = "1099: N"
            Else
              Ten99$ = "1099:  "
            End If

            If InRange Then
              Print #PRNFile, "Inv "; Left$(APLedgerRec(1).DOCNum, 20); Ten99$
              Print #PRNFile, Tab(15); QPTrim(APLedgerRec(1).Comment); Tab(54); Using(fmt$, Str$(APLedgerRec(1).Amt)); "  Void"

            End If
            '--02/27/97

          ElseIf APLedgerRec(1).TRCode = 4 Then
            POAmt# = 0
          'Have to calc the total of uncleared items, not just total of po
             nexdist& = APLedgerRec(1).FrstDist
             Do Until nexdist& = 0
               Get APDistFile, nexdist&, APDistRec(1)
               If APDistRec(1).DistStat <> "L" Then
                 POCredit# = Round(POCredit# + APDistRec(1).DistAmt)
                 POAmt# = Round(POAmt# + APDistRec(1).DistAmt)
               End If
               nexdist& = APDistRec(1).NextDist
              Loop
              'POCredit# = Round(POCredit# + APLedgerRec(1).Amt)
            If InRange Then
              If POAmt# <> APLedgerRec(1).Amt Then
                Print #PRNFile, "PO  "; Left$(APLedgerRec(1).DOCNum, 20); Tab(54); Using(fmt$, Str$(POAmt#)); "  Open Partial"
              Else
                Print #PRNFile, "PO  "; Left$(APLedgerRec(1).DOCNum, 20); Tab(54); Using(fmt$, Str$(APLedgerRec(1).Amt)); "  Open"
              End If
            End If
          ElseIf APLedgerRec(1).TRCode = -4 Then                'for cleared P
            PODebit# = Round(PODebit# + APLedgerRec(1).Amt)
            If InRange Then
              Print #PRNFile, "PO  "; Left$(APLedgerRec(1).DOCNum, 20); Tab(54); Using(fmt$, Str$(APLedgerRec(1).Amt)); "  Closed"

            End If
          ElseIf APLedgerRec(1).TRCode = 3 Then
            VTDebit# = Round(VTDebit# + APLedgerRec(1).Amt)
            If InRange Then
              Print #PRNFile, "Chk "; Left$(APLedgerRec(1).DOCNum, 20); Tab(42); Using(fmt$, Str$(APLedgerRec(1).Amt))

            End If
          ElseIf APLedgerRec(1).TRCode = -3 Then
            If InRange Then
              Print #PRNFile, "Chk "; Left$(APLedgerRec(1).DOCNum, 10); " VOID"; Tab(42); Using(fmt$, Str$(APLedgerRec(1).Amt))

            End If
          Else
            If InRange Then
              Print #PRNFile, "Code is:"; APLedgerRec(1).TRCode
            End If
                    End If
          If InRange Then
            Linecnt = Linecnt + 1
            'PRINT #PrnFile, "PSL: !"; APLedgerRec(1).PSLFlag; "!"
          End If

          'PRINT #PrnFile, "1099? "; APLedgerRec(1).Get1099
          'END IF
          NextTrans& = APLedgerRec(1).NextTrans

        Loop

        'dale
'        If TransCnt = 0 Or APLedgerRec(1).TRDATE <= BDate Then
'          Print #20, "Vendor: "; Vendor.VNAME; Tab(40);
'          If TransCnt = 0 Then
'            Print #20, "  NO TRANSACTIONS."
'            'NoHist = NoHist + 1
'          Else
'            Print #20, "last: "; Format(DateAdd("d", (APLedgerRec(1).TRDATE), "12-31-1979"), "mm/dd/yyyy")
'          End If
'
'          DLCnt = DLCnt + 1
'          If DLCnt > 58 Then
'            Print #20, Chr$(12)
'            DLCnt = 0
'          End If
'        End If
'        'dale

        If TransCnt = 0 Then
          Print #PRNFile, "  NO TRANSACTIONS."
          Print #PRNFile, Dash$
          Linecnt = Linecnt + 2
        Else
          VTBal# = Round(VTCredit# - VTDebit#)
          Print #PRNFile, "Current Balance :"; Using(fmt$, Str$(VTBal#));
          Print #PRNFile, Tab(40); "Range Invoice Total :"; Using(fmt$, Str$(TInvAmt#))
          Print #PRNFile, "Current On Order:"; Using(fmt$, Str$(POCredit#))

          Print #PRNFile, Dash$
          Linecnt = Linecnt + 2
        End If
      End If
      End If
    End If
  Next

  Print #PRNFile, FF$

  Print #PRNFile, Tab(40 - (Int(Len(User$) / 2))); User$
  Print #PRNFile, Tab(40 - (Int(Len(Header$) / 2))); Header$
  Print #PRNFile, "Report Parameters: "
  Print #PRNFile, "=============================================================================="

  Print #PRNFile, "From Vendor: "; Vend1St$; "   Thru Vendor: "; VendLst$
  If FromDate = -32768 Then
    FromDate = 0
  End If

  Print #PRNFile, "  From Date: "; fpDate1.Text; "     Thru Date: "; fpDate2.Text
  Print #PRNFile, FF$
  If NumVRecs < 1 Then
    FrmShowPctComp.ShowPctComp 1, 1
    Print #PRNFile, "No Information To Print"
  End If
  Close
  Erase VIndex

  ViewPrint ReportFile$, "Vendor History Report"
  KillFile ReportFile$
  Me.cmdExit.Enabled = True
  Me.cmdGo.Enabled = True
  EnableCloseButton Me.hwnd, True
  mnuOptions.Enabled = True
  'If ListCnt = 0 Then Exit Sub
  'add a trap here to display an error scrn if no matching ledger recs
  'to pay

ExitSelPayables:

  Exit Sub

VendHistHeader:
  Page = Page + 1
  Print #PRNFile, Tab(40 - (Int(Len(User$) / 2))); User$
  Print #PRNFile, Tab(40 - (Int(Len(Header$) / 2))); Header$
  Print #PRNFile,
  Print #PRNFile, "Report Date: "; Date$; Tab(67); "Page #"; Page
  Print #PRNFile, "Inv Date    TrCode  Desc                        Debit      Credit  Status"
  Print #PRNFile, "=============================================================================="
  Linecnt = 6
Return
CancelExit:
  Exit Sub

End Sub

Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub
