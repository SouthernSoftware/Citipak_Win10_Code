VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Begin VB.Form frmPrnBudHist 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Budget History"
   ClientHeight    =   8640
   ClientLeft      =   36
   ClientTop       =   540
   ClientWidth     =   12192
   Icon            =   "frmPrnBudHist.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   12192
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboRptType 
      Height          =   384
      Left            =   5136
      TabIndex        =   4
      Top             =   5712
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
      ColDesigner     =   "frmPrnBudHist.frx":08CA
   End
   Begin LpLib.fpCombo fpcboAcct1 
      Height          =   384
      Left            =   5136
      TabIndex        =   0
      Top             =   3216
      Width           =   4404
      _Version        =   196608
      _ExtentX        =   7768
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
      ColDesigner     =   "frmPrnBudHist.frx":0CA0
   End
   Begin LpLib.fpCombo fpcboAcct2 
      Height          =   384
      Left            =   5136
      TabIndex        =   1
      Top             =   3852
      Width           =   4404
      _Version        =   196608
      _ExtentX        =   7768
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
      ColDesigner     =   "frmPrnBudHist.frx":1113
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00D0D0D0&
      Caption         =   "F10 &Print"
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
      TabIndex        =   5
      Top             =   7488
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
      TabIndex        =   6
      Top             =   7488
      Width           =   1332
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   7
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
            TextSave        =   "5:29 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7133
            TextSave        =   "1/26/2006"
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
   Begin EditLib.fpDateTime txtDate1 
      Height          =   372
      Left            =   5136
      TabIndex        =   2
      Top             =   4476
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
      ButtonColor     =   14737632
      AutoMenu        =   0   'False
      StartMonth      =   4
      ButtonAlign     =   0
      BoundDataType   =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpDateTime txtDate2 
      Height          =   372
      Left            =   5136
      TabIndex        =   3
      Top             =   5100
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
      ButtonColor     =   14737632
      AutoMenu        =   0   'False
      StartMonth      =   4
      ButtonAlign     =   0
      BoundDataType   =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin VB.Label Label7 
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
      Top             =   5736
      Width           =   2388
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Starting Account:"
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
      TabIndex        =   12
      Top             =   3312
      Width           =   2004
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   3588
      Left            =   2304
      Top             =   2880
      Width           =   7500
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Print Budget History Report"
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
      TabIndex        =   11
      Top             =   1560
      Width           =   4332
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3216
      Top             =   1320
      Width           =   5772
   End
   Begin VB.Image Image1 
      Height          =   276
      Left            =   2496
      Picture         =   "frmPrnBudHist.frx":1586
      Top             =   3120
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
      TabIndex        =   10
      Top             =   5148
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
      TabIndex        =   9
      Top             =   4536
      Width           =   1668
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ending Account:"
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
      TabIndex        =   8
      Top             =   3924
      Width           =   1812
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D0D0D0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00D0D0D0&
      Height          =   972
      Left            =   3216
      Top             =   1200
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
Attribute VB_Name = "frmPrnBudHist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim GLSetup As GLSetupRecType
Dim GLAcct    As GLAcctRecType
Dim GLFundIdx As GLFundIndexType
Dim GLAcctidx As GLAcctIndexType
'Dim GLTrans   As GLTransRecType
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Dim FY1BegDate As Integer, FY1EndDate As Integer, FY2BegDate As Integer, FY2EndDate As Integer
Dim FirstFund As String, LastFund As String
Dim ActiveYear As Integer

Private Sub cmdExit_Click()
  frmGLReportsMenu.Show
  Unload frmPrnBudHist
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        ClearInUse PWcnt
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
      cmdExit_Click
      KeyCode = 0
    Case vbKeyF10:
      cmdPrint_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub
Private Sub fpcboAcct1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboAcct1.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcboAcct1.ListIndex = -1
    fpcboAcct1.Action = ActionClearSearchBuffer
  End If
  If fpcboAcct1.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpcboAcct2.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        cmdPrint.SetFocus
        KeyCode = 0
      End If
    End If
  End If

End Sub
Private Sub fpcboAcct2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboAcct2.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcboAcct2.ListIndex = -1
    fpcboAcct2.Action = ActionClearSearchBuffer
  End If
  If fpcboAcct2.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      txtDate1.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpcboAcct1.SetFocus
        KeyCode = 0
      End If
    End If
  End If

End Sub
Private Sub fpcboRptType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboRptType.ListDown = True
  End If
  If fpcboRptType.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      cmdPrint.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        txtDate2.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub
Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  StatusBar1.Panels.Item(1).Text = GLUserName
  Me.HelpContextID = hlpBudgetHistory
  BudAcctstwo fpcboAcct1, fpcboAcct2
  txtDate1.Text = Format(Now, "mm/dd/yyyy")
  txtDate2.Text = txtDate1.Text
  fpcboRptType.InsertRow = "Graphics"
  fpcboRptType.InsertRow = "Text"
  fpcboRptType.ListIndex = 0
End Sub

Private Sub Form_Resize()
'  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
'  End If
End Sub
Private Function ValidDate()
  Dim TempDate1 As Integer, TempDate2 As Integer
  GetFYDates FY1BegDate, FY1EndDate, FY2BegDate, FY2EndDate
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
Private Function ValidAccts()
  If fpcboAcct1.ListIndex <> -1 And fpcboAcct2.ListIndex <> -1 Then
    fpcboAcct1.Col = 1
    fpcboAcct2.Col = 1
    If fpcboAcct1.ColText > fpcboAcct2.ColText Then
      MsgBox "Invalid Account Selection, The Starting Account Should Be Less or Equal to Ending Account.", vbOKOnly, "Invalid Selection"
      ValidAccts = False
    Else
      ValidAccts = True
    End If
  Else
    MsgBox "You Must Select An Account, Retry", vbOKOnly, "Invalid Selection"
    ValidAccts = False
  End If
End Function

Private Sub cmdPrint_Click()
  If ValidAccts = True Then
    If ValidDate = True Then
      If fpcboRptType.ListIndex = 0 Then
        rptopt = 1
      ElseIf fpcboRptType.ListIndex = 1 Then
        rptopt = 2
      End If
      If rptopt = 1 Then
        PrintBgtHist
      ElseIf rptopt = 2 Then
        PrintBgtHist2
      End If
    End If
  End If
End Sub

Private Sub PrintBgtHist()
  Dim MaxLines As Integer, LookFor As String, CrLF As String, T As String
  Dim Linecnt As Integer, PRNFile As Integer, FundCnt As Integer
  Dim ReportFile As String, ToPrint As String, SumLine As String
  Dim FF As String, Header As String, FirstAcct As String, LastAcct As String
  Dim PRNFileNum As Integer, cnt As Integer, Howmany As Integer
  Dim FundCode As String, DivLine As String, DivLine2 As String
  Dim CommaFmt As String, TotalFmt As String, FundNumber As String
  Dim TotDr As Double, TotCR As Double, TranCashTot As Double, CalcBal As Double
  ReDim FundList(1) As String
  Dim OpenDate As String, IGuess As Integer, GrTotDr As Double, GrTotCr As Double
  Dim FundDr As Double, FundCr As Double, FundRecNum As Integer, NewAcct As Boolean
  Dim Found As Boolean, FundOutofBal As Boolean, Fund As Integer
  Dim FundIdxFileNum As Integer, NumFunds As Integer, EndDate As Integer, BegDate As Integer
  Dim AcctIdxFileNum As Integer, NumGLAccts As Integer, FundName As String
  Dim AcctFileNum As Integer, NumGLAcctRecs As Integer, RecNo As Integer
  Dim transfile As Integer, NumTrans As Long, NextTr As Long, AcctNum As String
  Dim Debit As String, Credit As String, Diff As Double, PYFundBal As Double
  Dim DrFwd As Double, CrFwd As Double, TotAcctDr As Double, TotAcctCr As Double
  Dim BalFwd As Double, AcctNumber As String, FYBeg As Integer, FwdFlag As Boolean
  Dim OutOfOrder As Boolean, cntT As Integer, AcctRunBal As Double, Trn As Integer
  Dim TmpSort As TrSortType
  Dim Opt As String, AcctBal As Double, BgtCol As Integer, BudgetAmt As Double
  Dim Var As Double, VarCol As Integer, VarText As String, HollyFlag As Boolean
  Dim Pitch12 As String, PageNum As Integer, NumAcctTrans As Long, RunBalFmt As String
  Dim BgtTransFile As Integer, NumBgtTrans As Integer, Newrp As String
  Dim ToPrintA As String, ToPrintB As String, ToPrintC As String
  ReDim Desc$(1)
  GetFYDates FY1BegDate, FY1EndDate, FY2BegDate, FY2EndDate

'''''Remember QPTrim$

  'End of Input
  '=====================================================
  'Start Report Processing

  CommaFmt$ = "###,###,###.##"  'format takes 13 chars
  TotalFmt$ = "#,###,###,###.##" 'format takes 14 chars
  RunBalFmt$ = "##########.##"
  SumLine$ = String$(16, "-")   'column summary line
  DivLine$ = String$(77, "-")   'dashed line
  DivLine2$ = String$(77, "=")  'Double Line
  CrLF$ = Chr$(13) + Chr$(10)
  BegDate = DateDiff("d", "12/31/1979", txtDate1)
  EndDate = DateDiff("d", "12/31/1979", txtDate2)
  Header$ = "Budget History"
  Desc$(1) = "Date       Description             Reference       Debit"
  fpcboAcct1.Col = 1
  fpcboAcct2.Col = 1
  FirstAcct$ = QPTrim$(fpcboAcct1.ColText)
  LastAcct$ = QPTrim(fpcboAcct2.ColText)
  TotDr# = 0
  TotCR# = 0
  OpenDate$ = Format(DateAdd("d", FY1BegDate, "12-31-1979"), "mm/dd/yy")
  'OpenDesc$ = "Balance Foward"
  Newrp = "BGTHST"
  GetRPTName Newrp
  ReportFile$ = Newrp
  PRNFile = FreeFile
  Open ReportFile$ For Output As #PRNFile
  OpenAcctIdx AcctIdxFileNum, NumGLAccts
  OpenAcctFile AcctFileNum
  NumGLAcctRecs = LOF(AcctFileNum) \ Len(GLAcct)

  Dim BgtTrans As GLTransRecType
  OpenBgtTransFile BgtTransFile, NumTrans
  If NumTrans = 0 Then
    Close
    MsgBox "No Transactions To Report.", vbOKOnly, "No Trans"
    Exit Sub
  End If
  FrmShowPctComp.Label1 = "Printing Budget History Report"
  FrmShowPctComp.Show , Me
  DoEvents
  DeActivateControls frmPrnBudHist, True
  'trying to set TrSort array to max needed
 
  ReDim Trsort(1 To 1) As TrSortType

  GrTotDr# = 0
  GrTotCr# = 0

  For cnt = 1 To NumGLAccts

    DrFwd# = 0
    CrFwd# = 0
    TotAcctDr# = 0
    TotAcctCr# = 0

    Get AcctIdxFileNum, cnt, GLAcctidx
    Get AcctFileNum, GLAcctidx.RecNum, GLAcct

    AcctNum$ = QPTrim$(GLAcct.Num)
    If AcctNum$ >= FirstAcct$ And AcctNum$ <= LastAcct$ Then
      If GLAcct.Typ = "R" Or GLAcct.Typ = "E" Then

        NextTr = GLAcct.FrstBTran 'get the first trans for this acct
        'AcctNumber$ = "Account " + AcctNum$ + " - " + QPTrim$(GLAcct.Title)
        AcctNumber$ = AcctNum$ + " - " + QPTrim$(GLAcct.Title)

        ToPrintA$ = Space$(80)
        ToPrintA$ = AcctNumber$
        'Print #PRNFile, ToPrint$

        Do Until NextTr = 0     'keep going 'til we run out of trans

          Get BgtTransFile, NextTr, BgtTrans

          If BgtTrans.TRDATE >= BegDate And BgtTrans.TRDATE <= EndDate Then
            '--within range - assign to array for sorting
            NumBgtTrans = NumBgtTrans + 1
            ReDim Preserve Trsort(1 To NumBgtTrans) As TrSortType
            Trsort(NumBgtTrans).TRDATE = BgtTrans.TRDATE
            Trsort(NumBgtTrans).Record = NextTr

          Else
            '--check the transaction to see if we need to carry it in
            '  the balance fwd
            'QPRintRC Num2Date(BgtTrans.TrDate), 24, 25, -1

            If BgtTrans.TRDATE < BegDate Then
              DrFwd# = DrFwd# + BgtTrans.DrAmt
              CrFwd# = CrFwd# + BgtTrans.CrAmt
              FwdFlag = -1
            End If
          End If

          NextTr = BgtTrans.NextTran            'Get the next transaction

        Loop
        ToPrintB$ = Space$(80)
        If FwdFlag Then
          '--
          ToPrintB$ = "Balance Forward"
          Select Case GLAcct.Typ
          Case "E"
            BalFwd# = DrFwd# - CrFwd#

            If BalFwd# >= 0 Then
              Debit$ = Using$(CommaFmt$, Str$(BalFwd#))
              Credit$ = "0"
            Else
              Credit$ = Using$(CommaFmt$, Str$(Abs(BalFwd#)))
              Debit$ = "0"
            End If

          Case "R"
            BalFwd# = CrFwd# - DrFwd#
            If BalFwd# >= 0 Then
              Credit$ = Using$(CommaFmt$, Str$(BalFwd#))
              Debit$ = "0"
            Else
              Debit$ = Using$(CommaFmt$, Str$(Abs(BalFwd#)))
              Credit$ = "0"
            End If

          End Select
          ToPrintB$ = ToPrintB$ + "~~" + Debit$ + "~" + Credit$
          Print #PRNFile, ToPrintA$ + "~~~" + ToPrintB$ + "~~~"
        Else
          ToPrintB$ = ""
          Debit$ = 0
          Credit$ = 0
        End If
        If NumBgtTrans > 0 Or BalFwd# <> 0 Then


         ' SortT TrSort(1), NumAcctTrans, 0, 6, 0, -1
  'Created a SortT Function in Main Module
          SortT Trsort(), NumBgtTrans
'''        Do
'''          OutOfOrder = False          'assume it's sorted
'''          For cntT = 1 To NumBgtTrans - 1
'''            If Trsort(cntT).TRDATE > Trsort(cntT + 1).TRDATE Then
'''              LSet TmpSort = Trsort(cntT)
'''              LSet Trsort(cntT) = Trsort(cntT + 1)
'''              LSet Trsort(cntT + 1) = TmpSort
'''              OutOfOrder = True       'we're not done yet
'''            End If
'''          Next
'''        Loop While OutOfOrder

     
          For Trn = 1 To NumBgtTrans
            Get BgtTransFile, Trsort(Trn).Record, BgtTrans

            ToPrintC$ = Space$(80)

            ToPrintC$ = "D~" + Format(DateAdd("d", BgtTrans.TRDATE, "12-31-1979"), "mm/dd/yy")
            ToPrintC$ = ToPrintC$ + "~" + BgtTrans.Desc
            ToPrintC$ = ToPrintC$ + "~" + BgtTrans.Ref

            If BgtTrans.DrAmt <> 0 Then
              ToPrintC$ = ToPrintC$ + "~" + Using$(CommaFmt$, Str$(BgtTrans.DrAmt))
            Else
              ToPrintC$ = ToPrintC$ + "~" + "0"
            End If
            If BgtTrans.CrAmt <> 0 Then
              ToPrintC$ = ToPrintC$ + "~" + Using$(CommaFmt$, Str$(BgtTrans.CrAmt))
            Else
              ToPrintC$ = ToPrintC$ + "~" + "0"
            End If
            ToPrintC$ = ToPrintC$ + "~" + QPTrim(BgtTrans.Src) + "~~"

            Print #PRNFile, ToPrintA$ + "~" + ToPrintC$

            TotAcctDr# = TotAcctDr# + BgtTrans.DrAmt
            TotAcctCr# = TotAcctCr# + BgtTrans.CrAmt

            GrTotDr# = GrTotDr# + BgtTrans.DrAmt
            GrTotCr# = GrTotCr# + BgtTrans.CrAmt

          Next

          '--Print summary lines
          'ToPrint$ = Space$(80)
          'Mid$(ToPrint$, 43) = SumLine$
          'Mid$(ToPrint$, 57) = SumLine$

          'Print #PRNFile, ToPrint$

          '--Print transaction totals
          If NumAcctTrans > 0 Then
'            ToPrint$ = Space$(80)
'            Mid$(ToPrint$, 1) = "Transaction Totals"
'            Mid$(ToPrint$, 43) = Using$(TotalFmt$, Str$(TotAcctDr#))
'            Mid$(ToPrint$, 57) = Using$(TotalFmt$, Str$(TotAcctCr#))
'            Print #PRNFile, ToPrint$
          End If

          '--Print ending balance
          ToPrint$ = Space$(80)
          
          Select Case GLAcct.Typ
          Case "E"
            AcctBal# = BalFwd# + TotAcctDr# - TotAcctCr#
            If AcctBal# >= 0 Then
              Debit$ = Using$(TotalFmt$, Str$(AcctBal#))
              Credit$ = "0"
            Else
              Credit$ = Using$(TotalFmt$, Str$(Abs(AcctBal#)))
              Debit$ = "0"
            End If

          Case "R"
            AcctBal# = BalFwd# + TotAcctCr# - TotAcctDr#
            If AcctBal# >= 0 Then
              Credit$ = Using$(TotalFmt$, Str$(AcctBal#))
              Debit$ = "0"
            Else
              Debit$ = Using$(TotalFmt$, Str$(Abs(AcctBal#)))
              Credit$ = "0"
            End If

          End Select
          'If AcctBal# <> 0 Then
            ToPrint$ = "Budget Balance"
          'Else
          '  ToPrint$ = ""
          'End If
          ToPrint$ = ToPrint$ + "~~~~~" + Debit$ + "~" + Credit$
          
          Print #PRNFile, ToPrintA$ + "~~~" + ToPrint$

          'ToPrint$ = String$(80, "*")
          'Print #PRNFile, ToPrint$

        Else
          ToPrint$ = Space$(80)
          ToPrint$ = ToPrintA$ + "~~~ -- No Activity -- ~~~~~~"
          Print #PRNFile, ToPrint$

          'ToPrint$ = String$(80, "*")
          'Print #PRNFile, ToPrint$

        End If

      End If
    End If      'Account is not of this fund
       FrmShowPctComp.ShowPctComp cnt, NumGLAccts
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ActivateControls frmPrnBudHist, True
      Unload FrmShowPctComp
      GoTo CancelExit
    End If
    NumBgtTrans = 0             'reset for next account
    BalFwd# = 0
    FwdFlag = 0
    TotAcctDr# = 0
    TotAcctCr# = 0

  Next
'  ToPrint$ = Space$(80)
'  LSet ToPrint$ = "Grand Total Debits"
'  Mid$(ToPrint$, 25) = Using$(TotalFmt$, Str$(GrTotDr#))
'  Print #PRNFile, ToPrint$
'
'  ToPrint$ = Space$(80)
'  LSet ToPrint$ = "Grand Total Credits"
'  Mid$(ToPrint$, 25) = Using$(TotalFmt$, Str$(GrTotCr#))
'  Print #PRNFile, ToPrint$

  Close
'  ViewPrint ReportFile$, "Budget History Report"
'  KillFile ReportFile$
  ActivateControls frmPrnBudHist, True
  Load frmLoadingRpt
  ARptBudHist.txtRptInfo = "Reporting - Account: " + FirstAcct$ + " thru " + LastAcct$ + " For Date Range: " + txtDate1 + " thru " + txtDate2
  ARptBudHist.txtDate = Now
  ARptBudHist.txtTown = GLUserName$
  ARptBudHist.GetName ReportFile$
  ARptBudHist.startrpt

CancelExit:
Exit Sub
End Sub


Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
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
Private Sub PrintBgtHist2()
  Dim MaxLines As Integer, LookFor As String, CrLF As String, T As String
  Dim Linecnt As Integer, PRNFile As Integer, FundCnt As Integer
  Dim ReportFile As String, ToPrint As String, SumLine As String
  Dim FF As String, Header As String, FirstAcct As String, LastAcct As String
  Dim PRNFileNum As Integer, cnt As Integer, Howmany As Integer
  Dim FundCode As String, DivLine As String, DivLine2 As String
  Dim CommaFmt As String, TotalFmt As String, FundNumber As String
  Dim TotDr As Double, TotCR As Double, TranCashTot As Double, CalcBal As Double
  ReDim FundList(1) As String
  Dim OpenDate As String, IGuess As Integer, GrTotDr As Double, GrTotCr As Double
  Dim FundDr As Double, FundCr As Double, FundRecNum As Integer, NewAcct As Boolean
  Dim Found As Boolean, FundOutofBal As Boolean, Fund As Integer
  Dim FundIdxFileNum As Integer, NumFunds As Integer, EndDate As Integer, BegDate As Integer
  Dim AcctIdxFileNum As Integer, NumGLAccts As Integer, FundName As String
  Dim AcctFileNum As Integer, NumGLAcctRecs As Integer, RecNo As Integer
  Dim transfile As Integer, NumTrans As Long, NextTr As Long, AcctNum As String
  Dim Debit As String, Credit As String, Diff As Double, PYFundBal As Double
  Dim DrFwd As Double, CrFwd As Double, TotAcctDr As Double, TotAcctCr As Double
  Dim BalFwd As Double, AcctNumber As String, FYBeg As Integer, FwdFlag As Boolean
  Dim OutOfOrder As Boolean, cntT As Integer, AcctRunBal As Double, Trn As Integer
  Dim TmpSort As TrSortType
  Dim Opt As String, AcctBal As Double, BgtCol As Integer, BudgetAmt As Double
  Dim Var As Double, VarCol As Integer, VarText As String, HollyFlag As Boolean
  Dim Pitch12 As String, PageNum As Integer, NumAcctTrans As Long, RunBalFmt As String
  Dim BgtTransFile As Integer, NumBgtTrans As Integer, Newrp As String
  ReDim Desc$(1)
  GetFYDates FY1BegDate, FY1EndDate, FY2BegDate, FY2EndDate

'''''Remember QPTrim$

  'End of Input
  '=====================================================
  'Start Report Processing

  CommaFmt$ = "###,###,###.##"  'format takes 13 chars
  TotalFmt$ = "#,###,###,###.##" 'format takes 14 chars
  RunBalFmt$ = "##########.##"
  SumLine$ = String$(16, "-")   'column summary line
  DivLine$ = String$(77, "-")   'dashed line
  DivLine2$ = String$(77, "=")  'Double Line
  CrLF$ = Chr$(13) + Chr$(10)
  BegDate = DateDiff("d", "12/31/1979", txtDate1)
  EndDate = DateDiff("d", "12/31/1979", txtDate2)
  Header$ = "Budget History"
  Desc$(1) = "Date       Description             Reference       Debit"
  fpcboAcct1.Col = 1
  fpcboAcct2.Col = 1
  FirstAcct$ = QPTrim$(fpcboAcct1.ColText)
  LastAcct$ = QPTrim(fpcboAcct2.ColText)
  TotDr# = 0
  TotCR# = 0
  OpenDate$ = Format(DateAdd("d", FY1BegDate, "12-31-1979"), "mm/dd/yy")
  'OpenDesc$ = "Balance Foward"
  Newrp = "BGTHST"
  GetRPTName Newrp
  ReportFile$ = Newrp
  PRNFile = FreeFile
  Open ReportFile$ For Output As #PRNFile
  OpenAcctIdx AcctIdxFileNum, NumGLAccts
  OpenAcctFile AcctFileNum
  NumGLAcctRecs = LOF(AcctFileNum) \ Len(GLAcct)

  Dim BgtTrans As GLTransRecType
  OpenBgtTransFile BgtTransFile, NumTrans
  If NumTrans = 0 Then
    Close
    MsgBox "No Transactions To Report.", vbOKOnly, "No Trans"
    Exit Sub
  End If
  FrmShowPctComp.Label1 = "Printing Budget History Report"
  FrmShowPctComp.Show , Me
  DoEvents
  DeActivateControls frmPrnBudHist, True
  'trying to set TrSort array to max needed
 
  ReDim Trsort(1 To 1) As TrSortType

  GrTotDr# = 0
  GrTotCr# = 0

  For cnt = 1 To NumGLAccts

    DrFwd# = 0
    CrFwd# = 0
    TotAcctDr# = 0
    TotAcctCr# = 0

    Get AcctIdxFileNum, cnt, GLAcctidx
    Get AcctFileNum, GLAcctidx.RecNum, GLAcct

    AcctNum$ = QPTrim$(GLAcct.Num)
    If AcctNum$ >= FirstAcct$ And AcctNum$ <= LastAcct$ Then
      If GLAcct.Typ = "R" Or GLAcct.Typ = "E" Then

        NextTr = GLAcct.FrstBTran 'get the first trans for this acct
        AcctNumber$ = "Account " + AcctNum$ + " - " + QPTrim$(GLAcct.Title)

        ToPrint$ = Space$(80)
        LSet ToPrint$ = AcctNumber$
        Print #PRNFile, ToPrint$

        Do Until NextTr = 0     'keep going 'til we run out of trans

          Get BgtTransFile, NextTr, BgtTrans

          If BgtTrans.TRDATE >= BegDate And BgtTrans.TRDATE <= EndDate Then
            '--within range - assign to array for sorting
            NumBgtTrans = NumBgtTrans + 1
            ReDim Preserve Trsort(1 To NumBgtTrans) As TrSortType
            Trsort(NumBgtTrans).TRDATE = BgtTrans.TRDATE
            Trsort(NumBgtTrans).Record = NextTr

          Else
            '--check the transaction to see if we need to carry it in
            '  the balance fwd
            'QPRintRC Num2Date(BgtTrans.TrDate), 24, 25, -1

            If BgtTrans.TRDATE < BegDate Then
              DrFwd# = DrFwd# + BgtTrans.DrAmt
              CrFwd# = CrFwd# + BgtTrans.CrAmt
              FwdFlag = -1
            End If
          End If

          NextTr = BgtTrans.NextTran            'Get the next transaction

        Loop

        If FwdFlag Then
          '--
          ToPrint$ = Space$(80)
          Mid$(ToPrint$, 1) = "Balance Forward"

          Select Case GLAcct.Typ
          Case "E"
            BalFwd# = DrFwd# - CrFwd#

            If BalFwd# >= 0 Then
              Debit$ = Using$(CommaFmt$, Str$(BalFwd#))
              Credit$ = ""
            Else
              Credit$ = Using$(CommaFmt$, Str$(Abs(BalFwd#)))
              Debit$ = ""
            End If

          Case "R"
            BalFwd# = CrFwd# - DrFwd#
            If BalFwd# >= 0 Then
              Credit$ = Using$(CommaFmt$, Str$(BalFwd#))
              Debit$ = ""
            Else
              Debit$ = Using$(CommaFmt$, Str$(Abs(BalFwd#)))
              Credit$ = ""
            End If

          End Select

          Mid$(ToPrint$, 45) = Debit$
          Mid$(ToPrint$, 59) = Credit$
          Print #PRNFile, ToPrint$

        End If

        If NumBgtTrans > 0 Or BalFwd# <> 0 Then


         ' SortT TrSort(1), NumAcctTrans, 0, 6, 0, -1
  'Created a SortT Function in Main Module
          SortT Trsort(), NumBgtTrans
'''        Do
'''          OutOfOrder = False          'assume it's sorted
'''          For cntT = 1 To NumBgtTrans - 1
'''            If Trsort(cntT).TRDATE > Trsort(cntT + 1).TRDATE Then
'''              LSet TmpSort = Trsort(cntT)
'''              LSet Trsort(cntT) = Trsort(cntT + 1)
'''              LSet Trsort(cntT + 1) = TmpSort
'''              OutOfOrder = True       'we're not done yet
'''            End If
'''          Next
'''        Loop While OutOfOrder

     
          For Trn = 1 To NumBgtTrans
            Get BgtTransFile, Trsort(Trn).Record, BgtTrans

            ToPrint$ = Space$(80)

            Mid$(ToPrint$, 1) = Format(DateAdd("d", BgtTrans.TRDATE, "12-31-1979"), "mm/dd/yy")
            Mid$(ToPrint$, 12) = BgtTrans.Desc
            Mid$(ToPrint$, 36) = BgtTrans.Ref

            If BgtTrans.DrAmt <> 0 Then
              Mid$(ToPrint$, 45) = Using$(CommaFmt$, Str$(BgtTrans.DrAmt))

            End If

            If BgtTrans.CrAmt <> 0 Then
              Mid$(ToPrint$, 59) = Using$(CommaFmt$, Str$(BgtTrans.CrAmt))
            End If

            Mid$(ToPrint$, 75) = Left$(BgtTrans.Src, 6)

            Print #PRNFile, ToPrint$; Trsort(Trn).Record

            TotAcctDr# = TotAcctDr# + BgtTrans.DrAmt
            TotAcctCr# = TotAcctCr# + BgtTrans.CrAmt

            GrTotDr# = GrTotDr# + BgtTrans.DrAmt
            GrTotCr# = GrTotCr# + BgtTrans.CrAmt

          Next

          '--Print summary lines
          ToPrint$ = Space$(80)
          Mid$(ToPrint$, 43) = SumLine$
          Mid$(ToPrint$, 57) = SumLine$

          Print #PRNFile, ToPrint$

          '--Print transaction totals
          If NumAcctTrans > 0 Then
            ToPrint$ = Space$(80)
            Mid$(ToPrint$, 1) = "Transaction Totals"
            Mid$(ToPrint$, 43) = Using$(TotalFmt$, Str$(TotAcctDr#))
            Mid$(ToPrint$, 57) = Using$(TotalFmt$, Str$(TotAcctCr#))
            Print #PRNFile, ToPrint$
          End If

          '--Print ending balance
          ToPrint$ = Space$(80)
          Mid$(ToPrint$, 1) = "Budget Balance"
          Select Case GLAcct.Typ
          Case "E"
            AcctBal# = BalFwd# + TotAcctDr# - TotAcctCr#
            If AcctBal# >= 0 Then
              Debit$ = Using$(TotalFmt$, Str$(AcctBal#))
              Credit$ = ""
            Else
              Credit$ = Using$(TotalFmt$, Str$(Abs(AcctBal#)))
              Debit$ = ""
            End If

          Case "R"
            AcctBal# = BalFwd# + TotAcctCr# - TotAcctDr#
            If AcctBal# >= 0 Then
              Credit$ = Using$(TotalFmt$, Str$(AcctBal#))
              Debit$ = ""
            Else
              Debit$ = Using$(TotalFmt$, Str$(Abs(AcctBal#)))
              Credit$ = ""
            End If

          End Select

          Mid$(ToPrint$, 43) = Debit$
          Mid$(ToPrint$, 57) = Credit$
          Print #PRNFile, ToPrint$

          ToPrint$ = String$(80, "*")
          Print #PRNFile, ToPrint$

        Else
          ToPrint$ = Space$(80)
          Mid$(ToPrint$, 5) = " -- No Activity --"
          Print #PRNFile, ToPrint$

          ToPrint$ = String$(80, "*")
          Print #PRNFile, ToPrint$

        End If

      End If
    End If      'Account is not of this fund
       FrmShowPctComp.ShowPctComp cnt, NumGLAccts
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ActivateControls frmPrnBudHist, True
      Unload FrmShowPctComp
      GoTo CancelExit
    End If
    NumBgtTrans = 0             'reset for next account
    BalFwd# = 0
    FwdFlag = 0
    TotAcctDr# = 0
    TotAcctCr# = 0

  Next
  ToPrint$ = Space$(80)
  LSet ToPrint$ = "Grand Total Debits"
  Mid$(ToPrint$, 25) = Using$(TotalFmt$, Str$(GrTotDr#))
  Print #PRNFile, ToPrint$

  ToPrint$ = Space$(80)
  LSet ToPrint$ = "Grand Total Credits"
  Mid$(ToPrint$, 25) = Using$(TotalFmt$, Str$(GrTotCr#))
  Print #PRNFile, ToPrint$

  Close
  ViewPrint ReportFile$, "Budget History Report"
  KillFile ReportFile$
  ActivateControls frmPrnBudHist, True

CancelExit:
Exit Sub
End Sub

