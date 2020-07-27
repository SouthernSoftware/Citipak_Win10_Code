VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmRptCMJournal 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Cash Management Journal"
   ClientHeight    =   8865
   ClientLeft      =   3930
   ClientTop       =   1890
   ClientWidth     =   12210
   Icon            =   "frmRptCMJournal.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   12210
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboSource 
      Height          =   375
      Left            =   5370
      TabIndex        =   2
      Top             =   3690
      Width           =   3540
      _Version        =   196608
      _ExtentX        =   6244
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
      ColDesigner     =   "frmRptCMJournal.frx":08CA
   End
   Begin LpLib.fpCombo fpcboPrintOrder 
      Height          =   375
      Left            =   5370
      TabIndex        =   4
      Top             =   4770
      Width           =   3615
      _Version        =   196608
      _ExtentX        =   6376
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
      ColDesigner     =   "frmRptCMJournal.frx":0BED
   End
   Begin LpLib.fpCombo fpcboRptType 
      Height          =   375
      Left            =   5370
      TabIndex        =   5
      Top             =   5310
      Width           =   3615
      _Version        =   196608
      _ExtentX        =   6376
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
      ColDesigner     =   "frmRptCMJournal.frx":0F10
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Include Description"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Left            =   5424
      TabIndex        =   17
      Top             =   6408
      Width           =   2340
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Show Detail"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Left            =   5448
      TabIndex        =   16
      Top             =   5856
      Width           =   2316
   End
   Begin VB.CommandButton cmdPrint 
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
      Left            =   8400
      TabIndex        =   6
      Top             =   7824
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
      Left            =   10080
      TabIndex        =   7
      Top             =   7824
      Width           =   1332
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   8
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
            TextSave        =   "7:57 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7144
            TextSave        =   "2/4/2020"
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
   Begin EditLib.fpDateTime txtDate2 
      Height          =   348
      Left            =   5376
      TabIndex        =   1
      Top             =   3156
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
      Left            =   5376
      TabIndex        =   0
      Top             =   2640
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
   Begin EditLib.fpText fptxtOperator 
      Height          =   348
      Left            =   5376
      TabIndex        =   3
      Top             =   4224
      Width           =   852
      _Version        =   196608
      _ExtentX        =   1503
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
      MaxLength       =   4
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Print CM Payment Journal"
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
      Left            =   3624
      TabIndex        =   15
      Top             =   984
      Width           =   5004
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   708
      Left            =   3192
      Top             =   816
      Width           =   5772
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
      Height          =   372
      Index           =   0
      Left            =   3720
      TabIndex        =   14
      Top             =   3204
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
      Height          =   372
      Left            =   3624
      TabIndex        =   13
      Top             =   2688
      Width           =   1668
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
      Left            =   2952
      TabIndex        =   12
      Top             =   5352
      Width           =   2388
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Source:"
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
      Left            =   3048
      TabIndex        =   11
      Top             =   3756
      Width           =   2244
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   4812
      Left            =   2328
      Top             =   2304
      Width           =   7284
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Printing Order:"
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
      Index           =   7
      Left            =   3576
      TabIndex        =   10
      Top             =   4824
      Width           =   1716
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Operator No:"
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
      Index           =   2
      Left            =   3216
      TabIndex        =   9
      Top             =   4284
      Width           =   2076
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   828
      Left            =   3192
      Top             =   696
      Width           =   5772
   End
End
Attribute VB_Name = "frmRptCMJournal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider

Private Sub cmdExit_Click()
  Load frmCMReportMenu
  DoEvents
  frmCMReportMenu.Show
  Unload Me
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        CMLog "Closed via RptCMTransJournal by " + PWUser$
        CitiTerminate
      End If
    End If
  End If
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
    fpcboSource.SetFocus
  End If
End Sub
Private Sub fpcboSource_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboSource.ListDown = True
  End If
  If fpcboSource.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fptxtOperator.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        txtDate2.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub

Private Sub fptxtOperator_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpcboPrintOrder.SetFocus
  End If
End Sub


Private Sub fpcboPrintOrder_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboPrintOrder.ListDown = True
  End If
  If fpcboPrintOrder.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpcboRptType.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fptxtOperator.SetFocus
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
        fpcboPrintOrder.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub


Private Sub cmdPrint_Click()
  If ValidDate = True Then
   If fpcboSource.ListIndex <> -1 Then
    DeActivateControls Me
    If fpcboRptType.ListIndex = 1 Then
      PrintJournal
      ActivateControls Me
    ElseIf fpcboRptType.ListIndex = 0 Then
      PrintJournal
    Else
      ActivateControls Me
    End If
   Else
    MsgBox "Invalid Source Selection.", vbOKOnly, "Invalid Selection"
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

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TownName$
  fptxtOperator = OperNum
  fpcboPrintOrder.AddItem "Entry Order"
  fpcboPrintOrder.AddItem "Name"
  fpcboPrintOrder.ListIndex = 0
  txtDate1.Text = Format(Now, "mm/dd/yyyy")
  txtDate2.Text = Format(Now, "mm/dd/yyyy")
  fpcboSource.AddItem "All"
  fpcboSource.AddItem "Utility Payment"
  fpcboSource.AddItem "Miscellaneous Payment"
  fpcboSource.AddItem "Business License Payment"
  fpcboSource.AddItem "Tax Payment"
  fpcboSource.AddItem "Decal Payment"
  fpcboSource.AddItem "Voids Only"
  fpcboSource.ListIndex = 0
  fpcboRptType.InsertRow = "Graphics"
  fpcboRptType.InsertRow = "Text - Condensed Print"
  fpcboRptType.ListIndex = 0
End Sub

Private Sub Form_Resize()
'  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
'  End If
End Sub

Private Sub PrintJournal()
    Dim BegDate As Integer, EndDate As Integer, FromDate As String
    Dim ThruDate As String, RecSource As String, OperatorNumber As Integer
    Dim SortOrder As String, ReportFile As String, Fmt1 As String, Fmt2 As String
    Dim Fmt3 As String, Fmt4 As String, TotDr As Double, TotCr As Double
    Dim CMTrRecLen As Integer, TRHandle As Integer, TrNumRecs As Long
    Dim Max As Long, Size As Long, Start As Integer, sDir As Integer
    Dim SSize As Integer, MOff As Integer, MSize As Integer, RptHandle As Integer
    Dim NumOfMiscRecs As Long, cnt As Long, RptType As Integer
    Dim TRType As String, TotalCash As Double, TotalCheck As Double
    Dim TotalCharge As Double, TxRev As Double, TRev As Integer
    Dim TotalAmount As Double, Change As Double, TotChange As Double
    Dim TotalReceipts As Long, PrintMiscFlag As Integer, MCnt As Integer
    Dim MiscRevAmt As Double, NumofRevs As Integer, RCnt As Integer
    Dim PrintUtilFlag As Integer, PrintTaxFlag As Integer, Header As String
    Dim Page As Integer, BegRecNo As Long, TransDate As Integer, PrintDecalFlag As Integer
    Dim GoodRecordFlag As Boolean, CntG As Long, UBSetupLen As Integer
    Dim RevCnt As Integer, OutOfOrder As Boolean, x As Integer
    Dim Temp2 As Integer, uCnt As Integer, dcnt As Integer, TotalUtilAmt As Double
    Dim TotalDepAmt As Double, TotCurTax As Double, TotPastTax As Double
    Dim TotCurInt As Double, TotCurPen As Double, TotStrmFee As Double
    Dim TotPastInt As Double, TotPastPen As Double, TotPastStrm As Double
    Dim TotalBLAmt As Double, TotalDCAmt As Double, TotalMiscCnt As Long
    Dim TCnt As Long, TotalMisc As Double, NumChks As Long, PrnOpr As String
    Dim VOnly As Boolean, VTot As Long, Vflag As Boolean, totPay As Double
    Dim TotPrincpleTx As Double, TotIntTax As Double, TotCollTax As Double
    Dim TotLateList As Double, TotOpt1 As Double, TotOpt2 As Double, TotOpt3 As Double
    Dim TotDiscTax As Double, TotPrePayTx As Double, TotTaxBills As Double
    Dim TxOpt1 As String, TxOpt2 As String, TxOpt3 As String, TMHandle As Integer
    Dim TxOpt1F As String, TxOpt2F As String, TxOpt3F As String, TotTaxBillsR#
    Dim TxPOpt1 As String, TxPOpt2 As String, TxPOpt3 As String
    Dim TxPOpt1F As String, TxPOpt2F As String, TxPOpt3F As String
    ReDim RevName$(10), TotalMiscRec$(200), TotalMiscDesc$(200), TotalMiscAmt#(200), MiscCodeGL$(200)
    Dim TotPenTax#, TotAllTax#, TotPrinc1Tx#, TotPrinc2Tx#, TotPrinc3Tx#, TotPenTaxP#
    Dim TotPrinc4Tx#, TotPrinc5Tx#, TotTaxBillsP As Double, TotOpt2P As Double, TotOpt3P As Double
    Dim TotDiscTaxP As Double, TotPrePayTxP As Double, TotOpt1P As Double, TotIntTaxP#
    ' ReDim DecalDesc$(100), Decaltot#(100)
    ReDim TotalUtilRevAmt#(15)
    ReDim TotalDepRevAmt#(15)
    ReDim RevText$(15)
    Dim MCFile As Integer
    ReDim UBSetUpRec(1) As UBSetupRecType
    ReDim DistArray(1 To 1) As DistArrayType
    FrmShowPctComp.Label1 = "Creating Cash Management Journal Report"
    FrmShowPctComp.Show , Me
    VOnly = False
    Vflag = False
    VTot = 0
    totPay = 0
    BegDate = Date2Num(txtDate1)
    FromDate$ = txtDate1
    EndDate = Date2Num(txtDate2)
    ThruDate$ = txtDate2
    If fpcboSource.ListIndex = 1 Then
        RecSource$ = "U"
    ElseIf fpcboSource.ListIndex = 2 Then
      RecSource$ = "M"
    ElseIf fpcboSource.ListIndex = 3 Then
      RecSource$ = "B"
    ElseIf fpcboSource.ListIndex = 4 Then
      RecSource$ = "T"
    ElseIf fpcboSource.ListIndex = 5 Then
      RecSource$ = "D"
    ElseIf fpcboSource.ListIndex = 6 Then
      RecSource$ = "A"
      VOnly = True
    Else
      RecSource$ = "A"
    End If
    OperatorNumber = Val(fptxtOperator)
    If OperatorNumber = 0 Then
      PrnOpr$ = "All"
    Else
      PrnOpr$ = Str$(OperatorNumber)
    End If
    If fpcboRptType.ListIndex = 0 Then
      RptType = 0
    Else
      RptType = 2
    End If
    If fpcboPrintOrder.ListIndex = 0 Then
      SortOrder$ = "Entry Order"
    Else
      SortOrder$ = "Name Order"
    End If
  
  
    Select Case intHasTaxes
    Case 1 'NC Taxes
        ReDim TaxMasterRec(1) As TaxMasterType
        OpenTaxSetUpFile TMHandle
        Get TMHandle, 1, TaxMasterRec(1)
        Close TMHandle
        TxOpt1 = Mid$(TaxMasterRec(1).OptRev1, 1, 5)
        TxOpt2 = Mid$(TaxMasterRec(1).OptRev2, 1, 5)
        TxOpt3 = Mid$(TaxMasterRec(1).OptRev3, 1, 5)
        TxOpt1F = QPTrim$(TaxMasterRec(1).OptRev1)
        TxOpt2F = QPTrim$(TaxMasterRec(1).OptRev2)
        TxOpt3F = QPTrim$(TaxMasterRec(1).OptRev3)
    Case 2 'VA Taxes
        ReDim TaxMaster(1) As VATaxMasterType
        OpenVATaxSetUpFile TMHandle
        Get TMHandle, 1, TaxMaster(1)
        Close TMHandle
        TxOpt1 = Mid$(TaxMaster(1).OptRev1, 1, 5)
        TxOpt2 = Mid$(TaxMaster(1).OptRev2, 1, 5)
        TxOpt3 = Mid$(TaxMaster(1).OptRev3, 1, 5)
        TxOpt1F = QPTrim$(TaxMaster(1).OptRev1)
        TxOpt2F = QPTrim$(TaxMaster(1).OptRev2)
        TxOpt3F = QPTrim$(TaxMaster(1).OptRev3)
        TxPOpt1 = Mid$(TaxMaster(1).POptRev1, 1, 5)
        TxPOpt2 = Mid$(TaxMaster(1).POptRev2, 1, 5)
        TxPOpt3 = Mid$(TaxMaster(1).POptRev3, 1, 5)
        TxPOpt1F = QPTrim$(TaxMaster(1).POptRev1)
        TxPOpt2F = QPTrim$(TaxMaster(1).POptRev2)
        TxPOpt3F = QPTrim$(TaxMaster(1).POptRev3)
    Case Else
        'doesn't have taxes.
    End Select
  
'  If Exist(UBPath$ + "CitiTaxes.EXE") Then
'    If Exist("TAXSETUP.DAT") Then
'      ReDim TaxMasterRec(1) As TaxMasterType
'      OpenTaxSetUpFile TMHandle
'      Get TMHandle, 1, TaxMasterRec(1)
'      Close TMHandle
'      TxOpt1 = Mid$(TaxMasterRec(1).OptRev1, 1, 5)
'      TxOpt2 = Mid$(TaxMasterRec(1).OptRev2, 1, 5)
'      TxOpt3 = Mid$(TaxMasterRec(1).OptRev3, 1, 5)
'      TxOpt1F = QPTrim$(TaxMasterRec(1).OptRev1)
'      TxOpt2F = QPTrim$(TaxMasterRec(1).OptRev2)
'      TxOpt3F = QPTrim$(TaxMasterRec(1).OptRev3)
'    End If
'  ElseIf Exist(UBPath$ + "VACitiTax.EXE") Then
'    If Exist("TAXSETUP.DAT") Then
'      ReDim TaxMaster(1) As VATaxMasterType
'      OpenVATaxSetUpFile TMHandle
'      Get TMHandle, 1, TaxMaster(1)
'      Close TMHandle
'      TxOpt1 = Mid$(TaxMaster(1).OptRev1, 1, 5)
'      TxOpt2 = Mid$(TaxMaster(1).OptRev2, 1, 5)
'      TxOpt3 = Mid$(TaxMaster(1).OptRev3, 1, 5)
'      TxOpt1F = QPTrim$(TaxMaster(1).OptRev1)
'      TxOpt2F = QPTrim$(TaxMaster(1).OptRev2)
'      TxOpt3F = QPTrim$(TaxMaster(1).OptRev3)
'      TxPOpt1 = Mid$(TaxMaster(1).POptRev1, 1, 5)
'      TxPOpt2 = Mid$(TaxMaster(1).POptRev2, 1, 5)
'      TxPOpt3 = Mid$(TaxMaster(1).POptRev3, 1, 5)
'      TxPOpt1F = QPTrim$(TaxMaster(1).POptRev1)
'      TxPOpt2F = QPTrim$(TaxMaster(1).POptRev2)
'      TxPOpt3F = QPTrim$(TaxMaster(1).POptRev3)
'    End If
'  End If
  'End of Input
  '=====================================================
  'Start Report Processing

  ReportFile$ = UBPath$ + "CMJOURNL.PRN"  'Report File Name
  Fmt1$ = "###,###.##"
  Fmt2$ = "###,###.##"
  Fmt3$ = "$#,###,###,###.##"
  Fmt4$ = "$#,###,###,###.##"

  FF$ = Chr$(12)
  If RptType = 2 Then
    MaxLines = 53
  Else
    MaxLines = 51
  End If
  LineCnt = 0
  TotDr# = 0
  TotCr# = 0
  ReDim CMTrRec(1) As CMTransRecType            ' open transaction file
  CMTrRecLen = Len(CMTrRec(1))
  TRHandle = FreeFile
  Open UBPath$ + "CMTRANS.DAT" For Random Access Read Write Shared As TRHandle Len = CMTrRecLen
  TrNumRecs& = LOF(TRHandle) \ CMTrRecLen

  Max& = TrNumRecs& '(FRE(-1) - 16000) \ 16
  Size = Max&

  Start = 1     'start at array element 1
  sDir = 0       'sort direction - use anything else for descending
  SSize = 16    'total size of each TYPE element
  MOff = 0      'offset into the TYPE for the key element
  MSize = 16    'size of the key element - coded as follows:

  '   -1 = integer
  '   -2 = long integer
  '   -3 = single precision
  '   -4 = double precision
  '   +N = TYPE array/fixed-length string of length N

  ReDim Array1(1 To Size) As struct

  GoSub GetReportInformation

  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle

'  ReDim OperRec(1) As CMOperRecType             ' opens operatorfile
'  OperRecLen = Len(OperRec(1))
'  OperFile = FreeFile
'  Open "CMOPER.DAT" For Random As OperFile Len = OperRecLen
'  NumOperRecs = LOF(OperFile) / OperRecLen
'  Get OperFile, OperRecNumb, OperRec(1)
  MCFile = FreeFile
  OpenMiscCodeFile NumOfMiscRecs     ' opens misc code file
  ReDim MiscCodeRec(1) As MiscCodeRecType
  If RptType = 2 Then
    Print #RptHandle, Chr$(27); Chr$(58);         ' oki 320 12 cpi
  End If
  GoSub PrintRptHeader

  For cnt = 1 To CntG
    FrmShowPctComp.ShowPctComp cnt, CntG
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ActivateControls Me
      Exit Sub
    End If

    Get TRHandle, Array1(cnt).RecNum, CMTrRec(1)
    If CMTrRec(1).TransDate >= BegDate And CMTrRec(1).TransDate <= EndDate Then
    totPay = Round(CMTrRec(1).TransCheck + CMTrRec(1).TransCash)
    If VOnly = True Then
      If Not CMTrRec(1).TransSource > 200 Then
        GoTo NotVSkip
      End If
    End If
    If OperatorNumber = 0 Or OperatorNumber = CMTrRec(1).TransOperNum Then
      If Check1.Value = 1 Then
        If Check2.Value = 1 Then
         If LineCnt >= MaxLines - 12 Then
              Print #RptHandle, FF$
              LineCnt = 1
            GoSub PrintRptHeader
          End If
        Else
          If LineCnt >= MaxLines - 8 Then
              Print #RptHandle, FF$
              LineCnt = 1
            GoSub PrintRptHeader
          End If
        End If
      Else
'        If Check2.Value = 1 Then
'          If Linecnt >= MaxLines - 10 Then
'            Print #RptHandle, FF$
'            GoSub PrintRptHeader
'          End If
'        Else
          If LineCnt >= MaxLines - 6 Then
            Print #RptHandle, FF$
            LineCnt = 1
            GoSub PrintRptHeader
          End If
      '  End If
      End If
      If CMTrRec(1).TransDate >= BegDate And CMTrRec(1).TransDate <= EndDate Then
        TRType$ = ""
        Vflag = False
        Select Case CMTrRec(1).TransSource
        Case 1
          TRType$ = "Misc."
        Case 201
          TRType$ = "V-Misc"
          Vflag = True
        Case 27
          TRType$ = "UT-Dep."
        Case 227
          TRType$ = "V-UTDep"
          Vflag = True
        Case 24
          TRType$ = "Util."
        Case 224
          TRType$ = "V-Util"
          Vflag = True
        Case 30 To 39, 131, 161, 171
          TRType$ = "Tax"
        Case 231, 261, 271
          TRType$ = "V-Tax"
          Vflag = True
        Case 40 To 49, 141
          TRType$ = "Lic."
        Case 241
          TRType$ = "V-Lic"
          Vflag = True
        Case 50 To 59, 151
          TRType$ = "Decal"
        Case 251
          TRType$ = "V-Decal"
          Vflag = True
        End Select
        
      '  If RptType = 2 Then
          Print #RptHandle, Num2Date(CMTrRec(1).TransDate); Tab(12); TRType$; Tab(20); Left$(CMTrRec(1).TransName, 18);
          Print #RptHandle, Tab(38); Using(Fmt1$, CMTrRec(1).TransCash);
          If CMTrRec(1).TransTender = 4 Then
            Print #RptHandle, Tab(49); Using(Fmt1$, 0#);
            Print #RptHandle, Tab(60); Using(Fmt1$, CMTrRec(1).TransCheck);
          Else
            Print #RptHandle, Tab(49); Using(Fmt1$, CMTrRec(1).TransCheck);
            Print #RptHandle, Tab(60); Using(Fmt1$, 0#);
          End If
          Print #RptHandle, Tab(73); Using(Fmt1$, CMTrRec(1).TransAmount); 'totPay); 'Round(CMTrRec(1).TransCheck + CMTrRec(1).TransCash)) 'CMTrRec(1).TransAmtOwed);
       ' Else
      '    ToPrint$ = Num2Date(CMTRRec(1).TransDate) + "~" + TRType$ + "~" + Left$(CMTRRec(1).TransName, 18)
      '    ToPrint$ = ToPrint$ + "~" + Using(Fmt1$, CMTRRec(1).TransCash) + "~" + Using(Fmt1$, CMTRRec(1).TransCheck) + "~" + Using(Fmt1$, CMTRRec(1).TransAmtOwed)
       ' End If
        TotalCash# = Round#(TotalCash# + CMTrRec(1).TransCash)
''        If CMTRRec(1).TransTender = 4 Then
''          Print #RptHandle, Tab(70); Using("######.##", 0);
''          Print #RptHandle, Tab(84); Using("######.##", UBPaymentRec(1).CHKAMT);
''        Else
''          Print #RptHandle, Tab(70); Using("######.##", UBPaymentRec(1).CHKAMT);
''          Print #RptHandle, Tab(84); Using("######.##", 0);
''        End If
        If CMTrRec(1).TransTender = 4 Then
          TotalCharge# = Round(TotalCharge# + CMTrRec(1).TransCheck)
        ElseIf CMTrRec(1).TransCheck <> 0 Then
          TotalCheck# = Round#(TotalCheck# + CMTrRec(1).TransCheck)
          If Not Vflag Then
            NumChks = NumChks + 1
          End If
        End If
        Select Case CMTrRec(1).TransSource
        Case 20 To 29, 224, 227
          TxRev# = 0
          For TRev = 1 To 15
            TxRev# = Round#(TxRev# + CMTrRec(1).TransRevAmt(TRev))
          Next
          TotalAmount# = Round#(TotalAmount# + TxRev#)
          Change# = Round#((CMTrRec(1).TransCash + CMTrRec(1).TransCheck) - TxRev#)
        Case 30 To 39
        'ifCMTrRec(1).TransSource >= 30 And CMTrRec(1).TransSource <= 39 Then
          TxRev# = 0
          For TRev = 1 To 9
            TxRev# = Round#(TxRev# + CMTrRec(1).TransRevAmt(TRev))
          Next
          TotalAmount# = Round#(TotalAmount# + TxRev#)
          Change# = Round#((CMTrRec(1).TransCash + CMTrRec(1).TransCheck) - TxRev#)
        Case 131, 231
          TxRev# = 0
          For TRev = 1 To 7
            TxRev# = Round#(TxRev# + CMTrRec(1).TransRevAmt(TRev))
          Next
          TxRev# = Round#(TxRev# + CMTrRec(1).TransRevAmt(9))
          'TxRev# = Round#(TxRev# - CMTrRec(1).TransRevAmt(8))
          TotalAmount# = Round#(TotalAmount# + TxRev#)
          Change# = Round#((CMTrRec(1).TransCash + CMTrRec(1).TransCheck) - TxRev#)
        Case 161, 261
          TxRev# = 0
          For TRev = 1 To 8
            TxRev# = Round#(TxRev# + CMTrRec(1).TransRevAmt(TRev))
          Next
          TxRev# = Round#(TxRev# + CMTrRec(1).TransRevAmt(10))
          'TxRev# = Round#(TxRev# - CMTrRec(1).TransRevAmt(9))
          TotalAmount# = Round#(TotalAmount# + TxRev#)
          Change# = Round#((CMTrRec(1).TransCash + CMTrRec(1).TransCheck) - TxRev#)
        Case 171, 271
          TxRev# = 0
          For TRev = 1 To 10
            TxRev# = Round#(TxRev# + CMTrRec(1).TransRevAmt(TRev))
          Next
          TxRev# = Round#(TxRev# + CMTrRec(1).TransRevAmt(12))
          'TxRev# = Round#(TxRev# - CMTrRec(1).TransRevAmt(11))
          TotalAmount# = Round#(TotalAmount# + TxRev#)
          Change# = Round#((CMTrRec(1).TransCash + CMTrRec(1).TransCheck) - TxRev#)
        Case 40 To 49, 141, 241
        'ElseIf CMTrRec(1).TransSource >= 40 And CMTrRec(1).TransSource <= 49 Then
          If CMTrRec(1).TransAmount < totPay Then 'CMTrRec(1).TransAmtOwed Then
            TotalAmount# = Round#(TotalAmount# + CMTrRec(1).TransAmount)
          Else
            TotalAmount# = Round#(TotalAmount# + totPay) 'CMTrRec(1).TransAmtOwed)
          End If
          Change# = Round#((CMTrRec(1).TransCash + CMTrRec(1).TransCheck) - CMTrRec(1).TransAmount)
        Case 1, 201
        'ElseIf CMTrRec(1).TransSource = 1 Then
          TotalAmount# = Round#(TotalAmount# + totPay) 'CMTrRec(1).TransAmtOwed)
          Change# = Round#((CMTrRec(1).TransCash + CMTrRec(1).TransCheck) - CMTrRec(1).TransAmount)
        Case Else
          TotalAmount# = Round#(TotalAmount# + CMTrRec(1).TransCash + CMTrRec(1).TransCheck)
        'End If
          Change# = Round#((CMTrRec(1).TransCash + CMTrRec(1).TransCheck) - CMTrRec(1).TransAmount)
        End Select
'        If CMTRRec(1).TransSource = 27 Then
'          CHANGE# = Round#((CMTRRec(1).TransCash + CMTRRec(1).TransCheck) - CMTRRec(1).TransAmtOwed)
'        Else
'          CHANGE# = Round#(CMTRRec(1).TransAmount - CMTRRec(1).TransAmtOwed)
'        End If
'?????????HEY LOOK HERE MADE THIS CHANGE WHY WOULDN'T WORK??????????????
        
        'If CHANGE# <> 0 Then CHANGE# = 0
        'IF Change# <> 0 THEN STOP
        TotChange# = Round#(TotChange# + Change#)
        'If RptType = 2 Then
          Print #RptHandle, Tab(86); Using(Fmt1$, Change#)
          If Check2.Value = 1 Then
            Print #RptHandle, "Account-"; CMTrRec(1).TransAcctNum; Tab(20); CMTrRec(1).TransDesc
            LineCnt = LineCnt + 1
          End If
        'PRINT #RptHandle, TAB(84); USING Fmt1$; ((CMTRRec(1).TransCheck + CMTRRec(1).TransCash) - CMTRRec(1).TransAmtOwed)
       ' Else
       '   ToPrint$ = ToPrint$ + "~" + Using(Fmt1$, CHANGE#)
       ' End If
        If Not Vflag Then
          TotalReceipts = TotalReceipts + 1
        Else
          VTot = VTot + 1
        End If
        LineCnt = LineCnt + 1
        Select Case CMTrRec(1).TransSource
        Case 1, 201
        'If CMTrRec(1).TransSource = 1 Then
          'Second Line of Print is Misc Code Breakdown Dist.****************
'          If Linecnt >= MaxLines - 7 Then
'          '  If RptType = 2 Then
'              Print #RptHandle, FF$
'          '  End If
'            GoSub PrintRptHeader
'          End If
 
          PrintMiscFlag = 0
         ' If RptType = 2 Then
            'Print #RptHandle, "Description: "; CMTrRec(1).TransDesc
         ' Else
         '   ToPrint$ = ToPrint$ + "~" + QPTrim$(CMTRRec(1).TransDesc)
         ' End If
          For MCnt = 1 To 5
            MiscRevAmt# = (CMTrRec(1).TransRevAmt(MCnt))
            MiscRevAmt# = Round#(MiscRevAmt#)
            If MiscRevAmt# <> 0 Then
              ' If There Is an Amount in Misc Rev 1-5 then get code record number from 6-10
              If CMTrRec(1).TransRevAmt(MCnt + 5) <> 0 Then
                Get MCFile, CMTrRec(1).TransRevAmt(MCnt + 5), MiscCodeRec(1)
 '               If RptType = 2 Then
                 If Check1.Value = 1 Then
                  Print #RptHandle, "Code BrkDwn:";
                  Print #RptHandle, Tab(14); MiscCodeRec(1).MiscCode;
                  Print #RptHandle, Tab(25); MiscCodeRec(1).Description;
                  Print #RptHandle, Tab(55); Using(Fmt1$, MiscRevAmt#)
                  PrintMiscFlag = 1
                  LineCnt = LineCnt + 1
                 End If
'                Else
'                  ToPrint$ = ToPrint$ + "~" + QPTrim$(MiscCodeRec(1).MiscCode)
'                  ToPrint$ = ToPrint$ + "~" + QPTrim$(MiscCodeRec(1).Description)
'                  ToPrint$ = ToPrint$ + "~" + Using(Fmt1$, MiscRevAmt#)
'                End If
                GoSub SubTotalMisc
              End If
            End If
          Next MCnt
         ' If RptType = 2 Then
            If PrintMiscFlag = 1 Then Print #RptHandle, String$(96, "-"): LineCnt = LineCnt + 1
         ' End If
          'End Misc Code Print on Second Line ********************************
        'End If
        Case 20 To 29, 224, 227
        'If CMTrRec(1).TransSource >= 20 And CMTrRec(1).TransSource <= 29 Then
          If CMTrRec(1).TransSource <> 27 And CMTrRec(1).TransSource <> 227 Then
            'Second Line of Print is Utility Breakdown Dist. *****************
            GoSub GetRevenueSources
            If NumofRevs <> 0 Then
             ' If RptType = 2 Then
'              If Linecnt >= MaxLines - 7 Then
'              '  If RptType = 2 Then
'                  Print #RptHandle, FF$
'              '  End If
'                GoSub PrintRptHeader
'              End If
               If Check1.Value = 1 Then
               Print #RptHandle, "Util BrkDwn:";
             ' End If
              For RCnt = 1 To NumofRevs Step 2
 '               If RptType = 2 Then
                  Print #RptHandle, Tab(15); RevText$(RCnt);
                  Print #RptHandle, Tab(40); Using(Fmt1$, CMTrRec(1).TransRevAmt(RCnt));
                  Print #RptHandle, Tab(55); RevText$(RCnt + 1);
                  Print #RptHandle, Tab(80); Using(Fmt1$, CMTrRec(1).TransRevAmt(RCnt + 1))
                  PrintUtilFlag = 1
                  LineCnt = LineCnt + 1
'                Else
'                  ToPrint$ = ToPrint$ + "~" + QPTrim$(RevText$(RCnt))
'                  ToPrint$ = ToPrint$ + "~" + Using(Fmt1$, CMTRRec(1).TransRevAmt(RCnt))
'                  ToPrint$ = ToPrint$ + "~" + QPTrim$(RevText$(RCnt + 1))
'                  ToPrint$ = ToPrint$ + "~" + Using(Fmt1$, CMTRRec(1).TransRevAmt(RCnt + 1))
'                End If
              Next RCnt
              End If
              GoSub SubTotalUtil
            End If
         '   If RptType = 2 Then
           '  If Check1.Value = 1 Then
              If PrintUtilFlag = 1 Then Print #RptHandle, String$(96, "-"): LineCnt = LineCnt + 1
            ' End If
            'End of Utility Print on Second Line *****************************
          Else
            GoSub SubTotalDep
         '   If RptType = 2 Then
          '   If Check1.Value = 1 Then
'              Print #RptHandle, String$(96, "-")
'              Linecnt = Linecnt + 1
          '   End If
          End If
        'End If
        Case 30 To 39
        'If CMTrRec(1).TransSource >= 30 And CMTrRec(1).TransSource <= 39 Then
          'Second Line of Print is Tax Breakdown Dist.     *****************
'          If Linecnt >= MaxLines - 3 Then
'          '  If RptType = 2 Then
'              Print #RptHandle, FF$
'          '  End If
'            GoSub PrintRptHeader
'          End If
          If Check1.Value = 1 Then
            Print #RptHandle, "Tax BrkDwn:";
            Print #RptHandle, Tab(15); "Tax:"; Using(Fmt1$, CMTrRec(1).TransRevAmt(1));
            Print #RptHandle, Tab(32); "Int.: "; Using(Fmt2$, CMTrRec(1).TransRevAmt(2));
            Print #RptHandle, Tab(50); "Pen.: "; Using(Fmt2$, CMTrRec(1).TransRevAmt(3));
            Print #RptHandle, Tab(65); "Strm: "; Using(Fmt2$, CMTrRec(1).TransRevAmt(4))
            Print #RptHandle, Tab(5); " Past Tax: "; Using(Fmt1$, CMTrRec(1).TransRevAmt(6));
            Print #RptHandle, Tab(32); "Int.: "; Using(Fmt2$, CMTrRec(1).TransRevAmt(7));
            Print #RptHandle, Tab(50); "Pen.: "; Using(Fmt2$, CMTrRec(1).TransRevAmt(8));
            Print #RptHandle, Tab(65); "Strm: "; Using(Fmt2$, CMTrRec(1).TransRevAmt(9))
            PrintTaxFlag = 1
            LineCnt = LineCnt + 2
          End If
          GoSub SubTotalTax
        'End If
        If PrintTaxFlag = 1 Then
          'If Check1.Value = 1 Then
            Print #RptHandle, String$(96, "-")
            LineCnt = LineCnt + 1
          'End If
        End If
        'End of Tax Print on Second Line *********************************
        Case 131, 231
          If Check1.Value = 1 Then
            Print #RptHandle, "Tax BrkDwn:";
            Print #RptHandle, Tab(15); "PrePay: "; Using(Fmt2$, CMTrRec(1).TransRevAmt(9));
            Print #RptHandle, Tab(32); "#of Bills: "; CInt(CMTrRec(1).TransRevAmt(10))
            Print #RptHandle, Tab(5); "Prcpl:"; Using(Fmt1$, CMTrRec(1).TransRevAmt(1));
            Print #RptHandle, Tab(32); "Ints: "; Using(Fmt2$, CMTrRec(1).TransRevAmt(2));
            Print #RptHandle, Tab(55); "Coll: "; Using(Fmt2$, CMTrRec(1).TransRevAmt(3));
            Print #RptHandle, Tab(75); "Late: "; Using(Fmt2$, CMTrRec(1).TransRevAmt(4))
            Print #RptHandle, Tab(5); TxOpt1$; ": "; Using(Fmt1$, CMTrRec(1).TransRevAmt(5));
            Print #RptHandle, Tab(32); TxOpt2$; ": "; Using(Fmt2$, CMTrRec(1).TransRevAmt(6));
            Print #RptHandle, Tab(55); TxOpt3$; ": "; Using(Fmt2$, CMTrRec(1).TransRevAmt(7));
            Print #RptHandle, Tab(75); "Discnt: "; Using(Fmt2$, CMTrRec(1).TransRevAmt(8))
            PrintTaxFlag = 1
            LineCnt = LineCnt + 3
          End If
          GoSub SubTotalTax2
          If PrintTaxFlag = 1 Then
            Print #RptHandle, String$(96, "-")
            LineCnt = LineCnt + 1
          End If
        Case 161, 261
          If Check1.Value = 1 Then
            Print #RptHandle, "Tax Real BrkDwn:";
            Print #RptHandle, Tab(15); "PrePay: "; Using(Fmt2$, CMTrRec(1).TransRevAmt(10));
            Print #RptHandle, Tab(32); "#of Bills: "; CInt(CMTrRec(1).TransRevAmt(11))
            Print #RptHandle, Tab(5); "Prcpl:"; Using(Fmt1$, CMTrRec(1).TransRevAmt(1));
            Print #RptHandle, Tab(32); "Ints: "; Using(Fmt2$, CMTrRec(1).TransRevAmt(2));
            Print #RptHandle, Tab(55); "Coll: "; Using(Fmt2$, CMTrRec(1).TransRevAmt(3));
            Print #RptHandle, Tab(75); "Late: "; Using(Fmt2$, CMTrRec(1).TransRevAmt(4))
            Print #RptHandle, Tab(5); "Penty: "; Using(Fmt1$, CMTrRec(1).TransRevAmt(5));
            Print #RptHandle, Tab(32); TxOpt1$; ": "; Using(Fmt2$, CMTrRec(1).TransRevAmt(6));
            Print #RptHandle, Tab(55); TxOpt2$; ": "; Using(Fmt2$, CMTrRec(1).TransRevAmt(7));
            Print #RptHandle, Tab(75); TxOpt3$; ": "; Using(Fmt2$, CMTrRec(1).TransRevAmt(8))
            Print #RptHandle, Tab(15); "Discnt: "; Using(Fmt2$, CMTrRec(1).TransRevAmt(9))
            PrintTaxFlag = 1
            LineCnt = LineCnt + 4
          End If
          GoSub SubTotalTaxVA
          If PrintTaxFlag = 1 Then
            Print #RptHandle, String$(96, "-")
            LineCnt = LineCnt + 1
          End If
        Case 171, 271
          If Check1.Value = 1 Then
            Print #RptHandle, "Tax Pers BrkDwn:";
            Print #RptHandle, Tab(15); "PrePay: "; Using(Fmt2$, CMTrRec(1).TransRevAmt(12));
            Print #RptHandle, Tab(32); "#of Bills: "; CInt(CMTrRec(1).TransRevAmt(13))
            Print #RptHandle, Tab(5); "Prcpl1:"; Using(Fmt1$, CMTrRec(1).TransRevAmt(1));
            Print #RptHandle, Tab(32); "Prcpl2: "; Using(Fmt2$, CMTrRec(1).TransRevAmt(2));
            Print #RptHandle, Tab(55); "Prcpl3: "; Using(Fmt2$, CMTrRec(1).TransRevAmt(3));
            Print #RptHandle, Tab(75); "Prcpl4: "; Using(Fmt2$, CMTrRec(1).TransRevAmt(4))
            Print #RptHandle, Tab(5); "Prcpl5: "; Using(Fmt1$, CMTrRec(1).TransRevAmt(5));
            Print #RptHandle, Tab(32); ; "Int: "; Using(Fmt2$, CMTrRec(1).TransRevAmt(6));
            Print #RptHandle, Tab(55); ; "Pen: "; Using(Fmt2$, CMTrRec(1).TransRevAmt(7));
            Print #RptHandle, Tab(75); TxPOpt1$; ": "; Using(Fmt2$, CMTrRec(1).TransRevAmt(8))
            Print #RptHandle, Tab(5); TxPOpt2$; ": "; Using(Fmt2$, CMTrRec(1).TransRevAmt(9));
            Print #RptHandle, Tab(32); TxPOpt3$; ": "; Using(Fmt2$, CMTrRec(1).TransRevAmt(10));
            Print #RptHandle, Tab(55); "Discnt: "; Using(Fmt2$, CMTrRec(1).TransRevAmt(11))
            PrintTaxFlag = 1
            LineCnt = LineCnt + 4
          End If
          GoSub SubTotalTaxVA
          If PrintTaxFlag = 1 Then
            Print #RptHandle, String$(96, "-")
            LineCnt = LineCnt + 1
          End If

        Case 40 To 49, 141, 241
'Go thru this to see what need for detail of categories later on
'        If CMTrRec(1).TransSource = 141 Or CMTrRec(1).TransSource = 241 Then
'          For MCnt = 1 To 5
'            BLRevAmt# = (CMTrRec(1).TransRevAmt(MCnt + 5))
'            BLRevAmt# = Round#(BLRevAmt#)
'            If BLRevAmt# > 0 Then
'              ' If There Is an Amount in Misc Rev 1-5 then get code record number from 6-10
'              If CMTrRec(1).TransRevAmt(MCnt + 5) >= 1 Then
'               GetCatDesc (QPTrim$(CustRec.BILLCAT1))
'                Get MCFile, CMTrRec(1).TransRevAmt(MCnt + 5), MiscCodeRec(1)
' '               If RptType = 2 Then
'                  Print #RptHandle, "Code BrkDwn:";
'                  Print #RptHandle, Tab(14); MiscCodeRec(1).MiscCode;
'                  Print #RptHandle, Tab(25); MiscCodeRec(1).Description;
'                  Print #RptHandle, Tab(55); Using(Fmt1$, MiscRevAmt#)
'                  PrintMiscFlag = 1
'                  Linecnt = Linecnt + 1
'              End If
'            End If
'          End If
          GoSub SubTotalBL
          'If Check1.Value = 1 Then
'            Print #RptHandle, String$(96, "-")
'            Linecnt = Linecnt + 1
          'End If
        'End If
        Case 50 To 59, 151, 251
'          If CMTrRec(1).TransSource >= 151 Then
         If Check1.Value = 1 Then PrintDecalFlag = 1
'          End If
          GoSub SubTotalDC
          If PrintDecalFlag = 1 Then Print #RptHandle, String$(96, "-"): LineCnt = LineCnt + 1
          'If Check1.Value = 1 Then
'            Print #RptHandle, String$(96, "-")
'            Linecnt = Linecnt + 1
          'End If
        End Select
      End If
    End If
    End If
NotVSkip:
  Next cnt

  GoSub PrintRptEnding
  If RptType = 2 Then
    Print #RptHandle, Chr$(18);   ' oki 320 12 cpi
  End If
  Close         'Close all open files now


  Erase RevName$, TotalMiscRec$, TotalMiscDesc$, TotalMiscAmt#
  Erase TotalUtilRevAmt#, MiscCodeGL$
  Erase Array1, CMTrRec, RevText$, MiscCodeRec, UBSetUpRec
  Erase DistArray
  If TotalReceipts > 0 Or VTot > 0 Then
    If RptType = 2 Then
      ViewPrint ReportFile$, Header$, True
    Else
      Load frmLoadingRpt
      frmLoadingRpt.setwherefrom frmRptCMJournal
        ARptLineRpt.GetName ReportFile$
        ARptLineRpt.startrpt
    End If
  Else
    Unload FrmShowPctComp
    MsgBox "No Transactions to Print", vbOKOnly, "No Transactions"
    ActivateControls Me
  End If
  Exit Sub

PrintRptHeader:
  Page = Page + 1
  'If RptType <> 2 Then
    Print #RptHandle, " "
    Print #RptHandle, " "
  'End If
  Print #RptHandle, Tab(27); "Cash Receipts Journal : Cash Management System"
  Print #RptHandle, "  Current Date: "; Now
  Print #RptHandle, "Beginning Date: "; FromDate$
  Print #RptHandle, "   Ending Date: "; ThruDate$
  Print #RptHandle, "      Operator: "; PrnOpr$; Tab(83); "Page #"; Page
  Print #RptHandle, " "
  Print #RptHandle, "   Date"; Tab(12); "Source"; Tab(20); "Name/Desc"; Tab(44); "Cash"; Tab(54); "Check"; Tab(64); "Charge"; Tab(75); "Amt Applied"; Tab(90); "Change"

  Print #RptHandle, String$(96, "=")
  LineCnt = 9
Return

PrintRptEnding:
  Print #RptHandle, FF$
  'If RptType <> 2 Then
    Print #RptHandle, " "
    Print #RptHandle, " "
  'End If
  Print #RptHandle, Tab(27); "Cash Receipts Journal : Cash Management System"
  Print #RptHandle, "  Current Date: "; Now
  Print #RptHandle, "Beginning Date: "; FromDate$
  Print #RptHandle, "   Ending Date: "; ThruDate$
  Print #RptHandle, "      Operator: "; PrnOpr$
  Print #RptHandle, " "
  Print #RptHandle, "Totals Page for Operator # "; PrnOpr$
  Print #RptHandle, " "
  Print #RptHandle, "   Total Cash Received: "; Using(Fmt3$, TotalCash#)
  Print #RptHandle, " Total Checks Received: "; Using(Fmt3$, TotalCheck#)
  Print #RptHandle, "Total Charges Received: "; Using(Fmt3$, TotalCharge#)
  Print #RptHandle, "                      -----------------"
  Print #RptHandle, " Total Amount Received: "; Using(Fmt3$, TotalCash# + TotalCheck#)
  Print #RptHandle, "Amount Applied to Acct: "; Using(Fmt3$, (TotalCash# + TotalCheck# + TotalCharge#) - TotChange#)
  Print #RptHandle, "    Total Change Given: "; Using(Fmt3$, TotChange#)
  Print #RptHandle, ""
  Print #RptHandle, "   Bank Deposit Amount: "; Using(Fmt3$, (TotalCash# + TotalCheck#) - TotChange#)
  Print #RptHandle, "    Number of Receipts: "; Using("##,###", TotalReceipts)
  Print #RptHandle, "      Number of Checks: "; Using("##,###", NumChks)
  Print #RptHandle, "       Number of Voids: "; Using("##,###", VTot)
  LineCnt = 22
  If RecSource$ = "M" Or RecSource$ = "A" Then
    GoSub PrintTotalMisc
  End If

  If RecSource$ = "U" Or RecSource$ = "A" Then
    GoSub PrintTotalUtil
  End If

  If RecSource$ = "T" Or RecSource$ = "A" Then
    GoSub PrintTotalTax
  End If

  If RecSource$ = "B" Or RecSource$ = "A" Then
    GoSub PrintBLTotal          ' Not Active Yet!!!
  End If
  If RecSource$ = "D" Or RecSource$ = "A" Then
    GoSub PrintDCTotal
  End If
  If RptType <> 0 Then
    Print #RptHandle, FF$
    LineCnt = 1
  End If
Return

GetReportInformation:
  BegRecNo& = 1   'TrNumRecs& - 7500    ' Move back 7500 records to beg
  If BegRecNo& < 1 Then BegRecNo& = 1         ' Don't Allow Less Than 1
  'If OperatorNumber = 0 Then BegRecNo& = 1

  For cnt& = BegRecNo& To TrNumRecs&
    Get TRHandle, cnt&, CMTrRec(1)
    TransDate = CMTrRec(1).TransDate
    GoodRecordFlag = False
    If VOnly = True Then
      If Not CMTrRec(1).TransSource > 200 Then
        GoTo okSkip
      End If
    End If
    If OperatorNumber = 0 Or OperatorNumber = CMTrRec(1).TransOperNum Then
    'IF CMTRRec(1).TransOperNum = OperRecNumb AND (TransDate >= BegDate AND TransDate <= EndDate) THEN
    If (TransDate >= BegDate And TransDate <= EndDate) Then
      If RecSource$ = "A" Then   'All
        GoodRecordFlag = True
      End If
      If RecSource$ = "M" Then   'Miscellaneous
       If CMTrRec(1).TransSource = 1 Or CMTrRec(1).TransSource = 201 Then
        GoodRecordFlag = True
       End If
      End If
      If RecSource$ = "U" Then  'Utilities
       Select Case CMTrRec(1).TransSource
       Case 20 To 29, 224, 227
        GoodRecordFlag = True
       End Select
      End If
      If RecSource$ = "T" Then  'Taxes
       Select Case CMTrRec(1).TransSource
       Case 30 To 39, 131, 231, 161, 261, 171, 271
        GoodRecordFlag = True
       End Select
      End If
      If RecSource$ = "B" Then 'Business License
       Select Case CMTrRec(1).TransSource
       Case 40 To 49, 141, 241
        GoodRecordFlag = True
       End Select
      End If
      If RecSource$ = "D" Then 'Decals
       Select Case CMTrRec(1).TransSource
       Case 50 To 59, 151, 251
        GoodRecordFlag = True
       End Select
      End If
    End If
    End If
    If GoodRecordFlag Then
      CntG = CntG + 1
      If CntG > Size Then
'        SaveScrn TempScrn()
'        DisplayUBScrn "ERRSCRN1"
'        QPrintRC "TO MANY TRANSACTIONS!", 10, 30, -1
'        UseW$ = "Will Display First:" + Str$(CntG) + " Transactions."
'        OffSet = ((80 - Len(UseW$)) / 2)
'        QPrintRC UseW$, 11, OffSet, -1
'        QPrintRC "Press any key to continue!", 13, 28, -1
'        WaitForAction
'        RestScrn TempScrn()
        CntG = Size
        Exit For
      End If

      'PrintHelp Help$ + " Count:" + STR$(Count)
      If SortOrder$ = "Entry Order" Then
        Array1(CntG).who = Str$(cnt)
      Else
        Array1(CntG).who = Left$(CMTrRec(1).TransName, 12)
      End If
      Array1(CntG).RecNum = cnt
    End If
okSkip:
  Next cnt
  If Not SortOrder$ = "Entry Order" Then
    NameEntryQSort Array1(), 1, CntG
  End If
'  SortT Array1(Start), CntG, Dir, SSize, MOff, MSize
Return
'GetDecalCodes:
'  Dim DCCatCodeRec As DCCatCodeRecType
'  Dim DCCatCodeRecLen As Integer, ghandle As Integer, cnt As Integer
'  Dim NumOFDCCatRecs As Integer
'  DCCatCodeRecLen = Len(DCCatCodeRec)
'  ghandle = FreeFile
'  Open "DCCODE.DAT" For Random Access Read Write Shared As ghandle Len = DCCatCodeRecLen
'  NumOFDCCatRecs = LOF(ghandle) \ DCCatCodeRecLen
'  ReDim DecalDesc$(1 To NumOFDCCatRecs)
'  ReDim Decaltot#(1 To NumOFDCCatRecs)
'  For cnt = 1 To NumOFDCCatRecs
'    Get #ghandle, cnt, DCCatCodeRec
'    DecalDesc$(cnt) = QPTrim$(DCCodeRec.CODEDESC)
'    'QPTrim$ (DCCodeRec.CATCODE)
'  Next
'  Close #ghandle
'Return

GetRevenueSources:

  NumofRevs = MaxRevsCnt
  ReDim UBSetUpRec(1) As UBSetupRecType
  ReDim DistArray(1 To MaxRevsCnt) As DistArrayType
  ReDim UBSetUpRec(1) As UBSetupRecType
  UBSetupLen = Len(UBSetUpRec(1))
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen

  For RevCnt = 1 To MaxRevsCnt
    RevText$(RevCnt) = Left$(QPTrim$(UBSetUpRec(1).Revenues(RevCnt).RevName), 14)
    DistArray(RevCnt).DistOrder = UBSetUpRec(1).Revenues(RevCnt).DistOr
    DistArray(RevCnt).DistCnt = RevCnt
    If Len(RevText$(RevCnt)) = 0 Then
      NumofRevs = RevCnt - 1
      Exit For
    End If
  Next

  ReDim Preserve DistArray(1 To NumofRevs) As DistArrayType

  Do
    OutOfOrder = False          'assume it's sorted
    For x = 1 To NumofRevs - 1
      If DistArray(x).DistOrder > DistArray(x + 1).DistOrder Then
        Temp2 = DistArray(x).DistOrder
        DistArray(x).DistOrder = DistArray(x + 1).DistOrder
        DistArray(x + 1).DistOrder = Temp2
        'SWAP DistArray(x), DistArray(x + 1)     'if we had to swap
        OutOfOrder = True       'we're not done yet
      End If
    Next
  Loop While OutOfOrder

  TownName$ = UBSetUpRec(1).UTILNAME
Return
SubTotalUtil:
  For uCnt = 1 To NumofRevs
    TotalUtilRevAmt#(uCnt) = TotalUtilRevAmt#(uCnt) + CMTrRec(1).TransRevAmt(uCnt)
  Next uCnt
Return
 


SubTotalDep:
  For dcnt = 1 To 15
    TotalDepRevAmt#(dcnt) = Round#(TotalDepRevAmt#(dcnt) + CMTrRec(1).TransRevAmt(dcnt))
  Next
Return

PrintTotalUtil:
  If LineCnt >= MaxLines - 8 Then
    Print #RptHandle, FF$
    Print #RptHandle, " "
    Print #RptHandle, " "
    Print #RptHandle, Tab(27); "Cash Receipts Journal : Cash Management System"
    Print #RptHandle, "  Current Date: "; Now
    Print #RptHandle, "Beginning Date: "; FromDate$
    Print #RptHandle, "   Ending Date: "; ThruDate$
    Print #RptHandle, "      Operator: "; PrnOpr$
    Print #RptHandle, " "
    Print #RptHandle, "Totals Page for Operator # "; PrnOpr$
    Print #RptHandle, " "
    Print #RptHandle, String$(96, "-")
    Print #RptHandle, "GRAND TOTAL Utilities Receipts Recap"
    LineCnt = 12
  Else
    Print #RptHandle, String$(96, "-")
    Print #RptHandle, "GRAND TOTAL Utilities Receipts Recap"
    LineCnt = LineCnt + 2
  End If
  For uCnt = 1 To NumofRevs
    If LineCnt >= MaxLines - 8 Then
      Print #RptHandle, FF$
      Print #RptHandle, ""
      Print #RptHandle, ""
      Print #RptHandle, Tab(27); "Cash Receipts Journal : Cash Management System"
      Print #RptHandle, "  Current Date: "; Now
      Print #RptHandle, "Beginning Date: "; FromDate$
      Print #RptHandle, "   Ending Date: "; ThruDate$
      Print #RptHandle, "Totals Page for Operator # "; PrnOpr$
      Print #RptHandle, "Total Utilities Receipts Cont'd"
      Print #RptHandle, ""
      LineCnt = 9
    End If
    Print #RptHandle, RevText$(uCnt); Tab(32); Using(Fmt4$, TotalUtilRevAmt#(uCnt))
    LineCnt = LineCnt + 1
    TotalUtilAmt# = TotalUtilAmt# + TotalUtilRevAmt#(uCnt)
  Next uCnt
  Print #RptHandle, "GRAND Total Utility Receipts ... "; Using(Fmt4$, TotalUtilAmt#)
  Print #RptHandle,
  LineCnt = LineCnt + 2
  TotalDepAmt# = 0
  For uCnt = 1 To 15
    'PRINT #RptHandle, RevText$(Cnt!); TAB(34); USING Fmt4$; TotalDepRevAmt#(Cnt!)
    TotalDepAmt# = Round#(TotalDepAmt# + TotalDepRevAmt#(uCnt))
  Next uCnt
  Print #RptHandle, "GRAND Total Utility Deposits ... "; Using(Fmt4$, TotalDepAmt#)
  LineCnt = LineCnt + 1
Return

SubTotalTax:
  TotCurTax# = TotCurTax# + CMTrRec(1).TransRevAmt(1)
  TotCurInt# = TotCurInt# + CMTrRec(1).TransRevAmt(2)
  TotCurPen# = TotCurPen# + CMTrRec(1).TransRevAmt(3)
  TotStrmFee# = TotStrmFee# + CMTrRec(1).TransRevAmt(4)
  TotPastTax# = TotPastTax# + CMTrRec(1).TransRevAmt(6)
  TotPastInt# = TotPastInt# + CMTrRec(1).TransRevAmt(7)
  TotPastPen# = TotPastPen# + CMTrRec(1).TransRevAmt(8)
  TotPastStrm# = TotPastStrm# + CMTrRec(1).TransRevAmt(9)
Return
SubTotalTax2:
  TotPrincpleTx# = TotPrincpleTx# + CMTrRec(1).TransRevAmt(1)
  TotIntTax# = TotIntTax# + CMTrRec(1).TransRevAmt(2)
  TotCollTax# = TotCollTax# + CMTrRec(1).TransRevAmt(3)
  TotLateList# = TotLateList# + CMTrRec(1).TransRevAmt(4)
  TotOpt1# = TotOpt1# + CMTrRec(1).TransRevAmt(5)
  TotOpt2# = TotOpt2# + CMTrRec(1).TransRevAmt(6)
  TotOpt3# = TotOpt3# + CMTrRec(1).TransRevAmt(7)
  TotDiscTax# = TotDiscTax# + CMTrRec(1).TransRevAmt(8)
  TotPrePayTx# = TotPrePayTx# + CMTrRec(1).TransRevAmt(9)
  TotTaxBills# = TotTaxBills# + CMTrRec(1).TransRevAmt(10)
Return
SubTotalTaxVA:
  If CMTrRec(1).TransSource = 161 Or CMTrRec(1).TransSource = 261 Then
    TotPrincpleTx# = TotPrincpleTx# + CMTrRec(1).TransRevAmt(1)
    TotIntTax# = TotIntTax# + CMTrRec(1).TransRevAmt(2)
    TotCollTax# = TotCollTax# + CMTrRec(1).TransRevAmt(3)
    TotLateList# = TotLateList# + CMTrRec(1).TransRevAmt(4)
    TotPenTax# = TotPenTax# + CMTrRec(1).TransRevAmt(5)
    TotOpt1# = TotOpt1# + CMTrRec(1).TransRevAmt(6)
    TotOpt2# = TotOpt2# + CMTrRec(1).TransRevAmt(7)
    TotOpt3# = TotOpt3# + CMTrRec(1).TransRevAmt(8)
    TotDiscTax# = TotDiscTax# + CMTrRec(1).TransRevAmt(9)
    TotPrePayTx# = TotPrePayTx# + CMTrRec(1).TransRevAmt(10)
    TotTaxBillsR# = TotTaxBillsR# + CMTrRec(1).TransRevAmt(11)
  ElseIf CMTrRec(1).TransSource = 171 Or CMTrRec(1).TransSource = 271 Then
    TotPrinc1Tx# = TotPrinc1Tx# + CMTrRec(1).TransRevAmt(1)
    TotPrinc2Tx# = TotPrinc2Tx# + CMTrRec(1).TransRevAmt(2)
    TotPrinc3Tx# = TotPrinc3Tx# + CMTrRec(1).TransRevAmt(3)
    TotPrinc4Tx# = TotPrinc4Tx# + CMTrRec(1).TransRevAmt(4)
    TotPrinc5Tx# = TotPrinc5Tx# + CMTrRec(1).TransRevAmt(5)
    TotIntTaxP# = TotIntTaxP# + CMTrRec(1).TransRevAmt(6)
    TotPenTaxP# = TotPenTaxP# + CMTrRec(1).TransRevAmt(7)
    TotOpt1P# = TotOpt1P# + CMTrRec(1).TransRevAmt(8)
    TotOpt2P# = TotOpt2P# + CMTrRec(1).TransRevAmt(9)
    TotOpt3P# = TotOpt3P# + CMTrRec(1).TransRevAmt(10)
    TotDiscTaxP# = TotDiscTaxP# + CMTrRec(1).TransRevAmt(11)
    TotPrePayTxP# = TotPrePayTxP# + CMTrRec(1).TransRevAmt(12)
    TotTaxBillsP# = TotTaxBillsP# + CMTrRec(1).TransRevAmt(13)
  End If
Return
PrintTotalTax:
  If TotTaxBillsR# > 0 Or TotPrePayTx# > 0 Then
    If LineCnt + 13 >= MaxLines - 6 Then
      Print #RptHandle, FF$
      Print #RptHandle, " "
      Print #RptHandle, "Totals Page for Operator # "; PrnOpr$
      Print #RptHandle, "Total Real Tax Receipts Cont'd"
      Print #RptHandle, " "
      LineCnt = 4
    End If
    Print #RptHandle, String$(96, "-")
    Print #RptHandle, "GRAND TOTAL Real Tax Receipts Recap"
    LineCnt = LineCnt + 2
    If LineCnt >= MaxLines - 6 Then
      Print #RptHandle, FF$
      Print #RptHandle, " "
      Print #RptHandle, "Totals Page for Operator # "; PrnOpr$
      Print #RptHandle, "Total Real Tax Receipts Cont'd"
      Print #RptHandle, " "
      LineCnt = 4
    End If
    Print #RptHandle, "Total Tax Principle Received ..... "; Tab(36); Using(Fmt3$, TotPrincpleTx#)
    Print #RptHandle, "Total Interest Received .......... "; Tab(36); Using(Fmt3$, TotIntTax#)
    Print #RptHandle, "Total Collection Fees Received ... "; Tab(36); Using(Fmt3$, TotCollTax#)
    Print #RptHandle, "Total Late List Fees Received .... "; Tab(36); Using(Fmt3$, TotLateList#)
    Print #RptHandle, "Total Penalties Received ......... "; Tab(36); Using(Fmt3$, TotPenTax#)
    Print #RptHandle, "Total " + TxOpt1F + " Received .. "; Tab(36); Using(Fmt3$, TotOpt1#)
    Print #RptHandle, "Total " + TxOpt2F + " Received .. "; Tab(36); Using(Fmt3$, TotOpt2#)
    Print #RptHandle, "Total " + TxOpt3F + " Received .. "; Tab(36); Using(Fmt3$, TotOpt3#)
    Print #RptHandle, "Total Discount Given ............. "; Tab(36); Using(Fmt3$, TotDiscTax#)
    Print #RptHandle, "Total PrePayment Received ........ "; Tab(36); Using(Fmt3$, TotPrePayTx#)
    Print #RptHandle, "Total Tax Bills Paid ............. "; Tab(36); CInt(TotTaxBillsR#)
    Print #RptHandle, "GRAND Total Real Tax Received .... "; Tab(36); Using(Fmt3$, ((TotPrincpleTx# + TotIntTax# + TotCollTax# + TotPenTax# + TotLateList# + TotOpt1# + TotOpt2# + TotOpt3# + TotPrePayTx#) - TotDiscTax#))
    Print #RptHandle, ""
    LineCnt = LineCnt + 13
  End If
  If TotTaxBillsP# > 0 Or TotPrePayTxP# > 0 Then
   If LineCnt + 13 >= MaxLines - 6 Then
      Print #RptHandle, FF$
      Print #RptHandle, " "
      Print #RptHandle, "Totals Page for Operator # "; PrnOpr$
      Print #RptHandle, "Total Pers Tax Receipts Cont'd"
      Print #RptHandle, " "
      LineCnt = 4
    End If
    Print #RptHandle, String$(96, "-")
    Print #RptHandle, "GRAND TOTAL Pers Tax Receipts Recap"
    LineCnt = LineCnt + 2
    If LineCnt >= MaxLines - 6 Then
      Print #RptHandle, FF$
      Print #RptHandle, " "
      Print #RptHandle, "Totals Page for Operator # "; PrnOpr$
      Print #RptHandle, "Total Tax Receipts Cont'd"
      Print #RptHandle, " "
      LineCnt = 4
    End If
    Print #RptHandle, "Total Tax Principle1 Received .... "; Tab(36); Using(Fmt3$, TotPrinc1Tx#)
    Print #RptHandle, "Total Tax Principle2 Received .... "; Tab(36); Using(Fmt3$, TotPrinc2Tx#)
    Print #RptHandle, "Total Tax Principle3 Received .... "; Tab(36); Using(Fmt3$, TotPrinc3Tx#)
    Print #RptHandle, "Total Tax Principle4 Received .... "; Tab(36); Using(Fmt3$, TotPrinc4Tx#)
    Print #RptHandle, "Total Tax Principle5 Received .... "; Tab(36); Using(Fmt3$, TotPrinc5Tx#)
    Print #RptHandle, "Total Interest Received .......... "; Tab(36); Using(Fmt3$, TotIntTaxP#)
    Print #RptHandle, "Total Penalties Received ......... "; Tab(36); Using(Fmt3$, TotPenTaxP#)
    Print #RptHandle, "Total " + TxPOpt1F + " Received .. "; Tab(36); Using(Fmt3$, TotOpt1P#)
    Print #RptHandle, "Total " + TxPOpt2F + " Received .. "; Tab(36); Using(Fmt3$, TotOpt2P#)
    Print #RptHandle, "Total " + TxPOpt3F + " Received .. "; Tab(36); Using(Fmt3$, TotOpt3P#)
    Print #RptHandle, "Total Discount Given ............. "; Tab(36); Using(Fmt3$, TotDiscTaxP#)
    Print #RptHandle, "Total PrePayment Received ........ "; Tab(36); Using(Fmt3$, TotPrePayTxP#)
    Print #RptHandle, "Total Tax Bills Paid ............. "; Tab(36); CInt(TotTaxBillsP#)
    Print #RptHandle, "GRAND Total Pers Tax Received .... "; Tab(36); Using(Fmt3$, ((TotPrinc1Tx# + TotPrinc2Tx# + TotPrinc3Tx# + TotPrinc4Tx# + TotPrinc5Tx# + TotIntTaxP# + TotPenTaxP# + TotOpt1P# + TotOpt2P# + TotOpt3P# + TotPrePayTxP#) - TotDiscTaxP#))
    Print #RptHandle, ""
    LineCnt = LineCnt + 13
  End If
  If TotTaxBillsR# > 0 Or TotTaxBillsP# > 0 Or TotPrePayTxP# > 0 Or TotPrePayTx# > 0 Then
    TotAllTax# = ((TotPrincpleTx# + TotIntTax# + TotCollTax# + TotPenTax# + TotLateList# + TotOpt1# + TotOpt2# + TotOpt3# + TotPrePayTx#) - TotDiscTax#)
    TotAllTax# = (TotAllTax# + ((TotPrinc1Tx# + TotPrinc2Tx# + TotPrinc3Tx# + TotPrinc4Tx# + TotPrinc5Tx# + TotIntTaxP# + TotPenTaxP# + TotOpt1P# + TotOpt2P# + TotOpt3P# + TotPrePayTxP#) - TotDiscTaxP#))
    Print #RptHandle, String$(96, "-")
    Print #RptHandle, "GRAND Total Tax Received ......... "; Tab(36); Using(Fmt3$, TotAllTax#)
    Print #RptHandle, ""
    LineCnt = LineCnt + 3
  ElseIf (TotCurTax# + TotCurInt# + TotCurPen# + TotPastTax# + TotPastInt# + TotPastPen# + TotStrmFee# + TotPastStrm#) <> 0 Then
    Print #RptHandle, String$(96, "-")
    Print #RptHandle, "GRAND TOTAL Tax Receipts Recap"
    LineCnt = LineCnt + 2
    If LineCnt + 8 >= MaxLines - 6 Then
      Print #RptHandle, FF$
      Print #RptHandle, " "
      Print #RptHandle, "Totals Page for Operator # "; PrnOpr$
      Print #RptHandle, "Total Tax Receipts Cont'd"
      Print #RptHandle, " "
      LineCnt = 4
    End If
    Print #RptHandle, "Total Current Taxes Received ..... "; Using(Fmt3$, TotCurTax#)
    Print #RptHandle, "Total Current Interest Received .. "; Using(Fmt3$, TotCurInt#)
    Print #RptHandle, "Total Current Penalty Received ... "; Using(Fmt3$, TotCurPen#)
    Print #RptHandle, "Total Storm Fee Received ......... "; Using(Fmt3$, TotStrmFee#)
    Print #RptHandle, "Total Past Taxes Received ........ "; Using(Fmt3$, TotPastTax#)
    Print #RptHandle, "Total Past Interest Received ..... "; Using(Fmt3$, TotPastInt#)
    Print #RptHandle, "Total Past Penalty Received ...... "; Using(Fmt3$, TotPastPen#)
    Print #RptHandle, "Total Past Storm Fee Received .... "; Using(Fmt3$, TotPastStrm#)
    Print #RptHandle, "GRAND Total Tax Received ......... "; Using(Fmt3$, (TotCurTax# + TotCurInt# + TotCurPen# + TotPastTax# + TotPastInt# + TotPastPen# + TotStrmFee# + TotPastStrm#))
    Print #RptHandle, ""
    LineCnt = LineCnt + 8
  ElseIf (TotPrincpleTx# + TotIntTax# + TotCollTax# + TotLateList# + TotOpt1# + TotOpt2# + TotOpt3# + TotDiscTax# + TotPrePayTx# + TotTaxBills#) <> 0 Then
    Print #RptHandle, String$(96, "-")
    Print #RptHandle, "GRAND TOTAL Tax Receipts Recap"
    LineCnt = LineCnt + 2
    If LineCnt + 12 >= MaxLines - 6 Then
      Print #RptHandle, FF$
      Print #RptHandle, " "
      Print #RptHandle, "Totals Page for Operator # "; PrnOpr$
      Print #RptHandle, "Total Tax Receipts Cont'd"
      Print #RptHandle, " "
      LineCnt = 4
    End If
    Print #RptHandle, "Total Tax Principle Received ..... "; Tab(36); Using(Fmt3$, TotPrincpleTx#)
    Print #RptHandle, "Total Interest Received .......... "; Tab(36); Using(Fmt3$, TotIntTax#)
    Print #RptHandle, "Total Collection Fees Received ... "; Tab(36); Using(Fmt3$, TotCollTax#)
    Print #RptHandle, "Total Late List Fees Received .... "; Tab(36); Using(Fmt3$, TotLateList#)
    Print #RptHandle, "Total " + TxOpt1F + " Received .. "; Tab(36); Using(Fmt3$, TotOpt1#)
    Print #RptHandle, "Total " + TxOpt2F + " Received .. "; Tab(36); Using(Fmt3$, TotOpt2#)
    Print #RptHandle, "Total " + TxOpt3F + " Received .. "; Tab(36); Using(Fmt3$, TotOpt3#)
    Print #RptHandle, "Total Discount Given ............. "; Tab(36); Using(Fmt3$, TotDiscTax#)
    Print #RptHandle, "Total PrePayment Received ........ "; Tab(36); Using(Fmt3$, TotPrePayTx#)
    Print #RptHandle, "Total Tax Bills Paid ............. "; Tab(36); CInt(TotTaxBills#)
    Print #RptHandle, "GRAND Total Tax Received ......... "; Tab(36); Using(Fmt3$, ((TotPrincpleTx# + TotIntTax# + TotCollTax# + TotLateList# + TotOpt1# + TotOpt2# + TotOpt3# + TotPrePayTx#) - TotDiscTax#))
    Print #RptHandle, ""
    LineCnt = LineCnt + 12
  End If
 
Return

SubTotalBL:
  If CMTrRec(1).TransAmount <> totPay Then 'CMTrRec(1).TransAmtOwed Then
    TotalBLAmt# = TotalBLAmt# + CMTrRec(1).TransAmount
  Else
    TotalBLAmt# = TotalBLAmt# + totPay 'CMTrRec(1).TransAmtOwed
  End If
  TotalBLAmt# = Round#(TotalBLAmt#)
Return
PrintBLTotal:
  If LineCnt + 3 >= MaxLines - 5 Then
    Print #RptHandle, FF$
    LineCnt = 1
  End If
  Print #RptHandle, String$(96, "-")
  Print #RptHandle, "GRAND TOTAL Business Licence Receipts Recap"
  Print #RptHandle, "GRAND Total Bus. Lic. Receipts .. "; Tab(35); Using("$##,###,###.##", TotalBLAmt#)
  LineCnt = LineCnt + 3
Return

SubTotalDC:
  TotalDCAmt# = TotalDCAmt# + CMTrRec(1).TransAmount 'totPay'CMTrRec(1).TransAmtOwed
  TotalDCAmt# = Round#(TotalDCAmt#)
Return

PrintDCTotal:
  If LineCnt + 3 >= MaxLines - 5 Then
    Print #RptHandle, FF$
    LineCnt = 1
  End If
  Print #RptHandle, String$(96, "-")
  Print #RptHandle, "GRAND TOTAL Vehicle Decals Receipts Recap"
  Print #RptHandle, "GRAND Total Veh. Dec. Receipts .. "; Tab(35); Using("$##,###,###.##", TotalDCAmt#)
  LineCnt = LineCnt + 3
Return

SubTotalMisc:
  If TotalMiscCnt = 0 Then
    TotalMiscCnt = 1
    TotalMiscRec$(1) = MiscCodeRec(1).MiscCode
    TotalMiscDesc$(1) = MiscCodeRec(1).Description
    TotalMiscAmt#(1) = MiscRevAmt#
    MiscCodeGL$(1) = MiscCodeRec(1).GlAcctNumb
  Else
    For TCnt = 1 To TotalMiscCnt
      If MiscCodeRec(1).MiscCode = TotalMiscRec$(TCnt) Then
        TotalMiscAmt#(TCnt) = TotalMiscAmt#(TCnt) + MiscRevAmt#: Return
      End If
    Next TCnt
    TotalMiscCnt = TotalMiscCnt + 1
    TotalMiscRec$(TotalMiscCnt) = MiscCodeRec(1).MiscCode
    TotalMiscDesc$(TotalMiscCnt) = MiscCodeRec(1).Description
    TotalMiscAmt#(TotalMiscCnt) = MiscRevAmt#
    MiscCodeGL$(TotalMiscCnt) = MiscCodeRec(1).GlAcctNumb
  End If
Return

PrintTotalMisc:
  If LineCnt + 12 >= MaxLines - 5 Then
    Print #RptHandle, FF$
    LineCnt = 1
    Print #RptHandle, " "
    Print #RptHandle, " "
    Print #RptHandle, Tab(27); "Cash Receipts Journal : Cash Management System"
    Print #RptHandle, "  Current Date: "; Now
    Print #RptHandle, "Beginning Date: "; FromDate$
    Print #RptHandle, "   Ending Date: "; ThruDate$
    Print #RptHandle, "      Operator: "; PrnOpr$
    Print #RptHandle, " "
    Print #RptHandle, "Totals Page for Operator # "; PrnOpr$
    Print #RptHandle, " "
    Print #RptHandle, String$(96, "-")
    Print #RptHandle, "GRAND TOTAL Misc Receipts Recap"
    LineCnt = LineCnt + 12
  Else
    Print #RptHandle, String$(96, "-")
    Print #RptHandle, "GRAND TOTAL Misc Receipts Recap"
    LineCnt = LineCnt + 2
  End If
  For TCnt = 1 To TotalMiscCnt
    If LineCnt >= MaxLines - 5 Then
      Print #RptHandle, FF$
      Print #RptHandle, ""
      Print #RptHandle, ""
      Print #RptHandle, Tab(27); "Cash Receipts Journal : Cash Management System"
      Print #RptHandle, "  Current Date: "; Now
      Print #RptHandle, "Beginning Date: "; FromDate$
      Print #RptHandle, "   Ending Date: "; ThruDate$
      Print #RptHandle, "Totals Page for Operator # "; PrnOpr$
      Print #RptHandle, "Total Misc Receipts Recap Cont'd"
      Print #RptHandle, ""
      LineCnt = 9
    End If
    Print #RptHandle, TotalMiscDesc$(TCnt);
    Print #RptHandle, Tab(35); Using("$##,###,###.##", TotalMiscAmt#(TCnt));
    Print #RptHandle, Tab(52); "GL# "; MiscCodeGL$(TCnt)
    TotalMisc# = TotalMisc# + TotalMiscAmt#(TCnt)
    LineCnt = LineCnt + 1
  Next TCnt
  Print #RptHandle, "GRAND Total Misc Receipts .... "; Tab(35); Using("$##,###,###.##", TotalMisc#)
  LineCnt = LineCnt + 1
Return

'OpenARFile:
'  Close ARFile
'  ARCustRecLen = Len(ARCustRec(1))
'  ARFile = FreeFile
'  Open "ARCUST.DAT" For Random Access Read Write Shared As ARFile Len = ARCustRecLen
'  NumOfArRecs = LOF(ARFile) \ ARCustRecLen
'  CatFile = FreeFile
'  Open "ARCODE.DAT" For Random As CatFile Len = CatCodeRecLen
'  NumOfCatRecs = LOF(CatFile) \ CatCodeRecLen
'  Return
'
End Sub

