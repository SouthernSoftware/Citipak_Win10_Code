VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmRptTransJournal 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customers Transaction Journal Report"
   ClientHeight    =   8640
   ClientLeft      =   30
   ClientTop       =   540
   ClientWidth     =   12195
   Icon            =   "frmRptTransJournal.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   12195
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboTransType 
      Height          =   375
      Left            =   5370
      TabIndex        =   4
      Top             =   3810
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
      ColDesigner     =   "frmRptTransJournal.frx":08CA
   End
   Begin LpLib.fpCombo fpcboPrintOrder 
      Height          =   375
      Left            =   5370
      TabIndex        =   8
      Top             =   5910
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
      ColDesigner     =   "frmRptTransJournal.frx":0C61
   End
   Begin LpLib.fpCombo fpcboDetail 
      Height          =   375
      Left            =   5370
      TabIndex        =   7
      Top             =   5370
      Width           =   840
      _Version        =   196608
      _ExtentX        =   1482
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
      ColDesigner     =   "frmRptTransJournal.frx":0FF8
   End
   Begin LpLib.fpCombo fpcboRptType 
      Height          =   375
      Left            =   5370
      TabIndex        =   9
      Top             =   6435
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
      ColDesigner     =   "frmRptTransJournal.frx":139A
   End
   Begin VB.CheckBox DelOnly 
      Caption         =   "Deleted Customers Transactions Only"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3720
      TabIndex        =   25
      Top             =   1176
      Width           =   5148
   End
   Begin VB.CheckBox QckSrch 
      BackColor       =   &H008F8265&
      Caption         =   "Quick Search"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5376
      TabIndex        =   10
      Top             =   6960
      Width           =   2052
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
      TabIndex        =   12
      Top             =   7560
      Width           =   1332
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
      TabIndex        =   11
      Top             =   7560
      Width           =   1332
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   13
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
            TextSave        =   "5:09 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7117
            TextSave        =   "5/13/2013"
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
      Top             =   2268
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
      Top             =   1752
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
   Begin EditLib.fpText fptxtRoute2 
      Height          =   348
      Left            =   5376
      TabIndex        =   3
      Top             =   3300
      Width           =   636
      _Version        =   196608
      _ExtentX        =   1122
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
   Begin EditLib.fpText fptxtRoute1 
      Height          =   348
      Left            =   5376
      TabIndex        =   2
      Top             =   2784
      Width           =   636
      _Version        =   196608
      _ExtentX        =   1122
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
   Begin EditLib.fpText fptxtCustType 
      Height          =   348
      Left            =   5376
      TabIndex        =   6
      Top             =   4860
      Width           =   1188
      _Version        =   196608
      _ExtentX        =   2096
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
      MaxLength       =   3
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
   Begin EditLib.fpText fptxtOperator 
      Height          =   348
      Left            =   5376
      TabIndex        =   5
      Top             =   4344
      Width           =   804
      _Version        =   196608
      _ExtentX        =   1418
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
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Type:"
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
      Index           =   1
      Left            =   3216
      TabIndex        =   24
      Top             =   4920
      Width           =   2076
   End
   Begin VB.Label LabelB2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "To Book:"
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
      Left            =   3912
      TabIndex        =   23
      Top             =   3360
      Width           =   1380
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
      TabIndex        =   22
      Top             =   4404
      Width           =   2076
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
      TabIndex        =   21
      Top             =   5964
      Width           =   1716
   End
   Begin VB.Label LabelB1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "From Book:"
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
      Left            =   3816
      TabIndex        =   20
      Top             =   2844
      Width           =   1476
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   5820
      Left            =   2328
      Top             =   1632
      Width           =   7284
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Show Detail: "
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
      Left            =   3336
      TabIndex        =   19
      Top             =   5436
      Width           =   2004
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Transaction Type:"
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
      TabIndex        =   18
      Top             =   3876
      Width           =   2244
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
      TabIndex        =   17
      Top             =   6480
      Width           =   2388
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
      TabIndex        =   16
      Top             =   1800
      Width           =   1668
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
      TabIndex        =   15
      Top             =   2316
      Width           =   1572
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   708
      Left            =   2718
      Top             =   312
      Width           =   6756
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Print Customers Transactions"
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
      Left            =   2766
      TabIndex        =   14
      Top             =   480
      Width           =   6612
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   828
      Left            =   2718
      Top             =   192
      Width           =   6756
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
Attribute VB_Name = "frmRptTransJournal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim BegRoute As String, EndRoute As String
Dim UseCycle As Boolean
Private Sub cmdExit_Click()
  frmUBEditMenu.Show
  Unload frmRptTransJournal
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        UBLog "Closed via RptTransJournal by " + "Util OPer"
        'CitiTerminate
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
    fptxtRoute1.SetFocus
  End If
End Sub
Private Sub fptxtRoute1_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fptxtRoute1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fptxtRoute2.SetFocus
  End If
End Sub
Private Sub fptxtRoute2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpcboTransType.SetFocus
  End If
End Sub

Private Sub fptxtRoute2_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub
Private Sub fpcboTransType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboTransType.ListDown = True
  End If
  If fpcboTransType.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fptxtOperator.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fptxtRoute2.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub

Private Sub fptxtOperator_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fptxtCustType.SetFocus
  End If
End Sub

Private Sub fptxtCustType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpcboDetail.SetFocus
  End If
End Sub
Private Sub fpcboDetail_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboDetail.ListDown = True
  End If
  If fpcboDetail.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpcboPrintOrder.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fptxtCustType.SetFocus
        KeyCode = 0
      End If
    End If
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
        fpcboDetail.SetFocus
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
Private Function ValidRoutes()
  If fptxtRoute1 <> "" And fptxtRoute2 <> "" Then
    If fptxtRoute1 > fptxtRoute2 Then
      MsgBox "Invalid Selection, The Beginning Value Should Be Less or Equal to Ending Value.", vbOKOnly, "Invalid Selection"
      ValidRoutes = False
    Else
      ValidRoutes = True
      BegRoute = QPTrim(fptxtRoute1)
      EndRoute = QPTrim(fptxtRoute2)
    End If
  Else
    MsgBox "Fields May Not Be Left Blank.", vbOKOnly, "Invalid Selection"
  End If
End Function


Private Sub cmdPrint_Click()
  If ValidDate = True Then
   If ValidRoutes Then
    DeActivateControls Me, True
    If fpcboTransType.ListIndex = 0 Then
      DetailedTransJournal3
      ActivateControls Me, True
     Else
    If fpcboRptType.ListIndex = 2 Then
      DetailedTransJournal
      ActivateControls Me, True
    ElseIf fpcboRptType.ListIndex = 1 Or fpcboRptType.ListIndex = 0 Then
      DetailedTransJournal2
    Else
      ActivateControls Me, True
    End If
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

Private Sub Form_Load()
'  Dim UBSetupreclen As Integer
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
  'Me.HelpContextID = hlpTransactionJournal
'  ReDim ubsetup(1) As UBSetupRecType
'  UBSetupreclen = Len(ubsetup(1))
 ' LoadUBSetUpFile ubsetup(), UBSetupreclen
''Do Not allow use cycle for transaction report
 ' If ubsetup(1).BILLCYCL = "Y" Then
'    UseCycle = True
'  End If
'  Erase ubsetup
'  If UseCycle Then
'    LabelB1.Caption = "From Cycle:"
'    LabelB2.Caption = "Thru Cycle:"
'  Else
    LabelB1.Caption = "From Book:"
    LabelB2.Caption = "Thru Book:"
'  End If
  fptxtRoute1 = "00"
  fptxtRoute2 = "99"
  fptxtOperator = ""
  fptxtCustType = ""
  fpcboDetail.AddItem "No"
  fpcboDetail.AddItem "Yes"
  fpcboDetail.ListIndex = 0
  fpcboPrintOrder.AddItem "Customer Name Order"
  fpcboPrintOrder.AddItem "Account Number Order"
  fpcboPrintOrder.AddItem "Location Number Order"
  fpcboPrintOrder.AddItem "Service Address Order"
  fpcboPrintOrder.ListIndex = 0
  txtDate1.Text = Format(Now, "mm/dd/yyyy")
  txtDate2.Text = Format(Now, "mm/dd/yyyy")
  fpcboTransType.AddItem " 0) - All"
  fpcboTransType.AddItem " 1) - Utility Bill"
  fpcboTransType.AddItem " 4) - Payment"
  fpcboTransType.AddItem " 5) - Applied Deposit"
  fpcboTransType.AddItem " 6) - Penalty Charge"
  fpcboTransType.AddItem " 7) - Deposit Payment"
  fpcboTransType.AddItem " 9) - Refunded Deposit"
  fpcboTransType.AddItem "11) - Up Adjustment"
  fpcboTransType.AddItem "12) - Down Adjustment"
  fpcboTransType.AddItem "33) - Payment Adjustment"
  fpcboTransType.AddItem "37) - Deposit Credit Removal"
  fpcboTransType.AddItem "39) - Deposit Payment Void"
  fpcboTransType.ListIndex = 0
  fpcboRptType.InsertRow = "Graphics - Landscape"
  fpcboRptType.InsertRow = "Graphics - Portrait"
  fpcboRptType.InsertRow = "Text - Condensed Print"
  fpcboRptType.ListIndex = 0
  QckSrch.Value = 1
  DelOnly.Value = 0
End Sub

Private Sub Form_Resize()
'  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
'  End If
End Sub
Private Sub DetailedTransJournal()
  Dim UBCustRecLen As Integer, UBSetupreclen As Integer
  Dim UsingBook As Boolean, UsingName As Boolean, UsingAddr As Boolean
  Dim CustomerCnt As Long, UBTransRecLen As Integer, UBTrans As Integer
  Dim IndexName As String, Handle As Integer, Dash120 As String
  Dim IdxRecLen As Integer, IdxFileSize As Long, UBRpt As Integer
  Dim IdxNumOfRecs As Long, NumOfRecs As Long, MaxRevenue As Integer
  Dim cnt As Long, UBCust As Integer, RCnt As Integer, UseType As Boolean
  Dim ThisType As String, CUSTTYPE As String, CustBook As Integer
  Dim FromBook As Integer, ThruBook As Integer, BadCount As Long
  Dim Trans As Long, UBTransLen As Integer, BegDate As Integer
  Dim EndDate As Integer, BegOperator As Integer, EndOperator As Integer
  Dim BegTrans As Integer, EndTrans As Integer, TransDesc As String
  Dim Amount As Double, TotalTrans As Double, TransCnt As Long
  Dim Detail As String, Date1 As String, Date2 As String, Operator As String
  Dim TotalRevsAmt As Double, EstCnt As Integer, TrType As String
  Dim TrTyp As Integer, OperatorNo As String, UsingAcct As Boolean
  Dim ReportFile As String, MoFlag As Boolean, Head As String
 'get report parameters
  GoSub CheckDetailParms
  MaxLines = 55
  PageNo = 0
  Dash120$ = String$(121, "-")
  FrmShowPctComp.Label1 = "Creating Transaction Journal"
  FrmShowPctComp.Show , Me
  DoEvents
  ''DeActivateControls Me, True
  ReDim RevTotals(1 To 15) As Double
  ReDim RevenueName(1 To 15) As String
  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))

  ReDim UBSetUp(1) As UBSetupRecType
  UBSetupreclen = Len(UBSetUp(1))

'  IF INSTR(UBSetUp(1).UTILNAME, "AUTRY") > 0 THEN
'    LptPort = 2
'  ELSE
'    LptPort = 1
'  END IF
'  ReDim UBTrans(1) As UBTransRecType
'  UBTransRecLen = Len(UBTrans(1))
  ReDim UBTransRec(1) As UBTransRecType
  UBTransRecLen = Len(UBTransRec(1))

  If UsingName Or UsingBook Then
    IdxRecLen = 4               'we are using a long integer
    IdxFileSize& = FileSize(IndexName$)
    IdxNumOfRecs = IdxFileSize& \ IdxRecLen
    ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
    'FGetAH IndexName$, IdxBuff(1), IdxRecLen, IdxNumOfRecs      'load it
    NumOfRecs = IdxNumOfRecs
    Handle = FreeFile
    Open IndexName$ For Random Shared As Handle Len = IdxRecLen
    For cnt& = 1 To IdxNumOfRecs
      Get #Handle, cnt&, IdxBuff(cnt&)
    Next
    Close Handle
  ElseIf UsingAddr Then
'unrem
    SortServiceAddrs frmRptTransJournal
    IdxRecLen = 4               'we are using a long integer
    IdxFileSize& = FileSize&(IndexName$)
    IdxNumOfRecs = IdxFileSize& \ IdxRecLen
    ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
    'FGetAH IndexName$, IdxBuff(1), IdxRecLen, IdxNumOfRecs      'load it
    NumOfRecs = IdxNumOfRecs
    Handle = FreeFile
    Open IndexName$ For Random Shared As Handle Len = IdxRecLen
    For cnt& = 1 To IdxNumOfRecs
      Get #Handle, cnt&, IdxBuff(cnt&)
    Next
    Close Handle

  Else
    NumOfRecs = FileSize(UBPath$ + "UBCUST.DAT") \ UBCustRecLen
  End If

  UBCust = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen
  UBTrans = FreeFile
  Open UBPath$ + "UBTRANS.DAT" For Random Shared As UBTrans Len = UBTransRecLen
  ReportFile$ = UBPath$ + "UBDJLIST.RPT"
  UBRpt = FreeFile
  Open ReportFile$ For Output As UBRpt
'  ubsetup = FreeFile
'  Open "UBSETUP.DAT" For Random Shared As ubsetup Len = UBSetupreclen
  LoadUBSetUpFile UBSetUp(), UBSetupreclen

  If Len(TOWNNAME$) = 0 Then
    TOWNNAME$ = "Undefined"
    ' Set Revenue Names to Nothing
    For RCnt = 1 To 15
      RevenueName$(RCnt) = "Not Set"
    Next RCnt
  Else
    'Get ubsetup, 1, ubsetup(1)
    For RCnt = 1 To 15
      RevenueName$(RCnt) = QPTrim$(UBSetUp(1).Revenues(RCnt).RevName)
    Next RCnt
    RCnt = 1
    Do While RCnt <= 15
      If RevenueName$(RCnt) = "" Then
        MaxRevenue = RCnt - 1
        Exit Do
      End If
      RCnt = RCnt + 1
    Loop
'    TownName$ = ubsetup(1).UTILNAME
'    TownLen = Len(RTrim$(TownName$))
'    TabStop = 40 - (TownLen / 2)
'    If TabStop < 1 Then TabStop = 1
  End If
  'Close ubsetup

  'Special Code just for ellenboro!!
'  If InStr(TownName$, "ELLENBO") > 0 Then
'    EllenFlag = True
'  End If
'  If InStr(TOWNNAME$, "MOORE") > 0 Or InStr(TOWNNAME$, "JOHNSTON") > 0 Then
'    MoFlag = True
'  End If
  If QckSrch.Value = 1 Then
    MoFlag = True
  Else
    MoFlag = False
  End If
'  BlockClear
'  ShowProcessingScrn "Detailed Journal Report."

  GoSub DoDetailedRptHeader

  For cnt = 1 To NumOfRecs
    FrmShowPctComp.ShowPctComp cnt, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ''ActivateControls Me, True
      GoTo ExitDetailedListing
    End If

    If UsingName Or UsingBook Or UsingAddr Then
      Get UBCust, IdxBuff(cnt).RecNum, UBCustRec(1)
    Else
      Get UBCust, cnt, UBCustRec(1)
    End If
    If DelOnly.Value = 1 Then
  'only list trans for deleted customers *)((*&(*&(*&(*&(&
      If UBCustRec(1).DelFlag = 0 Then
        GoTo SkipThisOne
      End If
    Else
      If UBCustRec(1).DelFlag <> 0 Then
        GoTo SkipThisOne
      End If
    End If
    If UseType Then
      ThisType$ = QPTrim$(UBCustRec(1).CUSTTYPE)
      If ThisType$ <> CUSTTYPE$ Then
        GoTo SkipThisOne
      End If
    End If
'    If UseCycle Then
'      CustBook = UBCustRec(1).BILLCYCL
'    Else
      CustBook = Val(UBCustRec(1).Book)
'    End If
    If CustBook < FromBook Or CustBook > ThruBook Then
      GoTo SkipThisOne
    End If

    If LineCnt > MaxLines Then
      Print #UBRpt, FF$
      GoSub DoDetailedRptHeader
    End If
'*************************************
'   Main Body of Printing goes here
    BadCount = 0
    Trans& = UBCustRec(1).LastTrans
    Do While Trans& <> 0
      Get UBTrans, Trans&, UBTransRec(1)
      'If Not EllenFlag Then
'        If UBTransRec(1).TransDate < BegDate Then
'          BadCount = BadCount + 1
'          If BadCount > 3 Then
'            Exit Do
'          End If
'        End If
      'End If
      If MoFlag Then
        If UBTransRec(1).TransDate < BegDate Then
          BadCount = BadCount + 1
          If BadCount > 3 Then
            Exit Do
          End If
        End If
      End If
      
      'Check Date, Operator and Trans Type
      
      If (UBTransRec(1).TransDate >= BegDate And UBTransRec(1).TransDate <= EndDate) Then
        If (UBTransRec(1).OperatorNumber >= BegOperator And UBTransRec(1).OperatorNumber <= EndOperator) Then
          If (UBTransRec(1).TransType >= BegTrans And UBTransRec(1).TransType <= EndTrans Or (UBTransRec(1).TransType >= BegTrans + 100 And UBTransRec(1).TransType <= EndTrans + 100)) Then
            GoSub DefineType
            Print #UBRpt, Num2Date$(UBTransRec(1).TransDate); Tab(11); Using("#####", UBTransRec(1).CustAcctNo);
            'PRINT #UBRpt, Num2Date$(UBTransRec(1).TransDate); TAB(11); ASC(UBTransRec(1).Posted2GL); 'USING "#####"; UBTrans(1).CustAcctNo;
            Print #UBRpt, Tab(20); Left$(UBCustRec(1).CustName, 33);
            Print #UBRpt, Tab(55); TransDesc$;
            'PRINT #UBRpt, TAB(55); Trans&;
            Print #UBRpt, Tab(80); UBTransRec(1).OperatorNumber;
            'PRINT #UBRpt, TAB(80); "!"; UBTransRec(1).Posted2GL; "!";
            Print #UBRpt, Tab(90); Left$(UBTransRec(1).TransDesc, 20);
            Print #UBRpt, Tab(110); Using("$###,###.##", Amount#)
            'PRINT #UBRpt, "  "; "!"; UBTransRec(1).Posted2GL; "!"
            LineCnt = LineCnt + 1
            TotalTrans# = Round#(TotalTrans# + Amount#)
            TransCnt& = TransCnt& + 1
            If Detail$ = "Y" Then
              Print #UBRpt, "Revenue Source Breakdown ........................"
              LineCnt = LineCnt + 1
              For RCnt = 1 To MaxRevenue Step 3
                Print #UBRpt, RevenueName$(RCnt); Tab(16); Using("#####.##", Round#(UBTransRec(1).RevAmt(RCnt) + UBTransRec(1).TaxAmt(RCnt)));
                Print #UBRpt, Tab(30); RevenueName$(RCnt + 1); Tab(46); Using("#####.##", Round#(UBTransRec(1).RevAmt(RCnt + 1) + UBTransRec(1).TaxAmt(RCnt + 1)));
                Print #UBRpt, Tab(60); RevenueName$(RCnt + 2); Tab(76); Using("#####.##", Round#(UBTransRec(1).RevAmt(RCnt + 2) + UBTransRec(1).TaxAmt(RCnt + 2)))
                LineCnt = LineCnt + 1
              Next RCnt
              'IF UBTransRec(1).TransType = TranUpwardAdjustment OR UBTransRec(1).TransType = TranDownwardAdjustment THEN
              '  FOR RCnt = 1 TO 7
              '    PRINT #UBRpt, RevenueName$(RCnt); TAB(16); USING "#####.##"; UBTransRec(1).RevAmt(RCnt);
              '  PRINT #UBRpt, TAB(30); RevenueName$(RCnt + 1); TAB(46); USING "#####.##"; UBTransRec(1).RevAmt(RCnt + 1);
              '  PRINT #UBRpt, TAB(60); RevenueName$(RCnt + 2); TAB(76); USING "#####.##"; UBTransRec(1).RevAmt(RCnt + 2)
              '  LineCnt = LineCnt + 1
              'NEXT RCnt
              Print #UBRpt, Dash120$
              LineCnt = LineCnt + 1
            End If
            For RCnt = 1 To MaxRevenue
              RevTotals(RCnt) = Round#(RevTotals(RCnt) + (UBTransRec(1).RevAmt(RCnt) + UBTransRec(1).TaxAmt(RCnt)))
            Next
            If LineCnt > MaxLines Then
              Print #UBRpt, FF$
              GoSub DoDetailedRptHeader
            End If
          End If
        End If
      End If
'      If AskAbandonPrint% Then
'        AbortFlag = True
'        Exit For
'      End If
      Trans& = UBTransRec(1).PrevTrans
    Loop
SkipThisOne:
'    ShowPctComp cnt, NumOfRecs
  Next

  GoSub DoDetailedRptFooter
  Print #UBRpt, FF$;

  Close

  Erase IdxBuff, UBCustRec
  ViewPrint ReportFile$, Head$, True

ExitDetailedListing:

  Exit Sub

DoDetailedRptHeader:
  PageNo = PageNo + 1
  Print #UBRpt, TOWNNAME$
  Print #UBRpt, Tab(28); Head$; Tab(90); "Page #"; PageNo
  Print #UBRpt, "Report Date: "; Date$
  Print #UBRpt, "Beginning Transaction Date: "; Date1$;
  If Val(Operator$) = 0 Then
    Print #UBRpt, Tab(90); " Operator #: ALL"
  Else
    Print #UBRpt, Tab(90); " Operator #: "; Operator$
  End If
  Print #UBRpt, "   Ending Transaction Date: "; Date2$;
  Print #UBRpt, Tab(90); "Show Detail: "; Detail$
  Print #UBRpt, "          Transaction Type: "; fpcboTransType.Text
  Print #UBRpt, "             Customer Type: ";

  If UseType Then
    Print #UBRpt, CUSTTYPE$
  Else
    Print #UBRpt, "N/A"
  End If

  Print #UBRpt,
  Print #UBRpt, "  Date"; Tab(11); "Acct #"; Tab(20); "Customer Name"; Tab(55); "Description"; Tab(80); "Oper#"; Tab(90); "Trans Type"; Tab(113); "   Amount"
  Print #UBRpt, Dash120$
  LineCnt = 10
  Return

DoDetailedRptFooter:
  Print #UBRpt, Dash120$
  Print #UBRpt, "Transactions: "; TransCnt&; "                                                       Total of Transactions: "; Using("$##,###,###.##", TotalTrans#)
  Print #UBRpt, FF$
  PageNo = PageNo + 1
  Print #UBRpt, TOWNNAME$
  Print #UBRpt, Tab(28); Head$; Tab(90); "Page #"; PageNo
  Print #UBRpt, "Report Date: "; Date$
  Print #UBRpt, "Beginning Transaction Date: "; Date1$;
  If Val(Operator$) = 0 Then
    Print #UBRpt, Tab(90); " Operator #: ALL"
  Else
    Print #UBRpt, Tab(90); " Operator #: "; Operator$
  End If
  Print #UBRpt, "   Ending Transaction Date: "; Date2$;
  Print #UBRpt, Tab(90); "Show Detail: "; Detail$
  Print #UBRpt, ""
  Print #UBRpt, "Revenue Summary"; Tab(38); "Amount"
  Print #UBRpt, Dash120$
  TotalRevsAmt# = 0
  For RCnt = 1 To MaxRevenue
    TotalRevsAmt# = Round#(TotalRevsAmt# + RevTotals(RCnt))
    Print #UBRpt, RevenueName$(RCnt), Tab(35); Using("########.##", RevTotals(RCnt))
  Next
  Print #UBRpt,
  Print #UBRpt, "Total Amount"; Tab(35); Using("########.##", TotalRevsAmt#)
  Return
DefineType:
  Select Case UBTransRec(1).TransType
  Case 1, 101
    TransDesc$ = "Util Bill"
    'EstFlag = False
    For EstCnt = 1 To 7
      If UBTransRec(1).ESTREAD(EstCnt) = "Y" Then
        'EstFlag = True
        TransDesc$ = TransDesc$ + "*E"
        Exit For
      End If
    Next
    Amount# = UBTransRec(1).Transamt
  Case 2, 102
    TransDesc$ = "Late Charge"
    Amount# = UBTransRec(1).Transamt
  Case 3
    TransDesc$ = "Reconnect"
    Amount# = UBTransRec(1).Transamt
  Case 4, 104
    TransDesc$ = "Reg Payment"
        Amount# = UBTransRec(1).Transamt
  Case 5, 105
    TransDesc$ = "Applied Dep"
    'Amount# = -UBTransRec(1).TransAmt
    Amount# = Abs(UBTransRec(1).Transamt)
  Case 6
    TransDesc$ = "Penalty Chg"
    Amount# = UBTransRec(1).Transamt
  Case 7, 107
    TransDesc$ = "Dep. Payment"
    Amount# = UBTransRec(1).Transamt
  Case 8
    TransDesc$ = "Draft Paymt"
    Amount# = UBTransRec(1).Transamt * -1
  Case 9, 109
    TransDesc$ = "Refunded Dep"
    Amount# = Abs(UBTransRec(1).Transamt)
  Case 10, 110
    TransDesc$ = "Beg Balance"
    Amount# = UBTransRec(1).Transamt
  Case 11, 111
    TransDesc$ = QPTrim$(UBTransRec(1).BillMsg)
    Amount# = UBTransRec(1).Transamt
  Case 12, 112
    TransDesc$ = QPTrim$(UBTransRec(1).BillMsg)
    Amount# = UBTransRec(1).Transamt
  Case 33
    TransDesc$ = QPTrim$(UBTransRec(1).BillMsg)
    Amount# = UBTransRec(1).Transamt
  Case 37
    TransDesc$ = QPTrim$(UBTransRec(1).BillMsg)
    Amount# = UBTransRec(1).Transamt
  Case 39
    TransDesc$ = QPTrim$(UBTransRec(1).BillMsg)
    Amount# = UBTransRec(1).Transamt
  Case 99
    TransDesc$ = "Misc Payment"
    Amount# = UBTransRec(1).Transamt
  Case Else
    TransDesc$ = "UNKNOWN"
    Amount# = UBTransRec(1).Transamt
  End Select
  Return

CheckDetailParms:

  Date1$ = txtDate1
  Date2$ = txtDate2

  BegDate = Date2Num%(Date1$)
  EndDate = Date2Num%(Date2$)

  FromBook = Val(BegRoute)
  ThruBook = Val(EndRoute)
  If fpcboTransType.ListIndex <> -1 Then
    TrType$ = QPTrim$(Left$(fpcboTransType.Text, 2))
    TrTyp = Val(TrType$)
  Else
    MsgBox "Invalid Transaction Type.", vbOKOnly, "Invalid Selection"
    fpcboTransType.SetFocus
    GoSub ExitDetailedListing
  End If
'this trtyp of 0 would only work if allowed all
'which we do not allow on transaction type - maybe in administrative section
  If TrTyp = 0 Then
    BegTrans = 1
    EndTrans = 999
  Else
    BegTrans = TrTyp
    EndTrans = TrTyp
  End If

  OperatorNo$ = fptxtOperator
  Operator = Val(OperatorNo$)
  If Operator = 0 Then
    BegOperator = 0
    EndOperator = 9999
  Else
    BegOperator = Operator
    EndOperator = Operator
  End If
  If DelOnly.Value = 1 Then
    Head$ = "Deleted Customers Transaction Journal"
  Else
    Head$ = "Transaction Journal"
  End If
  Detail$ = QPTrim$(Left$(fpcboDetail.Text, 1))

  CUSTTYPE$ = QPTrim$(fptxtCustType)
  If Len(CUSTTYPE$) > 0 Then
    UseType = True
  End If

  Select Case Left$(fpcboPrintOrder.Text, 1)
    Case "C"
    IndexName$ = NameIndexFile
    UsingName = True
  Case "A"
    IndexName$ = ""
    UsingAcct = True
  Case "L"
    IndexName$ = BookIndexFile
    UsingBook = True
  Case "S"
    IndexName$ = TempIndexName
    UsingAddr = True
  Case Else
  End Select
Return
End Sub
Private Sub DetailedTransJournal2()
  Dim UBCustRecLen As Integer, UBSetupreclen As Integer
  Dim UsingBook As Boolean, UsingName As Boolean, UsingAddr As Boolean
  Dim CustomerCnt As Long, UBTransRecLen As Integer, UBTrans As Integer
  Dim IndexName As String, Handle As Integer, Dash120 As String
  Dim IdxRecLen As Integer, IdxFileSize As Long, UBRpt As Integer
  Dim IdxNumOfRecs As Long, NumOfRecs As Long, MaxRevenue As Integer
  Dim cnt As Long, UBCust As Integer, RCnt As Integer, UseType As Boolean
  Dim ThisType As String, CUSTTYPE As String, CustBook As Integer
  Dim FromBook As Integer, ThruBook As Integer, BadCount As Long
  Dim Trans As Long, UBTransLen As Integer, BegDate As Integer
  Dim EndDate As Integer, BegOperator As Integer, EndOperator As Integer
  Dim BegTrans As Integer, EndTrans As Integer, TransDesc As String
  Dim Amount As Double, TotalTrans As Double, TransCnt As Long
  Dim Detail As String, Date1 As String, Date2 As String, Operator As String
  Dim TotalRevsAmt As Double, EstCnt As Integer, TrType As String
  Dim TrTyp As Integer, OperatorNo As String, UsingAcct As Boolean
  Dim ToPrint As String, PrnH1 As String, PrnH2 As String, PrnH3 As String
  Dim SumRpt As Integer, ToPrintD As String, DetFlag As Boolean, Head$
  Dim ReportFile As String, ReportSum As String, MoFlag As Boolean
 'get report parameters
  GoSub CheckDetailParms
  If fpcboDetail.ListIndex = 1 Then
    DetFlag = True
  Else
    DetFlag = False
  End If
  FrmShowPctComp.Label1 = "Creating Transaction Journal"
  FrmShowPctComp.Show , Me
  DoEvents
  ''DeActivateControls Me, True
  ReDim RevTotals(1 To 15) As Double
  ReDim RevenueName(1 To 15) As String
  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))

  ReDim UBSetUp(1) As UBSetupRecType
  UBSetupreclen = Len(UBSetUp(1))

'  IF INSTR(UBSetUp(1).UTILNAME, "AUTRY") > 0 THEN
'    LptPort = 2
'  ELSE
'    LptPort = 1
'  END IF
'  ReDim UBTrans(1) As UBTransRecType
'  UBTransRecLen = Len(UBTrans(1))
  ReDim UBTransRec(1) As UBTransRecType
  UBTransRecLen = Len(UBTransRec(1))

  If UsingName Or UsingBook Then
    IdxRecLen = 4               'we are using a long integer
    IdxFileSize& = FileSize(IndexName$)
    IdxNumOfRecs = IdxFileSize& \ IdxRecLen
    ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
    'FGetAH IndexName$, IdxBuff(1), IdxRecLen, IdxNumOfRecs      'load it
    NumOfRecs = IdxNumOfRecs
    Handle = FreeFile
    Open IndexName$ For Random Shared As Handle Len = IdxRecLen
    For cnt& = 1 To IdxNumOfRecs
      Get #Handle, cnt&, IdxBuff(cnt&)
    Next
    Close Handle
  ElseIf UsingAddr Then
'unrem
    SortServiceAddrs frmRptTransJournal
    IdxRecLen = 4               'we are using a long integer
    IdxFileSize& = FileSize&(IndexName$)
    IdxNumOfRecs = IdxFileSize& \ IdxRecLen
    ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
    'FGetAH IndexName$, IdxBuff(1), IdxRecLen, IdxNumOfRecs      'load it
    NumOfRecs = IdxNumOfRecs
    Handle = FreeFile
    Open IndexName$ For Random Shared As Handle Len = IdxRecLen
    For cnt& = 1 To IdxNumOfRecs
      Get #Handle, cnt&, IdxBuff(cnt&)
    Next
    Close Handle

  Else
    NumOfRecs = FileSize(UBPath$ + "UBCUST.DAT") \ UBCustRecLen
  End If

  UBCust = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen
  UBTrans = FreeFile
  Open UBPath$ + "UBTRANS.DAT" For Random Shared As UBTrans Len = UBTransRecLen
  ReportFile$ = UBPath$ + "UBDJLIST.RPT"
  UBRpt = FreeFile
  Open ReportFile$ For Output As UBRpt
  ReportSum$ = UBPath$ + "UBDJSUM.RPT"
  SumRpt = FreeFile
  Open ReportSum$ For Output As SumRpt
'  ubsetup = FreeFile
'  Open "UBSETUP.DAT" For Random Shared As ubsetup Len = UBSetupreclen
  LoadUBSetUpFile UBSetUp(), UBSetupreclen

  If Len(TOWNNAME$) = 0 Then
    TOWNNAME$ = "Undefined"
    ' Set Revenue Names to Nothing
    For RCnt = 1 To 15
      RevenueName$(RCnt) = "Not Set"
    Next RCnt
  Else
    'Get ubsetup, 1, ubsetup(1)
    For RCnt = 1 To 15
      RevenueName$(RCnt) = QPTrim$(UBSetUp(1).Revenues(RCnt).RevName)
    Next RCnt
    RCnt = 1
    Do While RCnt <= 15
      If RevenueName$(RCnt) = "" Then
        MaxRevenue = RCnt - 1
        Exit Do
      End If
      RCnt = RCnt + 1
    Loop
'    TownName$ = ubsetup(1).UTILNAME
'    TownLen = Len(RTrim$(TownName$))
'    TabStop = 40 - (TownLen / 2)
'    If TabStop < 1 Then TabStop = 1
  End If
  'Close ubsetup

  'Special Code just for ellenboro!!
'  If InStr(TownName$, "ELLENBO") > 0 Then
'    EllenFlag = True
'  End If
'  If InStr(TOWNNAME$, "MOORE") > 0 Or InStr(TOWNNAME$, "JOHNSTON") > 0 Then
'    MoFlag = True
'  End If
  If QckSrch.Value = 1 Then
    MoFlag = True
  Else
    MoFlag = False
  End If
'  BlockClear
'  ShowProcessingScrn "Detailed Journal Report."

'  GoSub DoDetailedRptHeader

  For cnt = 1 To NumOfRecs
    FrmShowPctComp.ShowPctComp cnt, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ''ActivateControls Me, True
      ActivateControls Me, True
      GoTo ExitDetailedListing
    End If

    If UsingName Or UsingBook Or UsingAddr Then
      Get UBCust, IdxBuff(cnt).RecNum, UBCustRec(1)
    Else
      Get UBCust, cnt, UBCustRec(1)
    End If
   If DelOnly.Value = 1 Then
  'only list trans for deleted customers *)((*&(*&(*&(*&(&
      If UBCustRec(1).DelFlag = 0 Then
        GoTo SkipThisOne
      End If
    Else
      If UBCustRec(1).DelFlag <> 0 Then
        GoTo SkipThisOne
      End If
    End If
    If UseType Then
      ThisType$ = QPTrim$(UBCustRec(1).CUSTTYPE)
      If ThisType$ <> CUSTTYPE$ Then
        GoTo SkipThisOne
      End If
    End If

    CustBook = Val(UBCustRec(1).Book)
    If CustBook < FromBook Or CustBook > ThruBook Then
      GoTo SkipThisOne
    End If

'    If Linecnt > MaxLines Then
'      Print #UBRpt, FF$
'      GoSub DoDetailedRptHeader
'    End If
'*************************************
'   Main Body of Printing goes here
    BadCount = 0
    Trans& = UBCustRec(1).LastTrans
    Do While Trans& <> 0
      Get UBTrans, Trans&, UBTransRec(1)
      'If Not EllenFlag Then
'        If UBTransRec(1).TransDate < BegDate Then
'          BadCount = BadCount + 1
'          If BadCount > 3 Then
'            Exit Do
'          End If
'        End If
      'End If
      If MoFlag Then
        If UBTransRec(1).TransDate < BegDate Then
          BadCount = BadCount + 1
          If BadCount > 3 Then
            Exit Do
          End If
        End If
      End If
      
      'Check Date, Operator and Trans Type
      If (UBTransRec(1).TransDate >= BegDate And UBTransRec(1).TransDate <= EndDate) Then
        If (UBTransRec(1).OperatorNumber >= BegOperator And UBTransRec(1).OperatorNumber <= EndOperator) Then
          If (UBTransRec(1).TransType >= BegTrans And UBTransRec(1).TransType <= EndTrans Or (UBTransRec(1).TransType >= BegTrans + 100 And UBTransRec(1).TransType <= EndTrans + 100)) Then
            GoSub DefineType
            ToPrint$ = Str$(Trans&) + "~" + Num2Date$(UBTransRec(1).TransDate) + "~" + Using("#####", UBTransRec(1).CustAcctNo)
            'PRINT #UBRpt, Num2Date$(UBTransRec(1).TransDate); TAB(11); ASC(UBTransRec(1).Posted2GL); 'USING "#####"; UBTrans(1).CustAcctNo;
            ToPrint$ = ToPrint$ + "~" + Left$(UBCustRec(1).CustName, 33)
            ToPrint$ = ToPrint$ + "~" + TransDesc$
            'PRINT #UBRpt, TAB(55); Trans&;
            ToPrint$ = ToPrint$ + "~" + Str$(UBTransRec(1).OperatorNumber)
            'PRINT #UBRpt, TAB(80); "!"; UBTransRec(1).Posted2GL; "!";
            ToPrint$ = ToPrint$ + "~" + Left$(UBTransRec(1).TransDesc, 20)
            ToPrint$ = ToPrint$ + "~" + Using("$###,###.##", Amount#)
            'PRINT #UBRpt, "  "; "!"; UBTransRec(1).Posted2GL; "!"
            'Linecnt = Linecnt + 1
            TotalTrans# = Round#(TotalTrans# + Amount#)
            TransCnt& = TransCnt& + 1
            If Detail$ = "Y" Then
             ' Print #UBRpt, "Revenue Source Breakdown ........................"
             ' Linecnt = Linecnt + 1
              For RCnt = 1 To 15 'MaxRevenue 'Step 3
                If UBTransRec(1).RevAmt(RCnt) <> 0 Then
                  ToPrintD$ = ToPrintD$ + RevenueName$(RCnt) + "~" + Str$(Round#(UBTransRec(1).RevAmt(RCnt) + UBTransRec(1).TaxAmt(RCnt))) + "~"
'                Print #UBRpt, Tab(30); RevenueName$(RCnt + 1); Tab(46); Using("#####.##", UBTransRec(1).RevAmt(RCnt + 1));
'                Print #UBRpt, Tab(60); RevenueName$(RCnt + 2); Tab(76); Using("#####.##", UBTransRec(1).RevAmt(RCnt + 2))
              '  Linecnt = Linecnt + 1
                Else
                  If Len(RevenueName$(RCnt)) > 0 Then
                    ToPrintD$ = ToPrintD$ + RevenueName$(RCnt) + "~0.00~"
                  Else
                    ToPrintD$ = ToPrintD$ + " ~  ~"
                  End If
                End If
              Next RCnt
              'IF UBTransRec(1).TransType = TranUpwardAdjustment OR UBTransRec(1).TransType = TranDownwardAdjustment THEN
              '  FOR RCnt = 1 TO 7
              '    PRINT #UBRpt, RevenueName$(RCnt); TAB(16); USING "#####.##"; UBTransRec(1).RevAmt(RCnt);
              '  PRINT #UBRpt, TAB(30); RevenueName$(RCnt + 1); TAB(46); USING "#####.##"; UBTransRec(1).RevAmt(RCnt + 1);
              '  PRINT #UBRpt, TAB(60); RevenueName$(RCnt + 2); TAB(76); USING "#####.##"; UBTransRec(1).RevAmt(RCnt + 2)
              '  LineCnt = LineCnt + 1
              'NEXT RCnt
'              Print #UBRpt, Dash120$
'              Linecnt = Linecnt + 1
            Else
              ToPrintD$ = "~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~"
            End If
            For RCnt = 1 To MaxRevenue
              RevTotals(RCnt) = Round#(RevTotals(RCnt) + (UBTransRec(1).RevAmt(RCnt) + UBTransRec(1).TaxAmt(RCnt)))
            Next
'            If Linecnt > MaxLines Then
'              Print #UBRpt, FF$
'              GoSub DoDetailedRptHeader
'            End If
            Print #UBRpt, ToPrint$ + "~" + ToPrintD$
            ToPrint$ = ""
            ToPrintD$ = ""
          End If
        End If
      End If
'      If AskAbandonPrint% Then
'        AbortFlag = True
'        Exit For
'      End If
      Trans& = UBTransRec(1).PrevTrans
    Loop
SkipThisOne:
'    ShowPctComp cnt, NumOfRecs
  Next
  GoSub DoDetailedRptHeader
  GoSub DoDetailedRptFooter
'  Print #UBRpt, FF$;

  Close

  Erase IdxBuff, UBCustRec
 '' ActivateControls Me, True
  'END

  'If Not AbortFlag Then
  '  PrintRptFile "Detailed Journal Report.", "UBDJLIST.RPT", LptPort, RetCode, EntryPoint
 ' End If
 ' ViewPrint "UBDJLIST.RPT", "Detailed Journal Report", True
  'KillFile "UBDJLIST.RPT"
  If TransCnt& > 0 Then

    Load frmLoadingRpt
    frmLoadingRpt.setwherefrom frmRptTransJournal
    If fpcboRptType.ListIndex = 0 Then
      ARptTransJournal.txtDate = Now
      ARptTransJournal.txtTown = TOWNNAME$
      ARptTransJournal.Title = Head$
      ARptTransJournal.txtRptParm1.Caption = PrnH1$
      ARptTransJournal.txtRptParm2.Caption = PrnH2$
      ARptTransJournal.txtPrnOrd = "In " + fpcboPrintOrder.Text
      ARptTransJournal.totCust = TransCnt&
      'ARptTransJournal.txtTotCur.DataValue = TCurrBalance#
      'ARptTransJournal.txtTotPast.DataValue = TPrevBalance#
      'ARptTransJournal.txtHead = fpcboRevenues.Text
      'ARptTransJournal.txtTotAcctBal.DataValue = Round#(TCurrBalance# + TPrevBalance#)
      ARptTransJournal.GetName ReportFile$, ReportSum$, DetFlag, MaxRevenue
      ARptTransJournal.startrpt
    ElseIf fpcboRptType.ListIndex = 1 Then
      ARptTransJPortrait.txtDate = Now
      ARptTransJPortrait.txtTown = TOWNNAME$
      ARptTransJPortrait.Title = Head$
      ARptTransJPortrait.txtRptParm1.Caption = PrnH1$
      ARptTransJPortrait.txtRptParm2.Caption = PrnH2$
      ARptTransJPortrait.txtPrnOrd = "In " + fpcboPrintOrder.Text
      ARptTransJPortrait.totCust = TransCnt&
      'ARptTransJournal.txtTotCur.DataValue = TCurrBalance#
      'ARptTransJournal.txtTotPast.DataValue = TPrevBalance#
      'ARptTransJournal.txtHead = fpcboRevenues.Text
      'ARptTransJournal.txtTotAcctBal.DataValue = Round#(TCurrBalance# + TPrevBalance#)
      ARptTransJPortrait.GetName ReportFile$, ReportSum$, DetFlag, MaxRevenue
       
      ARptTransJPortrait.startrpt
    End If
  Else
    MsgBox "No Information to print.", vbOKOnly, "No Information"
    ActivateControls Me, True
  End If

ExitDetailedListing:
  
  Exit Sub

DoDetailedRptHeader:
'  PageNo = PageNo + 1
'  Print #UBRpt, TownName$
'  Print #UBRpt, Tab(28); "Detailed Transaction Report"; Tab(90); "Page #"; PageNo
'  Print #UBRpt, "Report Date: "; Date$
  If UseType Then
    PrnH3$ = CUSTTYPE$
  Else
    PrnH3$ = "N/A"
  End If
  PrnH1$ = "   Beginning Transaction Date: " + Date1$ + "        From Book: " + BegRoute
  If Val(Operator$) = 0 Then
    PrnH1$ = PrnH1$ + "     Operator #: ALL" + "      Customer Type: " + PrnH3$
  Else
    PrnH1$ = PrnH1$ + "     Operator #: " + Mid$(Operator$, 1, 3) + "      Customer Type: " + PrnH3$
  End If
  PrnH2$ = "        Ending Transaction Date: " + Date2$ + "     Ending Book: " + EndRoute
  PrnH2$ = PrnH2$ + "     Show Detail: " + Detail$ + "        Transaction Type: " + fpcboTransType.Text
  'PrnH3$ = "          Transaction Type: " + fpcboTransType.Text
' PrnH3$ = PrnH3$ + "             Customer Type: "


'  Print #UBRpt,
'  Print #UBRpt, "  Date"; Tab(11); "Acct #"; Tab(20); "Customer Name"; Tab(55); "Description"; Tab(80); "Oper#"; Tab(90); "Trans Type"; Tab(113); "   Amount"
'  Print #UBRpt, Dash120$
'  Linecnt = 10
  Return

DoDetailedRptFooter:
'  Print #UBRpt, Dash120$
'  Print #UBRpt, "Transactions: "; TransCnt&; "                                                       Total of Transactions: "; Using("$##,###,###.##", TotalTrans#)
'  Print #UBRpt, FF$
'  PageNo = PageNo + 1
'  Print #UBRpt, TownName$
'  Print #UBRpt, Tab(28); "Detailed Transaction Report"; Tab(90); "Page #"; PageNo
'  Print #UBRpt, "Report Date: "; Date$
'  Print #UBRpt, "Beginning Transaction Date: "; Date1$;
'  If Val(Operator$) = 0 Then
'    Print #UBRpt, Tab(90); " Operator #: ALL"
'  Else
'    Print #UBRpt, Tab(90); " Operator #: "; Operator$
'  End If
'  Print #UBRpt, "   Ending Transaction Date: "; Date2$;
'  Print #UBRpt, Tab(90); "Show Detail: "; Detail$
'  Print #UBRpt, ""
'  Print #SumRpt, "Revenue Summary" + "~" + "Amount"
'  Print #UBRpt, Dash120$
  TotalRevsAmt# = 0
  For RCnt = 1 To MaxRevenue
    TotalRevsAmt# = Round#(TotalRevsAmt# + RevTotals(RCnt))
    Print #SumRpt, RevenueName$(RCnt) + "~" + Using("########.##", RevTotals(RCnt))
  Next
'  Print #UBRpt,
'  Print #UBRpt, "Total Amount"; Tab(35); Using("########.##", TotalRevsAmt#)
  Return
DefineType:
  Select Case UBTransRec(1).TransType
  Case 1, 101
    TransDesc$ = "Util Bill"
    'EstFlag = False
    For EstCnt = 1 To 7
      If UBTransRec(1).ESTREAD(EstCnt) = "Y" Then
        'EstFlag = True
        TransDesc$ = TransDesc$ + "*E"
        Exit For
      End If
    Next
    Amount# = UBTransRec(1).Transamt
  Case 2, 102
    TransDesc$ = "Late Charge"
    Amount# = UBTransRec(1).Transamt
  Case 3
    TransDesc$ = "Reconnect"
    Amount# = UBTransRec(1).Transamt
  Case 4, 104
    TransDesc$ = "Reg Payment"
        Amount# = UBTransRec(1).Transamt
  Case 5, 105
    TransDesc$ = "Applied Dep"
    'Amount# = -UBTransRec(1).TransAmt
    Amount# = Abs(UBTransRec(1).Transamt)
  Case 6
    TransDesc$ = "Penalty Chg"
    Amount# = UBTransRec(1).Transamt
  Case 7, 107
    TransDesc$ = "Dep. Payment"
    Amount# = UBTransRec(1).Transamt
  Case 8
    TransDesc$ = "Draft Paymt"
    Amount# = UBTransRec(1).Transamt * -1
  Case 9, 109
    TransDesc$ = "Refunded Dep"
    Amount# = Abs(UBTransRec(1).Transamt)
  Case 10, 110
    TransDesc$ = "Beg Balance"
    Amount# = UBTransRec(1).Transamt
  Case 11, 111
    TransDesc$ = UBTransRec(1).BillMsg
    Amount# = UBTransRec(1).Transamt
  Case 12, 112
    TransDesc$ = UBTransRec(1).BillMsg
    Amount# = UBTransRec(1).Transamt
  Case 33
    TransDesc$ = UBTransRec(1).BillMsg
    Amount# = UBTransRec(1).Transamt
  Case 37
    TransDesc$ = UBTransRec(1).BillMsg
    Amount# = UBTransRec(1).Transamt
  Case 39
    TransDesc$ = UBTransRec(1).BillMsg
    Amount# = UBTransRec(1).Transamt
  Case 99
    TransDesc$ = "Misc Payment"
    Amount# = UBTransRec(1).Transamt
  Case Else
    TransDesc$ = "UNKNOWN"
    Amount# = UBTransRec(1).Transamt
  End Select
  Return

CheckDetailParms:

  Date1$ = txtDate1
  Date2$ = txtDate2

  BegDate = Date2Num%(Date1$)
  EndDate = Date2Num%(Date2$)

  FromBook = Val(BegRoute)
  ThruBook = Val(EndRoute)
  If fpcboTransType.ListIndex <> -1 Then
    TrType$ = QPTrim$(Left$(fpcboTransType.Text, 2))
    TrTyp = Val(TrType$)
  Else
    MsgBox "Invalid Transaction Type.", vbOKOnly, "Invalid Selection"
    fpcboTransType.SetFocus
    GoSub ExitDetailedListing
  End If
'this trtyp of 0 would only work if allowed all
'which we do not allow on transaction type - maybe in administrative section
  If TrTyp = 0 Then
    BegTrans = 1
    EndTrans = 999
  Else
    BegTrans = TrTyp
    EndTrans = TrTyp
  End If

  OperatorNo$ = fptxtOperator
  Operator = Val(OperatorNo$)
  If Operator = 0 Then
    BegOperator = 0
    EndOperator = 9999
  Else
    BegOperator = Operator
    EndOperator = Operator
  End If
  If DelOnly.Value = 1 Then
    Head$ = "Deleted Customers Transaction Journal"
  Else
    Head$ = "Transaction Journal"
  End If
  Detail$ = QPTrim$(Left$(fpcboDetail.Text, 1))

  CUSTTYPE$ = QPTrim$(fptxtCustType)
  If Len(CUSTTYPE$) > 0 Then
    UseType = True
  End If

  Select Case Left$(fpcboPrintOrder.Text, 1)
    Case "C"
    IndexName$ = NameIndexFile
    UsingName = True
  Case "A"
    IndexName$ = ""
    UsingAcct = True
  Case "L"
    IndexName$ = BookIndexFile
    UsingBook = True
  Case "S"
    IndexName$ = TempIndexName
    UsingAddr = True
  Case Else
  End Select
Return
End Sub
Private Sub DetailedTransJournal3()
  Dim UBCustRecLen As Integer, UBSetupreclen As Integer
  Dim UsingBook As Boolean, UsingName As Boolean, UsingAddr As Boolean
  Dim CustomerCnt As Long, UBTransRecLen As Integer, UBTrans As Integer
  Dim IndexName As String, Handle As Integer, Dash120 As String
  Dim IdxRecLen As Integer, IdxFileSize As Long, UBRpt As Integer
  Dim IdxNumOfRecs As Long, NumOfRecs As Long, MaxRevenue As Integer
  Dim cnt As Long, UBCust As Integer, RCnt As Integer, UseType As Boolean
  Dim ThisType As String, CUSTTYPE As String, CustBook As Integer
  Dim FromBook As Integer, ThruBook As Integer, BadCount As Long
  Dim Trans As Long, UBTransLen As Integer, BegDate As Integer
  Dim EndDate As Integer, BegOperator As Integer, EndOperator As Integer
  Dim BegTrans As Integer, EndTrans As Integer, TransDesc As String
  Dim Amount As Double, TotalTrans As Double, TransCnt As Long
  Dim Detail As String, Date1 As String, Date2 As String, Operator As String
  Dim TotalRevsAmt As Double, EstCnt As Integer, TrType As String
  Dim TrTyp As Integer, OperatorNo As String, UsingAcct As Boolean
  Dim ToPrint As String, PrnH1 As String, PrnH2 As String, PrnH3 As String
  Dim SumRpt As Integer, ToPrintD As String, DetFlag As Boolean, Head$
  Dim ReportFile As String, ReportSum As String, MoFlag As Boolean
 'get report parameters
  GoSub CheckDetailParms
  If fpcboDetail.ListIndex = 1 Then
    DetFlag = True
  Else
    DetFlag = False
  End If
'  FrmShowPctComp.Label1 = "Creating Transaction Journal"
'  FrmShowPctComp.Show , Me
'  DoEvents
  ''DeActivateControls Me, True
  ReDim RevTotals(1 To 15) As Double
  ReDim RevenueName(1 To 15) As String
  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))

  ReDim UBSetUp(1) As UBSetupRecType
  UBSetupreclen = Len(UBSetUp(1))

'  IF INSTR(UBSetUp(1).UTILNAME, "AUTRY") > 0 THEN
'    LptPort = 2
'  ELSE
'    LptPort = 1
'  END IF
'  ReDim UBTrans(1) As UBTransRecType
'  UBTransRecLen = Len(UBTrans(1))
  ReDim UBTransRec(1) As UBTransRecType
  UBTransRecLen = Len(UBTransRec(1))

  If UsingName Or UsingBook Then
    IdxRecLen = 4               'we are using a long integer
    IdxFileSize& = FileSize(IndexName$)
    IdxNumOfRecs = IdxFileSize& \ IdxRecLen
   ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
    'FGetAH IndexName$, IdxBuff(1), IdxRecLen, IdxNumOfRecs      'load it
    NumOfRecs = IdxNumOfRecs
    Handle = FreeFile
    Open IndexName$ For Random Shared As Handle Len = IdxRecLen
    For cnt& = 1 To IdxNumOfRecs
      Get #Handle, cnt&, IdxBuff(cnt&)
    Next
    Close Handle
  ElseIf UsingAddr Then
'unrem
    SortServiceAddrs frmRptTransJournal
    IdxRecLen = 4               'we are using a long integer
    IdxFileSize& = FileSize&(IndexName$)
    IdxNumOfRecs = IdxFileSize& \ IdxRecLen
    ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
    'FGetAH IndexName$, IdxBuff(1), IdxRecLen, IdxNumOfRecs      'load it
    NumOfRecs = IdxNumOfRecs
    Handle = FreeFile
    Open IndexName$ For Random Shared As Handle Len = IdxRecLen
    For cnt& = 1 To IdxNumOfRecs
      Get #Handle, cnt&, IdxBuff(cnt&)
    Next
    Close Handle

  Else
    NumOfRecs = FileSize(UBPath$ + "UBCUST.DAT") \ UBCustRecLen
  End If

  UBCust = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen
  UBTrans = FreeFile
  Open UBPath$ + "UBTRANS.DAT" For Random Shared As UBTrans Len = UBTransRecLen
  ReportFile$ = UBPath$ + "UBDJLIST.RPT"
  UBRpt = FreeFile
  Open ReportFile$ For Output As UBRpt
  ReportSum$ = UBPath$ + "UBDJSUM.RPT"
  SumRpt = FreeFile
  Open ReportSum$ For Output As SumRpt
'  ubsetup = FreeFile
'  Open "UBSETUP.DAT" For Random Shared As ubsetup Len = UBSetupreclen
  LoadUBSetUpFile UBSetUp(), UBSetupreclen

  If Len(TOWNNAME$) = 0 Then
    TOWNNAME$ = "Undefined"
    ' Set Revenue Names to Nothing
    For RCnt = 1 To 15
      RevenueName$(RCnt) = "Not Set"
    Next RCnt
  Else
    'Get ubsetup, 1, ubsetup(1)
    For RCnt = 1 To 15
      RevenueName$(RCnt) = QPTrim$(UBSetUp(1).Revenues(RCnt).RevName)
    Next RCnt
    RCnt = 1
    Do While RCnt <= 15
      If RevenueName$(RCnt) = "" Then
        MaxRevenue = RCnt - 1
        Exit Do
      End If
      RCnt = RCnt + 1
    Loop
'    TownName$ = ubsetup(1).UTILNAME
'    TownLen = Len(RTrim$(TownName$))
'    TabStop = 40 - (TownLen / 2)
'    If TabStop < 1 Then TabStop = 1
  End If
  'Close ubsetup

  'Special Code just for ellenboro!!
'  If InStr(TownName$, "ELLENBO") > 0 Then
'    EllenFlag = True
'  End If
'  If InStr(TOWNNAME$, "MOORE") > 0 Or InStr(TOWNNAME$, "JOHNSTON") > 0 Then
'    MoFlag = True
'  End If
  If QckSrch.Value = 1 Then
    MoFlag = True
  Else
    MoFlag = False
  End If
'  BlockClear
'  ShowProcessingScrn "Detailed Journal Report."

'  GoSub DoDetailedRptHeader

'  For cnt = 1 To NumOfRecs
'    FrmShowPctComp.ShowPctComp cnt, NumOfRecs
'    If FrmShowPctComp.Out = True Then
'      Close
'      FrmShowPctComp.Out = False
'      ''ActivateControls Me, True
'      ActivateControls Me, True
'      GoTo ExitDetailedListing
'    End If
'
'    If UsingName Or UsingBook Or UsingAddr Then
'      Get UBCust, IdxBuff(cnt).RecNum, UBCustRec(1)
'    Else
'      Get UBCust, cnt, UBCustRec(1)
'    End If
'   If DelOnly.Value = 1 Then
'  'only list trans for deleted customers *)((*&(*&(*&(*&(&
'      If UBCustRec(1).DelFlag = 0 Then
'        GoTo SkipThisOne
'      End If
'    Else
'      If UBCustRec(1).DelFlag <> 0 Then
'        GoTo SkipThisOne
'      End If
'    End If
'    If UseType Then
'      ThisType$ = QPTrim$(UBCustRec(1).CUSTTYPE)
'      If ThisType$ <> CUSTTYPE$ Then
'        GoTo SkipThisOne
'      End If
'    End If
'
'    CustBook = Val(UBCustRec(1).Book)
'    If CustBook < FromBook Or CustBook > ThruBook Then
'      GoTo SkipThisOne
'    End If

'    If Linecnt > MaxLines Then
'      Print #UBRpt, FF$
'      GoSub DoDetailedRptHeader
'    End If
'*************************************
'   Main Body of Printing goes here
   BadCount = 0
    Trans& = 1
    Do While Trans& <> 0
    If Trans& < 920921 Then
      Get UBTrans, Trans&, UBTransRec(1)
      'If Not EllenFlag Then
'        If UBTransRec(1).TransDate < BegDate Then
'          BadCount = BadCount + 1
'          If BadCount > 3 Then
'            Exit Do
'          End If
'        End If
      'End If
'      If MoFlag Then
'        If UBTransRec(1).TransDate < BegDate Then
'          BadCount = BadCount + 1
'          If BadCount > 3 Then
'            Exit Do
'          End If
'        End If
'      End If
      
      'Check Date, Operator and Trans Type
      'If (UBTransRec(1).TransDate >= BegDate And UBTransRec(1).TransDate <= EndDate) Then
       ' If (UBTransRec(1).OperatorNumber >= BegOperator And UBTransRec(1).OperatorNumber <= EndOperator) Then
        '  If (UBTransRec(1).TransType >= BegTrans And UBTransRec(1).TransType <= EndTrans Or (UBTransRec(1).TransType >= BegTrans + 100 And UBTransRec(1).TransType <= EndTrans + 100)) Then
            Select Case UBTransRec(1).TransType
            Case Is > 1000, 0, Is < 0
            
            ToPrint$ = Str$(Trans&) + "~" + Num2Date$(UBTransRec(1).TransDate) + "~" + Using("#####", UBTransRec(1).CustAcctNo)
            'PRINT #UBRpt, Num2Date$(UBTransRec(1).TransDate); TAB(11); ASC(UBTransRec(1).Posted2GL); 'USING "#####"; UBTrans(1).CustAcctNo;
            ToPrint$ = ToPrint$ + "~" + Left$(UBCustRec(1).CustName, 33)
            ToPrint$ = ToPrint$ + "~" + Str(UBTransRec(1).TransType)
            'PRINT #UBRpt, TAB(55); Trans&;
            ToPrint$ = ToPrint$ + "~" + Str$(UBTransRec(1).OperatorNumber)
            'PRINT #UBRpt, TAB(80); "!"; UBTransRec(1).Posted2GL; "!";
            ToPrint$ = ToPrint$ + "~" + Left$(UBTransRec(1).TransDesc, 20)
            ToPrint$ = ToPrint$ + "~" + Str(UBTransRec(1).Transamt)
            'PRINT #UBRpt, "  "; "!"; UBTransRec(1).Posted2GL; "!"
            'Linecnt = Linecnt + 1
            'TotalTrans# = Round#(TotalTrans# + Dbl(UBTransRec(1).Transamt))
            TransCnt& = TransCnt& + 1
            If Detail$ = "Y" Then
             ' Print #UBRpt, "Revenue Source Breakdown ........................"
             ' Linecnt = Linecnt + 1
              For RCnt = 1 To 15 'MaxRevenue 'Step 3
                If UBTransRec(1).RevAmt(RCnt) <> 0 Then
                  ToPrintD$ = ToPrintD$ + RevenueName$(RCnt) + "~" + Str$(Round#(UBTransRec(1).RevAmt(RCnt) + UBTransRec(1).TaxAmt(RCnt))) + "~"
'                Print #UBRpt, Tab(30); RevenueName$(RCnt + 1); Tab(46); Using("#####.##", UBTransRec(1).RevAmt(RCnt + 1));
'                Print #UBRpt, Tab(60); RevenueName$(RCnt + 2); Tab(76); Using("#####.##", UBTransRec(1).RevAmt(RCnt + 2))
              '  Linecnt = Linecnt + 1
                Else
                  If Len(RevenueName$(RCnt)) > 0 Then
                    ToPrintD$ = ToPrintD$ + RevenueName$(RCnt) + "~0.00~"
                  Else
                    ToPrintD$ = ToPrintD$ + " ~  ~"
                  End If
                End If
              Next RCnt
              'IF UBTransRec(1).TransType = TranUpwardAdjustment OR UBTransRec(1).TransType = TranDownwardAdjustment THEN
              '  FOR RCnt = 1 TO 7
              '    PRINT #UBRpt, RevenueName$(RCnt); TAB(16); USING "#####.##"; UBTransRec(1).RevAmt(RCnt);
              '  PRINT #UBRpt, TAB(30); RevenueName$(RCnt + 1); TAB(46); USING "#####.##"; UBTransRec(1).RevAmt(RCnt + 1);
              '  PRINT #UBRpt, TAB(60); RevenueName$(RCnt + 2); TAB(76); USING "#####.##"; UBTransRec(1).RevAmt(RCnt + 2)
              '  LineCnt = LineCnt + 1
              'NEXT RCnt
'              Print #UBRpt, Dash120$
'              Linecnt = Linecnt + 1
            Else
              ToPrintD$ = "~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~"
            End If
            For RCnt = 1 To MaxRevenue
              RevTotals(RCnt) = Round#(RevTotals(RCnt) + (UBTransRec(1).RevAmt(RCnt) + UBTransRec(1).TaxAmt(RCnt)))
            Next
'            If Linecnt > MaxLines Then
'              Print #UBRpt, FF$
'              GoSub DoDetailedRptHeader
'            End If
            Print #UBRpt, ToPrint$ + "~" + ToPrintD$
            ToPrint$ = ""
            ToPrintD$ = ""
            Case Else
            End Select
      '    End If
     '   End If
    '  End If
'      If AskAbandonPrint% Then
'        AbortFlag = True
'        Exit For
      
      Trans& = Trans& + 1
      Else
      Trans& = 0
      End If
    Loop
SkipThisOne:
'    ShowPctComp cnt, NumOfRecs
  'Next
  GoSub DoDetailedRptHeader
  GoSub DoDetailedRptFooter
'  Print #UBRpt, FF$;

  Close

  Erase IdxBuff, UBCustRec
 '' ActivateControls Me, True
  'END

  'If Not AbortFlag Then
  '  PrintRptFile "Detailed Journal Report.", "UBDJLIST.RPT", LptPort, RetCode, EntryPoint
 ' End If
 ' ViewPrint "UBDJLIST.RPT", "Detailed Journal Report", True
  'KillFile "UBDJLIST.RPT"
  If TransCnt& > 0 Then

    Load frmLoadingRpt
    frmLoadingRpt.setwherefrom frmRptTransJournal
    If fpcboRptType.ListIndex = 0 Then
      ARptTransJournal.txtDate = Now
      ARptTransJournal.txtTown = TOWNNAME$
      ARptTransJournal.Title = Head$
      ARptTransJournal.txtRptParm1.Caption = PrnH1$
      ARptTransJournal.txtRptParm2.Caption = PrnH2$
      ARptTransJournal.txtPrnOrd = "In " + fpcboPrintOrder.Text
      ARptTransJournal.totCust = TransCnt&
      'ARptTransJournal.txtTotCur.DataValue = TCurrBalance#
      'ARptTransJournal.txtTotPast.DataValue = TPrevBalance#
      'ARptTransJournal.txtHead = fpcboRevenues.Text
      'ARptTransJournal.txtTotAcctBal.DataValue = Round#(TCurrBalance# + TPrevBalance#)
      ARptTransJournal.GetName ReportFile$, ReportSum$, DetFlag, MaxRevenue
      ARptTransJournal.startrpt
    ElseIf fpcboRptType.ListIndex = 1 Then
      ARptTransJPortrait.txtDate = Now
      ARptTransJPortrait.txtTown = TOWNNAME$
      ARptTransJPortrait.Title = Head$
      ARptTransJPortrait.txtRptParm1.Caption = PrnH1$
      ARptTransJPortrait.txtRptParm2.Caption = PrnH2$
      ARptTransJPortrait.txtPrnOrd = "In " + fpcboPrintOrder.Text
      ARptTransJPortrait.totCust = TransCnt&
      'ARptTransJournal.txtTotCur.DataValue = TCurrBalance#
      'ARptTransJournal.txtTotPast.DataValue = TPrevBalance#
      'ARptTransJournal.txtHead = fpcboRevenues.Text
      'ARptTransJournal.txtTotAcctBal.DataValue = Round#(TCurrBalance# + TPrevBalance#)
      ARptTransJPortrait.GetName ReportFile$, ReportSum$, DetFlag, MaxRevenue
       
      ARptTransJPortrait.startrpt
    End If
  Else
    MsgBox "No Information to print.", vbOKOnly, "No Information"
    ActivateControls Me, True
  End If

ExitDetailedListing:
  
  Exit Sub

DoDetailedRptHeader:
'  PageNo = PageNo + 1
'  Print #UBRpt, TownName$
'  Print #UBRpt, Tab(28); "Detailed Transaction Report"; Tab(90); "Page #"; PageNo
'  Print #UBRpt, "Report Date: "; Date$
  If UseType Then
    PrnH3$ = CUSTTYPE$
  Else
    PrnH3$ = "N/A"
  End If
  PrnH1$ = "   Beginning Transaction Date: " + Date1$ + "        From Book: " + BegRoute
  If Val(Operator$) = 0 Then
    PrnH1$ = PrnH1$ + "     Operator #: ALL" + "      Customer Type: " + PrnH3$
  Else
    PrnH1$ = PrnH1$ + "     Operator #: " + Mid$(Operator$, 1, 3) + "      Customer Type: " + PrnH3$
  End If
  PrnH2$ = "        Ending Transaction Date: " + Date2$ + "     Ending Book: " + EndRoute
  PrnH2$ = PrnH2$ + "     Show Detail: " + Detail$ + "        Transaction Type: " + fpcboTransType.Text
  'PrnH3$ = "          Transaction Type: " + fpcboTransType.Text
' PrnH3$ = PrnH3$ + "             Customer Type: "


'  Print #UBRpt,
'  Print #UBRpt, "  Date"; Tab(11); "Acct #"; Tab(20); "Customer Name"; Tab(55); "Description"; Tab(80); "Oper#"; Tab(90); "Trans Type"; Tab(113); "   Amount"
'  Print #UBRpt, Dash120$
'  Linecnt = 10
  Return

DoDetailedRptFooter:
'  Print #UBRpt, Dash120$
'  Print #UBRpt, "Transactions: "; TransCnt&; "                                                       Total of Transactions: "; Using("$##,###,###.##", TotalTrans#)
'  Print #UBRpt, FF$
'  PageNo = PageNo + 1
'  Print #UBRpt, TownName$
'  Print #UBRpt, Tab(28); "Detailed Transaction Report"; Tab(90); "Page #"; PageNo
'  Print #UBRpt, "Report Date: "; Date$
'  Print #UBRpt, "Beginning Transaction Date: "; Date1$;
'  If Val(Operator$) = 0 Then
'    Print #UBRpt, Tab(90); " Operator #: ALL"
'  Else
'    Print #UBRpt, Tab(90); " Operator #: "; Operator$
'  End If
'  Print #UBRpt, "   Ending Transaction Date: "; Date2$;
'  Print #UBRpt, Tab(90); "Show Detail: "; Detail$
'  Print #UBRpt, ""
'  Print #SumRpt, "Revenue Summary" + "~" + "Amount"
'  Print #UBRpt, Dash120$
  TotalRevsAmt# = 0
  For RCnt = 1 To MaxRevenue
    TotalRevsAmt# = Round#(TotalRevsAmt# + RevTotals(RCnt))
    Print #SumRpt, RevenueName$(RCnt) + "~" + Using("########.##", RevTotals(RCnt))
  Next
'  Print #UBRpt,
'  Print #UBRpt, "Total Amount"; Tab(35); Using("########.##", TotalRevsAmt#)
  Return
DefineType:
  Select Case UBTransRec(1).TransType
  
  Case 1, 101
    TransDesc$ = "Util Bill"
    'EstFlag = False
    For EstCnt = 1 To 7
      If UBTransRec(1).ESTREAD(EstCnt) = "Y" Then
        'EstFlag = True
        TransDesc$ = TransDesc$ + "*E"
        Exit For
      End If
    Next
    Amount# = UBTransRec(1).Transamt
  Case 2, 102
    TransDesc$ = "Late Charge"
    Amount# = UBTransRec(1).Transamt
  Case 3
    TransDesc$ = "Reconnect"
    Amount# = UBTransRec(1).Transamt
  Case 4, 104
    TransDesc$ = "Reg Payment"
        Amount# = UBTransRec(1).Transamt
  Case 5, 105
    TransDesc$ = "Applied Dep"
    'Amount# = -UBTransRec(1).TransAmt
    Amount# = Abs(UBTransRec(1).Transamt)
  Case 6
    TransDesc$ = "Penalty Chg"
    Amount# = UBTransRec(1).Transamt
  Case 7, 107
    TransDesc$ = "Dep. Payment"
    Amount# = UBTransRec(1).Transamt
  Case 8
    TransDesc$ = "Draft Paymt"
    Amount# = UBTransRec(1).Transamt * -1
  Case 9, 109
    TransDesc$ = "Refunded Dep"
    Amount# = Abs(UBTransRec(1).Transamt)
  Case 10, 110
    TransDesc$ = "Beg Balance"
    Amount# = UBTransRec(1).Transamt
  Case 11, 111
    TransDesc$ = UBTransRec(1).BillMsg
    Amount# = UBTransRec(1).Transamt
  Case 12, 112
    TransDesc$ = UBTransRec(1).BillMsg
    Amount# = UBTransRec(1).Transamt
  Case 33
    TransDesc$ = UBTransRec(1).BillMsg
    Amount# = UBTransRec(1).Transamt
  Case 37
    TransDesc$ = UBTransRec(1).BillMsg
    Amount# = UBTransRec(1).Transamt
  Case 39
    TransDesc$ = UBTransRec(1).BillMsg
    Amount# = UBTransRec(1).Transamt
  Case 99
    TransDesc$ = "Misc Payment"
    Amount# = UBTransRec(1).Transamt
  Case Else
    TransDesc$ = Str(UBTransRec(1).TransType)
    Amount# = UBTransRec(1).Transamt
  End Select
  Return

CheckDetailParms:

  Date1$ = txtDate1
  Date2$ = txtDate2

  BegDate = Date2Num%(Date1$)
  EndDate = Date2Num%(Date2$)

  FromBook = Val(BegRoute)
  ThruBook = Val(EndRoute)
  If fpcboTransType.ListIndex <> -1 Then
    TrType$ = QPTrim$(Left$(fpcboTransType.Text, 2))
    TrTyp = Val(TrType$)
  Else
    MsgBox "Invalid Transaction Type.", vbOKOnly, "Invalid Selection"
    fpcboTransType.SetFocus
    GoSub ExitDetailedListing
  End If
'this trtyp of 0 would only work if allowed all
'which we do not allow on transaction type - maybe in administrative section
  If TrTyp = 0 Then
    BegTrans = 1
    EndTrans = 999
  Else
    BegTrans = TrTyp
    EndTrans = TrTyp
  End If

  OperatorNo$ = fptxtOperator
  Operator = Val(OperatorNo$)
  If Operator = 0 Then
    BegOperator = 0
    EndOperator = 9999
  Else
    BegOperator = Operator
    EndOperator = Operator
  End If
  If DelOnly.Value = 1 Then
    Head$ = "Deleted Customers Transaction Journal"
  Else
    Head$ = "Transaction Journal"
  End If
  Detail$ = QPTrim$(Left$(fpcboDetail.Text, 1))

  CUSTTYPE$ = QPTrim$(fptxtCustType)
  If Len(CUSTTYPE$) > 0 Then
    UseType = True
  End If

  Select Case Left$(fpcboPrintOrder.Text, 1)
    Case "C"
    IndexName$ = NameIndexFile
    UsingName = True
  Case "A"
    IndexName$ = ""
    UsingAcct = True
  Case "L"
    IndexName$ = BookIndexFile
    UsingBook = True
  Case "S"
    IndexName$ = TempIndexName
    UsingAddr = True
  Case Else
  End Select
Return
End Sub


