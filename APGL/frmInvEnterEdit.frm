VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#3.5#0"; "SPR32X35.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInvEnterEdit 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "A/P Invoice Entry"
   ClientHeight    =   8850
   ClientLeft      =   30
   ClientTop       =   540
   ClientWidth     =   12195
   ClipControls    =   0   'False
   Icon            =   "frmInvEnterEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8850
   ScaleWidth      =   12195
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboAcctNumNa 
      Height          =   375
      Left            =   1395
      TabIndex        =   11
      Top             =   3630
      Width           =   4935
      _Version        =   196608
      _ExtentX        =   8705
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
      ColDesigner     =   "frmInvEnterEdit.frx":08CA
   End
   Begin LpLib.fpCombo fpcboPSL 
      Height          =   375
      Left            =   9930
      TabIndex        =   9
      Top             =   2010
      Width           =   825
      _Version        =   196608
      _ExtentX        =   1455
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
      Columns         =   2
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
      ListWidth       =   792
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
      EditMarginLeft  =   1
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmInvEnterEdit.frx":0CCD
   End
   Begin LpLib.fpCombo fpcbo1099 
      Height          =   375
      Left            =   9930
      TabIndex        =   10
      Top             =   2370
      Width           =   825
      _Version        =   196608
      _ExtentX        =   1455
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
      Columns         =   2
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
      ListWidth       =   792
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
      EditMarginLeft  =   1
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmInvEnterEdit.frx":1068
   End
   Begin LpLib.fpCombo fpcboVendName 
      Height          =   375
      Left            =   2715
      TabIndex        =   0
      Top             =   975
      Width           =   4140
      _Version        =   196608
      _ExtentX        =   7302
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
      ColDesigner     =   "frmInvEnterEdit.frx":1403
   End
   Begin LpLib.fpCombo fpcboTaxable 
      Height          =   375
      Left            =   2715
      TabIndex        =   5
      Top             =   2880
      Width           =   810
      _Version        =   196608
      _ExtentX        =   1429
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
      Columns         =   2
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
      ListWidth       =   792
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
      EditMarginLeft  =   1
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmInvEnterEdit.frx":17B6
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   2595
      Left            =   1710
      TabIndex        =   14
      Top             =   4350
      Width           =   8790
      _Version        =   196613
      _ExtentX        =   15505
      _ExtentY        =   4577
      _StockProps     =   64
      BackColorStyle  =   1
      ButtonDrawMode  =   4
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   6
      MaxRows         =   36
      OperationMode   =   2
      SelectBlockOptions=   2
      ShadowColor     =   13684944
      ShadowDark      =   8421504
      SpreadDesigner  =   "frmInvEnterEdit.frx":1B51
      VisibleCols     =   3
      VisibleRows     =   8
      ScrollBarTrack  =   1
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00D0D0D0&
      Caption         =   "Alt C &Clear"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   444
      Left            =   9552
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   3504
      Width           =   1212
   End
   Begin EditLib.fpText fpAPLegNum 
      Height          =   396
      Left            =   144
      TabIndex        =   56
      Top             =   3984
      Visible         =   0   'False
      Width           =   1020
      _Version        =   196608
      _ExtentX        =   1799
      _ExtentY        =   698
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
   Begin EditLib.fpText fpCode 
      Height          =   324
      Left            =   288
      TabIndex        =   55
      Top             =   3312
      Visible         =   0   'False
      Width           =   684
      _Version        =   196608
      _ExtentX        =   1206
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
   Begin EditLib.fpText fpDacrec 
      Height          =   324
      Left            =   240
      TabIndex        =   54
      Top             =   2712
      Visible         =   0   'False
      Width           =   924
      _Version        =   196608
      _ExtentX        =   1630
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
   Begin EditLib.fpText fpPODistNum 
      Height          =   324
      Left            =   240
      TabIndex        =   53
      Top             =   2112
      Visible         =   0   'False
      Width           =   948
      _Version        =   196608
      _ExtentX        =   1672
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
   Begin EditLib.fpText fptxtPo 
      Height          =   348
      Left            =   4152
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   1728
      Width           =   1476
      _Version        =   196608
      _ExtentX        =   2603
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
   Begin VB.CommandButton cmdPOList 
      BackColor       =   &H00D0D0D0&
      Caption         =   "F8 &PO List"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   2736
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   1752
      Width           =   1332
   End
   Begin VB.CommandButton cmdDelDist 
      BackColor       =   &H00D0D0D0&
      Caption         =   "F6 Del D&ist"
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
      Left            =   5675
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   7752
      Width           =   1428
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00D0D0D0&
      Caption         =   "F10 &Save"
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
      Left            =   9143
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7752
      Width           =   1236
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00D0D0D0&
      Caption         =   "F4 &Edit"
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
      Left            =   3203
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   7752
      Width           =   1044
   End
   Begin VB.CommandButton cmdNew 
      BackColor       =   &H00D0D0D0&
      Caption         =   "F2 &New"
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
      Left            =   530
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7752
      Width           =   1020
   End
   Begin VB.CommandButton cmdList 
      BackColor       =   &H00D0D0D0&
      Caption         =   "F5 &List"
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
      Left            =   4439
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   7752
      Width           =   1044
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00D0D0D0&
      Caption         =   "F3 &Delete"
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
      Left            =   1739
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   7752
      Width           =   1284
   End
   Begin VB.CommandButton cmdNewVend 
      BackColor       =   &H00D0D0D0&
      Caption         =   "F7 New &Vendor"
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
      Left            =   7283
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   7752
      Width           =   1668
   End
   Begin EditLib.fpText fpInvNum 
      Height          =   348
      Left            =   2712
      TabIndex        =   1
      Top             =   1356
      Width           =   3492
      _Version        =   196608
      _ExtentX        =   6159
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
      ButtonStyle     =   0
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
      AutoCase        =   0
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
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   25
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
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
      Caption         =   "F9 &Add Dist."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   468
      Left            =   8112
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3504
      Width           =   1332
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00D0D0D0&
      Caption         =   "Esc E&xit"
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
      Left            =   10572
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7752
      Width           =   1092
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   24
      Top             =   8565
      Width           =   12195
      _ExtentX        =   21511
      _ExtentY        =   503
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
            TextSave        =   "8:58 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7117
            TextSave        =   "6/18/2018"
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
   Begin EditLib.fpCurrency fpDist 
      Height          =   348
      Left            =   7752
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   7176
      Width           =   1812
      _Version        =   196608
      _ExtentX        =   3196
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
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   1
      ControlType     =   1
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   2
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   "$"
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
   Begin EditLib.fpCurrency fpUndist 
      Height          =   348
      Left            =   3792
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   7176
      Width           =   1812
      _Version        =   196608
      _ExtentX        =   3196
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
      ControlType     =   1
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   2
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   "$"
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
   Begin EditLib.fpCurrency fpDebAmt 
      Height          =   348
      Left            =   6480
      TabIndex        =   12
      Top             =   3624
      Width           =   1452
      _Version        =   196608
      _ExtentX        =   2561
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
      ButtonStyle     =   0
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483637
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   2
      AlignTextV      =   1
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
      ControlType     =   0
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   -1
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   "$"
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "999999999"
      MinValue        =   "-999999999"
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
   Begin EditLib.fpText fptxtDesc 
      Height          =   348
      Left            =   2712
      TabIndex        =   4
      Top             =   2496
      Width           =   2412
      _Version        =   196608
      _ExtentX        =   4254
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
      ButtonStyle     =   0
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
      AutoCase        =   0
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
      OnFocusNoSelect =   0   'False
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
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpCurrency fpInvAmt 
      Height          =   348
      Left            =   2712
      TabIndex        =   2
      Top             =   2112
      Width           =   1500
      _Version        =   196608
      _ExtentX        =   2646
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
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   2
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   "$"
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "99999999"
      MinValue        =   "-99999999"
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
      IncDec          =   1
      BorderGrayAreaColor=   -2147483637
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpDateTime fpInvDate 
      Height          =   348
      Left            =   9120
      TabIndex        =   6
      Top             =   960
      Width           =   1620
      _Version        =   196608
      _ExtentX        =   2857
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
      AllowNull       =   -1  'True
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
      Text            =   "10/01/2001"
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
   Begin EditLib.fpDateTime fpDueDate 
      Height          =   348
      Left            =   9120
      TabIndex        =   7
      Top             =   1296
      Width           =   1620
      _Version        =   196608
      _ExtentX        =   2857
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
      AllowNull       =   -1  'True
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
      Text            =   "10/01/2001"
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
   Begin EditLib.fpDateTime fpPostDate 
      Height          =   348
      Left            =   9120
      TabIndex        =   8
      Top             =   1632
      Width           =   1620
      _Version        =   196608
      _ExtentX        =   2857
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
      AllowNull       =   -1  'True
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
      Text            =   "10/01/2001"
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
   Begin EditLib.fpCurrency fpInvTotal 
      Height          =   348
      Left            =   9000
      TabIndex        =   39
      Top             =   2880
      Width           =   1740
      _Version        =   196608
      _ExtentX        =   3069
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
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   1
      ControlType     =   1
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   2
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   "$"
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "999999999"
      MinValue        =   "-999999999"
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
   Begin EditLib.fpCurrency fpSTax 
      Height          =   348
      Left            =   4464
      TabIndex        =   49
      Top             =   2880
      Width           =   1116
      _Version        =   196608
      _ExtentX        =   1968
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
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   1
      ControlType     =   1
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   2
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   "$"
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "999999999"
      MinValue        =   "-999999999"
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
      IncDec          =   1
      BorderGrayAreaColor=   -2147483637
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpCurrency fpCTax 
      Height          =   348
      Left            =   6480
      TabIndex        =   50
      Top             =   2880
      Width           =   1092
      _Version        =   196608
      _ExtentX        =   1926
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
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   1
      ControlType     =   1
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   2
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   "$"
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "999999999"
      MinValue        =   "-999999999"
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
      IncDec          =   1
      BorderGrayAreaColor=   -2147483637
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fpManPO 
      Height          =   348
      Left            =   5736
      TabIndex        =   3
      Top             =   2112
      Width           =   1476
      _Version        =   196608
      _ExtentX        =   2603
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
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Manual PO:"
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
      Left            =   4416
      TabIndex        =   58
      Top             =   2160
      Width           =   1260
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   6204
      Left            =   1296
      Top             =   864
      Width           =   9564
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CTax:"
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
      Left            =   5712
      TabIndex        =   48
      Top             =   2952
      Width           =   660
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "STax:"
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
      Left            =   3648
      TabIndex        =   47
      Top             =   2952
      Width           =   708
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Taxable:"
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
      Left            =   1656
      TabIndex        =   46
      Top             =   2928
      Width           =   948
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H80000016&
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   732
      Index           =   0
      Left            =   1296
      Top             =   3312
      Width           =   9564
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Inv Number:"
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
      Left            =   1392
      TabIndex        =   45
      Top             =   1404
      Width           =   1236
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "PO Number:"
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
      Left            =   1368
      TabIndex        =   44
      Top             =   1788
      Width           =   1260
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor:"
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
      Left            =   1776
      TabIndex        =   43
      Top             =   1032
      Width           =   852
   End
   Begin VB.Label Label3b 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Inv Desc:"
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
      Index           =   1
      Left            =   1632
      TabIndex        =   42
      Top             =   2544
      Width           =   972
   End
   Begin VB.Label Label2b 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Inv Amount:"
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
      Index           =   1
      Left            =   1320
      TabIndex        =   41
      Top             =   2184
      Width           =   1308
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Inv Total:"
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
      Left            =   7920
      TabIndex        =   40
      Top             =   2952
      Width           =   1044
   End
   Begin VB.Label Label3b 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "1099 Transaction:"
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
      Index           =   0
      Left            =   8088
      TabIndex        =   38
      Top             =   2424
      Width           =   1764
   End
   Begin VB.Label Label4b 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Due Date:"
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
      Index           =   1
      Left            =   8040
      TabIndex        =   37
      Top             =   1344
      Width           =   996
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Inv Date:"
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
      Index           =   1
      Left            =   8016
      TabIndex        =   36
      Top             =   1008
      Width           =   1020
   End
   Begin VB.Label Label33 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Include on PSL:"
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
      Index           =   0
      Left            =   8184
      TabIndex        =   35
      Top             =   2088
      Width           =   1668
   End
   Begin VB.Label lblNew 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Caption         =   "New Invoice"
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
      Left            =   864
      TabIndex        =   34
      Top             =   480
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.Label lblEdit 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Caption         =   "Edit Invoice"
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
      Left            =   9696
      TabIndex        =   33
      Top             =   480
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.Label Label3b 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Post Date:"
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
      Index           =   2
      Left            =   7896
      TabIndex        =   32
      Top             =   1704
      Width           =   1140
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "GL Account Number/Name"
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
      Index           =   0
      Left            =   2544
      TabIndex        =   31
      Top             =   3360
      Width           =   2652
   End
   Begin VB.Label Label3b 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Debit Amount"
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
      Index           =   5
      Left            =   6552
      TabIndex        =   30
      Top             =   3360
      Width           =   1332
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Distributions :"
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
      Left            =   1488
      TabIndex        =   29
      Top             =   4056
      Width           =   1500
   End
   Begin VB.Label lblDebits 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Undistributed"
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
      Index           =   2
      Left            =   2376
      TabIndex        =   28
      Top             =   7224
      Width           =   1284
   End
   Begin VB.Label lblCredits 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Distributed"
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
      Index           =   3
      Left            =   6528
      TabIndex        =   27
      Top             =   7224
      Width           =   1116
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000005&
      BorderWidth     =   3
      Height          =   492
      Left            =   1296
      Top             =   7056
      Width           =   9564
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D0D0D0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000005&
      BorderWidth     =   3
      Height          =   2772
      Left            =   1632
      Top             =   4296
      Width           =   8892
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Enter/Edit Invoices"
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
      TabIndex        =   23
      Top             =   360
      Width           =   4020
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000009&
      Height          =   636
      Left            =   2580
      Top             =   216
      Width           =   7020
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00D0D0D0&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   780
      Left            =   2592
      Top             =   96
      Width           =   7020
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
Attribute VB_Name = "frmInvEnterEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Dim LPDate As Integer, HPDate As Integer, TInvDate As String
Dim POControl As POControlRecType
Dim Vendor As VendorRecType
Dim VendorIdx As VendorIdxRecType
Dim APIED As APInv85Type
Dim APLedgerRec As APLedger81RecType
Dim DefDist As VendorDefDistRecType
Private Temp_Class As Resize_Class
Dim EMode As Boolean, RecNum As Integer, appolines As Integer, RecLok As Boolean
Dim GotTaxFile As Boolean, StaTaxFlag As Boolean, CtyTaxFlag As Boolean
Dim STATAX As Double, CTYTAX As Double, AutoDistFlag As Boolean, OldRec As Integer
Dim POTabStop As Boolean, StCode As Boolean, CnCode As Boolean
Dim PSLDef As Boolean, DupInvDef As Boolean
Private Sub NextNew()
  Dim APEditFile As Integer, NumEdTrans As Integer
  OpenAPEditFile APEditFile, NumEdTrans
  Close APEditFile
   If NumEdTrans > 0 Then
     RecNum = NumEdTrans + 1
   Else
     RecNum = 1
   End If
   EMode = False
   ClearScn
   SetScreen
   InitTaxInfo
   fpcboVendName.SetFocus
End Sub

Private Sub cmdClear_Click()
  fpcboAcctNumNa.ListIndex = -1
  fpDebAmt = 0
  fpcboAcctNumNa.Enabled = True
  fpPODistNum = 0
  fpDacrec = 0
  fpCode = "N"
  fpcboAcctNumNa.SetFocus
End Sub

Private Sub cmdDelete_Click()
  Dim APEditFile As Integer, NumEdTrans As Integer
  Dim InvBusy As Boolean
  InvBusy = False
  If Exist("APIED.DAT") Then InvBusy = GetAttr("APIED.DAT") And vbReadOnly
  If Not InvBusy Then
    If EMode = True Then
      If MsgBox("Are you sure you wish to delete this entry?", vbYesNo, "Delete GJEntry") = vbYes Then
        OpenAPEditFile APEditFile, NumEdTrans
        APIED.LOCKED = False
        APIED.DelFlag = -1
        Put APEditFile, RecNum, APIED
        Close APEditFile
        Call MainLog("Del Inv - " + fpInvNum)
        APIED.DelFlag = 0
        Call NextNew
      Else
        fpcboAcctNumNa.SetFocus
      End If
    Else
      If Changed = True Then
        If MsgBox("Are you sure you wish to delete this entry?", vbYesNo, "Delete GJEntry") = vbYes Then
          Call NextNew
        End If
      Else
        MsgBox "Nothing to Delete", vbOKOnly, "Delete Canceled"
      End If
    End If
  Else
    MsgBox "Posting Is In Progress, Editing Not Allowed At This Time.", vbOKOnly, "Canceled"
    frmInvProcessMenu.Show
    Unload frmInvEnterEdit
  End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbYes Then
      If Changed = False Then
        Undolok RecNum
        Call MainLog("Close via InvEnterEdit")
        ClearInUse PWcnt
      Else
        If MsgBox("Changes Have Been Made to the Current Record." & Chr(13) & Chr(13) & "                          Select OK to Abandon," & Chr(13) & Chr(13) & "       or Cancel to Remain on Entry/Edit Screen.", vbOKCancel, "Abandon Changes?") = vbOK Then
          Undolok RecNum
          Call MainLog("Close via InvEnterEdit")
          ClearInUse PWcnt
        Else
          Cancel = True
        End If
      End If
    Else
      Cancel = True
    End If
  End If
End Sub
Private Sub cmdNewVend_Click()
  frmAddVendor.Inv = True
  'frmLoadingRpt.Show
  DoEvents
  'Load frmAddVendor
  
  frmAddVendor.Show ' 1, Me
  'Unload frmLoadingRpt
End Sub
Public Sub FirstOpenInv()
  If RecLok = True Then
    frmInvListing.Show 1, frmInvEnterEdit
  End If
End Sub
Private Sub Undolok(OldRec)
  Dim APEditFile As Integer, NumEdTrans As Integer
  Dim InvBusy As Boolean
  If Exist("APIED.DAT") Then InvBusy = GetAttr("APIEd.DAT") And vbReadOnly
  If Not InvBusy Then
    OpenAPEditFile APEditFile, NumEdTrans
      If OldRec <= NumEdTrans Then
        Get APEditFile, OldRec, APIED
        APIED.LOCKED = False
        Put APEditFile, OldRec, APIED
      End If
      Close APEditFile
  Else
    MsgBox "Posting In Progress, Editing May Not Continue At This Time.", vbOKOnly, "Canceled"
    frmInvProcessMenu.Show
    Unload frmInvEnterEdit
  End If
End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  GetPostDates LPDate, HPDate  'In Main Module to get dates from setup
  GetInvDef POTabStop, PSLDef, DupInvDef
  StatusBar1.Panels.Item(1).Text = GLUserName
  Me.HelpContextID = hlpEnterInv
  VendCodeName fpcboVendName
  fpcboTaxable.AddItem "Yes"
  fpcboTaxable.AddItem "No"
  fpcbo1099.AddItem "Yes"
  fpcbo1099.AddItem "No"
  fpcboPSL.AddItem "Yes"
  fpcboPSL.AddItem "No"
  'added to setupfile a value to stop at po option on invoices
  If POTabStop = True Then
    cmdPOList.TabIndex = 2
    cmdPOList.TabStop = True
  End If
  If PSLDef = True Then
    fpcboPSL.ListIndex = 0
  Else
    fpcboPSL.ListIndex = 1
  End If
  InitTaxInfo
  If AutoDistFlag = True Then
    fpcboTaxable.ListIndex = 0
  Else
    fpcboTaxable.ListIndex = 1
  End If

  Fixspread
  FillAcctNumName fpcboAcctNumNa
  EdorNewEntry
End Sub
Public Function Rec2Form(TempRec)
  Dim APEditFile As Integer, NumEdTrans As Integer, Rec As Integer
  Dim FileName As String, EdLen As Integer, POUcnt As Integer
  Dim CurrRec As Integer, NextRec As Integer, cnt As Integer, Last As Integer
  Dim VendorFile As Integer, NumVRecs As Integer, vrec As Integer
  'If used list or edit the temprec was selected there, need to transfer to recnum
  Dim InvBusy As Boolean
  InvBusy = False
  If Exist("APIED.DAT") Then InvBusy = GetAttr("APIED.DAT") And vbReadOnly
  If Not InvBusy Then
    OldRec = RecNum
    RecNum = TempRec
    OpenAPEditFile APEditFile, NumEdTrans
    OpenVendorFile VendorFile, NumVRecs
    Get APEditFile, RecNum, APIED
    If APIED.LOCKED = False Then
      APIED.LOCKED = True
      Put APEditFile, RecNum, APIED
      vrec = Trim(APIED.VRecNum)
        fpcboVendName.SearchText = vrec
        fpcboVendName.ColumnSearch = 2
        fpcboVendName.Action = ActionSearch
        If fpcboVendName.SearchIndex <> -1 Then
          fpcboVendName.ListIndex = fpcboVendName.SearchIndex
        End If
        fpcboVendName.ColumnSearch = 0
      Close VendorFile
      
      fpInvDate = Format(DateAdd("d", (APIED.InvDate), "12-31-1979"), "mm/dd/yyyy")
      fpDueDate = Format(DateAdd("d", (APIED.DueDate), "12-31-1979"), "mm/dd/yyyy")
      fpPostDate = Format(DateAdd("d", (APIED.DISTDATE), "12-31-1979"), "mm/dd/yyyy")
      fptxtPo = APIED.PONum
      fpManPO = APIED.MPONum
      appolines = APIED.POLINES
      fpInvAmt = APIED.InvAmt
      fptxtDesc = APIED.INVDESC
      If APIED.TAXYN = "Y" Then
        fpcboTaxable.ListIndex = 0
      Else
        fpcboTaxable.ListIndex = 1
      End If
      fpSTax = APIED.STAXAMT
      fpCTax = APIED.CTAXAMT
      If APIED.Get1099 = "Y" Then
        fpcbo1099.ListIndex = 0
      Else
        fpcbo1099.ListIndex = 1
      End If
      If APIED.PSLFlag = "Y" Then
        fpcboPSL.ListIndex = 0
      Else
        fpcboPSL.ListIndex = 1
      End If
      fpInvTotal = APIED.GRANDTOT
      fpAPLegNum = APIED.POAPLRecNum
      
      Last = UBound(APIED.Dist)
      POUcnt = 0
      For cnt = 1 To Last
        If Val(APIED.Dist(cnt).DACN) <> 0 Then
          vaSpread1.Row = vaSpread1.DataRowCnt + 1
          vaSpread1.col = 1
          vaSpread1.Text = APIED.Dist(cnt).DISTNUM
          vaSpread1.col = 2
          vaSpread1.Text = APIED.Dist(cnt).DACREC
          vaSpread1.col = 3
          vaSpread1.Text = APIED.Dist(cnt).DACODE
          vaSpread1.col = 4
          vaSpread1.Text = APIED.Dist(cnt).DACN
          vaSpread1.col = 5
          vaSpread1.Text = APIED.Dist(cnt).DACNM
          vaSpread1.col = 6
          vaSpread1.Text = APIED.Dist(cnt).DAMT
          fpDist = (fpDist.DoubleValue + APIED.Dist(cnt).DAMT)
        Else
          Exit For
        End If
        If APIED.Dist(cnt).DISTNUM > 0 Then
          POUcnt = POUcnt + 1
        End If
      Next
      If POUcnt > 0 Then
        If APIED.POUSED <> POUcnt Then
          APIED.POUSED = POUcnt
        End If
      End If
      fpUndist = Round(fpInvTotal.DoubleValue - fpDist.DoubleValue)
      fpInvNum = APIED.InvNum
      Close APEditFile
      EMode = True
      fpCode = 0
      fpDacrec = 0
      fpPODistNum = 0
      Undolok OldRec
    Else
      MsgBox "Record Is Being Edited By Another User.", vbOKOnly, "Record Unavailable"
      RecNum = OldRec
      Close APEditFile
    End If
    SetScreen
  Else
    MsgBox "Posting Is In Progress, Editing Not Allowed At This Time.", vbOKOnly, "Canceled"
    frmInvProcessMenu.Show
    Unload frmInvEnterEdit
  End If
End Function

Private Function EDCheckforDupInv()
  Dim GotOne As Boolean, DistRecLEn As Integer, LedgerRecLen As Integer
  Dim VRecNum As Integer, InvNum As String, NextTrans As Long
  Dim VendorFile As Integer, NumVRecs As Integer
  Dim APLedgerFile As Integer, NumLdgTran As Long
  GotOne = False

  ReDim TVendor(1) As VendorRecType
  ReDim TAPLedgerRec(1) As APLedger81RecType
  ReDim TAPDistRec(1) As APDistRecType

  DistRecLEn = Len(TAPDistRec(1))
  LedgerRecLen = Len(TAPLedgerRec(1))
  fpcboVendName.col = 2
  VRecNum = fpcboVendName.ColText
  InvNum$ = QPTrim$(fpInvNum)

  OpenVendorFile VendorFile, NumVRecs
  Get VendorFile, VRecNum, TVendor(1)
  Close VendorFile

  OpenAPLedgerFile APLedgerFile, NumLdgTran&, LedgerRecLen
  NextTrans& = TVendor(1).FrstTran
  Do Until NextTrans& = 0
    Get APLedgerFile, NextTrans&, TAPLedgerRec(1)
    'TAPLedgerRec(1)
    If TAPLedgerRec(1).TRCode = 1 And QPTrim$(TAPLedgerRec(1).DOCNum) = InvNum$ Then
      GotOne = True
      Exit Do
    End If
    NextTrans& = TAPLedgerRec(1).NextTrans
  Loop

  Close APLedgerFile

  EDCheckforDupInv = GotOne

End Function


Private Sub fpDebAmt_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    cmdAddDist.SetFocus
  End If
End Sub

Private Sub fpDueDate_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpPostDate.SetFocus
  End If
End Sub

Private Sub fpInvAmt_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Or KeyCode = vbKeyRight Then
    fpManPO.SetFocus
  End If
End Sub
Private Sub fpManPO_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fptxtDesc.SetFocus
  End If
End Sub
Private Sub fpInvDate_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpDueDate.SetFocus
  End If
End Sub

Private Sub fpInvDate_LostFocus()
  fpDueDate = fpInvDate
  fpPostDate = fpInvDate
End Sub

Private Sub fpInvNum_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    If POTabStop = True Then
      cmdPOList.SetFocus
    Else
      fpInvAmt.SetFocus
    End If
  End If
End Sub

Private Sub fpInvNum_LostFocus()
  Dim APEditFile As Integer, NumEdTrans As Integer, CheckCnt As Integer
  Dim NumofEditRecs As Integer
  If fpcboVendName.ListIndex <> -1 Then
  If Not DupInvDef Then
    If EDCheckforDupInv Then
      MsgBox "Duplicate Invoice Number, Please Try Again.", vbOKOnly, "Duplicate Invoice"
      fpInvNum.SetFocus
    End If
  End If
  OpenAPEditFile APEditFile, NumEdTrans
  NumofEditRecs = LOF(APEditFile) / Len(APIED)
  For CheckCnt = 1 To NumofEditRecs
    Get APEditFile, CheckCnt, APIED           'write it
    If APIED.DelFlag = 0 Then
      fpcboVendName.col = 2
      If Not DupInvDef Then
      If QPTrim(APIED.InvNum) = fpInvNum And EMode = False And Val(APIED.VRecNum) = Val(fpcboVendName.ColText) Then
        MsgBox "Duplicate Invoice Number, Please Try Another.", vbOKOnly, "Duplicate Invoice"
        fpInvNum.SetFocus
      End If
      End If
    End If
  Next CheckCnt
  Close APEditFile
  End If

End Sub


Private Sub fpPostDate_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpcboPSL.SetFocus
  End If
End Sub

Private Sub fptxtDesc_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpcboTaxable.SetFocus
  End If
End Sub

Private Sub fptxtPo_Change()
  If Len(fptxtPo) <> 0 Then
    fpcboVendName.Enabled = False
  End If
End Sub

Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub vaSpread1_DblClick(ByVal col As Long, ByVal Row As Long)
  Dim TempAcct As String
  Dim TempCol As Long, TempRow As Long
  TempRow = Row
  TempCol = col
  If TempRow > 0 Then
    vaSpread1.Row = TempRow
    vaSpread1.col = 4
    TempAcct = QPTrim(vaSpread1.Text)
    If vaSpread1.Text <> "" Then
      If fpcboAcctNumNa.ListIndex <> -1 Or fpDebAmt <> 0 Then
        If MsgBox("Do You Wish To Abandon Current Distribution?, 'Yes' or 'No' Complete Distribution Entry.", vbYesNo, "Clear??") = vbNo Then
          cmdAddDist.SetFocus
          Exit Sub
        Else
          fpcboAcctNumNa.ListIndex = -1
          fpDebAmt = 0
        End If
      End If
    
        fpcboAcctNumNa.SearchText = QPStrip(TempAcct)
        fpcboAcctNumNa.Action = 0
        If fpcboAcctNumNa.SearchIndex <> -1 Then
          fpcboAcctNumNa.ListIndex = fpcboAcctNumNa.SearchIndex
        End If
          vaSpread1.col = 1
          fpPODistNum = vaSpread1.Text
          vaSpread1.col = 2
          fpcboAcctNumNa.col = 0
          fpcboAcctNumNa.ColText = vaSpread1.Text
          fpDacrec = vaSpread1.Text
          vaSpread1.col = 3
          fpCode = vaSpread1.Text
          vaSpread1.col = 4
          fpcboAcctNumNa.col = 1
          fpcboAcctNumNa.ColText = vaSpread1.Text
          vaSpread1.col = 5
          fpcboAcctNumNa.col = 2
          fpcboAcctNumNa.ColText = vaSpread1.Text
          vaSpread1.col = 6
          fpDebAmt = vaSpread1.Text
          fpDist = (fpDist.DoubleValue - fpDebAmt.DoubleValue)
          fpUndist = Round(fpInvTotal.DoubleValue - fpDist.DoubleValue)
          'vaSpread1.ClearRange TempCol, TempRow, 4, TempRow, True
          vaSpread1.DeleteRows TempRow, 1
          If fpPODistNum > 0 And fpCode = QPTrim("T") Then
            fpcboAcctNumNa.Enabled = False
            fpDebAmt.SetFocus
          Else
            fpcboAcctNumNa.SetFocus
          End If
    End If
  End If
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
   ' Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
   ' Me.SetFocus
  End If
End Sub
'Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'  If KeyCode = vbKeyReturn Then
'    SendKeys "{Tab}"
'    DoEvents
'    KeyCode = 0
'  End If
'End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
'    Case vbKeyReturn:
'      SendKeys "{Tab}"
'      KeyCode = 0
'      DoEvents
'    Case vbKeyUp:
'      SendKeys "+{Tab}"
'      KeyCode = 0
'      DoEvents
    Case vbKeyEscape:
      cmdExit_Click
      KeyCode = 0
      DoEvents
    Case vbKeyF10:
      cmdSave_Click
      KeyCode = 0
      DoEvents
    Case vbKeyF9:
      cmdAddDist_Click
      KeyCode = 0
      DoEvents
    Case vbKeyF2:
      cmdNew_Click
      KeyCode = 0
      DoEvents
    Case vbKeyF4:
      cmdEdit_Click
      KeyCode = 0
      DoEvents
    Case vbKeyF5:
      cmdList_Click
      KeyCode = 0
      DoEvents
    Case vbKeyF3:
      cmdDelete_Click
      KeyCode = 0
      DoEvents
    Case vbKeyF6:
      cmdDelDist_Click
      KeyCode = 0
      DoEvents
    Case vbKeyF7:
      cmdNewVend_Click
      KeyCode = 0
      DoEvents
    Case vbKeyF8:
      cmdPOList_Click
      KeyCode = 0
      DoEvents
    Case Else:
  End Select
  DoEvents
End Sub
'This is to fix spreadsheet for various resolutions
Public Function Fixspread()
'    Select Case screenW
'      Case 1280
'      If Screen.TwipsPerPixelX <> 12 Then
'        coladj = 28
'        vaSpread1.RowHeight(-1) = 22
'        vaSpread1.RowHeight(0) = 22
'      Else
'        coladj = 20.5
'        vaSpread1.RowHeight(-1) = 18
'        vaSpread1.RowHeight(0) = 18
'      End If
'      Case 1152
'      If Screen.TwipsPerPixelX <> 12 Then
'        coladj = 24.5
'        vaSpread1.RowHeight(0) = 18.5
'        vaSpread1.RowHeight(-1) = 18.5
'      Else
'        coladj = 17.4
'        vaSpread1.RowHeight(0) = 15
'        vaSpread1.RowHeight(-1) = 15
'      End If
'      Case 1024
'      If Screen.TwipsPerPixelX <> 12 Then
'        coladj = 20.5
'        vaSpread1.RowHeight(0) = 17.5
'        vaSpread1.RowHeight(-1) = 17.5
'      Else
'        coladj = 14.45
'      End If
'      Case 800
'        coladj = 14
'        vaSpread1.Font.Size = 10
'        vaSpread1.RowHeight(-1) = 12.2
'      Case Else
'        'don't worry be happpy
'    End Select
'    vaSpread1.ColWidth(-1) = vaSpread1.ColWidth(-1) + coladj
     vaSpread1.FontSize = 8
End Function
Private Function SetScreen()
  If EMode = False Then  'This is in New Mode (But Dale calls this Add Mode)
    cmdNew.Enabled = False
    cmdEdit.Enabled = True
    lblNew.Visible = True
    lblEdit.Visible = False
    fpcboVendName.Enabled = True
  Else               'This is in Edit Mode
    fpcboVendName.Enabled = False
    cmdNew.Enabled = True
    cmdEdit.Enabled = False
    cmdDelete.Enabled = True
    lblNew.Visible = False
    lblEdit.Visible = True
  End If
End Function
Private Sub EdorNewEntry()
'sets screen for New when first starts -changes if selects from list at that time
'this gets record number off to correct start
'also turns on recloc for opening procedure to show list if needed
  Dim APEditFile As Integer, NumEdTrans As Integer, Rec As Integer
  Dim FileName As String, EdLen As Integer
  RecLok = False
  OpenAPEditFile APEditFile, NumEdTrans
  If NumEdTrans > 0 Then
    For Rec = 1 To NumEdTrans
      Get APEditFile, Rec, APIED
      If APIED.DelFlag <> True Then
        'RecNum = Rec
        'EMode = True
        RecLok = True
        Exit For
      End If
    Next
  End If
  
  Close APEditFile
  If RecLok = True Then
    RecNum = NumEdTrans + 1
  Else
    RecNum = 1
  End If
    EMode = False
    SetScreen
    fpInvDate.Text = Format(Now, "mm/dd/yyyy")
    fpcboVendName.ListIndex = -1
    fpInvNum = ""
    fptxtPo = ""
    If PSLDef = True Then
      fpcboPSL.ListIndex = 0
    Else
      fpcboPSL.ListIndex = 1
    End If
    If AutoDistFlag = True Then
      fpcboTaxable.ListIndex = 0
    Else
      fpcboTaxable.ListIndex = 1
    End If
    fpcbo1099.ListIndex = 1
    fpcboAcctNumNa.ListIndex = -1
    
    fpCTax = 0
    fpDist = 0
    fpDueDate.Text = ""
    fpPostDate.Text = ""
'***** spreadsheet do Not have to set blank fields on load ..
    fpDebAmt = 0
    fptxtDesc = ""
    fpInvTotal = 0
    fpUndist = 0
    fpPODistNum = 0
    fpDacrec = 0
    fpCode = "N"
    fpAPLegNum = 0
  
End Sub

Private Sub cmdExit_Click()
  If Changed = False Then
    Undolok RecNum
    frmInvProcessMenu.Show
    Unload frmInvEnterEdit
    Call MainLog("Exit InvEnterEdit")
  Else
    If MsgBox("Changes Have Been Made to the Current Record." & Chr(13) & Chr(13) & "                          Select OK to Abandon," & Chr(13) & Chr(13) & "       or Cancel to Remain on Entry/Edit Screen.", vbOKCancel, "Abandon Changes?") = vbOK Then
      Undolok RecNum
      frmInvProcessMenu.Show
      Unload frmInvEnterEdit
      Call MainLog("Exit InvEnterEdit")
    End If
  End If
End Sub
Public Sub ClearScn()
    fpInvDate.Text = TInvDate
    fpcboVendName.ListIndex = -1
    fpInvNum = ""
    fpInvAmt = 0
    fptxtPo = ""
    fpManPO = ""
    fpcbo1099.ListIndex = -1
    fpcboAcctNumNa.ListIndex = -1
    If PSLDef = True Then
      fpcboPSL.ListIndex = 0
    Else
      fpcboPSL.ListIndex = 1
    End If
    If AutoDistFlag = True Then
      fpcboTaxable.ListIndex = 0
    Else
      fpcboTaxable.ListIndex = 1
    End If
    fpCTax = 0
    fpDist = 0
    fpDueDate.Text = "" 'Format(Now, "mm/dd/yyyy")
    fpPostDate.Text = "" 'Format(Now, "mm/dd/yyyy")
    fpDebAmt = 0
    fptxtDesc = ""
    fpInvTotal = 0
    fpUndist = 0
    vaSpread1.ClearRange 1, 1, 6, 36, True
    fpPODistNum = 0
    fpDacrec = 0
    fpCode = "N"
    fpAPLegNum = 0
End Sub

Private Sub fpcbo1099_LostFocus()
  If EMode = False And fpcboVendName.ListIndex <> -1 Then
    Check4DefDist
  End If
End Sub
Private Sub fpcbo1099_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcbo1099.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcbo1099.ListIndex = -1
    fpcbo1099.Action = ActionClearSearchBuffer
  End If
  If fpcbo1099.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      If fpcboAcctNumNa.Enabled = True Then
        fpcboAcctNumNa.SetFocus
      Else
        fpDebAmt.SetFocus
      End If
        KeyCode = 0
        DoEvents
    Else
      If KeyCode = vbKeyUp Then
        fpcboPSL.SetFocus
        KeyCode = 0
        DoEvents
      End If
    End If
  End If

End Sub
Private Sub fpcboPSL_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboPSL.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcboPSL.ListIndex = -1
    fpcboPSL.Action = ActionClearSearchBuffer
  End If
  If fpcboPSL.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
        fpcbo1099.SetFocus
        KeyCode = 0
        DoEvents
    Else
      If KeyCode = vbKeyUp Then
        fpPostDate.SetFocus
        KeyCode = 0
        DoEvents
      End If
    End If
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
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
        fpDebAmt.SetFocus
        KeyCode = 0
        DoEvents
    Else
      If KeyCode = vbKeyUp Then
        fpcbo1099.SetFocus
        KeyCode = 0
        DoEvents
      End If
    End If
  End If

End Sub

Private Sub fpcboAcctNumNa_LostFocus()
  fpcboAcctNumNa.Action = ActionClearSearchBuffer
  
End Sub

Private Sub fpcboTaxable_Click()
  RedoTax
End Sub
Private Sub fpcboTaxable_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboTaxable.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcboTaxable.ListIndex = -1
    fpcboTaxable.Action = ActionClearSearchBuffer
  End If
  If fpcboTaxable.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
        fpInvDate.SetFocus
        KeyCode = 0
        DoEvents
    Else
      If KeyCode = vbKeyUp Then
        fpInvAmt.SetFocus
        KeyCode = 0
        DoEvents
      End If
    End If
  End If
End Sub


Private Sub fpcboVendName_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboVendName.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcboVendName.ListIndex = -1
    fpcboVendName.Action = ActionClearSearchBuffer
  End If
  If fpcboVendName.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
        fpInvNum.SetFocus
        KeyCode = 0
        DoEvents
    Else
      If KeyCode = vbKeyUp Then
        fpDebAmt.SetFocus
        KeyCode = 0
        DoEvents
      End If
    End If
  End If
End Sub

Private Sub fpcboVendName_LostFocus()
  fpcboVendName.Action = ActionClearSearchBuffer
End Sub
Private Sub RedoTax()
  If GotTaxFile = True And fpcboTaxable.ListIndex = 0 Then
    fpSTax = Round#(fpInvAmt * (STATAX# / 100))
    fpCTax = Round#(fpInvAmt * (CTYTAX# / 100))
    fpInvTotal = (fpInvAmt.DoubleValue + (fpSTax.DoubleValue + fpCTax.DoubleValue))
  Else
    fpSTax = 0
    fpCTax = 0
    fpInvTotal = fpInvAmt
    fpcboTaxable.ListIndex = 1
  End If
  fpUndist = Round(fpInvTotal.DoubleValue - fpDist.DoubleValue)
End Sub
Private Sub fpInvAmt_LostFocus()
  RedoTax
End Sub


Private Sub mnuPrnScn_Click()
  PrintForm
End Sub
Private Sub InitTaxInfo()
  Dim APInvTax(1)  As APInvTaxRecType
  Dim InvTaxFileNum As Integer
  OpenInvTaxFile InvTaxFileNum
  If LOF(InvTaxFileNum) > 0 Then
    Get InvTaxFileNum, 1, APInvTax(1)
      If APInvTax(1).InvTax(1).TaxAmt > 0 Then
        GotTaxFile = True
        STATAX# = APInvTax(1).InvTax(1).TaxAmt
        StaTaxFlag = True
      End If
      If APInvTax(1).InvTax(2).TaxAmt > 0 Then
        CTYTAX# = APInvTax(1).InvTax(2).TaxAmt
        GotTaxFile = True
        CtyTaxFlag = True
      End If
    
    AutoDistFlag = APInvTax(1).AUTODIST = "Y"
  End If
  Close InvTaxFileNum
  If GotTaxFile = False Then
    fpcboTaxable.ListIndex = 1
    fpSTax = 0#
    fpSTax.ControlType = ControlTypeReadOnly
    fpCTax = 0#
    fpCTax.ControlType = ControlTypeReadOnly
  Else
    fpcboTaxable.ListIndex = 0
  End If
'Below from Old But was remarked out in old(PS 3-15-02)
  'IF GotTaxFile THEN
  '  IF StaTaxFlag THEN
  '    TaxActualAdj = TaxActualAdj + 1
  '    AcctRec = FindAcct(QPTrim$(InvTaxRec(1).InvTax(1).AcctNo))
  '    REDIM PRESERVE TaxDist(1 TO TaxActualAdj)  AS DistSumType
  '    TaxDist(TaxActualAdj).DistAcctNum = InvTaxRec(1).InvTax(1).AcctNo
  '    TaxDist(TaxActualAdj).AcctTitle = GetAcctTitle$(AcctRec)
  '  END IF
  '  IF CtyTaxFlag THEN
  '    TaxActualAdj = TaxActualAdj + 1
  '    AcctRec = FindAcct(QPTrim$(InvTaxRec(1).InvTax(2).AcctNo))
  '    REDIM PRESERVE TaxDist(1 TO TaxActualAdj)  AS DistSumType
  '    TaxDist(TaxActualAdj).DistAcctNum = InvTaxRec(1).InvTax(2).AcctNo
  '    TaxDist(TaxActualAdj).AcctTitle = GetAcctTitle$(AcctRec)
  '  END IF
  'END IF
End Sub

Private Sub fpcboVendName_Click()
  Dim VendorFile As Integer, NumVRecs As Integer, VRecNum As Integer
  If EMode = True Then
    Exit Sub
  End If
  If fpcboVendName.ListIndex <> -1 Then
    OpenVendorFile VendorFile, NumVRecs
    fpcboVendName.col = 2
    VRecNum = Val(fpcboVendName.ColText)
    Get VendorFile, VRecNum, Vendor
    If Vendor.Get1099 = "Y" Then
      fpcbo1099.ListIndex = 0
    Else
      fpcbo1099.ListIndex = 1
    End If
    
    Close
  End If
'need to check for entry change msg to change then clear screen
End Sub
Private Sub cmdPOList_Click()
  If frmInvEnterEdit.fpcboVendName.ListIndex = -1 Then
    MsgBox "You Must Select A Vendor First", vbOKOnly, "No Vendor"
    fpcboVendName.SetFocus
    Exit Sub
  End If
  If QPTrim(fptxtPo) <> "" Then
    If MsgBox("Abandon Current PO?", vbYesNo, "Change PO") = vbNo Then
      fpInvAmt.SetFocus
      Exit Sub
    End If
  End If
  frmPOs.Loadpos
  If frmPOs.fplstPOs.ListCount > 0 Then
    frmPOs.Show 1, Me
  Else
    Unload frmPOs
  End If
  fpInvAmt.SetFocus
End Sub
Private Sub cmdNew_Click()
  Dim InvBusy As Boolean
  InvBusy = False
  If Exist("APIED.DAT") Then InvBusy = GetAttr("APIED.DAT") And vbReadOnly
  If Not InvBusy Then
    If Changed = True Then
      If MsgBox("Changes Have Been Made to the Current Record." & Chr(13) & "Select OK to Abandon," & Chr(13) & "or Cancel to Remain on Current Record.", vbOKCancel, "Abandon Changes?") = vbCancel Then
        fpcboAcctNumNa.SetFocus
        Exit Sub
      End If
    End If
    Undolok RecNum
    NextNew
  Else
    MsgBox "Posting Is In Progress, Editing Not Allowed At This Time.", vbOKOnly, "Canceled"
    frmInvProcessMenu.Show
    Unload frmInvEnterEdit
  End If
End Sub

Private Sub cmdAddDist_Click()
  Dim DistRecLEn As Integer, APDistFile As Integer, NumDistRecs As Long
  OpenAPDistFile APDistFile, NumDistRecs&, DistRecLEn
  Close APDistFile
  If VerifyEntered = False Then
    MsgBox "The Information In The Top Section Must Be Completed Before Adding Distributions.", vbOKOnly, "Cash Disbursement"
  Else
    If fpcboAcctNumNa.Text <> "" And fpDebAmt.DoubleValue <> 0 Then
      If vaSpread1.DataRowCnt < 36 Then
       If fpPODistNum <= NumDistRecs& Then
        vaSpread1.Row = vaSpread1.DataRowCnt + 1
        vaSpread1.col = 1
        vaSpread1.Text = fpPODistNum
        vaSpread1.col = 2
        fpcboAcctNumNa.col = 0
        vaSpread1.Text = fpcboAcctNumNa.ColText
        vaSpread1.col = 3
        vaSpread1.Text = fpCode
        vaSpread1.col = 4
        fpcboAcctNumNa.col = 1
        vaSpread1.Text = fpcboAcctNumNa.ColText
        vaSpread1.col = 5
        fpcboAcctNumNa.col = 2
        vaSpread1.Text = fpcboAcctNumNa.ColText
        vaSpread1.col = 6
        fpDist = (fpDebAmt.DoubleValue + fpDist.DoubleValue)
        vaSpread1.Text = fpDebAmt
        fpcboAcctNumNa.ListIndex = -1
        fpDebAmt = 0
        fpUndist = Round(fpInvTotal.DoubleValue - fpDist.DoubleValue)
      Else
        MsgBox "Error with this distribution, Delete or Re-Enter or Select PO again.", vbOKOnly, "Distribution Error"
      End If
     Else
      MsgBox "Only 36 Distributions Per Invoice.", vbOKOnly, "Limit Reached"
     End If
    Else
      MsgBox "The Account and Amount Must Be Entered Before Adding To The Distribution List.", vbOKOnly, "Add Distribution Denied"
    End If
    fpcboAcctNumNa.Enabled = True
    fpcboAcctNumNa.SetFocus
  End If
End Sub
Private Sub cmdEdit_Click()
  If Changed = True Then
    If MsgBox("Changes Were Made to the Current Information on Screen and Not Saved." & Chr(13) & "Select OK to View List," & Chr(13) & "or Cancel to Remain on Current Record.", vbOKCancel, "View List?") = vbCancel Then
      fpcboAcctNumNa.SetFocus
      Exit Sub
    End If
  End If
  Undolok RecNum
  NextNew
  If Check4Trans = True Then
    frmInvListing.Show 1, frmInvEnterEdit
    If EMode = True Then
      SetScreen
      
      fpcboAcctNumNa.SetFocus
    End If
  Else
    MsgBox "No Entries To Display.", vbOKOnly, "No Entries"
    fpcboAcctNumNa.SetFocus
  End If

End Sub

Private Sub cmdList_Click()
  If Changed = True Then
    If MsgBox("Changes Were Made to the Current Information on Screen and Not Saved." & Chr(13) & "Select OK to View List," & Chr(13) & "or Cancel to Remain on Current Record.", vbOKCancel, "View List?") = vbCancel Then
      fpcboAcctNumNa.SetFocus
      Exit Sub
    End If
  End If
  Undolok RecNum
  NextNew
  If Check4Trans = True Then
    frmInvListing.Show 1, frmInvEnterEdit
    If EMode = True Then
      SetScreen
      fpcboAcctNumNa.SetFocus
    End If
  Else
    MsgBox "No Entries To Display.", vbOKOnly, "No Entries"
    fpcboAcctNumNa.SetFocus
  End If
End Sub
Private Function Check4Trans()
  Dim APEditFile As Integer, NumEdTrans As Integer
  Dim cnt As Integer, Good As Integer
  Good = 0
  If Exist("APIED.dat") Then
    OpenAPEditFile APEditFile, NumEdTrans
    If NumEdTrans > 0 Then
      For cnt = 1 To NumEdTrans
        Get APEditFile, cnt, APIED
        If APIED.DelFlag = 0 Then
          Good = Good + 1
        End If
      Next
    Else
      Check4Trans = False
    End If
  Else
    Check4Trans = False
  End If
  If Good > 0 Then
    Check4Trans = True
  Else
    Check4Trans = False
  End If
 Close APEditFile
 End Function
''Private Sub CHK4POUSED(TRNUM As Long)
''  Dim APEditFile As Integer, NumEdTrans As Integer
''  Dim cnt As Integer, Good As Integer, CntDist As Integer
''  Good = 0
''  If Exist("APIED.dat") Then
''    OpenAPEditFile APEditFile, NumEdTrans
''    If NumEdTrans > 0 Then
''      For cnt = 1 To NumEdTrans
''        Get APEditFile, cnt, APIED
''          If APIED.POAPLRecNum = TRNUM Then  'Used on another invoice this edit
''            If APIED.DELFLAG = 0 Then
''              Last = UBound(APIED.Dist)
''              For cnt = 1 To Last
''                If APIED.Dist(cnt).DISTNUM > 0 Then
''          APIED.Dist(cnt).DACODEvaSpread1.Row = vaSpread1.DataRowCnt + 1
''          vaSpread1.col = 1
''          vaSpread1.Text = APIED.Dist(cnt).DISTNUM
''          vaSpread1.col = 2
''          vaSpread1.Text = APIED.Dist(cnt).DACREC
''          vaSpread1.col = 3
''          vaSpread1.Text =
''          vaSpread1.col = 4
''          vaSpread1.Text = APIED.Dist(cnt).DACN
''          vaSpread1.col = 5
''          vaSpread1.Text = APIED.Dist(cnt).DACNM
''          vaSpread1.col = 6
''          vaSpread1.Text = APIED.Dist(cnt).DAMT
''          fpDist = (fpDist.DoubleValue + APIED.Dist(cnt).DAMT)
''        Else
''          Exit For
''        End If
''        NextTrans& = APDistRec(1).NextDist
''      Close APEditFile
''          End If
''         End If
''
''        End If
''      Next
''    Else
''      Check4Trans = False
''    End If
''  Else
''    Check4Trans = False
''  End If
''  If Good > 0 Then
''    Check4Trans = True
''  Else
''    Check4Trans = False
''  End If
'' Close APEditFile
''
''End Sub
Private Sub Check4DefDist()
  Dim Amt As String, DistAmt As Double, d As Integer
  Dim AmtDist As Double, DistPct As Double, ACCTNO As Integer
  Dim VendorFile As Integer, NumVRecs As Integer, VRecNum As Integer
  Dim APDefDistFile As Integer, NumDefRecs As Integer, DefRecLen As Integer
  ReDim vndr(1) As VendorRecType
  OpenVendorFile VendorFile, NumVRecs
  fpcboVendName.col = 2
  VRecNum = fpcboVendName.ColText
  Get VendorFile, VRecNum, vndr(1)

  If vndr(1).DefDist > 0 Then
    Amt$ = InputBox$("Please Enter Total Amount to Distribute", "Default Distribution")
    If Val(Amt$) = 0 Then
      Close
      Exit Sub
    Else
      
      DistAmt# = Val(Amt$)
      If DistAmt# <> 0 Then
        Amt$ = Using("#######.##", DistAmt#)
        OpenDefDistFile DefRecLen, APDefDistFile, NumDefRecs
        Get APDefDistFile, vndr(1).DefDist, DefDist
        'If Not ISEven(DistAmt#) Then
        DistPct# = DefDist.DefDist(1).DefPct
        If DistPct# < 100 Then
          If Not ISEven(Amt$) Then
            DistAmt# = Round(DistAmt# - 0.01)
          End If
        End If
        vaSpread1.ClearRange 1, 1, 6, 36, True
        fpUndist = Round(fpInvTotal)
        fpDist = 0
        'ReDim VdefDist(1) As VendorDefDistRecType
        For d = 1 To 8
          vaSpread1.Row = d
          If Len(QPTrim$(DefDist.DefDist(d).DefAcct)) > 0 Then
            ACCTNO = AcctFind(DefDist.DefDist(d).DefAcct)
            'GotDef = d       'set the flag to the last field
            vaSpread1.col = 1
            vaSpread1.Text = 0
            vaSpread1.col = 2
            vaSpread1.Text = ACCTNO
            vaSpread1.col = 3
            vaSpread1.Text = "N"
            vaSpread1.col = 4
            vaSpread1.Text = DefDist.DefDist(d).DefAcct
            vaSpread1.col = 5
            vaSpread1.Text = DefDist.DefDist(d).DefAcctName
            DistPct# = DefDist.DefDist(d).DefPct * 0.01 'Round#((DefDist.DefDist(d).DefPct * 0.01))
            AmtDist# = Round#(DistPct# * DistAmt#)
            vaSpread1.col = 6
            vaSpread1.Text = Str$(AmtDist#)
            fpDist = fpDist.DoubleValue + AmtDist#
            fpUndist = Round(fpInvTotal.DoubleValue - fpDist.DoubleValue)
          Else
            Exit For
          End If
        Next
      End If
    End If
  End If

  Close

End Sub

Public Sub GetPOInfo(TempDist As Long, TempLeg As Long)
  Dim POCnt As Integer, NextTrans As Long, NumDistRecs As Long
  Dim VendorFile As Integer, NumVRecs As Integer, tempstr As String
  Dim APLedgerFile As Integer, NumTrans As Long, LdRecLen As Integer
  Dim AcctFileNum As Integer, NumAccts As Integer, vrec As Integer
  Dim DistRecLEn As Integer, POLcnt As Integer
  Dim APDistFile As Integer
  Dim AcctFile As GLAcctRecType
  Dim APDistRec(1) As APDistRecType
  Dim APLedgerRec(1) As APLedger81RecType
  LdRecLen = Len(APLedgerRec(1))
  DistRecLEn = Len(APDistRec(1))
    OpenVendorFile VendorFile, NumVRecs
    OpenAPLedgerFile APLedgerFile, NumTrans, LdRecLen

    fpcboVendName.col = 2
    vrec = fpcboVendName.ColText
    Get VendorFile, vrec, Vendor
    NextTrans& = TempLeg
    'Do Until NextTrans& = 0
      Get APLedgerFile, NextTrans&, APLedgerRec(1)
      If APLedgerRec(1).TRCode = 4 Then
        POCnt = POCnt + 1
               
      End If
    'NextTrans& = APLedgerRec(1).NextTrans
  
  'Loop
  POLcnt = 0
  If POCnt <> 0 Then
  fpInvAmt = APLedgerRec(1).Amt
  fptxtPo = APLedgerRec(1).PONum
  RedoTax
  fpAPLegNum = TempLeg
  vaSpread1.ClearRange 1, 1, 6, 36, True
  fpDist = 0
  fpUndist = 0
  OpenAcctFile AcctFileNum, NumAccts
  OpenAPDistFile APDistFile, NumDistRecs&, DistRecLEn
  NextTrans& = TempDist
  Do Until NextTrans& = 0
    Get APDistFile, NextTrans&, APDistRec(1)
    If Val(APDistRec(1).DistStat) = 0 Then
      vaSpread1.Row = vaSpread1.DataRowCnt + 1
      If Len(QPTrim$(APDistRec(1).DistAcctNum)) > 0 Then
        If APDistRec(1).DistAcctRec > 0 Then
        vaSpread1.col = 1
        vaSpread1.Text = NextTrans&
        vaSpread1.col = 2
        vaSpread1.Text = APDistRec(1).DistAcctRec
        vaSpread1.col = 3
        If APDistRec(1).DistStat = "L" Then
          vaSpread1.Text = "L"
        Else
          vaSpread1.Text = "T"
        End If
        vaSpread1.col = 4
        vaSpread1.Text = APDistRec(1).DistAcctNum
        vaSpread1.col = 5
        Get AcctFileNum, APDistRec(1).DistAcctRec, AcctFile
        vaSpread1.Text = QPTrim(AcctFile.Title)
        vaSpread1.col = 6
        vaSpread1.Text = APDistRec(1).DistAmt
        fpDist = fpDist.DoubleValue + APDistRec(1).DistAmt
        fpUndist = Round(fpInvTotal.DoubleValue - fpDist.DoubleValue)
        POLcnt = POLcnt + 1
        End If
      Else
        Exit Do
      End If
    End If
    NextTrans& = APDistRec(1).NextDist
  Loop
  End If
  appolines = POLcnt

  Close
End Sub
Private Function Changed()
  Dim APEditFile As Integer, NumEdTrans As Integer, Rec As Integer
  Dim FileName As String, EdLen As Integer
  Dim cnt As Integer
  If EMode = False Then
    If fpcboVendName.ListIndex <> -1 Then
      Changed = True
      Exit Function
    Else
      If Len(QPTrim(fpInvNum)) <> 0 Then
        Changed = True
        Exit Function
      Else
        If Len(QPTrim(fptxtPo)) <> 0 Then
          Changed = True
          Exit Function
        Else
        If fpInvAmt <> 0 Then
          Changed = True
          Exit Function
        Else
          If fpDist <> 0 Then
            Changed = True
            Exit Function
          Else
            If fpUndist <> 0 Then
              Changed = True
              Exit Function
            Else
              Changed = False
            End If
          End If
          End If
        End If
      End If
    End If
  Else
    OpenAPEditFile APEditFile, NumEdTrans
    Get APEditFile, RecNum, APIED
    
    If fpInvDate <> Format(DateAdd("d", (APIED.InvDate), "12-31-1979"), "mm/dd/yyyy") Then
      Changed = True
      Close APEditFile
      Exit Function
    Else
    If fpDueDate <> Format(DateAdd("d", (APIED.DueDate), "12-31-1979"), "mm/dd/yyyy") Then
      Changed = True
      Close APEditFile
      Exit Function
    Else
    If fpPostDate <> Format(DateAdd("d", (APIED.DISTDATE), "12-31-1979"), "mm/dd/yyyy") Then
      Changed = True
      Close APEditFile
      Exit Function
    Else

      If QPTrim(fptxtDesc) <> QPTrim(APIED.INVDESC) Then
        Changed = True
        Close APEditFile
        Exit Function
      Else
        If QPTrim(fpInvNum) <> QPTrim(APIED.InvNum) Then
          Changed = True
          Close APEditFile
          Exit Function
        Else
          If fpInvAmt.DoubleValue() <> APIED.InvAmt Then
            Changed = True
            Close APEditFile
            Exit Function
          Else
            If Left$(fpcboPSL.Text, 1) <> APIED.PSLFlag Then
              Changed = True
              Close APEditFile
              Exit Function
            Else
              If Left$(fpcbo1099.Text, 1) <> APIED.Get1099 Then
                Changed = True
                Close APEditFile
                Exit Function
              Else
              If Left$(fpcboTaxable.Text, 1) <> APIED.TAXYN Then
                Changed = True
                Close APEditFile
              Else
              If fpSTax <> APIED.STAXAMT Then
                Changed = True
                Close APEditFile
              Else
              If fpCTax <> APIED.CTAXAMT Then
                Changed = True
                Close APEditFile
              Else
                Changed = False
              End If
              End If
              End If
              End If
              End If
              
              End If
            End If
          End If
        End If
      End If
    End If
    If fpDebAmt <> 0 Then
      Changed = True
      Close APEditFile
      Exit Function
    End If
    If fpcboAcctNumNa.ListIndex <> -1 Then
      Changed = True
      Close APEditFile
      Exit Function
    Else
      For cnt = 1 To 36
        vaSpread1.Row = cnt
        vaSpread1.col = 2
        If Val(vaSpread1.Text) = APIED.Dist(cnt).DACREC Then
          If Val(vaSpread1.Text) = 0 Then
            Changed = False
            Exit For
          Else
            vaSpread1.col = 1
            If vaSpread1.Text = APIED.Dist(cnt).DISTNUM Then
              vaSpread1.col = 3
              If vaSpread1.Text = APIED.Dist(cnt).DACODE Then
                vaSpread1.col = 4
                If vaSpread1.Text = APIED.Dist(cnt).DACN Then
                  vaSpread1.col = 5
                  If vaSpread1.Text = APIED.Dist(cnt).DACNM Then
                   vaSpread1.col = 6
                    If vaSpread1.Text = APIED.Dist(cnt).DAMT Then
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
              Else
                Changed = True
                Exit For
              End If
            Else
              Changed = True
              Exit For
            End If
          End If
        Else
          Changed = True
          Exit For
        End If
      Next
    Close APEditFile
    End If
  End If
End Function
Private Sub cmdSave_Click()
  If Ready2Save = True Then
    If SaveInvoice = True Then
      MsgBox "Save Invoice Completed.", vbOKOnly, "Invoice Saved"
      Call NextNew
    Else
      MsgBox "Save Canceled.", vbOKOnly, "Canceled"
    End If
  Else
    MsgBox "             Save Canceled.", vbOKOnly, "Invoice Entry"
  End If
End Sub
Private Function VerifyEntered()
  If fpcboVendName.ListIndex <> -1 Then
    If Len(QPTrim(fpInvNum.Text)) > 0 Then
      If fpInvAmt <> 0 Then
        If fpcboTaxable.ListIndex <> -1 Then
          If fpcboPSL.ListIndex <> -1 Then
            If fpcbo1099.ListIndex <> -1 Then
              VerifyEntered = True
            Else
              VerifyEntered = False
              fpcbo1099.SetFocus
              Exit Function
            End If
          Else
            VerifyEntered = False
            fpcboPSL.SetFocus
            Exit Function
          End If
        Else
          VerifyEntered = False
          fpcboTaxable.SetFocus
          Exit Function
        End If
      Else
        VerifyEntered = False
        fpInvAmt.SetFocus
        Exit Function
      End If
    Else
      VerifyEntered = False
      fpInvNum.SetFocus
      Exit Function
    End If
  Else
    VerifyEntered = False
    fpcboVendName.SetFocus
    Exit Function
  End If
End Function

Private Function Ready2Save()
  Dim TempDate As Integer, TempDate2 As Integer, cnt As Integer
  Dim TempDist As Double, TempDate3 As Integer
  TempDist = 0
  'Take care of Invalid Data and Messages in this Section
  'CheckValDate is in main module to verify dates entered w/correct format
  If CheckValDate(QPTrim(fpInvDate)) = True Then
    TempDate = DateDiff("d", "12/31/1979", fpInvDate)
    If CheckValDate(QPTrim(fpDueDate)) = True Then
      TempDate2 = DateDiff("d", "12/31/1979", fpDueDate)
      If CheckValDate(QPTrim(fpPostDate)) = True Then
        TempDate3 = DateDiff("d", "12/31/1979", fpPostDate)
      Else
        MsgBox "This Date Is Not Valid. Please Correct.", vbOKOnly, "Invalid Date"
        Ready2Save = False
        fpPostDate.SetFocus
        Exit Function
      End If
    Else
      MsgBox "This Date Is Not Valid. Please Correct.", vbOKOnly, "Invalid Date"
      Ready2Save = False
      fpDueDate.SetFocus
      Exit Function
    End If
  Else
    MsgBox "This Date Is Not Valid. Please Correct.", vbOKOnly, "Invalid Date"
    Ready2Save = False
    fpInvDate.SetFocus
    Exit Function
  End If
    If Len(QPTrim(fpInvNum)) = 0 Then
      MsgBox "You May Not Save An Invoice Without An Invoice Number.", vbOKOnly, "Invoice Entries"
      Ready2Save = False
      fpInvNum.SetFocus
      Exit Function
    End If
  'Also compare dates with Hi/Lo range
    If (TempDate < LPDate) Or (TempDate > HPDate) Then
      MsgBox "This Date Is Not Within Allowable Posting Range. Please Correct or Change Setup.", vbOKOnly, "Invalid Date"
      Ready2Save = False
      fpInvDate.SetFocus
      Exit Function
    Else
      Ready2Save = True
    End If
    If (TempDate2 < LPDate) Or (TempDate2 > HPDate) Then
      MsgBox "This Date Is Not Within Allowable Posting Range. Please Correct or Change Setup.", vbOKOnly, "Invalid Date"
      Ready2Save = False
      fpDueDate.SetFocus
      Exit Function
    Else
      Ready2Save = True
    End If
    If (TempDate3 < LPDate) Or (TempDate3 > HPDate) Then
      MsgBox "This Date Is Not Within Allowable Posting Range. Please Correct or Change Setup.", vbOKOnly, "Invalid Date"
      Ready2Save = False
      fpPostDate.SetFocus
      Exit Function
    Else
      Ready2Save = True
    End If

  'Not allow Zero Total or Unequal Distritbutions
  If fpInvAmt <> 0 Then
    If fpUndist <> 0 Or fpInvTotal <> fpDist Then
      MsgBox "The Total Distributed Does Not Equal The Amount of The Invoice." & Chr$(13) & "Please Correct Before Saving.", vbOKOnly, "Invoice Entry"
      Ready2Save = False
      Exit Function
    Else
      
      For cnt = 1 To 36
        vaSpread1.col = 6
        vaSpread1.Row = cnt
        If vaSpread1.Text <> "" Then
          TempDist = Round(vaSpread1.Text + TempDist)
        Else
          Exit For
        End If
      Next
      If Round#(TempDist) <> Round#(fpDist.DoubleValue) Or Round#(TempDist) <> Round(fpInvTotal.DoubleValue) Then
        MsgBox "Totals Are Not In Balance. Please Correct.", vbOKOnly, "Invoice Entry"
        Ready2Save = False
        Exit Function
      Else
        Ready2Save = True
      End If
     End If
  Else
    MsgBox "You May Not Save An Invoice With A $0.00 Total.", vbOKOnly, "Invoice Entries"
    Ready2Save = False
  End If
End Function

 Private Function SaveInvoice()
  Dim APEditFile As Integer, NumEdTrans As Integer, cnt As Integer
  Dim NextTrans As Long, NumDistRecs As Long, POUSED As Integer
  Dim VendorFile As Integer, NumVRecs As Integer, tempstr As String
  Dim APLedgerFile As Integer, NumTrans As Long, LdRecLen As Integer
  Dim AcctFileNum As Integer, NumAccts As Integer, vrec As Integer
  Dim DistRecLEn As Integer, APDistFile As Integer
  Dim AcctFile As GLAcctRecType
  Dim APDistRec(1) As APDistRecType
  Dim APLedgerRec(1) As APLedger81RecType
  Dim InvBusy As Boolean
  InvBusy = False
  If Exist("APIED.DAT") Then InvBusy = GetAttr("APIED.DAT") And vbReadOnly
  If Not InvBusy Then
    OpenAPDistFile APDistFile, NumDistRecs&, DistRecLEn
    Close APDistFile
    OpenAPEditFile APEditFile, NumEdTrans
    If EMode = False Then
      If NumEdTrans > 0 Then
        RecNum = NumEdTrans + 1
      Else
        RecNum = 1
      End If
      TInvDate = fpInvDate
    Else
      Get APEditFile, RecNum, APIED
      TInvDate = ""
    End If
    LdRecLen = Len(APLedgerRec(1))
    DistRecLEn = Len(APDistRec(1))
    APIED.DelFlag = 0
    fpcboVendName.col = 2
    APIED.VRecNum = fpcboVendName.ColText
    fpcboVendName.col = 0
    APIED.Vendor = QPTrim(fpcboVendName.ColText)
    fpcboVendName.col = 1
    APIED.VendName = QPTrim(fpcboVendName.ColText)
    APIED.DISTDATE = DateDiff("d", "12/31/1979", fpPostDate)
    APIED.DueDate = DateDiff("d", "12/31/1979", fpDueDate)
    APIED.InvDate = DateDiff("d", "12/31/1979", fpInvDate)
    APIED.INVDESC = Trim(fptxtDesc)
    APIED.InvNum = Trim(fpInvNum)
    APIED.InvAmt = fpInvAmt.DoubleValue
    APIED.PONum = Trim(fptxtPo)
    APIED.MPONum = Trim(fpManPO)
    APIED.TAXYN = Left$(fpcboTaxable.Text, 1)
    APIED.CTAXAMT = fpCTax.DoubleValue
    APIED.STAXAMT = fpSTax.DoubleValue
    APIED.Get1099 = Left$(fpcbo1099.Text, 1)
    APIED.PSLFlag = Left$(fpcboPSL.Text, 1)
    APIED.GRANDTOT = fpInvTotal.DoubleValue
    APIED.PAYCODE = 1
    APIED.POAPLRecNum = fpAPLegNum
    For cnt = 1 To 36
      vaSpread1.Row = cnt
      vaSpread1.col = 2
      If Val(vaSpread1.Text) = 0 Then
        APIED.Dist(cnt).DISTNUM = 0
        APIED.Dist(cnt).DACREC = 0
        APIED.Dist(cnt).DACODE = ""
        APIED.Dist(cnt).DACN = ""
        APIED.Dist(cnt).DACNM = ""
        APIED.Dist(cnt).DAMT = 0
      Else
        vaSpread1.col = 1
        If vaSpread1.Text <> "0" Then
          POUSED = POUSED + 1
        End If
        If QPTrim(vaSpread1.Text) > NumDistRecs& Then
          'this must be stopped
           MsgBox "There is an ERROR with this Distribution, You should delete or re-enter these distributions", vbOKOnly, "Invalid PO Distribution"
           Close
           SaveInvoice = False
           Exit Function
        End If
        APIED.Dist(cnt).DISTNUM = QPTrim(vaSpread1.Text)
        vaSpread1.col = 2
        APIED.Dist(cnt).DACREC = QPTrim(vaSpread1.Text)
        vaSpread1.col = 3
        APIED.Dist(cnt).DACODE = QPTrim(vaSpread1.Text)
        vaSpread1.col = 4
        APIED.Dist(cnt).DACN = QPTrim(vaSpread1.Text)
        vaSpread1.col = 5
        APIED.Dist(cnt).DACNM = QPTrim(vaSpread1.Text)
        vaSpread1.col = 6
        APIED.Dist(cnt).DAMT = vaSpread1.Text
      End If
    Next
    If Len(QPTrim(fptxtPo)) <> 0 Then
     APIED.POUSED = POUSED
     APIED.POLINES = appolines
     If POUSED < appolines Then
      If MsgBox("Only Part Of The Purchase Order Was Used, The Remainder Will Still Be Available. Press OK to Continue, or Cancel to Edit This Entry.", vbOKCancel, "Continue?") = vbCancel Then
        Close APEditFile
        SaveInvoice = False
        Exit Function
      End If
      APIED.POFLAG = 2
     Else
      APIED.POFLAG = 1
     End If
    Else
      APIED.POUSED = 0
      APIED.POLINES = 0
      APIED.POFLAG = 0
    End If
    APIED.LOCKED = False
    Put APEditFile, RecNum, APIED
    Close APEditFile
    Call MainLog("Saved Inv - " + fpInvNum)
    SaveInvoice = True
  Else
    MsgBox "Posting Is In Progress, Editing Not Allowed At This Time.", vbOKOnly, "Canceled"
    frmInvProcessMenu.Show
    Unload frmInvEnterEdit
  End If
End Function
Private Sub cmdDelDist_Click()
  If vaSpread1.ActiveRow > 0 Then
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.col = 4
    If vaSpread1.Text <> "" Then
      If MsgBox("You Wish to Delete this Distribution?", vbYesNo, "Delete Distribution") = vbYes Then
        vaSpread1.col = 6
        fpDist = Round(fpDist.DoubleValue - vaSpread1.Text)
        fpUndist = Round(fpInvTotal.DoubleValue - fpDist.DoubleValue)
        
        vaSpread1.DeleteRows vaSpread1.Row, 1
        'fpcboAcctNumNa.SetFocus
      End If
    Else
      MsgBox "There Must Be A Distribution Selected to Delete.", vbOKOnly, "Selection Blank"
    End If
  End If

End Sub


