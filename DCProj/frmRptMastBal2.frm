VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmRptMastBal2 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Master Balance Report"
   ClientHeight    =   8640
   ClientLeft      =   36
   ClientTop       =   540
   ClientWidth     =   12192
   Icon            =   "frmRptMastBal2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   12192
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpList fplstCodes 
      Height          =   1488
      Left            =   5874
      TabIndex        =   21
      Top             =   1968
      Width           =   3732
      _Version        =   196608
      _ExtentX        =   6583
      _ExtentY        =   2625
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
      Columns         =   3
      Sorted          =   0
      LineWidth       =   1
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   -1
      ColumnWidthScale=   2
      RowHeight       =   -1
      MultiSelect     =   1
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
      ColDesigner     =   "frmRptMastBal2.frx":08CA
   End
   Begin LpLib.fpCombo fpcboDetRev 
      Height          =   348
      Left            =   5874
      TabIndex        =   6
      Top             =   5364
      Width           =   828
      _Version        =   196608
      _ExtentX        =   1460
      _ExtentY        =   614
      Text            =   ""
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
      ColDesigner     =   "frmRptMastBal2.frx":0C82
   End
   Begin LpLib.fpCombo fpcboPrintOrder 
      Height          =   348
      Left            =   5874
      TabIndex        =   5
      Top             =   4968
      Width           =   3612
      _Version        =   196608
      _ExtentX        =   6371
      _ExtentY        =   614
      Text            =   ""
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
      ColDesigner     =   "frmRptMastBal2.frx":1058
   End
   Begin LpLib.fpCombo fpcboRptType 
      Height          =   348
      Left            =   5874
      TabIndex        =   8
      Top             =   6168
      Width           =   1908
      _Version        =   196608
      _ExtentX        =   3365
      _ExtentY        =   614
      Text            =   ""
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
      ColDesigner     =   "frmRptMastBal2.frx":1423
   End
   Begin LpLib.fpCombo fpcboBalType 
      Height          =   348
      Left            =   5874
      TabIndex        =   2
      Top             =   3756
      Width           =   3540
      _Version        =   196608
      _ExtentX        =   6244
      _ExtentY        =   614
      Text            =   ""
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
      ColDesigner     =   "frmRptMastBal2.frx":17F9
   End
   Begin LpLib.fpCombo fpcboCustStatus 
      Height          =   348
      Left            =   5874
      TabIndex        =   4
      Top             =   4560
      Width           =   2004
      _Version        =   196608
      _ExtentX        =   3535
      _ExtentY        =   614
      Text            =   ""
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
      ColDesigner     =   "frmRptMastBal2.frx":1BC4
   End
   Begin LpLib.fpCombo fpcboRevenues 
      Height          =   348
      Left            =   5874
      TabIndex        =   7
      Top             =   5760
      Width           =   3612
      _Version        =   196608
      _ExtentX        =   6371
      _ExtentY        =   614
      Text            =   ""
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
      ColDesigner     =   "frmRptMastBal2.frx":1F8F
   End
   Begin EditLib.fpText fptxtRoute2 
      Height          =   348
      Left            =   552
      TabIndex        =   1
      Top             =   1680
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
   Begin EditLib.fpText fptxtRoute1 
      Height          =   348
      Left            =   168
      TabIndex        =   0
      Top             =   1176
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
      Left            =   5874
      TabIndex        =   3
      Top             =   4152
      Width           =   1764
      _Version        =   196608
      _ExtentX        =   3111
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
      TabIndex        =   10
      Top             =   7248
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
      TabIndex        =   9
      Top             =   7248
      Width           =   1332
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   11
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
            TextSave        =   "5:12 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7133
            TextSave        =   "2/14/2005"
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
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   2304
      X2              =   9876
      Y1              =   3576
      Y2              =   3588
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "* Press SpaceBar or Mouse to Toggle"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   348
      Left            =   2298
      TabIndex        =   22
      Top             =   2544
      Width           =   3276
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Select Revenue:"
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
      Height          =   348
      Left            =   3666
      TabIndex        =   20
      Top             =   5784
      Width           =   2004
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Balance Type:"
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
      Height          =   396
      Left            =   3966
      TabIndex        =   19
      Top             =   3816
      Width           =   1740
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
      Height          =   372
      Left            =   3390
      TabIndex        =   18
      Top             =   6168
      Width           =   2388
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Detail Revenues: "
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
      Left            =   3774
      TabIndex        =   17
      Top             =   5424
      Width           =   2004
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   5052
      Left            =   2274
      Top             =   1680
      Width           =   7644
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Select Group Codes From List:"
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
      Index           =   8
      Left            =   2274
      TabIndex        =   16
      Top             =   2136
      Width           =   3420
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
      Height          =   324
      Index           =   7
      Left            =   3990
      TabIndex        =   15
      Top             =   5028
      Width           =   1716
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Minimum Balance:"
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
      Index           =   2
      Left            =   3630
      TabIndex        =   14
      Top             =   4224
      Width           =   2076
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Status:"
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
      Left            =   3630
      TabIndex        =   13
      Top             =   4644
      Width           =   2076
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
      Caption         =   "Print Master Balance Report"
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
      TabIndex        =   12
      Top             =   552
      Width           =   5004
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
Attribute VB_Name = "frmRptMastBal2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim BegRoute As String, EndRoute As String
Private Sub cmdExit_Click()
  frmUBReportsMenu.Show
  Unload frmRptMastBal
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        UBLog "Closed via RptMastBal by " + PWUser$
        CitiTerminate
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
      fpMinBal.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fptxtRoute2.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub


Private Sub fpcboCustStatus_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboCustStatus.ListDown = True
  End If
  If fpcboCustStatus.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpcboPrintOrder.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpMinBal.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub


Private Sub fpcboDetRev_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboDetRev.ListDown = True
  End If
  If fpcboDetRev.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpcboRevenues.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpcboPrintOrder.SetFocus
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
      fpcboDetRev.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpcboCustStatus.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub

Private Sub fpcboRevenues_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboRevenues.ListDown = True
  End If
  If fpcboRevenues.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpcboRptType.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpcboDetRev.SetFocus
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
        fpcboRevenues.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub

Private Function ValidRoutes()
  If fptxtRoute1 <> "" And fptxtRoute2 <> "" Then
    If fptxtRoute1 > fptxtRoute2 Then
      MsgBox "Invalid Route Selection, The Beginning Route Should Be Less or Equal to Ending Route.", vbOKOnly, "Invalid Selection"
      ValidRoutes = False
    Else
      ValidRoutes = True
      BegRoute = QPTrim(fptxtRoute1)
      EndRoute = QPTrim(fptxtRoute2)
    End If
  Else
    MsgBox "Route Selections May Not Be Left Blank.", vbOKOnly, "Invalid Selection"
  End If
End Function



Private Sub fpMinBal_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpcboCustStatus.SetFocus
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
    fpcboBalType.SetFocus
  End If
End Sub

Private Sub fptxtRoute2_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub

Private Sub cmdPrint_Click()
  If ValidRoutes Then
    DeActivateControls Me, True
    If fpcboRptType.ListIndex = 0 Then
      MasterBalanceListing2
    ElseIf fpcboRptType.ListIndex = 1 Then
      MasterBalanceListing
      ActivateControls Me, True
    Else
      ActivateControls Me, True
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
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
  CodesList fplstCodes
  fpcboBalType.AddItem "Total Balance"
  fpcboBalType.AddItem "Current Balance"
  fpcboBalType.AddItem "Past Due Balance"
  fpcboBalType.AddItem "Credit Balance"
  fpcboBalType.ListIndex = 0
  fpcboCustStatus.AddItem "ALL"
  fpcboCustStatus.AddItem "Active"
  fpcboCustStatus.AddItem "Inactive"
  fpcboCustStatus.AddItem "Balance"
  fpcboCustStatus.AddItem "Pending"
  fpcboCustStatus.AddItem "Delinquent"
  fpcboCustStatus.AddItem "Final"
  fpcboCustStatus.ListIndex = 0
  fpcboPrintOrder.AddItem "Customer Name Order"
  fpcboPrintOrder.AddItem "Account Number Order"
  fpcboPrintOrder.AddItem "Location Number Order"
  fpcboPrintOrder.ListIndex = 0
  fpcboDetRev.AddItem "Y"
  fpcboDetRev.AddItem "N"
  fpcboDetRev.ListIndex = 0
  FillRevList fpcboRevenues
  fpcboRevenues.ListIndex = 0
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
Private Sub MasterBalanceListing()
  Dim RCnt As Integer, UBCustRecLen As Integer, UBSetupreclen As Integer
  Dim MaxRevenue As Integer, TRevName As String, AndPos As String
  Dim UsingBook As Boolean, RStatus As String, UsingName As Boolean
  Dim PageNo As Integer, UseStatus As Boolean, AcctNo As Long
  Dim Dash80 As String, IndexName As String, RealBalance As Double
  Dim IdxRecLen As Integer, IdxFileSize As Long, OKToSkip As Boolean
  Dim IdxNumOfRecs As Long, NumOfRecs As Long, Handle As Integer
  Dim cnt As Long, UBCust As Integer, UBRpt As Integer, UBSetUp1 As Integer
  Dim RevChk As Integer, CStatus As String, Book As String
  Dim SEQNUMB As String, BalType As String, ChkBalance As Double
  Dim MinBal As Double, RevSource As Integer, TCurrBalance As Double
  Dim CustCnt As Long, TPrevBalance As Double, Detail As String
  Dim GTotal As Double, CoFlag As Boolean, Stat As String, UsingAcct As Boolean
  Dim POrder As String, Bal As String, DLineCnt As Integer, bk As Integer
  Dim TCnt As Integer, First As Integer, Last As Integer, Rev As String
  Dim TabStop As Integer, Det As Boolean, Order As String
  Dim ReportFile As String
  RCnt = RCnt + 1
  UsingAcct = False
  UseStatus = False
  UsingName = False
  UsingBook = False
  ReDim fmt$(1 To 3)
  fmt$(1) = "####,#.##"
  fmt$(2) = "#####"
  fmt$(3) = "######,#.##"
  'Main Body Start
  FrmShowPctComp.Label1 = "Creating Master Balance Listing"
  FrmShowPctComp.Show , Me

  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))
  ReDim UBSetUp(1) As UBSetupRecType
  LoadUBSetUpFile UBSetUp(), UBSetupreclen
  TOWNNAME$ = QPTrim$(UBSetUp(1).UTILNAME)
  MaxRevenue = 15
  GoSub CheckBalInfo 'This gets all stuff from screen for report values
  ReDim RevenueName(1 To 15) As String * 10
  For RCnt = 1 To 15
    TRevName$ = QPTrim$(UBSetUp(1).Revenues(RCnt).RevName)
    If Len(TRevName$) > 0 Then
          AndPos = InStr(TRevName$, "&")
      If AndPos Then
        Mid$(TRevName$, AndPos) = " "
      End If
      RevenueName(RCnt) = TRevName$
    Else
      MaxRevenue = RCnt - 1
      Exit For
    End If
  Next
  
  ReDim RevTotals(1 To MaxRevenue) As Double

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

  Else
    NumOfRecs = FileSize(UBPath$ + "UBCUST.DAT") \ UBCustRecLen
  End If

  MaxLines = 55
  Dash80$ = String$(80, "-")

  UBCust = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen
  ReportFile$ = UBPath$ + "UBBALIST.RPT"
  UBRpt = FreeFile
  Open ReportFile$ For Output As UBRpt

  'RecFile = FREEFILE
  'OPEN "RECLIST.IDX" FOR RANDOM AS RecFile LEN = 4

  'ShowProcessingScrn "Master Balance Listing."

  GoSub DoCustRptHeader

  For cnt = 1 To NumOfRecs
    FrmShowPctComp.ShowPctComp cnt, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      GoTo ExitBalRpt
    End If

    If UsingName Or UsingBook Then
      AcctNo& = IdxBuff(cnt).RecNum
    Else
      AcctNo& = cnt
    End If

    Get UBCust, AcctNo&, UBCustRec(1)

    If UBCustRec(1).DelFlag <> 0 Then
      GoTo BSkipEm
    End If

    RealBalance# = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)

    '110598 Old code could skew totals, A credit in one revenue and a debit
    '       in another revenue, that totaled to zero wouldn't show!!!
    'IF RealBalance# = 0 THEN
    '  GOTO BSkipEM
    'END IF

    '110598 Code to fix the above bug
        'If all there rev sources are "0" then skipem
    OKToSkip = True
    For RevChk = 1 To 15        'MaxRevenue '15
      If Round#(UBCustRec(1).CurrRevAmts(RevChk)) <> 0 Then
        OKToSkip = False
        Exit For
      End If
    Next

    If OKToSkip Then
      GoTo BSkipEm
    End If

    If LineCnt > MaxLines Then
      Print #UBRpt, FF$
      GoSub DoCustRptHeader
    End If

    If UseStatus Then           'if they care about the cust status, or want all.
      CStatus$ = Left$(QPTrim$(UBCustRec(1).Status), 1)
      If CStatus$ <> RStatus$ Then
         GoTo BSkipEm
      End If
    End If

    Book$ = QPTrim$(UBCustRec(1).Book)
    SEQNUMB$ = QPTrim$(UBCustRec(1).SEQNUMB)
    If Len(Book$) = 0 Then
      Book$ = "  "
    End If
    bk = Val(Book$)
    If bk < Val(BegRoute) Or bk > Val(EndRoute) Then
      GoTo BSkipEm
    End If

    'IF (RealBalance# = 0) AND (NOT Ok2Skip) THEN STOP

    Select Case BalType$
    Case "Pa"
      ChkBalance# = UBCustRec(1).PrevBalance
      If (ChkBalance# >= MinBal#) And (ChkBalance# > 0) Then
        If RealBalance# <= 0 Then
            GoTo BSkipEm
        End If
      Else
        GoTo BSkipEm
      End If
    Case "Cu"
      ChkBalance# = UBCustRec(1).CurrBalance
      If MinBal# > 0 Then
        If (ChkBalance# < MinBal#) Or (ChkBalance# <= 0) Then
          GoTo BSkipEm
        End If
      Else
        If ChkBalance# = 0 Then
          GoTo BSkipEm
        End If
      End If
    Case "To"
      ChkBalance# = RealBalance#
      If MinBal# > 0 Then
        If ChkBalance# < MinBal# Then
          GoTo BSkipEm
        End If
      End If
    Case "Cr"
      If RealBalance# >= 0 Then
        GoTo BSkipEm
      End If
    End Select

    If RevSource > 0 Then
      If UBCustRec(1).CurrRevAmts(RevSource) <> 0 Then
        Print #UBRpt, Using(fmt$(2), AcctNo&);
        Print #UBRpt, Tab(8); Book$; "-"; SEQNUMB$;
        Print #UBRpt, Tab(18); Left$(UBCustRec(1).CustName, 31);
        Print #UBRpt, Tab(60); Using(fmt$(1), Round#(UBCustRec(1).CurrRevAmts(RevSource)))
        LineCnt = LineCnt + 1
        TCurrBalance# = Round#(TCurrBalance# + UBCustRec(1).CurrRevAmts(RevSource))
        CustCnt = CustCnt + 1
      End If
    Else
      Print #UBRpt, Using(fmt$(2), AcctNo&);
      Print #UBRpt, Tab(8); Book$; "-"; SEQNUMB$;
      Print #UBRpt, Tab(18); Left$(UBCustRec(1).CustName, 31);
      Print #UBRpt, Tab(50); Using(fmt$(1), Round#(UBCustRec(1).CurrBalance));
      Print #UBRpt, Tab(61); Using(fmt$(1), Round#(UBCustRec(1).PrevBalance));
      Print #UBRpt, Tab(72); Using(fmt$(1), RealBalance#)
      LineCnt = LineCnt + 1
      TCurrBalance# = Round#(TCurrBalance# + UBCustRec(1).CurrBalance)
      TPrevBalance# = Round#(TPrevBalance# + UBCustRec(1).PrevBalance)
      CustCnt = CustCnt + 1
    End If

    GoSub PrintDetail
    
'    If AskAbandonPrint% Then
'      AbortFlag = True
'      Exit For
'    End If

BSkipEm:
    'ShowPctComp cnt, NumOfRecs
  Next

  GoSub DoCustRptFooter
  Close
  Erase IdxBuff, UBCustRec
  If CustCnt > 0 Then
 ' If Not AbortFlag Then
    ViewPrint ReportFile$, "Balance Listing Report."
 ' End If
  Else
    MsgBox "No Information to print.", vbOKOnly, "No Information"
    ActivateControls Me, True
  End If

  'Main Body Exit
ExitBalRpt:

  Exit Sub

DoCustRptHeader:
  PageNo = PageNo + 1
  Print #UBRpt, TOWNNAME$
  Print #UBRpt, Tab(26); "Customer Balance Listing Report"; Tab(70); "Page #"; PageNo
  Print #UBRpt, "Report Date: "; Date$
  Print #UBRpt, "Acct #"; Tab(9); "Location"; Tab(18); "Customer Name";
  If RevSource > 0 Then
    Print #UBRpt, Tab(60); fpcboRevenues.Text; " Amount"
  Else
     Print #UBRpt, Tab(52); "Cur Bal"; Tab(62); "Past Due"; Tab(73); "Acct Bal"
  End If
  Print #UBRpt, Dash80$
  LineCnt = 5
  Return

DoCustRptFooter:
  Print #UBRpt, Dash80$
  Print #UBRpt, "Totals:"; Tab(10); "Customers: "; Using("#####,#", CustCnt);
  If RevSource = 0 Then
    Print #UBRpt, Tab(48); Using(fmt$(3), TCurrBalance#);
    Print #UBRpt, Tab(59); Using(fmt$(3), TPrevBalance#);
    Print #UBRpt, Tab(70); Using(fmt$(3), Round#(TCurrBalance# + TPrevBalance#))
    'IF Detail THEN
    For cnt = 1 To MaxRevenue
      Detail$ = Space$(28)
      LSet Detail$ = RevenueName(cnt)
     ' Mid$(Detail$, 15) = fmt$(3)
      GTotal# = Round#(GTotal# + RevTotals(cnt))
      Print #UBRpt, QPTrim(Detail$); Tab(30); Using(fmt$(3), RevTotals(cnt))
    Next
    LSet Detail$ = "Grand Total:"
   ' Mid$(Detail$, 15) = fmt$(3)
    Print #UBRpt, QPTrim(Detail$); Tab(30); Using(fmt$(3), GTotal#)
    'ELSE
    '  PRINT #UBRpt,
    'END IF
  Else
    Print #UBRpt, Tab(58); Using(fmt$(3), TCurrBalance#)
  End If

  Print #UBRpt, "******************"
  Print #UBRpt, "Report Parameters"
  Print #UBRpt, "                Routes:"; Str$(BegRoute); " to"; Str$(EndRoute)
  If CoFlag Then
    Print #UBRpt, "       Minimum Balance: N/A       ";
  Else
    Print #UBRpt, "       Minimum Balance: "; Using("$######.##", MinBal#);
  End If
  Print #UBRpt, "            Customer Status:"; Stat$
  Print #UBRpt, "        Printing Order:"; POrder$;
  Print #UBRpt, "           Balance Type:"; Bal$;
    If RevSource > 0 Then
    Print #UBRpt, Tab(39); "Revenue Source: "; fpcboRevenues.Text;
  End If
  Print #UBRpt,
  Print #UBRpt, Dash80$
  LineCnt = LineCnt + 4

  Print #UBRpt, FF$
  Return

PrintDetail:
  DLineCnt = LineCnt
  TCnt = 0
  Detail$ = Space$(18)
  First = 1
  Last = MaxRevenue
  If MaxRevenue < Last Then
    Last = MaxRevenue
  End If
  For RCnt = First To Last
    TCnt = TCnt + 1
        TabStop = (TCnt * 21) - 20
    If TabStop > 81 Then
      LineCnt = LineCnt + 1
      TCnt = 1
      TabStop = (TCnt * 21) - 20
    End If
    LSet Detail$ = RevenueName(RCnt)
   ' Mid$(Detail$, 10) = "#####.##"
    RevTotals(RCnt) = Round#(RevTotals(RCnt) + UBCustRec(1).CurrRevAmts(RCnt))
    'IF RCnt = 15 THEN
    '  IF UBCustRec(1).CurrRevAmts(RCnt) <> 0 THEN STOP
    'END IF
    If Det Then
      Print #UBRpt, Tab(TabStop); QPTrim(Detail$); Using("#####.##", UBCustRec(1).CurrRevAmts(RCnt));
    End If
  Next

  If Det Then
    Print #UBRpt,
    Print #UBRpt, Dash80$
    LineCnt = LineCnt + 2
   Else
    LineCnt = DLineCnt
  End If
  Return


CheckBalInfo:
  BegRoute = fptxtRoute1
  EndRoute = fptxtRoute2
  BalType$ = Mid$(fpcboBalType.Text, 1, 2)
  MinBal# = fpMinBal.DoubleValue
  RStatus$ = Mid$(fpcboCustStatus.Text, 1, 2)
  Order$ = Mid$(fpcboPrintOrder.Text, 1, 1)
  If fpcboDetRev.ListIndex = 0 Then
    Det = True
  Else
    Det = False
  End If
'revenue listindex should be same as revenue number since
'first line (listindex of 0) is all revenues.
RevSource = fpcboRevenues.ListIndex
  If RevSource > 0 Then
    Det = False
  End If

  If BegRoute > EndRoute Then
    MsgBox "Invalid Route Order", vbOKOnly, "Invalid Parameter"
    fptxtRoute1.SetFocus
    GoTo ParmErrorRet
  End If
  Select Case BalType$
  Case "Pa"
    Bal$ = " PAST DUE"
  Case "Cu"
    Bal$ = " CURRENT"
  Case "To"
    Bal$ = " TOTAL BALANCE"
  Case "Cr"
    Bal$ = " CREDIT BALANCE"
    CoFlag = True
  Case Else
    MsgBox "Invalid Balance Type", vbOKOnly, "Invalid Parameter"
    fpcboBalType.SetFocus
    GoTo ParmErrorRet
  End Select

  Select Case Order$
  Case "C"
    IndexName$ = NameIndexFile
    UsingName = True
    POrder$ = " CUSTOMER NAME"
  Case "A"
    POrder$ = " ACCOUNT NUMBER"
        IndexName$ = ""
    UsingAcct = True
  Case "L"
    POrder$ = " LOCATION NUMBER"
    IndexName$ = BookIndexFile
    UsingBook = True
  Case Else
    MsgBox "Invalid Printing Order", vbOKOnly, "Invalid Parameter"
    fpcboPrintOrder.SetFocus
    GoTo ParmErrorRet
  End Select
  Select Case RStatus$
  Case "Ac"
    UseStatus = True
    Stat$ = " ACTIVE"
  Case "In"
      UseStatus = True
    Stat$ = " INACTIVE"
  Case "Ba"
    UseStatus = True
    Stat$ = " BALANCE DUE"
  Case "Pe"
    Stat$ = " PENDING"
    UseStatus = True
  Case "De"
    Stat$ = " DELINQUENT"
    UseStatus = True
  Case "Fi"
    Stat$ = " FINAL"
    UseStatus = True
  Case Else
    Stat$ = " ALL"
    UseStatus = False
  End Select
  RStatus$ = Mid$(fpcboCustStatus.Text, 1, 1)
  Return

ParmErrorRet:

  Exit Sub

End Sub

Private Sub MasterBalanceListing2()
  Dim RCnt As Integer, UBCustRecLen As Integer, UBSetupreclen As Integer
  Dim MaxRevenue As Integer, TRevName As String, AndPos As String
  Dim UsingBook As Boolean, RStatus As String, UsingName As Boolean
  Dim PageNo As Integer, UseStatus As Boolean, AcctNo As Long
  Dim Dash80 As String, IndexName As String, RealBalance As Double
  Dim IdxRecLen As Integer, IdxFileSize As Long, OKToSkip As Boolean
  Dim IdxNumOfRecs As Long, NumOfRecs As Long, Handle As Integer
  Dim cnt As Long, UBCust As Integer, UBRpt As Integer, UBSetUp1 As Integer
  Dim RevChk As Integer, CStatus As String, Book As String
  Dim SEQNUMB As String, BalType As String, ChkBalance As Double
  Dim MinBal As Double, RevSource As Integer, TCurrBalance As Double
  Dim CustCnt As Long, TPrevBalance As Double, Detail As String
  Dim GTotal As Double, CoFlag As Boolean, Stat As String, UsingAcct As Boolean
  Dim POrder As String, Bal As String, DLineCnt As Integer, bk As Integer
  Dim TCnt As Integer, First As Integer, Last As Integer, Rev As String
  Dim TabStop As Integer, Det As Boolean, Order As String
  Dim ToPrint As String, ToPrintD As String, ToPrintH1 As String
  Dim ToPrintH2 As String, UBRpt2 As Integer, ToPrintS As String
  Dim ToPrintD2 As String, DetFlag As Integer, ReportFile As String
  Dim Report2 As String
  RCnt = RCnt + 1
  UsingAcct = False
  UseStatus = False
  UsingName = False
  UsingBook = False
  ReDim fmt$(1 To 3)
  fmt$(1) = "####,#.##"
  fmt$(2) = "#####"
  fmt$(3) = "######,#.##"
  'Main Body Start
  FrmShowPctComp.Label1 = "Creating Master Balance Listing"
  FrmShowPctComp.Show , Me

  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))
  ReDim UBSetUp(1) As UBSetupRecType
  LoadUBSetUpFile UBSetUp(), UBSetupreclen
  TOWNNAME$ = QPTrim$(UBSetUp(1).UTILNAME)
  MaxRevenue = 15
  GoSub CheckBalInfo 'This gets all stuff from screen for report values
  ReDim RevenueName(1 To 15) As String * 10
  For RCnt = 1 To 15
    TRevName$ = QPTrim$(UBSetUp(1).Revenues(RCnt).RevName)
    If Len(TRevName$) > 0 Then
          AndPos = InStr(TRevName$, "&")
      If AndPos Then
        Mid$(TRevName$, AndPos) = " "
      End If
      RevenueName(RCnt) = TRevName$
    Else
      MaxRevenue = RCnt - 1
      Exit For
    End If
  Next
  If Det Then
    DetFlag = MaxRevenue
  Else
    DetFlag = 0
  End If

  ReDim RevTotals(1 To MaxRevenue) As Double

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

  Else
      NumOfRecs = FileSize(UBPath$ + "UBCUST.DAT") \ UBCustRecLen
  End If

  MaxLines = 55
  Dash80$ = String$(80, "-")

  UBCust = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen
  ReportFile$ = UBPath$ + "UBBALIST.RPT"
  UBRpt = FreeFile
  Open ReportFile$ For Output As UBRpt
  Report2$ = UBPath$ + "UBBALSub.RPT"
  UBRpt2 = FreeFile
  Open Report2$ For Output As UBRpt2

  'RecFile = FREEFILE
  'OPEN "RECLIST.IDX" FOR RANDOM AS RecFile LEN = 4

  'ShowProcessingScrn "Master Balance Listing."

  GoSub DoCustRptHeader

  For cnt = 1 To NumOfRecs
    FrmShowPctComp.ShowPctComp cnt, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ActivateControls Me, True
      GoTo ExitBalRpt
    End If

      If UsingName Or UsingBook Then
      AcctNo& = IdxBuff(cnt).RecNum
    Else
      AcctNo& = cnt
    End If

    Get UBCust, AcctNo&, UBCustRec(1)

    If UBCustRec(1).DelFlag <> 0 Then
      GoTo BSkipEm
    End If

    RealBalance# = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)

    '110598 Old code could skew totals, A credit in one revenue and a debit
    '       in another revenue, that totaled to zero wouldn't show!!!
    'IF RealBalance# = 0 THEN
    '  GOTO BSkipEM
    'END IF

    '110598 Code to fix the above bug
        'If all there rev sources are "0" then skipem
    OKToSkip = True
    For RevChk = 1 To 15        'MaxRevenue '15
      If Round#(UBCustRec(1).CurrRevAmts(RevChk)) <> 0 Then
        OKToSkip = False
        Exit For
      End If
    Next

    If OKToSkip Then
      GoTo BSkipEm
    End If

'    If Linecnt > MaxLines Then
'      Print #UBRpt, FF$
'
'    End If

    If UseStatus Then           'if they care about the cust status, or want all.
      CStatus$ = Left$(QPTrim$(UBCustRec(1).Status), 1)
      If CStatus$ <> RStatus$ Then
         GoTo BSkipEm
      End If
    End If

    Book$ = QPTrim$(UBCustRec(1).Book)
    SEQNUMB$ = QPTrim$(UBCustRec(1).SEQNUMB)
    If Len(Book$) = 0 Then
      Book$ = "  "
    End If
    bk = Val(Book$)
    If bk < Val(BegRoute) Or bk > Val(EndRoute) Then
      GoTo BSkipEm
    End If

    'IF (RealBalance# = 0) AND (NOT Ok2Skip) THEN STOP

    Select Case BalType$
    Case "Pa"
      ChkBalance# = UBCustRec(1).PrevBalance
      If (ChkBalance# >= MinBal#) And (ChkBalance# > 0) Then
        If RealBalance# <= 0 Then
            GoTo BSkipEm
        End If
      Else
        GoTo BSkipEm
      End If
    Case "Cu"
      ChkBalance# = UBCustRec(1).CurrBalance
      If MinBal# > 0 Then
        If (ChkBalance# < MinBal#) Or (ChkBalance# <= 0) Then
          GoTo BSkipEm
        End If
      Else
        If ChkBalance# = 0 Then
          GoTo BSkipEm
        End If
      End If
    Case "To"
      ChkBalance# = RealBalance#
      If MinBal# > 0 Then
        If ChkBalance# < MinBal# Then
          GoTo BSkipEm
        End If
      End If
    Case "Cr"
      If RealBalance# >= 0 Then
        GoTo BSkipEm
      End If
    End Select

    If RevSource > 0 Then
      If UBCustRec(1).CurrRevAmts(RevSource) <> 0 Then
        ToPrint$ = Str$(AcctNo&) + "~"
        ToPrint$ = ToPrint$ + Book$ + "-" + SEQNUMB$
        ToPrint$ = ToPrint$ + "~" + Left$(UBCustRec(1).CustName, 31)
        ToPrint$ = ToPrint$ + "~" + Str$(Round#(UBCustRec(1).CurrRevAmts(RevSource)))
        'Linecnt = Linecnt + 1
        TCurrBalance# = Round#(TCurrBalance# + UBCustRec(1).CurrRevAmts(RevSource))
        CustCnt = CustCnt + 1
        GoSub PrintDetail
      End If
    Else
      ToPrint$ = Str$(AcctNo&) + "~"
      ToPrint$ = ToPrint$ + Book$ + "-" + SEQNUMB$
      ToPrint$ = ToPrint$ + "~" + Left$(UBCustRec(1).CustName, 31)
      ToPrint$ = ToPrint$ + "~" + Str$(Round#(UBCustRec(1).CurrBalance))
      ToPrint$ = ToPrint$ + "~" + Str$(Round#(UBCustRec(1).PrevBalance))
      ToPrint$ = ToPrint$ + "~" + Str$(RealBalance#)
      'Linecnt = Linecnt + 1
      TCurrBalance# = Round#(TCurrBalance# + UBCustRec(1).CurrBalance)
      TPrevBalance# = Round#(TPrevBalance# + UBCustRec(1).PrevBalance)
      CustCnt = CustCnt + 1
      GoSub PrintDetail
    End If

   
    
'    If AskAbandonPrint% Then
'      AbortFlag = True
'      Exit For
'    End If

BSkipEm:
    'ShowPctComp cnt, NumOfRecs
  Next
  GoSub DoCustRptHeader
  GoSub DoCustRptFooter
  Close
  Erase IdxBuff, UBCustRec

   'ViewPrint "UBBALIST.RPT", "Balance Listing Report."
  If CustCnt > 0 Then
  Load frmLoadingRpt
  frmLoadingRpt.setwherefrom frmRptMastBal
  ARptMastBalList.txtDate = Now
  ARptMastBalList.txtTown = TOWNNAME$
  ARptMastBalList.Title = "Master Customer Balance Report"
  ARptMastBalList.txtRptParm1.Caption = ToPrintH1$
  ARptMastBalList.txtRptParm2.Caption = ToPrintH2$
  ARptMastBalList.txtTotCust = CustCnt
  ARptMastBalList.txtTotCur.DataValue = TCurrBalance#
  ARptMastBalList.txtTotPast.DataValue = TPrevBalance#
  ARptMastBalList.txtHead = fpcboRevenues.Text
  ARptMastBalList.txtTotAcctBal.DataValue = Round#(TCurrBalance# + TPrevBalance#)
  ARptMastBalList.GetName ReportFile$, Report2$, DetFlag, RevSource
  ARptMastBalList.startrpt
  Else
    MsgBox "No Information to print.", vbOKOnly, "No Information"
    ActivateControls Me, True
  End If
  'Main Body Exit
ExitBalRpt:

  Exit Sub

DoCustRptHeader:
  'PageNo = PageNo + 1
  'Print #UBRpt, TownName$
  'Print #UBRpt, Tab(26); "Customer Balance Listing Report"; Tab(70); "Page #"; PageNo
  'Print #UBRpt, "Report Date: "; Date$
  'Print #UBRpt, "Acct #"; Tab(9); "Location"; Tab(18); "Customer Name";
'  If RevSource > 0 Then
'    Print #UBRpt, Tab(60); fpcboRevenues.Text; " Amount"
'  Else
'     Print #UBRpt, Tab(52); "Cur Bal"; Tab(62); "Past Due"; Tab(73); "Acct Bal"
'  End If
'  Print #UBRpt, Dash80$
'  Linecnt = 5
  Return

DoCustRptFooter:
  ToPrintH1$ = ""
  ToPrintH2$ = ""
 ' Print #UBRpt, "Totals:"; Tab(10); "Customers: "; Using("#####,#", CustCnt);
  If RevSource = 0 Then
  '  Print #UBRpt, Tab(48); Using(fmt$(3), TCurrBalance#);
  '  Print #UBRpt, Tab(59); Using(fmt$(3), TPrevBalance#);
  '  Print #UBRpt, Tab(70); Using(fmt$(3), Round#(TCurrBalance# + TPrevBalance#))
    'IF Detail THEN
    ToPrintS$ = ""
    For cnt = 1 To MaxRevenue
      Detail$ = Space$(28)
      LSet Detail$ = RevenueName(cnt)
     ' Mid$(Detail$, 15) = fmt$(3)
      GTotal# = Round#(GTotal# + RevTotals(cnt))
      ToPrintS$ = ToPrintS$ + QPTrim(Detail$) + "~" + Str$(RevTotals(cnt)) + "~"
      Print #UBRpt2, ToPrintS$
      ToPrintS$ = ""
    Next
    LSet Detail$ = "Grand Total:"
   ' Mid$(Detail$, 15) = fmt$(3)
    ToPrintS$ = ToPrintS$ + QPTrim(Detail$) + "~" + Str$(GTotal#)
    Print #UBRpt2, ToPrintS$ 'ELSE
    '  PRINT #UBRpt,
    'END IF
  Else
   ' Print #UBRpt, Tab(58); Using(fmt$(3), TCurrBalance#)
  End If

  'Print #UBRpt, "******************"
  'Print #UBRpt, "Report Parameters"
  ToPrintH1$ = "Routes:" + Str$(BegRoute) + " to" + Str$(EndRoute) + "  Printing Order:" + POrder$
'  If CoFlag Then
'    Print #UBRpt, "       Minimum Balance: N/A       ";
'  Else
  ToPrintH1$ = ToPrintH1$ + "  Minimum Balance: " + Using("$######.##", MinBal#)
'  End If
  ToPrintH2$ = "Customer Status:" + Stat$
  ToPrintH2$ = ToPrintH2$ + "  Balance Type:" + Bal$ + "  Revenue Source: " + fpcboRevenues.Text
  '  If RevSource > 0 Then
    'Print #UBRpt, Tab(39);
 ' End If
 ' Print #UBRpt,
 ' Print #UBRpt, Dash80$
 ' Linecnt = Linecnt + 4

  'Print #UBRpt, FF$
  Return

PrintDetail:
  DLineCnt = LineCnt
  TCnt = 0
  Detail$ = Space$(18)
  First = 1
  ToPrintD$ = ""
  ToPrintD2$ = ""
  Last = 15
'  If MaxRevenue < Last Then
'    Last = MaxRevenue
'  End If
  For RCnt = First To Last
    TCnt = TCnt + 1
        TabStop = (TCnt * 21) - 20
    If TabStop > 81 Then
     ' Linecnt = Linecnt + 1
      TCnt = 1
      TabStop = (TCnt * 21) - 20
    End If
    If RCnt <= MaxRevenue Then
      LSet Detail$ = RevenueName(RCnt)
      ' Mid$(Detail$, 10) = "#####.##"
      RevTotals(RCnt) = Round#(RevTotals(RCnt) + UBCustRec(1).CurrRevAmts(RCnt))
      ToPrintD2$ = Str$(UBCustRec(1).CurrRevAmts(RCnt)) + "~"
      'IF RCnt = 15 THEN
      '  IF UBCustRec(1).CurrRevAmts(RCnt) <> 0 THEN STOP
    'End If
    Else
      LSet Detail$ = ""
      ToPrintD2$ = "~"
    End If
    If Det Then
      ToPrintD$ = ToPrintD$ + QPTrim(Detail$) + "~" + ToPrintD2$
    Else
      ToPrintD$ = ToPrintD$ + "~~~"
    End If
  
  Next

  If Det Then
    'Print #UBRpt,
    Print #UBRpt, ToPrint$ + "~" + ToPrintD$
   Else
    'Linecnt = DLineCnt
    Print #UBRpt, ToPrint$ + "~" + ToPrintD$
  End If
  ToPrint$ = ""
  ToPrintD$ = ""
  Return


CheckBalInfo:
  BegRoute = fptxtRoute1
  EndRoute = fptxtRoute2
  BalType$ = Mid$(fpcboBalType.Text, 1, 2)
  MinBal# = fpMinBal.DoubleValue
  RStatus$ = Mid$(fpcboCustStatus.Text, 1, 2)
  Order$ = Mid$(fpcboPrintOrder.Text, 1, 1)
  If fpcboDetRev.ListIndex = 0 Then
    Det = True
  Else
    Det = False
  End If
'revenue listindex should be same as revenue number since
'first line (listindex of 0) is all revenues.
RevSource = fpcboRevenues.ListIndex
  If RevSource > 0 Then
    Det = False
  End If

  If BegRoute > EndRoute Then
    MsgBox "Invalid Route Order", vbOKOnly, "Invalid Parameter"
    fptxtRoute1.SetFocus
    GoTo ParmErrorRet
  End If
  Select Case BalType$
  Case "Pa"
    Bal$ = " PAST DUE"
  Case "Cu"
    Bal$ = " CURRENT"
  Case "To"
    Bal$ = " TOTAL BALANCE"
  Case "Cr"
    Bal$ = " CREDIT BALANCE"
    CoFlag = True
  Case Else
    MsgBox "Invalid Balance Type", vbOKOnly, "Invalid Parameter"
    fpcboBalType.SetFocus
    GoTo ParmErrorRet
  End Select

  Select Case Order$
  Case "C"
    IndexName$ = NameIndexFile
    UsingName = True
    POrder$ = " CUSTOMER NAME"
  Case "A"
    POrder$ = " ACCOUNT NUMBER"
        IndexName$ = ""
    UsingAcct = True
  Case "L"
    POrder$ = " LOCATION NUMBER"
    IndexName$ = BookIndexFile
    UsingBook = True
  Case Else
    MsgBox "Invalid Printing Order", vbOKOnly, "Invalid Parameter"
    fpcboPrintOrder.SetFocus
    GoTo ParmErrorRet
  End Select
  Select Case RStatus$
  Case "Ac"
    UseStatus = True
    Stat$ = " ACTIVE"
  Case "In"
      UseStatus = True
    Stat$ = " INACTIVE"
  Case "Ba"
    UseStatus = True
    Stat$ = " BALANCE DUE"
  Case "Pe"
    Stat$ = " PENDING"
    UseStatus = True
  Case "De"
    Stat$ = " DELINQUENT"
    UseStatus = True
  Case "Fi"
    Stat$ = " FINAL"
    UseStatus = True
  Case Else
    Stat$ = " ALL"
    UseStatus = False
  End Select
  RStatus$ = Mid$(fpcboCustStatus.Text, 1, 1)
  Return

ParmErrorRet:

  Exit Sub

End Sub
Private Function CodesList(x As fpList)
  Dim GroupCde As GroupCodeRecType
  Dim GrpCodeRecLen As Integer, ghandle As Integer, cnt As Integer
  Dim NumofGrps As Integer
  GrpCodeRecLen = Len(GroupCde)
  
  ghandle = FreeFile
  Open UBPath$ + "UBGrpCde.DAT" For Random Shared As ghandle Len = GrpCodeRecLen
  NumofGrps = LOF(ghandle) \ GrpCodeRecLen
  x.Row = 0
  For cnt = 1 To NumofGrps
    Get #ghandle, cnt, GroupCde
    If GroupCde.Deleted = 0 Then
      x.AddItem Str$(cnt) & Chr$(9) & GroupCde.GroupCODE & Chr$(9) & GroupCde.GroupCodeName
'    Else
'      x.AddItem Str$(cnt) & Chr$(9) & GroupCde.GroupCODE & Chr$(9) & "Inactivated Code"
    End If
  Next
  Close
End Function

'Private Sub GetCodestoReport()
'  Dim FundsToClose As GLFundCloseRecType
'  Dim FundCloseListFile As Integer
'  Dim CloseListFileName As String
'  Dim PCnt As Integer, cnt As Integer
'  '--process the FUND LIST into a file of only the selected choices
'  CloseListFileName$ = "FCLOSE.LST"
'  KillFile CloseListFileName$
'  FundCloseListFile = FreeFile
'  Open CloseListFileName$ For Random As FundCloseListFile Len = Len(FundsToClose)
'  fplstFunds.ListIndex = 0
'  'If fplstFunds.Selected = True Then
'    For PCnt = 0 To fplstFunds.ListCount - 1
'      If fplstFunds.Selected(PCnt) Then
'        cnt = cnt + 1
'        fplstFunds.col = 0
'        fplstFunds.ListIndex = PCnt
'        FundsToClose.FundNum = QPTrim(fplstFunds.ColText)
'        Put FundCloseListFile, cnt, FundsToClose
'        'fplstFunds.Row = fplstFunds.NextSel
'      End If
'    Next
'  Close
'End Sub

