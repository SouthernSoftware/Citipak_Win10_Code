VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPrnSalesTaxRpt1 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sales Tax Report"
   ClientHeight    =   8892
   ClientLeft      =   36
   ClientTop       =   492
   ClientWidth     =   12192
   Icon            =   "frmPrnSalesTaxRpt1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8892
   ScaleWidth      =   12192
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboCoAcct 
      Height          =   384
      Left            =   5136
      TabIndex        =   2
      Top             =   3600
      Width           =   4740
      _Version        =   196608
      _ExtentX        =   8361
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
      ColDesigner     =   "frmPrnSalesTaxRpt1.frx":08CA
   End
   Begin LpLib.fpCombo fpcboStAcct 
      Height          =   384
      Left            =   5136
      TabIndex        =   1
      Top             =   2976
      Width           =   4740
      _Version        =   196608
      _ExtentX        =   8361
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
      ColDesigner     =   "frmPrnSalesTaxRpt1.frx":0CCD
   End
   Begin LpLib.fpCombo fpcboCombined 
      Height          =   384
      Left            =   5136
      TabIndex        =   0
      Top             =   2352
      Width           =   996
      _Version        =   196608
      _ExtentX        =   1757
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
      Columns         =   0
      Sorted          =   0
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   -1
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
      ColDesigner     =   "frmPrnSalesTaxRpt1.frx":10D0
   End
   Begin LpLib.fpCombo fpcboStateCode 
      Height          =   384
      Left            =   5124
      TabIndex        =   3
      Top             =   4200
      Width           =   996
      _Version        =   196608
      _ExtentX        =   1757
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
      ColDesigner     =   "frmPrnSalesTaxRpt1.frx":13FF
   End
   Begin LpLib.fpCombo fpcboRptType 
      Height          =   384
      Left            =   5112
      TabIndex        =   7
      Top             =   6648
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
      ColDesigner     =   "frmPrnSalesTaxRpt1.frx":172E
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
      TabIndex        =   8
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
      TabIndex        =   9
      Top             =   7512
      Width           =   1332
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   10
      Top             =   8532
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
            TextSave        =   "3:39 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7133
            TextSave        =   "7/23/2004"
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
      TabIndex        =   5
      Top             =   5424
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
      TabIndex        =   6
      Top             =   6036
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
   Begin EditLib.fpDoubleSingle fpdblFactor 
      Height          =   372
      Left            =   5136
      TabIndex        =   4
      Top             =   4824
      Width           =   852
      _Version        =   196608
      _ExtentX        =   1503
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
      Text            =   ""
      DecimalPlaces   =   3
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   1
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
      TabIndex        =   20
      Top             =   6696
      Width           =   2388
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "State Code:"
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
      Left            =   3504
      TabIndex        =   19
      Top             =   4248
      Width           =   1476
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "(Only if use Combined Acct)"
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
      Height          =   420
      Left            =   6192
      TabIndex        =   18
      Top             =   4872
      Width           =   2748
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
      Left            =   2976
      TabIndex        =   17
      Top             =   3636
      Width           =   2004
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Combined Account:"
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
      Left            =   2760
      TabIndex        =   16
      Top             =   2400
      Width           =   2220
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   5148
      Left            =   1920
      Top             =   2160
      Width           =   8316
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Print Sales Tax Report"
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
      TabIndex        =   15
      Top             =   1248
      Width           =   4332
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3216
      Top             =   1008
      Width           =   5772
   End
   Begin VB.Image Image1 
      Height          =   276
      Left            =   2280
      Picture         =   "frmPrnSalesTaxRpt1.frx":1A94
      Top             =   2376
      Width           =   288
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
      TabIndex        =   14
      Top             =   6084
      Width           =   1572
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "State Factor:"
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
      Left            =   3384
      TabIndex        =   13
      Top             =   4848
      Width           =   1596
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
      TabIndex        =   12
      Top             =   5464
      Width           =   1668
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "State Tax Account:"
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
      Top             =   3012
      Width           =   2196
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
Attribute VB_Name = "frmPrnSalesTaxRpt1"
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
  Unload frmPrnSalesTaxRpt
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
      SalesTaxReport
    ElseIf rptopt = 2 Then
      SalesTaxReport2
    End If

    fpcboCombined.SetFocus
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
  fpDate1.Text = Format(Now, "mm/dd/yyyy")
  fpDate2.Text = Format(Now, "mm/dd/yyyy")
  FillAccts fpcboCoAcct
  VendcoCodeList fpcboStateCode
  fpcboStateCode.ListIndex = 0
  fpcboRptType.InsertRow = "Graphics"
  fpcboRptType.InsertRow = "Text"
  fpcboRptType.ListIndex = 0
End Sub
Private Sub fpcboStAcct_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboStAcct.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcboStAcct.ListIndex = -1
    fpcboStAcct.Action = ActionClearSearchBuffer
  End If
  If fpcboStAcct.ListDown <> True Then
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

Private Sub fpcboCombined_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboCombined.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcboCombined.ListIndex = -1
    fpcboCombined.Action = ActionClearSearchBuffer
  End If
  If fpcboCombined.ListDown <> True Then
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

Private Sub fpcboCombined_LostFocus()
  If fpcboCombined.ListIndex = 0 Then
    fpdblFactor = ""
    fpdblFactor.Enabled = False
    fpcboCoAcct.Enabled = True
  Else
    fpdblFactor.Enabled = True
    fpcboCoAcct.ListIndex = -1
    fpcboCoAcct.Enabled = False
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
    If fpcboCombined.ListIndex = 0 Then
      If fpcboStAcct.ListIndex <> -1 Then
        If fpcboCoAcct.ListIndex <> -1 Then
          Oktogo = True
        Else
          MsgBox "Please Select County Account.", vbOKOnly, "Invalid Accounts"
          Oktogo = False
        End If
      Else
        MsgBox "Please Select State Account.", vbOKOnly, "Invalid Accounts"
        Oktogo = False
      End If
    Else
      If fpcboStAcct.ListIndex <> -1 Then
        If Val(fpdblFactor) <> 0 Then
          Oktogo = True
        Else
          Oktogo = False
          MsgBox "If Accounts are combined, You must enter the state factor percent.", vbOKOnly, "Invalid %"
        End If
      Else
        MsgBox "Please Select State Account.", vbOKOnly, "Invalid Accounts"
        Oktogo = False
      End If
    End If
End Function

Private Sub SalesTaxReport()
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
  If fpcboCombined.ListIndex = 1 Then Combined = True
  fpcboStAcct.col = 1
  fpcboCoAcct.col = 1
  User$ = QPTrim(GLUserName$)
  StateTaxRecAcct$ = QPTrim$(fpcboStAcct.ColText)
  CountyTaxRecAcct$ = QPTrim$(fpcboCoAcct.ColText)
  StateFactor# = Val(fpdblFactor)
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
   DeActivateControls frmPrnSalesTaxRpt, True
  ReDim StSalesTaxPaid#(0 To 999)

  If Not Combined Then
    ReDim CoSalesTaxPaid#(0 To 999)
  End If
  For State = 1 To TotList
    PrintFlag = 0
    FrmShowPctComp.ShowPctComp State, TotList
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ActivateControls frmPrnSalesTaxRpt, True
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
    CoTotTax# = 0
    For ll = 0 To 999
      StSalesTaxPaid#(ll) = 0
      If Not Combined Then
        CoSalesTaxPaid#(ll) = 0
      End If
    Next ll

    Dim Vendor As VendorRecType
    Close VendorFile
    OpenVendorFile VendorFile, NumVRecs

    Dim APLedger As APLedger81RecType
    LdRecLen = Len(APLedger)
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
          Get APLedgerFile, NextTran, APLedger
          If APLedger.TRCode = 1 Then
            '--check this out
            If APLedger.GLDistDate >= BegDate And APLedger.GLDistDate <= EndDate Then
              TCnt = TCnt + 1
              NextDist& = APLedger.FrstDist
              Do
                Get APDistFile, NextDist&, APDist

                If QPTrim$(APDist.DistAcctNum) = StateTaxRecAcct$ Then
                  Get VendorFile, APLedger.VRecNum, Vendor
                  'offset = State
                  
                  Coffset = Val(Vendor.CoCode)

                  If StateCd = QPTrim(Vendor.StCode) Then
                    SCnt = SCnt + 1
                    StTax# = StTax# + APDist.DistAmt
                    PrintFlag = 1
                    'offset = State
                    If all Then offset = 0
                    StSalesTaxPaid#(Coffset) = StSalesTaxPaid#(Coffset) + APDist.DistAmt
                    'LPRINT APDist.DistAmt
                  End If
                End If
                If Not Combined Then
                  If QPTrim$(APDist.DistAcctNum) = CountyTaxRecAcct$ Then
                    Get VendorFile, APLedger.VRecNum, Vendor
                    If StateCd = QPTrim(Vendor.StCode) Then
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
                End If
                NextDist& = APDist.NextDist
              Loop Until NextDist& = 0

            End If

          End If
          NextTran = APLedger.NextTrans
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
  ActivateControls frmPrnSalesTaxRpt, True
  ARptSalesTax.GetName ReportFile$
  ARptSalesTax.Label1.Caption = "Sales Tax Report"
  ARptSalesTax.txtDate.Caption = Now
  ARptSalesTax.txtTown.Caption = User$
  ARptSalesTax.Label17.Caption = "Reporting : " + fpDate1.Text + " thru " + fpDate2.Text
  ARptSalesTax.startrpt

 ' ViewPrint ReportFile$, Header$
 ' KillFile ReportFile$
  
  Exit Sub

PrintState:
  ToPrint$ = Space(80)
  
  ToPrint$ = StateCd + "~"

  If Combined Then
    For cnt = 0 To 999
      If StSalesTaxPaid#(cnt) > 0 Then
        ToPrint1$ = Space(80)
        StateAmt# = (StSalesTaxPaid#(cnt) * StateFactor#)
        CountyAmt# = (StSalesTaxPaid#(cnt) - StateAmt#)
        StTotTax# = StTotTax# + StSalesTaxPaid#(cnt)
        TotState# = TotState# + StateAmt#
        TotCounty# = TotCounty# + CountyAmt#
        ToPrint1$ = ToPrint$ + Str(cnt) + "~" + (Using("######.##", StateAmt#)) + "~" + (Using("######.##", CountyAmt#))
        Print #PRNFile, ToPrint1$
      End If
    Next
  Else
    For cnt = 0 To 999
      If StSalesTaxPaid#(cnt) > 0 Or CoSalesTaxPaid#(cnt) > 0 Then
        ToPrint1$ = Space(80)
        ToPrint1$ = ToPrint$ + Str(cnt) + "~" + (Using("######.##", StSalesTaxPaid#(cnt))) + "~" + (Using("######.##", CoSalesTaxPaid#(cnt)))
        Print #PRNFile, ToPrint1$
        StTotTax# = StTotTax# + StSalesTaxPaid#(cnt)
        CoTotTax# = CoTotTax# + CoSalesTaxPaid#(cnt)
      End If
    Next
  End If

  Return
CancelExit:
  Exit Sub
End Sub
Private Sub SalesTaxReport2()
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
  If fpcboCombined.ListIndex = 1 Then Combined = True
  fpcboStAcct.col = 1
  fpcboCoAcct.col = 1
  StateTaxRecAcct$ = QPTrim$(fpcboStAcct.ColText)
  CountyTaxRecAcct$ = QPTrim$(fpcboCoAcct.ColText)
  StateFactor# = Val(fpdblFactor)
  BegDate = DateDiff("d", "12/31/1979", fpDate1)
  EndDate = DateDiff("d", "12/31/1979", fpDate2)
  PRNFile = FreeFile
  Newrp = "SalTax"
  GetRPTName Newrp
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
   DeActivateControls frmPrnSalesTaxRpt, True
  ReDim StSalesTaxPaid#(0 To 999)

  If Not Combined Then
    ReDim CoSalesTaxPaid#(0 To 999)
  End If
  For State = 1 To TotList
    PrintFlag = 0
    FrmShowPctComp.ShowPctComp State, TotList
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ActivateControls frmPrnSalesTaxRpt, True
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
    CoTotTax# = 0
    For ll = 0 To 999
      StSalesTaxPaid#(ll) = 0
      If Not Combined Then
        CoSalesTaxPaid#(ll) = 0
      End If
    Next ll

    Dim Vendor As VendorRecType
    Close VendorFile
    OpenVendorFile VendorFile, NumVRecs

    Dim APLedger As APLedger81RecType
    LdRecLen = Len(APLedger)
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
          Get APLedgerFile, NextTran, APLedger
          If APLedger.TRCode = 1 Then
            '--check this out
            If APLedger.GLDistDate >= BegDate And APLedger.GLDistDate <= EndDate Then
              TCnt = TCnt + 1
              NextDist& = APLedger.FrstDist
              Do
                Get APDistFile, NextDist&, APDist

                If QPTrim$(APDist.DistAcctNum) = StateTaxRecAcct$ Then
                  Get VendorFile, APLedger.VRecNum, Vendor
                  'offset = State
                  
                  Coffset = Val(Vendor.CoCode)

                  If StateCd = QPTrim(Vendor.StCode) Then
                    SCnt = SCnt + 1
                    StTax# = StTax# + APDist.DistAmt
                    PrintFlag = 1
                    'offset = State
                    If all Then offset = 0
                    StSalesTaxPaid#(Coffset) = StSalesTaxPaid#(Coffset) + APDist.DistAmt
                    'LPRINT APDist.DistAmt
                  End If
                End If
                If Not Combined Then
                  If QPTrim$(APDist.DistAcctNum) = CountyTaxRecAcct$ Then
                    Get VendorFile, APLedger.VRecNum, Vendor
                    If StateCd = QPTrim(Vendor.StCode) Then
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
                End If
                NextDist& = APDist.NextDist
              Loop Until NextDist& = 0

            End If

          End If
          NextTran = APLedger.NextTrans
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
    Print #PRNFile, "No Information to Display"
  End If
  Close
  Unload FrmShowPctComp
  ActivateControls frmPrnSalesTaxRpt, True
  ViewPrint ReportFile$, Header$
  KillFile ReportFile$
  
  Exit Sub

PrintState:

  Print #PRNFile, "Sales Tax Report"
  Print #PRNFile, "State Code: "; StateCd
  Print #PRNFile,
  Print #PRNFile, "County Code          State          County"
  Print #PRNFile, "--------------------------------------------"

  If Combined Then
    For cnt = 0 To 999
      If StSalesTaxPaid#(cnt) > 0 Then
        StateAmt# = (StSalesTaxPaid#(cnt) * StateFactor#)
        CountyAmt# = (StSalesTaxPaid#(cnt) - StateAmt#)
        StTotTax# = StTotTax# + StSalesTaxPaid#(cnt)
        TotState# = TotState# + StateAmt#
        TotCounty# = TotCounty# + CountyAmt#
        Print #PRNFile, cnt; Tab(20); Using("######.##", StateAmt#); Tab(35); Using("######.##", CountyAmt#)
      End If
    Next
    Print #PRNFile, "--------------------------------------------"
    Print #PRNFile, "Totals"; Tab(20); Using("######.##", TotState#); Tab(35); Using("######.##", TotCounty#)
  Else
    For cnt = 0 To 999
      If StSalesTaxPaid#(cnt) > 0 Or CoSalesTaxPaid#(cnt) > 0 Then
        Print #PRNFile, cnt; Tab(20); Using("#####.##", StSalesTaxPaid#(cnt)); Tab(35); Using("#####.##", CoSalesTaxPaid#(cnt))
        StTotTax# = StTotTax# + StSalesTaxPaid#(cnt)
        CoTotTax# = CoTotTax# + CoSalesTaxPaid#(cnt)
      End If
    Next
    Print #PRNFile, "--------------------------------------------"
    Print #PRNFile, "State Totals"; Tab(19); Using("######.##", StTotTax#); Tab(34); Using("######.##", CoTotTax#)
  End If

  Print #PRNFile, Chr$(12)
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
   DeActivateControls frmPrnSalesTaxRpt, True
  ReDim StSalesTaxPaid#(0 To 999)
  If Not Combined Then
    ReDim CoSalesTaxPaid#(0 To 999)
  End If
  For State = 1 To TotList
    PrintFlag = 0
    FrmShowPctComp.ShowPctComp State, TotList
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ActivateControls frmPrnSalesTaxRpt, True
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

    Dim APLedger As APLedger81RecType
    LdRecLen = Len(APLedger)
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
          Get APLedgerFile, NextTran, APLedger
          If APLedger.TRCode = 1 Then
            '--check this out
            If APLedger.GLDistDate >= BegDate And APLedger.GLDistDate <= EndDate Then
              TCnt = TCnt + 1
              NextDist& = APLedger.FrstDist
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
                    Get VendorFile, APLedger.VRecNum, Vendor
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
          NextTran = APLedger.NextTrans
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
  ActivateControls frmPrnSalesTaxRpt, True
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

