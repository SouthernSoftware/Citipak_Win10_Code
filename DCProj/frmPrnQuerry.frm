VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "EDT32X30.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Begin VB.Form frmPrnQuery 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transaction Query"
   ClientHeight    =   8640
   ClientLeft      =   36
   ClientTop       =   540
   ClientWidth     =   12192
   Icon            =   "frmPrnQuerry.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   12192
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboRptType 
      Height          =   384
      Left            =   5376
      TabIndex        =   11
      Top             =   6864
      Width           =   1908
      _Version        =   196608
      _ExtentX        =   3365
      _ExtentY        =   677
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
      ColDesigner     =   "frmPrnQuerry.frx":08CA
   End
   Begin LpLib.fpCombo fpcboAcctType 
      Height          =   384
      Left            =   5376
      TabIndex        =   4
      Top             =   3552
      Width           =   1884
      _Version        =   196608
      _ExtentX        =   3323
      _ExtentY        =   677
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
      EditMarginLeft  =   1
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmPrnQuerry.frx":0C30
   End
   Begin LpLib.fpCombo fpcboFund1 
      Height          =   384
      Left            =   5376
      TabIndex        =   2
      Top             =   2376
      Width           =   3612
      _Version        =   196608
      _ExtentX        =   6371
      _ExtentY        =   677
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
      ColumnSearch    =   2
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
      ColDesigner     =   "frmPrnQuerry.frx":0F5F
   End
   Begin LpLib.fpCombo fpcboFund2 
      Height          =   384
      Left            =   5376
      TabIndex        =   3
      Top             =   2952
      Width           =   3612
      _Version        =   196608
      _ExtentX        =   6371
      _ExtentY        =   677
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
      Columns         =   3
      Sorted          =   0
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   2
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
      AutoSearchFillDelay=   500
      EditMarginLeft  =   25
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmPrnQuerry.frx":1312
   End
   Begin LpLib.fpCombo fpcboJournalRef 
      Height          =   384
      Left            =   5376
      TabIndex        =   9
      Top             =   6312
      Width           =   1068
      _Version        =   196608
      _ExtentX        =   1884
      _ExtentY        =   677
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
      ColDesigner     =   "frmPrnQuerry.frx":16C5
   End
   Begin VB.TextBox txtDescription 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   408
      Left            =   5376
      TabIndex        =   7
      Top             =   5160
      Width           =   2652
   End
   Begin VB.TextBox txtReference 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   408
      Left            =   5376
      TabIndex        =   8
      Top             =   5748
      Width           =   1164
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
            TextSave        =   "6:02 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7133
            TextSave        =   "4/2/03"
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
   Begin EditLib.fpDateTime txtDate2 
      Height          =   372
      Left            =   5376
      TabIndex        =   1
      Top             =   1800
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
   Begin EditLib.fpDateTime txtDate1 
      Height          =   372
      Left            =   5376
      TabIndex        =   0
      Top             =   1272
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
   Begin EditLib.fpDateTime txtPostDate 
      Height          =   372
      Left            =   7920
      TabIndex        =   10
      Top             =   6312
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
   Begin EditLib.fpMask txtAcctCode 
      Height          =   372
      Left            =   5376
      TabIndex        =   5
      Top             =   4104
      Width           =   1092
      _Version        =   196608
      _ExtentX        =   1926
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
      BackColor       =   -2147483634
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
      InvalidColor    =   -2147483634
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483639
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   0
      AllowOverflow   =   0   'False
      BestFit         =   0   'False
      ClipMode        =   0
      DataFormatEx    =   0
      Mask            =   ""
      PromptChar      =   "_"
      PromptInclude   =   0   'False
      RequireFill     =   -1  'True
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      AutoTab         =   0   'False
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpMask txtObjCode 
      Height          =   372
      Left            =   5376
      TabIndex        =   6
      Top             =   4632
      Width           =   1068
      _Version        =   196608
      _ExtentX        =   1884
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
      BackColor       =   -2147483634
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
      InvalidColor    =   -2147483634
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483639
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   0
      AllowOverflow   =   0   'False
      BestFit         =   0   'False
      ClipMode        =   0
      DataFormatEx    =   0
      Mask            =   ""
      PromptChar      =   "_"
      PromptInclude   =   0   'False
      RequireFill     =   -1  'True
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      AutoTab         =   -1  'True
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
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
      Left            =   2928
      TabIndex        =   28
      Top             =   6912
      Width           =   2388
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   6252
      Left            =   2376
      Top             =   1152
      Width           =   7452
   End
   Begin VB.Line Line3 
      X1              =   4272
      X2              =   4560
      Y1              =   2616
      Y2              =   2616
   End
   Begin VB.Line Line2 
      X1              =   4272
      X2              =   4800
      Y1              =   3096
      Y2              =   3096
   End
   Begin VB.Line Line1 
      X1              =   4272
      X2              =   4272
      Y1              =   3096
      Y2              =   2616
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "From:"
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
      Left            =   4608
      TabIndex        =   27
      Top             =   2460
      Width           =   636
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Journal Reference:"
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
      Left            =   3024
      TabIndex        =   26
      Top             =   6360
      Width           =   2220
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Post Date:"
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
      Index           =   6
      Left            =   6528
      TabIndex        =   25
      Top             =   6360
      Width           =   1260
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Reference:"
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
      Index           =   5
      Left            =   3552
      TabIndex        =   24
      Top             =   5796
      Width           =   1692
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Description:"
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
      Index           =   4
      Left            =   3552
      TabIndex        =   23
      Top             =   5232
      Width           =   1692
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
      Left            =   3576
      TabIndex        =   22
      Top             =   1320
      Width           =   1668
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Account Type:"
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
      Index           =   3
      Left            =   3552
      TabIndex        =   21
      Top             =   3600
      Width           =   1692
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Object Code:"
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
      Left            =   3552
      TabIndex        =   20
      Top             =   4704
      Width           =   1692
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Acct Code:"
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
      Left            =   3792
      TabIndex        =   19
      Top             =   4140
      Width           =   1452
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fund Codes"
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
      Left            =   2784
      TabIndex        =   18
      Top             =   2712
      Width           =   1428
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
      Height          =   324
      Index           =   0
      Left            =   3672
      TabIndex        =   17
      Top             =   1860
      Width           =   1572
   End
   Begin VB.Image Image1 
      Height          =   276
      Left            =   8064
      Picture         =   "frmPrnQuerry.frx":1A20
      Top             =   1512
      Width           =   288
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
      Caption         =   "Print Transaction Query"
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
      Caption         =   "To:"
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
      Index           =   1
      Left            =   4752
      TabIndex        =   15
      Top             =   2952
      Width           =   492
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
Attribute VB_Name = "frmPrnQuery"
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
Dim GLTrans   As GLTransRecType
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Dim FY1BegDate As Integer, FY1EndDate As Integer, FY2BegDate As Integer, FY2EndDate As Integer
Dim StartFund As String, EndFund As String, FYStartDate As Integer
Dim ActiveYear As Integer
Dim acctmsk As String, detmsk As String
Private Sub cmdExit_Click()
  frmGLReportsMenu.Show
  Unload frmPrnQuery
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
Private Function ValidFunds()
  If fpcboFund1.Text <> "" And fpcboFund2.Text <> "" Then
    fpcboFund1.Col = 0
    fpcboFund2.Col = 0
    If fpcboFund1.ColText > fpcboFund2.ColText Then
      MsgBox "Invalid Fund Selection, The Beginning Fund Should Be Less or Equal to Ending Fund.", vbOKOnly, "Invalid Selection"
      ValidFunds = False
    Else
      ValidFunds = True
      StartFund = QPTrim(fpcboFund1.ColText)
      EndFund = QPTrim(fpcboFund2.ColText)
    End If
  Else
    MsgBox "Fund Selections May Not Be Left Blank.", vbOKOnly, "Invalid Selection"
  End If
End Function

Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub

Private Sub txtAcctCode_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub txtAcctCode_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    txtObjCode.SetFocus
  End If
End Sub

Private Sub txtAcctCode_LostFocus()
  Dim Num As String
  Num = Trim(txtAcctCode)
  If (Len(Num)) > 1 Then
    If (Len(Num)) <> (Val(GLAcctLen)) Then
      MsgBox "Invalid Code.", vbOKOnly, "Invalid Data!"
      txtAcctCode.Mask = acctmsk
      txtAcctCode.SetFocus
    End If
  End If
End Sub

Private Sub txtDate1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    txtDate2.SetFocus
  End If
End Sub

Private Sub txtDate2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpcboFund1.SetFocus
  End If
End Sub

Private Sub txtDescription_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
    txtReference.SetFocus
  End If
  If KeyCode = vbKeyUp Then
    txtObjCode.SetFocus
  End If
End Sub

Private Sub txtObjCode_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub txtObjCode_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    txtDescription.SetFocus
  End If
End Sub

Private Sub txtObjCode_LostFocus()
  Dim Num As String
  Num = Trim(txtObjCode)
  If (Len(Num)) > 1 Then
    If (Len(Num)) <> (Val(GLDetLen)) Then
      MsgBox "Invalid Code.", vbOKOnly, "Invalid Data!"
      txtObjCode.Mask = detmsk
      txtObjCode.SetFocus
    End If
  End If
End Sub

Private Sub cmdPrint_Click()
  If ValidDate = True Then
    If ValidFunds = True Then
      If fpcboRptType.ListIndex = 0 Then
        rptopt = 1
      ElseIf fpcboRptType.ListIndex = 1 Then
        rptopt = 2
      End If
      If rptopt = 1 Then
        PrintAQuery
      ElseIf rptopt = 2 Then
        PrintAQuery2
      End If
    End If
  End If
End Sub
Private Sub fpcboFund2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboFund2.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcboFund2.ListIndex = -1
    fpcboFund2.Action = ActionClearSearchBuffer
  End If
  If fpcboFund2.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpcboAcctType.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpcboFund1.SetFocus
        KeyCode = 0
      End If
    End If
  End If

End Sub
Private Sub fpcboFund1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboFund1.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcboFund1.ListIndex = -1
    fpcboFund1.Action = ActionClearSearchBuffer
  End If
  If fpcboFund1.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpcboFund2.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        txtDate2.SetFocus
        KeyCode = 0
      End If
    End If
  End If

End Sub
Private Sub fpcboAcctType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboAcctType.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcboAcctType.ListIndex = -1
    fpcboAcctType.Action = ActionClearSearchBuffer
  End If
  If fpcboAcctType.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      txtAcctCode.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpcboFund2.SetFocus
        KeyCode = 0
      End If
    End If
  End If

End Sub
Private Sub fpcboJournalRef_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboJournalRef.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcboJournalRef.ListIndex = -1
    fpcboJournalRef.Action = ActionClearSearchBuffer
  End If
  If fpcboJournalRef.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      txtPostDate.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        txtReference.SetFocus
        KeyCode = 0
      End If
    End If
  End If

End Sub
Private Sub txtPostDate_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDelete Then
    txtPostDate.Text = ""
  End If
  If KeyCode = vbKeyReturn Then
    fpcboRptType.SetFocus
  End If
End Sub
Private Sub txtReference_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
    fpcboJournalRef.SetFocus
  End If
  If KeyCode = vbKeyUp Then
    txtDescription.SetFocus
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
        txtPostDate.SetFocus
        KeyCode = 0
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
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  detmsk = String(GLDetLen, "#")
  acctmsk = String(GLAcctLen, "#")
  StatusBar1.Panels.Item(1).Text = GLUserName
  FundstoList fpcboFund1
  FundstoList fpcboFund2
  txtDate1.Text = Format(Now, "mm/dd/yyyy")
  txtDate2.Text = Format(Now, "mm/dd/yyyy")
  fpcboAcctType.AddItem "All"
  fpcboAcctType.AddItem "Asset"
  fpcboAcctType.AddItem "Liability"
  fpcboAcctType.AddItem "Revenue"
  fpcboAcctType.AddItem "Expenditure"
  fpcboAcctType.ListIndex = 0
  txtAcctCode.Mask = acctmsk
  txtObjCode.Mask = detmsk
  fpcboJournalRef.AddItem "Any"
  fpcboJournalRef.AddItem "AR"
  fpcboJournalRef.AddItem "AP"
  fpcboJournalRef.AddItem "CD"
  fpcboJournalRef.AddItem "CK"
  fpcboJournalRef.AddItem "CR"
  fpcboJournalRef.AddItem "GJ"
  fpcboJournalRef.AddItem "IF"
  fpcboJournalRef.AddItem "PR"
  fpcboJournalRef.AddItem "VC"
  fpcboJournalRef.AddItem "VI"
  fpcboJournalRef.ListIndex = 0
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

Private Sub PrintAQuery()
  Dim CommaFmt As String, TotalFmt As String, RptTitle As String
  Dim SumLine As String, FF As String, PRNFile As Integer, Match As Long
  Dim MaxLines As Integer, TotDr As Double, TotCr As Double, RecNum As String
  Dim HollyFlag As Boolean, Pitch12 As String, AType As String, Newrp As String
  Dim ReportFile As String, AT As String, Src As String, Sc As String
  Dim FundIdxFileNum As Integer, NumFunds As Integer, Acct As Integer
  Dim AcctIdxFileNum As Integer, NumGLAccts As Integer, FundName As String
  Dim AcctFileNum As Integer, NumGLAcctRecs As Integer, Rec As Integer
  Dim TransFileNum As Integer, NumTrans As Long, NextTr As Long, AcctNum As String
  Dim Fund As String, Det As String, FundRec As Integer, NumFIdxRecs As Integer
  Dim Dept As String, cnt As Integer, AC As String, Obj As String, TSrch As String
  Dim AcctNumber As String, AcctCode As String, ObjCode As String, TRef As String
  Dim Pct As String, ToPrint As String, Linecnt As Integer, PostDate As String
  Dim LowDate As Integer, HighDate As Integer, TrDateStamp As String
  Dim Diff As Double, Bal As Double, DrCr As String, PageNum As Integer
'  If InStr(UCase$(GLUserName), "HOLLY SPR") > 0 Then
'    HollyFlag = True
'    Pitch12$ = Chr$(27) + Chr$(38) + Chr$(107) + Chr$(52) + Chr$(83)
'  End If
  LowDate = DateDiff("d", "12/31/1979", txtDate1)
  HighDate = DateDiff("d", "12/31/1979", txtDate2)
  CommaFmt$ = "###,###,###.##"  'format takes 14 chars
  TotalFmt$ = "#,###,###,###.##" 'format takes 16 chars
  SumLine$ = String$(16, "-")   'column summary line
  TSrch$ = QPTrim(txtDescription)
  PostDate$ = Mid$(txtPostDate.DateValue, 5, 4) + Mid$(txtPostDate.DateValue, 3, 2)
  MaxLines = 60
  FF$ = Chr$(12)
  RptTitle$ = "Transaction Query"
  TotDr# = 0
  TotCr# = 0
  TRef$ = QPTrim(txtReference)
  Src$ = QPTrim(fpcboJournalRef.Text)

  AType$ = fpcboAcctType.Text

  If AType$ = "All" Then
    AType$ = ""
    AT$ = "All"
  Else
    AType$ = Left$(AType$, 1)
    AT$ = AType$
  End If

  If Src$ = "Any" Then
    Src$ = ""
    Sc$ = "Any "
  Else
    Sc$ = Src$
  End If
  AcctCode$ = txtAcctCode.Text
  If AcctCode$ = "" Then
    AC$ = "All "
  Else
    AC$ = AcctCode$
  End If
  ObjCode$ = txtObjCode.Text
  If ObjCode$ = "" Then
    Obj$ = "All "
  Else
    Obj$ = ObjCode$
  End If

  'End of Input
  '=====================================================
  'Start Report Processing
  FrmShowPctComp.Label1 = "Printing Transaction Query Report"
  FrmShowPctComp.Show , Me
  DoEvents
  DeActivateControls frmPrnQuery, True
  Newrp = "Query"
  GetRPTName Newrp
  ReportFile$ = Newrp
  'ReportFile$ = Unique$(Path$)

  PRNFile = FreeFile
  Open ReportFile$ For Output As #PRNFile
  OpenAcctIdx AcctIdxFileNum, NumGLAccts
  OpenAcctFile AcctFileNum, NumGLAcctRecs
  OpenTransFile TransFileNum, NumTrans&
  OpenFundIdx FundIdxFileNum, NumFIdxRecs

  'GoSub PrintQueryPageHeader

  For cnt = 1 To NumGLAccts     'NumGLAccts
    Get AcctIdxFileNum, cnt, GLAcctidx

    AcctNumber$ = QPTrim$(GLAcctidx.AcctNum)

    Fund$ = Left$(AcctNumber$, GLFundLen)
    Dept$ = Mid$(AcctNumber$, GLFundLen + 2, GLAcctLen)
    Det$ = Right$(AcctNumber$, GLDetLen)

    If Fund$ >= StartFund$ And Fund$ <= EndFund$ Then
      If InStr(Dept$, AcctCode$) And InStr(Det$, ObjCode$) Then

        Get AcctFileNum, GLAcctidx.RecNum, GLAcct
        If InStr(GLAcct.Typ, AType$) Then

          NextTr& = GLAcct.FrstTran               'get the first trans for this
          Do Until NextTr& = 0  'keep going 'til we run out of trans
            Get TransFileNum, NextTr&, GLTrans

            If GLTrans.TRDATE >= LowDate And GLTrans.TRDATE <= HighDate Then
              TrDateStamp$ = Mid$(GLTrans.Src, 3, 6)
              If InStr(TrDateStamp$, PostDate$) Or PostDate$ = "000000" Then
                If InStr(GLTrans.Src, Src$) Then
                  If InStr(UCase$(GLTrans.Desc), UCase$(TSrch$)) Then
                    If InStr(1, GLTrans.Ref, TRef$) Then
                      Match = Match + 1
                      TotDr# = TotDr# + Round#(GLTrans.DrAmt)
                      TotCr# = TotCr# + Round#(GLTrans.CrAmt)

                      ToPrint$ = Space$(100)
                      ToPrint$ = Format(DateAdd("d", GLTrans.TRDATE, "12-31-1979"), "mm/dd/yy")
                      ToPrint$ = ToPrint$ + "~" + QPTrim(GLTrans.AcctNum)
                      ToPrint$ = ToPrint$ + "~" + QPTrim(Left$(GLTrans.Desc, 15))
                      ToPrint$ = ToPrint$ + "~" + QPTrim(GLTrans.Ref)
                      If GLTrans.DrAmt <> 0 Then
                        ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, QPTrim(Str$(GLTrans.DrAmt)))
                      Else
                        ToPrint$ = ToPrint$ + "~" + "0"
             End If
             If GLTrans.CrAmt <> 0 Then
               ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, QPTrim(Str$(GLTrans.CrAmt)))
             Else
              ToPrint$ = ToPrint$ + "~" + "0"
             End If
             'Mid$(ToPrint$, 82) = RecNum$
             ToPrint$ = ToPrint$ + "~" + QPTrim(GLTrans.Src)
             Print #PRNFile, ToPrint$
             
           End If
         End If        'kill me
       End If
     End If
   End If

   NextTr& = GLTrans.NextTran            'Get the next transaction

          Loop
        End If
      End If
    End If
    FrmShowPctComp.ShowPctComp cnt, NumGLAccts
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ActivateControls frmPrnQuery, True
      Unload FrmShowPctComp
      GoTo CancelExit
    End If
  Next          'Process next account

  ActivateControls frmPrnQuery, True

  Diff# = Round#(TotDr# - TotCr#)
  Bal# = Abs(Diff#)

  If Diff# = 0 Then
    DrCr$ = ""
  ElseIf Diff# > 0 Then
    DrCr$ = " Dr"
  Else
    DrCr$ = " Cr"
  End If

'  Print #PRNFile, String$(96, "-")
'  Print #PRNFile, "Totals";
'  Print #PRNFile, Tab(50);
'  Print #PRNFile, Tab(65);
'  Print #PRNFile,
'  Print #PRNFile, "Report Summary"
''''''''''''''''''''123456789012345678901234567890123456789012345678901234567890
'  Print #PRNFile, "Balance:            ",
'  Print #PRNFile, DrCr$
'
'  Print #PRNFile, "Number of transactions:    ",
'  Print #PRNFile, FF$

  Close
  Load frmLoadingRpt
  ARptQuery.totDebits = Using$(TotalFmt$, Round#(TotDr#))
  ARptQuery.totCredits = Using$(TotalFmt$, Round#(TotCr#))
  ARptQuery.totBal = Using$(TotalFmt$, Round#(Bal#))
  ARptQuery.totTrans = Using$("##,###", Match)
  ARptQuery.txtRptDate = "Date Range: " + txtDate1.Text + " thru " + txtDate2.Text + "  Module: " + Sc$ + "  Funds " + StartFund$ + " thru " + EndFund$ + "  Account: " + AC$ + "  Object: " + Obj$ + "  Account Type: " + AT$
  ARptQuery.txtDate = Now
  ARptQuery.txtTown = GLUserName$
  ARptQuery.GetName ReportFile$
  ARptQuery.startrpt

  Exit Sub
CancelExit:
  Exit Sub
End Sub
Private Sub PrintAQuery2()
  Dim CommaFmt As String, TotalFmt As String, RptTitle As String
  Dim SumLine As String, FF As String, PRNFile As Integer, Match As Long
  Dim MaxLines As Integer, TotDr As Double, TotCr As Double, RecNum As String
  Dim HollyFlag As Boolean, Pitch12 As String, AType As String, Newrp As String
  Dim ReportFile As String, AT As String, Src As String, Sc As String
  Dim FundIdxFileNum As Integer, NumFunds As Integer, Acct As Integer
  Dim AcctIdxFileNum As Integer, NumGLAccts As Integer, FundName As String
  Dim AcctFileNum As Integer, NumGLAcctRecs As Integer, Rec As Integer
  Dim TransFileNum As Integer, NumTrans As Long, NextTr As Long, AcctNum As String
  Dim Fund As String, Det As String, FundRec As Integer, NumFIdxRecs As Integer
  Dim Dept As String, cnt As Integer, AC As String, Obj As String, TSrch As String
  Dim AcctNumber As String, AcctCode As String, ObjCode As String, TRef As String
  Dim Pct As String, ToPrint As String, Linecnt As Integer, PostDate As String
  Dim LowDate As Integer, HighDate As Integer, TrDateStamp As String
  Dim Diff As Double, Bal As Double, DrCr As String, PageNum As Integer
'  If InStr(UCase$(GLUserName), "HOLLY SPR") > 0 Then
'    HollyFlag = True
'    Pitch12$ = Chr$(27) + Chr$(38) + Chr$(107) + Chr$(52) + Chr$(83)
'  End If
  LowDate = DateDiff("d", "12/31/1979", txtDate1)
  HighDate = DateDiff("d", "12/31/1979", txtDate2)
  CommaFmt$ = "###,###,###.##"  'format takes 14 chars
  TotalFmt$ = "#,###,###,###.##" 'format takes 16 chars
  SumLine$ = String$(16, "-")   'column summary line
  TSrch$ = QPTrim(txtDescription)
  PostDate$ = Mid$(txtPostDate.DateValue, 5, 4) + Mid$(txtPostDate.DateValue, 3, 2)
  MaxLines = 60
  FF$ = Chr$(12)
  RptTitle$ = "Transaction Query"
  TotDr# = 0
  TotCr# = 0
  TRef$ = QPTrim(txtReference)
  Src$ = QPTrim(fpcboJournalRef.Text)

  AType$ = fpcboAcctType.Text

  If AType$ = "All" Then
    AType$ = ""
    AT$ = "All"
  Else
    AType$ = Left$(AType$, 1)
    AT$ = AType$
  End If

  If Src$ = "Any" Then
    Src$ = ""
    Sc$ = "Any "
  Else
    Sc$ = Src$
  End If
  AcctCode$ = txtAcctCode.Text
  If AcctCode$ = "" Then
    AC$ = "All "
  Else
    AC$ = AcctCode$
  End If
  ObjCode$ = txtObjCode.Text
  If ObjCode$ = "" Then
    Obj$ = "All "
  Else
    Obj$ = ObjCode$
  End If

  'End of Input
  '=====================================================
  'Start Report Processing
  FrmShowPctComp.Label1 = "Printing Transaction Query Report"
  FrmShowPctComp.Show , Me
  DoEvents
  DeActivateControls frmPrnQuery, True
  Newrp = "Query"
  GetRPTName Newrp
  ReportFile$ = Newrp
  'ReportFile$ = Unique$(Path$)

  PRNFile = FreeFile
  Open ReportFile$ For Output As #PRNFile
  OpenAcctIdx AcctIdxFileNum, NumGLAccts
  OpenAcctFile AcctFileNum, NumGLAcctRecs
  OpenTransFile TransFileNum, NumTrans&
  OpenFundIdx FundIdxFileNum, NumFIdxRecs
  If HollyFlag Then
    Print #PRNFile, Pitch12$;
  End If

  GoSub PrintQueryPageHeader

  For cnt = 1 To NumGLAccts     'NumGLAccts
    Get AcctIdxFileNum, cnt, GLAcctidx

    AcctNumber$ = QPTrim$(GLAcctidx.AcctNum)

    Fund$ = Left$(AcctNumber$, GLFundLen)
    Dept$ = Mid$(AcctNumber$, GLFundLen + 2, GLAcctLen)
    Det$ = Right$(AcctNumber$, GLDetLen)

    If Fund$ >= StartFund$ And Fund$ <= EndFund$ Then
      If InStr(Dept$, AcctCode$) And InStr(Det$, ObjCode$) Then

        Get AcctFileNum, GLAcctidx.RecNum, GLAcct
        If InStr(GLAcct.Typ, AType$) Then

          NextTr& = GLAcct.FrstTran               'get the first trans for this
          Do Until NextTr& = 0  'keep going 'til we run out of trans
            Get TransFileNum, NextTr&, GLTrans

            If GLTrans.TRDATE >= LowDate And GLTrans.TRDATE <= HighDate Then
              TrDateStamp$ = Mid$(GLTrans.Src, 3, 6)
              If InStr(TrDateStamp$, PostDate$) Or PostDate$ = "000000" Then
                If InStr(GLTrans.Src, Src$) Then
                  If InStr(UCase$(GLTrans.Desc), UCase$(TSrch$)) Then
                    If InStr(1, GLTrans.Ref, TRef$) Then
                      Match = Match + 1
                      TotDr# = TotDr# + Round#(GLTrans.DrAmt)
                      TotCr# = TotCr# + Round#(GLTrans.CrAmt)

                      ToPrint$ = Space$(100)
                      LSet ToPrint$ = Format(DateAdd("d", GLTrans.TRDATE, "12-31-1979"), "mm/dd/yy")
                      Mid$(ToPrint$, 11) = QPTrim(GLTrans.AcctNum)
                      Mid$(ToPrint$, 23) = QPTrim(Left$(GLTrans.Desc, 15))
                      Mid$(ToPrint$, 41) = QPTrim(GLTrans.Ref)
                      If GLTrans.DrAmt <> 0 Then
                        Mid$(ToPrint$, 52) = Using$(CommaFmt$, QPTrim(Str$(GLTrans.DrAmt)))
                      Else
                        Mid$(ToPrint$, 52) = ""
             End If
             If GLTrans.CrAmt <> 0 Then
               Mid$(ToPrint$, 67) = Using$(CommaFmt$, QPTrim(Str$(GLTrans.CrAmt)))
             Else
               Mid$(ToPrint$, 67) = ""
             End If
             'Mid$(ToPrint$, 82) = RecNum$
             Mid$(ToPrint$, 88) = QPTrim(GLTrans.Src)
             Print #PRNFile, ToPrint$
             Linecnt = Linecnt + 1
             If Linecnt > MaxLines Then
               Print #PRNFile, FF$
               GoSub PrintQueryPageHeader
             End If
           End If
         End If        'kill me
       End If
     End If
   End If

   NextTr& = GLTrans.NextTran            'Get the next transaction

          Loop
        End If
      End If
    End If
    FrmShowPctComp.ShowPctComp cnt, NumGLAccts
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ActivateControls frmPrnQuery, True
      Unload FrmShowPctComp
      GoTo CancelExit
    End If
  Next          'Process next account

  ActivateControls frmPrnQuery, True

  Diff# = Round#(TotDr# - TotCr#)
  Bal# = Abs(Diff#)

  If Diff# = 0 Then
    DrCr$ = ""
  ElseIf Diff# > 0 Then
    DrCr$ = " Dr"
  Else
    DrCr$ = " Cr"
  End If

  Print #PRNFile, String$(96, "-")
  Print #PRNFile, "Totals";
  Print #PRNFile, Tab(50); Using$(TotalFmt$, Round#(TotDr#))
  Print #PRNFile, Tab(65); Using$(TotalFmt$, Round#(TotCr#))
  Print #PRNFile,
  Print #PRNFile, "Report Summary"
'''''''''''''''''''123456789012345678901234567890123456789012345678901234567890
  Print #PRNFile, "Balance:            ", Using$(TotalFmt$, Round#(Bal#))
  Print #PRNFile, DrCr$

  Print #PRNFile, "Number of transactions:    ", Using$("##,###", Match)
  Print #PRNFile, FF$

  Close

  ViewPrint ReportFile$, RptTitle$, True

  KillFile ReportFile$

  Exit Sub

PrintQueryPageHeader:
  PageNum = PageNum + 1
  Print #PRNFile, "Transaction Query"
  Print #PRNFile, Tab(68); "Page: "; PageNum
  Print #PRNFile, "Date Range: " + txtDate1.Text + " thru " + txtDate2.Text + "  " + Space$(2) + "Module: " + Sc$
  Print #PRNFile, "Funds " + StartFund$ + " thru " + EndFund$ + "  Account: " + AC$ + "  Object: " + Obj$ + "  Account Type: " + AT$
  Print #PRNFile, "Date        Acct No        Desc          Ref               Debit        Credit          Source"
  Print #PRNFile, String$(96, "-")
  Linecnt = 6
  Return

CancelExit:
  Exit Sub
End Sub

