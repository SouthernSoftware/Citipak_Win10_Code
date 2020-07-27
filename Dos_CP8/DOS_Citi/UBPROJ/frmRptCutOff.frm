VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmRptCutOff 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Cut-Off Report"
   ClientHeight    =   8640
   ClientLeft      =   36
   ClientTop       =   540
   ClientWidth     =   12192
   Icon            =   "frmRptCutOff.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   12192
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboRptFormat 
      Height          =   348
      Left            =   5544
      TabIndex        =   5
      Top             =   4920
      Width           =   3612
      _Version        =   196608
      _ExtentX        =   6371
      _ExtentY        =   614
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
      ColDesigner     =   "frmRptCutOff.frx":08CA
   End
   Begin LpLib.fpCombo fpcboBalType 
      Height          =   348
      Left            =   5544
      TabIndex        =   3
      Top             =   3876
      Width           =   3540
      _Version        =   196608
      _ExtentX        =   6244
      _ExtentY        =   614
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
      ColDesigner     =   "frmRptCutOff.frx":0C5D
   End
   Begin LpLib.fpCombo fpcboRptType 
      Height          =   348
      Left            =   5544
      TabIndex        =   7
      Top             =   5976
      Width           =   1908
      _Version        =   196608
      _ExtentX        =   3365
      _ExtentY        =   614
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
      ColDesigner     =   "frmRptCutOff.frx":0FF0
   End
   Begin LpLib.fpCombo fpcboPrintOrder 
      Height          =   348
      Left            =   5544
      TabIndex        =   4
      Top             =   4392
      Width           =   3612
      _Version        =   196608
      _ExtentX        =   6371
      _ExtentY        =   614
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
      ColDesigner     =   "frmRptCutOff.frx":138E
   End
   Begin LpLib.fpCombo fpcboMetInf 
      Height          =   348
      Left            =   5544
      TabIndex        =   6
      Top             =   5448
      Width           =   828
      _Version        =   196608
      _ExtentX        =   1460
      _ExtentY        =   614
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
      ColDesigner     =   "frmRptCutOff.frx":1721
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
      TabIndex        =   9
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
      TabIndex        =   8
      Top             =   7560
      Width           =   1332
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   10
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
            TextSave        =   "3:12 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7133
            TextSave        =   "6/17/2003"
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
   Begin EditLib.fpText fptxtRoute2 
      Height          =   348
      Left            =   5544
      TabIndex        =   1
      Top             =   2844
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
      Left            =   5544
      TabIndex        =   0
      Top             =   2328
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
      Left            =   5544
      TabIndex        =   2
      Top             =   3360
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
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Include Meter Info: "
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
      Left            =   2880
      TabIndex        =   19
      Top             =   5472
      Width           =   2556
   End
   Begin VB.Label LabelB2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Thru Book:"
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
      Left            =   3996
      TabIndex        =   18
      Top             =   2886
      Width           =   1380
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
      Left            =   3300
      TabIndex        =   17
      Top             =   3384
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
      Height          =   324
      Index           =   7
      Left            =   3660
      TabIndex        =   16
      Top             =   4452
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
      Height          =   324
      Left            =   3900
      TabIndex        =   15
      Top             =   2388
      Width           =   1476
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   4644
      Left            =   2472
      Top             =   1992
      Width           =   7284
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
      Left            =   3060
      TabIndex        =   14
      Top             =   6024
      Width           =   2388
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Print List Using:"
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
      Left            =   3636
      TabIndex        =   13
      Top             =   3882
      Width           =   1740
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Report Format:"
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
      Left            =   3336
      TabIndex        =   12
      Top             =   4950
      Width           =   2004
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3192
      Top             =   768
      Width           =   5772
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Print Customer Cut-Off Report"
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
      TabIndex        =   11
      Top             =   1008
      Width           =   5004
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   972
      Left            =   3192
      Top             =   648
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
Attribute VB_Name = "frmRptCutOff"
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
  frmUBReportsMenu.Show
  Unload frmRptCutOff
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        'ClearInUse PWcnt
      End If
    End If
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
    fpMinBal.SetFocus
  End If
End Sub

Private Sub fptxtRoute2_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub
Private Sub fpMinBal_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpcboBalType.SetFocus
  End If
End Sub
'this is Print using field
Private Sub fpcboBalType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboBalType.ListDown = True
  End If
  If fpcboBalType.ListDown <> True Then
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

Private Sub fpcboPrintOrder_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboPrintOrder.ListDown = True
  End If
  If fpcboPrintOrder.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpcboRptFormat.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpcboBalType.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub

Private Sub fpcboRptFormat_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboRptFormat.ListDown = True
  End If
  If fpcboRptFormat.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpcboMetInf.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpcboPrintOrder.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub
Private Sub fpcboMetInf_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboMetInf.ListDown = True
  End If
  If fpcboMetInf.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpcboRptType.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpcboRptFormat.SetFocus
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
        fpcboMetInf.SetFocus
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
      CustCutOffListing2
    ElseIf fpcboRptType.ListIndex = 1 Then
      CustCutOffListing
    End If
    ActivateControls Me, True
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
  Dim UBSetupreclen As Integer
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TownName$
  ReDim UBSetUp(1) As UBSetupRecType
  UBSetupreclen = Len(UBSetUp(1))
  LoadUBSetUpFile UBSetUp(), UBSetupreclen

  If UBSetUp(1).BILLCYCL = "Y" Then
    UseCycle = True
  End If

  If UseCycle Then
    LabelB1.Caption = "From Cycle"
    LabelB2.Caption = "Thru Cycle"
  Else
    LabelB1.Caption = "From Book"
    LabelB2.Caption = "Thru Book"
  End If
    fptxtRoute1 = "00"
    fptxtRoute2 = "99"

  fpcboBalType.AddItem "Total Balance"
  fpcboBalType.AddItem "Current Balance"
  fpcboBalType.AddItem "Past Due Balance"
  fpcboBalType.ListIndex = 0
  fpcboPrintOrder.AddItem "Customer Name Order"
  fpcboPrintOrder.AddItem "Account Number Order"
  fpcboPrintOrder.AddItem "Location Number Order"
  fpcboPrintOrder.AddItem "Sequence Number Order"
  fpcboPrintOrder.ListIndex = 0
  fpcboRptFormat.AddItem "1) Total Balance Listed"
  fpcboRptFormat.AddItem "2) All Balances Listed"
  fpcboRptFormat.ListIndex = 0
  fpcboMetInf.AddItem "Yes"
  fpcboMetInf.AddItem "No"
  fpcboMetInf.ListIndex = 0
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
Private Sub CustCutOffListing()
  Dim First As Integer, Last As Integer, MINAMT As Double, Order As String
  Dim BalanceType As String, RptTyp As Integer, IndexName As String
  Dim UsingName As Boolean, UsingAcct As Boolean, UsingBook As Boolean
  Dim UsingSeq As Boolean, Dash80 As String, UBCustRecLen As Integer
  Dim UBTransRecLen As Integer, UBSetupreclen As Integer
  Dim IdxRecLen As Integer, IdxFileSize As Long, cnt As Long
  Dim IdxNumOfRecs As Long, NumOfRecs As Long, Handle As Integer
  Dim UBCust As Integer, UBRpt As Integer, UBTrans As Integer
  Dim AcctNo As Long, CustCheck As Integer, CustomerCnt As Integer
  Dim CutOffBalance As Double, RealBalance As Double
  Dim TBalance As Double, MrtCnt As Integer, NOMetInf As Boolean
  Dim ReportFile As String
  PageNo = 0
  If fpcboMetInf.ListIndex = 0 Then
    NOMetInf = False
  Else
    NOMetInf = True
  End If
  First = BegRoute
  Last = EndRoute
  MINAMT# = fpMinBal.DoubleValue
  BalanceType$ = Left$(fpcboBalType.Text, 1)
  Order$ = Mid$(fpcboPrintOrder.Text, 1, 1)
  RptTyp = fpcboRptFormat.ListIndex + 1

      Select Case Order$
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
        IndexName$ = "UBTEMP.IDX"
        UsingSeq = True
      Case Else
        MsgBox "Invalid Printing Order.", vbOKOnly, "Invalid Selection"
        GoTo ExitCutOffListing
      End Select
      If RptTyp = 0 Then
        RptTyp = 1
      End If
  FrmShowPctComp.Label1 = "Creating Customer Cut-Off Report."
  FrmShowPctComp.Show , Me
  '***************
  MaxLines = 55

  PageNo = 0
  Dash80$ = String$(80, "-")

  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))
  ReDim UBTransRec(1) As UBTransRecType
  UBTransRecLen = Len(UBTransRec(1))
  ReDim UBSetUp(1) As UBSetupRecType
  UBSetupreclen = Len(UBSetUp(1))

  'AgeDate = Date2Num%(DATE$)

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

  ElseIf UsingSeq Then
    MakeSequenceIndex "Sequence Number", Me
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
    NumOfRecs = FileSize("UBCUST.DAT") \ UBCustRecLen
  End If

  UBCust = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen
  UBTrans = FreeFile
  Open UBPath$ + "UBTRANS.DAT" For Random Shared As UBTrans Len = UBTransRecLen
  ReportFile$ = UBPath$ + "UBCOLIST.RPT"
  UBRpt = FreeFile
  Open ReportFile$ For Output As UBRpt
'replace the longview flag with detail option on screen
'NOMetInf
''  If InStr(TownName$, "LONGVIEW") Then
''    LVFlag = True
''  End If

 
  GoSub DoCutOffRptHeader

  For cnt = 1 To NumOfRecs
    If UsingName Or UsingBook Or UsingSeq Then
      AcctNo& = IdxBuff(cnt).RecNum
    Else
      AcctNo& = cnt
    End If
    Get UBCust, AcctNo&, UBCustRec(1)
    FrmShowPctComp.ShowPctComp cnt, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      GoTo ExitCutOffListing
    End If

    'IF AcctNo& = 41 THEN STOP

    If UseCycle Then
      CustCheck = UBCustRec(1).BILLCYCL
    Else
      CustCheck = Val(UBCustRec(1).Book)
    End If
    If CustCheck >= First And CustCheck <= Last Then
      If UBCustRec(1).CUTOFFYN = "Y" Then
              If UBCustRec(1).Status = "A" Then
          Select Case BalanceType$
          Case "P"
            CutOffBalance# = UBCustRec(1).PrevBalance
          Case "C"
            CutOffBalance# = UBCustRec(1).CurrBalance
          Case "T"
            CutOffBalance# = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
          End Select
          RealBalance# = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
          If CutOffBalance# >= MINAMT# Then
            If RealBalance# > 0 And CutOffBalance# > 0 Then
              GoSub PrintLine
            End If
          End If
        End If
      End If
    End If
    If Linecnt > MaxLines Then
      Print #UBRpt, FF$
      GoSub DoCutOffRptHeader
    End If
  Next

  GoSub DoCutOffRptFooter:

  Close

  Erase IdxBuff, UBCustRec

    ViewPrint ReportFile$, "Customer Cut-Off Report."

  KillFile ReportFile$

ExitCutOffListing:
  Exit Sub

DoCutOffRptHeader:
  PageNo = PageNo + 1
  Print #UBRpt, TownName$
  Print #UBRpt, Tab(26); "Customer Cut Off Listing Report"; Tab(70); "Page #"; PageNo
  Print #UBRpt, "Report Date: "; Date$

  If RptTyp = 1 Then
    Print #UBRpt, "Location  Acct#  Customer Name"; Tab(43); "Service Address"; Tab(64); "Balance"
  Else
    Print #UBRpt, "Location  Acct#  Customer Name"; Tab(43); "Service Address";
    Print #UBRpt, Tab(47); "Previous     Current       Total"
  End If
  If Not NOMetInf Then
    Print #UBRpt, "Meter Number"; Tab(17); "Last Read"; Tab(30); "Reading at Cut-Off"
  End If

  Print #UBRpt, Dash80$
  Linecnt = 7
Return
DoCutOffRptFooter:
  Print #UBRpt, ""
  Print #UBRpt, "Total Customers to Cut Off: "; Using("#####,#", CustomerCnt)
  Print #UBRpt, "     Total Cut Off Balance: "; Using("#####.##", TBalance#)
  Print #UBRpt,
  Print #UBRpt, Tab(10); "Report Parameters *****************"
  Print #UBRpt, Tab(10); LabelB1.Caption; ": "; First; LabelB2.Caption; ": "; Last
  Print #UBRpt, Tab(10); "Printing Order: "; fpcboPrintOrder.Text; " "; "     Meter Info: "; fpcboMetInf.Text
  Print #UBRpt, Tab(10); "Balance Type: "; fpcboBalType.Text; "    Minimum Balance: "; Using("#####.##", MINAMT#)
  Print #UBRpt, Tab(10); "Report Format: "; fpcboRptFormat.Text
Return

PrintLine:
  TBalance# = Round#(TBalance# + RealBalance#)
  '*************************************
  '   Main body of Printing goes here
  Print #UBRpt, UBCustRec(1).Book; "-"; UBCustRec(1).SEQNUMB; Using("#####", AcctNo&);
  Print #UBRpt, Tab(16); Left$(UBCustRec(1).CustName, 22); " "; Left$(UBCustRec(1).SERVADDR, 22);
  If RptTyp = 2 Then
     Print #UBRpt,
    Print #UBRpt, Tab(46); Using("#####.##", UBCustRec(1).PrevBalance); Tab(58); Using("#####.##", UBCustRec(1).CurrBalance); Tab(70); Using("#####.##", RealBalance#)
  Else
    Print #UBRpt, Tab(70); Using("#####.##", RealBalance#)
  End If
  'PRINT #UBRpt, TAB(65); USING "($$#####.##)"; CutOffBalance#
  Linecnt = Linecnt + 2

  If NOMetInf Then
    GoTo Skip1
  End If

  For MrtCnt = 1 To 7
    If Len(QPTrim$(UBCustRec(1).LocMeters(MrtCnt).MTRType)) > 0 Then
      Print #UBRpt, UBCustRec(1).LocMeters(MrtCnt).MtrNum;
      Print #UBRpt, Tab(17); UBCustRec(1).LocMeters(MrtCnt).CurRead;
      Print #UBRpt, Tab(30); "_________________"
      Linecnt = Linecnt + 1
    End If
  Next
  Print #UBRpt, String$(79, "-")
   Linecnt = Linecnt + 1

Skip1:
  'If Landis Then
    Print #UBRpt,
    Linecnt = Linecnt + 2
  'End If
  CustomerCnt = CustomerCnt + 1
  '*************************************
  Return

End Sub
Private Sub CustCutOffListing2()
  Dim First As Integer, Last As Integer, MINAMT As Double
  Dim BalanceType As String, RptTyp As Integer, IndexName As String
  Dim UsingName As Boolean, UsingAcct As Boolean, UsingBook As Boolean
  Dim UsingSeq As Boolean, UBCustRecLen As Integer, Order As String
  Dim UBTransRecLen As Integer, UBSetupreclen As Integer
  Dim IdxRecLen As Integer, IdxFileSize As Long, cnt As Long
  Dim IdxNumOfRecs As Long, NumOfRecs As Long, Handle As Integer
  Dim UBCust As Integer, UBRpt As Integer, UBTrans As Integer
  Dim AcctNo As Long, CustCheck As Integer, CustomerCnt As Integer
  Dim CutOffBalance As Double, RealBalance As Double, ToPrint As String
  Dim TBalance As Double, MrtCnt As Integer, NOMetInf As Boolean
  Dim ToPrintM As String, ToPrintP As String, ToPrintP2 As String
  Dim ReportFile As String
  If fpcboMetInf.ListIndex = 0 Then
    NOMetInf = False
  Else
    NOMetInf = True
  End If
  First = BegRoute
  Last = EndRoute
  MINAMT# = fpMinBal.DoubleValue
  BalanceType$ = Left$(fpcboBalType.Text, 1)
  Order$ = Mid$(fpcboPrintOrder.Text, 1, 1)
  RptTyp = fpcboRptFormat.ListIndex + 1
  ToPrint$ = ""
  ToPrintM$ = ""
      Select Case Order$
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
        IndexName$ = "UBTEMP.IDX"
        UsingSeq = True
      Case Else
        MsgBox "Invalid Printing Order.", vbOKOnly, "Invalid Selection"
        GoTo ExitCutOffListing
      End Select
      If RptTyp = 0 Then
        RptTyp = 1
      End If
  FrmShowPctComp.Label1 = "Creating Customer Cut-Off Report."
  FrmShowPctComp.Show , Me
  '***************
  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))
  ReDim UBTransRec(1) As UBTransRecType
  UBTransRecLen = Len(UBTransRec(1))
  ReDim UBSetUp(1) As UBSetupRecType
  UBSetupreclen = Len(UBSetUp(1))
  
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

  ElseIf UsingSeq Then
    MakeSequenceIndex "Sequence Number", Me
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
    NumOfRecs = FileSize("UBCUST.DAT") \ UBCustRecLen
  End If

  UBCust = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen
  UBTrans = FreeFile
  Open UBPath$ + "UBTRANS.DAT" For Random Shared As UBTrans Len = UBTransRecLen
  ReportFile$ = UBPath$ + "UBCOLIST.RPT"
  UBRpt = FreeFile
  Open ReportFile$ For Output As UBRpt
'replace the longview flag with detail option on screen
'NOMetInf
''  If InStr(TownName$, "LONGVIEW") Then
''    LVFlag = True
''  End If


  For cnt = 1 To NumOfRecs
    If UsingName Or UsingBook Or UsingSeq Then
      AcctNo& = IdxBuff(cnt).RecNum
    Else
      AcctNo& = cnt
    End If
    Get UBCust, AcctNo&, UBCustRec(1)
    FrmShowPctComp.ShowPctComp cnt, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      GoTo ExitCutOffListing
    End If

    'IF AcctNo& = 41 THEN STOP

    If UseCycle Then
      CustCheck = UBCustRec(1).BILLCYCL
    Else
      CustCheck = Val(UBCustRec(1).Book)
    End If
    If CustCheck >= First And CustCheck <= Last Then
      If UBCustRec(1).CUTOFFYN = "Y" Then
              If UBCustRec(1).Status = "A" Then
          Select Case BalanceType$
          Case "P"
            CutOffBalance# = UBCustRec(1).PrevBalance
          Case "C"
            CutOffBalance# = UBCustRec(1).CurrBalance
          Case "T"
            CutOffBalance# = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
          End Select
          RealBalance# = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
          If CutOffBalance# >= MINAMT# Then
            If RealBalance# > 0 And CutOffBalance# > 0 Then
              GoSub PrintLine
              Print #UBRpt, ToPrint$ + "~" + ToPrintM$
              ToPrintM$ = ""
            End If
          End If
        End If
      End If
    End If
  Next

  Close

  Erase IdxBuff, UBCustRec
  If CustomerCnt > 0 Then
    Load frmLoadingRpt
    ARptCutOff.txtDate = Now
    ARptCutOff.txtTown = TownName$
    ARptCutOff.Title = "Customer Cut-Off Report."
    ARptCutOff.totCust = CustomerCnt
    ToPrintP$ = LabelB1.Caption + ": " + Str(First) + " " + LabelB2.Caption + ": " + Str(Last)
    ToPrintP$ = ToPrintP$ + "  Printing Order: " + fpcboPrintOrder.Text + " " + "     Meter Info: " + fpcboMetInf.Text
    ToPrintP2$ = "Balance Type: " + fpcboBalType.Text + "  Minimum Balance: " + Using("#####.##", MINAMT#)
    ToPrintP2$ = ToPrintP2$ + "  Report Format: " + fpcboRptFormat.Text
    ARptCutOff.txtRptParm1 = ToPrintP$
    ARptCutOff.txtRptParm2 = ToPrintP2$
    ARptCutOff.GetName ReportFile$, RptTyp, NOMetInf
    ARptCutOff.startrpt

  Else
    MsgBox "No Information to print.", vbOKOnly, "No Information"
  End If

ExitCutOffListing:
  Exit Sub

'DoCutOffRptHeader:
'  PageNo = PageNo + 1
'  Print #UBRpt, TownName$
'  Print #UBRpt, Tab(26); "Customer Cut Off Listing Report"; Tab(70); "Page #"; PageNo
'  Print #UBRpt, "Report Date: "; Date$
'
'  If RptTyp = 1 Then
'    Print #UBRpt, "Location  Acct#  Customer Name"; Tab(43); "Service Address"; Tab(64); "Balance"
'  Else
'    Print #UBRpt, "Location  Acct#  Customer Name"; Tab(43); "Service Address";
'    Print #UBRpt, Tab(47); "Previous     Current       Total"
'  End If
'  If Not NOMetInf Then
'    Print #UBRpt, "Meter Number"; Tab(17); "Last Read"; Tab(30); "Reading at Cut-Off"
'  End If
'
'  Print #UBRpt, Dash80$
'  Linecnt = 7
'Return
'DoCutOffRptFooter:
'  Print #UBRpt, ""
'  Print #UBRpt, "Total Customers to Cut Off: "; Using("#####,#", CustomerCnt)
'  Print #UBRpt, "     Total Cut Off Balance: "; Using("#####.##", TBalance#)
'  Print #UBRpt,
'  Print #UBRpt, Tab(10); "Report Parameters *****************"
'  Print #UBRpt, Tab(10); LabelB1.Caption; ": "; First; LabelB2.Caption; ": "; Last
'  Print #UBRpt, Tab(10); "Printing Order: "; fpcboPrintOrder.Text; " "; "     Meter Info: "; fpcboMetInf.Text
'  Print #UBRpt, Tab(10); "Balance Type: "; fpcboBalType.Text; "    Minimum Balance: "; Using("#####.##", MINAMT#)
'  Print #UBRpt, Tab(10); "Report Format: "; fpcboRptFormat.Text
'Return

PrintLine:
  ToPrint$ = ""
  TBalance# = Round#(TBalance# + RealBalance#)
  '*************************************
  '   Main body of Printing goes here
  ToPrint$ = UBCustRec(1).Book + "-" + UBCustRec(1).SEQNUMB + "~" + Str$(AcctNo&) + "~"
  ToPrint$ = ToPrint$ + Left$(UBCustRec(1).CustName, 30) + "~" + Left$(UBCustRec(1).SERVADDR, 22)
  If RptTyp = 2 Then
    ToPrint$ = ToPrint$ + "~" + Str$(UBCustRec(1).PrevBalance) + "~" + Str$(UBCustRec(1).CurrBalance) + "~" + Str$(RealBalance#)
  Else
    ToPrint$ = ToPrint$ + "~~~" + Str$(RealBalance#)
  End If
  'PRINT #UBRpt, TAB(65); USING "($$#####.##)"; CutOffBalance#
  'Linecnt = Linecnt + 2

  If NOMetInf Then
    GoTo Skip1
  End If

  For MrtCnt = 1 To 7
    
    If Len(QPTrim$(UBCustRec(1).LocMeters(MrtCnt).MTRType)) > 0 Then
      ToPrintM$ = ToPrintM$ + UBCustRec(1).LocMeters(MrtCnt).MtrNum + "~"
      ToPrintM$ = ToPrintM$ + Str$(UBCustRec(1).LocMeters(MrtCnt).CurRead) + "~"
      ToPrintM$ = ToPrintM$ + "__________________" + "~"
    Else
      ToPrintM$ = ToPrintM$ + " ~ ~ ~ "
    End If
  'ToPrintM$ = ToPrintM$ + "~" + ToPrintM$
  Next
  CustomerCnt = CustomerCnt + 1
Return
Skip1:
  CustomerCnt = CustomerCnt + 1
  '*************************************
  ToPrintM$ = " ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ "
  Return

End Sub

