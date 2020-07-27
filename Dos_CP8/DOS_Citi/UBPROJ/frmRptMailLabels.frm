VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmRptMailLabels 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mailing Labels"
   ClientHeight    =   8640
   ClientLeft      =   36
   ClientTop       =   540
   ClientWidth     =   12192
   Icon            =   "frmRptMailLabels.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   12192
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboSvcAddr 
      Height          =   348
      Left            =   5256
      TabIndex        =   6
      Top             =   5388
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
      ColDesigner     =   "frmRptMailLabels.frx":08CA
   End
   Begin LpLib.fpCombo fpcboRevenues 
      Height          =   348
      Left            =   5256
      TabIndex        =   4
      Top             =   4332
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
      ColDesigner     =   "frmRptMailLabels.frx":0C68
   End
   Begin LpLib.fpCombo fpcboLabelType 
      Height          =   348
      Left            =   5256
      TabIndex        =   7
      Top             =   5940
      Width           =   4548
      _Version        =   196608
      _ExtentX        =   8022
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
      ColDesigner     =   "frmRptMailLabels.frx":0FFB
   End
   Begin LpLib.fpCombo fpcboPrintOrder 
      Height          =   348
      Left            =   5256
      TabIndex        =   2
      Top             =   3300
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
      ColDesigner     =   "frmRptMailLabels.frx":1399
   End
   Begin LpLib.fpCombo fpcboIncludeInactive 
      Height          =   348
      Left            =   5256
      TabIndex        =   5
      Top             =   4848
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
      ColDesigner     =   "frmRptMailLabels.frx":172C
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
      Left            =   5244
      TabIndex        =   1
      Top             =   2796
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
      Left            =   5244
      TabIndex        =   0
      Top             =   2292
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
   Begin EditLib.fpText fptxtCustType 
      Height          =   348
      Left            =   5256
      TabIndex        =   3
      Top             =   3816
      Width           =   1188
      _Version        =   196608
      _ExtentX        =   2096
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
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Print Service Address:"
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
      Left            =   2394
      TabIndex        =   19
      Top             =   5454
      Width           =   2628
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Include Inactive: "
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
      Left            =   3090
      TabIndex        =   18
      Top             =   4950
      Width           =   2004
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
      Left            =   3300
      TabIndex        =   17
      Top             =   3366
      Width           =   1716
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
      Index           =   0
      Left            =   2940
      TabIndex        =   16
      Top             =   3876
      Width           =   2076
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ending Route:"
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
      Left            =   3162
      TabIndex        =   15
      Top             =   2862
      Width           =   1860
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Starting Route:"
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
      Index           =   8
      Left            =   3090
      TabIndex        =   14
      Top             =   2346
      Width           =   1932
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   4764
      Left            =   2148
      Top             =   1944
      Width           =   7908
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Select Label Type: "
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
      Left            =   2706
      TabIndex        =   13
      Top             =   5982
      Width           =   2388
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Revenue Source:"
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
      Left            =   3018
      TabIndex        =   12
      Top             =   4410
      Width           =   2004
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3192
      Top             =   912
      Width           =   5772
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Print Mailing Labels"
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
      Top             =   1152
      Width           =   5004
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   972
      Left            =   3192
      Top             =   792
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
Attribute VB_Name = "frmRptMailLabels"
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
  Unload frmRptMailLabels
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
    fpcboPrintOrder.SetFocus
  End If
End Sub

Private Sub fptxtRoute2_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub
Private Sub fpcboPrintOrder_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboPrintOrder.ListDown = True
  End If
  If fpcboPrintOrder.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fptxtCustType.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fptxtRoute2.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub
Private Sub fptxtCustType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpcboRevenues.SetFocus
  End If
End Sub

Private Sub fpcboRevenues_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboRevenues.ListDown = True
  End If
  If fpcboRevenues.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpcboIncludeInactive.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fptxtCustType.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub


Private Sub fpcboIncludeInactive_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboIncludeInactive.ListDown = True
  End If
  If fpcboIncludeInactive.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpcboSvcAddr.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpcboRevenues.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub
Private Sub fpcboSvcAddr_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboSvcAddr.ListDown = True
  End If
  If fpcboSvcAddr.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpcboLabelType.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpcboIncludeInactive.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub

Private Sub fpcboLabelType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboLabelType.ListDown = True
  End If
  If fpcboLabelType.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      cmdPrint.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpcboSvcAddr.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub

Private Function ValidRoutes()
  If fptxtRoute1 <> "" And fptxtRoute2 <> "" Then
    If fptxtRoute1 <> 0 And fptxtRoute2 <> 0 Then
    If fptxtRoute1 > fptxtRoute2 Then
      MsgBox "Invalid Route Selection, The Beginning Route Should Be Less or Equal to Ending Route.", vbOKOnly, "Invalid Selection"
      ValidRoutes = False
    Else
      ValidRoutes = True
      BegRoute = QPTrim(fptxtRoute1)
      EndRoute = QPTrim(fptxtRoute2)
    End If
  Else
    MsgBox "Route Selections Must Be Greater than 0.", vbOKOnly, "Invalid Selection"
  End If
  
  Else
    MsgBox "Route Selections May Not Be Left Blank.", vbOKOnly, "Invalid Selection"
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
    If fpcboLabelType.ListIndex = 0 Then
        MailingLabel
      ElseIf fpcboLabelType.ListIndex = 1 Then
        MailingLabel
      ElseIf fpcboLabelType.ListIndex = 2 Then
        MailingLabel
      Else
        MailingLabel2
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
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TownName$
  fptxtRoute1 = "01"
  fptxtRoute2 = "99"
  fpcboPrintOrder.AddItem "Customer Name Order"
  fpcboPrintOrder.AddItem "Account Number Order"
  fpcboPrintOrder.AddItem "Location Number Order"
  fpcboPrintOrder.AddItem "Service Address Order"
  fpcboPrintOrder.AddItem "Sequence Number Order"
  fpcboPrintOrder.AddItem "Zip Code Number Order"
  fpcboPrintOrder.AddItem "Postal Carrier Order"
  fpcboPrintOrder.ListIndex = 0
  fptxtCustType = ""
  FillRevList fpcboRevenues
  fpcboRevenues.ListIndex = 0
  fpcboIncludeInactive.AddItem "Yes"
  fpcboIncludeInactive.AddItem "No"
  fpcboIncludeInactive.ListIndex = 0
  fpcboSvcAddr.AddItem "Yes"
  fpcboSvcAddr.AddItem "No"
  fpcboSvcAddr.ListIndex = 1
  fpcboLabelType.InsertRow = "1) 1 X 3 1/2 (1-Label Wide)Text"
  fpcboLabelType.InsertRow = "2) 1 X 3 1/2 (3-Labels Wide)Text"
  fpcboLabelType.InsertRow = "3) 1 X 3 1/2 (4-Labels Wide)Text"
  fpcboLabelType.InsertRow = "4) 1 X 2 5/8 (Full Sheet 3 Wide)Graphics"
  fpcboLabelType.ListIndex = 0
End Sub

Private Sub Form_Resize()
'  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
'  End If
End Sub
Private Sub MailingLabel()
  Dim IdxRecLen As Integer, IdxFileSize As Long, UsingZip As Boolean
  Dim IdxNumOfRecs As Long, NumOfRecs As Long, Handle As Integer
  Dim Status As String, UsingSeq As Boolean, UsingPos As Boolean
  Dim UBCust As Integer, UBRpt As Integer, UsingBook As Boolean
  Dim UBSetupreclen  As Integer, ReportFile As String, cnt As Long
  Dim IndexName As String, UBCustRecLen As Integer, Order As String
  Dim BeechFlag As Boolean, UsingAcct As Boolean, UsingAddr As Boolean
  Dim UsingName As Boolean, CustomerType As String, xbar As Boolean
  Dim FirstBook As Integer, LastBook As Integer, IncInactive As Boolean
  Dim RType As Integer, AcctNumber As Long, CustBook As Integer
  Dim Rev As String, FlatCnt As Integer, CustPCnt As Long, Zip As String
  Dim Zip1 As String, ZipCnt As Integer, zz As Integer, ADDR2 As String
  Dim CustName As String, ADDR1 As String, DraftOnlyFlag As Boolean
  Dim CityStaZip As String, LType As Integer, DidCnt As Integer
  Dim LabelCnt As Integer, PCnt As Integer, ToPrint(1 To 5) As String * 132
  Dim Align As String, MaskLabel As String, Filler As String
'  ReDim UBSetUpRec(1) As UBSetupRecType
'  LoadUBSetUpFile UBSetUpRec(), UBSetuplen%
 Dim PrnArray() As UBPostalIndexType
'  If InStr(UBSetUpRec(1).UTILNAME, "BEECH MOUNTAIN") > 0 Then
'    BeechFlag = True
'  End If
'instead of beechmtn will use screen option of service address
'but use same flag
  If fpcboSvcAddr.ListIndex = 0 Then
    BeechFlag = True
  End If
'  DidCnt = 0
'  For RevCnt = 1 To 15
'    Rev$ = QPTrim$(UBSetUpRec(1).Revenues(RevCnt).REVNAME)
'    If Len(Rev$) > 0 Then
'      DidCnt = DidCnt + 1
'      Choice$(DidCnt, 1) = QPTrim$(Str$(DidCnt)) + ") " + Rev$
'    Else
'      Exit For
'    End If
'  Next
ToPrint(1) = ""
ToPrint(2) = ""
ToPrint(3) = ""
ToPrint(4) = ""
ToPrint(5) = ""
  ReDim OSet(1 To 4) As Integer

  OSet(1) = 1
  OSet(2) = 37
  OSet(3) = 74
  OSet(4) = 110
  UsingBook = False
  UsingAcct = False
  UsingName = False
  UsingAddr = False
  'AbortFlag = False
  PageNo = 0
  Filler$ = Space(23)

      Order$ = Left$(fpcboPrintOrder.Text, 3)
      Select Case Order$
      Case "Cus"
        IndexName$ = NameIndexFile
        UsingName = True
      Case "Acc"
        IndexName$ = ""
        UsingAcct = True
      Case "Loc"
        IndexName$ = BookIndexFile
        UsingBook = True
      Case "Ser"
        IndexName$ = TempIndexName
        UsingAddr = True
      Case "Seq"
        MakeSequenceIndex "Sequence Numbers", Me
        IndexName$ = TempIndexName
        UsingSeq = True
      Case "Zip"
        MakeMowZipCodeIndex "ZipCode"
        IndexName$ = TempIndexName
        UsingZip = True
      Case "Pos"
        MakePostalIndex "Postal Route"
        IndexName$ = TempIndexName
        UsingPos = True
      Case Else
      End Select



'***************

  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))

  ReDim UBSetUp(1) As UBSetupRecType
  UBSetupreclen = Len(UBSetUp(1))

  CustomerType$ = QPTrim$(fptxtCustType)

  If CustomerType$ = "DRFT" Then
    DraftOnlyFlag = True
  End If

  FirstBook = Val(BegRoute)
  LastBook = Val(EndRoute)
  If fpcboIncludeInactive.ListIndex = 0 Then
    IncInactive = True
  End If
  LType = fpcboLabelType.ListIndex + 1
  RType = fpcboRevenues.ListIndex
'$^%$^$^$^&%$^%$  Check the revenue value !!!!!
  If UsingAddr Then
    SortServiceAddrs frmRptMailLabels
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

  ElseIf UsingName Or UsingBook Or UsingSeq Or UsingZip Or UsingPos Then
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
  ReportFile$ = UBPath$ + "UBLABEL.RPT"
  UBRpt = FreeFile
  Open ReportFile$ For Output As UBRpt

'  BlockClear
'  ShowProcessingScrn "Mailing Labels"

  For cnt = 1 To NumOfRecs
    If UsingName Or UsingBook Or UsingSeq Or UsingZip Or UsingPos Then
      AcctNumber& = IdxBuff(cnt).RecNum
    Else
      AcctNumber& = cnt
    End If
    Get UBCust, AcctNumber&, UBCustRec(1)

    If UBCustRec(1).DelFlag = True Then
      GoTo NextLabel
    End If

    CustBook = Val(UBCustRec(1).Book)

    If CustBook < FirstBook Or CustBook > LastBook Then
      GoTo NextLabel
    End If

    If (UBCustRec(1).Status = "I") And (IncInactive = 0) Then
      GoTo NextLabel
    End If
 
    If DraftOnlyFlag Then
      If UBCustRec(1).USEDRAFT <> "Y" Then
        GoTo NextLabel
      End If
    ElseIf Len(CustomerType$) > 0 Then
      If UCase$(QPTrim$(UBCustRec(1).CUSTTYPE)) <> UCase$(CustomerType$) Then
        GoTo NextLabel
      End If
    End If

    If QPTrim$(UBCustRec(1).CustName) = "VACANT" Then
      GoTo NextLabel
    End If

    If RType > 0 Then
      Rev$ = QPTrim$(UBCustRec(1).Serv(RType).RATECODE)
      If Len(Rev$) > 0 Then
        GoTo GoodCust
      End If
      For FlatCnt = 1 To 4
        If UBCustRec(1).FlatRates(FlatCnt).REVSRC = RType Then
          GoTo GoodCust
        End If
      Next
      GoTo NextLabel
    
    End If
GoodCust:
    CustPCnt = CustPCnt + 1
    
    Zip$ = Left$(UBCustRec(1).ZIPCODE, 5) + "-" + Mid$(UBCustRec(1).ZIPCODE, 6)
    Zip$ = QPTrim$(Zip$)

    Zip1$ = Left$(UBCustRec(1).ZIPCODE, 5)
    If CustPCnt = 1 Then
      ZipCnt = ZipCnt + 1
      ReDim Preserve PrnArray(1 To ZipCnt) As UBPostalIndexType
      PrnArray(ZipCnt).ZIPCODE = Zip1$
      PrnArray(ZipCnt).RecNum = 1
    Else
      For zz = 1 To ZipCnt
        If InStr(PrnArray(zz).ZIPCODE, Zip1$) > 0 Then
          PrnArray(zz).RecNum = PrnArray(zz).RecNum + 1
          GoTo GotZip
        End If
      Next
      ZipCnt = ZipCnt + 1
      ReDim Preserve PrnArray(1 To ZipCnt) As UBPostalIndexType
      PrnArray(ZipCnt).ZIPCODE = Zip1$
      PrnArray(ZipCnt).RecNum = 1
    End If

GotZip:
   
    CustName$ = QPTrim$(UBCustRec(1).CustName)
    If Len(CustName$) < 1 Then GoTo NextLabel
    If BeechFlag Then
      ADDR1$ = QPTrim$(UBCustRec(1).SERVADDR)
      ADDR2$ = ""
    Else
      ADDR1$ = QPTrim$(UBCustRec(1).ADDR1)
      ADDR2$ = QPTrim$(UBCustRec(1).ADDR2)
    End If

    CityStaZip$ = Left$(QPTrim$(UBCustRec(1).CITY), 18) + ", " + UBCustRec(1).STATE + " " + Zip$

    Select Case LType
    Case 1
      Print #UBRpt, Filler$ '"Cust #" + STR$(AcctNumber&)
      Print #UBRpt, Left$(CustName$, 23)
      Print #UBRpt, Left$(ADDR1$, 23)
      If Len(ADDR2$) > 0 Then
        Print #UBRpt, Left$(ADDR2$, 23)
        Print #UBRpt, CityStaZip$
      Else
        Print #UBRpt, CityStaZip$
        Print #UBRpt, Filler$
      End If
      Print #UBRpt, Filler$
      DidCnt = DidCnt + 1
    Case 2
      LabelCnt = LabelCnt + 1
      '
      Mid$(ToPrint(1), OSet(LabelCnt)) = Filler$ '"   Cust #" + Str$(AcctNumber&)
      Mid$(ToPrint(2), OSet(LabelCnt)) = Left$(CustName$, 23)
      Mid$(ToPrint(3), OSet(LabelCnt)) = Left$(ADDR1$, 23)
      If Len(ADDR2$) > 0 Then
        Mid$(ToPrint(4), OSet(LabelCnt)) = Left$(ADDR2$, 23)
        Mid$(ToPrint(5), OSet(LabelCnt)) = CityStaZip$
      Else
        Mid$(ToPrint(4), OSet(LabelCnt)) = CityStaZip$
      End If
      If LabelCnt = 3 Then
        For PCnt = 1 To 5
          'LPRINT QPTrim$(ToPrint(PCnt))
          Print #UBRpt, ToPrint(PCnt)
          ToPrint(PCnt) = ""
        Next
        Print #UBRpt, Filler$
        LabelCnt = 0
      End If

    Case 3
      LabelCnt = LabelCnt + 1
      Mid$(ToPrint(1), OSet(LabelCnt)) = Filler$ '"Cust #" + STR$(AcctNumber&)
      Mid$(ToPrint(2), OSet(LabelCnt)) = Left$(CustName$, 23)
      Mid$(ToPrint(3), OSet(LabelCnt)) = Left$(ADDR1$, 23)
      If Len(ADDR2$) > 0 Then
        Mid$(ToPrint(4), OSet(LabelCnt)) = Left$(ADDR2$, 23)
        Mid$(ToPrint(5), OSet(LabelCnt)) = CityStaZip$
      Else
        Mid$(ToPrint(4), OSet(LabelCnt)) = CityStaZip$
      End If
      If LabelCnt = 4 Then
        For PCnt = 1 To 5
          'LPRINT QPTrim$(ToPrint(PCnt))
          Print #UBRpt, ToPrint(PCnt)
          ToPrint(PCnt) = ""
        Next
        Print #UBRpt,
        LabelCnt = 0
      End If

    End Select


NextLabel:
 '   ShowPctComp cnt, NumOfRecs
    'IF CustPCnt > 60 THEN EXIT FOR
  Next
  If LType = 2 Or LType = 3 Then
    If LabelCnt > 0 Then
      For PCnt = 1 To 5
        Print #UBRpt, QPTrim$(ToPrint(PCnt))
      Next
      Print #UBRpt, Filler$
    End If
  End If
  PCnt = 0

 ' SortT PrnArray(1), ZipCnt, 0, 16, 0, 14
'  For cnt = 1 To ZipCnt
'    'PCnt = PCnt + 1
'    Print #UBRpt, PrnArray(cnt).ZIPCODE; Tab(40); PrnArray(cnt).RecNum
'    'IF PCnt = 5 THEN
'    '  PRINT #UBRpt,
'    '  PCnt = 0
'    'END IF
'  Next
  LSet ToPrint(1) = "Total:" + Str$(CustPCnt)
  Print #UBRpt, ToPrint(1)
  Print #UBRpt, Chr$(12);
  Close

  Erase IdxBuff, UBCustRec, ToPrint
  Erase UBSetUp, OSet 'frm, Form$, Fld,

  GoSub DoAlignLabelMask

  
  ViewPrint ReportFile$, "Mailing Labels", xbar, , True, MaskLabel$
  
'  If Not AbortFlag Then
'    PrintRptFile , , 1, RetCode, EntryPoint
'  End If

  'KillFile "UBLABEL.RPT"

ExitMailLabListing:

  Exit Sub

DoAlignLabelMask:
  ReDim OSet(1 To 4) As Integer

  OSet(1) = 1
  OSet(2) = 37
  OSet(3) = 74
  OSet(4) = 110
  xbar = False
'  ReDim TempScrn(0)
'  SaveScrn TempScrn()
  Align$ = String$(34, "X")
  UBRpt = FreeFile
  MaskLabel$ = UBPath$ + "UBLblA.RPT"
  Open UBPath$ + "UBLblA.RPT" For Output As UBRpt
  Select Case LType
  Case 1
    For cnt = 1 To 5
      Print #UBRpt, Align$
    Next
    Print #UBRpt,
    xbar = False
  Case 2
    For cnt = 1 To 5
      Print #UBRpt, Align$; Tab(OSet(2)); Align$; Tab(OSet(3)); Align$
    Next
    Print #UBRpt,
    xbar = False
  Case 3
    For cnt = 1 To 5
      Print #UBRpt, Align$; Tab(OSet(2)); Align$; Tab(OSet(3)); Align$; Tab(OSet(4)); Align$
    Next
    Print #UBRpt,
    xbar = True
  End Select
  Close UBRpt

Return

End Sub
Private Sub MailingLabel2()
  Dim IdxRecLen As Integer, IdxFileSize As Long, UsingZip As Boolean
  Dim IdxNumOfRecs As Long, NumOfRecs As Long, Handle As Integer
  Dim Status As String, UsingSeq As Boolean, UsingPos As Boolean
  Dim UBCust As Integer, UBRpt As Integer, UsingBook As Boolean
  Dim UBSetupreclen  As Integer, ReportFile As String, cnt As Long
  Dim IndexName As String, UBCustRecLen As Integer, Order As String
  Dim BeechFlag As Boolean, UsingAcct As Boolean, UsingAddr As Boolean
  Dim UsingName As Boolean, CustomerType As String, xbar As Boolean
  Dim FirstBook As Integer, LastBook As Integer, IncInactive As Boolean
  Dim RType As Integer, AcctNumber As Long, CustBook As Integer
  Dim Rev As String, FlatCnt As Integer, CustPCnt As Long, Zip As String
  Dim Zip1 As String, ZipCnt As Integer, zz As Integer, ADDR2 As String
  Dim CustName As String, ADDR1 As String, DraftOnlyFlag As Boolean
  Dim CityStaZip As String, LType As Integer, DidCnt As Integer
  Dim LabelCnt As Integer, ToPrint1 As String, ToPrint2 As String
  Dim ToPrint As String, ToPrint4 As String, ToPrint3 As String
  
  If fpcboSvcAddr.ListIndex = 0 Then
    BeechFlag = True
  End If
  Dim PrnArray() As UBPostalIndexType
 
  ToPrint = ""
  ToPrint1 = ""
  ToPrint2 = ""
  ToPrint3 = ""
  ToPrint4 = ""
  'ToPrint(5) = ""
  UsingBook = False
  UsingAcct = False
  UsingName = False
  UsingAddr = False

      Order$ = Left$(fpcboPrintOrder.Text, 3)
      Select Case Order$
      Case "Cus"
        IndexName$ = NameIndexFile
        UsingName = True
      Case "Acc"
        IndexName$ = ""
        UsingAcct = True
      Case "Loc"
        IndexName$ = BookIndexFile
        UsingBook = True
      Case "Ser"
        IndexName$ = TempIndexName
        UsingAddr = True
      Case "Seq"
        MakeSequenceIndex "Sequence Numbers", Me
        IndexName$ = TempIndexName
        UsingSeq = True
      Case "Zip"
        MakeMowZipCodeIndex "ZipCode"
        IndexName$ = TempIndexName
        UsingZip = True
      Case "Pos"
        MakePostalIndex "Postal Route"
        IndexName$ = TempIndexName
        UsingPos = True
      Case Else
      End Select
'***************
  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))

  ReDim UBSetUp(1) As UBSetupRecType
  UBSetupreclen = Len(UBSetUp(1))

  CustomerType$ = QPTrim$(fptxtCustType)

  If CustomerType$ = "DRFT" Then
    DraftOnlyFlag = True
  End If

  FirstBook = Val(fptxtRoute1)
  LastBook = Val(fptxtRoute2)
  If fpcboIncludeInactive.ListIndex = 0 Then
    IncInactive = True
  End If
  LType = fpcboLabelType.ListIndex + 1
  RType = fpcboRevenues.ListIndex
'$^%$^$^$^&%$^%$  Check the revenue value !!!!!
  If UsingAddr Then
    SortServiceAddrs frmRptMailLabels
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

  ElseIf UsingName Or UsingBook Or UsingSeq Or UsingZip Or UsingPos Then
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
  ReportFile$ = UBPath$ + "UBLABEL.RPT"
  UBRpt = FreeFile
  Open ReportFile$ For Output As UBRpt
  For cnt = 1 To NumOfRecs
    If UsingName Or UsingBook Or UsingSeq Or UsingZip Or UsingPos Then
      AcctNumber& = IdxBuff(cnt).RecNum
    Else
      AcctNumber& = cnt
    End If
    Get UBCust, AcctNumber&, UBCustRec(1)

    If UBCustRec(1).DelFlag = True Then
      GoTo NextLabel
    End If

    CustBook = Val(UBCustRec(1).Book)

    If CustBook < FirstBook Or CustBook > LastBook Then
      GoTo NextLabel
    End If

    If (UBCustRec(1).Status = "I") And (IncInactive = 0) Then
      GoTo NextLabel
    End If

    If DraftOnlyFlag Then
      If UBCustRec(1).USEDRAFT <> "Y" Then
        GoTo NextLabel
      End If
    ElseIf Len(CustomerType$) > 0 Then
      If UCase$(QPTrim$(UBCustRec(1).CUSTTYPE)) <> UCase$(CustomerType$) Then
        GoTo NextLabel
      End If
    End If
    If QPTrim$(UBCustRec(1).CustName) = "VACANT" Then
      GoTo NextLabel
    End If

    If RType > 0 Then
      Rev$ = QPTrim$(UBCustRec(1).Serv(RType).RATECODE)
      If Len(Rev$) > 0 Then
        GoTo GoodCust
      End If
      For FlatCnt = 1 To 4
        If UBCustRec(1).FlatRates(FlatCnt).REVSRC = RType Then
          GoTo GoodCust
        End If
      Next
      GoTo NextLabel
    
    End If
GoodCust:
    CustPCnt = CustPCnt + 1
    
    Zip$ = Left$(UBCustRec(1).ZIPCODE, 5) + "-" + Mid$(UBCustRec(1).ZIPCODE, 6)
    Zip$ = QPTrim$(Zip$)

    Zip1$ = Left$(UBCustRec(1).ZIPCODE, 5)
    If CustPCnt = 1 Then
      ZipCnt = ZipCnt + 1
      ReDim Preserve PrnArray(1 To ZipCnt) As UBPostalIndexType
      PrnArray(ZipCnt).ZIPCODE = Zip1$
      PrnArray(ZipCnt).RecNum = 1
    Else
      For zz = 1 To ZipCnt
        If InStr(PrnArray(zz).ZIPCODE, Zip1$) > 0 Then
          PrnArray(zz).RecNum = PrnArray(zz).RecNum + 1
          GoTo GotZip
        End If
      Next
      ZipCnt = ZipCnt + 1
      ReDim Preserve PrnArray(1 To ZipCnt) As UBPostalIndexType
      PrnArray(ZipCnt).ZIPCODE = Zip1$
      PrnArray(ZipCnt).RecNum = 1
    End If

GotZip:
   
    CustName$ = QPTrim$(UBCustRec(1).CustName)
    If Len(CustName$) < 1 Then GoTo NextLabel
    If BeechFlag Then
      ADDR1$ = QPTrim$(UBCustRec(1).SERVADDR)
      ADDR2$ = ""
    Else
      ADDR1$ = QPTrim$(UBCustRec(1).ADDR1)
      ADDR2$ = QPTrim$(UBCustRec(1).ADDR2)
    End If

    CityStaZip$ = Left$(QPTrim$(UBCustRec(1).CITY), 18) + ", " + UBCustRec(1).STATE + " " + Zip$

      LabelCnt = LabelCnt + 1
      '
      'Mid$(ToPrint(1), OSet(LabelCnt)) = Filler$ '"   Cust #" + Str$(AcctNumber&)
      ToPrint1 = ToPrint1 + Left$(CustName$, 23) + "~"
      ToPrint2 = ToPrint2 + Left$(ADDR1$, 23) + "~"
      If Len(ADDR2$) > 0 Then
        ToPrint3 = ToPrint3 + Left$(ADDR2$, 23) + "~"
        ToPrint4 = ToPrint4 + CityStaZip$ + "~"
      Else
       ToPrint3 = ToPrint3 + CityStaZip$ + "~"
       ToPrint4 = ToPrint4 + " ~"
      End If
      If LabelCnt = 3 Then
       ' For cnt = 1 To 4
          'LPRINT QPTrim$(ToPrint(PCnt))
          Print #UBRpt, ToPrint1 + ToPrint2 + ToPrint3 + ToPrint4
          ToPrint1 = ""
          ToPrint2 = ""
          ToPrint3 = ""
          ToPrint4 = ""
'        Next
'        Print #UBRpt, Filler$
        LabelCnt = 0
      Else
        
      End If



NextLabel:
 '   ShowPctComp cnt, NumOfRecs
    'IF CustPCnt > 60 THEN EXIT FOR
  Next
'  If LType = 2 Or LType = 3 Then
'    If LabelCnt > 0 Then
'      For PCnt = 1 To 5
'        Print #UBRpt, QPTrim$(ToPrint(PCnt))
'      Next
'      Print #UBRpt, Filler$
'    End If
'  End If
'  PCnt = 0

 ' SortT PrnArray(1), ZipCnt, 0, 16, 0, 14
'  For cnt = 1 To ZipCnt
'    'PCnt = PCnt + 1
'    Print #UBRpt, PrnArray(cnt).ZIPCODE; Tab(40); PrnArray(cnt).RecNum
'    'IF PCnt = 5 THEN
'    '  PRINT #UBRpt,
'    '  PCnt = 0
'    'END IF
'  Next
'  LSet ToPrint(1) = "Total:" + Str$(CustPCnt)
'  Print #UBRpt, ToPrint(1)
'  Print #UBRpt, Chr$(12);
  Close

  Erase IdxBuff, UBCustRec ', 'ToPrint
  Erase UBSetUp ', OSet 'frm, Form$, Fld,

'  GoSub DoAlignLabelMask
  Load frmLoadingRpt
  ARptMailLabels.GetName ReportFile$
  ARptMailLabels.startrpt

  
'  ViewPrint , "Mailing Labels", xbar, , True, MaskLabel$

ExitMailLabListing:
  Exit Sub
End Sub

