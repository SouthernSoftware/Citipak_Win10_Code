VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmEmpORBITEdit 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ORBIT Edit"
   ClientHeight    =   8835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11640
   Icon            =   "frmEmpORBITEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8640
   ScaleMode       =   0  'User
   ScaleWidth      =   28980.62
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpList fpListEmp 
      Height          =   4560
      Left            =   120
      TabIndex        =   60
      ToolTipText     =   "The first column is the check date."
      Top             =   2160
      Width           =   2295
      _Version        =   196608
      _ExtentX        =   4048
      _ExtentY        =   8043
      TextAlias       =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
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
      Columns         =   1
      Sorted          =   0
      LineWidth       =   1
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   -1
      ColumnWidthScale=   3
      RowHeight       =   -1
      MultiSelect     =   0
      WrapList        =   0   'False
      WrapWidth       =   200
      SelMax          =   -1
      AutoSearch      =   1
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
      BorderStyle     =   1
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
      ScrollBarH      =   0
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
      ColDesigner     =   "frmEmpORBITEdit.frx":08CA
   End
   Begin LpLib.fpCombo cmbContractPd 
      Height          =   360
      Left            =   9600
      TabIndex        =   19
      Top             =   3600
      Width           =   1770
      _Version        =   196608
      _ExtentX        =   3122
      _ExtentY        =   635
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
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
      AutoSearchFill  =   -1  'True
      AutoSearchFillDelay=   200
      EditMarginLeft  =   1
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   1
      EditAlignV      =   0
      ColDesigner     =   "frmEmpORBITEdit.frx":0BBA
   End
   Begin LpLib.fpCombo cmbPayType 
      Height          =   360
      Left            =   9600
      TabIndex        =   18
      Top             =   3240
      Width           =   1770
      _Version        =   196608
      _ExtentX        =   3122
      _ExtentY        =   635
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
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
      AutoSearchFill  =   -1  'True
      AutoSearchFillDelay=   200
      EditMarginLeft  =   1
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   1
      EditAlignV      =   0
      ColDesigner     =   "frmEmpORBITEdit.frx":0F15
   End
   Begin LpLib.fpCombo cmbAdjustment 
      Height          =   360
      Left            =   9600
      TabIndex        =   17
      Top             =   2880
      Width           =   1770
      _Version        =   196608
      _ExtentX        =   3122
      _ExtentY        =   635
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
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
      AutoSearchFill  =   -1  'True
      AutoSearchFillDelay=   200
      EditMarginLeft  =   1
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   1
      EditAlignV      =   0
      ColDesigner     =   "frmEmpORBITEdit.frx":1270
   End
   Begin LpLib.fpCombo cmbPlanCode 
      Height          =   360
      Left            =   9600
      TabIndex        =   15
      ToolTipText     =   "Select the Employee's Gender from the pick list."
      Top             =   2160
      Width           =   1770
      _Version        =   196608
      _ExtentX        =   3122
      _ExtentY        =   635
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
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
      AutoSearchFill  =   -1  'True
      AutoSearchFillDelay=   200
      EditMarginLeft  =   1
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   1
      EditAlignV      =   0
      ColDesigner     =   "frmEmpORBITEdit.frx":15CB
   End
   Begin LpLib.fpCombo cmbGender 
      Height          =   360
      Left            =   9600
      TabIndex        =   14
      ToolTipText     =   "Select the Employee's Gender from the pick list."
      Top             =   1800
      Width           =   1770
      _Version        =   196608
      _ExtentX        =   3122
      _ExtentY        =   635
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
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
      AutoSearchFill  =   -1  'True
      AutoSearchFillDelay=   200
      EditMarginLeft  =   1
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   1
      EditAlignV      =   0
      ColDesigner     =   "frmEmpORBITEdit.frx":1926
   End
   Begin LpLib.fpCombo cmbJobClassID 
      Height          =   360
      Left            =   9600
      TabIndex        =   16
      ToolTipText     =   "Select the Employee's Gender from the pick list."
      Top             =   2520
      Width           =   1770
      _Version        =   196608
      _ExtentX        =   3122
      _ExtentY        =   635
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
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
      AutoSearchFill  =   -1  'True
      AutoSearchFillDelay=   200
      EditMarginLeft  =   1
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   1
      EditAlignV      =   0
      ColDesigner     =   "frmEmpORBITEdit.frx":1C81
   End
   Begin LpLib.fpCombo cmbTermCode 
      Height          =   360
      Left            =   9600
      TabIndex        =   20
      Top             =   3960
      Width           =   1770
      _Version        =   196608
      _ExtentX        =   3122
      _ExtentY        =   635
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
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
      AutoSearchFill  =   -1  'True
      AutoSearchFillDelay=   200
      EditMarginLeft  =   1
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   1
      EditAlignV      =   0
      ColDesigner     =   "frmEmpORBITEdit.frx":1FDC
   End
   Begin LpLib.fpCombo cmbSuffix 
      Height          =   360
      Left            =   4080
      TabIndex        =   3
      ToolTipText     =   "Select the Employee's Gender from the pick list."
      Top             =   2895
      Width           =   1770
      _Version        =   196608
      _ExtentX        =   3122
      _ExtentY        =   635
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
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
      AutoSearchFill  =   -1  'True
      AutoSearchFillDelay=   200
      EditMarginLeft  =   1
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   1
      EditAlignV      =   0
      ColDesigner     =   "frmEmpORBITEdit.frx":2337
   End
   Begin LpLib.fpCombo cmbState 
      Height          =   360
      Left            =   4080
      TabIndex        =   8
      ToolTipText     =   "Required if 'Out Of Country Address' has not been populated."
      Top             =   4320
      Width           =   930
      _Version        =   196608
      _ExtentX        =   1640
      _ExtentY        =   635
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
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
      ColumnEdit      =   0
      ColumnBound     =   -1
      Style           =   2
      MaxDrop         =   8
      ListWidth       =   -1
      EditHeight      =   -1
      GrayAreaColor   =   -2147483633
      ListLeftOffset  =   0
      ComboGap        =   -2
      MaxEditLen      =   2
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
      AutoSearchFillDelay=   200
      EditMarginLeft  =   1
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   1
      EditAlignV      =   0
      ColDesigner     =   "frmEmpORBITEdit.frx":2692
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdBatch 
      Height          =   375
      Left            =   360
      TabIndex        =   76
      Top             =   7800
      Width           =   1815
      _Version        =   131072
      _ExtentX        =   3201
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   0   'False
      Value           =   0   'False
      GroupID         =   0
      GroupSelect     =   0
      DrawFocusRect   =   2
      DrawFocusRectCell=   -1
      GrayAreaPictureStyle=   0
      Static          =   0   'False
      BackStyle       =   1
      AutoSize        =   0
      AutoSizeOffsetTop=   0
      AutoSizeOffsetBottom=   0
      AutoSizeOffsetLeft=   0
      AutoSizeOffsetRight=   0
      DropShadowOffsetX=   3
      DropShadowOffsetY=   3
      DropShadowType  =   0
      DropShadowColor =   0
      Redraw          =   -1  'True
      ButtonDesigner  =   "frmEmpORBITEdit.frx":29ED
      Begin fpBtnAtlLibCtl.fpBtn fpBtn2 
         Height          =   375
         Left            =   0
         TabIndex        =   77
         Top             =   4200
         Width           =   1815
         _Version        =   131072
         _ExtentX        =   3201
         _ExtentY        =   661
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         GrayAreaColor   =   12632256
         BorderShowDefault=   -1  'True
         ButtonType      =   0
         NoPointerFocus  =   0   'False
         Value           =   0   'False
         GroupID         =   0
         GroupSelect     =   0
         DrawFocusRect   =   2
         DrawFocusRectCell=   -1
         GrayAreaPictureStyle=   0
         Static          =   0   'False
         BackStyle       =   1
         AutoSize        =   0
         AutoSizeOffsetTop=   0
         AutoSizeOffsetBottom=   0
         AutoSizeOffsetLeft=   0
         AutoSizeOffsetRight=   0
         DropShadowOffsetX=   3
         DropShadowOffsetY=   3
         DropShadowType  =   0
         DropShadowColor =   0
         Redraw          =   -1  'True
         ButtonDesigner  =   "frmEmpORBITEdit.frx":2BCE
      End
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdAmend 
      Height          =   375
      Left            =   4200
      TabIndex        =   71
      ToolTipText     =   $"frmEmpORBITEdit.frx":2DAF
      Top             =   8160
      Width           =   1575
      _Version        =   131072
      _ExtentX        =   2778
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   0   'False
      Value           =   0   'False
      GroupID         =   0
      GroupSelect     =   0
      DrawFocusRect   =   2
      DrawFocusRectCell=   -1
      GrayAreaPictureStyle=   0
      Static          =   0   'False
      BackStyle       =   1
      AutoSize        =   0
      AutoSizeOffsetTop=   0
      AutoSizeOffsetBottom=   0
      AutoSizeOffsetLeft=   0
      AutoSizeOffsetRight=   0
      DropShadowOffsetX=   3
      DropShadowOffsetY=   3
      DropShadowType  =   0
      DropShadowColor =   0
      Redraw          =   -1  'True
      ButtonDesigner  =   "frmEmpORBITEdit.frx":2E54
   End
   Begin EditLib.fpCurrency fpCurrRetSalary 
      Height          =   375
      Left            =   4320
      TabIndex        =   28
      Top             =   7440
      Width           =   1815
      _Version        =   196608
      _ExtentX        =   3201
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
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
   Begin fpBtnAtlLibCtl.fpBtn cmdByNum 
      Height          =   375
      Left            =   360
      TabIndex        =   62
      Top             =   7320
      Width           =   1815
      _Version        =   131072
      _ExtentX        =   3201
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   0   'False
      Value           =   0   'False
      GroupID         =   0
      GroupSelect     =   0
      DrawFocusRect   =   2
      DrawFocusRectCell=   -1
      GrayAreaPictureStyle=   0
      Static          =   0   'False
      BackStyle       =   1
      AutoSize        =   0
      AutoSizeOffsetTop=   0
      AutoSizeOffsetBottom=   0
      AutoSizeOffsetLeft=   0
      AutoSizeOffsetRight=   0
      DropShadowOffsetX=   3
      DropShadowOffsetY=   3
      DropShadowType  =   0
      DropShadowColor =   0
      Redraw          =   -1  'True
      ButtonDesigner  =   "frmEmpORBITEdit.frx":3033
      Begin fpBtnAtlLibCtl.fpBtn fpBtn1 
         Height          =   375
         Left            =   240
         TabIndex        =   75
         Top             =   480
         Width           =   1815
         _Version        =   131072
         _ExtentX        =   3201
         _ExtentY        =   661
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         GrayAreaColor   =   12632256
         BorderShowDefault=   -1  'True
         ButtonType      =   0
         NoPointerFocus  =   0   'False
         Value           =   0   'False
         GroupID         =   0
         GroupSelect     =   0
         DrawFocusRect   =   2
         DrawFocusRectCell=   -1
         GrayAreaPictureStyle=   0
         Static          =   0   'False
         BackStyle       =   1
         AutoSize        =   0
         AutoSizeOffsetTop=   0
         AutoSizeOffsetBottom=   0
         AutoSizeOffsetLeft=   0
         AutoSizeOffsetRight=   0
         DropShadowOffsetX=   3
         DropShadowOffsetY=   3
         DropShadowType  =   0
         DropShadowColor =   0
         Redraw          =   -1  'True
         ButtonDesigner  =   "frmEmpORBITEdit.frx":3215
      End
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdByName 
      Height          =   375
      Left            =   360
      TabIndex        =   61
      Top             =   6840
      Width           =   1815
      _Version        =   131072
      _ExtentX        =   3201
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   0   'False
      Value           =   0   'False
      GroupID         =   0
      GroupSelect     =   0
      DrawFocusRect   =   2
      DrawFocusRectCell=   -1
      GrayAreaPictureStyle=   0
      Static          =   0   'False
      BackStyle       =   1
      AutoSize        =   0
      AutoSizeOffsetTop=   0
      AutoSizeOffsetBottom=   0
      AutoSizeOffsetLeft=   0
      AutoSizeOffsetRight=   0
      DropShadowOffsetX=   3
      DropShadowOffsetY=   3
      DropShadowType  =   0
      DropShadowColor =   0
      Redraw          =   -1  'True
      ButtonDesigner  =   "frmEmpORBITEdit.frx":33F7
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   495
      Left            =   9840
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Press ESC to exit this screen."
      Top             =   8040
      Width           =   1455
      _Version        =   131072
      _ExtentX        =   2566
      _ExtentY        =   873
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   0   'False
      Value           =   0   'False
      GroupID         =   0
      GroupSelect     =   0
      DrawFocusRect   =   2
      DrawFocusRectCell=   -1
      GrayAreaPictureStyle=   0
      Static          =   0   'False
      BackStyle       =   1
      AutoSize        =   0
      AutoSizeOffsetTop=   0
      AutoSizeOffsetBottom=   0
      AutoSizeOffsetLeft=   0
      AutoSizeOffsetRight=   0
      DropShadowOffsetX=   3
      DropShadowOffsetY=   3
      DropShadowType  =   0
      DropShadowColor =   0
      Redraw          =   -1  'True
      ButtonDesigner  =   "frmEmpORBITEdit.frx":35D7
   End
   Begin EditLib.fpText fptxtFirstName 
      Height          =   360
      Left            =   4080
      TabIndex        =   1
      Top             =   2160
      Width           =   3300
      _Version        =   196608
      _ExtentX        =   5821
      _ExtentY        =   635
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
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
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   50
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
   Begin EditLib.fpText fptxtLastName 
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   1800
      Width           =   3300
      _Version        =   196608
      _ExtentX        =   5821
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
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
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   50
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
   Begin EditLib.fpMask fpMaskSoc 
      Height          =   375
      Left            =   4080
      TabIndex        =   10
      ToolTipText     =   "Enter the Employee's Social Security Number here."
      Top             =   4695
      Width           =   1770
      _Version        =   196608
      _ExtentX        =   3122
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
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
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483637
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   1
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
      AllowOverflow   =   0   'False
      BestFit         =   0   'False
      ClipMode        =   0
      DataFormatEx    =   0
      Mask            =   "###-##-####"
      PromptChar      =   "_"
      PromptInclude   =   0   'False
      RequireFill     =   0   'False
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
   Begin EditLib.fpText fptxtMemberID 
      Height          =   375
      Left            =   4080
      TabIndex        =   12
      ToolTipText     =   "Required for existing (not a new hire) employees."
      Top             =   5445
      Width           =   1770
      _Version        =   196608
      _ExtentX        =   3122
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
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
      Text            =   ""
      CharValidationText=   "0, 1, 2, 3, 4, 5, 6, 7, 8, 9,"
      MaxLength       =   9
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
   Begin EditLib.fpText fptxtInitial 
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      ToolTipText     =   "Enter a Complete or Partial Last Name here. Entering ""Mc"" will find ""McCoy, McDonald"". Press (F5) to do Look-Up, (ESC) to Cancel. "
      Top             =   2520
      Width           =   3300
      _Version        =   196608
      _ExtentX        =   5821
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
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
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   50
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
   Begin EditLib.fpText fptxtAdd1 
      Height          =   360
      Left            =   4080
      TabIndex        =   4
      Top             =   3240
      Width           =   3300
      _Version        =   196608
      _ExtentX        =   5821
      _ExtentY        =   635
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
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
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   50
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
   Begin EditLib.fpText fptxtAdd2 
      Height          =   360
      Left            =   4080
      TabIndex        =   5
      Top             =   3600
      Width           =   3300
      _Version        =   196608
      _ExtentX        =   5821
      _ExtentY        =   635
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
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
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   50
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
   Begin EditLib.fpText fptxtCity 
      Height          =   360
      Left            =   4080
      TabIndex        =   7
      Top             =   3960
      Width           =   3300
      _Version        =   196608
      _ExtentX        =   5821
      _ExtentY        =   635
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
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
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   50
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
   Begin EditLib.fpText fptxtOutOfCountry 
      Height          =   360
      Left            =   4080
      TabIndex        =   11
      Top             =   5085
      Width           =   3300
      _Version        =   196608
      _ExtentX        =   5821
      _ExtentY        =   635
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
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
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   50
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
   Begin EditLib.fpText fptxtDeptNum 
      Height          =   375
      Left            =   4080
      TabIndex        =   13
      Top             =   5805
      Width           =   1770
      _Version        =   196608
      _ExtentX        =   3122
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
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
      Text            =   ""
      CharValidationText=   "0, 1, 2, 3, 4, 5, 6, 7, 8, 9,"
      MaxLength       =   6
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
   Begin EditLib.fpDateTime fpMaskBDay 
      Height          =   375
      Left            =   9600
      TabIndex        =   23
      ToolTipText     =   "Enter the Employee's Date of Birth here."
      Top             =   5160
      Width           =   1770
      _Version        =   196608
      _ExtentX        =   3122
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
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
      AlignTextH      =   1
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
      Text            =   "10/01/2001"
      DateCalcMethod  =   0
      DateTimeFormat  =   5
      UserDefinedFormat=   "mm/dd/yyyy"
      DateMax         =   "20350101"
      DateMin         =   "19200101"
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
   Begin EditLib.fpDateTime fpdtHireDate 
      Height          =   375
      Left            =   9600
      TabIndex        =   22
      Top             =   4800
      Width           =   1770
      _Version        =   196608
      _ExtentX        =   3122
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
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
      AlignTextH      =   1
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
      OnFocusAlignH   =   2
      OnFocusAlignV   =   0
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   "10/01/2001"
      DateCalcMethod  =   0
      DateTimeFormat  =   5
      UserDefinedFormat=   "mm/dd/yyyy"
      DateMax         =   "20350101"
      DateMin         =   "19200101"
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
   Begin EditLib.fpDateTime fpdtEligibleDate 
      Height          =   375
      Left            =   9600
      TabIndex        =   24
      ToolTipText     =   "Required if a waiting period is used by your municipality."
      Top             =   5520
      Width           =   1770
      _Version        =   196608
      _ExtentX        =   3122
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
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
      AlignTextH      =   1
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
      DateMin         =   "19200101"
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
   Begin EditLib.fpDateTime fpdtTermDate 
      Height          =   375
      Left            =   9600
      TabIndex        =   21
      ToolTipText     =   "If terminated, enter the date this employee was terminated."
      Top             =   4440
      Width           =   1770
      _Version        =   196608
      _ExtentX        =   3122
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
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
      AlignTextH      =   1
      AlignTextV      =   0
      AllowNull       =   -1  'True
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
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
      DateMin         =   "19200101"
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
   Begin fpBtnAtlLibCtl.fpBtn cmdSave 
      Height          =   495
      Left            =   8040
      TabIndex        =   63
      TabStop         =   0   'False
      ToolTipText     =   "Press ESC to exit this screen."
      Top             =   8040
      Width           =   1455
      _Version        =   131072
      _ExtentX        =   2566
      _ExtentY        =   873
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   0   'False
      Value           =   0   'False
      GroupID         =   0
      GroupSelect     =   0
      DrawFocusRect   =   2
      DrawFocusRectCell=   -1
      GrayAreaPictureStyle=   0
      Static          =   0   'False
      BackStyle       =   1
      AutoSize        =   0
      AutoSizeOffsetTop=   0
      AutoSizeOffsetBottom=   0
      AutoSizeOffsetLeft=   0
      AutoSizeOffsetRight=   0
      DropShadowOffsetX=   3
      DropShadowOffsetY=   3
      DropShadowType  =   0
      DropShadowColor =   0
      Redraw          =   -1  'True
      ButtonDesigner  =   "frmEmpORBITEdit.frx":37B3
   End
   Begin EditLib.fpText fptxtSearch 
      Height          =   375
      Left            =   120
      TabIndex        =   64
      Top             =   1680
      Width           =   2250
      _Version        =   196608
      _ExtentX        =   3969
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
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
      AutoCase        =   1
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
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   6
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
   Begin EditLib.fpCurrency fpCurrOTPay 
      Height          =   375
      Left            =   10080
      TabIndex        =   33
      ToolTipText     =   "Use this field for calculations only. This field is not reported."
      Top             =   7440
      Width           =   1335
      _Version        =   196608
      _ExtentX        =   2355
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
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
   Begin EditLib.fpCurrency fpCurrRegPay 
      Height          =   375
      Left            =   8160
      TabIndex        =   29
      ToolTipText     =   "Use this field for calculations only. This field is not reported."
      Top             =   6600
      Width           =   1335
      _Version        =   196608
      _ExtentX        =   2355
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
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
   Begin EditLib.fpCurrency fpCurrEmpCont 
      Height          =   375
      Left            =   8160
      TabIndex        =   31
      Top             =   7440
      Width           =   1335
      _Version        =   196608
      _ExtentX        =   2355
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
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
   Begin EditLib.fpMask txtZip 
      Height          =   360
      Left            =   6075
      TabIndex        =   9
      ToolTipText     =   "Required if 'Out Of Country Address' has not been populated."
      Top             =   4320
      Width           =   1305
      _Version        =   196608
      _ExtentX        =   2302
      _ExtentY        =   635
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
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
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   1
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
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
      AllowOverflow   =   0   'False
      BestFit         =   0   'False
      ClipMode        =   0
      DataFormatEx    =   0
      Mask            =   "#####-####"
      PromptChar      =   "_"
      PromptInclude   =   0   'False
      RequireFill     =   0   'False
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      AutoTab         =   0   'False
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpDateTime fpdtPayPdBegDate 
      Height          =   375
      Left            =   4320
      TabIndex        =   26
      Top             =   6600
      Width           =   1815
      _Version        =   196608
      _ExtentX        =   3201
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
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
      AlignTextH      =   1
      AlignTextV      =   0
      AllowNull       =   -1  'True
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
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
      DateMin         =   "19200101"
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
   Begin EditLib.fpDateTime fpdtPayPdEndDate 
      Height          =   375
      Left            =   4320
      TabIndex        =   27
      Top             =   7022
      Width           =   1815
      _Version        =   196608
      _ExtentX        =   3201
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
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
      AlignTextH      =   1
      AlignTextV      =   0
      AllowNull       =   -1  'True
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
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
      DateMin         =   "19200101"
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
   Begin EditLib.fpCurrency fpCurrEmployerCont 
      Height          =   375
      Left            =   8160
      TabIndex        =   30
      Top             =   7042
      Width           =   1335
      _Version        =   196608
      _ExtentX        =   2355
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
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
   Begin EditLib.fpDoubleSingle fptxtVacHrsPaid 
      Height          =   375
      Left            =   10515
      TabIndex        =   25
      Top             =   6000
      Width           =   855
      _Version        =   196608
      _ExtentX        =   1508
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
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
      Text            =   "0.0"
      DecimalPlaces   =   1
      DecimalPoint    =   "."
      FixedPoint      =   -1  'True
      LeadZero        =   2
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
   Begin fpBtnAtlLibCtl.fpBtn cmdCheckNum 
      Height          =   375
      Left            =   360
      TabIndex        =   78
      Top             =   8280
      Width           =   1815
      _Version        =   131072
      _ExtentX        =   3201
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   0   'False
      Value           =   0   'False
      GroupID         =   0
      GroupSelect     =   0
      DrawFocusRect   =   2
      DrawFocusRectCell=   -1
      GrayAreaPictureStyle=   0
      Static          =   0   'False
      BackStyle       =   1
      AutoSize        =   0
      AutoSizeOffsetTop=   0
      AutoSizeOffsetBottom=   0
      AutoSizeOffsetLeft=   0
      AutoSizeOffsetRight=   0
      DropShadowOffsetX=   3
      DropShadowOffsetY=   3
      DropShadowType  =   0
      DropShadowColor =   0
      Redraw          =   -1  'True
      ButtonDesigner  =   "frmEmpORBITEdit.frx":398F
   End
   Begin EditLib.fpDoubleSingle fpCheckNum 
      Height          =   375
      Left            =   10080
      TabIndex        =   32
      Top             =   6840
      Width           =   1335
      _Version        =   196608
      _ExtentX        =   2355
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
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
      AlignTextH      =   1
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
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
      Text            =   "0"
      DecimalPlaces   =   0
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   2
      MaxValue        =   "9000000000"
      MinValue        =   "-9000000000"
      NegFormat       =   2
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
      IncDec          =   1
      BorderGrayAreaColor=   -2147483637
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
   Begin fpBtnAtlLibCtl.fpBtn cmdGetOld 
      Height          =   375
      Left            =   6000
      TabIndex        =   80
      Top             =   8160
      Width           =   1695
      _Version        =   131072
      _ExtentX        =   2990
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   0   'False
      Value           =   0   'False
      GroupID         =   0
      GroupSelect     =   0
      DrawFocusRect   =   2
      DrawFocusRectCell=   -1
      GrayAreaPictureStyle=   0
      Static          =   0   'False
      BackStyle       =   1
      AutoSize        =   0
      AutoSizeOffsetTop=   0
      AutoSizeOffsetBottom=   0
      AutoSizeOffsetLeft=   0
      AutoSizeOffsetRight=   0
      DropShadowOffsetX=   3
      DropShadowOffsetY=   3
      DropShadowType  =   0
      DropShadowColor =   0
      Redraw          =   -1  'True
      ButtonDesigner  =   "frmEmpORBITEdit.frx":3B72
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdDelete 
      Height          =   375
      Left            =   2640
      TabIndex        =   81
      Top             =   8160
      Width           =   1335
      _Version        =   131072
      _ExtentX        =   2355
      _ExtentY        =   661
      Enabled         =   0   'False
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   0   'False
      Value           =   0   'False
      GroupID         =   0
      GroupSelect     =   0
      DrawFocusRect   =   2
      DrawFocusRectCell=   -1
      GrayAreaPictureStyle=   0
      Static          =   0   'False
      BackStyle       =   1
      AutoSize        =   0
      AutoSizeOffsetTop=   0
      AutoSizeOffsetBottom=   0
      AutoSizeOffsetLeft=   0
      AutoSizeOffsetRight=   0
      DropShadowOffsetX=   3
      DropShadowOffsetY=   3
      DropShadowType  =   0
      DropShadowColor =   0
      Redraw          =   -1  'True
      ButtonDesigner  =   "frmEmpORBITEdit.frx":3D54
   End
   Begin VB.Label Label36 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Check Num:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   270
      Left            =   10080
      TabIndex        =   79
      Top             =   6600
      Width           =   1170
   End
   Begin VB.Label Label35 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Employer Contrib:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   270
      Left            =   6360
      TabIndex        =   74
      Top             =   7162
      Width           =   1650
   End
   Begin VB.Label Label34 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Pay Period End:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   270
      Left            =   2640
      TabIndex        =   73
      Top             =   7142
      Width           =   1530
   End
   Begin VB.Label Label33 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Pay Period Begin:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   270
      Left            =   2520
      TabIndex        =   72
      Top             =   6720
      Width           =   1650
   End
   Begin VB.Line Line3 
      BorderColor     =   &H0080FFFF&
      X1              =   19420
      X2              =   19420
      Y1              =   7745.195
      Y2              =   8449.304
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0080FFFF&
      X1              =   6274.154
      X2              =   28681.85
      Y1              =   6336.978
      Y2              =   6336.978
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080FFFF&
      X1              =   6274.154
      X2              =   28681.85
      Y1              =   7745.195
      Y2              =   7745.195
   End
   Begin VB.Label Label32 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Contrib:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   270
      Left            =   6360
      TabIndex        =   70
      Top             =   7560
      Width           =   1650
   End
   Begin VB.Label Label31 
      BackColor       =   &H008F8265&
      Caption         =   "*Required Fields:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2400
      TabIndex        =   69
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Label30 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Regular Pay:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   270
      Left            =   6600
      TabIndex        =   68
      Top             =   6720
      Width           =   1170
   End
   Begin VB.Label Label29 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "OT Pay:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   270
      Left            =   10200
      TabIndex        =   67
      Top             =   7200
      Width           =   930
   End
   Begin VB.Label Label26 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Gross Ret Pay:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   270
      Left            =   2880
      TabIndex        =   66
      Top             =   7560
      Width           =   1290
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Search:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   270
      Left            =   600
      TabIndex        =   65
      Top             =   1440
      Width           =   1290
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      Height          =   6975
      Left            =   2520
      Top             =   1680
      Width           =   9015
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H008F8265&
      Caption         =   "*Plan Code:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   8040
      TabIndex        =   59
      Top             =   2205
      Width           =   1335
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H008F8265&
      Caption         =   "*Gender:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   8040
      TabIndex        =   58
      Top             =   1845
      Width           =   1335
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H008F8265&
      Caption         =   "*Birth Date:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   8040
      TabIndex        =   57
      Top             =   5265
      Width           =   1335
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "*Job Class ID:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   270
      Left            =   7725
      TabIndex        =   56
      Top             =   2595
      Width           =   1650
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      BackColor       =   &H008F8265&
      Caption         =   "*Employment Date:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   7680
      TabIndex        =   55
      Top             =   4875
      Width           =   1695
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      BackColor       =   &H008F8265&
      Caption         =   "Eligibility Date:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   8040
      TabIndex        =   54
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Adjustment:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   270
      Left            =   7725
      TabIndex        =   53
      Top             =   2955
      Width           =   1650
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "*Pay Type:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   270
      Left            =   7725
      TabIndex        =   52
      Top             =   3315
      Width           =   1650
   End
   Begin VB.Label Label23 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Vacation Hours Paid (upon termination only):"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   270
      Left            =   6240
      TabIndex        =   51
      Top             =   6120
      Width           =   4170
   End
   Begin VB.Label Label24 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "*Employment Period:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   270
      Left            =   7365
      TabIndex        =   50
      Top             =   3675
      Width           =   2010
   End
   Begin VB.Label Label27 
      Alignment       =   1  'Right Justify
      BackColor       =   &H008F8265&
      Caption         =   "Termination Date:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   7680
      TabIndex        =   49
      Top             =   4515
      Width           =   1695
   End
   Begin VB.Label Label28 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Termination Code:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   270
      Left            =   7320
      TabIndex        =   48
      Top             =   4035
      Width           =   2010
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H008F8265&
      Caption         =   "*Last Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2520
      TabIndex        =   47
      Top             =   1875
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H008F8265&
      Caption         =   "*First Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   2520
      TabIndex        =   46
      Top             =   2205
      Width           =   1335
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "*SS Num:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2445
      TabIndex        =   45
      Top             =   4755
      Width           =   1410
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Member ID:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   270
      Left            =   2685
      TabIndex        =   44
      Top             =   5520
      Width           =   1170
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H008F8265&
      Caption         =   "Middle Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   2160
      TabIndex        =   43
      Top             =   2565
      Width           =   1695
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H008F8265&
      Caption         =   "Suffix:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2520
      TabIndex        =   42
      Top             =   2925
      Width           =   1335
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackColor       =   &H008F8265&
      Caption         =   "*Address 1:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2520
      TabIndex        =   41
      Top             =   3285
      Width           =   1335
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackColor       =   &H008F8265&
      Caption         =   "Address 2:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2520
      TabIndex        =   40
      Top             =   3645
      Width           =   1335
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackColor       =   &H008F8265&
      Caption         =   "*City:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2520
      TabIndex        =   39
      ToolTipText     =   "Required even if  'Out Of Country Address' is populated."
      Top             =   4005
      Width           =   1335
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackColor       =   &H008F8265&
      Caption         =   "*State:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   2520
      TabIndex        =   38
      Top             =   4380
      Width           =   1335
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Zip*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   5520
      TabIndex        =   37
      Top             =   4365
      Width           =   450
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackColor       =   &H008F8265&
      Caption         =   "Out Of Cntry Addrs:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   2520
      TabIndex        =   36
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Dept Num:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   270
      Left            =   2565
      TabIndex        =   35
      Top             =   5880
      Width           =   1290
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Retirement/ORBIT Edit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2933
      TabIndex        =   34
      Top             =   600
      Width           =   6015
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   855
      Index           =   1
      Left            =   1493
      Top             =   360
      Width           =   8655
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   1493
      Top             =   240
      Width           =   8655
   End
End
Attribute VB_Name = "frmEmpORBITEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Dim NameIdx() As String
Dim NameCnt As Integer
Dim NumIdx() As String
Dim NumCnt As Integer
Dim GIdx As Integer
Public GRecNum As Integer
Dim AddNew As Boolean
Dim GEmpNum As String
Dim GEmpRecNum As Long
Dim TerminateS As String
Dim SaveOld As Boolean
Dim SortType As String
Dim TempRetSalary As Double
Dim TempRegPay As Double
Dim TempEmpCont As Double
Dim TempOTPay As Double

Private Sub cmdAmend_Click()
  If GIdx = -1 Then
    MsgBox ("Please select an employee from the employee list on the left.")
    Exit Sub
  End If
  fpCurrRetSalary.Value = 0
  fpCurrRegPay.Value = 0
  fpCurrOTPay.Value = 0
  fpCurrEmpCont.Value = 0
  fpCurrEmployerCont.Value = 0
  AddNew = True
  GRecNum = 0
  GIdx = -1
  cmdDelete.Enabled = False
  frmMessage.Label1.Caption = "Please make certain that the combined totals for the amended and edited transactions equal the same total as the original record."
  frmMessage.Label1.Top = 700
  frmMessage.Show vbModal
  MainLog ("User warned to make sure that the amended and edited totals equal the same as the original totals for " & QPTrim$(fptxtFirstName.Text) & " " & QPTrim$(fptxtLastName.Text) & " with check number " & QPTrim$(fpCheckNum.Text) & " and pay period beginning on " & fpdtPayPdBegDate.Text & " and ending on " & fpdtPayPdEndDate.Text & ".")
End Sub

Private Sub cmdBatch_Click()
  Dim SaveIdx As String
  If GIdx > 0 Then
    SaveIdx = fpListEmp.ColText
  End If
  Call BatchSort
  fpListEmp.SearchText = SaveIdx
  fpListEmp.Action = 0
  GIdx = fpListEmp.SearchIndex

End Sub

Private Sub cmdByName_Click()
  Dim SaveIdx As String
  If GIdx > 0 Then
    SaveIdx = fpListEmp.ColText
  End If
  Call NameSort
  fpListEmp.SearchText = SaveIdx
  fpListEmp.Action = 0
  GIdx = fpListEmp.SearchIndex
End Sub

Private Sub cmdByNum_Click()
  Dim SaveIdx As String
  If GIdx > 0 Then
    SaveIdx = fpListEmp.ColText
  End If
  Call NumberSort
  fpListEmp.SearchText = SaveIdx
  fpListEmp.Action = 0
  GIdx = fpListEmp.SearchIndex
End Sub

Private Sub cmdClear_Click()
  fptxtFirstName.Text = ""
  fptxtLastName.Text = ""
  fptxtInitial.Text = ""
  cmbSuffix.Text = ""
  fptxtAdd1.Text = ""
  fptxtAdd2.Text = ""
  fptxtCity.Text = ""
  cmbState.Text = "NC"
  txtZip.Text = ""
  fpMaskSoc.Text = ""
  fptxtOutOfCountry.Text = ""
  fptxtMemberID.Text = ""
  cmbGender.Text = "Female"
  cmbPlanCode.Text = ""
  cmbJobClassID.Text = ""
  cmbAdjustment.Text = ""
  cmbPayType.Text = ""
  cmbContractPd.Text = ""
  cmbTermCode.Text = ""
  fpdtTermDate.Text = ""
  fpdtHireDate.Text = ""
  fpMaskBDay.Text = ""
  fpCurrRetSalary.Value = 0
  fpCurrRegPay.Value = 0
  fpCurrOTPay.Value = 0
  fpCurrEmpCont.Value = 0
  fpCurrEmployerCont.Value = 0
  AddNew = True
  SaveOld = False
  GRecNum = 0
  GEmpRecNum = 0
  GIdx = -1
  cmdDelete.Enabled = False
End Sub

Private Sub cmdCheckNum_Click()
  Dim SaveIdx As String
  If GIdx > 0 Then
    SaveIdx = fpListEmp.ColText
  End If
  Call CheckSort
  fpListEmp.SearchText = SaveIdx
  fpListEmp.Action = 0
  GIdx = fpListEmp.SearchIndex

End Sub

Private Sub cmdDelete_Click()
  Dim ORec As OrbitDetail
  Dim OHandle As Integer
  Dim NumOfORecs As Integer
  Dim Answer As VbMsgBoxResult
  
  On Error GoTo ERRORSTUFF
 
  Answer = MsgBox("Are you sure you wish to delete this transaction?", vbYesNo) = vbNo
  If Answer = No Then
    Close
    Exit Sub
  End If
  
  If AddNew = False Then
    OpenOrbDetail OHandle, NumOfORecs
    Get OHandle, GRecNum, ORec
    ORec.Deleted = True
    Put OHandle, GRecNum, ORec
    Close
    Call cmdClear_Click
  End If
  
  Call NameSort
  
  MsgBox ("This transaction has been deleted successfully.")
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmORBITEdit", "cmdDelete_Click", Erl)
     Case emrExitProc:
       Resume Proc_Exit
     Case emrResume:
       Resume
     Case emrResumeNext:
       Resume Next
     Case Else
      '--- Technically, this should never happen.
       Resume Proc_Exit
   End Select
  
Proc_Exit:
  '--- Cleanup code goes here...
    Close
  
End Sub

Private Sub cmdGetOld_Click()
  Call cmdClear_Click
  frmEmpOrbitOldTransList.Show vbModal
  DoEvents
End Sub

Private Sub cmdSave_Click()
  Dim ORec As OrbitDetail
  Dim OHandle As Integer
  Dim NumOfORecs As Integer
  Dim OERec As OrbitEmpData
  Dim OEHandle As Integer
  Dim NumOfOERecs As Integer
  Dim x As Integer
  Dim CheckDate As String
  Dim saveHere As Long
  Dim Answer As VbMsgBoxResult
  Dim EmpRec2 As EmpData2Type
  Dim EHandle As Integer
  Dim SaveRow As Integer
  
  On Error GoTo ERRORSTUFF
  SaveRow = fpListEmp.ListIndex
  If Check4MiddleName(fptxtFirstName.Text) = False Then
    Exit Sub
  End If
  
  If Check4Reqd = False Then
    Exit Sub
  End If
  
  If QPTrim$(fptxtOutOfCountry.Text) <> "" Then
    If QPTrim$(cmbState.Text) <> "" Then
      Answer = MsgBox("When the 'Out of Country Address' field is used the State field must be blank. Do you wish to delete the state?", vbYesNoCancel)
      If Answer = vbCancel Then
        cmbState.SetFocus
        Exit Sub
      ElseIf Answer = vbYes Then
        cmbState.Text = ""
      Else
        MainLog ("User warned that when editing ORBIT data (cust# " & GEmpNum & ") that when 'Out of Country is used the state field must be left blank. The user ignored this warning.")
      End If
    End If
    If QPTrim$(ReplaceString(txtZip.Text, "-", "")) <> "" Then
      Answer = MsgBox("When the 'Out of Country Address' field is used the Zip field must be blank. Do you wish to delete the zip?", vbYesNoCancel)
      If Answer = vbCancel Then
        txtZip.SetFocus
        Exit Sub
      ElseIf Answer = vbYes Then
        txtZip.Text = ""
      Else
        MainLog ("User warned that when editing ORBIT data (cust# " & GEmpNum & ") that when 'Out of Country is used the zip field must be left blank. The user ignored this warning.")
      End If
    End If
  End If
  OpenOrbDetail OHandle, NumOfORecs
  If AddNew = True Then
    saveHere = NumOfORecs + 1
  Else
    saveHere = GRecNum
  End If
  If saveHere = GRecNum Then
    Get OHandle, GRecNum, ORec
  End If
  If saveHere = 0 Then
    MsgBox ("ERROR: There is a problem saving this data. Please select an existing record from the list at the left or from an old record.")
    Exit Sub
  End If
  
  If fpCurrRetSalary.Value <> TempRetSalary Or fpCurrRegPay.Value <> TempRegPay Or fpCurrEmpCont.Value <> TempEmpCont Or fpCurrOTPay.Value <> TempOTPay Then
    frmMessageWOpts.Label1.Caption = "Editing transactions on this screen will NOT change the earnings history for the employee in the Payroll program and may have W2 consequences that will require an adjustment of the employee's W2 at the end of the year. Do you wish to save anyway?"
    frmMessageWOpts.Label1.Top = 650
    frmMessageWOpts.cmdCont.Text = "F10 Save Anyway"
    frmMessageWOpts.cmdExit.Text = "ESC Review"
    frmMessageWOpts.Show vbModal
    If frmMessageWOpts.fptxtChoice.Text = "abort" Then
      Unload frmMessageWOpts
      fpCurrRetSalary.SetFocus
      Exit Sub
    Else
      MainLog ("User warned about editing transactions in ORBIT not affecting earnings history and may have W2 consequences for " & QPTrim$(fptxtFirstName.Text) & " " & QPTrim$(fptxtLastName.Text) & " with check number " & QPTrim$(fpCheckNum.Text) & " and pay period beginning on " & fpdtPayPdBegDate.Text & " and ending on " & fpdtPayPdEndDate.Text & ". User chose to save anyway.")
    End If
  End If
 
  ORec.FirstName = QPTrim$(fptxtFirstName.Text)
  ORec.LastName = QPTrim$(fptxtLastName.Text)
  ORec.MiddleName = QPTrim$(fptxtInitial.Text)
  ORec.Suffix = QPTrim$(cmbSuffix.Text)
  ORec.AddLine1 = QPTrim$(fptxtAdd1.Text)
  ORec.AddLine2 = QPTrim$(fptxtAdd2.Text)
  ORec.City = QPTrim$(fptxtCity.Text)
  ORec.State = cmbState.Text
  ORec.Zip = QPTrim$(ReplaceString(txtZip.Text, "-", ""))
  ORec.SSN = ReplaceString(fpMaskSoc.Text, "-", "")
  ORec.OutOfCntryAdd = QPTrim$(fptxtOutOfCountry.Text)
  ORec.MemberID = QPTrim$(fptxtMemberID.Text)
  ORec.DeptNum = QPTrim$(fptxtDeptNum.Text)
  ORec.Gender = QPTrim$(cmbGender.Text)
  ORec.PlanCode = QPTrim$(cmbPlanCode.Text)
  ORec.JobClass = QPTrim$(cmbJobClassID.Text)
  ORec.Adjustment = QPTrim$(cmbAdjustment.Text)
  ORec.PayType = QPTrim$(cmbPayType.Text)
  ORec.ContrPdEmpPrd = QPTrim$(cmbContractPd.Text)
  ORec.TermType = QPTrim$(cmbTermCode.Text)
  If QPTrim$(fpdtTermDate.Text) <> "" Then
    ORec.TerminationDate = FormatThisDate(fpdtTermDate.Text, 2)
  Else
    ORec.TerminationDate = "0"
  End If
  ORec.EmployDate = FormatThisDate(fpdtHireDate.Text, 2)
  ORec.DateOfBirth = FormatThisDate(fpMaskBDay.Text, 2)
  If QPTrim$(fpdtEligibleDate.Text) <> "" Then
    ORec.EligibleDate = FormatThisDate(fpdtEligibleDate.Text, 2)
  Else
    ORec.EligibleDate = "0"
  End If
  ORec.VacHours = CStr(fptxtVacHrsPaid.Value)
  ORec.CheckNum = fpCheckNum.Value
'  ORec.Salary = Abs(fpCurrRetSalary.Value)
  ORec.Salary = fpCurrRetSalary.Value
  If fpCurrRetSalary.Value >= 0 Then
    ORec.IncDecSalary = "+"
  ElseIf fpCurrRetSalary.Value < 0 Then
    ORec.IncDecSalary = "-"
  End If
  ORec.RegPay = fpCurrRegPay.Value
  ORec.OTPay = fpCurrOTPay.Value
'  ORec.EmployeeCntrb = Abs(fpCurrEmpCont.Value)
  ORec.EmployeeCntrb = fpCurrEmpCont.Value
  If fpCurrEmpCont.Value >= 0 Then
    ORec.IncDecEmpleeCntrb = "+"
  ElseIf fpCurrEmpCont.Value < 0 Then
    ORec.IncDecEmpleeCntrb = "-"
  End If
  ORec.EmployerCntrb = fpCurrEmployerCont.Value
  ORec.EmpRecNum = GEmpRecNum 'saveHere
  ORec.EmpNum = GEmpNum
  ORec.PayPrdBeginDate = FormatThisDate(fpdtPayPdBegDate.Text, 2)
  ORec.PayPrdEndDate = FormatThisDate(fpdtPayPdEndDate.Text, 2)
  GRecNum = saveHere
  Put OHandle, saveHere, ORec
  Close OHandle
  
'  OpenOrbEmpData OEHandle, NumOfOERecs
'  For x = 1 To NumOfOERecs
'    Get OEHandle, x, OERec
'    If OERec.EmpNum = ORec.EmpNum Then
'      OERec.FirstName = QPTrim$(fptxtFirstName.Text)
'      OERec.LastName = QPTrim$(fptxtLastName.Text)
'      OERec.MiddleName = QPTrim$(fptxtInitial.Text)
'      OERec.Suffix = QPTrim$(cmbSuffix.Text)
'      OERec.AddLine1 = QPTrim$(fptxtAdd1.Text)
'      OERec.AddLine2 = QPTrim$(fptxtAdd2.Text)
'      OERec.City = QPTrim$(fptxtCity.Text)
'      OERec.State = cmbState.Text
'      OERec.Zip = QPTrim$(ReplaceString(txtZip.Text, "-", ""))
'      OERec.SSN = ReplaceString(fpMaskSoc.Text, "-", "")
'      OERec.OutOfCntryAdd = QPTrim$(fptxtOutOfCountry.Text)
'      OERec.MemberID = QPTrim$(fptxtMemberID.Text)
'      OERec.DeptNum = QPTrim$(fptxtDeptNum.Text)
'      OERec.Gender = QPTrim$(cmbGender.Text)
'      OERec.PlanCode = QPTrim$(cmbPlanCode.Text)
'      OERec.JobClass = QPTrim$(cmbJobClassID.Text)
'      OERec.Adjustment = QPTrim$(cmbAdjustment.Text)
'      OERec.PayType = QPTrim$(cmbPayType.Text)
'      OERec.ContrPdEmpPrd = QPTrim$(cmbContractPd.Text)
'      OERec.TermType = QPTrim$(cmbTermCode.Text)
'      If QPTrim$(fpdtTermDate.Text) <> "" Then
'        OERec.TerminationDate = FormatThisDate(fpdtTermDate.Text, 2)
'      Else
'        OERec.TerminationDate = "0"
'      End If
'      If TerminateS <> QPTrim$(fpdtTermDate.Text) Then
'        Answer = MsgBox("Do you wish to update the termination date in this employee's primary record?", vbYesNo)
'        If Answer = vbYes Then
'          OpenEmpData2File EHandle
'          Get EHandle, OERec.EmpRecNum, EmpRec2
'          EmpRec2.EMPTDATE = Date2Num(fpdtTermDate.Text)
'          Put EHandle, OERec.EmpRecNum, EmpRec2
'          TerminateS = QPTrim$(fpdtTermDate.Text)
'          Close EHandle
'        Else
'          MainLog ("The termination date on the ORBIT Edit screen was updated and the user elected to NOT update the termination date on the Employee Maintenance screen.")
'        End If
'      End If
'      OERec.EmployDate = FormatThisDate(fpdtHireDate.Text, 2)
'      OERec.DateOfBirth = FormatThisDate(fpMaskBDay.Text, 2)
'      If QPTrim$(fpdtEligibleDate.Text) <> "" Then
'        OERec.EligibleDate = FormatThisDate(fpdtEligibleDate.Text, 2)
'      Else
'        OERec.EligibleDate = "0"
'      End If
'      OERec.VacHours = CStr(fptxtVacHrsPaid.Value)
'      OERec.EmpNum = GEmpNum
'      ORec.CheckNum = fpCheckNum.Value
'      Put OEHandle, x, OERec
'      Exit For
'    End If
'  Next x
'  Close OEHandle
'
  Select Case SortType
    Case "Number"
      Call NumberSort
    Case "Name"
      Call NameSort
    Case "Check"
      Call CheckSort
    Case "Batch"
      Call BatchSort
    Case Else
      Call NumberSort
  End Select
  
'  If NumCnt > 0 Then
'    Call NumberSort
'  Else
'    Call NameSort
'  End If

  If AddNew = True Then
    Call cmdClear_Click
  End If
  
  fpListEmp.ListIndex = SaveRow
  AddNew = False
  MsgBox ("The employee ORBIT data for the current submission has been updated.")
  cmdDelete.Enabled = True
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmORBITEdit", "cmdSave_Click", Erl)
     Case emrExitProc:
       Resume Proc_Exit
     Case emrResume:
       Resume
     Case emrResumeNext:
       Resume Next
     Case Else
      '--- Technically, this should never happen.
       Resume Proc_Exit
   End Select
  
Proc_Exit:
  '--- Cleanup code goes here...
    Close
  
  
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
      SendKeys "%x"
      Call cmdExit_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%S"
      Call cmdSave_Click
      KeyCode = 0
    Case Else:
  End Select

End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Call LoadMe
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    ''Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("Payroll.exe terminated via menu bar on frmEmpORBITEdit.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub cmdExit_Click()
  frmORBITMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub LoadMe()
 fpListEmp.ScrollBarV = ScrollBarVShow
 Call NameSort
 GIdx = -1
 cmbPlanCode.AddItem ("STG")
 cmbPlanCode.AddItem ("STL")
 cmbPlanCode.AddItem ("STMAX")
 cmbPlanCode.AddItem ("STRS")
 cmbPlanCode.AddItem ("STRE")
 cmbPlanCode.AddItem ("STDIS")
 cmbPlanCode.AddItem ("LOCG")
 cmbPlanCode.AddItem ("LOCL")
 cmbPlanCode.AddItem ("LOCF")
 cmbPlanCode.AddItem ("LOCMAX")
 cmbPlanCode.AddItem ("LOCWP")
 cmbPlanCode.AddItem ("LOCROD")
 cmbPlanCode.AddItem ("LOCRS")
 cmbPlanCode.AddItem ("JUD1")
 cmbPlanCode.AddItem ("JUD2")
 cmbPlanCode.AddItem ("JUD3")
 cmbPlanCode.AddItem ("LEGL")
 cmbPlanCode.AddItem ("ORPG")
 cmbPlanCode.AddItem ("ORPMAX")
 cmbSuffix.AddItem ("II")
 cmbSuffix.AddItem ("III")
 cmbSuffix.AddItem ("IV")
 cmbSuffix.AddItem ("V")
 cmbSuffix.AddItem ("JR")
 cmbSuffix.AddItem ("SR")
 cmbSuffix.AddItem ("MD")
 cmbGender.AddItem ("Male")
 cmbGender.AddItem ("Female")
 cmbGender.AddItem ("Unknown")
 cmbState.AddItem ("NC")
 cmbState.AddItem ("SC")
 cmbState.AddItem ("VA")
 cmbState.AddItem ("GA")
 cmbState.AddItem ("TN")
 cmbState.AddItem ("AL")
 cmbState.AddItem ("OH")
 cmbState.AddItem ("AK")
 cmbJobClassID.AddItem ("100")
 cmbJobClassID.AddItem ("102")
 cmbJobClassID.AddItem ("103")
 cmbJobClassID.AddItem ("104")
 cmbJobClassID.AddItem ("105")
 cmbJobClassID.AddItem ("200")
 cmbJobClassID.AddItem ("201")
 cmbJobClassID.AddItem ("202")
 cmbJobClassID.AddItem ("203")
 cmbJobClassID.AddItem ("204")
 cmbJobClassID.AddItem ("300")
 cmbJobClassID.AddItem ("301")
 cmbJobClassID.AddItem ("302")
 cmbJobClassID.AddItem ("303")
 cmbJobClassID.AddItem ("400")
 cmbJobClassID.AddItem ("401")
 cmbJobClassID.AddItem ("500")
 cmbJobClassID.AddItem ("501")
 cmbJobClassID.AddItem ("502")
 cmbJobClassID.AddItem ("503")
 cmbJobClassID.AddItem ("504")
 cmbJobClassID.AddItem ("505")
 cmbJobClassID.AddItem ("506")
 cmbJobClassID.AddItem ("507")
 cmbJobClassID.AddItem ("508")
 cmbJobClassID.AddItem ("509")
 cmbJobClassID.AddItem ("600")
 cmbJobClassID.AddItem ("601")
 cmbAdjustment.AddItem ("NA")
 cmbAdjustment.AddItem ("PRIOR")
 cmbAdjustment.AddItem ("RETRO")
 cmbPayType.AddItem ("REG")
 cmbPayType.AddItem ("BONUS")
 cmbPayType.AddItem ("ANNLONG")
 cmbPayType.AddItem ("ANNLEAVE")
 cmbPayType.AddItem ("OVERTIME")
 cmbPayType.AddItem ("WORKCOMP")
 cmbPayType.AddItem ("LEAVEPAY")
 cmbPayType.AddItem ("SUMMERPAY")
 cmbContractPd.AddItem ("08")
 cmbContractPd.AddItem ("09")
 cmbContractPd.AddItem ("10")
 cmbContractPd.AddItem ("11")
 cmbContractPd.AddItem ("12")
 cmbTermCode.AddItem ("NA")
 cmbTermCode.AddItem ("RETIRE")
 cmbTermCode.AddItem ("DEATH")
 cmbTermCode.AddItem ("VOL")
 cmbTermCode.AddItem ("INVOL")
End Sub

Private Sub NameSort()
  Dim ORec As OrbitDetail
  Dim OHandle As Integer
  Dim NumOfORecs As Integer
  Dim x As Integer
  Dim Big As String
  Dim Thisx As Integer
  Dim BigSave As String
  Dim NextRec As Integer
  Dim SaveName As String
  Dim SaveNum As String
  Dim TempCnt As Integer
  Dim SaveBig As String
  Dim ThisRec As Integer
  Dim CheckDate As String
  Dim BegDate As String
  Dim EndDate As String
  Dim PayPeriod As String
  
  On Error GoTo ERRORSTUFF
  SortType = "Name"
  NumCnt = 0
  OpenOrbDetail OHandle, NumOfORecs
  ReDim TempOIDX(1 To 1) As String
  ReDim TempNameIDX(1 To 1) As String
  TempCnt = 0
  For x = 1 To NumOfORecs
    Get OHandle, x, ORec
'    If ORec.EmpNum = 560020 Then Stop
    If ORec.Deleted = True Then GoTo Nope
    TempCnt = TempCnt + 1
    ReDim Preserve TempOIDX(1 To TempCnt) As String
    ReDim Preserve TempNameIDX(1 To TempCnt) As String
    BegDate = Mid(ORec.PayPrdBeginDate, 5, 2) & "/" & Mid(ORec.PayPrdBeginDate, 7, 2)
    EndDate = Mid(ORec.PayPrdEndDate, 5, 2) & "/" & Mid(ORec.PayPrdEndDate, 7, 2)
    PayPeriod = BegDate & " - " & EndDate
    
    CheckDate = MakeRegDate(ORec.CheckDate)
    TempOIDX(TempCnt) = PayPeriod & " " & CStr(ORec.CheckNum) & " " & QPTrim(ORec.EmpNum) & " " & QPTrim$(ORec.LastName) & ", " & QPTrim$(ORec.FirstName) & " " & Using$("$##,###.##", ORec.Salary)
    TempNameIDX(TempCnt) = QPTrim$(ORec.LastName) & ", " & QPTrim$(ORec.FirstName) & " " & Using$("$##,###.##", ORec.Salary)
Nope:
  Next x
  
  Big = ""
  For x = 1 To NumOfORecs
    Get OHandle, x, ORec
    If ORec.Deleted = True Then GoTo NoNo
    If ORec.LastName > Big Then
      Big = QPTrim$(ORec.LastName) & ", " & QPTrim$(ORec.FirstName)
    End If
NoNo:
  Next x
  Close OHandle
  SaveBig = Big + "z"
  
  Big = SaveBig
  NextRec = 1
  Do
    For x = NextRec To TempCnt
      If TempNameIDX(x) < Big Then
        Big = TempNameIDX(x)
        ThisRec = x
      End If
    Next x
    SaveName = TempNameIDX(NextRec)
    SaveNum = TempOIDX(NextRec)
    TempNameIDX(NextRec) = TempNameIDX(ThisRec)
    TempOIDX(NextRec) = TempOIDX(ThisRec)
    TempNameIDX(ThisRec) = SaveName
    TempOIDX(ThisRec) = SaveNum
    NextRec = NextRec + 1
    If NextRec > TempCnt Then Exit Do
    Big = SaveBig
  Loop
  NameCnt = TempCnt
  ReDim NameIdx(1 To NameCnt) As String
  For x = 1 To NameCnt
     NameIdx(x) = TempOIDX(x)
  Next x
  fpListEmp.Clear
  For x = 1 To NameCnt
    fpListEmp.AddItem (NameIdx(x))
  Next x
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmORBITEdit", "NameSort", Erl)
     Case emrExitProc:
       Resume Proc_Exit
     Case emrResume:
       Resume
     Case emrResumeNext:
       Resume Next
     Case Else
      '--- Technically, this should never happen.
       Resume Proc_Exit
   End Select
  
Proc_Exit:
  '--- Cleanup code goes here...
    Close
  
End Sub

Private Sub NumberSort()
  Dim ORec As OrbitDetail
  Dim OHandle As Integer
  Dim NumOfORecs As Integer
  Dim x As Integer
  Dim Big As Long
  Dim Thisx As Integer
  Dim NextRec As Integer
  Dim SaveName As String
  Dim SaveNum As Long
  Dim TempCnt As Integer
  Dim SaveBig As Long
  Dim ThisRec As Integer
  Dim CheckDate As String
  Dim BegDate As String
  Dim EndDate As String
  Dim PayPeriod As String
  
  On Error GoTo ERRORSTUFF
  SortType = "Number"
  NameCnt = 0
  OpenOrbDetail OHandle, NumOfORecs
  ReDim TempOIDX(1 To 1) As String
  ReDim TempNumIDX(1 To 1) As Long
  TempCnt = 0
  For x = 1 To NumOfORecs
    Get OHandle, x, ORec
    If ORec.Deleted = True Then GoTo Nope
    TempCnt = TempCnt + 1
    ReDim Preserve TempOIDX(1 To TempCnt) As String
    ReDim Preserve TempNumIDX(1 To TempCnt) As Long
    BegDate = Mid(ORec.PayPrdBeginDate, 5, 2) & "/" & Mid(ORec.PayPrdBeginDate, 7, 2)
    EndDate = Mid(ORec.PayPrdEndDate, 5, 2) & "/" & Mid(ORec.PayPrdEndDate, 7, 2)
    PayPeriod = BegDate & " - " & EndDate
    
    CheckDate = MakeRegDate(ORec.CheckDate)
    TempOIDX(TempCnt) = PayPeriod & " " & CStr(ORec.CheckNum) & " " & QPTrim(ORec.EmpNum) & " " & QPTrim$(ORec.LastName) & ", " & QPTrim$(ORec.FirstName) & " " & Using$("$##,###.##", ORec.Salary)
    TempNumIDX(TempCnt) = CLng(ORec.EmpNum)
Nope:
  Next x
  
  Big = 0
  For x = 1 To NumOfORecs
    Get OHandle, x, ORec
    If ORec.Deleted = True Then GoTo NoNo
    If ORec.EmpNum > Big Then
      Big = ORec.EmpNum
    End If
NoNo:
  Next x
  Close OHandle
  SaveBig = Big + "1"
  
  Big = SaveBig
  NextRec = 1
  Do
    For x = NextRec To TempCnt
      If TempNumIDX(x) < Big Then
        Big = TempNumIDX(x)
        ThisRec = x
      End If
    Next x
    SaveName = TempOIDX(NextRec)
    SaveNum = TempNumIDX(NextRec)
    TempNumIDX(NextRec) = TempNumIDX(ThisRec)
    TempOIDX(NextRec) = TempOIDX(ThisRec)
    TempNumIDX(ThisRec) = SaveNum
    TempOIDX(ThisRec) = SaveName
    NextRec = NextRec + 1
    If NextRec > TempCnt Then Exit Do
    Big = SaveBig
  Loop
  NumCnt = TempCnt
  ReDim NumIdx(1 To NumCnt) As String
  For x = 1 To NumCnt
     NumIdx(x) = TempOIDX(x)
  Next x
  fpListEmp.Clear
  For x = 1 To NumCnt
    fpListEmp.AddItem (NumIdx(x))
  Next x
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmORBITEdit", "NumberSort", Erl)
     Case emrExitProc:
       Resume Proc_Exit
     Case emrResume:
       Resume
     Case emrResumeNext:
       Resume Next
     Case Else
      '--- Technically, this should never happen.
       Resume Proc_Exit
   End Select
  
Proc_Exit:
  '--- Cleanup code goes here...
    Close
  
End Sub

Private Sub fpdtTermDate_Change()
  If fpdtTermDate.Text <> "" Then
   fptxtVacHrsPaid.Enabled = True
   cmbTermCode.Enabled = True
   Label28.Caption = "*Termination Code:"
  Else
   fptxtVacHrsPaid.Enabled = False
   fptxtVacHrsPaid.Text = "0"
   cmbTermCode.Enabled = False
   cmbTermCode.Text = "NA"
   Label28.Caption = "Termination Code:"
  End If
End Sub

Private Sub fpListEmp_DblClick()
  Static Repeat As Boolean
  Dim Answer As VbMsgBoxResult
  
  On Error GoTo ERRORSTUFF
  
  If Repeat = True Then
    Repeat = False
    Exit Sub
  End If
  Repeat = False
  
  If QPTrim$(fpdtTermDate.Text) <> "" Then
    If QPTrim$(cmbTermCode.Text) = "NA" Then
      MsgBox ("A termination date has been set but the required termination code has not been set. Please correct this situation.")
      cmbTermCode.SetFocus
      Close
      Exit Sub
    End If
  End If
  
  If GIdx > -1 And AddNew = False Then
    If Check4Changes = True Then
      Answer = MsgBox("Changes have been made. Do you wish to save them?", vbYesNoCancel)
      If Answer = vbCancel Then
        Repeat = True
        fpListEmp.ListIndex = GIdx
        Exit Sub
      ElseIf Answer = vbYes Then
        Call cmdSave_Click
      End If
    End If
  End If
  GIdx = fpListEmp.ListIndex
  Call LoadMeEmp
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmORBITEdit", "fpListEmp_Click", Erl)
     Case emrExitProc:
       Resume Proc_Exit
     Case emrResume:
       Resume
     Case emrResumeNext:
       Resume Next
     Case Else
      '--- Technically, this should never happen.
       Resume Proc_Exit
   End Select
  
Proc_Exit:
  '--- Cleanup code goes here...
    Close
  
  
  
  
End Sub

Private Sub LoadMeEmp()
  Dim ORec As OrbitDetail
  Dim OHandle As Integer
  Dim NumOfORecs As Integer
  Dim x As Integer
  Dim MatchThis As String
  Dim BegDate As String
  Dim EndDate As String
  Dim PayPeriod As String
  Dim CheckDate As String
  
  On Error GoTo ERRORSTUFF
  
  MatchThis = fpListEmp.ColText
  cmdDelete.Enabled = True
  OpenOrbDetail OHandle, NumOfORecs
  
  AddNew = False
  For x = 1 To NumOfORecs
    Get OHandle, x, ORec
    If ORec.Deleted = True Then GoTo Deleted
'    If ORec.EmpNum = 560020 Then Stop
    BegDate = Mid(ORec.PayPrdBeginDate, 5, 2) & "/" & Mid(ORec.PayPrdBeginDate, 7, 2)
    EndDate = Mid(ORec.PayPrdEndDate, 5, 2) & "/" & Mid(ORec.PayPrdEndDate, 7, 2)
    PayPeriod = BegDate & " - " & EndDate
    CheckDate = MakeRegDate(ORec.CheckDate)
    If MatchThis = PayPeriod & " " & CStr(ORec.CheckNum) & " " & QPTrim(ORec.EmpNum) & " " & QPTrim$(ORec.LastName) & ", " & QPTrim$(ORec.FirstName) & " " & Using$("$##,###.##", ORec.Salary) Then
      GRecNum = x ' ORec.EmpRecNum
      GEmpRecNum = ORec.EmpRecNum
      GEmpNum = ORec.EmpNum
      fptxtFirstName.Text = QPTrim$(ORec.FirstName)
      fptxtLastName.Text = QPTrim$(ORec.LastName)
      fptxtInitial.Text = QPTrim$(ORec.MiddleName)
      cmbSuffix.Text = QPTrim$(ORec.Suffix)
      fptxtAdd1.Text = QPTrim$(ORec.AddLine1)
      fptxtAdd2.Text = QPTrim$(ORec.AddLine2)
      fptxtCity.Text = QPTrim$(ORec.City)
      cmbState.Text = ORec.State
      txtZip.Text = QPTrim$(ORec.Zip)
      fpMaskSoc.Text = ORec.SSN
      fptxtOutOfCountry.Text = QPTrim$(ORec.OutOfCntryAdd)
      fptxtMemberID.Text = QPTrim$(ORec.MemberID)
      cmbGender.Text = QPTrim$(ORec.Gender)
      cmbPlanCode.Text = QPTrim$(ORec.PlanCode)
      cmbJobClassID.Text = QPTrim$(ORec.JobClass)
      cmbAdjustment.Text = QPTrim$(ORec.Adjustment)
      cmbPayType.Text = QPTrim$(ORec.PayType)
      cmbContractPd.Text = QPTrim$(ORec.ContrPdEmpPrd)
      cmbTermCode.Text = QPTrim$(ORec.TermType)
      fptxtVacHrsPaid.Value = ORec.VacHours
      If ORec.TerminationDate <> 0 Then
       fpdtTermDate.Text = Mid(ORec.TerminationDate, 5, 2) & "/" & Mid(ORec.TerminationDate, 7, 2) & "/" & Mid(ORec.TerminationDate, 1, 4)
      Else
       fpdtTermDate.Text = ""
      End If
      TerminateS = QPTrim$(fpdtTermDate.Text)
      If ORec.EmployDate <> 0 Then
       fpdtHireDate.Text = Mid(ORec.EmployDate, 5, 2) & "/" & Mid(ORec.EmployDate, 7, 2) & "/" & Mid(ORec.EmployDate, 1, 4)
      Else
       fpdtHireDate.Text = ""
      End If
      If ORec.DateOfBirth <> "0" Then
        fpMaskBDay.Text = Mid(ORec.DateOfBirth, 5, 2) & "/" & Mid(ORec.DateOfBirth, 7, 2) & "/" & Mid(ORec.DateOfBirth, 1, 4)
      Else
        fpMaskBDay.Text = ""
      End If
      If ORec.PayPrdBeginDate <> "0" Then
        fpdtPayPdBegDate.Text = Mid(ORec.PayPrdBeginDate, 5, 2) & "/" & Mid(ORec.PayPrdBeginDate, 7, 2) & "/" & Mid(ORec.PayPrdBeginDate, 1, 4)
      Else
        fpdtPayPdBegDate.Text = ""
      End If
      If ORec.PayPrdEndDate <> "0" Then
        fpdtPayPdEndDate.Text = Mid(ORec.PayPrdEndDate, 5, 2) & "/" & Mid(ORec.PayPrdEndDate, 7, 2) & "/" & Mid(ORec.PayPrdEndDate, 1, 4)
      Else
        fpdtPayPdEndDate.Text = ""
      End If
      If QPTrim$(ORec.EligibleDate) <> "0" Then
        fpdtEligibleDate.Text = Mid(ORec.EligibleDate, 5, 2) & "/" & Mid(ORec.EligibleDate, 7, 2) & "/" & Mid(ORec.EligibleDate, 1, 4)
      Else
        fpdtEligibleDate.Text = ""
      End If
      fpCurrRetSalary.Value = ORec.Salary
      fpCurrRegPay.Value = ORec.RegPay
      fpCurrOTPay.Value = ORec.OTPay
      fpCurrEmpCont.Value = ORec.EmployeeCntrb
      fpCurrEmployerCont.Value = ORec.EmployerCntrb
      fpCheckNum.Value = ORec.CheckNum
      fptxtDeptNum.Text = ORec.DeptNum
    End If
Deleted:
  Next x
  Close OHandle
  Call LoadTemps

  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmORBITEdit", "LoadMeEmp", Erl)
     Case emrExitProc:
       Resume Proc_Exit
     Case emrResume:
       Resume
     Case emrResumeNext:
       Resume Next
     Case Else
      '--- Technically, this should never happen.
       Resume Proc_Exit
   End Select
  
Proc_Exit:
  '--- Cleanup code goes here...
    Close
  
End Sub

Private Sub fptxtSearch_Change()
  Dim x As Integer
  On Error GoTo ERRORSTUFF

  For x = 0 To fpListEmp.ListCount - 1
    fpListEmp.ListIndex = x
    If InStr(1, fpListEmp.ColText, fptxtSearch.Text) > 0 Then
      fpListEmp.ListIndex = x
      Exit For
    End If
  Next x
  If x > fpListEmp.ListCount - 1 Then
    fpListEmp.ListIndex = -1
  End If
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmORBITEdit", "fptxtSearch_Change", Erl)
     Case emrExitProc:
       Resume Proc_Exit
     Case emrResume:
       Resume
     Case emrResumeNext:
       Resume Next
     Case Else
      '--- Technically, this should never happen.
       Resume Proc_Exit
   End Select
  
Proc_Exit:
  '--- Cleanup code goes here...
    Close
  
  
  
End Sub

Private Function Check4Reqd() As Boolean
  Dim ThisText As String
  
  On Error GoTo ERRORSTUFF
  
  Check4Reqd = True
  
  If fptxtLastName.Text = "" Then
    MsgBox ("A value for Employee Last Name is required.")
    fptxtLastName.SetFocus
    Check4Reqd = False
    Exit Function
  End If
  
  If fptxtFirstName.Text = "" Then
    MsgBox ("A value for Employee First Name is required.")
    fptxtFirstName.SetFocus
    Check4Reqd = False
    Exit Function
  End If
  
  If fptxtAdd1.Text = "" Then
    MsgBox ("A value for Address 1 is required.")
    fptxtAdd1.SetFocus
    Check4Reqd = False
    Exit Function
  End If
  
  If fptxtCity.Text = "" Then
    MsgBox ("A value for City is required.")
    fptxtCity.SetFocus
    Check4Reqd = False
    Exit Function
  End If
 
  If QPTrim$(cmbState.Text) = "" Then
    MsgBox ("A value for State is required.")
    cmbState.SetFocus
    Check4Reqd = False
    Exit Function
  End If
 
  ThisText = ReplaceString(txtZip.Text, "-", "")
  If ThisText = "" Then
    MsgBox ("A value for Zip is required.")
    txtZip.SetFocus
    Check4Reqd = False
    Exit Function
  End If
  
  ThisText = ReplaceString(fpMaskSoc.Text, "-", "")
  If ThisText = "" Then
    MsgBox ("A value for Social Security Number is required.")
    fpMaskSoc.SetFocus
    Check4Reqd = False
    Exit Function
  End If
  
  If Len(ThisText) <> 9 Then
    MsgBox ("The Social Security Number must be nine digits.")
    fpMaskSoc.SetFocus
    Check4Reqd = False
    Exit Function
  End If
  
  If cmbGender.Text = "" Then
    MsgBox ("A value for Gender is required.")
    cmbGender.SetFocus
    Check4Reqd = False
    Exit Function
  End If
  
  If cmbPlanCode.Text = "" Then
    MsgBox ("A value for Plan Code is required.")
    cmbPlanCode.SetFocus
    Check4Reqd = False
    Exit Function
  End If
  
  If cmbJobClassID.Text = "" Then
    MsgBox ("A value for Job Class ID is required.")
    cmbJobClassID.SetFocus
    Check4Reqd = False
    Exit Function
  End If
  
  If cmbPayType.Text = "" Then
    MsgBox ("A value for Pay Type is required.")
    cmbPayType.SetFocus
    Check4Reqd = False
    Exit Function
  End If
  
  If cmbContractPd.Text = "" Then
    MsgBox ("A value for Employment Period is required.")
    cmbContractPd.SetFocus
    Check4Reqd = False
    Exit Function
  End If
  
  If fptxtMemberID.Text = "" Then
    If MsgBox("There is no value for Member ID. This is permissible if this is a new employee. Otherwise this field must not be left empty. Do you wish to leave this field empty?", vbYesNo) = vbNo Then
      fptxtMemberID.SetFocus
      Check4Reqd = False
      Exit Function
    End If
  End If
  
  If fpdtHireDate.Text = "" Then
    MsgBox ("A value for Employment Date is required.")
    fpdtHireDate.SetFocus
    Check4Reqd = False
    Exit Function
  End If
  
  If fpMaskBDay.Text = "" Then
    MsgBox ("A value for Birth Date is required.")
    fpMaskBDay.SetFocus
    Check4Reqd = False
    Exit Function
  End If
  
  If fpdtTermDate.Text <> "" Then
    If cmbTermCode.Text = "NA" Then
      MsgBox ("A value for Termination Code is required if a Termination Date exists.")
      fpdtTermDate.SetFocus
      Check4Reqd = False
      Exit Function
    End If
  End If
  
  If fpdtPayPdBegDate.Text = "" Then
    MsgBox ("A value for Pay Period Begin is required.")
    fpdtPayPdBegDate.SetFocus
    Check4Reqd = False
    Exit Function
  End If
  
  If fpdtPayPdEndDate.Text = "" Then
    MsgBox ("A value for Pay Period End is required.")
    fpdtPayPdEndDate.SetFocus
    Check4Reqd = False
    Exit Function
  End If
  
  If fpdtEligibleDate.Text <> "" Then
    If Date2Num(fpdtEligibleDate.Text) < Date2Num(fpdtHireDate.Text) Then
      MsgBox ("The Eligibility Date cannot come before the Hire Date.")
      fpdtEligibleDate.SetFocus
      Check4Reqd = False
      Exit Function
    End If
  End If
  
  If QPTrim$(fptxtAdd1.Text) = "" And QPTrim$(fptxtAdd2.Text) <> "" Then
    MsgBox ("Address 2 must be accompanied with an Address 1.")
    fptxtAdd2.SetFocus
    Check4Reqd = False
    Exit Function
  End If
  
  If QPTrim$(fpdtTermDate.Text) <> "" Then
    If Date2Num(fpdtTermDate.Text) < Date2Num(fpdtHireDate.Text) Then
      MsgBox ("The termination date cannot come before the employment date.")
      fpdtTermDate.SetFocus
      Check4Reqd = False
      Exit Function
    End If
  End If
  
  If fpCurrRetSalary.Value < 0 Or fpCurrEmployerCont.Value < 0 Then
    If cmbAdjustment.Text <> "PRIOR" Then
      MsgBox ("If reporting a negative gross retirement pay the Adjustment value must be set to 'PRIOR'.")
      cmbAdjustment.SetFocus
      Check4Reqd = False
      Exit Function
    End If
  End If
  
  If fpCurrEmpCont.Value < 0 Or fpCurrEmployerCont.Value < 0 Then
    If cmbAdjustment.Text <> "PRIOR" Then
      MsgBox ("If reporting a negative employee contribution amount the Adjustment value must be set to 'PRIOR'.")
      cmbAdjustment.SetFocus
      Check4Reqd = False
      Exit Function
    End If
  End If
  
  Exit Function
  
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmORBITEdit", "Check4Reqd", Erl)
     Case emrExitProc:
       Resume Proc_Exit
     Case emrResume:
       Resume
     Case emrResumeNext:
       Resume Next
     Case Else
      '--- Technically, this should never happen.
       Resume Proc_Exit
   End Select
  
Proc_Exit:
  '--- Cleanup code goes here...
    Close
  
  
  
End Function

Private Sub cmbAdjustment_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    cmbAdjustment.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    cmbAdjustment.ListIndex = -1
  End If
  If cmbAdjustment.ListDown <> True Then
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

Private Sub cmbContractPd_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    cmbContractPd.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    cmbContractPd.ListIndex = -1
  End If
  If cmbContractPd.ListDown <> True Then
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

Private Sub cmbGender_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    cmbGender.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    cmbGender.ListIndex = -1
  End If
  If cmbGender.ListDown <> True Then
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

Private Sub cmbJobClassID_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    cmbJobClassID.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    cmbJobClassID.ListIndex = -1
  End If
  If cmbJobClassID.ListDown <> True Then
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

Private Sub cmbPayType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    cmbPayType.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    cmbPayType.ListIndex = -1
  End If
  If cmbPayType.ListDown <> True Then
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

Private Sub cmbPlanCode_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    cmbPlanCode.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    cmbPlanCode.ListIndex = -1
  End If
  If cmbPlanCode.ListDown <> True Then
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

Private Sub cmbState_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    cmbState.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    cmbState.ListIndex = -1
  End If
  If cmbState.ListDown <> True Then
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

Private Sub cmbSuffix_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    cmbSuffix.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    cmbSuffix.ListIndex = -1
  End If
  If cmbSuffix.ListDown <> True Then
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

Private Sub cmbTermCode_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    cmbTermCode.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    cmbTermCode.ListIndex = -1
  End If
  If cmbTermCode.ListDown <> True Then
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

Private Function FormatThisDate(ByRef ThisDate As String, ByVal Vers As Integer) As String
  Dim ch As String
  Dim DateLen As Integer
  Dim FSPstn As Integer
  Dim x As Integer
  Dim ThisDay As String
  Dim ThisMonth As String
  Dim ThisYear As String
  
  On Error GoTo ERRORSTUFF
  
  FSPstn = 0
  DateLen = Len(ThisDate)
  For x = 1 To DateLen
    ch = Mid(ThisDate, x, 1)
    If ch = "/" Then
      FSPstn = x
      Exit For
    End If
  Next x
  
  ThisMonth = Mid(ThisDate, 1, FSPstn - 1)
  If Len(ThisMonth) = 1 Then ThisMonth = "0" & ThisMonth
  ThisDay = Mid(ThisDate, FSPstn + 1, 2)
  If Len(ThisDay) = 1 Then ThisDay = "0" + ThisDay
  ThisYear = Mid(ThisDate, DateLen - 3, DateLen)
  If Vers = 2 Then
    ThisDate = ThisYear & ThisMonth & ThisDay
  ElseIf Vers = 1 Then
    ThisDate = ThisYear & ThisMonth
  End If
  FormatThisDate = ThisDate
  Exit Function
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmORBITEdit", "FormatThisDate", Erl)
     Case emrExitProc:
       Resume Proc_Exit
     Case emrResume:
       Resume
     Case emrResumeNext:
       Resume Next
     Case Else
      '--- Technically, this should never happen.
       Resume Proc_Exit
   End Select
  
Proc_Exit:
  '--- Cleanup code goes here...
    Close
  
  
  
End Function

Private Function Check4Changes() As Boolean
  Dim ORec As OrbitDetail
  Dim OHandle As Integer
  Dim NumOfORecs As Integer
  Dim Change As Boolean
  Dim ThisAmt As Double
  Dim ThatAmt As Double
  
  On Error GoTo ERRORSTUFF
  
  Change = False
  Check4Changes = False
  OpenOrbDetail OHandle, NumOfORecs
  Get OHandle, GRecNum, ORec
  
  If QPTrim$(fptxtFirstName.Text) <> QPTrim$(ORec.FirstName) Then
    Change = True
    GoTo ChangesMade
  End If
  If QPTrim$(fptxtLastName.Text) <> QPTrim$(ORec.LastName) Then
    Change = True
    GoTo ChangesMade
  End If
  If QPTrim$(fptxtInitial.Text) <> QPTrim$(ORec.MiddleName) Then
    Change = True
    GoTo ChangesMade
  End If
  If QPTrim$(cmbSuffix.Text) <> QPTrim$(ORec.Suffix) Then
    Change = True
    GoTo ChangesMade
  End If
  If QPTrim$(fptxtAdd1.Text) <> QPTrim$(ORec.AddLine1) Then
    Change = True
    GoTo ChangesMade
  End If
  If QPTrim$(fptxtAdd2.Text) <> QPTrim$(ORec.AddLine2) Then
    Change = True
    GoTo ChangesMade
  End If
  If QPTrim$(fptxtCity.Text) <> QPTrim$(ORec.City) Then
    Change = True
    GoTo ChangesMade
  End If
  If QPTrim$(cmbState.Text) <> QPTrim$(ORec.State) Then
    Change = True
    GoTo ChangesMade
  End If
  If ReplaceString(txtZip.Text, "-", "") <> QPTrim$(ORec.Zip) Then
    Change = True
    GoTo ChangesMade
  End If
 
  If ReplaceString$(fpMaskSoc.Text, "-", "") <> QPTrim$(ORec.SSN) Then
    Change = True
    GoTo ChangesMade
  End If
  If QPTrim$(fptxtOutOfCountry.Text) <> QPTrim$(ORec.OutOfCntryAdd) Then
    Change = True
    GoTo ChangesMade
  End If
  
  If QPTrim$(fptxtMemberID.Text) <> QPTrim$(ORec.MemberID) Then
    Change = True
    GoTo ChangesMade
  End If
  If QPTrim$(cmbGender.Text) <> QPTrim$(ORec.Gender) Then
    Change = True
    GoTo ChangesMade
  End If
  If QPTrim$(cmbPlanCode.Text) <> QPTrim$(ORec.PlanCode) Then
    Change = True
    GoTo ChangesMade
  End If
  If QPTrim$(cmbJobClassID.Text) <> QPTrim$(ORec.JobClass) Then
    Change = True
    GoTo ChangesMade
  End If
  
  If QPTrim$(cmbAdjustment.Text) <> QPTrim$(ORec.Adjustment) Then
    Change = True
    GoTo ChangesMade
  End If
 
  If QPTrim$(cmbPayType.Text) <> QPTrim$(ORec.PayType) Then
    Change = True
    GoTo ChangesMade
  End If
  
  If QPTrim$(cmbContractPd.Text) <> QPTrim$(ORec.ContrPdEmpPrd) Then
    Change = True
    GoTo ChangesMade
  End If

  If QPTrim$(cmbTermCode.Text) <> QPTrim$(ORec.TermType) Then
    Change = True
    GoTo ChangesMade
  End If
  
  If QPTrim$(fpdtTermDate.Text) = "" Then
    If QPTrim$(ORec.TerminationDate) <> "0" And QPTrim$(ORec.TerminationDate) <> "" Then
      Change = True
      GoTo ChangesMade
    End If
    GoTo Blank1:
  End If
  
  If FormatThisDate(fpdtTermDate.Text, 2) <> QPTrim$(ORec.TerminationDate) Then
    Change = True
    GoTo ChangesMade
  End If
Blank1:
  If QPTrim$(fpdtHireDate.Text) = "" Then
    If QPTrim$(ORec.EmployDate) <> "0" Or QPTrim$(ORec.EmployDate) <> "" Then
      Change = True
      GoTo ChangesMade
    End If
    GoTo Blank2:
  End If
  If FormatThisDate(fpdtHireDate.Text, 2) <> QPTrim$(ORec.EmployDate) Then
    Change = True
    GoTo ChangesMade
  End If
Blank2:
  If QPTrim$(fpMaskBDay.Text) = "" Then
    If QPTrim$(ORec.DateOfBirth) <> "0" Or QPTrim$(ORec.DateOfBirth) <> "" Then
      Change = True
      GoTo ChangesMade
    End If
    GoTo Blank3
  End If
  If FormatThisDate(fpMaskBDay.Text, 2) <> QPTrim$(ORec.DateOfBirth) Then
    Change = True
    GoTo ChangesMade
  End If
Blank3:
  ThisAmt = fpCurrRetSalary.Value
  ThatAmt = CDbl(ORec.Salary)
  If ThisAmt <> ThatAmt Then
    Change = True
    GoTo ChangesMade
  End If
  ThisAmt = fpCurrEmpCont.Value
  ThatAmt = CDbl(ORec.EmployeeCntrb)
  
  If ThisAmt <> ThatAmt Then
    Change = True
    GoTo ChangesMade
  End If
  
  If QPTrim$(fpdtPayPdBegDate.Text) = "" Then
    If QPTrim$(ORec.PayPrdBeginDate) <> "0" Or QPTrim$(ORec.PayPrdBeginDate) <> "" Then
      Change = True
      GoTo ChangesMade
    End If
    GoTo Blank4:
  End If
  If FormatThisDate(fpdtPayPdBegDate.Text, 2) <> QPTrim$(ORec.PayPrdBeginDate) Then
    Change = True
    GoTo ChangesMade
  End If
Blank4:

  If QPTrim$(fpdtPayPdEndDate.Text) = "" Then
    If QPTrim$(ORec.PayPrdEndDate) <> "0" Or QPTrim$(ORec.PayPrdEndDate) <> "" Then
      Change = True
      GoTo ChangesMade
    End If
    GoTo Blank5:
  End If
  If FormatThisDate(fpdtPayPdEndDate.Text, 2) <> QPTrim$(ORec.PayPrdEndDate) Then
    Change = True
    GoTo ChangesMade
  End If
Blank5:

ChangesMade:
  If Change = True Then
    Check4Changes = True
  End If
  
  Exit Function
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmORBITEdit", "Check4Changes", Erl)
     Case emrExitProc:
       Resume Proc_Exit
     Case emrResume:
       Resume
     Case emrResumeNext:
       Resume Next
     Case Else
      '--- Technically, this should never happen.
       Resume Proc_Exit
   End Select
  
Proc_Exit:
  '--- Cleanup code goes here...
    Close
  
  
  
End Function

Private Function Check4MiddleName(ByVal Name As String) As Boolean
  Dim ch As String * 1
  Dim x As Integer
  Dim ThisLen As Integer
  Dim Answer As VbMsgBoxResult
  
  On Error GoTo ERRORSTUFF
  
  Check4MiddleName = True
  If QPTrim$(fptxtInitial.Text) <> "" Then Exit Function
  ThisLen = Len(QPTrim$(Name))
  For x = 1 To ThisLen
    ch = Mid(Name, x, 1)
    If ch = " " Then
      Answer = MsgBox("Please note that this employee's first name might be separated into a first name and a middle name. Do you wish to edit this employee's name?", vbYesNo)
      If Answer = vbYes Then
        Check4MiddleName = False
        fptxtFirstName.SetFocus
        Exit Function
      End If
      Exit Function
    End If
  Next x
  Exit Function
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmORBITEdit", "Check4MiddleName", Erl)
     Case emrExitProc:
       Resume Proc_Exit
     Case emrResume:
       Resume
     Case emrResumeNext:
       Resume Next
     Case Else
      '--- Technically, this should never happen.
       Resume Proc_Exit
   End Select
  
Proc_Exit:
  '--- Cleanup code goes here...
    Close
  
  
      
End Function

Private Sub BatchSort()
  Dim ORec As OrbitDetail
  Dim OHandle As Integer
  Dim NumOfORecs As Integer
  Dim x As Integer
  Dim Big As Long
  Dim Thisx As Integer
  Dim NextRec As Integer
  Dim SaveName As String
  Dim SaveNum As Long
  Dim TempCnt As Integer
  Dim SaveBig As Long
  Dim ThisRec As Integer
  Dim CheckDate As String
  Dim BegDate As String
  Dim EndDate As String
  Dim PayPeriod As String
  Dim PayPrdCompare As Integer
  Dim Year As String
  Dim Month As String
  Dim Day As String
  Dim BegWholeDate As String
  Dim EndWholeDate As String
  
  On Error GoTo ERRORSTUFF
  SortType = "Batch"
  NameCnt = 0
  OpenOrbDetail OHandle, NumOfORecs
  ReDim TempOIDX(1 To 1) As String
  ReDim TempBatchIDX(1 To 1) As Long
  TempCnt = 0
  For x = 1 To NumOfORecs
    Get OHandle, x, ORec
    If ORec.Deleted = True Then GoTo Nope
    TempCnt = TempCnt + 1
    ReDim Preserve TempOIDX(1 To TempCnt) As String
    ReDim Preserve TempBatchIDX(1 To TempCnt) As Long
    BegDate = Mid(ORec.PayPrdBeginDate, 5, 2) & "/" & Mid(ORec.PayPrdBeginDate, 7, 2)
    EndDate = Mid(ORec.PayPrdEndDate, 5, 2) & "/" & Mid(ORec.PayPrdEndDate, 7, 2)
    PayPeriod = BegDate & " - " & EndDate
    GoSub FormatDate
    PayPrdCompare = Date2Num(BegWholeDate) + Date2Num(EndWholeDate)
    CheckDate = MakeRegDate(ORec.CheckDate)
    TempOIDX(TempCnt) = PayPeriod & " " & CStr(ORec.CheckNum) & " " & QPTrim(ORec.EmpNum) & " " & QPTrim$(ORec.LastName) & ", " & QPTrim$(ORec.FirstName) & " " & Using$("$##,###.##", ORec.Salary)
    TempBatchIDX(TempCnt) = PayPrdCompare
Nope:
  Next x
  
  Big = 0
  For x = 1 To NumOfORecs
    Get OHandle, x, ORec
    If ORec.Deleted = True Then GoTo NoNo
    GoSub FormatDate
    PayPrdCompare = Date2Num(BegWholeDate) + Date2Num(EndWholeDate)
    If PayPrdCompare > Big Then
      Big = PayPrdCompare
    End If
NoNo:
  Next x
  Close OHandle
  SaveBig = Big + 1
  
  Big = SaveBig
  NextRec = 1
  Do
    For x = NextRec To TempCnt
      If TempBatchIDX(x) < Big Then
        Big = TempBatchIDX(x)
        ThisRec = x
      End If
    Next x
    SaveName = TempOIDX(NextRec)
    SaveNum = TempBatchIDX(NextRec)
    TempBatchIDX(NextRec) = TempBatchIDX(ThisRec)
    TempOIDX(NextRec) = TempOIDX(ThisRec)
    TempBatchIDX(ThisRec) = SaveNum
    TempOIDX(ThisRec) = SaveName
    NextRec = NextRec + 1
    If NextRec > TempCnt Then Exit Do
    Big = SaveBig
  Loop
  NumCnt = TempCnt
  ReDim NumIdx(1 To NumCnt) As String
  For x = 1 To NumCnt
     NumIdx(x) = TempOIDX(x)
  Next x
  fpListEmp.Clear
  For x = 1 To NumCnt
    fpListEmp.AddItem (NumIdx(x))
  Next x
  Exit Sub
  
FormatDate:
  Year = Mid(ORec.PayPrdBeginDate, 1, 4)
  Month = Mid(ORec.PayPrdBeginDate, 5, 2)
  Day = Mid(ORec.PayPrdBeginDate, 7, 2)
  BegWholeDate = Month & "/" & Day & "/" & Year
  Year = Mid(ORec.PayPrdEndDate, 1, 4)
  Month = Mid(ORec.PayPrdEndDate, 5, 2)
  Day = Mid(ORec.PayPrdEndDate, 7, 2)
  EndWholeDate = Month & "/" & Day & "/" & Year
  
  Return
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmORBITEdit", "BatchSort", Erl)
     Case emrExitProc:
       Resume Proc_Exit
     Case emrResume:
       Resume
     Case emrResumeNext:
       Resume Next
     Case Else
      '--- Technically, this should never happen.
       Resume Proc_Exit
   End Select
  
Proc_Exit:
  '--- Cleanup code goes here...
    Close
  
  
  
End Sub

Private Sub CheckSort()
  Dim ORec As OrbitDetail
  Dim OHandle As Integer
  Dim NumOfORecs As Integer
  Dim x As Integer
  Dim Big As Long
  Dim Thisx As Integer
  Dim NextRec As Integer
  Dim SaveName As String
  Dim SaveNum As Long
  Dim TempCnt As Integer
  Dim SaveBig As Long
  Dim ThisRec As Integer
  Dim CheckDate As String
  Dim BegDate As String
  Dim EndDate As String
  Dim CheckNum As String
  Dim Year As String
  Dim Month As String
  Dim Day As String
  Dim BegWholeDate As String
  Dim PayPeriod As String
  Dim EndWholeDate As String
  
  On Error GoTo ERRORSTUFF
  SortType = "Check"
  NameCnt = 0
  OpenOrbDetail OHandle, NumOfORecs
  ReDim TempOIDX(1 To 1) As String
  ReDim TempCheckIDX(1 To 1) As Long
  TempCnt = 0
  For x = 1 To NumOfORecs
    Get OHandle, x, ORec
    If ORec.Deleted = True Then GoTo Nope
    TempCnt = TempCnt + 1
    ReDim Preserve TempOIDX(1 To TempCnt) As String
    ReDim Preserve TempCheckIDX(1 To TempCnt) As Long
    BegDate = Mid(ORec.PayPrdBeginDate, 5, 2) & "/" & Mid(ORec.PayPrdBeginDate, 7, 2)
    EndDate = Mid(ORec.PayPrdEndDate, 5, 2) & "/" & Mid(ORec.PayPrdEndDate, 7, 2)
    PayPeriod = BegDate & " - " & EndDate
    
    CheckNum = ORec.CheckNum
    TempOIDX(TempCnt) = PayPeriod & " " & CStr(ORec.CheckNum) & " " & QPTrim(ORec.EmpNum) & " " & QPTrim$(ORec.LastName) & ", " & QPTrim$(ORec.FirstName) & " " & Using$("$##,###.##", ORec.Salary)
    TempCheckIDX(TempCnt) = CheckNum
Nope:
  Next x
  
  Big = 0
  For x = 1 To NumOfORecs
    Get OHandle, x, ORec
    If ORec.Deleted = True Then GoTo NoNo
    If ORec.CheckNum > Big Then
      Big = ORec.CheckNum
    End If
NoNo:
  Next x
  Close OHandle
  SaveBig = Big + 1
  
  Big = SaveBig
  NextRec = 1
  Do
    For x = NextRec To TempCnt
      If TempCheckIDX(x) < Big Then
        Big = TempCheckIDX(x)
        ThisRec = x
      End If
    Next x
    SaveName = TempOIDX(NextRec)
    SaveNum = TempCheckIDX(NextRec)
    TempCheckIDX(NextRec) = TempCheckIDX(ThisRec)
    TempOIDX(NextRec) = TempOIDX(ThisRec)
    TempCheckIDX(ThisRec) = SaveNum
    TempOIDX(ThisRec) = SaveName
    NextRec = NextRec + 1
    If NextRec > TempCnt Then Exit Do
    Big = SaveBig
  Loop
  NumCnt = TempCnt
  ReDim NumIdx(1 To NumCnt) As String
  For x = 1 To NumCnt
     NumIdx(x) = TempOIDX(x)
  Next x
  fpListEmp.Clear
  For x = 1 To NumCnt
    fpListEmp.AddItem (NumIdx(x))
  Next x
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmORBITEdit", "CheckSort", Erl)
     Case emrExitProc:
       Resume Proc_Exit
     Case emrResume:
       Resume
     Case emrResumeNext:
       Resume Next
     Case Else
      '--- Technically, this should never happen.
       Resume Proc_Exit
   End Select
  
Proc_Exit:
  '--- Cleanup code goes here...
    Close
  
End Sub

Public Sub LoadMeEmpFromOld(ThisTRec As Long, ThisORec As Long)
  Dim ORec As OrbitEmpData
  Dim OHandle As Integer
  Dim NumOfORecs As Integer
  Dim x As Integer
  Dim TransRec As TransRecType
  Dim THandle As Integer
  
  On Error GoTo ERRORSTUFF
  
  AddNew = True
  OpenOrbEmpData OHandle, NumOfORecs
  Get OHandle, ThisORec, ORec
  Close OHandle
  OpenTransHistFile THandle
  Get THandle, ThisTRec, TransRec
  Close THandle
  GEmpRecNum = ORec.EmpRecNum
  GEmpNum = ORec.EmpNum
  fptxtFirstName.Text = QPTrim$(ORec.FirstName)
  fptxtLastName.Text = QPTrim$(ORec.LastName)
  fptxtInitial.Text = QPTrim$(ORec.MiddleName)
  cmbSuffix.Text = QPTrim$(ORec.Suffix)
  fptxtAdd1.Text = QPTrim$(ORec.AddLine1)
  fptxtAdd2.Text = QPTrim$(ORec.AddLine2)
  fptxtCity.Text = QPTrim$(ORec.City)
  cmbState.Text = ORec.State
  txtZip.Text = QPTrim$(ORec.Zip)
  fpMaskSoc.Text = ORec.SSN
  fptxtOutOfCountry.Text = QPTrim$(ORec.OutOfCntryAdd)
  fptxtMemberID.Text = QPTrim$(ORec.MemberID)
  cmbGender.Text = QPTrim$(ORec.Gender)
  cmbPlanCode.Text = QPTrim$(ORec.PlanCode)
  cmbJobClassID.Text = QPTrim$(ORec.JobClass)
  cmbAdjustment.Text = "PRIOR"
  cmbPayType.Text = QPTrim$(ORec.PayType)
  cmbContractPd.Text = QPTrim$(ORec.ContrPdEmpPrd)
  cmbTermCode.Text = QPTrim$(ORec.TermType)
  fptxtVacHrsPaid.Value = ORec.VacHours
  If QPTrim$(ORec.TerminationDate) <> "" Then
    fpdtTermDate.Text = Mid(ORec.TerminationDate, 5, 2) & "/" & Mid(ORec.TerminationDate, 7, 2) & "/" & Mid(ORec.TerminationDate, 1, 4)
  Else
    fpdtTermDate.Text = ""
  End If
  TerminateS = QPTrim$(fpdtTermDate.Text)
  If QPTrim$(ORec.EmployDate) <> "" Then
    fpdtHireDate.Text = Mid(ORec.EmployDate, 5, 2) & "/" & Mid(ORec.EmployDate, 7, 2) & "/" & Mid(ORec.EmployDate, 1, 4)
  Else
    fpdtHireDate.Text = ""
  End If
  If QPTrim$(ORec.DateOfBirth) <> "" Then
    fpMaskBDay.Text = Mid(ORec.DateOfBirth, 5, 2) & "/" & Mid(ORec.DateOfBirth, 7, 2) & "/" & Mid(ORec.DateOfBirth, 1, 4)
  Else
    fpMaskBDay.Text = ""
  End If
  fpdtPayPdBegDate.Text = MakeRegDate(TransRec.PayPdStart)
  fpdtPayPdEndDate.Text = MakeRegDate(TransRec.PayPdEnd)
  If QPTrim$(ORec.EligibleDate) <> "" Then
    fpdtEligibleDate.Text = Mid(ORec.EligibleDate, 5, 2) & "/" & Mid(ORec.EligibleDate, 7, 2) & "/" & Mid(ORec.EligibleDate, 1, 4)
  Else
    fpdtEligibleDate.Text = ""
  End If
  fpCurrRetSalary.Value = TransRec.RetGrossPay
  fpCurrRegPay.Value = TransRec.TotRegWage
  fpCurrOTPay.Value = TransRec.TotOTWage
  fpCurrEmpCont.Value = TransRec.RetireAmt
  fpCurrEmployerCont.Value = TransRec.MatchRetAmt
  fpCheckNum.Value = TransRec.CheckNum
  Call LoadTemps
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmORBITEdit", "LoadMeEmpFromOld", Erl)
     Case emrExitProc:
       Resume Proc_Exit
     Case emrResume:
       Resume
     Case emrResumeNext:
       Resume Next
     Case Else
      '--- Technically, this should never happen.
       Resume Proc_Exit
   End Select
  
Proc_Exit:
  '--- Cleanup code goes here...
    Close
  
End Sub

Private Sub LoadTemps()
  TempRetSalary = fpCurrRetSalary.Value
  TempRegPay = fpCurrRegPay.Value
  TempEmpCont = fpCurrEmpCont.Value
  TempOTPay = fpCurrOTPay.Value
End Sub
