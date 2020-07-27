VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmIntFaceMaint 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Interface Maintenance"
   ClientHeight    =   8670
   ClientLeft      =   1905
   ClientTop       =   120
   ClientWidth     =   11655
   Icon            =   "frmIntFaceMaint.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8449.949
   ScaleMode       =   0  'User
   ScaleWidth      =   11652
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo comboSplitDeDuc 
      Height          =   405
      Left            =   3210
      TabIndex        =   10
      ToolTipText     =   "No help for this field"
      Top             =   3555
      Width           =   1080
      _Version        =   196608
      _ExtentX        =   1905
      _ExtentY        =   714
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
      ColDesigner     =   "frmIntFaceMaint.frx":08CA
   End
   Begin LpLib.fpCombo comboInterface 
      Height          =   405
      Left            =   3930
      TabIndex        =   0
      ToolTipText     =   "Enter ""C"" for Central Depository System, ""I"" for Imprest Payroll Account, or ""P"" for Pooled cash with no Central Depository."
      Top             =   1005
      Width           =   2175
      _Version        =   196608
      _ExtentX        =   3836
      _ExtentY        =   714
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
      AutoSearchFill  =   -1  'True
      AutoSearchFillDelay=   200
      EditMarginLeft  =   1
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmIntFaceMaint.frx":0BF9
   End
   Begin LpLib.fpCombo fpcomboGLCheckYN 
      Height          =   405
      Left            =   960
      TabIndex        =   9
      ToolTipText     =   "Select ""Y"" to have all General Ledger number entries validated using the General Ledger list of valid numbers."
      Top             =   3555
      Width           =   1065
      _Version        =   196608
      _ExtentX        =   1879
      _ExtentY        =   714
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
      AutoSearchFill  =   -1  'True
      AutoSearchFillDelay=   500
      EditMarginLeft  =   1
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   1
      EditAlignV      =   0
      ColDesigner     =   "frmIntFaceMaint.frx":0F28
   End
   Begin LpLib.fpCombo fpcomboChkStyle 
      Height          =   405
      Left            =   5040
      TabIndex        =   63
      Top             =   3555
      Width           =   3615
      _Version        =   196608
      _ExtentX        =   6376
      _ExtentY        =   714
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
      ColDesigner     =   "frmIntFaceMaint.frx":1257
   End
   Begin EditLib.fpText txtSocSecWHNum 
      Height          =   396
      Left            =   8928
      TabIndex        =   6
      ToolTipText     =   "Enter the liability account for Social Security withholding."
      Top             =   1416
      Width           =   2172
      _Version        =   196608
      _ExtentX        =   3831
      _ExtentY        =   698
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
      ControlType     =   0
      Text            =   ""
      CharValidationText=   "1 2 3 4 5 6 7 8 9 0 -"
      MaxLength       =   14
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
   Begin EditLib.fpText txtFedWHNum 
      Height          =   396
      Left            =   3936
      TabIndex        =   4
      ToolTipText     =   "Enter the liability account for the Federal Tax withholding."
      Top             =   2688
      Width           =   2172
      _Version        =   196608
      _ExtentX        =   3831
      _ExtentY        =   698
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
      ControlType     =   0
      Text            =   ""
      CharValidationText=   "1 2 3 4 5 6 7 8 9 0 -"
      MaxLength       =   14
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
   Begin EditLib.fpCurrency fptxtIndirectCost 
      Height          =   365
      Left            =   3270
      TabIndex        =   16
      Top             =   5956
      Width           =   2130
      _Version        =   196608
      _ExtentX        =   3757
      _ExtentY        =   644
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
   Begin EditLib.fpCurrency fptxtFringeCost 
      Height          =   365
      Left            =   3270
      TabIndex        =   12
      Top             =   4380
      Width           =   2130
      _Version        =   196608
      _ExtentX        =   3757
      _ExtentY        =   644
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
   Begin EditLib.fpText txtWagesCentDep 
      Height          =   396
      Left            =   3936
      TabIndex        =   3
      ToolTipText     =   $"frmIntFaceMaint.frx":1586
      Top             =   2256
      Width           =   2172
      _Version        =   196608
      _ExtentX        =   3831
      _ExtentY        =   698
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
      Text            =   "0"
      CharValidationText=   "1 2 3 4 5 6 7 8 9 0 -"
      MaxLength       =   14
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
   Begin EditLib.fpText txtImprest 
      Height          =   396
      Left            =   3936
      TabIndex        =   2
      ToolTipText     =   "For imprest or central depository funds, Enter the account number credited for net pay."
      Top             =   1824
      Width           =   2172
      _Version        =   196608
      _ExtentX        =   3831
      _ExtentY        =   698
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
      Text            =   "0"
      CharValidationText=   "1 2 3 4 5 6 7 8 9 0 -"
      MaxLength       =   14
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
   Begin EditLib.fpText txtCashWagesDue 
      Height          =   396
      Left            =   3936
      TabIndex        =   1
      ToolTipText     =   $"frmIntFaceMaint.frx":163C
      Top             =   1392
      Width           =   2172
      _Version        =   196608
      _ExtentX        =   3831
      _ExtentY        =   698
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
      CharValidationText=   "1 2 3 4 5 6 7 8 9 0 -"
      MaxLength       =   14
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
   Begin EditLib.fpText txtExpAllocation 
      Height          =   396
      Left            =   9456
      TabIndex        =   11
      ToolTipText     =   "Input the method of calculating matching expenses."
      Top             =   3552
      Width           =   1068
      _Version        =   196608
      _ExtentX        =   1884
      _ExtentY        =   698
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
      AlignTextH      =   1
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
      CharValidationText=   "1 2 3 4 5 6 7 8 9 0 -"
      MaxLength       =   1
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
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
   Begin EditLib.fpText txtIndirectFundCredit 
      Height          =   365
      Left            =   3270
      TabIndex        =   19
      ToolTipText     =   "Enter the Indirect Account Credited."
      Top             =   7169
      Width           =   2130
      _Version        =   196608
      _ExtentX        =   3757
      _ExtentY        =   644
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
      CharValidationText=   "1 2 3 4 5 6 7 8 9 0 -"
      MaxLength       =   14
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
   Begin EditLib.fpText txtIndirectFundCash 
      Height          =   365
      Left            =   3270
      TabIndex        =   18
      ToolTipText     =   "Enter the Indirect Account Debited."
      Top             =   6765
      Width           =   2130
      _Version        =   196608
      _ExtentX        =   3757
      _ExtentY        =   644
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
      CharValidationText=   "1 2 3 4 5 6 7 8 9 0 -"
      MaxLength       =   14
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
   Begin EditLib.fpText txtIndirectExp 
      Height          =   365
      Left            =   4125
      TabIndex        =   17
      ToolTipText     =   "Enter the object code for Indirect Costs."
      Top             =   6356
      Width           =   1260
      _Version        =   196608
      _ExtentX        =   2222
      _ExtentY        =   644
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
      CharValidationText=   "1 2 3 4 5 6 7 8 9 0 -"
      MaxLength       =   7
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
   Begin EditLib.fpText txtFringeFundCredit 
      Height          =   365
      Left            =   3270
      TabIndex        =   15
      ToolTipText     =   "Enter the Fringe account credited."
      Top             =   5547
      Width           =   2130
      _Version        =   196608
      _ExtentX        =   3757
      _ExtentY        =   644
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
      CharValidationText=   "1 2 3 4 5 6 7 8 9 0 -"
      MaxLength       =   14
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
   Begin EditLib.fpText txtFringeFundCash 
      Height          =   365
      Left            =   3270
      TabIndex        =   14
      ToolTipText     =   "Enter the Fringe Account Debited."
      Top             =   5153
      Width           =   2130
      _Version        =   196608
      _ExtentX        =   3757
      _ExtentY        =   644
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
      CharValidationText=   "1 2 3 4 5 6 7 8 9 0 -"
      MaxLength       =   14
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
   Begin EditLib.fpText txtFringeExp 
      Height          =   365
      Left            =   4080
      TabIndex        =   13
      ToolTipText     =   "Enter the object code for Fringe Costs."
      Top             =   4769
      Width           =   1305
      _Version        =   196608
      _ExtentX        =   2302
      _ExtentY        =   644
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
      CharValidationText=   "1 2 3 4 5 6 7 8 9 0 -"
      MaxLength       =   7
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
   Begin EditLib.fpText txtGLTranAcctLength 
      Height          =   360
      Left            =   8790
      TabIndex        =   27
      ToolTipText     =   "Please don't adjust. This has been set by Southern Software."
      Top             =   7169
      Width           =   2130
      _Version        =   196608
      _ExtentX        =   3757
      _ExtentY        =   644
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
      CharValidationText=   "1 2 3 4 5 6 7 8 9 0 -"
      MaxLength       =   14
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
   Begin EditLib.fpText txtFundandAcctLength 
      Height          =   360
      Left            =   8790
      TabIndex        =   26
      ToolTipText     =   "Enter the total number of digits in your FUND and ACCOUNT code.  i.e. 10-420-02 would be 5. Omit dashes (xx-xxxx is 6)"
      Top             =   6765
      Width           =   2130
      _Version        =   196608
      _ExtentX        =   3757
      _ExtentY        =   644
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
      CharValidationText=   "1 2 3 4 5 6 7 8 9 0 -"
      MaxLength       =   14
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
   Begin EditLib.fpText txtRetirementLiab 
      Height          =   360
      Left            =   8790
      TabIndex        =   25
      ToolTipText     =   "Enter the Retirement Expense Liability Account Number."
      Top             =   6356
      Width           =   2130
      _Version        =   196608
      _ExtentX        =   3757
      _ExtentY        =   644
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
      CharValidationText=   "1 2 3 4 5 6 7 8 9 0 -"
      MaxLength       =   14
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
   Begin EditLib.fpText txtRetirementExp 
      Height          =   360
      Left            =   8790
      TabIndex        =   24
      ToolTipText     =   "Enter you retirement expense account code."
      Top             =   5956
      Width           =   2130
      _Version        =   196608
      _ExtentX        =   3757
      _ExtentY        =   644
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
      CharValidationText=   "1 2 3 4 5 6 7 8 9 0 -"
      MaxLength       =   14
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
   Begin EditLib.fpText txtMedicareLiab 
      Height          =   360
      Left            =   8790
      TabIndex        =   23
      ToolTipText     =   "Enter the Medicare Matching Expense Liability account number."
      Top             =   5547
      Width           =   2130
      _Version        =   196608
      _ExtentX        =   3757
      _ExtentY        =   644
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
      CharValidationText=   "1 2 3 4 5 6 7 8 9 0 -"
      MaxLength       =   14
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
   Begin EditLib.fpText txtMedicareExp 
      Height          =   360
      Left            =   8790
      TabIndex        =   22
      ToolTipText     =   "Enter the object code for Medicare expense."
      Top             =   5153
      Width           =   2130
      _Version        =   196608
      _ExtentX        =   3757
      _ExtentY        =   644
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
      CharValidationText=   "1 2 3 4 5 6 7 8 9 0 -"
      MaxLength       =   14
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
   Begin EditLib.fpText txtSocSecLiab 
      Height          =   360
      Left            =   8790
      TabIndex        =   21
      ToolTipText     =   "Enter the Social Security Matching Liability account number."
      Top             =   4769
      Width           =   2130
      _Version        =   196608
      _ExtentX        =   3757
      _ExtentY        =   644
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
      CharValidationText=   "1 2 3 4 5 6 7 8 9 0 -"
      MaxLength       =   14
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
   Begin EditLib.fpText txtSocSecExp 
      Height          =   365
      Left            =   8796
      TabIndex        =   20
      ToolTipText     =   "Enter the oject code for Social Security matching expenses."
      Top             =   4368
      Width           =   2136
      _Version        =   196608
      _ExtentX        =   3768
      _ExtentY        =   644
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
      CharValidationText=   "1 2 3 4 5 6 7 8 9 0 -"
      MaxLength       =   14
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
   Begin EditLib.fpText txtStateWHNum 
      Height          =   396
      Left            =   8928
      TabIndex        =   5
      ToolTipText     =   "Enter the liability account for State withholding."
      Top             =   984
      Width           =   2172
      _Version        =   196608
      _ExtentX        =   3831
      _ExtentY        =   698
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
      ControlType     =   0
      Text            =   ""
      CharValidationText=   "1 2 3 4 5 6 7 8 9 0 -"
      MaxLength       =   14
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
   Begin EditLib.fpText txtMedicareWHNum 
      Height          =   396
      Left            =   8928
      TabIndex        =   7
      ToolTipText     =   "Enter the liability account for Medicare withholding."
      Top             =   1836
      Width           =   2172
      _Version        =   196608
      _ExtentX        =   3831
      _ExtentY        =   698
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
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   14
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
   Begin EditLib.fpText txtRetireWHNum 
      Height          =   396
      Left            =   8928
      TabIndex        =   8
      ToolTipText     =   "Enter the liability account for Retirement withholding."
      Top             =   2268
      Width           =   2172
      _Version        =   196608
      _ExtentX        =   3831
      _ExtentY        =   698
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
      ControlType     =   0
      Text            =   ""
      CharValidationText=   "1 2 3 4 5 6 7 8 9 0 -"
      MaxLength       =   14
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
   Begin fpBtnAtlLibCtl.fpBtn cmdList 
      Height          =   495
      Left            =   5850
      TabIndex        =   60
      TabStop         =   0   'False
      ToolTipText     =   "Press to bring up a list of all General Ledger numbers."
      Top             =   7815
      Width           =   1485
      _Version        =   131072
      _ExtentX        =   2619
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
      ButtonDesigner  =   "frmIntFaceMaint.frx":16E3
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSave 
      Height          =   495
      Left            =   7695
      TabIndex        =   61
      TabStop         =   0   'False
      ToolTipText     =   "Press to commit this data to memory."
      Top             =   7813
      Width           =   1470
      _Version        =   131072
      _ExtentX        =   2593
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
      ButtonDesigner  =   "frmIntFaceMaint.frx":18C3
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   495
      Left            =   9540
      TabIndex        =   62
      TabStop         =   0   'False
      ToolTipText     =   "Press to bring up a list of all General Ledger numbers."
      Top             =   7813
      Width           =   1470
      _Version        =   131072
      _ExtentX        =   2593
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
      ButtonDesigner  =   "frmIntFaceMaint.frx":1A9F
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   2268
      Left            =   336
      Top             =   912
      Width           =   10980
   End
   Begin VB.Label Label33 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "GL Validation?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   240
      Left            =   672
      TabIndex        =   59
      Top             =   3264
      Width           =   1668
   End
   Begin VB.Label Label32 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Check Style"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   252
      Left            =   6216
      TabIndex        =   58
      Top             =   3264
      Width           =   1140
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Type 1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   360
      Left            =   2220
      TabIndex        =   49
      Top             =   4080
      Width           =   1650
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   630
      Index           =   1
      Left            =   1440
      Top             =   113
      Width           =   8655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Interface Maintenance"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   2880
      TabIndex        =   28
      Top             =   240
      Width           =   6012
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   720
      Left            =   1440
      Top             =   48
      Width           =   8652
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Exp Allocation Code"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   252
      Left            =   8928
      TabIndex        =   57
      Top             =   3264
      Width           =   1956
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Indirect Fund Cash Acct"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   270
      Left            =   615
      TabIndex        =   56
      Top             =   6850
      Width           =   2340
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Indirect Exp Acct     XX-XXX-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   270
      Left            =   1005
      TabIndex        =   55
      Top             =   6454
      Width           =   2970
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Indirect Cost Rate"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   270
      Left            =   615
      TabIndex        =   54
      Top             =   6034
      Width           =   2340
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fringe Fund Credit Acct"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   270
      Left            =   615
      TabIndex        =   53
      Top             =   5642
      Width           =   2340
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fringe Fund Cash Acct"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   270
      Left            =   615
      TabIndex        =   52
      Top             =   5254
      Width           =   2340
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fringe Exp Acct    XX-XXX-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   270
      Left            =   615
      TabIndex        =   51
      Top             =   4849
      Width           =   3345
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fringe Cost Rate"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   270
      Left            =   630
      TabIndex        =   50
      Top             =   4440
      Width           =   2340
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Indirect Fund Credit Acct"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   270
      Left            =   615
      TabIndex        =   48
      Top             =   7265
      Width           =   2340
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Type 2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   330
      Left            =   7830
      TabIndex        =   47
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   3570
      Left            =   6090
      Top             =   4080
      Width           =   5220
   End
   Begin VB.Label Label31 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fund && Acct Length"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   270
      Left            =   6285
      TabIndex        =   46
      Top             =   6850
      Width           =   2175
   End
   Begin VB.Label Label30 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Retirement Liab Acct"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   6270
      TabIndex        =   45
      Top             =   6454
      Width           =   2175
   End
   Begin VB.Label Label29 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Retirement Exp Acct"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   330
      Left            =   6270
      TabIndex        =   44
      Top             =   6034
      Width           =   2175
   End
   Begin VB.Label Label28 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Medicare Liab Acct"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   330
      Left            =   6285
      TabIndex        =   43
      Top             =   5642
      Width           =   2175
   End
   Begin VB.Label Label27 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Medicare Exp Acct"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   336
      Left            =   6624
      TabIndex        =   42
      Top             =   5254
      Width           =   1836
   End
   Begin VB.Label Label26 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Soc Sec Liab Acct"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   225
      Left            =   6810
      TabIndex        =   41
      Top             =   4849
      Width           =   1650
   End
   Begin VB.Label Label25 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Soc Sec Exp Acct"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   270
      Left            =   6810
      TabIndex        =   40
      Top             =   4459
      Width           =   1650
   End
   Begin VB.Label Label24 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "GL Tran Acct Len"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   6270
      TabIndex        =   39
      Top             =   7265
      Width           =   2175
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Split Deductions?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   240
      Left            =   2880
      TabIndex        =   38
      Top             =   3264
      Width           =   1668
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   828
      Left            =   336
      Top             =   3216
      Width           =   10980
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Retirement W/H Acct No"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   276
      Left            =   6324
      TabIndex        =   37
      Top             =   2328
      Width           =   2412
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Medicare W/H Acct No"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   324
      Left            =   6660
      TabIndex        =   36
      Top             =   1896
      Width           =   2076
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "SocSec W/H Acct No"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   6564
      TabIndex        =   35
      Top             =   1488
      Width           =   2172
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "State W/H Acct No"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   276
      Left            =   6420
      TabIndex        =   34
      Top             =   1104
      Width           =   2292
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Interface Code"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   276
      Left            =   1920
      TabIndex        =   33
      Top             =   1104
      Width           =   1692
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cash/Wages Pay/Due to Cent Dep"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   480
      TabIndex        =   32
      Top             =   1488
      Width           =   3132
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Imprest/Cent Dep Cash Acct (Cr.)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   324
      Left            =   432
      TabIndex        =   31
      Top             =   1896
      Width           =   3228
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Wages Pay/Cent Dep Due from (Dr.)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   276
      Left            =   240
      TabIndex        =   30
      Top             =   2328
      Width           =   3444
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Federal W/H Acct No"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   276
      Left            =   1104
      TabIndex        =   29
      Top             =   2712
      Width           =   2532
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   3570
      Left            =   330
      Top             =   4080
      Width           =   5415
   End
   Begin VB.Menu mnuoptions 
      Caption         =   "Options"
      Begin VB.Menu mnuPrnScn 
         Caption         =   "Prin&t Screen"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmIntFaceMaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Dim changeFlag As Boolean
Dim GLTranNum As Long
Dim SysFileLen As Integer
Dim BadGLNum As Boolean
Dim OverrideFlag As Boolean
Dim PICType As String
Dim GLCheckFlag As Boolean

Private Sub cmdList_Click()
'  If QPTrim$(txtCitipakDir.Text) = "" Then
'    MsgBox "No Citipak directory path has been saved."
'    Exit Sub
'  Else
    frmGLPickList.Show vbModal
'  End If
End Sub

Private Sub cmdSave_Click()
   Dim SysHandle As Integer
   Dim SysFileRec As RegDSysFileRecType
   Dim FileSize As Long
   Dim tempICode As String
   Dim tempSocLiab As String, tempSocExp As String, tempMedicareExp As String
   Dim tempMedicareLiab As String, tempRetireExp As String, tempRetireLiab As String, tempFringeCost As Double, tempFringeExp As String
   Dim tempIndirectCost As Double, tempIndirectExp As String, tempFringeCredit As String, tempFringeCash As String
   Dim TempESCType As Integer, TaxBase As Double
   Dim tempIndirectCredit As String, tempIndirectCash As String, tempCitiDir As String, tempCashAcct As String
   Dim tempSplitFlag As String, tempExpAlloc As String, tempFedWH As String
   Dim tempStateWH As String, tempSocSecWH As String
   Dim tempMedicareWH As String, tempRetireWH As String
   Dim tempWagesDueFrom As String, tempImprestCash As String
   Dim tempAcctCnt As Integer, tempGLAcctLen As Integer
   Dim tempCheckStyle As Integer
   Dim tempGLCheckYN As String * 1
   Dim FundLength As Integer, AcctLength As Integer, DetLength As Integer '8/5
'   Dim tempVAC2SICK As String * 1
   
   'do we want to check for valid GL nums?
   If GLCheckFlag = True And OverrideFlag = False Then
'     If QPTrim$(txtCitipakDir.Text) <> "" Then
       Call CheckForValidWHNum
'     End If
   End If

   OverrideFlag = False
   If BadGLNum = True Then
     BadGLNum = False
     Exit Sub
   End If
   
   tempICode = comboInterface.Text
   If tempICode = "" Then
      MsgBox "Please enter an Interface Code"
      comboInterface.SetFocus
      GoTo BadUnitData
   End If
   
   tempCashAcct = txtCashWagesDue.Text
   If tempCashAcct = "" Then
     MsgBox "Please enter a Cash/WagesPay/Due To Central Deposit Amount"
     txtCashWagesDue.SetFocus
     GoTo BadUnitData
   End If
   
   If PICType = "C" Then
     tempImprestCash = QPTrim$(txtImprest.Text)
     If tempImprestCash = "" Then
       MsgBox "Please enter a Imprest/Cent Dep Cash Acct (Cr.) Value"
       txtImprest.SetFocus
       GoTo BadUnitData
     End If
   
     If Val(ReplaceString(txtWagesCentDep.Text, "-", "")) <= 0 Then
       MsgBox "Please enter a valid number in the Wages Pay/Cent Dep Due from (Dr.) field"
       txtWagesCentDep.SetFocus
       Exit Sub
     End If
   
     tempWagesDueFrom = QPTrim$(txtWagesCentDep.Text)
     If tempWagesDueFrom = "" Then
       MsgBox "Please enter a Wages Pay/Cent Dep Due from (Dr.) Value"
       txtWagesCentDep.SetFocus
       GoTo BadUnitData
     End If
   Else
     tempImprestCash = QPTrim$(txtImprest.Text)
     tempWagesDueFrom = QPTrim$(txtWagesCentDep.Text)
   End If
   
   
   tempFedWH = QPTrim$(txtFedWHNum.Text)
   If tempFedWH = "" Then
     MsgBox "Please enter a Federal Withholding Account Number"
     txtFedWHNum.SetFocus
     GoTo BadUnitData
   End If
   tempStateWH = QPTrim$(txtStateWHNum.Text)
   If tempStateWH = "" Then
     MsgBox "Please enter a State Withholding Account Number"
     txtStateWHNum.SetFocus
     GoTo BadUnitData
   End If
   
   tempSocSecWH = QPTrim$(txtSocSecWHNum.Text)
   If tempSocSecWH = "" Then
     MsgBox "Please enter a Social Security Withholding Account Number"
     txtSocSecWHNum.SetFocus
     GoTo BadUnitData
   End If
   tempMedicareWH = QPTrim$(txtMedicareWHNum.Text)
   If tempMedicareWH = "" Then
     MsgBox "Please enter a Medicare Withholding Account Number"
     txtMedicareWHNum.SetFocus
     GoTo BadUnitData
   End If
   
   tempRetireWH = QPTrim$(txtRetireWHNum.Text)
   If tempRetireWH = "" Then
     MsgBox "Please enter a Retirement Withholding Account Number"
     txtRetireWHNum.SetFocus
     GoTo BadUnitData
   End If
   
'   If QPTrim$(txtCitipakDir.Text) = "" Then GoTo NoGLFiles '11/27/02
'   tempCitiDir = QPTrim$(txtCitipakDir.Text)
'   If tempCitiDir = "0" Or Len(tempCitiDir) = 0 Then
'     MsgBox "Please enter a CitiPak Working Directory"
'     txtCitipakDir.SetFocus
'     GoTo BadUnitData
'   End If
NoGLFiles:
   tempSplitFlag = comboSplitDeDuc.Text
   If tempSplitFlag = "0" Or Len(tempSplitFlag) = 0 Then
     MsgBox "Please indicate if you want to Split Deductions"
     comboSplitDeDuc.SetFocus
     GoTo BadUnitData
   End If
   tempExpAlloc = QPTrim$(txtExpAllocation.Text)
   If tempExpAlloc = "0" Or Len(tempExpAlloc) = 0 Then
     MsgBox "Please enter an Expense Allocation Code"
     txtExpAllocation.SetFocus
     GoTo BadUnitData
   End If
   
'   tempVAC2SICK = Mid(fpcomboVAC2SICK.Text, 3, 1)
   
   tempCheckStyle = Val(Mid(fpcomboChkStyle.Text, 3, 1))
   If tempCheckStyle = 0 Then
     MsgBox "Please enter a Check Style"
     fpcomboChkStyle.SetFocus
     GoTo BadUnitData
   ElseIf tempCheckStyle = 6 Then
     MsgBox "Check style #6 is not supported at this time. Please select another style."
     fpcomboChkStyle.SetFocus
     GoTo BadUnitData
   End If
   
   tempFringeCost = Val(QPTrim$(fptxtFringeCost.Text))
   If Len(tempFringeCost) = 0 Then
     MsgBox "Please enter a Fringe Cost Rate"
     fptxtFringeCost.SetFocus
     GoTo BadUnitData
   End If
   tempFringeExp = QPTrim$(txtFringeExp.Text)
   If tempFringeExp = "" Then
     MsgBox "Please enter a Fringe Expense Account Number"
     txtFringeExp.SetFocus
     GoTo BadUnitData
   End If
   
   tempFringeCash = QPTrim$(txtFringeFundCash.Text)
   If tempFringeCash = "" Then
     MsgBox "Please enter a Fringe Fund Cash Account Number"
     txtFringeFundCash.SetFocus
     GoTo BadUnitData
   End If
   tempFringeCredit = QPTrim$(txtFringeFundCredit.Text)
   If tempFringeCredit = "" Then
     MsgBox "Please enter a Fringe Fund Credit Account Number"
     txtFringeFundCredit.SetFocus
     GoTo BadUnitData
   End If
    
   tempIndirectCost = Val(QPTrim$(fptxtIndirectCost.Text))
   If Len(tempIndirectCost) = 0 Then
     MsgBox "Please enter an Indirect Cost Rate"
     fptxtIndirectCost.SetFocus
     GoTo BadUnitData
   End If
   tempIndirectExp = QPTrim$(txtIndirectExp.Text)
   If tempIndirectExp = "" Then
     MsgBox "Please enter a Indirect Expense Account Number"
     txtIndirectExp.SetFocus
     GoTo BadUnitData
   End If
    
   tempIndirectCash = QPTrim$(txtIndirectFundCash.Text)
   If tempIndirectCash = "" Then
     MsgBox "Please enter an Indirect Fund Cash Account Number"
     txtIndirectFundCash.SetFocus
     GoTo BadUnitData
   End If
   tempIndirectCredit = QPTrim$(txtIndirectFundCredit.Text)
   If tempIndirectCredit = "" Then
     MsgBox "Please enter a Fringe Fund Credit Account Number"
     txtIndirectFundCredit.SetFocus
     GoTo BadUnitData
   End If
   
  tempSocExp = QPTrim$(txtSocSecExp.Text)
   If tempSocExp = "" Then
     MsgBox "Please enter a Social Security Expense Account Number"
     txtSocSecExp.SetFocus
     GoTo BadUnitData
   End If
   tempSocLiab = QPTrim$(txtSocSecLiab.Text)
   If tempSocLiab = "" Then
     MsgBox "Please enter a Social Security Liability Account Number"
     txtSocSecLiab.SetFocus
     GoTo BadUnitData
   End If
   
   tempMedicareExp = QPTrim$(txtMedicareExp.Text)
   If tempMedicareExp = "" Then
     MsgBox "Please enter a Medicare Expense Account Number"
     txtMedicareExp.SetFocus
     GoTo BadUnitData
   End If
   tempMedicareLiab = QPTrim$(txtMedicareLiab.Text)
   If tempMedicareLiab = "" Then
     MsgBox "Please enter a Medicare Liability Account Number"
     txtMedicareLiab.SetFocus
     GoTo BadUnitData
   End If
   
   tempRetireExp = QPTrim$(txtRetirementExp.Text)
   If tempRetireExp = "" Then
     MsgBox "Please enter a Retirement Expense Account Number"
     txtRetirementExp.SetFocus
     GoTo BadUnitData
   End If
   tempRetireLiab = QPTrim$(txtRetirementLiab.Text)
   If tempRetireLiab = "" Then
     MsgBox "Please enter a Retirement Liability Account Number"
     txtRetirementLiab.SetFocus
     GoTo BadUnitData
   End If

'   Call GetAcctStruct(txtCitipakDir.Text, FundLength, AcctLength, DetLength) '8/5
   Call GetAcctStruct(CurrCitiPath, FundLength, AcctLength, DetLength)
   If FundLength + AcctLength <> Val(txtFundandAcctLength.Text) Then
     If MsgBox("The fund length plus the account length do not equal the value entered for 'Fund & Acct Length'. Do you want to save anyway?", vbYesNo) = vbNo Then
       Close
       txtFundandAcctLength.SetFocus
       Exit Sub
     End If
   End If
'   txtFundandAcctLength.Text = FundLength + AcctLength '8/5
   If (FundLength + AcctLength + DetLength) <> Val(txtGLTranAcctLength.Text) Then
     If MsgBox("The fund length plus the account length plus the detail length do not equal the value entered for 'GL Tran Acct Len'. Do you want to save it anyway?", vbYesNo) = vbNo Then
       Close
       txtGLTranAcctLength.SetFocus
       Exit Sub
     End If
   End If
'   txtGLTranAcctLength.Text = FundLength + AcctLength + DetLength '8/5
   tempAcctCnt = Val(QPTrim$(txtFundandAcctLength.Text))
   If tempAcctCnt = 0 Then
     MsgBox "Please enter a Fund and Account Length Number."
     txtFundandAcctLength.SetFocus
     GoTo BadUnitData
   End If
   
   tempGLAcctLen = Val(QPTrim$(txtGLTranAcctLength.Text))
   If tempGLAcctLen = 0 Then
     MsgBox "Please enter a GL Tran Acct Length Number."
     txtGLTranAcctLength.SetFocus
     GoTo BadUnitData
   End If
   
   tempGLCheckYN = QPTrim$(fpcomboGLCheckYN.Text)
   
   SysFileRec.USEIMP = tempICode
   SysFileRec.CashAcct = tempCashAcct
   SysFileRec.IDRACCT = tempImprestCash
   SysFileRec.ICRACCT = tempWagesDueFrom
   
   SysFileRec.Liab(1).Acct = tempFedWH
   SysFileRec.Liab(2).Acct = tempStateWH
   SysFileRec.Liab(3).Acct = tempSocSecWH
   SysFileRec.Liab(4).Acct = tempMedicareWH
   SysFileRec.Liab(5).Acct = tempRetireWH
   
   SysFileRec.CITIDIR = "" 'tempCitiDir
   SysFileRec.SplitFlag = tempSplitFlag
   SysFileRec.EXPMETHD = tempExpAlloc
   SysFileRec.FRNGRATE = tempFringeCost
   SysFileRec.FRNGEXP = tempFringeExp
   
   SysFileRec.FRNGDR = tempFringeCash
   SysFileRec.FRNGCR = tempFringeCredit
   SysFileRec.INDRATE = tempIndirectCost
   SysFileRec.INDEXP = tempIndirectExp
   SysFileRec.INDDR = tempIndirectCash
   
   SysFileRec.INDCR = tempIndirectCredit
   SysFileRec.SOCEXP = tempSocExp
   SysFileRec.SOCLIAB = tempSocLiab
   SysFileRec.MEDEXP = tempMedicareExp
   SysFileRec.MEDLIAB = tempMedicareLiab
   
   SysFileRec.RETEXP = tempRetireExp
   SysFileRec.RETLIAB = tempRetireLiab
   SysFileRec.AcctCnt = tempAcctCnt
   SysFileRec.GLActLen = tempGLAcctLen
   SysFileRec.CheckStyle = tempCheckStyle
   SysFileRec.GLCheckYN = tempGLCheckYN
'   If QPTrim$(txtCitipakDir.Text) = "" Then SysFileRec.GLCheckYN = "N"
'   SysFileRec.VAC2SICK = tempVAC2SICK
   'save to file
   OpenSysFile SysHandle
   Put SysHandle, 1, SysFileRec
   Close SysHandle

   MsgBox "Your Information has been saved.", vbOKOnly
   frmControlFileMaint.Show
   DoEvents
   Unload frmIntFaceMaint
BadUnitData:
  MainLog ("System Interface data was saved.")
End Sub

Private Sub comboInterface_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    comboInterface.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    comboInterface.ListIndex = -1
  End If
  If comboInterface.ListDown = False Then
    If KeyCode = vbKeyDown Then
      SendKeys "{Tab}"
      KeyCode = 0
    ElseIf KeyCode = vbKeyUp Then
      SendKeys "+{Tab}"
      KeyCode = 0
    End If
  End If
End Sub

Private Sub comboInterface_LostFocus()
  comboInterface.Action = ActionClearSearchBuffer
  If QPTrim$(comboInterface.Text) <> QPTrim$(PICType) Then
    PICType = QPTrim$(comboInterface.Text)
    If QPTrim$(comboInterface.Text) = "P" Then
      txtImprest.Enabled = False
      txtWagesCentDep.Enabled = False
    ElseIf QPTrim$(comboInterface.Text) = "C" Or QPTrim$(comboInterface.Text) = "I" Then
      txtImprest.Enabled = True
      txtWagesCentDep.Enabled = True
    End If
  End If
End Sub

Private Sub comboSplitDeDuc_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    comboSplitDeDuc.ListDown = True
  End If
  If comboSplitDeDuc.ListDown = False Then
    If KeyCode = vbKeyDown Then
      SendKeys "{Tab}"
      KeyCode = 0
    ElseIf KeyCode = vbKeyUp Then
      SendKeys "+{Tab}"
      KeyCode = 0
    End If
  End If
End Sub

Private Sub comboSplitDeDuc_LostFocus()
  comboSplitDeDuc.Action = ActionClearSearchBuffer
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
      Call cmdExit_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%S"
      Call cmdSave_Click
      KeyCode = 0
    Case vbKeyF12:
      SendKeys "%G"
      Call cmdList_Click
      KeyCode = 0
    Case Else:
  End Select
  
End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  OverrideFlag = False
  LoadIFMFile
  MainLog ("System Interface Maintenance screen accessed.")
  Me.HelpContextID = hlpSystemFile
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If

End Sub

Private Sub cmdExit_Click()
'   On Error Resume Next
   changeFlag = False
   Dim SysHandle As Integer
   Dim DoWhatFlag As SaveChangeOptions1
   Dim SysFileRec As RegDSysFileRecType
   Dim SysFileLen As Integer
   
   'do we want to check for valid GL nums?
   If GLCheckFlag = True And OverrideFlag = False Then
     Call CheckForValidWHNum
   End If
   OverrideFlag = False
   If BadGLNum = True Then
     BadGLNum = False
     Exit Sub
   End If
   OpenSysFile SysHandle
   SysFileLen = LOF(SysHandle) \ Len(SysFileRec)
   
   Get SysHandle, 1, SysFileRec
   Close SysHandle

   If QPTrim$(SysFileRec.USEIMP) <> QPTrim$(comboInterface.Text) Then
      If QPTrim$(comboInterface.Text) = "0" And SysFileLen = 0 Then GoTo Next1
      changeFlag = True
      comboInterface.SetFocus
   End If
Next1:
   If QPTrim$(SysFileRec.CashAcct) <> QPTrim$(txtCashWagesDue.Text) Then
      If QPTrim$(txtCashWagesDue.Text) = "0" And SysFileLen = 0 Then GoTo Next2
      changeFlag = True
      txtCashWagesDue.SetFocus
   End If
Next2:
   If QPTrim$(PICType) <> "C" Then GoTo Next4
   If QPTrim$(SysFileRec.IDRACCT) <> QPTrim$(txtImprest.Text) Then
      If QPTrim$(txtImprest.Text) = "0" And SysFileLen = 0 Then GoTo Next3
      changeFlag = True
      txtImprest.SetFocus
   End If
Next3:
   If QPTrim$(SysFileRec.ICRACCT) <> QPTrim$(txtWagesCentDep.Text) Then
      If QPTrim$(txtWagesCentDep.Text) = "0" And SysFileLen = 0 Then GoTo Next4
      changeFlag = True
      txtWagesCentDep.SetFocus
   End If
Next4:
   If QPTrim$(ReplaceString(SysFileRec.Liab(1).Acct, "-", "")) <> QPTrim$(ReplaceString(txtFedWHNum.Text, "-", "")) Then
      If QPTrim$(txtFedWHNum.Text) = "0" And SysFileLen = 0 Then GoTo Next5
      changeFlag = True
      txtFedWHNum.SetFocus
   End If
Next5:
   If QPTrim$(ReplaceString(SysFileRec.Liab(2).Acct, "-", "")) <> QPTrim$(ReplaceString(txtStateWHNum.Text, "-", "")) Then
      If QPTrim$(txtStateWHNum.Text) = "0" And SysFileLen = 0 Then GoTo Next6
      changeFlag = True
      txtStateWHNum.SetFocus
   End If
Next6:
   If QPTrim$(ReplaceString(SysFileRec.Liab(3).Acct, "-", "")) <> QPTrim$(ReplaceString(txtSocSecWHNum.Text, "-", "")) Then
      If QPTrim$(txtSocSecWHNum.Text) = "0" And SysFileLen = 0 Then GoTo Next7
      changeFlag = True
      txtSocSecWHNum.SetFocus
   End If
Next7:
   If QPTrim$(ReplaceString(SysFileRec.Liab(4).Acct, "-", "")) <> QPTrim$(ReplaceString(txtMedicareWHNum.Text, "-", "")) Then
      If QPTrim$(txtMedicareWHNum.Text) = "0" And SysFileLen = 0 Then GoTo Next8
      changeFlag = True
      txtMedicareWHNum.SetFocus
   End If
Next8:
   If QPTrim$(ReplaceString(SysFileRec.Liab(5).Acct, "-", "")) <> QPTrim$(ReplaceString(txtRetireWHNum.Text, "-", "")) Then
      If QPTrim$(txtRetireWHNum.Text) = "0" And SysFileLen = 0 Then GoTo Next9
      changeFlag = True
      txtRetireWHNum.SetFocus
   End If
Next9:
'   If QPTrim$(SysFileRec.CITIDIR) <> QPTrim$(txtCitipakDir.Text) Then
'      If QPTrim$(txtCitipakDir.Text) = "0" And SysFileLen = 0 Then GoTo Next10
'      changeFlag = True
'      txtCitipakDir.SetFocus
'   End If
Next10:
   If QPTrim$(SysFileRec.SplitFlag) <> QPTrim$(comboSplitDeDuc.Text) Then
      If QPTrim$(comboSplitDeDuc.Text) = "0" And SysFileLen = 0 Then GoTo Next11
      changeFlag = True
      comboSplitDeDuc.SetFocus
   End If
Next11:
   If QPTrim$(SysFileRec.EXPMETHD) <> QPTrim$(txtExpAllocation.Text) Then
      If QPTrim$(txtExpAllocation.Text) = "0" And SysFileLen = 0 Then GoTo Next12
      changeFlag = True
      txtExpAllocation.SetFocus
   End If
Next12:
   If SysFileRec.FRNGRATE <> fptxtFringeCost.Text Then
      If QPTrim$(fptxtFringeCost.Text) = "0" And SysFileLen = 0 Then GoTo Next13
      changeFlag = True
      fptxtFringeCost.SetFocus
   End If
Next13:
   If QPTrim$(SysFileRec.FRNGEXP) <> QPTrim$(txtFringeExp.Text) Then
      If QPTrim$(txtFringeExp.Text) = "0" And SysFileLen = 0 Then GoTo Next14
      changeFlag = True
      txtFringeExp.SetFocus
   End If
Next14:
   If QPTrim$(SysFileRec.FRNGDR) <> QPTrim$(txtFringeFundCash.Text) Then
      If QPTrim$(txtFringeFundCash.Text) = "0" And SysFileLen = 0 Then GoTo Next15
      changeFlag = True
      txtFringeFundCash.SetFocus
   End If
Next15:
   If QPTrim$(SysFileRec.FRNGCR) <> QPTrim$(txtFringeFundCredit.Text) Then
      If QPTrim$(txtFringeFundCredit.Text) = "0" And SysFileLen = 0 Then GoTo Next16
      changeFlag = True
      txtFringeFundCredit.SetFocus
   End If
Next16:
   If SysFileRec.INDRATE <> fptxtIndirectCost.Text Then
      If QPTrim$(fptxtIndirectCost.Text) = "0" And SysFileLen = 0 Then GoTo Next17
      changeFlag = True
      fptxtIndirectCost.SetFocus
   End If
Next17:
   If QPTrim$(SysFileRec.INDEXP) <> QPTrim$(txtIndirectExp.Text) Then
      If QPTrim$(txtIndirectExp.Text) = "0" And SysFileLen = 0 Then GoTo Next18
      changeFlag = True
      txtIndirectExp.SetFocus
   End If
Next18:
   If QPTrim$(SysFileRec.INDDR) <> QPTrim$(txtIndirectFundCash.Text) Then
      If QPTrim$(txtIndirectFundCash.Text) = "0" And SysFileLen = 0 Then GoTo Next19
      changeFlag = True
      txtIndirectFundCash.SetFocus
   End If
Next19:
   If QPTrim$(SysFileRec.INDCR) <> QPTrim$(txtIndirectFundCredit.Text) Then
      If QPTrim$(txtIndirectFundCredit.Text) = "0" And SysFileLen = 0 Then GoTo Next20
      changeFlag = True
      txtIndirectFundCredit.SetFocus
   End If
Next20:
   If QPTrim$(SysFileRec.SOCEXP) <> QPTrim$(txtSocSecExp.Text) Then
      If QPTrim$(txtSocSecExp.Text) = "0" And SysFileLen = 0 Then GoTo Next21
      changeFlag = True
      txtSocSecExp.SetFocus
   End If
Next21:
   If QPTrim$(SysFileRec.SOCLIAB) <> QPTrim$(txtSocSecLiab.Text) Then
      If QPTrim$(txtSocSecLiab.Text) = "0" And SysFileLen = 0 Then GoTo Next22
      changeFlag = True
      txtSocSecLiab.SetFocus
   End If
Next22:
   If QPTrim$(SysFileRec.MEDEXP) <> QPTrim$(txtMedicareExp.Text) Then
      If QPTrim$(txtMedicareExp.Text) = "0" And SysFileLen = 0 Then GoTo Next23
      changeFlag = True
      txtMedicareExp.SetFocus
   End If
Next23:
   If QPTrim$(SysFileRec.MEDLIAB) <> QPTrim$(txtMedicareLiab.Text) Then
      If QPTrim$(txtMedicareLiab.Text) = "0" And SysFileLen = 0 Then GoTo Next24
      changeFlag = True
      txtMedicareLiab.SetFocus
   End If
Next24:
   If QPTrim$(SysFileRec.RETEXP) <> QPTrim$(txtRetirementExp.Text) Then
      If QPTrim$(txtRetirementExp.Text) = "0" And SysFileLen = 0 Then GoTo Next25
      changeFlag = True
      txtRetirementExp.SetFocus
   End If
Next25:
   If QPTrim$(SysFileRec.RETLIAB) <> QPTrim$(txtRetirementLiab.Text) Then
      If QPTrim$(txtRetirementLiab.Text) = "0" And SysFileLen = 0 Then GoTo Next26
      changeFlag = True
      txtRetirementLiab.SetFocus
   End If
Next26:
   If SysFileRec.AcctCnt <> Val(txtFundandAcctLength.Text) Then
      If QPTrim$(txtFundandAcctLength.Text) = "0" And SysFileLen = 0 Then GoTo Next27
      changeFlag = True
      txtFundandAcctLength.SetFocus
   End If
Next27:
   If QPTrim$(SysFileRec.GLCheckYN) <> QPTrim$(fpcomboGLCheckYN.Text) Then
      If QPTrim$(fpcomboGLCheckYN.Text) = "0" And SysFileLen = 0 Then GoTo Next28
      changeFlag = True
      fpcomboGLCheckYN.SetFocus
   End If
Next28:
   If Mid(fpcomboChkStyle.Text, 3, 1) <> SysFileRec.CheckStyle Then
     If QPTrim$(fpcomboChkStyle.Text) = "" And SysFileRec.CheckStyle = 0 Then GoTo Next29
     If QPTrim$(fpcomboChkStyle.Text) = "" And SysFileLen = 0 Then GoTo Next29
     changeFlag = True
     fpcomboChkStyle.SetFocus
   End If
Next29:
   If SysFileRec.GLActLen <> Val(txtGLTranAcctLength.Text) Then
      If QPTrim$(txtGLTranAcctLength.Text) = "0" And SysFileLen = 0 Then GoTo Next30
      changeFlag = True
      txtGLTranAcctLength.SetFocus
   End If
Next30:
'   If Mid(fpcomboVAC2SICK.Text, 3, 1) <> SysFileRec.VAC2SICK Then
'     changeFlag = True
'     fpcomboVAC2SICK.SetFocus
'   End If
   
   If changeFlag = False Then 'no changes detected
      frmControlFileMaint.Show
      DoEvents
      Unload frmIntFaceMaint
      GoTo endClick
   Else
      DoWhatFlag = PromptSaveChanges(Me)
      Select Case DoWhatFlag
      Case SaveChangeOptions1.scoSaveChanges 'save changes
        Call cmdSave_Click
      Case SaveChangeOptions1.scoReviewChanges 'review is just bringing back the current form
      Case SaveChangeOptions1.scoAbandonChanges 'abandon
        frmControlFileMaint.Show
        DoEvents
        Unload frmIntFaceMaint
      Case Else:
        'Do nothing because we don't know about any options except
        'save, review or abandon
      End Select
         
   End If
endClick:
  MainLog ("System Interface Maintenance screen exited.")
End Sub

Private Sub LoadIFMFile()
'   On Error Resume Next
   Dim SysHandle As Integer
   Dim SysFileRec As RegDSysFileRecType
   Dim FileSize As Long
   Dim FundLength As Integer, AcctLength As Integer, DetLength As Integer
   comboInterface.AddItem "P"
   comboInterface.AddItem "I"
   comboInterface.AddItem "C"
   comboSplitDeDuc.AddItem "Y"
   comboSplitDeDuc.AddItem "N"
   fpcomboGLCheckYN.AddItem "Y"
   fpcomboGLCheckYN.AddItem "N"
   
   OpenSysFile SysHandle
   FileSize = LOF(SysHandle) / Len(SysFileRec)
   SysFileLen = FileSize
   If FileSize = 0 Then
      'file is zero bytes
     Close SysHandle
     GLTranNum = 0
     comboInterface.Text = "0"
     txtCashWagesDue.Text = "0"
     txtImprest.Text = "0"
     txtWagesCentDep.Text = "0"
     txtFedWHNum.Text = "0"
     txtStateWHNum.Text = "0"
     txtSocSecWHNum.Text = "0"
     txtMedicareWHNum.Text = "0"
     txtRetireWHNum.Text = "0"
     
'     txtCitipakDir.Text = ""
     fpcomboGLCheckYN.Text = "Y"
     GLCheckFlag = True
     comboSplitDeDuc.Text = "N"
     txtExpAllocation.Text = "0"
     fptxtFringeCost.Text = "0"
     txtFringeExp.Text = "0"
     txtFringeFundCash.Text = "0"
     txtFringeFundCredit.Text = "0"
     fptxtIndirectCost.Text = "0"
     txtIndirectExp.Text = "0"
     txtIndirectFundCash.Text = "0"
     txtIndirectFundCredit.Text = "0"
     txtSocSecExp.Text = "0"
     txtSocSecLiab.Text = "0"
     txtMedicareExp.Text = "0"
     txtMedicareLiab.Text = "0"
     txtRetirementExp.Text = "0"
     txtRetirementLiab.Text = "0"
     txtFundandAcctLength.Text = "0"
     txtGLTranAcctLength.Text = "0"
     fpcomboChkStyle.InsertRow = "  1  Blank Top Stub 39 Line" '9013-39
     fpcomboChkStyle.InsertRow = "  2  Blank Top Stub 42 Line" '9013-42 Carthage.bi
     fpcomboChkStyle.InsertRow = "  3  Product Code 9028" '9028 Stdchk.bi
     fpcomboChkStyle.InsertRow = "  4  Product Code 9007"
     fpcomboChkStyle.InsertRow = "  5  Top and bottom stub laser"
     fpcomboChkStyle.InsertRow = "  6  Middle and bottom stub laser"
     fpcomboChkStyle.InsertRow = "  7  Custom 42 Line"
'     fpcomboVAC2SICK.Text = "N"
'     fpcomboVAC2SICK.InsertRow = " Y  Add vac over max to sick"
'     fpcomboVAC2SICK.InsertRow = " N  Don't add vac over max to sick"
     GoTo NoUnitFileYet
   Else
     Get SysHandle, 1, SysFileRec
     Close SysHandle
   End If
   GLTranNum = SysFileRec.GLActLen
   
   If SysFileRec.CheckStyle = 1 Then
     fpcomboChkStyle.Text = "  1  Blank Top Stub 39 Line"
   ElseIf SysFileRec.CheckStyle = 2 Then
     fpcomboChkStyle.Text = "  2  Blank Top Stub 42 Line"
   ElseIf SysFileRec.CheckStyle = 3 Then
     fpcomboChkStyle.Text = "  3  Product Code 9028"
   ElseIf SysFileRec.CheckStyle = 4 Then
     fpcomboChkStyle.Text = "  4  Product Code 9007"
   ElseIf SysFileRec.CheckStyle = 5 Then
     fpcomboChkStyle.Text = "  5  Top and bottom stub laser"
   ElseIf SysFileRec.CheckStyle = 6 Then
     fpcomboChkStyle.Text = "  6  Middle and bottom stub laser"
   ElseIf SysFileRec.CheckStyle = 7 Then
     fpcomboChkStyle.Text = "  7  Custom 42 Line"
   Else
     fpcomboChkStyle.Text = ""
   End If
   
   fpcomboChkStyle.InsertRow = "  1  Blank Top Stub 39 Line"
   fpcomboChkStyle.InsertRow = "  2  Blank Top Stub 42 Line"
   fpcomboChkStyle.InsertRow = "  3  Product Code 9028"
   fpcomboChkStyle.InsertRow = "  4  Product Code 9007"
   fpcomboChkStyle.InsertRow = "  5  Top and bottom stub laser"
   fpcomboChkStyle.InsertRow = "  6  Middle and bottom stub laser"
   fpcomboChkStyle.InsertRow = "  7  Custom 42 Line"
   
'   If SysFileRec.VAC2SICK = "N" Then fpcomboVAC2SICK.Text = "  " & SysFileRec.VAC2SICK & "  Don't add vac over max to sick"
'   If SysFileRec.VAC2SICK = "Y" Then fpcomboVAC2SICK.Text = "  " & SysFileRec.VAC2SICK & "  Add vac over max to sick"
'
'   fpcomboVAC2SICK.InsertRow = "  Y  Add vac over max to sick"
'   fpcomboVAC2SICK.InsertRow = "  N  Don't add vac over max to sick"
   
   comboInterface.Text = QPTrim$(SysFileRec.USEIMP)
   PICType = comboInterface.Text
   txtCashWagesDue.Text = QPTrim$(SysFileRec.CashAcct)
   If QPTrim$(comboInterface.Text) = "P" Then
     txtImprest.Enabled = False
     txtWagesCentDep.Enabled = False
     GoTo PIsOn
   End If
   txtImprest.Text = QPTrim$(SysFileRec.IDRACCT)
   txtWagesCentDep.Text = QPTrim$(SysFileRec.ICRACCT)
PIsOn:
   txtFedWHNum.Text = QPTrim$(SysFileRec.Liab(1).Acct)
   txtStateWHNum.Text = QPTrim$(SysFileRec.Liab(2).Acct)
   txtSocSecWHNum.Text = QPTrim$(SysFileRec.Liab(3).Acct)
   txtMedicareWHNum.Text = QPTrim$(SysFileRec.Liab(4).Acct)
   txtRetireWHNum.Text = QPTrim$(SysFileRec.Liab(5).Acct)
'   txtCitipakDir.Text = QPTrim$(SysFileRec.CITIDIR)
'   CurrCitiPath = QPTrim$(txtCitipakDir.Text)
   fpcomboGLCheckYN.Text = QPTrim$(SysFileRec.GLCheckYN)
   If QPTrim$(fpcomboGLCheckYN.Text) = "N" Then
     GLCheckFlag = False
   Else
     GLCheckFlag = True
   End If
   comboSplitDeDuc.Text = QPTrim$(SysFileRec.SplitFlag)
   txtExpAllocation.Text = QPTrim$(SysFileRec.EXPMETHD)
   fptxtFringeCost.Text = SysFileRec.FRNGRATE
   txtFringeExp.Text = QPTrim$(SysFileRec.FRNGEXP)
   txtFringeFundCash.Text = QPTrim$(SysFileRec.FRNGDR)
   txtFringeFundCredit.Text = QPTrim$(SysFileRec.FRNGCR)
   fptxtIndirectCost.Text = SysFileRec.INDRATE
   txtIndirectExp.Text = QPTrim$(SysFileRec.INDEXP)
   txtIndirectFundCash.Text = QPTrim$(SysFileRec.INDDR)
   txtIndirectFundCredit.Text = QPTrim$(SysFileRec.INDCR)
   txtSocSecExp.Text = QPTrim$(SysFileRec.SOCEXP)
   txtSocSecLiab.Text = QPTrim$(SysFileRec.SOCLIAB)
   txtMedicareExp.Text = QPTrim$(SysFileRec.MEDEXP)
   txtMedicareLiab.Text = QPTrim$(SysFileRec.MEDLIAB)
   txtRetirementExp.Text = QPTrim$(SysFileRec.RETEXP)
   txtRetirementLiab.Text = QPTrim$(SysFileRec.RETLIAB)
   
'   Call GetAcctStruct(txtCitipakDir.Text, FundLength, AcctLength, DetLength)
   
   txtFundandAcctLength.Text = Val(SysFileRec.AcctCnt) '  FundLength + AcctLength
   txtGLTranAcctLength.Text = Val(SysFileRec.GLActLen) ' + AcctLength + DetLength
   
'   If FundLength + AcctLength = 0 Then '7/26
'     MsgBox "The account lengths may not be accurate if the Citipak directory is wrong." '7/26
'   End If '7/26
NoUnitFileYet:
End Sub

Private Sub fpcomboChkStyle_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcomboChkStyle.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcomboChkStyle.ListIndex = -1
  End If
  If fpcomboChkStyle.ListDown = False Then
    If KeyCode = vbKeyDown Then
      SendKeys "{Tab}"
      KeyCode = 0
    ElseIf KeyCode = vbKeyUp Then
      SendKeys "+{Tab}"
      KeyCode = 0
    End If
  End If
End Sub

Private Sub fpcomboGLCheckYN_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcomboGLCheckYN.ListDown = True
  End If
  If fpcomboGLCheckYN.ListDown = False Then
    If KeyCode = vbKeyDown Then
      SendKeys "{Tab}"
      KeyCode = 0
    ElseIf KeyCode = vbKeyUp Then
      SendKeys "+{Tab}"
      KeyCode = 0
    End If
  End If

End Sub

Private Sub fpcomboGLCheckYN_LostFocus()
  If QPTrim$(fpcomboGLCheckYN.Text) = "N" Then
    GLCheckFlag = False
  Else
    GLCheckFlag = True
  End If
End Sub

Private Sub fptxtFringeCost_LostFocus()
  If fptxtFringeCost.Text = "" Then fptxtFringeCost.Text = "0"
End Sub

Private Sub fptxtIndirectCost_LostFocus()
  If fptxtIndirectCost.Text = "" Then fptxtIndirectCost.Text = "0"
End Sub

Private Sub mnuExit_Click()
  Call cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
  MainLog ("System Interface screen printed.")
End Sub

Private Sub txtCashWagesDue_DblClick(Button As Integer)
  txtCashWagesDue = Clipboard.GetText

End Sub

Private Sub txtCashWagesDue_LostFocus()
  If txtCashWagesDue.Text = "" Then txtCashWagesDue.Text = "0"
End Sub

'Private Sub txtCitipakDir_LostFocus()
'
'  If QPTrim$(txtCitipakDir.Text) = "" Then GoTo NoGLFiles
'  If CheckCitiDir(txtCitipakDir.Text) = True Then
'    If Not Exist(txtCitipakDir.Text + "GLACCT.IDX") And Not Exist(txtCitipakDir.Text + "\GLACCT.IDX") Then
'      MsgBox "The GLACCT.IDX file could not be found in this Citipak directory."
'      fpcomboGLCheckYN.SetFocus
'      Exit Sub
'    End If
'
'    If Not Exist(txtCitipakDir.Text + "GLACCT.DAT") And Not Exist(txtCitipakDir.Text + "\GLACCT.DAT") Then
'      MsgBox "The GLACCT.DAT file could not be found in this Citipak directory."
'      fpcomboGLCheckYN.SetFocus
'      Exit Sub
'    End If
'  Else
'    MsgBox "This Citipak path cannot be found."
'    fpcomboGLCheckYN.SetFocus
'    Exit Sub
'  End If
'NoGLFiles:
'  CurrCitiPath = QPTrim$(txtCitipakDir.Text)
'End Sub

Private Sub txtExpAllocation_LostFocus()
  If txtExpAllocation.Text = "" Then txtExpAllocation.Text = "0"
End Sub

Private Sub txtFedWHNum_DblClick(Button As Integer)
  txtFedWHNum.Text = Clipboard.GetText
End Sub

Private Sub txtFedWHNum_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    SendKeys "{Tab}"
    KeyCode = 0
  ElseIf KeyCode = vbKeyUp Then
    SendKeys "+{Tab}"
    KeyCode = 0
  End If

End Sub

Private Sub txtFedWHNum_LostFocus()
  If txtFedWHNum.Text = "" Then txtFedWHNum.Text = "0"
End Sub

Private Sub txtFringeExp_LostFocus()
  If txtFringeExp.Text = "" Then txtFringeExp.Text = "0"
End Sub

Private Sub txtFringeFundCash_DblClick(Button As Integer)
 txtFringeFundCash = Clipboard.GetText

End Sub

Private Sub txtFringeFundCash_LostFocus()
  If txtFringeFundCash.Text = "" Then txtFringeFundCash.Text = "0"
End Sub

Private Sub txtFringeFundCredit_Change()
  If txtFringeFundCredit.Text = "" Then txtFringeFundCredit.Text = "0"
End Sub

Private Sub txtFringeFundCredit_DblClick(Button As Integer)
 txtFringeFundCredit = Clipboard.GetText

End Sub

Private Sub txtFundandAcctLength_LostFocus()
  If CheckFor2ManyDecimals(txtFundandAcctLength.Text) = True Then
    MsgBox "Invalid number entered"
    txtFundandAcctLength.SetFocus
    Exit Sub
  End If
  If txtFundandAcctLength.Text = "" Then txtFundandAcctLength.Text = "0"
End Sub

Private Sub txtGLTranAcctLength_KeyPress(KeyAscii As Integer)
  If KeyAscii = 40 Then
    comboInterface.SetFocus
  End If

End Sub

Private Sub txtGLTranAcctLength_LostFocus()
  If CheckFor2ManyDecimals(txtGLTranAcctLength.Text) = True Then
    MsgBox "Invalid number entered"
    txtGLTranAcctLength.SetFocus
    Exit Sub
  End If
  If txtGLTranAcctLength.Text = "" Then txtGLTranAcctLength.Text = "0"
 
End Sub

Private Sub txtImprest_DblClick(Button As Integer)
  txtImprest = Clipboard.GetText
End Sub

Private Sub txtImprest_LostFocus()
  If txtImprest.Text = "" Then txtImprest.Text = "0"

End Sub

Private Sub txtIndirectExp_LostFocus()
  If txtIndirectExp.Text = "" Then txtIndirectExp.Text = "0"
End Sub

Private Sub txtIndirectFundCash_DblClick(Button As Integer)
 txtIndirectFundCash = Clipboard.GetText

End Sub

Private Sub txtIndirectFundCash_LostFocus()
  If txtIndirectFundCash.Text = "" Then txtIndirectFundCash.Text = "0"
End Sub

Private Sub txtIndirectFundCredit_DblClick(Button As Integer)
  txtIndirectFundCredit = Clipboard.GetText

End Sub

Private Sub txtIndirectFundCredit_LostFocus()
  If txtIndirectFundCredit.Text = "" Then txtIndirectFundCredit.Text = "0"

End Sub

Private Sub txtMedicareExp_DblClick(Button As Integer)
  txtMedicareExp = Clipboard.GetText

End Sub

Private Sub txtMedicareExp_LostFocus()
  If txtMedicareExp.Text = "" Then txtMedicareExp.Text = "0"

End Sub

Private Sub txtMedicareLiab_DblClick(Button As Integer)
  txtMedicareLiab = Clipboard.GetText

End Sub

Private Sub txtMedicareLiab_LostFocus()
  If txtMedicareLiab.Text = "" Then txtMedicareLiab.Text = "0"
End Sub

Private Sub txtMedicareWHNum_DblClick(Button As Integer)
  txtMedicareWHNum.Text = Clipboard.GetText

End Sub

Private Sub txtMedicareWHNum_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    SendKeys "{Tab}"
    KeyCode = 0
  ElseIf KeyCode = vbKeyUp Then
    SendKeys "+{Tab}"
    KeyCode = 0
  End If

End Sub

Private Sub txtMedicareWHNum_LostFocus()
  If txtMedicareWHNum.Text = "" Then txtMedicareWHNum.Text = "0"
End Sub

Private Sub txtRetirementExp_DblClick(Button As Integer)
  txtRetirementExp = Clipboard.GetText

End Sub

Private Sub txtRetirementExp_LostFocus()
  If txtRetirementExp.Text = "" Then txtRetirementExp.Text = "0"
End Sub

Private Sub txtRetirementLiab_DblClick(Button As Integer)
  txtRetirementLiab = Clipboard.GetText

End Sub

Private Sub txtRetirementLiab_LostFocus()
  If txtRetirementLiab.Text = "" Then txtRetirementLiab.Text = "0"
End Sub

Private Sub txtRetireWHNum_DblClick(Button As Integer)
  txtRetireWHNum = Clipboard.GetText
End Sub

Private Sub txtRetireWHNum_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    SendKeys "{Tab}"
    KeyCode = 0
  ElseIf KeyCode = vbKeyUp Then
    SendKeys "+{Tab}"
    KeyCode = 0
  End If

End Sub

Private Sub txtRetireWHNum_LostFocus()
  If txtRetireWHNum.Text = "" Then txtRetireWHNum.Text = "0"
End Sub

Private Sub txtSocSecExp_DblClick(Button As Integer)
 txtSocSecExp = Clipboard.GetText
End Sub

Private Sub txtSocSecExp_LostFocus()
  If txtSocSecExp.Text = "" Then txtSocSecExp.Text = "0"
End Sub

Private Sub txtSocSecLiab_DblClick(Button As Integer)
  txtSocSecLiab = Clipboard.GetText
End Sub

Private Sub txtSocSecLiab_LostFocus()
  If txtSocSecLiab.Text = "" Then txtSocSecLiab.Text = "0"
End Sub

Private Sub txtSocSecWHNum_DblClick(Button As Integer)
  txtSocSecWHNum = Clipboard.GetText
End Sub

Private Sub txtSocSecWHNum_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    SendKeys "{Tab}"
    KeyCode = 0
  ElseIf KeyCode = vbKeyUp Then
    SendKeys "+{Tab}"
    KeyCode = 0
  End If

End Sub

Private Sub txtSocSecWHNum_LostFocus()
  If txtSocSecWHNum.Text = "" Then txtSocSecWHNum.Text = "0"
End Sub

Private Sub txtStateWHNum_DblClick(Button As Integer)
  txtStateWHNum = Clipboard.GetText
End Sub

Private Sub txtStateWHNum_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    SendKeys "{Tab}"
    KeyCode = 0
  ElseIf KeyCode = vbKeyUp Then
    SendKeys "+{Tab}"
    KeyCode = 0
  End If

End Sub

Private Sub txtStateWHNum_LostFocus()
  If txtStateWHNum.Text = "" Then txtStateWHNum.Text = "0"
End Sub

Private Sub txtWagesCentDep_DblClick(Button As Integer)
  txtWagesCentDep = Clipboard.GetText

End Sub

Private Sub txtWagesCentDep_LostFocus()
  If txtWagesCentDep.Text = "" Then txtWagesCentDep.Text = "0"

End Sub
Private Sub CheckForValidWHNum()
   Dim Split$
   Dim JGLIdxRec(1) As JGLAcctIdxType
   Dim GLIdxNum$
   Dim GLDHandle As Integer
   Dim GLIdxRecLen As Integer
   Dim GLDescRecLen As Integer
   Dim TotalAccts As Integer
   Dim Nextx As Integer, x As Integer
   Dim GLIDATDesc$
   Dim GLDesc(1) As GLAcctRecType
   Dim GLIdxHandle As Integer
   Dim WHNum(1 To 5) As String
   Dim FixedDDNum(1 To 5) As String 'always on and just dept/detail nums
   Dim FixedFull As String 'this is imprest only which is always the entire gl num
   Dim FixedDetail(1 To 3)
   Dim DoWhatFlag As BadGLNUMOption
   Dim n As Integer
   Dim FundLength As Integer
   Dim AcctLength As Integer
   Dim DetLength As Integer
   
   If QPTrim$(PICType) = "I" Then Exit Sub
   On Error GoTo ERRORSTUFF
   
'   Call GetAcctStruct(txtCitipakDir.Text, FundLength, AcctLength, DetLength)
   Call GetAcctStruct(CurrCitiPath, FundLength, AcctLength, DetLength)
   'No need to trap because they are not using our gl package
   If FundLength = 0 And AcctLength = 0 And DetLength = 0 Then
     Exit Sub
   End If
   
   Split$ = comboSplitDeDuc.Text
   
   BadGLNum = False
   
   WHNum(1) = QPTrim$(ReplaceString(txtFedWHNum.Text, "-", ""))
   WHNum(2) = QPTrim$(ReplaceString(txtStateWHNum.Text, "-", ""))
   WHNum(3) = QPTrim$(ReplaceString(txtSocSecWHNum.Text, "-", ""))
   WHNum(4) = QPTrim$(ReplaceString(txtMedicareWHNum.Text, "-", ""))
   WHNum(5) = QPTrim$(ReplaceString(txtRetireWHNum.Text, "-", ""))
   'FixedDDNum 1 and 2 are reversed according to the location on the screen
   'txtWagesCentDep is only valid if "C" or "I" is selected so it is easier
   'to start with it when validating GL numbers if "C" is chosen
   'or start with 2 if "C" or "I" is not chosen
   FixedDDNum(1) = QPTrim$(ReplaceString(txtWagesCentDep.Text, "-", ""))
   FixedDDNum(2) = QPTrim$(ReplaceString(txtCashWagesDue.Text, "-", ""))
   FixedDDNum(3) = QPTrim$(ReplaceString(txtSocSecLiab.Text, "-", ""))
   FixedDDNum(4) = QPTrim$(ReplaceString(txtMedicareLiab.Text, "-", ""))
   FixedDDNum(5) = QPTrim$(ReplaceString(txtRetirementLiab.Text, "-", ""))
   FixedFull = QPTrim$(ReplaceString(txtImprest.Text, "-", ""))
   FixedDetail(1) = QPTrim$(ReplaceString(txtSocSecExp.Text, "-", ""))
   FixedDetail(2) = QPTrim$(ReplaceString(txtMedicareExp.Text, "-", ""))
   FixedDetail(3) = QPTrim$(ReplaceString(txtRetirementExp.Text, "-", ""))
   
   'we are looking at whatever is in the current citipak directory field
   'assuming it to be the most up-to-date path
'   If Exist(txtCitipakDir.Text + "GLACCT.IDX") Then
'     GLIdxNum$ = txtCitipakDir.Text + "GLACCT.IDX"
'   ElseIf Exist(txtCitipakDir.Text + "\GLACCT.IDX") Then
'     GLIdxNum$ = txtCitipakDir.Text + "\GLACCT.IDX"
   If Exist(CurrCitiPath + "GLACCT.IDX") Then
     GLIdxNum$ = CurrCitiPath + "GLACCT.IDX"
   ElseIf Exist(CurrCitiPath + "\GLACCT.IDX") Then
     GLIdxNum$ = CurrCitiPath + "\GLACCT.IDX"
   Else
     MsgBox "No G/L account number validation possible...GLACCT.IDX could not be found."
     Exit Sub
   End If

'   If Exist(txtCitipakDir.Text + "GLACCT.DAT") Then
'     GLIDATDesc$ = txtCitipakDir.Text + "GLACCT.DAT"
'   ElseIf Exist(txtCitipakDir.Text + "\GLACCT.DAT") Then
'     GLIDATDesc$ = txtCitipakDir.Text + "\GLACCT.DAT"
   If Exist(CurrCitiPath + "GLACCT.DAT") Then
     GLIDATDesc$ = CurrCitiPath + "GLACCT.DAT"
   ElseIf Exist(CurrCitiPath + "\GLACCT.DAT") Then
     GLIDATDesc$ = CurrCitiPath + "\GLACCT.DAT"
   Else
     MsgBox "No G/L account number validation possible...GLACCT.DAT could not be found."
     Exit Sub
   End If
   
   GLIdxRecLen = Len(JGLIdxRec(1))
   GLDescRecLen = Len(GLDesc(1))
   TotalAccts = FileSize(GLIDATDesc$) \ GLDescRecLen
   
   If TotalAccts = 0 Then Exit Sub
   ReDim DescBuff(1 To TotalAccts)
   GLIdxHandle = FreeFile
   Open GLIdxNum$ For Random As GLIdxHandle Len = GLIdxRecLen
   For x = 1 To TotalAccts
     Get GLIdxHandle, x, JGLIdxRec(1)
     DescBuff(x) = JGLIdxRec(1).RecNo
   Next x
   Close GLIdxHandle
   GLDHandle = FreeFile
   Open GLIDATDesc$ For Random As GLDHandle Len = GLDescRecLen
   
   'the following Fixed numbers are always checked regardless
   'of split option choice..if "P' then we don't need to check the
   'unenabled txtWagesCentDep (#1 FixedDDNum)
   If QPTrim$(PICType) = "P" Then
     Nextx = 2
   Else
     Nextx = 1
   End If
   'go thru each number one at a time and compare against all gl nums
   'if "C" or "I" is chosen then begin at 1 else begin at 2
   For Nextx = Nextx To 5
     For x = 1 To TotalAccts
       If DescBuff(x) = 0 Then GoTo DescBuffIsZero
         Get GLDHandle, DescBuff(x), GLDesc(1)
         If Nextx = 1 Then '7/26 Nextx is 1 only if "C" is used
         'this checks only the Wages/Pay Cent Dep Due field for the fund and acct length
         '6/29/2004 added Mid to FixedDDNum(Nextx) to compare only the first 2 components
         'because there is an issue with including fund numbers at the end but sometimes
         'a "0" pad is needed if the detail length is longer than the fund length...example...
         'fund length is 2 but detail length is 4, if you want to add a fund to the end of
         'this GL number than you cannot just add the fund, you have to include the fund number
         'embedded with a "00" pad to make it equal the correct GL number length
           If Mid(FixedDDNum(Nextx), 1, (FundLength + AcctLength)) = QPTrim$(ReplaceString(Mid(GLDesc(1).Num, 1, FundLength + AcctLength + 1), "-", "")) Then
             Exit For
           End If
         Else '7/26
         'This If checks the value in the field against all GL fund and detail
         'lengths only
           If FixedDDNum(Nextx) = QPTrim$(ReplaceString(Mid(GLDesc(1).Num, FundLength + 2, AcctLength + DetLength + 1), "-", "")) Then
             Exit For
           End If
         End If '7/26
DescBuffIsZero:
         'been all the way through all 5 tests and no match has
         'been found so it is a BadGLNum
         If x = TotalAccts Then
           Select Case Nextx
             Case 1:
               txtWagesCentDep.SetFocus
             Case 2:
               txtCashWagesDue.SetFocus
             Case 3:
               txtSocSecLiab.SetFocus
             Case 4:
               txtMedicareLiab.SetFocus
             Case 5:
               txtRetirementLiab.SetFocus
           End Select
           
           DoWhatFlag = PromptBadGLNum(Me)
           Select Case DoWhatFlag
           Case BadGLNUMOption.badglExit
             BadGLNum = True
             Close
             frmControlFileMaint.Show
             DoEvents
             Unload frmIntFaceMaint
             Exit Sub
           Case BadGLNUMOption.badglReturn
             Close
             BadGLNum = True
             Exit Sub
           Case BadGLNUMOption.badglSave
             Close
             OverrideFlag = True
             Exit Sub
           Case Else:
              'Do nothing because we don't know about any options except
              'save, review or abandon...used as a placeholder for adding
              'other options at a later date
           End Select
           Close GLDHandle
           Exit Sub
         End If
      Next x
   Next Nextx
   
   If QPTrim$(PICType) = "C" Then
     For x = 1 To TotalAccts
       If DescBuff(x) = 0 Then GoTo DescBuffIsEmpty
         Get GLDHandle, DescBuff(x), GLDesc(1)
         If FixedFull = QPTrim$(ReplaceString(GLDesc(1).Num, "-", "")) Then
           Exit For
         End If
DescBuffIsEmpty:
     Next x
   End If
   If x = TotalAccts + 1 Then
     txtImprest.SetFocus
     DoWhatFlag = PromptBadGLNum(Me)
     Select Case DoWhatFlag
     Case BadGLNUMOption.badglExit
       Close
       BadGLNum = True
       frmControlFileMaint.Show
       DoEvents
       Unload frmIntFaceMaint
       Exit Sub
     Case BadGLNUMOption.badglReturn
       Close
       BadGLNum = True
       Exit Sub
     Case BadGLNUMOption.badglSave
       Close
       OverrideFlag = True
       Exit Sub
     Case Else:
       'Do nothing because we don't know about any options except
       'save, review or abandon...used as a placeholder for adding
       'other options at a later date
     End Select
     Close GLDHandle
     Exit Sub
   End If
     
   Nextx = 1
   For Nextx = Nextx To 3
     For x = 1 To TotalAccts
       If DescBuff(x) = 0 Then GoTo DescBuffIsNatta
         Get GLDHandle, DescBuff(x), GLDesc(1)
         If FixedDetail(Nextx) = QPTrim$(ReplaceString(Mid(GLDesc(1).Num, FundLength + AcctLength + 3, DetLength), "-", "")) Then
           Exit For
         End If
DescBuffIsNatta:
         'been all the way through all 3 tests and no match has
         'been found so it is a BadGLNum
         If x = TotalAccts Then
           Select Case Nextx
             Case 1:
               txtSocSecExp.SetFocus
             Case 2:
               txtMedicareExp.SetFocus
             Case 3:
               txtRetirementExp.SetFocus
           End Select
           
           DoWhatFlag = PromptBadGLNum(Me)
           Select Case DoWhatFlag
           Case BadGLNUMOption.badglExit
             Close
             BadGLNum = True
             frmControlFileMaint.Show
             DoEvents
             Unload frmIntFaceMaint
             Exit Sub
           Case BadGLNUMOption.badglReturn
             Close
             BadGLNum = True
             Exit Sub
           Case BadGLNUMOption.badglSave
             Close
             OverrideFlag = True
             Exit Sub
           Case Else:
              'Do nothing because we don't know about any options except
              'save, review or abandon...used as a placeholder for adding
              'other options at a later date
           End Select
'           BadGLNum = True
           Close GLDHandle
           Exit Sub
         End If
      Next x
   Next Nextx
     
   Nextx = 1
   If Split$ = "Y" Then
     GoTo SplitY
   End If
   For Nextx = Nextx To 5
     For x = 1 To TotalAccts
       If DescBuff(x) = 0 Then GoTo DescBuffIs0N
         Get GLDHandle, DescBuff(x), GLDesc(1)
         If WHNum(Nextx) = QPTrim$(ReplaceString(GLDesc(1).Num, "-", "")) Then
           Exit For
         End If
DescBuffIs0N:
         'been all the way through all 5 tests and no match has
         'been found so it is a BadGLNum
         If x = TotalAccts Then
           Select Case Nextx
             Case 1:
               txtFedWHNum.SetFocus
             Case 2:
               txtStateWHNum.SetFocus
             Case 3:
               txtSocSecWHNum.SetFocus
             Case 4:
               txtMedicareWHNum.SetFocus
             Case 5:
               txtRetireWHNum.SetFocus
           End Select
           
           DoWhatFlag = PromptBadGLNum(Me)
           Select Case DoWhatFlag
           Case BadGLNUMOption.badglExit
             Close
             BadGLNum = True
             frmControlFileMaint.Show
             DoEvents
             Unload frmIntFaceMaint
             Exit Sub
           Case BadGLNUMOption.badglReturn
             Close
             BadGLNum = True
             Exit Sub
           Case BadGLNUMOption.badglSave
             Close
             OverrideFlag = True
             Exit Sub
           Case Else:
              'Do nothing because we don't know about any options except
              'save, review or abandon...used as a placeholder for adding
              'other options at a later date
           End Select
'           BadGLNum = True
           Close GLDHandle
           Exit Sub
         End If
      Next x
   Next Nextx
   GoTo SplitN
SplitY:
   For Nextx = Nextx To 5
     For x = 1 To TotalAccts
        If DescBuff(x) = 0 Then GoTo DescBuffIs0Y
          Get GLDHandle, DescBuff(x), GLDesc(1)
          If WHNum(Nextx) = QPTrim$(ReplaceString(Mid(GLDesc(1).Num, FundLength + 2, AcctLength + DetLength + 2), "-", "")) Then
            Exit For
          End If
DescBuffIs0Y:
          If x = TotalAccts Then
            Select Case Nextx
              Case 1:
                txtFedWHNum.SetFocus
              Case 2:
                txtStateWHNum.SetFocus
              Case 3:
                txtSocSecWHNum.SetFocus
              Case 4:
                txtMedicareWHNum.SetFocus
              Case 5:
                txtRetireWHNum.SetFocus
            End Select
            DoWhatFlag = PromptBadGLNum(Me)
            Select Case DoWhatFlag
            Case BadGLNUMOption.badglExit
               Close
               BadGLNum = True
               frmControlFileMaint.Show
               DoEvents
               Unload frmIntFaceMaint
               Exit Sub
            Case BadGLNUMOption.badglReturn
              BadGLNum = True
              Close
              Exit Sub
            Case BadGLNUMOption.badglSave
              OverrideFlag = True
              Close
              Exit Sub
            Case Else:
              'Do nothing because we don't know about any options except
              'save, review or abandon...used as a placeholder for adding
              'other options at a later date
            End Select
'            BadGLNum = True
            Close GLDHandle
            Exit Sub
         End If
     Next x
   Next Nextx

SplitN:
  Close GLDHandle
  
  Exit Sub
   
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmIntFaceMaint", "CheckForValidWHNum", Erl)
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("Payroll.exe terminated via menu bar on frmIntFaceMaint.")
      Call Terminate
      End
    End If
  End If
End Sub

