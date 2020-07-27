VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "EDT32X30.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Begin VB.Form frm401KMag 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "401K Magnetic Media Report"
   ClientHeight    =   8868
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   11652
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   30550.05
   ScaleMode       =   0  'User
   ScaleWidth      =   36263.11
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcomboDrive 
      Height          =   384
      Left            =   4656
      TabIndex        =   1
      Top             =   2640
      Width           =   924
      _Version        =   196608
      _ExtentX        =   1630
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
      Style           =   0
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
      ColDesigner     =   "frm401KMag.frx":0000
   End
   Begin LpLib.fpCombo fpcomboMonth 
      Height          =   384
      Left            =   4656
      TabIndex        =   2
      Top             =   3216
      Width           =   2508
      _Version        =   196608
      _ExtentX        =   4424
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
      Style           =   0
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
      ColDesigner     =   "frm401KMag.frx":02F7
   End
   Begin LpLib.fpCombo fpcomboLoanPayCode 
      Height          =   384
      Left            =   4656
      TabIndex        =   5
      Top             =   4944
      Width           =   2508
      _Version        =   196608
      _ExtentX        =   4424
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
      Style           =   0
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
      ColDesigner     =   "frm401KMag.frx":05EE
   End
   Begin LpLib.fpCombo fpcomboVolCode 
      Height          =   384
      Left            =   4656
      TabIndex        =   4
      Top             =   4368
      Width           =   2508
      _Version        =   196608
      _ExtentX        =   4424
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
      Style           =   0
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
      ColDesigner     =   "frm401KMag.frx":08E5
   End
   Begin EditLib.fpText fptxtYear 
      Height          =   396
      Left            =   4656
      TabIndex        =   6
      Top             =   3792
      Width           =   1164
      _Version        =   196608
      _ExtentX        =   2053
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
      CharValidationText=   "1, 2, 3, 4, 5, 6, 7, 8, 9, 0"
      MaxLength       =   4
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
   Begin EditLib.fpText fptxtCodeGPct 
      Height          =   396
      Left            =   4656
      TabIndex        =   13
      Top             =   5520
      Width           =   1164
      _Version        =   196608
      _ExtentX        =   2053
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
      CharValidationText=   "1, 2, 3, 4, 5, 6, 7, 8, 9, 0"
      MaxLength       =   4
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
   Begin EditLib.fpText fptxtCodeLPct 
      Height          =   396
      Left            =   4656
      TabIndex        =   14
      Top             =   6096
      Width           =   1164
      _Version        =   196608
      _ExtentX        =   2053
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
      CharValidationText=   "1, 2, 3, 4, 5, 6, 7, 8, 9, 0"
      MaxLength       =   4
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
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Code L Pct:"
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
      Height          =   300
      Left            =   2304
      TabIndex        =   12
      Top             =   6192
      Width           =   2124
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Code G Pct:"
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
      Height          =   300
      Left            =   2304
      TabIndex        =   11
      Top             =   5616
      Width           =   2124
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Loan Payment Code:"
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
      Height          =   300
      Left            =   2064
      TabIndex        =   10
      Top             =   5040
      Width           =   2364
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Voluntary Code:"
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
      Height          =   300
      Left            =   2304
      TabIndex        =   9
      Top             =   4464
      Width           =   2124
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Enter the Year:"
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
      Height          =   300
      Left            =   2304
      TabIndex        =   8
      Top             =   3888
      Width           =   2124
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Reporting Month:"
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
      Height          =   300
      Left            =   2304
      TabIndex        =   7
      Top             =   3312
      Width           =   2124
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Drive (A-B):"
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
      Height          =   300
      Left            =   2304
      TabIndex        =   3
      Top             =   2736
      Width           =   2124
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "401K Magnetic Media Report"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   396
      Left            =   3504
      TabIndex        =   0
      Top             =   1680
      Width           =   4620
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   6636
      Left            =   1500
      Top             =   1116
      Width           =   8652
   End
End
Attribute VB_Name = "frm401KMag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyUp:
      SendKeys "+{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%C"
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%P"
      KeyCode = 0
    Case Else:
  End Select
End Sub
Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Call LoadThisForm
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If

End Sub

Public Sub LoadThisForm()
'DECLARE SUB VertMenu401 (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS Any)

  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'PRFiles.bi'
  '$INCLUDE: 'PREmpRec.bi'
  '$INCLUDE: 'PRTRANS.Bi'
  '$INCLUDE: 'PRUNIT.Bi'
  '$INCLUDE: 'DedCodes.Bi'
  '$INCLUDE: 'PRRpts.BI'
  '$INCLUDE: 'PR401k.BI'

'  COMMON SHARED Cnf  AS Config

  '$INCLUDE: 'SetCnf.BI'

'  CONST False = 0, True = NOT False

'  ReDim TRec(1) As TransRecType
'  ReDim E2Rec(1) As EmpData2Type
'  ReDim Item$(1)
'  ReDim Unit(1) As UnitFileRecType
  Dim TRec(1) As TransRecType
  Dim E2Rec(1) As EmpData2Type
  Dim Item$(1)
  Dim Unit(1) As UnitFileRecType
  Dim UnitFileName As Integer
  Dim MaxLen As Integer
  Dim Image1$, Image2$
  Dim TRecSize As Long
  Dim EmpRecSize As Integer
  Dim NumKeys$, DrvKeys$
  Dim q$, Edit$
  
  
  
  OpenUnitFile UnitFileName
  Get UnitFileName, 1, Unit(1)
'  FGetAH UnitFileName, Unit(1), Len(Unit(1)), 1

  GoSub LoadDedCodes

  MaxLen = 15

  Image1$ = "######.##"
  Image2$ = "######"

  TRecSize = Len(TRec(1))
  EmpRecSize = Len(E2Rec(1))

  NumKeys$ = "1234567890"
  DrvKeys$ = "AaBbcC"
  q$ = Chr$(34)

  Color 15, 1

top:
  Do
    Cls
    LOCATE 2, 15
    Print "401K BB&T Magnetic Media Report"
    Edit$ = " "
    LOCATE 6, 9
    Print "Enter Drive(A-B): ";
    WInput Edit$, DrvKeys$, 6, 29, ExitCode
    If ExitCode = -27 Or Len(Edit$) = 0 Then
      ExitFlag = True
      Exit Do
    End If
    Drive$ = Edit$

GetLastMonth:
    Edit$ = "  "
    LOCATE 8, 10
    Print "Reporting Month:     ";
    WInput Edit$, NumKeys$, 8, 29, ExitCode
    If ExitCode = -27 Or Len(Edit$) = 0 Then
      ExitFlag = True
      Exit Do
    End If
    EMonth = QPValI(Edit$)
    If (EMonth < 1 Or EMonth > 12) Or BMonth > EMonth Then
      LOCATE 12, 10
      Print "Invalid Month Specification."
      LOCATE 14, 11
      Print "Press any key to continue."
      dodo = BiosKey
      CTop = 8
      GoSub ClearArea
      GoTo GetLastMonth
    End If
    '-----

GetYear:
    Edit$ = "    "
    LOCATE 9, 11
    Print "Enter the Year:       ";
    WInput Edit$, NumKeys$, 9, 29, ExitCode
    If ExitCode = -27 Then
      ExitFlag = True
      Exit Do
    End If

    Year = QPValI(Edit$)
    If Year <= 0 Then
      LOCATE 12, 10
      Print "Invalid Year Specifcation."
      LOCATE 14, 10
      Print "Press any key to continue."
      dodo = BiosKey
      CTop = 9
      GoSub ClearArea
      GoTo GetYear
    End If

GetVCode:
    LOCATE 10, 11
    Print "Voluntary Code:                   ";
    LOCATE 10, 34
    VertMenu401 Item$(), Choice, MaxLen, 17, Ky$, 0, Cnf
    If Ky$ = Chr$(27) Then
      GoTo EndTheProg
    Else
      LOCATE 10, 28
      Print Item$(Choice)
      VCodeNum = Choice
    End If

GetLCode:
    LOCATE 11, 8
    Print "Loan Payment Code:                ";
    LOCATE 11, 34
    VertMenu401 Item$(), Choice, MaxLen, 18, Ky$, 0, Cnf
    If Ky$ = Chr$(27) Then
      GoTo EndTheProg
    Else
      LOCATE 11, 28
      Print Item$(Choice)
      LCodeNum = Choice
    End If

GetGPct:
    Edit$ = "    "
    LOCATE 12, 15
    Print "Code G Pct:           ";
    WInput Edit$, NumKeys$ + ".", 12, 29, ExitCode
    If ExitCode = -27 Then
      ExitFlag = True
      Exit Do
    End If

    GPct# = Val(Edit$)
    If GPct# <= 0 Then
      CTop = 12
      GoSub ClearArea
      GoTo GetGPct
    End If

GetLPct:
    Edit$ = "    "
    LOCATE 13, 15
    Print "Code L Pct:           ";
    WInput Edit$, NumKeys$ + ".", 13, 29, ExitCode
    If ExitCode = -27 Then
      ExitFlag = True
      Exit Do
    End If

    LPct# = Val(Edit$)
    If LPct# <= 0 Then
      CTop = 13
      GoSub ClearArea
      GoTo GetLPct
    End If
    Exit Do
  Loop

  If ExitFlag Then
    GoTo EndTheProg
  End If

  If EMonth < 10 Then
    EMonth$ = "0" + LTrim$(Str$(EMonth))
  Else
    EMonth$ = LTrim$(Str$(EMonth))
  End If

  BMonth$ = EMonth$

  Year$ = LTrim$(Str$(Year))

  LowDate = Date2Num(BMonth$ + "-" + "01" + "-" + Year$)

  Select Case EMonth
  Case 2
    HiDate = Date2Num(EMonth$ + "-" + "28" + "-" + Year$)
  Case 4, 6, 9, 11
    HiDate = Date2Num(EMonth$ + "-" + "30" + "-" + Year$)
  Case 1, 3, 5, 7, 8, 10, 12
    HiDate = Date2Num(EMonth$ + "-" + "31" + "-" + Year$)
  End Select

  IdxRecLen = 2
  IdxFileSize& = FileSize(EmpIdxNName)
  NumOfRecs = IdxFileSize& \ IdxRecLen

  If DosError Then
    LOCATE 15, 10
    Print "Unable to Find/Open Transaction History file!"
    LOCATE 16, 10
    Print "Press any key to return to system."
    dodo = BiosKey
    GoTo EndTheProg
  End If

  ReDim IdxBuff(1 To NumOfRecs)
  FGetAH EmpIdxNName, IdxBuff(1), IdxRecLen, NumOfRecs

  RptName$ = Drive$ + ":\NC401K"

  FCreate RptName$
  If DosError Then
    LOCATE 15, 10
    Print "Unable to Open/Create report file!"
    LOCATE 16, 10
    Print "Press any key to return to system."
    dodo = BiosKey
    GoTo EndTheProg
  End If
  '*****************
  'make disk report here

  CrLf$ = Chr$(13) + Chr$(10)

  ReDim TransHRec(1) As TransRecType
  ReDim Emp2Rec(1) As EmpData2Type

  EmpRecSize = Len(Emp2Rec(1))
  TRecSize = Len(TransHRec(1))

  IdxRecLen = 2

  IdxFileSize& = FileSize(EmpIdxLName)
  NumOfRecs = IdxFileSize& \ IdxRecLen

  ReDim IdxBuff(1 To NumOfRecs)
  FGetAH EmpIdxLName, IdxBuff(1), IdxRecLen, NumOfRecs

  ReDim D401kRec(1) As DetailRecType
  ReDim T401kRec(1) As TrailerRecType
  D401Len = Len(D401kRec(1))
  T401Len = Len(T401kRec(1))

  'got input here

  RptFile = FreeFile
  Open RptName$ For Output As #RptFile
  Close RptFile

  RptFile = FreeFile
  Open RptName$ For Random As #RptFile Len = D401Len
  HFile = FreeFile
  Open TransHistFileName For Random As #HFile Len = TRecSize
  EFile = FreeFile
  Open EmpData2Name For Random As #EFile Len = EmpRecSize

  For RecNo = 1 To NumOfRecs
    UsingThisOne = False
    VCalcAmt# = 0
    LCalcAmt# = 0
    GCalcAmt# = 0


    Get #EFile, IdxBuff(RecNo), Emp2Rec(1)

    If Emp2Rec(1).LastTransRec <= 0 Then
      GoTo SkipEm
    End If

    TransRecNum& = Emp2Rec(1).LastTransRec
    Do
      Get #HFile, TransRecNum&, TransHRec(1)

      Select Case TransHRec(1).CheckDate

      Case LowDate To HiDate

        If VCodeNum > 0 Then
          If TransHRec(1).DAmt(VCodeNum) <> 0 Then
            VCalcAmt# = RoundDbl#(VCalcAmt# + TransHRec(1).DAmt(VCodeNum))
            UsingThisOne = True
          End If
        End If
        If LCodeNum > 0 Then
          If TransHRec(1).DAmt(LCodeNum) > 0 Then
            LCalcAmt# = RoundDbl#(LCalcAmt# + TransHRec(1).DAmt(LCodeNum))
            UsingThisOne = True
          End If
        End If
        EmpRType$ = UCase$(Left$(LTrim$(Emp2Rec(1).EMPRETTP), 1))
        If EmpRType$ = "L" Or EmpRType$ = "G" Then
          GCalcAmt# = RoundDbl#(GCalcAmt# + TransHRec(1).GrossPay)
          UsingThisOne = True
        End If
      Case Else
      End Select

      If TransHRec(1).PrevTransRec <= 0 Then
        If UsingThisOne Then
          If EmpRType$ = "L" Then
            EPct# = LPct#
          Else
            EPct# = GPct#
          End If
          GoSub PrintThisOne
        End If
        Exit Do
      Else
        TransRecNum& = CLng(TransHRec(1).PrevTransRec)
      End If

    Loop

SkipEm:
    LOCATE 15, 1
    Print "Processing: "; Int((RecNo / NumOfRecs) * 100);
  Next


  GoSub DoTrailerRec

  Close

  '*****************
  LOCATE 15, 1
  Print Space$(79);
  LOCATE 15, 12
  Print "Report Completed."
  Print
  Print "Press any key to continue."
  dodo = BiosKey

EndTheProg:
  RUN "PR"
  End



PrintThisOne:

  If EPct# > 0 Or VCalcAmt# > 0 Or LCalcAmt# > 0 Then
    EPrinted = EPrinted + 1
    ReDim D401kRec(1) As DetailRecType

    TMatchAmt# = RoundDbl#((GCalcAmt# * EPct#) * 0.01)
    If EmpRType$ = "G" And VCalcAmt# = 0 Then
      If LCalcAmt# = 0 Then
        GoTo SkipEMBubba
      Else
        TMatchAmt# = 0
      End If
    ElseIf EmpRType$ = "G" Then
      If TMatchAmt# > VCalcAmt# Then
        TMatchAmt# = VCalcAmt#
      End If
    End If

    TotalVAmt# = RoundDbl#(TotalVAmt# + VCalcAmt#)
    TotalLAmt# = RoundDbl#(TotalLAmt# + LCalcAmt#)

    TotalMatchAmt# = RoundDbl#(TotalMatchAmt# + TMatchAmt#)

    LSet D401kRec(1).ID = "D"
    LSet D401kRec(1).Batch = "01001"
    LSet D401kRec(1).PCN = QPTrim$(Unit(1).BBTCNTNO)
    LSet D401kRec(1).ProcDate = EMonth$ + "31" + LTrim$(Str$(Year))
    LSet D401kRec(1).SSN = Emp2Rec(1).EmpSSN
    LSet D401kRec(1).EmpName = QPTrim$(Emp2Rec(1).EMPFNAME) + " " + QPTrim$(Emp2Rec(1).EMPLNAME)

    VolDed$ = RSet0$(VCalcAmt#, 7)
    LSet D401kRec(1).EmpVolDed = VolDed$        ''AS STRING * 8
    LoanDed$ = RSet0$(LCalcAmt#, 7)
    LSet D401kRec(1).EmpLoanPay = LoanDed$      ''AS STRING * 8
    ContDed$ = RSet0$(TMatchAmt#, 7)
    LSet D401kRec(1).EmpContAmt = ContDed$      ''AS STRING * 8
    D401kRec(1).CrLf = CrLf$

    Put #RptFile, , D401kRec(1)

  End If

SkipEMBubba:
  Return

DoTrailerRec:

  LSet T401kRec(1).ID = "T"

  TVolDed$ = RSet0$(TotalVAmt#, 10)
  LSet T401kRec(1).TotVolDED = TVolDed$       ''AS STRING * 11

  TLoanDed$ = RSet0$(TotalLAmt#, 10)
  LSet T401kRec(1).TotLoanAmt = TLoanDed$       ''AS STRING * 11

  TContDed$ = RSet0$(TotalMatchAmt#, 10)
  LSet T401kRec(1).TotContAmt = TContDed$       ''AS STRING * 11
  LSet T401kRec(1).Filler = ""

  TDetRecs$ = FUsing$(Str$(EPrinted), "###")
  TDetRecs$ = "000000" + QPTrim$(TDetRecs$)
  T401kRec(1).TotDRecs = Right$(TDetRecs$, 6)
  LSet T401kRec(1).CrLf = CrLf$

  Put #RptFile, , T401kRec(1)

  Return




ClearArea:
  T$ = Space$(60)
  For cnt = CTop To 18
    LOCATE cnt, 1: Print T$;
  Next
Return

LoadDedCodes:
  ReDim DedCode(1) As DedCodeRecType
  DedLen = Len(DedCode(1))
  DedFile = FreeFile
  Open DedCodeFileName For Random Shared As #DedFile Len = DedLen
  NumOfDed = LOF(DedFile) / DedLen
  ReDim Item$(1 To NumOfDed)
  For cnt = 1 To NumOfDed
    Get DedFile, cnt, DedCode(1)
    Item$(cnt) = Str$(cnt) + ") " + DedCode(1).DCDESC1
  Next
  Close
Return



End Sub

