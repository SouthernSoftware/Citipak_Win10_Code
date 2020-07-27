VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#3.5#0"; "SPR32X35.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmDeductionCodes 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Deduction Codes"
   ClientHeight    =   8565
   ClientLeft      =   30
   ClientTop       =   585
   ClientWidth     =   11655
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDeductionCodes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8347.614
   ScaleMode       =   0  'User
   ScaleWidth      =   11652
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcomboFWT 
      Height          =   405
      Left            =   1605
      TabIndex        =   2
      ToolTipText     =   "Is this deduction federal tax deferred?"
      Top             =   3315
      Width           =   885
      _Version        =   196608
      _ExtentX        =   1561
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
      Style           =   2
      MaxDrop         =   8
      ListWidth       =   -1
      EditHeight      =   -1
      GrayAreaColor   =   -2147483633
      ListLeftOffset  =   0
      ComboGap        =   -2
      MaxEditLen      =   1
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
      ColDesigner     =   "frmDeductionCodes.frx":08CA
   End
   Begin LpLib.fpCombo fpcomboSWT 
      Height          =   405
      Left            =   3435
      TabIndex        =   3
      ToolTipText     =   "Is this deduction State tax deferred?"
      Top             =   3315
      Width           =   870
      _Version        =   196608
      _ExtentX        =   1535
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
      Style           =   2
      MaxDrop         =   8
      ListWidth       =   -1
      EditHeight      =   -1
      GrayAreaColor   =   -2147483633
      ListLeftOffset  =   0
      ComboGap        =   -2
      MaxEditLen      =   1
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
      ColDesigner     =   "frmDeductionCodes.frx":0B51
   End
   Begin LpLib.fpCombo fpcomboSOC 
      Height          =   405
      Left            =   5205
      TabIndex        =   4
      ToolTipText     =   "Is this deduction Social Security Tax deferred?"
      Top             =   3315
      Width           =   885
      _Version        =   196608
      _ExtentX        =   1561
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
      Style           =   2
      MaxDrop         =   8
      ListWidth       =   -1
      EditHeight      =   -1
      GrayAreaColor   =   -2147483633
      ListLeftOffset  =   0
      ComboGap        =   -2
      MaxEditLen      =   1
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
      ColDesigner     =   "frmDeductionCodes.frx":0DD8
   End
   Begin LpLib.fpCombo fpcomboMED 
      Height          =   405
      Left            =   6960
      TabIndex        =   5
      ToolTipText     =   "Is this deduction Medicare Tax deferred?"
      Top             =   3315
      Width           =   870
      _Version        =   196608
      _ExtentX        =   1535
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
      Style           =   2
      MaxDrop         =   8
      ListWidth       =   -1
      EditHeight      =   -1
      GrayAreaColor   =   -2147483633
      ListLeftOffset  =   0
      ComboGap        =   -2
      MaxEditLen      =   1
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
      ColDesigner     =   "frmDeductionCodes.frx":105F
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdList 
      Height          =   372
      Left            =   1176
      TabIndex        =   20
      TabStop         =   0   'False
      ToolTipText     =   "Press F12 to bring up a General Ledger number list."
      Top             =   4440
      Width           =   3024
      _Version        =   131072
      _ExtentX        =   5334
      _ExtentY        =   656
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
      ButtonDesigner  =   "frmDeductionCodes.frx":12E6
   End
   Begin EditLib.fpText fpDescription 
      Height          =   375
      Left            =   2310
      TabIndex        =   0
      ToolTipText     =   "Enter a description for this deduction."
      Top             =   2370
      Width           =   5535
      _Version        =   196608
      _ExtentX        =   9763
      _ExtentY        =   661
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
      MaxLength       =   10
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
   Begin FPSpread.vaSpread vaSpreadDeductionCodes 
      Height          =   2964
      Left            =   1140
      TabIndex        =   6
      ToolTipText     =   "Double click any row to bring it up in the Data Entry section."
      Top             =   4896
      Width           =   9432
      _Version        =   196613
      _ExtentX        =   16669
      _ExtentY        =   5636
      _StockProps     =   64
      AutoSize        =   -1  'True
      ButtonDrawMode  =   4
      ColsFrozen      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   13684944
      MaxCols         =   6
      MaxRows         =   50
      ProcessTab      =   -1  'True
      ScrollBars      =   2
      ShadowColor     =   13684944
      SpreadDesigner  =   "frmDeductionCodes.frx":14C6
      StartingColNumber=   0
      VisibleCols     =   6
      VisibleRows     =   10
      ScrollBarTrack  =   1
   End
   Begin EditLib.fpText fptxtLiabAcct 
      Height          =   390
      Left            =   3480
      TabIndex        =   1
      ToolTipText     =   "Enter a liability number for this deduction."
      Top             =   2835
      Width           =   4380
      _Version        =   196608
      _ExtentX        =   7726
      _ExtentY        =   688
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
      CharValidationText=   ""
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
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   615
      Left            =   8520
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   "Press ESC to exit this screen."
      Top             =   2159
      Width           =   2055
      _Version        =   131072
      _ExtentX        =   3625
      _ExtentY        =   1085
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
      ButtonDesigner  =   "frmDeductionCodes.frx":1B0D
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSaveContinue 
      Height          =   615
      Left            =   8520
      TabIndex        =   17
      TabStop         =   0   'False
      ToolTipText     =   "Press F10 to save the last entry but leave the screen open to allow further editing."
      Top             =   2819
      Width           =   2055
      _Version        =   131072
      _ExtentX        =   3625
      _ExtentY        =   1085
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
      ButtonDesigner  =   "frmDeductionCodes.frx":1CE9
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSaveExit 
      Height          =   615
      Left            =   8520
      TabIndex        =   18
      TabStop         =   0   'False
      ToolTipText     =   "Press F11 to exit this screen after saving all data on this screen."
      Top             =   3494
      Width           =   2055
      _Version        =   131072
      _ExtentX        =   3625
      _ExtentY        =   1085
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
      ButtonDesigner  =   "frmDeductionCodes.frx":1ED3
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdClear 
      Height          =   375
      Left            =   3045
      TabIndex        =   19
      TabStop         =   0   'False
      ToolTipText     =   "Press to delete data from the fields above."
      Top             =   3795
      Width           =   3030
      _Version        =   131072
      _ExtentX        =   5345
      _ExtentY        =   661
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
      ButtonDesigner  =   "frmDeductionCodes.frx":20B8
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Data Entry"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2295
      TabIndex        =   10
      Top             =   1890
      Width           =   3735
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   375
      Left            =   6525
      Top             =   4440
      Width           =   3735
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   2295
      Left            =   600
      Top             =   1965
      Width           =   10455
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   8067.922
      X2              =   8067.922
      Y1              =   1915.127
      Y2              =   4136.285
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H008F8265&
      Caption         =   "MED"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   6240
      TabIndex        =   15
      Top             =   3435
      Width           =   630
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H008F8265&
      Caption         =   "SOC"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   4395
      TabIndex        =   14
      Top             =   3435
      Width           =   690
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H008F8265&
      Caption         =   "SWT"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2550
      TabIndex        =   13
      Top             =   3435
      Width           =   750
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H008F8265&
      Caption         =   "FWT"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   750
      TabIndex        =   12
      Top             =   3435
      Width           =   735
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H008F8265&
      Caption         =   "Liability Acct Number"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   795
      TabIndex        =   11
      Top             =   2970
      Width           =   2385
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H008F8265&
      Caption         =   "Description"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   750
      TabIndex        =   9
      Top             =   2520
      Width           =   1425
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Withholding Exemptions"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   6525
      TabIndex        =   8
      Top             =   4440
      Width           =   3735
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   1095
      Index           =   1
      Left            =   1605
      Top             =   450
      Width           =   8655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Deduction Codes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2820
      TabIndex        =   7
      Top             =   750
      Width           =   6015
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   1215
      Left            =   1605
      Top             =   330
      Width           =   8655
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
Attribute VB_Name = "frmDeductionCodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Dim ClickFlag As Boolean
Dim RowFlag As Integer
Dim PriorRowNum As Integer
Dim ContinueFlag As Boolean
Dim SaveAndExitFlag As Boolean
Dim JustExitFlag As Boolean
Dim BadDataFlag As Boolean
Dim NoDataFlag As Boolean
Dim changeFlag As Boolean
Dim PriorDesc As String
Dim PriorLibNum As String
Dim PriorFWT As String
Dim PriorSWT As String
Dim PriorSOC As String
Dim ClearFieldsFlag As Boolean
Dim PriorMED As String
Dim FirstTimeThru As Boolean
'Dim DCcnt As Integer '8/19 commented out
Dim DupFlag As Boolean
Dim BadGLNum As Boolean
Dim CheckGLFlag As Boolean
Dim SaveAnyWayFlag As Boolean
Dim GLFundLen%, GLAcctLen%, GLDetLen% '1/18/05
Dim Split As Boolean

Private Sub cmdClear_Click()
'This sub was designed to allow the user to clear all fields
'after having been editing existing entries...if the user
'has been editing then in order to enter a brand new entry
'they would have to highlight the last empty row and then
'enter new data...this way the program knows this is a new entry
'and automatically saves new data in the next empty row
'before entering new data we must check to see if the last
'entry was saved properly and if not give the user the option
'to save or not to save
  Dim DedCodeFileHandle As Integer, x As Integer, FileLen As Integer
  Dim DedCodeFileRec As DedCodeRecType
  ClearFieldsFlag = True 'if we're here it's because
  'the clear flag command button was used
  
  'store the current data temporarily
  PriorDesc = QPTrim$(fpDescription.Text)
  PriorLibNum = QPTrim$(fptxtLiabAcct.Text) 'comboLiabilityNum.Text)
  PriorFWT = QPTrim$(fpcomboFWT.Text)
  PriorSWT = QPTrim$(fpcomboSWT.Text)
  PriorSOC = QPTrim$(fpcomboSOC.Text)
  PriorMED = QPTrim$(fpcomboMED.Text)
  'if there's nothing there then exit
  If PriorDesc = "" And PriorLibNum = "" And PriorFWT = "" And PriorSWT = "" And PriorSOC = "" And PriorMED = "" Then Exit Sub
  'take a look to see if something has been edited
  Call CheckForChanges
  
  'OK we found a change
  If changeFlag = True Then
    If MsgBox("Your last edit was not saved. Do you want to save it?", vbYesNo) = vbYes Then
      Call cmdSaveContinue_Click
      'if we don't exit here if there is a problem
      'with data (i.e. empty field), the program clears and resets
      'all fields without giving the user the chance
      'to correct mistakes
      If BadDataFlag = True Then Exit Sub 'if BadDataFlag is true then
      'the data was not saved...cmdSaveContinue kicks it out
      ClearFieldsFlag = False
    End If
  End If
  'this is where we know it's OK to clear the fields
  If ClearFieldsFlag = True Then
     fpDescription.Text = ""
     fptxtLiabAcct.Text = ""
     fpcomboFWT.Text = ""
     fpcomboSWT.Text = ""
     fpcomboSOC.Text = ""
     fpcomboMED.Text = ""
  End If
  'reset all flags (except the ClearFieldsFlag) when cleared
  BadDataFlag = False
  ContinueFlag = False
  NoDataFlag = False
  SaveAndExitFlag = False
  JustExitFlag = False
  DupFlag = False
  changeFlag = False
  ClickFlag = False
  SaveAnyWayFlag = False
  'Set RowFlag to accept a new entry
  OpenDedCodeFile DedCodeFileHandle
  FileLen = LOF(DedCodeFileHandle) / Len(DedCodeFileRec)
  Close DedCodeFileHandle
  RowFlag = FileLen + 1
  fpDescription.SetFocus
End Sub

Private Sub cmdList_Click()
  frmGLPickList.Show vbModal
End Sub

Private Sub comboLiabilityNum_KeyDown(KeyCode As Integer, Shift As Integer)
End Sub

Private Sub comboLiabilityNum_LostFocus()
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'  this causes all characters to be capitalized
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub

Private Sub fpcomboFWT_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcomboFWT.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcomboFWT.ListIndex = -1
  End If
  If fpcomboFWT.ListDown = False Then
    If KeyCode = vbKeyUp Then
      SendKeys "+{Tab}"
      KeyCode = 0
    ElseIf KeyCode = vbKeyDown Then
      SendKeys "{Tab}"
      KeyCode = 0
    End If
  End If
End Sub

Private Sub fpcomboFWT_KeyPress(KeyAscii As Integer)
'  this causes all characters to be capitalized
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub

Private Sub fpcomboFWT_LostFocus()
  fpcomboFWT.Action = ActionClearSearchBuffer
End Sub

Private Sub fpcomboMED_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcomboMED.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcomboMED.ListIndex = -1
  End If
  If fpcomboMED.ListDown = False Then
    If KeyCode = vbKeyDown Then
      fpDescription.SetFocus
      KeyCode = 0
    ElseIf KeyCode = vbKeyUp Then
      SendKeys "+{Tab}"
      KeyCode = 0
    End If
  End If
End Sub

Private Sub fpcomboMED_KeyPress(KeyAscii As Integer)
'  this causes all characters to be capitalized
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub

Private Sub fpcomboMED_LostFocus()
  fpcomboMED.Action = ActionClearSearchBuffer

End Sub

Private Sub fpcomboSOC_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcomboSOC.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcomboSOC.ListIndex = -1
  End If
  If fpcomboSOC.ListDown = False Then
    If KeyCode = vbKeyUp Then
      SendKeys "+{Tab}"
      KeyCode = 0
    ElseIf KeyCode = vbKeyDown Then
      SendKeys "{Tab}"
      KeyCode = 0
    End If
  End If

End Sub

Private Sub fpcomboSOC_KeyPress(KeyAscii As Integer)
'  this causes all characters to be capitalized
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub

Private Sub fpcomboSOC_LostFocus()
  fpcomboSOC.Action = ActionClearSearchBuffer

End Sub

Private Sub fpcomboSWT_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcomboSWT.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcomboSWT.ListIndex = -1
  End If
  If fpcomboSWT.ListDown = False Then
    If KeyCode = vbKeyUp Then
      SendKeys "+{Tab}"
      KeyCode = 0
    ElseIf KeyCode = vbKeyDown Then
      SendKeys "{Tab}"
      KeyCode = 0
    End If
  End If

End Sub

Private Sub fpcomboSWT_KeyPress(KeyAscii As Integer)
'  this causes all characters to be capitalized
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%X"
      Call cmdExit_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%C"
      Call cmdSaveContinue_Click
      KeyCode = 0
    Case vbKeyF11:
      SendKeys "%E"
      Call cmdSaveExit_Click
      KeyCode = 0
    Case vbKeyF12:
      SendKeys "%G"
      Call cmdList_Click
      KeyCode = 0
    Case Else:
  End Select

End Sub
Private Sub Form_Load()

  Dim ScrWidth As Long
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  FirstTimeThru = True 'when we save and continue the
  'combo boxes do not need to be reloaded when the form
  'returns to a clean screen by going back thru LoadDCFile
  'so FirstTimeThru turns false if something has been saved or
  'if a row has been double clicked for edit...
  Call FixSpread
  Me.HelpContextID = hlpDeductionCode
  LoadDCFile
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
   Dim A%, b%, C%, D%, E%, f%
   
   A = Len(QPTrim$(fpDescription.Text))
   b = Len(QPTrim$(fptxtLiabAcct.Text))
   C = Len(QPTrim$(fpcomboFWT.Text))
   D = Len(QPTrim$(fpcomboSWT.Text))
   E = Len(QPTrim$(fpcomboSOC.Text))
   f = Len(QPTrim$(fpcomboMED.Text))
   If A + b + C + D + E + f = 0 Then GoTo EmptyFields 'nothing there so exit
   
   'user wants to bail out so set JustExitFlag
   JustExitFlag = True
   Call CheckForChanges
   If changeFlag = True Then
     If MsgBox("A change has been made. Do you want to exit without saving it?", vbYesNo) = vbNo Then
       'oops...no don't exit yet
       If FirstTimeThru = True Or ClearFieldsFlag = True Then
         'this justs sets the focus back to where a change was made
         'when we know the fields were empty to start with
         If A <> 0 Then fpDescription.SetFocus
         If b <> 0 Then fptxtLiabAcct.SetFocus
         If C <> 0 Then fpcomboFWT.SetFocus
         If D <> 0 Then fpcomboSWT.SetFocus
         If E <> 0 Then fpcomboSOC.SetFocus
         If f <> 0 Then fpcomboMED.SetFocus
       Else
         fpDescription.SetFocus
       End If
       Exit Sub
     End If
   End If
EmptyFields:
   KillFile "prdeductions.dat"
   frmControlFileMaint.Show
   DoEvents
   Unload frmDeductionCodes
End Sub

Private Sub cmdSaveContinue_Click()
   Dim DedCodeFileHandle As Integer, x As Integer, FileLen As Integer
   Dim DedCodeFileRec As DedCodeRecType
   Dim RowCount As Integer
   Dim A$, b$, C$, D$, E$, f$
   Dim TempRowFlag As Integer
   Dim DedAlert As TempDedAlertType
   Dim DHandle As Integer
   Dim NumOfDedAlerts As Integer
   Dim TotGLLen As Integer
   Dim GLIsRightLen As Boolean
   
   On Error GoTo ERRORSTUFF
   'this sub is designed to allow multiple saves without
   'exiting the screen
   If SaveAndExitFlag = True Then GoTo SkipGLChk 'GL checked in
   'SaveandExit
   If SaveAnyWayFlag = True Then
     SaveAnyWayFlag = False
     GoTo SkipGLChk 'GL already checked and user knows it's
     'wrong at this point
   End If
   GLIsRightLen = True '1/18/2005
   TotGLLen = GLFundLen% + GLAcctLen% + GLDetLen '1/18/05
   If Split = True Then '1/18/05 entire if statement added
     If Len(ReplaceString(fptxtLiabAcct.Text, "-", "")) = TotGLLen Then
        frmMessageWOpts.Label1.Caption = "Since the split accounting method is being used the liability account number should not include the fund number. Please examine the liability number entered to make sure it only includes the account number and the detail number. Do you wish to save this number anyway?"
        frmMessageWOpts.Label1.Top = 650
        frmMessageWOpts.cmdCont.Text = "F10 Save Anyway"
        frmMessageWOpts.cmdExit.Text = "ESC Review"
        frmMessageWOpts.Show vbModal
        If frmMessageWOpts.fptxtChoice.Text = "abort" Then
          Unload frmMessageWOpts
          Close
          fptxtLiabAcct.SetFocus
          Exit Sub
        Else
          GLIsRightLen = False '1/18/05
          fptxtLiabAcct.SetFocus
          MainLog ("User warned that the liability account number entered for deduction " + QPTrim$(fpDescription.Text) + " as " + QPTrim$(fptxtLiabAcct.Text) + " is of the wrong length for a split account. User chose to save this number anyway.")
        End If
      End If
  Else
     If Len(ReplaceString(fptxtLiabAcct.Text, "-", "")) <> TotGLLen Then
        frmMessageWOpts.Label1.Caption = "Since the non-split accounting method is being used the liability account number should include the fund number, account number and the detail number. Please examine the liability number entered to make sure it includes all three numbers. Do you wish to save this number anyway?"
        frmMessageWOpts.Label1.Top = 650
        frmMessageWOpts.cmdCont.Text = "F10 Save Anyway"
        frmMessageWOpts.cmdExit.Text = "ESC Review"
        frmMessageWOpts.Show vbModal
        If frmMessageWOpts.fptxtChoice.Text = "abort" Then
          Unload frmMessageWOpts
          Close
          fptxtLiabAcct.SetFocus
          Exit Sub
        Else
          GLIsRightLen = False '1/18/05
          fptxtLiabAcct.SetFocus
          MainLog ("User warned that the liability account number entered for deduction " + QPTrim$(fpDescription.Text) + " as " + QPTrim$(fptxtLiabAcct.Text) + " is of the wrong length for a non-split account. User chose to save this number anyway.")
        End If
      End If
   End If
   
   If CheckGLFlag = True And GLIsRightLen = True Then 'the data to set this flag is set in the system
   'interface screen 'added GLIsRightLen on 1/18/05 after inserting the GL length trap
     Call CheckForValidWHNum
   End If
   If BadGLNum = True Then 'CheckForValidWHNum could not
   'find a valid match so exit sub
     BadGLNum = False
     Exit Sub
   End If
SkipGLChk:
   'Since there can be any number of deduction code descriptions we now
   'have inserted a duplicate description name check to make sure all
   'descriptions are unique
   If DescInUseCheck(fpDescription.Text, PriorRowNum) = True Then
      MsgBox "This Description is already in use. Please choose another Description or press Exit to escape."
      fpDescription.SetFocus
      DupFlag = True
      Exit Sub
   End If
   'Looking for empty required fields here
   A = Len(QPTrim$(fpDescription.Text))
   b = Len(QPTrim$(fptxtLiabAcct.Text))
   C = Len(QPTrim$(fpcomboFWT.Text))
   D = Len(QPTrim$(fpcomboSWT.Text))
   E = Len(QPTrim$(fpcomboSOC.Text))
   f = Len(QPTrim$(fpcomboMED.Text))
   'If the user wants to save and then exit the screen
   'we do not turn on the ContinueFlag
   If SaveAndExitFlag = True Then
      ContinueFlag = False
   Else
      ContinueFlag = True
   End If
   'because more than one field needs to have the focus set each
   'If statement below must handle code individually instead of sending
   'error traps to a goto line as done in the next series of If
   'statements
   If A <> 0 Then
      If b = 0 Then
         fptxtLiabAcct.SetFocus
         MsgBox "All fields must be filled out if the Description field is filled out."
         BadDataFlag = True
         GoTo ExitTran
      End If
      If C = 0 Then
         fpcomboFWT.SetFocus
         MsgBox "All fields must be filled out if the Description field is filled out."
         BadDataFlag = True
         GoTo ExitTran
      End If
      If D = 0 Then
         fpcomboSWT.SetFocus
         MsgBox "All fields must be filled out if the Description field is filled out."
         BadDataFlag = True
         GoTo ExitTran
      End If
      If E = 0 Then
         fpcomboSOC.SetFocus
         MsgBox "All fields must be filled out if the Description field is filled out."
         BadDataFlag = True
         GoTo ExitTran
      End If
      If f = 0 Then
         fpcomboMED.SetFocus
         MsgBox "All fields must be filled out if the Description field is filled out."
         BadDataFlag = True
         GoTo ExitTran
      End If
   End If
   'If nothing has been entered and the user tries to save
   'a message box alerts them to this because we do not want
   'to save empty fields
   If A + b + C + D + E + f = 0 Then
      'NoDataFlag is set to 1 so if the user wanted to
      'exit the screen after this save then that procedure
      'will behave with the save response
      NoDataFlag = True
      If SaveAndExitFlag = False Then
         MsgBox "No new or edited data to save"
      End If
      Exit Sub
   End If
   'if the description field is empty and any other field
   'is not empty this if statement traps this error
   If A = 0 Then
      If b <> 0 Then GoTo BadDataEntry
      If C <> 0 Then GoTo BadDataEntry
      If D <> 0 Then GoTo BadDataEntry
      If E <> 0 Then GoTo BadDataEntry
      If f <> 0 Then GoTo BadDataEntry
      Else: GoTo EntryDataOK
BadDataEntry:
       MsgBox "Please complete the Description field, double click an existing account to edit or delete all fields to continue."
       BadDataFlag = True
       fpDescription.SetFocus
       GoTo ExitTran
   End If
   
EntryDataOK:
   OpenDedCodeFile DedCodeFileHandle
   FileLen = LOF(DedCodeFileHandle) / Len(DedCodeFileRec)
   If FileLen = 0 Then 'first ever save
      vaSpreadDeductionCodes.Col = 1
      vaSpreadDeductionCodes.Row = 1
      DedCodeFileRec.DCDESC1 = QPTrim$(fpDescription.Text)
      vaSpreadDeductionCodes.Col = 2
      vaSpreadDeductionCodes.Row = 1
      DedCodeFileRec.DCACCT1 = QPTrim$(fptxtLiabAcct.Text)
      vaSpreadDeductionCodes.Col = 3
      vaSpreadDeductionCodes.Row = 1
      DedCodeFileRec.DCFWT1 = QPTrim$(fpcomboFWT.Text)
      vaSpreadDeductionCodes.Col = 4
      vaSpreadDeductionCodes.Row = 1
      DedCodeFileRec.DCSWT1 = QPTrim$(fpcomboSWT.Text)
      vaSpreadDeductionCodes.Col = 5
      vaSpreadDeductionCodes.Row = 1
      DedCodeFileRec.DCSOC1 = QPTrim$(fpcomboSOC.Text)
      vaSpreadDeductionCodes.Col = 6
      vaSpreadDeductionCodes.Row = 1
      DedCodeFileRec.DCMED1 = QPTrim$(fpcomboMED.Text)
      Put DedCodeFileHandle, 1, DedCodeFileRec
      Close DedCodeFileHandle
      GoTo ClickSave
   End If
   'ClickFlag denotes we are here because the user double clicked a row
   'to edit it
   'If RowFlag is not revalued to the row that was just changed
   'the change takes place but in the row that is now in focus
   'causing data to be saved to the wrong row...if an edit by
   'double clicking is what we want and a change has been made
   'in existing data then we want the existing row's data saved
   'in the same row with updated data ...so we have to make
   'sure the updated data is not saved in a different row with
   'the old data still residing in the old row where the update
   'should be
   If changeFlag = True Then
      TempRowFlag = RowFlag 'save current row setting
      RowFlag = PriorRowNum 'reset row to the one that was changed
   End If
   'if ClearFieldsFlag is true then we do not want to save anything
   'until we find the first empty row...
   If ClearFieldsFlag = True And RowFlag > FileLen Then GoTo NonEditEntry
   If ClickFlag = True Then 'save row that was double clicked for edit
      Get DedCodeFileHandle, RowFlag, DedCodeFileRec
      vaSpreadDeductionCodes.Col = 1
      vaSpreadDeductionCodes.Row = RowFlag
      DedCodeFileRec.DCDESC1 = QPTrim$(fpDescription.Text)
      vaSpreadDeductionCodes.Col = 2
      vaSpreadDeductionCodes.Row = RowFlag
      DedCodeFileRec.DCACCT1 = QPTrim$(fptxtLiabAcct.Text)
      vaSpreadDeductionCodes.Col = 3
      vaSpreadDeductionCodes.Row = RowFlag
      DedCodeFileRec.DCFWT1 = QPTrim$(fpcomboFWT.Text)
      vaSpreadDeductionCodes.Col = 4
      vaSpreadDeductionCodes.Row = RowFlag
      DedCodeFileRec.DCSWT1 = QPTrim$(fpcomboSWT.Text)
      vaSpreadDeductionCodes.Col = 5
      vaSpreadDeductionCodes.Row = RowFlag
      DedCodeFileRec.DCSOC1 = QPTrim$(fpcomboSOC.Text)
      vaSpreadDeductionCodes.Col = 6
      vaSpreadDeductionCodes.Row = RowFlag
      DedCodeFileRec.DCMED1 = QPTrim$(fpcomboMED.Text)
      Put DedCodeFileHandle, RowFlag, DedCodeFileRec
      Close DedCodeFileHandle
      'change RowFlag back to original value so we can continue
      'editing with the new row now needing editing
      If changeFlag = True Then
        changeFlag = False
        RowFlag = TempRowFlag
      Else
        RowFlag = 0
      End If
      GoTo ClickSave
   End If
   'save data from fields at top of form
NonEditEntry:
   For x = 1 To 50
      vaSpreadDeductionCodes.Col = 1
      vaSpreadDeductionCodes.Row = x
      If Len(QPTrim$(vaSpreadDeductionCodes.Value)) = 0 Then
      'save in the next empty row
         RowCount = x
         Exit For
      End If
   Next
   If x > 50 Then
     MsgBox "You have reached the maximum allowable deductions"
     Close DedCodeFileHandle
     Exit Sub
   End If
   OpenDedAlertFile DHandle
   NumOfDedAlerts = LOF(DHandle) / Len(DedAlert)
   
   Get DedCodeFileHandle, RowCount, DedCodeFileRec
   vaSpreadDeductionCodes.Col = 1
   vaSpreadDeductionCodes.Row = RowCount
   
   DedAlert.DCDESC1 = QPTrim$(fpDescription.Text)
   DedAlert.Number = RowCount
   Put DHandle, NumOfDedAlerts + 1, DedAlert
   Close DHandle
   
   DedCodeFileRec.DCDESC1 = QPTrim$(fpDescription.Text)
   vaSpreadDeductionCodes.Col = 2
   vaSpreadDeductionCodes.Row = RowCount
   DedCodeFileRec.DCACCT1 = QPTrim$(fptxtLiabAcct.Text)
   vaSpreadDeductionCodes.Col = 3
   vaSpreadDeductionCodes.Row = RowCount
   DedCodeFileRec.DCFWT1 = QPTrim$(fpcomboFWT.Text)
   vaSpreadDeductionCodes.Col = 4
   vaSpreadDeductionCodes.Row = RowCount
   DedCodeFileRec.DCSWT1 = QPTrim$(fpcomboSWT.Text)
   vaSpreadDeductionCodes.Col = 5
   vaSpreadDeductionCodes.Row = RowCount
   DedCodeFileRec.DCSOC1 = QPTrim$(fpcomboSOC.Text)
   vaSpreadDeductionCodes.Col = 6
   vaSpreadDeductionCodes.Row = RowCount
   DedCodeFileRec.DCMED1 = QPTrim$(fpcomboMED.Text)
   Put DedCodeFileHandle, RowCount, DedCodeFileRec
   Close DedCodeFileHandle
ClickSave: 'jump here if editing was done by double clicking
   'a row with existing data because we don't want a new row
   'added...we just want to save the new data on an existing row
   BadDataFlag = False
   'SaveAndExit command button used so we don't need anything between here
   'and ExitTran
   
   If SaveAndExitFlag = True Then GoTo ExitTran ' this save is coming from the
   'the exit and save routine that has already performed everything
   'from here to ExitTran
   MsgBox "Your Information has been saved.", vbOKOnly
   Call LoadDCFile
   fpDescription.SetFocus
   FirstTimeThru = False 'once data has been saved and accepted then
   'first time thru is always false
ExitTran:
   MainLog ("Deduction code data was saved.")
  Exit Sub
   
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmDeductionCodes", "cmdSaveContinue", Erl)
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

Private Sub cmdSaveExit_Click()
   Dim TotGLLen As Integer
   Dim GLIsRightLen As Boolean
   
   On Error GoTo ERRORSTUFF
   GLIsRightLen = True '1/18/2005
   TotGLLen = GLFundLen% + GLAcctLen% + GLDetLen '1/18/05
   If Split = True Then '1/18/05 entire if statement added
     If Len(ReplaceString(fptxtLiabAcct.Text, "-", "")) = TotGLLen Then
        frmMessageWOpts.Label1.Caption = "Since the split accounting method is being used the liability account number should not include the fund number. Please examine the liability number entered to make sure it does not include the fund number. Do you wish to save this number anyway?"
        frmMessageWOpts.Label1.Top = 650
        frmMessageWOpts.cmdCont.Text = "F10 Save Anyway"
        frmMessageWOpts.cmdExit.Text = "ESC Review"
        frmMessageWOpts.Show vbModal
        If frmMessageWOpts.fptxtChoice.Text = "abort" Then
          Unload frmMessageWOpts
          Close
          fptxtLiabAcct.SetFocus
          Exit Sub
        Else
          GLIsRightLen = False '1/18/05
          fptxtLiabAcct.SetFocus
        End If
      End If
  Else
     If Len(ReplaceString(fptxtLiabAcct.Text, "-", "")) <> TotGLLen Then
        frmMessageWOpts.Label1.Caption = "Since the non-split accounting method is being used the liability account number should include the fund number, account number and the detail number. Please examine the liability number entered to make sure it includes all three numbers. Do you wish to save this number anyway?"
        frmMessageWOpts.Label1.Top = 550
        frmMessageWOpts.cmdCont.Text = "F10 Save Anyway"
        frmMessageWOpts.cmdExit.Text = "ESC Review"
        frmMessageWOpts.Show vbModal
        If frmMessageWOpts.fptxtChoice.Text = "abort" Then
          Unload frmMessageWOpts
          Close
          fptxtLiabAcct.SetFocus
          Exit Sub
        Else
          GLIsRightLen = False '1/18/05
          fptxtLiabAcct.SetFocus
        End If
      End If
   End If
   
   If CheckGLFlag = True And GLIsRightLen = True Then 'the data to set this flag is set in the system
   'interface screen 'added GLIsRightLen on 1/18/05 after inserting the GL length trap
     Call CheckForValidWHNum
   End If
   
   If BadGLNum = True Then
     BadGLNum = False
     Exit Sub
   End If
   
   SaveAndExitFlag = True
   Call cmdSaveContinue_Click
   If DupFlag = True Then 'DupFlag comes up if the user has tried to save
   'data with a description that's already been saved
      DupFlag = False
      Exit Sub
   End If
   If BadDataFlag = True Then 'when BadDataFlag is set to true a
   'message box alerts the user to the problem and the program
   'focuses on the box that needs attention...so we can't exit until the
   'problem is fixed
      SaveAndExitFlag = False
      GoTo ExitTran
   End If
      If NoDataFlag = True Then
      MsgBox "No new or edited data to save"
      GoTo NoData
   End If
   MsgBox "Your Information has been saved.", vbOKOnly
NoData:
   SaveAndExitFlag = False
   BadDataFlag = False
   frmControlFileMaint.Show
   DoEvents
   Unload frmDeductionCodes
ExitTran:
  KillFile "prdeductions.dat"
  Exit Sub
   
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmDeductionCodes", "cmdSaveExit", Erl)
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

Private Sub LoadDCFile()
   'all fields in the upper block must be cleared for the
   'ClickFlag to work properly
   Dim DedCodeFileHandle As Integer, x As Integer, FileLen As Integer
   Dim DedCodeFileRec As DedCodeRecType
   Dim SysRec As RegDSysFileRecType
   Dim SysHandle As Integer
   Dim One As Integer
   Dim AHandle As Integer
  
   One = 1 '01/18/05
   AHandle = FreeFile '01/18/05
   Open "prdeductions.dat" For Output As AHandle '01/18/05
   Print #AHandle, One '01/18/05
   Close AHandle '01/18/05
   
   OpenSysFile SysHandle
   Get SysHandle, 1, SysRec
   Close SysHandle
   
   Split = False
   If SysRec.SplitFlag = "Y" Then
     Split = True
   End If
   Call GetAcctStruct(CurrCitiPath, GLFundLen%, GLAcctLen%, GLDetLen%) '01/18/05
   
   If QPTrim$(CurrCitiPath) = "" Then cmdList.Visible = False '11/27/02 'changed to CurrCitiPath 8/12/04
   If QPTrim$(SysRec.GLCheckYN) = "N" Then
     CheckGLFlag = False
   Else
     CheckGLFlag = True
   End If
   
   SaveAnyWayFlag = False
   NoDataFlag = False
   ClickFlag = False
   
   fpDescription.Text = ""
   fptxtLiabAcct.Text = ""
   fpcomboFWT.Text = ""
   fpcomboSWT.Text = ""
   fpcomboSOC.Text = ""
   fpcomboMED.Text = ""
   'load the combo boxes in the upper block..if we are reloading from
   'the Save and Continue button then we don't need to reload the combo
   'boxes because the form was never unloaded
   If FirstTimeThru = True Then 'if FirstTimeThru is false then
   'do not load the combo boxes again because everytime this
   'sub is called the combo boxes keep adding to whatever is already loaded
     fpcomboFWT.Clear
     fpcomboSWT.Clear
     fpcomboSOC.Clear
     fpcomboMED.Clear
     fpcomboFWT.AddItem "Y"
     fpcomboFWT.AddItem "N"
     fpcomboSWT.AddItem "Y"
     fpcomboSWT.AddItem "N"
     fpcomboSOC.AddItem "Y"
     fpcomboSOC.AddItem "N"
     fpcomboMED.AddItem "Y"
     fpcomboMED.AddItem "N"
   End If
   OpenDedCodeFile DedCodeFileHandle
   FileLen = LOF(DedCodeFileHandle) / Len(DedCodeFileRec)
   'This for loop loads all data stored on file plus it loads "N" in
   'the FWT, SWT, SOC and MED fields if no description is on that row
   For x = 1 To FileLen
     Get DedCodeFileHandle, x, DedCodeFileRec
     'load form info
     vaSpreadDeductionCodes.Col = 1
     vaSpreadDeductionCodes.Row = x
     vaSpreadDeductionCodes.Text = QPTrim$(DedCodeFileRec.DCDESC1)
     vaSpreadDeductionCodes.Col = 2
     vaSpreadDeductionCodes.Row = x
     vaSpreadDeductionCodes.Text = QPTrim$(DedCodeFileRec.DCACCT1)
     vaSpreadDeductionCodes.Col = 3
     vaSpreadDeductionCodes.Row = x
     vaSpreadDeductionCodes.Text = QPTrim$(DedCodeFileRec.DCFWT1)
     vaSpreadDeductionCodes.Col = 4
     vaSpreadDeductionCodes.Row = x
     vaSpreadDeductionCodes.Text = QPTrim$(DedCodeFileRec.DCSWT1)
     vaSpreadDeductionCodes.Col = 5
     vaSpreadDeductionCodes.Row = x
     vaSpreadDeductionCodes.Text = QPTrim$(DedCodeFileRec.DCSOC1)
     vaSpreadDeductionCodes.Col = 6
     vaSpreadDeductionCodes.Text = x
     vaSpreadDeductionCodes.Value = QPTrim$(DedCodeFileRec.DCMED1)
   Next
   
   Close DedCodeFileHandle
   BadDataFlag = False
End Sub

Private Sub fpcomboSWT_LostFocus()
  fpcomboSWT.Action = ActionClearSearchBuffer

End Sub

Private Sub fpDescription_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyUp Then
    fpcomboMED.SetFocus
    KeyCode = 0
  ElseIf KeyCode = vbKeyDown Then
    fptxtLiabAcct.SetFocus
    KeyCode = 0
  End If
End Sub

Private Sub fptxtLiabAcct_DblClick(Button As Integer)
  fptxtLiabAcct = Clipboard.GetText
End Sub

Private Sub mnuExit_Click()
  Call cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
  MainLog ("Deduction code control screen printed.")
End Sub

Private Sub vaSpreadDeductionCodes_DblClick(ByVal Col As Long, ByVal Row As Long)
  Dim DedCodeFileHandle As Integer, x As Integer, FileLen As Integer
  Dim DedCodeFileRec As DedCodeRecType
  'save all data before the doubleclick removed them
  OpenDedCodeFile DedCodeFileHandle
  FileLen = LOF(DedCodeFileHandle) / Len(DedCodeFileRec)
  Close DedCodeFileHandle
  
  If Row > FileLen Then
    MsgBox "Empty rows cannot be edited"
    Exit Sub
  End If

  FirstTimeThru = False 'once a row is double clicked then
  'first time thru will always be false
'  DCcnt = DCcnt + 1
'  If DCcnt > 1 Then FirstTimeThru = False
  changeFlag = False
  RowFlag = Row
  'the next 6 lines saves data currently loaded on the screen
  PriorDesc = QPTrim$(fpDescription.Text)
  PriorLibNum = QPTrim$(fptxtLiabAcct.Text)
  PriorFWT = QPTrim$(fpcomboFWT.Text)
  PriorSWT = QPTrim$(fpcomboSWT.Text)
  PriorSOC = QPTrim$(fpcomboSOC.Text)
  PriorMED = QPTrim$(fpcomboMED.Text)
  'if ClearFieldsFlag is true we've already checked for changes
  If ClearFieldsFlag = True Then
     ClearFieldsFlag = False
     GoTo NoChangeCheck
  End If
  'the user is double clicking without saving the existing data
  'put there by double clicking...the next if statement traps for
  'changes in the first edit that the user might want to save
  'before reloading with new double clicked data
  If ClickFlag = True Then
    Call CheckForChanges
      If changeFlag = True Then
        If MsgBox("Your last edit was not saved. Do you want to save it?", vbYesNo) = vbYes Then
          Call cmdSaveContinue_Click
          changeFlag = False
        End If
      End If
   End If
'This routine allows the user to double click a specific row
'that places that row's data in the edit fields
NoChangeCheck:
   ClickFlag = True
   'load the fields in the upper block with the data for
   'the file numbered as Row (the row double clicked)
   vaSpreadDeductionCodes.Col = 1
   vaSpreadDeductionCodes.Row = Row
   fpDescription.Text = QPTrim$(vaSpreadDeductionCodes.Value)
   vaSpreadDeductionCodes.Col = 2
   vaSpreadDeductionCodes.Row = Row
   fptxtLiabAcct.Text = QPTrim$(vaSpreadDeductionCodes.Value)
   vaSpreadDeductionCodes.Col = 3
   vaSpreadDeductionCodes.Row = Row
   fpcomboFWT.Text = QPTrim$(vaSpreadDeductionCodes.Text)
   vaSpreadDeductionCodes.Col = 4
   vaSpreadDeductionCodes.Row = Row
   fpcomboSWT.Text = QPTrim$(vaSpreadDeductionCodes.Text)
   vaSpreadDeductionCodes.Col = 5
   vaSpreadDeductionCodes.Row = Row
   fpcomboSOC.Text = QPTrim$(vaSpreadDeductionCodes.Text)
   vaSpreadDeductionCodes.Col = 6
   vaSpreadDeductionCodes.Row = Row
   fpcomboMED.Text = QPTrim$(vaSpreadDeductionCodes.Text)
   PriorRowNum = RowFlag 'saves the current row number and
   'is used in CheckForChanges to compare what is to be saved
   'against what was already on the screen
End Sub

Private Sub CheckForChanges()
'This routine compares data in the row that just lost focus with the data
'that is in the appropriate row in the spreadsheet...if a change
'has been made it will be detected here

'if FirstTimeThru or JustExitFlag are true then we must use the
'current value in the non-spreadsheet (upper) fields because
   changeFlag = False
   If FirstTimeThru = True Or JustExitFlag = True Then PriorDesc = QPTrim$(fpDescription.Text)
   vaSpreadDeductionCodes.Col = 1
   vaSpreadDeductionCodes.Row = PriorRowNum
   'next if statements were done solely to refocus properly
   If QPTrim$(vaSpreadDeductionCodes.Text) <> QPTrim$(PriorDesc) Then
     changeFlag = True
     fpDescription.SetFocus
   End If
   If FirstTimeThru = True Or JustExitFlag = True Then PriorLibNum = QPTrim$(fptxtLiabAcct.Text) 'comboLiabilityNum.Text)
   vaSpreadDeductionCodes.Col = 2
   vaSpreadDeductionCodes.Row = PriorRowNum
   If QPTrim$(vaSpreadDeductionCodes.Text) <> QPTrim$(PriorLibNum) Then
     changeFlag = True
     fptxtLiabAcct.SetFocus
   End If
   If FirstTimeThru = True Or JustExitFlag = True Then PriorFWT = QPTrim$(fpcomboFWT.Text)
   vaSpreadDeductionCodes.Col = 3
   vaSpreadDeductionCodes.Row = PriorRowNum
   If QPTrim$(vaSpreadDeductionCodes.Text) <> QPTrim$(PriorFWT) Then
     changeFlag = True
     fpcomboFWT.SetFocus
   End If
   If FirstTimeThru = True Or JustExitFlag = True Then PriorSWT = QPTrim$(fpcomboSWT.Text)
   vaSpreadDeductionCodes.Col = 4
   vaSpreadDeductionCodes.Row = PriorRowNum
   If QPTrim$(vaSpreadDeductionCodes.Text) <> QPTrim$(PriorSWT) Then
     changeFlag = True
     fpcomboSWT.SetFocus
   End If
   If FirstTimeThru = True Or JustExitFlag = True Then PriorSOC = QPTrim$(fpcomboSOC.Text)
   vaSpreadDeductionCodes.Col = 5
   vaSpreadDeductionCodes.Row = PriorRowNum
   If QPTrim$(vaSpreadDeductionCodes.Text) <> QPTrim$(PriorSOC) Then
     changeFlag = True
     fpcomboSOC.SetFocus
   End If
   If FirstTimeThru = True Or JustExitFlag = True Then PriorMED = QPTrim$(fpcomboMED.Text)
   vaSpreadDeductionCodes.Col = 6
   vaSpreadDeductionCodes.Row = PriorRowNum
   If QPTrim$(vaSpreadDeductionCodes.Text) <> QPTrim$(PriorMED) Then
     changeFlag = True
     fpcomboMED.SetFocus
   End If
'   FirstTimeThru = False
   JustExitFlag = False
End Sub

Private Function DescInUseCheck(Desc As String, ThisRow As Integer) As Boolean
   Dim DedCodeFileHandle As Integer, x As Integer, FileLen As Integer
   Dim DedCodeFileRec As DedCodeRecType
   'we do not want duplicate descriptions used so this function
   'looks at all the existing descriptions and compares the
   'description to be used with what's already on file
   
   If QPTrim$(Desc) = "" Then Exit Function
   
   DescInUseCheck = False
   OpenDedCodeFile DedCodeFileHandle
   FileLen = LOF(DedCodeFileHandle) / Len(DedCodeFileRec)
   For x = 1 To FileLen
      If x = ThisRow Then GoTo RowInEdit 'must skip the row you are
      'now on because it's always a duplicate...duh
      Get DedCodeFileHandle, x, DedCodeFileRec
      If QPTrim$(Desc) = QPTrim$(DedCodeFileRec.DCDESC1) Then 'found a match
         DescInUseCheck = True
         Exit For
      End If
RowInEdit:
   Next x
   Close DedCodeFileHandle
End Function
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
   Dim WHNum As String
   Dim SysRec As RegDSysFileRecType
   Dim SysHandle As Integer
   Dim DoWhatFlag As BadGLNUMOption
   Dim n As Integer
   Dim FundLength As Integer
   Dim AcctLength As Integer
   Dim DetLength As Integer
   
   On Error GoTo ERRORSTUFF
   'no need to verify nothing
   If Len(QPTrim$(fptxtLiabAcct.Text)) = 0 Then Exit Sub '7/26
     
   OpenSysFile SysHandle
   Get SysHandle, 1, SysRec
   Close SysHandle
   
   'GetAcctStruct reads in general ledger numbers and distributes
   'this number into three fields based on three different lengths
'   Call GetAcctStruct(QPTrim$(SysRec.CITIDIR), FundLength, AcctLength, DetLength)
   Call GetAcctStruct(CurrCitiPath, FundLength, AcctLength, DetLength)
   'No need to trap because they are not using our gl package
   If FundLength = 0 And AcctLength = 0 And DetLength = 0 Then
     Exit Sub
   End If
   
   Split$ = QPTrim$(SysRec.SplitFlag)
   
   BadGLNum = False
   'WHNum is the liability number without hyphens
   WHNum = QPTrim$(ReplaceString(fptxtLiabAcct.Text, "-", ""))
   
   If Exist(GetCitiDirFolder + "GLACCT.IDX") Then
     GLIdxNum$ = GetCitiDirFolder + "GLACCT.IDX"
   Else
     MsgBox "No G/L account number validation possible...GLACCT.IDX could not be found."
     Exit Sub
   End If
   
   If Exist(GetCitiDirFolder + "GLACCT.DAT") Then
     GLIDATDesc$ = GetCitiDirFolder + "GLACCT.DAT"
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
   'read gl numbers into an array
   For x = 1 To TotalAccts
     Get GLIdxHandle, x, JGLIdxRec(1)
     DescBuff(x) = JGLIdxRec(1).RecNo
   Next x
   Close GLIdxHandle
   GLDHandle = FreeFile
   Open GLIDATDesc$ For Random As GLDHandle Len = GLDescRecLen
   
   Nextx = 1
   If Split$ = "Y" Then
     GoTo SplitY
   End If
   For x = 1 To TotalAccts
   If DescBuff(x) = 0 Then GoTo DescBuffIs0N
     Get GLDHandle, DescBuff(x), GLDesc(1)
     'go thru each gl number until we find one that
     'matches what has been entered to the screen
     If WHNum = QPTrim$(ReplaceString(GLDesc(1).Num, "-", "")) Then
       Exit For 'OK we found it so no need to look further
     End If
DescBuffIs0N:
     'been all the way through all 5 tests and no match has
     'been found so it is a BadGLNum
     If x = TotalAccts Then
       fptxtLiabAcct.SetFocus
       DoWhatFlag = PromptBadGLNum(Me)
       Select Case DoWhatFlag
       Case BadGLNUMOption.badglExit
         MainLog ("Deductions maintenance screen: user warned that the GL number entered, " + QPTrim$(fptxtLiabAcct.Text) + " was not found in the GL List. User chose to exit without saving.")
         frmControlFileMaint.Show
         DoEvents
         Unload frmDeductionCodes
       Case BadGLNUMOption.badglSave '8/16
         MainLog ("Deductions maintenance screen: user warned that the GL number entered, " + QPTrim$(fptxtLiabAcct.Text) + " was not found in the GL List. User chose to save anyway.")
         SaveAnyWayFlag = True '8/16
         Call cmdSaveContinue_Click '8/16
       Case BadGLNUMOption.badglReturn
         MainLog ("Deductions maintenance screen: user warned that the GL number entered, " + QPTrim$(fptxtLiabAcct.Text) + " was not found in the GL List. User chose to review.")
       Case Else:
          'Do nothing because we don't know about any options except
          'save, review or abandon...used as a placeholder for adding
          'other options at a later date
       End Select
       BadGLNum = True
       Close GLDHandle
       Exit Sub
     End If
   Next x
   GoTo SplitN
SplitY:
     For x = 1 To TotalAccts
        If DescBuff(x) = 0 Then GoTo DescBuffIs0Y
          Get GLDHandle, DescBuff(x), GLDesc(1)
          'go thru each gl number until a match is found
          If WHNum = QPTrim$(ReplaceString(Mid(GLDesc(1).Num, FundLength + 1), "-", "")) Then
            Exit For 'found it so exit
          End If
DescBuffIs0Y:
          If x = TotalAccts Then 'no match found
          'so ask user what to do next
            fptxtLiabAcct.SetFocus
            DoWhatFlag = PromptBadGLNum(Me)
            Select Case DoWhatFlag
            Case BadGLNUMOption.badglExit
              MainLog ("Deductions maintenance screen: user warned that the GL number entered, " + QPTrim$(fptxtLiabAcct.Text) + " was not found in the GL List. User chose to exit without saving.")
              frmControlFileMaint.Show
              DoEvents
              Unload frmDeductionCodes
            Case BadGLNUMOption.badglSave '8/16
              MainLog ("Deductions maintenance screen: user warned that the GL number entered, " + QPTrim$(fptxtLiabAcct.Text) + " was not found in the GL List. User chose to save anyway.")
              SaveAnyWayFlag = True '8/16
              Call cmdSaveContinue_Click '8/16
            Case BadGLNUMOption.badglReturn
              MainLog ("Deductions maintenance screen: user warned that the GL number entered, " + QPTrim$(fptxtLiabAcct.Text) + " was not found in the GL List. User chose to review.")
            Case Else:
              'Do nothing because we don't know about any options except
              'save, review or abandon...used as a placeholder for adding
              'other options at a later date
            End Select
            BadGLNum = True
            Close GLDHandle
            Exit Sub
         End If
     Next x

SplitN:
  Close GLDHandle
  
  Exit Sub
   
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmDeductionCodes", "CheckForValidWHNum", Erl)
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

Private Function CheckDupDedDesc(Text As String) As Boolean
  Dim DedRec As DedCodeRecType
  Dim DedCnt As Integer
  Dim DHandle As Integer
  Dim x As Integer
  
  CheckDupDedDesc = False
  OpenDedCodeFile DHandle
  DedCnt = LOF(DHandle) / Len(DedRec)
  For x = 1 To DedCnt
    Get DHandle, x, DedRec
    If QPTrim$(Text) <> "" Then GoTo DupFound
    If UCase(QPTrim$(Text)) = UCase(QPTrim$(DedRec.DCDESC1)) Then
      CheckDupDedDesc = True
      GoTo DupFound
    End If
  Next x
  
DupFound:
End Function
Private Function FixSpread()
  Dim COne As Integer
  Dim CTwo As Integer
  Dim CThree As Integer
  Dim CFour As Integer
  Dim CFive As Integer
  Dim CSix As Integer
  '-1 means all rows or all columns....0 means headers
    Select Case ScreenW
      Case 1280
        COne = 8
        CTwo = 8
        coladj = 2
        vaSpreadDeductionCodes.RowHeight(-1) = 18
        vaSpreadDeductionCodes.RowHeight(0) = 18
        vaSpreadDeductionCodes.FontSize = 12
      Case 1152
        COne = 2
        CTwo = 2
        coladj = 2
        vaSpreadDeductionCodes.RowHeight(0) = 15
        vaSpreadDeductionCodes.RowHeight(-1) = 15
      Case 1024
        COne = 1
        CTwo = 1
'        coladj = 0.5
'        COne = 10
'        CTwo = 10
        coladj = 4.75
      Case 800
'        COne = 4
'        CTwo = 4
'        coladj = 1
'        vaSpreadDeductionCodes.Font.Size = 10
'        vaSpreadDeductionCodes.RowHeight(-1) = 12.2
      Case Else
       
    End Select
    vaSpreadDeductionCodes.ColWidth(1) = vaSpreadDeductionCodes.ColWidth(1) + COne
    vaSpreadDeductionCodes.ColWidth(2) = vaSpreadDeductionCodes.ColWidth(2) + CTwo
    vaSpreadDeductionCodes.ColWidth(3) = vaSpreadDeductionCodes.ColWidth(3) + coladj
    vaSpreadDeductionCodes.ColWidth(4) = vaSpreadDeductionCodes.ColWidth(4) + coladj
    vaSpreadDeductionCodes.ColWidth(5) = vaSpreadDeductionCodes.ColWidth(5) + coladj
    vaSpreadDeductionCodes.ColWidth(6) = vaSpreadDeductionCodes.ColWidth(6) + coladj

End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      KillFile "prdeductions.dat"
      MainLog ("Payroll.exe terminated via menu bar on frmDeductionCodes.")
      Call Terminate
      End
    End If
  End If
End Sub

