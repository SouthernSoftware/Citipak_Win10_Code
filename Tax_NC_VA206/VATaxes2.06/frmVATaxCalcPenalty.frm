VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmVATaxCalcPenalty 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Penalty Calculation"
   ClientHeight    =   8760
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   11655
   Icon            =   "frmVATaxCalcPenalty.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8760
   ScaleMode       =   0  'User
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcmbPrintOrder 
      Height          =   405
      Left            =   4920
      TabIndex        =   5
      Top             =   6840
      Width           =   3375
      _Version        =   196608
      _ExtentX        =   5953
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
      EditAlignH      =   1
      EditAlignV      =   0
      ColDesigner     =   "frmVATaxCalcPenalty.frx":08CA
   End
   Begin LpLib.fpCombo fpcmbPrintOpt 
      Height          =   405
      Left            =   4920
      TabIndex        =   4
      Top             =   6240
      Width           =   3330
      _Version        =   196608
      _ExtentX        =   5874
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
      BackColor       =   16777215
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
      ColDesigner     =   "frmVATaxCalcPenalty.frx":0BF9
   End
   Begin LpLib.fpCombo fpcmbType 
      Height          =   405
      Left            =   3960
      TabIndex        =   0
      Top             =   1560
      Width           =   3375
      _Version        =   196608
      _ExtentX        =   5953
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
      EditAlignH      =   1
      EditAlignV      =   0
      ColDesigner     =   "frmVATaxCalcPenalty.frx":0F28
   End
   Begin LpLib.fpList fpList 
      Height          =   1080
      Left            =   2220
      TabIndex        =   13
      Top             =   3840
      Width           =   7215
      _Version        =   196608
      _ExtentX        =   12726
      _ExtentY        =   1905
      TextAlias       =   ""
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
      Columns         =   4
      Sorted          =   0
      LineWidth       =   1
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   -1
      ColumnWidthScale=   2
      RowHeight       =   -1
      MultiSelect     =   0
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
      ColDesigner     =   "frmVATaxCalcPenalty.frx":1257
   End
   Begin VB.CheckBox chkPPTRAYN 
      BackColor       =   &H000000FF&
      Caption         =   "Exclude PPTRA Discount from Balance"
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
      Height          =   252
      Left            =   3960
      TabIndex        =   20
      Top             =   5640
      Width           =   3732
   End
   Begin VB.CheckBox ChkAll 
      BackColor       =   &H008F8265&
      Caption         =   "Calculate For Entire Year"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   372
      Left            =   4560
      TabIndex        =   17
      Top             =   2880
      Width           =   2772
   End
   Begin EditLib.fpDateTime fptxtCurrRYear 
      Height          =   372
      Left            =   4140
      TabIndex        =   2
      Top             =   2160
      Width           =   1020
      _Version        =   196608
      _ExtentX        =   1799
      _ExtentY        =   656
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
      Text            =   "2018"
      DateCalcMethod  =   0
      DateTimeFormat  =   5
      UserDefinedFormat=   "yyyy"
      DateMax         =   "20350101"
      DateMin         =   "19800101"
      TimeMax         =   "000000"
      TimeMin         =   "000000"
      TimeString1159  =   ""
      TimeString2359  =   ""
      DateDefault     =   "20010101"
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
      ButtonColor     =   13684944
      AutoMenu        =   0   'False
      StartMonth      =   4
      ButtonAlign     =   0
      BoundDataType   =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
      Height          =   495
      Left            =   7215
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   7800
      Width           =   2055
      _Version        =   131072
      _ExtentX        =   3625
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
      ButtonDesigner  =   "frmVATaxCalcPenalty.frx":15CB
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   492
      Left            =   2388
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   7800
      Width           =   2052
      _Version        =   131072
      _ExtentX        =   3619
      _ExtentY        =   868
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
      ButtonDesigner  =   "frmVATaxCalcPenalty.frx":17AA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdTable 
      Height          =   492
      Left            =   4788
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   7800
      Width           =   2052
      _Version        =   131072
      _ExtentX        =   3619
      _ExtentY        =   868
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
      ButtonDesigner  =   "frmVATaxCalcPenalty.frx":1986
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPrintList 
      Height          =   492
      Left            =   4380
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   5040
      Width           =   2892
      _Version        =   131072
      _ExtentX        =   5101
      _ExtentY        =   868
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
      ButtonDesigner  =   "frmVATaxCalcPenalty.frx":1B66
   End
   Begin EditLib.fpDateTime fptxtCurrPYear 
      Height          =   372
      Left            =   9180
      TabIndex        =   3
      Top             =   2160
      Width           =   1020
      _Version        =   196608
      _ExtentX        =   1799
      _ExtentY        =   656
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
      Text            =   "2018"
      DateCalcMethod  =   0
      DateTimeFormat  =   5
      UserDefinedFormat=   "yyyy"
      DateMax         =   "20350101"
      DateMin         =   "19800101"
      TimeMax         =   "000000"
      TimeMin         =   "000000"
      TimeString1159  =   ""
      TimeString2359  =   ""
      DateDefault     =   "20010101"
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
      ButtonColor     =   13684944
      AutoMenu        =   0   'False
      StartMonth      =   4
      ButtonAlign     =   0
      BoundDataType   =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Personal Tax Year To Penalize:"
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
      Left            =   5940
      TabIndex        =   19
      Top             =   2268
      Width           =   3060
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Please Note: Only one real and one personal penalty calculation combination can be processed at a time."
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
      Height          =   276
      Left            =   1656
      TabIndex        =   18
      Top             =   8520
      Width           =   8556
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   1917.529
      X2              =   9710.486
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   3252
      Left            =   1920
      Top             =   2760
      Width           =   7812
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Backup File Name"
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
      Left            =   4320
      TabIndex        =   16
      Top             =   3480
      Width           =   1860
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Post Date"
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
      Left            =   2280
      TabIndex        =   15
      Top             =   3480
      Width           =   1620
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Select A Billing Type"
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
      Left            =   4560
      TabIndex        =   11
      Top             =   1200
      Width           =   1980
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Real Tax Year To Penalize:"
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
      Left            =   1380
      TabIndex        =   8
      Top             =   2268
      Width           =   2580
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Report Type:"
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
      Height          =   360
      Left            =   3120
      TabIndex        =   7
      Top             =   6360
      Width           =   1812
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
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
      Height          =   360
      Left            =   3000
      TabIndex        =   6
      Top             =   6960
      Width           =   1812
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   612
      Left            =   1200
      Top             =   2040
      Width           =   9252
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   1452
      Left            =   2880
      Top             =   6000
      Width           =   5652
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   660
      Index           =   1
      Left            =   1500
      Top             =   468
      Width           =   8652
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Penalty Calculations"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Left            =   3120
      TabIndex        =   1
      Top             =   636
      Width           =   5292
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   780
      Left            =   1500
      Top             =   360
      Width           =   8652
   End
End
Attribute VB_Name = "frmVATaxCalcPenalty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim CurrRTaxYear As Integer
  Dim CurrPTaxYear As Integer
  Dim Over As clsTextBoxOverRider
  Dim PrincInt As Boolean
  Dim IntInt As Boolean
  Dim AdvColInt As Boolean
  Dim LateListInt As Boolean
  Dim Opt1Int As Boolean
  Dim Opt2Int As Boolean
  Dim Opt3Int As Boolean
  Dim Years() As Integer
  Dim YrCnt As Integer
  Public BillType As String
  Dim ThisOpt$
  Private Temp_Class As Resize_Class
  Dim GCustList() As Long
  Dim GCustCnt As Long
  Dim GFirstTrans As Long
  Dim GLastTrans As Long
  Dim PPTRAYN As String * 1

Private Sub ChkAll_Click()
  If ChkAll.Value = 1 Then
    fpList.Enabled = False
    fpcmbPrintOrder.Enabled = True
  Else
    fpList.Enabled = True
    fpcmbPrintOrder.Enabled = False
  End If
End Sub

Private Sub cmdExit_Click()
  KillFile "C:\CPWork\pencalc.dat"

  frmVATaxPenaltyMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdPrintList_Click()
  Dim PRRec As VARETaxBillType
  Dim PPRec As VAPPTaxBillType
  Dim PRHandle As Integer
  Dim PPHandle As Integer
  Dim NumOfPRRecs As Long
  Dim NumOfPPRecs As Long
  Dim MyPath$, ThisFile$
  Dim x As Integer
  Dim RptFile$
  Dim RptHandle As Integer
  Dim dlm$
  Dim TotAmt#
  Dim PCnt As Long
  Dim TransDate$
  
  dlm$ = "~"
  
  If fpList.SelCount = 0 Then
    Call TaxMsg(900, "Please make a selection from the list.")
    Exit Sub
  End If
  MyPath = StartPath + "\TAXBILLBU\"
  fpList.Col = 1
  fpList.Selected(fpList.ListIndex) = True
  fpList.Row = fpList.ListIndex
  ThisFile = QPTrim$(fpList.ColText)
  fpList.Col = 2
  GFirstTrans = CLng(fpList.ColText)
  fpList.Col = 3
  GLastTrans = CLng(fpList.ColText)
  
  RptFile = "TAXRPTS\OLDBILLS.RPT"
  RptHandle = FreeFile
  Open RptFile$ For Output As #RptHandle
  PCnt = 0
  TotAmt = 0
  TransDate = ""
  If fpcmbType.Text = "REAL" Then
    OpenRealPostedReprintFile PRHandle, NumOfPRRecs, ThisFile
    For x = 1 To NumOfPRRecs
      Get PRHandle, x, PRRec
      If CInt(PRRec.BillNumber) < 0 Then GoTo SkipR
      If TransDate = "" Then
        TransDate = MakeRegDate(PRRec.PostDate)
      End If
      PCnt = PCnt + 1
      
      TotAmt = OldRound(TotAmt + PRRec.TotalBillDue)
      '                       0                      1                   2
      Print #RptHandle, PRRec.BillNumber; dlm; PRRec.CustRec; dlm; QPTrim$(PRRec.CustName); dlm;
      '                       3                     4                  5            6           7
      Print #RptHandle, PRRec.TaxYear; dlm; PRRec.TotalBillDue; dlm; "REAL"; dlm; TotAmt; dlm; PCnt; dlm;
      '                     8
      Print #RptHandle, TransDate
SkipR:
    Next x
    Close PRHandle
  ElseIf fpcmbType.Text = "PERSONAL" Then
    OpenPersPostedReprintFile PPHandle, NumOfPPRecs, ThisFile
    For x = 1 To NumOfPPRecs
      Get PPHandle, x, PPRec
      If CInt(PPRec.BillNumber) < 0 Then GoTo SkipP
      If TransDate = "" Then
        TransDate = MakeRegDate(PPRec.PostDate)
      End If
      PCnt = PCnt + 1
      TotAmt = OldRound(TotAmt + PPRec.TotalBillDue)
      '                       0                      1                   2
      Print #RptHandle, PPRec.BillNumber; dlm; PPRec.CustRec; dlm; QPTrim$(PPRec.CustName); dlm;
      '                       3                     4                    5              6           7
      Print #RptHandle, PPRec.TaxYear; dlm; PPRec.TotalBillDue; dlm; "PERSONAL"; dlm; TotAmt; dlm; PCnt; dlm;
      '                     8
      Print #RptHandle, TransDate
SkipP:
    Next x
    Close PPHandle
  End If
  
  Close
  arVATaxOldBillTransList.Show
  
End Sub

Private Sub cmdProcess_Click()
  If ChkAll.Value = 0 And fpList.SelCount = 0 Then
    Call TaxMsg(800, "Please make a selection from the list or check 'Calculate For Entire Year'.")
    Exit Sub
  End If
  If fpcmbType.Text = "REAL" And (BillType = "R" Or BillType = "B") Then
    If fpcmbPrintOpt.Text = "Graphical" Then
      Call ProcessRealGraphics
    ElseIf fpcmbPrintOpt.Text = "Text" Then
      Call ProcessRealText
    End If
  ElseIf fpcmbType.Text = "PERSONAL" And (BillType = "P" Or BillType = "B") Then
    If fpcmbPrintOpt.Text = "Graphical" Then
      Call ProcessPersGraphics
    ElseIf fpcmbPrintOpt.Text = "Text" Then
      Call ProcessPersText
    End If
  End If
  
End Sub

Private Sub cmdTable_Click()
  Dim One As Integer
  Dim AHandle As Integer
  Dim ThisAns$
  Dim Message$
  
  If fpcmbType.Text = "PERSONAL" Then
    One = 1
    AHandle = FreeFile
    Open "C:\CPWork\pencalc.dat" For Output As AHandle
    Print #AHandle, One
    Close AHandle
    frmVATaxPPenRateSetUpTbl.Show
  ElseIf fpcmbType.Text = "REAL" Then
    One = 1
    AHandle = FreeFile
    Open "C:\CPWork\pencalc.dat" For Output As AHandle
    Print #AHandle, One
    Close AHandle
    frmVATaxPenRateSetUpTbl.Show
  End If
  
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%E"
      Call cmdExit_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%P"
      Call cmdProcess_Click
      KeyCode = 0
    Case vbKeyF9:
      SendKeys "%S"
      Call cmdPrintList_Click
      KeyCode = 0
    Case vbKeyF5:
      SendKeys "%V"
      Call cmdTable_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Me.HelpContextID = hlpTaxPenalty
  Call LoadMe
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("CitiTaxes.exe terminated via menu bar on frmVATaxCalcPenalty.")
      Call Terminate
      End
    End If
  End If

End Sub
Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    'Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
    DoEvents
  End If
End Sub

Private Sub LoadMe()
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim TblRec As PenaltyRateTablesType
  Dim PRHandle As Integer
  Dim NumOfPRRecs As Integer
  Dim x As Integer
  
  PPTRAYN = "N"
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  fptxtCurrRYear.Text = CStr(TaxMasterRec.RTaxYear)
  CurrRTaxYear = TaxMasterRec.RTaxYear
  
  fptxtCurrPYear.Text = CStr(TaxMasterRec.PTaxYear)
  CurrPTaxYear = TaxMasterRec.PTaxYear
  
  fptxtCurrPYear.Enabled = False
  fptxtCurrRYear.Enabled = True
  
  fpcmbPrintOrder.Enabled = False
  
  OpenTaxPenRateTbls PRHandle, NumOfPRRecs
  For x = 1 To NumOfPRRecs
    Get PRHandle, x, TblRec
    If x = 1 Then
      If TblRec.BillType = "R" Then
        fpcmbType.Text = "REAL"
        If NumOfPRRecs = 1 Then
          BillType = "R"
        End If
      ElseIf TblRec.BillType = "P" Then
        fpcmbType.Text = "PERSONAL"
        If NumOfPRRecs = 1 Then
          BillType = "P"
        End If
      End If
    ElseIf x = 2 Then
      If fpcmbType.Text = "REAL" Then
        If TblRec.BillType = "P" Then
          BillType = "B"
        End If
      ElseIf fpcmbType.Text = "PERSONAL" Then
        If TblRec.BillType = "R" Then
          BillType = "B"
        End If
      End If
    End If
  Next x
  
  fpcmbType.AddItem "REAL"
  fpcmbType.AddItem "PERSONAL"
  
  fpcmbPrintOrder.Text = "1) Account Number Order"
  fpcmbPrintOrder.AddItem "1) Account Number Order"
  fpcmbPrintOrder.AddItem "2) Customer Name Order"
  fpcmbPrintOrder.AddItem "3) Search Name Order"
  ThisOpt = QPTrim$(TaxMasterRec.OptSrchCust)
  If ThisOpt <> "" Then
    fpcmbPrintOrder.AddItem "4) " + ThisOpt + " Order"
  End If
  
  fpcmbPrintOpt.Text = "Graphical"
  fpcmbPrintOpt.AddItem "Graphical"
  fpcmbPrintOpt.AddItem "Text"
  Call LoadList
  ChkAll.Value = 0
End Sub

Private Sub fpcmbPrintOpt_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbPrintOpt.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbPrintOpt.ListIndex = -1
  End If
  If fpcmbPrintOpt.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbPrintOrder.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbPrintOrder_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbPrintOrder.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbPrintOrder.ListIndex = -1
  End If
  If fpcmbPrintOrder.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbType.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbType_Change()
  If fpcmbType.Text = "REAL" Then
    fptxtCurrPYear.Enabled = False
    fptxtCurrRYear.Enabled = True
    chkPPTRAYN.Visible = False
    PPTRAYN = "N"
  ElseIf fpcmbType.Text = "PERSONAL" Then
    fptxtCurrPYear.Enabled = True
    fptxtCurrRYear.Enabled = False
    If Check4PPTRADisc(CInt(fptxtCurrPYear.Text)) = True Then
      chkPPTRAYN.Visible = True
'      PPTRAYN = "Y"
    Else
      chkPPTRAYN.Visible = False
    End If
  End If
End Sub

Private Sub fpcmbType_Click()
  Dim One As Integer
  Dim AHandle As Integer
  Dim ThisAns$
  Dim Message$
  
  If BillType <> "B" Then
    If BillType = "R" And fpcmbType.Text = "PERSONAL" Then
      Message = "No personal penalty rate tables have been set up. Would you like to jump to the personal penalty set up screen now?"
      If TaxMsgWOpts(800, Message, "F10 Jump", "ESC Exit") = "abort" Then
        Unload frmVATaxMsgW3Opts
        Exit Sub
      Else
        Unload frmVATaxMsgW3Opts
        One = 1
        AHandle = FreeFile
        Open "C:\CPWork\pencalc.dat" For Output As AHandle
        Print #AHandle, One
        Close AHandle
        frmVATaxPPenRateSetUpTbl.Show
      End If
    ElseIf BillType = "P" And fpcmbType.Text = "REAL" Then
      Message = "No real penalty rate tables have been set up. Would you like to jump to the real penalty set up screen now?"
      If TaxMsgWOpts(800, Message, "F10 Jump", "ESC Exit") = "abort" Then
        Unload frmVATaxMsgW3Opts
        Exit Sub
      Else
        Unload frmVATaxMsgW3Opts
        One = 1
        AHandle = FreeFile
        Open "C:\CPWork\pencalc.dat" For Output As AHandle
        Print #AHandle, One
        Close AHandle
        frmVATaxPenRateSetUpTbl.Show
      End If
    End If
  End If
  Call LoadList
  
End Sub

Private Sub ProcessRealGraphics()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long
  Dim TblRec As PenaltyRateTablesType
  Dim PRHandle As Integer
  Dim NumOfPRRecs As Integer
  Dim Balance As Double
  Dim ChargeThis As Double
  Dim PenRec As PenaltyRecType
  Dim INHandle As Integer
  Dim NumOfINRecs As Long
  Dim CurTaxYr As Integer
  Dim GCurTaxYr As Integer
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim NextRec As Long, PenRecord&
  Dim y As Long, sb As Integer
  Dim BillNumber$, CustAcct&
  Dim ThisRec As Integer
  Dim RptHandle As Integer
  Dim RptFile$, TownName$
  Dim dlm$
  Dim TotCnt As Long
  Dim TotPen As Double
  Dim TotBal As Double
  Dim IdxRec As CustNameIdxType
  Dim IdxHandle As Integer
  Dim NumOfIdx As Long
  Dim IdxFlag As Boolean
  Dim SearchRec As SrchNameIdxType
  Dim SHandle As Integer
  Dim OptFlag As Boolean
  Dim OptRec As OptCustIdxType
  Dim OHandle As Integer
  Dim NumOfORecs As Long
  Dim PrntCnt As Long
  
  dlm$ = "~"
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close
  GCurTaxYr = TaxMasterRec.RTaxYear
  TownName = QPTrim$(TaxMasterRec.Name)
  CurTaxYr = CInt(fptxtCurrRYear.Text)
  
  If GCurTaxYr <> CurTaxYr Then
    If TaxMsgWOpts(700, "The system real tax year (" + CStr(GCurTaxYr) + ") is not the same as the tax year entered on this form (" + CStr(CurTaxYr) + "). If you wish to continue anyway then press F10. Otherwise, press ESC to edit.", "F10 Continue", "ESC Exit") = "abort" Then
      Unload frmVATaxMsgWOpts
      fptxtCurrRYear.SetFocus
      Exit Sub
    Else
      Unload frmVATaxMsgWOpts
      MainLog ("WARNING: Issued because when calculating real penalty amounts the user used the year " + CStr(CurTaxYr) + " instead of the system current real tax year of " + CStr(GCurTaxYr) + ".")
    End If
  End If
  
  OpenTaxPenRateTbls PRHandle, NumOfPRRecs
  For x = 1 To NumOfPRRecs
    Get PRHandle, x, TblRec
    If TblRec.BillType = "R" Then
      ThisRec = x
    End If
  Next x
  If ThisRec = 0 Then
    Call TaxMsg(900, "ERROR: There was a problem determining which rate table to use. Please save the real rate tables again.")
    Close
    Exit Sub
  End If
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  If NumOfTCRecs = 0 Then
    Call TaxMsg(900, "No customers have been saved")
    Close
    Exit Sub
  End If
  
  If Exist(TaxRPenFile) Then
    KillFile TaxRPenFile
  End If
  
  OpenRPenRecFile INHandle, NumOfINRecs
  OpenTaxTransFile TTHandle, NumOfTTRecs
  RptFile$ = "TAXRPTS\TAXRPEN.RPT"     'Report File Name
  RptHandle = FreeFile
  Open RptFile$ For Output As #RptHandle
  
  If ChkAll.Value = 0 Then GoTo PrintSingleBill
  
  IdxFlag = False
  OptFlag = False
  If QPTrim$(fpcmbPrintOrder.Text) = "2) Customer Name Order" Then
    OpenNameIdxFile IdxHandle, NumOfIdx
    ReDim IdxArray(1 To NumOfIdx) As Long

    For x = 1 To NumOfIdx
      Get IdxHandle, x, IdxRec
      IdxArray(x) = IdxRec.CustRec
    Next x
    Close IdxHandle
    IdxFlag = True
    NumOfTCRecs = NumOfIdx
  ElseIf QPTrim$(fpcmbPrintOrder.Text) = "3) Search Name Order" Then
    OpenSrchNameIdxFile SHandle, NumOfIdx
    If NumOfIdx = 0 Then
      frmVATaxMsg.Label1.Caption = "There are no search names indexed."
      frmVATaxMsg.Label1.Top = 900
      frmVATaxMsg.Show
      fpcmbPrintOrder.SetFocus
      Close SHandle
      Exit Sub
    End If
    ReDim IdxArray(1 To NumOfIdx) As Long

    For x = 1 To NumOfIdx
      Get SHandle, x, SearchRec
      IdxArray(x) = SearchRec.CustRec
    Next x
    Close SHandle
    IdxFlag = True
    NumOfTCRecs = NumOfIdx
  ElseIf QPTrim$(fpcmbPrintOrder.Text) = "4) " + ThisOpt + " Order" Then
    OpenCustOptSearchFile OHandle, NumOfIdx
    If NumOfIdx = 0 Then
      Call TaxMsg(900, "There are no " + ThisOpt + "descriptions indexed.")
      Close OHandle
      Exit Sub
    End If
    ReDim IdxArray(1 To NumOfIdx) As Long
    
    For x = 1 To NumOfIdx
      Get OHandle, x, OptRec
      IdxArray(x) = OptRec.CustRec
    Next x
    Close OHandle
    IdxFlag = True
    OptFlag = True
    NumOfTCRecs = NumOfIdx
  End If
  
'  If Exist(TaxRPenFile) Then
'    KillFile TaxRPenFile
'  End If
  
  frmVATaxShowPctComp.Label1 = "Calculating Real Penalty"
  frmVATaxShowPctComp.Show , Me
  EnableCloseButton Me.hwnd, False
  cmdTable.Enabled = False
  cmdExit.Enabled = False
  cmdProcess.Enabled = False
  PrntCnt = 0
  For x = 1 To NumOfTCRecs
    If IdxFlag = True Then
      Get TCHandle, IdxArray(x), TaxCust
    Else
      Get TCHandle, x, TaxCust
    End If
    If TaxCust.Deleted <> 0 Then GoTo SkipIt2
    If TaxCust.Penalty = "N" Then GoTo SkipIt2
    CustAcct = TaxCust.Acct
    If TaxCust.LastTrans > 0 Then
      NextRec = TaxCust.LastTrans
      
      Do While NextRec > 0
        Get TTHandle, NextRec, TaxTrans
        If TaxTrans.TaxYear = CurTaxYr Then
          If TaxTrans.TranType = 1 And TaxTrans.BillType = "R" Then
            Balance = TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Collection
            Balance = OldRound(Balance + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2)
            Balance = OldRound(Balance + TaxTrans.Revenue.RevOpt3)
            Balance = OldRound(Balance - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.CollectionPd))
            Balance = OldRound(Balance - (TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd))
            Balance = OldRound(Balance - TaxTrans.Revenue.RevOpt3Pd)
            If Balance > 0 Then
              ChargeThis = FigurePenalty(Balance, PRHandle, ThisRec)
              ChargeThis = OldRound(ChargeThis)
              BillNumber$ = TaxTrans.Description
              sb = InStr(BillNumber$, "Bill #")
              If sb > 0 Then
                BillNumber$ = Mid$(TaxTrans.Description, sb + 6, 10)
              End If
              If ChargeThis > 0 Then
                PenRec.CustRec = CustAcct&
                PenRec.CustName = TaxCust.CustName
                PenRec.TaxYear = TaxTrans.TaxYear
                PenRec.Amount = ChargeThis
                PenRec.BillNumber = QPTrim$(BillNumber$)
                PenRec.BillRec = NextRec&
                PenRec.CurYear = CurTaxYr
                PenRec.Balance = Balance
                PenRec.BillType = "R"
                PenRec.CustPin = TaxCust.PIN
                PenRec.RealPin = TaxTrans.RealPin
                PenRecord& = PenRecord& + 1
                Put INHandle, PenRecord, PenRec
                TotPen = OldRound(TotPen + ChargeThis)
                TotBal = OldRound(TotBal + Balance#)
                TotCnt = TotCnt + 1
                GoSub PrintIt
                PrntCnt = PrntCnt + 1
              End If
            End If
          End If
        End If
        NextRec = TaxTrans.LastTrans
      Loop
    End If
SkipIt2:
    frmVATaxShowPctComp.ShowPctComp x, NumOfTCRecs
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      cmdTable.Enabled = True
      cmdExit.Enabled = True
      cmdProcess.Enabled = True
      EnableCloseButton Me.hwnd, True
      Exit Sub
    End If
  Next x
  
  Unload frmVATaxShowPctComp
  cmdTable.Enabled = True
  cmdExit.Enabled = True
  cmdProcess.Enabled = True
  EnableCloseButton Me.hwnd, True
  Close
  If PrntCnt = 0 Then
    Call TaxMsg(800, "No customers qualified for a penalty calculation with the parameters entered.")
    Exit Sub
  End If
  arVATaxRPenRpt.Show

  Exit Sub
  
  
PrintIt:
  '                     0             1              2                       3
  Print #RptHandle, TownName; dlm; Balance#; dlm; CustAcct&; dlm; QPTrim$(TaxCust.CustName); dlm;
  '                        4                   5                6              7            8       9
  Print #RptHandle, TaxTrans.TaxYear; dlm; ChargeThis; dlm; BillNumber; dlm; TotPen; dlm; TotBal; dlm; TotCnt
  
  Return
  

PrintSingleBill:
'  If Exist(TaxRPenFile) Then
'    KillFile TaxRPenFile
'  End If
  Call GetGCustList
  
  frmVATaxShowPctComp.Label1 = "Calculating Real Penalty"
  frmVATaxShowPctComp.Show , Me
  EnableCloseButton Me.hwnd, False
  cmdTable.Enabled = False
  cmdExit.Enabled = False
  cmdProcess.Enabled = False
  PrntCnt = 0
  For x = 1 To GCustCnt
    Get TCHandle, GCustList(x), TaxCust
    If TaxCust.Deleted <> 0 Then GoTo SkipIt
    If TaxCust.Penalty = "N" Then GoTo SkipIt
    CustAcct = TaxCust.Acct
    If TaxCust.LastTrans > 0 Then
      NextRec = TaxCust.LastTrans
      Do While NextRec > 0
        Get TTHandle, NextRec, TaxTrans
        If NextRec >= GFirstTrans And NextRec <= GLastTrans Then
          If TaxTrans.TranType = 1 And TaxTrans.BillType = "R" Then
            Balance = TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Collection
            Balance = OldRound(Balance + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2)
            Balance = OldRound(Balance + TaxTrans.Revenue.RevOpt3)
            Balance = OldRound(Balance - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.CollectionPd))
            Balance = OldRound(Balance - (TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd))
            Balance = OldRound(Balance - TaxTrans.Revenue.RevOpt3Pd)
            If Balance > 0 Then
              ChargeThis = FigurePenalty(Balance, PRHandle, ThisRec)
              ChargeThis = OldRound(ChargeThis)
              BillNumber$ = TaxTrans.Description
              sb = InStr(BillNumber$, "Bill #")
              If sb > 0 Then
                BillNumber$ = Mid$(TaxTrans.Description, sb + 6, 10)
              End If
              If ChargeThis > 0 Then
                PenRec.CustRec = CustAcct&
                PenRec.CustName = TaxCust.CustName
                PenRec.TaxYear = TaxTrans.TaxYear
                PenRec.Amount = ChargeThis
                PenRec.BillNumber = QPTrim$(BillNumber$)
                PenRec.BillRec = NextRec&
                PenRec.CurYear = CurTaxYr
                PenRec.Balance = Balance
                PenRec.CustPin = TaxCust.PIN
                PenRec.RealPin = TaxTrans.RealPin
                PenRec.BillType = "R"
                PenRecord& = PenRecord& + 1
                Put INHandle, PenRecord, PenRec
                TotPen = OldRound(TotPen + ChargeThis)
                TotBal = OldRound(TotBal + Balance#)
                TotCnt = TotCnt + 1
                GoSub PrintIt
                PrntCnt = PrntCnt + 1
              End If
            End If
          End If
        End If
        NextRec = TaxTrans.LastTrans
      Loop
    End If
SkipIt:
    frmVATaxShowPctComp.ShowPctComp x, GCustCnt
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      cmdTable.Enabled = True
      cmdExit.Enabled = True
      cmdProcess.Enabled = True
      EnableCloseButton Me.hwnd, True
      Exit Sub
    End If
  Next x
  
  Unload frmVATaxShowPctComp
  cmdTable.Enabled = True
  cmdExit.Enabled = True
  cmdProcess.Enabled = True
  EnableCloseButton Me.hwnd, True
  Close
  If PrntCnt = 0 Then
    Call TaxMsg(800, "No customers qualified for a penalty calculation with the parameters entered.")
    Exit Sub
  End If
  arVATaxRPenRpt.Show
  Exit Sub

End Sub

Private Sub ProcessPersGraphics()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long
  Dim TblRec As PenaltyRateTablesType
  Dim PRHandle As Integer
  Dim NumOfPRRecs As Integer
  Dim Balance As Double
  Dim ChargeThis As Double
  Dim PenRec As PenaltyRecType
  Dim INHandle As Integer
  Dim NumOfINRecs As Long
  Dim CurTaxYr As Integer
  Dim GCurTaxYr As Integer
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim NextRec As Long, PenRecord&
  Dim y As Long, sb As Integer
  Dim BillNumber$, CustAcct&
  Dim ThisRec As Integer
  Dim dlm$, TownName$
  Dim RptHandle As Integer
  Dim RptFile$
  Dim TotPen As Double
  Dim TotBal As Double
  Dim TotCnt As Long
  Dim IdxRec As CustNameIdxType
  Dim IdxHandle As Integer
  Dim NumOfIdx As Long
  Dim IdxFlag As Boolean
  Dim SearchRec As SrchNameIdxType
  Dim SHandle As Integer
  Dim OptFlag As Boolean
  Dim OptRec As OptCustIdxType
  Dim OHandle As Integer
  Dim NumOfORecs As Long
  Dim PrntCnt As Long
  Dim ThisBal As Double
  
  dlm = "~"
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close
  GCurTaxYr = TaxMasterRec.RTaxYear
  TownName = QPTrim$(TaxMasterRec.Name)
  
  CurTaxYr = CInt(fptxtCurrPYear.Text)
  If GCurTaxYr <> CurTaxYr Then
    If TaxMsgWOpts(700, "The system personal tax year (" + CStr(GCurTaxYr) + ") is not the same as the tax year entered on this form (" + CStr(CurTaxYr) + "). If you wish to continue anyway then press F10. Otherwise, press ESC to edit.", "F10 Continue", "ESC Exit") = "abort" Then
      Unload frmVATaxMsgWOpts
      fptxtCurrPYear.SetFocus
      Exit Sub
    Else
      Unload frmVATaxMsgWOpts
      MainLog ("WARNING: Issued because when calculating personal penalty amounts the user used the year " + CStr(CurTaxYr) + " instead of the system current personal tax year of " + CStr(GCurTaxYr) + ".")
    End If
  End If
  
  OpenTaxPenRateTbls PRHandle, NumOfPRRecs
  For x = 1 To NumOfPRRecs
    Get PRHandle, x, TblRec
    If TblRec.BillType = "P" Then
      ThisRec = x
    End If
  Next x
  If ThisRec = 0 Then
    Call TaxMsg(900, "ERROR: There was a problem determining which rate table to use. Please save the real rate tables again.")
    Close
    Exit Sub
  End If
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  If NumOfTCRecs = 0 Then
    Call TaxMsg(900, "No customers have been saved")
    Close
    Exit Sub
  End If
  
  If Exist(TaxPPenFile) Then
    KillFile TaxPPenFile
  End If
  
  OpenPPenRecFile INHandle, NumOfINRecs
  OpenTaxTransFile TTHandle, NumOfTTRecs
  
  RptFile$ = "TAXRPTS\TAXPPEN.RPT"     'Report File Name
  RptHandle = FreeFile
  Open RptFile$ For Output As #RptHandle
  
  If ChkAll.Value = 0 Then GoTo PrintSingleBill
  
  IdxFlag = False
  OptFlag = False
  If QPTrim$(fpcmbPrintOrder.Text) = "2) Customer Name Order" Then
    OpenNameIdxFile IdxHandle, NumOfIdx
    ReDim IdxArray(1 To NumOfIdx) As Long

    For x = 1 To NumOfIdx
      Get IdxHandle, x, IdxRec
      IdxArray(x) = IdxRec.CustRec
    Next x
    Close IdxHandle
    IdxFlag = True
    NumOfTCRecs = NumOfIdx
  ElseIf QPTrim$(fpcmbPrintOrder.Text) = "3) Search Name Order" Then
    OpenSrchNameIdxFile SHandle, NumOfIdx
    If NumOfIdx = 0 Then
      frmVATaxMsg.Label1.Caption = "There are no search names indexed."
      frmVATaxMsg.Label1.Top = 900
      frmVATaxMsg.Show
      fpcmbPrintOrder.SetFocus
      Close SHandle
      Exit Sub
    End If
    ReDim IdxArray(1 To NumOfIdx) As Long

    For x = 1 To NumOfIdx
      Get SHandle, x, SearchRec
      IdxArray(x) = SearchRec.CustRec
    Next x
    Close SHandle
    IdxFlag = True
    NumOfTCRecs = NumOfIdx
  ElseIf QPTrim$(fpcmbPrintOrder.Text) = "4) " + ThisOpt + " Order" Then
    OpenCustOptSearchFile OHandle, NumOfIdx
    If NumOfIdx = 0 Then
      Call TaxMsg(900, "There are no " + ThisOpt + "descriptions indexed.")
      Close OHandle
      Exit Sub
    End If
    ReDim IdxArray(1 To NumOfIdx) As Long
    
    For x = 1 To NumOfIdx
      Get OHandle, x, OptRec
      IdxArray(x) = OptRec.CustRec
    Next x
    Close OHandle
    IdxFlag = True
    OptFlag = True
    NumOfTCRecs = NumOfIdx
  End If
  
  frmVATaxShowPctComp.Label1 = "Calculating Personal Penalty"
  frmVATaxShowPctComp.Show , Me
  EnableCloseButton Me.hwnd, False
  cmdTable.Enabled = False
  cmdExit.Enabled = False
  cmdProcess.Enabled = False
  
  PrntCnt = 0
  For x = 1 To NumOfTCRecs
    If IdxFlag = True Then
      Get TCHandle, IdxArray(x), TaxCust
    Else
      Get TCHandle, x, TaxCust
    End If
    If TaxCust.Deleted <> 0 Then GoTo SkipIt
'    If TaxCust.FirstPersRec = 0 Then GoTo SkipIt
    If TaxCust.Penalty = "N" Then GoTo SkipIt
    ThisBal = GetCustPersBalance(TaxCust.Acct, -1) 'added 1/24/07 because if the user
    'chooses to not include PPTRA discounts then if the customer has paid in full
    'including the PPTRA discount it appears as if they still have a balance
    If ThisBal = 0 Then GoTo SkipIt
    CustAcct = TaxCust.Acct
    If TaxCust.LastTrans > 0 Then
      NextRec = TaxCust.LastTrans
      Do While NextRec > 0
        Get TTHandle, NextRec, TaxTrans
        If TaxTrans.TaxYear = CurTaxYr Then
          If TaxTrans.TranType = 1 And TaxTrans.BillType = "P" Then
            If PPTRAYN = "Y" Then
              If TaxTrans.PPTRADisc > 0 Then
                TaxTrans.PPTRADisc = 0
              ElseIf TaxTrans.Revenue.Future1 > 0 Then
                TaxTrans.Revenue.Principle1 = OldRound(TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Future1)
              End If
            End If
            Balance = TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3
            Balance = OldRound(Balance + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
            Balance = OldRound(Balance + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2)
            Balance = OldRound(Balance + TaxTrans.Revenue.RevOpt3)
            Balance = OldRound(Balance - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd))
            Balance = OldRound(Balance - (TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd))
            Balance = OldRound(Balance - (TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd))
            Balance = OldRound(Balance - TaxTrans.Revenue.RevOpt3Pd)
            Balance = OldRound(Balance - OldRound(TaxTrans.PPTRADisc - TaxTrans.PPTRARmvl))
            If Balance > 0 Then
              ChargeThis = FigurePenalty(Balance, PRHandle, ThisRec)
              ChargeThis = OldRound(ChargeThis)
              BillNumber$ = TaxTrans.Description
              sb = InStr(BillNumber$, "Bill #")
              If sb > 0 Then
                BillNumber$ = Mid$(TaxTrans.Description, sb + 6, 10)
              End If
              If ChargeThis > 0 Then
                PenRec.CustRec = CustAcct&
                PenRec.CustName = TaxCust.CustName
                PenRec.TaxYear = TaxTrans.TaxYear
                PenRec.Amount = ChargeThis
                PenRec.BillNumber = QPTrim$(BillNumber$)
                PenRec.BillRec = NextRec&
                PenRec.CurYear = CurTaxYr
                PenRec.BillType = "P"
                PenRec.Balance = Balance
                PenRec.CustPin = TaxCust.PIN
                PenRec.PersPin = TaxTrans.PersPin
                PenRecord& = PenRecord& + 1
                Put INHandle, PenRecord, PenRec
                TotCnt = TotCnt + 1
                TotPen = OldRound(TotPen + ChargeThis)
                TotBal = OldRound(TotBal + Balance#)
                GoSub PrintIt
                PrntCnt = PrntCnt + 1
              End If
            End If
          End If
        End If
        NextRec = TaxTrans.LastTrans
      Loop
    End If
SkipIt:
    frmVATaxShowPctComp.ShowPctComp x, NumOfTCRecs
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      cmdTable.Enabled = True
      cmdExit.Enabled = True
      cmdProcess.Enabled = True
      EnableCloseButton Me.hwnd, True
      Exit Sub
    End If
  Next x
  Unload frmVATaxShowPctComp
  cmdTable.Enabled = True
  cmdExit.Enabled = True
  cmdProcess.Enabled = True
  EnableCloseButton Me.hwnd, True
  
  Close
  If PrntCnt = 0 Then
    Call TaxMsg(800, "No customers qualified for a penalty calculation with the parameters entered.")
    Exit Sub
  End If
  
  arVATaxPPenRpt.Show
  
  Exit Sub
  
  
PrintIt:
  '                     0             1              2                       3
  Print #RptHandle, TownName; dlm; Balance#; dlm; CustAcct&; dlm; QPTrim$(TaxCust.CustName); dlm;
  '                        4                   5                6               7            8           9
  Print #RptHandle, TaxTrans.TaxYear; dlm; ChargeThis; dlm; BillNumber; dlm; TotPen; dlm; TotBal; dlm; TotCnt
  
  Return


PrintSingleBill:
  Call GetGCustList
  
  frmVATaxShowPctComp.Label1 = "Calculating Personal Penalty"
  frmVATaxShowPctComp.Show , Me
  EnableCloseButton Me.hwnd, False
  cmdTable.Enabled = False
  cmdExit.Enabled = False
  cmdProcess.Enabled = False
  PrntCnt = 0
  For x = 1 To GCustCnt
    Get TCHandle, GCustList(x), TaxCust
    If TaxCust.Deleted <> 0 Then GoTo SkipIt2
    If TaxCust.Penalty = "N" Then GoTo SkipIt2
    CustAcct = TaxCust.Acct
    If TaxCust.LastTrans > 0 Then
      NextRec = TaxCust.LastTrans
      Do While NextRec > 0
        Get TTHandle, NextRec, TaxTrans
        If NextRec >= GFirstTrans And NextRec <= GLastTrans Then
          If TaxTrans.TranType = 1 And TaxTrans.BillType = "P" Then
            Balance = TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3
            Balance = OldRound(Balance + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
            Balance = OldRound(Balance + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2)
            Balance = OldRound(Balance + TaxTrans.Revenue.RevOpt3)
            Balance = OldRound(Balance - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd))
            Balance = OldRound(Balance - (TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd))
            Balance = OldRound(Balance - (TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd))
            Balance = OldRound(Balance - TaxTrans.Revenue.RevOpt3Pd)
            Balance = OldRound(Balance - OldRound(TaxTrans.PPTRADisc - TaxTrans.PPTRARmvl))
            If Balance > 0 Then
              ChargeThis = FigurePenalty(Balance, PRHandle, ThisRec)
              ChargeThis = OldRound(ChargeThis)
              BillNumber$ = TaxTrans.Description
              sb = InStr(BillNumber$, "Bill #")
              If sb > 0 Then
                BillNumber$ = Mid$(TaxTrans.Description, sb + 6, 10)
              End If
              If ChargeThis > 0 Then
                PenRec.CustRec = CustAcct&
                PenRec.CustName = TaxCust.CustName
                PenRec.TaxYear = TaxTrans.TaxYear
                PenRec.Amount = ChargeThis
                PenRec.BillNumber = QPTrim$(BillNumber$)
                PenRec.BillRec = NextRec&
                PenRec.CurYear = CurTaxYr
                PenRec.BillType = "P"
                PenRec.Balance = Balance
                PenRec.CustPin = TaxCust.PIN
                PenRec.PersPin = TaxTrans.PersPin
                PenRecord& = PenRecord& + 1
                Put INHandle, PenRecord, PenRec
                TotCnt = TotCnt + 1
                TotPen = OldRound(TotPen + ChargeThis)
                TotBal = OldRound(TotBal + Balance#)
                GoSub PrintIt
                PrntCnt = PrntCnt + 1
              End If
            End If
          End If
        End If
        NextRec = TaxTrans.LastTrans
      Loop
    End If
SkipIt2:
    frmVATaxShowPctComp.ShowPctComp x, GCustCnt
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      cmdTable.Enabled = True
      cmdExit.Enabled = True
      cmdProcess.Enabled = True
      EnableCloseButton Me.hwnd, True
      Exit Sub
    End If
  Next x
  Unload frmVATaxShowPctComp
  cmdTable.Enabled = True
  cmdExit.Enabled = True
  cmdProcess.Enabled = True
  EnableCloseButton Me.hwnd, True
  
  Close
  If PrntCnt = 0 Then
    Call TaxMsg(800, "No customers qualified for a penalty calculation with the parameters entered.")
    Exit Sub
  End If
  
  arVATaxPPenRpt.Show
  
  Exit Sub

End Sub

Private Function FigurePenalty(BaseAmt As Double, ThisHandle As Integer, ThisRec As Integer) As Double
  Dim TblRec As PenaltyRateTablesType
  Dim x As Integer
  Dim ThisPct As Double
  
  FigurePenalty = 0
  Get ThisHandle, ThisRec, TblRec
  For x = 1 To 10
    If TblRec.RateType(x) = "F" Then
      If BaseAmt >= TblRec.FromAmt(x) And BaseAmt < TblRec.ToAmt(x) Then
        FigurePenalty = TblRec.TaxFAmt(x)
        Exit For
      End If
    ElseIf TblRec.RateType(x) = "P" Then
      If BaseAmt >= TblRec.FromAmt(x) And BaseAmt < TblRec.ToAmt(x) Then
        FigurePenalty = OldRound(BaseAmt * (TblRec.TaxPAmt(x) / 100))
        Exit For
      End If
    End If
  Next x
    
'  If TblRec.RateType = "F" Then
'    FigurePenalty = TblRec.FlatAmt
'  ElseIf TblRec.RateType = "S" Then
'    For x = 1 To 10
'  ElseIf TblRec.RateType = "P" Then
'    For x = 1 To 10
'    Next x
'  End If
  
End Function

Private Sub fpcmbType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbType.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbType.ListIndex = -1
  End If
  If fpcmbType.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fptxtCurrRYear.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub
Private Sub ProcessPersText()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long
  Dim TblRec As PenaltyRateTablesType
  Dim PRHandle As Integer
  Dim NumOfPRRecs As Integer
  Dim Balance As Double
  Dim ChargeThis As Double
  Dim PenRec As PenaltyRecType
  Dim INHandle As Integer
  Dim NumOfINRecs As Long
  Dim CurTaxYr As Integer
  Dim GCurTaxYr As Integer
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim NextRec As Long, PenRecord&
  Dim y As Long, sb As Integer
  Dim BillNumber$, CustAcct&
  Dim ThisRec As Integer
  Dim LineCnt As Integer
  Dim FF$, MaxLines As Integer
  Dim RptHandle As Integer
  Dim RptFile$, TownName$, Page As Integer
  Dim TotCnt As Long
  Dim TotPen As Double
  Dim TotBal As Double
  Dim CustName As String * 45
  Dim IdxRec As CustNameIdxType
  Dim IdxHandle As Integer
  Dim NumOfIdx As Long
  Dim IdxFlag As Boolean
  Dim SearchRec As SrchNameIdxType
  Dim SHandle As Integer
  Dim OptFlag As Boolean
  Dim OptRec As OptCustIdxType
  Dim OHandle As Integer
  Dim NumOfORecs As Long
  Dim PrntCnt As Long
  Dim ThisBal As Double
  
  MaxLines = 58
  FF$ = Chr(12)
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close
  GCurTaxYr = TaxMasterRec.RTaxYear
  TownName = QPTrim$(TaxMasterRec.Name)
  
  CurTaxYr = CInt(fptxtCurrPYear.Text)
  If GCurTaxYr <> CurTaxYr Then
    If TaxMsgWOpts(700, "The system personal tax year (" + CStr(GCurTaxYr) + ") is not the same as the tax year entered on this form (" + CStr(CurTaxYr) + "). If you wish to continue anyway then press F10. Otherwise, press ESC to edit.", "F10 Continue", "ESC Exit") = "abort" Then
      Unload frmVATaxMsgWOpts
      fptxtCurrPYear.SetFocus
      Exit Sub
    Else
      Unload frmVATaxMsgWOpts
      MainLog ("WARNING: Issued because when calculating personal penalty amounts the user used the year " + CStr(CurTaxYr) + " instead of the system current tax year of " + CStr(GCurTaxYr) + ".")
    End If
  End If
  
  OpenTaxPenRateTbls PRHandle, NumOfPRRecs
  For x = 1 To NumOfPRRecs
    Get PRHandle, x, TblRec
    If TblRec.BillType = "P" Then
      ThisRec = x
    End If
  Next x
  If ThisRec = 0 Then
    Call TaxMsg(900, "ERROR: There was a problem determining which rate table to use. Please save the real rate tables again.")
    Close
    Exit Sub
  End If
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  If NumOfTCRecs = 0 Then
    Call TaxMsg(900, "No customers have been saved")
    Close
    Exit Sub
  End If
  
  If Exist(TaxPPenFile) Then
    KillFile TaxPPenFile
  End If

  OpenPPenRecFile INHandle, NumOfINRecs
  OpenTaxTransFile TTHandle, NumOfTTRecs
  RptFile$ = "TAXRPTS\TAXPPEN.PRN"     'Report File Name
  RptHandle = FreeFile
  Open RptFile$ For Output As #RptHandle
  
  If ChkAll.Value = 0 Then GoTo PrintSingleBill
  
  IdxFlag = False
  OptFlag = False
  If QPTrim$(fpcmbPrintOrder.Text) = "2) Customer Name Order" Then
    OpenNameIdxFile IdxHandle, NumOfIdx
    ReDim IdxArray(1 To NumOfIdx) As Long
    For x = 1 To NumOfIdx
      Get IdxHandle, x, IdxRec
      IdxArray(x) = IdxRec.CustRec
    Next x
    Close IdxHandle
    IdxFlag = True
    NumOfTCRecs = NumOfIdx
  ElseIf QPTrim$(fpcmbPrintOrder.Text) = "3) Search Name Order" Then
    OpenSrchNameIdxFile SHandle, NumOfIdx
    If NumOfIdx = 0 Then
      frmVATaxMsg.Label1.Caption = "There are no search names indexed."
      frmVATaxMsg.Label1.Top = 900
      frmVATaxMsg.Show
      fpcmbPrintOrder.SetFocus
      Close SHandle
      Exit Sub
    End If
    ReDim IdxArray(1 To NumOfIdx) As Long

    For x = 1 To NumOfIdx
      Get SHandle, x, SearchRec
      IdxArray(x) = SearchRec.CustRec
    Next x
    Close SHandle
    IdxFlag = True
    NumOfTCRecs = NumOfIdx
  ElseIf QPTrim$(fpcmbPrintOrder.Text) = "4) " + ThisOpt + " Order" Then
    OpenCustOptSearchFile OHandle, NumOfIdx
    If NumOfIdx = 0 Then
      Call TaxMsg(900, "There are no " + ThisOpt + "descriptions indexed.")
      Close OHandle
      Exit Sub
    End If
    ReDim IdxArray(1 To NumOfIdx) As Long
    
    For x = 1 To NumOfIdx
      Get OHandle, x, OptRec
      IdxArray(x) = OptRec.CustRec
    Next x
    Close OHandle
    IdxFlag = True
    OptFlag = True
    NumOfTCRecs = NumOfIdx
  End If
  
  frmVATaxShowPctComp.Label1 = "Calculating Personal Penalty"
  frmVATaxShowPctComp.Show , Me
  EnableCloseButton Me.hwnd, False
  cmdTable.Enabled = False
  cmdExit.Enabled = False
  cmdProcess.Enabled = False
  GoSub PrintHeader
  PrntCnt = 0
  For x = 1 To NumOfTCRecs
    If IdxFlag = True Then
      Get TCHandle, IdxArray(x), TaxCust
    Else
      Get TCHandle, x, TaxCust
    End If
    If TaxCust.Deleted <> 0 Then GoTo SkipIt
    If TaxCust.Penalty = "N" Then GoTo SkipIt
    ThisBal = GetCustPersBalance(TaxCust.Acct, -1) 'added 1/24/07 because if the user
    'chooses to not include PPTRA discounts then if the customer has paid in full
    'including the PPTRA discount it appears as if they still have a balance
    If ThisBal = 0 Then GoTo SkipIt
    CustAcct = TaxCust.Acct
    If TaxCust.LastTrans > 0 Then
      NextRec = TaxCust.LastTrans
      Do While NextRec > 0
        Get TTHandle, NextRec, TaxTrans
        If TaxTrans.TaxYear = CurTaxYr Then
          If TaxTrans.TranType = 1 And TaxTrans.BillType = "P" Then
            If PPTRAYN = "Y" Then
              If TaxTrans.PPTRADisc > 0 Then
                TaxTrans.PPTRADisc = 0
              ElseIf TaxTrans.Revenue.Future1 > 0 Then
                TaxTrans.Revenue.Principle1 = OldRound(TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Future1)
              End If
            End If
            Balance = TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3
            Balance = OldRound(Balance + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
            Balance = OldRound(Balance + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2)
            Balance = OldRound(Balance + TaxTrans.Revenue.RevOpt3)
            Balance = OldRound(Balance - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd))
            Balance = OldRound(Balance - (TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd))
            Balance = OldRound(Balance - (TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd))
            Balance = OldRound(Balance - TaxTrans.Revenue.RevOpt3Pd)
            Balance = OldRound(Balance - OldRound(TaxTrans.PPTRADisc - TaxTrans.PPTRARmvl))
            If Balance > 0 Then
              ChargeThis = FigurePenalty(Balance, PRHandle, ThisRec)
              ChargeThis = OldRound(ChargeThis)
              BillNumber$ = TaxTrans.Description
              sb = InStr(BillNumber$, "Bill #")
              If sb > 0 Then
                RSet BillNumber$ = Mid$(TaxTrans.Description, sb + 6, 10)
              End If
              If ChargeThis > 0 Then
                PenRec.CustRec = CustAcct&
                PenRec.CustName = TaxCust.CustName
                LSet CustName = QPTrim$(TaxCust.CustName)
                PenRec.TaxYear = TaxTrans.TaxYear
                PenRec.Amount = ChargeThis
                PenRec.BillNumber = QPTrim$(BillNumber$)
                PenRec.BillRec = NextRec&
                PenRec.CurYear = CurTaxYr
                PenRec.BillType = "P"
                PenRec.Balance = Balance
                PenRec.CustPin = TaxCust.PIN
                PenRec.PersPin = TaxTrans.PersPin
                PenRecord& = PenRecord& + 1
                Put INHandle, PenRecord, PenRec
                TotCnt = TotCnt + 1
                TotPen = OldRound(TotPen + ChargeThis)
                TotBal = OldRound(TotBal + Balance#)
                If LineCnt >= MaxLines Then
                  Print #RptHandle, FF$
                  GoSub PrintHeader
                End If
                GoSub PrintIt
                PrntCnt = PrntCnt + 1
              End If
            End If
          End If
        End If
        NextRec = TaxTrans.LastTrans
      Loop
    End If
SkipIt:
    frmVATaxShowPctComp.ShowPctComp x, NumOfTCRecs
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      cmdExit.Enabled = True
      cmdTable.Enabled = True
      cmdProcess.Enabled = True
      EnableCloseButton Me.hwnd, True
      Exit Sub
    End If
  Next x
  Unload frmVATaxShowPctComp
  cmdTable.Enabled = True
  cmdExit.Enabled = True
  cmdProcess.Enabled = True
  EnableCloseButton Me.hwnd, True
  If LineCnt > MaxLines - 10 Then
    Print #RptHandle, FF$
    GoSub PrintHeader
  End If
  GoSub PrintSummary
  Close
  If PrntCnt = 0 Then
    Call TaxMsg(800, "No customers qualified for a penalty calculation with the parameters entered.")
    Exit Sub
  End If
  
  ViewPrint RptFile$, "Processing Personal Penalty Amounts", True
  
  Exit Sub
  
PrintHeader:
  Page = Page + 1
  Print #RptHandle, Tab(20); "Property Tax Billing: Personal Penalty Calculations For Tax Year " + Using$("###0", CurTaxYr)
  Print #RptHandle, "Town Of: " + TownName
  Print #RptHandle, "Date: " + CStr(Date)
  Print #RptHandle, "Acct #"; Tab(10); "Customer Name"; Tab(54); "Tax Year"; Tab(68); "Bill #"; Tab(80); "Base Balance"; Tab(96); "Penalty"
  Print #RptHandle, String(102, "-")
  LineCnt = 5
  Return
  
PrintSummary:
  Print #RptHandle,
  Print #RptHandle, String(102, "-")
  Print #RptHandle, "Total Transaction Count: " + Using$("#####0", TotCnt)
  Print #RptHandle, "Total Base Balance:    " + Using$("$###,###,##0.00", TotBal)
  Print #RptHandle, "Total Penalty Charges: " + Using$("$###,###,##0.00", TotPen)
  
  Return
  
PrintIt:
  Print #RptHandle, Using$("#######0", CustAcct); Tab(10); CustName; Tab(56); Using$("###0", TaxTrans.TaxYear);
  Print #RptHandle, Tab(66); Using$("######0", CInt(BillNumber)); Tab(79); Using$("$#,###,##0.00", Balance#); Tab(92); Using$("$###,##0.00", ChargeThis)
  LineCnt = LineCnt + 1
  
  Return
  
PrintSingleBill:
  Call GetGCustList
  
  frmVATaxShowPctComp.Label1 = "Calculating Personal Penalty"
  frmVATaxShowPctComp.Show , Me
  EnableCloseButton Me.hwnd, False
  cmdTable.Enabled = False
  cmdExit.Enabled = False
  cmdProcess.Enabled = False
  GoSub PrintHeader
  PrntCnt = 0
  For x = 1 To GCustCnt
    Get TCHandle, GCustList(x), TaxCust
    If TaxCust.Deleted <> 0 Then GoTo SkipIt2
    If TaxCust.Penalty = "N" Then GoTo SkipIt2
    CustAcct = TaxCust.Acct
    If TaxCust.LastTrans > 0 Then
      NextRec = TaxCust.LastTrans
      Do While NextRec > 0
        Get TTHandle, NextRec, TaxTrans
        If NextRec >= GFirstTrans And NextRec <= GLastTrans Then
          If TaxTrans.TranType = 1 And TaxTrans.BillType = "P" Then
            Balance = TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3
            Balance = OldRound(Balance + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
            Balance = OldRound(Balance + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2)
            Balance = OldRound(Balance + TaxTrans.Revenue.RevOpt3)
            Balance = OldRound(Balance - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd))
            Balance = OldRound(Balance - (TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd))
            Balance = OldRound(Balance - (TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd))
            Balance = OldRound(Balance - TaxTrans.Revenue.RevOpt3Pd)
            Balance = OldRound(Balance - OldRound(TaxTrans.PPTRADisc - TaxTrans.PPTRARmvl))
            If Balance > 0 Then
              ChargeThis = FigurePenalty(Balance, PRHandle, ThisRec)
              ChargeThis = OldRound(ChargeThis)
              BillNumber$ = TaxTrans.Description
              sb = InStr(BillNumber$, "Bill #")
              If sb > 0 Then
                RSet BillNumber$ = Mid$(TaxTrans.Description, sb + 6, 10)
              End If
              If ChargeThis > 0 Then
                PenRec.CustRec = CustAcct&
                PenRec.CustName = TaxCust.CustName
                LSet CustName = QPTrim$(TaxCust.CustName)
                PenRec.TaxYear = TaxTrans.TaxYear
                PenRec.Amount = ChargeThis
                PenRec.BillNumber = QPTrim(BillNumber$)
                PenRec.BillRec = NextRec&
                PenRec.CurYear = CurTaxYr
                PenRec.BillType = "P"
                PenRec.Balance = Balance
                PenRec.CustPin = TaxCust.PIN
                PenRec.PersPin = TaxTrans.PersPin
                PenRecord& = PenRecord& + 1
                Put INHandle, PenRecord, PenRec
                TotCnt = TotCnt + 1
                TotPen = OldRound(TotPen + ChargeThis)
                TotBal = OldRound(TotBal + Balance#)
                If LineCnt >= MaxLines Then
                  Print #RptHandle, FF$
                  GoSub PrintHeader
                End If
                GoSub PrintIt
                PrntCnt = PrntCnt + 1
              End If
            End If
          End If
        End If
        NextRec = TaxTrans.LastTrans
      Loop
    End If
SkipIt2:
    frmVATaxShowPctComp.ShowPctComp x, GCustCnt
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      cmdExit.Enabled = True
      cmdTable.Enabled = True
      cmdProcess.Enabled = True
      EnableCloseButton Me.hwnd, True
      Exit Sub
    End If
  Next x
  Unload frmVATaxShowPctComp
  cmdTable.Enabled = True
  cmdExit.Enabled = True
  cmdProcess.Enabled = True
  EnableCloseButton Me.hwnd, True
  If LineCnt > MaxLines - 10 Then
    Print #RptHandle, FF$
    GoSub PrintHeader
  End If
  GoSub PrintSummary
  
  Close
  If PrntCnt = 0 Then
    Call TaxMsg(800, "No customers qualified for a penalty calculation with the parameters entered.")
    Exit Sub
  End If
  
  ViewPrint RptFile$, "Processing Personal Penalty Amounts", True
  
  Exit Sub

  
End Sub

Private Sub ProcessRealText()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long
  Dim TblRec As PenaltyRateTablesType
  Dim PRHandle As Integer
  Dim NumOfPRRecs As Integer
  Dim Balance As Double
  Dim ChargeThis As Double
  Dim PenRec As PenaltyRecType
  Dim INHandle As Integer
  Dim NumOfINRecs As Long
  Dim CurTaxYr As Integer
  Dim GCurTaxYr As Integer
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim NextRec As Long, PenRecord&
  Dim y As Long, sb As Integer
  Dim BillNumber$, CustAcct&
  Dim ThisRec As Integer
  Dim RptHandle As Integer, FF$
  Dim RptFile$, TownName$, Page As Integer
  Dim MaxLines As Integer
  Dim LineCnt As Integer
  Dim TotCnt As Long
  Dim TotPen As Double
  Dim TotBal As Double
  Dim CustName As String * 45
  Dim IdxRec As CustNameIdxType
  Dim IdxHandle As Integer
  Dim NumOfIdx As Long
  Dim IdxFlag As Boolean
  Dim SearchRec As SrchNameIdxType
  Dim SHandle As Integer
  Dim OptFlag As Boolean
  Dim OptRec As OptCustIdxType
  Dim OHandle As Integer
  Dim NumOfORecs As Long
  Dim PrntCnt As Long
  
  MaxLines = 58
  FF$ = Chr(12)
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close
  GCurTaxYr = TaxMasterRec.RTaxYear
  TownName = QPTrim$(TaxMasterRec.Name)
  CurTaxYr = CInt(fptxtCurrRYear.Text)
  If GCurTaxYr <> CurTaxYr Then
    If TaxMsgWOpts(700, "The system real tax year (" + CStr(GCurTaxYr) + ") is not the same as the tax year entered on this form (" + CStr(CurTaxYr) + "). If you wish to continue anyway then press F10. Otherwise, press ESC to edit.", "F10 Continue", "ESC Exit") = "abort" Then
      Unload frmVATaxMsgWOpts
      fptxtCurrRYear.SetFocus
      Exit Sub
    Else
      Unload frmVATaxMsgWOpts
      MainLog ("WARNING: Issued because when calculating real penalty amounts the user used the year " + CStr(CurTaxYr) + " instead of the system current real tax year of " + CStr(GCurTaxYr) + ".")
    End If
  End If
  
  OpenTaxPenRateTbls PRHandle, NumOfPRRecs
  For x = 1 To NumOfPRRecs
    Get PRHandle, x, TblRec
    If TblRec.BillType = "R" Then
      ThisRec = x
    End If
  Next x
  If ThisRec = 0 Then
    Call TaxMsg(900, "ERROR: There was a problem determining which rate table to use. Please save the real rate tables again.")
    Close
    Exit Sub
  End If
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  If NumOfTCRecs = 0 Then
    Call TaxMsg(900, "No customers have been saved")
    Close
    Exit Sub
  End If
  
  If Exist(TaxRPenFile) Then
    KillFile TaxRPenFile
  End If
  
  OpenRPenRecFile INHandle, NumOfINRecs
  OpenTaxTransFile TTHandle, NumOfTTRecs
  RptFile$ = "TAXRPTS\TAXRPEN.RPT"     'Report File Name
  RptHandle = FreeFile
  Open RptFile$ For Output As #RptHandle
  
  If ChkAll.Value = 0 Then GoTo PrintSingleBill
  
  IdxFlag = False
  OptFlag = False
  If QPTrim$(fpcmbPrintOrder.Text) = "2) Customer Name Order" Then
    OpenNameIdxFile IdxHandle, NumOfIdx
    ReDim IdxArray(1 To NumOfIdx) As Long

    For x = 1 To NumOfIdx
      Get IdxHandle, x, IdxRec
      IdxArray(x) = IdxRec.CustRec
    Next x
    Close IdxHandle
    IdxFlag = True
    NumOfTCRecs = NumOfIdx
  ElseIf QPTrim$(fpcmbPrintOrder.Text) = "3) Search Name Order" Then
    OpenSrchNameIdxFile SHandle, NumOfIdx
    If NumOfIdx = 0 Then
      frmVATaxMsg.Label1.Caption = "There are no search names indexed."
      frmVATaxMsg.Label1.Top = 900
      frmVATaxMsg.Show
      fpcmbPrintOrder.SetFocus
      Close SHandle
      Exit Sub
    End If
    ReDim IdxArray(1 To NumOfIdx) As Long

    For x = 1 To NumOfIdx
      Get SHandle, x, SearchRec
      IdxArray(x) = SearchRec.CustRec
    Next x
    Close SHandle
    IdxFlag = True
    NumOfTCRecs = NumOfIdx
  ElseIf QPTrim$(fpcmbPrintOrder.Text) = "4) " + ThisOpt + " Order" Then
    OpenCustOptSearchFile OHandle, NumOfIdx
    If NumOfIdx = 0 Then
      Call TaxMsg(900, "There are no " + ThisOpt + "descriptions indexed.")
      Close OHandle
      Exit Sub
    End If
    ReDim IdxArray(1 To NumOfIdx) As Long
    
    For x = 1 To NumOfIdx
      Get OHandle, x, OptRec
      IdxArray(x) = OptRec.CustRec
    Next x
    Close OHandle
    IdxFlag = True
    OptFlag = True
    NumOfTCRecs = NumOfIdx
  End If
  
  GoSub PrintHeader
  frmVATaxShowPctComp.Label1 = "Calculating Real Penalty"
  frmVATaxShowPctComp.Show , Me
  EnableCloseButton Me.hwnd, False
  cmdTable.Enabled = False
  cmdExit.Enabled = False
  cmdProcess.Enabled = False
  PrntCnt = 0
  For x = 1 To NumOfTCRecs
    If IdxFlag = True Then
      Get TCHandle, IdxArray(x), TaxCust
    Else
      Get TCHandle, x, TaxCust
    End If
    If TaxCust.Deleted <> 0 Then GoTo SkipIt
    If TaxCust.FirstPropRec = 0 Then GoTo SkipIt
    If TaxCust.Penalty = "N" Then GoTo SkipIt
    CustAcct = TaxCust.Acct
    If TaxCust.LastTrans > 0 Then
      NextRec = TaxCust.LastTrans
      
      Do While NextRec > 0
        Get TTHandle, NextRec, TaxTrans
        If TaxTrans.TaxYear = CurTaxYr Then
          If TaxTrans.TranType = 1 And TaxTrans.BillType = "R" Then
            Balance = TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Collection
            Balance = OldRound(Balance + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2)
            Balance = OldRound(Balance + TaxTrans.Revenue.RevOpt3)
            Balance = OldRound(Balance - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.CollectionPd))
            Balance = OldRound(Balance - (TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd))
            Balance = OldRound(Balance - TaxTrans.Revenue.RevOpt3Pd)
            If Balance > 0 Then
              ChargeThis = FigurePenalty(Balance, PRHandle, ThisRec)
              ChargeThis = OldRound(ChargeThis)
              BillNumber$ = TaxTrans.Description
              sb = InStr(BillNumber$, "Bill #")
              If sb > 0 Then
                BillNumber$ = Mid$(TaxTrans.Description, sb + 6, 10)
              End If
              If ChargeThis > 0 Then
                PenRec.CustRec = CustAcct&
                PenRec.CustName = TaxCust.CustName
                LSet CustName = QPTrim$(TaxCust.CustName)
                PenRec.TaxYear = TaxTrans.TaxYear
                PenRec.Amount = ChargeThis
                PenRec.BillNumber = QPTrim$(BillNumber$)
                PenRec.BillRec = NextRec&
                PenRec.CurYear = CurTaxYr
                PenRec.BillType = "R"
                PenRec.Balance = Balance
                PenRec.CustPin = TaxCust.PIN
                PenRec.RealPin = TaxTrans.RealPin
                PenRecord& = PenRecord& + 1
                Put INHandle, PenRecord, PenRec
                TotCnt = TotCnt + 1
                TotPen = OldRound(TotPen + ChargeThis)
                TotBal = OldRound(TotBal + Balance#)
                If LineCnt > MaxLines Then
                  Print #RptHandle, FF$
                  GoSub PrintHeader
                End If
                GoSub PrintIt
                PrntCnt = PrntCnt + 1
              End If
            End If
          End If
        End If
        NextRec = TaxTrans.LastTrans
      Loop
    End If
SkipIt:
    frmVATaxShowPctComp.ShowPctComp x, NumOfTCRecs
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      cmdTable.Enabled = True
      cmdExit.Enabled = True
      cmdProcess.Enabled = True
      EnableCloseButton Me.hwnd, True
      Exit Sub
    End If
  Next x
  
  Unload frmVATaxShowPctComp
  cmdTable.Enabled = True
  cmdExit.Enabled = True
  cmdProcess.Enabled = True
  EnableCloseButton Me.hwnd, True
  
  GoSub PrintSummary
  Close
  If PrntCnt = 0 Then
    Call TaxMsg(800, "No customers qualified for a penalty calculation with the parameters entered.")
    Exit Sub
  End If
  
  ViewPrint RptFile$, "Processing Real Penalty Amounts", True
  
  Exit Sub
  
PrintHeader:
  Page = Page + 1
  Print #RptHandle, Tab(20); "Property Tax Billing: Real Penalty Calculations For Tax Year " + Using$("###0", CurTaxYr)
  Print #RptHandle, "Town Of: " + TownName
  Print #RptHandle, "Date: " + CStr(Date)
  Print #RptHandle, "Acct #"; Tab(10); "Customer Name"; Tab(54); "Tax Year"; Tab(68); "Bill #"; Tab(80); "Base Balance"; Tab(96); "Penalty"
  Print #RptHandle, String(102, "-")
  LineCnt = 5
  
  Return
  
PrintSummary:
  Print #RptHandle,
  Print #RptHandle, String(102, "-")
  Print #RptHandle, "Total Transaction Count: " + Using$("#####0", TotCnt)
  Print #RptHandle, "Total Base Balance:    " + Using$("$###,###,##0.00", TotBal)
  Print #RptHandle, "Total Penalty Charges: " + Using$("$###,###,##0.00", TotPen)
  
  Return
  
PrintIt:
  Print #RptHandle, Using$("#######0", CustAcct); Tab(10); CustName; Tab(56); Using$("###0", TaxTrans.TaxYear);
  Print #RptHandle, Tab(66); Using$("######0", CInt(BillNumber)); Tab(79); Using$("$#,###,##0.00", Balance#); Tab(92); Using$("$###,##0.00", ChargeThis)
  LineCnt = LineCnt + 1
  
  Return
  
PrintSingleBill:
  Call GetGCustList
  
  GoSub PrintHeader
  frmVATaxShowPctComp.Label1 = "Calculating Real Penalty"
  frmVATaxShowPctComp.Show , Me
  EnableCloseButton Me.hwnd, False
  cmdTable.Enabled = False
  cmdExit.Enabled = False
  cmdProcess.Enabled = False
  PrntCnt = 0
  For x = 1 To GCustCnt
    Get TCHandle, GCustList(x), TaxCust
    If TaxCust.Deleted <> 0 Then GoTo SkipIt
'    If TaxCust.FirstPropRec = 0 Then GoTo SkipIt
    If TaxCust.Penalty = "N" Then GoTo SkipIt
    CustAcct = TaxCust.Acct
    If TaxCust.LastTrans > 0 Then
      NextRec = TaxCust.LastTrans
      
      Do While NextRec > 0
        Get TTHandle, NextRec, TaxTrans
        If NextRec >= GFirstTrans And NextRec <= GLastTrans Then
          If TaxTrans.TranType = 1 And TaxTrans.BillType = "R" Then
            Balance = TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Collection
            Balance = OldRound(Balance + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2)
            Balance = OldRound(Balance + TaxTrans.Revenue.RevOpt3)
            Balance = OldRound(Balance - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.CollectionPd))
            Balance = OldRound(Balance - (TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd))
            Balance = OldRound(Balance - TaxTrans.Revenue.RevOpt3Pd)
            If Balance > 0 Then
              ChargeThis = FigurePenalty(Balance, PRHandle, ThisRec)
              ChargeThis = OldRound(ChargeThis)
              BillNumber$ = TaxTrans.Description
              sb = InStr(BillNumber$, "Bill #")
              If sb > 0 Then
                BillNumber$ = Mid$(TaxTrans.Description, sb + 6, 10)
              End If
              If ChargeThis > 0 Then
                PenRec.CustRec = CustAcct&
                PenRec.CustName = TaxCust.CustName
                LSet CustName = QPTrim$(TaxCust.CustName)
                PenRec.TaxYear = TaxTrans.TaxYear
                PenRec.Amount = ChargeThis
                PenRec.BillNumber = QPTrim$(BillNumber$)
                PenRec.BillRec = NextRec&
                PenRec.CurYear = CurTaxYr
                PenRec.BillType = "R"
                PenRec.Balance = Balance
                PenRec.CustPin = TaxCust.PIN
                PenRec.RealPin = TaxTrans.RealPin
                PenRecord& = PenRecord& + 1
                Put INHandle, PenRecord, PenRec
                TotCnt = TotCnt + 1
                TotPen = OldRound(TotPen + ChargeThis)
                TotBal = OldRound(TotBal + Balance#)
                If LineCnt > MaxLines Then
                  Print #RptHandle, FF$
                  GoSub PrintHeader
                End If
                GoSub PrintIt
                PrntCnt = PrntCnt + 1
              End If
            End If
          End If
        End If
        NextRec = TaxTrans.LastTrans
      Loop
    End If
SkipIt2:
    frmVATaxShowPctComp.ShowPctComp x, GCustCnt
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      cmdTable.Enabled = True
      cmdExit.Enabled = True
      cmdProcess.Enabled = True
      EnableCloseButton Me.hwnd, True
      Exit Sub
    End If
  Next x
  
  Unload frmVATaxShowPctComp
  cmdTable.Enabled = True
  cmdExit.Enabled = True
  cmdProcess.Enabled = True
  EnableCloseButton Me.hwnd, True
  
  GoSub PrintSummary
  Close
  If PrntCnt = 0 Then
    Call TaxMsg(800, "No customers qualified for a penalty calculation with the parameters entered.")
    Exit Sub
  End If
  
  ViewPrint RptFile$, "Processing Real Penalty Amounts", True
  
  Exit Sub

End Sub

Private Sub LoadList()
  Dim PostRec As TaxBillPostDateType
  Dim PostHandle As Integer
  Dim NumOfPostRecs As Long
  Dim x As Long
  Dim CurrYear As Integer
  Dim ThisType As String * 1
  
  fpList.Action = ActionClear
  ThisType = Mid(fpcmbType.Text, 1, 1)
  If ThisType = "R" Then
    CurrYear = CInt(fptxtCurrRYear.Text)
  ElseIf ThisType = "P" Then
    CurrYear = CInt(fptxtCurrPYear.Text)
  End If
  
  If Exist(TaxBillPostDateFile) Then
    OpenBillPostDateFile PostHandle, NumOfPostRecs
    For x = 1 To NumOfPostRecs
      Get PostHandle, x, PostRec
      If PostRec.PostYear = CurrYear And PostRec.BillType = ThisType Then
        fpList.InsertRow = MakeRegDate(PostRec.PostDate) + Chr(9) + PostRec.BackUpName + Chr(9) + CStr(PostRec.FirstTrans) + Chr(9) + CStr(PostRec.LastTrans)
      End If
    Next x
    Close PostHandle
  End If
  
  If fpList.ListCount > 0 Then
    fpList.Selected(0) = True
    fpList.ListIndex = 0
  End If
  
End Sub

Private Sub GetGCustList()
  Dim PRRec As VARETaxBillType
  Dim PPRec As VAPPTaxBillType
  Dim PRHandle As Integer
  Dim PPHandle As Integer
  Dim NumOfPRRecs As Long
  Dim NumOfPPRecs As Long
  Dim MyPath$, ThisFile$
  Dim x As Integer, y As Integer, yCnt As Integer
  Dim RCnt As Integer '3/10/08
  Dim PCnt As Integer '3/10/08
  Dim TCustRec As Long
  
  MyPath = StartPath + "\TAXBILLBU\"
  fpList.Col = 1
  fpList.Selected(fpList.ListIndex) = True
  fpList.Row = fpList.ListIndex
  ThisFile = QPTrim$(fpList.ColText)
  fpList.Col = 2
  GFirstTrans = CLng(fpList.ColText)
  fpList.Col = 3
  GLastTrans = CLng(fpList.ColText)
  GCustCnt = 0
  If fpcmbType.Text = "REAL" Then
    OpenRealPostedReprintFile PRHandle, NumOfPRRecs, ThisFile
'    ReDim GCustList(1 To NumOfPRRecs) As Long 'commented 3/10/08
    For x = 1 To NumOfPRRecs
      Get PRHandle, x, PRRec
      For y = 1 To yCnt '3/10/08
        If PRRec.CustRec = GCustList(y) Then
          Exit For
        End If
      Next y
      If y > yCnt Then 'added 3/10/08
        yCnt = yCnt + 1
        ReDim Preserve GCustList(1 To yCnt) As Long
        GCustList(yCnt) = PRRec.CustRec
        GCustCnt = yCnt 'NumOfPRRecs 'added 3/10/08
      End If
    Next x
'    GCustCnt = RCnt 'NumOfPRRecs 'commented 3/10/08
    Close PRHandle
  ElseIf fpcmbType.Text = "PERSONAL" Then
    OpenPersPostedReprintFile PPHandle, NumOfPPRecs, ThisFile
'    ReDim GCustList(1 To NumOfPPRecs) As Long'commented 3/10/08
    For x = 1 To NumOfPPRecs
      Get PPHandle, x, PPRec
      For y = 1 To yCnt '3/10/08
        If PPRec.CustRec = GCustList(y) Then '3/10/08
          Exit For '3/10/08
        End If '3/10/08
      Next y '3/10/08
      If y > yCnt Then 'added 3/10/08
        yCnt = yCnt + 1 '3/10/08
        ReDim Preserve GCustList(1 To yCnt) As Long '3/10/08
        GCustList(yCnt) = PPRec.CustRec
        GCustCnt = yCnt 'NumOfPRRecs 'added 3/10/08
      End If
    Next x
'    GCustCnt = NumOfPPRecs 'commented 3/10/08
    Close PPHandle
  End If

End Sub

Private Sub fptxtCurrPYear_Change()
  Call LoadList
End Sub

Private Sub fptxtCurrRYear_Change()
  Call LoadList
End Sub

Private Function Check4PPTRADisc(TaxYear As Integer) As Boolean
  Dim TransRec As TaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim x As Long
  
  Check4PPTRADisc = False
  OpenTaxTransFile THandle, NumOfTRecs
  If NumOfTRecs = 0 Then Exit Function
  For x = 1 To NumOfTRecs
    Get THandle, x, TransRec
    If TransRec.TaxYear = TaxYear Then
      If TransRec.PPTRADisc > 0 Or TransRec.Revenue.Future1 > 0 Then
        Check4PPTRADisc = True
        Exit For
      End If
    End If
  Next x
  
  Close THandle
End Function

Private Sub chkPPTRAYN_Click()
  If chkPPTRAYN.Value = 1 Then
    PPTRAYN = "Y"
  Else
    PPTRAYN = "N"
  End If
  
End Sub

