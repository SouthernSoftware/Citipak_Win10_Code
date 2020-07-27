VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmVATaxPenaltyTransRpt 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Penalty Transaction Report"
   ClientHeight    =   8760
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   11655
   Icon            =   "frmVATaxPenaltyTransRpt.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcmbType 
      Height          =   384
      Left            =   3960
      TabIndex        =   0
      Top             =   1920
      Width           =   3372
      _Version        =   196608
      _ExtentX        =   5948
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
      ColDesigner     =   "frmVATaxPenaltyTransRpt.frx":08CA
   End
   Begin LpLib.fpCombo fpcmbPrintOpt 
      Height          =   384
      Left            =   3960
      TabIndex        =   1
      Top             =   6000
      Width           =   3336
      _Version        =   196608
      _ExtentX        =   5884
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
      ColDesigner     =   "frmVATaxPenaltyTransRpt.frx":0BC1
   End
   Begin EditLib.fpDateTime fptxtCurrRYear 
      Height          =   372
      Left            =   6840
      TabIndex        =   2
      Top             =   2880
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
      ControlType     =   1
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
      Left            =   6015
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   7200
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
      ButtonDesigner  =   "frmVATaxPenaltyTransRpt.frx":0EB8
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   492
      Left            =   3228
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   7200
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
      ButtonDesigner  =   "frmVATaxPenaltyTransRpt.frx":1097
   End
   Begin EditLib.fpDateTime fptxtRealYr 
      Height          =   372
      Left            =   6840
      TabIndex        =   10
      Top             =   3960
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
      ControlType     =   1
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
   Begin EditLib.fpDateTime fptxtPersYr 
      Height          =   372
      Left            =   6840
      TabIndex        =   12
      Top             =   4440
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
      ControlType     =   1
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
   Begin EditLib.fpDateTime fptxtCurrPYear 
      Height          =   372
      Left            =   6840
      TabIndex        =   3
      Top             =   3360
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
      ControlType     =   1
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
   Begin VB.Line Line1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   3000
      X2              =   8280
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "System Current Personal Tax Year:"
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
      Left            =   3240
      TabIndex        =   14
      Top             =   3468
      Width           =   3420
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Calculation Personal Tax Year:"
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
      Left            =   3720
      TabIndex        =   13
      Top             =   4548
      Width           =   2940
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Calculation Real Tax Year:"
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
      Left            =   4080
      TabIndex        =   11
      Top             =   4080
      Width           =   2580
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Penalty Transaction Report"
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
      TabIndex        =   9
      Top             =   636
      Width           =   5292
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   660
      Index           =   1
      Left            =   1500
      Top             =   468
      Width           =   8652
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   1212
      Left            =   3600
      Top             =   5400
      Width           =   4092
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   2172
      Left            =   3000
      Top             =   2760
      Width           =   5292
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
      Left            =   4800
      TabIndex        =   8
      Top             =   5580
      Width           =   1812
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "System Current Real Tax Year:"
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
      Left            =   3600
      TabIndex        =   7
      Top             =   2988
      Width           =   3060
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Select A Type"
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
      Left            =   4680
      TabIndex        =   6
      Top             =   1560
      Width           =   1860
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
Attribute VB_Name = "frmVATaxPenaltyTransRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  'Private Temp_Class As Resize_Class
'  Dim CurrTaxYear As Integer
  Dim ThisOpt$
  Dim TownName$
  Dim IncReal As Boolean
  Dim IncPers As Boolean

Private Sub cmdExit_Click()
  frmVATaxPenaltyMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdProcess_Click()
  If fpcmbType.Text = "NO REAL" Then
    Call TaxMsg(900, "No data to report.")
    fpcmbType.SetFocus
    Exit Sub
  ElseIf fpcmbType.Text = "NO PERSONAL" Then
    Call TaxMsg(900, "No data to report.")
    fpcmbType.SetFocus
    Exit Sub
  ElseIf fpcmbType.Text = "REAL" Then
    If fpcmbPrintOpt.Text = "Graphical" Then
      Call PrintRealGraphics
    ElseIf fpcmbPrintOpt.Text = "Text" Then
      Call PrintRealText
    End If
  ElseIf fpcmbType.Text = "PERSONAL" Then
    If fpcmbPrintOpt.Text = "Graphical" Then
      Call PrintPersGraphics
    ElseIf fpcmbPrintOpt.Text = "Text" Then
      Call PrintPersText
    End If
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
    Case Else:
  End Select

End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  'Set Temp_Class = New Resize_Class
  'Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Me.HelpContextID = hlpCalculateI
  Call LoadMe
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("CitiTaxes.exe terminated via menu bar on frmVATaxPenaltyTransRpt.")
      Call Terminate
      End
    End If
  End If

End Sub
Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    'Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
    DoEvents
  End If
End Sub
Private Sub LoadMe()
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim PenTrans As PenaltyRecType
  Dim NumOfRPNRecs As Long
  Dim RPNHandle As Integer
  Dim NumOfPPNRecs As Long
  Dim PPNHandle As Integer
  Dim x As Long, y As Long
  
  NumOfRPNRecs = 0
  NumOfPPNRecs = 0
  IncReal = False
  IncPers = False
  If Exist(TaxRPenFile) Then
    OpenRPenRecFile RPNHandle, NumOfRPNRecs
    IncReal = True
  End If
  If Exist(TaxPPenFile) Then
    OpenPPenRecFile PPNHandle, NumOfPPNRecs
    IncPers = True
  End If
  
  For x = 1 To NumOfRPNRecs
    Get RPNHandle, x, PenTrans
    If PenTrans.DelFlag = False Then
      fptxtRealYr.Text = CStr(PenTrans.CurYear)
      Exit For
    End If
  Next x
  For y = 1 To NumOfPPNRecs
    Get PPNHandle, y, PenTrans
    If PenTrans.DelFlag = False Then
      fptxtPersYr.Text = CStr(PenTrans.CurYear)
      Exit For
    End If
  Next y
  
  If NumOfRPNRecs > 0 And NumOfPPNRecs > 0 Then
    If x <= NumOfRPNRecs And y <= NumOfPPNRecs Then
      fpcmbType.Text = "REAL"
      fpcmbType.AddItem "REAL"
      fpcmbType.AddItem "PERSONAL"
    ElseIf x <= NumOfRPNRecs And y > NumOfPPNRecs Then
      fpcmbType.Text = "REAL"
      fpcmbType.AddItem "REAL"
      fpcmbType.AddItem "NO PERSONAL"
      fptxtPersYr.Text = "NA"
      IncPers = False
    ElseIf x > NumOfRPNRecs And y <= NumOfPPNRecs Then
      fpcmbType.Text = "PERSONAL"
      fpcmbType.AddItem "NO REAL"
      fpcmbType.AddItem "PERSONAL"
      fptxtRealYr.Text = "NA"
      IncReal = False
    End If
  ElseIf NumOfRPNRecs > 0 And NumOfPPNRecs = 0 Then
    If x <= NumOfRPNRecs Then
      fpcmbType.Text = "REAL"
      fpcmbType.AddItem "REAL"
      fpcmbType.AddItem "NO PERSONAL"
      fptxtPersYr.Text = "NA"
      IncPers = False
    End If
  ElseIf NumOfRPNRecs = 0 And NumOfPPNRecs > 0 Then
    If y <= NumOfPPNRecs Then
      fpcmbType.Text = "PERSONAL"
      fpcmbType.AddItem "NO REAL"
      fpcmbType.AddItem "PERSONAL"
      fptxtRealYr.Text = "NA"
      IncReal = False
    End If
  End If

  Close RPNHandle
  Close PPNHandle
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  TownName = QPTrim$(TaxMasterRec.Name)
  
  fptxtCurrRYear.Text = CStr(TaxMasterRec.RTaxYear)
'  CurrTaxYear = TaxMasterRec.RTaxYear
  
  fptxtCurrPYear.Text = CStr(TaxMasterRec.PTaxYear)
'  CurrTaxYear = TaxMasterRec.PTaxYear
  
  fpcmbPrintOpt.Text = "Graphical"
  fpcmbPrintOpt.AddItem "Graphical"
  fpcmbPrintOpt.AddItem "Text"

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


Private Sub PrintRealGraphics()
  Dim PenRec As PenaltyRecType
  Dim INHandle As Integer
  Dim NumOfINRecs As Long
  Dim x As Long
  Dim RptHandle As Integer
  Dim RptFile$, TownName$
  Dim dlm$
  Dim TotCnt As Long
  Dim TotPen As Double
  Dim TotBal As Double
  
  dlm$ = "~"
  RptFile$ = "TAXRPTS\TAXRPEN.RPT"     'Report File Name
  RptHandle = FreeFile
  Open RptFile$ For Output As #RptHandle
  OpenRPenRecFile INHandle, NumOfINRecs
  For x = 1 To NumOfINRecs
    Get INHandle, x, PenRec
    If PenRec.DelFlag = True Then GoTo Skip
    TotPen = OldRound(TotPen + PenRec.Amount)
    TotBal = OldRound(TotBal + PenRec.Balance)
    TotCnt = TotCnt + 1
    '                     0                1                    2                          3
    Print #RptHandle, TownName; dlm; PenRec.Balance; dlm; PenRec.CustRec; dlm; QPTrim$(PenRec.CustName); dlm;
    '                        4                   5                    6                  7             8           9
    Print #RptHandle, PenRec.TaxYear; dlm; PenRec.Amount; dlm; PenRec.BillNumber; dlm; TotPen; dlm; TotBal; dlm; TotCnt
Skip:
  Next x
  
  Close
  arVATaxRPenRpt.Show
  
End Sub
Private Sub PrintRealText()
  Dim PenRec As PenaltyRecType
  Dim INHandle As Integer
  Dim NumOfINRecs As Long
  Dim x As Long
  Dim BillNumber$, CustAcct&
  Dim ThisRec As Integer
  Dim LineCnt As Integer
  Dim FF$, MaxLines As Integer
  Dim RptHandle As Integer
  Dim RptFile$, Page As Integer
  Dim TotCnt As Long
  Dim TotPen As Double
  Dim TotBal As Double
  Dim CustName As String * 45
  Dim CurTaxYr As Integer
  
  CurTaxYr = CInt(fptxtCurrRYear.Text)
  MaxLines = 58
  FF$ = Chr(12)
  OpenRPenRecFile INHandle, NumOfINRecs
  RptFile$ = "TAXRPTS\TAXRPEN.PRN"     'Report File Name
  RptHandle = FreeFile
  Open RptFile$ For Output As #RptHandle
  GoSub PrintHeader
  
  For x = 1 To NumOfINRecs
    Get INHandle, x, PenRec
    If PenRec.DelFlag = True Then GoTo Skip
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintHeader
    End If
    TotCnt = TotCnt + 1
    TotPen = OldRound(TotPen + PenRec.Amount)
    TotBal = OldRound(TotBal + PenRec.Balance)
    LSet CustName = QPTrim$(PenRec.CustName)
    Print #RptHandle, Using$("#######0", PenRec.CustRec); Tab(10); CustName; Tab(56); Using$("###0", PenRec.TaxYear);
    Print #RptHandle, Tab(66); Using$("######0", CInt(PenRec.BillNumber)); Tab(79); Using$("$#,###,##0.00", PenRec.Balance); Tab(92); Using$("$###,##0.00", PenRec.Amount)
    LineCnt = LineCnt + 1
Skip:
  Next x
  
  If LineCnt > MaxLines - 10 Then
    Print #RptHandle, FF$
    GoSub PrintHeader
  End If
  GoSub PrintSummary
  Close
  
  ViewPrint RptFile$, "Printing Real Penalty Amounts", True
  
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

End Sub
Private Sub PrintPersGraphics()
  Dim PenRec As PenaltyRecType
  Dim INHandle As Integer
  Dim NumOfINRecs As Long
  Dim x As Long
  Dim RptHandle As Integer
  Dim RptFile$, TownName$
  Dim dlm$
  Dim TotCnt As Long
  Dim TotPen As Double
  Dim TotBal As Double
  
  dlm$ = "~"
  RptFile$ = "TAXRPTS\TAXPPEN.RPT"     'Report File Name
  RptHandle = FreeFile
  Open RptFile$ For Output As #RptHandle
  OpenPPenRecFile INHandle, NumOfINRecs
  For x = 1 To NumOfINRecs
    Get INHandle, x, PenRec
    If PenRec.DelFlag = True Then GoTo Skip
    TotPen = OldRound(TotPen + PenRec.Amount)
    TotBal = OldRound(TotBal + PenRec.Balance)
    TotCnt = TotCnt + 1
    '                     0                 1                    2                       3
    Print #RptHandle, TownName; dlm; PenRec.Balance; dlm; PenRec.CustRec; dlm; QPTrim$(PenRec.CustName); dlm;
    '                        4                   5                    6                  7             8           9
    Print #RptHandle, PenRec.TaxYear; dlm; PenRec.Amount; dlm; PenRec.BillNumber; dlm; TotPen; dlm; TotBal; dlm; TotCnt
Skip:
  Next x
  
  Close
  arVATaxPPenRpt.Show

End Sub
Private Sub PrintPersText()
  Dim PenRec As PenaltyRecType
  Dim INHandle As Integer
  Dim NumOfINRecs As Long
  Dim x As Long
  Dim BillNumber$, CustAcct&
  Dim ThisRec As Integer
  Dim LineCnt As Integer
  Dim FF$, MaxLines As Integer
  Dim RptHandle As Integer
  Dim RptFile$, Page As Integer
  Dim TotCnt As Long
  Dim TotPen As Double
  Dim TotBal As Double
  Dim CustName As String * 45
  Dim CurTaxYr As Integer
  
  CurTaxYr = CInt(fptxtCurrPYear.Text)
  MaxLines = 58
  FF$ = Chr(12)
  OpenPPenRecFile INHandle, NumOfINRecs
  RptFile$ = "TAXRPTS\TAXPPEN.PRN"     'Report File Name
  RptHandle = FreeFile
  Open RptFile$ For Output As #RptHandle
  GoSub PrintHeader
  
  For x = 1 To NumOfINRecs
    Get INHandle, x, PenRec
    If PenRec.DelFlag = True Then GoTo Skip
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintHeader
    End If
    TotCnt = TotCnt + 1
    TotPen = OldRound(TotPen + PenRec.Amount)
    TotBal = OldRound(TotBal + PenRec.Balance)
    LSet CustName = QPTrim$(PenRec.CustName)
    Print #RptHandle, Using$("#######0", PenRec.CustRec); Tab(10); CustName; Tab(56); Using$("###0", PenRec.TaxYear);
    Print #RptHandle, Tab(66); Using$("######0", CInt(PenRec.BillNumber)); Tab(79); Using$("$#,###,##0.00", PenRec.Balance); Tab(92); Using$("$###,##0.00", PenRec.Amount)
    LineCnt = LineCnt + 1
Skip:
  Next x
  
  If LineCnt > MaxLines - 10 Then
    Print #RptHandle, FF$
    GoSub PrintHeader
  End If
  GoSub PrintSummary
  Close
  
  ViewPrint RptFile$, "Printing Personal Penalty Amounts", True
  
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
  
End Sub

Private Sub fpcmbType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbType.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbType.ListIndex = -1
  End If
  If fpcmbType.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbPrintOpt.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

