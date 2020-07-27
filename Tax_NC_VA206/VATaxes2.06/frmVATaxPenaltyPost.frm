VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmVATaxPenaltyPost 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Penalty Calculations Post"
   ClientHeight    =   8760
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   11655
   Icon            =   "frmVATaxPenaltyPost.frx":0000
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
      Left            =   4140
      TabIndex        =   6
      Top             =   2160
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
      ColDesigner     =   "frmVATaxPenaltyPost.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPost 
      Height          =   492
      Left            =   6240
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   7680
      Width           =   1812
      _Version        =   131072
      _ExtentX        =   3196
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
      ButtonDesigner  =   "frmVATaxPenaltyPost.frx":0BC1
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   492
      Left            =   3600
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   7680
      Width           =   1812
      _Version        =   131072
      _ExtentX        =   3196
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
      ButtonDesigner  =   "frmVATaxPenaltyPost.frx":0D9D
   End
   Begin EditLib.fpDateTime fptxtCurrRYear 
      Height          =   372
      Left            =   7020
      TabIndex        =   7
      Top             =   3000
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
   Begin EditLib.fpDateTime fptxtRealYr 
      Height          =   372
      Left            =   7020
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
      Left            =   7020
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
      Left            =   7020
      TabIndex        =   14
      Top             =   3480
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
   Begin VB.Label Label7 
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
      Left            =   3420
      TabIndex        =   15
      Top             =   3588
      Width           =   3420
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Year For Personal:"
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
      Left            =   4500
      TabIndex        =   13
      Top             =   4548
      Width           =   2340
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Year For Real:"
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
      Left            =   4980
      TabIndex        =   11
      Top             =   4068
      Width           =   1860
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   2052
      Left            =   3360
      Top             =   2880
      Width           =   4932
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
      Left            =   3420
      TabIndex        =   9
      Top             =   3108
      Width           =   3420
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Make A Post Selection"
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
      Left            =   4620
      TabIndex        =   8
      Top             =   1800
      Width           =   2340
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "Press F10 To Post."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   372
      Left            =   2688
      TabIndex        =   4
      Top             =   6588
      Width           =   3132
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "Press ESC To Exit."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   372
      Left            =   5808
      TabIndex        =   3
      Top             =   6588
      Width           =   3132
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   2052
      Left            =   2040
      Top             =   5268
      Width           =   7572
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Penalty Calculations Post"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3828
      TabIndex        =   2
      Top             =   960
      Width           =   4020
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000009&
      Height          =   840
      Left            =   2304
      Top             =   720
      Width           =   7020
   End
   Begin VB.Shape Shape6 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   996
      Left            =   2316
      Top             =   600
      Width           =   7020
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "                                                                                Ready to Post Penalty Calculations Transactions? "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   2052
      Left            =   2040
      TabIndex        =   5
      Top             =   5268
      Width           =   7572
   End
End
Attribute VB_Name = "frmVATaxPenaltyPost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class
  Dim IncReal As Boolean
  Dim IncPers As Boolean

Private Sub cmdExit_Click()
  frmVATaxPenaltyMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdPost_Click()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim cnt As Long, Previous&
  Dim DidSome As Long
  Dim TaxTrans As TaxTransactionType
  Dim NewTaxTrans As TaxTransactionType
  Dim ClearTaxTrans As TaxTransactionType
  Dim TaxPenRec As PenaltyRecType
  Dim RPenTrans As PenaltyRecType
  Dim PPenTrans As PenaltyRecType
  Dim NumOfRPRecs As Long
  Dim NumOfPPRecs As Long
  Dim RPHandle As Integer
  Dim PPHandle As Integer
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim x As Long, NextRecord&
  Dim PostReal As Boolean
  Dim PostPers As Boolean
  
  On Error GoTo ERRORSTUFF
  
  PostReal = False
  PostPers = False
  If fpcmbType.Text = "REAL ONLY" Then
    If TaxMsgWOpts(800, "If you are sure you are ready to post REAL ONLY then press F10 to continue. Otherwise, press ESC to abort the post attempt.", "F10 Continue", "ESC Abort") = "abort" Then
      Unload frmVATaxMsgWOpts
      Call TaxMsg(900, "Real penalty post attempt aborted.")
      Close
      Exit Sub
    Else
      Unload frmVATaxMsgWOpts
      PostReal = True
    End If
  ElseIf fpcmbType.Text = "PERSONAL ONLY" Then
    If TaxMsgWOpts(800, "If you are sure you are ready to post PERSONAL ONLY then press F10 to continue. Otherwise, press ESC to abort the post attempt.", "F10 Continue", "ESC Abort") = "abort" Then
      Unload frmVATaxMsgWOpts
      Call TaxMsg(900, "Personal penalty post attempt aborted.")
      Close
      Exit Sub
    Else
      Unload frmVATaxMsgWOpts
      PostPers = True
    End If
  ElseIf fpcmbType.Text = "POST BOTH" Then
    If TaxMsgWOpts(800, "If you are sure you are ready to post both PERSONAL AND REAL then press F10 to continue. Otherwise, press ESC to abort the post attempt.", "F10 Continue", "ESC Abort") = "abort" Then
      Unload frmVATaxMsgWOpts
      Call TaxMsg(800, "Personal and real penalty post combination attempt aborted.")
      Close
      Exit Sub
    Else
      Unload frmVATaxMsgWOpts
      PostReal = True
      PostPers = True
    End If
  End If
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  If IncReal = True Then
    OpenRPenRecFile RPHandle, NumOfRPRecs
  End If
  If IncPers = True Then
    OpenPPenRecFile PPHandle, NumOfPPRecs
  End If
  
  OpenTaxTransFile TTHandle, NumOfTTRecs
  
  If PostReal = True Then
    frmVATaxShowPctComp.Label1 = "Posting Real Penalty"
    frmVATaxShowPctComp.Show , Me
    frmVATaxShowPctComp.cmdCancel.Visible = False
    cmdExit.Enabled = False
    cmdPost.Enabled = False
  
    For cnt& = 1 To NumOfRPRecs
      Get RPHandle, cnt&, TaxPenRec
      If TaxPenRec.DelFlag = 0 Then
      'Update the Bill transaction first
      'TaxPenRec(1).BillRec
        Get TTHandle, TaxPenRec.BillRec, TaxTrans 'get bill trans
        TaxTrans.Revenue.Penalty = OldRound#(TaxTrans.Revenue.Penalty + TaxPenRec.Amount)
        Put #TTHandle, TaxPenRec.BillRec, TaxTrans 'put it back
      'Now make a new clean transaction
        NewTaxTrans = ClearTaxTrans
        NewTaxTrans.TransDate = Date2Num%(Date$)
        NewTaxTrans.TaxYear = TaxPenRec.TaxYear
        NewTaxTrans.TranType = 5       '5=Penalty
        NewTaxTrans.BillType = "R"     'R=Real P=Personal Property C=Combined (NC/GA)
        NewTaxTrans.Amount = TaxPenRec.Amount  'Total Transaction Amount
        NewTaxTrans.Revenue.Penalty = TaxPenRec.Amount
        NewTaxTrans.Description = "Tax Pen on Bill# " + QPTrim$(TaxPenRec.BillNumber)
        NewTaxTrans.Posted2GL = "N"
        NewTaxTrans.CustomerRec = TaxPenRec.CustRec
        NewTaxTrans.CustPin = TaxPenRec.CustPin
        NewTaxTrans.RealPin = TaxPenRec.RealPin
        NewTaxTrans.PersPin = 0
        NewTaxTrans.LastTrans = 0
        NewTaxTrans.BelongTo = TaxPenRec.BillRec
        NewTaxTrans.Revenue.PrePaidAmt = 0
        NewTaxTrans.Revenue.PrePaidBal = OldRound(GetCustRealBalance(TaxPenRec.CustRec, -1))
        If NewTaxTrans.Revenue.PrePaidBal > 0 Then
          NewTaxTrans.Revenue.PrePaidBal = 0
        End If
        NewTaxTrans.Revenue.PrePaidUsed = 0
        NewTaxTrans.OperNum = OperNum
        LSet NewTaxTrans.Padding = ""
        'Increment Transaction File Record Count
        NextRecord& = (LOF(TTHandle) / Len(NewTaxTrans)) + 1
        Put TTHandle, NextRecord&, NewTaxTrans
        'Update the Customer Pointers Now
        Get TCHandle, TaxPenRec.CustRec, TaxCust
      
        If TaxCust.LastTrans = 0 Then
          TaxCust.LastTrans = NextRecord&
          Put TCHandle, TaxPenRec.CustRec, TaxCust
        Else
          Previous& = TaxCust.LastTrans
          TaxCust.LastTrans = NextRecord&
          Put TCHandle, TaxPenRec.CustRec, TaxCust
          Get TTHandle, NextRecord&, NewTaxTrans
          NewTaxTrans.LastTrans = Previous&
          Put TTHandle, NextRecord&, NewTaxTrans
        End If
      End If
      frmVATaxShowPctComp.ShowPctComp cnt, NumOfRPRecs
      If frmVATaxShowPctComp.Out = True Then
        Close
        frmVATaxShowPctComp.Out = False
        Unload frmVATaxShowPctComp
        EnableCloseButton Me.hwnd, True
        cmdExit.Enabled = True
        cmdPost.Enabled = True
        Exit Sub
      End If
    Next
    
    Unload frmVATaxShowPctComp
    EnableCloseButton Me.hwnd, True
    cmdExit.Enabled = True
    cmdPost.Enabled = True
  End If
  
  If PostPers = True Then
    frmVATaxShowPctComp.Label1 = "Posting Personal Penalty"
    frmVATaxShowPctComp.Show , Me
    frmVATaxShowPctComp.cmdCancel.Visible = False
    cmdExit.Enabled = False
    cmdPost.Enabled = False
  
    For cnt& = 1 To NumOfPPRecs
      Get PPHandle, cnt&, TaxPenRec
      If TaxPenRec.DelFlag = 0 Then
      'Update the Bill transaction first
      'TaxPenRec(1).BillRec
        Get TTHandle, TaxPenRec.BillRec, TaxTrans 'get bill trans
        TaxTrans.Revenue.Penalty = OldRound#(TaxTrans.Revenue.Penalty + TaxPenRec.Amount)
        Put #TTHandle, TaxPenRec.BillRec, TaxTrans 'put it back
      'Now make a new clean transaction
        NewTaxTrans = ClearTaxTrans
        NewTaxTrans.TransDate = Date2Num%(Date$)
        NewTaxTrans.TaxYear = TaxPenRec.TaxYear
        NewTaxTrans.TranType = 5       '5=Penalty
        NewTaxTrans.BillType = "P"     'R=Real P=Personal Property C=Combined (NC/GA)
        NewTaxTrans.Amount = TaxPenRec.Amount  'Total Transaction Amount
        NewTaxTrans.Revenue.Penalty = TaxPenRec.Amount
        NewTaxTrans.Description = "Tax Pen on Bill# " + QPTrim$(TaxPenRec.BillNumber)
        NewTaxTrans.Posted2GL = "N"
        NewTaxTrans.CustomerRec = TaxPenRec.CustRec
        NewTaxTrans.CustPin = TaxPenRec.CustPin
        NewTaxTrans.RealPin = 0
        NewTaxTrans.PersPin = TaxPenRec.PersPin
        NewTaxTrans.LastTrans = 0
        NewTaxTrans.BelongTo = TaxPenRec.BillRec
        NewTaxTrans.Revenue.PrePaidAmt = 0
        NewTaxTrans.Revenue.PrePaidBal = OldRound(GetCustPersBalance(TaxPenRec.CustRec, -1))
        If NewTaxTrans.Revenue.PrePaidBal > 0 Then
          NewTaxTrans.Revenue.PrePaidBal = 0
        End If
        NewTaxTrans.Revenue.PrePaidUsed = 0
        NewTaxTrans.OperNum = OperNum
        LSet NewTaxTrans.Padding = ""
        'Increment Transaction File Record Count
        NextRecord& = (LOF(TTHandle) / Len(NewTaxTrans)) + 1
        Put TTHandle, NextRecord&, NewTaxTrans
        'Update the Customer Pointers Now
        Get TCHandle, TaxPenRec.CustRec, TaxCust
      
        If TaxCust.LastTrans = 0 Then
          TaxCust.LastTrans = NextRecord&
          Put TCHandle, TaxPenRec.CustRec, TaxCust
        Else
          Previous& = TaxCust.LastTrans
          TaxCust.LastTrans = NextRecord&
          Put TCHandle, TaxPenRec.CustRec, TaxCust
          Get TTHandle, NextRecord&, NewTaxTrans
          NewTaxTrans.LastTrans = Previous&
          Put TTHandle, NextRecord&, NewTaxTrans
        End If
      End If
      frmVATaxShowPctComp.ShowPctComp cnt, NumOfPPRecs
      If frmVATaxShowPctComp.Out = True Then
        Close
        frmVATaxShowPctComp.Out = False
        Unload frmVATaxShowPctComp
        EnableCloseButton Me.hwnd, True
        cmdExit.Enabled = True
        cmdPost.Enabled = True
        Exit Sub
      End If
    Next
  
    Unload frmVATaxShowPctComp
    EnableCloseButton Me.hwnd, True
    cmdExit.Enabled = True
    cmdPost.Enabled = True
  End If
  Close
  
  'Now Log and Delete the Tax Bill File so Duplicate's Cannot Be Reproduced
      
  If PostReal = True And PostPers = False Then
    Call Savemsg(900, "The real calculations have been posted successfully.")
    MainLog ("Real penalty calculations posted.")
    KillFile TaxRPenFile
  ElseIf PostReal = False And PostPers = True Then
    Call Savemsg(900, "The personal calculations have been posted successfully.")
    MainLog ("Personal penalty calculations posted.")
    KillFile TaxPPenFile
  ElseIf PostReal = True And PostPers = True Then
    Call Savemsg(900, "The real and personal calculations have been posted successfully.")
    MainLog ("Personal penalty and real penalty combined calculations posted.")
    KillFile TaxRPenFile
    KillFile TaxPPenFile
  End If
  
  frmVATaxPenaltyMenu.Show
  DoEvents
  Unload Me
  
  Exit Sub

  
  Return

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPenaltyPost", "cmdPost_Click", Erl)
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
    ClearInUse PWcnt
    Terminate
  

End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Call LoadMe
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("CitiTaxes.exe terminated via menu bar on frmVATaxPenaltyPost.")
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
      Call cmdPost_Click
      KeyCode = 0
  End Select

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
      fpcmbType.Text = "REAL ONLY"
      fpcmbType.AddItem "REAL ONLY"
      fpcmbType.AddItem "PERSONAL ONLY"
      fpcmbType.AddItem "POST BOTH"
    ElseIf x <= NumOfRPNRecs And y > NumOfPPNRecs Then
      fpcmbType.Text = "REAL ONLY"
      fpcmbType.AddItem "REAL ONLY"
      fpcmbType.AddItem "NO PERSONAL"
      fptxtPersYr.Text = "NA"
      IncPers = False
    ElseIf x > NumOfRPNRecs And y <= NumOfPPNRecs Then
      fpcmbType.Text = "PERSONAL ONLY"
      fpcmbType.AddItem "NO REAL"
      fpcmbType.AddItem "PERSONAL ONLY"
      fptxtRealYr.Text = "NA"
      IncReal = False
    End If
  ElseIf NumOfRPNRecs > 0 And NumOfPPNRecs = 0 Then
    If x <= NumOfRPNRecs Then
      fpcmbType.Text = "REAL ONLY"
      fpcmbType.AddItem "REAL ONLY"
      fpcmbType.AddItem "NO PERSONAL"
      fptxtPersYr.Text = "NA"
      IncPers = False
    End If
  ElseIf NumOfRPNRecs = 0 And NumOfPPNRecs > 0 Then
    If y <= NumOfPPNRecs Then
      fpcmbType.Text = "PERSONAL ONLY"
      fpcmbType.AddItem "NO REAL"
      fpcmbType.AddItem "PERSONAL ONLY"
      fptxtRealYr.Text = "NA"
      IncReal = False
    End If
  End If

  Close RPNHandle
  Close PPNHandle
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  fptxtCurrRYear.Text = CStr(TaxMasterRec.RTaxYear)
  fptxtCurrPYear.Text = CStr(TaxMasterRec.PTaxYear)
  
End Sub
