VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Begin VB.Form frmPrnAcctHist 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Account History"
   ClientHeight    =   8640
   ClientLeft      =   30
   ClientTop       =   540
   ClientWidth     =   12195
   Icon            =   "frmPrnAcctHist.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   12195
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboRptType 
      Height          =   405
      Left            =   5130
      TabIndex        =   5
      Top             =   5910
      Width           =   1920
      _Version        =   196608
      _ExtentX        =   3387
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
      BackColor       =   16777215
      ForeColor       =   0
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
      ThreeDInsideHighlightColor=   14737632
      ThreeDInsideShadowColor=   0
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   14737632
      ThreeDOutsideShadowColor=   8421504
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   0
      BorderWidth     =   1
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   12632256
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   4210752
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
      GrayAreaColor   =   12632256
      ListLeftOffset  =   0
      ComboGap        =   -2
      MaxEditLen      =   150
      VirtualPageSize =   0
      VirtualPagesAhead=   0
      ExtendCol       =   0
      ColumnLevels    =   1
      ListGrayAreaColor=   12632256
      GroupHeaderHeight=   -1
      GroupHeaderShow =   0   'False
      AllowGrpResize  =   0
      AllowGrpDragDrop=   0
      MergeAdjustView =   0   'False
      ColumnHeaderShow=   0   'False
      ColumnHeaderHeight=   -1
      GrpsFrozen      =   0
      BorderGrayAreaColor=   12632256
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
      ColDesigner     =   "frmPrnAcctHist.frx":08CA
   End
   Begin LpLib.fpCombo fpcboAcct1 
      Height          =   405
      Left            =   5130
      TabIndex        =   0
      Top             =   2805
      Width           =   4740
      _Version        =   196608
      _ExtentX        =   8361
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
      ForeColor       =   0
      Text            =   ""
      Columns         =   4
      Sorted          =   0
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   3
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
      ThreeDInsideHighlightColor=   14737632
      ThreeDInsideShadowColor=   0
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   14737632
      ThreeDOutsideShadowColor=   8421504
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   0
      BorderWidth     =   1
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   12632256
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   4210752
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
      ColumnEdit      =   -1
      ColumnBound     =   -1
      Style           =   2
      MaxDrop         =   8
      ListWidth       =   -1
      EditHeight      =   -1
      GrayAreaColor   =   12632256
      ListLeftOffset  =   0
      ComboGap        =   -2
      MaxEditLen      =   150
      VirtualPageSize =   0
      VirtualPagesAhead=   0
      ExtendCol       =   0
      ColumnLevels    =   1
      ListGrayAreaColor=   12632256
      GroupHeaderHeight=   -1
      GroupHeaderShow =   0   'False
      AllowGrpResize  =   0
      AllowGrpDragDrop=   0
      MergeAdjustView =   0   'False
      ColumnHeaderShow=   0   'False
      ColumnHeaderHeight=   -1
      GrpsFrozen      =   0
      BorderGrayAreaColor=   12632256
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
      ColDesigner     =   "frmPrnAcctHist.frx":0C44
   End
   Begin LpLib.fpCombo fpcboAcct2 
      Height          =   405
      Left            =   5130
      TabIndex        =   1
      Top             =   3420
      Width           =   4740
      _Version        =   196608
      _ExtentX        =   8361
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
      ForeColor       =   0
      Text            =   ""
      Columns         =   4
      Sorted          =   0
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   3
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
      ThreeDInsideHighlightColor=   14737632
      ThreeDInsideShadowColor=   0
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   14737632
      ThreeDOutsideShadowColor=   8421504
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   0
      BorderWidth     =   1
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   12632256
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   4210752
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
      ColumnEdit      =   -1
      ColumnBound     =   -1
      Style           =   2
      MaxDrop         =   8
      ListWidth       =   -1
      EditHeight      =   -1
      GrayAreaColor   =   12632256
      ListLeftOffset  =   0
      ComboGap        =   -2
      MaxEditLen      =   150
      VirtualPageSize =   0
      VirtualPagesAhead=   0
      ExtendCol       =   0
      ColumnLevels    =   1
      ListGrayAreaColor=   12632256
      GroupHeaderHeight=   -1
      GroupHeaderShow =   0   'False
      AllowGrpResize  =   0
      AllowGrpDragDrop=   0
      MergeAdjustView =   0   'False
      ColumnHeaderShow=   0   'False
      ColumnHeaderHeight=   -1
      GrpsFrozen      =   0
      BorderGrayAreaColor=   12632256
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
      ColDesigner     =   "frmPrnAcctHist.frx":105B
   End
   Begin LpLib.fpCombo txtRepType 
      Height          =   405
      Left            =   5130
      TabIndex        =   4
      Top             =   5280
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
      BackColor       =   16777215
      ForeColor       =   0
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
      SearchMethod    =   2
      VirtualMode     =   0   'False
      VRowCount       =   0
      DataSync        =   3
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   14737632
      ThreeDInsideShadowColor=   0
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   14737632
      ThreeDOutsideShadowColor=   8421504
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   0
      BorderWidth     =   1
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   12632256
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   4210752
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
      GrayAreaColor   =   12632256
      ListLeftOffset  =   0
      ComboGap        =   -2
      MaxEditLen      =   150
      VirtualPageSize =   0
      VirtualPagesAhead=   0
      ExtendCol       =   0
      ColumnLevels    =   1
      ListGrayAreaColor=   12632256
      GroupHeaderHeight=   -1
      GroupHeaderShow =   0   'False
      AllowGrpResize  =   0
      AllowGrpDragDrop=   0
      MergeAdjustView =   0   'False
      ColumnHeaderShow=   0   'False
      ColumnHeaderHeight=   -1
      GrpsFrozen      =   0
      BorderGrayAreaColor=   12632256
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
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmPrnAcctHist.frx":1472
   End
   Begin VB.CheckBox chkLDesc 
      BackColor       =   &H00D0D0D0&
      Caption         =   "Include Additional Description"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   5136
      TabIndex        =   16
      Top             =   6456
      Width           =   3180
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00D0D0D0&
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
      Left            =   10032
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7296
      Width           =   1332
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00D0D0D0&
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
      Left            =   8256
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7296
      Width           =   1332
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   8
      Top             =   8280
      Width           =   12192
      _ExtentX        =   21511
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7117
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7117
            TextSave        =   "4:03 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7117
            TextSave        =   "7/15/2009"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin EditLib.fpDateTime txtDate1 
      Height          =   372
      Left            =   5136
      TabIndex        =   2
      Top             =   4044
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
      BackColor       =   16777215
      ForeColor       =   0
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   12632256
      ThreeDInsideShadowColor=   0
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   14737632
      ThreeDOutsideShadowColor=   8421504
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   0
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
      ThreeDTextHighlightColor=   14737632
      ThreeDTextShadowColor=   8421504
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
      InvalidColor    =   16777215
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   16777215
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
      BorderGrayAreaColor=   12632256
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   12632256
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   8421504
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
      ButtonColor     =   12632256
      AutoMenu        =   0   'False
      StartMonth      =   4
      ButtonAlign     =   0
      BoundDataType   =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpDateTime txtDate2 
      Height          =   372
      Left            =   5136
      TabIndex        =   3
      Top             =   4668
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
      BackColor       =   16777215
      ForeColor       =   0
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   14737632
      ThreeDInsideShadowColor=   8421504
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   14737632
      ThreeDOutsideShadowColor=   8421504
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   0
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
      ThreeDTextHighlightColor=   14737632
      ThreeDTextShadowColor=   8421504
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
      InvalidColor    =   16777215
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   16777215
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
      BorderGrayAreaColor=   12632256
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   12632256
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   8421504
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
      ButtonColor     =   12632256
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
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   2664
      TabIndex        =   15
      Top             =   5928
      Width           =   2388
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ending Account:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   3168
      TabIndex        =   14
      Top             =   3492
      Width           =   1812
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
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   3312
      TabIndex        =   13
      Top             =   4104
      Width           =   1668
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Run Balance or Source:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   2400
      TabIndex        =   12
      Top             =   5328
      Width           =   2580
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
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   3408
      TabIndex        =   11
      Top             =   4716
      Width           =   1572
   End
   Begin VB.Image Image1 
      Height          =   345
      Left            =   2490
      Picture         =   "frmPrnAcctHist.frx":17B5
      Top             =   2970
      Width           =   360
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3216
      Top             =   1176
      Width           =   5772
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Print Account History Report"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3984
      TabIndex        =   10
      Top             =   1416
      Width           =   4332
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   4524
      Left            =   1920
      Top             =   2544
      Width           =   8316
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Starting Account:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   2976
      TabIndex        =   9
      Top             =   2880
      Width           =   2004
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D0D0D0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00D0D0D0&
      FillColor       =   &H00D0D0D0&
      Height          =   972
      Left            =   3216
      Top             =   1056
      Width           =   5772
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Optio&ns"
      Begin VB.Menu mnuPrnScn 
         Caption         =   "Prin&t Screen"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmPrnAcctHist"
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
Dim FirstFund As String, LastFund As String
Dim ActiveYear As Integer

Private Sub cmdExit_Click()
  frmGLReportsMenu.Show
  Unload frmPrnAcctHist
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
Private Function ValidAccts()
  If fpcboAcct1.ListIndex <> -1 And fpcboAcct2.ListIndex <> -1 Then
    fpcboAcct1.Col = 1
    fpcboAcct2.Col = 1
    If fpcboAcct1.ColText > fpcboAcct2.ColText Then
      MsgBox "Invalid Account Selection, The Starting Account Should Be Less or Equal to Ending Account.", vbOKOnly, "Invalid Selection"
      ValidAccts = False
    Else
      ValidAccts = True
    End If
  Else
    MsgBox "You Must Select An Account, Retry", vbOKOnly, "Invalid Selection"
    ValidAccts = False
  End If
End Function

Private Sub cmdPrint_Click()
  If ValidAccts = True Then
    If ValidDate = True Then
      If fpcboRptType.ListIndex = 0 Then
        rptopt = 1
      ElseIf fpcboRptType.ListIndex = 1 Then
        rptopt = 2
      End If
      If rptopt = 1 Then
        PrintAcctHist
      ElseIf rptopt = 2 Then
        PrintAcctHist2
      End If
    End If
  End If
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

Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub


Private Sub txtDate1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    txtDate2.SetFocus
  End If
End Sub


Private Sub txtDate2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    txtRepType.SetFocus
  End If
End Sub

Private Sub txtRepType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    txtRepType.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    txtRepType.ListIndex = -1
    txtRepType.Action = ActionClearSearchBuffer
  End If
  If txtRepType.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpcboRptType.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        txtDate2.SetFocus
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
        txtRepType.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub
Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  Me.HelpContextID = hlpAccountHistory
  StatusBar1.Panels.Item(1).Text = GLUserName
  FillAcctstwo fpcboAcct1, fpcboAcct2
  txtDate1.Text = Format(Now, "mm/dd/yyyy")
  txtDate2.Text = txtDate1.Text
  txtRepType.AddItem "Running Balance"
  txtRepType.AddItem "Source"
  fpcboRptType.InsertRow = "Graphics"
  fpcboRptType.InsertRow = "Text"
  fpcboRptType.ListIndex = 0
End Sub
Private Sub fpcboAcct1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboAcct1.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcboAcct1.ListIndex = -1
    fpcboAcct1.Action = ActionClearSearchBuffer
  End If
  If fpcboAcct1.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpcboAcct2.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        cmdPrint.SetFocus
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcboAcct2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboAcct2.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcboAcct2.ListIndex = -1
    fpcboAcct2.Action = ActionClearSearchBuffer
  End If
  If fpcboAcct2.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      txtDate1.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpcboAcct1.SetFocus
        KeyCode = 0
      End If
    End If
  End If

End Sub
Private Sub Form_Resize()
'  If Me.Visible Then
'  If Me.cmdExit.Enabled = False Then
'    Me.WindowState = vbMaximized
'  Else
  
    Temp_Class.ResizeControls Me
    DoEvents
 ' End If
End Sub

Private Sub PrintAcctHist()
  Dim MaxLines As Integer, LookFor As String, CrLF As String, T As String
  Dim Linecnt As Integer, PRNFile As Integer, FundCnt As Integer, Newrp As String
  Dim ReportFile As String, ToPrint As String, SumLine As String
  Dim FF As String, Header As String, StartAcct As String, EndAcct As String
  Dim PRNFileNum As Integer, cnt As Integer, Howmany As Integer
  Dim FundCode As String, DivLine As String, DivLine2 As String, HAcct As String
  Dim CommaFmt As String, TotalFmt As String, FundNumber As String
  Dim TotDr As Double, TotCr As Double, TranCashTot As Double, CalcBal As Double
  ReDim FundList(1) As String
  Dim OpenDate As String, IGuess As Integer, GrTotDr As Double, GrTotCr As Double
  Dim FundDr As Double, FundCr As Double, FundRecNum As Integer, NewAcct As Boolean
  Dim Found As Boolean, FundOutofBal As Boolean, Fund As Integer
  Dim FundIdxFileNum As Integer, NumFunds As Integer, EndDate As Integer, StartDate As Integer
  Dim AcctIdxFileNum As Integer, NumGLAccts As Integer, FundName As String
  Dim AcctFile As Integer, NumAccts As Integer, RecNo As Integer
  Dim transfile As Integer, NumTrans As Long, NextTr As Long, AcctNum As String
  Dim Debit As String, Credit As String, Diff As Double, PYFundBal As Double
  Dim DrFwd As Double, CrFwd As Double, TotAcctDr As Double, TotAcctCr As Double
  Dim BalFwd As Double, AcctNumber As String, FYBeg As Integer, FwdFlag As Boolean
  Dim OutOfOrder As Boolean, cntT As Integer, AcctRunBal As Double, Trn As Integer
  Dim TmpSort As TrSortType
  Dim Opt As String, AcctBal As Double, BgtCol As Integer, BudgetAmt As Double
  Dim Var As Double, VarCol As Integer, VarText As String, HollyFlag As Boolean
  Dim Pitch12 As String, PageNum As Integer, NumAcctTrans As Long, RunBalFmt As String
'  If InStr(UCase$(GLUserName), "HOLLY SPR") > 0 Then
'    HollyFlag = True
'    Pitch12$ = Chr$(27) + Chr$(38) + Chr$(107) + Chr$(52) + Chr$(83)
'  End If
  'Pitch12$ = Chr$(27) + "&k4S"
  GetFYDates FY1BegDate, FY1EndDate, FY2BegDate, FY2EndDate
    
  If txtRepType.Text = "Running Balance" Then
    Opt$ = "R"
  Else
    Opt$ = "S"
  End If
'  '=====================================================
  'Start Report Processing

  CommaFmt$ = "###,###,###.##"  'format takes 13 chars
  TotalFmt$ = "#,###,###,###.##" 'format takes 14 chars
  RunBalFmt$ = "##########.##"
  SumLine$ = String$(16, "-")   'column summary line
  DivLine$ = String$(77, "-")   'dashed line
  DivLine2$ = String$(77, "=")  'Double Line
  FF$ = Chr$(12)
  MaxLines = 53
  Linecnt = 0
  TotDr# = 0
  TotCr# = 0
  StartDate = DateDiff("d", "12/31/1979", txtDate1)
  EndDate = DateDiff("d", "12/31/1979", txtDate2)
  fpcboAcct1.Col = 1
  StartAcct = fpcboAcct1.ColText
  fpcboAcct2.Col = 1
  EndAcct = fpcboAcct2.ColText
  T$ = "Account History " + txtDate1 + "  " + "Thru  " + txtDate2 + "                           Page:"
  ToPrint$ = Space$(80)
  FrmShowPctComp.Label1 = "Printing Account History Report"
  FrmShowPctComp.Show , Me
  DoEvents
  DeActivateControls frmPrnAcctHist, True
  If EndDate >= FY2BegDate Then
    ActiveYear = 2
    OpenDate$ = Format(DateAdd("d", FY2BegDate, "12-31-1979"), "mm/dd/yy")
    FYBeg = FY2BegDate
  Else
    ActiveYear = 1
    OpenDate$ = Format(DateAdd("d", FY1BegDate, "12-31-1979"), "mm/dd/yy")
    FYBeg = FY1BegDate
  End If
  Newrp = "ActHst"
  GetRPTName Newrp
  ReportFile$ = Newrp

  PRNFile = FreeFile
  Open ReportFile$ For Output As #PRNFile
'  If HollyFlag Then
'    Print #PRNFile, Pitch12$;
'  End If

  OpenAcctIdx AcctIdxFileNum, NumGLAccts
  OpenAcctFile AcctFile
  NumAccts = LOF(AcctFile) / Len(GLAcct)
  OpenTransFile transfile, NumTrans&

  If NumTrans& = 0 Then
    Close
    ActivateControls frmPrnAcctHist, True
    Unload FrmShowPctComp
    MsgBox "NO Transactions To Report.", vbOKOnly, "No Trans"
    GoTo CancelExit
  End If

  '--trying to set TrSort array to max needed
'  If NumTrans& > 10000 Then
'    IGuess = Int(NumTrans& * 0.5)
'  Else
'    IGuess = NumTrans&
'  End If

  'ReDim TrSort(1 To IGuess) As TrSortType
  '|-----this works--------------------------|
  'ReDim Trsort(1 To 20000) As TrSortType         '|
  '|-----------------------------------------|
  ReDim Trsort(1 To 1) As TrSortType         '|

  GrTotDr# = 0
  GrTotCr# = 0

  NewAcct = True
  'GoSub PrintAHPageHeader

  For cnt = 1 To NumGLAccts

    DrFwd# = 0
    CrFwd# = 0
    TotAcctDr# = 0
    TotAcctCr# = 0

    Get AcctIdxFileNum, cnt, GLAcctidx
    Get AcctFile, GLAcctidx.RecNum, GLAcct

    AcctNum$ = QPTrim$(GLAcct.Num)

    If AcctNum$ >= StartAcct$ And AcctNum$ <= EndAcct$ Then

      BalFwd# = Round#(GLAcct.BegBal)             'get the beginning balance
      NextTr& = GLAcct.FrstTran   'get the first trans for this acct
      AcctNumber$ = AcctNum$ + " - " + QPTrim$(GLAcct.Title)
      ToPrint$ = "H~Account~" & AcctNumber$ & "~~~~"
      Print #PRNFile, ToPrint$
      'HAcct$ = "H~Account~" & AcctNumber$
      Linecnt = Linecnt + 1

      'IF LineCnt > MaxLines THEN
      '  PRINT #PrnFile, FF$
      '  GOSUB PrintAHPageHeader
      'END IF

      '--Deal with the bal fwd field
      If BalFwd# <> 0 Then
        '--Convert the balance fwd to a debit or a credit amt
        Select Case GLAcct.Typ
        Case "A", "E"
          If BalFwd# >= 0 Then
            Debit$ = Abs(BalFwd#)
            Credit$ = ""
            DrFwd# = BalFwd#
          Else
            Credit$ = Abs(BalFwd#)
            Debit$ = ""
            CrFwd# = Abs(BalFwd#)
          End If
        Case "L", "R"
          If BalFwd# >= 0 Then
            Credit$ = Abs(BalFwd#)
            Debit$ = ""
            CrFwd# = BalFwd#
          Else
            Debit$ = Abs(BalFwd#)
            Credit$ = ""
            DrFwd# = Abs(BalFwd#)
          End If
        End Select

        '--If we're running a history with a Report Begin Date that is prior
        '  to the FY Beginning date then print as opening balance.
        If StartDate <= FYBeg Then
          ToPrint$ = ""
'          Mid$(ToPrint$, 1) = OpenDate$
'          Mid$(ToPrint$, 12) = "Balance Foward"
'          Mid$(ToPrint$, 34) = Debit$
'          Mid$(ToPrint$, 51) = Credit$
          ToPrint$ = "H~" & OpenDate$ & "~" & "Balance Forward~~" & Debit$ & "~" & Credit$ & "~"
          Print #PRNFile, ToPrint$
          Linecnt = Linecnt + 1
          If Linecnt > MaxLines Then
            Linecnt = 0
            NewAcct = False
            'GoSub PrintAHPageHeader
          End If
        Else
          '--otherwise set a flag so we'll know we have a bal fwd
          '--to print
          FwdFlag = -1
        End If
      End If

      '--search thru acct transactions
      Do Until NextTr& = 0
        Get transfile, NextTr&, GLTrans
        '--see if transaction is in date range
        If GLTrans.TRDATE >= StartDate And GLTrans.TRDATE <= EndDate Then
          '--within range - assign to array for sorting
          NumAcctTrans = NumAcctTrans + 1
          ReDim Preserve Trsort(1 To NumAcctTrans) As TrSortType         '|
          Trsort(NumAcctTrans).TRDATE = GLTrans.TRDATE
          Trsort(NumAcctTrans).Record = NextTr&
        Else
          '--check the transaction to see if we need to carry it in
          '  the balance fwd
          If GLTrans.TRDATE <= StartDate Then
            '--Check here for dual year to bring asset & liab acct
            '  balances fwd but not rev & exp acct balances
            Select Case ActiveYear
            Case 1              '--Year 1
              DrFwd# = DrFwd# + GLTrans.DrAmt
              CrFwd# = CrFwd# + GLTrans.CrAmt
              FwdFlag = -1
            Case 2              '--Year 2
              Select Case GLAcct.Typ
              '--Get Asset & Liab acct balances as bal fwd
              Case "A", "L"
                DrFwd# = DrFwd# + GLTrans.DrAmt
                CrFwd# = CrFwd# + GLTrans.CrAmt
                FwdFlag = -1
              '--Get Rev & Liab acct balances
              Case "R", "E"
                If GLTrans.TRDATE >= FY2BegDate Then               'STOP
                  DrFwd# = DrFwd# + GLTrans.DrAmt
                  CrFwd# = CrFwd# + GLTrans.CrAmt
                  FwdFlag = -1
                End If
              End Select
            End Select
          End If                'IF Trans.TrDate < BegDate THEN
        End If
        NextTr& = GLTrans.NextTran                'Get the next transaction
      Loop

      If FwdFlag Then
        '--Print the balance forward
        ToPrint$ = ""
        ToPrint$ = "H~****~" & "Balance Forward~"
        Select Case GLAcct.Typ
        Case "A", "E"
          BalFwd# = Round(DrFwd# - CrFwd#)
          If BalFwd# >= 0 Then
            Debit$ = BalFwd#
            Credit$ = ""
          Else
            Credit$ = Abs(BalFwd#)
            Debit$ = ""
          End If
        Case "L", "R"
          BalFwd# = Round(CrFwd# - DrFwd#)
          If BalFwd# >= 0 Then
            Credit$ = BalFwd#
            Debit$ = ""
          Else
            Debit$ = Abs(BalFwd#)
            Credit$ = ""
          End If
        End Select

        ToPrint$ = ToPrint$ & "~" & Debit$ & "~" & Credit$ & "~"
        

        Print #PRNFile, ToPrint$
        Linecnt = Linecnt + 1
        If Linecnt > MaxLines Then
          NewAcct = False
          Linecnt = 0
          'Print #PRNFile, FF$
          'GoSub PrintAHPageHeader
        End If

      End If

    If NumAcctTrans > 0 Or BalFwd# <> 0 Then
      SortT Trsort(), NumAcctTrans
''''''' Start here with sortt replacement
'''      Do
'''        OutOfOrder = False          'assume it's sorted
'''        For cntT = 1 To NumAcctTrans - 1
'''          If Trsort(cntT).TRDATE > Trsort(cntT + 1).TRDATE Then
'''            LSet TmpSort = Trsort(cntT)
'''            LSet Trsort(cntT) = Trsort(cntT + 1)
'''            LSet Trsort(cntT + 1) = TmpSort
'''            OutOfOrder = True       'we're not done yet
'''          End If
'''        Next
'''      Loop While OutOfOrder
''''''' The SortT replaced with above statements
''''''' SortT TrSort(1), NumAcctTrans, 0, 6, 0, -1
        AcctRunBal# = BalFwd#
        For Trn = 1 To NumAcctTrans
          Get transfile, Trsort(Trn).Record, GLTrans
          ToPrint$ = ""
          ToPrint$ = "D~" & Format(DateAdd("d", GLTrans.TRDATE, "12-31-1979"), "mm/dd/yy")
          If chkLDesc.Value = 1 Then
            If Len(QPTrim$(GLTrans.LDesc)) > 0 Then
              ToPrint$ = ToPrint$ & "~" & Mid$(QPTrim$(GLTrans.Desc), 1, 20) + " " + (QPTrim$(GLTrans.LDesc))
            Else
              ToPrint$ = ToPrint$ & "~" & Mid$(QPTrim$(GLTrans.Desc), 1, 20)
            End If
          Else
            ToPrint$ = ToPrint$ & "~" & Mid$(QPTrim$(GLTrans.Desc), 1, 20)
          End If
          ToPrint$ = ToPrint$ & "~" & Mid$(QPTrim$(GLTrans.Ref), 1, 8)
         ' If GLTrans.DrAmt <> 0 Then
            ToPrint$ = ToPrint$ & "~" & Using$(RunBalFmt$, Str$(GLTrans.DrAmt))
         ' End If
         ' If GLTrans.CrAmt <> 0 Then
            ToPrint$ = ToPrint$ & "~" & Using$(RunBalFmt$, Str$(GLTrans.CrAmt))
         ' End If
          Select Case Opt$
          Case "R"
          Select Case GLAcct.Typ
            Case "A", "E"
              AcctRunBal# = Round(AcctRunBal# + GLTrans.DrAmt - GLTrans.CrAmt)
            Case "L", "R"
              AcctRunBal# = Round(AcctRunBal# + GLTrans.CrAmt - GLTrans.DrAmt)
            End Select
            ToPrint$ = ToPrint$ & "~" & Using$(RunBalFmt$, Str$(AcctRunBal#))
          Case "S"
             ToPrint$ = ToPrint$ & "~" & Left$(GLTrans.Src, 6)
          End Select

          Print #PRNFile, ToPrint$
          Linecnt = Linecnt + 1
          If Linecnt > MaxLines Then
            NewAcct = False
            Linecnt = 0
            'Print #PRNFile, FF$
            'GoSub PrintAHPageHeader
          End If
          TotAcctDr# = Round(TotAcctDr# + GLTrans.DrAmt)
          TotAcctCr# = Round(TotAcctCr# + GLTrans.CrAmt)
          GrTotDr# = Round(GrTotDr# + GLTrans.DrAmt)
          GrTotCr# = Round(GrTotCr# + GLTrans.CrAmt)
        Next

        '--Print summary lines
'        LSet ToPrint$ = ""
'        Mid$(ToPrint$, 34) = SumLine$
'        Mid$(ToPrint$, 51) = SumLine$
'        Print #PRNFile, ToPrint$
        Linecnt = Linecnt + 1
        'IF LineCnt > MaxLines THEN
       ' ToPrint$ = "L~~~~~~"
       '   Print #PRNFile, ToPrint$
        '  GOSUB PrintAHPageHeader
        'END IF

        '--Print transaction totals
        If NumAcctTrans > 0 Then
          ToPrint$ = "T~~~~.................~..................~"
          Print #PRNFile, ToPrint$
          ToPrint$ = "T~*****~Transaction Totals~~" & Using$(RunBalFmt$, Str$(TotAcctDr#)) & "~" & Using$(RunBalFmt$, Str$(TotAcctCr#)) & "~"
          Print #PRNFile, ToPrint$
          Linecnt = Linecnt + 1
        End If

        '--Print ending balance
        'Print #PRNFile,
        
          ToPrint$ = "T~~~~.................~..................~"
          Print #PRNFile, ToPrint$

        ToPrint$ = "T~*****~Ending Account Balance~"
        Select Case GLAcct.Typ
        Case "A", "E"
          AcctBal# = Round(BalFwd# + TotAcctDr# - TotAcctCr#)
          If AcctBal# >= 0 Then
            Debit$ = AcctBal#
            Credit$ = ""
          Else
            Credit$ = Abs(AcctBal#)
            Debit$ = ""
          End If
        Case "L", "R"
          AcctBal# = Round(BalFwd# + TotAcctCr# - TotAcctDr#)
          If AcctBal# >= 0 Then
            Credit$ = AcctBal#
            Debit$ = ""
          Else
            Debit$ = Abs(AcctBal#)
            Credit$ = ""
          End If
        End Select
        ToPrint$ = ToPrint$ & "~" & Debit$ & "~" & Credit$ & "~"
        Print #PRNFile, ToPrint$
        Linecnt = Linecnt + 1

        If GLAcct.Typ = "R" Or GLAcct.Typ = "E" Then
          If GLAcct.Typ = "R" Then
            BgtCol = 51
          Else
            BgtCol = 34
          End If

          Select Case ActiveYear
          Case 1
            BudgetAmt# = GLAcct.Bgt
          Case 2
            BudgetAmt# = GLAcct.NYApp
          End Select
          ToPrint$ = "T~*****~Budget~"
          ToPrint$ = ToPrint$ & Abs(BudgetAmt#) & "~~~"
          Print #PRNFile, ToPrint$
          Linecnt = Linecnt + 1
          Var# = Round(BudgetAmt# - AcctBal#)
          Select Case GLAcct.Typ
          Case "R"
            If Var# > 0 Then
              VarCol = 51
              VarText$ = "Uncollected Balance"
            Else
              VarCol = 51
              VarText$ = "Revenues Exceeding Budget"
            End If
          Case "E"
            If Var# > 0 Then
              VarCol = 34
              VarText$ = "Appropriation Remaining"
            Else
              VarCol = 34
              VarText$ = "Over Spent"
            End If
          End Select
          ToPrint$ = "T~*****~" & VarText$
          ToPrint$ = ToPrint$ & "~~" & Using$(RunBalFmt$, Str$(Abs(Var#))) & "~~"
          Print #PRNFile, ToPrint$
          Linecnt = Linecnt + 1
        End If

        ToPrint$ = "TE~------------------~------------------------------------------~~~~"
        Print #PRNFile, ToPrint$
        Linecnt = Linecnt + 1

        '--Don't break up summary section
        If Linecnt > MaxLines Then
          NewAcct = True
          Linecnt = 0
          'Print #PRNFile, FF$
          'GoSub PrintAHPageHeader
        End If
      Else
        ToPrint$ = ""
        ToPrint$ = "TE~-- No Activity --~-------------------------------------------~~~~"
        Print #PRNFile, ToPrint$
        Linecnt = Linecnt + 1
        'ToPrint$ = "T~~~~~~End"
        'Print #PRNFile, ToPrint$
        'Linecnt = Linecnt + 1
        If Linecnt > MaxLines Then
          NewAcct = True
          Linecnt = 0
          'Print #PRNFile, FF$
          'GoSub PrintAHPageHeader
        End If
      End If
    End If      'Account is not of this fund

    NumAcctTrans = 0            'reset for next account
       FrmShowPctComp.ShowPctComp cnt, NumGLAccts
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ActivateControls frmPrnAcctHist, True
      Unload FrmShowPctComp
      GoTo CancelExit
    End If
  Next
 ' ToPrint$ = ""
  ToPrint$ = "T~=================~====================================================================~====================~======================~===================~"
  Print #PRNFile, ToPrint$
  ToPrint$ = "T~*****~Grand Total Debits~~" & Using$(RunBalFmt$, Str$(GrTotDr#)) & "~~"
 
  Print #PRNFile, ToPrint$
 ' ToPrint$ = ""
  ToPrint$ = "T~*****~Grand Total Credits~~~" & Using$(RunBalFmt$, Str$(GrTotCr#)) & "~"
  
  Print #PRNFile, ToPrint$

  If ActiveYear = 2 Then
    ToPrint$ = "TE~Note:Unaudited balances~prior year not closed~~~~"
    Print #PRNFile, ToPrint$
  Else
    ToPrint$ = "TE~---------------------~-------------------------------------------~~~~"
    Print #PRNFile, ToPrint$
  End If

  'Print #PRNFile, FF$
  Close
  Load frmLoadingRpt
  Erase Trsort
  ARptAcctHist.GetName ReportFile$
  GoSub PrintAHPageHeader
   
  ActivateControls frmPrnAcctHist, True
  ARptAcctHist.startrpt

  'End Report Processing
  '===========================================================================
 '  ViewPrint ReportFile$, "Account History Report"
 ' KillFile ReportFile$
   

  Exit Sub

PrintAHPageHeader:
  'PageNum = PageNum + 1
  ARptAcctHist.txtTown = GLUserName$
  ARptAcctHist.txtDate = Now
  ARptAcctHist.Label1.Caption = "Date"
  ARptAcctHist.Label2.Caption = "Description"
  ARptAcctHist.Label3.Caption = "Reference"
  ARptAcctHist.Label4.Caption = "Debit"
  ARptAcctHist.Label5.Caption = "Credit"
  ARptAcctHist.Label7.Caption = T$
  
  Select Case Opt$
  Case "S", " "
    ARptAcctHist.Label6.Caption = "Source"
    ARptAcctHist.Label14.Caption = "Source"
  Case "R"
    ARptAcctHist.Label6.Caption = "Run Balance"
    ARptAcctHist.BalSrc.Alignment = ddTXRight
    ARptAcctHist.Label14.Caption = "Run Balance"
  End Select
  'Print #PRNFile, String$(80, "-")
  If NewAcct Then
    Linecnt = 5
  Else
    'ARptAcctHist.Label7.Caption = AcctNumber$ + " continued."
    Linecnt = 5
    NewAcct = True
  End If
  Return

CancelExit:
Exit Sub

End Sub
Private Sub PrintAcctHist2()
  Dim MaxLines As Integer, LookFor As String, CrLF As String, T As String
  Dim Linecnt As Integer, PRNFile As Integer, FundCnt As Integer, Newrp As String
  Dim ReportFile As String, ToPrint As String, SumLine As String
  Dim FF As String, Header As String, StartAcct As String, EndAcct As String
  Dim PRNFileNum As Integer, cnt As Integer, Howmany As Integer
  Dim FundCode As String, DivLine As String, DivLine2 As String
  Dim CommaFmt As String, TotalFmt As String, FundNumber As String
  Dim TotDr As Double, TotCr As Double, TranCashTot As Double, CalcBal As Double
  ReDim FundList(1) As String
  Dim OpenDate As String, IGuess As Integer, GrTotDr As Double, GrTotCr As Double
  Dim FundDr As Double, FundCr As Double, FundRecNum As Integer, NewAcct As Boolean
  Dim Found As Boolean, FundOutofBal As Boolean, Fund As Integer
  Dim FundIdxFileNum As Integer, NumFunds As Integer, EndDate As Integer, StartDate As Integer
  Dim AcctIdxFileNum As Integer, NumGLAccts As Integer, FundName As String
  Dim AcctFile As Integer, NumAccts As Integer, RecNo As Integer
  Dim transfile As Integer, NumTrans As Long, NextTr As Long, AcctNum As String
  Dim Debit As String, Credit As String, Diff As Double, PYFundBal As Double
  Dim DrFwd As Double, CrFwd As Double, TotAcctDr As Double, TotAcctCr As Double
  Dim BalFwd As Double, AcctNumber As String, FYBeg As Integer, FwdFlag As Boolean
  Dim OutOfOrder As Boolean, cntT As Integer, AcctRunBal As Double, Trn As Integer
  Dim TmpSort As TrSortType
  Dim Opt As String, AcctBal As Double, BgtCol As Integer, BudgetAmt As Double
  Dim Var As Double, VarCol As Integer, VarText As String, HollyFlag As Boolean
  Dim Pitch12 As String, PageNum As Integer, NumAcctTrans As Long, RunBalFmt As String
  If InStr(UCase$(GLUserName), "HOLLY SPR") > 0 Then
    HollyFlag = True
    Pitch12$ = Chr$(27) + Chr$(38) + Chr$(107) + Chr$(52) + Chr$(83)
  End If
  'Pitch12$ = Chr$(27) + "&k4S"
  GetFYDates FY1BegDate, FY1EndDate, FY2BegDate, FY2EndDate
    
  If txtRepType.Text = "Running Balance" Then
    Opt$ = "R"
  Else
    Opt$ = "S"
  End If
'  '=====================================================
  'Start Report Processing

  CommaFmt$ = "###,###,###.##"  'format takes 13 chars
  TotalFmt$ = "#,###,###,###.##" 'format takes 14 chars
  RunBalFmt$ = "##########.##"
  SumLine$ = String$(16, "-")   'column summary line
  DivLine$ = String$(77, "-")   'dashed line
  DivLine2$ = String$(77, "=")  'Double Line
  FF$ = Chr$(12)
  MaxLines = 53
  Linecnt = 0
  TotDr# = 0
  TotCr# = 0
  StartDate = DateDiff("d", "12/31/1979", txtDate1)
  EndDate = DateDiff("d", "12/31/1979", txtDate2)
  fpcboAcct1.Col = 1
  StartAcct = fpcboAcct1.ColText
  fpcboAcct2.Col = 1
  EndAcct = fpcboAcct2.ColText
  T$ = "Account History " + txtDate1 + "  " + "Thru  " + txtDate2 + "                           Page:"
  ToPrint$ = Space$(80)
  FrmShowPctComp.Label1 = "Printing Account History Report"
  FrmShowPctComp.Show , Me
  DoEvents
  DeActivateControls frmPrnAcctHist, True
  If EndDate >= FY2BegDate Then
    ActiveYear = 2
    OpenDate$ = Format(DateAdd("d", FY2BegDate, "12-31-1979"), "mm/dd/yy")
    FYBeg = FY2BegDate
  Else
    ActiveYear = 1
    OpenDate$ = Format(DateAdd("d", FY1BegDate, "12-31-1979"), "mm/dd/yy")
    FYBeg = FY1BegDate
  End If
  Newrp = "ActHst"
  GetRPTName Newrp
  ReportFile$ = Newrp

  PRNFile = FreeFile
  Open ReportFile$ For Output As #PRNFile
  If HollyFlag Then
    Print #PRNFile, Pitch12$;
  End If

  OpenAcctIdx AcctIdxFileNum, NumGLAccts
  OpenAcctFile AcctFile
  NumAccts = LOF(AcctFile) / Len(GLAcct)
  OpenTransFile transfile, NumTrans&

  If NumTrans& = 0 Then
    Close
    ActivateControls frmPrnAcctHist, True
    Unload FrmShowPctComp
    MsgBox "NO Transactions To Report.", vbOKOnly, "No Trans"
    GoTo CancelExit
  End If

  '--trying to set TrSort array to max needed
'  If NumTrans& > 10000 Then
'    IGuess = Int(NumTrans& * 0.5)
'  Else
'    IGuess = NumTrans&
'  End If

  'ReDim TrSort(1 To IGuess) As TrSortType
  '|-----this works--------------------------|
  'ReDim Trsort(1 To 20000) As TrSortType         '|
  '|-----------------------------------------|
  ReDim Trsort(1 To 1) As TrSortType         '|

  GrTotDr# = 0
  GrTotCr# = 0

  NewAcct = True
  GoSub PrintAHPageHeader

  For cnt = 1 To NumGLAccts

    DrFwd# = 0
    CrFwd# = 0
    TotAcctDr# = 0
    TotAcctCr# = 0

    Get AcctIdxFileNum, cnt, GLAcctidx
    Get AcctFile, GLAcctidx.RecNum, GLAcct

    AcctNum$ = QPTrim$(GLAcct.Num)

    If AcctNum$ >= StartAcct$ And AcctNum$ <= EndAcct$ Then

      BalFwd# = Round#(GLAcct.BegBal)             'get the beginning balance
      NextTr& = GLAcct.FrstTran   'get the first trans for this acct
      AcctNumber$ = "Account " + AcctNum$ + " - " + QPTrim$(GLAcct.Title)
      LSet ToPrint$ = AcctNumber$
      Print #PRNFile, ToPrint$
      Linecnt = Linecnt + 1

      'IF LineCnt > MaxLines THEN
      '  PRINT #PrnFile, FF$
      '  GOSUB PrintAHPageHeader
      'END IF

      '--Deal with the bal fwd field
      If BalFwd# <> 0 Then
        '--Convert the balance fwd to a debit or a credit amt
        Select Case GLAcct.Typ
        Case "A", "E"
          If BalFwd# >= 0 Then
            Debit$ = Using$(RunBalFmt$, Str$(Abs(BalFwd#)))
            Credit$ = ""
            DrFwd# = BalFwd#
          Else
            Credit$ = Using$(RunBalFmt$, Str$(Abs(BalFwd#)))
            Debit$ = ""
            CrFwd# = Abs(BalFwd#)
          End If
        Case "L", "R"
          If BalFwd# >= 0 Then
            Credit$ = Using$(RunBalFmt$, Str$(Abs(BalFwd#)))
            Debit$ = ""
            CrFwd# = BalFwd#
          Else
            Debit$ = Using$(RunBalFmt$, Str$(Abs(BalFwd#)))
            Credit$ = ""
            DrFwd# = Abs(BalFwd#)
          End If
        End Select

        '--If we're running a history with a Report Begin Date that is prior
        '  to the FY Beginning date then print as opening balance.
        If StartDate <= FYBeg Then
          LSet ToPrint$ = ""
          Mid$(ToPrint$, 1) = OpenDate$
          Mid$(ToPrint$, 12) = "Balance Foward"
          Mid$(ToPrint$, 34) = Debit$
          Mid$(ToPrint$, 51) = Credit$
          Print #PRNFile, ToPrint$
          Linecnt = Linecnt + 1
          If Linecnt > MaxLines Then
            Print #PRNFile, FF$
            NewAcct = False
            GoSub PrintAHPageHeader
          End If
        Else
          '--otherwise set a flag so we'll know we have a bal fwd
          '--to print
          FwdFlag = -1
        End If
      End If

      '--search thru acct transactions
      Do Until NextTr& = 0
        Get transfile, NextTr&, GLTrans
        '--see if transaction is in date range
        If GLTrans.TRDATE >= StartDate And GLTrans.TRDATE <= EndDate Then
          '--within range - assign to array for sorting
          NumAcctTrans = NumAcctTrans + 1
          ReDim Preserve Trsort(1 To NumAcctTrans) As TrSortType         '|
          Trsort(NumAcctTrans).TRDATE = GLTrans.TRDATE
          Trsort(NumAcctTrans).Record = NextTr&
        Else
          '--check the transaction to see if we need to carry it in
          '  the balance fwd
          If GLTrans.TRDATE <= StartDate Then
            '--Check here for dual year to bring asset & liab acct
            '  balances fwd but not rev & exp acct balances
            Select Case ActiveYear
            Case 1              '--Year 1
              DrFwd# = DrFwd# + GLTrans.DrAmt
              CrFwd# = CrFwd# + GLTrans.CrAmt
              FwdFlag = -1
            Case 2              '--Year 2
              Select Case GLAcct.Typ
              '--Get Asset & Liab acct balances as bal fwd
              Case "A", "L"
                DrFwd# = DrFwd# + GLTrans.DrAmt
                CrFwd# = CrFwd# + GLTrans.CrAmt
                FwdFlag = -1
              '--Get Rev & Liab acct balances
              Case "R", "E"
                If GLTrans.TRDATE >= FY2BegDate Then               'STOP
                  DrFwd# = DrFwd# + GLTrans.DrAmt
                  CrFwd# = CrFwd# + GLTrans.CrAmt
                  FwdFlag = -1
                End If
              End Select
            End Select
          End If                'IF Trans.TrDate < BegDate THEN
        End If
        NextTr& = GLTrans.NextTran                'Get the next transaction
      Loop

      If FwdFlag Then
        '--Print the balance forward
        LSet ToPrint$ = ""
        Mid$(ToPrint$, 1) = "Balance Forward"
        Select Case GLAcct.Typ
        Case "A", "E"
          BalFwd# = DrFwd# - CrFwd#
          If BalFwd# >= 0 Then
            Debit$ = Using$(TotalFmt$, Str$(BalFwd#))
            Credit$ = ""
          Else
            Credit$ = Using$(TotalFmt$, Str$(Abs(BalFwd#)))
            Debit$ = ""
          End If
        Case "L", "R"
          BalFwd# = CrFwd# - DrFwd#
          If BalFwd# >= 0 Then
            Credit$ = Using$(TotalFmt$, Str$(BalFwd#))
            Debit$ = ""
          Else
            Debit$ = Using$(TotalFmt$, Str$(Abs(BalFwd#)))
            Credit$ = ""
          End If
        End Select

        Mid$(ToPrint$, 34) = Debit$
        Mid$(ToPrint$, 51) = Credit$

        Print #PRNFile, ToPrint$
        Linecnt = Linecnt + 1
        If Linecnt > MaxLines Then
          NewAcct = False
          Print #PRNFile, FF$
          GoSub PrintAHPageHeader
        End If

      End If

    If NumAcctTrans > 0 Or BalFwd# <> 0 Then
      SortT Trsort(), NumAcctTrans
''''''' Start here with sortt replacement
'''      Do
'''        OutOfOrder = False          'assume it's sorted
'''        For cntT = 1 To NumAcctTrans - 1
'''          If Trsort(cntT).TRDATE > Trsort(cntT + 1).TRDATE Then
'''            LSet TmpSort = Trsort(cntT)
'''            LSet Trsort(cntT) = Trsort(cntT + 1)
'''            LSet Trsort(cntT + 1) = TmpSort
'''            OutOfOrder = True       'we're not done yet
'''          End If
'''        Next
'''      Loop While OutOfOrder
''''''' The SortT replaced with above statements
''''''' SortT TrSort(1), NumAcctTrans, 0, 6, 0, -1
        AcctRunBal# = BalFwd#
        For Trn = 1 To NumAcctTrans
          Get transfile, Trsort(Trn).Record, GLTrans
          LSet ToPrint$ = ""
          Mid$(ToPrint$, 1) = Format(DateAdd("d", GLTrans.TRDATE, "12-31-1979"), "mm/dd/yy")
          Mid$(ToPrint$, 10, 15) = QPTrim$(GLTrans.Desc)
          Mid$(ToPrint$, 26, 8) = QPTrim$(GLTrans.Ref)
          If GLTrans.DrAmt <> 0 Then
            Mid$(ToPrint$, 37) = Using$(RunBalFmt$, Str$(GLTrans.DrAmt))
          End If
          If GLTrans.CrAmt <> 0 Then
            Mid$(ToPrint$, 54) = Using$(RunBalFmt$, Str$(GLTrans.CrAmt))
          End If
          Select Case Opt$
          Case "R"
          Select Case GLAcct.Typ
            Case "A", "E"
              AcctRunBal# = AcctRunBal# + GLTrans.DrAmt - GLTrans.CrAmt
            Case "L", "R"
              AcctRunBal# = AcctRunBal# + GLTrans.CrAmt - GLTrans.DrAmt
            End Select
            Mid$(ToPrint$, 68) = Using$(RunBalFmt$, Str$(AcctRunBal#))
          Case "S"
            Mid$(ToPrint$, 73) = Left$(GLTrans.Src, 6)
          End Select

          Print #PRNFile, ToPrint$
          If chkLDesc.Value = 1 Then
            Print #PRNFile, Tab(10); QPTrim$(GLTrans.LDesc)
            Linecnt = Linecnt + 1
          End If
          Linecnt = Linecnt + 1
          If Linecnt > MaxLines Then
            NewAcct = False
            Print #PRNFile, FF$
            GoSub PrintAHPageHeader
          End If
          TotAcctDr# = TotAcctDr# + GLTrans.DrAmt
          TotAcctCr# = TotAcctCr# + GLTrans.CrAmt
          GrTotDr# = GrTotDr# + GLTrans.DrAmt
          GrTotCr# = GrTotCr# + GLTrans.CrAmt
        Next

        '--Print summary lines
        LSet ToPrint$ = ""
        Mid$(ToPrint$, 34) = SumLine$
        Mid$(ToPrint$, 51) = SumLine$
        Print #PRNFile, ToPrint$
        Linecnt = Linecnt + 1
        'IF LineCnt > MaxLines THEN
        '  NewAcct = False
        '  PRINT #PrnFile, FF$
        '  GOSUB PrintAHPageHeader
        'END IF

        '--Print transaction totals
        If NumAcctTrans > 0 Then
          LSet ToPrint$ = "Transaction Totals"
          Mid$(ToPrint$, 34) = Using$(TotalFmt$, Str$(TotAcctDr#))
          Mid$(ToPrint$, 51) = Using$(TotalFmt$, Str$(TotAcctCr#))
          Print #PRNFile, ToPrint$
          Linecnt = Linecnt + 1
        End If

        '--Print ending balance
        Print #PRNFile,
        LSet ToPrint$ = "Ending Balance"
        Select Case GLAcct.Typ
        Case "A", "E"
          AcctBal# = BalFwd# + TotAcctDr# - TotAcctCr#
          If AcctBal# >= 0 Then
            Debit$ = Using$(TotalFmt$, Str$(AcctBal#))
            Credit$ = ""
          Else
            Credit$ = Using$(TotalFmt$, Str$(Abs(AcctBal#)))
            Debit$ = ""
          End If
        Case "L", "R"
          AcctBal# = BalFwd# + TotAcctCr# - TotAcctDr#
          If AcctBal# >= 0 Then
            Credit$ = Using$(TotalFmt$, Str$(AcctBal#))
            Debit$ = ""
          Else
            Debit$ = Using$(TotalFmt$, Str$(Abs(AcctBal#)))
            Credit$ = ""
          End If
        End Select
        Mid$(ToPrint$, 34) = Debit$
        Mid$(ToPrint$, 51) = Credit$
        Print #PRNFile, ToPrint$
        Linecnt = Linecnt + 1

        If GLAcct.Typ = "R" Or GLAcct.Typ = "E" Then
          If GLAcct.Typ = "R" Then
            BgtCol = 51
          Else
            BgtCol = 34
          End If

          Select Case ActiveYear
          Case 1
            BudgetAmt# = GLAcct.Bgt
          Case 2
            BudgetAmt# = GLAcct.NYApp
          End Select
          LSet ToPrint$ = "Budget"
          Mid$(ToPrint$, BgtCol) = Using$(TotalFmt$, Str$(Abs(BudgetAmt#)))
          Print #PRNFile, ToPrint$
          Linecnt = Linecnt + 1
          Var# = BudgetAmt# - AcctBal#
          Select Case GLAcct.Typ
          Case "R"
            If Var# > 0 Then
              VarCol = 51
              VarText$ = "Uncollected Balance"
            Else
              VarCol = 51
              VarText$ = "Revenues Exceeding Budget"
            End If
          Case "E"
            If Var# > 0 Then
              VarCol = 34
              VarText$ = "Appropriation Remaining"
            Else
              VarCol = 34
              VarText$ = "Over Spent"
            End If
          End Select
          LSet ToPrint$ = VarText$
          Mid$(ToPrint$, VarCol) = Using$(TotalFmt$, Str$(Abs(Var#)))
          Print #PRNFile, ToPrint$
          Linecnt = Linecnt + 1
        End If

        LSet ToPrint$ = String$(80, "=")
        Print #PRNFile, ToPrint$
        Linecnt = Linecnt + 1

        '--Don't break up summary section
        If Linecnt > MaxLines Then
          NewAcct = True
          Print #PRNFile, FF$
          GoSub PrintAHPageHeader
        End If
      Else
        LSet ToPrint$ = ""
        Mid$(ToPrint$, 5) = " -- No Activity --"
        Print #PRNFile, ToPrint$
        Linecnt = Linecnt + 1
        LSet ToPrint$ = String$(80, "=")
        Print #PRNFile, ToPrint$
        Linecnt = Linecnt + 1
        If Linecnt > MaxLines Then
          NewAcct = True
          Print #PRNFile, FF$
          GoSub PrintAHPageHeader
        End If
      End If
    End If      'Account is not of this fund

    NumAcctTrans = 0            'reset for next account
       FrmShowPctComp.ShowPctComp cnt, NumGLAccts
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ActivateControls frmPrnAcctHist, True
      Unload FrmShowPctComp
      GoTo CancelExit
    End If
  Next
   ActivateControls frmPrnAcctHist, True
  Load frmLoadingRpt
  LSet ToPrint$ = "Grand Total Debits"
  Mid$(ToPrint$, 25) = Using$(TotalFmt$, Str$(GrTotDr#))
  Print #PRNFile, ToPrint$

  LSet ToPrint$ = "Grand Total Credits"
  Mid$(ToPrint$, 25) = Using$(TotalFmt$, Str$(GrTotCr#))
  Print #PRNFile, ToPrint$

  If ActiveYear = 2 Then
    LSet ToPrint$ = "Note: Unaudited balances - prior year has not been closed"
    Print #PRNFile, ToPrint$
  End If

  Print #PRNFile, FF$
  Close

  Erase Trsort

  'End Report Processing
  '===========================================================================
   ViewPrint ReportFile$, "Account History Report"
  KillFile ReportFile$

  Exit Sub

PrintAHPageHeader:
  PageNum = PageNum + 1
  Print #PRNFile, GLUserName; Tab(43); "Run Date: " + Date$
  Print #PRNFile, T$; PageNum
  Print #PRNFile,
  Select Case Opt$
  Case "S", " "
    Print #PRNFile, "Date       Description        Reference       Debit        Credit       Source"
  Case "R"
    Print #PRNFile, "Date       Description        Reference       Debit        Credit    Run Balance"
  End Select
  Print #PRNFile, String$(80, "-")
  If NewAcct Then
    Linecnt = 5
  Else
    Print #PRNFile, AcctNumber$ + " continued."
    Linecnt = 5
    NewAcct = True
  End If
  Return

CancelExit:
Exit Sub

End Sub



