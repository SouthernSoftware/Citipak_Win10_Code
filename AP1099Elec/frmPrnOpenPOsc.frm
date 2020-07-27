VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPrnClosedPOs 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Purchase Order List"
   ClientHeight    =   8844
   ClientLeft      =   36
   ClientTop       =   540
   ClientWidth     =   12192
   Icon            =   "frmPrnOpenPOsc.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8844
   ScaleWidth      =   12192
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboRptType 
      Height          =   384
      Left            =   5880
      TabIndex        =   5
      Top             =   5880
      Width           =   1908
      _Version        =   196608
      _ExtentX        =   3365
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
      ColDesigner     =   "frmPrnOpenPOsc.frx":08CA
   End
   Begin LpLib.fpCombo fpcboEncumber 
      Height          =   384
      Left            =   5880
      TabIndex        =   3
      Top             =   4707
      Width           =   996
      _Version        =   196608
      _ExtentX        =   1757
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
      ColumnSearch    =   1
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
      ColDesigner     =   "frmPrnOpenPOsc.frx":0CF0
   End
   Begin LpLib.fpCombo fpcboDepartment 
      Height          =   384
      Left            =   5880
      TabIndex        =   4
      Top             =   5292
      Width           =   2364
      _Version        =   196608
      _ExtentX        =   4170
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
      Columns         =   3
      Sorted          =   0
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   1
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
      ColDesigner     =   "frmPrnOpenPOsc.frx":10DF
   End
   Begin LpLib.fpCombo fpcboSort 
      Height          =   384
      Left            =   5880
      TabIndex        =   2
      Top             =   4122
      Width           =   2340
      _Version        =   196608
      _ExtentX        =   4128
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
      ColumnSearch    =   1
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
      ColDesigner     =   "frmPrnOpenPOsc.frx":1552
   End
   Begin VB.CommandButton cmdOk 
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
      Height          =   468
      Left            =   6420
      TabIndex        =   6
      Top             =   7224
      Width           =   1236
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
      Height          =   468
      Left            =   8256
      TabIndex        =   7
      Top             =   7224
      Width           =   1236
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   8
      Top             =   8484
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
            TextSave        =   "1:22 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7133
            TextSave        =   "6/20/2005"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin EditLib.fpDateTime fpDate1 
      Height          =   372
      Left            =   5880
      TabIndex        =   0
      Top             =   2976
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
   Begin EditLib.fpDateTime fpDate2 
      Height          =   372
      Left            =   5880
      TabIndex        =   1
      Top             =   3549
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
   Begin VB.Label Label6 
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
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   3936
      TabIndex        =   15
      Top             =   3036
      Width           =   1668
   End
   Begin VB.Label Label5 
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
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   4032
      TabIndex        =   14
      Top             =   3614
      Width           =   1572
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Select Report Type:"
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
      Left            =   3216
      TabIndex        =   13
      Top             =   5928
      Width           =   2388
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Sort By:"
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
      Left            =   4440
      TabIndex        =   12
      Top             =   4192
      Width           =   1164
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Show Encumbrances:"
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
      Left            =   2856
      TabIndex        =   11
      Top             =   4770
      Width           =   2748
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Department:"
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
      Left            =   4248
      TabIndex        =   10
      Top             =   5348
      Width           =   1356
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Closed Purchase Orders"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3684
      TabIndex        =   9
      Top             =   1488
      Width           =   4836
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000009&
      Height          =   852
      Left            =   2580
      Top             =   1248
      Width           =   7020
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   4092
      Left            =   2322
      Top             =   2544
      Width           =   7548
   End
   Begin VB.Shape Shape6 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   996
      Left            =   2592
      Top             =   1128
      Width           =   7020
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
Attribute VB_Name = "frmPrnClosedPOs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim GLSetup As GLSetupRecType
Dim Acct    As GLAcctRecType
Dim GLFundIdx As GLFundIndexType
Dim AcctIdx As GLAcctIndexType
Dim Vendor As VendorRecType
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer

Private Sub cmdExit_Click()
  frmAPReportsMenu.Show
  Unload Me
End Sub

Private Sub cmdOk_Click()
 If ValidDate = True Then
  DeActivateControls frmPrnClosedPOs, True
  If fpcboRptType.ListIndex = 0 Then
    rptopt = 1
  ElseIf fpcboRptType.ListIndex = 1 Then
    rptopt = 2
  End If
  If rptopt = 1 Then
    PrintClosedPos
  ElseIf rptopt = 2 Then
    PrintClosedPos2
  End If
 End If
End Sub
Private Sub fpcboRptType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboRptType.ListDown = True
  End If
  If fpcboRptType.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      cmdOk.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpcboDepartment.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = True Then
      If MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        MainLog "Close AP"
        ClearInUse PWcnt
      End If
    Else
      Cancel = True
    End If
  End If
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
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%P"
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  StatusBar1.Panels.Item(1).Text = GLUserName
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Me.HelpContextID = hlpOpenPO
  fpDate1.Text = Format(Now, "mm/dd/yyyy")
  fpDate2.Text = fpDate1.Text
  DeptList fpcboDepartment
  fpcboDepartment.ListIndex = 0
  fpcboSort.AddItem "PO Number"
  fpcboSort.AddItem "Vendor"
  fpcboSort.ListIndex = 0
  fpcboEncumber.AddItem "Yes"
  fpcboEncumber.AddItem "No"
  fpcboEncumber.ListIndex = 0
  fpcboRptType.InsertRow = "Graphics"
  fpcboRptType.InsertRow = "Text"
  fpcboRptType.ListIndex = 0
End Sub
Private Function ValidDate()
  Dim TempDate1 As Integer, TempDate2 As Integer
  If CheckValDate(fpDate1) = False And CheckValDate(fpDate2) = False Then
    MsgBox "Date Is Not Valid. Please Correct.", vbOKOnly, "Invalid Date"
    ValidDate = False
  Else
    TempDate1 = DateDiff("d", "12/31/1979", fpDate1)
    TempDate2 = DateDiff("d", "12/31/1979", fpDate2)
    If TempDate1 > TempDate2 Then
      ValidDate = False
      MsgBox "The Starting And Ending Dates Must Be In Chronological Order Or Equal", vbOKOnly, "Invalid Date"
    Else
      ValidDate = True
    End If
  End If
End Function

Private Sub Form_Resize()
'  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
'  End If
End Sub
Private Sub PrintClosedPos()
  Dim ColTitle As String, Header As String, CommaFmt As String
  Dim RptFund As Integer, User As String, RptFundName As String
  Dim MaxPO As Integer, NumFunds As Integer, APLRecLen As Integer
  Dim APLedgerFile As Integer, NumTran As Long, APDRecLen As Integer
  Dim APDistFile As Integer, NumDistRecs As Long, VendorFile As Integer
  Dim NumVRecs As Integer, RptFile As Integer, RptFileName As String
  Dim OhShoot As Boolean, cnt As Long, NumPo As Integer, Dept As String
  Dim Linecnt As Integer, Rec As Long, RunTotal As Double, Newrp As String
  Dim ToPrint As String, Encumb As Double, UnEncumb As Double, PODet As Boolean
  Dim TotEnc As Double, TotUn As Double, NextDist As Long, DistAmt As Double
  Dim Found As Boolean, Fund As Integer, FundNum As String, FCnt As Integer
  Dim ToPrintV As String, ToPrintD As String, ColTitle2 As String
  Dim BegDate As Integer, EndDate As Integer
  If fpcboDepartment.ListIndex = -1 Then
    MsgBox "You Must Select All or A Valid Dept Number", vbOKOnly, "Department Required"
    fpcboDepartment.SetFocus
  End If
  If fpcboEncumber.ListIndex = 1 Then
    PODet = False
  Else
    PODet = True
  End If
  BegDate = DateDiff("d", "12/31/1979", fpDate1)
  EndDate = DateDiff("d", "12/31/1979", fpDate2)
  FrmShowPctComp.Label1 = "Creating Closed Purchase Orders Report"
  FrmShowPctComp.Show , Me
  DoEvents
'  If fpcboSort.ListIndex = 1 Then
'    ColTitle$ = "Vendor Code,Name"
'    ColTitle2$ = "PO Number"
'    Header$ = "Closed Purchase Orders by Vendor"
'  Else
    ColTitle$ = "PO Number"
    ColTitle2$ = "Vendor Code,Name"
    Header$ = "Closed Purchase Orders"
'  End If
  CommaFmt$ = "###,###,###.##"
  User$ = QPTrim$(GLUserName$)
  
 ' MaxPO = 400
 ' ReDim POList(1 To MaxPO) As GLAcctIndexType   '--borrowing this type
  ReDim POList(1 To 1) As ChkSortType   'use this so will have long rec num
  '--Get a list of active funds
  ReDim FundList(1) As String
  GetFundList FundList$(), NumFunds
  ReDim FundGrdTot#(1 To NumFunds)

  Dim APLedger As APLedger81RecType
  APLRecLen = Len(APLedger)
  OpenAPLedgerFile APLedgerFile, NumTran&, APLRecLen

  Dim APDist As APDistRecType
  APDRecLen = Len(APDist)
  OpenAPDistFile APDistFile, NumDistRecs&, APDRecLen

  OpenVendorFile VendorFile, NumVRecs

  RptFile = FreeFile
  Newrp = "OPO"
  GetRPTName Newrp
  RptFileName$ = Newrp
  Open RptFileName$ For Output As RptFile
  RptFund = FreeFile
  RptFundName$ = "s" & Newrp
  Open RptFundName$ For Output As RptFund

  If fpcboDepartment.ListIndex <> -1 Then
    fpcboDepartment.col = 1
    Dept$ = QPTrim$(fpcboDepartment.ColText)
  End If
  'OhShoot = False
  For cnt = 1 To NumTran&
   ' Pct$ = Str$(Int((cnt / NumTran&) * 100))
   ' QPrintRC "Reading..." + Pct$ + "%", 25, 2, -1
    FrmShowPctComp.ShowPctComp cnt, NumTran&
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ActivateControls frmPrnClosedPOs, True
      Unload FrmShowPctComp
      GoTo CancelExit
    End If

    Get APLedgerFile, cnt, APLedger
    If APLedger.TRCode = -4 Then
    ' used below line to test for problem with subreport not printing for Maxton on 4/30/03
    ' was because of split report footer - set to keep together and printed fine.
     ' If Val(APLedger.PONum) >= 3000 And Val(APLedger.PONum) <= 3008 Then
      If fpcboDepartment.ListIndex = 0 Or APLedger.DeptNumb = Val(Dept$) Then
       If APLedger.TRDATE >= BegDate And APLedger.TRDATE <= EndDate Then
        NumPo = NumPo + 1
        ReDim Preserve POList(1 To NumPo) As ChkSortType   'use this so will have long rec num
'        If NumPo = MaxPO Then
'          OhShoot = True
'          Exit For
'        Else
          POList(NumPo).Record = cnt
          If fpcboSort.ListIndex = 0 Then
            
            POList(NumPo).CHKinfo = Left$(APLedger.PONum, 14)
          Else
            POList(NumPo).CHKinfo = QPTrim$(APLedger.VendorCode)
          End If
'        End If
      End If
     End If
    End If
  Next

'  If OhShoot = True Then
'    Close
'    MsgBox "Error: Available elements exceed needs. Unable to run report.", vbOKOnly, "Error"
'    Exit Sub
'  End If

'  If NumPO > 0 Then
'
'    If ShowEnc Then
'      GoSub ClearEnc
'    End If
'
'  Else
  If NumPo > 0 Then
    GoSub SortPO
    GoSub PrintPOList
  Else
    MsgBox "No Closed Purchase Orders", vbOKOnly
    Close
    ActivateControls frmPrnClosedPOs, True
    Exit Sub
  End If

  Close
  Load frmLoadingRpt
  If PODet = True Then
    If fpcboSort.ListIndex = 1 Then
    'when duplicate po nums the vendor didn't print so rem this line
      'ARptOpnPOs1.GroupHeader2.DataField = "OPO2"
    End If
    ARptOpnPOs1.Label8.Caption = ColTitle$
    ARptOpnPOs1.Label11.Caption = ColTitle2$
    ARptOpnPOs1.Label1.Caption = Header$
    ARptOpnPOs1.Label2.Caption = Dept$
    ARptOpnPOs1.Label14.Caption = "Total Closed PO's"
    ARptOpnPOs1.Label13.Caption = "Canceled"
    ARptOpnPOs1.Label18.Caption = "Canceled"
    ARptOpnPOs1.GetName RptFileName$, RptFundName$
    ActivateControls frmPrnClosedPOs, True
    ARptOpnPOs1.txtTown.Caption = GLUserName$
    ARptOpnPOs1.txtDate.Caption = Now
    ARptOpnPOs1.startrpt
  Else
    ARptOpnPOs2.Label8.Caption = ColTitle$
    ARptOpnPOs2.Label11.Caption = ColTitle2$
    ARptOpnPOs2.Label1.Caption = Header$
    ARptOpnPOs2.Label2.Caption = Dept$
    ARptOpnPOs2.Label14.Caption = "Total Closed PO's"
    ARptOpnPOs2.GetName RptFileName$
    ActivateControls frmPrnClosedPOs, True
    ARptOpnPOs2.txtTown.Caption = GLUserName$
    ARptOpnPOs2.txtDate.Caption = Now
    ARptOpnPOs2.startrpt

  End If

  Exit Sub


SortPO:
  ReDim Preserve POList(1 To NumPo) As ChkSortType
  Dim lngCurLow As Long, lngCurHigh As Long
  lngCurLow = LBound(POList)
  lngCurHigh = UBound(POList)
  QPOSort POList(), lngCurLow, lngCurHigh

 
  Return


'OpenPOPageHdr:
'  Page = Page + 1
'  Print #RptFile, Tab(40 - (Int(Len(User$) / 2))); User$
'  Print #RptFile, Tab(40 - (Int(Len(Header$) / 2))); Header$
'  Print #RptFile,
'  Print #RptFile, "Dept Number: "; Dept$
'  Print #RptFile, "Report Date: "; Date$; Tab(67); "Page #"; Page
'  Print #RptFile, ColTitle$
'  Print #RptFile, String$(80, "=")
'  Linecnt = 4
'  Return


PrintPOList:
  For cnt = 1 To NumPo
   ' Pct$ = Str$(Int((cnt / NumPo) * 100))
   ' QPrintRC "Writing..." + Pct$ + "%", 25, 2, -1

    Rec = POList(cnt).Record
    Get APLedgerFile, Rec, APLedger
    Get VendorFile, APLedger.VRecNum, Vendor
    'IF APLedger.Amt < 100000 THEN
    '  'STOP
    '  APLedger.Amt = 0
    'END IF
    RunTotal# = RunTotal# + APLedger.Amt
'    Vendor.DelFlag = 0
'    PUT VendorFile, APLedger.VRecNum, Vendor
    Dim dodo As String
    
    ToPrintV$ = ""
'    If fpcboSort.ListIndex = 1 Then
'      ToPrintV$ = QPTrim(Vendor.vnum) + "," + QPTrim$(Vendor.VNAME) + "~" + Left$(APLedger.PONum, 15) + "~"
'    Else
      ToPrintV$ = Left$(APLedger.PONum, 15) + "~" + QPTrim(Vendor.vnum) + "," + QPTrim$(Vendor.VNAME) + "~"
     'dodo$ = QPTrim$(APLedger.PONum)
     'Stop
     'ToPrintV$ = dodo$ + "~" + QPTrim$(Vendor.VNAME) + "~"
     'ToPrintV$ = "1" + "~" + QPTrim$(Vendor.VNAME) + "~"
'    End If
    ToPrintV$ = ToPrintV$ + Format(DateAdd("d", (APLedger.TRDATE), "12-31-1979"), "mm/dd/yyyy") + "~"
    ToPrintV$ = ToPrintV$ + Using(CommaFmt$, Str$(APLedger.Amt)) + "~"
'    If Linecnt > MaxLines Then
'      Print #RptFile, FF$
'      GoSub OpenPOPageHdr
'    End If

    If fpcboEncumber.ListIndex = 0 Then
      '--Now print the distribution

      NextDist& = APLedger.FrstDist
      DistAmt# = 0

      If NextDist& > 0 Then  'and nextdist& < Then
'        Print #RptFile,
'        Print #RptFile, Tab(50); "Encumbered Accounts:"
'        LineCnt = LineCnt + 2
        Encumb = 0
        UnEncumb = 0

        Do
          Get APDistFile, NextDist&, APDist

          DistAmt# = DistAmt# + APDist.DistAmt
          NextDist& = APDist.NextDist

          ToPrintD$ = Space$(80)
          If APDist.DistStat <> "L" Then
            ToPrintD$ = "Canceled" + "~" + QPTrim$(APDist.DistAcctNum) + "~" + Using(CommaFmt$, Str$(APDist.DistAmt))
            Encumb = Encumb + APDist.DistAmt
          Else
            ToPrintD$ = "Liquidated" + "~" + QPTrim$(APDist.DistAcctNum) + "~" + Using(CommaFmt$, Str$(APDist.DistAmt))
            UnEncumb = UnEncumb + APDist.DistAmt
          End If
          Linecnt = Linecnt + 1
'          If Linecnt > MaxLines Then
'            Print #RptFile, FF$
'            GoSub OpenPOPageHdr
'          End If
          'ToPrint$ = ToPrintV$ + ToPrintD$
          ToPrint$ = ToPrintV$ + ToPrintD$
          Print #RptFile, ToPrint$
          '--summarize by fund
          Found = False
          For Fund = 1 To NumFunds
            FundNum$ = Left$(APDist.DistAcctNum, GLFundLen)
            If FundNum$ = FundList$(Fund) Then
              Found = True
              FundGrdTot#(Fund) = Round#(FundGrdTot#(Fund) + APDist.DistAmt)
              Exit For
            End If
          Next
        Loop Until NextDist& = 0
          TotEnc = TotEnc + Encumb
          TotUn = TotUn + UnEncumb
        'Print #RptFile, Tab(50); "Total:"; Tab(64); Using(CommaFmt$, Str$(DistAmt#))
       ' If fpcboEncumber.ListIndex = 0 Then
         ' Print #RptFile, Tab(46); "Liquidated:"; Tab(64); Using(CommaFmt$, Str$(UnEncumb))
         ' Print #RptFile, Tab(46); "Encumbered:"; Tab(64); Using(CommaFmt$, Str$(Encumb))
       ' End If
        '--showing encumbrances
      End If
    Else
      ToPrint$ = ToPrintV$ + "~~"
      Print #RptFile, ToPrint$
    End If
  Next

 ' Print #RptFile, "Total Open PO's: " + Using(CommaFmt$, Str$(RunTotal#))
  If fpcboEncumber.ListIndex = 0 Then
    'Print #RptFile, "Tot Encumbered:  " + Using(CommaFmt$, Str$(TotEnc))
    'Print #RptFile, "Tot Liquidated:  " + Using(CommaFmt$, Str$(TotUn))
    For FCnt = 1 To NumFunds
      If FundGrdTot#(FCnt) > 0 Then
        Print #RptFund, FundList$(FCnt) + "~" + Using(CommaFmt$, Str$(FundGrdTot#(FCnt)))
       'used line below to test error
       ' Print #RptFund, Str$(FCnt) + "~" + Str$(20)
      End If
    Next
  End If
  Return

''UpdateGLAcct:   'Reseting the Encumbered Amt
''  Amt# = APDist.DistAmt
''  DistAcctRec = FindAcct(APDist.DistAcctNum)
''  If DistAcctRec > 0 Then
''    OpenAcctFile AcctFileNum, NumGLAcctRecs
''    Get AcctFileNum, DistAcctRec, Acct
''    Acct.Encumb = Acct.Encumb + Amt#
''    Put AcctFileNum, DistAcctRec, Acct
''    Close AcctFileNum
''  End If
''  Return
''
''
''ClearEnc:
''  OpenAcctFile AcctFileNum, NumGLAcctRecs
''  For Cnt1! = 1 To NumGLAcctRecs
''    Get AcctFileNum, Cnt1!, Acct
''    Acct.Encumb = 0
''    Put AcctFileNum, Cnt1!, Acct
''  Next Cnt1!
''  Close AcctFileNum
''  Return
CancelExit:
  Exit Sub
End Sub
Private Sub PrintClosedPos2()
  Dim ColTitle As String, Header As String, CommaFmt As String
  Dim MaxLines As Integer, User As String, Page As Integer, FF As String
  Dim MaxPO As Integer, NumFunds As Integer, APLRecLen As Integer
  Dim APLedgerFile As Integer, NumTran As Long, APDRecLen As Integer
  Dim APDistFile As Integer, NumDistRecs As Long, VendorFile As Integer
  Dim NumVRecs As Integer, RptFile As Integer, RptFileName As String
  Dim OhShoot As Boolean, cnt As Long, NumPo As Integer, Dept As String
  Dim Linecnt As Integer, Rec As Long, RunTotal As Double, Newrp As String
  Dim ToPrint As String, Encumb As Double, UnEncumb As Double
  Dim TotEnc As Double, TotUn As Double, NextDist As Long, DistAmt As Double
  Dim Found As Boolean, Fund As Integer, FundNum As String, FCnt As Integer
  Dim BegDate As Integer, EndDate As Integer
  If fpcboDepartment.ListIndex = -1 Then
    MsgBox "You Must Select All or A Valid Dept Number", vbOKOnly, "Department Required"
    fpcboDepartment.SetFocus
  End If
  BegDate = DateDiff("d", "12/31/1979", fpDate1)
  EndDate = DateDiff("d", "12/31/1979", fpDate2)
  FrmShowPctComp.Label1 = "Creating Closed Purchase Orders Report"
  FrmShowPctComp.Show , Me
  DoEvents
'  If fpcboSort.ListIndex = 1 Then
'    ColTitle$ = " Vendor                           PO Num          Date             Amt"
'    Header$ = "Closed Purchase Orders by Vendor"
'  Else
    ColTitle$ = " PO Num           Vendor                          Date             Amt"
    Header$ = "Closed Purchase Orders"
 ' End If
  CommaFmt$ = "###,###,###.##"
  MaxLines = 55
  User$ = QPTrim$(GLUserName$)
  Page = 0
  FF$ = Chr$(12)
  'MaxPO = 400
  'ReDim POList(1 To MaxPO) As GLAcctIndexType   '--borrowing this type
  ReDim POList(1 To 1) As ChkSortType    'use this for long recnum
  '--Get a list of active funds
  ReDim FundList(1) As String
  GetFundList FundList$(), NumFunds
  ReDim FundGrdTot#(1 To NumFunds)

  Dim APLedger As APLedger81RecType
  APLRecLen = Len(APLedger)
  OpenAPLedgerFile APLedgerFile, NumTran&, APLRecLen

  Dim APDist As APDistRecType
  APDRecLen = Len(APDist)
  OpenAPDistFile APDistFile, NumDistRecs&, APDRecLen

  OpenVendorFile VendorFile, NumVRecs

  RptFile = FreeFile
  Newrp = "OPO"
  GetRPTName Newrp
  RptFileName$ = Newrp
  Open RptFileName$ For Output As RptFile

  GoSub OpenPOPageHdr
  If fpcboDepartment.ListIndex <> -1 Then
    fpcboDepartment.col = 1
    Dept$ = QPTrim$(fpcboDepartment.ColText)
  End If
  'OhShoot = False
  For cnt = 1 To NumTran&
   ' Pct$ = Str$(Int((cnt / NumTran&) * 100))
   ' QPrintRC "Reading..." + Pct$ + "%", 25, 2, -1
    FrmShowPctComp.ShowPctComp cnt, NumTran&
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ActivateControls frmPrnClosedPOs, True
      Unload FrmShowPctComp
      GoTo CancelExit
    End If

    Get APLedgerFile, cnt, APLedger
    If APLedger.TRCode = -4 Then
      If fpcboDepartment.ListIndex = 0 Or APLedger.DeptNumb = Val(Dept$) Then
       If APLedger.TRDATE >= BegDate And APLedger.TRDATE <= EndDate Then
        NumPo = NumPo + 1
        ReDim Preserve POList(1 To NumPo) As ChkSortType    'use this for long recnum
'        If NumPo = MaxPO Then
'          OhShoot = True
'          Exit For
'        Else
          POList(NumPo).Record = cnt
          If fpcboSort.ListIndex = 0 Then
            POList(NumPo).CHKinfo = Left$(APLedger.PONum, 14)
          Else
            POList(NumPo).CHKinfo = APLedger.VendorCode
          End If
        End If
      End If
    End If
  Next

'  If OhShoot = True Then
'    Close
'    MsgBox "Error: Available elements exceed needs. Unable to run report.", vbOKOnly, "Error"
'    Exit Sub
'  End If

'  If NumPO > 0 Then
'
'    If ShowEnc Then
'      GoSub ClearEnc
'    End If
'
'  Else
  If NumPo > 0 Then
    GoSub SortPO
    GoSub PrintPOList
  Else
    Print #RptFile, "No Open Purchase Orders"
  End If

  Close
  ActivateControls frmPrnClosedPOs, True
  ViewPrint RptFileName$, Header$
  KillFile RptFileName$

  Exit Sub


SortPO:
  ReDim Preserve POList(1 To NumPo) As ChkSortType
  Dim lngCurLow As Long, lngCurHigh As Long
  lngCurLow = LBound(POList)
  lngCurHigh = UBound(POList)
  QPOSort POList(), lngCurLow, lngCurHigh

 
  Return


OpenPOPageHdr:
  Page = Page + 1
  Print #RptFile, Tab(40 - (Int(Len(User$) / 2))); User$
  Print #RptFile, Tab(40 - (Int(Len(Header$) / 2))); Header$
  Print #RptFile,
  Print #RptFile, "Dept Number: "; Dept$
  Print #RptFile, "Report Date: "; Date$; Tab(67); "Page #"; Page
  Print #RptFile, ColTitle$
  Print #RptFile, String$(80, "=")
  Linecnt = 4
  Return


PrintPOList:
  For cnt = 1 To NumPo
   ' Pct$ = Str$(Int((cnt / NumPo) * 100))
   ' QPrintRC "Writing..." + Pct$ + "%", 25, 2, -1

    Rec = POList(cnt).Record
    Get APLedgerFile, Rec, APLedger
    Get VendorFile, APLedger.VRecNum, Vendor
    'IF APLedger.Amt < 100000 THEN
    '  'STOP
    '  APLedger.Amt = 0
    'END IF
    RunTotal# = RunTotal# + APLedger.Amt
'    Vendor.DelFlag = 0
'    PUT VendorFile, APLedger.VRecNum, Vendor

    ToPrint$ = Space$(80)
'    If fpcboSort.ListIndex = 1 Then
'      Mid$(ToPrint$, 2) = QPTrim(Vendor.vnum) + "," + QPTrim(Vendor.VNAME)      'APLedger.VendorCode
'      Mid$(ToPrint$, 35) = Left$(APLedger.PONum, 15)
'    Else
      Mid$(ToPrint$, 2) = Left$(APLedger.PONum, 15)
      'IF INSTR(APLedger.PONUM, "57097") > 0 THEN STOP
      'PUT VendorFile, APLedger.VRecNum, Vendor
      Mid$(ToPrint$, 19) = QPTrim(Vendor.vnum) + "," + QPTrim(Vendor.VNAME)         'APLedger.VendorCode
'    End If
    Mid$(ToPrint$, 51) = Format(DateAdd("d", (APLedger.TRDATE), "12-31-1979"), "mm/dd/yyyy")
    Mid$(ToPrint$, 64) = Using(CommaFmt$, Str$(APLedger.Amt))
    Print #RptFile, ToPrint$
    Linecnt = Linecnt + 1
    If Linecnt > MaxLines Then
      Print #RptFile, FF$
      GoSub OpenPOPageHdr
    End If

    If fpcboEncumber.ListIndex = 0 Then
      '--Now print the distribution

      NextDist& = APLedger.FrstDist
      DistAmt# = 0

      If NextDist& > 0 Then
'        Print #RptFile,
'        Print #RptFile, Tab(50); "Encumbered Accounts:"
'        LineCnt = LineCnt + 2
        Encumb = 0
        UnEncumb = 0

        Do
          Get APDistFile, NextDist&, APDist

          DistAmt# = DistAmt# + APDist.DistAmt
          NextDist& = APDist.NextDist

          ToPrint$ = Space$(80)
          If APDist.DistStat <> "L" Then
            Mid$(ToPrint$, 35) = "Canceled"
            Mid$(ToPrint$, 50) = APDist.DistAcctNum
            Mid$(ToPrint$, 64) = Using(CommaFmt$, Str$(APDist.DistAmt))
            Print #RptFile, ToPrint$
            Encumb = Encumb + APDist.DistAmt
          Else
            Mid$(ToPrint$, 35) = "Liquidated"
            Mid$(ToPrint$, 50) = APDist.DistAcctNum
            Mid$(ToPrint$, 64) = Using(CommaFmt$, Str$(APDist.DistAmt))
            Print #RptFile, ToPrint$
            UnEncumb = UnEncumb + APDist.DistAmt
          End If
          Linecnt = Linecnt + 1
          If Linecnt > MaxLines Then
            Print #RptFile, FF$
            GoSub OpenPOPageHdr
          End If

          '--summarize by fund
          Found = False
          For Fund = 1 To NumFunds
            FundNum$ = Left$(APDist.DistAcctNum, GLFundLen)
            If FundNum$ = FundList$(Fund) Then
              Found = True
              FundGrdTot#(Fund) = Round#(FundGrdTot#(Fund) + APDist.DistAmt)
              Exit For
            End If
          Next
        Loop Until NextDist& = 0
          TotEnc = TotEnc + Encumb
          TotUn = TotUn + UnEncumb

        Print #RptFile, Tab(64); "------------------"
        Print #RptFile, Tab(50); "Total:"; Tab(64); Using(CommaFmt$, Str$(DistAmt#))
        If fpcboEncumber.ListIndex = 0 Then
          Print #RptFile, Tab(64); "------------------"
          Print #RptFile, Tab(46); "Liquidated:"; Tab(64); Using(CommaFmt$, Str$(UnEncumb))
          Print #RptFile, Tab(46); "  Canceled:"; Tab(64); Using(CommaFmt$, Str$(Encumb))
        End If
        Linecnt = Linecnt + 4
        Print #RptFile, String$(80, "-")
        '--showing encumbrances
      End If
    End If
  Next

  Print #RptFile,
  Print #RptFile, "Total Closed PO's: " + Using(CommaFmt$, Str$(RunTotal#))
  If fpcboEncumber.ListIndex = 0 Then
    Print #RptFile, "   Total Canceled:  " + Using(CommaFmt$, Str$(TotEnc))
    Print #RptFile, " Total Liquidated:  " + Using(CommaFmt$, Str$(TotUn))

    Print #RptFile, "By Fund:"
    For FCnt = 1 To NumFunds
      If FundGrdTot#(FCnt) > 0 Then
        Print #RptFile, "Fund: "; FundList$(FCnt); " " + Using(CommaFmt$, Str$(FundGrdTot#(FCnt)))
      End If
    Next
  End If

  Print #RptFile, FF$

  Return

''UpdateGLAcct:   'Reseting the Encumbered Amt
''  Amt# = APDist.DistAmt
''  DistAcctRec = FindAcct(APDist.DistAcctNum)
''  If DistAcctRec > 0 Then
''    OpenAcctFile AcctFileNum, NumGLAcctRecs
''    Get AcctFileNum, DistAcctRec, Acct
''    Acct.Encumb = Acct.Encumb + Amt#
''    Put AcctFileNum, DistAcctRec, Acct
''    Close AcctFileNum
''  End If
''  Return
''
''
''ClearEnc:
''  OpenAcctFile AcctFileNum, NumGLAcctRecs
''  For Cnt1! = 1 To NumGLAcctRecs
''    Get AcctFileNum, Cnt1!, Acct
''    Acct.Encumb = 0
''    Put AcctFileNum, Cnt1!, Acct
''  Next Cnt1!
''  Close AcctFileNum
''  Return
CancelExit:
  Exit Sub
End Sub

Private Sub QPOSort(Idxbuff() As ChkSortType, lLBound, lUBound)
  Dim lngCurMid As Long, lngCurLow As Long, lngCurHigh As Long
  Dim Temp As ChkSortType
  Dim Temp2 As ChkSortType
  lngCurLow = lLBound
  lngCurHigh = lUBound
  'this is to exit loop if high and low are equal
  If lUBound <= lLBound Then Exit Sub 'GoTo Proc_Exit
    'find the midpoint
    lngCurMid = (lUBound + lLBound) \ 2
    '
    Temp = Idxbuff(lngCurMid)
    Do While (lngCurLow <= lngCurHigh)
      Do While Idxbuff(lngCurLow).CHKinfo < Temp.CHKinfo
        lngCurLow = lngCurLow + 1
        If lngCurLow = lUBound Then Exit Do
      Loop
      Do While Temp.CHKinfo < Idxbuff(lngCurHigh).CHKinfo
        lngCurHigh = lngCurHigh - 1
        If lngCurHigh = lLBound Then Exit Do
      Loop
      'if low is <= high then swap
      If (lngCurLow <= lngCurHigh) Then
        Temp2 = Idxbuff(lngCurLow)
        Idxbuff(lngCurLow) = Idxbuff(lngCurHigh)
        Idxbuff(lngCurHigh) = Temp2
    '
        lngCurLow = lngCurLow + 1
        lngCurHigh = lngCurHigh - 1
      End If
    Loop
  'recurse if necessary
    If lLBound < lngCurHigh Then
      QPOSort Idxbuff(), lLBound, lngCurHigh
    End If
  'recurse if necessary
    If lngCurLow < lUBound Then
      QPOSort Idxbuff(), lngCurLow, lUBound
    End If
End Sub




Private Sub fpcboDepartment_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboDepartment.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcboDepartment.ListIndex = -1
    fpcboDepartment.Action = ActionClearSearchBuffer
  End If
  If fpcboDepartment.ListDown <> True Then
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

Private Sub fpcboEncumber_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboEncumber.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcboEncumber.ListIndex = -1
    fpcboEncumber.Action = ActionClearSearchBuffer
  End If
  If fpcboEncumber.ListDown <> True Then
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

Private Sub fpcboSort_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboSort.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcboSort.ListIndex = -1
    fpcboSort.Action = ActionClearSearchBuffer
  End If
  If fpcboSort.ListDown <> True Then
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

Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub
