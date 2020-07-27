VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPrnAPChkList 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "A/P Check Listing"
   ClientHeight    =   8892
   ClientLeft      =   36
   ClientTop       =   492
   ClientWidth     =   12192
   Icon            =   "frmPrnAPChkList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8892
   ScaleWidth      =   12192
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboRptType 
      Height          =   384
      Left            =   5856
      TabIndex        =   5
      Top             =   6048
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
      ColDesigner     =   "frmPrnAPChkList.frx":08CA
   End
   Begin LpLib.fpCombo fpcboVoid 
      Height          =   384
      Left            =   5856
      TabIndex        =   4
      Top             =   5448
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
      ColDesigner     =   "frmPrnAPChkList.frx":0D6F
   End
   Begin LpLib.fpCombo fpcboDistributions 
      Height          =   384
      Left            =   5868
      TabIndex        =   3
      Top             =   4848
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
      ColDesigner     =   "frmPrnAPChkList.frx":11DD
   End
   Begin LpLib.fpCombo fpcboSort 
      Height          =   384
      Left            =   5868
      TabIndex        =   2
      Top             =   4260
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
      ColDesigner     =   "frmPrnAPChkList.frx":164B
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
      Top             =   7152
      Width           =   1332
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00D0D0D0&
      Caption         =   "F10 &Ok"
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
      Top             =   7152
      Width           =   1332
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   8
      Top             =   8532
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
            TextSave        =   "9:50 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7133
            TextSave        =   "11/18/2004"
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
      Left            =   5868
      TabIndex        =   0
      Top             =   3132
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
      Left            =   5868
      TabIndex        =   1
      Top             =   3684
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
      Left            =   3408
      TabIndex        =   15
      Top             =   6072
      Width           =   2388
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Voided Checks Only:"
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
      Left            =   3144
      TabIndex        =   14
      Top             =   5496
      Width           =   2580
   End
   Begin VB.Label Label4 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   4392
      TabIndex        =   13
      Top             =   4356
      Width           =   1164
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Show Distributions:"
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
      Left            =   3528
      TabIndex        =   12
      Top             =   4920
      Width           =   2196
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
      Left            =   4032
      TabIndex        =   11
      Top             =   3168
      Width           =   1668
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
      Left            =   4128
      TabIndex        =   10
      Top             =   3756
      Width           =   1572
   End
   Begin VB.Image Image1 
      Height          =   276
      Left            =   2496
      Picture         =   "frmPrnAPChkList.frx":1AB9
      Top             =   2736
      Width           =   288
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3216
      Top             =   936
      Width           =   5772
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A/P Check Listing"
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
      Left            =   3984
      TabIndex        =   9
      Top             =   1176
      Width           =   4332
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   4116
      Left            =   2172
      Top             =   2580
      Width           =   7860
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D0D0D0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00D0D0D0&
      Height          =   972
      Left            =   3216
      Top             =   816
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
Attribute VB_Name = "frmPrnAPChkList"
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
Dim VendorIdx As VendorIdxRecType
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer

Private Sub cmdExit_Click()
  frmAPReportsMenu.Show
  Unload frmPrnAPChkList
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
        fpcboVoid.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub
Private Sub cmdOk_Click()
 If Oktogo = True Then
   If fpcboRptType.ListIndex = 0 Then
    rptopt = 1
  ElseIf fpcboRptType.ListIndex = 1 Then
    rptopt = 2
  End If
  If rptopt = 1 Then
    PrintChkList
  ElseIf rptopt = 2 Then
    PrintChkList2
  End If
 End If
End Sub
Private Sub fpcboDistributions_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboDistributions.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcboDistributions.ListIndex = -1
    fpcboDistributions.Action = ActionClearSearchBuffer
  End If
  If fpcboDistributions.ListDown <> True Then
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

Private Sub fpcboVoid_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboVoid.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcboVoid.ListIndex = -1
    fpcboVoid.Action = ActionClearSearchBuffer
  End If
  If fpcboVoid.ListDown <> True Then
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
      SendKeys "%O"
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
  Me.HelpContextID = hlpAPChklisting
  fpcboSort.AddItem "Check Number Order"
  fpcboSort.AddItem "Vendor Order"
  fpcboSort.ListIndex = 0
  fpcboDistributions.AddItem "No"
  fpcboDistributions.AddItem "Yes"
  fpcboDistributions.ListIndex = 0
  fpDate1.Text = Format(Now, "mm/dd/yyyy")
  fpDate2.Text = Format(Now, "mm/dd/yyyy")
  fpcboVoid.AddItem "No"
  fpcboVoid.AddItem "Yes"
  fpcboVoid.ListIndex = 0
  fpcboRptType.InsertRow = "Graphics"
  fpcboRptType.InsertRow = "Text"
  fpcboRptType.ListIndex = 0
End Sub
Private Function Oktogo()
Dim TempDate1 As Integer, TempDate2 As Integer
    If CheckValDate(fpDate1) = False And CheckValDate(fpDate2) = False Then
      MsgBox "Date Is Not Valid. Please Correct.", vbOKOnly, "Invalid Date"
      Oktogo = False
    Else
      TempDate1 = DateDiff("d", "12/31/1979", fpDate1)
      TempDate2 = DateDiff("d", "12/31/1979", fpDate2)
      If TempDate1 > TempDate2 Then
        Oktogo = False
        MsgBox "The Starting And Ending Dates Must Be In Chronological Order Or Equal", vbOKOnly, "Invalid Date"
      Else
        Oktogo = True
      End If
    End If
End Function
Private Sub Form_Resize()
'  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
'  End If
End Sub

Private Sub PrintChkList()
  Dim BegDate As Integer, EndDate As Integer, SortSpec As String
  Dim ShowfundDist As Boolean, cnt As Long, ChkVen As String
  Dim Page As Integer, NumFunds As Integer, PRNFile As String
  Dim ColTitle As String, VendorFile As Integer, APDistRecLen As Integer
  Dim Header As String, A As String, CommaFmt As String, User As String
  Dim APLRecLen As Integer, APLedgerFile As Integer, NumTran As Long
  Dim APDRecLen As Integer, APDistFile As Integer, NumDistRecs As Long
  Dim RptFile As Integer, RptFileName As String, VRecLen As Integer
  Dim NumVRecs As Integer, OhShoot As Boolean, NumChks As Integer
  Dim Linecnt As Integer, Rec As Long, RunTotal As Double, Last As String
  Dim ToPrint As String, NextDist As Long, DistAmt As Double, Fundline As String
  Dim Found As Boolean, Fund As Integer, FundNum As String, FCnt As Integer
  Dim lngCurLow As Long, lngCurHigh As Long, RunToV As Double
  Dim vv As Boolean, distv As Double, RptFund As Integer, RptFundName As String
  BegDate = DateDiff("d", "12/31/1979", fpDate1)
  EndDate = DateDiff("d", "12/31/1979", fpDate2)
  SortSpec$ = Left$(fpcboSort.Text, 1)
  ShowfundDist = False
  FrmShowPctComp.Label1 = "Creating Check Listing Report"
  FrmShowPctComp.Show , Me
  DoEvents
  DeActivateControls frmPrnAPChkList, True
  If fpcboDistributions.ListIndex = 1 Then
    ShowfundDist = True 'unrem for new release
  Else
    ShowfundDist = False
  End If
  If fpcboSort.ListIndex = 1 Then
    'ColTitle$ = " Vendor                           Chk Num         Date              Amt"
    Header$ = "A/P Checks by Vendor"
  Else
    'ColTitle$ = " Chk Num  Vendor                                  Date              Amt"
    Header$ = "A/P Check Listing by Check"
  End If

  A$ = Space$(14)
  CommaFmt$ = "##,###,###.##"
  'FF$ = Chr$(12)
  'MaxLines = 55
  User$ = QPTrim$(GLUserName$)
  'Page = 0
  '--Get a list of active funds
  ReDim FundList(1) As String
  GetFundList FundList$(), NumFunds
  ReDim FundGrdTot#(1 To NumFunds)
  ReDim FundGrdV#(1 To NumFunds)
  Dim APLedger As APLedger81RecType
  APLRecLen = Len(APLedger)
  OpenAPLedgerFile APLedgerFile, NumTran&, APLRecLen

  'ReDim ChkList(1 To 1) As GLAcctIndexType      '--borrowing this type
'Needed file type with long record number
  ReDim ChkList(1 To 1) As ChkSortType
  
  Dim APDist As APDistRecType
  APDRecLen = Len(APDist)
  OpenAPDistFile APDistFile, NumDistRecs&, APDRecLen

  RptFile = FreeFile
  RptFileName$ = "apchks.prn"
  Open RptFileName$ For Output As RptFile
  RptFund = FreeFile
  RptFundName$ = "apchkFund.prn"
  Open RptFundName$ For Output As RptFund

  Dim Vendor As VendorRecType
  VRecLen = Len(Vendor)
  OpenVendorFile VendorFile, NumVRecs

  'GoSub OpenChkPageHdr

  OhShoot = False
  For cnt = 1 To NumTran&
    FrmShowPctComp.ShowPctComp cnt, NumTran&
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ActivateControls frmPrnAPChkList, True
      Unload FrmShowPctComp
      GoTo CancelExit
    End If
    If fpcboVoid.ListIndex = 0 Then
      Get APLedgerFile, cnt, APLedger
      If APLedger.VRecNum > 0 Then
        Get VendorFile, APLedger.VRecNum, Vendor
        If APLedger.TRCode = 3 Or APLedger.TRCode = -3 Then
          If APLedger.TRDATE >= BegDate And APLedger.TRDATE <= EndDate Then
            NumChks = NumChks + 1
            ReDim Preserve ChkList(1 To NumChks) As ChkSortType
            ChkList(NumChks).Record = cnt
            If fpcboSort.ListIndex = 0 Then
              RSet A$ = QPTrim$(APLedger.DOCNum)
              ChkList(NumChks).CHKinfo = A$
            Else
              ChkList(NumChks).CHKinfo = Vendor.VNAME
            End If
          End If
        End If
      End If
    Else
      Get APLedgerFile, cnt, APLedger
      If APLedger.VRecNum > 0 Then
        Get VendorFile, APLedger.VRecNum, Vendor
        If APLedger.TRCode = -3 Then
          If APLedger.TRDATE >= BegDate And APLedger.TRDATE <= EndDate Then
            NumChks = NumChks + 1
            ReDim Preserve ChkList(1 To NumChks) As ChkSortType
            ChkList(NumChks).Record = cnt
            If fpcboSort.ListIndex = 0 Then
              RSet A$ = QPTrim$(APLedger.DOCNum)
              ChkList(NumChks).CHKinfo = A$
            Else
              ChkList(NumChks).CHKinfo = Vendor.VNAME
            End If
          End If
        End If
      End If
    End If
  Next
  If NumTran& < 1 Then
    FrmShowPctComp.ShowPctComp 1, 1
  End If
  If NumChks > 0 Then
    lngCurLow = LBound(ChkList)
    lngCurHigh = UBound(ChkList)
    FrmShowPctComp.Label1 = "Sorting Checks"
    FrmShowPctComp.Show , Me
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ActivateControls frmPrnAPChkList, True
      Unload FrmShowPctComp
      GoTo CancelExit
    End If

    QCSort ChkList(), lngCurLow, lngCurHigh

    GoSub PrintChkList
  Else
    'Print #RptFile, "  No Checks on file."
  End If
   If ShowfundDist Then
   ' Print #RptFile, "           Total Checks     Voids          Check Total"
    For FCnt = 1 To NumFunds
      If FundGrdTot#(FCnt) > 0 Or FundGrdV#(FCnt) > 0 Then
        Fundline$ = FundList$(FCnt) + "~" + Using(CommaFmt$, Str$(FundGrdTot#(FCnt)))
        Fundline$ = Fundline$ + "~" + Using(CommaFmt$, Str(FundGrdV#(FCnt))) + "~"
        Fundline$ = Fundline$ + Using(CommaFmt$, Str(Round((FundGrdTot#(FCnt) - (FundGrdV#(FCnt))))))
        Print #RptFund, Fundline$

      End If
    Next
  End If

 
  Close
  
  'ViewPrint RptFileName$, Header$
  Load frmLoadingRpt
  ActivateControls frmPrnAPChkList, True
  If ShowfundDist = True Then
    If fpcboSort.ListIndex = 1 Then
      ARptAPChkList1.Order = 1
    Else
      ARptAPChkList1.Order = 2
    End If
    ARptAPChkList1.GetName RptFileName$, RptFundName$
    ARptAPChkList1.Label1.Caption = Header$
    ARptAPChkList1.txtTown.Caption = User$
    ARptAPChkList1.txtDate.Caption = Now
    ARptAPChkList1.Label17.Caption = "Report Dates : " + fpDate1 + "-" + fpDate2
    ARptAPChkList1.totChecks = Using(CommaFmt$, Str$(RunTotal#))
    ARptAPChkList1.totVoids = Using(CommaFmt$, Str$(RunToV#))
    ARptAPChkList1.Total = Using(CommaFmt$, Str$(Round#(RunTotal# - RunToV#)))
    ARptAPChkList1.startrpt
  Else
    If fpcboSort.ListIndex = 1 Then
      ARptAPCHkList2.Order = 1
    Else
      ARptAPCHkList2.Order = 2
    End If
    ARptAPCHkList2.GetName RptFileName$, RptFundName$
    ARptAPCHkList2.Label1.Caption = Header$
    ARptAPCHkList2.txtTown.Caption = User$
    ARptAPCHkList2.txtDate.Caption = Now
    ARptAPCHkList2.Label17.Caption = "Report Dates : " + fpDate1 + "-" + fpDate2
    ARptAPCHkList2.totChecks = Using(CommaFmt$, Str$(RunTotal#))
    ARptAPCHkList2.totVoids = Using(CommaFmt$, Str$(RunToV#))
    ARptAPCHkList2.Total = Using(CommaFmt$, Str$(Round#(RunTotal# - RunToV#)))
    ARptAPCHkList2.startrpt

   End If
  
  Exit Sub

PrintChkList:
  For cnt = 1 To NumChks
    FrmShowPctComp.ShowPctComp cnt, NumChks
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ActivateControls frmPrnAPChkList, True
      Unload FrmShowPctComp
      GoTo CancelExit
    End If

    A$ = Space$(14)
    'IF Cnt = 1 THEN
    '  ThisNum& = VAL(ChkList(Cnt).AcctNum)
    'ELSE
    '  IF ThisNum& = VAL(ChkList(Cnt).AcctNum) THEN
    '    Rec = ChkList(Cnt).RecNum
    '    GET APLedgerFile, Rec, APLedger
    '    APLedger.VRecNum = -1
    '    'APLedger.TrCode = -99
    '    PUT APLedgerFile, Rec, APLedger
    '    GOTO SkipIt
    '    'STOP
    '  ELSE
    '    ThisNum& = VAL(ChkList(Cnt).AcctNum)
    '  END IF
    'END IF

    Rec = ChkList(cnt).Record
    Get APLedgerFile, Rec, APLedger
    Get VendorFile, APLedger.VRecNum, Vendor
    If APLedger.TRCode = -3 Then
      RunToV# = RunToV# + APLedger.Amt
    End If
    RunTotal# = RunTotal# + APLedger.Amt
    ChkVen$ = Space$(80)
    If fpcboSort.ListIndex = 1 Then
      ChkVen$ = Vendor.VNAME + "~" + Left$(APLedger.DOCNum, 15) + "~"
    Else
      ChkVen$ = Left$(APLedger.DOCNum, 15) + "~" + Vendor.VNAME + "~"
    End If
    ChkVen$ = ChkVen$ + Format(DateAdd("d", (APLedger.TRDATE), "12-31-1979"), "mm/dd/yyyy")
    ChkVen$ = ChkVen$ + "~" + Using(CommaFmt$, Str$(APLedger.Amt))

    'IF INSTR(APLedger.DOCNum, "44628") > 0 THEN STOP

    If APLedger.TRCode = -3 Then
      ChkVen$ = ChkVen$ + "~Void~"
    Else
      ChkVen$ = ChkVen$ + "~    ~"
    End If
'    Print #RptFile, ToPrint$

'    Linecnt = Linecnt + 1
'    If Linecnt > MaxLines Then
'      Print #RptFile, FF$
'      GoSub OpenChkPageHdr
'    End If

    If ShowfundDist Then
      NextDist& = APLedger.FrstDist
      If APLedger.TRCode = -3 Then
        vv = True
      Else
        vv = False
      End If
      DistAmt# = 0
      distv# = 0
      If NextDist& > 0 Then
        'Print #RptFile,
        'Print #RptFile, Tab(50); "Fund Distribution:"
        'Linecnt = Linecnt + 2

        Do
          Get APDistFile, NextDist&, APDist
          DistAmt# = Round#(DistAmt# + APDist.DistAmt)
          If vv = True Then
            distv# = Round#(distv# + APDist.DistAmt)
          End If
          ToPrint$ = Space$(30)
          ToPrint$ = Left$(APDist.DistAcctNum, GLFundLen)
          ToPrint$ = ToPrint$ + "~" + Using(CommaFmt$, Str$(APDist.DistAmt))
          Print #RptFile, ChkVen$ + ToPrint$
          'Linecnt = Linecnt + 1
          'If Linecnt > MaxLines Then
          '  Print #RptFile, FF$
          '  GoSub OpenChkPageHdr
          'End If

          '--summarize by fund
          Found = False
          For Fund = 1 To NumFunds
            FundNum$ = Left$(APDist.DistAcctNum, GLFundLen)
            If FundNum$ = FundList$(Fund) Then
              Found = True
              If vv Then
                FundGrdV#(Fund) = Round#(FundGrdV#(Fund) + APDist.DistAmt)
              End If
              FundGrdTot#(Fund) = Round#(FundGrdTot#(Fund) + APDist.DistAmt)
              Exit For
            End If
          Next

          NextDist& = APDist.NextDist

        Loop Until NextDist& = 0
        'Print #RptFile, String$(78, "-")
      End If    '--showing Distribution
    Else
      Print #RptFile, ChkVen$
    End If
  Next
Return
'SkipIt:
'

'  Print #RptFile,
'  Print #RptFile, "Total Checks Listed: " + Using(CommaFmt$, Str$(RunTotal#))
'  Print #RptFile, "Total Checks Voided: " + Using(CommaFmt$, Str$(RunToV#))
'  Print #RptFile, "                     ==============="
'  Print #RptFile, "Total Check Amount : " + Using(CommaFmt$, Str$(Round#(RunTotal# - RunToV#)))
'  Print #RptFile,
  'Print #RptFile, FF$

  
CancelExit:
  Exit Sub
  
End Sub
Private Sub PrintChkList2()
  Dim BegDate As Integer, EndDate As Integer, SortSpec As String
  Dim ShowfundDist As Boolean, cnt As Long, FF As String, MaxLines As Integer
  Dim Page As Integer, NumFunds As Integer, PRNFile As String
  Dim ColTitle As String, VendorFile As Integer, APDistRecLen As Integer
  Dim Header As String, A As String, CommaFmt As String, User As String
  Dim APLRecLen As Integer, APLedgerFile As Integer, NumTran As Long
  Dim APDRecLen As Integer, APDistFile As Integer, NumDistRecs As Long
  Dim RptFile As Integer, RptFileName As String, VRecLen As Integer
  Dim NumVRecs As Integer, OhShoot As Boolean, NumChks As Integer
  Dim Linecnt As Integer, Rec As Long, RunTotal As Double
  Dim ToPrint As String, NextDist As Long, DistAmt As Double
  Dim Found As Boolean, Fund As Integer, FundNum As String, FCnt As Integer
  Dim lngCurLow As Long, lngCurHigh As Long, RunToV As Double
  Dim vv As Boolean, distv As Double
  BegDate = DateDiff("d", "12/31/1979", fpDate1)
  EndDate = DateDiff("d", "12/31/1979", fpDate2)
  SortSpec$ = Left$(fpcboSort.Text, 1)
  ShowfundDist = False
  FrmShowPctComp.Label1 = "Creating Check Listing Report"
  FrmShowPctComp.Show , Me
  DoEvents
  DeActivateControls frmPrnAPChkList, True
  If fpcboDistributions.ListIndex = 1 Then
    ShowfundDist = True 'unrem for new release
  Else
    ShowfundDist = False
  End If
  If fpcboSort.ListIndex = 1 Then
    ColTitle$ = " Vendor                           Chk Num         Date              Amt"
    Header$ = "A/P Checks by Vendor"
  Else
    ColTitle$ = " Chk Num  Vendor                                  Date              Amt"
    Header$ = "A/P Check Listing"
  End If

  A$ = Space$(14)
  CommaFmt$ = "##,###,###.##"
  FF$ = Chr$(12)
  MaxLines = 55
  User$ = QPTrim$(GLUserName$)
  Page = 0
  '--Get a list of active funds
  ReDim FundList(1) As String
  GetFundList FundList$(), NumFunds
  ReDim FundGrdTot#(1 To NumFunds)
  ReDim FundGrdV#(1 To NumFunds)
  Dim APLedger As APLedger81RecType
  APLRecLen = Len(APLedger)
  OpenAPLedgerFile APLedgerFile, NumTran&, APLRecLen

  'ReDim ChkList(1 To 1) As GLAcctIndexType      '--borrowing this type
  ReDim ChkList(1 To 1) As ChkSortType    'use for long record
  Dim APDist As APDistRecType
  APDRecLen = Len(APDist)
  OpenAPDistFile APDistFile, NumDistRecs&, APDRecLen

  RptFile = FreeFile
  RptFileName$ = "apchks.prn"
  Open RptFileName$ For Output As RptFile

  Dim Vendor As VendorRecType
  VRecLen = Len(Vendor)
  OpenVendorFile VendorFile, NumVRecs

  GoSub OpenChkPageHdr

  OhShoot = False
  For cnt = 1 To NumTran&
    FrmShowPctComp.ShowPctComp cnt, NumTran&
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ActivateControls frmPrnAPChkList, True
      Unload FrmShowPctComp
      GoTo CancelExit
    End If
    If fpcboVoid.ListIndex = 0 Then
      Get APLedgerFile, cnt, APLedger
      If APLedger.VRecNum > 0 Then
        Get VendorFile, APLedger.VRecNum, Vendor
        If APLedger.TRCode = 3 Or APLedger.TRCode = -3 Then
          If APLedger.TRDATE >= BegDate And APLedger.TRDATE <= EndDate Then
            NumChks = NumChks + 1
            ReDim Preserve ChkList(1 To NumChks) As ChkSortType
            ChkList(NumChks).Record = cnt
            If fpcboSort.ListIndex = 0 Then
              RSet A$ = QPTrim$(APLedger.DOCNum)
              ChkList(NumChks).CHKinfo = A$
            Else
              ChkList(NumChks).CHKinfo = Vendor.VNAME
            End If
          End If
        End If
      End If
    Else
      Get APLedgerFile, cnt, APLedger
      If APLedger.VRecNum > 0 Then
        Get VendorFile, APLedger.VRecNum, Vendor
        If APLedger.TRCode = -3 Then
          If APLedger.TRDATE >= BegDate And APLedger.TRDATE <= EndDate Then
            NumChks = NumChks + 1
            ReDim Preserve ChkList(1 To NumChks) As ChkSortType
            ChkList(NumChks).Record = cnt
            If fpcboSort.ListIndex = 0 Then
              RSet A$ = QPTrim$(APLedger.DOCNum)
              ChkList(NumChks).CHKinfo = A$
            Else
              ChkList(NumChks).CHKinfo = Vendor.VNAME
            End If
          End If
        End If
      End If
    End If
  Next
  If NumTran& < 1 Then
    FrmShowPctComp.ShowPctComp 1, 1
  End If
  If NumChks > 0 Then
    lngCurLow = LBound(ChkList)
    lngCurHigh = UBound(ChkList)
    FrmShowPctComp.Label1 = "Sorting Checks"
    FrmShowPctComp.Show , Me

    QCSort ChkList(), lngCurLow, lngCurHigh

    GoSub PrintChkList
  Else
    Print #RptFile, "  No Checks on file."
  End If
  Close
  ActivateControls frmPrnAPChkList, True
  ViewPrint RptFileName$, Header$
  Exit Sub

OpenChkPageHdr:
  Page = Page + 1
  Print #RptFile, Tab(40 - (Int(Len(User$) / 2))); User$
  Print #RptFile, Tab(40 - (Int(Len(Header$) / 2))); Header$
  Print #RptFile,
  Print #RptFile, "Report Date: "; Date$; Tab(67); "Page #"; Page
  Print #RptFile, ColTitle$
  Print #RptFile, String$(78, "=")
  Linecnt = 5
  Return

PrintChkList:
  For cnt = 1 To NumChks
    FrmShowPctComp.ShowPctComp cnt, NumChks
    A$ = Space$(14)
    'IF Cnt = 1 THEN
    '  ThisNum& = VAL(ChkList(Cnt).AcctNum)
    'ELSE
    '  IF ThisNum& = VAL(ChkList(Cnt).AcctNum) THEN
    '    Rec = ChkList(Cnt).RecNum
    '    GET APLedgerFile, Rec, APLedger
    '    APLedger.VRecNum = -1
    '    'APLedger.TrCode = -99
    '    PUT APLedgerFile, Rec, APLedger
    '    GOTO SkipIt
    '    'STOP
    '  ELSE
    '    ThisNum& = VAL(ChkList(Cnt).AcctNum)
    '  END IF
    'END IF

    Rec = ChkList(cnt).Record
    Get APLedgerFile, Rec, APLedger
    Get VendorFile, APLedger.VRecNum, Vendor
    If APLedger.TRCode = -3 Then
      RunToV# = RunToV# + APLedger.Amt
    End If
    RunTotal# = RunTotal# + APLedger.Amt
    ToPrint$ = Space$(80)
    If fpcboSort.ListIndex = 1 Then
    Mid$(ToPrint$, 2) = Vendor.VNAME
      Mid$(ToPrint$, 35) = Left$(APLedger.DOCNum, 15)
    Else
      Mid$(ToPrint$, 2) = Left$(APLedger.DOCNum, 15)
      Mid$(ToPrint$, 10) = Vendor.VNAME
    End If
    Mid$(ToPrint$, 51) = Format(DateAdd("d", (APLedger.TRDATE), "12-31-1979"), "mm/dd/yyyy")
    Mid$(ToPrint$, 62) = Using(CommaFmt$, Str$(APLedger.Amt))

    'IF INSTR(APLedger.DOCNum, "44628") > 0 THEN STOP

    If APLedger.TRCode = -3 Then
      Mid$(ToPrint$, 76) = "Void"
    End If

    Print #RptFile, ToPrint$

    Linecnt = Linecnt + 1
    If Linecnt > MaxLines Then
      Print #RptFile, FF$
      GoSub OpenChkPageHdr
    End If

    If ShowfundDist Then
      NextDist& = APLedger.FrstDist
      If APLedger.TRCode = -3 Then
        vv = True
      Else
        vv = False
      End If
      DistAmt# = 0
      distv# = 0
      If NextDist& > 0 Then
        Print #RptFile,
        Print #RptFile, Tab(50); "Fund Distribution:"
        Linecnt = Linecnt + 2

        Do
          Get APDistFile, NextDist&, APDist
          DistAmt# = Round#(DistAmt# + APDist.DistAmt)
          If vv = True Then
            distv# = Round#(distv# + APDist.DistAmt)
          End If
          ToPrint$ = Space$(80)
          Mid$(ToPrint$, 50) = Left$(APDist.DistAcctNum, GLFundLen)
          Mid$(ToPrint$, 62) = Using(CommaFmt$, Str$(APDist.DistAmt))
          Print #RptFile, ToPrint$
          Linecnt = Linecnt + 1
          If Linecnt > MaxLines Then
            Print #RptFile, FF$
            GoSub OpenChkPageHdr
          End If

          '--summarize by fund
          Found = False
          For Fund = 1 To NumFunds
            FundNum$ = Left$(APDist.DistAcctNum, GLFundLen)
            If FundNum$ = FundList$(Fund) Then
              Found = True
              If vv Then
                FundGrdV#(Fund) = Round#(FundGrdV#(Fund) + APDist.DistAmt)
              End If
              FundGrdTot#(Fund) = Round#(FundGrdTot#(Fund) + APDist.DistAmt)
              Exit For
            End If
          Next

          NextDist& = APDist.NextDist

        Loop Until NextDist& = 0
        Print #RptFile, String$(78, "-")
      End If    '--showing Distribution
    End If
SkipIt:
  Next

  Print #RptFile,
  Print #RptFile, "Total Checks Listed: " + Using(CommaFmt$, Str$(RunTotal#))
  Print #RptFile, "Total Checks Voided: " + Using(CommaFmt$, Str$(RunToV#))
  Print #RptFile, "                     ==============="
  Print #RptFile, "Total Check Amount : " + Using(CommaFmt$, Str$(Round#(RunTotal# - RunToV#)))
  Print #RptFile,
  If ShowfundDist Then
    Print #RptFile, "           Total Checks     Voids          Check Total"
    For FCnt = 1 To NumFunds
      If FundGrdTot#(FCnt) > 0 Or FundGrdV#(FCnt) > 0 Then
        Print #RptFile, "Fund: "; FundList$(FCnt); " " + Using(CommaFmt$, Str$(FundGrdTot#(FCnt))) + " (" + Using(CommaFmt$, Str(FundGrdV#(FCnt))) + ") = " + Using(CommaFmt$, Str(Round((FundGrdTot#(FCnt) - (FundGrdV#(FCnt))))))

      End If
    Next
  End If

  Print #RptFile, FF$

  Return
CancelExit:
  Exit Sub
  
End Sub


Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub
