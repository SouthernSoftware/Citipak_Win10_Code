VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Begin VB.Form frmPrnBudPrepWork 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Budget Preparation Worksheet Report Options"
   ClientHeight    =   8640
   ClientLeft      =   30
   ClientTop       =   540
   ClientWidth     =   12195
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   7.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPrnBudPrepWork.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   12195
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboPagebrk 
      Height          =   405
      Left            =   5910
      TabIndex        =   6
      Top             =   6270
      Width           =   1395
      _Version        =   196608
      _ExtentX        =   2461
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
      ColDesigner     =   "frmPrnBudPrepWork.frx":08CA
   End
   Begin LpLib.fpCombo fpcboFund2 
      Height          =   405
      Left            =   5475
      TabIndex        =   1
      Top             =   3525
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
      Columns         =   3
      Sorted          =   0
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   2
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
      ColDesigner     =   "frmPrnBudPrepWork.frx":0CF1
   End
   Begin LpLib.fpCombo fpcboFund1 
      Height          =   405
      Left            =   5475
      TabIndex        =   0
      Top             =   2955
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
      Object.TabStop         =   -1  'True
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Text            =   ""
      Columns         =   3
      Sorted          =   0
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   2
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
      ColDesigner     =   "frmPrnBudPrepWork.frx":119C
   End
   Begin LpLib.fpCombo fpcboRptType 
      Height          =   405
      Left            =   5910
      TabIndex        =   5
      Top             =   5715
      Width           =   1905
      _Version        =   196608
      _ExtentX        =   3360
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
      ColDesigner     =   "frmPrnBudPrepWork.frx":1647
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
      Left            =   9456
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7344
      Width           =   1524
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
      Left            =   7488
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7368
      Width           =   1572
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   9
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
            TextSave        =   "12:39 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7117
            TextSave        =   "1/24/2008"
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
   Begin EditLib.fpMask txtAcctCode 
      Height          =   372
      Left            =   5904
      TabIndex        =   3
      Top             =   4632
      Width           =   1092
      _Version        =   196608
      _ExtentX        =   1926
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
      NullColor       =   -2147483639
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   0
      AllowOverflow   =   0   'False
      BestFit         =   0   'False
      ClipMode        =   0
      DataFormatEx    =   0
      Mask            =   ""
      PromptChar      =   "_"
      PromptInclude   =   0   'False
      RequireFill     =   -1  'True
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
   Begin EditLib.fpMask txtDetCode 
      Height          =   372
      Left            =   5904
      TabIndex        =   4
      Top             =   5160
      Width           =   1068
      _Version        =   196608
      _ExtentX        =   1884
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
      NullColor       =   -2147483639
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   0
      AllowOverflow   =   0   'False
      BestFit         =   0   'False
      ClipMode        =   0
      DataFormatEx    =   0
      Mask            =   ""
      PromptChar      =   "_"
      PromptInclude   =   0   'False
      RequireFill     =   -1  'True
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      AutoTab         =   -1  'True
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpDateTime txtDate 
      Height          =   372
      Left            =   5880
      TabIndex        =   2
      Top             =   4104
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
      ButtonColor     =   14737632
      AutoMenu        =   0   'False
      StartMonth      =   4
      ButtonAlign     =   0
      BoundDataType   =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Page Break on Dept:"
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
      Height          =   324
      Index           =   3
      Left            =   3408
      TabIndex        =   19
      Top             =   6288
      Width           =   2364
   End
   Begin VB.Label Label4 
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
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   3456
      TabIndex        =   18
      Top             =   5760
      Width           =   2388
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Calculate Actual Thru:"
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
      Height          =   324
      Index           =   2
      Left            =   3144
      TabIndex        =   17
      Top             =   4140
      Width           =   2628
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Acct Code:"
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
      Height          =   324
      Index           =   0
      Left            =   4320
      TabIndex        =   16
      Top             =   4668
      Width           =   1452
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Detail Code:"
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
      Height          =   324
      Index           =   2
      Left            =   4080
      TabIndex        =   15
      Top             =   5232
      Width           =   1692
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "To:"
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
      Height          =   324
      Index           =   1
      Left            =   4848
      TabIndex        =   14
      Top             =   3588
      Width           =   492
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fund Codes"
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
      Index           =   1
      Left            =   2880
      TabIndex        =   13
      Top             =   3300
      Width           =   1428
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "From:"
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
      Height          =   324
      Index           =   8
      Left            =   4704
      TabIndex        =   12
      Top             =   3000
      Width           =   636
   End
   Begin VB.Line Line1 
      X1              =   4368
      X2              =   4368
      Y1              =   3684
      Y2              =   3204
   End
   Begin VB.Line Line2 
      X1              =   4368
      X2              =   4896
      Y1              =   3684
      Y2              =   3684
   End
   Begin VB.Line Line3 
      X1              =   4368
      X2              =   4656
      Y1              =   3204
      Y2              =   3204
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Select An Individual Fund Or All Funds For The Budget Prep Worksheet Rpt:"
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
      Height          =   708
      Index           =   0
      Left            =   3936
      TabIndex        =   11
      Top             =   2160
      Width           =   4500
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3240
      Top             =   648
      Width           =   5772
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Budget Prep Worksheet Rpt Options"
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
      Left            =   3312
      TabIndex        =   10
      Top             =   888
      Width           =   5652
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   4956
      Left            =   2688
      Top             =   1920
      Width           =   6972
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D0D0D0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00D0D0D0&
      Height          =   972
      Left            =   3240
      Top             =   528
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
Attribute VB_Name = "frmPrnBudPrepWork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim GLSetup As GLSetupRecType
Dim GLFundIdx As GLFundIndexType
Dim Acct    As GLAcctRecType
Dim AcctIdx As GLAcctIndexType
Dim GLTrans   As GLTransRecType
Dim GLFund As GLFundRecType

Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Dim StartFund As String, EndFund As String, AcctCode As String, Detcode As String
Dim acctmsk As String, detmsk As String
Dim FY1BegDate As Integer, FY1EndDate As Integer, FY2BegDate As Integer, FY2EndDate As Integer
Dim FYStartDate As Integer, ActiveYear As Integer
Private Sub cmdExit_Click()
  frmBudgetPrepMenu.Show
  Unload Me
End Sub
Private Function ValidDate()
  Dim TempDate As Integer
  GetFYDates FY1BegDate, FY1EndDate, FY2BegDate, FY2EndDate
  If CheckValDate(txtDate) = True Then
    TempDate = DateDiff("d", "12/31/1979", txtDate)
    ValidDate = True
    If TempDate >= FY2BegDate Then
      ActiveYear = 2
      FYStartDate = FY2BegDate
    Else
      ActiveYear = 1
      FYStartDate = FY1BegDate
    End If
  Else
    MsgBox "Date Is Not Valid. Please Correct.", vbOKOnly, "Invalid Date"
    ValidDate = False
    Exit Function
  
  End If
End Function

Private Function ValidFunds()
  If fpcboFund1.Text <> "" And fpcboFund2.Text <> "" Then
    fpcboFund1.Col = 0
    fpcboFund2.Col = 0
    If fpcboFund1.ColText > fpcboFund2.ColText Then
      MsgBox "Invalid Fund Selection, The Beginning Fund Should Be Less or Equal to Ending Fund.", vbOKOnly, "Invalid Selection"
      ValidFunds = False
    Else
      ValidFunds = True
      fpcboFund1.Col = 0
      fpcboFund2.Col = 0
      StartFund = QPTrim(fpcboFund1.ColText)
      EndFund = QPTrim(fpcboFund2.ColText)
    End If
  Else
    MsgBox "Fund Selections May Not Be Left Blank.", vbOKOnly, "Invalid Selection"
  End If
End Function

Private Sub cmdPrint_Click()
  If ValidFunds And ValidDate Then
    If fpcboRptType.ListIndex = 1 Then
      PrintWorksheetTxt
    Else
      PrintWorksheetGph
    End If
  End If
End Sub

'Private Sub cmdDisplay_Click()
'If ValidFunds = True Then
'  frmBudPrepMaint.SetOptions StartFund, EndFund, AcctCode, Detcode
'  Call MainLog("BudPrepOptions: " + StartFund + "," + EndFund + "," + AcctCode + "," + Detcode)
'  frmBudPrepMaint.Show
'  Unload frmBudPrepOptions
'End If
'End Sub

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
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      ClearInUse PWcnt
    End If
  End If
End Sub

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  detmsk = String(GLDetLen, "#")
  acctmsk = String(GLAcctLen, "#")
  txtDate.Text = Format(Now, "mm/dd/yyyy")
  StatusBar1.Panels.Item(1).Text = GLUserName
  txtAcctCode.Mask = acctmsk
  txtDetCode.Mask = detmsk
  FundstoList fpcboFund1
  FundstoList fpcboFund2
  AcctCode = ""
  Detcode = ""
  fpcboRptType.InsertRow = "Graphics"
  fpcboRptType.InsertRow = "Text"
  fpcboRptType.ListIndex = 0
  fpcboPagebrk.AddItem "Yes"
  fpcboPagebrk.AddItem "No"
  fpcboPagebrk.ListIndex = 1
  Me.HelpContextID = hlpBudPrepWSOptions
End Sub

Private Sub Form_Resize()
'  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
'  End If
End Sub

Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub txtAcctCode_LostFocus()
  Dim Num As String
  Num = Trim(txtAcctCode)
  If (Len(Num)) > 1 Then
    If (Len(Num)) <> (Val(GLAcctLen)) Then
      MsgBox "Invalid Code.", vbOKOnly, "Invalid Data!"
      txtAcctCode.Mask = acctmsk
      txtAcctCode.SetFocus
    Else
      AcctCode = Num
    End If
  Else
    AcctCode = ""
  End If
End Sub
Private Sub txtAcctCode_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub
Private Sub txtDetCode_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub txtDetCode_LostFocus()
  Dim Num As String
  Num = Trim(txtDetCode)
  If (Len(Num)) > 1 Then
    If (Len(Num)) <> (Val(GLDetLen)) Then
      MsgBox "Invalid Code.", vbOKOnly, "Invalid Data!"
      txtDetCode.Mask = detmsk
      txtDetCode.SetFocus
    Else
      Detcode = Num
    End If
  Else
    Detcode = ""
  End If
End Sub

Private Sub fpcboFund1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboFund1.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcboFund1.ListIndex = -1
    fpcboFund1.Action = ActionClearSearchBuffer
  End If
  If fpcboFund1.ListDown <> True Then
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
Private Sub fpcboFund2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboFund2.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcboFund2.ListIndex = -1
    fpcboFund2.Action = ActionClearSearchBuffer
  End If
  If fpcboFund2.ListDown <> True Then
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

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub
Private Sub PrintWorksheetTxt()
  Dim CommaFmt As String, TotalFmt As String, RunBalFmt As String
  Dim SumLine As String, BgtFmt As String, BSumLine As String, PSumLine As String
  Dim DivLine As String, DivLine2 As String, FF As String, RptTitle As String
  Dim MaxLines As Integer, Col1 As Integer, Col2 As Integer, Col3 As Integer
  Dim Col4 As Integer, Col5 As Integer, EndDate As Integer, PRNFile As Integer
  Dim M As String, HollyFlag As Boolean, Pitch12 As String, ThisFund As String
  Dim DoingDetail As Boolean, SubTotalRevenues As Boolean, DeptOnNewPage As Boolean
  Dim WhichReport As Integer, GetMonth As Boolean, GetQtr As Boolean
  Dim RptMonth As String, ReportFile As String, FundCode As String
  Dim FundIdxFile As Integer, NumFunds As Integer, GLAcct As Integer
  Dim AcctIdxFileNum As Integer, NumGLAccts As Integer, FundName As String
  Dim AcctFileNum As Integer, NumGLAcctRecs As Integer, Rec As Integer
  Dim TransFileNum As Integer, NumTrans As Long, NextTr As Long, AcctNum As String
  Dim BGTBal As Double, YTDBal As Double, MTDBal As Double, UsingFund As Boolean
  Dim ECnt As Integer, RCnt As Integer, d As String, TransMonth As String
  Dim InThisQtr As Boolean, FirstTime As Boolean, Fund As Integer, FundRec As Integer
  Dim Dept As String, LastDept As String, LastDeptName As String, cnt As Integer
  Dim Account As String, BudgetAmt As Double, DeptRecNum As Integer, DeptName As String
  Dim Pct As String, Variance As Double, ToPrint As String, Linecnt As Integer
  Dim MTDSum As Double, BgtSum As Double, YTDSum As Double, DeptBgtSum As Double
  Dim DeptYTDSum As Double, DeptENCSum As Double, DeptMTDSum As Double
  Dim FundRevMTD As Double, FundRevBgt As Double, FundRevYTD As Double
  Dim EncSum As Double, FundExpMTD As Double, FundExpBgt As Double
  Dim FundExpYTD As Double, FundEncYTD As Double, EncBal As Double
  Dim DeptSummary As String, PageNum As Integer, Newrp As String
  Dim Underline As String, CrLF As String, Col6 As Integer, Col7 As Integer
  Dim GLFile As Integer, GLSetUp1 As Integer, Category As String
  Dim DoingExp As Long, FirstTimeThru As Integer, PYActSum As Double
  Dim NYEstSum As Double, NYReqSum As Double, NYRecSum As Double
  Dim NYAppSum As Double, FundRevPYActSum As Double, FundRevNYEstSum As Double
  Dim FundRevNYReqSum As Double, FundRevNYRecSum As Double, FundRevNYAppSum As Double
  Dim NewPage  As Boolean, Break As Long, ObjCode As String, BreakFlag As Boolean
  Dim DeptPYActSum As Double, AcctNumber As String
  Dim DeptNYEstSum As Double, Det As String
  Dim DeptNYReqSum As Double
  Dim DeptNYRecSum As Double
  Dim DeptNYAppSum As Double

  Dim DeptBgtSum1 As Double
  Dim DeptYTDSum1 As Double
  Dim DeptPYActSum1 As Double
  Dim DeptNYEstSum1 As Double
  Dim DeptNYReqSum1 As Double
  Dim DeptNYRecSum1 As Double
  Dim DeptNYAppSum1 As Double
  Dim PYActBal As Double
  Dim NYEstBal As Double
  Dim NYReqBal As Double
  Dim NYRecBal As Double
  Dim NYAppBal As Double
  Dim FundRevPYSum As Double
  Dim DeptPYSum As Double
  ReDim Beg(100), BEnd(100), Desc$(100)

'  GetFundCodes FirstFund$, LastFund$
'
'
'  Form$(1, 0) = Date$
'  Form$(2, 0) = FirstFund$
'  Form$(3, 0) = LastFund$
'  Form$(4, 0) = "R"
'  Form$(5, 0) = "Y"
'

'  Do
'
'      Select Case Frm.KeyCode
'        '--SaveButton
'        Case F10Key
'          EndDate = Date2Num(Form$(1, 0))
'          StartFund$ = QPTrim$(Form$(2, 0))
'          EndFund$ = QPTrim$(Form$(3, 0))
'
'          '--print worksheet or completed report
'          Select Case Form$(4, 0)
'            Case "W"
'              WorkSheet = True
'            Case "R", ""
'              WorkSheet = False
'          End Select
'          '--print depts on separate pages
'          Select Case Form$(5, 0)
'            Case "Y"
'               DeptOnNewPage = True
'            Case "N", ""
'              DeptOnNewPage = False
'          End Select
'          '--Screen or printer
'          If Len(LTrim$(RTrim$(Form$(6, 0)))) = 0 Then
'             Dev$ = "S"
'             LPTNo = 1
'          Else
'              Dev$ = Left$(Form$(6, 0), 1)
'             LPTNo = Val(Right$(RTrim$(Form$(6, 0)), 1))
'          End If
'        Case EscKey
'          Exit Sub
'      End Select
'
'  Loop Until Frm.KeyCode = F10Key
    If fpcboPagebrk.Text = "Yes" Then
      DeptOnNewPage = True
    End If

  EndDate = DateDiff("d", "12/31/1979", txtDate)
  ReportFile$ = "BGTPREP.PRN"    'Report File Name
  CommaFmt$ = "###,###,###.##"    'format takes 14 chars
  BgtFmt$ = "#,###,###,###"         'format takes 13 chars
  TotalFmt$ = "###,###,###.##"   'format takes 14 chars
  SumLine$ = String$(12, "-")   'column summary line
  BSumLine$ = String$(10, "-")  'summary line for budget columns
  PSumLine$ = "----"            'summary line for Pct columns
  DivLine$ = String$(100, "-")   'dashed line
  DivLine2$ = String$(100, "=")  'Double Line
  Underline$ = "______________"
  
  CrLF$ = Chr$(13) + Chr$(10)
  FF$ = Chr$(12)
  MaxLines = 55

  '--Column offsets for printing amounts
    Col1 = 30
    Col2 = 43
    Col3 = 58
    Col4 = 73
    Col5 = 88
    Col6 = 103
    Col7 = 118

  PRNFile = FreeFile
  Open ReportFile$ For Output As #PRNFile

  OpenAcctIdx AcctIdxFileNum, NumGLAccts
  OpenAcctFile AcctFileNum, NumGLAcctRecs
  OpenTransFile TransFileNum, NumTrans&
  OpenFundIdx FundIdxFile, NumFunds

'  GLFile = FreeFile
'    GlSetuplen = Len(GLSetUp1)
'  Open "GLSETSUM.DAT" For Random Access Read Write Shared As GLFile Len = GlSetuplen
'  If LOF(GLFile) > 0 Then
'   Get GLFile, 1, GLSetUp1
'   Beg(1) = Val(GLSetUp1.Beg1)
'   BEnd(1) = Val(GLSetUp1.End1)
'   Desc$(1) = GLSetUp1.DESC1
'   Beg(2) = Val(GLSetUp1.Beg2)
'   BEnd(2) = Val(GLSetUp1.End2)
'   Desc$(2) = GLSetUp1.DESC2
'   Beg(3) = Val(GLSetUp1.Beg3)
'   BEnd(3) = Val(GLSetUp1.End3)
'   Desc$(3) = GLSetUp1.DESC3
'   Beg(4) = Val(GLSetUp1.Beg4)
'   BEnd(4) = Val(GLSetUp1.End4)
'   Desc$(4) = GLSetUp1.DESC4
'   Beg(5) = Val(GLSetUp1.Beg5)
'   BEnd(5) = Val(GLSetUp1.End5)
'   Desc$(5) = GLSetUp1.DESC5
'   Beg(6) = Val(GLSetUp1.Beg6)
'   BEnd(6) = Val(GLSetUp1.End6)
'    Desc$(6) = GLSetUp1.Desc6
'   Beg(7) = Val(GLSetUp1.Beg7)
'   BEnd(7) = Val(GLSetUp1.End7)
'   Desc$(7) = GLSetUp1.Desc7
'   Beg(8) = Val(GLSetUp1.Beg8)
'   BEnd(8) = Val(GLSetUp1.End8)
'   Desc$(8) = GLSetUp1.Desc8
'   Beg(9) = Val(GLSetUp1.Beg9)
'   BEnd(9) = Val(GLSetUp1.End9)
'   Desc$(9) = GLSetUp1.Desc9
'   BreakFlag = True
'   Else
   BreakFlag = False
'   Close GLFile
'  End If



  ReDim RevAccts%(1 To NumGLAccts)      'Holds all rev acct record nums
  ReDim ExpAccts%(1 To NumGLAccts)      'Holds all exp acct record nums
'  ReDim FundList$(1 To NumFunds)        'List of all active Funds
'  '--Initialize FundList
'  If NumFunds = 0 Then
'    ok = MsgBox("GL", "NOFUNDS")
'    Close
'    Exit Sub
'  End If
'  For Fund = 1 To NumFunds
'    Get FundIdxFile, Fund, FundIdx
'    FundList$(Fund) = QPTrim$(FundIdx.FundNum)
'  Next
'  Close FundIdxFile
  ReDim FundList(1) As String                           'List of all active Funds
  GetFundList FundList$(), NumFunds

  'PrintHelp "Processing:"

  '--Calculate balances thru ending date
  For GLAcct = 1 To NumGLAccts

    '--Initialize
    BGTBal# = 0
    YTDBal# = 0
    Get AcctIdxFileNum, GLAcct, AcctIdx

    Get AcctFileNum, AcctIdx.RecNum, Acct
    AcctNumber$ = QPTrim$(Acct.Num)
    Dept$ = Mid$(AcctNumber$, GLFundLen + 2, GLAcctLen)
    Det$ = Right$(AcctNumber$, GLDetLen)

'    IF INSTR(Acct.Num, "60-9840-980") > 0 THEN STOP

    '--Find what fund this account is in
    FundCode$ = Left$(Acct.Num, GLFundLen)

    '--See if the account is in a fund we want to see
    If FundCode$ >= StartFund$ And FundCode$ <= EndFund$ Then
    If InStr(Dept$, AcctCode$) And InStr(Det$, Detcode$) Then
      '--Account is in fund, check to see if its proper type
      '--We want only revenue or expenditure accounts
      If Acct.Typ = "R" Or Acct.Typ = "E" Then

        '--Print the account on screen for user to see what's going on
       ' QPrintRC Acct.Num, 25, 14, -1

        '--Assign the Account Record Number to proper list
        Select Case Acct.Typ
        Case "E"
          ECnt = ECnt + 1
          ExpAccts%(ECnt) = AcctIdx.RecNum
        Case "R"
          RCnt = RCnt + 1
          RevAccts%(RCnt) = AcctIdx.RecNum
        End Select

        '--Get account balances
        YTDBal# = Round#(Acct.BegBal)     'get the beginning balance
        NextTr& = Acct.FrstTran           'get the first trans for this acct

        Do Until NextTr& = 0              'keep going 'til we run out

          Get TransFileNum, NextTr&, GLTrans

          '--Get balance thru end date
          If GLTrans.TRDATE >= FYStartDate And GLTrans.TRDATE <= EndDate Then

            Select Case Acct.Typ
            Case "E"
              YTDBal# = YTDBal# + Round#(GLTrans.DrAmt) - Round#(GLTrans.CrAmt)
             Case "R"
              YTDBal# = YTDBal# + Round#(GLTrans.CrAmt) - Round#(GLTrans.DrAmt)
            End Select
          End If

          NextTr& = GLTrans.NextTran              'Get the next transaction

        Loop

        '--Put the new totals in the file
        Acct.YTD = Round#(YTDBal#)
        Put AcctFileNum, AcctIdx.RecNum, Acct

      End If    '--test for rev or exp accts
    End If      'dept/acct
    End If      '--End of acct in fund range test
  Next          'Process next account
If ECnt > 0 Or RCnt > 0 Then
  '--Now write the report to file.
  'PrintHelp "Generating Report..."
  FirstTimeThru = -1
  '--Check for fund in range we want
  For Fund = 1 To NumFunds
    ThisFund$ = FundList$(Fund)
    If ThisFund$ >= StartFund$ And ThisFund$ <= EndFund$ Then
      UsingFund = True
      FundRec = FindFund(ThisFund$)  'Get the fund name
      FundName$ = QPTrim$(GetFundTitle(FundRec))
      Dept$ = ""
      LastDept$ = ""
      LastDeptName$ = ""
    Else
      UsingFund = False
    End If

    If UsingFund Then
      Category$ = " Revenues"
      DoingExp = 0
      '--Advance page for each new fund except first fund
      If FirstTimeThru Then
        FirstTimeThru = 0
      Else
          Print #PRNFile, FF$
      End If
      GoSub PrintBGTWksPageHeader

      '--Search thru list of revenue account record numbers
      '--to see if we find one in this fund.
      For cnt = 1 To RCnt
        Rec = RevAccts%(cnt)
        Get AcctFileNum, Rec, Acct
        FundCode$ = Left$(Acct.Num, GLFundLen)

        '--Yup.. found one
        If FundCode$ = ThisFund$ Then

          '--Print Account detail
          ToPrint$ = Space$(132)
          Account$ = QPTrim$(Acct.Num) + "  " + QPTrim$(Acct.Title)
          LSet ToPrint$ = Left$(Account$, 36)
          Mid$(ToPrint$, Col1) = Using$(BgtFmt$, Str$(Acct.Bgt))
          Mid$(ToPrint$, Col2) = Using$(CommaFmt$, Str$(Acct.YTD))
          Mid$(ToPrint$, Col3) = Using$(CommaFmt$, Str$(Acct.PYAct))
           'IF WorkSheet = False THEN
          '  MID$(ToPrint$, Col4) = FUsing$(STR$(Acct.NYEst), CommaFmt$)
          '  MID$(ToPrint$, Col5) = FUsing$(STR$(Acct.NYReq), CommaFmt$)
          '  MID$(ToPrint$, Col6) = FUsing$(STR$(Acct.NYRec), CommaFmt$)
          '  MID$(ToPrint$, Col7) = FUsing$(STR$(Acct.NYApp), CommaFmt$)
          'ELSE
            If Acct.NYEst <> 0 Then
              Mid$(ToPrint$, Col4) = Using$(CommaFmt$, Str$(Acct.NYEst))
            Else
              Mid$(ToPrint$, Col4) = Underline$
            End If
            If Acct.NYReq <> 0 Then
              Mid$(ToPrint$, Col5) = Using$(CommaFmt$, Str$(Acct.NYReq))
            Else
              Mid$(ToPrint$, Col5) = Underline$
            End If
            If Acct.NYRec <> 0 Then
              Mid$(ToPrint$, Col6) = Using$(CommaFmt$, Str$(Acct.NYRec))
            Else
              Mid$(ToPrint$, Col6) = Underline$
            End If
              If Acct.NYApp <> 0 Then
              Mid$(ToPrint$, Col7) = Using$(CommaFmt$, Str$(Acct.NYApp))
            Else
              Mid$(ToPrint$, Col7) = Underline$
            End If
          'END IF
          Print #PRNFile, ToPrint$
          Linecnt = Linecnt + 1
          If Linecnt > MaxLines Then
            Print #PRNFile, FF$
            GoSub PrintBGTWksPageHeader
          End If

          '--Add account to running totals
          BgtSum# = BgtSum# + Acct.Bgt
          YTDSum# = YTDSum# + Acct.YTD
          PYActSum# = PYActSum# + Acct.PYAct
          NYEstSum# = NYEstSum# + Acct.NYEst
          NYReqSum# = NYReqSum# + Acct.NYReq
          NYRecSum# = NYRecSum# + Acct.NYRec
          NYAppSum# = NYAppSum# + Acct.NYApp
          End If

      Next cnt     'Revenue Acct

      '--Summarize Revenues
      GoSub PrintSummaryLines

      ToPrint$ = Space$(132)
      LSet ToPrint$ = "  Total Revenues"
      'MID$(ToPrint$, Col1) = FUsing$(STR$(BgtSum#), TotalFmt$)
      Mid$(ToPrint$, Col1) = Using$(BgtFmt$, Str$(BgtSum#))
      Mid$(ToPrint$, Col2) = Using$(TotalFmt$, Str$(YTDSum#))
      Mid$(ToPrint$, Col3) = Using$(CommaFmt$, Str$(PYActSum#))
      'IF WorkSheet = False THEN
      '  MID$(ToPrint$, Col4) = FUsing$(STR$(NYEstSum#), CommaFmt$)
      '  MID$(ToPrint$, Col5) = FUsing$(STR$(NYReqSum#), CommaFmt$)
      '  MID$(ToPrint$, Col6) = FUsing$(STR$(NYRecSum#), CommaFmt$)
      '  MID$(ToPrint$, Col7) = FUsing$(STR$(NYAppSum#), CommaFmt$)
      '
      'ELSE
        If NYEstSum# <> 0 Then
          Mid$(ToPrint$, Col4) = Using$(CommaFmt$, Str$(NYEstSum#))
        Else
          Mid$(ToPrint$, Col4) = Underline$
        End If
        If NYReqSum# <> 0 Then
          Mid$(ToPrint$, Col5) = Using$(CommaFmt$, Str$(NYReqSum#))
        Else
          Mid$(ToPrint$, Col5) = Underline$
        End If
        If NYRecSum# <> 0 Then
          Mid$(ToPrint$, Col6) = Using$(CommaFmt$, Str$(NYRecSum#))
        Else
          Mid$(ToPrint$, Col6) = Underline$
        End If
        If NYAppSum# <> 0 Then
          Mid$(ToPrint$, Col7) = Using$(CommaFmt$, Str$(NYAppSum#))
        Else
          Mid$(ToPrint$, Col7) = Underline$
        End If
      'END IF
      Print #PRNFile, ToPrint$
      Linecnt = Linecnt + 1
      If Linecnt > MaxLines Then
        Print #PRNFile, FF$
        GoSub PrintBGTWksPageHeader
      End If

      '--Assign Summary totals for revenues to New variables so that we
      '--can reuse summary vars for exp
      FundRevBgt# = BgtSum#
      FundRevYTD# = YTDSum#
      FundRevPYActSum# = PYActSum#
      FundRevNYEstSum# = NYEstSum#
      FundRevNYReqSum# = NYReqSum#
      FundRevNYRecSum# = NYRecSum#
      FundRevNYAppSum# = NYAppSum#

      BgtSum# = 0
      YTDSum# = 0
      PYActSum# = 0
      NYEstSum# = 0
      NYReqSum# = 0
      NYRecSum# = 0
      NYAppSum# = 0

      '--Now process expenditures ********************************
      DoingExp = True   '-Flag used in printing report header
      Category$ = " Expenditures"

'        If DeptOnNewPage = False Then
'          Print #PRNFile, FF$
'          NewPage = True  '-Flag used in printing report header
'          GoSub PrintBGTWksPageHeader
'        End If

      'initialize dept variables
      Break = 0
      DeptBgtSum# = 0
      DeptYTDSum# = 0
      DeptPYActSum# = 0
      DeptNYEstSum# = 0
      DeptNYReqSum# = 0
      DeptNYRecSum# = 0
      DeptNYAppSum# = 0

      DeptBgtSum1# = 0
      DeptYTDSum1# = 0
      DeptPYActSum1# = 0
      DeptNYEstSum1# = 0
      DeptNYReqSum1# = 0
      DeptNYRecSum1# = 0
      DeptNYAppSum1# = 0

      LastDept$ = ""
      LastDeptName$ = ""


      '--Search exp account list for accounts in this fund
      For cnt = 1 To ECnt
        Rec = ExpAccts%(cnt)
        Get AcctFileNum, Rec, Acct
        FundCode$ = Left$(Acct.Num, GLFundLen)
        ObjCode$ = Mid$(Acct.Num, GLFundLen% + GLAcctLen% + 3, GLDetLen)
        If FundCode$ = ThisFund$ Then
          Account$ = QPTrim$(Acct.Num) + "  " + QPTrim$(Acct.Title)

          '--Extract the Dept$ from the G/L Acct
          Dept$ = Mid$(Acct.Num, GLFundLen + 2, GLAcctLen)

          '--Get the Department name from the Department name file
          If Dept$ <> LastDept$ Then
            DeptRecNum = FindDept(Dept$)
            If DeptRecNum > 0 Then
              DeptName$ = QPTrim$(GetDeptTitle$(DeptRecNum))
            Else
              DeptName$ = "Department " + Dept$
            End If
          End If

          '--Print Department Header first time thru
          If Len(LastDeptName$) = 0 Then
            LastDeptName$ = DeptName$
            LastDept$ = Dept$
            GoSub PrintDeptHeader
             If Linecnt > MaxLines Then
              Print #PRNFile, FF$
              GoSub PrintBGTWksPageHeader
            End If
          End If

         '--Put Budget Summary Breaks Here
         If BreakFlag Then
           If Break = 0 Then
            Break = 1
            Print #PRNFile, Desc$(Break)
           End If
          If Val(ObjCode$) > BEnd(Break) Then
           GoSub PrintBreakEnds
'           'Check to See Where This One Falls
'           For LL = 1 To 9
'           If Val(ObjCode$) > Beg(LL) And Val(ObjCode$) < BEnd(LL) Then
'           Break = LL
'           Exit For
'           End If
'           Next LL
            Print #PRNFile, Desc$(Break)
          End If
         End If

          '--see if we need to subtotal dept
          If Len(LastDept$) > 0 Then
            If Dept$ <> LastDept$ Then
                  If BreakFlag Then
                        GoSub PrintBreakEnds
                        Break = 0
                   End If
              GoSub PrintDeptTotals
              GoSub PrintDeptHeader
                  If BreakFlag Then
                        Break = 1
                        Print #PRNFile, Desc$(Break)
                  End If

              If Linecnt > MaxLines Then
                Print #PRNFile, FF$
                GoSub PrintBGTWksPageHeader
                End If
            End If
          End If

          ToPrint$ = Space$(132)
          LSet ToPrint$ = Left$(Account$, 36)
          Mid$(ToPrint$, Col1) = Using$(BgtFmt$, Str$(Acct.Bgt))
          Mid$(ToPrint$, Col2) = Using$(CommaFmt$, Str$(Acct.YTD))
          Mid$(ToPrint$, Col3) = Using$(CommaFmt$, Str$(Acct.PYAct))
          'IF WorkSheet = False THEN
          '  MID$(ToPrint$, Col4) = FUsing$(STR$(Acct.NYEst), CommaFmt$)
          '  MID$(ToPrint$, Col5) = FUsing$(STR$(Acct.NYReq), CommaFmt$)
          '  MID$(ToPrint$, Col6) = FUsing$(STR$(Acct.NYRec), CommaFmt$)
          '  MID$(ToPrint$, Col7) = FUsing$(STR$(Acct.NYApp), CommaFmt$)
          'ELSE
            If Acct.NYEst <> 0 Then
              Mid$(ToPrint$, Col4) = Using$(CommaFmt$, Str$(Acct.NYEst))
            Else
              Mid$(ToPrint$, Col4) = Underline$
            End If
            If Acct.NYReq <> 0 Then
                  Mid$(ToPrint$, Col5) = Using$(CommaFmt$, Str$(Acct.NYReq))
            Else
              Mid$(ToPrint$, Col5) = Underline$
            End If
            If Acct.NYRec <> 0 Then
              Mid$(ToPrint$, Col6) = Using$(CommaFmt$, Str$(Acct.NYRec))
            Else
              Mid$(ToPrint$, Col6) = Underline$
            End If
            If Acct.NYApp <> 0 Then
              Mid$(ToPrint$, Col7) = Using$(CommaFmt$, Str$(Acct.NYApp))
            Else
              Mid$(ToPrint$, Col7) = Underline$
            End If

          'END IF
          Print #PRNFile, ToPrint$
          Linecnt = Linecnt + 1
          If Linecnt > MaxLines Then
            Print #PRNFile, FF$
            GoSub PrintBGTWksPageHeader
           End If

          BgtSum# = BgtSum# + Acct.Bgt
          YTDSum# = YTDSum# + Acct.YTD
          PYActSum# = PYActSum# + Acct.PYAct
          NYEstSum# = NYEstSum# + Acct.NYEst
          NYReqSum# = NYReqSum# + Acct.NYReq
          NYRecSum# = NYRecSum# + Acct.NYRec
          NYAppSum# = NYAppSum# + Acct.NYApp

          DeptBgtSum# = DeptBgtSum# + Acct.Bgt
          DeptYTDSum# = DeptYTDSum# + Acct.YTD
          DeptPYActSum# = DeptPYActSum# + Acct.PYAct
          DeptNYEstSum# = DeptNYEstSum# + Acct.NYEst
          DeptNYReqSum# = DeptNYReqSum# + Acct.NYReq
          DeptNYRecSum# = DeptNYRecSum# + Acct.NYRec
          DeptNYAppSum# = DeptNYAppSum# + Acct.NYApp

          DeptBgtSum1# = DeptBgtSum1# + Acct.Bgt
          DeptYTDSum1# = DeptYTDSum1# + Acct.YTD
          DeptPYActSum1# = DeptPYActSum1# + Acct.PYAct
          DeptNYEstSum1# = DeptNYEstSum1# + Acct.NYEst
          DeptNYReqSum1# = DeptNYReqSum1# + Acct.NYReq
          DeptNYRecSum1# = DeptNYRecSum1# + Acct.NYRec
          DeptNYAppSum1# = DeptNYAppSum1# + Acct.NYApp

          LastDept$ = Dept$
          LastDeptName$ = DeptName$

        End If
      Next cnt
         If BreakFlag Then
           GoSub PrintBreakEnds
         End If


      '--Summarize last Dept after loop
      GoSub PrintDeptTotals

      '--Now summarize all expenditures for fund
      GoSub PrintSummaryLines   'Print dashed line after last
      ToPrint$ = Space$(132)
      LSet ToPrint$ = "Total Expenditures for Fund:"
      Mid$(ToPrint$, Col1) = Using$(BgtFmt$, Str$(BgtSum#))
      'MID$(ToPrint$, Col1) = FUsing$(STR$(BgtSum#), TotalFmt$)
      Mid$(ToPrint$, Col2) = Using$(CommaFmt$, Str$(YTDSum#))
      Mid$(ToPrint$, Col3) = Using$(CommaFmt$, Str$(PYActSum#))
      'IF WorkSheet = False THEN
      '  MID$(ToPrint$, Col4) = FUsing$(STR$(NYEstSum#), CommaFmt$)
      '  MID$(ToPrint$, Col5) = FUsing$(STR$(NYReqSum#), CommaFmt$)
      '  MID$(ToPrint$, Col6) = FUsing$(STR$(NYRecSum#), CommaFmt$)
      '  MID$(ToPrint$, Col7) = FUsing$(STR$(NYAppSum#), CommaFmt$)
      'ELSE
            If NYEstSum# <> 0 Then
              Mid$(ToPrint$, Col4) = Using$(CommaFmt$, Str$(NYEstSum#))
            Else
              Mid$(ToPrint$, Col4) = Underline$
            End If
            If NYReqSum# <> 0 Then
              Mid$(ToPrint$, Col5) = Using$(CommaFmt$, Str$(NYReqSum#))
            Else
              Mid$(ToPrint$, Col5) = Underline$
            End If
                      If NYRecSum# <> 0 Then
              Mid$(ToPrint$, Col6) = Using$(CommaFmt$, Str$(NYRecSum#))
            Else
              Mid$(ToPrint$, Col6) = Underline$
            End If
            If NYAppSum# <> 0 Then
              Mid$(ToPrint$, Col7) = Using$(CommaFmt$, Str$(NYAppSum#))
            Else
              Mid$(ToPrint$, Col7) = Underline$
            End If

      'END IF
      Print #PRNFile, ToPrint$

      '--Summarize balances
      BGTBal# = Round#(FundRevBgt# - BgtSum#)
      YTDBal# = Round#(FundRevYTD# - YTDSum#)
      PYActBal# = FundRevPYActSum# - PYActSum#
      NYEstBal# = FundRevNYEstSum# - NYEstSum#
      NYReqBal# = FundRevNYReqSum# - NYReqSum#
      NYRecBal# = FundRevNYRecSum# - NYRecSum#
      NYAppBal# = FundRevNYAppSum# - NYAppSum#
      Print #PRNFile,

      ToPrint$ = Space$(132)
      LSet ToPrint$ = "Revenues Over/(Under) Expenditures"
      Mid$(ToPrint$, Col1) = Using$(BgtFmt$, Str$(BGTBal#))
      'MID$(ToPrint$, Col1) = FUsing$(STR$(BGTBal#), TotalFmt$)
      Mid$(ToPrint$, Col2) = Using$(CommaFmt$, Str$(YTDBal#))
      Mid$(ToPrint$, Col3) = Using$(CommaFmt$, Str$(PYActBal#))
      'IF WorkSheet = False THEN
      '  MID$(ToPrint$, Col4) = FUsing$(STR$(NYEstBal#), CommaFmt$)
      '  MID$(ToPrint$, Col5) = FUsing$(STR$(NYReqBal#), CommaFmt$)
      '  MID$(ToPrint$, Col6) = FUsing$(STR$(NYRecBal#), CommaFmt$)
      '  MID$(ToPrint$, Col7) = FUsing$(STR$(NYAppBal#), CommaFmt$)
      'ELSE
            If NYEstBal# <> 0 Then
              Mid$(ToPrint$, Col4) = Using$(CommaFmt$, Str$(NYEstBal#))
            Else
              Mid$(ToPrint$, Col4) = Underline$
            End If
            If NYReqBal# <> 0 Then
              Mid$(ToPrint$, Col5) = Using$(CommaFmt$, Str$(NYReqBal#))
            Else
              Mid$(ToPrint$, Col5) = Underline$
            End If
            If NYRecBal# <> 0 Then
              Mid$(ToPrint$, Col6) = Using$(CommaFmt$, Str$(NYRecBal#))
            Else
              Mid$(ToPrint$, Col6) = Underline$
            End If
            If NYAppBal# <> 0 Then
              Mid$(ToPrint$, Col7) = Using$(CommaFmt$, Str$(NYAppBal#))
            Else
              Mid$(ToPrint$, Col7) = Underline$
            End If

      'END IF
      Print #PRNFile, ToPrint$

      '--Reset variables for next fund
      FundRevBgt# = 0
      FundRevYTD# = 0
      FundRevPYSum# = 0
      FundRevNYEstSum# = 0
      FundRevNYReqSum# = 0
      FundRevNYRecSum# = 0
      FundRevNYAppSum# = 0
      BgtSum# = 0
      YTDSum# = 0
      PYActSum# = 0
      NYEstSum# = 0
      NYReqSum# = 0
      NYRecSum# = 0
      NYAppSum# = 0

      DeptBgtSum# = 0
      DeptYTDSum# = 0
      DeptPYSum# = 0
      DeptNYEstSum# = 0
      DeptNYReqSum# = 0
      DeptNYRecSum# = 0
      DeptNYAppSum# = 0
    End If '--Using fund test
  Next '--fund

  Print #PRNFile, FF$

  Close

  '--End Report Creation

'  Select Case Dev$
'    Case "S"
'      EntryPoint = 2
'    Case "P"
'      EntryPoint = 5
'  End Select
  'PrintRptFile Header$, ReportFile$, LPTNo, RetCode, EntryPoint
   ViewPrint ReportFile$, "Bud Prep Worksheet", True
  Close
  'KILL ReportFile$
End If
Exit Sub
PrintBGTWksPageHeader:
  Print #PRNFile, "Budget Preparation"
  Print #PRNFile, FundName$ + Category$
  Print #PRNFile, "Period Ending: " + txtDate
  Print #PRNFile,
  Print #PRNFile, "Description                         Budget       Actual      Prior Year        Est            Req            Rec          Apprv'd"
  Print #PRNFile, "-------------------------------------------------------------------------------------------------------------------------------------"
  If DoingExp Then
    If NewPage = False Then
      Print #PRNFile, DeptName$ + " Continued"
      Linecnt = 7
    Else
      Linecnt = 6
    End If
  Else
    Linecnt = 6
  End If
Return


PrintSummaryLines:
  '--Print summary lines
  ToPrint$ = Space$(132)
  Mid$(ToPrint$, Col1) = BSumLine$
  Mid$(ToPrint$, Col2 + 1) = SumLine$
  Mid$(ToPrint$, Col3 + 1) = SumLine$
  Mid$(ToPrint$, Col4 + 1) = SumLine$
  Mid$(ToPrint$, Col5 + 1) = SumLine$
  Mid$(ToPrint$, Col6 + 1) = SumLine$
  Mid$(ToPrint$, Col7 + 1) = SumLine$
  Print #PRNFile, ToPrint$
  Linecnt = Linecnt + 1
Return

PrintBreakEnds:
  GoSub PrintSummaryLines
  DeptSummary$ = " Totals " + Desc$(Break)
  ToPrint$ = Space$(132)
  LSet ToPrint$ = QPTrim$(DeptSummary$)
  Mid$(ToPrint$, Col1) = Using$(BgtFmt$, Str$(DeptBgtSum1#))
  Mid$(ToPrint$, Col2) = Using$(CommaFmt$, Str$(DeptYTDSum1#))
  Mid$(ToPrint$, Col3) = Using$(CommaFmt$, Str$(DeptPYActSum1#))
  'IF WorkSheet = False THEN
  '  MID$(ToPrint$, Col4) = FUsing$(STR$(DeptNYEstSum1#), CommaFmt$)
  '  MID$(ToPrint$, Col5) = FUsing$(STR$(DeptNYReqSum1#), CommaFmt$)
  '  MID$(ToPrint$, Col6) = FUsing$(STR$(DeptNYRecSum1#), CommaFmt$)
  '  MID$(ToPrint$, Col7) = FUsing$(STR$(DeptNYAppSum1#), CommaFmt$)
  'ELSE

  'END IF
  Print #PRNFile, ToPrint$
  Print #PRNFile,
  Linecnt = Linecnt + 2
  DeptBgtSum1# = 0
  DeptYTDSum1# = 0
  DeptPYActSum1# = 0
  DeptNYEstSum1# = 0
  DeptNYReqSum1# = 0
  DeptNYRecSum1# = 0
  DeptNYAppSum1# = 0

Return
PrintDeptTotals:
  GoSub PrintSummaryLines
  DeptSummary$ = LastDeptName$ + " Totals"
  ToPrint$ = Space$(132)
  LSet ToPrint$ = DeptSummary$
  Mid$(ToPrint$, Col1) = Using$(BgtFmt$, Str$(DeptBgtSum#))
  'MID$(ToPrint$, Col1) = FUsing$(STR$(DeptBgtSum#), TotalFmt$)
  Mid$(ToPrint$, Col2) = Using$(CommaFmt$, Str$(DeptYTDSum#))
  Mid$(ToPrint$, Col3) = Using$(CommaFmt$, Str$(DeptPYActSum#))
  'IF WorkSheet = False THEN
  '  MID$(ToPrint$, Col4) = FUsing$(STR$(DeptNYEstSum#), CommaFmt$)
  '  MID$(ToPrint$, Col5) = FUsing$(STR$(DeptNYReqSum#), CommaFmt$)
  '  MID$(ToPrint$, Col6) = FUsing$(STR$(DeptNYRecSum#), CommaFmt$)
  '  MID$(ToPrint$, Col7) = FUsing$(STR$(DeptNYAppSum#), CommaFmt$)
  'ELSE
      If DeptNYEstSum# <> 0 Then
        Mid$(ToPrint$, Col4) = Using$(CommaFmt$, Str$(DeptNYEstSum#))
      Else
        Mid$(ToPrint$, Col4) = Underline$
      End If
      If DeptNYReqSum# <> 0 Then
        Mid$(ToPrint$, Col5) = Using$(CommaFmt$, Str$(DeptNYReqSum#))
      Else
        Mid$(ToPrint$, Col5) = Underline$
      End If
      If DeptNYRecSum# <> 0 Then
        Mid$(ToPrint$, Col6) = Using$(CommaFmt$, Str$(DeptNYRecSum#))
      Else
        Mid$(ToPrint$, Col6) = Underline$
      End If
      If DeptNYAppSum# <> 0 Then
        Mid$(ToPrint$, Col7) = Using$(CommaFmt$, Str$(DeptNYAppSum#))
      Else
        Mid$(ToPrint$, Col7) = Underline$
      End If

  'END IF
  Print #PRNFile, ToPrint$
  Print #PRNFile,
  Linecnt = Linecnt + 2
  DeptBgtSum# = 0
  DeptYTDSum# = 0
  DeptPYActSum# = 0
  DeptNYEstSum# = 0
  DeptNYReqSum# = 0
  DeptNYRecSum# = 0
  DeptNYAppSum# = 0
Return


PrintDeptHeader:
  If DeptOnNewPage Then
    Print #PRNFile, FF$
    NewPage = True  '--flag not to print Dept continued
    GoSub PrintBGTWksPageHeader
    ToPrint$ = Space$(132)
    LSet ToPrint$ = DeptName$
    Print #PRNFile, ToPrint$
    Linecnt = Linecnt + 1
  Else
    '--Don't print header if not room for at least one trans
    If Linecnt < 53 Then
      NewPage = False
      ToPrint$ = Space$(132)
        LSet ToPrint$ = DeptName$
      Print #PRNFile, ToPrint$
      Linecnt = Linecnt + 1
    Else
      Print #PRNFile, FF$
      NewPage = True  '--flag not to print Dept continued
      GoSub PrintBGTWksPageHeader
      ToPrint$ = Space$(132)
      LSet ToPrint$ = DeptName$
      Print #PRNFile, ToPrint$
      Linecnt = Linecnt + 1
    End If
  End If
Return

GotErr:
'  Errorcode$ = Str$(Err)
'  Select Case Err
'    Case 70
'      Cls
'            QPrintRC "Access Denied.  Try again later.", 12, 1, 12
'      QPrintRC "Error Code:" + Errorcode$, 13, 1, 12
'      QPrintRC "Press any key exit.", 14, 1, 11
'    Case Else
'      Cls
'      QPrintRC "An Error has halted the system, Error Code: " + Errorcode$, 12, 1, 14
'      QPrintRC "Press any key exit.", 12, 1, 14
'   End Select
   'K$ = INPUT$(1)
   Exit Sub
Return

End Sub
  
Private Sub PrintWorksheetGph()
  Dim CommaFmt As String, TotalFmt As String, RunBalFmt As String
  Dim SumLine As String, BgtFmt As String, BSumLine As String, PSumLine As String
  Dim DivLine As String, DivLine2 As String, FF As String, RptTitle As String
  Dim MaxLines As Integer, Col1 As Integer, Col2 As Integer, Col3 As Integer
  Dim Col4 As Integer, Col5 As Integer, EndDate As Integer, PRNFile As Integer
  Dim M As String, HollyFlag As Boolean, Pitch12 As String, ThisFund As String
  Dim DoingDetail As Boolean, SubTotalRevenues As Boolean, DeptOnNewPage As Boolean
  Dim WhichReport As Integer, GetMonth As Boolean, GetQtr As Boolean
  Dim RptMonth As String, ReportFile As String, FundCode As String
  Dim FundIdxFile As Integer, NumFunds As Integer, GLAcct As Integer
  Dim AcctIdxFileNum As Integer, NumGLAccts As Integer, FundName As String
  Dim AcctFileNum As Integer, NumGLAcctRecs As Integer, Rec As Integer
  Dim TransFileNum As Integer, NumTrans As Long, NextTr As Long, AcctNum As String
  Dim BGTBal As Double, YTDBal As Double, MTDBal As Double, UsingFund As Boolean
  Dim ECnt As Integer, RCnt As Integer, d As String, TransMonth As String
  Dim InThisQtr As Boolean, FirstTime As Boolean, Fund As Integer, FundRec As Integer
  Dim Dept As String, LastDept As String, LastDeptName As String, cnt As Integer
  Dim Account As String, BudgetAmt As Double, DeptRecNum As Integer, DeptName As String
  Dim Pct As String, Variance As Double, ToPrint As String, Linecnt As Integer
  Dim MTDSum As Double, BgtSum As Double, YTDSum As Double, DeptBgtSum As Double
  Dim DeptYTDSum As Double, DeptENCSum As Double, DeptMTDSum As Double
  Dim FundRevMTD As Double, FundRevBgt As Double, FundRevYTD As Double
  Dim EncSum As Double, FundExpMTD As Double, FundExpBgt As Double
  Dim FundExpYTD As Double, FundEncYTD As Double, EncBal As Double
  Dim DeptSummary As String, PageNum As Integer, Newrp As String
  Dim Underline As String, CrLF As String, Col6 As Integer, Col7 As Integer
  Dim GLFile As Integer, GLSetUp1 As Integer, Category As String
  Dim DoingExp As Long, FirstTimeThru As Integer, PYActSum As Double
  Dim NYEstSum As Double, NYReqSum As Double, NYRecSum As Double
  Dim NYAppSum As Double, FundRevPYActSum As Double, FundRevNYEstSum As Double
  Dim FundRevNYReqSum As Double, FundRevNYRecSum As Double, FundRevNYAppSum As Double
  Dim NewPage  As Boolean, Break As Long, ObjCode As String, BreakFlag As Boolean
  Dim DeptPYActSum As Double, AcctNumber As String, NP As String
  Dim DeptNYEstSum As Double, Det As String, First4 As String
  Dim DeptNYReqSum As Double
  Dim DeptNYRecSum As Double
  Dim DeptNYAppSum As Double

  Dim DeptBgtSum1 As Double
  Dim DeptYTDSum1 As Double
  Dim DeptPYActSum1 As Double
  Dim DeptNYEstSum1 As Double
  Dim DeptNYReqSum1 As Double
  Dim DeptNYRecSum1 As Double
  Dim DeptNYAppSum1 As Double
  Dim PYActBal As Double
  Dim NYEstBal As Double
  Dim NYReqBal As Double
  Dim NYRecBal As Double
  Dim NYAppBal As Double
  Dim FundRevPYSum As Double
  Dim DeptPYSum As Double
  ReDim Beg(100), BEnd(100), Desc$(100)
    
    If fpcboPagebrk.Text = "Yes" Then
      DeptOnNewPage = True
      NP$ = "NP"
    Else
      NP$ = " "
    End If

  EndDate = DateDiff("d", "12/31/1979", txtDate)
  ReportFile$ = "BGTPREPwk.PRN"    'Report File Name
  CommaFmt$ = "###,###,###.##"    'format takes 14 chars
  BgtFmt$ = "#,###,###,###"         'format takes 13 chars
  TotalFmt$ = "###,###,###.##"   'format takes 14 chars
  SumLine$ = String$(20, "*")   'column summary line
  BSumLine$ = String$(10, "-")  'summary line for budget columns
  PSumLine$ = "----"            'summary line for Pct columns
  DivLine$ = String$(100, "-")   'dashed line
  DivLine2$ = String$(100, "=")  'Double Line
  Underline$ = " "
  
  CrLF$ = Chr$(13) + Chr$(10)
  FF$ = Chr$(12)
  MaxLines = 55

  '--Column offsets for printing amounts
    Col1 = 30
    Col2 = 43
    Col3 = 58
    Col4 = 73
    Col5 = 88
    Col6 = 103
    Col7 = 118

  PRNFile = FreeFile
  Open ReportFile$ For Output As #PRNFile

  OpenAcctIdx AcctIdxFileNum, NumGLAccts
  OpenAcctFile AcctFileNum, NumGLAcctRecs
  OpenTransFile TransFileNum, NumTrans&
  OpenFundIdx FundIdxFile, NumFunds

  BreakFlag = False
  ReDim RevAccts%(1 To NumGLAccts)      'Holds all rev acct record nums
  ReDim ExpAccts%(1 To NumGLAccts)      'Holds all exp acct record nums
  ReDim FundList(1) As String                           'List of all active Funds
  GetFundList FundList$(), NumFunds
  '--Calculate balances thru ending date
  For GLAcct = 1 To NumGLAccts
    '--Initialize
    BGTBal# = 0
    YTDBal# = 0
    Get AcctIdxFileNum, GLAcct, AcctIdx

    Get AcctFileNum, AcctIdx.RecNum, Acct
    AcctNumber$ = QPTrim$(Acct.Num)
    Dept$ = Mid$(AcctNumber$, GLFundLen + 2, GLAcctLen)
    Det$ = Right$(AcctNumber$, GLDetLen)

    '--Find what fund this account is in
    FundCode$ = Left$(Acct.Num, GLFundLen)

    '--See if the account is in a fund we want to see
    If FundCode$ >= StartFund$ And FundCode$ <= EndFund$ Then
    If InStr(Dept$, AcctCode$) And InStr(Det$, Detcode$) Then
      '--Account is in fund, check to see if its proper type
      '--We want only revenue or expenditure accounts
      If Acct.Typ = "R" Or Acct.Typ = "E" Then

        '--Print the account on screen for user to see what's going on
       ' QPrintRC Acct.Num, 25, 14, -1

        '--Assign the Account Record Number to proper list
        Select Case Acct.Typ
        Case "E"
          ECnt = ECnt + 1
          ExpAccts%(ECnt) = AcctIdx.RecNum
        Case "R"
          RCnt = RCnt + 1
          RevAccts%(RCnt) = AcctIdx.RecNum
        End Select

        '--Get account balances
        YTDBal# = Round#(Acct.BegBal)     'get the beginning balance
        NextTr& = Acct.FrstTran           'get the first trans for this acct

        Do Until NextTr& = 0              'keep going 'til we run out

          Get TransFileNum, NextTr&, GLTrans

          '--Get balance thru end date
          If GLTrans.TRDATE >= FYStartDate And GLTrans.TRDATE <= EndDate Then

            Select Case Acct.Typ
            Case "E"
              YTDBal# = YTDBal# + Round#(GLTrans.DrAmt) - Round#(GLTrans.CrAmt)
             Case "R"
              YTDBal# = YTDBal# + Round#(GLTrans.CrAmt) - Round#(GLTrans.DrAmt)
            End Select
          End If

          NextTr& = GLTrans.NextTran              'Get the next transaction

        Loop

        '--Put the new totals in the file
        Acct.YTD = Round#(YTDBal#)
        Put AcctFileNum, AcctIdx.RecNum, Acct

      End If    '--test for rev or exp accts
    End If      'dept/acct
    End If      '--End of acct in fund range test
  Next          'Process next account
  If ECnt > 0 Or RCnt > 0 Then
  '--Now write the report to file.
  'PrintHelp "Generating Report..."
  FirstTimeThru = -1
  '--Check for fund in range we want
  For Fund = 1 To NumFunds
    ThisFund$ = FundList$(Fund)
    If ThisFund$ >= StartFund$ And ThisFund$ <= EndFund$ Then
      UsingFund = True
      FundRec = FindFund(ThisFund$)  'Get the fund name
      FundName$ = QPTrim$(GetFundTitle(FundRec))
      Dept$ = ""
      LastDept$ = ""
      LastDeptName$ = ""
    Else
      UsingFund = False
    End If

    If UsingFund Then
      Category$ = " Revenues"
      DoingExp = 0
      '--Advance page for each new fund except first fund
'      If FirstTimeThru Then
'        FirstTimeThru = 0
'      Else
'          Print #PRNFile, FF$
'      End If
      GoSub PrintBGTWksPageHeader

      '--Search thru list of revenue account record numbers
      '--to see if we find one in this fund.
      For cnt = 1 To RCnt
        Rec = RevAccts%(cnt)
        Get AcctFileNum, Rec, Acct
        FundCode$ = Left$(Acct.Num, GLFundLen)
        AcctNumber$ = QPTrim$(Acct.Num)
        Dept$ = Mid$(AcctNumber$, GLFundLen + 2, GLAcctLen)
        Det$ = Right$(AcctNumber$, GLDetLen)
        First4$ = FundCode$ + "~" + Acct.Typ + "~" + Dept$ + "~" + Det$
        '--Yup.. found one
        If FundCode$ = ThisFund$ Then

          '--Print Account detail
          ToPrint$ = ""
          Account$ = QPTrim$(Acct.Num) + "  " + QPTrim$(Acct.Title)
          ToPrint$ = First4$ + "~" + Left$(Account$, 36)
          ToPrint$ = ToPrint$ + "~" + Using$(BgtFmt$, Str$(Acct.Bgt))
          ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, Str$(Acct.YTD))
          ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, Str$(Acct.PYAct))
           'IF WorkSheet = False THEN
          '  MID$(ToPrint$, Col4) = FUsing$(STR$(Acct.NYEst), CommaFmt$)
          '  MID$(ToPrint$, Col5) = FUsing$(STR$(Acct.NYReq), CommaFmt$)
          '  MID$(ToPrint$, Col6) = FUsing$(STR$(Acct.NYRec), CommaFmt$)
          '  MID$(ToPrint$, Col7) = FUsing$(STR$(Acct.NYApp), CommaFmt$)
          'ELSE
            If Acct.NYEst <> 0 Then
              ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, Str$(Acct.NYEst))
            Else
              ToPrint$ = ToPrint$ + "~" + Underline$
            End If
            If Acct.NYReq <> 0 Then
              ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, Str$(Acct.NYReq))
            Else
              ToPrint$ = ToPrint$ + "~" + Underline$
            End If
            If Acct.NYRec <> 0 Then
              ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, Str$(Acct.NYRec))
            Else
              ToPrint$ = ToPrint$ + "~" + Underline$
            End If
            If Acct.NYApp <> 0 Then
              ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, Str$(Acct.NYApp))
            Else
              ToPrint$ = ToPrint$ + "~" + Underline$
            End If
          'END IF
          Print #PRNFile, ToPrint$
          Linecnt = Linecnt + 1
          If Linecnt > MaxLines Then
            'Print #PRNFile, FF$
            GoSub PrintBGTWksPageHeader
          End If

          '--Add account to running totals
          BgtSum# = BgtSum# + Acct.Bgt
          YTDSum# = YTDSum# + Acct.YTD
          PYActSum# = PYActSum# + Acct.PYAct
          NYEstSum# = NYEstSum# + Acct.NYEst
          NYReqSum# = NYReqSum# + Acct.NYReq
          NYRecSum# = NYRecSum# + Acct.NYRec
          NYAppSum# = NYAppSum# + Acct.NYApp
          End If

      Next cnt     'Revenue Acct

      '--Summarize Revenues
      GoSub PrintSummaryLines

      ToPrint$ = ""
      ToPrint$ = ThisFund$ + "~" + NP$ + "~ ~ ~" + "  Total Revenues"
      'MID$(ToPrint$, Col1) = FUsing$(STR$(BgtSum#), TotalFmt$)
      ToPrint$ = ToPrint$ + "~" + Using$(BgtFmt$, Str$(BgtSum#))
      ToPrint$ = ToPrint$ + "~" + Using$(TotalFmt$, Str$(YTDSum#))
      ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, Str$(PYActSum#))
      'IF WorkSheet = False THEN
      '  MID$(ToPrint$, Col4) = FUsing$(STR$(NYEstSum#), CommaFmt$)
      '  MID$(ToPrint$, Col5) = FUsing$(STR$(NYReqSum#), CommaFmt$)
      '  MID$(ToPrint$, Col6) = FUsing$(STR$(NYRecSum#), CommaFmt$)
      '  MID$(ToPrint$, Col7) = FUsing$(STR$(NYAppSum#), CommaFmt$)
      '
      'ELSE
        If NYEstSum# <> 0 Then
          ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, Str$(NYEstSum#))
        Else
          ToPrint$ = ToPrint$ + "~" + Underline$
        End If
        If NYReqSum# <> 0 Then
          ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, Str$(NYReqSum#))
        Else
          ToPrint$ = ToPrint$ + "~" + Underline$
        End If
        If NYRecSum# <> 0 Then
          ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, Str$(NYRecSum#))
        Else
          ToPrint$ = ToPrint$ + "~" + Underline$
        End If
        If NYAppSum# <> 0 Then
          ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, Str$(NYAppSum#))
        Else
          ToPrint$ = ToPrint$ + "~" + Underline$
        End If
      'END IF
      Print #PRNFile, ToPrint$
      Linecnt = Linecnt + 1
      If Linecnt > MaxLines Then
        'Print #PRNFile, FF$
        GoSub PrintBGTWksPageHeader
      End If

      '--Assign Summary totals for revenues to New variables so that we
      '--can reuse summary vars for exp
      FundRevBgt# = BgtSum#
      FundRevYTD# = YTDSum#
      FundRevPYActSum# = PYActSum#
      FundRevNYEstSum# = NYEstSum#
      FundRevNYReqSum# = NYReqSum#
      FundRevNYRecSum# = NYRecSum#
      FundRevNYAppSum# = NYAppSum#

      BgtSum# = 0
      YTDSum# = 0
      PYActSum# = 0
      NYEstSum# = 0
      NYReqSum# = 0
      NYRecSum# = 0
      NYAppSum# = 0

      '--Now process expenditures ********************************
      DoingExp = True   '-Flag used in printing report header
      Category$ = " Expenditures"

'        If DeptOnNewPage = False Then
'          Print #PRNFile, FF$
'          NewPage = True  '-Flag used in printing report header
'          GoSub PrintBGTWksPageHeader
'        End If

      'initialize dept variables
      Break = 0
      DeptBgtSum# = 0
      DeptYTDSum# = 0
      DeptPYActSum# = 0
      DeptNYEstSum# = 0
      DeptNYReqSum# = 0
      DeptNYRecSum# = 0
      DeptNYAppSum# = 0

      DeptBgtSum1# = 0
      DeptYTDSum1# = 0
      DeptPYActSum1# = 0
      DeptNYEstSum1# = 0
      DeptNYReqSum1# = 0
      DeptNYRecSum1# = 0
      DeptNYAppSum1# = 0

      LastDept$ = ""
      LastDeptName$ = ""


      '--Search exp account list for accounts in this fund
      For cnt = 1 To ECnt
        Rec = ExpAccts%(cnt)
        Get AcctFileNum, Rec, Acct
        FundCode$ = Left$(Acct.Num, GLFundLen)
        ObjCode$ = Mid$(Acct.Num, GLFundLen% + GLAcctLen% + 3, GLDetLen)
        If FundCode$ = ThisFund$ Then
          Account$ = QPTrim$(Acct.Num) + "  " + QPTrim$(Acct.Title)

          '--Extract the Dept$ from the G/L Acct
          Dept$ = Mid$(Acct.Num, GLFundLen + 2, GLAcctLen)

          '--Get the Department name from the Department name file
          If Dept$ <> LastDept$ Then
            DeptRecNum = FindDept(Dept$)
            If DeptRecNum > 0 Then
              DeptName$ = QPTrim$(GetDeptTitle$(DeptRecNum))
            Else
              DeptName$ = "Department " + Dept$
            End If
          End If

          '--Print Department Header first time thru
          If Len(LastDeptName$) = 0 Then
            LastDeptName$ = DeptName$
            LastDept$ = Dept$
            GoSub PrintDeptHeader
             If Linecnt > MaxLines Then
              'Print #PRNFile, FF$
              GoSub PrintBGTWksPageHeader
            End If
          End If

         '--Put Budget Summary Breaks Here
         If BreakFlag Then
           If Break = 0 Then
            Break = 1
            'Print #PRNFile, Desc$(Break)
           End If
          If Val(ObjCode$) > BEnd(Break) Then
           GoSub PrintBreakEnds
'           'Check to See Where This One Falls
'           For LL = 1 To 9
'           If Val(ObjCode$) > Beg(LL) And Val(ObjCode$) < BEnd(LL) Then
'           Break = LL
'           Exit For
'           End If
'           Next LL
            'Print #PRNFile, Desc$(Break)
          End If
         End If

          '--see if we need to subtotal dept
          If Len(LastDept$) > 0 Then
            If Dept$ <> LastDept$ Then
                  If BreakFlag Then
                        GoSub PrintBreakEnds
                        Break = 0
                   End If
              GoSub PrintDeptTotals
              GoSub PrintDeptHeader
                  If BreakFlag Then
                        Break = 1
                        Print #PRNFile, Desc$(Break)
                  End If

              If Linecnt > MaxLines Then
               'Print #PRNFile, FF$
                GoSub PrintBGTWksPageHeader
                End If
            End If
          End If
        AcctNumber$ = QPTrim$(Acct.Num)
        Dept$ = Mid$(AcctNumber$, GLFundLen + 2, GLAcctLen)
        Det$ = Right$(AcctNumber$, GLDetLen)
        First4$ = FundCode$ + "~" + Acct.Typ + "~" + Dept$ + "~" + Det$
          ToPrint$ = ""
          ToPrint$ = First4$ + "~" + Left$(Account$, 36)
          ToPrint$ = ToPrint$ + "~" + Using$(BgtFmt$, Str$(Acct.Bgt))
          ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, Str$(Acct.YTD))
          ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, Str$(Acct.PYAct))
          'IF WorkSheet = False THEN
          '  MID$(ToPrint$, Col4) = FUsing$(STR$(Acct.NYEst), CommaFmt$)
          '  MID$(ToPrint$, Col5) = FUsing$(STR$(Acct.NYReq), CommaFmt$)
          '  MID$(ToPrint$, Col6) = FUsing$(STR$(Acct.NYRec), CommaFmt$)
          '  MID$(ToPrint$, Col7) = FUsing$(STR$(Acct.NYApp), CommaFmt$)
          'ELSE
            If Acct.NYEst <> 0 Then
              ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, Str$(Acct.NYEst))
            Else
              ToPrint$ = ToPrint$ + "~" + Underline$
            End If
            If Acct.NYReq <> 0 Then
              ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, Str$(Acct.NYReq))
            Else
              ToPrint$ = ToPrint$ + "~" + Underline$
            End If
            If Acct.NYRec <> 0 Then
              ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, Str$(Acct.NYRec))
            Else
              ToPrint$ = ToPrint$ + "~" + Underline$
            End If
            If Acct.NYApp <> 0 Then
              ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, Str$(Acct.NYApp))
            Else
              ToPrint$ = ToPrint$ + "~" + Underline$
            End If

          'END IF
          Print #PRNFile, ToPrint$
          Linecnt = Linecnt + 1
          If Linecnt > MaxLines Then
            'Print #PRNFile, FF$
            GoSub PrintBGTWksPageHeader
           End If

          BgtSum# = BgtSum# + Acct.Bgt
          YTDSum# = YTDSum# + Acct.YTD
          PYActSum# = PYActSum# + Acct.PYAct
          NYEstSum# = NYEstSum# + Acct.NYEst
          NYReqSum# = NYReqSum# + Acct.NYReq
          NYRecSum# = NYRecSum# + Acct.NYRec
          NYAppSum# = NYAppSum# + Acct.NYApp

          DeptBgtSum# = DeptBgtSum# + Acct.Bgt
          DeptYTDSum# = DeptYTDSum# + Acct.YTD
          DeptPYActSum# = DeptPYActSum# + Acct.PYAct
          DeptNYEstSum# = DeptNYEstSum# + Acct.NYEst
          DeptNYReqSum# = DeptNYReqSum# + Acct.NYReq
          DeptNYRecSum# = DeptNYRecSum# + Acct.NYRec
          DeptNYAppSum# = DeptNYAppSum# + Acct.NYApp

          DeptBgtSum1# = DeptBgtSum1# + Acct.Bgt
          DeptYTDSum1# = DeptYTDSum1# + Acct.YTD
          DeptPYActSum1# = DeptPYActSum1# + Acct.PYAct
          DeptNYEstSum1# = DeptNYEstSum1# + Acct.NYEst
          DeptNYReqSum1# = DeptNYReqSum1# + Acct.NYReq
          DeptNYRecSum1# = DeptNYRecSum1# + Acct.NYRec
          DeptNYAppSum1# = DeptNYAppSum1# + Acct.NYApp

          LastDept$ = Dept$
          LastDeptName$ = DeptName$

        End If
      Next cnt
         If BreakFlag Then
           GoSub PrintBreakEnds
         End If


      '--Summarize last Dept after loop
      GoSub PrintDeptTotals

      '--Now summarize all expenditures for fund
      GoSub PrintSummaryLines   'Print dashed line after last
      ToPrint$ = ""
      ToPrint$ = ThisFund$ + "~ ~ ~ ~**Total Expenditures for Fund: " + ThisFund$
      ToPrint$ = ToPrint$ + "~" + Using$(BgtFmt$, Str$(BgtSum#))
      'MID$(ToPrint$, Col1) = FUsing$(STR$(BgtSum#), TotalFmt$)
      ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, Str$(YTDSum#))
      ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, Str$(PYActSum#))
      'IF WorkSheet = False THEN
      '  MID$(ToPrint$, Col4) = FUsing$(STR$(NYEstSum#), CommaFmt$)
      '  MID$(ToPrint$, Col5) = FUsing$(STR$(NYReqSum#), CommaFmt$)
      '  MID$(ToPrint$, Col6) = FUsing$(STR$(NYRecSum#), CommaFmt$)
      '  MID$(ToPrint$, Col7) = FUsing$(STR$(NYAppSum#), CommaFmt$)
      'ELSE
            If NYEstSum# <> 0 Then
              ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, Str$(NYEstSum#))
            Else
              ToPrint$ = ToPrint$ + "~" + Underline$
            End If
            If NYReqSum# <> 0 Then
              ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, Str$(NYReqSum#))
            Else
              ToPrint$ = ToPrint$ + "~" + Underline$
            End If
            If NYRecSum# <> 0 Then
              ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, Str$(NYRecSum#))
            Else
              ToPrint$ = ToPrint$ + "~" + Underline$
            End If
            If NYAppSum# <> 0 Then
              ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, Str$(NYAppSum#))
            Else
              ToPrint$ = ToPrint$ + "~" + Underline$
            End If

      'END IF
      Print #PRNFile, ToPrint$

      '--Summarize balances
      BGTBal# = Round#(FundRevBgt# - BgtSum#)
      YTDBal# = Round#(FundRevYTD# - YTDSum#)
      PYActBal# = FundRevPYActSum# - PYActSum#
      NYEstBal# = FundRevNYEstSum# - NYEstSum#
      NYReqBal# = FundRevNYReqSum# - NYReqSum#
      NYRecBal# = FundRevNYRecSum# - NYRecSum#
      NYAppBal# = FundRevNYAppSum# - NYAppSum#
      'Print #PRNFile,

      ToPrint$ = ""
      ToPrint$ = ThisFund$ + "~ ~ ~ ~**Revenues Over/(Under) Expenditures"
      ToPrint$ = ToPrint$ + "~" + Using$(BgtFmt$, Str$(BGTBal#))
      'MID$(ToPrint$, Col1) = FUsing$(STR$(BGTBal#), TotalFmt$)
      ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, Str$(YTDBal#))
      ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, Str$(PYActBal#))
      'IF WorkSheet = False THEN
      '  MID$(ToPrint$, Col4) = FUsing$(STR$(NYEstBal#), CommaFmt$)
      '  MID$(ToPrint$, Col5) = FUsing$(STR$(NYReqBal#), CommaFmt$)
      '  MID$(ToPrint$, Col6) = FUsing$(STR$(NYRecBal#), CommaFmt$)
      '  MID$(ToPrint$, Col7) = FUsing$(STR$(NYAppBal#), CommaFmt$)
      'ELSE
            If NYEstBal# <> 0 Then
              ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, Str$(NYEstBal#))
            Else
              ToPrint$ = ToPrint$ + "~" + Underline$
            End If
            If NYReqBal# <> 0 Then
              ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, Str$(NYReqBal#))
            Else
              ToPrint$ = ToPrint$ + "~" + Underline$
            End If
            If NYRecBal# <> 0 Then
              ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, Str$(NYRecBal#))
            Else
              ToPrint$ = ToPrint$ + "~" + Underline$
            End If
            If NYAppBal# <> 0 Then
              ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, Str$(NYAppBal#))
            Else
              ToPrint$ = ToPrint$ + "~" + Underline$
            End If

      'END IF
      Print #PRNFile, ToPrint$

      '--Reset variables for next fund
      FundRevBgt# = 0
      FundRevYTD# = 0
      FundRevPYSum# = 0
      FundRevNYEstSum# = 0
      FundRevNYReqSum# = 0
      FundRevNYRecSum# = 0
      FundRevNYAppSum# = 0
      BgtSum# = 0
      YTDSum# = 0
      PYActSum# = 0
      NYEstSum# = 0
      NYReqSum# = 0
      NYRecSum# = 0
      NYAppSum# = 0

      DeptBgtSum# = 0
      DeptYTDSum# = 0
      DeptPYSum# = 0
      DeptNYEstSum# = 0
      DeptNYReqSum# = 0
      DeptNYRecSum# = 0
      DeptNYAppSum# = 0
    End If '--Using fund test
  Next '--fund

  'Print #PRNFile, FF$

  Close

  '--End Report Creation

'  Select Case Dev$
'    Case "S"
'      EntryPoint = 2
'    Case "P"
'      EntryPoint = 5
'  End Select
  'PrintRptFile Header$, ReportFile$, LPTNo, RetCode, EntryPoint
 '  ViewPrint ReportFile$, "Bud Prep Worksheet", True
  Close
  'KILL ReportFile$
'   ARptBudVAct.Label4.Caption = lab14
 '  ARptBudVAct.rptnum = NumRpt
    If DeptOnNewPage = True Then
      ARptBudgetWorksheet.deptpage = True
    Else
      ARptBudgetWorksheet.deptpage = False
    End If

   ARptBudgetWorksheet.labelEnd.Caption = ("Ending Date: " + txtDate)
   ARptBudgetWorksheet.txtDate = Now
   ARptBudgetWorksheet.txtTown = GLUserName$
   ARptBudgetWorksheet.GetName ReportFile$
   'ARptBudVAct.Visible = False
   ARptBudgetWorksheet.startrpt
End If
Exit Sub
PrintBGTWksPageHeader:
'  Print #PRNFile, "Budget Preparation"
'  Print #PRNFile, FundName$ + Category$
'  Print #PRNFile, "Period Ending: " + txtDate
'  Print #PRNFile,
'  Print #PRNFile, "Description                         Budget       Actual      Prior Year        Est            Req            Rec          Apprv'd"
'  Print #PRNFile, "-------------------------------------------------------------------------------------------------------------------------------------"
'  If DoingExp Then
'    If NewPage = False Then
'      Print #PRNFile, DeptName$ + " Continued"
'      Linecnt = 7
'    Else
'      Linecnt = 6
'    End If
'  Else
'    Linecnt = 6
'  End If
Return


PrintSummaryLines:
  '--Print summary lines
'  ToPrint$ = ""
'  ToPrint$ = ToPrint$ + "~" + BSumLine$
'  ToPrint$ = ToPrint$ + "~" + SumLine$
'  ToPrint$ = ToPrint$ + "~" + SumLine$
'  ToPrint$ = ToPrint$ + "~" + SumLine$
'  ToPrint$ = ToPrint$ + "~" + SumLine$
'  ToPrint$ = ToPrint$ + "~" + SumLine$
'  ToPrint$ = ToPrint$ + "~" + SumLine$
'  Print #PRNFile, ToPrint$
'  Linecnt = Linecnt + 1
Return

PrintBreakEnds:
    GoSub PrintSummaryLines
    DeptSummary$ = "** Totals " + Desc$(Break)
    ToPrint$ = ""
    ToPrint$ = QPTrim$(DeptSummary$)
    ToPrint$ = ToPrint$ + "~" + Using$(BgtFmt$, Str$(DeptBgtSum1#))
    ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, Str$(DeptYTDSum1#))
    ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, Str$(DeptPYActSum1#))
    'IF WorkSheet = False THEN
    '  MID$(ToPrint$, Col4) = FUsing$(STR$(DeptNYEstSum1#), CommaFmt$)
    '  MID$(ToPrint$, Col5) = FUsing$(STR$(DeptNYReqSum1#), CommaFmt$)
    '  MID$(ToPrint$, Col6) = FUsing$(STR$(DeptNYRecSum1#), CommaFmt$)
    '  MID$(ToPrint$, Col7) = FUsing$(STR$(DeptNYAppSum1#), CommaFmt$)
    'ELSE
  
    'END IF
    Print #PRNFile, ToPrint$
    'Print #PRNFile,
    Linecnt = Linecnt + 2
    DeptBgtSum1# = 0
    DeptYTDSum1# = 0
    DeptPYActSum1# = 0
    DeptNYEstSum1# = 0
    DeptNYReqSum1# = 0
    DeptNYRecSum1# = 0
    DeptNYAppSum1# = 0
  
Return
PrintDeptTotals:
  'GoSub PrintSummaryLines
  DeptSummary$ = "****" + LastDeptName$ + " Totals"
  ToPrint$ = ""
  ToPrint$ = ThisFund$ + "~" + NP$ + "~ ~ ~ " + DeptSummary$
  ToPrint$ = ToPrint$ + "~" + Using$(BgtFmt$, Str$(DeptBgtSum#))
  'MID$(ToPrint$, Col1) = FUsing$(STR$(DeptBgtSum#), TotalFmt$)
  ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, Str$(DeptYTDSum#))
  ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, Str$(DeptPYActSum#))
  'IF WorkSheet = False THEN
  '  MID$(ToPrint$, Col4) = FUsing$(STR$(DeptNYEstSum#), CommaFmt$)
  '  MID$(ToPrint$, Col5) = FUsing$(STR$(DeptNYReqSum#), CommaFmt$)
  '  MID$(ToPrint$, Col6) = FUsing$(STR$(DeptNYRecSum#), CommaFmt$)
  '  MID$(ToPrint$, Col7) = FUsing$(STR$(DeptNYAppSum#), CommaFmt$)
  'ELSE
      If DeptNYEstSum# <> 0 Then
        ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, Str$(DeptNYEstSum#))
      Else
        ToPrint$ = ToPrint$ + "~" + Underline$
      End If
      If DeptNYReqSum# <> 0 Then
        ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, Str$(DeptNYReqSum#))
      Else
        ToPrint$ = ToPrint$ + "~" + Underline$
      End If
      If DeptNYRecSum# <> 0 Then
        ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, Str$(DeptNYRecSum#))
      Else
        ToPrint$ = ToPrint$ + "~" + Underline$
      End If
      If DeptNYAppSum# <> 0 Then
        ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, Str$(DeptNYAppSum#))
      Else
        ToPrint$ = ToPrint$ + "~" + Underline$
      End If

  'END IF
  Print #PRNFile, ToPrint$
  If DeptOnNewPage = False Then
    Print #PRNFile, ThisFund$ + "~" + SumLine$ + "~" + SumLine$ + "~" + SumLine$ + "~" + SumLine$ + "~" + SumLine$ + "~" + SumLine$ + " ~" + SumLine$ + " ~" + SumLine$ + " ~" + SumLine$ + " ~" + SumLine$ + " ~" + SumLine$
    Linecnt = Linecnt + 1
  End If
  Linecnt = Linecnt + 2
  DeptBgtSum# = 0
  DeptYTDSum# = 0
  DeptPYActSum# = 0
  DeptNYEstSum# = 0
  DeptNYReqSum# = 0
  DeptNYRecSum# = 0
  DeptNYAppSum# = 0
Return


PrintDeptHeader:
'  If DeptOnNewPage Then
'    Print #PRNFile, FF$
'    NewPage = True  '--flag not to print Dept continued
'    GoSub PrintBGTWksPageHeader
'    ToPrint$ = ""
'    LSet ToPrint$ = DeptName$
'    Print #PRNFile, ToPrint$
'    Linecnt = Linecnt + 1
'  Else
'    '--Don't print header if not room for at least one trans
'    If Linecnt < 53 Then
'      NewPage = False
'      ToPrint$ = ""
'      LSet ToPrint$ = DeptName$
'      Print #PRNFile, ToPrint$
'      Linecnt = Linecnt + 1
'    Else
'      Print #PRNFile, FF$
'      NewPage = True  '--flag not to print Dept continued
'      GoSub PrintBGTWksPageHeader
'      ToPrint$ = ""
'      LSet ToPrint$ = DeptName$
'      Print #PRNFile, ToPrint$
'      Linecnt = Linecnt + 1
'    End If
'  End If
Return


GotErr:
'  Errorcode$ = Str$(Err)
'  Select Case Err
'    Case 70
'      Cls
'            QPrintRC "Access Denied.  Try again later.", 12, 1, 12
'      QPrintRC "Error Code:" + Errorcode$, 13, 1, 12
'      QPrintRC "Press any key exit.", 14, 1, 11
'    Case Else
'      Cls
'      QPrintRC "An Error has halted the system, Error Code: " + Errorcode$, 12, 1, 14
'      QPrintRC "Press any key exit.", 12, 1, 14
'   End Select
   'K$ = INPUT$(1)
   Exit Sub
Return


End Sub

