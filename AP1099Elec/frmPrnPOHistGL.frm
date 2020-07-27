VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPrnPOHistGL 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "P/O History By G/L Account"
   ClientHeight    =   8895
   ClientLeft      =   30
   ClientTop       =   495
   ClientWidth     =   12195
   Icon            =   "frmPrnPOHistGL.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8895
   ScaleWidth      =   12195
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboRptType 
      Height          =   405
      Left            =   5115
      TabIndex        =   4
      Top             =   5370
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
      ColDesigner     =   "frmPrnPOHistGL.frx":08CA
   End
   Begin LpLib.fpCombo fpcboAcct2 
      Height          =   405
      Left            =   5130
      TabIndex        =   1
      Top             =   3570
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
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
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
      ColDesigner     =   "frmPrnPOHistGL.frx":0C30
   End
   Begin LpLib.fpCombo fpcboAcct1 
      Height          =   405
      Left            =   5130
      TabIndex        =   0
      Top             =   2955
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
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
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
      ColDesigner     =   "frmPrnPOHistGL.frx":1033
   End
   Begin VB.CommandButton cmdOk 
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
      TabIndex        =   6
      Top             =   7512
      Width           =   1332
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
      Height          =   492
      Left            =   10032
      TabIndex        =   5
      Top             =   7512
      Width           =   1332
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   7
      Top             =   8532
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
            TextSave        =   "2:51 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7117
            TextSave        =   "10/19/2006"
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
   Begin EditLib.fpDateTime fpDate1 
      Height          =   372
      Left            =   5136
      TabIndex        =   2
      Top             =   4188
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
      Left            =   5136
      TabIndex        =   3
      Top             =   4788
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
      Left            =   2664
      TabIndex        =   13
      Top             =   5400
      Width           =   2388
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
      ForeColor       =   &H8000000E&
      Height          =   420
      Left            =   2976
      TabIndex        =   12
      Top             =   3000
      Width           =   2004
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   3396
      Left            =   1920
      Top             =   2616
      Width           =   8316
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "P/O History by G/L Account"
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
      TabIndex        =   11
      Top             =   1248
      Width           =   4332
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3216
      Top             =   1008
      Width           =   5772
   End
   Begin VB.Image Image1 
      Height          =   345
      Left            =   2490
      Picture         =   "frmPrnPOHistGL.frx":1436
      Top             =   2805
      Width           =   360
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
      ForeColor       =   &H8000000E&
      Height          =   420
      Left            =   3408
      TabIndex        =   10
      Top             =   4836
      Width           =   1572
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
      ForeColor       =   &H8000000E&
      Height          =   420
      Left            =   3312
      TabIndex        =   9
      Top             =   4224
      Width           =   1668
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
      ForeColor       =   &H8000000E&
      Height          =   420
      Left            =   3168
      TabIndex        =   8
      Top             =   3612
      Width           =   1812
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   972
      Left            =   3216
      Top             =   888
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
Attribute VB_Name = "frmPrnPOHistGL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim GLSetup As GLSetupRecType
Dim GLFundIdx As GLFundIndexType
Dim AcctIdx As GLAcctIndexType
Dim Vendor As VendorRecType
Dim VendorIdx As VendorIdxRecType
Dim Acct    As GLAcctRecType
Dim Trsort() As TrSortType2
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer

Private Sub cmdExit_Click()
  frmAPReportsMenu.Show
  Unload frmPrnPOHistGL
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
Private Function ValidAccts()
  If fpcboAcct1.ListIndex <> -1 And fpcboAcct2.ListIndex <> -1 Then
    fpcboAcct1.col = 1
    fpcboAcct2.col = 1
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

Private Sub cmdOk_Click()
  If ValidAccts = True Then
    If ValidDate = True Then
      If fpcboRptType.ListIndex = 0 Then
        rptopt = 1
      ElseIf fpcboRptType.ListIndex = 1 Then
        rptopt = 2
      End If
      If rptopt = 1 Then
        pohistgl
      ElseIf rptopt = 2 Then
        pohistgl2
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
        fpDate2.SetFocus
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
  Me.HelpContextID = hlpPOHistGLAcct
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  StatusBar1.Panels.Item(1).Text = GLUserName
  FillAcctstwo fpcboAcct1, fpcboAcct2
  fpcboAcct1.ListIndex = 0
  fpcboAcct2.ListIndex = fpcboAcct2.ListCount - 1
  fpDate1.Text = Format(Now, "mm/dd/yyyy")
  fpDate2.Text = fpDate1.Text
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

Private Sub fpcboAcct1_LostFocus()
  fpcboAcct1.Action = ActionClearSearchBuffer
End Sub
Private Sub fpcboAcct2_LostFocus()
  fpcboAcct2.Action = ActionClearSearchBuffer
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
Private Sub Form_Resize()
'  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
'  End If
End Sub
Private Sub pohistgl()
  Dim AcctIdxFileNum As Integer, NumGLAccts As Integer, PRNFile As Integer
  Dim AcctFileNum As Integer, NumGLAcctRecs As Integer, TransLen As Integer
  Dim ReportFile As String, TotalFmt As String, SumLine As String
  Dim NumTrans As Integer, Header As String, Desc As String
  Dim OpenDate As String, CommaFmt As String, OpenDesc As String
  Dim BegDate As Integer, EndDate As Integer, TransFile As Integer
  Dim GrTotDr As Double, GrTotCr As Double, cnt As Integer, Newrp As String
  Dim DrFwd As Double, CrFwd As Double, TotAcctDr As Double
  Dim TotAcctCr As Double, AcctNum As String, NextTr As Long
  Dim FirstAcct As String, LastAcct As String, AcctNumber As String
  Dim ToPrint As String, FwdFlag As Boolean, BalFwd As Double
  Dim Debit As String, Credit As String, Trn As Long, AcctBal As Double
  Dim lngCurLow As Long, lngCurHigh As Long, NumAcctTrans As Long
  Dim ToPrintA As String, ToPrintB As String, ToPrintT As String, ToPrintE As String
  BegDate = DateDiff("d", "12/31/1979", fpDate1)
  EndDate = DateDiff("d", "12/31/1979", fpDate2)
  fpcboAcct1.col = 1
  fpcboAcct2.col = 1
  FirstAcct$ = fpcboAcct1.ColText
  LastAcct$ = fpcboAcct2.ColText
  Newrp = "POH"
  GetRPTName Newrp
  ReportFile$ = Newrp  'Report File Name
  CommaFmt$ = "##,###,###.##"    'format takes 13 chars
  TotalFmt$ = "###,###,###.##"   'format takes 14 chars
  SumLine$ = String$(13, "-")   'column summary line
  'DivLine$ = STRING$(77, "-")   'dashed line
  'DivLine2$ = STRING$(77, "=")  'Double Line

  Header$ = "Purchase Order History"
  Desc$ = "Date       Description             Reference       Debit        Credit  Post Ref"

  OpenDate$ = fpDate1.Text
  OpenDesc$ = "Opening Balance"
  PRNFile = FreeFile
  Open ReportFile$ For Output As #PRNFile

  'PrintHelp "   Processing:"

  OpenAcctIdx AcctIdxFileNum, NumGLAccts
  OpenAcctFile AcctFileNum, NumGLAcctRecs

  Dim Trans As GLTransRecType
  TransLen = Len(Trans)
  TransFile = FreeFile
  Open "potrans.dat" For Random Access Read Write Shared As TransFile Len = TransLen
  NumTrans = LOF(TransFile) \ TransLen

  If NumTrans = 0 Then
    Close
    MsgBox "No Purchase Order Transactions To Report.", vbOKOnly, "No PO's"
    fpcboAcct1.SetFocus
    Exit Sub
  End If
  FrmShowPctComp.Label1 = "Creating Purchase Order History Report"
  FrmShowPctComp.Show , Me
  DoEvents
  DeActivateControls frmPrnPOHistGL, True
'10-20-2000
'Changed to correct subscript out of range error.
'modified TRSortType2 to use 8 bytes
'  REDIM TrSort(1 TO NumTrans)  AS TrSortType2

  GrTotDr# = 0
  GrTotCr# = 0

  For cnt = 1 To NumGLAccts
    FrmShowPctComp.ShowPctComp cnt, NumGLAccts
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ActivateControls frmPrnPOHistGL, True
      Unload FrmShowPctComp
      GoTo CancelExit
    End If

    DrFwd# = 0
    CrFwd# = 0
    TotAcctDr# = 0
    TotAcctCr# = 0

    Get AcctIdxFileNum, cnt, AcctIdx
    Get AcctFileNum, AcctIdx.RecNum, Acct

    AcctNum$ = QPTrim$(Acct.Num)
    'PrintHelp "Processing account: " + AcctNum$

    If AcctNum$ >= FirstAcct$ And AcctNum$ <= LastAcct$ Then
      If Acct.Typ = "E" Then
        NextTr = Acct.FrstPTran 'get the first trans for this acct
        AcctNumber$ = AcctNum$ + " - " + QPTrim$(Acct.Title)

        ToPrintA$ = Space$(80)
        ToPrintA$ = AcctNumber$
        'Print #PRNFile, ToPrint$

        Do Until NextTr = 0     'keep going 'til we run out of trans

          Get TransFile, NextTr, Trans

          If Trans.TRDATE >= BegDate And Trans.TRDATE <= EndDate Then
            '--within range - assign to array for sorting
            NumTrans = NumTrans + 1
            ReDim Preserve Trsort(1 To NumTrans) As TrSortType2
            Trsort(NumTrans).TRDATE = Trans.TRDATE
            Trsort(NumTrans).Record = NextTr

          Else
            '--check the transaction to see if we need to carry it in
            '  the balance fwd
            If Trans.TRDATE < BegDate Then
              DrFwd# = DrFwd# + Trans.DrAmt
              CrFwd# = CrFwd# + Trans.CrAmt
              FwdFlag = -1
            End If
          End If

          NextTr = Trans.NextTran               'Get the next transaction

        Loop

        If FwdFlag Then
          '--
          'FwdFlag = 0
          ToPrintB$ = Space$(80)

          BalFwd# = DrFwd# - CrFwd#
          If BalFwd# >= 0 Then
            Debit$ = Using$(CommaFmt$, Str$(BalFwd#))
            Credit$ = ""
          Else
            Credit$ = Using$(CommaFmt$, Str$(Abs(BalFwd#)))
            Debit$ = ""
          End If

          ToPrintB$ = Debit$ + "~" + Credit$
        Else
          ToPrintB$ = "" + "~" + ""
        End If

        If NumTrans > 0 Or BalFwd# <> 0 Then

          If NumTrans > 1 Then
            'PrintHelp "Sorting Transactions..."
            lngCurLow = LBound(Trsort)
            lngCurHigh = UBound(Trsort)
            Q2Sort Trsort(), lngCurLow, lngCurHigh
          End If

          'PrintHelp "Writing to report file..."
          For Trn = 1 To NumTrans
            Get TransFile, Trsort(Trn).Record, Trans

            ToPrintT$ = Space$(80)

            ToPrintT$ = Format(DateAdd("d", (Trans.TRDATE), "12-31-1979"), "mm/dd/yyyy") + "~"
            ToPrintT$ = ToPrintT$ + Trans.Desc + "~" + Trans.Ref + "~"

            If Trans.DrAmt <> 0 Then
              ToPrintT$ = ToPrintT$ + Using$(CommaFmt$, Str$(Trans.DrAmt)) + "~" + "" + "~"
            ElseIf Trans.CrAmt <> 0 Then
              ToPrintT$ = ToPrintT$ + "" + "~" + Using$(CommaFmt$, Str$(Trans.CrAmt)) + "~"
            Else
              ToPrintT$ = ToPrintT$ + "" + "~~"
              End If

            ToPrintT$ = ToPrintT$ + Left$(Trans.Src, 6) + "~" + "" + "~" + ""
            ToPrint$ = ToPrintA$ + "~" + ToPrintB$ + "~" + ToPrintT$
        '''$#$#$$ Write Record Here
            Print #PRNFile, ToPrint$

            TotAcctDr# = TotAcctDr# + Trans.DrAmt
            TotAcctCr# = TotAcctCr# + Trans.CrAmt
            GrTotDr# = GrTotDr# + Trans.DrAmt
            GrTotCr# = GrTotCr# + Trans.CrAmt

          Next

          '--Print summary lines
          'ToPrint$ = Space$(80)
         ' Mid$(ToPrint$, 45) = SumLine$
          'Mid$(ToPrint$, 59) = SumLine$
          'Print #PRNFile, ToPrint$

          '--Print transaction totals
'          If NumAcctTrans > 0 Then
'            ToPrint$ = Space$(80)
'            Mid$(ToPrint$, 1) = "Transaction Totals"
'            Mid$(ToPrint$, 44) = Using$(TotalFmt$, Str$(TotAcctDr#))
'            Mid$(ToPrint$, 58) = Using$(TotalFmt$, Str$(TotAcctCr#))
'            Print #PRNFile, ToPrint$
'          End If

          '--Print ending balance
          ToPrint$ = Space$(80)
          'Mid$(ToPrint$, 1) = "Encumbered Balance"
          Select Case Acct.Typ
          Case "E"
            AcctBal# = BalFwd# + TotAcctDr# - TotAcctCr#
            If AcctBal# >= 0 Then
              Debit$ = Using$(TotalFmt$, Str$(AcctBal#))
              Credit$ = ""
            Else
              Credit$ = Using$(TotalFmt$, Str$(Abs(AcctBal#)))
              Debit$ = ""
            End If

          Case "R"
            AcctBal# = BalFwd# + TotAcctCr# - TotAcctDr#
            If AcctBal# >= 0 Then
              Credit$ = Using$(TotalFmt$, Str$(AcctBal#))
              Debit$ = ""
            Else
              Debit$ = Using$(TotalFmt$, Str$(Abs(AcctBal#)))
              Credit$ = ""
            End If

          End Select

          ToPrint$ = ToPrintA$ + "~" + ToPrintB$ + "~" + "~~~~~~" + Debit$ + "~" + Credit$
          
          Print #PRNFile, ToPrint$

        Else
          ToPrint$ = Space$(80)
          ToPrint$ = ToPrintA$ + "~" + ToPrintB$ + "~~-- No Activity --~~~~~~~~"
          Print #PRNFile, ToPrint$
'
'          ToPrint$ = String$(80, "*")
'          Print #PRNFile, ToPrint$
        End If
        
      End If
    End If      'Account is not of this fund
    NumTrans = 0                'reset for next account
    BalFwd# = 0
    FwdFlag = 0
    TotAcctDr# = 0
    TotAcctCr# = 0
  Next
  If NumGLAccts < 1 Then
    FrmShowPctComp.ShowPctComp 1, 1
  End If
  'ToPrint$ = SPACE$(80)
  'LSET ToPrint$ = "Grand Total Debits"
  'MID$(ToPrint$, 25) = FUsing$(STR$(GrTotDr#), "##########,.##")
  'PRINT #PrnFile, ToPrint$

  'ToPrint$ = SPACE$(80)
  'LSET ToPrint$ = "Grand Total Credits"
  'MID$(ToPrint$, 25) = FUsing$(STR$(GrTotCr#), "##########,.##")
  'PRINT #PrnFile, ToPrint$

  Close
  Load frmLoadingRpt
  ARptPOHistGL.GetName ReportFile$
  ActivateControls frmPrnPOHistGL, True
  ARptPOHistGL.txtTown.Caption = GLUserName$
  ARptPOHistGL.txtDate.Caption = Now
  ARptPOHistGL.Label1.Caption = "P/O HISTORY BY GL ACCOUNT"
  ARptPOHistGL.startrpt


 ' KillFile ReportFile$
CancelExit:
  Exit Sub
End Sub
Private Sub pohistgl2()
  Dim AcctIdxFileNum As Integer, NumGLAccts As Integer, PRNFile As Integer
  Dim AcctFileNum As Integer, NumGLAcctRecs As Integer, TransLen As Integer
  Dim ReportFile As String, TotalFmt As String, SumLine As String
  Dim NumTrans As Integer, Header As String, Desc As String
  Dim OpenDate As String, CommaFmt As String, OpenDesc As String
  Dim BegDate As Integer, EndDate As Integer, TransFile As Integer
  Dim GrTotDr As Double, GrTotCr As Double, cnt As Integer, Newrp As String
  Dim DrFwd As Double, CrFwd As Double, TotAcctDr As Double
  Dim TotAcctCr As Double, AcctNum As String, NextTr As Long
  Dim FirstAcct As String, LastAcct As String, AcctNumber As String
  Dim ToPrint As String, FwdFlag As Boolean, BalFwd As Double
  Dim Debit As String, Credit As String, Trn As Long, AcctBal As Double
  Dim lngCurLow As Long, lngCurHigh As Long, NumAcctTrans As Long
  BegDate = DateDiff("d", "12/31/1979", fpDate1)
  EndDate = DateDiff("d", "12/31/1979", fpDate2)
  fpcboAcct1.col = 1
  fpcboAcct2.col = 1
  FirstAcct$ = fpcboAcct1.ColText
  LastAcct$ = fpcboAcct2.ColText
  Newrp = "POH"
  GetRPTName Newrp
  ReportFile$ = Newrp  'Report File Name
  CommaFmt$ = "##,###,###.##"    'format takes 13 chars
  TotalFmt$ = "###,###,###.##"   'format takes 14 chars
  SumLine$ = String$(13, "-")   'column summary line
  'DivLine$ = STRING$(77, "-")   'dashed line
  'DivLine2$ = STRING$(77, "=")  'Double Line

  Header$ = "Purchase Order History"
  Desc$ = "Date       Description             Reference       Debit        Credit  Post Ref"

  OpenDate$ = fpDate1.Text
  OpenDesc$ = "Opening Balance"
  PRNFile = FreeFile
  Open ReportFile$ For Output As #PRNFile

  'PrintHelp "   Processing:"

  OpenAcctIdx AcctIdxFileNum, NumGLAccts
  OpenAcctFile AcctFileNum, NumGLAcctRecs

  Dim Trans As GLTransRecType
  TransLen = Len(Trans)
  TransFile = FreeFile
  Open "potrans.dat" For Random Access Read Write Shared As TransFile Len = TransLen
  NumTrans = LOF(TransFile) \ TransLen

  If NumTrans = 0 Then
    Close
    MsgBox "No Purchase Order Transactions To Report.", vbOKOnly, "No PO's"
    fpcboAcct1.SetFocus
    Exit Sub
  End If
  FrmShowPctComp.Label1 = "Creating Purchase Order History Report"
  FrmShowPctComp.Show , Me
  DoEvents
  DeActivateControls frmPrnPOHistGL, True
'10-20-2000
'Changed to correct subscript out of range error.
'modified TRSortType2 to use 8 bytes
'  REDIM TrSort(1 TO NumTrans)  AS TrSortType2

  GrTotDr# = 0
  GrTotCr# = 0

  For cnt = 1 To NumGLAccts
    FrmShowPctComp.ShowPctComp cnt, NumGLAccts
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ActivateControls frmPrnPOHistGL, True
      Unload FrmShowPctComp
      GoTo CancelExit
    End If

    DrFwd# = 0
    CrFwd# = 0
    TotAcctDr# = 0
    TotAcctCr# = 0

    Get AcctIdxFileNum, cnt, AcctIdx
    Get AcctFileNum, AcctIdx.RecNum, Acct

    AcctNum$ = QPTrim$(Acct.Num)
    'PrintHelp "Processing account: " + AcctNum$

    If AcctNum$ >= FirstAcct$ And AcctNum$ <= LastAcct$ Then
      If Acct.Typ = "E" Then
        NextTr = Acct.FrstPTran 'get the first trans for this acct
        AcctNumber$ = "Account " + AcctNum$ + " - " + QPTrim$(Acct.Title)

        ToPrint$ = Space$(80)
        LSet ToPrint$ = AcctNumber$
        Print #PRNFile, ToPrint$

        Do Until NextTr = 0     'keep going 'til we run out of trans

          Get TransFile, NextTr, Trans

          If Trans.TRDATE >= BegDate And Trans.TRDATE <= EndDate Then
            '--within range - assign to array for sorting
            NumTrans = NumTrans + 1
            ReDim Preserve Trsort(1 To NumTrans) As TrSortType2
            Trsort(NumTrans).TRDATE = Trans.TRDATE
            Trsort(NumTrans).Record = NextTr

          Else
            '--check the transaction to see if we need to carry it in
            '  the balance fwd
            If Trans.TRDATE < BegDate Then
              DrFwd# = DrFwd# + Trans.DrAmt
              CrFwd# = CrFwd# + Trans.CrAmt
              FwdFlag = -1
            End If
          End If

          NextTr = Trans.NextTran               'Get the next transaction

        Loop

        If FwdFlag Then
          '--
          'FwdFlag = 0
          ToPrint$ = Space$(80)
          Mid$(ToPrint$, 1) = "Balance Forward"

          BalFwd# = DrFwd# - CrFwd#
          If BalFwd# >= 0 Then
            Debit$ = Using$(CommaFmt$, Str$(BalFwd#))
            Credit$ = ""
          Else
            Credit$ = Using$(CommaFmt$, Str$(Abs(BalFwd#)))
            Debit$ = ""
          End If

          Mid$(ToPrint$, 45) = Debit$
          Mid$(ToPrint$, 59) = Credit$
          Print #PRNFile, ToPrint$

        End If

        If NumTrans > 0 Or BalFwd# <> 0 Then

          If NumTrans > 1 Then
            'PrintHelp "Sorting Transactions..."
            lngCurLow = LBound(Trsort)
            lngCurHigh = UBound(Trsort)
            Q2Sort Trsort(), lngCurLow, lngCurHigh
          End If

          'PrintHelp "Writing to report file..."
          For Trn = 1 To NumTrans
            Get TransFile, Trsort(Trn).Record, Trans

            ToPrint$ = Space$(80)

            Mid$(ToPrint$, 1) = Format(DateAdd("d", (Trans.TRDATE), "12-31-1979"), "mm/dd/yyyy")
            Mid$(ToPrint$, 12) = Trans.Desc
            Mid$(ToPrint$, 36) = Trans.Ref

            If Trans.DrAmt <> 0 Then
              Mid$(ToPrint$, 45) = Using$(CommaFmt$, Str$(Trans.DrAmt))
            End If

            If Trans.CrAmt <> 0 Then
              Mid$(ToPrint$, 59) = Using$(CommaFmt$, Str$(Trans.CrAmt))
            End If

            Mid$(ToPrint$, 74) = Left$(Trans.Src, 6)

            Print #PRNFile, ToPrint$

            TotAcctDr# = TotAcctDr# + Trans.DrAmt
            TotAcctCr# = TotAcctCr# + Trans.CrAmt
            GrTotDr# = GrTotDr# + Trans.DrAmt
            GrTotCr# = GrTotCr# + Trans.CrAmt

          Next

          '--Print summary lines
          ToPrint$ = Space$(80)
          Mid$(ToPrint$, 45) = SumLine$
          Mid$(ToPrint$, 59) = SumLine$
          Print #PRNFile, ToPrint$

          '--Print transaction totals
          If NumAcctTrans > 0 Then
            ToPrint$ = Space$(80)
            Mid$(ToPrint$, 1) = "Transaction Totals"
            Mid$(ToPrint$, 44) = Using$(TotalFmt$, Str$(TotAcctDr#))
            Mid$(ToPrint$, 58) = Using$(TotalFmt$, Str$(TotAcctCr#))
            Print #PRNFile, ToPrint$
          End If

          '--Print ending balance
          ToPrint$ = Space$(80)
          Mid$(ToPrint$, 1) = "Encumbered Balance"
          Select Case Acct.Typ
          Case "E"
            AcctBal# = BalFwd# + TotAcctDr# - TotAcctCr#
            If AcctBal# >= 0 Then
              Debit$ = Using$(TotalFmt$, Str$(AcctBal#))
              Credit$ = ""
            Else
              Credit$ = Using$(TotalFmt$, Str$(Abs(AcctBal#)))
              Debit$ = ""
            End If

          Case "R"
            AcctBal# = BalFwd# + TotAcctCr# - TotAcctDr#
            If AcctBal# >= 0 Then
              Credit$ = Using$(TotalFmt$, Str$(AcctBal#))
              Debit$ = ""
            Else
              Debit$ = Using$(TotalFmt$, Str$(Abs(AcctBal#)))
              Credit$ = ""
            End If

          End Select

          Mid$(ToPrint$, 44) = Debit$
          Mid$(ToPrint$, 58) = Credit$
          Print #PRNFile, ToPrint$

          ToPrint$ = String$(80, "*")
          Print #PRNFile, ToPrint$
        Else
          ToPrint$ = Space$(80)
          Mid$(ToPrint$, 5) = " -- No Activity --"
          Print #PRNFile, ToPrint$

          ToPrint$ = String$(80, "*")
          Print #PRNFile, ToPrint$
        End If
      End If
    End If      'Account is not of this fund
    NumTrans = 0                'reset for next account
    BalFwd# = 0
    FwdFlag = 0
    TotAcctDr# = 0
    TotAcctCr# = 0
  Next
  If NumGLAccts < 1 Then
    FrmShowPctComp.ShowPctComp 1, 1
  End If
  'ToPrint$ = SPACE$(80)
  'LSET ToPrint$ = "Grand Total Debits"
  'MID$(ToPrint$, 25) = FUsing$(STR$(GrTotDr#), "##########,.##")
  'PRINT #PrnFile, ToPrint$

  'ToPrint$ = SPACE$(80)
  'LSET ToPrint$ = "Grand Total Credits"
  'MID$(ToPrint$, 25) = FUsing$(STR$(GrTotCr#), "##########,.##")
  'PRINT #PrnFile, ToPrint$

  Close
  ActivateControls frmPrnPOHistGL, True
  ViewPrint ReportFile$, Header$
  

  KillFile ReportFile$
CancelExit:
  Exit Sub
End Sub

Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub
