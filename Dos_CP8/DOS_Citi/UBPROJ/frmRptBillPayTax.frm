VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmRptBillPayTax 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Utility Tax Bill/Payment Report"
   ClientHeight    =   8640
   ClientLeft      =   36
   ClientTop       =   540
   ClientWidth     =   12192
   Icon            =   "frmRptBillPayTax.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   12192
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboRptType 
      Height          =   348
      Left            =   6360
      TabIndex        =   4
      Top             =   5136
      Width           =   1908
      _Version        =   196608
      _ExtentX        =   3365
      _ExtentY        =   614
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.2
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
      ColDesigner     =   "frmRptBillPayTax.frx":08CA
   End
   Begin LpLib.fpCombo fpcboRptOn 
      Height          =   348
      Left            =   6360
      TabIndex        =   2
      Top             =   4056
      Width           =   1884
      _Version        =   196608
      _ExtentX        =   3323
      _ExtentY        =   614
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.2
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
      ColDesigner     =   "frmRptBillPayTax.frx":0C68
   End
   Begin LpLib.fpCombo fpcboDetail 
      Height          =   348
      Left            =   6384
      TabIndex        =   3
      Top             =   4608
      Width           =   828
      _Version        =   196608
      _ExtentX        =   1460
      _ExtentY        =   614
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.2
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
      ThreeDOutsideStyle=   2
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   -2147483642
      BorderWidth     =   1
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
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
      ColDesigner     =   "frmRptBillPayTax.frx":0FCF
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
      Left            =   10080
      TabIndex        =   6
      Top             =   7296
      Width           =   1332
   End
   Begin VB.CommandButton cmdPrint 
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
      Left            =   8400
      TabIndex        =   5
      Top             =   7296
      Width           =   1332
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   7
      Top             =   8280
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
            TextSave        =   "3:12 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7133
            TextSave        =   "6/17/2003"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin EditLib.fpDateTime txtDate2 
      Height          =   348
      Left            =   6360
      TabIndex        =   1
      Top             =   3516
      Width           =   1692
      _Version        =   196608
      _ExtentX        =   2984
      _ExtentY        =   614
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.8
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
   Begin EditLib.fpDateTime txtDate1 
      Height          =   348
      Left            =   6360
      TabIndex        =   0
      Top             =   2976
      Width           =   1692
      _Version        =   196608
      _ExtentX        =   2984
      _ExtentY        =   614
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.8
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
      Caption         =   "Detail: "
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
      Left            =   4170
      TabIndex        =   13
      Top             =   4644
      Width           =   2004
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
      Left            =   3786
      TabIndex        =   12
      Top             =   5184
      Width           =   2388
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   3540
      Left            =   2790
      Top             =   2400
      Width           =   6612
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Beginning Date:"
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
      Left            =   4146
      TabIndex        =   11
      Top             =   3024
      Width           =   1956
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Report On:"
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
      Index           =   3
      Left            =   4410
      TabIndex        =   10
      Top             =   4104
      Width           =   1692
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
      Height          =   372
      Index           =   0
      Left            =   4530
      TabIndex        =   9
      Top             =   3564
      Width           =   1572
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3210
      Top             =   984
      Width           =   5772
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Print Tax Bill/Payment Report"
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
      Left            =   3618
      TabIndex        =   8
      Top             =   1224
      Width           =   5004
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   972
      Left            =   3210
      Top             =   864
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
Attribute VB_Name = "frmRptBillPayTax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Private Sub cmdExit_Click()
  frmUBReportsMenu.Show
  Unload frmRptBillPayTax
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
        fpcboDetail.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub
Private Sub fpcboRptOn_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboRptOn.ListDown = True
  End If
  If fpcboRptOn.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpcboDetail.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        txtDate2.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub
Private Sub fpcboDetail_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboDetail.ListDown = True
  End If
  If fpcboDetail.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpcboRptType.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpcboRptOn.SetFocus
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

Private Sub txtDate1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    txtDate2.SetFocus
  End If
End Sub

Private Sub txtDate2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpcboRptOn.SetFocus
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
        'ClearInUse PWcnt
      End If
    End If
  End If
End Sub

Private Function ValidDate()
  Dim TempDate1 As Integer, TempDate2 As Integer
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

Private Sub cmdPrint_Click()
If ValidDate Then
  DeActivateControls Me, True
  If fpcboRptType.ListIndex = 0 Then
    CustTaxReport2
  ElseIf fpcboRptType.ListIndex = 1 Then
    CustTaxReport
  End If
  ActivateControls Me, True
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

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TownName$
  txtDate1.Text = Format(Now, "mm/dd/yyyy")
  txtDate2.Text = Format(Now, "mm/dd/yyyy")
  fpcboDetail.InsertRow = "Yes"
  fpcboDetail.InsertRow = "No"
  fpcboDetail.ListIndex = 0
  fpcboRptOn.InsertRow = "Payments"
  fpcboRptOn.InsertRow = "Billing"
  fpcboRptOn.ListIndex = 0
  fpcboRptType.InsertRow = "Graphics"
  fpcboRptType.InsertRow = "Text"
  fpcboRptType.ListIndex = 0
End Sub

Private Sub Form_Resize()
'  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
'  End If
End Sub
Private Sub CustTaxReport()
  Dim PageNo As Integer, Title As String, MaxLines As Integer
  Dim Dash80 As String, IndexName As String, UBCustRecLen As Integer
  Dim UBSetupreclen As Integer, IdxRecLen As Integer, IdxFileSize As Long
  Dim IdxNumOfRecs As Long, NumCust As Long, Handle As Integer
  Dim CCnt As Long, UBCust As Integer, UBRpt As Integer, UBFile As Integer
  Dim zz As Integer, cnt As Long, TheDate As String, Trans As Long
  Dim UBTransRecLen As Integer, TempRev As String, LastRev As Integer
  Dim GotTaxFlag As Boolean, NumOfRecs As Long, UBTrans As Integer
  Dim TransType As Integer, BegDate As Integer, EndDate As Integer
  Dim CustTax As Double, RCnt As Integer, Diff As Double, Tax As Double
  Dim TTax As Double, DetailFlag As Boolean, rpt As String
  Dim ReportFile As String
  TheDate$ = Date$
  Dash80$ = String$(80, "-")
  MaxLines = 60

  ReDim RevText$(1 To MaxRevsCnt)
  ReDim TaxRates(1 To 15) As Single
  ReDim RevTotals(1 To 15) As Double
  ReDim TaxAmt(1 To 15) As Double

  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))
  
  ReDim UBTransRec(1) As UBTransRecType
  UBTransRecLen = Len(UBTransRec(1))
  ReDim UBSetUp(1) As UBSetupRecType
  LoadUBSetUpFile UBSetUp(), UBSetupreclen

  For cnt = 1 To MaxRevsCnt
    TempRev$ = QPTrim$(UBSetUp(1).Revenues(cnt).REVNAME)
    If Len(TempRev$) = 0 Then
      LastRev = cnt - 1
      Exit For
    Else
      RevText$(cnt) = TempRev$
      TaxRates(cnt) = UBSetUp(1).Revenues(cnt).TAXRATE
      If TaxRates(cnt) > 0 Then
        GotTaxFlag = True
      End If
    End If
  Next
  If Not GotTaxFlag Then
  'Put msg here !!!!!
'    QPrintRC , 10, 21, -1
  MsgBox "You do not have any taxes to report on. This report is ONLY for taxed revenues.", vbOKOnly, "No Data To Print"
'    QPrintRC "Press any key to continue.", 13, 27, -1
'    ShowCursor
'    WaitForAction
    GoTo ExitTaxReport
  End If

  FrmShowPctComp.Label1 = "Creating Bill/Payment Tax Report."
  FrmShowPctComp.Show , Me
  GoSub CheckInfo
  UBCust = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen

  UBTrans = FreeFile
  Open UBPath$ + "UBTRANS.DAT" For Random Shared As UBTrans Len = UBTransRecLen
  ReportFile$ = UBPath$ + "UBTAXRPT.RPT"
  UBRpt = FreeFile
  Open ReportFile$ For Output As UBRpt
  GoSub DoTaxRptHeader

  NumOfRecs& = LOF(UBCust) / UBCustRecLen
  For cnt& = 1 To NumOfRecs&
     FrmShowPctComp.ShowPctComp cnt, NumOfRecs
     If FrmShowPctComp.Out = True Then
       Close
       FrmShowPctComp.Out = False
       GoTo ExitTaxReport
     End If

    Get UBCust, cnt&, UBCustRec(1)
    'IF Cnt& = 2016 THEN STOP

    If Linecnt > MaxLines Then
      Print #UBRpt, FF$
      GoSub DoTaxRptHeader
    End If
    Trans& = UBCustRec(1).LastTrans
    Do While Trans& <> 0
      Get UBTrans, Trans&, UBTransRec(1)
      If UBTransRec(1).TransType = TransType Then
        If (UBTransRec(1).TransDate >= BegDate And UBTransRec(1).TransDate <= EndDate) Then
          If UBTransRec(1).TaxExempt <> "Y" Then
            CustTax# = 0
            For RCnt = 1 To LastRev
              If TaxRates(RCnt) > 0 Then
                Diff# = Round#(UBTransRec(1).RevAmt(RCnt) / (1 + TaxRates(RCnt)))
                Tax# = Round#(UBTransRec(1).RevAmt(RCnt) - Diff#)
                CustTax# = Round#(CustTax# + Tax#)
                TaxAmt(RCnt) = Round#(TaxAmt(RCnt) + Tax#)
              End If
            Next
            TTax# = Round#(TTax# + CustTax#)
            '*********************
            Print #UBRpt, Num2Date$(UBTransRec(1).TransDate); Tab(11); Using("#####", UBTransRec(1).CustAcctNo);
            If UBCustRec(1).DelFlag Then
              Print #UBRpt, Tab(24); "*";
            End If
            Print #UBRpt, Tab(25); Left$(UBCustRec(1).CustName, 33);
            Print #UBRpt, Tab(65); Using("$###,###.##", CustTax#)
            Linecnt = Linecnt + 1
            If DetailFlag Then
              For RCnt = 1 To LastRev
                If TaxRates(RCnt) > 0 Then
                  Diff# = Round#(UBTransRec(1).RevAmt(RCnt) / (1 + TaxRates(RCnt)))
                  Tax# = Round#(UBTransRec(1).RevAmt(RCnt) - Diff#)
                  Print #UBRpt, RevText$(RCnt); Tab(16); Using("#####.##", Tax#)
                  Linecnt = Linecnt + 1
                End If
              Next RCnt
              Print #UBRpt, Dash80$
              Linecnt = Linecnt + 1
            End If
            If Linecnt > MaxLines Then
              Print #UBRpt, FF$
              GoSub DoTaxRptHeader:
            End If
          End If
        End If
      End If

'      If AskAbandonPrint% Then
'        AbortFlag = True
'        Exit For
'      End If
      Trans& = UBTransRec(1).PrevTrans
    Loop
    'ShowPctComp cnt&, NumOfRecs&
SkipThisOne:
  Next

  Print #UBRpt, Dash80$
  Print #UBRpt, "Total Tax:"; Tab(65); Using("$###,###.##", TTax#)
  Print #UBRpt,
  Print #UBRpt, "Tax Breakdown:"
  For RCnt = 1 To LastRev
    If TaxRates(RCnt) > 0 Then
      Print #UBRpt, Tab(5); RevText$(RCnt); Tab(20); Using("$###,###.##", TaxAmt(RCnt))
    End If
  Next
  Print #UBRpt,
  Print #UBRpt, "Report Parameters"
  Print #UBRpt, "     Report Type: "; fpcboRptOn.Text
  Print #UBRpt, "      Start Date: "; txtDate1
  Print #UBRpt, "     Ending Date: "; txtDate2
  Print #UBRpt, "          Detail: "; fpcboDetail.Text

  Close

'  If Not AbortFlag Then
'    PrintRptFile "Transaction Tax Report.", "UBTAXRPT.RPT", 1, RetCode, EntryP
'  End If
  ViewPrint ReportFile$, "Transaction Tax Report"
ExitTaxReport:
  Exit Sub

CheckInfo:
  BegDate = Date2Num(txtDate1)
  EndDate = Date2Num(txtDate2)
  rpt$ = Left$(fpcboRptOn.Text, 1)
  If fpcboDetail.ListIndex = 0 Then
    DetailFlag = True
  Else
    DetailFlag = False
  End If
    Select Case rpt$
    Case "P"
      TransType = TranBillPayment
    Case "B"
      TransType = TranUtilityBill
    End Select

  Return

'ShowTaxParmErr:
'  SaveScrn TempScrn()
'  DisplayUBScrn "ERRSCRN1"
'  Select Case ErrCode
'  Case 1
'    QPrintRC "Invalid Start/Ending Dates!", 10, 26, -1
'  Case 2
'    QPrintRC "Invalid Report Type!", 10, 29, -1
'  End Select
'  QPrintRC "Correct and try again.", 13, 29, -1
'  WaitForAction
'  RestScrn TempScrn()
'  Action = 1
'  Return

DoTaxRptHeader:
  PageNo = PageNo + 1
  Print #UBRpt, Tab(30); "Transaction Tax Report"
  Print #UBRpt, TownName$; Tab(70); "Page #"; PageNo
  Print #UBRpt, "Report Date: "; TheDate$
  Print #UBRpt, "  Date      Acct #             Customer Name"
  Print #UBRpt, Dash80$
  Linecnt = 5
Return


End Sub
Private Sub CustTaxReport2()
  Dim Title As String, IndexName As String, UBCustRecLen As Integer
  Dim UBSetupreclen As Integer, IdxRecLen As Integer, IdxFileSize As Long
  Dim IdxNumOfRecs As Long, NumCust As Long, Handle As Integer
  Dim CCnt As Long, UBCust As Integer, UBRpt As Integer, UBFile As Integer
  Dim zz As Integer, cnt As Long, TheDate As String, Trans As Long
  Dim UBTransRecLen As Integer, TempRev As String, LastRev As Integer
  Dim GotTaxFlag As Boolean, NumOfRecs As Long, UBTrans As Integer
  Dim TransType As Integer, BegDate As Integer, EndDate As Integer
  Dim CustTax As Double, RCnt As Integer, Diff As Double, Tax As Double
  Dim TTax As Double, DetailFlag As Boolean, rpt As String
  Dim ToPrint As String, ToPrintD As String, ToPrintT As String
  Dim ToPrintH1 As String, ToPrintH2 As String, UBSub As Integer
  Dim RevTot As Integer, ReportSub As String, ReportFile As String
  TheDate$ = Date$

  ReDim RevText$(1 To MaxRevsCnt)
  ReDim TaxRates(1 To 15) As Single
  ReDim RevTotals(1 To 15) As Double
  ReDim TaxAmt(1 To 15) As Double

  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))
  
  ReDim UBTransRec(1) As UBTransRecType
  UBTransRecLen = Len(UBTransRec(1))
  ReDim UBSetUp(1) As UBSetupRecType
  LoadUBSetUpFile UBSetUp(), UBSetupreclen

  For cnt = 1 To MaxRevsCnt
    TempRev$ = QPTrim$(UBSetUp(1).Revenues(cnt).REVNAME)
    If Len(TempRev$) = 0 Then
      LastRev = cnt - 1
      Exit For
    Else
      RevText$(cnt) = TempRev$
      TaxRates(cnt) = UBSetUp(1).Revenues(cnt).TAXRATE
      If TaxRates(cnt) > 0 Then
        GotTaxFlag = True
      End If
    End If
  Next
  If Not GotTaxFlag Then
  'Put msg here !!!!!
'    QPrintRC , 10, 21, -1
  MsgBox "You do not have any taxes to report on. This report is ONLY for taxed revenues.", vbOKOnly, "No Data To Print"
'    QPrintRC "Press any key to continue.", 13, 27, -1
'    ShowCursor
'    WaitForAction
    GoTo ExitTaxReport
  End If

  FrmShowPctComp.Label1 = "Creating Bill/Payment Tax Report."
  FrmShowPctComp.Show , Me
  GoSub CheckInfo
  UBCust = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen

  UBTrans = FreeFile
  Open UBPath$ + "UBTRANS.DAT" For Random Shared As UBTrans Len = UBTransRecLen
  ReportFile$ = UBPath$ + "UBTAXRPT.RPT"
  UBRpt = FreeFile
  Open ReportFile$ For Output As UBRpt
  ReportSub$ = UBPath$ + "UBTAXSUB.RPT"
  UBSub = FreeFile
  Open ReportSub$ For Output As UBSub
  
  GoSub DoTaxRptHeader

  NumOfRecs& = LOF(UBCust) / UBCustRecLen
  For cnt& = 1 To NumOfRecs&
     FrmShowPctComp.ShowPctComp cnt, NumOfRecs
     If FrmShowPctComp.Out = True Then
       Close
       FrmShowPctComp.Out = False
       GoTo ExitTaxReport
     End If

    Get UBCust, cnt&, UBCustRec(1)
    'IF Cnt& = 2016 THEN STOP

    Trans& = UBCustRec(1).LastTrans
    Do While Trans& <> 0
      Get UBTrans, Trans&, UBTransRec(1)
      If UBTransRec(1).TransType = TransType Then
        If (UBTransRec(1).TransDate >= BegDate And UBTransRec(1).TransDate <= EndDate) Then
          If UBTransRec(1).TaxExempt <> "Y" Then
            CustTax# = 0
            For RCnt = 1 To LastRev
              If TaxRates(RCnt) > 0 Then
                Diff# = Round#(UBTransRec(1).RevAmt(RCnt) / (1 + TaxRates(RCnt)))
                Tax# = Round#(UBTransRec(1).RevAmt(RCnt) - Diff#)
                CustTax# = Round#(CustTax# + Tax#)
                TaxAmt(RCnt) = Round#(TaxAmt(RCnt) + Tax#)
              End If
            Next
            TTax# = Round#(TTax# + CustTax#)
            '*********************
            ToPrint$ = Num2Date$(UBTransRec(1).TransDate) + "~" + Using("#####", UBTransRec(1).CustAcctNo) + "~"
            If UBCustRec(1).DelFlag Then
              ToPrint$ = ToPrint$ + "*"
            End If
            ToPrint$ = ToPrint$ + Left$(UBCustRec(1).CustName, 33)
            ToPrint$ = ToPrint$ + "~" + Using("$###,###.##", CustTax#)
            If DetailFlag Then
              RevTot = 0
              For RCnt = 1 To 15 'LastRev
                If TaxRates(RCnt) > 0 Then
                  Diff# = Round#(UBTransRec(1).RevAmt(RCnt) / (1 + TaxRates(RCnt)))
                  Tax# = Round#(UBTransRec(1).RevAmt(RCnt) - Diff#)
                  ToPrintD$ = ToPrintD$ + RevText$(RCnt) + "~" + Using("#####.##", Tax#) + "~"
                  RevTot = RevTot + 1
                Else
                  ToPrintD$ = ToPrintD$ + " ~ ~ ~"
                End If
              Next RCnt
            Else
              ToPrintD$ = " ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~"
            End If
            Print #UBRpt, ToPrint$ + "~" + ToPrintD$
            ToPrint$ = ""
            ToPrintD$ = ""
          End If
        End If
      End If

'      If AskAbandonPrint% Then
'        AbortFlag = True
'        Exit For
'      End If
      Trans& = UBTransRec(1).PrevTrans
    Loop
    'ShowPctComp cnt&, NumOfRecs&
SkipThisOne:
  Next

'  Print #UBRpt, Dash80$
'  Print #UBRpt, "Total Tax:"; Tab(65); Using("$###,###.##", TTax#)
'  Print #UBRpt,
'  Print #UBRpt, "Tax Breakdown:"
  For RCnt = 1 To LastRev
    If TaxRates(RCnt) > 0 Then
      Print #UBSub, RevText$(RCnt) + "~" + Using("$###,###.##", TaxAmt(RCnt))
    End If
  Next
'  Print #UBRpt,
'  Print #UBRpt, "Report Parameters"
   ToPrintH1$ = "         Start Date: " + txtDate1 + "     Report Type: " + fpcboRptOn.Text
   ToPrintH2$ = "     Ending Date: " + txtDate2 + "                Detail: " + fpcboDetail.Text
'  Print #UBRpt,
'  Print #UBRpt,

  Close

'  ViewPrint "UBTAXRPT.RPT", "Transaction Tax Report"
    Load frmLoadingRpt
    ARptBillPayTax.txtDate = Now
    ARptBillPayTax.txtTown = TownName$
    ARptBillPayTax.Title = "Transaction Tax Report"
    ARptBillPayTax.txtRptParm1.Caption = ToPrintH1$
    ARptBillPayTax.txtRptParm2.Caption = ToPrintH2$
    ARptBillPayTax.GetName ReportFile$, ReportSub$, DetailFlag, RevTot
    ARptBillPayTax.startrpt

ExitTaxReport:
  Exit Sub

CheckInfo:
  BegDate = Date2Num(txtDate1)
  EndDate = Date2Num(txtDate2)
  rpt$ = Left$(fpcboRptOn.Text, 1)
  If fpcboDetail.ListIndex = 0 Then
    DetailFlag = True
  Else
    DetailFlag = False
  End If
    Select Case rpt$
    Case "P"
      TransType = TranBillPayment
    Case "B"
      TransType = TranUtilityBill
    End Select

  Return

'ShowTaxParmErr:
'  SaveScrn TempScrn()
'  DisplayUBScrn "ERRSCRN1"
'  Select Case ErrCode
'  Case 1
'    QPrintRC "Invalid Start/Ending Dates!", 10, 26, -1
'  Case 2
'    QPrintRC "Invalid Report Type!", 10, 29, -1
'  End Select
'  QPrintRC "Correct and try again.", 13, 29, -1
'  WaitForAction
'  RestScrn TempScrn()
'  Action = 1
'  Return

DoTaxRptHeader:
'  PageNo = PageNo + 1
'  Print #UBRpt, Tab(30); "Transaction Tax Report"
'  Print #UBRpt, TownName$; Tab(70); "Page #"; PageNo
'  Print #UBRpt, "Report Date: "; TheDate$
'  Print #UBRpt, "  Date      Acct #             Customer Name"
'  Print #UBRpt, Dash80$
'  Linecnt = 5
Return


End Sub

