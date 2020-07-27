VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmRptFinalApply 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Credit-Applied/Final Report"
   ClientHeight    =   8640
   ClientLeft      =   30
   ClientTop       =   540
   ClientWidth     =   12195
   Icon            =   "frmRptFinalApply.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   12195
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboRptType 
      Height          =   375
      Left            =   5370
      TabIndex        =   6
      Top             =   6630
      Width           =   3615
      _Version        =   196608
      _ExtentX        =   6376
      _ExtentY        =   661
      Text            =   ""
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
      ColDesigner     =   "frmRptFinalApply.frx":08CA
   End
   Begin LpLib.fpCombo fpcboDetail 
      Height          =   375
      Left            =   5370
      TabIndex        =   4
      Top             =   5640
      Width           =   840
      _Version        =   196608
      _ExtentX        =   1482
      _ExtentY        =   661
      Text            =   ""
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
      ColDesigner     =   "frmRptFinalApply.frx":0BF8
   End
   Begin LpLib.fpCombo fpcboPrintOrder 
      Height          =   375
      Left            =   5370
      TabIndex        =   5
      Top             =   6135
      Width           =   3615
      _Version        =   196608
      _ExtentX        =   6376
      _ExtentY        =   661
      Text            =   ""
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
      EditMarginLeft  =   2
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmRptFinalApply.frx":0F26
   End
   Begin VB.CheckBox CheckRefOnly 
      BackColor       =   &H00988E74&
      Caption         =   "Refunds Only"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7350
      MaskColor       =   &H00C0FFFF&
      TabIndex        =   22
      Top             =   5640
      Width           =   1635
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
      Left            =   9330
      TabIndex        =   8
      Top             =   7350
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
      Left            =   7680
      TabIndex        =   7
      Top             =   7350
      Width           =   1332
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
            TextSave        =   "1:18 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7117
            TextSave        =   "4/5/2010"
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
   Begin EditLib.fpText fptxtCustType 
      Height          =   345
      Left            =   5370
      TabIndex        =   3
      Top             =   5145
      Width           =   1185
      _Version        =   196608
      _ExtentX        =   2096
      _ExtentY        =   614
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   2
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
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      AutoCase        =   1
      CaretInsert     =   2
      CaretOverWrite  =   2
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
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   3
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fptxtOperator 
      Height          =   345
      Left            =   5370
      TabIndex        =   2
      Top             =   4665
      Width           =   810
      _Version        =   196608
      _ExtentX        =   1418
      _ExtentY        =   614
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   2
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
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      AutoCase        =   0
      CaretInsert     =   2
      CaretOverWrite  =   2
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
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   ""
      CharValidationText=   "0123456789"
      MaxLength       =   4
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpDateTime txtDate2 
      Height          =   345
      Left            =   5355
      TabIndex        =   1
      Top             =   4170
      Width           =   1695
      _Version        =   196608
      _ExtentX        =   2984
      _ExtentY        =   614
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
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
      Height          =   345
      Left            =   5355
      TabIndex        =   0
      Top             =   3690
      Width           =   1695
      _Version        =   196608
      _ExtentX        =   2984
      _ExtentY        =   614
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
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
   Begin VB.Line Line1 
      X1              =   1440
      X2              =   10950
      Y1              =   3420
      Y2              =   3420
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "The customers most recent transaction is an applied deposit from a final billing."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   675
      Left            =   3060
      TabIndex        =   21
      Top             =   1950
      Width           =   7335
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "The customer/transaction falls within the criteria below."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   675
      Left            =   3120
      TabIndex        =   20
      Top             =   2670
      Width           =   5415
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "IF the following applies:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   405
      Left            =   2130
      TabIndex        =   19
      Top             =   1530
      Width           =   5475
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "**This report will generate a list of customers with an applied deposit."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   405
      Left            =   1800
      TabIndex        =   18
      Top             =   1140
      Width           =   9555
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
      Height          =   375
      Index           =   0
      Left            =   3690
      TabIndex        =   17
      Top             =   4230
      Width           =   1575
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
      Height          =   375
      Left            =   3600
      TabIndex        =   16
      Top             =   3735
      Width           =   1665
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Type:"
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
      Height          =   375
      Index           =   1
      Left            =   3210
      TabIndex        =   15
      Top             =   5205
      Width           =   2070
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Operator No:"
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
      Height          =   375
      Index           =   2
      Left            =   3210
      TabIndex        =   14
      Top             =   4725
      Width           =   2070
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
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
      Height          =   375
      Index           =   7
      Left            =   3570
      TabIndex        =   13
      Top             =   6195
      Width           =   1710
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   6390
      Left            =   1440
      Top             =   810
      Width           =   9570
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Show Detail: "
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
      Height          =   375
      Left            =   3330
      TabIndex        =   12
      Top             =   5700
      Width           =   2010
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
      Height          =   375
      Left            =   2955
      TabIndex        =   11
      Top             =   6675
      Width           =   2385
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   705
      Left            =   2775
      Top             =   105
      Width           =   6645
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Print Final Cust w/Credits"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2475
      TabIndex        =   10
      Top             =   300
      Width           =   7230
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   825
      Left            =   2775
      Top             =   0
      Width           =   6645
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
Attribute VB_Name = "frmRptFinalApply"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim BegRoute As String, EndRoute As String
Dim UseCycle As Boolean
Private Sub cmdExit_Click()
  
  Unload Me
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        UBLog "Closed via RptFinalApply by " + PWUser$
        CitiTerminate
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

Private Sub txtDate1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    txtDate2.SetFocus
  End If
End Sub

Private Sub txtDate2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fptxtOperator.SetFocus
  End If
End Sub

Private Sub fptxtOperator_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fptxtCustType.SetFocus
  End If
End Sub

Private Sub fptxtCustType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpcboDetail.SetFocus
  End If
End Sub
Private Sub fpcboDetail_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboDetail.ListDown = True
  End If
  If fpcboDetail.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpcboPrintOrder.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fptxtCustType.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub

Private Sub fpcboPrintOrder_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboPrintOrder.ListDown = True
  End If
  If fpcboPrintOrder.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpcboRptType.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpcboDetail.SetFocus
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
        fpcboPrintOrder.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub


Private Sub cmdPrint_Click()
  If ValidDate Then
    DeActivateControls Me, True
    If fpcboRptType.ListIndex = 1 Then
      DetailedTransJournalXX
      ActivateControls Me, True
    Else
      DetailedTransJournalXX
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
'  Dim UBSetupreclen As Integer
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
  fptxtOperator = ""
  fptxtCustType = ""
  fpcboDetail.AddItem "No"
  fpcboDetail.AddItem "Yes"
  fpcboDetail.ListIndex = 0
  fpcboPrintOrder.AddItem "Customer Name Order"
  fpcboPrintOrder.AddItem "Account Number Order"
  fpcboPrintOrder.AddItem "Service Address Order"
  fpcboPrintOrder.ListIndex = 0
  txtDate1.Text = Format(Now, "mm/dd/yyyy")
  txtDate2.Text = Format(Now, "mm/dd/yyyy")

 ' fpcboTransType.AddItem " 1) - Utility Bill"
 ' fpcboTransType.AddItem " 5) - Applied Deposit"
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
Private Sub DetailedTransJournal()
  Dim UBCustRecLen As Integer, UBSetupreclen As Integer
  Dim UsingBook As Boolean, UsingName As Boolean, UsingAddr As Boolean
  Dim CustomerCnt As Long, UBTransRecLen As Integer, UBTrans As Integer
  Dim IndexName As String, Handle As Integer, Dash120 As String
  Dim IdxRecLen As Integer, IdxFileSize As Long, UBRpt As Integer
  Dim IdxNumOfRecs As Long, NumOfRecs As Long, MaxRevenue As Integer
  Dim cnt As Long, UBCust As Integer, RCnt As Integer, UseType As Boolean
  Dim ThisType As String, CUSTTYPE As String, CustBook As Integer
  Dim FromBook As Integer, ThruBook As Integer, BadCount As Long
  Dim Trans As Long, UBTransLen As Integer, BegDate As Integer
  Dim EndDate As Integer, BegOperator As Integer, EndOperator As Integer
  Dim BegTrans As Integer, EndTrans As Integer, TransDesc As String
  Dim Amount As Double, TotalTrans As Double, TransCnt As Long
  Dim Detail As String, Date1 As String, Date2 As String, Operator As String
  Dim TotalRevsAmt As Double, EstCnt As Integer, TrType As String
  Dim TrTyp As Integer, OperatorNo As String, UsingAcct As Boolean
  Dim ReportFile As String, MoFlag As Boolean, nexttr As Long
 'get report parameters
  GoSub CheckDetailParms
  MaxLines = 45
  PageNo = 0
  'MaxRevenue = 15
  Dash120$ = String$(80, "-")
  FrmShowPctComp.Label1 = "Creating Transactions"
  FrmShowPctComp.Show , Me
  DoEvents
  ''DeActivateControls Me, True
  ReDim RevTotals(1 To 15) As Double
  ReDim RevenueName(1 To 15) As String
  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))

  ReDim UBSetUp(1) As UBSetupRecType
  UBSetupreclen = Len(UBSetUp(1))
  ReDim UBTransRec(1) As UBTransRecType
  UBTransRecLen = Len(UBTransRec(1))
  Date1$ = txtDate1
  Date2$ = txtDate2

  BegDate = Date2Num%(Date1$)
  EndDate = Date2Num%(Date2$)

  If UsingName Or UsingBook Then
    IdxRecLen = 4               'we are using a long integer
    IdxFileSize& = FileSize(IndexName$)
    IdxNumOfRecs = IdxFileSize& \ IdxRecLen
    ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
    'FGetAH IndexName$, IdxBuff(1), IdxRecLen, IdxNumOfRecs      'load it
    NumOfRecs = IdxNumOfRecs
    Handle = FreeFile
    Open IndexName$ For Random Shared As Handle Len = IdxRecLen
    For cnt& = 1 To IdxNumOfRecs
      Get #Handle, cnt&, IdxBuff(cnt&)
    Next
    Close Handle
  ElseIf UsingAddr Then
'unrem
    SortServiceAddrs frmRptMastCust
    IdxRecLen = 4               'we are using a long integer
    IdxFileSize& = FileSize&(IndexName$)
    IdxNumOfRecs = IdxFileSize& \ IdxRecLen
    ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
    'FGetAH IndexName$, IdxBuff(1), IdxRecLen, IdxNumOfRecs      'load it
    NumOfRecs = IdxNumOfRecs
    Handle = FreeFile
    Open IndexName$ For Random Shared As Handle Len = IdxRecLen
    For cnt& = 1 To IdxNumOfRecs
      Get #Handle, cnt&, IdxBuff(cnt&)
    Next
    Close Handle

  Else
    NumOfRecs = FileSize(UBPath$ + "UBCUST.DAT") \ UBCustRecLen
  End If

  UBCust = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen
  UBTrans = FreeFile
  Open UBPath$ + "UBTRANS.DAT" For Random Shared As UBTrans Len = UBTransRecLen
  ReportFile$ = UBPath$ + "UBCRLIST.RPT"
  UBRpt = FreeFile
  Open ReportFile$ For Output As UBRpt
'  ubsetup = FreeFile
'  Open "UBSETUP.DAT" For Random Shared As ubsetup Len = UBSetupreclen
  LoadUBSetUpFile UBSetUp(), UBSetupreclen

  If Len(TOWNNAME$) = 0 Then
    TOWNNAME$ = "Undefined"
    ' Set Revenue Names to Nothing
    For RCnt = 1 To 15
      RevenueName$(RCnt) = "Not Set"
    Next RCnt
  Else
    'Get ubsetup, 1, ubsetup(1)
    For RCnt = 1 To 15
      RevenueName$(RCnt) = QPTrim$(UBSetUp(1).Revenues(RCnt).RevName)
    Next RCnt
    RCnt = 1
    Do While RCnt <= 15
      If RevenueName$(RCnt) = "" Then
        MaxRevenue = RCnt - 1
        Exit Do
      End If
      RCnt = RCnt + 1
    Loop
  End If

  GoSub DoDetailedRptHeader

  For cnt = 1 To NumOfRecs
    FrmShowPctComp.ShowPctComp cnt, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ''ActivateControls Me, True
      GoTo ExitDetailedListing
    End If

    If UsingName Or UsingAddr Then
      Get UBCust, IdxBuff(cnt).RecNum, UBCustRec(1)
    Else
      Get UBCust, cnt, UBCustRec(1)
    End If

    If UBCustRec(1).DelFlag <> 0 Or UBCustRec(1).Status = "A" Then
      GoTo SkipThisOne
    End If

    If UseType Then
      ThisType$ = QPTrim$(UBCustRec(1).CUSTTYPE)
      If ThisType$ <> CUSTTYPE$ Then
        GoTo SkipThisOne
      End If
    End If

    If LineCnt > MaxLines Then
      Print #UBRpt, FF$
      GoSub DoDetailedRptHeader
    End If
'*************************************
'   Main Body of Printing goes here
    BadCount = 0
    Trans& = UBCustRec(1).LastTrans
   ' Do While Trans& <> 0
   If Trans& > 0 Then
      Get UBTrans, Trans&, UBTransRec(1)
                'TransDesc$ = "Applied Dep"
                'Amount# = Abs(UBTransRec(1).Transamt)
      If (UBTransRec(1).TransDate >= BegDate And UBTransRec(1).TransDate <= EndDate) Then
        If (UBTransRec(1).OperatorNumber >= BegOperator And UBTransRec(1).OperatorNumber <= EndOperator) Then
          If UBTransRec(1).TransType = 5 Or UBTransRec(1).TransType = 105 Then
             If UBTransRec(1).PrevTrans > 0 Then
              Trans& = UBTransRec(1).PrevTrans
              Get UBTrans, Trans&, UBTransRec(1)
                If UBTransRec(1).TransType = 1 Or UBTransRec(1).TransType = 101 And UBTransRec(1).ApplyDepFlag = "Y" Then
                  If Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance) < 0 Then
                    TransDesc$ = "Credit Balance"
                    Amount# = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
                    Print #UBRpt, Num2Date$(UBTransRec(1).TransDate); Tab(13); Using("#####", UBTransRec(1).CustAcctNo);
                    
                    Print #UBRpt, Tab(25); QPTrim$(UBCustRec(1).CustName);
                    Print #UBRpt, Tab(70); Using("$###,###.##", Amount#)
                    LineCnt = LineCnt + 1
                    Print #UBRpt, Tab(25); QPTrim$(UBCustRec(1).ADDR1)
                    LineCnt = LineCnt + 1
                    If Len(QPTrim$(UBCustRec(1).ADDR2)) > 0 Then
                      Print #UBRpt, Tab(25); QPTrim$(UBCustRec(1).ADDR2)
                      LineCnt = LineCnt + 1
                    End If
                    Print #UBRpt, Tab(25); QPTrim$(UBCustRec(1).CITY); " "; QPTrim$(UBCustRec(1).STATE); " "; QPTrim$(UBCustRec(1).ZIPCODE)
                    LineCnt = LineCnt + 1
                    If fpcboDetail.ListIndex = 1 Then
                      Print #UBRpt, "Final Billing Breakdown ........................"
                      LineCnt = LineCnt + 1
                      For RCnt = 1 To MaxRevenue Step 3
                        Print #UBRpt, RevenueName$(RCnt); Tab(16); Using("#####.##", Round#(UBTransRec(1).RevAmt(RCnt) + UBTransRec(1).TaxAmt(RCnt)));
                        Print #UBRpt, Tab(30); RevenueName$(RCnt + 1); Tab(46); Using("#####.##", Round#(UBTransRec(1).RevAmt(RCnt + 1) + UBTransRec(1).TaxAmt(RCnt + 1)));
                        Print #UBRpt, Tab(58); RevenueName$(RCnt + 2); Tab(70); Using("#####.##", Round#(UBTransRec(1).RevAmt(RCnt + 2) + UBTransRec(1).TaxAmt(RCnt + 2)))
                        LineCnt = LineCnt + 1
                      Next RCnt
                     
                      Print #UBRpt, Dash120$
                      LineCnt = LineCnt + 1
                      
                      'get apply dep trans again
                      Trans& = UBCustRec(1).LastTrans
                      If Trans& > 0 Then Get UBTrans, Trans&, UBTransRec(1)

                      Print #UBRpt, "Apply Deposit Breakdown ........................"
                      LineCnt = LineCnt + 1
                      For RCnt = 1 To MaxRevenue Step 3
                        Print #UBRpt, RevenueName$(RCnt); Tab(16); Using("#####.##", Round#(UBTransRec(1).RevAmt(RCnt) + UBTransRec(1).TaxAmt(RCnt)));
                        Print #UBRpt, Tab(30); RevenueName$(RCnt + 1); Tab(46); Using("#####.##", Round#(UBTransRec(1).RevAmt(RCnt + 1) + UBTransRec(1).TaxAmt(RCnt + 1)));
                        Print #UBRpt, Tab(58); RevenueName$(RCnt + 2); Tab(70); Using("#####.##", Round#(UBTransRec(1).RevAmt(RCnt + 2) + UBTransRec(1).TaxAmt(RCnt + 2)))
                        LineCnt = LineCnt + 1
                      Next RCnt
                     
                      Print #UBRpt, Dash120$
                      LineCnt = LineCnt + 1


                      Print #UBRpt, "Revenue Balances after Apply Deposit ........................"
                      LineCnt = LineCnt + 1
                      For RCnt = 1 To MaxRevenue Step 3
                        Print #UBRpt, RevenueName$(RCnt); Tab(16); Using("#####.##", Round#(UBCustRec(1).CurrRevAmts(RCnt)));
                        Print #UBRpt, Tab(30); RevenueName$(RCnt + 1); Tab(46); Using("#####.##", Round#(UBCustRec(1).CurrRevAmts(RCnt + 1)));
                        Print #UBRpt, Tab(58); RevenueName$(RCnt + 2); Tab(70); Using("#####.##", Round#(UBCustRec(1).CurrRevAmts(RCnt + 2)))
                        LineCnt = LineCnt + 1
                      Next RCnt
                      Print #UBRpt, Dash120$
                      LineCnt = LineCnt + 1
                      For RCnt = 1 To MaxRevenue
                        RevTotals(RCnt) = Round#(RevTotals(RCnt) + (UBCustRec(1).CurrRevAmts(RCnt)))
                      Next
                    Else
                      Print #UBRpt, Dash120$
                      LineCnt = LineCnt + 1
                    End If
                    TotalTrans# = Round#(TotalTrans# + Amount#)
                    TransCnt& = TransCnt& + 1
                    If LineCnt > MaxLines Then
                      Print #UBRpt, FF$
                      GoSub DoDetailedRptHeader
                    End If
                 'Exit Do
                 End If
               End If
          End If
        End If
      End If
     End If
   End If '   Trans& = UBTransRec(1).PrevTrans
   ' Loop
SkipThisOne:
'    ShowPctComp cnt, NumOfRecs
  Next

  GoSub DoDetailedRptFooter
  If fpcboRptType.ListIndex = 1 Then Print #UBRpt, FF$

  Close

  Erase IdxBuff, UBCustRec
 ''' ActivateControls Me, True
  'END
 If fpcboRptType.ListIndex <> 1 Then
    Load frmLoadingRpt
    frmLoadingRpt.setwherefrom frmRptFinalApply
    ARptLineRpt.GetName ReportFile$
    ARptLineRpt.startrpt
 Else
  'If Not AbortFlag Then
  '  PrintRptFile "Detailed Journal Report.", "UBDJLIST.RPT", LptPort, RetCode, EntryPoint
 ' End If
    ViewPrint ReportFile$, "Credit-Applied/Final Report", True
  'KillFile "UBDJLIST.RPT"
  End If
ExitDetailedListing:

  Exit Sub

DoDetailedRptHeader:
  PageNo = PageNo + 1
  Print #UBRpt,
  Print #UBRpt,
  Print #UBRpt, TOWNNAME$
  Print #UBRpt, "Credit-Applied/Final Report"; Tab(70); "Page #"; PageNo
    Print #UBRpt, "               Report Date: "; Date$
    Print #UBRpt, "Beginning Transaction Date: "; Date1$
    Print #UBRpt, "   Ending Transaction Date: "; Date2$
  If Val(Operator$) = 0 Then
    Print #UBRpt, "                Operator #: ALL"
  Else
    Print #UBRpt, "                Operator #: "; Operator$
  End If
    Print #UBRpt, "               Show Detail: "; Detail$
    Print #UBRpt, "             Customer Type: ";

  If UseType Then
    Print #UBRpt, CUSTTYPE$
  Else
    Print #UBRpt, "N/A"
  End If

  Print #UBRpt,
  Print #UBRpt, "  Date"; Tab(15); "Acct #"; Tab(30); "Customer Name/Address"; Tab(65); "Credit Balance"
  Print #UBRpt, Dash120$
  LineCnt = 13
  Return

DoDetailedRptFooter:

  Print #UBRpt, Dash120$
  Print #UBRpt, "Transactions: "; TransCnt&; "                       Total of Credit Balances: "; Using("$##,###,###.##", TotalTrans#)
  'If fpcboRptType.ListIndex = 1 Then
  If fpcboDetail.ListIndex = 1 Then
    Print #UBRpt, FF$
    PageNo = PageNo + 1
    Print #UBRpt,
    Print #UBRpt,
    Print #UBRpt, TOWNNAME$
    Print #UBRpt, "Credit-Applied/Final Report"; Tab(70); "Page #"; PageNo
      Print #UBRpt, "               Report Date: "; Date$
      Print #UBRpt, "Beginning Transaction Date: "; Date1$
      Print #UBRpt, "   Ending Transaction Date: "; Date2$
    If Val(Operator$) = 0 Then
      Print #UBRpt, "                Operator #: ALL"
    Else
      Print #UBRpt, "                Operator #: "; Operator$
    End If
      Print #UBRpt, "               Show Detail: "; Detail$
      Print #UBRpt, "             Customer Type: ";
  
    If UseType Then
      Print #UBRpt, CUSTTYPE$
    Else
      Print #UBRpt, "N/A"
    End If
    Print #UBRpt, ""
    Print #UBRpt, "Revenue Summary"; Tab(38); "Amount"
    Print #UBRpt, Dash120$
    TotalRevsAmt# = 0
    For RCnt = 1 To MaxRevenue
      TotalRevsAmt# = Round#(TotalRevsAmt# + RevTotals(RCnt))
      Print #UBRpt, RevenueName$(RCnt), Tab(35); Using("########.##", RevTotals(RCnt))
    Next
    Print #UBRpt,
    Print #UBRpt, "Total Credit Balance Amount"; Tab(35); Using("########.##", TotalRevsAmt#)
  End If
  Return

CheckDetailParms:



  OperatorNo$ = fptxtOperator
  Operator = Val(OperatorNo$)
  If Operator = 0 Then
    BegOperator = 0
    EndOperator = 9999
  Else
    BegOperator = Operator
    EndOperator = Operator
  End If
  
  Detail$ = QPTrim$(Left$(fpcboDetail.Text, 1))

  CUSTTYPE$ = QPTrim$(fptxtCustType)
  If Len(CUSTTYPE$) > 0 Then
    UseType = True
  End If

  Select Case Left$(fpcboPrintOrder.Text, 1)
    Case "C"
    IndexName$ = NameIndexFile
    UsingName = True
  Case "A"
    IndexName$ = ""
    UsingAcct = True
  Case "S"
    IndexName$ = TempIndexName
    UsingAddr = True
  Case Else
  End Select
Return
End Sub
Private Sub DetailedTransJournalXX()
  Dim UBCustRecLen As Integer, UBSetupreclen As Integer
  Dim UsingBook As Boolean, UsingName As Boolean, UsingAddr As Boolean
  Dim CustomerCnt As Long, UBTransRecLen As Integer, UBTrans As Integer
  Dim IndexName As String, Handle As Integer, Dash120 As String
  Dim IdxRecLen As Integer, IdxFileSize As Long, UBRpt As Integer
  Dim IdxNumOfRecs As Long, NumOfRecs As Long, MaxRevenue As Integer
  Dim cnt As Long, UBCust As Integer, RCnt As Integer, UseType As Boolean
  Dim ThisType As String, CUSTTYPE As String, CustBook As Integer
  Dim FromBook As Integer, ThruBook As Integer, BadCount As Long
  Dim Trans As Long, UBTransLen As Integer, BegDate As Integer
  Dim EndDate As Integer, BegOperator As Integer, EndOperator As Integer
  Dim BegTrans As Integer, EndTrans As Integer, TransDesc As String
  Dim Amount As Double, TotalTrans As Double, TransCnt As Long
  Dim Detail As String, Date1 As String, Date2 As String, Operator As String
  Dim TotalRevsAmt As Double, EstCnt As Integer, TrType As String, TransBal As Long, TransRef As Long
  Dim TrTyp As Integer, OperatorNo As String, UsingAcct As Boolean, totBal As Double
  Dim ReportFile As String, MoFlag As Boolean, nexttr As Long, TotRef As Double
 'get report parameters
  GoSub CheckDetailParms
  MaxLines = 45
  PageNo = 0
  Dash120$ = String$(80, "-")
  FrmShowPctComp.Label1 = "Creating Transactions"
  FrmShowPctComp.Show , Me
  DoEvents
 ' MaxRevenue = 15
  ''DeActivateControls Me, True
  ReDim RevTotals(1 To 15) As Double
  ReDim RevTotsbef(1 To 15) As Double
  ReDim RevenueName(1 To 15) As String
  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))

  ReDim UBSetUp(1) As UBSetupRecType
  UBSetupreclen = Len(UBSetUp(1))
  ReDim UBTransRec(1) As UBTransRecType
  UBTransRecLen = Len(UBTransRec(1))
  Date1$ = txtDate1
  Date2$ = txtDate2

  BegDate = Date2Num%(Date1$)
  EndDate = Date2Num%(Date2$)
  totBal = 0
  TotRef = 0
  TransRef = 0
  TransBal = 0
  If UsingName Or UsingBook Then
    IdxRecLen = 4               'we are using a long integer
    IdxFileSize& = FileSize(IndexName$)
    IdxNumOfRecs = IdxFileSize& \ IdxRecLen
    ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
    'FGetAH IndexName$, IdxBuff(1), IdxRecLen, IdxNumOfRecs      'load it
    NumOfRecs = IdxNumOfRecs
    Handle = FreeFile
    Open IndexName$ For Random Shared As Handle Len = IdxRecLen
    For cnt& = 1 To IdxNumOfRecs
      Get #Handle, cnt&, IdxBuff(cnt&)
    Next
    Close Handle
  ElseIf UsingAddr Then
'unrem
    SortServiceAddrs frmRptMastCust
    IdxRecLen = 4               'we are using a long integer
    IdxFileSize& = FileSize&(IndexName$)
    IdxNumOfRecs = IdxFileSize& \ IdxRecLen
    ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
    'FGetAH IndexName$, IdxBuff(1), IdxRecLen, IdxNumOfRecs      'load it
    NumOfRecs = IdxNumOfRecs
    Handle = FreeFile
    Open IndexName$ For Random Shared As Handle Len = IdxRecLen
    For cnt& = 1 To IdxNumOfRecs
      Get #Handle, cnt&, IdxBuff(cnt&)
    Next
    Close Handle

  Else
    NumOfRecs = FileSize(UBPath$ + "UBCUST.DAT") \ UBCustRecLen
  End If

  UBCust = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen
  UBTrans = FreeFile
  Open UBPath$ + "UBTRANS.DAT" For Random Shared As UBTrans Len = UBTransRecLen
  ReportFile$ = UBPath$ + "UBCRLIST.RPT"
  UBRpt = FreeFile
  Open ReportFile$ For Output As UBRpt
'  ubsetup = FreeFile
'  Open "UBSETUP.DAT" For Random Shared As ubsetup Len = UBSetupreclen
  LoadUBSetUpFile UBSetUp(), UBSetupreclen

  If Len(TOWNNAME$) = 0 Then
    TOWNNAME$ = "Undefined"
    ' Set Revenue Names to Nothing
    For RCnt = 1 To 15
      RevenueName$(RCnt) = "Not Set"
    Next RCnt
  Else
    'Get ubsetup, 1, ubsetup(1)
    For RCnt = 1 To 15
      RevenueName$(RCnt) = QPTrim$(UBSetUp(1).Revenues(RCnt).RevName)
    Next RCnt
    RCnt = 1
    Do While RCnt <= 15
      If RevenueName$(RCnt) = "" Then
        MaxRevenue = RCnt - 1
        Exit Do
      End If
      RCnt = RCnt + 1
    Loop
  End If

  GoSub DoDetailedRptHeader

  For cnt = 1 To NumOfRecs
    FrmShowPctComp.ShowPctComp cnt, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ''ActivateControls Me, True
      GoTo ExitDetailedListing
    End If

    If UsingName Or UsingAddr Then
      Get UBCust, IdxBuff(cnt).RecNum, UBCustRec(1)
    Else
      Get UBCust, cnt, UBCustRec(1)
    End If

    If UBCustRec(1).DelFlag <> 0 Or UBCustRec(1).Status = "A" Then
      GoTo SkipThisOne
    End If

    If UseType Then
      ThisType$ = QPTrim$(UBCustRec(1).CUSTTYPE)
      If ThisType$ <> CUSTTYPE$ Then
        GoTo SkipThisOne
      End If
    End If

    If LineCnt > MaxLines Then
      Print #UBRpt, FF$
      GoSub DoDetailedRptHeader
    End If
'*************************************
'   Main Body of Printing goes here
    BadCount = 0
    Trans& = UBCustRec(1).LastTrans
   ' Do While Trans& <> 0
   If Trans& > 0 Then
     Get UBTrans, Trans&, UBTransRec(1)
                'TransDesc$ = "Applied Dep"
                'Amount# = Abs(UBTransRec(1).Transamt)
      If (UBTransRec(1).TransDate >= BegDate And UBTransRec(1).TransDate <= EndDate) Then
        If (UBTransRec(1).OperatorNumber >= BegOperator And UBTransRec(1).OperatorNumber <= EndOperator) Then
          If UBTransRec(1).TransType = 5 Or UBTransRec(1).TransType = 105 Then
             If UBTransRec(1).PrevTrans > 0 Then
              Trans& = UBTransRec(1).PrevTrans
              Get UBTrans, Trans&, UBTransRec(1)
              If UBTransRec(1).TransType = 1 Or UBTransRec(1).TransType = 101 And UBTransRec(1).ApplyDepFlag = "Y" Then
                Amount# = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)

                  If Amount# < 0 Then
                    TransDesc$ = "Credit Balance"
                    TotRef# = Round#(TotRef# + Amount#)
                    TransRef& = TransRef& + 1
                  ElseIf Amount# = 0 Then
                    TransDesc$ = "Zero Balance"
                    TransCnt& = TransCnt& + 1
                  ElseIf Amount# > 0 Then
                    TransDesc$ = "Balance Owed"
                    totBal# = Round#(totBal# + Amount#)
                    TransBal& = TransBal& + 1
                  End If
                  If CheckRefOnly.Value = 1 And Amount# < 0 Then

                    Print #UBRpt, Num2Date$(UBTransRec(1).TransDate); Tab(13); Using("#####", UBTransRec(1).CustAcctNo);
                    
                    Print #UBRpt, Tab(25); QPTrim$(UBCustRec(1).CustName);
                    Print #UBRpt, Tab(70); Using("$###,###.##", Amount#)
                    LineCnt = LineCnt + 1
                    Print #UBRpt, Tab(25); QPTrim$(UBCustRec(1).ADDR1)
                    LineCnt = LineCnt + 1
                    If Len(QPTrim$(UBCustRec(1).ADDR2)) > 0 Then
                      Print #UBRpt, Tab(25); QPTrim$(UBCustRec(1).ADDR2)
                      LineCnt = LineCnt + 1
                    End If
                    Print #UBRpt, Tab(25); QPTrim$(UBCustRec(1).CITY); " "; QPTrim$(UBCustRec(1).STATE); " "; QPTrim$(UBCustRec(1).ZIPCODE)
                    LineCnt = LineCnt + 1
                    If fpcboDetail.ListIndex = 1 Then
                      Print #UBRpt, "Final Billing Breakdown ........................"
                      LineCnt = LineCnt + 1
                      For RCnt = 1 To MaxRevenue Step 3
                        Print #UBRpt, RevenueName$(RCnt); Tab(16); Using("#####.##", Round#(UBTransRec(1).RevAmt(RCnt) + UBTransRec(1).TaxAmt(RCnt)));
                        Print #UBRpt, Tab(30); RevenueName$(RCnt + 1); Tab(46); Using("#####.##", Round#(UBTransRec(1).RevAmt(RCnt + 1) + UBTransRec(1).TaxAmt(RCnt + 1)));
                        Print #UBRpt, Tab(58); RevenueName$(RCnt + 2); Tab(70); Using("#####.##", Round#(UBTransRec(1).RevAmt(RCnt + 2) + UBTransRec(1).TaxAmt(RCnt + 2)))
                        LineCnt = LineCnt + 1
                      Next RCnt
                     
                      Print #UBRpt, Dash120$
                      LineCnt = LineCnt + 1
                      
                      'get apply dep trans again
                      Trans& = UBCustRec(1).LastTrans
                      If Trans& > 0 Then Get UBTrans, Trans&, UBTransRec(1)
                      
                      Print #UBRpt, "Balances Before Apply Dep........................"
                      LineCnt = LineCnt + 1
                      For RCnt = 1 To MaxRevenue Step 3
                        Print #UBRpt, RevenueName$(RCnt); Tab(16); Using("#####.##", Round#(UBCustRec(1).CurrRevAmts(RCnt)) + (UBTransRec(1).RevAmt(RCnt) + UBTransRec(1).TaxAmt(RCnt)));
                        Print #UBRpt, Tab(30); RevenueName$(RCnt + 1); Tab(46); Using("#####.##", Round#(UBCustRec(1).CurrRevAmts(RCnt + 1)) + (UBTransRec(1).RevAmt(RCnt + 1) + UBTransRec(1).TaxAmt(RCnt + 1)));
                        Print #UBRpt, Tab(58); RevenueName$(RCnt + 2); Tab(70); Using("#####.##", Round#(UBCustRec(1).CurrRevAmts(RCnt + 2)) + (UBTransRec(1).RevAmt(RCnt + 2) + UBTransRec(1).TaxAmt(RCnt + 2)))
                        LineCnt = LineCnt + 1
                      Next RCnt
                      For RCnt = 1 To MaxRevenue
                        RevTotsbef(RCnt) = Round#(RevTotsbef(RCnt) + Round#(UBCustRec(1).CurrRevAmts(RCnt)) + (UBTransRec(1).RevAmt(RCnt) + UBTransRec(1).TaxAmt(RCnt)))
                      Next

                      Print #UBRpt, Dash120$
                      LineCnt = LineCnt + 1

                      Print #UBRpt, "Apply Deposit Breakdown ........................"
                      LineCnt = LineCnt + 1
                      For RCnt = 1 To MaxRevenue Step 3
                        Print #UBRpt, RevenueName$(RCnt); Tab(16); Using("#####.##", Round#(UBTransRec(1).RevAmt(RCnt) + UBTransRec(1).TaxAmt(RCnt)));
                        Print #UBRpt, Tab(30); RevenueName$(RCnt + 1); Tab(46); Using("#####.##", Round#(UBTransRec(1).RevAmt(RCnt + 1) + UBTransRec(1).TaxAmt(RCnt + 1)));
                        Print #UBRpt, Tab(58); RevenueName$(RCnt + 2); Tab(70); Using("#####.##", Round#(UBTransRec(1).RevAmt(RCnt + 2) + UBTransRec(1).TaxAmt(RCnt + 2)))
                        LineCnt = LineCnt + 1
                      Next RCnt
                      For RCnt = 1 To MaxRevenue
                        RevTotals(RCnt) = Round#(RevTotals(RCnt) + Round#(UBTransRec(1).RevAmt(RCnt) + UBTransRec(1).TaxAmt(RCnt)))
                      Next

                      Print #UBRpt, Dash120$
                      LineCnt = LineCnt + 1


                      Print #UBRpt, "Revenue Balances after Apply Deposit ........................"
                      LineCnt = LineCnt + 1
                      For RCnt = 1 To MaxRevenue Step 3
                        Print #UBRpt, RevenueName$(RCnt); Tab(16); Using("#####.##", Round#(UBCustRec(1).CurrRevAmts(RCnt)));
                        Print #UBRpt, Tab(30); RevenueName$(RCnt + 1); Tab(46); Using("#####.##", Round#(UBCustRec(1).CurrRevAmts(RCnt + 1)));
                        Print #UBRpt, Tab(58); RevenueName$(RCnt + 2); Tab(70); Using("#####.##", Round#(UBCustRec(1).CurrRevAmts(RCnt + 2)))
                        LineCnt = LineCnt + 1
                      Next RCnt
                      Print #UBRpt, Dash120$
                      Print #UBRpt, Dash120$
                      LineCnt = LineCnt + 2
                    TotalTrans# = Round#(TotalTrans# + Amount#)
                    'TransCnt& = TransCnt& + 1
                    If LineCnt > MaxLines - 15 Then
                      Print #UBRpt, FF$
                      GoSub DoDetailedRptHeader
                    End If
                 
                  Else
                    Print #UBRpt, Dash120$
                    LineCnt = LineCnt + 2
                 'Exit Do
                 End If
                 ElseIf CheckRefOnly.Value <> 1 Then
                    Print #UBRpt, Num2Date$(UBTransRec(1).TransDate); Tab(13); Using("#####", UBTransRec(1).CustAcctNo);
                    
                    Print #UBRpt, Tab(25); QPTrim$(UBCustRec(1).CustName);
                    Print #UBRpt, Tab(70); Using("$###,###.##", Amount#)
                    LineCnt = LineCnt + 1
                    Print #UBRpt, Tab(25); QPTrim$(UBCustRec(1).ADDR1)
                    LineCnt = LineCnt + 1
                    If Len(QPTrim$(UBCustRec(1).ADDR2)) > 0 Then
                      Print #UBRpt, Tab(25); QPTrim$(UBCustRec(1).ADDR2)
                      LineCnt = LineCnt + 1
                    End If
                    Print #UBRpt, Tab(25); QPTrim$(UBCustRec(1).CITY); " "; QPTrim$(UBCustRec(1).STATE); " "; QPTrim$(UBCustRec(1).ZIPCODE)
                    LineCnt = LineCnt + 1
                    If fpcboDetail.ListIndex = 1 Then
                      Print #UBRpt, "Final Billing Breakdown ........................"
                      LineCnt = LineCnt + 1
                      For RCnt = 1 To MaxRevenue Step 3
                        Print #UBRpt, RevenueName$(RCnt); Tab(16); Using("#####.##", Round#(UBTransRec(1).RevAmt(RCnt) + UBTransRec(1).TaxAmt(RCnt)));
                        Print #UBRpt, Tab(30); RevenueName$(RCnt + 1); Tab(46); Using("#####.##", Round#(UBTransRec(1).RevAmt(RCnt + 1) + UBTransRec(1).TaxAmt(RCnt + 1)));
                        Print #UBRpt, Tab(58); RevenueName$(RCnt + 2); Tab(70); Using("#####.##", Round#(UBTransRec(1).RevAmt(RCnt + 2) + UBTransRec(1).TaxAmt(RCnt + 2)))
                        LineCnt = LineCnt + 1
                      Next RCnt
                     
                      Print #UBRpt, Dash120$
                      LineCnt = LineCnt + 1
                      
                      'get apply dep trans again
                      Trans& = UBCustRec(1).LastTrans
                      If Trans& > 0 Then Get UBTrans, Trans&, UBTransRec(1)
                      
                      Print #UBRpt, "Balances Before Apply Dep........................"
                      LineCnt = LineCnt + 1
                      For RCnt = 1 To MaxRevenue Step 3
                        Print #UBRpt, RevenueName$(RCnt); Tab(16); Using("#####.##", Round#(UBCustRec(1).CurrRevAmts(RCnt)) + (UBTransRec(1).RevAmt(RCnt) + UBTransRec(1).TaxAmt(RCnt)));
                        Print #UBRpt, Tab(30); RevenueName$(RCnt + 1); Tab(46); Using("#####.##", Round#(UBCustRec(1).CurrRevAmts(RCnt + 1)) + (UBTransRec(1).RevAmt(RCnt + 1) + UBTransRec(1).TaxAmt(RCnt + 1)));
                        Print #UBRpt, Tab(58); RevenueName$(RCnt + 2); Tab(70); Using("#####.##", Round#(UBCustRec(1).CurrRevAmts(RCnt + 2)) + (UBTransRec(1).RevAmt(RCnt + 2) + UBTransRec(1).TaxAmt(RCnt + 2)))
                        LineCnt = LineCnt + 1
                      Next RCnt
                      For RCnt = 1 To MaxRevenue
                        RevTotsbef(RCnt) = Round#(RevTotsbef(RCnt) + Round#(UBCustRec(1).CurrRevAmts(RCnt)) + (UBTransRec(1).RevAmt(RCnt) + UBTransRec(1).TaxAmt(RCnt)))
                      Next

                      Print #UBRpt, Dash120$
                      LineCnt = LineCnt + 1

                      Print #UBRpt, "Apply Deposit Breakdown ........................"
                      LineCnt = LineCnt + 1
                      For RCnt = 1 To MaxRevenue Step 3
                        Print #UBRpt, RevenueName$(RCnt); Tab(16); Using("#####.##", Round#(UBTransRec(1).RevAmt(RCnt) + UBTransRec(1).TaxAmt(RCnt)));
                        Print #UBRpt, Tab(30); RevenueName$(RCnt + 1); Tab(46); Using("#####.##", Round#(UBTransRec(1).RevAmt(RCnt + 1) + UBTransRec(1).TaxAmt(RCnt + 1)));
                        Print #UBRpt, Tab(58); RevenueName$(RCnt + 2); Tab(70); Using("#####.##", Round#(UBTransRec(1).RevAmt(RCnt + 2) + UBTransRec(1).TaxAmt(RCnt + 2)))
                        LineCnt = LineCnt + 1
                      Next RCnt
                      For RCnt = 1 To MaxRevenue
                        RevTotals(RCnt) = Round#(RevTotals(RCnt) + Round#(UBTransRec(1).RevAmt(RCnt) + UBTransRec(1).TaxAmt(RCnt)))
                      Next

                      Print #UBRpt, Dash120$
                      LineCnt = LineCnt + 1


                      Print #UBRpt, "Revenue Balances after Apply Deposit ........................"
                      LineCnt = LineCnt + 1
                      For RCnt = 1 To MaxRevenue Step 3
                        Print #UBRpt, RevenueName$(RCnt); Tab(16); Using("#####.##", Round#(UBCustRec(1).CurrRevAmts(RCnt)));
                        Print #UBRpt, Tab(30); RevenueName$(RCnt + 1); Tab(46); Using("#####.##", Round#(UBCustRec(1).CurrRevAmts(RCnt + 1)));
                        Print #UBRpt, Tab(58); RevenueName$(RCnt + 2); Tab(70); Using("#####.##", Round#(UBCustRec(1).CurrRevAmts(RCnt + 2)))
                        LineCnt = LineCnt + 1
                      Next RCnt
                      Print #UBRpt, Dash120$
                      Print #UBRpt, Dash120$
                      LineCnt = LineCnt + 2
                    TotalTrans# = Round#(TotalTrans# + Amount#)
                    'TransCnt& = TransCnt& + 1
                    If LineCnt > MaxLines - 15 Then
                      Print #UBRpt, FF$
                      GoSub DoDetailedRptHeader
                    End If
                 
                  Else
                    Print #UBRpt, Dash120$
                    LineCnt = LineCnt + 2
                 'Exit Do
                 End If
                End If
               End If
          End If
        End If
      End If
     End If
   End If '   Trans& = UBTransRec(1).PrevTrans
   ' Loop
SkipThisOne:
'    ShowPctComp cnt, NumOfRecs
  Next

  GoSub DoDetailedRptFooter
  If fpcboRptType.ListIndex = 1 Then Print #UBRpt, FF$

  Close

  Erase IdxBuff, UBCustRec
 ''' ActivateControls Me, True
  'END
 If fpcboRptType.ListIndex <> 1 Then
    Load frmLoadingRpt
    frmLoadingRpt.setwherefrom frmRptFinalApply
    ARptLineRpt.GetName ReportFile$
    ARptLineRpt.startrpt
 Else
  'If Not AbortFlag Then
  '  PrintRptFile "Detailed Journal Report.", "UBDJLIST.RPT", LptPort, RetCode, EntryPoint
 ' End If
    ViewPrint ReportFile$, "Applied/Final Report", True
  'KillFile "UBDJLIST.RPT"
  End If
ExitDetailedListing:

  Exit Sub

DoDetailedRptHeader:
  PageNo = PageNo + 1
  Print #UBRpt,
  Print #UBRpt,
  Print #UBRpt, TOWNNAME$
  Print #UBRpt, " Applied/Final Report"; Tab(70); "Page #"; PageNo
    Print #UBRpt, "               Report Date: "; Date$
    Print #UBRpt, "Beginning Transaction Date: "; Date1$
    Print #UBRpt, "   Ending Transaction Date: "; Date2$
  If Val(Operator$) = 0 Then
    Print #UBRpt, "                Operator #: ALL"
  Else
    Print #UBRpt, "                Operator #: "; Operator$
  End If
    Print #UBRpt, "               Show Detail: "; Detail$
    Print #UBRpt, "             Customer Type: ";

  If UseType Then
    Print #UBRpt, CUSTTYPE$
  Else
    Print #UBRpt, "N/A"
  End If

  Print #UBRpt,
  Print #UBRpt, "  Date"; Tab(15); "Acct #"; Tab(30); "Customer Name/Address"; Tab(65); "     Balance"
  Print #UBRpt, Dash120$
  LineCnt = 13
  Return

DoDetailedRptFooter:

  Print #UBRpt, Dash120$
  Print #UBRpt, "Transactions: "; TransRef&; "                       Total of Credit Balances: "; Using("$##,###,###.##", TotRef#)
  Print #UBRpt, "Transactions: "; TransBal&; "                         Total of Balances Owed: "; Using("$##,###,###.##", totBal#)
  Print #UBRpt, "Transactions: "; TransCnt&; "                Total Customers with Balance of: "; Using("$##,###,###.##", 0)
  
  
  'If fpcboRptType.ListIndex = 1 Then
  If fpcboDetail.ListIndex = 1 Then
    Print #UBRpt, FF$
    PageNo = PageNo + 1
    Print #UBRpt,
    Print #UBRpt,
    Print #UBRpt, TOWNNAME$
    Print #UBRpt, " Applied/Final Report"; Tab(70); "Page #"; PageNo
      Print #UBRpt, "               Report Date: "; Date$
      Print #UBRpt, "Beginning Transaction Date: "; Date1$
      Print #UBRpt, "   Ending Transaction Date: "; Date2$
    If Val(Operator$) = 0 Then
      Print #UBRpt, "                Operator #: ALL"
    Else
      Print #UBRpt, "                Operator #: "; Operator$
    End If
      Print #UBRpt, "               Show Detail: "; Detail$
      Print #UBRpt, "             Customer Type: ";
  
    If UseType Then
      Print #UBRpt, CUSTTYPE$
    Else
      Print #UBRpt, "N/A"
    End If
    Print #UBRpt, ""
    Print #UBRpt, "Revenue Summary of Applied Deposits"; Tab(38); "Amount"
    Print #UBRpt, Dash120$
    TotalRevsAmt# = 0
    For RCnt = 1 To MaxRevenue
      TotalRevsAmt# = Round#(TotalRevsAmt# + RevTotals(RCnt))
      Print #UBRpt, RevenueName$(RCnt), Tab(35); Using("########.##", RevTotals(RCnt))
    Next
    Print #UBRpt,
    Print #UBRpt, "Total Deposits Applied"; Tab(35); Using("########.##", TotalRevsAmt#)
    Print #UBRpt,
    Print #UBRpt, "***********"
    Print #UBRpt, "Balances Before Applied Deposits"; Tab(38); "Amount"
    Print #UBRpt, Dash120$
    TotalRevsAmt# = 0
    For RCnt = 1 To MaxRevenue
      TotalRevsAmt# = Round#(TotalRevsAmt# + RevTotsbef(RCnt))
      Print #UBRpt, RevenueName$(RCnt), Tab(35); Using("########.##", RevTotsbef(RCnt))
    Next
    Print #UBRpt,
    Print #UBRpt, "Total Revenues Before Applied"; Tab(35); Using("########.##", TotalRevsAmt#)

    'Print #UBRpt, "Total Balance Owed"; Tab(35);
  End If
  Return

CheckDetailParms:



  OperatorNo$ = fptxtOperator
  Operator = Val(OperatorNo$)
  If Operator = 0 Then
    BegOperator = 0
    EndOperator = 9999
  Else
    BegOperator = Operator
    EndOperator = Operator
  End If
  
  Detail$ = QPTrim$(Left$(fpcboDetail.Text, 1))

  CUSTTYPE$ = QPTrim$(fptxtCustType)
  If Len(CUSTTYPE$) > 0 Then
    UseType = True
  End If

  Select Case Left$(fpcboPrintOrder.Text, 1)
    Case "C"
    IndexName$ = NameIndexFile
    UsingName = True
  Case "A"
    IndexName$ = ""
    UsingAcct = True
  Case "S"
    IndexName$ = TempIndexName
    UsingAddr = True
  Case Else
  End Select
Return
End Sub

