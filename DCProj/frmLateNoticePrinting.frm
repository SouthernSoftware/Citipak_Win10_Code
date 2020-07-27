VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmLateNoticePrinting 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Late Notice Printing"
   ClientHeight    =   8868
   ClientLeft      =   3924
   ClientTop       =   1884
   ClientWidth     =   12216
   Icon            =   "frmLateNoticePrinting.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   12216
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboBalType 
      Height          =   348
      Left            =   5388
      TabIndex        =   5
      Top             =   3828
      Width           =   3612
      _Version        =   196608
      _ExtentX        =   6371
      _ExtentY        =   614
      Text            =   ""
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
      BorderDropShadowWidth=   1
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
      ColDesigner     =   "frmLateNoticePrinting.frx":08CA
   End
   Begin LpLib.fpCombo fpcboPrintOrder 
      Height          =   348
      Left            =   5388
      TabIndex        =   6
      Top             =   4248
      Width           =   3612
      _Version        =   196608
      _ExtentX        =   6371
      _ExtentY        =   614
      Text            =   ""
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
      ColDesigner     =   "frmLateNoticePrinting.frx":0C99
   End
   Begin LpLib.fpCombo fpcboActiveOnly 
      Height          =   348
      Left            =   5400
      TabIndex        =   7
      Top             =   4656
      Width           =   924
      _Version        =   196608
      _ExtentX        =   1630
      _ExtentY        =   614
      Text            =   ""
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
      ColDesigner     =   "frmLateNoticePrinting.frx":1068
   End
   Begin EditLib.fpText fptxtMessage 
      Height          =   324
      Index           =   0
      Left            =   4716
      TabIndex        =   8
      Top             =   5664
      Width           =   4644
      _Version        =   196608
      _ExtentX        =   8191
      _ExtentY        =   572
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.2
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
      HideSelection   =   0   'False
      InvalidColor    =   -2147483637
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483637
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   25
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
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
      TabIndex        =   12
      Top             =   7776
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
      Left            =   10080
      TabIndex        =   13
      Top             =   7776
      Width           =   1332
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   14
      Top             =   8508
      Width           =   12216
      _ExtentX        =   21548
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7154
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7154
            TextSave        =   "10:56 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7154
            TextSave        =   "5/19/2005"
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
   Begin EditLib.fpDateTime txtPayDate 
      Height          =   348
      Left            =   5388
      TabIndex        =   3
      Top             =   3009
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
   Begin EditLib.fpCurrency fpMinBal 
      Height          =   348
      Left            =   5388
      TabIndex        =   4
      Top             =   3420
      Width           =   1476
      _Version        =   196608
      _ExtentX        =   2603
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      CaretInsert     =   0
      CaretOverWrite  =   3
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
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   0
      ControlType     =   0
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   -1
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   ""
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "-9000000000"
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
      IncDec          =   1
      BorderGrayAreaColor=   -2147483637
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
   Begin EditLib.fpText fptxtBook2 
      Height          =   348
      Left            =   5388
      TabIndex        =   1
      Top             =   2187
      Width           =   636
      _Version        =   196608
      _ExtentX        =   1122
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
      MaxLength       =   2
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
   Begin EditLib.fpText fptxtBook1 
      Height          =   348
      Left            =   5388
      TabIndex        =   0
      Top             =   1776
      Width           =   636
      _Version        =   196608
      _ExtentX        =   1122
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
      MaxLength       =   2
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
   Begin EditLib.fpDateTime txtNoticeDate 
      Height          =   348
      Left            =   5388
      TabIndex        =   2
      Top             =   2598
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
   Begin EditLib.fpText fptxtMessage 
      Height          =   324
      Index           =   1
      Left            =   4716
      TabIndex        =   9
      Top             =   5988
      Width           =   4644
      _Version        =   196608
      _ExtentX        =   8191
      _ExtentY        =   572
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.2
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
      HideSelection   =   0   'False
      InvalidColor    =   -2147483637
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483637
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   25
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fptxtMessage 
      Height          =   324
      Index           =   2
      Left            =   4716
      TabIndex        =   10
      Top             =   6312
      Width           =   4644
      _Version        =   196608
      _ExtentX        =   8191
      _ExtentY        =   572
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.2
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
      HideSelection   =   0   'False
      InvalidColor    =   -2147483637
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483637
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   25
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fptxtMessage 
      Height          =   324
      Index           =   3
      Left            =   4716
      TabIndex        =   11
      Top             =   6636
      Width           =   4644
      _Version        =   196608
      _ExtentX        =   8191
      _ExtentY        =   572
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.2
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
      HideSelection   =   0   'False
      InvalidColor    =   -2147483637
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483637
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   25
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin VB.Label Labelthru 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Thru Book:"
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
      Left            =   3804
      TabIndex        =   27
      Top             =   2211
      Width           =   1428
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Print Notices On:"
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
      Left            =   3156
      TabIndex        =   26
      Top             =   3855
      Width           =   2076
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Late Notice Information"
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
      Left            =   3894
      TabIndex        =   25
      Top             =   768
      Width           =   4428
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3222
      Top             =   528
      Width           =   5772
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Notice Date:"
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
      Left            =   3540
      TabIndex        =   24
      Top             =   2622
      Width           =   1692
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Only Active Accounts:"
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
      Index           =   2
      Left            =   2508
      TabIndex        =   23
      Top             =   4680
      Width           =   2724
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
      Height          =   372
      Index           =   3
      Left            =   2556
      TabIndex        =   22
      Top             =   4266
      Width           =   2676
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Pay By Date:"
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
      Left            =   3516
      TabIndex        =   21
      Top             =   3033
      Width           =   1716
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Minimum Balance:"
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
      Index           =   4
      Left            =   2628
      TabIndex        =   20
      Top             =   3444
      Width           =   2604
   End
   Begin VB.Label Labelm1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Message Line 1:"
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
      Left            =   2364
      TabIndex        =   19
      Top             =   5664
      Width           =   2244
   End
   Begin VB.Label Labelm4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Message Line 4:"
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
      Left            =   2508
      TabIndex        =   18
      Top             =   6672
      Width           =   2100
   End
   Begin VB.Label Labelm2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Message Line 2:"
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
      Left            =   2436
      TabIndex        =   17
      Top             =   6000
      Width           =   2172
   End
   Begin VB.Label Labelfrom 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "From Book:"
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
      Left            =   3588
      TabIndex        =   16
      Top             =   1800
      Width           =   1644
   End
   Begin VB.Label Labelm3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Message Line 3:"
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
      Left            =   2388
      TabIndex        =   15
      Top             =   6336
      Width           =   2220
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   5748
      Left            =   2496
      Top             =   1512
      Width           =   7236
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   972
      Left            =   3222
      Top             =   408
      Width           =   5772
   End
   Begin VB.Line Line1 
      X1              =   2496
      X2              =   9708
      Y1              =   5376
      Y2              =   5376
   End
End
Attribute VB_Name = "frmLateNoticePrinting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim CycleFlag As Boolean, OKFlag As Boolean, BadDate As Boolean
Dim ErFlag As Boolean, LNType As Integer, LPIFlag As Boolean
Dim OkiMode As Integer ' 1 for not ibm, 2 for ibm
Dim Rteflag As Boolean, AcctBar As Boolean
Private Sub cmdExit_Click()
  Load frmUBLateNoticeMenu
  DoEvents
  frmUBLateNoticeMenu.Show
  Unload Me
  DoEvents
End Sub

Private Sub cmdPrint_Click()
  CheckFields
  LPIFlag = False
  If OKFlag = True Then
'do print stuff here
'depending on which late notice they have selected in setup
  If LNType = 1 Then
    frmReportOpt.Show 1
    DeActivateControls Me
    If rptopt = 1 Then
    'do the graphics
     PrintLateNotices True
    ElseIf rptopt = 2 Then
    'do the text
      PrintLateNotices False
      ActivateControls Me
    End If
  ElseIf LNType = 8 Or LNType = 9 Then
    PrintLateNotices True
  Else
    DeActivateControls Me
    PrintLateNotices False
    ActivateControls Me
  End If
  End If
End Sub

Private Sub fptxtBook1_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub
Private Sub fptxtBook2_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub
Private Sub fptxtbook1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fptxtBook2.SetFocus
  End If
End Sub
Private Sub fptxtbook2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    txtNoticeDate.SetFocus
  End If
End Sub


Private Sub txtNoticeDate_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    txtPayDate.SetFocus
  End If
End Sub
Private Sub txtPayDate_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpMinBal.SetFocus
  End If
End Sub

Private Sub fpMinBal_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpcboBalType.SetFocus
  End If
End Sub
Private Sub fpcboBalType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboBalType.ListDown = True
  End If
  If fpcboBalType.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpcboPrintOrder.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpMinBal.SetFocus
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
      fpcboActiveOnly.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpcboBalType.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub
Private Sub fpcboActiveOnly_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboActiveOnly.ListDown = True
  End If
  If fpcboActiveOnly.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      If fptxtMessage(0).Visible = True Then
        fptxtMessage(0).SetFocus
      Else
        cmdPrint.SetFocus
      End If
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpcboPrintOrder.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub

Private Sub fptxtMessage_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    If Index < 3 Then
     If fptxtMessage(Index + 1).Visible = True Then
      fptxtMessage(Index + 1).SetFocus
     Else
      cmdPrint.SetFocus
     End If
    Else
      cmdPrint.SetFocus
    End If
  End If
End Sub

Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        UBLog "Closed via LateNoticePrinting by " + PWUser$
        CitiTerminate
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
      KeyCode = 0
      DoEvents
      Call cmdExit_Click
    Case vbKeyF10:
      KeyCode = 0
      DoEvents
      Call cmdPrint_Click
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Dim UBSetupreclen As Integer, cnt As Integer, UBBillSetuplen As Integer
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
  ReDim UBSetUp(1) As UBSetupRecType
  UBSetupreclen = Len(UBSetUp(1))
  LoadUBSetUpFile UBSetUp(), UBSetupreclen
  fpcboPrintOrder.AddItem "Customer Name Order"
  fpcboPrintOrder.AddItem "Account Number Order"
  fpcboPrintOrder.AddItem "Location Number Order"
  fpcboPrintOrder.AddItem "ZipCode Order"
  fpcboPrintOrder.AddItem "Zip/Location Order"
  fpcboActiveOnly.ListIndex = 0
  fpcboActiveOnly.AddItem "Yes"
  fpcboActiveOnly.AddItem "No"
  fpcboActiveOnly.ListIndex = 0
  If UBSetUp(1).BILLCYCL = "Y" Then
    CycleFlag = True
    Labelfrom.Caption = "From Cycle:"
    Labelthru.Caption = "Thru Cycle:"
  End If
  txtNoticeDate.Text = Format(Now, "mm/dd/yyyy")
  txtPayDate.Text = Format(Now, "mm/dd/yyyy")
  fpcboBalType.AddItem "Current Balance"
  fpcboBalType.AddItem "Previous Balance"
  fpcboBalType.AddItem "Total Balance"
  fpcboBalType.ListIndex = 0
  'get late notice type from setup and store integer
  'at same time get the OkiMode 1 is not ibm, 2 is ibm
  ReDim UBBillSetup(1) As UBBillSetupType
  UBBillSetuplen = Len(UBBillSetup(1))
  LoadUBBillSetUpFile UBBillSetup(), UBBillSetuplen
  LNType = UBBillSetup(1).LateNotice
  If UBBillSetup(1).RtePrint = 1 Then
    Rteflag = True
  Else
    Rteflag = False
  End If
'  If UBBillSetup(1).PostBar = "Y" Then
'    PostBar = True
'  Else
'    PostBar = False
'  End If
  If UBBillSetup(1).AcctBar = "Y" Then
    AcctBar = True
  Else
    AcctBar = False
  End If
  OkiMode = 2
  Select Case LNType
    Case 1:
      For cnt = 0 To 3
        fptxtMessage(cnt).Visible = False
      Next
      Labelm1.Visible = False
      Labelm2.Visible = False
      Labelm3.Visible = False
      Labelm4.Visible = False
    Case 2:
      For cnt = 0 To 3
        fptxtMessage(cnt).Maxlength = 55
      Next
    Case 3:
      For cnt = 0 To 3
        fptxtMessage(cnt).Maxlength = 30
      Next
    Case 4:
      For cnt = 1 To 3
        fptxtMessage(cnt).Visible = False
      Next
      Labelm2.Visible = False
      Labelm3.Visible = False
      Labelm4.Visible = False
    Case 5 To 7:
      Labelm2.Caption = "   1 Continued:"
      Labelm3.Caption = "Message Line 2:"
      Labelm4.Caption = "   2 Continued:"
    Case 8:
      Labelm3.Visible = False
      Labelm4.Visible = False
      fptxtMessage(2).Visible = False
      fptxtMessage(3).Visible = False
      fptxtMessage(0).Maxlength = 30
      fptxtMessage(1).Maxlength = 30
    Case 9:
      For cnt = 0 To 3
        fptxtMessage(cnt).Maxlength = 50
      Next
    Case Else
  End Select
     
End Sub

Private Sub Form_Resize()
  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
  End If
End Sub
Private Sub CheckPostDate()
  Dim PenDate As String, PenDate2 As String
  PenDate$ = txtNoticeDate.Text
  PenDate2$ = txtPayDate.Text
  If Val(Left$(PenDate$, 2)) < 1 Or Val(Left$(PenDate$, 2)) > 12 Then
    If Val(Mid$(PenDate$, 4, 2)) < 1 Or Val(Mid$(PenDate$, 4, 2)) > 31 Then
      BadDate = True
    Else
      BadDate = False
    End If
  Else
    BadDate = False
  End If
  If Val(Left$(PenDate2$, 2)) < 1 Or Val(Left$(PenDate2$, 2)) > 12 Then
    If Val(Mid$(PenDate2$, 4, 2)) < 1 Or Val(Mid$(PenDate2$, 4, 2)) > 31 Then
      BadDate = True
    Else
      BadDate = False
    End If
  Else
    BadDate = False
  End If
End Sub
Private Function CheckFields()
  ReDim MsgText(0 To 5) As String
  Dim FntSize As Integer
  MsgText(2) = ""
  OKFlag = False
  CheckPostDate
  If BadDate = True Then
    MsgText(3) = "You Have Not Entered a Proper"
    MsgText(4) = "Date!"
  ElseIf fpcboPrintOrder.ListIndex = -1 Then
    MsgText(3) = "Invalid Printing Order."
    MsgText(4) = "Correct and try again."
  ElseIf fpcboBalType.ListIndex = -1 Then
    MsgText(3) = "Must identify Balance source."
    MsgText(4) = "Correct and try again."
  ElseIf fpcboActiveOnly.ListIndex = -1 Then
    MsgText(3) = "Do Not Leave Active Only field blank."
    MsgText(4) = "Correct and try again."
  ElseIf Val(fptxtBook1.Text) > Val(fptxtBook2.Text) Then 'Val(fptxtBook1.Text) = 0 And Val(fptxtBook2.Text) = 0 Or
    MsgText(3) = "Invalid Book or Cycle range!"
    MsgText(4) = "Correct and try again."
  Else
    OKFlag = True
  End If
  If Not OKFlag Then
    frmMsgDialog.RetLabel = "-2"
    FntSize = frmMsgDialog.Label(2).FontSize
    frmMsgDialog.Label(1).FontSize = (FntSize + 2)
    frmMsgDialog.Label(2).FontSize = (FntSize + 2)
    frmMsgDialog.Label(3).FontSize = (FntSize + 2)
    MsgText(0) = "ERROR:"
    MsgText(1) = ""
    MsgText(5) = ""
    GetOKorNot MsgText(), True
  End If

End Function

Private Sub PrintLateNotices(Grpt As Boolean) '(NoticeInfo As NoticeInfoType)
  Dim PDate As String, NDate As String, NMonth As String
  Dim LongPDate As String, LongNDate As String, UBSetupLen As Integer
  Dim PSAFlag As Boolean, UseCycle As Boolean, FromBC As Integer
  Dim ThruBC As Integer, MinBalance As Double, IndexName As String
  Dim NoIndex As Boolean, UBCustRecLen As Integer, TBooks As Integer
  Dim NumOfRecs As Long, IdxNumOfRecs As Long, Handle As Integer
  Dim cnt As Long, UBCst As Integer, UBRpt As Integer, UBLRec As Integer
  Dim Next2Print As Integer, AcctNo As Long, GotWater As Boolean
  Dim CustBC As Integer, Location As String, Acct As String, Zip As String
  Dim ZipLen As Integer, TotalBal As Double, CustBal As Double
  Dim Print1 As Integer, PrnCnt As Integer, NIfile As Integer
  Dim fmt1 As String, fmt2 As String, ReportFile As String, PCnt As Integer
  Dim LaLe As Integer, lenlate As Integer, cntll As Integer, Ext As String
  Dim DeDate As String, lenNI As Integer, AcctNum As Long, Totalamt As Double
  Dim AcctLen As Integer, Previous As Double, Current As Double
  Dim WRevCnt As Integer, PZip As String, ZDigit As String, ToPrint As String
  Dim CustMsg As String, MPCnt As Integer, tmprev As Double, LNcnt As Integer
  Dim ToPrint2 As String, endit As Boolean, UBRptA As Integer, MaskNotice As String
  Dim Fmt10 As String, Fmt10a As String, Fmt15 As String, Today As String
  Dim SCSFileName As String, ChkName As String, BillOutRecLen As Integer
  FrmShowPctComp.Label1 = "Creating Late Notices"
  FrmShowPctComp.Show , Me
  endit = False
  'if lntype = 1 then
  If Exist(UBPath$ + "UBLatLet.dat") Then
    ReDim latelet(1) As UBLateLetterType
    LaLe = FreeFile
    lenlate = Len(latelet(1))
    Open UBPath$ + "UBLatLet.dat" For Random Shared As LaLe Len = lenlate
    Get LaLe, 1, latelet(1)
    Close
  End If
  If Grpt Then
    MaxLines = 48
  Else
    MaxLines = 53
  End If
  If LNType = 98 Then
    CrLf$ = Chr$(13) + Chr$(10)
    Fmt10$ = "##########"
    Fmt10a$ = "#######.##"
    Fmt15$ = "############.##"
  Today$ = Date$
  Ext$ = ".LNT"
  ReDim PrintRec(1) As BillOutRecType
  BillOutRecLen = Len(PrintRec(1))
  SCSFileName$ = Left$(Today$, 2) + Mid$(Today$, 4, 2) + Right$(Today$, 2) + "N"
  For cnt = 1 To 9
    ChkName$ = SCSFileName$ + QPTrim$(Str$(cnt)) + Ext$
    If Exist(ChkName$) = False Then
      SCSFileName$ = ChkName$
      Exit For
    End If
  Next

  UBRpt = FreeFile
  Open SCSFileName$ For Random Shared As UBRpt Len = BillOutRecLen
  ReportFile$ = SCSFileName$
  End If

'endif
  fmt1$ = String$(80, "-")
  fmt2$ = "$###,###,###.##"
  PDate$ = txtPayDate
  NDate$ = txtNoticeDate
  LNcnt = 0
  ToPrint2$ = "~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~"
  'NMonth$ = Left$(MakeMonth$(NDate$), 3) + "."
  ToPrint$ = ""
  LongPDate$ = FormatDateTime(PDate$, vbLongDate)
  LongNDate$ = FormatDateTime(NDate$, vbLongDate)
  
  'load setup file
  ReDim UBSetUpRec(1) As UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen
  TOWNNAME$ = QPTrim$(UCase$(UBSetUpRec(1).UTILNAME))
  If InStr(TOWNNAME$, "GILES") > 0 Then
    PSAFlag = True
    GoTo PissySkipCycle
  End If

  If UBSetUpRec(1).BILLCYCL = "Y" Then
    UseCycle = True
  End If

PissySkipCycle:

  FromBC = Val(fptxtBook1)
  ThruBC = Val(fptxtBook2)

  MinBalance# = fpMinBal.DoubleValue

  PageNo = 0

  'Section to check for customer modifications
  ReDim UBSetUpRec(1) As UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen      'load setup file
  TOWNNAME$ = UBSetUpRec(1).UTILNAME

  Select Case fpcboPrintOrder.ListIndex
  Case 0
    IndexName$ = NameIndexFile
  Case 1
    NoIndex = True
  Case 2
    IndexName$ = BookIndexFile
  Case 3
    IndexName$ = UBPath$ + "UBTEMP.IDX"
    MakeMowZipCodeIndex "ZipCode"
  Case 4
    IndexName$ = UBPath$ + "UBTEMP.IDX"
    MakeZipLocationIndex "ZipLocation"
  End Select

  OKFlag = True

  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))

  TBooks = 0
  If NoIndex = False Then
    NumOfRecs = FileSize(IndexName$) \ 4
    ReDim IndexArray(1 To NumOfRecs) As UBCustIndexRecType
    'FGetAH IndexName$, IndexArray(1), , NumOfRecs
    IdxNumOfRecs = NumOfRecs
    Handle = FreeFile
    Open IndexName$ For Random Shared As Handle Len = 4 'IdxRecLen
    For cnt& = 1 To IdxNumOfRecs
      Get #Handle, cnt&, IndexArray(cnt&)
    Next
    Close Handle

  Else
    NumOfRecs = FileSize(UBPath$ + "UBCUST.DAT") \ UBCustRecLen
  End If
'''

  UBCst = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCst Len = UBCustRecLen
  If Not LNType = 98 Then
    UBRpt = FreeFile
    ReportFile$ = UBPath$ + "UBLATNOT.RPT"
    Open ReportFile$ For Output As UBRpt
  End If
'01-07-99 Added record list of late notices printed (for mailing labels)
  KillFile UBPath$ + "UBLNIDX.DAT"

  UBLRec = FreeFile
  Open UBPath$ + "UBLNIDX.DAT" For Random Shared As UBLRec Len = 4

'  BlockClear
'  ShowProcessingScrn "Processing Late Notices"

  Next2Print = 1

  For cnt = 1 To NumOfRecs
  'If cnt = NumOfRecs Then Stop
    FrmShowPctComp.ShowPctComp cnt, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      GoTo ExitLatePrint
    End If

    If NoIndex Then
      AcctNo& = cnt
    Else
      AcctNo& = IndexArray(cnt).RecNum
    End If
    Get UBCst, AcctNo&, UBCustRec(1)
    GotWater = False
    If UBCustRec(1).DelFlag = 0 And UBCustRec(1).CUTOFFYN = "Y" Then
      If UseCycle Then
        CustBC = UBCustRec(1).BILLCYCL
      Else
        CustBC = Val(UBCustRec(1).Book)
      End If
      If CustBC < FromBC Or CustBC > ThruBC Then
        GoTo SkipEm
      End If

      If UBCustRec(1).CurrRevAmts(1) > 0 Then
        GotWater = True
      End If
      Location$ = UBCustRec(1).Book + "-" + UBCustRec(1).SEQNUMB
      Acct$ = QPTrim$(Str$(AcctNo&))
      Zip$ = QPTrim$(UBCustRec(1).ZIPCODE)
      ZipLen = Len(Zip$)
      Select Case ZipLen
      Case 9, 10
        Zip$ = Left$(Zip$, 5) + "-" + Right$(Zip$, 4)
      Case Else
        Zip$ = Left$(Zip$, 5)
      End Select

      TotalBal# = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
      Select Case fpcboBalType.ListIndex
      Case 0
        CustBal# = Round#(UBCustRec(1).CurrBalance)
        If (CustBal# >= MinBalance#) And (CustBal# > 0) Then
          If TotalBal# > 0 Then
            If fpcboActiveOnly.ListIndex = 0 Then
              If UBCustRec(1).Status = "A" Then
                Print1 = Print1 + 1
                GoSub PrintThemOne
              End If
            Else
              Print1 = Print1 + 1
              GoSub PrintThemOne
            End If
          End If
        End If
      Case 1
        CustBal# = Round#(UBCustRec(1).PrevBalance)
        If (CustBal# >= MinBalance#) And (CustBal# > 0) Then
          If TotalBal# > 0 Then
            If fpcboActiveOnly.ListIndex = 0 Then
              If UBCustRec(1).Status = "A" Then
                Print1 = Print1 + 1
                GoSub PrintThemOne
              End If
            Else
              Print1 = Print1 + 1
              GoSub PrintThemOne
            End If
          End If
        End If
      Case 2
        CustBal# = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
        If CustBal# >= MinBalance# Then
          If TotalBal# > 0 Then
            If fpcboActiveOnly.ListIndex = 0 Then
              If UBCustRec(1).Status = "A" Then
                Print1 = Print1 + 1
                GoSub PrintThemOne
              End If
            Else
              Print1 = Print1 + 1
              GoSub PrintThemOne
            End If
          End If
        End If
      End Select
    End If

    If Next2Print = 4 Then
      Next2Print = 1
      Print #UBRpt, Chr$(12)
    End If
'    If AskAbandonPrint% Then
'      AbortFlag = True
'      Exit For
'    End If
SkipEm:
   ' ShowPctComp cnt, NumOfRecs

    If InStr(Command$, "TEST") > 0 Then
      If PrnCnt > 1 Then Exit For
    End If

  Next
    
  If LNType = 8 Then
    endit = True
    GoSub Dblcheck
  End If

  Close
  
  NIfile = FreeFile
  ReDim NoticeInfo(1) As NoticeInfoType
  lenNI = Len(NoticeInfo(1))
  Open UBPath$ + "UBLNINFO.DAT" For Random Shared As NIfile Len = lenNI
  
  NoticeInfo(1).FromBC = Val(fptxtBook1)
  NoticeInfo(1).ThruBC = Val(fptxtBook2)
  NoticeInfo(1).NoticeDate = Date2Num(txtNoticeDate)
  NoticeInfo(1).PayByDate = Date2Num(txtPayDate)
  NoticeInfo(1).MinBalance# = fpMinBal.DoubleValue
  NoticeInfo(1).BalanceType = fpcboBalType.ListIndex + 1
  NoticeInfo(1).PrnOrder = fpcboPrintOrder.ListIndex + 1
  If fpcboActiveOnly.ListIndex = 0 Then
    NoticeInfo(1).UseAFlag = True
  Else
    NoticeInfo(1).UseAFlag = False
  End If
  NoticeInfo(1).MsgLine1 = fptxtMessage(0)
  NoticeInfo(1).MsgLine2 = fptxtMessage(1)
  NoticeInfo(1).MsgLine3 = fptxtMessage(2)
  NoticeInfo(1).MsgLine4 = fptxtMessage(3)
  NoticeInfo(1).PrnCnt = PrnCnt

  Put NIfile, 1, NoticeInfo(1)
  Close
  GoSub DoLNMask
'If lntype = 1 then
  If Grpt Then
    Load frmLoadingRpt
    frmLoadingRpt.setwherefrom frmLateNoticePrinting
    If LNType = 1 Then
      ARptLineRpt.GetName ReportFile$
      ARptLineRpt.startrpt
    ElseIf LNType = 8 Then
      ARptLateNotice1.GetName ReportFile$
      ARptLateNotice1.startrpt
    ElseIf LNType = 9 Then
      ARptLglLateNotice17.GetName ReportFile$
      ARptLglLateNotice17.startrpt
    End If
  Else
    If Not LNType = 1 And Not LNType = 98 Then
      ViewPrint ReportFile$, "Late Notice Printing", False, , True, MaskNotice$
    ElseIf Not LNType = 98 Then
      ViewPrint ReportFile$, "Late Notice Printing"
    ElseIf LNType = 98 Then
      MsgBox "Late Notice File Created - " + ReportFile$, vbOKOnly
      
    End If
  End If
 'endif
  
  GoTo ExitLatePrint

PrintThemOne:
   If PSAFlag Then
     If fpcboActiveOnly.ListIndex = 0 Then
       If UBCustRec(1).Status <> "A" Then
         GoTo NoPSAPrint
       End If
     ElseIf fpcboActiveOnly.ListIndex = 1 Then
       If UBCustRec(1).Status <> "I" Then
         GoTo NoPSAPrint
       End If
     End If
   End If
   Select Case LNType
    Case 1:  'Letter Format
      GoSub PrintLetterFormat
    Case 2:
      GoSub PrintNewStandV1
    Case 3:
      GoSub PrintNewStandBar
    Case 4:
      GoSub PrnStand21Line
    Case 5:
      GoSub PrintNewStandRmStamp
    Case 6:
      GoSub PrnStand24L2Bx
    Case 7:
      GoSub PrnStand24L3Bx
    Case 8:
      GoSub PrintLaserNotice1
      Grpt = True
    Case 9:
      GoSub PrnLaserLegalLN
      Grpt = True
    Case 98:
      GoSub CreateSCSFileTransfer
    Case Else
   End Select
   
   PrnCnt = PrnCnt + 1
  Put UBLRec, , AcctNo&

NoPSAPrint:
Return

PrintLetterFormat:   '1 'this can be text or graphic
  Print #UBRpt, " "
  Print #UBRpt, " "
  Print #UBRpt, " "
  Print #UBRpt, " "
  Print #UBRpt, " "
  Print #UBRpt, " "
  Print #UBRpt, Tab(20); latelet(1).Head1
  Print #UBRpt, Tab(20); latelet(1).Head2
  Print #UBRpt, Tab(20); latelet(1).Head3
  Print #UBRpt, Tab(20); latelet(1).Head4
  Print #UBRpt, Tab(20); latelet(1).Head5
  Print #UBRpt, " "
  Print #UBRpt, Tab(5); QPTrim(UBCustRec(1).CustName)
  Print #UBRpt, Tab(5); QPTrim(UBCustRec(1).ADDR1)
  Print #UBRpt, Tab(5); QPTrim(UBCustRec(1).ADDR2)
  Print #UBRpt, Tab(5); QPTrim$(UBCustRec(1).CITY); ", "; UBCustRec(1).STATE; " "; UBCustRec(1).ZIPCODE
  Print #UBRpt, " "
  Print #UBRpt, Tab(5); "Notice Date: " + NDate$; Tab(42); "Due Date: " + PDate$
  Print #UBRpt, " "
  For cntll = 1 To 10
    Print #UBRpt, Tab(5); RTrim(latelet(1).Body(cntll))
  Next
  Print #UBRpt, " "
  Print #UBRpt, Tab(5); "   Customer Account: "; Acct$; Tab(42); "  Prev: "; Using("$###,###.##", UBCustRec(1).PrevBalance)
  Print #UBRpt, Tab(5); "           Location: "; Location$; Tab(42); "  Curr: "; Using("$###,###.##", UBCustRec(1).CurrBalance)
  Print #UBRpt, Tab(42); " Total: "; Using("$###,###.##", TotalBal#)
  Print #UBRpt, " "
  For cntll = 11 To 20
    Print #UBRpt, Tab(5); RTrim(latelet(1).Body(cntll))
  Next
  For cntll = 40 To MaxLines
    Print #UBRpt, " "
  Next
  Print #UBRpt, FF$

Return

PrintNewStandV1:    '2
'New Utility Bill format 10-28-96 BAR CODE PRINTABLE
'
'MUST SHOW BOTH METERS OR, TOTAL CONSUMPTION ON THIS BILL
'1-7 revs listed, 8-15 total as other
    If Not LPIFlag Then
      LPIFlag = -2
    If OkiMode = 1 Then
      Print #UBRpt, Chr$(27); Chr$(48); Chr$(27); Chr$(77);
    Else
      Print #UBRpt, Chr$(27); Chr$(48); Chr$(27); Chr$(58);
    End If
'        Print #UBRpt, Chr$(27); Chr$(48); Chr$(27); Chr$(58);
'     ' Print #UBRpt, Chr$(27); Chr$(48); Chr$(27); Chr$(77);
'      ' put printer in     8 lpi             12 cpi  oki mode
    End If

    AcctNum = AcctNo&
    Acct$ = QPTrim$(Str$(AcctNum))
    Select Case AcctNum
    Case Is < 10
      Acct$ = "00" + Acct$
    Case Is < 100
      Acct$ = "0" + Acct$
    End Select
    AcctLen = Len(Acct$)
    Previous# = UBCustRec(1).PrevBalance
    Current# = UBCustRec(1).CurrBalance
    Totalamt# = Round#(Previous# + Current#)

    Print #UBRpt, "~"  '; TAB(30); USING "########"; FBillNO& + PrintedCnt
    Print #UBRpt,
    Print #UBRpt,
    Print #UBRpt,
    Print #UBRpt,
    Print #UBRpt,
    Print #UBRpt,
    Print #UBRpt, Using("##########", AcctNo&);
    Print #UBRpt, Tab(15); Left$(UBCustRec(1).ServAddr, 19); Tab(50); Using("########", AcctNum);
    Print #UBRpt, Tab(62); NDate$
    Print #UBRpt,

    Print #UBRpt, Tab(50); PDate$; Tab(64); Using("#####.##", Abs(Totalamt#))
    Print #UBRpt, Tab(3); NDate$; 'TAB(15); PrevDate$; TAB(26); DateRead$;
    Print #UBRpt,
    Print #UBRpt,
    Print #UBRpt,
    Print #UBRpt,
    Print #UBRpt,
    Print #UBRpt,
    PCnt = 0
    For WRevCnt = 1 To 7
      PCnt = PCnt + 1
      If UBCustRec(1).CurrRevAmts(WRevCnt) <> 0 Then
        Print #UBRpt, " "; Left$(UBSetUpRec(1).Revenues(WRevCnt).RevName, 5);
        Print #UBRpt, Tab(36); Using("#####.##", UBCustRec(1).CurrRevAmts(WRevCnt));
      End If
      Select Case PCnt
      Case 2
    'print bar code acct
    ')))))))))))))))))))))))
        If AcctBar = True Then
    '*************For Okidata to print Bar code
          Print #UBRpt, Tab(47); Chr$(27); Chr$(16); "A";
          Print #UBRpt, Chr$(8);
          Print #UBRpt, Chr$(2); "0";
          Print #UBRpt, "0"; Chr$(2);
          Print #UBRpt, Chr$(1); Chr$(1);
          Print #UBRpt, Chr$(1); Chr$(2);
          Print #UBRpt, Chr$(27); Chr$(16); "B"; Chr$(AcctLen); Acct$
    '**************************
        Else
          Print #UBRpt, " "
        End If
    '))))))))))))))))))))))))))
      Case 4
        Print #UBRpt, Tab(47); Left$(UBCustRec(1).CustName, 29)
      Case 5
        Print #UBRpt, Tab(47); QPTrim$(UBCustRec(1).ADDR1)
      Case 6
        Print #UBRpt, Tab(47); QPTrim$(UBCustRec(1).ADDR2)
      Case 7
        Print #UBRpt, Tab(47); Left$(UBCustRec(1).CITY, 14); " "; UBCustRec(1).STATE; " "; Left$(UBCustRec(1).ZIPCODE, 5)
      Case Else
        Print #UBRpt,
      End Select
    Next
    tmprev# = 0
    For WRevCnt = 8 To 15
      If UBCustRec(1).CurrRevAmts(WRevCnt) <> 0 Then
        tmprev# = tmprev# + UBCustRec(1).CurrRevAmts(WRevCnt)
      End If
    Next
    If tmprev# > 0 Then
      Print #UBRpt, " Other";
      Print #UBRpt, Tab(36); Using("#####.##", tmprev#)
    Else
    'Zip$ = QPTrim$(UBCustRec(1).ZipCode) + "@"
    'Ziplen = LEN(Zip$)
    'PRINT #UBRpt, STRING$(47, " "); CHR$(27); CHR$(10); CHR$(67); CHR$(10); Zip$
     Print #UBRpt,
    End If
    'Zip$ = QPTrim$(LEFT$(UBCustRec(1).ZipCode, 5))
    'Ziplen = LEN(Zip$)
    'PRINT #UBRpt, STRING$(47, " "); CHR$(27); CHR$(16); "A";
    'PRINT #UBRpt, CHR$(8);
    'PRINT #UBRpt, CHR$(2); CHR$(0);           '
    'PRINT #UBRpt, CHR$(0); CHR$(2);           'Line 12
    'PRINT #UBRpt, CHR$(1); CHR$(1);           '
    'PRINT #UBRpt, CHR$(1); CHR$(2);
    'PRINT #UBRpt, CHR$(27); CHR$(16); "B"; CHR$(Ziplen); Zip$

'    If Previous# <> 0 Then
'      Print #UBRpt, "                  Previous:  "; Using("$###,###.##", Previous#)
'    Else
      Print #UBRpt, " "
'    End If
'    Print #UBRpt, "                   Current:  "; Using("$###,###.##", Current#)
    Print #UBRpt, " "
    Print #UBRpt, "                              --------------"
    Print #UBRpt,
    Print #UBRpt, "                      Past Due:  "; Using("$###,###.##", Totalamt#)

    Print #UBRpt, "  "; fptxtMessage(0); Tab(60); UBCustRec(1).Book; "-"; UBCustRec(1).SEQNUMB
    Print #UBRpt, "  "; fptxtMessage(1)
    Print #UBRpt, "  "; fptxtMessage(2);
    If Rteflag Then
      Print #UBRpt, Tab(60); "Route: "; QPTrim$(UBCustRec(1).POSTRTE)
    Else
      Print #UBRpt, " "
    End If
    Print #UBRpt, "  "; fptxtMessage(3)
    Print #UBRpt, "~"
Return

PrintNewStandBar:  '3
'Hamlet Bill format 01-28-99 BAR CODE PRINTABLE
'1st thru 6th rev will list 7-15 total as other
    Zip$ = QPTrim$(UBCustRec(1).ZIPCODE)
    PZip$ = Zip$
    PZip$ = Left$(PZip$, 5) + "-" + Mid$(PZip$, 6)
    ZDigit$ = GetZipEDigit$(Zip$)
    Zip$ = Zip$ + ZDigit$
    CustMsg$ = QPTrim$(UBCustRec(1).BILLCMNT)

'    IF NOT LPIFlag THEN
'      LPIFlag = -2
'    If OkiMode = 1 Then
'      Print #UBRpt, Chr$(27); Chr$(48); Chr$(27); Chr$(77);
'    Else
'      Print #UBRpt, Chr$(27); Chr$(48); Chr$(27); Chr$(58);
'    End If
'    END IF
    If Not LPIFlag Then
      LPIFlag = -2
      If InStr(TOWNNAME$, "PEACHLAND") Then
        Print #UBRpt, Chr$(27); Chr$(48); Chr$(27); Chr$(77);
      Else
        Print #UBRpt, Chr$(27); Chr$(48); Chr$(27); Chr$(58);
      'PRINT #UBRpt, CHR$(27); CHR$(48); CHR$(27); CHR$(77);
      ' put printer in     8 lpi             12 cpi  oki mode
      End If
    End If                                                 ':  M

'    AcctNum = UBCustRec(1).CustAcctNo
    Previous# = UBCustRec(1).PrevBalance
    Current# = UBCustRec(1).CurrBalance
    Totalamt# = Round#(Previous# + Current#)

    Acct$ = QPTrim$(UBCustRec(1).ZIPCODE)
    AcctLen = Len(Acct$)
    Print #UBRpt, "~" '; Tab(50); Using("########", FBillNO& + PrintedCnt)
    Print #UBRpt, " "
    Print #UBRpt, " "
    Print #UBRpt, " "
    Print #UBRpt, " "
    Print #UBRpt, " "
    Print #UBRpt, " "

    MPCnt = 1
    PCnt = 0
    For WRevCnt = 1 To 6
      PCnt = PCnt + 1
      If UBCustRec(1).CurrRevAmts(WRevCnt) <> 0 Then
        Print #UBRpt, Left$(UBSetUpRec(1).Revenues(WRevCnt).RevName, 3);
      End If
      If UBCustRec(1).CurrRevAmts(WRevCnt) <> 0 Then
        Print #UBRpt, Tab(33); Using("#####.##", UBCustRec(1).CurrRevAmts(WRevCnt));
      End If
      Select Case PCnt
      Case 1
        Print #UBRpt, Tab(44); Using("##########", AcctNo&)
      Case 5
        Print #UBRpt, Tab(49); Left$(UBCustRec(1).ServAddr, 26)
      Case Else
        Print #UBRpt, " "
      End Select
    Next
    tmprev# = 0
    For WRevCnt = 7 To 15
      If UBCustRec(1).CurrRevAmts(WRevCnt) <> 0 Then
        tmprev# = tmprev# + UBCustRec(1).CurrRevAmts(WRevCnt)
      End If
    Next
    If tmprev# > 0 Then
      Print #UBRpt, "Other";
      Print #UBRpt, Tab(33); Using("#####.##", tmprev#)
    Else
      Print #UBRpt, " "
    End If
    Print #UBRpt, " "
    Print #UBRpt, Tab(14); "  Past Due:      "; Using("$###,###.##", Totalamt#);
    Print #UBRpt, Tab(45); NDate$; Tab(60); PDate$
    Print #UBRpt, " "
    Print #UBRpt, Tab(2); fptxtMessage(0)
    Print #UBRpt, Tab(2); fptxtMessage(1)
    Print #UBRpt, Tab(2); fptxtMessage(2); Tab(36); "        "; Using("$###,###.##", Totalamt#)
    Print #UBRpt, Tab(2); fptxtMessage(3)
    Print #UBRpt, " "
    Print #UBRpt, Tab(22); Left$(UBCustRec(1).CustName, 29)
    Print #UBRpt, Tab(22); QPTrim(UBCustRec(1).ADDR1)
    Print #UBRpt, Tab(22); QPTrim(UBCustRec(1).ADDR2)
    Print #UBRpt, Using("##########", AcctNo&);
    Print #UBRpt, Tab(22); Left$(UBCustRec(1).CITY, 14); " "; UBCustRec(1).STATE; " "; PZip$
    Print #UBRpt, " "
    Print #UBRpt, " "
    Print #UBRpt, Using("#######.##", Totalamt#)
    Print #UBRpt, " "
    Print #UBRpt, " "
    Print #UBRpt, Using("#######.##", Totalamt#);
    Print #UBRpt, Tab(22); Chr$(27); Chr$(16); "C"; Chr$(Len(Zip$)); Zip$
    Print #UBRpt, " "
    Print #UBRpt, "~"
Return
PrnStand21Line: '4
'Prints in 10cpi
'Will list 1 thru 5 revenues then 6 to 15 totaled under other
    Zip$ = QPTrim$(UBCustRec(1).ZIPCODE)
    PZip$ = Zip$
    PZip$ = Left$(PZip$, 5) + "-" + Mid$(PZip$, 6)
    'ZDigit$ = GetZipEDigit$(Zip$)
    'Zip$ = Zip$ + ZDigit$
    CustMsg$ = QPTrim$(fptxtMessage(0).Text)

    Previous# = UBCustRec(1).PrevBalance
    Current# = UBCustRec(1).CurrBalance
    Totalamt# = Round#(Previous# + Current#)

    Print #UBRpt, "" '; Using; "#####"; FBillNO& + PrintedCnt
    Print #UBRpt, "~"
    Print #UBRpt, " "
    Print #UBRpt, " "
    Print #UBRpt, " "
    Print #UBRpt, " "
    Print #UBRpt, Tab(17); Left$(NDate$, 2); Tab(22); Mid$(NDate$, 4, 2); Tab(27); Right$(NDate$, 2);
    Print #UBRpt, Tab(40); PDate$
    Print #UBRpt, " " '; Tab(60);
    Print #UBRpt, " "
    Print #UBRpt, Tab(35); Left$(QPTrim$(UBCustRec(1).ServAddr), 24)
    Print #UBRpt, " "
    Print #UBRpt, Tab(34); Left$(QPTrim$(UBCustRec(1).CustName), 25)

    If UBCustRec(1).CurrRevAmts(1) <> 0 Then
      Print #UBRpt, Tab(3); QPTrim(UBSetUpRec(1).Revenues(1).RevName); Tab(20); Tab(20); Using("#####.##", UBCustRec(1).CurrRevAmts(1));
    End If
    Print #UBRpt, Tab(34); Mid$(QPTrim$(UBCustRec(1).ADDR1), 1, 25)

    If UBCustRec(1).CurrRevAmts(2) <> 0 Then
      Print #UBRpt, Tab(3); QPTrim(UBSetUpRec(1).Revenues(2).RevName); Tab(20); Using("#####.##", UBCustRec(1).CurrRevAmts(2));
    End If
    Print #UBRpt, Tab(34); Mid$(QPTrim$(UBCustRec(1).ADDR2), 1, 25)

    If UBCustRec(1).CurrRevAmts(3) <> 0 Then
      Print #UBRpt, Tab(3); QPTrim(UBSetUpRec(1).Revenues(3).RevName);
      Print #UBRpt, Tab(20); Using("#####.##", UBCustRec(1).CurrRevAmts(3));
    End If
    Print #UBRpt, Tab(34); Left$(QPTrim$(UBCustRec(1).CITY), 14); " "; QPTrim(UBCustRec(1).STATE); " "; QPTrim(UBCustRec(1).ZIPCODE)

    If UBCustRec(1).CurrRevAmts(4) <> 0 Then
      Print #UBRpt, Tab(3); QPTrim(UBSetUpRec(1).Revenues(4).RevName);
      Print #UBRpt, Tab(20); Using("#####.##", UBCustRec(1).CurrRevAmts(4));
    End If
    Print #UBRpt, Tab(34); String$(24, "-")

    If UBCustRec(1).CurrRevAmts(5) <> 0 Then
      Print #UBRpt, Tab(3); QPTrim(UBSetUpRec(1).Revenues(5).RevName);
      Print #UBRpt, Tab(20); Using("#####.##", UBCustRec(1).CurrRevAmts(5));
    End If
    Print #UBRpt, Tab(34); CustMsg$
    tmprev# = 0
    For PCnt = 6 To 15
      If UBCustRec(1).CurrRevAmts(PCnt) <> 0 Then
        tmprev# = tmprev + UBCustRec(1).CurrRevAmts(PCnt)
      End If
    Next
    If tmprev# <> 0 Then
      Print #UBRpt, Tab(3); "Other";
      Print #UBRpt, Tab(20); Using("#####.##", tmprev#)
    Else
      Print #UBRpt, " "
    End If
    Print #UBRpt, " "
    Print #UBRpt, " "
    Print #UBRpt, Tab(5); Acct$; Tab(19); Using("######.##", Totalamt#);
    Print #UBRpt, Tab(37); Acct$; Tab(50); Using("######.##", Totalamt#)
    'Print #UBRpt, "~" 'Per Dale
Return
PrintNewStandRmStamp: '5
    If Not LPIFlag Then
      LPIFlag = -2
    If OkiMode = 1 Then
      Print #UBRpt, Chr$(27); Chr$(48); Chr$(27); Chr$(77);
    Else
      Print #UBRpt, Chr$(27); Chr$(48); Chr$(27); Chr$(58);
    End If
    End If
    Zip$ = QPTrim$(UBCustRec(1).ZIPCODE)
    PZip$ = Zip$
    PZip$ = Left$(PZip$, 5) + "-" + Mid$(PZip$, 6)
    ZDigit$ = GetZipEDigit$(Zip$)
    Zip$ = Zip$ + ZDigit$
    Previous# = UBCustRec(1).PrevBalance
    Current# = UBCustRec(1).CurrBalance
    Totalamt# = Round#(Previous# + Current#)


'*** Look for meter readings
'  PrevRead& = 0
'  CurrRead& = 0
'  UsageAmt& = 0
'  DidWMrt = False
'  For WMtrCnt = 1 To 7
'    Select Case UBBillRec(1).MtrTypes(WMtrCnt)
'    Case MtrWaterOnly, MtrSewerOnly, MtrCombined, MtrTouchRead
'      If UBBillRec(1).PrevRead(WMtrCnt) < 0 Then
'        UBBillRec(1).PrevRead(WMtrCnt) = 0
'      End If
'      If UBBillRec(1).CurRead(WMtrCnt) < 0 Then
'        UBBillRec(1).CurRead(WMtrCnt) = 0
'      End If
'      PrevRead& = UBBillRec(1).PrevRead(WMtrCnt)
'      CurrRead& = UBBillRec(1).CurRead(WMtrCnt)
'      UsageAmt& = CurrRead& - PrevRead&
'      If UsageAmt& < 0 Then
'        MaxMeterAmt& = 10& ^ (Len(Str$(UBBillRec(1).PrevRead(WMtrCnt))) - 1)
'        UsageAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(WMtrCnt)) + UBBillRec(1).CurRead(WMtrCnt)
'      End If
'      Exit For
'    End Select
'  Next
'**** Find a meter

'    Zero$ = "0"

'    AcctNum = UBBillRec(1).CustAcctNo
'    Previous# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
'    Totalamt# = Round#(Previous# + UBBillRec(1).TransAmt)

'    If FinalFlag And CDeposit# Then
'      Totalamt# = Round#(Totalamt# - UBCustRec(1).DepositAmt)
'    End If

'    If Totalamt# > 0 Then
'      'TenPct# = 10
'      TenPct# = 2 'Round#(TotalAmt# * .1)
'    Else
'      TenPct# = 0
'    End If

'    AcctNum = UBBillRec(1).CustAcctNo
    Acct$ = QPTrim$(Str$(AcctNum))
    Select Case AcctNum
    Case Is < 10
      Acct$ = "00" + Acct$
    Case Is < 100
      Acct$ = "0" + Acct$
    End Select
    AcctLen = Len(Acct$)

    Print #UBRpt, " " 'Tab(50); Using; "########"; FBillNO& + PrintedCnt
    Print #UBRpt, " "
    Print #UBRpt, " "
    Print #UBRpt, " "
    Print #UBRpt, " "
    Print #UBRpt, " "
    Print #UBRpt, " "
    Print #UBRpt, " "
    Print #UBRpt, " "
    Print #UBRpt, Using("##########", AcctNum);
    Print #UBRpt, Tab(15); Left$(QPTrim$(UBCustRec(1).ServAddr), 19); Tab(50); Using("########", AcctNum);
    Print #UBRpt, Tab(62); NDate$
    Print #UBRpt, " "

    Print #UBRpt, Tab(50); PDate$; Tab(64); Using("#####.##", Totalamt#)
    Print #UBRpt, Tab(3); NDate$; 'Tab(15); PrevDate$; Tab(26); DateRead$;
     'Only Print Days if Greater than 0
'     If DaysINRead > 0 Then
'       Print #UBRpt, Tab(40); Using; "####"; DaysINRead
'     Else
       Print #UBRpt, " "
'     End If

    Print #UBRpt, Tab(50); PDate$; Tab(64); Using("#####.##", Totalamt#) ' + TenPct#)

    Print #UBRpt, " "
    Print #UBRpt, " "
    Print #UBRpt, " "
    Print #UBRpt, " "
    PCnt = 0
    For WRevCnt = 1 To 7
      PCnt = PCnt + 1
      If UBCustRec(1).CurrRevAmts(WRevCnt) <> 0 Then
        Print #UBRpt, " "; Left$(UBSetUpRec(1).Revenues(WRevCnt).RevName, 5);
        Print #UBRpt, Tab(36); Using("#####.##", UBCustRec(1).CurrRevAmts(WRevCnt));
      End If
      Select Case PCnt
      Case 4
        Print #UBRpt, Tab(47); Left$(UBCustRec(1).CustName, 29)
      Case 5
        Print #UBRpt, Tab(47); QPTrim$(UBCustRec(1).ADDR1)
      Case 6
        Print #UBRpt, Tab(47); QPTrim$(UBCustRec(1).ADDR2)
      Case 7
        Print #UBRpt, Tab(47); Left$(UBCustRec(1).CITY, 14); " "; UBCustRec(1).STATE; " "; Left$(UBCustRec(1).ZIPCODE, 5)
      Case Else
        Print #UBRpt, " "
      End Select
    Next
    tmprev# = 0
    For WRevCnt = 8 To 15

      If UBCustRec(1).CurrRevAmts(PCnt) <> 0 Then
        tmprev# = tmprev + UBCustRec(1).CurrRevAmts(PCnt)
      End If
    Next
    If tmprev# <> 0 Then
      Print #UBRpt, " Other"; 'Tab(3);
      Print #UBRpt, Tab(36); Using("#####.##", tmprev#)
    Else
      Print #UBRpt, " "
    End If

'    If TotalTax# > 0 Then
'       Print #UBRpt, "                       TAX:  "; Using; "$$,######.##"; TotalTax#
'    Else
      Print #UBRpt, " "
'    End If

'    If Previous# <> 0 Then
'      Print #UBRpt, "                  Previous:  "; Using; "$$,######.##"; Previous#
'    Else
      Print #UBRpt, " "
'    End If
'      Print #UBRpt, "                   Current:  "; Using; "$$,######.##"; UBBillRec(1).TransAmt
'    Print #UBRpt, "                           --------------"

'    If FinalFlag And CDeposit# Then
'      Print #UBRpt, "                   Deposit:  "; Using; "$$,######.##"; -UBCustRec(1).DepositAmt
'    Else
      Print #UBRpt, " "
'    End If

'    If Totalamt# < 0 And FinalFlag Then
'      Print #UBRpt, "                Refund Due:  "; Using; "$$,######.##"; Abs(Totalamt#)
'    Else
      'STOP
      Print #UBRpt, "                      Past Due:  "; Using("$###,###.##", Totalamt#)
'    End If

    Print #UBRpt, " "
    Print #UBRpt, "  "; QPTrim$(fptxtMessage(0).Text); " "; QPTrim$(fptxtMessage(1).Text)
    Print #UBRpt, "  "; QPTrim$(fptxtMessage(2).Text); " "; QPTrim$(fptxtMessage(3).Text)
    'Print #UBRpt, " "
    Print #UBRpt, "~"
Return
PrnStand24L2Bx:  '6
  GoSub PrnStand24L3Bx
Return
PrnStand24L3Bx:   '7
'    If OkiMode = 1 Then
'      Print #UBRpt, Chr$(27); Chr$(48); Chr$(27); Chr$(77);
'    Else
'      Print #UBRpt, Chr$(27); Chr$(48); Chr$(27); Chr$(58);
'    End If
    Zip$ = QPTrim$(UBCustRec(1).ZIPCODE)
    PZip$ = Zip$
    PZip$ = Left$(PZip$, 5) + "-" + Mid$(PZip$, 6)
    ZDigit$ = GetZipEDigit$(Zip$)
    Zip$ = Zip$ + ZDigit$
    Previous# = UBCustRec(1).PrevBalance
    Current# = UBCustRec(1).CurrBalance
    Totalamt# = Round#(Previous# + Current#)

'    If Totalamt# > 0 Then
'      Bucks2# = 2
'    Else
'      Bucks2# = 0
'    End If

    Print #UBRpt, "~" '; Using; "#####"; FBillNO& + PrintedCnt
    Print #UBRpt, " "
    Print #UBRpt, " "
    Print #UBRpt, " "
    Print #UBRpt, Tab(3); Left$(NDate$, 2); Tab(8); Mid$(NDate$, 4, 2); Tab(13); Right$(NDate$, 2);
    Print #UBRpt, Tab(17); Left$(PDate$, 2); Tab(22); Mid$(PDate$, 4, 2); Tab(27); Right$(PDate$, 2);
    Print #UBRpt, " "
    Print #UBRpt, Tab(40); PDate$
    Print #UBRpt, " "
    Print #UBRpt, " "
    'Print #UBRpt, Tab(2); Using; "#########"; UBBillRec(1).PrevRead(1);
    'Print #UBRpt, Tab(12); Using; "#########"; UBBillRec(1).CurRead(1);
    'Print #UBRpt, Tab(22); Using; "########"; UsageAmt&;

    Print #UBRpt, Tab(35); Left$(QPTrim$(UBCustRec(1).ServAddr), 24)
    Print #UBRpt, " "
    Print #UBRpt, " "
    Print #UBRpt, Tab(34); Left$(QPTrim$(UBCustRec(1).CustName), 25)
    PCnt = 0
    For WRevCnt = 1 To 4
      PCnt = PCnt + 1
      If UBCustRec(1).CurrRevAmts(WRevCnt) <> 0 Then
        Print #UBRpt, Tab(3); Left$(UBSetUpRec(1).Revenues(WRevCnt).RevName, 5);
        Print #UBRpt, Tab(20); Using("#####.##", UBCustRec(1).CurrRevAmts(WRevCnt));
      End If
      Select Case PCnt
      Case 1
        Print #UBRpt, Tab(34); Left$(QPTrim(UBCustRec(1).ADDR1), 25)
      Case 2
        Print #UBRpt, Tab(34); Left$(QPTrim(UBCustRec(1).ADDR2), 25)
      Case 3
        Print #UBRpt, Tab(34); Left$(QPTrim(UBCustRec(1).CITY), 14); " "; QPTrim(UBCustRec(1).STATE); " "; Zip$
      Case Else
        Print #UBRpt, " "
      End Select
    Next
    tmprev# = 0
    For WRevCnt = 5 To 15
      If UBCustRec(1).CurrRevAmts(WRevCnt) > 0 Then
        tmprev# = tmprev + UBCustRec(1).CurrRevAmts(WRevCnt)
      End If
    Next
    If tmprev# <> 0 Then
      Print #UBRpt, Tab(3); "Other";
      Print #UBRpt, Tab(20); Using("#####.##", tmprev#)
    Else
      Print #UBRpt, " "
    End If

'    If Previous# <> 0 Then
'      Print #UBRpt, Tab(3); "Previous:"; Tab(20); Using; "######.##"; Previous#
'    Else
      Print #UBRpt, " "
'    End If

'    If FinalFlag And CDeposit# Then
'      Print #UBRpt, Tab(4); "Deposit:"; Tab(20); Using; "######.##"; -UBCustRec(1).DepositAmt
'      Totalamt# = Round#(Totalamt# - UBCustRec(1).DepositAmt)
'    Else
      Print #UBRpt, " "
'    End If

    Print #UBRpt, " "
    Print #UBRpt, Tab(5); AcctNo&; Tab(20); Using("#####.##", Totalamt#);
    If LNType = 6 Then
      Print #UBRpt, Tab(37); AcctNo&; Tab(50); Using("#####.##", Totalamt#)
    Else
      Print #UBRpt, Tab(35); AcctNo&; Tab(42); Using("#####.##", Totalamt#) '; Tab(51); Round#(Totalamt# + Bucks2#)
    End If
    Print #UBRpt, " "; QPTrim$(fptxtMessage(0).Text); " "; QPTrim$(fptxtMessage(1).Text)
    Print #UBRpt, " "; QPTrim$(fptxtMessage(2).Text); " "; QPTrim$(fptxtMessage(3).Text)
    Print #UBRpt, "~"
Return
PrintLaserNotice1:
  LNcnt = LNcnt + 1
  
    Zip$ = QPTrim$(UBCustRec(1).ZIPCODE)
    PZip$ = Zip$
    PZip$ = Left$(PZip$, 5) + "-" + Mid$(PZip$, 6)
    ZDigit$ = GetZipEDigit$(Zip$)
    Zip$ = Zip$ + ZDigit$
    
    Previous# = UBCustRec(1).PrevBalance
    Current# = UBCustRec(1).CurrBalance
    Totalamt# = Round#(Previous# + Current#)
    ToPrint$ = ToPrint$ + Acct$ + "~" + Left$(QPTrim$(UBCustRec(1).CustName), 25)
    ToPrint$ = ToPrint$ + "~" + Left$(NDate$, 2) + "~" + Mid$(NDate$, 4, 2) + "~" + Right$(NDate$, 2)
    ToPrint$ = ToPrint$ + "~" + Left$(PDate$, 2) + "~" + Mid$(PDate$, 4, 2) + "~" + Right$(PDate$, 2)
    ToPrint$ = ToPrint$ + "~" + PDate$
    ToPrint$ = ToPrint$ + "~" + Left$(QPTrim$(UBCustRec(1).ServAddr), 24)
    ToPrint$ = ToPrint$ + "~" + Mid$(QPTrim$(UBCustRec(1).ADDR1), 1, 25)
    ToPrint$ = ToPrint$ + "~" + Mid$(QPTrim$(UBCustRec(1).ADDR2), 1, 25)
    ToPrint$ = ToPrint$ + "~" + Left$(QPTrim$(UBCustRec(1).CITY), 14) + " " + QPTrim(UBCustRec(1).STATE) + " " + Zip$

    If UBCustRec(1).CurrRevAmts(1) <> 0 Then
      ToPrint$ = ToPrint$ + "~" + QPTrim(UBSetUpRec(1).Revenues(1).RevName) + "~" + Using("#####.##", UBCustRec(1).CurrRevAmts(1))
    Else
      ToPrint$ = ToPrint$ + "~ ~ "
    End If
    If UBCustRec(1).CurrRevAmts(2) <> 0 Then
      ToPrint$ = ToPrint$ + "~" + QPTrim(UBSetUpRec(1).Revenues(2).RevName) + "~" + Using("#####.##", UBCustRec(1).CurrRevAmts(2))
    Else
      ToPrint$ = ToPrint$ + "~ ~ "
    End If
    If UBCustRec(1).CurrRevAmts(3) <> 0 Then
      ToPrint$ = ToPrint$ + "~" + QPTrim(UBSetUpRec(1).Revenues(3).RevName) + "~" + Using("#####.##", UBCustRec(1).CurrRevAmts(3))
    Else
      ToPrint$ = ToPrint$ + "~ ~ "
    End If
    If UBCustRec(1).CurrRevAmts(4) <> 0 Then
      ToPrint$ = ToPrint$ + "~" + QPTrim(UBSetUpRec(1).Revenues(4).RevName) + "~" + Using("#####.##", UBCustRec(1).CurrRevAmts(4))
    Else
      ToPrint$ = ToPrint$ + "~ ~ "
    End If
    If UBCustRec(1).CurrRevAmts(5) <> 0 Then
      ToPrint$ = ToPrint$ + "~" + QPTrim(UBSetUpRec(1).Revenues(5).RevName) + "~" + Using("#####.##", UBCustRec(1).CurrRevAmts(5))
    Else
      ToPrint$ = ToPrint$ + "~ ~ "
    End If
    If UBCustRec(1).CurrRevAmts(6) <> 0 Then
      ToPrint$ = ToPrint$ + "~" + QPTrim(UBSetUpRec(1).Revenues(6).RevName) + "~" + Using("#####.##", UBCustRec(1).CurrRevAmts(6))
    Else
      ToPrint$ = ToPrint$ + "~ ~ "
    End If
    tmprev# = 0
    For PCnt = 7 To 15
      If UBCustRec(1).CurrRevAmts(PCnt) <> 0 Then
        tmprev# = tmprev + UBCustRec(1).CurrRevAmts(PCnt)
      End If
    Next
    If tmprev# <> 0 Then
      ToPrint$ = ToPrint$ + "~" + "Other" + "~" + Using("#####.##", tmprev#)
    Else
      ToPrint$ = ToPrint$ + "~ ~ "
    End If
    ToPrint$ = ToPrint$ + "~" + Using("#######.##", Totalamt#)
    ToPrint$ = ToPrint$ + "~" + QPTrim$(fptxtMessage(0).Text) + "~" + QPTrim$(fptxtMessage(1).Text) + "~"
Dblcheck:
    If LNcnt = 3 Then
      Print #UBRpt, ToPrint$
      ToPrint$ = ""
      LNcnt = 0
    ElseIf LNcnt = 1 And endit = True Then
      ToPrint$ = ToPrint$ + ToPrint2$ + ToPrint2$
      Print #UBRpt, ToPrint$
      ToPrint$ = ""
      LNcnt = 0
    ElseIf LNcnt = 2 And endit = True Then
      ToPrint$ = ToPrint$ + ToPrint2$
      Print #UBRpt, ToPrint$
      ToPrint$ = ""
      LNcnt = 0
    End If
Return
PrnLaserLegalLN:
'Late Notice Laser(Uses Blank Stock)BAR CODE PRINTABLE
  Zip$ = QPTrim$(UBCustRec(1).ZIPCODE)
  ZDigit$ = GetZipEDigit$(Zip$)
  Zip$ = Zip$ + ZDigit$
  Previous# = UBCustRec(1).PrevBalance
  Current# = UBCustRec(1).CurrBalance
  Totalamt# = Round#(Previous# + Current#)

  AcctNum = AcctNo&
  Acct$ = QPTrim$(Str$(AcctNum))
  Select Case AcctNum
  Case Is < 10
    Acct$ = "00" + Acct$
  Case Is < 100
    Acct$ = "0" + Acct$
  End Select
  ToPrint$ = Acct$   'Using("########",'(FBillNO& + PrintedCnt))
  ToPrint$ = ToPrint$ + "~" + NDate$ + "~" + PDate$
    If UBCustRec(1).CurrRevAmts(1) <> 0 Then
      ToPrint$ = ToPrint$ + "~" + QPTrim(UBSetUpRec(1).Revenues(1).RevName) + "~" + Using("#####.##", UBCustRec(1).CurrRevAmts(1))
    Else
      ToPrint$ = ToPrint$ + "~ ~ "
    End If
    If UBCustRec(1).CurrRevAmts(2) <> 0 Then
      ToPrint$ = ToPrint$ + "~" + QPTrim(UBSetUpRec(1).Revenues(2).RevName) + "~" + Using("#####.##", UBCustRec(1).CurrRevAmts(2))
    Else
      ToPrint$ = ToPrint$ + "~ ~ "
    End If
    If UBCustRec(1).CurrRevAmts(3) <> 0 Then
      ToPrint$ = ToPrint$ + "~" + QPTrim(UBSetUpRec(1).Revenues(3).RevName) + "~" + Using("#####.##", UBCustRec(1).CurrRevAmts(3))
    Else
      ToPrint$ = ToPrint$ + "~ ~ "
    End If
    If UBCustRec(1).CurrRevAmts(4) <> 0 Then
      ToPrint$ = ToPrint$ + "~" + QPTrim(UBSetUpRec(1).Revenues(4).RevName) + "~" + Using("#####.##", UBCustRec(1).CurrRevAmts(4))
    Else
      ToPrint$ = ToPrint$ + "~ ~ "
    End If
    If UBCustRec(1).CurrRevAmts(5) <> 0 Then
      ToPrint$ = ToPrint$ + "~" + QPTrim(UBSetUpRec(1).Revenues(5).RevName) + "~" + Using("#####.##", UBCustRec(1).CurrRevAmts(5))
    Else
      ToPrint$ = ToPrint$ + "~ ~ "
    End If
    If UBCustRec(1).CurrRevAmts(6) <> 0 Then
      ToPrint$ = ToPrint$ + "~" + QPTrim(UBSetUpRec(1).Revenues(6).RevName) + "~" + Using("#####.##", UBCustRec(1).CurrRevAmts(6))
    Else
      ToPrint$ = ToPrint$ + "~ ~ "
    End If
    If UBCustRec(1).CurrRevAmts(7) <> 0 Then
      ToPrint$ = ToPrint$ + "~" + QPTrim(UBSetUpRec(1).Revenues(7).RevName) + "~" + Using("#####.##", UBCustRec(1).CurrRevAmts(7))
    Else
      ToPrint$ = ToPrint$ + "~ ~ "
    End If
'    If UBCustRec(1).CurrRevAmts(8) <> 0 Then
'      ToPrint$ = ToPrint$ + "~" + QPTrim(UBSetUpRec(1).Revenues(8).RevName) + "~" + Using("#####.##", UBCustRec(1).CurrRevAmts(8))
'    Else
'      ToPrint$ = ToPrint$ + "~ ~ "
'    End If
    tmprev# = 0
    For PCnt = 8 To 15
      If UBCustRec(1).CurrRevAmts(PCnt) <> 0 Then
        tmprev# = tmprev + UBCustRec(1).CurrRevAmts(PCnt)
      End If
    Next
    If tmprev# <> 0 Then
      ToPrint$ = ToPrint$ + "~" + "Other" + "~" + Using("#####.##", tmprev#)
    Else
      ToPrint$ = ToPrint$ + "~ ~ "
    End If


    ToPrint$ = ToPrint$ + "~" + QPTrim$(fptxtMessage(0).Text)
    ToPrint$ = ToPrint$ + "~" + QPTrim$(fptxtMessage(1).Text)
    ToPrint$ = ToPrint$ + "~" + QPTrim$(fptxtMessage(2).Text)
    ToPrint$ = ToPrint$ + "~" + QPTrim$(fptxtMessage(3).Text)
  

    ToPrint$ = ToPrint$ + "~" + "TOTAL DUE" + "~" + Using("$#,###,###.##", Totalamt#)
    ToPrint$ = ToPrint$ + "~" + Acct$ 'Using("##########", UBBillRec(1).CustAcctNo)
    ToPrint$ = ToPrint$ + "~" + Left$(UBCustRec(1).ServAddr, 26)
    ToPrint$ = ToPrint$ + "~" + Left$(UBCustRec(1).CustName, 29)
    ToPrint$ = ToPrint$ + "~" + QPTrim(UBCustRec(1).ADDR1)
    ToPrint$ = ToPrint$ + "~" + QPTrim(UBCustRec(1).ADDR2)
    ToPrint$ = ToPrint$ + "~" + Left$(UBCustRec(1).CITY, 14) + " " + UBCustRec(1).STATE + " " + UBCustRec(1).ZIPCODE
'    If FinalFlag Then
'      ToPrint$ = ToPrint$ + "~" + Using("#######.##", Round#(Totalamt#))
'    Else
    If Not Totalamt# > 0 Then
      ToPrint$ = ToPrint$ + "~ "
    Else
      ToPrint$ = ToPrint$ + "~" + Using("$#,###,###.##", Round#(Totalamt#))
    End If
    ToPrint$ = ToPrint$ + "~" + Zip$ + "~" + UBCustRec(1).Book + "-" + UBCustRec(1).SEQNUMB
    Print #UBRpt, ToPrint$
    ToPrint$ = ""
Return
CreateSCSFileTransfer:
  Dim serv As Integer

  Previous# = UBCustRec(1).PrevBalance
  Current# = UBCustRec(1).CurrBalance
  Totalamt# = Round#(Previous# + Current#)

    ReDim PrintRec(1) As BillOutRecType

    PrintRec(1).AcctNo = Using("########", Str$(AcctNo&))
    PrintRec(1).LocationNum = UBCustRec(1).Book + "-" + UBCustRec(1).SEQNUMB
    RSet PrintRec(1).CustName = QPTrim$(UBCustRec(1).CustName)
    RSet PrintRec(1).ADDR1 = QPTrim$(UBCustRec(1).ADDR1)
    RSet PrintRec(1).ADDR2 = QPTrim$(UBCustRec(1).ADDR2)
    RSet PrintRec(1).ServAddr = QPTrim$(UBCustRec(1).ServAddr)
    RSet PrintRec(1).CITY = QPTrim$(UBCustRec(1).CITY)
    RSet PrintRec(1).STATE = QPTrim$(UBCustRec(1).STATE)
    RSet PrintRec(1).ZIPCODE = QPTrim$(Zip$)
    PrintRec(1).BillType = "N"
    PrintRec(1).DepAppAmt = ""
    PrintRec(1).PrevDue = Using(Fmt15$, Str$(Previous#))
    PrintRec(1).CurrDue = Using(Fmt15$, Str$(UBCustRec(1).CurrBalance))
    PrintRec(1).TotalDue = Using(Fmt15$, Str$(Totalamt#))

    For serv = 1 To 15
      PrintRec(1).ServInfo(serv).ServText = QPTrim$(UBSetUpRec(1).Revenues(serv).RevName)
      PrintRec(1).ServInfo(serv).ServAmt = Using(Fmt10a$, Str$(UBCustRec(1).CurrRevAmts(serv)))
    Next
    PrintRec(1).PastDueDate = PDate$
    PrintRec(1).BillDate = NDate$
   RSet PrintRec(1).MsgLine1 = fptxtMessage(0).Text
   RSet PrintRec(1).MsgLine2 = fptxtMessage(1).Text
   RSet PrintRec(1).MsgLine3 = fptxtMessage(2).Text
   RSet PrintRec(1).MsgLine4 = fptxtMessage(3).Text

  PrintRec(1).CrLf = CrLf$
  Put #UBRpt, , PrintRec(1)
   
Return

DoLNMask:
  UBRptA = FreeFile
  MaskNotice$ = UBPath$ + "UBLNA.RPT"
  Open MaskNotice$ For Output As UBRptA
  Select Case LNType
  Case 2
    GoSub PrintNewStandV1Mask
  Case 3
    GoSub PrintNewStandBarMask
  Case 4
    GoSub PrnStand21LineMask
  Case 5
    GoSub PrintNewStandRmStampMask
  Case 6
    GoSub PrnStand24L2BxMask
  Case 7
    GoSub PrnStand24L3BxMask
  Case Else
    'NO MASK
  End Select
  Close UBRptA
Return
PrintNewStandV1Mask:      '2
    If OkiMode = 1 Then
      Print #UBRptA, Chr$(27); Chr$(48); Chr$(27); Chr$(77);
    Else
      Print #UBRptA, Chr$(27); Chr$(48); Chr$(27); Chr$(58);
    End If

    Print #UBRptA, "~"
    Print #UBRptA,
    Print #UBRptA,
    Print #UBRptA,
    Print #UBRptA,
    Print #UBRptA,
    Print #UBRptA,
    Print #UBRptA, "##########";
    Print #UBRptA, Tab(15); "XXXXXXXXXXXXXXXXXXX"; Tab(50); "########";
    Print #UBRptA, Tab(62); "XX/XX/XXXX"
    Print #UBRptA,

    Print #UBRptA, Tab(50); "XX/XX/XXXX"; Tab(64); "#####.##"
    Print #UBRptA, Tab(3); "XX/XX/XXXX";
    Print #UBRptA,
    Print #UBRptA,
    Print #UBRptA,
    Print #UBRptA,
    Print #UBRptA,
    Print #UBRptA,
    PCnt = 0
    For PCnt = 1 To 8
        Print #UBRptA, " "; "XXXXX";
        Print #UBRptA, Tab(36); "#####.##";
     
      Select Case PCnt
      Case 4
        Print #UBRptA, Tab(47); "XXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
      Case 5
        Print #UBRptA, Tab(47); "XXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
      Case 6
        Print #UBRptA, Tab(47); "XXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
      Case 7
        Print #UBRptA, Tab(47); "XXXXXXXXXXXXXX"; " "; "XX"; " "; "XXXXX"
      Case Else
        Print #UBRptA, " "
      End Select
    Next
    Print #UBRptA, " "
    Print #UBRptA, " "
    Print #UBRptA, "                              --------------"
    Print #UBRptA,
    Print #UBRptA, "                      Past Due:  "; "$###,###.##"

    Print #UBRptA, "  "; "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
    Print #UBRptA, "  "; "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
    Print #UBRptA, "  "; "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
    Print #UBRptA, "  "; "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
    Print #UBRptA, "~"
Return
PrintNewStandBarMask:  '3
'    If OkiMode = 1 Then
'      Print #UBRptA, Chr$(27); Chr$(48); Chr$(27); Chr$(77);
'    Else
'      Print #UBRptA, Chr$(27); Chr$(48); Chr$(27); Chr$(58);
'    End If
      If InStr(TOWNNAME$, "PEACHLAND") Then
        Print #UBRptA, Chr$(27); Chr$(48); Chr$(27); Chr$(77);
      Else
        Print #UBRptA, Chr$(27); Chr$(48); Chr$(27); Chr$(58);
      'PRINT #UBRpt, CHR$(27); CHR$(48); CHR$(27); CHR$(77);
      ' put printer in     8 lpi             12 cpi  oki mode
      End If

    Print #UBRptA, "~"
    Print #UBRptA, " "
    Print #UBRptA, " "
    Print #UBRptA, " "
    Print #UBRptA, " "
    Print #UBRptA, " "
    Print #UBRptA, " "

    MPCnt = 1
    PCnt = 0
    For PCnt = 1 To 7
      Print #UBRptA, "XXX";
      Print #UBRptA, Tab(33); "#####.##";
      Select Case PCnt
      Case 1
        Print #UBRptA, Tab(44); "##########"
      Case 5
        Print #UBRptA, Tab(49); "XXXXXXXXXXXXXXXXXXXXXXXXXX"
      Case Else
        Print #UBRptA, " "
      End Select
    Next
    Print #UBRptA, " "
    Print #UBRptA, Tab(14); "  Past Due:      "; "$###,###.##";
    Print #UBRptA, Tab(45); "XX/XX/XXXX"; Tab(60); "XX/XX/XXXX"
    Print #UBRptA, " "
    Print #UBRptA, Tab(2); "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
    Print #UBRptA, Tab(2); "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
    Print #UBRptA, Tab(2); "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"; Tab(36); "        "; "$###,###.##"
    Print #UBRptA, Tab(2); "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
    Print #UBRptA, " "
    Print #UBRptA, Tab(22); "XXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
    Print #UBRptA, Tab(22); "XXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
    Print #UBRptA, Tab(22); "XXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
    Print #UBRptA, "##########";
    Print #UBRptA, Tab(22); "XXXXXXXXXXXXXX"; " XX  XXXXX"
    Print #UBRptA, " "
    Print #UBRptA, " "
    Print #UBRptA, "#######.##"
    Print #UBRptA, " "
    Print #UBRptA, " "
    Print #UBRptA, "#######.##";
    Print #UBRptA, Tab(22); Chr$(27); Chr$(16); "C"; Chr$(Len("11111")); "11111"
    Print #UBRptA, " "
    Print #UBRptA, "~"
Return
PrnStand21LineMask: '4
    CustMsg$ = "XXXXXXXXXXXXXXXXXXXXXXXXX"

    Print #UBRptA, ""
    Print #UBRptA, "~"
    Print #UBRptA, " "
    Print #UBRptA, " "
    Print #UBRptA, " "
    Print #UBRptA, " "
    Print #UBRptA, Tab(17); "XX"; Tab(22); "XX"; Tab(27); "XX";
    Print #UBRptA, Tab(40); "XX/XX/XXXX"
    Print #UBRptA, " "
    Print #UBRptA, " "
    Print #UBRptA, Tab(35); "XXXXXXXXXXXXXXXXXXXXXXXX"
    Print #UBRptA, " "
    Print #UBRptA, Tab(34); "XXXXXXXXXXXXXXXXXXXXXXXXX"
    Print #UBRptA, Tab(3); "XXXXXXXX"; Tab(20); "#####.##";
    Print #UBRptA, Tab(34); "XXXXXXXXXXXXXXXXXXXXXXXXX"
    Print #UBRptA, Tab(3); "XXXXXXXX"; Tab(20); "#####.##";
    Print #UBRptA, Tab(34); "XXXXXXXXXXXXXXXXXXXXXXXXX"
    Print #UBRptA, Tab(3); "XXXXXXXX"; Tab(20); "#####.##";
    Print #UBRptA, Tab(34); "XXXXXXXXXXXXXX  XX XXXXX"
    Print #UBRptA, Tab(3); "XXXXXXXX"; Tab(20); "#####.##";
    Print #UBRptA, Tab(34); String$(24, "-")
    Print #UBRptA, Tab(3); "XXXXXXXX"; Tab(20); "#####.##";
    Print #UBRptA, Tab(34); CustMsg$
    Print #UBRptA, Tab(3); "XXXXXXXX"; Tab(20); "#####.##"
    Print #UBRptA, " "
    Print #UBRptA, " "
    Print #UBRptA, Tab(5); "XXXXX"; Tab(19); "######.##";
    Print #UBRptA, Tab(37); "XXXXX"; Tab(50); "######.##"
    'Print #UBRpt, "~" 'Per Dale
Return
PrintNewStandRmStampMask: '5
    If OkiMode = 1 Then
      Print #UBRptA, Chr$(27); Chr$(48); Chr$(27); Chr$(77);
    Else
      Print #UBRptA, Chr$(27); Chr$(48); Chr$(27); Chr$(58);
    End If

    Print #UBRptA, " "
    Print #UBRptA, " "
    Print #UBRptA, " "
    Print #UBRptA, " "
    Print #UBRptA, " "
    Print #UBRptA, " "
    Print #UBRptA, " "
    Print #UBRptA, " "
    Print #UBRptA, " "
    Print #UBRptA, "   ######";
    Print #UBRptA, Tab(15); "XXXXXXXXXXXXXXXXXXX"; Tab(50); "   #####";
    Print #UBRptA, Tab(62); "XX/XX/XXXX"
    Print #UBRptA, " "
    Print #UBRptA, Tab(50); "XX/XX/XXXX"; Tab(64); "#####.##"
    Print #UBRptA, Tab(3); "XX/XX/XXXX";
    Print #UBRptA, " "
    Print #UBRptA, Tab(50); "XX/XX/XXXX"; Tab(64); "#####.##"
    Print #UBRptA, " "
    Print #UBRptA, " "
    Print #UBRptA, " "
    Print #UBRptA, " "
    PCnt = 0
    For PCnt = 1 To 8
      Print #UBRptA, " "; "XXXXX";
      Print #UBRptA, Tab(36); "#####.##";
      
      Select Case PCnt
      Case 4
        Print #UBRptA, Tab(47); "XXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
      Case 5
        Print #UBRptA, Tab(47); "XXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
      Case 6
        Print #UBRptA, Tab(47); "XXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
      Case 7
        Print #UBRptA, Tab(47); "XXXXXXXXXXXXXX XX ZZZZZ"
      Case Else
        Print #UBRptA, " "
      End Select
    Next
      Print #UBRptA, " "
      Print #UBRptA, " "
      Print #UBRptA, " "
      Print #UBRptA, "                     Past Due:  "; "$###,###.##"

    Print #UBRptA, " "
    Print #UBRptA, "  "; "XXXXXXXXXXXXXXXXXXXXXXXXX"; " "; "XXXXXXXXXXXXXXXXXXXXXXXXX"
    Print #UBRptA, "  "; "XXXXXXXXXXXXXXXXXXXXXXXXX"; " "; "XXXXXXXXXXXXXXXXXXXXXXXXX"
    Print #UBRptA, "~"
Return
PrnStand24L2BxMask:  '6
  GoSub PrnStand24L3BxMask
Return
PrnStand24L3BxMask:   '7

    Print #UBRptA, "~"
    Print #UBRptA, " "
    Print #UBRptA, " "
    Print #UBRptA, " "
    Print #UBRptA, Tab(3); "XX"; Tab(8); "XX"; Tab(13); "XX";
    Print #UBRptA, Tab(17); "XX"; Tab(22); "XX"; Tab(27); "XX";
    Print #UBRptA, " "
    Print #UBRptA, Tab(40); "XX/XX/XXXX"
    Print #UBRptA, " "
    Print #UBRptA, " "
    'Print #UBRpt, Tab(2); Using; "#########"; UBBillRec(1).PrevRead(1);
    'Print #UBRpt, Tab(12); Using; "#########"; UBBillRec(1).CurRead(1);
    'Print #UBRpt, Tab(22); Using; "########"; UsageAmt&;

    Print #UBRptA, Tab(35); "XXXXXXXXXXXXXXXXXXXXXXXX"
    Print #UBRptA, " "
    Print #UBRptA, " "
    Print #UBRptA, Tab(34); "XXXXXXXXXXXXXXXXXXXXXXXXX"
    PCnt = 0
    For PCnt = 1 To 5
        Print #UBRptA, Tab(3); "XXXXX";
        Print #UBRptA, Tab(20); "#####.##";
      Select Case PCnt
      Case 1
        Print #UBRptA, Tab(34); "XXXXXXXXXXXXXXXXXXXXXXXXX"
      Case 2
        Print #UBRptA, Tab(34); "XXXXXXXXXXXXXXXXXXXXXXXXX"
      Case 3
        Print #UBRptA, Tab(34); "XXXXXXXXXXXXXXX XX XXXXX"
      Case Else
        Print #UBRptA, " "
      End Select
    Next
    Print #UBRptA, " "
    Print #UBRptA, " "
    Print #UBRptA, " "
    Print #UBRptA, Tab(5); "XXXXX"; Tab(20); "#####.##";
    If LNType = 6 Then
      Print #UBRptA, Tab(37); "XXXXX"; Tab(50); "#####.##"
    Else
      Print #UBRptA, Tab(35); "XXXXX"; Tab(42); "#####.##"
    End If
    Print #UBRptA, " "; "XXXXXXXXXXXXXXXXXXXXXXXXX"; " "; "XXXXXXXXXXXXXXXXXXXXXXXXX"
    Print #UBRptA, " "; "XXXXXXXXXXXXXXXXXXXXXXXXX"; " "; "XXXXXXXXXXXXXXXXXXXXXXXXX"
    Print #UBRptA, "~"

Return
ExitLatePrint:

End Sub

