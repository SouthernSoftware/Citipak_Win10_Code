VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmRptCustbyRate 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customers by Rate Listing"
   ClientHeight    =   8868
   ClientLeft      =   3924
   ClientTop       =   2172
   ClientWidth     =   12192
   ClipControls    =   0   'False
   Icon            =   "frmCustbyRateRpt.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   12192
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboRptType 
      Height          =   348
      Left            =   5376
      TabIndex        =   5
      Top             =   5760
      Width           =   1908
      _Version        =   196608
      _ExtentX        =   3365
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
      ColDesigner     =   "frmCustbyRateRpt.frx":08CA
   End
   Begin LpLib.fpCombo fpComboRates 
      Height          =   348
      Left            =   5388
      TabIndex        =   0
      Top             =   3144
      Width           =   3924
      _Version        =   196608
      _ExtentX        =   6921
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
      Columns         =   0
      Sorted          =   0
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   -1
      ColumnWidthScale=   2
      RowHeight       =   -1
      WrapList        =   0   'False
      WrapWidth       =   0
      AutoSearch      =   0
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
      AutoMenu        =   0   'False
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmCustbyRateRpt.frx":0BF8
   End
   Begin LpLib.fpCombo fpCombo1 
      Height          =   348
      Left            =   5376
      TabIndex        =   4
      Top             =   5220
      Width           =   3924
      _Version        =   196608
      _ExtentX        =   6921
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
      Columns         =   0
      Sorted          =   0
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   -1
      ColumnWidthScale=   2
      RowHeight       =   -1
      WrapList        =   0   'False
      WrapWidth       =   0
      AutoSearch      =   0
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
      AutoMenu        =   0   'False
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmCustbyRateRpt.frx":0EEF
   End
   Begin LpLib.fpCombo fpcboCustStatus 
      Height          =   348
      Left            =   5376
      TabIndex        =   3
      Top             =   4680
      Width           =   2004
      _Version        =   196608
      _ExtentX        =   3535
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
      ColDesigner     =   "frmCustbyRateRpt.frx":11E6
   End
   Begin EditLib.fpBoolean fpPrintCode 
      Height          =   348
      Left            =   5388
      TabIndex        =   1
      Top             =   3648
      Width           =   324
      _Version        =   196608
      _ExtentX        =   572
      _ExtentY        =   614
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      AutoToggle      =   -1  'True
      BooleanStyle    =   1
      ToggleFalse     =   "Nn"
      TextFalse       =   "N"
      BooleanPicture  =   0
      AlignPictureH   =   3
      AlignPictureV   =   1
      GroupId         =   0
      GroupTag        =   0
      GroupSelect     =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      MultiLine       =   0   'False
      AlignTextH      =   1
      AlignTextV      =   0
      ToggleTrue      =   "Yy"
      TextTrue        =   "Y"
      Value           =   0
      BooleanMode     =   0
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      BorderGrayAreaColor=   -2147483637
      ToggleGrayed    =   ""
      TextGrayed      =   ""
      AllowMnemonic   =   0   'False
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDOnFocusInvert=   0   'False
      Caption         =   "N"
      ThreeDFrameColor=   -2147483633
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      BooleanDataType =   0
      OLEDropMode     =   0
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
      Left            =   9408
      TabIndex        =   7
      Top             =   7368
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
      Left            =   7752
      TabIndex        =   6
      Top             =   7368
      Width           =   1332
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   336
      Left            =   0
      TabIndex        =   10
      Top             =   8532
      Width           =   12192
      _ExtentX        =   21505
      _ExtentY        =   593
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
            TextSave        =   "1:14 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7133
            TextSave        =   "5/11/2005"
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
   Begin EditLib.fpBoolean fpMtrNum 
      Height          =   348
      Left            =   5400
      TabIndex        =   2
      Top             =   4152
      Width           =   324
      _Version        =   196608
      _ExtentX        =   572
      _ExtentY        =   614
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      AutoToggle      =   -1  'True
      BooleanStyle    =   1
      ToggleFalse     =   "Nn"
      TextFalse       =   "N"
      BooleanPicture  =   0
      AlignPictureH   =   3
      AlignPictureV   =   1
      GroupId         =   0
      GroupTag        =   0
      GroupSelect     =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      MultiLine       =   0   'False
      AlignTextH      =   1
      AlignTextV      =   0
      ToggleTrue      =   "Yy"
      TextTrue        =   "Y"
      Value           =   0
      BooleanMode     =   0
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      BorderGrayAreaColor=   -2147483637
      ToggleGrayed    =   ""
      TextGrayed      =   ""
      AllowMnemonic   =   0   'False
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDOnFocusInvert=   0   'False
      Caption         =   "N"
      ThreeDFrameColor=   -2147483633
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      BooleanDataType =   0
      OLEDropMode     =   0
   End
   Begin VB.Label Label7 
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
      Left            =   3384
      TabIndex        =   15
      Top             =   5244
      Width           =   1788
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Status:"
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
      Left            =   3096
      TabIndex        =   14
      Top             =   4716
      Width           =   2076
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Print Meter Number: "
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
      Left            =   2904
      TabIndex        =   13
      Top             =   4200
      Width           =   2388
   End
   Begin VB.Label Label5 
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
      Left            =   2892
      TabIndex        =   12
      Top             =   5784
      Width           =   2388
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Print Code:"
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
      Left            =   3972
      TabIndex        =   11
      Top             =   3696
      Width           =   1284
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Rate Code:"
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
      Left            =   3948
      TabIndex        =   9
      Top             =   3192
      Width           =   1332
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3210
      Top             =   1368
      Width           =   5772
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Quick Customer by Rate Report"
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
      Left            =   3888
      TabIndex        =   8
      Top             =   1608
      Width           =   4452
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   3660
      Left            =   2688
      Top             =   2832
      Width           =   6828
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   972
      Left            =   3210
      Top             =   1248
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
Attribute VB_Name = "frmRptCustbyRate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim RateRec As Integer
Dim PrintRate As Boolean
Dim ActiveOnly As Boolean
Dim PrnMtrNum As Boolean
Private Sub cmdExit_Click()
  Load frmUBCustMenu
  DoEvents
  frmUBCustMenu.Show
  Unload frmRptCustbyRate
  'LoadDisplayForm frmUBCustMenu, Me
End Sub

Private Sub cmdPrint_Click()
  RateRec = fpComboRates.ListIndex + 1
  PrintRate = fpPrintCode.Text = "Y"
  PrnMtrNum = fpMtrNum.Text = "Y"
  DeActivateControls Me
  If fpcboRptType.ListIndex = 0 Then
    UBCustbyRateRpt2
  ElseIf fpcboRptType.ListIndex = 1 Then
    UBCustbyRateRpt
    ActivateControls Me
    Me.fpComboRates.SetFocus
  Else
    ActivateControls Me
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


Private Sub fpcboRptType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboRptType.ListDown = True
  End If
  If fpcboRptType.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      cmdPrint.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpCombo1.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        UBLog "Closed via RptCustbyRate by " + PWUser$
        CitiTerminate
      End If
    End If
  End If
End Sub

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
  
  Dim RCode As String
  Dim Handle As Integer, cnt As Integer
  Dim UBRateTblRecLen As Integer, NumOfRateRecs As Integer
  ReDim UBRateTblRec(1) As UBRateTblRecType
  
  RCode$ = Space$(10)
  UBRateTblRecLen = Len(UBRateTblRec(1))
  NumOfRateRecs = GetNumRateRecs
  Handle = FreeFile
  Open UBPath$ + "UBRATE.DAT" For Random Shared As Handle Len = UBRateTblRecLen
  For cnt = 1 To NumOfRateRecs
    Get Handle, cnt, UBRateTblRec(1)
    LSet RCode$ = QPTrim$(UBRateTblRec(1).Ratecode)
    fpComboRates.AddItem RCode$ + QPTrim$(UBRateTblRec(1).RATEDESC)
  Next
  Close
  fpComboRates.ListIndex = 0
  fpPrintCode.Text = "N"
  fpMtrNum.Text = "N"
  fpcboCustStatus.AddItem "ALL"
  fpcboCustStatus.AddItem "Active"
  fpcboCustStatus.AddItem "Inactive"
  fpcboCustStatus.AddItem "Balance"
  fpcboCustStatus.AddItem "Pending"
  fpcboCustStatus.AddItem "Delinquent"
  fpcboCustStatus.AddItem "Final"
  fpcboCustStatus.ListIndex = 0

  fpCombo1.AddItem "Location Number Order"
  fpCombo1.AddItem "Account Number Order"
  fpCombo1.AddItem "Customer Name Order"
  fpCombo1.ListIndex = 0

  fpcboRptType.InsertRow = "Graphics"
  fpcboRptType.InsertRow = "Text"
  fpcboRptType.ListIndex = 0

End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
   ' Me.Visible = False
    Temp_Class.ResizeControls Me
   ' Me.Visible = True
   ' Me.SetFocus
  End If
  DoEvents
End Sub

Private Sub UBCustbyRateRpt()
  Dim Dash80 As String, IdxName As String
  Dim Title As String, ReportFile As String
  Dim UBCustRecLen As Integer, IdxRecLen As Integer
  Dim IdxFileSize As Long, cnt As Long
  Dim IdxNumOfRecs As Long, LineCnt As Integer
  Dim Handle As Integer, UBCust As Integer
  Dim UBRateTblRecLen As Integer, NumOfRateRecs As Integer
  Dim UBRpt As Integer, CustCnt As Long
  Dim Active As Long, Final As Long, InActive As Long
  Dim Balance As Long, UnKnown As Long, Delinquent As Long
  Dim RetCode As Integer, Pending As Long, AcctLu As Long
  Dim AbortFlag As Boolean, UBSetupLen As Integer
  Dim CBeach As Boolean, SCnt As Integer, MtrCnt As Integer
  Dim ThisRate As String, WhatRate As String, NumofRecs As Long
  Dim MtrNum As String, Mtr As String, UseStatus As Boolean
  Dim Fmt4 As String, RStatus As String, Stat As String
  Dim UsingBook As Boolean, UsingName As Boolean, UsingAcct As Boolean
  RStatus$ = Mid$(fpcboCustStatus.Text, 1, 2)
  Select Case RStatus$
  Case "Ac"
    UseStatus = True
    Stat$ = " ACTIVE"
  Case "In"
      UseStatus = True
    Stat$ = " INACTIVE"
  Case "Ba"
    UseStatus = True
    Stat$ = " BALANCE DUE"
  Case "Pe"
    Stat$ = " PENDING"
    UseStatus = True
  Case "De"
    Stat$ = " DELINQUENT"
    UseStatus = True
  Case "Fi"
    Stat$ = " FINAL"
    UseStatus = True
  Case Else
    Stat$ = " ALL"
    UseStatus = False
  End Select
  RStatus$ = Mid$(fpcboCustStatus.Text, 1, 1)
  
  MaxLines = 59
  PageNo = 0
  Dash80$ = String$(80, "-")
  Fmt4$ = "####"
  Title$ = "Customer Listing by Rate Code."
  
  FrmShowPctComp.Label1 = Title$
  FrmShowPctComp.Show

  ReDim UBSetUpRec(1) As UBSetupRecType
  ReportFile$ = UBPath$ + "UBCSBYRT.RPT"
  
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen
'  If InStr(UCase$(UBSetUpRec(1).UTILNAME), "CAROLINA BEACH") > 0 Then
'    CBeach = True
'  End If
'Instead of Carolina Beach Only - Have option on screen to select
'for Print Meter NUm's if yes then use cbeach value of true
  If PrnMtrNum = True Then
    CBeach = True
  End If
  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))
  IdxRecLen = 4 'we are using a long integer
 
  Select Case fpCombo1.ListIndex
   Case 0
    UsingBook = True
    IdxName$ = UBPath$ + "UBCUSTBK.IDX"
    Title$ = "Customer by Rate in Location Order."
   Case 1
    UsingAcct = True
    IdxName$ = ""
    Title$ = "Customer by Rate in Account Order."
   Case 2
    UsingName = True
    IdxName$ = UBPath$ + "UBCUSTNM.IDX"
    Title$ = "Customer by Rate in Name Order."
  End Select
  If Not UsingAcct Then
    IdxFileSize& = FileSize(IdxName$)
    IdxNumOfRecs = IdxFileSize& \ IdxRecLen
    NumofRecs = IdxNumOfRecs
    ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
    Handle = FreeFile
    Open IdxName$ For Random Shared As Handle Len = IdxRecLen
    For cnt& = 1 To IdxNumOfRecs
      Get #Handle, cnt&, IdxBuff(cnt&)
    Next
    Close Handle
  Else
    NumofRecs = FileSize(UBPath$ + "UBCUST.DAT") \ UBCustRecLen
  End If
  UBCust = FreeFile
  Open UBCustFile For Random Shared As UBCust Len = UBCustRecLen
'*****************

  ReDim UBRateTblRec(1) As UBRateTblRecType
  UBRateTblRecLen = Len(UBRateTblRec(1))
  NumOfRateRecs = GetNumRateRecs
  
  If NumOfRateRecs <= 0 Then
    GoTo ExitCustByRate
  End If
  
  Handle = FreeFile
  Open UBPath$ + "UBRATE.DAT" For Random Shared As Handle Len = UBRateTblRecLen
  Get Handle, RateRec, UBRateTblRec(1)
  Close Handle
  WhatRate$ = QPTrim$(UBRateTblRec(1).Ratecode)
  UBRpt = FreeFile
  Open ReportFile$ For Output As UBRpt

  GoSub CustByRateHeader

  For cnt = 1 To NumofRecs
    If Not UsingAcct Then
      AcctLu = IdxBuff(cnt).RecNum
    Else
      AcctLu = cnt
    End If
    Get UBCust, AcctLu, UBCustRec(1)
    If UBCustRec(1).DelFlag = 0 Then
      If UseStatus Then
        If UBCustRec(1).Status <> RStatus$ Then
          GoTo SkipCustRate
        End If
      End If
      
      For SCnt = 1 To 15
        ThisRate$ = QPTrim$(UBCustRec(1).serv(SCnt).Ratecode)
        If WhatRate$ = ThisRate$ Then
          'Mtr$ = "   MTR NO: "
          Mtr$ = ""
          Print #UBRpt, UBCustRec(1).Book; "-"; UBCustRec(1).SEQNUMB; " "; Using$("#####", AcctLu); "  "; Left$(UBCustRec(1).CustName, 25); "  "; Left$(UBCustRec(1).ServAddr, 30); " "; UBCustRec(1).Status
          CustCnt = CustCnt + 1
          If CBeach Then
            'Mtr$=
            For MtrCnt = 1 To 7
              MtrNum$ = QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrNum)
              If Len(MtrNum$) > 0 Then
                Mtr$ = Mtr$ + MtrNum$ + " "
              End If
            Next
            Print #UBRpt, Mtr$
            LineCnt = LineCnt + 2
          Else
            LineCnt = LineCnt + 1
          End If
        Select Case UBCustRec(1).Status
          Case "A"
            Active = Active + 1
          Case "F"
            Final = Final + 1
          Case "I"
            InActive = InActive + 1
          Case "B"
            Balance = Balance + 1
          Case "P"
            Pending = Pending + 1
          Case "D"
            Delinquent = Delinquent + 1
          Case Else
            UnKnown = UnKnown + 1
          End Select
          Exit For
        End If
      Next

      If LineCnt > MaxLines Then
        Print #UBRpt, Chr$(12)
        GoSub CustByRateHeader
      End If
    End If
SkipCustRate:
    FrmShowPctComp.ShowPctComp cnt, NumofRecs
    If FrmShowPctComp.Out Then
      Close
      Unload FrmShowPctComp
      GoTo ExitCustByRate
    End If
  Next

  GoSub CustByRateTotals

  Erase IdxBuff, UBCustRec   'free up memory
  
  If PrintRate Then
    GoSub PrintRateCode
  End If

  Close

  If Not AbortFlag Then
    ViewPrint ReportFile$, Title$
  End If

ExitCustByRate:
  Exit Sub

CustByRateHeader:
  PageNo = PageNo + 1
  Print #UBRpt, Title$; "Date: "; Date$; Tab(70); "Page: "; PageNo
  Print #UBRpt, "RATE CODE: "; WhatRate$
  Print #UBRpt, "Location   Acct.  Customer Name             Service Address             Status"
  Print #UBRpt, "Meter No:"
  Print #UBRpt, Dash80$
  LineCnt = 5
Return

CustByRateTotals:
  'PageNo = PageNo + 1
  Print #UBRpt,
  Print #UBRpt, Dash80$
  Print #UBRpt, "Customer Summary"
  Print #UBRpt,
  Print #UBRpt, "  Active: "; Using(Fmt4$, Active)
  Print #UBRpt, "   Final: "; Using(Fmt4$, Final)
  Print #UBRpt, "Inactive: "; Using(Fmt4$, InActive)
  Print #UBRpt, " Balance: "; Using(Fmt4$, Balance)
  Print #UBRpt, " Pending: "; Using(Fmt4$, Pending)
  Print #UBRpt, "Delinqnt: "; Using(Fmt4$, Delinquent)
  Print #UBRpt, " Unknown: "; Using(Fmt4$, UnKnown)
  Print #UBRpt,
  Print #UBRpt, "   TOTAL: "; Using(Fmt4$, CustCnt)
  Print #UBRpt, Chr$(12)
Return

PrintRateCode:

  ReDim UBRateTblRec(1) As UBRateTblRecType
  UBRateTblRecLen = Len(UBRateTblRec(1))

  NumOfRateRecs = FileSize("UBRATE.DAT") \ UBRateTblRecLen

  If NumOfRateRecs = 0 Then
    GoTo PrintRateExit
  End If

  Handle = FreeFile
  Open UBPath$ + "UBRATE.DAT" For Random Shared As Handle Len = UBRateTblRecLen
  Get Handle, RateRec, UBRateTblRec(1)
  Close Handle

  ReDim StepText(1 To 10) As String * 40

  GoSub PrintRateHeader
    Print #UBRpt, "       Rate Code:  "; UBRateTblRec(1).Ratecode
    Print #UBRpt, "     Description:  "; UBRateTblRec(1).RATEDESC
    Print #UBRpt, "  Minimum Charge:"; Using("#######.##", UBRateTblRec(1).MINAMT)
    Print #UBRpt, "   Minimum Units:"; Using("##########", UBRateTblRec(1).MINUNITS)
    Print #UBRpt, "      Max Amount:"; Using("######.##", UBRateTblRec(1).MaxAmt)
    Print #UBRpt, "      [ Step ]        [ Beg Unit ]     [ Amount/Unit ]"
    For cnt = 1 To 10
      LSet StepText$(cnt) = ""
      If UBRateTblRec(1).TblBreaks(cnt).UNITS >= 0 Then
        Mid$(StepText$(cnt), 8) = UBRateTblRec(1).TblBreaks(cnt).UNITS 'Using("########", UBRateTblRec(1).TblBreaks(cnt).UNITS)
      End If
      If UBRateTblRec(1).TblBreaks(cnt).UNITAMT >= 0 Then
        Mid$(StepText$(cnt), 25) = UBRateTblRec(1).TblBreaks(cnt).UNITAMT 'Using("####.######", UBRateTblRec(1).TblBreaks(cnt).UNITAMT)
      End If
    Next
    Print #UBRpt, "     First Break:"; StepText$(1)
    Print #UBRpt, "    Second Break:"; StepText$(2)
    Print #UBRpt, "     Third Break:"; StepText$(3)
    Print #UBRpt, "    Fourth Break:"; StepText$(4)
    Print #UBRpt, "     Fifth Break:"; StepText$(5)
    Print #UBRpt, "     Sixth Break:"; StepText$(6)
    Print #UBRpt, "   Seventh Break:"; StepText$(7)
    Print #UBRpt, "    Eighth Break:"; StepText$(8)
    Print #UBRpt, "     Ninth Break:"; StepText$(9)
    Print #UBRpt, "        All Over:"; StepText$(10)
    Print #UBRpt,
    Print #UBRpt, Dash80$
    Print #UBRpt, Chr$(12)

  Erase UBRateTblRec, StepText

  GoTo PrintRateExit

PrintRateHeader:
  PageNo = PageNo + 1
  Print #UBRpt, "Utility Rate Table Listing."
  Print #UBRpt, "RATE CODE: "; WhatRate$; Tab(70); "Page:"; PageNo
  Print #UBRpt, "Report Date: "; Date$
  Print #UBRpt, Dash80$
 ' NumPrinted = 0
Return

PrintRateExit:

Return
End Sub

Private Sub UBCustbyRateRpt2()
  Dim IdxName As String, ToPrint As String
  Dim Title As String, ReportFile As String
  Dim UBCustRecLen As Integer, IdxRecLen As Integer
  Dim IdxFileSize As Long, cnt As Long, DeletedCnt As Long
  Dim IdxNumOfRecs As Long, LineCnt As Integer
  Dim Handle As Integer, UBCust As Integer, NumofRecs As Long
  Dim UBRateTblRecLen As Integer, NumOfRateRecs As Integer
  Dim UBRpt As Integer, CustCnt As Long, Delinquent As Long
  Dim Active As Long, Final As Long, InActive As Long
  Dim Balance As Long, UnKnown As Long, Pending As Long
  Dim RetCode As Integer, SubRpt As Long, SubRptFile As String
  Dim AbortFlag As Boolean, UBSetupLen As Integer
  Dim CBeach As Boolean, SCnt As Integer, MtrCnt As Integer
  Dim ThisRate As String, WhatRate As String, Stat As String
  Dim MtrNum As String, Mtr As String, AcctLu As Long
  Dim Fmt4 As String, RStatus As String, UseStatus As Boolean
  Dim UsingBook As Boolean, UsingName As Boolean, UsingAcct As Boolean

  RStatus$ = Mid$(fpcboCustStatus.Text, 1, 2)
  Select Case RStatus$
  Case "Ac"
    UseStatus = True
    Stat$ = " ACTIVE"
  Case "In"
      UseStatus = True
    Stat$ = " INACTIVE"
  Case "Ba"
    UseStatus = True
    Stat$ = " BALANCE DUE"
  Case "Pe"
    Stat$ = " PENDING"
    UseStatus = True
  Case "De"
    Stat$ = " DELINQUENT"
    UseStatus = True
  Case "Fi"
    Stat$ = " FINAL"
    UseStatus = True
  Case Else
    Stat$ = " ALL"
    UseStatus = False
  End Select
  RStatus$ = Mid$(fpcboCustStatus.Text, 1, 1)

  Fmt4$ = "####"
  Title$ = "Customer Listing by Rate Code."
  
  FrmShowPctComp.Label1 = Title$
  FrmShowPctComp.Show

  ReDim UBSetUpRec(1) As UBSetupRecType
  ReportFile$ = UBPath$ + "UBCSBYRT.RPT"
  
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen
'  If InStr(UCase$(UBSetUpRec(1).UTILNAME), "CAROLINA BEACH") > 0 Then
'    CBeach = True
'  End If
'Instead of Carolina Beach Only - Have option on screen to select
'for Print Meter NUm's if yes then use cbeach value of true
  If PrnMtrNum = True Then
    CBeach = True
  End If
  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))
  IdxRecLen = 4 'we are using a long integer

  Select Case fpCombo1.ListIndex
   Case 0
    UsingBook = True
    IdxName$ = UBPath$ + "UBCUSTBK.IDX"
    Title$ = "Customer by Rate in Location Order."
   Case 1
    UsingAcct = True
    IdxName$ = ""
    Title$ = "Customer by Rate in Account Order."
   Case 2
    UsingName = True
    IdxName$ = UBPath$ + "UBCUSTNM.IDX"
    Title$ = "Customer by Rate in Name Order."
  End Select
  If Not UsingAcct Then
    IdxFileSize& = FileSize(IdxName$)
    IdxNumOfRecs = IdxFileSize& \ IdxRecLen
    NumofRecs = IdxNumOfRecs
    ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
    Handle = FreeFile
    Open IdxName$ For Random Shared As Handle Len = IdxRecLen
    For cnt& = 1 To IdxNumOfRecs
      Get #Handle, cnt&, IdxBuff(cnt&)
    Next
    Close Handle
  Else
    NumofRecs = FileSize(UBPath$ + "UBCUST.DAT") \ UBCustRecLen
  End If
  UBCust = FreeFile
  Open UBCustFile For Random Shared As UBCust Len = UBCustRecLen

'*****************

  ReDim UBRateTblRec(1) As UBRateTblRecType
  UBRateTblRecLen = Len(UBRateTblRec(1))
  NumOfRateRecs = GetNumRateRecs
  
  If NumOfRateRecs <= 0 Then
    GoTo ExitCustByRate
  End If
  
  Handle = FreeFile
  Open UBPath$ + "UBRATE.DAT" For Random Shared As Handle Len = UBRateTblRecLen
  Get Handle, RateRec, UBRateTblRec(1)
  Close Handle
  WhatRate$ = QPTrim$(UBRateTblRec(1).Ratecode)
  
  
  UBRpt = FreeFile
  Open ReportFile$ For Output As UBRpt

  For cnt = 1 To NumofRecs
    If Not UsingAcct Then
      AcctLu = IdxBuff(cnt).RecNum
    Else
      AcctLu = cnt
    End If
    Get UBCust, AcctLu, UBCustRec(1)
    If UBCustRec(1).DelFlag = 0 Then
      If UseStatus Then
        If UBCustRec(1).Status <> RStatus$ Then
          GoTo SkipCustRate
        End If
      End If
      
      For SCnt = 1 To 15
        ThisRate$ = QPTrim$(UBCustRec(1).serv(SCnt).Ratecode)
        If WhatRate$ = ThisRate$ Then
          'Mtr$ = "   MTR NO: "
          Mtr$ = ""
          ToPrint$ = UBCustRec(1).Book + "-" + UBCustRec(1).SEQNUMB + "~" + Str(AcctLu) + "~" + Left$(UBCustRec(1).CustName, 25) + "~" + Left$(UBCustRec(1).ServAddr, 30) + "~" + UBCustRec(1).Status + "~ "
          CustCnt = CustCnt + 1
          Print #UBRpt, ToPrint$
          If CBeach Then
            'now select Yes to print meter# with options on screen
            For MtrCnt = 1 To 7
              MtrNum$ = QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrNum)
              If Len(MtrNum$) > 0 Then
                Mtr$ = "Meter Num: "
                Mtr$ = " ~ ~" + Mtr$ + MtrNum$ + "~ ~ ~"
                Print #UBRpt, Mtr$
              End If
              
            Next
'
'            Linecnt = Linecnt + 2
          'Else
'            Linecnt = Linecnt + 1
          End If
        'Print #UBRpt, ToPrint$
        ToPrint$ = ""
        Select Case UBCustRec(1).Status
          Case "A"
            Active = Active + 1
          Case "F"
            Final = Final + 1
          Case "I"
            InActive = InActive + 1
          Case "B"
            Balance = Balance + 1
          Case "P"
            Pending = Pending + 1
          Case "D"
            Delinquent = Delinquent + 1
          Case Else
            UnKnown = UnKnown + 1
          End Select
          Exit For
        End If
      Next

    End If
SkipCustRate:
    FrmShowPctComp.ShowPctComp cnt, NumofRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ActivateControls Me
      GoTo ExitCustByRate
    End If
  Next
  GoSub CustByRateHeader
  GoSub CustByRateTotals

  Erase IdxBuff, UBCustRec   'free up memory
  
  If PrintRate Then
    GoSub PrintRateCode
  Else
    SubRptFile$ = ""
  End If

  Close
  Load frmLoadingRpt
  frmLoadingRpt.setwherefrom frmRptCustbyRate
  ARptQCustList.txtDate = Now
  ARptQCustList.txtTown = TOWNNAME$
  ARptQCustList.Title = Title$
  ARptQCustList.Label6.Visible = False
  ARptQCustList.GetName ReportFile$, SubRptFile$
  ARptQCustList.startrpt

'  If Not AbortFlag Then
'    ViewPrint ReportFile$, Title$
'  End If

ExitCustByRate:
  Exit Sub
  
CustByRateHeader:
 ARptQCustList.lableRate.Visible = True
 ARptQCustList.lableRate = "RATE CODE: " + WhatRate$
 ARptQCustList.lblRptOpt.Visible = False
Return

CustByRateTotals:
  ARptQCustList.totActive = Using$("#####", Active)
  ARptQCustList.totFinal = Using$("#####", Final)
  ARptQCustList.totInactive = Using$("#####", InActive)
  ARptQCustList.totBalance = Using$("#####", Balance)
  ARptQCustList.totPending = Using$("#####", Pending)
  ARptQCustList.totDelinquent = Using$("#####", Delinquent)
  ARptQCustList.totUnknown = Using$("#####", UnKnown)
  ARptQCustList.totDeleted = Using$("#####", DeletedCnt)
  ARptQCustList.totTotal = Using$("#####", CustCnt)
Return

PrintRateCode:
    SubRptFile$ = UBPath$ + "SubRate.RPT"
    SubRpt = FreeFile
    Open SubRptFile$ For Output As SubRpt

  ReDim UBRateTblRec(1) As UBRateTblRecType
  UBRateTblRecLen = Len(UBRateTblRec(1))

  NumOfRateRecs = FileSize(UBPath$ + "UBRATE.DAT") \ UBRateTblRecLen

  If NumOfRateRecs = 0 Then
    GoTo PrintRateExit
  End If

  Handle = FreeFile
  Open UBPath$ + "UBRATE.DAT" For Random Shared As Handle Len = UBRateTblRecLen
  Get Handle, RateRec, UBRateTblRec(1)
  Close Handle

  ReDim StepText(1 To 10) As String * 40

  
    Print #SubRpt, "Rate Code: ~" + UBRateTblRec(1).Ratecode + "~"
    Print #SubRpt, "Description: ~" + UBRateTblRec(1).RATEDESC + "~"
    Print #SubRpt, "Minimum Charge: ~" + Using("#######.##", UBRateTblRec(1).MINAMT) + "~"
    Print #SubRpt, "Minimum Units: ~" + Using("##########", UBRateTblRec(1).MINUNITS) + "~"
    Print #SubRpt, "Max Amount: ~" + Using("######.##", UBRateTblRec(1).MaxAmt) + "~"
    Print #SubRpt, "      [ Step ]   ~     [ Beg Unit ]~     [ Amount/Unit ]"
    For cnt = 1 To 10
      LSet StepText$(cnt) = ""
      If UBRateTblRec(1).TblBreaks(cnt).UNITS >= 0 Then
        Mid$(StepText$(cnt), 8) = Str(UBRateTblRec(1).TblBreaks(cnt).UNITS) + "~" 'Using("########", UBRateTblRec(1).TblBreaks(cnt).UNITS) + "~"
      Else
        Mid$(StepText$(cnt), 8) = " ~"
      End If
      If UBRateTblRec(1).TblBreaks(cnt).UNITAMT >= 0 Then
        Mid$(StepText$(cnt), 25) = UBRateTblRec(1).TblBreaks(cnt).UNITAMT 'Using("####.######", UBRateTblRec(1).TblBreaks(cnt).UNITAMT)
      End If
    Next
    Print #SubRpt, "First Break: ~" + StepText$(1)
    Print #SubRpt, "Second Break: ~" + StepText$(2)
    Print #SubRpt, "Third Break: ~" + StepText$(3)
    Print #SubRpt, "Fourth Break: ~" + StepText$(4)
    Print #SubRpt, "Fifth Break: ~" + StepText$(5)
    Print #SubRpt, "Sixth Break: ~" + StepText$(6)
    Print #SubRpt, "Seventh Break: ~" + StepText$(7)
    Print #SubRpt, "Eighth Break: ~" + StepText$(8)
    Print #SubRpt, "Ninth Break: ~" + StepText$(9)
    Print #SubRpt, "All Over: ~" + StepText$(10)

  Erase UBRateTblRec, StepText

  GoTo PrintRateExit


PrintRateExit:

Return
End Sub

Private Sub fpComboRates_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpComboRates.ListDown = True
  End If
  If fpComboRates.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpPrintCode.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        cmdExit.SetFocus
        KeyCode = 0
      End If
    End If
  End If

End Sub
Private Sub fpcboCustStatus_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboCustStatus.ListDown = True
  End If
  If fpcboCustStatus.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpCombo1.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpMtrNum.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub
Private Sub fpCombo1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpCombo1.ListDown = True
  End If
  If fpCombo1.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpcboRptType.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpcboCustStatus.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub

Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPnScn_Click()
  PrintForm
End Sub
