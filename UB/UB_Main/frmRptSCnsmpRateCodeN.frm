VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmRptSCnsmpRateCodeN 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consumption By Rate Code Report"
   ClientHeight    =   8640
   ClientLeft      =   30
   ClientTop       =   540
   ClientWidth     =   12195
   Icon            =   "frmRptSCnsmpRateCodeN.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   12195
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboRptType 
      Height          =   375
      Left            =   3885
      TabIndex        =   2
      Top             =   1680
      Width           =   1905
      _Version        =   196608
      _ExtentX        =   3360
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
      ColDesigner     =   "frmRptSCnsmpRateCodeN.frx":08CA
   End
   Begin LpLib.fpList fplstRates 
      Height          =   855
      Left            =   6435
      TabIndex        =   3
      Top             =   1245
      Width           =   4935
      _Version        =   196608
      _ExtentX        =   8705
      _ExtentY        =   1508
      TextAlias       =   ""
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
      Columns         =   2
      Sorted          =   0
      LineWidth       =   1
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   -1
      ColumnWidthScale=   2
      RowHeight       =   -1
      MultiSelect     =   1
      WrapList        =   0   'False
      WrapWidth       =   0
      SelMax          =   -1
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
      DataField       =   ""
      OLEDragMode     =   0
      OLEDropMode     =   0
      Redraw          =   -1  'True
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      ColDesigner     =   "frmRptSCnsmpRateCodeN.frx":0C30
   End
   Begin LpLib.fpList fplstGCodes 
      Height          =   1200
      Left            =   6990
      TabIndex        =   11
      Top             =   5790
      Width           =   4350
      _Version        =   196608
      _ExtentX        =   7673
      _ExtentY        =   2117
      TextAlias       =   ""
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
      Columns         =   3
      Sorted          =   0
      LineWidth       =   1
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   -1
      ColumnWidthScale=   2
      RowHeight       =   -1
      MultiSelect     =   1
      WrapList        =   0   'False
      WrapWidth       =   0
      SelMax          =   -1
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
      DataField       =   ""
      OLEDragMode     =   0
      OLEDropMode     =   0
      Redraw          =   -1  'True
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      ColDesigner     =   "frmRptSCnsmpRateCodeN.frx":0F4C
   End
   Begin VB.OptionButton optAllCust 
      BackColor       =   &H00C0C0C0&
      Caption         =   "All Customers"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Left            =   1608
      TabIndex        =   5
      Top             =   2568
      Width           =   2100
   End
   Begin EditLib.fpText fptxtBook 
      Height          =   948
      Left            =   6984
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   3432
      Width           =   4332
      _Version        =   196608
      _ExtentX        =   7641
      _ExtentY        =   1672
      Enabled         =   0   'False
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
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
      AutoAdvance     =   0   'False
      AutoBeep        =   0   'False
      AutoCase        =   0
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
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   1
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   255
      MultiLine       =   -1  'True
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
   Begin VB.OptionButton OptGroup 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Group"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   444
      Left            =   1098
      TabIndex        =   10
      Top             =   5784
      Width           =   1260
   End
   Begin VB.OptionButton OptCycle 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cycle"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   444
      Left            =   1098
      TabIndex        =   8
      Top             =   4608
      Width           =   1260
   End
   Begin VB.OptionButton OptBook 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Book"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   444
      Left            =   1098
      TabIndex        =   6
      Top             =   3432
      Width           =   1260
   End
   Begin EditLib.fpLongInteger fptxtCycleSel 
      Height          =   348
      Left            =   4098
      TabIndex        =   9
      Top             =   4680
      Width           =   636
      _Version        =   196608
      _ExtentX        =   1122
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
      ButtonMin       =   1
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
      Text            =   "0"
      MaxValue        =   "2147483647"
      MinValue        =   "-2147483648"
      NegFormat       =   1
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
      BorderGrayAreaColor=   -2147483637
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
   Begin VB.CheckBox PageBrk 
      Alignment       =   1  'Right Justify
      BackColor       =   &H008F8265&
      Caption         =   "Page Break on Rates:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1410
      TabIndex        =   4
      Top             =   2064
      Width           =   2652
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
      Left            =   8568
      TabIndex        =   12
      Top             =   7320
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
      Left            =   10176
      TabIndex        =   13
      Top             =   7320
      Width           =   1332
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   14
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
            TextSave        =   "5:14 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7117
            TextSave        =   "11/17/2011"
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
   Begin EditLib.fpDateTime txtDate2 
      Height          =   348
      Left            =   3888
      TabIndex        =   1
      Top             =   1272
      Width           =   1692
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
      Height          =   348
      Left            =   3888
      TabIndex        =   0
      Top             =   864
      Width           =   1692
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
   Begin EditLib.fpLongInteger fptxtBookSel 
      Height          =   348
      Left            =   4098
      TabIndex        =   7
      Top             =   3456
      Width           =   636
      _Version        =   196608
      _ExtentX        =   1122
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
      ButtonMin       =   1
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
      Text            =   "0"
      MaxValue        =   "2147483647"
      MinValue        =   "-2147483648"
      NegFormat       =   1
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
      BorderGrayAreaColor=   -2147483637
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
   Begin EditLib.fpText fptxtcycle 
      Height          =   948
      Left            =   6984
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   4608
      Width           =   4332
      _Version        =   196608
      _ExtentX        =   7641
      _ExtentY        =   1672
      Enabled         =   0   'False
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
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
      AutoAdvance     =   0   'False
      AutoBeep        =   0   'False
      AutoCase        =   0
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
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   1
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   255
      MultiLine       =   -1  'True
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
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Search:"
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
      Height          =   300
      Index           =   4
      Left            =   48
      TabIndex        =   29
      Top             =   2592
      Width           =   1452
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   1032
      X2              =   11508
      Y1              =   3312
      Y2              =   3324
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   1032
      X2              =   11520
      Y1              =   5688
      Y2              =   5688
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   1008
      X2              =   11496
      Y1              =   4512
      Y2              =   4512
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   4620
      Left            =   528
      Top             =   2472
      Width           =   11004
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Selected Cycles:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   348
      Index           =   2
      Left            =   4512
      TabIndex        =   28
      Top             =   4656
      Width           =   2340
   End
   Begin VB.Label lblGrp1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Select Group Codes From List:"
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
      Left            =   3402
      TabIndex        =   26
      Top             =   5880
      Width           =   3420
   End
   Begin VB.Label lblGrp2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "* Press SpaceBar or Mouse to Toggle"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   348
      Left            =   3618
      TabIndex        =   25
      Top             =   6216
      Width           =   3132
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Or Select Search Option Below: "
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
      Height          =   396
      Left            =   456
      TabIndex        =   24
      Top             =   3000
      Width           =   3852
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Cycle:"
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
      Height          =   300
      Index           =   3
      Left            =   2562
      TabIndex        =   23
      Top             =   4728
      Width           =   1452
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Book:"
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
      Height          =   300
      Index           =   1
      Left            =   2562
      TabIndex        =   22
      Top             =   3504
      Width           =   1452
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Selected Books:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   348
      Index           =   0
      Left            =   4896
      TabIndex        =   21
      Top             =   3480
      Width           =   1956
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Select Rate Codes From List:"
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
      Left            =   5712
      TabIndex        =   19
      Top             =   936
      Width           =   3324
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
      Height          =   324
      Index           =   0
      Left            =   2160
      TabIndex        =   18
      Top             =   1320
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
      Left            =   2064
      TabIndex        =   17
      Top             =   912
      Width           =   1668
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   1836
      Left            =   528
      Top             =   648
      Width           =   11004
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
      Left            =   1464
      TabIndex        =   16
      Top             =   1704
      Width           =   2340
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   588
      Left            =   3216
      Top             =   72
      Width           =   5772
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Consumption By Rate Report"
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
      Left            =   3216
      TabIndex        =   15
      Top             =   216
      Width           =   5772
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   684
      Left            =   3216
      Top             =   -24
      Width           =   5772
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   1008
      X2              =   1020
      Y1              =   3312
      Y2              =   7044
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
Attribute VB_Name = "frmRptSCnsmpRateCodeN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim RateRec As Integer, RCnt As Integer
Dim Doall As Boolean, Allcust As Boolean
Dim CycleCnt As Integer, BookCnt As Integer
Dim Cycle(1 To 30) As Integer
Dim Book(1 To 30) As Integer
Dim NumOfcdsrpt As Integer, Codefile As Integer, GCnt As Integer
Dim CodeName As String, Grp As String, RateD As String, RCName As String
Dim NumOfratesrpt As Integer, RCfile As Integer
Dim CodestoRpt As GroupCodeRptType
Dim RCstoRpt As RateCodeRptType

Private Sub cmdExit_Click()
  frmUBStatReportsMenu.Show
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
        UBLog "Closed via RptSCnsmpExtra by " + PWUser$
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
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
  Dim RCode As String
  Dim Handle As Integer, cnt As Integer
  Dim UBRateTblRecLen As Integer, NumOfRateRecs As Integer
  Me.HelpContextID = hlpConsumptionByRateCyc
  ReDim UBRateTblRec(1) As UBRateTblRecType
  RCode$ = Space$(10)
  UBRateTblRecLen = Len(UBRateTblRec(1))
  NumOfRateRecs = GetNumRateRecs
  Handle = FreeFile
  fplstRates.AddItem "ALL" & Chr$(9) & "-Print All Rates"
  Open UBPath$ + "UBRATE.DAT" For Random Shared As Handle Len = UBRateTblRecLen
  For cnt = 1 To NumOfRateRecs
    Get Handle, cnt, UBRateTblRec(1)
    LSet RCode$ = QPTrim$(UBRateTblRec(1).Ratecode)
    fplstRates.AddItem RCode$ & Chr$(9) & QPTrim$(UBRateTblRec(1).RATEDESC)
  Next
  Close
  fplstRates.ListIndex = 1
  fplstRates.Selected(1) = True
  GCodesList fplstGCodes
  txtDate1.Text = Format(Now, "mm/dd/yyyy")
  txtDate2.Text = Format(Now, "mm/dd/yyyy")
  fpcboRptType.InsertRow = "Graphics"
  fpcboRptType.InsertRow = "Text"
  fpcboRptType.ListIndex = 0
  optAllCust.Value = True
  Erase Cycle()
  Erase Book()
  BookCnt = 0
  CycleCnt = 0
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

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub

'Private Sub OptBook_Click()
'  If OptBook.Value = True Then
'
'End Sub

Private Sub txtDate1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    txtDate2.SetFocus
  End If
End Sub

Private Sub txtDate2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fplstRates.SetFocus
  End If
End Sub


Private Function ValidDate()
  Dim TempDate1 As Integer, TempDate2 As Integer
  If CheckValDate(txtDate1) = False Then
    MsgBox "Date Is Not Valid. Please Correct.", vbOKOnly, "Invalid Date"
    ValidDate = False
  ElseIf CheckValDate(txtDate2) = False Then
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
  Dim RptType As Integer
'RptOpt 1 is Consumption by rate code
'2 is Irrigation Consumption Report
  If fplstRates.SelCount > 0 Then
    RateRec = fplstRates.ListIndex
    If RateRec = 0 Then
      Doall = True
    Else
      Doall = False
      GetRCodestoReport fplstRates
    End If
  Else
    MsgBox "Error with Rate selection, please select a Rate option from the list.", vbOKOnly, "Invalid Selection"
    Exit Sub
  End If
  If optAllCust.Value = True Then
    Allcust = True
  Else
    Allcust = False
  End If
  If OptBook.Value = True Then
    If Not CheckBooks% Then
      MsgBox "Error with Book selection, please enter books again.", vbOKOnly, "Invalid Selection"
      Exit Sub
    End If
  ElseIf OptCycle.Value = True Then
    If Not CheckCycles% Then
      MsgBox "Error with Cycle selection, please enter cycles again.", vbOKOnly, "Invalid Selection"
      Exit Sub
    End If
  ElseIf OptGroup.Value = True Then
    If Not fplstGCodes.SelCount > 0 Then
      MsgBox "Error with Group selection, please select groups again.", vbOKOnly, "Invalid Selection"
      Exit Sub
    Else
      GetGCodestoReport fplstGCodes
    End If
  End If
  If ValidDate Then
'    DeActivateControls Me, True
    RptType = fpcboRptType.ListIndex
        ConsumpUnitStep RptType
 '     ActivateControls Me, True
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
        fplstRates.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub
Private Sub fptxtBookSel_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim cnt As Integer
  If KeyCode = vbKeyReturn Then
    If Len(fptxtBookSel.Text) <> 0 Then
      getbooklist
    Else
     'cmdOk.SetFocus
    End If
  End If
End Sub
Private Sub getbooklist()
  Dim TBok As String
  Dim ThisBook As Integer
  Dim cnt As Integer
 
  TBok$ = QPTrim$(fptxtBookSel.Text)
  If TBok$ = "0" Then
    fptxtBook.Text = ""
    BookCnt = 0
    Erase Book
    'cmdOk.SetFocus
  Else
    If Len(TBok$) > 0 Then
      ThisBook = Val(fptxtBookSel.Text)
      For cnt = 1 To 30
        If ThisBook = Book(cnt) Then
          GoTo DupeExit
        End If
      Next
      BookCnt = BookCnt + 1
      If BookCnt > 30 Then
        BookCnt = 30
        GoTo DupeExit
      End If
      Book(BookCnt) = ThisBook
      fptxtBook.Text = ""
      For cnt = 1 To BookCnt
        If cnt = BookCnt Then
          fptxtBook.Text = fptxtBook.Text & Book(cnt)
        Else
          fptxtBook.Text = fptxtBook.Text & Book(cnt) & ","
        End If
      Next
    End If
  End If
DupeExit:
  fptxtBookSel.Text = ""
End Sub
Private Function CheckBooks%()
  
  Dim BooksOK As Boolean
  BooksOK = False
  For RCnt = 1 To 30
    If Book(RCnt) > 0 Then
      BooksOK = True
      Exit For
    End If
  Next
  
  If Not BooksOK Then 'duh nothing to export
'    frmMsgDialog.RetLabel = "-2"
'    frmMsgDialog.Caption = "ERROR:"
'    For RCnt = 0 To 4
'      frmMsgDialog.Label(RCnt).Caption = ""
'      frmMsgDialog.Label(RCnt).FontSize = frmMsgDialog.Label(RCnt).FontSize + 2
'    Next
'    frmMsgDialog.Label(1).Caption = "NO CYCLES ENTERED TO EXPORT."
'    frmMsgDialog.Label(2).Caption = "Please call Southern Software for"
'    frmMsgDialog.Label(3).Caption = "additional Information."
'    frmMsgDialog.Show vbModal
'    Unload frmMsgDialog
    GoTo CheckBooksExit
  End If

CheckBooksExit:
  
  CheckBooks% = BooksOK
End Function

Private Sub fptxtCycleSel_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim cnt As Integer
  If KeyCode = vbKeyReturn Then
    If Len(fptxtCycleSel.Text) <> 0 Then
      getcyclelist
    Else
     'cmdOk.SetFocus
    End If
  End If
End Sub

Private Sub getcyclelist()
  Dim TCyc As String
  Dim ThisCycle As Integer
  Dim cnt As Integer
 
  TCyc$ = QPTrim$(fptxtCycleSel.Text)
  If TCyc$ = "0" Then
    fptxtcycle.Text = ""
    CycleCnt = 0
    Erase Cycle
    'cmdOk.SetFocus
  Else
    If Len(TCyc$) > 0 Then
      ThisCycle = Val(fptxtCycleSel.Text)
      For cnt = 1 To 30
        If ThisCycle = Cycle(cnt) Then
          GoTo DupeExit
        End If
      Next
      CycleCnt = CycleCnt + 1
      If CycleCnt > 30 Then
        CycleCnt = 30
        GoTo DupeExit
      End If
      Cycle(CycleCnt) = ThisCycle
      fptxtcycle.Text = ""
      For cnt = 1 To CycleCnt
        If cnt = CycleCnt Then
          fptxtcycle.Text = fptxtcycle.Text & Cycle(cnt)
        Else
          fptxtcycle.Text = fptxtcycle.Text & Cycle(cnt) & ","
        End If
      Next
    End If
  End If
DupeExit:
  fptxtCycleSel.Text = ""
End Sub
Private Function CheckCycles%()
  
  Dim CyclesOK As Boolean
  CyclesOK = False
  For RCnt = 1 To 30
    If Cycle(RCnt) > 0 Then
      CyclesOK = True
      Exit For
    End If
  Next
  
  If Not CyclesOK Then 'duh nothing to export
'    frmMsgDialog.RetLabel = "-2"
'    frmMsgDialog.Caption = "ERROR:"
'    For RCnt = 0 To 4
'      frmMsgDialog.Label(RCnt).Caption = ""
'      frmMsgDialog.Label(RCnt).FontSize = frmMsgDialog.Label(RCnt).FontSize + 2
'    Next
'    frmMsgDialog.Label(1).Caption = "NO CYCLES ENTERED TO EXPORT."
'    frmMsgDialog.Label(2).Caption = "Please call Southern Software for"
'    frmMsgDialog.Label(3).Caption = "additional Information."
'    frmMsgDialog.Show vbModal
'    Unload frmMsgDialog
    GoTo CheckCyclesExit
  End If

CheckCyclesExit:
  
  CheckCycles% = CyclesOK
End Function

Private Sub ConsumpUnitStep(RptType)
  Dim Dash80 As String, IdxName As String, IdxRecLen As Integer
  'Dim MINAMT(1 To 1) As Double
  'Dim RATECODE(1 To 1) As String
  Dim UBCustRecLen As Integer, UBCust As Integer, RCnt As Integer
  Dim IdxFileSize As Long, IdxNumOfRecs As Long, cnt As Long
  Dim UBTransRecLen As Integer, UBSetupreclen As Integer
  Dim Handle As Integer, UBTrans As Integer, NumOfRecs As Long
  Dim UBRateTblRecLen As Integer, NumOfRates As Integer
  Dim UBRpt As Integer, UBSetUp As Integer, ValidCustomer As Integer
  Dim BegDate As Integer, EndDate As Integer, Snt As Integer
  Dim TownLen As Integer, TabStop As Integer, MeterConsp As Long
  Dim UBSetupLen As Integer, RateFile As Integer, MT As Integer
  Dim Greater As Boolean, MaxMeterAmt As Long, RCode As String
  Dim Tnt As Integer, NMinAMT As Double, ToPrint As String
  Dim CustomerRecord  As Long, MCnt As Integer, GTMeterConsp As Double
  Dim Multi As Long, Cubic As Boolean, ChkMtr As Boolean
  Dim MTRType As String, MType As String, TMeterConsp As Double
  Dim NonUpdated As Integer, LL As Integer, BigUTotal As Double
  Dim ReportFile As String, MinGT As Double, GBBigUTotal As Double
  Dim GBMinGT As Double, GBGTMeterConsp As Double, BigTotCust As Long
  Dim GBCustTot As Long, Bcnt As Integer, CCnt As Integer, GCnt As Integer
  Dim RCstoRpt As RateCodeRptType
  Dim RCfile As Integer, RptInfo As String, RptInfo2 As String
  Dim RCName As String, Tempcalccnsp As Double, ToPrintI As String
  Dim PCnt As Integer, NewMtrConsp As Double, NumUser As Long
  Dim NTAmt As Double, UNITS As Long, UntPrc As Double
  Dim MinBillAmt As Double, TAmt As Double, MaxFlag As Boolean
  Dim MinimumConsp As Long, numflag As Boolean, NTMeterConsp As Double
  PageNo = 0
  MaxLines = 56
  Dash80$ = String$(80, "-")
  NumOfRates = GetNumRateRecs%
  numflag = False
  If Not Doall Then
    RCName$ = "Ratecds.LST"
    If Not Exist(RCName$) Then GoTo ExitConsStep
    RCfile = FreeFile
    Open RCName$ For Random As RCfile Len = Len(RCstoRpt)
    NumOfratesrpt = FileSize(RCName$) \ Len(RCstoRpt)
    Close RCfile
    ReDim Rate2Rpt(1 To NumOfratesrpt) As Integer
    ReDim UBRateTbls(1 To NumOfratesrpt) As UBRateTblRecType
    ReDim MINAMT(1 To NumOfratesrpt) As Double
    ReDim minunt(1 To NumOfratesrpt) As Double
    ReDim MaxAmt(1 To NumOfratesrpt) As Double
    ReDim Ratecode(1 To NumOfratesrpt) As String
    ReDim MaxStep(1 To NumOfratesrpt) As Integer
    ReDim TblBreak(1 To NumOfratesrpt, 11) As Long
    ReDim TblBreakfr(1 To NumOfratesrpt, 11) As Long
    ReDim TblUnitVal(1 To NumOfratesrpt, 11) As Double
    ReDim tblbrkchg(1 To NumOfratesrpt, 11) As Double
    ReDim TotalConsp(1 To NumOfratesrpt, 11) As Double
    ReDim TotalCust(1 To NumOfratesrpt) As Long
    ReDim numofuser(1 To NumOfratesrpt) As Long
 Else
    NumOfratesrpt = NumOfRates
    ReDim UBRateTbls(1 To NumOfRates) As UBRateTblRecType
    ReDim MINAMT(1 To NumOfRates) As Double
    ReDim minunt(1 To NumOfRates) As Double
    ReDim MaxAmt(1 To NumOfRates) As Double
    ReDim Ratecode(1 To NumOfRates) As String
    ReDim MaxStep(1 To NumOfRates) As Integer
    ReDim TblBreak(1 To NumOfRates, 11) As Long
    ReDim TblBreakfr(1 To NumOfRates, 11) As Long
    ReDim TblUnitVal(1 To NumOfRates, 11) As Double
    ReDim tblbrkchg(1 To NumOfRates, 11) As Double
    ReDim TotalConsp(1 To NumOfRates, 11) As Double
    ReDim TotalCust(1 To NumOfRates) As Long
    ReDim numofuser(1 To NumOfRates) As Long
  End If
  ReDim UBSetUpRec(1) As UBSetupRecType

  CodeName$ = "grpcds.LST"
 ' RCName$ = "Ratecds.lst"
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen
  FrmShowPctComp.Label1 = "Creating Consumption Report"
  FrmShowPctComp.Show
  BegDate = Date2Num%(txtDate1)
  EndDate = Date2Num%(txtDate2)
  IdxRecLen = 4 'we are using a long integer
  IdxFileSize& = FileSize(BookIndexFile)
  IdxNumOfRecs = IdxFileSize& \ IdxRecLen
  ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
  Handle = FreeFile
  Open BookIndexFile For Random Shared As Handle Len = IdxRecLen
  For cnt = 1 To IdxNumOfRecs
    Get #Handle, cnt, IdxBuff(cnt)
  Next
  Close #Handle
  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))
  ReDim UBTransRec(1) As UBTransRecType
  UBTransRecLen = Len(UBTransRec(1))
  ReDim UBSetUpRec(1) As UBSetupRecType
  UBSetupreclen = Len(UBSetUpRec(1))
  UBCust = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen
  UBTrans = FreeFile
  Open UBPath$ + "UBTRANS.DAT" For Random Shared As UBTrans Len = UBTransRecLen
  NumOfRecs& = LOF(UBTrans) / UBTransRecLen
  ReportFile$ = UBPath$ + "UBBKCNSP.RPT"
  UBRpt = FreeFile
  Open ReportFile$ For Output As UBRpt
  Rem Report Goes Here
  UBSetUp = FreeFile
  Open UBPath$ + "UBSETUP.DAT" For Random Access Read Write Shared As UBSetUp Len = UBSetupreclen
  If LOF(UBSetUp) / UBSetupreclen = 0 Then
    TOWNNAME$ = "Undefined"
  Else
    Get UBSetUp, 1, UBSetUpRec(1)
    TOWNNAME$ = UBSetUpRec(1).UTILNAME
    TownLen = Len(RTrim$(TOWNNAME$))
    TabStop = 40 - (TownLen / 2)
    If TabStop < 1 Then TabStop = 1
  End If
  Close UBSetUp
  If Not Doall Then
    RCfile = FreeFile
    Open RCName$ For Random As RCfile Len = Len(RCstoRpt)
    UBRateTblRecLen = Len(UBRateTbls(1))
    RateFile = FreeFile
    Open UBPath$ + "UBRATE.DAT" For Random Shared As RateFile Len = UBRateTblRecLen
    For PCnt = 1 To NumOfratesrpt
       Get RCfile, PCnt, RCstoRpt
       RateRec = RCstoRpt.RecordNum
       Get RateFile, RateRec, UBRateTbls(PCnt)
       'Rate2Rpt(PCnt) = RateRec
       RptInfo$ = RptInfo$ + QPTrim$(UBRateTbls(PCnt).Ratecode) + " "
       Ratecode(PCnt) = QPTrim$(UBRateTbls(PCnt).Ratecode)
     Next PCnt
  Close RateFile
  Else
    RptInfo$ = "All Rates"
    UBRateTblRecLen = Len(UBRateTbls(1))
    RateFile = FreeFile
    Open UBPath$ + "UBRATE.DAT" For Random Shared As RateFile Len = UBRateTblRecLen
       For RateRec = 1 To NumOfratesrpt
        Greater = False
        Get RateFile, RateRec, UBRateTbls(RateRec)
        Ratecode(RateRec) = QPTrim$(UBRateTbls(RateRec).Ratecode)
      Next RateRec
    Close RateFile
  End If

  If Not Doall Then
    RCfile = FreeFile
    Open RCName$ For Random As RCfile Len = Len(RCstoRpt)
  
    UBRateTblRecLen = Len(UBRateTbls(1))
    RateFile = FreeFile
    Open UBPath$ + "UBRATE.DAT" For Random Shared As RateFile Len = UBRateTblRecLen
    For PCnt = 1 To NumOfratesrpt
       Get RCfile, PCnt, RCstoRpt
       RateRec = RCstoRpt.RecordNum
       Get RateFile, RateRec, UBRateTbls(PCnt)

      Ratecode(PCnt) = QPTrim$(UBRateTbls(PCnt).Ratecode)
      MINAMT(PCnt) = UBRateTbls(PCnt).MINAMT
      minunt(PCnt) = UBRateTbls(PCnt).MINUNITS
      MaxAmt(PCnt) = UBRateTbls(PCnt).MaxAmt
      TblBreak&(PCnt, 1) = minunt(PCnt)
      TblBreakfr&(PCnt, 1) = 0
      TblUnitVal#(PCnt, 1) = 0
      tblbrkchg#(PCnt, 1) = 0
      MaxStep(PCnt) = 1
      For Tnt = 1 To 10
        If UBRateTbls(PCnt).TblBreaks(Tnt).UNITS <= 0 And UBRateTbls(PCnt).TblBreaks(Tnt).UNITAMT <= 0 Then
          Exit For
        End If
        If Tnt = 10 Then
          If UBRateTbls(PCnt).TblBreaks(Tnt).UNITS > 0 And UBRateTbls(PCnt).TblBreaks(Tnt).UNITAMT > 0 Then
            MaxStep(PCnt) = Tnt + 1
            TblBreakfr&(PCnt, Tnt + 1) = UBRateTbls(PCnt).TblBreaks(Tnt).UNITS
            TblBreak&(PCnt, Tnt + 1) = 99999999
            TblUnitVal#(PCnt, Tnt + 1) = UBRateTbls(PCnt).TblBreaks(Tnt).UNITAMT
            tblbrkchg#(PCnt, Tnt + 1) = 0
            Exit For
'          ElseIf UBRateTbls(1).TblBreaks(Tnt).UNITAMT <= 0 Then
'            Exit For
          End If
        Else

        If UBRateTbls(PCnt).TblBreaks(Tnt + 1).UNITS > 0 Then
          If UBRateTbls(PCnt).TblBreaks(Tnt).UNITS >= 0 Then
            If UBRateTbls(PCnt).TblBreaks(Tnt).UNITS = 0 Then
              TblBreak&(PCnt, Tnt + 1) = UBRateTbls(PCnt).TblBreaks(Tnt + 1).UNITS - 1
              TblBreakfr&(PCnt, Tnt + 1) = UBRateTbls(PCnt).TblBreaks(Tnt).UNITS
              TblUnitVal#(PCnt, Tnt + 1) = UBRateTbls(PCnt).TblBreaks(Tnt).UNITAMT
              tblbrkchg#(PCnt, Tnt + 1) = 0
              MaxStep(PCnt) = Tnt + 1
            Else
             ' If Tnt > 1 Then
                TblBreakfr&(PCnt, Tnt + 1) = UBRateTbls(PCnt).TblBreaks(Tnt).UNITS
                TblBreak&(PCnt, Tnt + 1) = UBRateTbls(PCnt).TblBreaks(Tnt + 1).UNITS - 1
             ' Else
             '   TblBreakfr&(1, Tnt + 1) = minunt(1)
             '   TblBreak&(1, Tnt + 1) = UBRateTbls(1).TblBreaks(Tnt).UNITS
             ' End If
              TblUnitVal#(PCnt, Tnt + 1) = UBRateTbls(PCnt).TblBreaks(Tnt).UNITAMT
              tblbrkchg#(PCnt, Tnt + 1) = 0
              MaxStep(PCnt) = Tnt + 1
            End If
          Else
            MaxStep(PCnt) = Tnt + 1
            TblBreakfr&(PCnt, Tnt + 1) = 0
            TblBreak&(PCnt, Tnt + 1) = 99999999
            TblUnitVal#(PCnt, Tnt + 1) = UBRateTbls(PCnt).TblBreaks(Tnt).UNITAMT
            tblbrkchg#(PCnt, Tnt + 1) = 0
          End If
        Else
          MaxStep(PCnt) = Tnt + 1
          TblBreakfr&(PCnt, Tnt + 1) = UBRateTbls(PCnt).TblBreaks(Tnt).UNITS
          TblBreak&(PCnt, Tnt + 1) = 99999999
          TblUnitVal#(PCnt, Tnt + 1) = UBRateTbls(PCnt).TblBreaks(Tnt).UNITAMT
          tblbrkchg#(PCnt, Tnt + 1) = 0
          Exit For
        End If
        End If
      Next Tnt
    Next PCnt
  Close RateFile
  Else
    UBRateTblRecLen = Len(UBRateTbls(1))
    RateFile = FreeFile
    Open UBPath$ + "UBRATE.DAT" For Random Shared As RateFile Len = UBRateTblRecLen
       For RateRec = 1 To NumOfratesrpt
        Greater = False
        Get RateFile, RateRec, UBRateTbls(RateRec)
       ' Ratecode(RateRec) = QPTrim$(UBRateTbls(RateRec).Ratecode)
        MINAMT(RateRec) = UBRateTbls(RateRec).MINAMT
        minunt(RateRec) = UBRateTbls(RateRec).MINUNITS
        MaxAmt(RateRec) = UBRateTbls(RateRec).MaxAmt
        TblBreakfr&(RateRec, 1) = 0
        TblBreak&(RateRec, 1) = minunt(RateRec)
        TblUnitVal#(RateRec, 1) = 0
        tblbrkchg#(RateRec, 1) = 0
        MaxStep(RateRec) = 1
        For Tnt = 1 To 10
          If UBRateTbls(RateRec).TblBreaks(Tnt).UNITS <= 0 And UBRateTbls(RateRec).TblBreaks(Tnt).UNITAMT <= 0 Then
            Exit For
          End If
          If Tnt = 10 Then
            If UBRateTbls(RateRec).TblBreaks(Tnt).UNITS > 0 And UBRateTbls(RateRec).TblBreaks(Tnt).UNITAMT > 0 Then
              MaxStep(RateRec) = Tnt + 1
              TblBreakfr&(RateRec, Tnt + 1) = UBRateTbls(RateRec).TblBreaks(Tnt).UNITS
              TblBreak&(RateRec, Tnt + 1) = 99999999
              TblUnitVal#(RateRec, Tnt + 1) = UBRateTbls(RateRec).TblBreaks(Tnt).UNITAMT
              tblbrkchg#(RateRec, Tnt + 1) = 0
              Exit For
'            ElseIf UBRateTbls(RateRec).TblBreaks(Tnt).UNITAMT <= 0 Then
'              Exit For
            End If
          Else
          If UBRateTbls(RateRec).TblBreaks(Tnt + 1).UNITS > 0 Then
            If UBRateTbls(RateRec).TblBreaks(Tnt).UNITS >= 0 Then
              If UBRateTbls(RateRec).TblBreaks(Tnt).UNITS = 0 Then
                TblBreakfr&(RateRec, Tnt + 1) = UBRateTbls(RateRec).TblBreaks(Tnt).UNITS
                TblBreak&(RateRec, Tnt + 1) = UBRateTbls(RateRec).TblBreaks(Tnt + 1).UNITS - 1
                TblUnitVal#(RateRec, Tnt + 1) = UBRateTbls(RateRec).TblBreaks(Tnt).UNITAMT
                tblbrkchg#(RateRec, Tnt + 1) = 0
                MaxStep(RateRec) = Tnt + 1
              Else
 '               If Tnt > 1 Then
                 TblBreakfr&(RateRec, Tnt + 1) = UBRateTbls(RateRec).TblBreaks(Tnt).UNITS
                 TblBreak&(RateRec, Tnt + 1) = UBRateTbls(RateRec).TblBreaks(Tnt + 1).UNITS - 1
 '               Else
 '                TblBreakfr&(RateRec, Tnt + 1) = minunt(RateRec)
 '                TblBreak&(RateRec, Tnt + 1) = UBRateTbls(RateRec).TblBreaks(Tnt).UNITS
 '               End If
                TblUnitVal#(RateRec, Tnt + 1) = UBRateTbls(RateRec).TblBreaks(Tnt).UNITAMT
                tblbrkchg#(RateRec, Tnt + 1) = 0
                MaxStep(RateRec) = Tnt + 1
              End If
            Else
              MaxStep(RateRec) = Tnt + 1
              TblBreakfr&(RateRec, Tnt + 1) = 0
              TblBreak&(RateRec, Tnt + 1) = 99999999
              TblUnitVal#(RateRec, Tnt + 1) = UBRateTbls(RateRec).TblBreaks(Tnt).UNITAMT
              tblbrkchg#(RateRec, Tnt + 1) = 0
            End If
          Else
            MaxStep(RateRec) = Tnt + 1
            TblBreakfr&(RateRec, Tnt + 1) = UBRateTbls(RateRec).TblBreaks(Tnt).UNITS
            TblBreak&(RateRec, Tnt + 1) = 99999999
            TblUnitVal#(RateRec, Tnt + 1) = UBRateTbls(RateRec).TblBreaks(Tnt).UNITAMT
            tblbrkchg#(RateRec, Tnt + 1) = 0
            Exit For
          End If
          End If
        Next Tnt
      Next RateRec
    Close RateFile
  End If
  If OptBook.Value = True Then
    For Bcnt = 1 To BookCnt
      RptInfo2$ = RptInfo2$ + Str(Book(Bcnt)) + " "
    Next Bcnt
    RptInfo2$ = "By Book - " + RptInfo2$
  ElseIf OptCycle.Value = True Then
    For CCnt = 1 To CycleCnt
      RptInfo2$ = RptInfo2$ + Str(Cycle(CCnt)) + " "
   Next
   RptInfo2$ = "By Cycle - " + RptInfo2$
  ElseIf OptGroup.Value = True Then
    Codefile = FreeFile
    Open CodeName$ For Random As Codefile Len = Len(CodestoRpt)
    NumOfcdsrpt = FileSize(CodeName$) \ Len(CodestoRpt)
    For GCnt = 1 To NumOfcdsrpt
      Get Codefile, GCnt, CodestoRpt
      RptInfo2$ = RptInfo2$ + QPTrim$(CodestoRpt.GroupCode) + " "
    Next
    Close Codefile
    RptInfo2$ = "By Group - " + RptInfo2$
  ElseIf optAllCust.Value = True Then
   RptInfo2$ = "All Customers"
  End If
  GoSub DoRptHeader
    For RCnt = 1 To NumOfratesrpt
      NMinAMT# = MINAMT(RCnt)
      RCode$ = Ratecode(RCnt)
      RateRec = RCnt
      GoSub DoRateHeader
      GoSub DoEachRate
      GoSub DoUnitStepFooter
    Next
  
  GoSub DoGrandFooter
  Close

  Erase TblBreak&, TotalConsp#, TotalCust
  Doall = False
  If RptType = 0 Then
    Load frmLoadingRpt
    frmLoadingRpt.setwherefrom frmRptSCnsmpRateCodeN
    ARptSCnsmpRate.txtDate = Now
    ARptSCnsmpRate.txtTown = TOWNNAME$
    ARptSCnsmpRate.FldDisclaimer = "**Totals are Estimated(Based on Current Settings)."
    ARptSCnsmpRate.Title = "Consumption by Rate Code"
    ARptSCnsmpRate.txtDate1 = txtDate1
    ARptSCnsmpRate.txtDate2 = txtDate2
    ARptSCnsmpRate.FldRptInfo = "Rates - " + RptInfo$ + RptInfo2$
 '   ARptSCnsmpRate.totCust = Using("###,###,###,###", GBCustTot)
    ARptSCnsmpRate.totConsump = Using("###,###,###,###", GBGTMeterConsp#)
   ' ARptSCnsmpRate.totUsage = Using(" $ ##,###,###.##", GBBigUTotal#)
  '  ARptSCnsmpRate.totMin = Using(" $ ##,###,###.##", GBMinGT#)
  '  ARptSCnsmpRate.totcharges = Using(" $ ###,###,###.##", (Round(GBMinGT# + GBBigUTotal#)))
    If PageBrk.Value = 1 Then
      ARptSCnsmpRate.GetName ReportFile$, True
    Else
      ARptSCnsmpRate.GetName ReportFile$, False
    End If
    ARptSCnsmpRate.startrpt
  Else
    ViewPrint ReportFile$, "Consumption by RateCode"
  'KillFile "UBBKCNSP.RPT"
  End If
  Exit Sub
DoEachRate:
  FrmShowPctComp.Label1 = "Processing Rate " + RCode$
  FrmShowPctComp.Show
If OptBook.Value = True Then
'for book

  For cnt& = 1 To NumOfRecs&
    FrmShowPctComp.ShowPctComp cnt, NumOfRecs&
    If FrmShowPctComp.Out Then
      Close
      Unload FrmShowPctComp
      GoTo ExitConsStep
    End If
   Get UBTrans, cnt&, UBTransRec(1)

      If UBTransRec(1).TransDate >= BegDate And UBTransRec(1).TransDate <= EndDate Then
      If UBTransRec(1).TransType = 1 Or UBTransRec(1).TransType = 101 Then
        'Set Valid Flag to Off to Check Each Customer For Valid Rate Code
        ValidCustomer = 0
        If RptType = 1 Then
        If LineCnt > MaxLines Then
          Print #UBRpt, Chr$(12)
          GoSub DoRptHeader
          GoSub DoRateHeader
        End If
        End If
        'IF Valid Customer Then Get the Customer and Check to See if Rate Appl
        CustomerRecord = UBTransRec(1).CustAcctNo

        If CustomerRecord > 0 Then
          Get UBCust, CustomerRecord, UBCustRec(1)
          For Bcnt = 1 To BookCnt
            If Val(UBCustRec(1).Book) <> Book(Bcnt) Then ' And (UBCustRec(1).Status <> "F") Then
              If Bcnt >= BookCnt Then GoTo skipemB
            Else
              Exit For
            End If
          Next Bcnt
 
            For Snt = 1 To 15
              If QPTrim$(UBCustRec(1).serv(Snt).Ratecode) = RCode$ Then
                MTRType$ = UBCustRec(1).serv(Snt).RMtrType
                Select Case MTRType$
                Case "W"
                  MT = 1
                Case "S"
                  MT = 2
                Case "C"
                  MT = 3
                Case "D"
                  MT = 5
                Case "G"
                  MT = 6
                Case "T"
                  MT = 7
                Case "L"
                  MT = 8
                Case "I"
                  MT = 9
                Case Else
                  MT = 4
                End Select
                ValidCustomer = 1
                Exit For
              End If
            Next Snt
        End If
        'IF Customer Has a Rate Code Match then Analyze Meter Type to Get Usag
        If ValidCustomer = 1 Then
          Multi& = 0
          Cubic = False
          For MCnt = 1 To 7
            MType$ = QPTrim$(UBCustRec(1).LocMeters(MCnt).MTRType)
            If MTRType$ = MType$ Then
            'IF LEN(MType$) > 0 THEN
              Multi& = UBCustRec(1).LocMeters(MCnt).MTRMulti
              If Multi& = 0 Then
                Multi& = 1
              End If
              If UBCustRec(1).LocMeters(MCnt).MtrUnit = "C" Then
                Cubic = True
              End If
              Exit For
            End If
          Next
          If Multi& <= 0 Then Multi& = 1
          'IF WhatRev > 0 THEN
            ChkMtr = True
          'ELSE
          '  ChkMtr = False
          'END IF
          For MCnt = 1 To 7
            If ChkMtr = True Then
              If UBTransRec(1).MtrTypes(MCnt) <> MT Then
                GoTo SkipThisMtrB
              End If
            End If
            MeterConsp& = MeterConsp& + UBTransRec(1).CurRead(MCnt) - UBTransRec(1).PrevRead(MCnt)
            If MeterConsp& < 0 Then
              MaxMeterAmt& = 10& ^ (Len(Str$(UBTransRec(1).PrevRead(MCnt))) - 1)
              MeterConsp& = (MaxMeterAmt& - UBTransRec(1).PrevRead(MCnt)) + UBTransRec(1).CurRead(MCnt)
            End If
            If Cubic Then
              MeterConsp& = MeterConsp& * 7.481
            End If
            MeterConsp& = MeterConsp& * Multi&
            'IF MeterConsp& = 1 THEN STOP
            TMeterConsp# = TMeterConsp# + MeterConsp&
            GTMeterConsp# = GTMeterConsp# + MeterConsp&
            ''If MeterConsp& > 0 Then Stop
            'LPRINT CustomerRecord
            'STOP
            'END IF
            NumUser& = UBCustRec(1).LocMeters(MCnt).NumUser
            MeterConsp& = 0
            'END IF
SkipThisMtrB:
          Next MCnt
        End If
        If (ValidCustomer = 1) Then
          
          NonUpdated = 1        'Set Flag to Let Me Know When this Cust Cons U
          GoSub CalcBrkConsump
'          NewMtrConsp# = TMeterConsp#
'          For LL = 1 To MaxStep(RateRec)
'            If NewMtrConsp# > TblBreak&(RateRec, LL - 1) And NewMtrConsp# <= TblBreak&(RateRec, LL) Then
'              NewMtrConsp# = (TMeterConsp# - TblBreak&(RateRec, LL - 1))
'              TotalConsp#(RateRec, LL) = TotalConsp#(RateRec, LL) + NewMtrConsp#
'              TotalConsp#(RateRec, LL - 1) = (TotalConsp#(RateRec, LL - 1) + TblBreak&(RateRec, LL - 1))
'              TotalCust(RateRec, LL) = TotalCust(RateRec, LL) + 1
'              NonUpdated = 0
'              Exit For
'            End If
'          Next LL
'          If NonUpdated = 1 Then
'            TotalConsp#(RateRec, MaxStep(RateRec)) = TotalConsp#(RateRec, MaxStep(RateRec)) + NewMtrConsp#
'            TotalCust(RateRec, MaxStep(RateRec)) = TotalCust(RateRec, MaxStep(RateRec)) + 1
'          End If
        End If
        TMeterConsp# = 0
      End If
    End If
skipemB:
  Next
ElseIf OptCycle.Value = True Then
'for cycle
  For cnt& = 1 To NumOfRecs&
    FrmShowPctComp.ShowPctComp cnt, NumOfRecs&
    If FrmShowPctComp.Out Then
      Close
      Unload FrmShowPctComp
      GoTo ExitConsStep
    End If
   Get UBTrans, cnt&, UBTransRec(1)

  
    If UBTransRec(1).TransDate >= BegDate And UBTransRec(1).TransDate <= EndDate Then
      If UBTransRec(1).TransType = 1 Or UBTransRec(1).TransType = 101 Then
        'Set Valid Flag to Off to Check Each Customer For Valid Rate Code
        ValidCustomer = 0
        If RptType = 1 Then
        If LineCnt > MaxLines Then
          Print #UBRpt, Chr$(12)
          GoSub DoRptHeader
          GoSub DoRateHeader
        End If
        End If
        'IF Valid Customer Then Get the Customer and Check to See if Rate Appl
        CustomerRecord = UBTransRec(1).CustAcctNo

        If CustomerRecord > 0 Then
          Get UBCust, CustomerRecord, UBCustRec(1)
            For CCnt = 1 To CycleCnt
              If Val(UBCustRec(1).BILLCYCL) <> Cycle(CCnt) Then 'And (UBCustRec(1).Status <> "F") Then
                If CCnt >= CycleCnt Then GoTo skipemC
              Else
                Exit For
              End If
            Next

            For Snt = 1 To 15
              If QPTrim$(UBCustRec(1).serv(Snt).Ratecode) = RCode$ Then
                MTRType$ = UBCustRec(1).serv(Snt).RMtrType
                Select Case MTRType$
                Case "W"
                  MT = 1
                Case "S"
                  MT = 2
                Case "C"
                  MT = 3
                Case "D"
                  MT = 5
                Case "G"
                  MT = 6
                Case "T"
                  MT = 7
                Case "L"
                  MT = 8
                Case "I"
                  MT = 9
                Case Else
                  MT = 4
                End Select
                ValidCustomer = 1
                Exit For
              End If
            Next Snt
        End If
        'IF Customer Has a Rate Code Match then Analyze Meter Type to Get Usag
        If ValidCustomer = 1 Then
          Multi& = 0
          Cubic = False
          For MCnt = 1 To 7
            MType$ = QPTrim$(UBCustRec(1).LocMeters(MCnt).MTRType)
            If MTRType$ = MType$ Then
            'IF LEN(MType$) > 0 THEN
              Multi& = UBCustRec(1).LocMeters(MCnt).MTRMulti
              If Multi& = 0 Then
                Multi& = 1
              End If
              If UBCustRec(1).LocMeters(MCnt).MtrUnit = "C" Then
                Cubic = True
              End If
              Exit For
            End If
          Next
          If Multi& <= 0 Then Multi& = 1
          'IF WhatRev > 0 THEN
            ChkMtr = True
          'ELSE
          '  ChkMtr = False
          'END IF
          For MCnt = 1 To 7
            If ChkMtr = True Then
              If UBTransRec(1).MtrTypes(MCnt) <> MT Then
                GoTo SkipThisMtrC
              End If
            End If
            MeterConsp& = MeterConsp& + UBTransRec(1).CurRead(MCnt) - UBTransRec(1).PrevRead(MCnt)
            If MeterConsp& < 0 Then
              MaxMeterAmt& = 10& ^ (Len(Str$(UBTransRec(1).PrevRead(MCnt))) - 1)
              MeterConsp& = (MaxMeterAmt& - UBTransRec(1).PrevRead(MCnt)) + UBTransRec(1).CurRead(MCnt)
            End If
            If Cubic Then
              MeterConsp& = MeterConsp& * 7.481
            End If
            MeterConsp& = MeterConsp& * Multi&
            'IF MeterConsp& = 1 THEN STOP
            TMeterConsp# = TMeterConsp# + MeterConsp&
            GTMeterConsp# = GTMeterConsp# + MeterConsp&
            ''If MeterConsp& > 0 Then Stop
            'LPRINT CustomerRecord
            'STOP
            'END IF
            NumUser& = UBCustRec(1).LocMeters(MCnt).NumUser

            MeterConsp& = 0
            'END IF
SkipThisMtrC:
          Next MCnt
        End If
        If (ValidCustomer = 1) Then
          'NewMtrConsp# = TMeterConsp#
          NonUpdated = 1        'Set Flag to Let Me Know When this Cust Cons U
          GoSub CalcBrkConsump

'          For LL = 1 To MaxStep(RateRec)
'            If NewMtrConsp# > TblBreak&(RateRec, LL - 1) And NewMtrConsp# <= TblBreak&(RateRec, LL) Then
'              NewMtrConsp# = (TMeterConsp# - TblBreak&(RateRec, LL - 1))
'              TotalConsp#(RateRec, LL) = TotalConsp#(RateRec, LL) + NewMtrConsp#
'              TotalConsp#(RateRec, LL - 1) = (TotalConsp#(RateRec, LL - 1) + TblBreak&(RateRec, LL - 1))
'              TotalCust(RateRec, LL) = TotalCust(RateRec, LL) + 1
'              NonUpdated = 0
'              Exit For
'            End If
'          Next LL
'          If NonUpdated = 1 Then
'            TotalConsp#(RateRec, MaxStep(RateRec)) = TotalConsp#(RateRec, MaxStep(RateRec)) + NewMtrConsp#
'            TotalCust(RateRec, MaxStep(RateRec)) = TotalCust(RateRec, MaxStep(RateRec)) + 1
'          End If
        End If
        TMeterConsp# = 0
      End If
    End If
skipemC:
  Next
ElseIf OptGroup.Value = True Then
'for group
  Codefile = FreeFile
  Open CodeName$ For Random As Codefile Len = Len(CodestoRpt)
  NumOfcdsrpt = FileSize(CodeName$) \ Len(CodestoRpt)

  For cnt& = 1 To NumOfRecs&
    FrmShowPctComp.ShowPctComp cnt, NumOfRecs&
    If FrmShowPctComp.Out Then
      Close
      Unload FrmShowPctComp
      GoTo ExitConsStep
    End If
   Get UBTrans, cnt&, UBTransRec(1)


     If UBTransRec(1).TransDate >= BegDate And UBTransRec(1).TransDate <= EndDate Then
      If UBTransRec(1).TransType = 1 Or UBTransRec(1).TransType = 101 Then
        'Set Valid Flag to Off to Check Each Customer For Valid Rate Code
        ValidCustomer = 0
        If RptType = 1 Then
        If LineCnt > MaxLines Then
          Print #UBRpt, Chr$(12)
          GoSub DoRptHeader
          GoSub DoRateHeader
        End If
        End If
        'IF Valid Customer Then Get the Customer and Check to See if Rate Appl
        CustomerRecord = UBTransRec(1).CustAcctNo

        If CustomerRecord > 0 Then
          Get UBCust, CustomerRecord, UBCustRec(1)
            For GCnt = 1 To NumOfcdsrpt
              Get Codefile, GCnt, CodestoRpt
              If UBCustRec(1).GroupCodeRec <> CodestoRpt.RecordNum Then
                If GCnt >= NumOfcdsrpt Then GoTo bskipem
              Else
                Exit For
              End If
            Next
            Grp$ = QPTrim$(CodestoRpt.GroupCode)
            For Snt = 1 To 15
              If QPTrim$(UBCustRec(1).serv(Snt).Ratecode) = RCode$ Then
                MTRType$ = UBCustRec(1).serv(Snt).RMtrType
                Select Case MTRType$
                Case "W"
                  MT = 1
                Case "S"
                  MT = 2
                Case "C"
                  MT = 3
                Case "D"
                  MT = 5
                Case "G"
                  MT = 6
                Case "T"
                  MT = 7
                Case "L"
                  MT = 8
                Case "I"
                  MT = 9
                Case Else
                  MT = 4
                End Select
                ValidCustomer = 1
                Exit For
              End If
            Next Snt
        End If
        'IF Customer Has a Rate Code Match then Analyze Meter Type to Get Usag
        If ValidCustomer = 1 Then
          Multi& = 0
          Cubic = False
          For MCnt = 1 To 7
            MType$ = QPTrim$(UBCustRec(1).LocMeters(MCnt).MTRType)
            If MTRType$ = MType$ Then
            'IF LEN(MType$) > 0 THEN
              Multi& = UBCustRec(1).LocMeters(MCnt).MTRMulti
              If Multi& = 0 Then
                Multi& = 1
              End If
              If UBCustRec(1).LocMeters(MCnt).MtrUnit = "C" Then
                Cubic = True
              End If
              Exit For
            End If
          Next
          If Multi& <= 0 Then Multi& = 1
          'IF WhatRev > 0 THEN
            ChkMtr = True
          'ELSE
          '  ChkMtr = False
          'END IF
          For MCnt = 1 To 7
            If ChkMtr = True Then
              If UBTransRec(1).MtrTypes(MCnt) <> MT Then
                GoTo SkipThisMtrG
              End If
            End If
            MeterConsp& = MeterConsp& + UBTransRec(1).CurRead(MCnt) - UBTransRec(1).PrevRead(MCnt)
            If MeterConsp& < 0 Then
              MaxMeterAmt& = 10& ^ (Len(Str$(UBTransRec(1).PrevRead(MCnt))) - 1)
              MeterConsp& = (MaxMeterAmt& - UBTransRec(1).PrevRead(MCnt)) + UBTransRec(1).CurRead(MCnt)
            End If
            If Cubic Then
              MeterConsp& = MeterConsp& * 7.481
            End If
            MeterConsp& = MeterConsp& * Multi&
            'IF MeterConsp& = 1 THEN STOP
            TMeterConsp# = TMeterConsp# + MeterConsp&
            GTMeterConsp# = GTMeterConsp# + MeterConsp&
            ''If MeterConsp& > 0 Then Stop
            'LPRINT CustomerRecord
            'STOP
            'END IF
            NumUser& = UBCustRec(1).LocMeters(MCnt).NumUser

            MeterConsp& = 0
            'END IF
SkipThisMtrG:
          Next MCnt
        End If
        If (ValidCustomer = 1) Then
'          NewMtrConsp# = TMeterConsp#
          NonUpdated = 1        'Set Flag to Let Me Know When this Cust Cons U
          GoSub CalcBrkConsump

'          For LL = 1 To MaxStep(RateRec)
'            If NewMtrConsp# > TblBreak&(RateRec, LL - 1) And NewMtrConsp# <= TblBreak&(RateRec, LL) Then
'              NewMtrConsp# = (TMeterConsp# - TblBreak&(RateRec, LL - 1))
'              TotalConsp#(RateRec, LL) = TotalConsp#(RateRec, LL) + NewMtrConsp#
'              TotalConsp#(RateRec, LL - 1) = (TotalConsp#(RateRec, LL - 1) + TblBreak&(RateRec, LL - 1))
'              TotalCust(RateRec, LL) = TotalCust(RateRec, LL) + 1
'              NonUpdated = 0
'              Exit For
'            End If
'          Next LL
'          If NonUpdated = 1 Then
'            TotalConsp#(RateRec, MaxStep(RateRec)) = TotalConsp#(RateRec, MaxStep(RateRec)) + NewMtrConsp#
'            TotalCust(RateRec, MaxStep(RateRec)) = TotalCust(RateRec, MaxStep(RateRec)) + 1
'          End If
        End If
        TMeterConsp# = 0
      End If
    End If
bskipem:
  Next
ElseIf optAllCust.Value = True Then
  For cnt& = 1 To NumOfRecs&
    FrmShowPctComp.ShowPctComp cnt, NumOfRecs&
    If FrmShowPctComp.Out Then
      Close
      Unload FrmShowPctComp
      GoTo ExitConsStep
    End If
    Get UBTrans, cnt&, UBTransRec(1)
    If UBTransRec(1).TransDate >= BegDate And UBTransRec(1).TransDate <= EndDate Then
      If UBTransRec(1).TransType = 1 Or UBTransRec(1).TransType = 101 Then
        'Set Valid Flag to Off to Check Each Customer For Valid Rate Code
        ValidCustomer = 0
        If RptType = 1 Then
          If LineCnt > MaxLines Then
            Print #UBRpt, Chr$(12)
            GoSub DoRptHeader
            GoSub DoRateHeader
          End If
        End If
        'IF Valid Customer Then Get the Customer and Check to See if Rate Appl
        CustomerRecord = UBTransRec(1).CustAcctNo

        If CustomerRecord > 0 Then
          Get UBCust, CustomerRecord, UBCustRec(1)
            For Snt = 1 To 15
              If QPTrim$(UBCustRec(1).serv(Snt).Ratecode) = RCode$ Then
                MTRType$ = UBCustRec(1).serv(Snt).RMtrType
                Select Case MTRType$
                Case "W"
                  MT = 1
                Case "S"
                  MT = 2
                Case "C"
                  MT = 3
                Case "D"
                  MT = 5
                Case "G"
                  MT = 6
                Case "T"
                  MT = 7
                Case "L"
                  MT = 8
                Case "I"
                  MT = 9
                Case Else
                  MT = 4
                End Select
                ValidCustomer = 1
                Exit For
              End If
            Next Snt
        End If
        'IF Customer Has a Rate Code Match then Analyze Meter Type to Get Usag
        If ValidCustomer = 1 Then
          Multi& = 0
          Cubic = False
          For MCnt = 1 To 7
            MType$ = QPTrim$(UBCustRec(1).LocMeters(MCnt).MTRType)
            If MTRType$ = MType$ Then
            'IF LEN(MType$) > 0 THEN
              Multi& = UBCustRec(1).LocMeters(MCnt).MTRMulti
              If Multi& = 0 Then
                Multi& = 1
              End If
              If UBCustRec(1).LocMeters(MCnt).MtrUnit = "C" Then
                Cubic = True
              End If
              Exit For
            End If
          Next
          If Multi& <= 0 Then Multi& = 1
          'IF WhatRev > 0 THEN
            ChkMtr = True
          'ELSE
          '  ChkMtr = False
          'END IF
          For MCnt = 1 To 7
            If ChkMtr = True Then
              If UBTransRec(1).MtrTypes(MCnt) <> MT Then
                GoTo SkipThisMtrA
              End If
            End If
            MeterConsp& = MeterConsp& + UBTransRec(1).CurRead(MCnt) - UBTransRec(1).PrevRead(MCnt)
            If MeterConsp& < 0 Then
              MaxMeterAmt& = 10& ^ (Len(Str$(UBTransRec(1).PrevRead(MCnt))) - 1)
              MeterConsp& = (MaxMeterAmt& - UBTransRec(1).PrevRead(MCnt)) + UBTransRec(1).CurRead(MCnt)
            End If
            If Cubic Then
              MeterConsp& = MeterConsp& * 7.481
            End If
            MeterConsp& = MeterConsp& * Multi&
            'IF MeterConsp& = 1 THEN STOP
            TMeterConsp# = TMeterConsp# + MeterConsp&
            GTMeterConsp# = GTMeterConsp# + MeterConsp&
            ''If MeterConsp& > 0 Then Stop
            'LPRINT CustomerRecord
            'STOP
            'END IF
'        AddRevAmt# = 0
'        TMaxAmt# = 0
'        If UBRateTbls(WhatTbl).MaxAmt > 0 Then
'          TMaxAmt# = UBRateTbls(WhatTbl).MaxAmt
'        End If
'        If UBCustRec(1).LocMeters(MeterLocNum).NumUser > 1 Then
'          TMaxAmt# = Round#(UBRateTbls(WhatTbl).MaxAmt * UBCustRec(1).LocMeters(MeterLocNum).NumUser)
          'adjust min consumption for calc below
          NumUser& = UBCustRec(1).LocMeters(MCnt).NumUser
'          AddRevAmt# = NumUser& * UBRateTbls(WhatTbl).MINAMT
'          MinimumConsp& = NumUser& * UBRateTbls(WhatTbl).MINUNITS
'          TMeterConsp& = TMeterConsp& - MinimumConsp&
'          If (TMeterConsp& - UBRateTbls(WhatTbl).MINUNITS) <= 0 Then
'            GoTo GotAmt
'          End If
'        Else
'          NumUser& = 1
'        End If



            MeterConsp& = 0
            'END IF
SkipThisMtrA:
          Next MCnt
        End If
        If (ValidCustomer = 1) Then
'          NewMtrConsp# = TMeterConsp#
          NonUpdated = 1        'Set Flag to Let Me Know When this Cust Cons U
          GoSub CalcBrkConsump
'          For LL = 1 To MaxStep(RateRec)
'            If NewMtrConsp# > TblBreak&(RateRec, LL - 1) And NewMtrConsp# <= TblBreak&(RateRec, LL) Then
'              NewMtrConsp# = (TMeterConsp# - TblBreak&(RateRec, LL - 1))
'              TotalConsp#(RateRec, LL) = TotalConsp#(RateRec, LL) + NewMtrConsp#
'              TotalConsp#(RateRec, LL - 1) = (TotalConsp#(RateRec, LL - 1) + TblBreak&(RateRec, LL - 1))
'              TotalCust(RateRec, LL) = TotalCust(RateRec, LL) + 1
'              NonUpdated = 0
'              Exit For
'            End If
'          Next LL
'          If NonUpdated = 1 Then
'            TotalConsp#(RateRec, MaxStep(RateRec)) = TotalConsp#(RateRec, MaxStep(RateRec)) + NewMtrConsp#
'            TotalCust(RateRec, MaxStep(RateRec)) = TotalCust(RateRec, MaxStep(RateRec)) + 1
'          End If
        End If
        TMeterConsp# = 0
      End If
    End If

  Next


 End If
Return
CalcBrkConsump:
MinBillAmt# = 0
NTMeterConsp# = 0
MinimumConsp& = 0
NewMtrConsp# = 0
        If NumUser& > 1 Then
          numflag = True
          MinBillAmt# = Round(MINAMT#(RateRec) * NumUser&)
        Else
          NumUser& = 1
          numflag = False
          MinBillAmt# = MINAMT#(RateRec)
        End If
          MinimumConsp& = Round(NumUser& * minunt(RateRec))
          NTMeterConsp# = TMeterConsp# - MinimumConsp&
          If NTMeterConsp# <= 0 Then
            NTMeterConsp# = 0
          End If

        numofuser(RateRec) = numofuser(RateRec) + NumUser&
        TotalCust(RateRec) = TotalCust(RateRec) + 1
        NewMtrConsp# = TMeterConsp#
        TAmt# = 0
        If MaxAmt#(RateRec) > 0 Then
          MaxFlag = True
        Else
          MaxFlag = False
        End If
      If MaxStep(RateRec) >= 2 Then
        If numflag Then
          If NTMeterConsp# >= TblBreakfr&(RateRec, 1) And NTMeterConsp# <= TblBreakfr&(RateRec, 2) Then
            UNITS& = NTMeterConsp#
            If MinimumConsp& > TMeterConsp# Then
               TotalConsp#(RateRec, 1) = TotalConsp#(RateRec, 1) + TMeterConsp#
            Else
              TotalConsp#(RateRec, 1) = TotalConsp#(RateRec, 1) + MinimumConsp&
            End If
            NewMtrConsp# = NTMeterConsp#
'            If UNITS& <= 0 Then UNITS& = NTMeterConsp#
'            'TotalConsp#(RateRec, 1) = TotalConsp#(RateRec, 1) + UNITS&
'            If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 1))) >= MaxAmt(RateRec) And MaxFlag Then
'              NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
'              tblbrkchg(RateRec, 1) = (tblbrkchg(RateRec, 1) + NTAmt#)
'            Else
'              tblbrkchg(RateRec, 1) = (tblbrkchg(RateRec, 1) + Round#(UNITS& * TblUnitVal#(RateRec, 1)))
'            End If
'            TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 1)))
'            GoTo GOTIT
          ElseIf NTMeterConsp# = 0 Then
            TotalConsp#(RateRec, 1) = TotalConsp#(RateRec, 1) + NewMtrConsp#
            GoTo GOTIT
          ElseIf NTMeterConsp# > TblBreakfr&(RateRec, 2) Then
            NewMtrConsp# = NTMeterConsp#
          '  UNITS& = (NTMeterConsp# - TblBreak&(RateRec, 1))
            TotalConsp#(RateRec, 1) = TotalConsp#(RateRec, 1) + MinimumConsp&
'            If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 1))) >= MaxAmt(RateRec) And MaxFlag Then
'              NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
'              tblbrkchg(RateRec, 1) = (tblbrkchg(RateRec, 1) + NTAmt#)
'            Else
'              tblbrkchg(RateRec, 1) = (tblbrkchg(RateRec, 1) + Round#(UNITS& * TblUnitVal#(RateRec, 1)))
'            End If
'            TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 1)))
          End If

        End If
        If Not numflag And NewMtrConsp# >= TblBreakfr&(RateRec, 1) And NewMtrConsp# <= TblBreakfr&(RateRec, 2) Then
          UNITS& = (NewMtrConsp# - TblBreak&(RateRec, 1))
          If UNITS& <= 0 Then UNITS& = NewMtrConsp#
          TotalConsp#(RateRec, 1) = TotalConsp#(RateRec, 1) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 1))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 1) = (tblbrkchg(RateRec, 1) + NTAmt#)
          Else
            tblbrkchg(RateRec, 1) = (tblbrkchg(RateRec, 1) + Round#(UNITS& * TblUnitVal#(RateRec, 1)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 1)))
          GoTo GOTIT
        ElseIf Not numflag And NewMtrConsp# < TblBreakfr&(RateRec, 1) Then
          UNITS& = (TblBreak&(RateRec, 1) - TblBreakfr&(RateRec, 1))
          TotalConsp#(RateRec, 1) = TotalConsp#(RateRec, 1) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 1))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 1) = (tblbrkchg(RateRec, 1) + NTAmt#)
          Else
            tblbrkchg(RateRec, 1) = (tblbrkchg(RateRec, 1) + Round#(UNITS& * TblUnitVal#(RateRec, 1)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 1)))
          GoTo GOTIT
        ElseIf Not numflag And NewMtrConsp# > TblBreakfr&(RateRec, 2) Then
          UNITS& = (TblBreak&(RateRec, 1) - TblBreakfr&(RateRec, 1))
          TotalConsp#(RateRec, 1) = TotalConsp#(RateRec, 1) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 1))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 1) = (tblbrkchg(RateRec, 1) + NTAmt#)
          Else
            tblbrkchg(RateRec, 1) = (tblbrkchg(RateRec, 1) + Round#(UNITS& * TblUnitVal#(RateRec, 1)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 1)))
        End If
      Else          'no other rate breaks
        If numflag Then
          UNITS& = NewMtrConsp#
        Else
          UNITS& = (NewMtrConsp# - TblBreak&(RateRec, 1))
        End If
        TotalConsp#(RateRec, 1) = TotalConsp#(RateRec, 1) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 1))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 1) = (tblbrkchg(RateRec, 1) + NTAmt#)
          Else
            tblbrkchg(RateRec, 1) = (tblbrkchg(RateRec, 1) + Round#(UNITS& * TblUnitVal#(RateRec, 1)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 1)))
        GoTo GOTIT
      End If
    
      'Break 2
      If MaxStep(RateRec) >= 3 Then
        If NewMtrConsp# >= TblBreakfr&(RateRec, 2) And NewMtrConsp# <= TblBreakfr&(RateRec, 3) Then
          If numflag Then
            UNITS& = NewMtrConsp#
          Else
            UNITS& = (NewMtrConsp# - TblBreak&(RateRec, 1))
          End If
          TotalConsp#(RateRec, 2) = TotalConsp#(RateRec, 2) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 2))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 2) = (tblbrkchg(RateRec, 2) + NTAmt#)
          Else
            tblbrkchg(RateRec, 2) = (tblbrkchg(RateRec, 2) + Round#(UNITS& * TblUnitVal#(RateRec, 2)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 2)))
          GoTo GOTIT
        ElseIf NewMtrConsp# < TblBreakfr&(RateRec, 2) Then
          If TblBreakfr&(RateRec, 2) < 1 Then
            UNITS& = (TblBreakfr&(RateRec, 3) - 1)
          Else
            UNITS& = (TblBreakfr&(RateRec, 3) - TblBreakfr&(RateRec, 2))
          End If
          TotalConsp#(RateRec, 2) = TotalConsp#(RateRec, 2) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 2))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 2) = (tblbrkchg(RateRec, 2) + NTAmt#)
          Else
            tblbrkchg(RateRec, 2) = (tblbrkchg(RateRec, 2) + Round#(UNITS& * TblUnitVal#(RateRec, 2)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 2)))
          GoTo GOTIT
        ElseIf NewMtrConsp# > TblBreakfr&(RateRec, 3) Then
          If TblBreakfr&(RateRec, 2) < 1 Then
            UNITS& = (TblBreakfr&(RateRec, 3) - 1)
          Else
            UNITS& = (TblBreakfr&(RateRec, 3) - TblBreakfr&(RateRec, 2))
          End If
          TotalConsp#(RateRec, 2) = TotalConsp#(RateRec, 2) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 2))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 2) = (tblbrkchg(RateRec, 2) + NTAmt#)
          Else
            tblbrkchg(RateRec, 2) = (tblbrkchg(RateRec, 2) + Round#(UNITS& * TblUnitVal#(RateRec, 2)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 2)))
        End If
      Else
        If numflag Then
          UNITS& = NewMtrConsp#
        Else
          UNITS& = (NewMtrConsp# - TblBreak&(RateRec, 1))
        End If
        TotalConsp#(RateRec, 2) = TotalConsp#(RateRec, 2) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 2))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 2) = (tblbrkchg(RateRec, 2) + NTAmt#)
          Else
            tblbrkchg(RateRec, 2) = (tblbrkchg(RateRec, 2) + Round#(UNITS& * TblUnitVal#(RateRec, 2)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 2)))
        GoTo GOTIT
      End If
    
      'Break 3
      If MaxStep(RateRec) >= 4 Then
        If NewMtrConsp# >= TblBreakfr&(RateRec, 3) And NewMtrConsp# <= TblBreakfr&(RateRec, 4) Then
          UNITS& = (NewMtrConsp# - TblBreak&(RateRec, 2))
          TotalConsp#(RateRec, 3) = TotalConsp#(RateRec, 3) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 3))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 3) = (tblbrkchg(RateRec, 3) + NTAmt#)
          Else
            tblbrkchg(RateRec, 3) = (tblbrkchg(RateRec, 3) + Round#(UNITS& * TblUnitVal#(RateRec, 3)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 3)))
          GoTo GOTIT
        ElseIf NewMtrConsp# < TblBreakfr&(RateRec, 3) Then
          UNITS& = (TblBreakfr&(RateRec, 4) - TblBreakfr&(RateRec, 3))
          TotalConsp#(RateRec, 3) = TotalConsp#(RateRec, 3) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 3))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 3) = (tblbrkchg(RateRec, 3) + NTAmt#)
          Else
            tblbrkchg(RateRec, 3) = (tblbrkchg(RateRec, 3) + Round#(UNITS& * TblUnitVal#(RateRec, 3)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 3)))
          GoTo GOTIT
        ElseIf NewMtrConsp# > TblBreakfr&(RateRec, 4) Then
          UNITS& = (TblBreakfr&(RateRec, 4) - TblBreakfr&(RateRec, 3))
          TotalConsp#(RateRec, 3) = TotalConsp#(RateRec, 3) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 3))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 3) = (tblbrkchg(RateRec, 3) + NTAmt#)
          Else
            tblbrkchg(RateRec, 3) = (tblbrkchg(RateRec, 3) + Round#(UNITS& * TblUnitVal#(RateRec, 3)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 3)))
        End If
      Else
        UNITS& = (NewMtrConsp# - TblBreak&(RateRec, 2))
        TotalConsp#(RateRec, 3) = TotalConsp#(RateRec, 3) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 3))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 3) = (tblbrkchg(RateRec, 3) + NTAmt#)
          Else
            tblbrkchg(RateRec, 3) = (tblbrkchg(RateRec, 3) + Round#(UNITS& * TblUnitVal#(RateRec, 3)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 3)))
        GoTo GOTIT
      End If
    
      'Break 4
     If MaxStep(RateRec) >= 5 Then
       If NewMtrConsp# >= TblBreakfr&(RateRec, 4) And NewMtrConsp# <= TblBreakfr&(RateRec, 5) Then
         UNITS& = (NewMtrConsp# - TblBreak&(RateRec, 3))
         TotalConsp#(RateRec, 4) = TotalConsp#(RateRec, 4) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 4))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 4) = (tblbrkchg(RateRec, 4) + NTAmt#)
          Else
            tblbrkchg(RateRec, 4) = (tblbrkchg(RateRec, 4) + Round#(UNITS& * TblUnitVal#(RateRec, 4)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 4)))
         GoTo GOTIT
       ElseIf NewMtrConsp# < TblBreakfr&(RateRec, 4) Then
         UNITS& = (TblBreakfr&(RateRec, 5) - TblBreakfr&(RateRec, 4))
         TotalConsp#(RateRec, 4) = TotalConsp#(RateRec, 4) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 4))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 4) = (tblbrkchg(RateRec, 4) + NTAmt#)
          Else
            tblbrkchg(RateRec, 4) = (tblbrkchg(RateRec, 4) + Round#(UNITS& * TblUnitVal#(RateRec, 4)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 4)))
         GoTo GOTIT
       ElseIf NewMtrConsp# > TblBreakfr&(RateRec, 5) Then
         UNITS& = (TblBreakfr&(RateRec, 5) - TblBreakfr&(RateRec, 4))
         TotalConsp#(RateRec, 4) = TotalConsp#(RateRec, 4) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 4))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 4) = (tblbrkchg(RateRec, 4) + NTAmt#)
          Else
            tblbrkchg(RateRec, 4) = (tblbrkchg(RateRec, 4) + Round#(UNITS& * TblUnitVal#(RateRec, 4)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 4)))
       End If
     Else
       UNITS& = (NewMtrConsp# - TblBreak&(RateRec, 3))
       TotalConsp#(RateRec, 4) = TotalConsp#(RateRec, 4) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 4))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 4) = (tblbrkchg(RateRec, 4) + NTAmt#)
          Else
            tblbrkchg(RateRec, 4) = (tblbrkchg(RateRec, 4) + Round#(UNITS& * TblUnitVal#(RateRec, 4)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 4)))
       GoTo GOTIT
     End If
    
     'break 5
     If MaxStep(RateRec) >= 6 Then
       If NewMtrConsp# >= TblBreakfr&(RateRec, 5) And NewMtrConsp# <= TblBreakfr&(RateRec, 6) Then
         UNITS& = (NewMtrConsp# - TblBreak&(RateRec, 4))
         TotalConsp#(RateRec, 5) = TotalConsp#(RateRec, 5) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 5))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 5) = (tblbrkchg(RateRec, 5) + NTAmt#)
          Else
            tblbrkchg(RateRec, 5) = (tblbrkchg(RateRec, 5) + Round#(UNITS& * TblUnitVal#(RateRec, 5)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 5)))
         GoTo GOTIT
        ElseIf NewMtrConsp# < TblBreakfr&(RateRec, 5) Then
          UNITS& = (TblBreakfr&(RateRec, 6) - TblBreakfr&(RateRec, 5))
          TotalConsp#(RateRec, 5) = TotalConsp#(RateRec, 5) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 5))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 5) = (tblbrkchg(RateRec, 5) + NTAmt#)
          Else
            tblbrkchg(RateRec, 5) = (tblbrkchg(RateRec, 5) + Round#(UNITS& * TblUnitVal#(RateRec, 5)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 5)))
          GoTo GOTIT
        ElseIf NewMtrConsp# > TblBreakfr&(RateRec, 6) Then
          UNITS& = (TblBreakfr&(RateRec, 6) - TblBreakfr&(RateRec, 5))
          TotalConsp#(RateRec, 5) = TotalConsp#(RateRec, 5) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 5))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 5) = (tblbrkchg(RateRec, 5) + NTAmt#)
          Else
            tblbrkchg(RateRec, 5) = (tblbrkchg(RateRec, 5) + Round#(UNITS& * TblUnitVal#(RateRec, 5)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 5)))
        End If
      Else
        UNITS& = (NewMtrConsp# - TblBreak&(RateRec, 4))
        TotalConsp#(RateRec, 5) = TotalConsp#(RateRec, 5) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 5))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 5) = (tblbrkchg(RateRec, 5) + NTAmt#)
          Else
            tblbrkchg(RateRec, 5) = (tblbrkchg(RateRec, 5) + Round#(UNITS& * TblUnitVal#(RateRec, 5)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 5)))
        GoTo GOTIT
      End If
    
      'break 6
      If MaxStep(RateRec) >= 7 Then
        If NewMtrConsp# >= TblBreakfr&(RateRec, 6) And NewMtrConsp# <= TblBreakfr&(RateRec, 7) Then
          UNITS& = (NewMtrConsp# - TblBreak&(RateRec, 5))
          TotalConsp#(RateRec, 6) = TotalConsp#(RateRec, 6) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 6))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 6) = (tblbrkchg(RateRec, 6) + NTAmt#)
          Else
            tblbrkchg(RateRec, 6) = (tblbrkchg(RateRec, 6) + Round#(UNITS& * TblUnitVal#(RateRec, 6)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 6)))
          GoTo GOTIT
        ElseIf NewMtrConsp# < TblBreakfr&(RateRec, 6) Then
          UNITS& = (TblBreakfr&(RateRec, 7) - TblBreakfr&(RateRec, 6))
          TotalConsp#(RateRec, 6) = TotalConsp#(RateRec, 6) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 6))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 6) = (tblbrkchg(RateRec, 6) + NTAmt#)
          Else
            tblbrkchg(RateRec, 6) = (tblbrkchg(RateRec, 6) + Round#(UNITS& * TblUnitVal#(RateRec, 6)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 6)))
          GoTo GOTIT
        ElseIf NewMtrConsp# > TblBreakfr&(RateRec, 7) Then
          UNITS& = (TblBreakfr&(RateRec, 7) - TblBreakfr&(RateRec, 6))
          TotalConsp#(RateRec, 6) = TotalConsp#(RateRec, 6) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 6))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 6) = (tblbrkchg(RateRec, 6) + NTAmt#)
          Else
            tblbrkchg(RateRec, 6) = (tblbrkchg(RateRec, 6) + Round#(UNITS& * TblUnitVal#(RateRec, 6)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 6)))
        End If
      Else
        UNITS& = (NewMtrConsp# - TblBreak&(RateRec, 5))
        TotalConsp#(RateRec, 6) = TotalConsp#(RateRec, 6) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 6))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 6) = (tblbrkchg(RateRec, 6) + NTAmt#)
          Else
            tblbrkchg(RateRec, 6) = (tblbrkchg(RateRec, 6) + Round#(UNITS& * TblUnitVal#(RateRec, 6)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 6)))
        GoTo GOTIT
      End If
    
      'break 7
      If MaxStep(RateRec) >= 8 Then
        If NewMtrConsp# >= TblBreakfr&(RateRec, 7) And NewMtrConsp# <= TblBreakfr&(RateRec, 8) Then
          UNITS& = (NewMtrConsp# - TblBreak&(RateRec, 6))
          TotalConsp#(RateRec, 7) = TotalConsp#(RateRec, 7) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 7))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 7) = (tblbrkchg(RateRec, 7) + NTAmt#)
          Else
            tblbrkchg(RateRec, 7) = (tblbrkchg(RateRec, 7) + Round#(UNITS& * TblUnitVal#(RateRec, 7)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 7)))
          GoTo GOTIT
        ElseIf NewMtrConsp# < TblBreakfr&(RateRec, 7) Then
          UNITS& = (TblBreakfr&(RateRec, 8) - TblBreakfr&(RateRec, 7))
          TotalConsp#(RateRec, 7) = TotalConsp#(RateRec, 7) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 7))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 7) = (tblbrkchg(RateRec, 7) + NTAmt#)
          Else
            tblbrkchg(RateRec, 7) = (tblbrkchg(RateRec, 7) + Round#(UNITS& * TblUnitVal#(RateRec, 7)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 7)))
          GoTo GOTIT
        ElseIf NewMtrConsp# > TblBreakfr&(RateRec, 8) Then
          UNITS& = (TblBreakfr&(RateRec, 8) - TblBreakfr&(RateRec, 7))
          TotalConsp#(RateRec, 7) = TotalConsp#(RateRec, 7) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 7))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 7) = (tblbrkchg(RateRec, 7) + NTAmt#)
          Else
            tblbrkchg(RateRec, 7) = (tblbrkchg(RateRec, 7) + Round#(UNITS& * TblUnitVal#(RateRec, 7)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 7)))
        End If
      Else
        UNITS& = (NewMtrConsp# - TblBreak&(RateRec, 6))
        TotalConsp#(RateRec, 7) = TotalConsp#(RateRec, 7) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 7))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 7) = (tblbrkchg(RateRec, 7) + NTAmt#)
          Else
            tblbrkchg(RateRec, 7) = (tblbrkchg(RateRec, 7) + Round#(UNITS& * TblUnitVal#(RateRec, 7)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 7)))
        GoTo GOTIT
      End If
      'break 8
      If MaxStep(RateRec) >= 9 Then
        If NewMtrConsp# >= TblBreakfr&(RateRec, 8) And NewMtrConsp# <= TblBreakfr&(RateRec, 9) Then
          UNITS& = (NewMtrConsp# - TblBreak&(RateRec, 7))
          TotalConsp#(RateRec, 8) = TotalConsp#(RateRec, 8) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 8))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 8) = (tblbrkchg(RateRec, 8) + NTAmt#)
          Else
            tblbrkchg(RateRec, 8) = (tblbrkchg(RateRec, 8) + Round#(UNITS& * TblUnitVal#(RateRec, 8)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 8)))
          GoTo GOTIT
        ElseIf NewMtrConsp# < TblBreakfr&(RateRec, 8) Then
          UNITS& = (TblBreakfr&(RateRec, 9) - TblBreakfr&(RateRec, 8))
          TotalConsp#(RateRec, 8) = TotalConsp#(RateRec, 8) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 8))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 8) = (tblbrkchg(RateRec, 8) + NTAmt#)
          Else
            tblbrkchg(RateRec, 8) = (tblbrkchg(RateRec, 8) + Round#(UNITS& * TblUnitVal#(RateRec, 8)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 8)))
          GoTo GOTIT
        ElseIf NewMtrConsp# > TblBreakfr&(RateRec, 9) Then
          UNITS& = (TblBreakfr&(RateRec, 9) - TblBreakfr&(RateRec, 8))
          TotalConsp#(RateRec, 8) = TotalConsp#(RateRec, 8) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 8))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 8) = (tblbrkchg(RateRec, 8) + NTAmt#)
          Else
            tblbrkchg(RateRec, 8) = (tblbrkchg(RateRec, 8) + Round#(UNITS& * TblUnitVal#(RateRec, 8)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 8)))
        End If
      Else
        UNITS& = (NewMtrConsp# - TblBreak&(RateRec, 7))
        TotalConsp#(RateRec, 8) = TotalConsp#(RateRec, 8) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 8))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 8) = (tblbrkchg(RateRec, 8) + NTAmt#)
          Else
            tblbrkchg(RateRec, 8) = (tblbrkchg(RateRec, 8) + Round#(UNITS& * TblUnitVal#(RateRec, 8)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 8)))
        GoTo GOTIT
      End If
    
      'break 9
      If MaxStep(RateRec) >= 10 Then
        If NewMtrConsp# >= TblBreakfr&(RateRec, 9) And NewMtrConsp# <= TblBreakfr&(RateRec, 10) Then
          UNITS& = (NewMtrConsp# - TblBreak&(RateRec, 8))
          TotalConsp#(RateRec, 9) = TotalConsp#(RateRec, 9) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 9))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 9) = (tblbrkchg(RateRec, 9) + NTAmt#)
          Else
            tblbrkchg(RateRec, 9) = (tblbrkchg(RateRec, 9) + Round#(UNITS& * TblUnitVal#(RateRec, 9)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 9)))
          GoTo GOTIT
        ElseIf NewMtrConsp# < TblBreakfr&(RateRec, 9) Then
          UNITS& = (TblBreakfr&(RateRec, 10) - TblBreakfr&(RateRec, 9))
          TotalConsp#(RateRec, 9) = TotalConsp#(RateRec, 9) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 9))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 9) = (tblbrkchg(RateRec, 9) + NTAmt#)
          Else
            tblbrkchg(RateRec, 9) = (tblbrkchg(RateRec, 9) + Round#(UNITS& * TblUnitVal#(RateRec, 9)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 9)))
          GoTo GOTIT
        ElseIf NewMtrConsp# > TblBreakfr&(RateRec, 10) Then
          UNITS& = (TblBreakfr&(RateRec, 10) - TblBreakfr&(RateRec, 9))
          TotalConsp#(RateRec, 9) = TotalConsp#(RateRec, 9) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 9))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 9) = (tblbrkchg(RateRec, 9) + NTAmt#)
          Else
            tblbrkchg(RateRec, 9) = (tblbrkchg(RateRec, 9) + Round#(UNITS& * TblUnitVal#(RateRec, 9)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 9)))
        End If
      Else
        UNITS& = (NewMtrConsp# - TblBreak&(RateRec, 8))
        TotalConsp#(RateRec, 9) = TotalConsp#(RateRec, 9) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 9))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 9) = (tblbrkchg(RateRec, 9) + NTAmt#)
          Else
            tblbrkchg(RateRec, 9) = (tblbrkchg(RateRec, 9) + Round#(UNITS& * TblUnitVal#(RateRec, 9)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 9)))
       GoTo GOTIT
      End If
    
      If MaxStep(RateRec) >= 11 Then
        If NewMtrConsp# >= TblBreakfr&(RateRec, 10) And NewMtrConsp# <= TblBreakfr&(RateRec, 11) Then
          UNITS& = (NewMtrConsp# - TblBreak&(RateRec, 9))
          TotalConsp#(RateRec, 10) = TotalConsp#(RateRec, 10) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 10))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 10) = (tblbrkchg(RateRec, 10) + NTAmt#)
          Else
            tblbrkchg(RateRec, 10) = (tblbrkchg(RateRec, 10) + Round#(UNITS& * TblUnitVal#(RateRec, 10)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 10)))
          GoTo GOTIT
        ElseIf NewMtrConsp# < TblBreakfr&(RateRec, 10) Then
          UNITS& = (TblBreakfr&(RateRec, 11) - TblBreakfr&(RateRec, 10))
          TotalConsp#(RateRec, 10) = TotalConsp#(RateRec, 10) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 10))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 10) = (tblbrkchg(RateRec, 10) + NTAmt#)
          Else
            tblbrkchg(RateRec, 10) = (tblbrkchg(RateRec, 10) + Round#(UNITS& * TblUnitVal#(RateRec, 10)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 10)))
          GoTo GOTIT
        ElseIf NewMtrConsp# > TblBreakfr&(RateRec, 11) Then
          UNITS& = (TblBreakfr&(RateRec, 11) - TblBreakfr&(RateRec, 10))
          TotalConsp#(RateRec, 10) = TotalConsp#(RateRec, 10) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 10))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 10) = (tblbrkchg(RateRec, 10) + NTAmt#)
          Else
            tblbrkchg(RateRec, 10) = (tblbrkchg(RateRec, 10) + Round#(UNITS& * TblUnitVal#(RateRec, 10)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 10)))
    '      '*****
          UNITS& = (NewMtrConsp# - TblBreak&(RateRec, 10))
          TotalConsp#(RateRec, 11) = TotalConsp#(RateRec, 11) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 11))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 11) = (tblbrkchg(RateRec, 11) + NTAmt#)
          Else
            tblbrkchg(RateRec, 11) = (tblbrkchg(RateRec, 11) + Round#(UNITS& * TblUnitVal#(RateRec, 11)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 11)))
        End If
      Else
        UNITS& = (NewMtrConsp# - TblBreak&(RateRec, 9))
        TotalConsp#(RateRec, 10) = TotalConsp#(RateRec, 10) + UNITS&
        If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 10))) >= MaxAmt(RateRec) And MaxFlag Then
          NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
          tblbrkchg(RateRec, 10) = (tblbrkchg(RateRec, 10) + NTAmt#)
        Else
          tblbrkchg(RateRec, 10) = (tblbrkchg(RateRec, 10) + Round#(UNITS& * TblUnitVal#(RateRec, 10)))
        End If
        TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 10)))
       GoTo GOTIT
      End If
GOTIT:
      TMeterConsp# = 0
Return

DoRptHeader:
If RptType = 1 Then
  PageNo = PageNo + 1
  Print #UBRpt, Tab(29); "Consumption by RateCode"; Tab(70); "Page #"; PageNo
  Print #UBRpt, TOWNNAME$
  Print #UBRpt, "     Report Date: "; Now
  Print #UBRpt, "Period Beginning: "; txtDate1
  Print #UBRpt, "   Period Ending: "; txtDate2
  Print #UBRpt, "      Report Opt: "; "Rates - "; RptInfo$; RptInfo2$
  Print #UBRpt, " "
  Print #UBRpt, "**Totals are Estimated (Based on Current Settings)."
  Print #UBRpt, Dash80$
  LineCnt = 5
End If
Return
DoRateHeader:
  'PageNo = PageNo + 1
'  Print #UBRpt, Tab(29); "Consumption by RateCode"; Tab(70); "Page #"; PageNo
'  Print #UBRpt, TOWNNAME$
'  Print #UBRpt, "Report Date: "; Now
 If RptType = 1 Then
  Print #UBRpt, " "
  Print #UBRpt, "    For Rate Code: "; RCode$
  'Print #UBRpt, " Period Beginning: "; txtDate1
  'Print #UBRpt, "    Period Ending: "; txtDate2
  Print #UBRpt, ; Tab(52); "Usage Charge"; Tab(68); "  Min Charge"
  Print #UBRpt, Dash80$
  LineCnt = LineCnt + 6
End If
Return

DoUnitStepFooter:
  UntPrc = 0
  TblBreak&(RateRec, MaxStep(RateRec)) = 99999999
  If RptType = 1 Then
    Print #UBRpt, "MinAmount - "; Using("###,###.##", MINAMT(RateRec)); "  MinUnits - "; minunt(RateRec);
    If MaxAmt(RateRec) > 0 Then
      Print #UBRpt, "  MaxAmount - "; Using("###,###.##", MaxAmt(RateRec))
    Else
      Print #UBRpt,
    End If
    Print #UBRpt, "--------------------------------------------------------------"
  For LL = 1 To MaxStep(RateRec)
    If LL = 1 Then
      Print #UBRpt, "Min -";
    Else
      Print #UBRpt, "Step # "; LL - 1;
    End If
    Print #UBRpt, Tab(12); "From "; TblBreakfr&(RateRec, LL); " to "; TblBreak&(RateRec, LL)
    Print #UBRpt, "Consumption = "; Tab(18); Using("#########,#", TotalConsp#(RateRec, LL));
    'Print #UBRpt, Tab(29); " # of Trans = "; Using("#####,#", TotalCust(RateRec, LL));
    
    If TblUnitVal#(RateRec, LL) > 0 Then
     'Tempcalccnsp# = Round(TotalCust(RateRec, LL) * minunt(RateRec))
    Else
      TblUnitVal#(RateRec, LL) = 0
    End If
    UntPrc# = tblbrkchg#(RateRec, LL)
    BigUTotal# = Round#(BigUTotal# + UntPrc#)
    'MinGT# = Round#(MinGT# + Round#(NMinAMT# * TotalCust(RateRec)))
    If LL = 1 Then
      Print #UBRpt, Tab(51); Using("###,###,###.##", UntPrc#); Tab(68); Using("  ###,###.##", Round#(NMinAMT# * numofuser(RateRec)))
    Else
      Print #UBRpt, Tab(51); Using("###,###,###.##", UntPrc#)
    End If
    MinGT# = Round#(MinGT# + Round#(NMinAMT# * numofuser(RateRec)))
    BigTotCust = BigTotCust + TotalCust(RateRec)
    'If TotalCust(RateRec, LL) > 0 Then
      
      'PRINT #UBRpt, "  Avg Use= "; USING "#####,#.##"; TotalConsp#(LL) / Tota
   ' Else
   '   Print #UBRpt, ""
   ' End If
    Print #UBRpt, Dash80$
    LineCnt = LineCnt + 4
  Next LL
  Print #UBRpt, "Rate Totals: "; Using("###,###,###,###", GTMeterConsp#); 'Tab(41); "  "; Using("#####,#", BigTotCust);
  Print #UBRpt, Tab(51); Using("###,###,###.##", BigUTotal#); Tab(67); Using("##,###,###.##", Round#(NMinAMT# * numofuser(RateRec)))
  Print #UBRpt, "# of Trans - "; Using("###,###,###", TotalCust(RateRec))
  LineCnt = LineCnt + 2
  If PageBrk = 1 Then
    Print #UBRpt, Chr$(12);
  Else
    Print #UBRpt,
    LineCnt = LineCnt + 1
  End If
  BigTotCust = BigTotCust + TotalCust(RateRec)
  GBBigUTotal# = Round#(GBBigUTotal# + BigUTotal#)
  GBMinGT# = Round#(NMinAMT# * numofuser(RateRec))
  GBGTMeterConsp# = Round#(GBGTMeterConsp# + GTMeterConsp#)
  GBCustTot = GBCustTot + BigTotCust
  BigUTotal# = 0
  MinGT# = 0
  BigTotCust = 0
  GTMeterConsp# = 0

  Else
  If MaxAmt(RateRec) > 0 Then
    ToPrintI$ = "Min Amount - " + Using("###,###.##", MINAMT(RateRec)) + " MinUnits - " + Str(minunt(RateRec)) + "   MaxAmount - " + Using("###,###.##", MaxAmt(RateRec))
  Else
    ToPrintI$ = "Min Amount - " + Using("###,###.##", MINAMT(RateRec)) + " MinUnits - " + Str(minunt(RateRec))
  End If
  For LL = 1 To MaxStep(RateRec)
    If LL = 1 Then
      ToPrint$ = "Min - "
      ToPrint$ = ToPrint$ + "~" + Str(TblBreakfr&(RateRec, LL)) + "~" + Str(TblBreak&(RateRec, LL))
      ToPrint$ = ToPrint$ + "~" + Using("#,###,###,###", TotalConsp#(RateRec, LL))
      ToPrint$ = ToPrint$ + "~" + Using("###,###", TotalCust(RateRec))
    Else
      ToPrint$ = "Step # " + Str(LL - 1)
      ToPrint$ = ToPrint$ + "~" + Str(TblBreakfr&(RateRec, LL)) + "~" + Str(TblBreak&(RateRec, LL))
      ToPrint$ = ToPrint$ + "~" + Using("#,###,###,###", TotalConsp#(RateRec, LL))
      ToPrint$ = ToPrint$ + "~" + " "
    End If
   ' Using("  ###,###.##", Round#(NMinAMT# * Totalcust(RateRec))
    If TblUnitVal#(RateRec, LL) > 0 Then
    Else
      TblUnitVal#(RateRec, LL) = 0
    End If
    UntPrc# = tblbrkchg#(RateRec, LL)
'    If UntPrc# > MaxAmt(RateRec) Then
'      UntPrc# = MaxAmt(RateRec)
'    End If
    BigUTotal# = Round#(BigUTotal# + UntPrc#)
    'MinGT# = Round#(MinGT# + Round#(NMinAMT# * TotalCust(RateRec)))
    If LL = 1 Then
      ToPrint$ = ToPrint$ + "~" + Using("###,###,###.##", UntPrc#) + "~" + Using("  ###,###.##", Round#(NMinAMT# * numofuser(RateRec)))
    Else
      ToPrint$ = ToPrint$ + "~" + Using("###,###,###.##", UntPrc#) + "~" + " "
    End If
    Print #UBRpt, RCode$ + "~" + ToPrint$ + "~" + ToPrintI$
    ToPrint$ = ""
  Next LL
  BigTotCust = BigTotCust + TotalCust(RateRec)
  GBBigUTotal# = Round#(GBBigUTotal# + BigUTotal#)
  GBMinGT# = Round#(NMinAMT# * numofuser(RateRec))
  GBGTMeterConsp# = Round#(GBGTMeterConsp# + GTMeterConsp#)
  GBCustTot = GBCustTot + BigTotCust
  BigUTotal# = 0
  MinGT# = 0
  BigTotCust = 0
  GTMeterConsp# = 0
  End If
Return
DoGrandFooter:
If RptType = 1 Then
  If PageBrk = 1 Then
    GoSub DoRptHeader
  End If
  Print #UBRpt,
 ' Print #UBRpt, "Grand Total Customers     : "; Using("###,###,###,###", GBCustTot)
  Print #UBRpt, "Grand Total Consumption   : "; Using("###,###,###,###", GBGTMeterConsp#)
  Print #UBRpt, "Grand Total Usage Charge  : "; Using(" $ ##,###,###.##", GBBigUTotal#)
  Print #UBRpt, "Grand Total Minimum Charge: "; Using(" $ ##,###,###.##", GBMinGT#)
  Print #UBRpt, "Grand Total Charges       : "; Using(" $ ##,###,###.##", (Round(GBMinGT# + GBBigUTotal#)))
  Print #UBRpt,
End If
Return
  GoTo ExitConsStep
ExitConsStep:
  Close
Exit Sub
End Sub
