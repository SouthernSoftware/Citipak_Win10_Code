VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#3.5#0"; "SPR32X35.ocx"
Begin VB.Form frmBudgetEntEdit 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Budget Entry Edit"
   ClientHeight    =   8640
   ClientLeft      =   30
   ClientTop       =   540
   ClientWidth     =   12195
   Icon            =   "frmBudgetEntEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   12195
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboAcctNumNa 
      Height          =   405
      Left            =   2445
      TabIndex        =   5
      Top             =   2370
      Width           =   5745
      _Version        =   196608
      _ExtentX        =   10134
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
      SearchMethod    =   0
      VirtualMode     =   0   'False
      VRowCount       =   0
      DataSync        =   3
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   16777215
      ThreeDInsideShadowColor=   -2147483627
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   16777215
      ThreeDOutsideShadowColor=   12632256
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   -2147483642
      BorderWidth     =   1
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   0
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   8421504
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
      BorderGrayAreaColor=   14737632
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
      ColDesigner     =   "frmBudgetEntEdit.frx":08CA
   End
   Begin LpLib.fpCombo txtEType 
      Height          =   405
      Left            =   6990
      TabIndex        =   2
      Top             =   1410
      Width           =   1440
      _Version        =   196608
      _ExtentX        =   2540
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
      ThreeDOutsideHighlightColor=   16777215
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
      BorderDropShadowColor=   8421504
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
      GrayAreaColor   =   12632256
      ListLeftOffset  =   0
      ComboGap        =   -2
      MaxEditLen      =   150
      VirtualPageSize =   0
      VirtualPagesAhead=   0
      ExtendCol       =   0
      ColumnLevels    =   1
      ListGrayAreaColor=   14737632
      GroupHeaderHeight=   -1
      GroupHeaderShow =   0   'False
      AllowGrpResize  =   0
      AllowGrpDragDrop=   0
      MergeAdjustView =   0   'False
      ColumnHeaderShow=   0   'False
      ColumnHeaderHeight=   -1
      GrpsFrozen      =   0
      BorderGrayAreaColor=   14737632
      ExtendRow       =   0
      ListPosition    =   1
      ButtonThreeDAppearance=   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      Redraw          =   -1  'True
      AutoSearchFill  =   0   'False
      AutoSearchFillDelay=   500
      EditMarginLeft  =   5
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmBudgetEntEdit.frx":0CF0
   End
   Begin VB.CommandButton cmdExit 
      Appearance      =   0  'Flat
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
      Height          =   540
      Left            =   9552
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7536
      Width           =   1668
   End
   Begin VB.CommandButton cmdDelDist 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "F6 Del &Entry From List"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   5520
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7536
      Width           =   1668
   End
   Begin VB.CommandButton cmdSave 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "F10 &Save Bgt Entries"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   7536
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7536
      Width           =   1668
   End
   Begin VB.CommandButton cmdDelete 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "F3 &Delete"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   588
      Left            =   10224
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2208
      Width           =   900
   End
   Begin VB.CommandButton cmdUpdate 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "F9 &Add To Budget List"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   588
      Left            =   8544
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2208
      Width           =   1596
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   11
      Top             =   8340
      Width           =   12192
      _ExtentX        =   21511
      _ExtentY        =   529
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
            TextSave        =   "1:54 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7117
            TextSave        =   "5/14/2018"
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
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   4155
      Left            =   1080
      TabIndex        =   7
      Top             =   2910
      Width           =   10200
      _Version        =   196613
      _ExtentX        =   18018
      _ExtentY        =   7355
      _StockProps     =   64
      AutoSize        =   -1  'True
      BackColorStyle  =   1
      ButtonDrawMode  =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   13684944
      GridColor       =   8421504
      MaxCols         =   10
      OperationMode   =   3
      ScrollBarExtMode=   -1  'True
      SelectBlockOptions=   0
      ShadowColor     =   13684944
      ShadowDark      =   8421504
      ShadowText      =   0
      SpreadDesigner  =   "frmBudgetEntEdit.frx":1094
      VisibleCols     =   7
      VisibleRows     =   12
   End
   Begin EditLib.fpText txtRef 
      Height          =   372
      Left            =   4272
      TabIndex        =   1
      Top             =   1392
      Width           =   1236
      _Version        =   196608
      _ExtentX        =   2180
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
      ButtonStyle     =   0
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
      AutoCase        =   0
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
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   8
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   14737632
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   14737632
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   8421504
      BorderDropShadowWidth=   3
      ButtonColor     =   12632256
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText txtDesc 
      Height          =   372
      Left            =   5472
      TabIndex        =   4
      Top             =   1872
      Width           =   2940
      _Version        =   196608
      _ExtentX        =   5186
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
      ThreeDInsideHighlightColor=   16777215
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   16777215
      ThreeDOutsideShadowColor=   8421504
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
      ThreeDTextHighlightColor=   16777215
      ThreeDTextShadowColor=   4210752
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
      InvalidColor    =   -2147483643
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   16777215
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   20
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   14737632
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   14737632
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   8421504
      BorderDropShadowWidth=   3
      ButtonColor     =   12632256
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpCurrency txtAmount 
      Height          =   372
      Left            =   2184
      TabIndex        =   3
      Top             =   1872
      Width           =   1812
      _Version        =   196608
      _ExtentX        =   3196
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
      ThreeDInsideHighlightColor=   16777215
      ThreeDInsideShadowColor=   0
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   16777215
      ThreeDOutsideShadowColor=   12632256
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
      ButtonStyle     =   0
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   16777215
      ThreeDTextShadowColor=   12632256
      ThreeDTextOffset=   1
      AlignTextH      =   2
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
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   2
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   "$"
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "999999999"
      MinValue        =   "-999999999"
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
      IncDec          =   1
      BorderGrayAreaColor=   8421504
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   8421504
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   12632256
      BorderDropShadowWidth=   3
      ButtonColor     =   12632256
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpDateTime txtDate 
      Height          =   372
      Left            =   1968
      TabIndex        =   0
      Top             =   1392
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
      ThreeDInsideHighlightColor=   16777215
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   -2147483628
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
      ThreeDTextHighlightColor=   16777215
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
      ThreeDFrameColor=   8421504
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
   Begin EditLib.fpCurrency txtDebit 
      Height          =   372
      Left            =   8208
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   6960
      Width           =   1476
      _Version        =   196608
      _ExtentX        =   2603
      _ExtentY        =   656
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
      BackColor       =   16777215
      ForeColor       =   0
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
      ThreeDTextHighlightColor=   14737632
      ThreeDTextShadowColor=   8421504
      ThreeDTextOffset=   1
      AlignTextH      =   2
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
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   1
      ControlType     =   1
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   2
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   "$"
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
      BorderGrayAreaColor=   12632256
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   12632256
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   8421504
      BorderDropShadowWidth=   3
      ButtonColor     =   12632256
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpCurrency txtCredit 
      Height          =   372
      Left            =   9648
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   6960
      Width           =   1476
      _Version        =   196608
      _ExtentX        =   2603
      _ExtentY        =   656
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
      BackColor       =   16777215
      ForeColor       =   0
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   12632256
      ThreeDInsideShadowColor=   0
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   16777215
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
      ButtonStyle     =   0
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   14737632
      ThreeDTextShadowColor=   8421504
      ThreeDTextOffset=   1
      AlignTextH      =   2
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
      ControlType     =   1
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   2
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   "$"
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
      BorderGrayAreaColor=   12632256
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   12632256
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   8421504
      BorderDropShadowWidth=   3
      ButtonColor     =   14737632
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fptxtDesc2 
      Height          =   300
      Left            =   8544
      TabIndex        =   10
      Top             =   1680
      Width           =   2556
      _Version        =   196608
      _ExtentX        =   4508
      _ExtentY        =   529
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   7.5
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
      ThreeDOutsideShadowColor=   8421504
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
      ThreeDTextHighlightColor=   -2147483637
      ThreeDTextShadowColor=   8421504
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
      InvalidColor    =   -2147483643
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483643
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   32
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   8421504
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin VB.Label Label3b 
      BackStyle       =   0  'Transparent
      Caption         =   "(Opt.)Additional Desc."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   2
      Left            =   8544
      TabIndex        =   25
      Top             =   1392
      Width           =   2556
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Double-Click Row To Edit Information."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   300
      Left            =   720
      TabIndex        =   24
      Top             =   7584
      Width           =   4332
   End
   Begin VB.Label Label3b 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Totals"
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
      Index           =   1
      Left            =   6960
      TabIndex        =   23
      Top             =   7008
      Width           =   948
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00D0D0D0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      FillColor       =   &H00D0D0D0&
      Height          =   4068
      Left            =   1056
      Top             =   2880
      Width           =   10140
   End
   Begin VB.Label Label4b 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
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
      Index           =   0
      Left            =   3936
      TabIndex        =   20
      Top             =   1920
      Width           =   1452
   End
   Begin VB.Label Label2b 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
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
      Index           =   1
      Left            =   1056
      TabIndex        =   19
      Top             =   1920
      Width           =   972
   End
   Begin VB.Label Label4b 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Entry Type"
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
      Index           =   1
      Left            =   5664
      TabIndex        =   18
      Top             =   1440
      Width           =   1212
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
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
      Height          =   300
      Index           =   1
      Left            =   1152
      TabIndex        =   17
      Top             =   1488
      Width           =   612
   End
   Begin VB.Label Label3b 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ref"
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
      Index           =   0
      Left            =   3504
      TabIndex        =   16
      Top             =   1440
      Width           =   684
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "GL Account"
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
      Index           =   0
      Left            =   984
      TabIndex        =   15
      Top             =   2400
      Width           =   1356
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      Height          =   852
      Left            =   2928
      Top             =   240
      Width           =   6492
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Budget Entry Edit"
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
      Left            =   3792
      TabIndex        =   14
      Top             =   480
      Width           =   4812
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D0D0D0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      Height          =   972
      Left            =   2928
      Top             =   144
      Width           =   6492
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000005&
      BorderWidth     =   3
      Height          =   6156
      Left            =   1056
      Top             =   1200
      Width           =   10140
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
Attribute VB_Name = "frmBudgetEntEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim GLSetup As GLSetupRecType
Dim GLAcctidx As GLAcctIndexType
Dim GLAcct As GLAcctRecType
Dim BgtEdit As TrEditRecType
Dim Emode As Boolean, RecNum As Integer, Verify As String, Change As Boolean
Dim FY1BegDate As Integer, FY1EndDate As Integer, FY2BegDate As Integer, FY2EndDate As Integer
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Dim BgtEditFileNum As Integer, NumEdTrans As Integer
'This is to fix spreadsheet for various resolutions
Public Function Fixspread()
    Select Case screenW
      Case 1280
      If Screen.TwipsPerPixelX <> 12 Then
        coladj = 9.8
        vaSpread1.RowHeight(-1) = 22.6
      Else
        coladj = 6.1
        vaSpread1.RowHeight(-1) = 18.5
      End If
      Case 1152
      If Screen.TwipsPerPixelX <> 12 Then
        coladj = 8
        vaSpread1.RowHeight(-1) = 20
      Else
        coladj = 4.6
        vaSpread1.RowHeight(-1) = 15.5
      End If
      Case 1024
      If Screen.TwipsPerPixelX <> 12 Then
        coladj = 6
        vaSpread1.RowHeight(0) = 18
        vaSpread1.RowHeight(-1) = 18
      Else
        coladj = 3.1
      End If
      Case 800
        coladj = 2.9
        'vaSpread1.Font.Size = 8
        vaSpread1.RowHeight(-1) = 14
      Case Else
        'don't worry be happpy
    End Select
    vaSpread1.ColWidth(-1) = vaSpread1.ColWidth(-1) + coladj
End Function
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbYes Then
        If VerifyEntered = True Then
          If MsgBox("          Select OK To Close And Abandon Entry." & Chr(13) & Chr(13) & "CANCEL to Remain on Current Screen.", vbOKCancel, "Abandon And Close?") = vbCancel Then
            Cancel = True
          End If
        Else
          If Change = True Then
            If MsgBox("Close Without Saving Changes, Yes or No?", vbYesNo, "Budget Entry") = vbNo Then
              Cancel = True
            End If
          End If
        End If
        Close BgtEditFileNum
        KillFileD "BGTED.opn"
        ClearInUse PWcnt
      Else
        Cancel = True
      End If
    End If
  End If
End Sub

Private Sub Form_Load()
  Dim cnt As Integer, CntB As Integer
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  GetFYDates FY1BegDate, FY1EndDate, FY2BegDate, FY2EndDate
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  Fixspread
  StatusBar1.Panels.Item(1).Text = GLUserName
  Me.HelpContextID = hlpEditBudgetMenu
  BudAcctNumName fpcboAcctNumNa
  OpenBgtEditFile BgtEditFileNum, NumEdTrans
  If BgtEditFileNum < 0 Then
    frmBudgetMaintMenu.Show
    Unload frmBudgetEntEdit
    Exit Sub
  End If
  Change = False
  If NumEdTrans > 0 Then
    For cnt = 1 To NumEdTrans
      Get BgtEditFileNum, cnt, BgtEdit
      If BgtEdit.Deleted <> 0 Then
        CntB = CntB + 1
      Else
        Emode = True
        Exit For
      End If
    Next
  Else
    Emode = False
  End If
  If CntB = NumEdTrans Then
    Close BgtEditFileNum
    KillFile "BGTED.DAT"
    OpenBgtEditFile BgtEditFileNum, NumEdTrans
    If BgtEditFileNum < 0 Then
      frmBudgetMaintMenu.Show
      Unload frmBudgetEntEdit
      Exit Sub
    End If
  End If
  If Emode = True Then
    Rec2Form
  Else
    RecNum = 1
    txtDebit = 0
    txtCredit = 0
  End If
  txtDate.Text = Format(Now, "mm/dd/yyyy")
  txtDesc = ""
  fptxtDesc2 = ""
  txtRef = ""
  txtAmount = 0
  txtEType.AddItem "Increase"
  txtEType.AddItem "Decrease"
  txtEType.ListIndex = -1
  fpcboAcctNumNa.ListIndex = -1
  '***** spreadsheet do Not have to set blank fields on load ..
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    ''Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub

Private Sub fpcboAcctNumNa_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboAcctNumNa.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcboAcctNumNa.ListIndex = -1
    fpcboAcctNumNa.Action = ActionClearSearchBuffer
  End If
  If fpcboAcctNumNa.ListDown <> True Then
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
Private Sub fpcboAcctNumNa_LostFocus()
  fpcboAcctNumNa.Action = ActionClearSearchBuffer
End Sub

Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub

Private Sub txtDate_LostFocus()
  If CheckValDate(txtDate) = False Then
    MsgBox "Invalid Date, Please Retry.", vbOKOnly, "Budget Entry"
    txtDate.SetFocus
  End If
End Sub
Private Sub cmdDelDist_Click()
  If vaSpread1.ActiveRow > 0 Then
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = 4
    If vaSpread1.Text <> "" Then
      If MsgBox("You Wish to Delete Selected Entry?", vbYesNo, "Delete Entry") = vbYes Then
        vaSpread1.Col = 8
        txtDebit = (txtDebit.DoubleValue - vaSpread1.Text)
        vaSpread1.Col = 9
        txtCredit = (txtCredit.DoubleValue - vaSpread1.Text)
        vaSpread1.DeleteRows vaSpread1.Row, 1
        fpcboAcctNumNa.SetFocus
      End If
    End If
  End If

End Sub

Private Sub cmdDelete_Click()
  If VerifyEntered = True Then
   Change = True
   If MsgBox("Are you sure you wish to delete this entry?", vbYesNo, "Delete BgtEntry") = vbYes Then
      Call NextNew
    Else
      txtDate.SetFocus
      Exit Sub
    End If
    
  Else
    MsgBox "You Must First Select An Entry To Delete.", vbOKOnly, "Deletion Denied"
  End If
End Sub

Private Sub cmdUpdate_Click()
  Dim AcctRec As Integer, EType As String
  If VerifyEntered = True Then
    If fpcboAcctNumNa.Text <> "" And txtAmount.DoubleValue > 0 Then
      vaSpread1.Row = vaSpread1.DataRowCnt + 1
      vaSpread1.Col = 1
      vaSpread1.Text = txtDate
      vaSpread1.Col = 2
      fpcboAcctNumNa.Col = 0
      AcctRec = Val(fpcboAcctNumNa.ColText)
      vaSpread1.Text = fpcboAcctNumNa.ColText
      vaSpread1.Col = 3
      fpcboAcctNumNa.Col = 1
      vaSpread1.Text = fpcboAcctNumNa.ColText
      vaSpread1.Col = 4
      fpcboAcctNumNa.Col = 2
      vaSpread1.Text = fpcboAcctNumNa.ColText
      vaSpread1.Col = 5
      vaSpread1.Text = txtRef
      vaSpread1.Col = 6
      vaSpread1.Text = txtDesc
      vaSpread1.Col = 7
      vaSpread1.Text = Mid$(txtEType.Text, 1, 1)
      EType = vaSpread1.Text
      If GetType(AcctRec, EType) = True Then
        vaSpread1.Col = 8
        vaSpread1.Text = txtAmount.DoubleValue
        txtDebit = (txtDebit.DoubleValue + txtAmount.DoubleValue)
        vaSpread1.Col = 9
        vaSpread1.Text = 0
      Else
        vaSpread1.Col = 8
        vaSpread1.Text = 0
        vaSpread1.Col = 9
        vaSpread1.Text = txtAmount.DoubleValue
        txtCredit = (txtCredit.DoubleValue + txtAmount.DoubleValue)
      End If
      vaSpread1.Col = 10
      vaSpread1.Text = QPTrim$(fptxtDesc2)
      Change = True
      NextNew
    End If
  Else
    MsgBox Verify$, vbOKOnly, "Budget Entry Error"
  End If
End Sub
Public Function GetType(AcctRec As Integer, EType As String)
  Dim AcctFile As Integer, NumAccts As Integer
  OpenAcctFile AcctFile
  NumAccts = LOF(AcctFile) / Len(GLAcct)
    Get AcctFile, AcctRec, GLAcct
    If GLAcct.Deleted = 0 Then
      If EType = "I" Then
        If GLAcct.Typ = "E" Then
          GetType = True
        Else
          GetType = False
        End If
      End If
      If EType = "D" Then
        If GLAcct.Typ = "E" Then
          GetType = False
        Else
          GetType = True
        End If
      End If
       
    End If
  Close AcctFile
End Function
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyUp:
      SendKeys "+{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      cmdExit_Click
      KeyCode = 0
    Case vbKeyF6:
      cmdDelDist_Click
      KeyCode = 0
    Case vbKeyF3:
      cmdDelete_Click
      KeyCode = 0
    Case vbKeyF10:
      cmdSave_Click
      KeyCode = 0
    Case vbKeyF9:
      cmdUpdate_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub
Private Sub cmdExit_Click()
  If MsgBox("Are You Sure You Wish To Exit?", vbYesNo, "Budget Entry") = vbYes Then
  If VerifyEntered = True Then
    If MsgBox("          Select OK To Abandon Entry." & Chr(13) & Chr(13) & "CANCEL to Remain on Current Screen.", vbOKCancel, "Abandon Current Entry?") = vbCancel Then
      txtDate.SetFocus
      Exit Sub
    Else
      cmdDelete_Click
      Exit Sub
    End If
  End If
  If Change = False Then
    frmBudgetMaintMenu.Show
    Unload frmBudgetEntEdit
    Close BgtEditFileNum
  Else
    If MsgBox("Save Changes Before Exiting, Yes or No?", vbYesNo, "Budget Entry") = vbYes Then
       SaveBudgetList
    End If
    frmBudgetMaintMenu.Show
    Unload frmBudgetEntEdit
    Close BgtEditFileNum
  End If
  KillFileD "BGTED.opn"
  End If
End Sub
Private Sub cmdSave_Click()
  

  If VerifyEntered = True Then
    If MsgBox("         Select OK to Abandon Entry." & Chr(13) & Chr(13) & "CANCEL to Remain on Current Screen.", vbOKCancel, "Abandon Current Entry?") = vbCancel Then
      txtDate.SetFocus
      Exit Sub
    Else
      cmdDelete_Click
      Exit Sub
    End If
  End If
  
'    If vaSpread1.DataRowCnt = 0 Then
'      If MsgBox("                  The Budget List Is Empty." & Chr(13) & Chr(13) & "Select YES And The Edit File Will Be Cleared" & Chr(13) & Chr(13) & "              Or NO To Abandon Changes.", vbYesNo, "Budget Entry") = vbNo Then
'        frmBudgetMaintMenu.Show
'        Unload frmBudgetEntEdit
'        Close BgtEditFileNum
'        Exit Sub
'      End If
'    End If
    SaveBudgetList
    MsgBox "Your Budget Entries Have Been Saved.", vbOKOnly, "Budget Saved"
  KillFileD "BGTED.opn"
  frmBudgetMaintMenu.Show
  Unload frmBudgetEntEdit
  'Close is done in SaveBudgetList procedure
End Sub

Private Sub SaveBudgetList()
  Dim cnt As Integer
 ' If Exist("BGTED.DAT") Then
    Close BgtEditFileNum
    KillFile "BGTED.DAT"
  'End If
  OpenBgtEditFile BgtEditFileNum, NumEdTrans
  If BgtEditFileNum < 0 Then
    frmBudgetMaintMenu.Show
    Unload frmBudgetEntEdit
    Exit Sub
  End If
  If vaSpread1.DataRowCnt > 0 Then
    For cnt = 1 To vaSpread1.DataRowCnt
      vaSpread1.Row = cnt
      BgtEdit.Deleted = 0
      vaSpread1.Col = 1
      BgtEdit.TRDATE = DateDiff("d", "12/31/1979", vaSpread1.Text)
      vaSpread1.Col = 2
      BgtEdit.AcctRec = QPTrim(vaSpread1.Text)
      vaSpread1.Col = 3
      BgtEdit.AcctNum = QPTrim(vaSpread1.Text)
      vaSpread1.Col = 4
      BgtEdit.AcctName = QPTrim(vaSpread1.Text)
      vaSpread1.Col = 5
      BgtEdit.Ref = QPTrim(vaSpread1.Text)
      vaSpread1.Col = 6
      BgtEdit.Desc = QPTrim(vaSpread1.Text)
      vaSpread1.Col = 7
      BgtEdit.EType = QPTrim(vaSpread1.Text)
      vaSpread1.Col = 8
      BgtEdit.DrAmt = QPTrim(vaSpread1.Text)
      vaSpread1.Col = 9
      BgtEdit.CrAmt = QPTrim(vaSpread1.Text)
      vaSpread1.Col = 10
      BgtEdit.LDesc = QPTrim(vaSpread1.Text)
      RecNum = cnt
      Put BgtEditFileNum, RecNum, BgtEdit
    Next
  Close BgtEditFileNum
  Call MainLog("Budget Saved.")
  Else
    Close BgtEditFileNum
    
  End If
End Sub

Public Function Rec2Form()
  Dim AcctRec As Integer
  Dim CurrRec As Integer, NextRec As Integer, cnt As Integer, Last As Integer
  'OpenBgtEditFile BgtEditFileNum, NumEdTrans
  For RecNum = 1 To NumEdTrans
    Get BgtEditFileNum, RecNum, BgtEdit
    If BgtEdit.Deleted = 0 Then
      AcctRec = AcctFind(BgtEdit.AcctNum)
      If AcctRec > 0 Then
        vaSpread1.Row = vaSpread1.DataRowCnt + 1
        vaSpread1.Col = 1
        vaSpread1.Text = Format(DateAdd("d", (BgtEdit.TRDATE), "12-31-1979"), "mm/dd/yyyy")
        vaSpread1.Col = 2
        vaSpread1.Text = AcctRec
        vaSpread1.Col = 3
        vaSpread1.Text = BgtEdit.AcctNum
        vaSpread1.Col = 4
        vaSpread1.Text = BgtEdit.AcctName
        vaSpread1.Col = 5
        vaSpread1.Text = BgtEdit.Ref
        vaSpread1.Col = 6
        vaSpread1.Text = BgtEdit.Desc
        vaSpread1.Col = 7
        vaSpread1.Text = BgtEdit.EType
        vaSpread1.Col = 8
        vaSpread1.Text = BgtEdit.DrAmt
        vaSpread1.Col = 9
        vaSpread1.Text = BgtEdit.CrAmt
        vaSpread1.Col = 10
        vaSpread1.Text = BgtEdit.LDesc
        txtDebit = (txtDebit.DoubleValue + BgtEdit.DrAmt)
        txtCredit = (txtCredit.DoubleValue + BgtEdit.CrAmt)
      Else
        MsgBox "An Invalid Account Record Was Encountered And Will Not Be Loaded.", vbOKOnly, "Invalid Account"
      End If
    End If
  Next
  Emode = True
End Function

Private Sub NextNew()
'  Dim BgtEditFileNum As Integer, NumEdTrans As Integer
'  OpenBgtEditFile BgtEditFileNum, NumEdTrans
'  Close BgtEditFileNum
'   If NumEdTrans > 0 Then
'     RecNum = NumEdTrans + 1
'   Else
'     RecNum = 1
'   End If

   Emode = False
   ClearFields
   txtDate.SetFocus
End Sub

Private Function VerifyEntered()
  Dim TempDate As Integer
  If CheckValDate(txtDate) = True Then
  TempDate = DateDiff("d", "12/31/1979", txtDate)
    If txtEType.Text <> "" Then
      If txtAmount > 0 Then
        If txtDesc <> "" Then
          If fpcboAcctNumNa.Text <> "" Then
            VerifyEntered = True
            'Also compare date with First Fiscal Year range
            If (TempDate < FY1BegDate) Or (TempDate > FY1EndDate) Then
              Verify$ = "This Date Is Not Within First Fiscal Year Date Range. Please Correct."
              VerifyEntered = False
            End If
          Else
            VerifyEntered = False
            Verify$ = "An Account Must Be Selected Before Adding To List."
            fpcboAcctNumNa.SetFocus
            Exit Function
          End If
        Else
          VerifyEntered = False
          Verify$ = "Description May Not Be Left Blank."
          txtDesc.SetFocus
          Exit Function
        End If
      Else
        VerifyEntered = False
        Verify$ = "Amount May Not Be Zero."
        txtAmount.SetFocus
        Exit Function
      End If
    Else
      VerifyEntered = False
      Verify$ = "You Must Select Either Increase or Decrease."
      txtEType.SetFocus
      Exit Function
    End If
  Else
    VerifyEntered = False
    Verify$ = "Invalid Date."
    txtDate.SetFocus
    Exit Function
  End If
End Function


Private Sub txtEType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    txtEType.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    txtEType.ListIndex = -1
    txtEType.Action = ActionClearSearchBuffer
  End If
  If txtEType.ListDown <> True Then
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

Private Sub vaSpread1_DblClick(ByVal Col As Long, ByVal Row As Long)
  Dim TempAcct As String
  Dim TempCol As Long, TempRow As Long
  TempRow = Row
  TempCol = Col
  vaSpread1.Col = 1
  vaSpread1.Row = Row
  If TempRow > 0 And vaSpread1.Text <> "" Then
    If fpcboAcctNumNa.Text <> "" Then
      If MsgBox("Overwrite and Delete Current Entry, 'OK' or 'Cancel'", vbOKCancel, "Budget Entry") = vbCancel Then
        Exit Sub
      Else
       If MsgBox("Are you sure you wish to delete this entry?", vbYesNo, "Delete BgtEntry") = vbYes Then
        Call NextNew
       Else
        txtDate.SetFocus
        Exit Sub
       End If

      End If
    End If
    vaSpread1.Row = TempRow
    vaSpread1.Col = 3
    TempAcct = QPTrim(vaSpread1.Text)
    If vaSpread1.Text <> "" Then
      fpcboAcctNumNa.SearchText = QPStrip(TempAcct)
      fpcboAcctNumNa.Action = 0
      If fpcboAcctNumNa.SearchIndex <> -1 Then
        fpcboAcctNumNa.ListIndex = fpcboAcctNumNa.SearchIndex
      End If
      vaSpread1.Col = 2
      fpcboAcctNumNa.Col = 0
      fpcboAcctNumNa.ColText = vaSpread1.Text
      vaSpread1.Col = 3
      fpcboAcctNumNa.Col = 1
      fpcboAcctNumNa.ColText = vaSpread1.Text
      vaSpread1.Col = 4
      fpcboAcctNumNa.Col = 2
      fpcboAcctNumNa.ColText = vaSpread1.Text
      vaSpread1.Col = 1
      txtDate = vaSpread1.Text
      vaSpread1.Col = 5
      txtRef = vaSpread1.Text
      vaSpread1.Col = 6
      txtDesc = vaSpread1.Text
      vaSpread1.Col = 7
      If vaSpread1.Text = "I" Then
        txtEType.Text = "Increase"
      Else
        txtEType.Text = "Decrease"
      End If
      vaSpread1.Col = 8
      If vaSpread1.Value > 0 Then
        txtAmount = vaSpread1.Text
        txtDebit = (txtDebit.DoubleValue - txtAmount.DoubleValue)
      Else
        vaSpread1.Col = 9
        txtAmount = vaSpread1.Text
        txtCredit = (txtCredit.DoubleValue - txtAmount.DoubleValue)
      End If
      vaSpread1.Col = 10
      fptxtDesc2 = vaSpread1.Text
      'vaSpread1.ClearRange TempCol, TempRow, 4, TempRow, True
      vaSpread1.DeleteRows TempRow, 1
      fpcboAcctNumNa.SetFocus
    End If
  End If
End Sub
Public Sub ClearFields()
  txtRef = ""
  txtAmount = 0
  txtEType.Text = ""
  fpcboAcctNumNa.ListIndex = -1
  'vaSpread1.ClearRange 1, 1, 4, 36, True
End Sub

