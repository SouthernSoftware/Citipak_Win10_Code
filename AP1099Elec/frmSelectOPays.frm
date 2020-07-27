VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSelectOPays 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Open Payables"
   ClientHeight    =   8844
   ClientLeft      =   36
   ClientTop       =   540
   ClientWidth     =   12192
   Icon            =   "frmSelectOPays.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8844
   ScaleWidth      =   12192
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboVendFirst 
      Height          =   360
      Left            =   5088
      TabIndex        =   0
      Top             =   3192
      Width           =   4140
      _Version        =   196608
      _ExtentX        =   7302
      _ExtentY        =   635
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.8
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
      ColumnSearch    =   0
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
      EditMarginLeft  =   5
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmSelectOPays.frx":08CA
   End
   Begin LpLib.fpCombo fpcboVendLast 
      Height          =   360
      Left            =   5112
      TabIndex        =   1
      Top             =   3840
      Width           =   4140
      _Version        =   196608
      _ExtentX        =   7302
      _ExtentY        =   635
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.8
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
      ColumnSearch    =   0
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
      EditMarginLeft  =   5
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmSelectOPays.frx":0D05
   End
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   4020
      Left            =   2328
      TabIndex        =   12
      Top             =   2232
      Visible         =   0   'False
      Width           =   7548
      _Version        =   196609
      _ExtentX        =   13314
      _ExtentY        =   7091
      _StockProps     =   70
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14737632
      Caption         =   ""
      Picture         =   "frmSelectOPays.frx":1140
      Begin LpLib.fpList lstInvoices 
         Height          =   2688
         Left            =   216
         TabIndex        =   20
         Top             =   360
         Width           =   7116
         _Version        =   196608
         _ExtentX        =   12552
         _ExtentY        =   4741
         TextAlias       =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   10.8
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
         Columns         =   6
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
         ColDesigner     =   "frmSelectOPays.frx":115C
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00D0D0D0&
         Caption         =   "Alt-C &Clear All"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   396
         Left            =   5856
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   3504
         UseMaskColor    =   -1  'True
         Width           =   1500
      End
      Begin VB.CommandButton cmdMark 
         BackColor       =   &H00D0D0D0&
         Caption         =   "Alt-M &Mark All"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   396
         Left            =   4272
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   3504
         UseMaskColor    =   -1  'True
         Width           =   1500
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Vendor"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   228
         Left            =   648
         TabIndex        =   17
         Top             =   24
         Width           =   660
      End
      Begin VB.Shape Shape2 
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   3156
         Left            =   120
         Top             =   240
         Width           =   7308
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   3
         Left            =   6084
         TabIndex        =   16
         Top             =   24
         Width           =   852
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   1
         Left            =   4068
         TabIndex        =   15
         Top             =   24
         Width           =   1308
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   7
         Left            =   2652
         TabIndex        =   14
         Top             =   24
         Width           =   732
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Spacebar or Click to Toggle, F10 to Continue. "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   516
         Left            =   144
         TabIndex        =   13
         Top             =   3432
         Width           =   4140
      End
   End
   Begin VB.CommandButton cmdAll 
      Caption         =   "F2 &All"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Left            =   8448
      TabIndex        =   7
      Top             =   5160
      Width           =   900
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
      Height          =   468
      Left            =   7860
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7152
      Width           =   1236
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
      Height          =   468
      Left            =   9720
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7152
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
            TextSave        =   "4:58 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7133
            TextSave        =   "3/7/2005"
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
   Begin EditLib.fpDateTime fpDueDate 
      Height          =   348
      Left            =   5112
      TabIndex        =   2
      Top             =   4800
      Width           =   1620
      _Version        =   196608
      _ExtentX        =   2857
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
      Text            =   "10/01/2001"
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
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   3444
      Left            =   2370
      Top             =   2268
      Width           =   7452
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Last-"
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
      Left            =   4488
      TabIndex        =   11
      Top             =   3912
      Width           =   564
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "First-"
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
      Height          =   276
      Left            =   4488
      TabIndex        =   10
      Top             =   3264
      Width           =   612
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Select Invoices For Payment"
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
      Left            =   3678
      TabIndex        =   9
      Top             =   1032
      Width           =   4836
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000009&
      Height          =   852
      Left            =   2580
      Top             =   792
      Width           =   7020
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor Range:"
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
      Left            =   2976
      TabIndex        =   6
      Top             =   2712
      Width           =   1692
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice Due Date:"
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
      Left            =   2976
      TabIndex        =   5
      Top             =   4800
      Width           =   2124
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00D0D0D0&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   996
      Left            =   2592
      Top             =   672
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
Attribute VB_Name = "frmSelectOPays"
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
Dim TPayList As TPayListType
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer

Private Sub cmdAll_Click()
  fpcboVendFirst.ListIndex = 0
  fpcboVendLast.ListIndex = fpcboVendLast.ListCount - 1
  fpDueDate.Text = Format(Now, "mm/dd/yyyy")
End Sub

Private Sub cmdExit_Click()
  On Local Error Resume Next
  frmAPChkProcessMenu.Show
  Unload frmSelectOPays
End Sub

Private Sub cmdOk_Click()
On Local Error GoTo Cancel
10:
 ' MsgBox "is imprint visible", vbOKOnly
 If vaImprint1.Visible = False Then
 ' MsgBox "vaimprint.visible = false", vbOKOnly
  fpcboVendFirst.col = 0
  fpcboVendLast.col = 0
20:
 ' MsgBox "col ok now try see if vend sel ok", vbOKOnly
  If QPTrim$(fpcboVendFirst.ColText) <= QPTrim$(fpcboVendLast.ColText) Then
    LoadInvList
22:
    vaImprint1.Visible = True
23:
    lstInvoices.SetFocus
24:
    fpcboVendFirst.Visible = False
25:
    fpcboVendLast.Visible = False
26:
  Else
    MsgBox "First and Last Vendors must be Equal or in Correct Order.", vbOKOnly, "Invalid Vendor Selection"
  End If
27:
Else
 ' MsgBox "getinv ", vbOKOnly
  GetSelInvs
28:
  cmdExit_Click
End If
Cancel:
  If Err > 0 Then
    MsgBox "Error Code Was " + Err.Description + Str$(Err) + "(OK - Line:" & Erl & ")"
  End If
  Close
  Exit Sub

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      KillFile ("APCHK.opn")
      MainLog "Close AP"
      ClearInUse PWcnt
    End If
  End If
End Sub
Private Sub fpcboVendFirst_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboVendFirst.ListDown = True
  End If
End Sub

Private Sub fpcboVendFirst_LostFocus()
  fpcboVendFirst.Action = ActionClearSearchBuffer
End Sub

Private Sub fpcboVendLast_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboVendLast.ListDown = True
  End If
End Sub

Private Sub fpcboVendLast_LostFocus()
  fpcboVendLast.Action = ActionClearSearchBuffer
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Local Error GoTo Cancel
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
10:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyUp:
20:
      SendKeys "+{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
30:
      'cmdExit_Click
      SendKeys "%X"
      KeyCode = 0
    Case vbKeyF2:
40:
      'cmdAll_Click
      SendKeys "%A"
      KeyCode = 0
    Case vbKeyF10:
50:
      'cmdOk_Click
      SendKeys "%O"
      KeyCode = 0
    Case Else:
  End Select
Cancel:
  If Err > 0 Then
    MsgBox "Error Code Was " + Err.Description + Str$(Err) + "(key - Line:" & Erl & ")"
  End If
  Close
  Exit Sub

End Sub
Private Sub fpDueDate_LostFocus()
On Local Error Resume Next
  If CheckValDate(fpDueDate) = False Then
    MsgBox "Invalid Date, Please Correct.", vbOKOnly, "Invalid Date"
    fpDueDate.SetFocus
 ' Else
    'cmdOk.SetFocus
  End If
End Sub

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  StatusBar1.Panels.Item(1).Text = GLUserName
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Me.HelpContextID = hlpInvforPay
  VendCodeName2 fpcboVendFirst, fpcboVendLast
  fpcboVendFirst.ListIndex = 0
  fpcboVendLast.ListIndex = fpcboVendLast.ListCount - 1
  fpDueDate.Text = Format(Now, "mm/dd/yyyy")
End Sub
Private Sub cmdClear_Click()
  lstInvoices.Action = ActionDeselectAll
End Sub

Private Sub cmdMark_Click()
  lstInvoices.Action = ActionSelectAll
End Sub

Private Sub Form_Resize()
'  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
'  End If
End Sub
Private Sub LoadInvList()
'this fills the listbox with iNVOICE Entries for User to Select
  Dim cnt As Integer, FirstVendRecNum As Long, disdate As String
  Dim fmt As String, tempstr As String, strInfo As String
  Dim LdRecLen As Integer, VendorFile As Integer, NumVRecs As Integer
  Dim APLedgerFile As Integer, NumTran As Long, NextTrans As Long
  Dim LastVendRecNum As Long, DueDate As Integer, TempAmt As String
  Dim VendorIdxFile As Integer, NumActiveVendors As Integer
  Dim APLedgerRec(1) As APLedger81RecType
  Dim VendorIdx As VendorIdxRecType
  Dim Vendor As VendorRecType
 ' MsgBox "We got to load", vbOKOnly
  On Local Error GoTo Cancel
  FrmShowPctComp.Label1 = "Creating Invoice List"
  FrmShowPctComp.Show , Me
  DoEvents
10:
  DeActivateControls frmSelectOPays
11:
  LdRecLen = Len(APLedgerRec(1))
  fmt = "$##########.##"
12:
  DueDate = DateDiff("d", "12/31/1979", fpDueDate)
  fpcboVendFirst.col = 2
  fpcboVendLast.col = 2
15:
  FirstVendRecNum = fpcboVendFirst.ColText
  LastVendRecNum = fpcboVendLast.ColText
17:
  OpenVendorFile VendorFile, NumVRecs
18:
  OpenAPLedgerFile APLedgerFile, NumTran&, LdRecLen
19:
  OpenVendorIdx VendorIdxFile, NumActiveVendors
20:
  lstInvoices.Clear
21:
  For cnt = FirstVendRecNum To LastVendRecNum
        FrmShowPctComp.ShowPctComp cnt, LastVendRecNum
        If FrmShowPctComp.Out = True Then
          Close
          FrmShowPctComp.Out = False
          ActivateControls frmSelectOPays
          Unload FrmShowPctComp
          Exit For
        End If
28:
    Get VendorIdxFile, cnt, VendorIdx
    If VendorIdx.RecNum > 0 Then
    Get VendorFile, VendorIdx.RecNum, Vendor
    NextTrans& = Vendor.FrstTran
      Do Until NextTrans& = 0
      Get APLedgerFile, NextTrans&, APLedgerRec(1)
      If APLedgerRec(1).TRCode = 1 And APLedgerRec(1).PAYCODE = 1 Then
        If APLedgerRec(1).DueDate <= DueDate Then
          tempstr = Space$(15)
          TempAmt$ = Using$(fmt$, Str$(APLedgerRec(1).Amt))
          lstInvoices.AddItem Left$(Vendor.VNAME, 15) & Chr$(9) & QPTrim$(APLedgerRec(1).DOCNum) & Chr$(9) & Format(DateAdd("d", (APLedgerRec(1).TRDATE), "12-31-1979"), "mm/dd/yyyy") & Chr$(9) & Right$(TempAmt$, 20) & Chr$(9) & NextTrans & Chr$(9) & VendorIdx.RecNum
          lstInvoices.ListApplyTo = ListApplyToIndividual
          lstInvoices.col = 3
          lstInvoices.AlignH = AlignHRight
        End If
      End If
29:
    NextTrans& = APLedgerRec(1).NextTrans
      Loop
      End If
    Next
  Close
33:
  lstInvoices.Action = ActionSelectAll
34:
  ActivateControls frmSelectOPays
Cancel:
'  If Erl = 60 Or Erl = 62 Then
'    Resume Next
'  End If
  If Err > 0 Then
    MsgBox "Error Code Was " + Err.Description + Str$(Err) + "(LoadInv - Line:" & Erl & ")"
  End If
  Close
  Exit Sub

End Sub
  
Private Sub GetSelInvs()
  Dim TPayListFile As Integer, PayListRecLen As Integer
  Dim APLedgerFile As Integer, NumTran As Long, RecLen As Integer
'    MsgBox "we got here ", vbOKOnly
  
  On Local Error GoTo Cancel
10:
  KillFile "TPAYLIST.LST"
11:
  KillFile "TPayList2.lst"
12:
  KillFile "TPayListD.lst"
13:
  Dim Pcnt As Integer, cnt As Integer, x As Integer, y As Integer, s As String
  PayListRecLen = Len(TPayList)
  TPayListFile = FreeFile
14:
  Open "TPAYLIST.LST" For Random Shared As TPayListFile Len = PayListRecLen
15:
  OpenAPLedgerFile APLedgerFile, NumTran&, RecLen
  FrmShowPctComp.Label1 = "Creating Payment List"
  FrmShowPctComp.Show , Me
  DoEvents
17:
  DeActivateControls frmSelectOPays


'the following works if set multiselect property to simple
'but need to use the extended property !!! Figure out!!!
'  lstInvoices.ListIndex = 0
'    For Pcnt = 0 To lstInvoices.ListCount - 1
'      If lstInvoices.Selected(Pcnt) Then
'        cnt = cnt + 1
'        lstInvoices.col = 4
'        lstInvoices.ListIndex = Pcnt
'        TPayList.LedgerRecNum = QPTrim(lstInvoices.ColText)
'        lstInvoices.col = 5
'        TPayList.VendorRecNum = QPTrim(lstInvoices.ColText)
'        Put TPayListFile, cnt, TPayList
'
'      End If
'    Next
'got example from BOB at FarPoint used row and collist  !!Works!!
20:
    For Pcnt = 0 To lstInvoices.ListCount - 1
        FrmShowPctComp.ShowPctComp Pcnt, lstInvoices.ListCount - 1
        If FrmShowPctComp.Out = True Then
          Close
          FrmShowPctComp.Out = False
          ActivateControls frmSelectOPays
          Unload FrmShowPctComp
          Exit For
        End If
30:
      If lstInvoices.Selected(Pcnt) Then
        cnt = cnt + 1
        lstInvoices.col = 4
        lstInvoices.Row = Pcnt
        TPayList.LedgerRecNum = QPTrim(lstInvoices.ColList)
        lstInvoices.col = 5
        TPayList.VendorRecNum = QPTrim(lstInvoices.ColList)
40:
        Debug.Print lstInvoices.ColList
41:
        Put TPayListFile, cnt, TPayList
        
      End If
    Next

  Close
50:
  ActivateControls frmSelectOPays
Cancel:
'  If Erl = 60 Or Erl = 62 Then
'    Resume Next
'  End If
  If Err > 0 Then
    MsgBox "Error Code Was " + Err.Description + Str$(Err) + "(GetInv - Line:" & Erl & ")"
  End If
  Close
  Exit Sub

End Sub

Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub
