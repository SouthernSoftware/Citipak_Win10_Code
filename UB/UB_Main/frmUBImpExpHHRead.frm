VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmUBImpExpHHRead 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Handheld Import/Export Processing"
   ClientHeight    =   8865
   ClientLeft      =   3930
   ClientTop       =   2175
   ClientWidth     =   12210
   Icon            =   "frmUBImpExpHHRead.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   12210
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpMtrImpExpFlag 
      Height          =   375
      Left            =   5040
      TabIndex        =   0
      Top             =   2670
      Width           =   3345
      _Version        =   196608
      _ExtentX        =   5900
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
      ColDesigner     =   "frmUBImpExpHHRead.frx":08CA
   End
   Begin LpLib.fpCombo fpMtrType 
      Height          =   375
      Left            =   5340
      TabIndex        =   5
      Top             =   5040
      Visible         =   0   'False
      Width           =   3180
      _Version        =   196608
      _ExtentX        =   5609
      _ExtentY        =   661
      Text            =   "fpCombo1"
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
      Text            =   "fpCombo1"
      Columns         =   2
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
      ColumnEdit      =   0
      ColumnBound     =   -1
      Style           =   0
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
      ColDesigner     =   "frmUBImpExpHHRead.frx":0C30
   End
   Begin EditLib.fpLongInteger fpMtrInterNumb 
      Height          =   348
      Left            =   5064
      TabIndex        =   2
      Top             =   4032
      Visible         =   0   'False
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
   Begin MSMask.MaskEdBox fpReadDate 
      Height          =   348
      Left            =   7560
      TabIndex        =   3
      Top             =   4032
      Visible         =   0   'False
      Width           =   1932
      _ExtentX        =   3413
      _ExtentY        =   609
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/\2\0##"
      PromptChar      =   "_"
   End
   Begin EditLib.fpLongInteger fpMtrRoute 
      Height          =   345
      Left            =   5340
      TabIndex        =   6
      Top             =   5700
      Visible         =   0   'False
      Width           =   630
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
   Begin EditLib.fpText fpPathOut 
      Height          =   348
      Left            =   3408
      TabIndex        =   1
      Top             =   3528
      Width           =   6084
      _Version        =   196608
      _ExtentX        =   10731
      _ExtentY        =   614
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
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
      AlignTextV      =   1
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      AutoCase        =   0
      CaretInsert     =   0
      CaretOverWrite  =   3
      UserEntry       =   1
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
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   48
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   -1  'True
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
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3072
      Top             =   7752
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
      Left            =   9456
      TabIndex        =   9
      Top             =   7608
      Width           =   1332
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
      Left            =   7776
      TabIndex        =   8
      Top             =   7608
      Width           =   1332
   End
   Begin EditLib.fpText fptxtCycleSel 
      Height          =   330
      Left            =   5340
      TabIndex        =   7
      Top             =   5685
      Visible         =   0   'False
      Width           =   630
      _Version        =   196608
      _ExtentX        =   1122
      _ExtentY        =   572
      Enabled         =   -1  'True
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
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   12
      Top             =   8505
      Width           =   12210
      _ExtentX        =   21537
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7144
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7144
            TextSave        =   "7:27 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7144
            TextSave        =   "2/7/2020"
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
   Begin EditLib.fpText fptxtcycle 
      Height          =   945
      Left            =   5340
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   6270
      Visible         =   0   'False
      Width           =   4335
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
   Begin EditLib.fpLongInteger fpMtrCycle 
      Height          =   345
      Left            =   5340
      TabIndex        =   4
      Top             =   5160
      Visible         =   0   'False
      Width           =   630
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
      Left            =   3720
      TabIndex        =   23
      Top             =   5250
      Visible         =   0   'False
      Width           =   1470
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "30 Max."
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
      Height          =   276
      Index           =   2
      Left            =   3312
      TabIndex        =   22
      Top             =   6768
      Visible         =   0   'False
      Width           =   1908
   End
   Begin VB.Label lblHHInfo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Interrogator:"
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
      Index           =   4
      Left            =   3144
      TabIndex        =   20
      Top             =   4080
      Visible         =   0   'False
      Width           =   1644
   End
   Begin VB.Label lblHHInfo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Read Date:"
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
      Index           =   3
      Left            =   5928
      TabIndex        =   19
      Top             =   4080
      Visible         =   0   'False
      Width           =   1428
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Import/Export Meter Readings"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3534
      TabIndex        =   18
      Top             =   936
      Width           =   5148
   End
   Begin VB.Label lblHHInfo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Selected Meter Type:"
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
      Index           =   2
      Left            =   2376
      TabIndex        =   17
      Top             =   5088
      Visible         =   0   'False
      Width           =   2820
   End
   Begin VB.Label lblHHInfo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Default Path:"
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
      Left            =   2496
      TabIndex        =   16
      Top             =   3120
      Width           =   2436
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Height          =   2655
      Left            =   2250
      Top             =   4740
      Visible         =   0   'False
      Width           =   7620
   End
   Begin VB.Label lblHHInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Wait. . ."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   324
      Index           =   0
      Left            =   3072
      TabIndex        =   15
      Top             =   2160
      Width           =   6060
   End
   Begin VB.Label lblWhatHH 
      Caption         =   "0"
      Height          =   324
      Left            =   2736
      TabIndex        =   11
      Top             =   7752
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   228
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Selected Routes:"
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
      Height          =   276
      Index           =   0
      Left            =   3288
      TabIndex        =   14
      Top             =   6504
      Visible         =   0   'False
      Width           =   1908
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Height          =   2628
      Left            =   2256
      Top             =   1992
      Width           =   7620
   End
   Begin VB.Label lblHHop 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Handheld Operation:"
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
      Left            =   2688
      TabIndex        =   13
      Top             =   2688
      Width           =   2268
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3228
      Top             =   720
      Width           =   5772
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Route:"
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
      Left            =   3720
      TabIndex        =   10
      Top             =   5736
      Visible         =   0   'False
      Width           =   1476
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   972
      Left            =   3228
      Top             =   600
      Width           =   5772
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuPrnScn 
         Caption         =   "Prin&t Screen"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmUBImpExpHHRead"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Temp_Class As Resize_Class
Dim TRDate As Integer
Dim Over As clsTextBoxOverRider
Dim BegRoute As String, EndRoute As String, HHType As String
Dim UseCycle As Boolean, FirstInit As Boolean, Perquim As Boolean, Middle As Boolean
Dim CycleCnt As Integer, HHAction As Integer, applachflag As Boolean, SaludaCoF As Boolean
Dim Alberta As Boolean, CitySaluda As Boolean
Dim schlnglat As Boolean
Dim Cycle(1 To 30) As Integer

Dim UBCustRec(1) As NewUBCustRecType
Dim UBCustRecLen As Integer, UBSetupLen As Integer
Dim NumPC3000RdRecs As Integer, NumPC3000GetRdRecs As Integer

Dim UBSetUpRec(1) As UBSetupRecType
Dim HuskyErr As String
Dim ImpExpFlag As Boolean
Dim HuskyPort As String
Dim UBHHPath As UBHHPathRecType
Dim RCnt As Integer
Dim IdxFileSize As Long
Dim IdxRecLen As Integer
Dim IdxNumOfRecs As Long
Dim OkORNotFlag As Integer
Dim MsgText(0 To 5) As String
Dim InterrNum As Integer         'sensus
Dim HHPathInOut As String
Dim SensusIOFile As String
Dim PathData As String
Dim ttt$

'Badger B HH extra export file is comma separated file
' Added 01-2014  called tmc file, use that as extention with route as name

Private Sub Form_Load()
  
  Dim HHType As String
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
  Me.HelpContextID = hlpImportExport
  fpMtrImpExpFlag.InsertRow = "Import Meter Readings."
  fpMtrImpExpFlag.InsertRow = "Export Meter Readings."
  fpMtrImpExpFlag.ListIndex = 0
  
  fpMtrType.InsertRow = "All Meters"
  fpMtrType.InsertRow = "Water/Sewer"
  fpMtrType.InsertRow = "Electric"
  fpMtrType.InsertRow = "Gas Meters"
  fpMtrType.ListIndex = 0
  If InStr(TOWNNAME$, "APPALACHIA") Or InStr(TOWNNAME$, "SHENANDOAH") Or InStr(TOWNNAME$, "ASHLEY") Then
    applachflag = True
  Else
    applachflag = False
  End If
  If InStr(TOWNNAME$, "GILBERT SUMMIT") Then
    schlnglat = True
  Else
    schlnglat = False
  End If
  If InStr(TOWNNAME$, "MIDDLETOWN") _
  Or InStr(TOWNNAME$, "BISCOE") _
  Or InStr(TOWNNAME$, "SUMMERTON") _
  Or InStr(TOWNNAME$, "GILBERT SUMMIT") _
  Or InStr(TOWNNAME$, "DOBSON") _
  Or InStr(TOWNNAME$, "PLYMOUTH") _
  Or InStr(TOWNNAME$, "TIMMONSVILLE") _
  Or InStr(TOWNNAME$, "BEECH") _
  Or InStr(TOWNNAME$, "SEVEN DEVILS") _
  Or InStr(TOWNNAME$, "CLARKTON") _
  Or InStr(TOWNNAME$, "BETHEL") _
  Or InStr(TOWNNAME$, "OAKBORO") _
  Or InStr(TOWNNAME$, "LINCOLN") _
  Or InStr(TOWNNAME$, "STANLEY") _
  Or InStr(TOWNNAME$, "NORLINA") _
  Or InStr(TOWNNAME$, "REMINGTON") _
  Or InStr(TOWNNAME$, "TROY") _
  Or InStr(TOWNNAME$, "SPRING") _
  Then
    Middle = True
  Else
    Middle = False
  End If
  If InStr(TOWNNAME$, "SALUDA CO") Then
    SaludaCoF = True
  Else
    SaludaCoF = False
  End If
  If InStr(TOWNNAME$, "PERQUIMANS") Or InStr(TOWNNAME$, "WADESBORO") _
  Or InStr(TOWNNAME$, "GILEAD") Or InStr(TOWNNAME$, "NORTHWEST CL") Then
    Perquim = True
  Else
    Perquim = False
  End If
  If InStr(TOWNNAME$, "ALBERTA") Then
    Perquim = True
    Alberta = True
  Else
    Alberta = False
  End If
  '
  If InStr(TOWNNAME$, "CITY OF SALUDA") Then
    CitySaluda = True
  Else
    CitySaluda = False
  End If
  
  frmUBImpExpHHRead.Timer1.Enabled = True
  
  Erase Cycle
  
  Call GetHandHeldPathWay
  DoEvents

End Sub

Private Sub fpMtrImpExpFlag_Change()
  Dim WhatAction As Integer
  If Not FirstInit Then
    FirstInit = True
  Else
    WhatAction = fpMtrImpExpFlag.ListIndex
    Select Case WhatAction
    Case 0
      ImpExpFlag = False
    Case 1
      ImpExpFlag = True
    Case Else
    End Select
  End If
End Sub

Private Sub fpMtrImpExpFlag_LostFocus()
  SetupImportExportScrn ImpExpFlag
End Sub

Private Sub fpMtrImpExpFlag_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpMtrImpExpFlag.ListDown = True
  End If
  If fpMtrImpExpFlag.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      KeyCode = 0
      If Me.fpPathOut.Visible Then
        Me.fpPathOut.SetFocus
      Else
'        SetupImportExportScrn ImpExpFlag
'        If ImpExpFlag = True Then
'          fpMtrType.SetFocus
'        Else
'          cmdOk.SetFocus
'        End If
      End If
    Else
      If KeyCode = vbKeyUp Then
        KeyCode = 0
      End If
    End If
  End If
'this traps the ImpExp field keyboard events.
End Sub


Private Sub fpMtrType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpMtrType.ListDown = True
  End If
  If fpMtrType.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      KeyCode = 0
      Me.fptxtCycleSel.SetFocus
    Else
      If KeyCode = vbKeyUp Then
        If Me.fpPathOut.Visible Then Me.fpPathOut.SetFocus
        KeyCode = 0
      End If
    End If
  End If
'this traps the ImpExp field key board events.
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
      Call cmdOk_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_Resize()
  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
  End If
End Sub

Private Sub SetupImportExportScrn(ByVal ImpExpFlag As Boolean)
  If ImpExpFlag Then     'EXPORTING METER READINGS
    Select Case HHType
    Case "H", "W", "D", "Z", "J"  'Husky, Intermec, data genie, Jett
      Me.Shape3.Visible = True
      Me.lblHHInfo(2).Caption = "Selected Meter Type:"
      Me.lblHHInfo(2).Visible = True
      Me.fptxtCycleSel.Visible = True
      Me.Label3(0).Visible = True
      Me.Label3(1).Visible = True
      Me.fptxtcycle.Visible = True
      Me.fpMtrType.Visible = True
    
    Case "S", "E", "U"         'Sensus, ESensus , old esensus
      Me.Shape3.Visible = True
      Me.lblHHInfo(4).Caption = "Interrogator:"
      Me.lblHHInfo(4).Visible = True
      Me.lblHHInfo(2).Visible = True
      Me.fpMtrInterNumb.Visible = True
      Me.Label3(0).Visible = True
      Me.Label3(1).Visible = True
      Me.fptxtCycleSel.Visible = True
      Me.fptxtcycle.Visible = True
      Me.fpMtrType.Visible = True
    
    Case "C" 'Syscon
      Me.Shape3.Visible = True
      Me.lblHHInfo(2).Caption = "Route to Process:"
      Me.lblHHInfo(2).Visible = True
      Me.fpMtrInterNumb.Visible = True
    
    Case "T" 'Badger Beacon
    
    Case "L"
      Me.Shape3.Visible = True
      Me.Label3(0).Visible = True
      Me.Label3(1).Visible = True
      Me.fptxtCycleSel.Visible = True
      Me.fptxtcycle.Visible = True
    
  '  Case "I"  'Itron

    Case "P"  'Performance Meter
      Me.Shape3.Visible = True
      Me.Label3(1).Visible = True
      'Me.fpMtrRoute.Visible = True
      Me.fpMtrCycle.Visible = True
      Me.Label3(0).Visible = True
      Me.Label3(1).Visible = True
      Me.Label3(3).Visible = True
      Me.fptxtCycleSel.Visible = True
      Me.fptxtcycle.Visible = True
  '
    Case "B"
      Me.Shape3.Visible = True
      Me.Label3(1).Visible = True
      Me.Label3(3).Visible = True
      Me.fpMtrRoute.Visible = True
      Me.fpMtrCycle.Visible = True
      Me.lblHHInfo(3).Visible = True
      fpReadDate.Visible = True
    Case "X", "G", "Y", "A"
      Me.Shape3.Visible = True
      Me.lblHHInfo(2).Caption = "Route to Process:"
      Me.fptxtCycleSel.Visible = True
      Me.Label3(0).Visible = True
      Me.Label3(1).Visible = True
  
      Me.fptxtcycle.Visible = True
    Case Else 'No handheld device.
    End Select
  Else  'IMPORTING METER READINGS
    Call MtrReadExportOFF  ' Turns off the export part.
  End If
End Sub

Private Sub fpPathOut_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    KeyCode = 0
    Sendkeys "{DOWN}"
  End If
End Sub

Private Sub fpReadDate_Validate(Cancel As Boolean)
  TRDate = Date2Num(fpReadDate.Text)
  If TRDate < 0 Then
    Me.fpReadDate.Text = "__/__/20__"
  End If
End Sub


Private Sub fptxtCycleSel_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim cnt As Integer
  If KeyCode = vbKeyReturn Then
    If Len(fptxtCycleSel.Text) <> 0 Then
      getcyclelist
    Else
      cmdOk.SetFocus
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
    cmdOk.SetFocus
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

Private Sub mnuExit_Click()
  Call cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
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

Private Sub Timer1_Timer()
  Timer1.Enabled = False
  
  HHType = lblWhatHH.Caption
  Select Case HHType
  Case "H"   'Husky
    Me.lblHHInfo(0).Caption = "Husky Meter Reading Device"
    Me.lblHHInfo(1).Caption = "Path to Husky:"
  
  Case "E", "S", "U"
    Select Case HHType
    Case "E"
      Me.lblHHInfo(0).Caption = "ESensus Meter Reading Device"
    Case "S"
      Me.lblHHInfo(0).Caption = "Sensus Meter Reading Device"
    Case "U"
      Me.lblHHInfo(0).Caption = "OESensus Meter Reading Device"
    End Select
    
    Me.lblHHInfo(1).Caption = "Path to Sensus:"
    Me.lblHHInfo(3).Visible = True
    Me.lblHHInfo(4).Visible = True
    Me.fpMtrInterNumb.Visible = True
    Me.fpReadDate.Visible = True
  
  Case "C" 'Syscon
    Me.lblHHInfo(0).Caption = "Syscon Meter Reading Device"
  
  Case "J"  'Jett
    Me.lblHHInfo(0).Caption = "Jett Meter Reading Device."
    Me.fpReadDate.Visible = False
    Me.lblHHInfo(3).Visible = False
    Me.lblHHInfo(1).Visible = False
    Me.fpPathOut.Visible = False
    
  Case "D"  'DataGenie
    Me.fpReadDate.Visible = True
    Me.lblHHInfo(3).Visible = True
  
    Me.lblHHInfo(0).Caption = "Data Genie Meter Reading Device"
    Me.lblHHInfo(1).Caption = "Path to Genie:"
  
  Case "A"
    Me.lblHHInfo(0).Caption = "Badger Beacon V2 Meter Reading Device"
    Me.lblHHInfo(1).Caption = "Path to Beacon:"
  
  Case "T"
    Me.lblHHInfo(0).Caption = "Badger Beacon Meter Reading Device"
    Me.lblHHInfo(1).Caption = "Path to Beacon:"
  
  Case "L"
    Me.lblHHInfo(0).Caption = "Logicon Meter Reading Device"
    Me.lblHHInfo(1).Caption = "Path to Logicon:"
    Me.lblHHInfo(4).Caption = "Route ID:"
    Me.lblHHInfo(3).Visible = True
    Me.lblHHInfo(4).Visible = True
    Me.fpMtrInterNumb.Visible = True
    Me.fpReadDate.Visible = True
      
  Case "I"
    Me.lblHHInfo(0).Caption = "Itron Meter Reading Device"
    Me.lblHHInfo(1).Caption = "Path to Itron:"
    
  Case "Z"
    Me.lblHHInfo(0).Caption = "Schulmberger Meter Reading Device"
    Me.lblHHInfo(1).Caption = "Path to Schulmberger:"
  Case "B"
    Me.lblHHInfo(0).Caption = "Badger Meter Reading Device"
    Me.lblHHInfo(1).Caption = "Path to Badger:"
  Case "G"
    Me.lblHHInfo(0).Caption = "BadgerV2 Meter Reading Device"
    Me.lblHHInfo(1).Caption = "Path to Badger:"
  Case "Y"
    Me.lblHHInfo(0).Caption = "BadgerV3 Meter Reading Device"
    Me.lblHHInfo(1).Caption = "Path to Badger:"
  Case "W"
    Me.lblHHInfo(0).Caption = "Intermec CE Meter Reading Device"
    Me.lblHHInfo(1).Caption = "Path to Intermec:"
  Case "P"
    Me.lblHHInfo(0).Caption = "Performance Meter Reading Device"
'    Stop
  Case Else 'No handheld device.
  End Select
  
End Sub

Private Sub MtrReadExportOFF()
  
  Me.Shape3.Visible = False
  'Me.lblHHInfo(2).Caption = "Selected Meter Type:"
  Me.lblHHInfo(2).Visible = False
  Me.fptxtCycleSel.Visible = False
  Me.Label3(0).Visible = False
  Me.Label3(1).Visible = False
  Me.Label3(3).Visible = False
  fpMtrCycle.Visible = False
  Select Case HHType
  
  Case "E", "S", "L", "U"
  'If HHType = "E" Or HHType = "S" Or HHType = "L" Then
  Case "B", "J", "P"
    Me.lblHHInfo(3).Visible = False
    fpReadDate.Visible = False
  Case "X", "G", "Y", "A"
    Me.lblHHInfo(3).Visible = True
    fpReadDate.Visible = True
  Case Else
    Me.fpMtrInterNumb.Visible = False
  End Select
  Me.fptxtcycle.Visible = False
  Me.fpMtrType.Visible = False
  Me.fpMtrRoute.Visible = False
  DoEvents
  If HHType = "J" Then
    If fpMtrType.Visible Then fpMtrType.SetFocus
  End If

End Sub

Private Sub cmdExit_Click()
  FirstInit = False
  CycleCnt = 0
  
  Load frmUBHHMenu
  DoEvents
  frmUBHHMenu.Show
  Unload frmUBImpExpHHRead
End Sub

Private Function CheckInterrNum%()
  
  Dim InterrFlag As Boolean
  InterrFlag = False
  
  InterrNum = Me.fpMtrInterNumb.Value
  
  If InterrNum > 0 Then
    InterrFlag = True
  Else
    frmMsgDialog.RetLabel = "-2"
    frmMsgDialog.Caption = "ERROR:"
    For RCnt = 0 To 4
      frmMsgDialog.Label(RCnt).Caption = ""
      frmMsgDialog.Label(RCnt).FontSize = frmMsgDialog.Label(RCnt).FontSize + 2
    Next
    If lblWhatHH.Caption = "L" Then
      frmMsgDialog.Label(1).Caption = "Invalid Route ID Number."
    Else
      frmMsgDialog.Label(1).Caption = "Invalid Interrogator Number."
    End If
    frmMsgDialog.Label(2).Caption = "Please call Southern Software for"
    frmMsgDialog.Label(3).Caption = "additional Information."
    frmMsgDialog.Show vbModal
    Unload frmMsgDialog
  End If

  CheckInterrNum% = InterrFlag
End Function

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
    If HHType <> "J" Then
      frmMsgDialog.RetLabel = "-2"
      frmMsgDialog.Caption = "ERROR:"
      For RCnt = 0 To 4
        frmMsgDialog.Label(RCnt).Caption = ""
        frmMsgDialog.Label(RCnt).FontSize = frmMsgDialog.Label(RCnt).FontSize + 2
      Next
      frmMsgDialog.Label(1).Caption = "NO CYCLES ENTERED TO EXPORT."
      frmMsgDialog.Label(2).Caption = "Please call Southern Software for"
      frmMsgDialog.Label(3).Caption = "additional Information."
      frmMsgDialog.Show vbModal
      Unload frmMsgDialog
    End If
    GoTo CheckCyclesExit
  End If

CheckCyclesExit:
  
  CheckCycles% = CyclesOK
End Function

Private Sub cmdOk_Click()
  Dim UBHHPathRecLen As Integer, UBPathWayF As Integer
  UBHHPathRecLen = Len(UBHHPath)
  Call fpMtrImpExpFlag_Change

  DoEvents
  If HHType <> "J" Then
    HHPathInOut = QPTrim$(UBHHPath.PathWay)
  
    If HHPathInOut <> QPTrim$(fpPathOut.Text) Then
      HHPathInOut = QPTrim$(fpPathOut.Text)
      If Len(HHPathInOut) > 2 Then  'this allows entry as drive letter only 'C:'
        If Right$(HHPathInOut, 1) <> "\" Then
          HHPathInOut = HHPathInOut + "\"
        End If
      End If
      UBHHPath.PathWay = HHPathInOut
    End If

    If ChkHHPathWay(HHPathInOut) Then
      UBPathWayF = FreeFile
      Open UBPath$ + UBHHPathWayFile For Random Shared As UBPathWayF Len = UBHHPathRecLen
      Put UBPathWayF, 1, UBHHPath
      Close UBPathWayF
    Else
      MsgText(0) = "ERROR:"
      MsgText(1) = "INVALID PATH TO HANDHELD"
      MsgText(2) = ""
      MsgText(3) = "Please call Southern Software for"
      MsgText(4) = "additional Information."
      MsgText(5) = ""
      GetOKorNot% MsgText(), True, False
      GoTo cmdOkErrorExit
    End If
'Stop
  End If
  Select Case HHType
    Case "W"
      If ImpExpFlag Then     'EXPORTING METER READINGS
        If Not CheckCycles% Then  'cycles err
          GoTo cmdOkErrorExit
        End If
'      Else  'Importing readings
      End If
      DeActivateControls Me
      ImpExpIMecHHInfo ImpExpFlag
      ActivateControls Me
    Case "H"   'Husky
      If ImpExpFlag Then     'EXPORTING METER READINGS
        If Not LoadHuskyCFGFile Then
          frmMsgDialog.RetLabel = "-2"
          frmMsgDialog.Caption = "ERROR"
          For RCnt = 0 To 4
            frmMsgDialog.Label(RCnt).Caption = ""
            frmMsgDialog.Label(RCnt).FontSize = frmMsgDialog.Label(RCnt).FontSize + 2
          Next
          frmMsgDialog.Label(1).Caption = "CAN NOT FIND THE FILE 'UBHUSKY.CFG'"
          frmMsgDialog.Label(2).Caption = "Please call Southern Software for"
          frmMsgDialog.Label(3).Caption = "additional Information."
          frmMsgDialog.Show vbModal
          GoTo cmdOkErrorExit
        End If
        If Not CheckCycles% Then  'cycles err
          GoTo cmdOkErrorExit
        End If
      End If
      DeActivateControls Me
      ImpExpHuskyHHInfo ImpExpFlag
      ActivateControls Me
  
  Case "E", "S", "U"       'ESensus
      If Not CheckInterrNum% Then
        GoTo cmdOkErrorExit
      End If
      If ImpExpFlag Then          'if exporting
        If Not CheckCycles% Then  'cycles err
          GoTo cmdOkErrorExit
        End If
      End If
      
      DeActivateControls Me
      If HHType = "E" Then
        ImpExpESenHHInfo ImpExpFlag
      ElseIf HHType = "S" Then
        ImpExpOSenHHInfo ImpExpFlag
      Else
        ImpExpUSenHHInfo ImpExpFlag
      End If
      
      ActivateControls Me
      
  Case "P" 'Performance
      If ImpExpFlag Then          'if exporting
'        If Not CheckCycles% Then  'cycles err
'          GoTo cmdOkErrorExit
'        End If
      End If
      
      DeActivateControls Me
      ImpExpPMHHInfo ImpExpFlag
      ActivateControls Me

  Case "C" 'Syscon
  
  Case "D"  'DataGenie
    If ImpExpFlag Then     'EXPORTING METER READINGS
'      If Not CheckCycles% Then  'cycles err
'        GoTo cmdOkErrorExit
'      End If
'      Else  'Importing readings
    End If
    DeActivateControls Me
    ImpExpGenieHHInfo ImpExpFlag
    ActivateControls Me
 
'dale dale
  
  Case "L"
    If Not CheckInterrNum% Then 'route id check for logicon
      GoTo cmdOkErrorExit
    End If
    If ImpExpFlag Then     'EXPORTING METER READINGS
      If Not CheckCycles% Then  'cycles err
        GoTo cmdOkErrorExit
      End If
    End If
    DeActivateControls Me
    ImpExpLogiconHHInfo ImpExpFlag
    ActivateControls Me
  
  Case "I"
  Case "Z"
    If ImpExpFlag Then     'EXPORTING METER READINGS
      If Not CheckCycles% Then  'cycles err
        GoTo cmdOkErrorExit
      End If
    End If
    DeActivateControls Me
    ImpExpSchulmHHInfo ImpExpFlag
    ActivateControls Me
  Case "X"
    If ImpExpFlag Then     'EXPORTING METER READINGS
      If Not CheckCycles% Then  'cycles err
        GoTo cmdOkErrorExit
      End If
    End If
    DeActivateControls Me
    ImpExpBadgerHHInfo ImpExpFlag
    ActivateControls Me
  Case "Y"
     If ImpExpFlag Then     'EXPORTING METER READINGS
      If Not CheckCycles% Then  'cycles err
        GoTo cmdOkErrorExit
      End If
    End If
    DeActivateControls Me
    ImpExpBadgerV3HHInfo ImpExpFlag
    ActivateControls Me
  Case "G"
    If ImpExpFlag Then     'EXPORTING METER READINGS
      If Not CheckCycles% Then  'cycles err
        GoTo cmdOkErrorExit
      End If
    End If
    DeActivateControls Me
    ImpExpBadgerV2HHInfo ImpExpFlag
    ActivateControls Me
  Case "B"
    If ImpExpFlag Then     'EXPORTING METER READINGS
'      If Not CheckCycles% Then  'cycles err
'        GoTo cmdOkErrorExit
'      End If
    End If
    DeActivateControls Me
    If applachflag Then
      ImpExpRadioReadHHInfo ImpExpFlag
    Else
      ImpExpRadioReadHHInfo2 ImpExpFlag
      
    End If
    ActivateControls Me
  Case "J"
    If ImpExpFlag Then     'EXPORTING METER READINGS
      If Not CheckCycles% Then  'cycles err
        GoTo cmdOkErrorExit
      End If
    End If
    DeActivateControls Me
    ImpExpJettHHInfo ImpExpFlag
    ActivateControls Me
  
  Case "A"  'Beacon V2
    If ImpExpFlag Then     'EXPORTING METER READINGS
      If Not CheckCycles% Then  'cycles err
        GoTo cmdOkErrorExit
      End If
    End If
    DeActivateControls Me
    ImpExpBeaconV2HHInfo ImpExpFlag
    ActivateControls Me
  
  Case "T"   'Badger beacon
    DeActivateControls Me
    ImpExpBeaconHHInfo ImpExpFlag
    ActivateControls Me
  
  Case Else 'No handheld device.
  
  End Select

cmdOkErrorExit:
End Sub

Private Sub GetBeaCurReadAndDate(ByRef BeaconTXT As String, ByRef CurrRead As Long, ByRef ReadDate As Integer)
    Dim wrkRead As Long, wrkDate As Integer
    Dim bPos1 As Integer, bPos2 As Integer
    Dim xcnt As Integer, BTLen As Integer
    Dim Vbar As String ', BeaText As String
    Dim tWrk As String
    
    xcnt = 0
    wrkRead = 0
    ReadDate = 0
    Vbar = "|"
    
    'If InStr(BeaconTXT, "311-1") > 0 Then Stop

    BTLen = Len(BeaconTXT)
    If BTLen <= 1 Then
      wrkRead = 0
    Else
      bPos1 = InStr(BeaconTXT, Vbar)
      xcnt = xcnt + 1
      
      Do While bPos1 > 0
        If (bPos1 + 1) >= BTLen Then
          wrkRead = 0
          Exit Do
        End If
        
        bPos1 = InStr(bPos1 + 1, BeaconTXT, Vbar)
        xcnt = xcnt + 1
        If xcnt = 103 Then
          bPos2 = InStr(bPos1 + 1, BeaconTXT, Vbar)
          tWrk$ = QPTrim(Mid$(BeaconTXT, bPos1 + 1, bPos2 - bPos1))
          wrkRead = Val(Left$(tWrk$, Len(tWrk$) - 1))
          CurrRead = wrkRead
          bPos1 = bPos2
          'point to field 104 read date
          bPos2 = InStr(bPos1 + 1, BeaconTXT, Vbar)
          If bPos2 - bPos1 >= 10 Then
            tWrk$ = Left$(QPTrim(Mid$(BeaconTXT, bPos1 + 1, bPos2 - bPos1)), 10)
            'tWrk$ = Mid$(tWrk$, 6) + "-" + Left$(tWrk$, 4)
            ReadDate = Date2Num(tWrk$)
          Else
            ReadDate = -1
          End If
        End If
      Loop
    End If

'    If MeterRecord < 1 Then
'      MeterRecord = 1
'    End If
'    BCCustRec = CustRec

End Sub


Private Sub GetBeaCustAndMtrID(ByRef BeaconTXT As String, ByRef BCCustRec As Long, ByRef MeterRecord As Integer)
    Dim CustRec As Long
    Dim bPos1 As Integer, bPos2 As Integer
    Dim xcnt As Integer, BTLen As Integer
    Dim Vbar As String ', BeaText As String
    Dim tWrk As String
    
    xcnt = 0
    CustRec = 0
    Vbar = "|"

    BTLen = Len(BeaconTXT)
    If BTLen <= 1 Then
      CustRec = 0
    Else
      bPos1 = InStr(BeaconTXT, Vbar)
      xcnt = xcnt + 1
      
      Do While bPos1 > 0
        If (bPos1 + 1) >= BTLen Then
          CustRec = 0
          Exit Do
        End If
        
        bPos1 = InStr(bPos1 + 1, BeaconTXT, Vbar)
        xcnt = xcnt + 1
        If xcnt = 97 Then
          bPos2 = InStr(bPos1 + 1, BeaconTXT, Vbar)
          tWrk$ = QPTrim(Mid$(BeaconTXT, bPos1 + 1, (bPos2 - bPos1) - 1))
          CustRec = Val(Left$(tWrk$, Len(tWrk$) - 2))
          MeterRecord = Val(Right$(tWrk$, 1))
        '  Stop
          Exit Do
        End If
      Loop
     
    End If
    If MeterRecord < 1 Then
      MeterRecord = 1
    End If
    BCCustRec = CustRec

'GetBeaconCustRec = CustRec

End Sub

'ImpExpBeaconHHInfo ImpExpFlag

Private Function ChkHHPathWay%(THHPathWay As String)
  On Local Error Resume Next
  
  Dim UBPathWayF As Integer
  UBPathWayF = FreeFile
  Open QPTrim$(THHPathWay) + "chkhhpth.tmp" For Random As UBPathWayF Len = 2
  
  If Err Then
    ChkHHPathWay% = False
  Else
    Close
    KillFile QPTrim$(THHPathWay) + "chkhhpth.tmp"
    ChkHHPathWay% = True
  End If
  On Local Error GoTo 0
End Function

Private Function GetHandHeldPathWay%()
  On Local Error Resume Next
  Dim UBHHPathRecLen As Integer, UBPathWayF As Integer
  UBHHPathRecLen = Len(UBHHPath)
  If Exist(UBPath$ + UBHHPathWayFile) Then
    UBPathWayF = FreeFile
    Open UBPath$ + UBHHPathWayFile For Random Shared As UBPathWayF Len = UBHHPathRecLen
    If LOF(UBPathWayF) > 0 Then
      Get UBPathWayF, 1, UBHHPath
      If Len(QPTrim$(UBHHPath.PathWay)) = 0 Then
        UBHHPath.PathWay = UBPath$
        Put UBPathWayF, 1, UBHHPath
      End If
    Else
      UBHHPath.PathWay = UBPath$
      Put UBPathWayF, 1, UBHHPath
    End If
  Else
    UBPathWayF = FreeFile
    Open UBPath$ + UBHHPathWayFile For Random Shared As UBPathWayF Len = UBHHPathRecLen
    UBHHPath.PathWay = UBPath$
    Put UBPathWayF, 1, UBHHPath
  End If

ExitGetHHPathway:
  fpPathOut.Text = QPTrim$(UBHHPath.PathWay)
  On Error GoTo 0
  
End Function

Private Sub FGetAH(FileName As String, IdxBuff() As UBCustIndexRecType, ByVal IdxRecLen As Integer, ByVal IdxNumOfRecs As Long)
  Dim ICnt As Long
  Dim IdxFile As Integer
  IdxFile = FreeFile
  Open FileName For Random Shared As IdxFile Len = IdxRecLen
  For ICnt = 1 To IdxNumOfRecs
    Get IdxFile, ICnt, IdxBuff(ICnt).RecNum
  Next
  Close IdxFile
End Sub

Private Function LoadHuskyCFGFile%()
  Dim CFGFile As Integer
  If Exist(UBPath$ + "UBHUSKY.CFG") Then
    CFGFile = FreeFile
    Open UBPath$ + "UBHUSKY.CFG" For Input As #CFGFile
    Line Input #CFGFile, HuskyPort
    Close CFGFile
    LoadHuskyCFGFile% = True
  Else
    LoadHuskyCFGFile% = False
  End If
End Function

Private Sub ImpExpGenieHHInfo(ByVal ImpExpFlag As Boolean)
  
  ReDim UBDGRec(1) As UBDGRecType
  Dim UBDGRecLen As Integer, UBDGFile As Integer
  Dim UBGenieIOFile As Integer, NumGenieRecs As Integer
  Dim GenieIOFile As String
  
  ReDim UBDGRdRec(1) As UBDGHHRecType
  
  Dim HighVar As Integer, LowVar As Integer
  Dim WhatTypes As String, CustAcc As String
  Dim UBFile As Integer, UBSenIOFile As Integer
  Dim UBSenRdRecLen As Integer, NumSenRdRecs As Integer
  Dim BookCnt As Integer, MtrCnt As Integer
  Dim RMCnt As Long, WhatRMRec As Long
  Dim Account As String
  Dim Average As Double, LowRead As Double
  Dim MeterID As String, MRDate As String
  Dim MeterOK As Boolean, MTRType As String, MeterType As String
  Dim HighRead As Double, ILowRead As Double
  Dim UBDGRdRecLen As Integer, NumDGRdRecs As Integer
  Dim UBDGRdFile As Integer, NumberofRoutes As Integer
  Dim cnt As Long, Prec As Long
  Dim MeterReadDate As String
  Dim DashPos As Integer, MeterRecord As Integer
  Dim CurReading As Double
  Dim TimesRead As Long
  
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen
  
  HighVar = UBSetUpRec(1).HighRead
  LowVar = UBSetUpRec(1).LowRead
  UBCustRecLen = Len(UBCustRec(1))
  
  If ImpExpFlag Then     'EXPORTING METER READINGS
    WhatTypes$ = Left$(Me.fpMtrType.Text, 1)
    GoSub SendGenie
    Call cmdExit_Click
  Else
    GoSub GetGenie
    Call cmdExit_Click
  End If
  Exit Sub

GetGenie:
  
  GenieIOFile = HHPathInOut + "UBCUSTTR.DAT"
  
  MRDate$ = QPTrim$(fpReadDate.Text)
  
  MsgText(0) = ""
  MsgText(1) = "Import Data Genie Reading File."
  MsgText(2) = ""
  MsgText(3) = ""
  MsgText(4) = "Ready to Proceed?"
  MsgText(5) = ""
  
  Select Case GetOKorNot%(MsgText(), False, True, 1)
  Case False
    GoTo ErrorGenieExit:
  End Select

'Open and Initialize the DG Genie Read Information File
  ReDim UBDGGetRDRec(1) As UBDGHHRecType
  'UBDGGetRDRec (1)
  UBDGRdRecLen = Len(UBDGGetRDRec(1))
  
  UBGenieIOFile = FreeFile
  Open GenieIOFile For Random Shared As UBGenieIOFile Len = UBDGRdRecLen
  
  NumGenieRecs = LOF(UBGenieIOFile) / UBDGRdRecLen
  'Open and Initialize the DG Genie Read Information File
  If NumGenieRecs = 0 Then
    Close
    MsgText(0) = "ERROR:"
    MsgText(1) = "IMPORT FILE NOT FOUND"
    MsgText(2) = "Make sure that UBCUSTTR.DAT"
    MsgText(3) = "is in the directory!"
    MsgText(4) = "Please call Southern Software for"
    MsgText(5) = "additional Information."
    GetOKorNot% MsgText(), True, False
    GoTo ErrorGenieExit
  End If
  
  FrmShowPctComp.Label1 = "Importing Meter Reading Information."
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show '1, Parent
  
  UBFile = FreeFile
  Open UBPath + "UBCUST.DAT" For Random Shared As UBFile Len = UBCustRecLen
  
  cnt& = 1      ' Initialize File Counter to 1
  Do
    Get UBGenieIOFile, cnt&, UBDGGetRDRec(1)
    Prec& = Val(QPTrim$(UBDGGetRDRec(1).Account))
    If Prec& > 0 Then
      Get UBFile, Prec&, UBCustRec(1)
      GoSub DGExtractRecord
    End If
    cnt& = cnt& + 1
    FrmShowPctComp.ShowPctComp cnt, NumGenieRecs
  Loop Until cnt& > NumGenieRecs
  Close
  
  'Done show import complete
  MsgText(0) = "Data Genie Operation"
  MsgText(1) = "Import Complete."
  MsgText(2) = ""
  MsgText(3) = ""
  MsgText(4) = ""
  MsgText(5) = ""
  GetOKorNot% MsgText(), True, True

ErrorGenieExit:
Return

DGExtractRecord:
  If Len(QPTrim$(UBDGGetRDRec(1).NewRdng)) > 0 Then
    MeterRecord = Val(Right$((QPTrim$(UBDGGetRDRec(1).Account)), 1))
    ' Check Meter Updated Flag
    ' Update Meter W/Reading
    CurReading# = Val(UBDGGetRDRec(1).NewRdng)

    MeterReadDate = Left$(UBDGGetRDRec(1).Date, 2) + "/" + Mid$(UBDGGetRDRec(1).Date, 3, 2) + "/20" + Right$(UBDGGetRDRec(1).Date, 2)

    If UBCustRec(1).LocMeters(MeterRecord).ReadFlag = "Y" Then
      UBCustRec(1).LocMeters(MeterRecord).CurRead = CurReading#
      UBCustRec(1).LocMeters(MeterRecord).CurDate = Date2Num%(MeterReadDate)
    Else
      UBCustRec(1).LocMeters(MeterRecord).PrevRead = UBCustRec(1).LocMeters(MeterRecord).CurRead
      UBCustRec(1).LocMeters(MeterRecord).PastDate = UBCustRec(1).LocMeters(MeterRecord).CurDate
      UBCustRec(1).LocMeters(MeterRecord).ReadFlag = "Y"
      UBCustRec(1).LocMeters(MeterRecord).CurDate = Date2Num%(MeterReadDate)
      UBCustRec(1).LocMeters(MeterRecord).CurRead = CurReading#
    End If
    Put UBFile, Prec&, UBCustRec(1)
  End If
Return

Return

SendGenie:
  'GoSub OpenCustFile      'Open Customer Data File
  
  FrmShowPctComp.Label1 = "Exporting Meter Reading Information."
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show '1, Parent
  
  UBFile = FreeFile
  Open UBPath + "UBCUST.DAT" For Random Shared As UBFile Len = UBCustRecLen

  'Open and Initialize the DG Genie Read Information File
  ReDim UBDGRdRec(1) As UBDGHHRecType
  UBDGRdRecLen = Len(UBDGRdRec(1))
  UBDGRdFile = FreeFile
  
  KillFile "UBCUSTTR.DAT"
  UBDGRdFile = FreeFile
  Open "UBCUSTTR.DAT" For Random Shared As UBDGRdFile Len = UBDGRdRecLen
  NumDGRdRecs = LOF(UBDGRdFile) / UBDGRdRecLen
  'Open the Location Order for Reading
  IdxRecLen = 4           'we are using a integer
  IdxFileSize& = FileSize&("UBCUSTBK.IDX")
  IdxNumOfRecs = IdxFileSize& \ IdxRecLen
  ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
  FGetAH "UBCUSTBK.IDX", IdxBuff(), IdxRecLen, IdxNumOfRecs            'load it
  
  cnt = 1
  Do
    Prec& = IdxBuff(cnt).RecNum
    If Prec& > 0 Then
      Get UBFile, Prec&, UBCustRec(1)
      If CycleCnt > 0 Then
        For BookCnt = 1 To CycleCnt
        'For BookCnt = 1 To NumberofRoutes
          If Val(UBCustRec(1).Book) = Cycle(BookCnt) And (UBCustRec(1).Status <> "F") Then
            GoSub DGWriteRecord
          End If
        Next
      Else
       If (UBCustRec(1).Status <> "F") And Val(UBCustRec(1).Book) > 0 Then
        GoSub DGWriteRecord
       End If
      End If
    End If
    cnt = cnt + 1
    FrmShowPctComp.ShowPctComp cnt&, IdxNumOfRecs
  Loop Until cnt > IdxNumOfRecs
  
  Close
  MsgText(0) = "Data Genie Operation"
  MsgText(1) = "Export Complete."
  MsgText(2) = ""
  'MsgText(3) = "Exported:" + Str$(FileSize(SensusIOFile) / UBSenRdRecLen) + " Readings."
  MsgText(4) = ""
  MsgText(5) = ""
  GetOKorNot% MsgText(), True, True

Return

DGWriteRecord:
  'May Have Up to 7 Meters to Read
  For MtrCnt = 1 To 7
    MeterOK = False
    Account$ = Str$(Prec&)
    Account$ = Left$(Account$, 6) + "-" + Right$(Str$(MtrCnt), 1)
    If Asc(UBCustRec(1).LocMeters(MtrCnt).MTRType) > 32 Then
      MTRType$ = UBCustRec(1).LocMeters(MtrCnt).MTRType
      If MTRType$ = "W" Or MTRType$ = "S" Or MTRType$ = "C" Or MTRType$ = "E" Or MTRType$ = "D" Or MTRType$ = "G" Then
        Select Case WhatTypes$
        Case "W"                'water/sewer
          If MTRType$ = "W" Or MTRType$ = "S" Or MTRType$ = "C" Then
            MeterType$ = "W"
            MeterOK = True
          End If
        Case "E"                'electric & demand elec.
          If MTRType$ = "E" Or MTRType$ = "D" Then
            MeterOK = True
          End If
        Case "G"                'gas
          If MTRType$ = "G" Then
            MeterType$ = "G"
            MeterOK = True
          End If
        Case "A", " "           'all meters
          If MTRType$ = "W" Or MTRType$ = "S" Or MTRType$ = "C" Then
            MeterType$ = "W"
          End If
          If MTRType$ = "E" Or MTRType$ = "D" Then
            MeterType$ = "E"
          End If
          MeterOK = True
        End Select

        If MeterOK = True Then
          ' Determine High and Low Reading
          Average# = UBCustRec(1).LocMeters(MtrCnt).AvgUse
          TimesRead = UBCustRec(1).LocMeters(MtrCnt).UseCnt
          'ILowRead$ = Right$(Str$((UBCustRec(1).LocMeters(MtrCnt).CurRead)), 8)
          ILowRead# = UBCustRec(1).LocMeters(MtrCnt).CurRead  'Val(ILowRead$)
          LowRead# = Fix(ILowRead#)
          HighRead# = Fix(Average# * (HighVar / 100)) + UBCustRec(1).LocMeters(MtrCnt).CurRead
          
          If HighRead# < 0 Or HighRead# > 99999999 Then
            HighRead# = 0
          End If
          MeterID$ = LTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrNum)
          MeterID$ = RTrim$(MeterID$)
          
          If Val(MeterID$) = 0 Then
            MeterID$ = UBCustRec(1).Book + UBCustRec(1).SEQNUMB
          End If
          If Len(MeterID$) < 8 Then
            MeterID$ = String$(8 - Len(MeterID$), "0") + MeterID$
          End If
          MeterID$ = Left$(MeterID$, 8)

          'Set Record Fields and Put On Disk
          UBDGRdRec(1).RouteID = LTrim$(Str$(UBCustRec(1).Seq))
          UBDGRdRec(1).SvcTyp = MeterType$
          UBDGRdRec(1).CustName = UBCustRec(1).CustName
          UBDGRdRec(1).SvcLoc = UBCustRec(1).ServAddr
          UBDGRdRec(1).MeterSN = MeterID$
          UBDGRdRec(1).MeterType = "C"
          UBDGRdRec(1).High = Str$(HighRead#)
          UBDGRdRec(1).Low = Str$(LowRead#)
          UBDGRdRec(1).Msg = UBCustRec(1).HHMSG1 + " " + UBCustRec(1).HHMSG2 + " " + UBCustRec(1).HHMSG2 + " " + UBCustRec(1).HHMSG3
          UBDGRdRec(1).Account = Account$
          UBDGRdRec(1).NewRdng = ""
          UBDGRdRec(1).Account = Account$
          UBDGRdRec(1).NewRdng = ""
          UBDGRdRec(1).NewDmnd = ""
          UBDGRdRec(1).Date = ""
          UBDGRdRec(1).Time = ""
          UBDGRdRec(1).NewAcctRte = ""
          Put UBDGRdFile, (LOF(UBDGRdFile) / UBDGRdRecLen) + 1, UBDGRdRec(1)
        End If
      End If
    End If
  Next
Return

End Sub

Private Sub ImpExpLogiconHHInfo(ByVal ImpExpFlag As Boolean)
  Dim HighVar As Integer, LowVar As Integer
  Dim WhatTypes As String, CustAcc As String
  Dim UBFile As Integer, RRDate As Integer
  Dim LogiconIOFile As String
  Dim UBLogRdRecLen As Integer, UBLogGetRdRecLen As Integer
  Dim UBLogiconRDFile As Integer, NumLogGetRdRecs As Integer
  Dim BookCnt As Integer, MtrCnt As Integer
  Dim RMCnt As Long, WhatRMRec As Long
  Dim Account As String, IHighRead As String
  Dim Average As Double, LowRead As Double
  Dim MeterID As String, MRDate As String
  Dim HighRead As Double, ILowRead As Double
  Dim MeterReadDate As String, TAcct As String
  Dim WhatBook As Integer, Prec As Long
  Dim KaKa As String, MtrMult As Long
  WhatBook = -1
  
  Dim CurReading As Double, MeterRecord As Integer
  
  UBCustRecLen = Len(UBCustRec(1))
  
  ReDim UBLogRdRec(1) As UBLogiconReadRecType
  UBLogRdRecLen = Len(UBLogRdRec(1))

  ' Check For Device Type and Run Appropriate Program
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen
  
  HighVar = UBSetUpRec(1).HighRead
  LowVar = UBSetUpRec(1).LowRead
  
  If HighVar < 100 Then
    HighVar = 100           'make sure
  End If
  If LowVar > HighVar Then
    LowVar = HighVar
  End If
  
  If ImpExpFlag Then     'EXPORTING METER READINGS
    'WhatTypes$ = Left$(Me.fpMtrType.Text, 1)
    GoSub LogiconSendRead
    Call cmdExit_Click
  Else
    GoSub LogiconGetRead
    Call cmdExit_Click
  End If

Exit Sub

LogiconSendRead:
  
  'Open Customer Data File
  LogiconIOFile = HHPathInOut + "WBLOGNO" + QPTrim(Str$(InterrNum)) + ".DAT"
  
  UBFile = FreeFile
  Open UBPath + "UBCUST.DAT" For Random Shared As UBFile Len = UBCustRecLen

  'Open Logicon Date File
  ReDim UBLogRdRec(1) As UBLogiconReadRecType
  UBLogRdRecLen = Len(UBLogRdRec(1))
'  UBLogiconRDFile = FreeFile
'  Open LogiconIOFile For Random Shared As UBLogiconRDFile Len = UBLogRdRecLen
'  Close UBLogiconRDFile
  
  FrmShowPctComp.Label1 = "Exporting Meter Reading Information."
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show '1, Parent
  
  IdxRecLen = 4         'we are using a integer
  IdxFileSize& = FileSize&("UBCUSTBK.IDX")
  IdxNumOfRecs = IdxFileSize& \ IdxRecLen
  ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
  'FGetAH "UBTEMP.IDX", IdxBuff(), 4, IdxNumOfRecs
  
  FGetAH "UBCUSTBK.IDX", IdxBuff(), IdxRecLen, IdxNumOfRecs
  
  KillFile LogiconIOFile
  
  UBLogiconRDFile = FreeFile
  Open LogiconIOFile For Random Shared As UBLogiconRDFile Len = UBLogRdRecLen
  'NumLogRdRecs = LOF(UBLogiconRDFile) / UBLogRdRecLen

'   'Write First Record With Route Information
'  UBLogRdRec(1).RecType = "H"
'  UBLogRdRec(1).RouteNo = QPTrim(Str$(InterrNum))
'  UBLogRdRec(1).AcctNo = ""
'  UBLogRdRec(1).RecName = ""
'  UBLogRdRec(1).ServAddress = ""
'  UBLogRdRec(1).ReadDate = ""
'  UBLogRdRec(1).ReadTime = ""
'  UBLogRdRec(1).Consumption = ""
'  UBLogRdRec(1).PrevRead = ""
'  UBLogRdRec(1).CurRead = ""
'  UBLogRdRec(1).LowRead = ""
'  UBLogRdRec(1).HighRead = ""
'  UBLogRdRec(1).MtrNumb = ""
'  UBLogRdRec(1).CountChg = ""
'  UBLogRdRec(1).ForceFlag = ""
'  UBLogRdRec(1).ReportCode = ""
'  UBLogRdRec(1).Remark = ""
'  UBLogRdRec(1).Label = ""
'  UBLogRdRec(1).PrintFlag = ""
'  UBLogRdRec(1).MessageOut = ""
'  UBLogRdRec(1).Book = ""
'  UBLogRdRec(1).Future = ""
'  UBLogRdRec(1).Recend = "X"
'  UBLogRdRec(1).CrLf = Chr$(13) + Chr$(10)
'  Put UBLogiconRDFile, (LOF(UBLogiconRDFile) / UBLogRdRecLen) + 1, UBLogRdRec(1)

  RMCnt = 1

  Do
    Prec& = IdxBuff(RMCnt).RecNum
    If Prec& > 0 Then
      Get UBFile, Prec&, UBCustRec(1)
      For BookCnt = 1 To CycleCnt
     ' For BookCnt! = 1 To NumberofRoutes
        
        If Val(UBCustRec(1).Book) = Cycle(BookCnt) And (UBCustRec(1).Status <> "F") Then
          If WhatBook <> Val(UBCustRec(1).Book) Then
            GoSub MakeHdrRecord
          End If
          GoSub WriteLogiconRec
        End If
      Next
    End If
    RMCnt = RMCnt + 1
    FrmShowPctComp.ShowPctComp RMCnt, IdxNumOfRecs
  Loop Until RMCnt > IdxNumOfRecs

  Close
  
  MsgText(0) = "Logicon Operation"
  MsgText(1) = "Export Complete."
  MsgText(2) = ""
  MsgText(3) = "Exported:" + Str$(FileSize(LogiconIOFile) / UBLogRdRecLen) + " Readings."
  MsgText(4) = ""
  MsgText(5) = ""
  GetOKorNot% MsgText(), True, True

Return

WriteLogiconRec:
  'May Have Up to 7 Meters to Read
  MtrCnt = 1

  Account$ = Space$(6)
  LSet Account$ = QPTrim$(Str$(Prec&))

  'Account$ = LEFT$(Account$, 6) + "-" + RIGHT$(STR$(MtrCnt!), 1)

  While MtrCnt < 8

    If (Asc(UBCustRec(1).LocMeters(MtrCnt).MTRType) > 32) Then
      Select Case UBCustRec(1).LocMeters(MtrCnt).MTRType
      Case "C", "W", "T", "S", "I"
      
       'If UBCustRec(1).LocMeters(MtrCnt).MtrType = "C" Or UBCustRec(1).LocMeters(MtrCnt).MtrType = "W" Or UBCustRec(1).LocMeters(MtrCnt).MtrType = "T" Or UBCustRec(1).LocMeters(MtrCnt).MtrType = "S" Then
        Mid$(Account$, 6, 1) = QPTrim$(Str$(MtrCnt))
        ' Determine High and Low Reading
        Average# = UBCustRec(1).LocMeters(MtrCnt).AvgUse
        MtrMult = UBCustRec(1).LocMeters(MtrCnt).MTRMulti
        If MtrMult < 1 Then
          MtrMult = 1
        End If
        'ILowRead$ = Right$(Str$((UBCustRec(1).LocMeters(MtrCnt).CurRead)), 8)
        If Average# < 1 Then
          Average# = 1
        End If
        ILowRead# = UBCustRec(1).LocMeters(MtrCnt).CurRead
        HighRead# = Fix(Average# * (HighVar / 100)) + UBCustRec(1).LocMeters(MtrCnt).CurRead
        If Fix(HighRead#) = ILowRead# Then
          HighRead# = HighRead# + 12000
        End If
        IHighRead$ = Str$(HighRead#)
        IHighRead$ = Right$(IHighRead$, 8)
        MeterID$ = LTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrNum)

        MeterID$ = QPTrim$(MeterID$)
        
        'If Val(MeterID$) = 0 Then
        '  MeterID$ = UBCustRec(1).Book + UBCustRec(1).SEQNUMB
        'End If
        If Len(MeterID$) < 8 Then
          MeterID$ = String$(8 - Len(MeterID$), "0") + MeterID$
        End If
        
        MeterID$ = Left$(MeterID$, 8)

        UBLogRdRec(1).RecType = "A"
        UBLogRdRec(1).RouteNo = QPTrim(Str$(InterrNum))
        UBLogRdRec(1).AcctNo = Account$
        UBLogRdRec(1).RecName = UBCustRec(1).CustName
        UBLogRdRec(1).ServAddress = UBCustRec(1).ServAddr
        UBLogRdRec(1).ReadDate = "      "
        UBLogRdRec(1).ReadTime = "      "
        UBLogRdRec(1).Consumption = "        "
        UBLogRdRec(1).PrevRead = Str$(ILowRead#)
        UBLogRdRec(1).CurRead = "XXXXXXXX"
        UBLogRdRec(1).LowRead = Str$(ILowRead#)
        UBLogRdRec(1).HighRead = IHighRead$
        UBLogRdRec(1).MtrNumb = MeterID$
        UBLogRdRec(1).CountChg = "0"
        UBLogRdRec(1).ForceFlag = " "
        UBLogRdRec(1).ReportCode = "--"
        UBLogRdRec(1).Remark = ""
        UBLogRdRec(1).Label = ""
        UBLogRdRec(1).PrintFlag = ""
        KaKa = QPTrim$(UBCustRec(1).HHMSG1) + " " + QPTrim$(UBCustRec(1).HHMSG2) + " " + QPTrim$(UBCustRec(1).HHMSG3)
        UBLogRdRec(1).MessageOut = KaKa
        UBLogRdRec(1).Book = UBCustRec(1).Book
        UBLogRdRec(1).MtrSize = QPTrim$(UBCustRec(1).USERCODE2)
        
        LSet UBLogRdRec(1).AvgUse = QPTrim$(Str$(Average# * MtrMult))
        UBLogRdRec(1).Future = ""
        UBLogRdRec(1).Recend = "X"
        UBLogRdRec(1).CrLf = Chr$(13) + Chr$(10)
        Put UBLogiconRDFile, (LOF(UBLogiconRDFile) / UBLogRdRecLen) + 1, UBLogRdRec(1)
      End Select
    End If

SkipEmLC:
    MtrCnt = MtrCnt + 1
  Wend

Return


LogiconGetRead:
  
  'PathWay$ = QPTrim$(PathWay$)
  'LogiconIOFile = PathWay$ + "WBLOGNO" + LTrim$(RouteID$) + ".DAT"
  'build sensus output file name
  
  LogiconIOFile = HHPathInOut + "WBLOGNO" + QPTrim(Str$(InterrNum)) + ".DAT"
  
  MRDate$ = QPTrim$(fpReadDate.Text)
  RRDate = Date2Num(MRDate$)
  
  MsgText(0) = "Import Logicon Reading File."
  MsgText(1) = ""
  MsgText(2) = "Import File:"
  MsgText(3) = LogiconIOFile
  MsgText(4) = "Ready to Proceed?"
  MsgText(5) = ""
  
  Select Case GetOKorNot%(MsgText(), False, True, 1)
  Case False
    GoTo LogicGetExit
  End Select
  
  ReDim UBLogGetRdRec(1) As UBLogiconGetReadRecType
  UBLogGetRdRecLen = Len(UBLogGetRdRec(1))
  
  UBFile = FreeFile
  Open UBPath + "UBCUST.DAT" For Random Shared As UBFile Len = UBCustRecLen

  UBLogiconRDFile = FreeFile
  Open LogiconIOFile For Random Shared As UBLogiconRDFile Len = UBLogGetRdRecLen
  NumLogGetRdRecs = LOF(UBLogiconRDFile) / UBLogGetRdRecLen

  If NumLogGetRdRecs = 0 Then
    Close
    MsgText(0) = "ERROR:"
    MsgText(1) = "IMPORT FILE NOT FOUND"
    MsgText(2) = "Make sure that: " + "'" + "WBLOGNO" + QPTrim(Str$(InterrNum)) + ".DAT" + "'"
    MsgText(3) = "is in the Logicon directory!"
    MsgText(4) = "Please call Southern Software for"
    MsgText(5) = "additional Information."
    GetOKorNot% MsgText(), True, False
    GoTo LogicGetExit
  End If

  RMCnt = 1                ' Initialize File Counter to 1
  Do
    Get UBLogiconRDFile, RMCnt, UBLogGetRdRec(1)

    TAcct$ = Left$(UBLogGetRdRec(1).AcctNo, 5)
    Prec& = Val(TAcct$)

    If Left$(UBLogGetRdRec(1).CurRead, 1) <> "X" Then
      If Prec& > 0 Then
        Get UBFile, Prec&, UBCustRec(1)
        GoSub ExtractRecordLC
      End If
    End If
    RMCnt = RMCnt + 1

  Loop Until RMCnt > NumLogGetRdRecs
  
  Close
  
  'Done show import complete
  MsgText(0) = ""
  MsgText(1) = "Logicon Operation"
  MsgText(2) = ""
  MsgText(3) = "Reading Import Complete."
  MsgText(4) = ""
  MsgText(5) = ""
  GetOKorNot% MsgText(), True, True

LogicGetExit:
Return

ExtractRecordLC:
  
  MeterRecord = Val(Right$(UBLogGetRdRec(1).AcctNo, 1))
  
  CurReading# = Val(UBLogGetRdRec(1).CurRead)
  If RRDate > 0 Then
    MeterReadDate$ = Num2Date$(RRDate)
  Else
    MeterReadDate$ = Mid$(UBLogGetRdRec(1).ReadDate, 3, 2) + "/" + Mid$(UBLogGetRdRec(1).ReadDate, 5, 2) + "/" + Right$(Date$, 4)
  End If

  If UBCustRec(1).LocMeters(MeterRecord).ReadFlag = "Y" Then
    UBCustRec(1).LocMeters(MeterRecord).CurRead = CurReading#
    UBCustRec(1).LocMeters(MeterRecord).CurDate = Date2Num(MeterReadDate$)
  Else
    UBCustRec(1).LocMeters(MeterRecord).PrevRead = UBCustRec(1).LocMeters(MeterRecord).CurRead
    UBCustRec(1).LocMeters(MeterRecord).PastDate = UBCustRec(1).LocMeters(MeterRecord).CurDate
    UBCustRec(1).LocMeters(MeterRecord).ReadFlag = "Y"
    UBCustRec(1).LocMeters(MeterRecord).CurDate = Date2Num(MeterReadDate$)
    UBCustRec(1).LocMeters(MeterRecord).CurRead = CurReading#
  End If
  Put UBFile, Prec&, UBCustRec(1)
Return

MakeHdrRecord:
  ReDim UBLogRdRec(1) As UBLogiconReadRecType
  WhatBook = Val(UBCustRec(1).Book)
  UBLogRdRec(1).RecType = "H"
  UBLogRdRec(1).RouteNo = QPTrim(Str$(WhatBook))
  UBLogRdRec(1).AcctNo = ""
  UBLogRdRec(1).RecName = ""
  UBLogRdRec(1).ServAddress = ""
  UBLogRdRec(1).ReadDate = ""
  UBLogRdRec(1).ReadTime = ""
  UBLogRdRec(1).Consumption = ""
  UBLogRdRec(1).PrevRead = ""
  UBLogRdRec(1).CurRead = ""
  UBLogRdRec(1).LowRead = ""
  UBLogRdRec(1).HighRead = ""
  UBLogRdRec(1).MtrNumb = ""
  UBLogRdRec(1).CountChg = ""
  UBLogRdRec(1).ForceFlag = ""
  UBLogRdRec(1).ReportCode = ""
  UBLogRdRec(1).Remark = ""
  UBLogRdRec(1).Label = ""
  UBLogRdRec(1).PrintFlag = ""
  UBLogRdRec(1).MessageOut = ""
  UBLogRdRec(1).Book = ""
  UBLogRdRec(1).MtrSize = ""
  UBLogRdRec(1).AvgUse = ""
  UBLogRdRec(1).Future = ""
  UBLogRdRec(1).Recend = "X"
  UBLogRdRec(1).CrLf = Chr$(13) + Chr$(10)
  Put UBLogiconRDFile, (LOF(UBLogiconRDFile) / UBLogRdRecLen) + 1, UBLogRdRec(1)
Return

End Sub

'************************************************************************************
Private Sub ImpExpUSenHHInfo(ByVal ImpExpFlag As Boolean)
  Dim HighVar As Integer, LowVar As Integer
  Dim WhatTypes As String, CustAcc As String
  Dim UBFile As Integer, UBSenIOFile As Integer
  Dim UBSenRdRecLen As Integer, NumSenRdRecs As Integer
  Dim BookCnt As Integer, MtrCnt As Integer
  Dim RMCnt As Long, WhatRMRec As Long
  Dim Account As String, SensusType As String
  Dim Average As Double, LowRead As Double
  Dim MeterID As String, MRDate As String, UCode1 As String
  'Dim ReadLowI As String, PrevRead As String
  'Dim NCurRead As String, PrevDate As String
  Dim HighRead As Double, ILowRead As Double
  Dim UBSenGetRecLen As Integer, NumSenGetRecs As Integer
  Dim MeterReadDate As String
  Dim DashPos As Integer
  Dim CurReading As Double
  
  UBCustRecLen = Len(UBCustRec(1))
  
  ReDim UBSenRdRec(1) As UBOESensusReadRecType
  
  ' Check For Device Type and Run Appropriate Program
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen
  
  HighVar = UBSetUpRec(1).HighRead
  LowVar = UBSetUpRec(1).LowRead
  
  If HighVar < 100 Then
    HighVar = 100           'make sure
  End If
  If LowVar > HighVar Then
    LowVar = HighVar
  End If
  
  If ImpExpFlag Then     'EXPORTING METER READINGS
    WhatTypes$ = Left$(Me.fpMtrType.Text, 1)
    GoSub ESendSensus
    Call cmdExit_Click
  Else
    GoSub EGetSensus
    Call cmdExit_Click
  End If
  
Exit Sub
  
EGetSensus:

  SensusIOFile = HHPathInOut + "exssi00" + QPTrim(Str$(InterrNum)) + ".DAT"

  MRDate$ = QPTrim$(fpReadDate.Text)
  
  MsgText(0) = "Import Sensus Reading File."
  MsgText(1) = ""
  MsgText(2) = "Import File:"
  MsgText(3) = SensusIOFile
  MsgText(4) = "Ready to Proceed?"
  MsgText(5) = ""
  
  Select Case GetOKorNot%(MsgText(), False, True, 1)
  Case False
    GoTo ErrorEGetSensusExit
    'Stop
  End Select
    
  ReDim UBSenGetRdRec(1) As UBOESensusGetReadRecType
  UBSenGetRecLen = Len(UBSenGetRdRec(1))
  
  UBSenIOFile = FreeFile
  Open SensusIOFile For Random Shared As UBSenIOFile Len = UBSenGetRecLen
  
  NumSenGetRecs = LOF(UBSenIOFile) / UBSenGetRecLen
  
  UBFile = FreeFile
  Open UBPath + "UBCUST.DAT" For Random Shared As UBFile Len = UBCustRecLen
        
  If NumSenGetRecs = 0 Then
    Close
    MsgText(0) = "ERROR:"
    MsgText(1) = "IMPORT FILE NOT FOUND"
    MsgText(2) = "Make sure that: " + "'" + "exssi00" + QPTrim(Str$(InterrNum)) + ".DAT" + "'"
    MsgText(3) = "is in the Sensus directory!"
    MsgText(4) = "Please call Southern Software for"
    MsgText(5) = "additional Information."
    GetOKorNot% MsgText(), True, False
    GoTo ErrorEGetSensusExit
  End If
  
  FrmShowPctComp.Label1 = "Importing Meter Reading Information."
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show '1, Parent
  
  RMCnt& = 1                ' Initialize File Counter to 1
  Do
    Get UBSenIOFile, RMCnt&, UBSenGetRdRec(1)
    WhatRMRec = Val(QPTrim$(UBSenGetRdRec(1).Account))
    If WhatRMRec > 0 Then
      GoSub EExtractRecord
    End If
    RMCnt& = RMCnt& + 1
    FrmShowPctComp.ShowPctComp RMCnt&, IdxNumOfRecs
  Loop Until RMCnt& > NumSenGetRecs
  
  Close
  
  'Done show import complete
  MsgText(0) = "Sensus Operation"
  MsgText(1) = "Sensus Import Complete."
  MsgText(2) = ""
  MsgText(3) = "Imported:" + Str$(FileSize(SensusIOFile) / UBSenGetRecLen) + " Readings."
  MsgText(4) = ""
  MsgText(5) = ""
  GetOKorNot% MsgText(), True, True

ErrorEGetSensusExit:

Return

EExtractRecord:
  
  Get UBFile, WhatRMRec, UBCustRec(1)
  DashPos = InStr(UBSenGetRdRec(1).Account, "-")
  MtrCnt = Val(Mid$(UBSenGetRdRec(1).Account, DashPos + 1))
  
  If MtrCnt = 0 Then MtrCnt = 1
  ' Check Meter Updated Flag
  ' Update Meter W/Reading
  CurReading# = Val(UBSenGetRdRec(1).CurRead)
  MeterReadDate$ = Left$(UBSenGetRdRec(1).ReadDate, 2) + "/" + Mid$(UBSenGetRdRec(1).ReadDate, 3, 2) + "/" + Mid$(MRDate$, 7, 2) + Right$(UBSenGetRdRec(1).ReadDate, 2)
  If Date2Num(MeterReadDate$) < 0 Then
    MeterReadDate$ = MRDate$
  End If
  
  If UBCustRec(1).LocMeters(MtrCnt).ReadFlag = "Y" Then
    UBCustRec(1).LocMeters(MtrCnt).CurRead = CurReading#
    UBCustRec(1).LocMeters(MtrCnt).CurDate = Date2Num(MeterReadDate$)
  Else
    UBCustRec(1).LocMeters(MtrCnt).PrevRead = UBCustRec(1).LocMeters(MtrCnt).CurRead
    UBCustRec(1).LocMeters(MtrCnt).PastDate = UBCustRec(1).LocMeters(MtrCnt).CurDate
    UBCustRec(1).LocMeters(MtrCnt).ReadFlag = "Y"
    UBCustRec(1).LocMeters(MtrCnt).CurDate = Date2Num(MeterReadDate$)
    UBCustRec(1).LocMeters(MtrCnt).CurRead = CurReading#
  End If
  
  Put UBFile, WhatRMRec, UBCustRec(1)
  
Return

'************************* Send info to sensus
ESendSensus:

  If CycleCnt > 0 Then
    'build sensus output file name
    SensusIOFile = HHPathInOut + "SSI00" + QPTrim(Str$(InterrNum)) + ".RTE"
      
    KillFile SensusIOFile 'kill old if there
      
    If UBSetUpRec(1).UseSeq = "Y" Then 'if they are using sequence numbers
      IdxRecLen = 4         'we are using a integer
      MakeSequenceIndex "Sequence Number", Me
      IdxNumOfRecs = FileSize&("UBTEMP.IDX") \ 4
      ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
      FGetAH "UBTEMP.IDX", IdxBuff(), 4, IdxNumOfRecs
    Else                               'use location number index
      IdxRecLen = 4         'we are using a integer
      IdxFileSize& = FileSize&("UBCUSTBK.IDX")
      IdxNumOfRecs = IdxFileSize& \ IdxRecLen
      ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
      FGetAH "UBCUSTBK.IDX", IdxBuff(), IdxRecLen, IdxNumOfRecs            'load it
    End If
    
    FrmShowPctComp.Label1 = "Exporting Meter Reading Information."
    FrmShowPctComp.cmdCancel.Enabled = False
    FrmShowPctComp.Show '1, Parent
    
    UBSenRdRecLen = Len(UBSenRdRec(1))
    
    UBSenIOFile = FreeFile
    Open SensusIOFile For Random Shared As UBSenIOFile Len = UBSenRdRecLen
    
    NumSenRdRecs = LOF(UBSenIOFile) / UBSenRdRecLen
        
    UBFile = FreeFile
    Open UBPath + "UBCUST.DAT" For Random Shared As UBFile Len = UBCustRecLen
    
    RMCnt& = 1
    Do
      WhatRMRec& = IdxBuff(RMCnt&).RecNum
      If WhatRMRec& > 0 Then
        Get UBFile, WhatRMRec&, UBCustRec(1)
        For BookCnt = 1 To CycleCnt
          If Val(UBCustRec(1).Book) = Cycle(BookCnt) And (UBCustRec(1).Status <> "F") Then
            GoSub EWriteRecord
            Exit For
          End If
        Next
      End If
      RMCnt& = RMCnt& + 1
      FrmShowPctComp.ShowPctComp RMCnt&, IdxNumOfRecs
    Loop Until RMCnt& > IdxNumOfRecs
    
    Close   'done with output file.
    
    MsgText(0) = "Sensus Operation"
    MsgText(1) = "Sensus Export Complete."
    MsgText(2) = ""
    MsgText(3) = "Exported:" + Str$(FileSize(SensusIOFile) / UBSenRdRecLen) + " Readings."
    MsgText(4) = ""
    MsgText(5) = ""
    GetOKorNot% MsgText(), True, True
  
  End If

Return

EWriteRecord:
'*****
  Account$ = Str$(WhatRMRec&)
  For MtrCnt = 1 To 7     'look at all possiable meters
    
    Select Case UBCustRec(1).LocMeters(MtrCnt).MTRType
    Case "C", "S", "W", "T", "E", "D", "P", "I"   'here dale
      If (UBCustRec(1).LocMeters(MtrCnt).MTRType = "T" And Val(QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrNum)) = 0) Then
        GoTo ESkipEm
      Else
        ' Determine Sensus Meter Type
        Select Case UBCustRec(1).LocMeters(MtrCnt).MTRType
        Case "T", "I"
            UCode1$ = QPTrim$(UBCustRec(1).USERCODE1)
            If InStr(UCode1$, "P") > 0 Or InStr(UCode1$, "p") > 0 Then
              SensusType$ = "P"
            Else
              SensusType$ = "B"
            End If
       ' Case "P"
          'SensusType$ = "P"
        Case Else
          SensusType$ = "M"
        End Select
        'Determine High and Low Reading
        Average# = UBCustRec(1).LocMeters(MtrCnt).AvgUse
        If Average# < 0 Then
          Average# = 0
        End If
        
        ILowRead# = Val(QPTrim$(Str$(UBCustRec(1).LocMeters(MtrCnt).CurRead)))
        HighRead# = Fix(Average# * (HighVar / 100)) + UBCustRec(1).LocMeters(MtrCnt).CurRead
        
        If Fix(HighRead#) = ILowRead# Then
          HighRead# = HighRead# + 5
        End If
        
        MeterID$ = LTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrNum)
        MeterID$ = RTrim$(MeterID$)
        If Val(MeterID$) = 0 Then
          MeterID$ = UBCustRec(1).Book + UBCustRec(1).SEQNUMB
        End If
        If Len(MeterID$) < 8 Then
          MeterID$ = String$(8 - Len(MeterID$), "0") + MeterID$
        End If
        
        MeterID$ = Left$(MeterID$, 8)
'Set Record Fields and Put On Disk
'************ clear old info
        UBSenRdRec(1).CustLastName = ""
        UBSenRdRec(1).CustFirstName = ""
        UBSenRdRec(1).MeterID = ""
        UBSenRdRec(1).Account = ""
        UBSenRdRec(1).LowRead = ""
        UBSenRdRec(1).HighRead = ""
        UBSenRdRec(1).SensusType = ""
        UBSenRdRec(1).PastRead = ""
        UBSenRdRec(1).CurRead = ""
        UBSenRdRec(1).ServAddress = ""
        UBSenRdRec(1).LocationNumber = ""
        UBSenRdRec(1).Message = ""
'***************************
        
        UBSenRdRec(1).ServAddress = QPTrim$(UBCustRec(1).ServAddr)
        UBSenRdRec(1).MeterID = MeterID$
        UBSenRdRec(1).LowRead = Str$(ILowRead#)
        UBSenRdRec(1).HighRead = Str$(HighRead#)
        UBSenRdRec(1).Account = Account$ + "-" + QPTrim$(Str$(MtrCnt))
        UBSenRdRec(1).SensusType = SensusType$
        UBSenRdRec(1).CustLastName = QPTrim$(UBCustRec(1).CustName)
        UBSenRdRec(1).CustFirstName = ""
        UBSenRdRec(1).LocationNumber = QPTrim$(UBCustRec(1).Book + UBCustRec(1).SEQNUMB)
        UBSenRdRec(1).Message = QPTrim$(UBCustRec(1).HHMSG1)
        
        Put UBSenIOFile, , UBSenRdRec(1)
      End If
    Case Else
      'no meter in this slot.
    End Select
ESkipEm:
  Next
Return

End Sub
'************************************************************************************

Private Sub ImpExpESenHHInfo(ByVal ImpExpFlag As Boolean)
  Dim HighVar As Integer, LowVar As Integer
  Dim WhatTypes As String, CustAcc As String
  Dim UBFile As Integer, UBSenIOFile As Integer
  Dim UBSenRdRecLen As Integer, NumSenRdRecs As Integer
  Dim BookCnt As Integer, MtrCnt As Integer
  Dim RMCnt As Long, WhatRMRec As Long
  Dim Account As String, SensusType As String
  Dim Average As Double, LowRead As Double
  Dim MeterID As String, MRDate As String, UCode1 As String
  'Dim ReadLowI As String, PrevRead As String
  'Dim NCurRead As String, PrevDate As String
  Dim HighRead As Double, ILowRead As Double
  Dim UBSenGetRecLen As Integer, NumSenGetRecs As Integer
  Dim MeterReadDate As String
  Dim DashPos As Integer
  Dim CurReading As Double
  Dim UBCustcount As Long
  UBCustRecLen = Len(UBCustRec(1))
  
  ReDim UBSenRdRec(1) As UBGilSensusReadRecType
  
  ' Check For Device Type and Run Appropriate Program
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen
  
'  If InStr(UBSetUpRec(1).UTILNAME, "PENNINGTON") > 0 Then
'    PenGapFlag = True
'  End If
  
  HighVar = UBSetUpRec(1).HighRead
  LowVar = UBSetUpRec(1).LowRead
  
  If HighVar < 100 Then
    HighVar = 100           'make sure
  End If
  If LowVar > HighVar Then
    LowVar = HighVar
  End If
  
  If ImpExpFlag Then     'EXPORTING METER READINGS
    WhatTypes$ = Left$(Me.fpMtrType.Text, 1)
    GoSub ESendSensus
    Call cmdExit_Click
  Else
    GoSub EGetSensus
    Call cmdExit_Click
  End If
  
Exit Sub
  
EGetSensus:
  If InterrNum < 10 Then
    SensusIOFile = HHPathInOut + "exssi00" + QPTrim(Str$(InterrNum)) + ".DAT"
  Else
    SensusIOFile = HHPathInOut + "exssi0" + QPTrim(Str$(InterrNum)) + ".DAT"
  End If

  MRDate$ = QPTrim$(fpReadDate.Text)
  
  MsgText(0) = "Import Sensus Reading File."
  MsgText(1) = ""
  MsgText(2) = "Import File:"
  MsgText(3) = SensusIOFile
  MsgText(4) = "Ready to Proceed?"
  MsgText(5) = ""
  
  Select Case GetOKorNot%(MsgText(), False, True, 1)
  Case False
    GoTo ErrorEGetSensusExit
    'Stop
  End Select
    
  ReDim UBSenGetRdRec(1) As UBGilSensusGetReadRecType
  UBSenGetRecLen = Len(UBSenGetRdRec(1))
  
  UBSenIOFile = FreeFile
  Open SensusIOFile For Random Shared As UBSenIOFile Len = UBSenGetRecLen
  
  NumSenGetRecs = LOF(UBSenIOFile) / UBSenGetRecLen
  
  UBFile = FreeFile
  Open UBPath + "UBCUST.DAT" For Random Shared As UBFile Len = UBCustRecLen
  UBCustcount = LOF(UBFile) / UBCustRecLen
  If NumSenGetRecs = 0 Then
    Close
    MsgText(0) = "ERROR:"
    MsgText(1) = "IMPORT FILE NOT FOUND"
    MsgText(2) = "Make sure that: " + "'" + "exssi00" + QPTrim(Str$(InterrNum)) + ".DAT" + "'"
    MsgText(3) = "is in the Sensus directory!"
    MsgText(4) = "Please call Southern Software for"
    MsgText(5) = "additional Information."
    GetOKorNot% MsgText(), True, False
    GoTo ErrorEGetSensusExit
  End If
  
  FrmShowPctComp.Label1 = "Importing Meter Reading Information."
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show '1, Parent
  
  RMCnt& = 1                ' Initialize File Counter to 1
  Do
    Get UBSenIOFile, RMCnt&, UBSenGetRdRec(1)
    WhatRMRec = Val(QPTrim$(UBSenGetRdRec(1).Account))
    If WhatRMRec > 0 Then
      'If WhatRMRec = 4913 Then Stop
      GoSub EExtractRecord
      
    End If
   FrmShowPctComp.ShowPctComp RMCnt&, NumSenGetRecs
    RMCnt& = RMCnt& + 1
  Loop Until RMCnt& > NumSenGetRecs
  
  Close
  
  'Done show import complete
  MsgText(0) = "Sensus Operation"
  MsgText(1) = "Sensus Import Complete."
  MsgText(2) = ""
  MsgText(3) = "Imported:" + Str$(FileSize(SensusIOFile) / UBSenGetRecLen) + " Readings."
  MsgText(4) = ""
  MsgText(5) = ""
  GetOKorNot% MsgText(), True, True

ErrorEGetSensusExit:

Return

EExtractRecord:
  If WhatRMRec <= UBCustcount Then
  Get UBFile, WhatRMRec, UBCustRec(1)
  DashPos = InStr(UBSenGetRdRec(1).Account, "-")
  MtrCnt = Val(Mid$(UBSenGetRdRec(1).Account, DashPos + 1))
  
  If MtrCnt = 0 Then MtrCnt = 1
  ' Check Meter Updated Flag
  ' Update Meter W/Reading
  CurReading# = Val(UBSenGetRdRec(1).CurRead)
'  MeterReadDate$ = Left$(UBSenGetRdRec(1).ReadDate, 2) + "/" + Mid$(UBSenGetRdRec(1).ReadDate, 3, 2) + "/" + Mid$(Form$(2, 0), 7, 2) + Right$(UBSenGetRdRec(1).ReadDate, 2)
  MeterReadDate$ = Left$(UBSenGetRdRec(1).ReadDate, 2) + "/" + Mid$(UBSenGetRdRec(1).ReadDate, 3, 2) + "/" + Mid$(MRDate$, 7, 2) + Right$(UBSenGetRdRec(1).ReadDate, 2)
  If Date2Num(MeterReadDate$) <= 0 Then
    MeterReadDate$ = MRDate$
  End If
  If Val(QPTrim$(UBSenGetRdRec(1).MtrLat)) <> 0 Then
    UBCustRec(1).LocMeters(MtrCnt).MtrLat = Val(QPTrim$(UBSenGetRdRec(1).MtrLat))
  End If
  If Val(QPTrim$(UBSenGetRdRec(1).MtrLng)) <> 0 Then
    UBCustRec(1).LocMeters(MtrCnt).MtrLng = Val(QPTrim$(UBSenGetRdRec(1).MtrLng))
  End If
  If UBCustRec(1).LocMeters(MtrCnt).ReadFlag = "Y" Then
    UBCustRec(1).LocMeters(MtrCnt).CurRead = CurReading#
    UBCustRec(1).LocMeters(MtrCnt).CurDate = Date2Num(MeterReadDate$)
  Else
    UBCustRec(1).LocMeters(MtrCnt).PrevRead = UBCustRec(1).LocMeters(MtrCnt).CurRead
    UBCustRec(1).LocMeters(MtrCnt).PastDate = UBCustRec(1).LocMeters(MtrCnt).CurDate
    UBCustRec(1).LocMeters(MtrCnt).ReadFlag = "Y"
    UBCustRec(1).LocMeters(MtrCnt).CurDate = Date2Num(MeterReadDate$)
    UBCustRec(1).LocMeters(MtrCnt).CurRead = CurReading#
  End If
  
  Put UBFile, WhatRMRec, UBCustRec(1)
  Else
    FrmShowPctComp.ShowPctComp NumSenGetRecs + 1, NumSenGetRecs
    'Unload FrmShowPctComp
    MsgBox "Invalid Account Number" + Str(WhatRMRec), vbCritical, "Error"
    Exit Sub
  End If
Return

'************************* Send info to sensus
ESendSensus:

  If CycleCnt > 0 Then
    'build sensus output file name
    SensusIOFile = HHPathInOut + "SSI00" + QPTrim(Str$(InterrNum)) + ".RTE"
      
    KillFile SensusIOFile 'kill old if there
      
    If UBSetUpRec(1).UseSeq = "Y" Then 'if they are using sequence numbers
      IdxRecLen = 4         'we are using a integer
      MakeSequenceIndex "Sequence Number", Me
      IdxNumOfRecs = FileSize&("UBTEMP.IDX") \ 4
      ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
      FGetAH "UBTEMP.IDX", IdxBuff(), 4, IdxNumOfRecs
    Else                               'use location number index
      IdxRecLen = 4         'we are using a integer
      IdxFileSize& = FileSize&("UBCUSTBK.IDX")
      IdxNumOfRecs = IdxFileSize& \ IdxRecLen
      ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
      FGetAH "UBCUSTBK.IDX", IdxBuff(), IdxRecLen, IdxNumOfRecs            'load it
    End If
    
    FrmShowPctComp.Label1 = "Exporting Meter Reading Information."
    FrmShowPctComp.cmdCancel.Enabled = False
    FrmShowPctComp.Show '1, Parent
    
    UBSenRdRecLen = Len(UBSenRdRec(1))
    
    UBSenIOFile = FreeFile
    Open SensusIOFile For Random Shared As UBSenIOFile Len = UBSenRdRecLen
    
    NumSenRdRecs = LOF(UBSenIOFile) / UBSenRdRecLen
        
    UBFile = FreeFile
    Open UBPath + "UBCUST.DAT" For Random Shared As UBFile Len = UBCustRecLen
    
    RMCnt& = 1
    Do
      WhatRMRec& = IdxBuff(RMCnt&).RecNum
      If WhatRMRec& > 0 Then
        Get UBFile, WhatRMRec&, UBCustRec(1)
        For BookCnt = 1 To CycleCnt
          If Val(UBCustRec(1).Book) = Cycle(BookCnt) And (UBCustRec(1).Status <> "F") Then
            GoSub EWriteRecord
            Exit For
          End If
        Next
      End If
      RMCnt& = RMCnt& + 1
      FrmShowPctComp.ShowPctComp RMCnt&, IdxNumOfRecs
    Loop Until RMCnt& > IdxNumOfRecs
    
    Close   'done with output file.
    
    MsgText(0) = "Sensus Operation"
    MsgText(1) = "Sensus Export Complete."
    MsgText(2) = ""
    MsgText(3) = "Exported:" + Str$(FileSize(SensusIOFile) / UBSenRdRecLen) + " Readings."
    MsgText(4) = ""
    MsgText(5) = ""
    GetOKorNot% MsgText(), True, True
  
  End If

Return

EWriteRecord:
'*****
  Account$ = Str$(WhatRMRec&)
  For MtrCnt = 1 To 7     'look at all possiable meters
    
    Select Case UBCustRec(1).LocMeters(MtrCnt).MTRType
    Case "C", "S", "W", "T", "E", "D", "P", "I", "G"  'here dale
      If (UBCustRec(1).LocMeters(MtrCnt).MTRType = "T" And Len(QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrNum)) = 0) Then
        GoTo ESkipEm
      Else
        ' Determine Sensus Meter Type
        Select Case UBCustRec(1).LocMeters(MtrCnt).MTRType
        Case "T", "I"
          UCode1$ = QPTrim$(UBCustRec(1).USERCODE1)
          If InStr(UCode1$, "P") > 0 Or InStr(UCode1$, "p") > 0 Then
            SensusType$ = "P"
          Else
            SensusType$ = "B"
          End If
'
'        Case "P"
'          SensusType$ = "P"
        Case Else
          SensusType$ = "M"
        End Select
        'Determine High and Low Reading
        Average# = UBCustRec(1).LocMeters(MtrCnt).AvgUse
        If Average# < 0 Then
          Average# = 0
        End If
        
        ILowRead# = Val(QPTrim$(Str$(UBCustRec(1).LocMeters(MtrCnt).CurRead)))
        HighRead# = Fix(Average# * (HighVar / 100)) + UBCustRec(1).LocMeters(MtrCnt).CurRead
        
        If Fix(HighRead#) = ILowRead# Then
          HighRead# = HighRead# + 5
        End If
        
        MeterID$ = LTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrNum)
        MeterID$ = RTrim$(MeterID$)
        If Len(MeterID$) = 0 Then
          MeterID$ = UBCustRec(1).Book + UBCustRec(1).SEQNUMB
        End If
        If Len(MeterID$) < 8 Then
          MeterID$ = String$(8 - Len(MeterID$), "0") + MeterID$
        End If
        
        MeterID$ = Left$(MeterID$, 8)
'Set Record Fields and Put On Disk
'************ clear old info
        UBSenRdRec(1).CustLastName = ""
        UBSenRdRec(1).CustFirstName = ""
        UBSenRdRec(1).MeterID = ""
        UBSenRdRec(1).Account = ""
        UBSenRdRec(1).LowRead = ""
        UBSenRdRec(1).HighRead = ""
        UBSenRdRec(1).SensusType = ""
        UBSenRdRec(1).PastRead = ""
        UBSenRdRec(1).CurRead = ""
        UBSenRdRec(1).ServAddress = ""
        UBSenRdRec(1).LocationNumber = ""
        UBSenRdRec(1).Message = ""
        UBSenRdRec(1).MtrIDMST = ""
        UBSenRdRec(1).MtrIDNO = ""
'***************************
        
        UBSenRdRec(1).ServAddress = QPTrim$(UBCustRec(1).ServAddr)
        UBSenRdRec(1).MeterID = MeterID$
        UBSenRdRec(1).LowRead = Str$(ILowRead#)
        UBSenRdRec(1).HighRead = Str$(HighRead#)
        UBSenRdRec(1).Account = Account$ + "-" + QPTrim$(Str$(MtrCnt))
        UBSenRdRec(1).SensusType = SensusType$
        UBSenRdRec(1).CustLastName = QPTrim$(UBCustRec(1).CustName)
        UBSenRdRec(1).CustFirstName = ""
        UBSenRdRec(1).LocationNumber = QPTrim$(UBCustRec(1).Book + UBCustRec(1).SEQNUMB)
        UBSenRdRec(1).Message = QPTrim$(UBCustRec(1).HHMSG1)
        'reuse the MeterID$ variable
        MeterID$ = QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrIDNO)
        If Len(MeterID$) > 0 Then
          UBSenRdRec(1).MtrIDMST = Left$(MeterID$, Len(MeterID$) - 1)
          UBSenRdRec(1).MtrIDNO = Right$(MeterID$, 1)
        End If
           
        UBSenRdRec(1).MtrLat = MakeExpCoordinate$(UBCustRec(1).LocMeters(MtrCnt).MtrLat)
        UBSenRdRec(1).MtrLng = MakeExpCoordinate$(UBCustRec(1).LocMeters(MtrCnt).MtrLng)
        
        Put UBSenIOFile, , UBSenRdRec(1)
      End If
    Case Else
      'no meter in this slot.
    End Select
ESkipEm:
  Next
Return

End Sub

Private Sub ImpExpSchulmHHInfo(ByVal ImpExpFlag As Boolean)
  Dim ExportFileName As String, ImportFileName As String
  Dim FF As String, q As String, cb As String
  Dim HighVar As Integer, LowVar As Integer
  Dim WhatTypes As String
  ReDim UBSchlumHHRec(1) As SchlumHHType
  Dim UBFile As Integer, UBSchlFile As Integer
  Dim RptHandle As Integer
  Dim UBSchlumHHRecLen As Integer
  Dim RMCnt As Long, WhatRMRec As Long
  Dim PageCnt As Long, WriteCnt As Long
  Dim MtrCnt As Integer
  Dim MeterOK As Boolean, HasZ As Boolean
  Dim BadDate As Boolean
  Dim Account As String, MTRType As String
  Dim MeterID As String
  Dim SRouteID As String
  Dim WalkSeq As String, PageNum As String
  Dim Dials As Integer, Page As Integer, LineCnt As Integer
  Dim RecStat As String
  Dim Average As Double, LowRead As Double
  Dim ReadLowI As String, PrevRead As String
  Dim NCurRead As String, PrevDate As String
  Dim HighRead As Double, ILowRead As Double
  Dim HiRead As String
  Dim ReportFile As String
  Dim cnt As Long
  Dim RecNumb As String, SchlSeq As String
  Dim MeterSlot  As Integer, WhatYear As Integer
  Dim CurReading As Double, Multi As Double
  Dim UCode1 As String, UCode2 As String
  Dim RYear As String, ReadYear As String
  Dim DateRead As String
  Dim ReadDate As Integer
  Dim c1 As String
  Dim s1 As String
  Dim BookSeq As String
  Dim BookCnt As Integer, Rptcnt As Long
  Dim NumTRGetRecs As Integer
  Dim HarryFlag As Boolean
'  ExportFileName$ = "C:\ezroute\HOST2PC.IMP"
'  ImportFileName$ = "C:\ezroute\PC2HOST.EXP"
  Rptcnt = 0
  ExportFileName$ = HHPathInOut + "HOST2PC.IMP"
  ImportFileName$ = HHPathInOut + "PC2HOST.EXP"
  
  CrLf$ = Chr$(13) + Chr$(10)
  FF$ = Chr$(12)
  q$ = Chr$(34)
  cb$ = Space$(45)

  LoadUBSetUpFile UBSetUpRec(), UBSetupLen
  If InStr(TOWNNAME$, "HARRISBURG") Then
    If InStr(UBSetUpRec(1).DEFSTATE, "NC") Then
      HarryFlag = True
    Else
      HarryFlag = False
    End If
  ElseIf InStr(TOWNNAME$, "STEPHENS CITY") Then
     HarryFlag = True
  End If
  
  HighVar = UBSetUpRec(1).HighRead
  LowVar = UBSetUpRec(1).LowRead
  
  If ImpExpFlag Then     'EXPORTING METER READINGS
    WhatTypes$ = Left$(Me.fpMtrType.Text, 1)
    GoSub ExportSchlum
    Call cmdExit_Click
  Else
    GoSub ImportSchlum
    Call cmdExit_Click
  End If
  
Exit Sub

ExportSchlum:
  
  If CycleCnt > 0 Then
    Call KillFile(ExportFileName$)
    
    UBCustRecLen = Len(UBCustRec(1))
    UBSchlumHHRecLen = Len(UBSchlumHHRec(1))
    
    UBFile = FreeFile
    Open UBPath + "UBCUST.DAT" For Random Shared As UBFile Len = UBCustRecLen
    
    UBSchlFile = FreeFile
    Open ExportFileName$ For Random Shared As UBSchlFile Len = UBSchlumHHRecLen
    
    'Open the Correct Order for Reading
    If UBSetUpRec(1).UseSeq = "Y" Then
      IdxRecLen = 4         'we are using a integer
      MakeSequenceIndex "Sequence Number", Me
      IdxNumOfRecs = FileSize&("UBTEMP.IDX") \ 4
      ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
      FGetAH "UBTEMP.IDX", IdxBuff(), 4, IdxNumOfRecs
    Else
      IdxRecLen = 4         'we are using a integer
      IdxFileSize& = FileSize&("UBCUSTBK.IDX")
      IdxNumOfRecs = IdxFileSize& \ IdxRecLen
      ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
      FGetAH "UBCUSTBK.IDX", IdxBuff(), IdxRecLen, IdxNumOfRecs            'load it
    End If

'*************************
  
    FrmShowPctComp.Label1 = "Exporting Meter Reading Information."
    FrmShowPctComp.cmdCancel.Enabled = False
    FrmShowPctComp.Show '1, Parent
  
    WriteCnt& = 0
    PageCnt& = 0
    RMCnt = 1
    
    Do
      
      WhatRMRec& = IdxBuff(RMCnt).RecNum
'      RMCnt = IdxBuff(cnt!).RecNum
      If WhatRMRec& > 0 Then
        Get UBFile, WhatRMRec&, UBCustRec(1)
        If InStr(UBCustRec(1).HHMSG1, "NOREAD") > 0 Then
          GoTo HWriteSkip
        End If
        For BookCnt = 1 To CycleCnt
          If Val(UBCustRec(1).Book) = Cycle(BookCnt) And (UBCustRec(1).Status <> "F") Then
            GoSub SchlumWriteRec
            Exit For
          End If
        Next
      End If

HWriteSkip:
      
      FrmShowPctComp.ShowPctComp RMCnt&, IdxNumOfRecs
      RMCnt = RMCnt + 1
    Loop Until RMCnt > IdxNumOfRecs
    Close
    DoEvents
    
    MsgText(0) = "Schlumberger Export Operation."
    MsgText(1) = ""
    MsgText(2) = "FILE: " + ExportFileName$
    MsgText(3) = ""
    MsgText(4) = "Completed."
    MsgText(5) = ""
    GetOKorNot% MsgText(), True, True
  End If
 
Return

ImportSchlum:
  MsgText(0) = "Import Schlumberger Reading File."
  MsgText(1) = "Make sure you have Exported current readings"
  MsgText(2) = "using the Schlumberger reading export utility."
  MsgText(3) = "The file 'PC2HOST.EXP' must be in the specified"
  MsgText(4) = "directory."
  MsgText(5) = ""
  
  Select Case GetOKorNot%(MsgText(), False, True)
  Case Not True
    GoTo SchlumbergerGetExit
  End Select
  
  ReportFile$ = "IMPREAD.RPT"
  MaxLines = 55
  
  ReDim UBSchlumHHRec(1) As SchlumHHType
  UBSchlumHHRecLen = Len(UBSchlumHHRec(1))
  
  UBSchlFile = FreeFile
  Open ImportFileName$ For Random Shared As UBSchlFile Len = UBSchlumHHRecLen
  
  NumTRGetRecs = LOF(UBSchlFile) / UBSchlumHHRecLen
  
  If NumTRGetRecs = 0 Then
    Close
    MsgText(0) = "ERROR:"
    MsgText(1) = "IMPORT FILE NOT FOUND"
    MsgText(2) = "Make sure that: 'PC2HOST.EXP'"
    MsgText(3) = "is in the specified directory!"
    MsgText(4) = "Please call Southern Software for"
    MsgText(5) = "additional Information."
    GetOKorNot% MsgText(), True, False
    GoTo SchlumbergerGetExit
  End If
  
  FrmShowPctComp.Label1 = "Importing Meter Reading Information."
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show '1, Parent
  
  UBCustRecLen = Len(UBCustRec(1))
  
  UBFile = FreeFile
  Open UBPath + "UBCUST.DAT" For Random Shared As UBFile Len = UBCustRecLen
  
  RptHandle = FreeFile
  Open ReportFile$ For Output As RptHandle
  
  GoSub ImpRptHeader
  
  For cnt& = 1 To NumTRGetRecs
    Get UBSchlFile, cnt&, UBSchlumHHRec(1)
    
    RecNumb$ = QPTrim$(UBSchlumHHRec(1).UBAcctNo)
    RMCnt = Val(RecNumb$)

    SchlSeq$ = QPTrim$(Left$(UBSchlumHHRec(1).Notes8, 8))
    If Len(SchlSeq$) = 0 Then
      GoTo BadSkip
    End If
    
    If RMCnt > 0 Then
      Get UBFile, RMCnt, UBCustRec(1)
      GoSub SchlumExtractRecord
      If LineCnt > MaxLines Then
        Print #RptHandle, FF$
        GoSub ImpRptHeader
      End If
    Else
      
    End If
BadSkip:
    FrmShowPctComp.ShowPctComp cnt&, NumTRGetRecs
  Next
  
  GoSub ImpRptTotal
  Print #RptHandle, FF$
  
  Close
  
  MsgText(0) = "Import Readings."
  MsgText(1) = ""
  MsgText(2) = "Readings Updated Successfully."
  MsgText(3) = ""
  MsgText(4) = " IMPORTED: " + Str$(Rptcnt&) + " Readings"
  MsgText(5) = ""
  
  GetOKorNot% MsgText(), True, True
  frmReportOpt.Show 1
  If rptopt = 1 Then
    'do the graphics
    Load frmLoadingRpt
    frmLoadingRpt.setwherefrom frmUBImpExpHHRead
    ARptLineRpt.GetName ReportFile$
    ARptLineRpt.startrpt
  ElseIf rptopt = 2 Then
    'do the text
    ViewPrint ReportFile$, "Imported Meter Readings"
  End If

SchlumbergerGetExit:

Return

ImpRptHeader:
  Page = Page + 1
  Print #RptHandle,
  Print #RptHandle, "Date: "; Date$; Tab(27); "Imported Meter Reading Report"; Tab(70); "Page:"; Page
  Print #RptHandle, "Location Account  Customer Name               Reading       ReadDate     Status"
  Print #RptHandle, String$(80, "-")
  LineCnt = 4
Return

ImpRptTotal:
  Print #RptHandle, String$(80, "-")
  Print #RptHandle, "Imported Count: " + Using("######", Rptcnt&, False)
Return

SchlumExtractRecord:
  BadDate = False
  
  MeterSlot = Val(QPTrim$(Mid$(UBSchlumHHRec(1).Notes8, 9, 2)))
  
  If Len(QPTrim$(UBSchlumHHRec(1).MtrRead)) = 0 Then
    GoTo NoUpdate
  End If
  Rptcnt = Rptcnt + 1
  CurReading# = Val(QPTrim$(UBSchlumHHRec(1).MtrRead))
  
  If CurReading# < 0 Then CurReading# = 0

  'UCode1$ = QPTrim$(UBCustRec(1).UserCode1)
  UCode2$ = QPTrim$(UBCustRec(1).USERCODE2)
  
  If Len(UCode2$) = 0 Then
    Multi# = 1
  Else
    Multi# = 10 'UBCustRec(1).LocMeters(MeterSlot).MTRMulti
  End If


  UCode1 = Val(QPTrim$(UBCustRec(1).USERCODE1))
  Select Case UCode1
  Case 0
    If Middle Then   'this is for middleton and Bicoe and Summerton REMINGTON
      Multi# = 1
    End If
  Case 1
    Multi# = 10
  Case 2
    Multi# = 100
  Case 3
    Multi# = 1000
  End Select

'Meansville 062519
  If InStr(TOWNNAME$, "MEANS") Then
    Multi# = 1
  End If
  
  'IF Multi# = 0 THEN Multi# = 1
  If Not HarryFlag Then
    CurReading# = CurReading# * Multi#
  End If
  RYear$ = QPTrim$(Right$(UBSchlumHHRec(1).ReadDate, 2))
  
  If Len(RYear$) < 2 Then
    BadDate = True
  End If
  
  WhatYear = Val(RYear$)
  If WhatYear < 65 Then
    ReadYear$ = "-20" + RYear$
  Else
    ReadYear$ = "-19" + RYear$
  End If
  DateRead$ = Left$(UBSchlumHHRec(1).ReadDate, 2) + "-" + Mid$(UBSchlumHHRec(1).ReadDate, 3, 2) + ReadYear$
  ReadDate = Date2Num(DateRead$)
  
  If CurReading# >= 9999999999# Then
    CurReading# = 999999999#
  End If
  
  If UBCustRec(1).LocMeters(MeterSlot).ReadFlag = "Y" Then
    UBCustRec(1).LocMeters(MeterSlot).CurRead = CurReading#
    UBCustRec(1).LocMeters(MeterSlot).CurDate = ReadDate
  Else
    UBCustRec(1).LocMeters(MeterSlot).PrevRead = UBCustRec(1).LocMeters(MeterSlot).CurRead
    UBCustRec(1).LocMeters(MeterSlot).PastDate = UBCustRec(1).LocMeters(MeterSlot).CurDate
    UBCustRec(1).LocMeters(MeterSlot).ReadFlag = "Y"
    UBCustRec(1).LocMeters(MeterSlot).CurDate = ReadDate
    UBCustRec(1).LocMeters(MeterSlot).CurRead = CurReading#
  End If
  If schlnglat = True Then
    UBCustRec(1).LocMeters(MeterSlot).MtrLat = Val(Mid$(QPTrim$(UBSchlumHHRec(1).Notes4), 5, (Len(QPTrim$(UBSchlumHHRec(1).Notes4)) - 4)))
    UBCustRec(1).LocMeters(MeterSlot).MtrLng = Val(Mid$(QPTrim$(UBSchlumHHRec(1).Notes5), 6, (Len(QPTrim$(UBSchlumHHRec(1).Notes5)) - 5)))
  End If
  
  c1$ = QPTrim$(UBCustRec(1).HHMSG1)
  s1$ = QPTrim$(UBSchlumHHRec(1).Notes1)
  If Len(s1$) > 0 Then
    If s1$ <> c1$ Then
      UBCustRec(1).NewNotes = True
      UBCustRec(1).HHMSG1 = s1$
    End If
  End If
  c1$ = QPTrim$(UBCustRec(1).HHMSG2)
  s1$ = QPTrim$(UBSchlumHHRec(1).Notes2)
  If Len(s1$) > 0 Then
    If s1$ <> c1$ Then
      UBCustRec(1).NewNotes = True
      UBCustRec(1).HHMSG2 = s1$
    End If
  End If
  c1$ = QPTrim$(UBCustRec(1).HHMSG3)
  s1$ = QPTrim$(UBSchlumHHRec(1).Notes3)
  If Len(s1$) > 0 Then
    If s1$ <> c1$ Then
      UBCustRec(1).NewNotes = True
      UBCustRec(1).HHMSG3 = s1$
    End If
  End If
  
  BookSeq$ = Left$(UBSchlumHHRec(1).Notes8, 2) + "-" + Mid$(UBSchlumHHRec(1).Notes8, 3, 6)
  Print #RptHandle, BookSeq$; Tab(10); Using("######", RMCnt, False); Tab(19);
  Print #RptHandle, UBSchlumHHRec(1).HHDisp4; Tab(45); Using("#########", CurReading#, False);
  Print #RptHandle, Tab(60); DateRead$; Tab(74); UBSchlumHHRec(1).ReadType
  LineCnt = LineCnt + 1
  Put UBFile, RMCnt, UBCustRec(1)
NoUpdate:
Return


SchlumWriteRec:  'May Have Up to 7 Meters to Read
  For MtrCnt = 1 To 7
    MeterOK = False
    Account$ = Str$(WhatRMRec&)
    Account$ = Left$(Account$, 6) + "-" + Right$(Str$(MtrCnt), 1)
    MTRType$ = QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MTRType)
    If Len(MTRType$) > 0 Then
      If MTRType$ = "W" Or MTRType$ = "S" Or MTRType$ = "C" Or MTRType$ = "E" Or MTRType$ = "D" Or MTRType$ = "G" Or MTRType$ = "T" Or MTRType$ = "I" Then
        Select Case WhatTypes$
        Case "W", "T"               'water/sewer
          If MTRType$ = "W" Or MTRType$ = "S" Or MTRType$ = "C" Or MTRType$ = "T" Or MTRType$ = "I" Then
            MeterOK = True
          End If
        Case "E"                'electric & demand elec.
          If MTRType$ = "E" Or MTRType$ = "D" Then
            MeterOK = True
          End If
        Case "G"                'gas
          If MTRType$ = "G" Then
            MeterOK = True
          End If
        Case "A", " "           'all meters
          MeterOK = True
        End Select
        
        If MeterOK = True Then  ' Determine High and Low Reading
          HasZ = False
          SRouteID$ = UBCustRec(1).Book + UBCustRec(1).SEQNUMB + "0" + QPTrim$(Str$(MtrCnt))

          LSet UBSchlumHHRec(1).Route = QPTrim$(UBCustRec(1).Book)
          PageCnt& = PageCnt& + 1
          WriteCnt& = WriteCnt& + 1
          
          WalkSeq$ = "0000" + QPTrim$(Str$(WriteCnt&))
          PageNum$ = "0000" + QPTrim$(Str$(PageCnt&))

          LSet UBSchlumHHRec(1).WalkSeq$ = ""
          
          UBSchlumHHRec(1).PageNum = Right$(PageNum$, 4)
          UBSchlumHHRec(1).ReadSeq = "01"       'UBSchlumHHRec(1).WalkSeq$
          UBSchlumHHRec(1).HHID = ""            'HH Number 'no information
          UBSchlumHHRec(1).ReadDir = "L"
          Dials = Val(UBCustRec(1).USERCODE2)
          'IF Dials = 0 THEN Dials = 7             'Default for Caldwell Cty
          UBSchlumHHRec(1).NumDial = QPTrim$(Str$(Dials))
          
          MeterID$ = UCase$(QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrNum))
          
          RecStat$ = Right$(MeterID$, 1)
          Select Case RecStat$
          Case "Z", "A", "M"
            HasZ = True
            MeterID$ = Mid$(MeterID$, 1, (Len(MeterID$) - 1))
          Case Else
          End Select
          
          LSet UBSchlumHHRec(1).IDExpected = MeterID$
          LSet UBSchlumHHRec(1).IDCaptured = ""
          LSet UBSchlumHHRec(1).IDOverride = ""
          LSet UBSchlumHHRec(1).Decimals = ""
          LSet UBSchlumHHRec(1).MtrRead = ""    'current from HH
          LSet UBSchlumHHRec(1).ReadOVRide = ""
          
          Average# = UBCustRec(1).LocMeters(MtrCnt).AvgUse
          ReadLowI$ = QPTrim$(Str$((UBCustRec(1).LocMeters(MtrCnt).CurRead + 1)))
          PrevRead$ = QPTrim$(Str$(UBCustRec(1).LocMeters(MtrCnt).PrevRead))
          NCurRead$ = QPTrim$(Str$(UBCustRec(1).LocMeters(MtrCnt).CurRead))

          PrevDate$ = Num2Date(UBCustRec(1).LocMeters(MtrCnt).CurDate)
          PrevDate$ = Left$(PrevDate$, 2) + Mid$(PrevDate$, 4, 2) + Right$(PrevDate$, 2)
          If LowRead# < 0 Then LowRead# = 0
          LowRead# = Fix(ILowRead#)
          
          HighRead# = Fix((Average# * (HighVar / 100) - Average#) + UBCustRec(1).LocMeters(MtrCnt).CurRead)
          If HighRead# < 0 Then HighRead# = 0
          HiRead$ = QPTrim$(Str$(HighRead#))
          
         
          LSet UBSchlumHHRec(1).HighLimit = HiRead$
          LSet UBSchlumHHRec(1).LowLimit = ReadLowI$
          LSet UBSchlumHHRec(1).Date2Read = ""
          LSet UBSchlumHHRec(1).Date2Exp = ""
          LSet UBSchlumHHRec(1).NoteCodes = ""
          LSet UBSchlumHHRec(1).LocatCode = ""
          LSet UBSchlumHHRec(1).MtrRCode = ""
          LSet UBSchlumHHRec(1).RecType = "EU"
          If HasZ Then
            LSet UBSchlumHHRec(1).RecStatus = RecStat$
          Else
            LSet UBSchlumHHRec(1).RecStatus = ""
          End If
          LSet UBSchlumHHRec(1).ReadDate = ""
          LSet UBSchlumHHRec(1).ReadTime = ""
          LSet UBSchlumHHRec(1).ReadType = ""
          LSet UBSchlumHHRec(1).NetNumb = ""
          LSet UBSchlumHHRec(1).ReadAtmpt = ""
          LSet UBSchlumHHRec(1).UserChar = ""
          LSet UBSchlumHHRec(1).HHManufac = ""
          LSet UBSchlumHHRec(1).ActStatus = UBCustRec(1).Status
          LSet UBSchlumHHRec(1).MTRType = ""
          LSet UBSchlumHHRec(1).ReadFailCode = ""
          LSet UBSchlumHHRec(1).PrevRead = NCurRead$
          LSet UBSchlumHHRec(1).PrevDate = PrevDate$
          LSet UBSchlumHHRec(1).HHDisp1 = QPTrim$(UBCustRec(1).ServAddr)
          LSet UBSchlumHHRec(1).HHDisp2 = QPTrim$(UBCustRec(1).CustName)
          LSet UBSchlumHHRec(1).HHDisp3 = QPTrim$(UBCustRec(1).HHMSG1)
          LSet UBSchlumHHRec(1).HHDisp4 = MeterID$
          
          LSet UBSchlumHHRec(1).Notes1 = QPTrim$(UBCustRec(1).HHMSG1)
          LSet UBSchlumHHRec(1).Notes2 = QPTrim$(UBCustRec(1).HHMSG2)
          LSet UBSchlumHHRec(1).Notes3 = QPTrim$(UBCustRec(1).HHMSG3)
        If schlnglat = True Then
          LSet UBSchlumHHRec(1).Notes4 = "Lat=" + QPTrim$(Str((UBCustRec(1).LocMeters(MtrCnt).MtrLat)))
          LSet UBSchlumHHRec(1).Notes5 = "Long=" + QPTrim$(Str((UBCustRec(1).LocMeters(MtrCnt).MtrLng)))
        Else
          LSet UBSchlumHHRec(1).Notes4 = ""
          LSet UBSchlumHHRec(1).Notes5 = ""
        End If
          LSet UBSchlumHHRec(1).Notes6 = ""
          LSet UBSchlumHHRec(1).Notes7 = ""
          LSet UBSchlumHHRec(1).Notes8 = SRouteID$
          LSet UBSchlumHHRec(1).OpCode = ""
          LSet UBSchlumHHRec(1).UBAcctNo = QPTrim$(Str$(WhatRMRec&))
          LSet UBSchlumHHRec(1).MtrSlot = QPTrim$(Str$(MtrCnt))
          LSet UBSchlumHHRec(1).UtilFld = ""
          LSet UBSchlumHHRec(1).CrLf = CrLf$
          Put UBSchlFile, (LOF(UBSchlFile) / UBSchlumHHRecLen) + 1, UBSchlumHHRec(1)
        End If
      End If
    End If
  Next 'meter

Return

End Sub 'END OF SCHULM

Private Sub ImpExpIMecHHInfo(ByVal ImpExpFlag As Boolean)

  Dim CEReadFile As String
  Dim HighVar As Integer, LowVar As Integer
  Dim UBFile As Integer, HHFile As Integer
  Dim UBInterRDRecLen As Integer
  Dim RMCnt As Long, WhatRMRec As Long
  Dim Ok2DoIt As Boolean
  Dim IdxFileName As String
  
  Dim BookCnt As Integer
  Dim MtrCnt As Integer
  Dim MeterOK As Boolean
  Dim Account As String, MTRType As String
  Dim WhatTypes As String, MeterID As String
  Dim Average As Double, LowRead As Double, HighRead As Double
  
'  ReDim MsgText(0 To 5) As String
  ReDim UBInterRDRec(1) As UBIntermecHHRecType
  Dim C As String, ThisDate As String
  Dim ReadingDate As Integer
  Dim MMsg1 As String, MMsg2 As String, MMsg3 As String
  Dim CMsg1 As String, CMsg2 As String, CMsg3 As String
  Dim NumIntermecRdRecs As Long
  Dim MeterRecord As Long
  Dim CurReading As Double
  
  CEReadFile$ = UBPath$ + "CEMTREAD.DAT"
  
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen

  HighVar = UBSetUpRec(1).HighRead
  LowVar = UBSetUpRec(1).LowRead
  
  If ImpExpFlag Then     'EXPORTING METER READINGS
    WhatTypes$ = Left$(Me.fpMtrType.Text, 1)
    GoSub ExportIntermec
    Call cmdExit_Click
  Else
    GoSub ImportIntermec
    Call cmdExit_Click
  End If

Exit Sub

ExportIntermec:

  If CycleCnt > 0 Then
    
    Call KillFile(CEReadFile$)
    
    UBInterRDRec(1).CEVariant = Chr$(8) + Chr$(0)
    UBInterRDRec(1).CEStrLen = Chr$(165) + Chr$(0)
    UBInterRDRecLen = Len(UBInterRDRec(1))
    
    UBCustRecLen = Len(UBCustRec(1))
    UBFile = FreeFile               'Open Customer Data File
    Open UBPath$ + "UBCUST.DAT" For Random Shared As UBFile Len = UBCustRecLen

    HHFile = FreeFile
    Open CEReadFile$ For Random Shared As HHFile Len = UBInterRDRecLen
        
    If UBSetUpRec(1).UseSeq = "Y" Then
      MakeSequenceIndex "Sequence Number", Me
      IdxFileName = UBPath$ + "UBTEMP.IDX"
    Else
      IdxFileName = UBPath$ + "UBCUSTBK.IDX"
    End If
    
    IdxRecLen = 4         'we are using a integer
    IdxFileSize& = FileSize&(IdxFileName)
    IdxNumOfRecs = IdxFileSize& \ IdxRecLen
    
    ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
    FGetAH IdxFileName, IdxBuff(), IdxRecLen, IdxNumOfRecs            'load it
    'Open the Correct Order for Reading
  End If
  
  FrmShowPctComp.Label1 = "Exporting Reading Information."
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show '1, Parent
  
  RMCnt = 1

  Do
    WhatRMRec& = IdxBuff(RMCnt).RecNum
    If Not (WhatRMRec&) = 0 Then
      Get UBFile, WhatRMRec&, UBCustRec(1)
      If InStr(UBCustRec(1).HHMSG1, "NOREAD") > 0 Then
        GoTo IWriteSkip
      End If
      For BookCnt = 1 To CycleCnt
        If Val(UBCustRec(1).Book) = Cycle(BookCnt) And (UBCustRec(1).Status <> "F") Then
          GoSub IntermecWriteRec
          Exit For
        End If
      Next
    End If
IWriteSkip:
    FrmShowPctComp.ShowPctComp RMCnt&, IdxNumOfRecs
    RMCnt = RMCnt + 1
  Loop Until RMCnt > IdxNumOfRecs
  
  Close  'Close files

  MsgText(0) = "Intermec Export Operation."
  MsgText(1) = ""
  MsgText(2) = "Reading file 'CEMTREAD.DAT' completed"
  MsgText(3) = ""
  MsgText(4) = ""
  MsgText(5) = ""
  
  GetOKorNot% MsgText(), True, True

  Return
  
ImportIntermec:

  MsgText(0) = "Import Intermec Reading File."
  MsgText(1) = ""
  MsgText(2) = "Make sure the file 'CEMTREAD.DAT' is in the"
  MsgText(3) = "CITIPAK folder."
  MsgText(4) = "Ready to Proceed?"
  MsgText(5) = ""
  
  Select Case GetOKorNot%(MsgText(), False, True)
  Case Not True
    GoTo IntermecGetExit
  End Select

  UBCustRecLen = Len(UBCustRec(1))
  UBFile = FreeFile
  Open UBPath + "UBCUST.DAT" For Random Shared As UBFile Len = UBCustRecLen

  UBInterRDRecLen = Len(UBInterRDRec(1))

  HHFile = FreeFile
  Open CEReadFile$ For Random Shared As HHFile Len = UBInterRDRecLen

  NumIntermecRdRecs = LOF(HHFile) / UBInterRDRecLen

  If NumIntermecRdRecs = 0 Then
    Close
    MsgText(0) = "ERROR:"
    MsgText(1) = ""
    MsgText(2) = "The file 'CEMTREAD.DAT' must be"
    MsgText(3) = "in the Citipak folder."
    MsgText(4) = "Please call Southern Software for"
    MsgText(5) = "additional Information."
    GetOKorNot% MsgText(), True, False
    GoTo IntermecGetExit:
  End If

  FrmShowPctComp.Label1 = "Importing Meter Reading Information."
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show '1, Parent
  
  For RMCnt& = 1 To NumIntermecRdRecs
    Get HHFile, RMCnt, UBInterRDRec(1)
    WhatRMRec& = Val(QPTrim$(UBInterRDRec(1).Account))
    If Not (WhatRMRec&) = 0 Then
      Get UBFile, WhatRMRec&, UBCustRec(1)
      GoSub IntermecExtractRec
    End If
    FrmShowPctComp.ShowPctComp RMCnt&, NumIntermecRdRecs
  Next
  Close
  
  MsgText(0) = "Import Readings."
  MsgText(1) = "Readings Updated Successfully."
  MsgText(2) = ""
  MsgText(3) = " IMPORTED: " + Str$(RMCnt&) + " Readings"
  MsgText(4) = ""
  MsgText(5) = ""
  GetOKorNot% MsgText(), True, True
  

IntermecGetExit:
  Return

IntermecExtractRec:

  If UBInterRDRec(1).ReadFlag = "Y" Then
    MeterRecord = Val(Right$((QPTrim$(UBInterRDRec(1).Account)), 1))

' Check Meter Updated Flag
' Update Meter W/Reading
''NOTE: New current reading stored in the
' pastreading field from CEMTREAD.DAT
    
    CurReading# = Val(UBInterRDRec(1).PastRead)
    ThisDate$ = Left$(UBInterRDRec(1).ReadDate, 2) + "/" + Mid$(UBInterRDRec(1).ReadDate, 3, 2) + "/" + Right$(UBInterRDRec(1).ReadDate, 4)
    ReadingDate = Date2Num(ThisDate$)

    If UBCustRec(1).LocMeters(MeterRecord).ReadFlag = "Y" Then
      UBCustRec(1).LocMeters(MeterRecord).CurRead = CurReading#
      UBCustRec(1).LocMeters(MeterRecord).CurDate = ReadingDate
    Else
      UBCustRec(1).LocMeters(MeterRecord).PrevRead = UBCustRec(1).LocMeters(MeterRecord).CurRead
      UBCustRec(1).LocMeters(MeterRecord).PastDate = UBCustRec(1).LocMeters(MeterRecord).CurDate
      UBCustRec(1).LocMeters(MeterRecord).ReadFlag = "Y"
      UBCustRec(1).LocMeters(MeterRecord).CurDate = ReadingDate
      UBCustRec(1).LocMeters(MeterRecord).CurRead = CurReading#
    End If

    MMsg1$ = QPTrim$(UBInterRDRec(1).Note1)
    CMsg1$ = QPTrim$(UBCustRec(1).HHMSG1)
    MMsg2$ = QPTrim$(UBInterRDRec(1).Note2)
    CMsg2$ = QPTrim$(UBCustRec(1).HHMSG2)
    MMsg3$ = QPTrim$(UBInterRDRec(1).Note3)
    CMsg3$ = QPTrim$(UBCustRec(1).HHMSG3)
    
    If MMsg1$ <> CMsg1$ Then
      GoSub UpDateNoteInfo
      GoTo DoneINNotes
    End If
    If MMsg2$ <> CMsg2$ Then
      GoSub UpDateNoteInfo
      GoTo DoneINNotes
    End If
    If MMsg3$ <> CMsg3$ Then
      GoSub UpDateNoteInfo
    End If

DoneINNotes:
    Put UBFile, WhatRMRec&, UBCustRec(1)
  End If
Return

UpDateNoteInfo:
  UBCustRec(1).NewNotes = True
  UBCustRec(1).HHMSG1 = MMsg1$
  UBCustRec(1).HHMSG2 = MMsg2$
  UBCustRec(1).HHMSG3 = MMsg3$
Return

IntermecWriteRec:  'May Have Up to 7 Meters to Read
  For MtrCnt = 1 To 7
    MeterOK = False
    Account$ = Str$(WhatRMRec&)
    Account$ = Left$(Account$, 6) + "-" + Right$(Str$(MtrCnt), 1)
      
    MTRType$ = UBCustRec(1).LocMeters(MtrCnt).MTRType
    If MTRType$ = "W" Or MTRType$ = "S" Or MTRType$ = "C" Or MTRType$ = "E" Or MTRType$ = "D" Or MTRType$ = "G" Then
      Select Case WhatTypes$
      Case "W"                'water/sewer
        If MTRType$ = "W" Or MTRType$ = "S" Or MTRType$ = "C" Then
          MeterOK = True
        End If
      Case "E"                'electric & demand elec.
        If MTRType$ = "E" Or MTRType$ = "D" Then
          MeterOK = True
        End If
      Case "G"                'gas
        If MTRType$ = "G" Then
          MeterOK = True
        End If
      Case "A", " "           'all meters
        MeterOK = True
      End Select

      If MeterOK = True Then
'          ' Determine High and Low Reading
        Average = UBCustRec(1).LocMeters(MtrCnt).AvgUse
        LowRead# = Fix(LowRead#)
        HighRead# = Fix(Average# * (HighVar / 100)) + UBCustRec(1).LocMeters(MtrCnt).CurRead

        MeterID$ = LTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrNum)
        MeterID$ = RTrim$(MeterID$)

        If Val(MeterID$) = 0 Then
          MeterID$ = UBCustRec(1).Book + UBCustRec(1).SEQNUMB
        End If
        If Len(MeterID$) < 8 Then
          MeterID$ = String$(8 - Len(MeterID$), "0") + MeterID$
        End If
        MeterID$ = Left$(MeterID$, 8)
'          'Set Record Fields and Put On Disk
        UBInterRDRec(1).CustName = UBCustRec(1).CustName
        UBInterRDRec(1).ServAddress = Left$(UBCustRec(1).ServAddr, 16)
        UBInterRDRec(1).ReadDate = ""
        C$ = QPTrim$(UBCustRec(1).USERCODE1)
        If Len(C$) > 0 Then
          Mid$(UBInterRDRec(1).ServAddress, 19, 1) = Left$(C$, 1)
        End If
        C$ = QPTrim$(UBCustRec(1).USERCODE2)
        If Len(C$) > 0 Then
          Mid$(UBInterRDRec(1).ServAddress, 20, 1) = Left$(C$, 1)
        End If

        UBInterRDRec(1).MeterID = MeterID$
        UBInterRDRec(1).LowRead = QPTrim$(Str$(LowRead#))
        UBInterRDRec(1).HighRead = QPTrim$(Str$(HighRead#))
        UBInterRDRec(1).Account = Account$
        UBInterRDRec(1).MeterType$ = UBCustRec(1).LocMeters(MtrCnt).MTRType
        UBInterRDRec(1).Book = UBCustRec(1).Book
        UBInterRDRec(1).CurRead = QPTrim$(Str$(UBCustRec(1).LocMeters(MtrCnt).CurRead))
        UBInterRDRec(1).PastRead = "0"
        UBInterRDRec(1).ReadFlag = "N"

        UBInterRDRec(1).Note1 = QPTrim$(UBCustRec(1).HHMSG1)
        UBInterRDRec(1).Note2 = QPTrim$(UBCustRec(1).HHMSG2)
        UBInterRDRec(1).Note3 = QPTrim$(UBCustRec(1).HHMSG3)
        '         ^^^
        UBInterRDRec(1).NoteStatus = ""
        Put HHFile, (LOF(HHFile) / UBInterRDRecLen) + 1, UBInterRDRec(1)
      End If
    End If
  Next
  
Return

End Sub

Private Sub ImpExpHuskyHHInfo(ByVal ImpExpFlag As Boolean)

  Dim UBFile As Integer, UBPC3000RdFile As Integer, UBPC3000GetRdFile As Integer
  Dim UBPC3000RDRec(1) As UBPC3000ReadRecType
  Dim UBPC3000RdRecLen As Integer, UBPC3000GetRdRecLen As Integer
  Dim RMCnt As Long, WhatRMRec As Long
  Dim UBPC3000GetRDRec(1) As UBPC3000ReadRecType
  Dim IdxFileName As String
  Dim BookCnt As Integer, MtrCnt As Integer
  Dim MeterOK As Boolean
  Dim Account As String, MTRType As String
  Dim WhatTypes As String
  Dim HighVar As Integer
  Dim Average As Double, LowRead As Double, HighRead As Double
  Dim ILowRead As String, MeterID As String
  Dim C As String
  Dim UpdCnt As Long
  Dim CustMTRRec As Integer
  Dim CurReading As Double
  
  UBPC3000RdRecLen = Len(UBPC3000RDRec(1))
  
  If ImpExpFlag Then     'EXPORTING METER READINGS
    WhatTypes$ = Left$(Me.fpMtrType.Text, 1)
    GoSub SendHusky
    Call cmdExit_Click
  Else
    GoSub GetHusky
    Call cmdExit_Click
  End If
  
Exit Sub
  
SendHusky:
  If CycleCnt > 0 Then
    
    MsgText(0) = "WARNING. . ."
    MsgText(1) = "Make sure the HUSKY is connected and ready to"
    MsgText(2) = "transfer files. The HUSKY should be at the"
    MsgText(3) = "'Husky File Transfer Utility' screen. Type 'H'"
    MsgText(4) = "and press the 'Yes' key at the C: prompt"
    MsgText(5) = "IS THE HUSKY READY TO PROCEED?"
  
    DoEvents
  
    Select Case GetOKorNot%(MsgText(), False, True)
    Case Not True
      GoTo SendHuskyOKExit
    End Select

    LoadUBSetUpFile UBSetUpRec(), UBSetupLen
    HighVar = UBSetUpRec(1).HighRead
    If UBSetUpRec(1).UseSeq = "Y" Then
      MakeSequenceIndex "Sequence Number", Me
      IdxFileName = UBPath$ + "UBTEMP.IDX"
    Else
      IdxFileName = UBPath$ + "UBCUSTBK.IDX"
    End If
    IdxRecLen = 4         'we are using a integer
    IdxFileSize& = FileSize&(IdxFileName)
    IdxNumOfRecs = IdxFileSize& \ IdxRecLen
    ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
    FGetAH IdxFileName, IdxBuff(), IdxRecLen, IdxNumOfRecs            'load it
  End If
  
  FrmShowPctComp.Label1 = "Exporting Reading Information."
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show '1, Parent
  
  Call KillFile(UBPath$ + "UBCUSTTR.DAT") 'kill old hh reading file
  UBCustRecLen = Len(UBCustRec(1))
  
  UBPC3000RdFile = FreeFile
  Open UBPath$ + "UBCUSTTR.DAT" For Random Shared As UBPC3000RdFile Len = UBPC3000RdRecLen
  
  UBFile = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBFile Len = UBCustRecLen
  
  RMCnt& = 1
  
  Do
    WhatRMRec& = IdxBuff(RMCnt&).RecNum
    If Not (WhatRMRec&) = 0 Then
      Get UBFile, WhatRMRec&, UBCustRec(1)
      If InStr(UBCustRec(1).HHMSG1, "NOREAD") > 0 Then
        GoTo HWriteSkip
      End If
      For BookCnt = 1 To CycleCnt
        If Val(UBCustRec(1).Book) = Cycle(BookCnt) And (UBCustRec(1).Status <> "F") Then
          GoSub HuskyWriteRec
        End If
      Next
    End If
HWriteSkip:
    RMCnt& = RMCnt& + 1
    FrmShowPctComp.ShowPctComp RMCnt&, IdxNumOfRecs
  Loop Until RMCnt& > IdxNumOfRecs
    
  Close
  
  MsgText(0) = "Husky File Transfer"
  MsgText(1) = "WARNING. . ."
  MsgText(2) = "WAIT UNTIL THE HUSKY FILE"
  MsgText(3) = "TRANNSFER UTILITY HAS COMPLETED"
  MsgText(4) = "Then Click 'OK' to Continue."
  
  DoEvents
  
  Shell "HCOMW32 " + HuskyPort + " /tx=ubcusttr.dat /abort", vbNormalFocus
    
  OkORNotFlag = GetOKorNotHH%(MsgText())
  
  DoEvents
  Call Chk4HuskyError%

SendHuskyOKExit:

Return

Exit Sub ' this should never happen

HuskyWriteRec:
'  'May Have Up to 7 Meters to Read
  For MtrCnt = 1 To 7
    MeterOK = False
    Account$ = Str$(WhatRMRec&)
    Account$ = Left$(Account$, 6) + "-" + Right$(Str$(MtrCnt), 1)
    MTRType$ = Left$(UBCustRec(1).LocMeters(MtrCnt).MTRType, 1)
    If Len(QPTrim$(MTRType$)) > 0 Then
      Select Case MTRType$
      Case "W", "S", "C", "E", "D", "G"
        Select Case WhatTypes$
        Case "W"                'water/sewer
          If MTRType$ = "W" Or MTRType$ = "S" Or MTRType$ = "C" Then
            MeterOK = True
          End If
        Case "E"                'electric & demand elec.
          If MTRType$ = "E" Or MTRType$ = "D" Then
            MeterOK = True
          End If
        Case "G"                'gas
          If MTRType$ = "G" Then
            MeterOK = True
          End If
        Case "A", " "           'all meters
          MeterOK = True
        End Select
      End Select

      Average = UBCustRec(1).LocMeters(MtrCnt).AvgUse
      ILowRead = Right$(Str$((UBCustRec(1).LocMeters(MtrCnt).CurRead)), 8)
      LowRead = Fix(ILowRead)
      HighRead = Fix(Average * (HighVar / 100)) + UBCustRec(1).LocMeters(MtrCnt).CurRead

      MeterID = LTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrNum)
      MeterID = RTrim$(MeterID$)

      If Val(MeterID) = 0 Then
        MeterID = UBCustRec(1).Book + UBCustRec(1).SEQNUMB
      End If
      If Len(MeterID) < 8 Then
        MeterID = String$(8 - Len(MeterID), "0") + MeterID
      End If
      MeterID = Left$(MeterID, 8)

      UBPC3000RDRec(1).CustName = UBCustRec(1).CustName
      UBPC3000RDRec(1).ServAddress = Left$(UBCustRec(1).ServAddr, 16)

      C$ = QPTrim$(UBCustRec(1).USERCODE1)
      If Len(C$) > 0 Then
        Mid$(UBPC3000RDRec(1).ServAddress, 19, 1) = Left$(C$, 1)
      End If
      C$ = QPTrim$(UBCustRec(1).USERCODE2)
      If Len(C$) > 0 Then
        Mid$(UBPC3000RDRec(1).ServAddress, 20, 1) = Left$(C$, 1)
      End If

      UBPC3000RDRec(1).MeterID = MeterID
      UBPC3000RDRec(1).LowRead = LowRead
      UBPC3000RDRec(1).HighRead = HighRead
      UBPC3000RDRec(1).Account = Account
      UBPC3000RDRec(1).MeterType = UBCustRec(1).LocMeters(MtrCnt).MTRType
      UBPC3000RDRec(1).Book = Val(UBCustRec(1).Book)
      UBPC3000RDRec(1).CurRead = UBCustRec(1).LocMeters(MtrCnt).CurRead
      UBPC3000RDRec(1).PastRead = 0
      UBPC3000RDRec(1).ReadFlag = "N"
      'Modifed 04-28-97
      UBPC3000RDRec(1).Note1 = UBCustRec(1).HHMSG1
      UBPC3000RDRec(1).Note2 = UBCustRec(1).HHMSG2
      UBPC3000RDRec(1).Note3 = UBCustRec(1).HHMSG3
      '         ^^^
      UBPC3000RDRec(1).NoteStatus = ""
      Put UBPC3000RdFile, (LOF(UBPC3000RdFile) / UBPC3000RdRecLen) + 1, UBPC3000RDRec(1)
    End If

NoMeterTypeRet:
  Next MtrCnt

Return

GetHusky:
    
'  ReDim MsgText(0 To 5) As String
  
  MsgText(0) = "WARNING. . ."
  MsgText(1) = "Make sure the HUSKY is connected and ready to"
  MsgText(2) = "transfer files. The HUSKY should be at the"
  MsgText(3) = "'Husky File Transfer Utility' screen. Type 'H'"
  MsgText(4) = "and press the 'Yes' key at the C: prompt"
  MsgText(5) = "IS THE HUSKY READY TO PROCEED?"
  
  DoEvents
  
  Select Case GetOKorNot%(MsgText(), False, True)
  Case Not True
    GoTo GetHuskyOKExit
  End Select
  
  DeActivateControls Me
  
  MsgText(0) = "Husky File Transfer"
  MsgText(1) = "WARNING. . ."
  MsgText(2) = "PLEASE WAIT UNTIL THE HUSKY FILE"
  MsgText(3) = "TRANNSFER UTILITY HAS COMPLETED!"
  MsgText(4) = "Then Click 'OK' to Continue."

  DoEvents

  Shell "HCOMW32 " + HuskyPort + " /rx=ubcusttr.dat /abort", vbNormalFocus
  OkORNotFlag = GetOKorNotHH%(MsgText())
  
  ActivateControls Me
  
  DoEvents
  
  If OkORNotFlag Then
    If Chk4HuskyError% Then
      GoTo GetHuskyOKExit
    Else
      GoTo HuskyImportReadings
    End If
  Else 'cancled
    GoTo GetHuskyOKExit
  End If


HuskyImportReadings:

  UBCustRecLen = Len(UBCustRec(1))
  UBPC3000GetRdRecLen = Len(UBPC3000GetRDRec(1))
  
  UBFile = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBFile Len = UBCustRecLen
  
  UBPC3000GetRdFile = FreeFile
  Open UBPath$ + "UBCUSTTR.DAT" For Random Shared As UBPC3000GetRdFile Len = UBPC3000RdRecLen

  NumPC3000GetRdRecs = LOF(UBPC3000GetRdFile) / UBPC3000GetRdRecLen
  
  If NumPC3000GetRdRecs = 0 Then
    Close
    MsgText(0) = "ERROR:"
    MsgText(1) = "NO READINGS FOUND"
    MsgText(2) = " Check the handheld connection"
    MsgText(3) = " and try the transfer again!!!"
    MsgText(4) = "   Press any key to continue. "
    
    DoEvents
    GetOKorNot% MsgText(), True, True
    GoTo GetHuskyOKExit
  End If
  
  FrmShowPctComp.Label1 = "Processing Meter Readings."
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show '1, Parent
  
  For RMCnt& = 1 To NumPC3000GetRdRecs
    Get UBPC3000GetRdFile, RMCnt&, UBPC3000GetRDRec(1)
    WhatRMRec& = Val(QPTrim$(UBPC3000GetRDRec(1).Account))
    If Not (WhatRMRec&) = 0 Then
      Get UBFile, WhatRMRec&, UBCustRec(1)
      GoSub HuskyExtractRecord
    End If
    FrmShowPctComp.ShowPctComp RMCnt&, NumPC3000GetRdRecs
  Next

  Close
  MsgText(0) = "Import Readings."
  MsgText(1) = "Readings Updated Successfully."
  MsgText(2) = ""
  MsgText(3) = " IMPORTED: " + Str$(RMCnt&) + " Readings"
  MsgText(4) = ""
  MsgText(5) = ""
  
  DoEvents
  GetOKorNot% MsgText(), True, True
  

GetHuskyOKExit:

Return

HuskyExtractRecord:
  UpdCnt& = UpdCnt& + 1
  CustMTRRec = Val(Right$((QPTrim$(UBPC3000GetRDRec(1).Account)), 1))
  CurReading# = UBPC3000GetRDRec(1).CurRead
  If UBCustRec(1).LocMeters(CustMTRRec).ReadFlag = "Y" Then
    UBCustRec(1).LocMeters(CustMTRRec).CurRead = CurReading#
    UBCustRec(1).LocMeters(CustMTRRec).CurDate = UBPC3000GetRDRec(1).ReadDate
  Else
    UBCustRec(1).LocMeters(CustMTRRec).PrevRead = UBCustRec(1).LocMeters(CustMTRRec).CurRead
    UBCustRec(1).LocMeters(CustMTRRec).PastDate = UBCustRec(1).LocMeters(CustMTRRec).CurDate
    UBCustRec(1).LocMeters(CustMTRRec).ReadFlag = "Y"
    UBCustRec(1).LocMeters(CustMTRRec).CurDate = UBPC3000GetRDRec(1).ReadDate
    UBCustRec(1).LocMeters(CustMTRRec).CurRead = CurReading#
  End If
  'Modifed 04-28-97
  If UBPC3000GetRDRec(1).NoteStatus = "P" Then
    UBCustRec(1).NewNotes = True
    UBCustRec(1).HHMSG1 = UBPC3000GetRDRec(1).Note1
    UBCustRec(1).HHMSG2 = UBPC3000GetRDRec(1).Note2
    UBCustRec(1).HHMSG3 = UBPC3000GetRDRec(1).Note3
  End If
  Put UBFile, WhatRMRec&, UBCustRec(1)
Return

End Sub

Private Function GetOKorNotHH%(MsgText() As String)
  Dim zz As Integer, RetValue As Integer
  frmHHMsgInfo.Caption = MsgText(0)
  For zz = 1 To 4
    frmHHMsgInfo.Label(zz) = MsgText(zz)
  Next
  zz = Screen.TwipsPerPixelX
  Select Case screenW
  Case 800 To 1023
    If zz = 12 Then
      frmHHMsgInfo.Left = 4300
    Else
      frmHHMsgInfo.Left = 6600
    End If
  Case 1024 To 1279
    If zz = 12 Then
      frmHHMsgInfo.Left = 6900
    Else
      frmHHMsgInfo.Left = 9900
    End If
  Case Is >= 1280
    If zz = 12 Then
      frmHHMsgInfo.Left = 9900
    Else
      frmHHMsgInfo.Left = 13500
    End If
  Case Else
    frmHHMsgInfo.Left = 0
  End Select
  frmHHMsgInfo.Show vbModal
  RetValue = Val(frmHHMsgInfo.RetLabel)
  DoEvents
  Unload frmHHMsgInfo
  GetOKorNotHH% = RetValue
End Function


'**********************
'old sensus format a
'Special flag for Saluda Co add 3 0's onto read on import
Private Sub ImpExpOSenHHInfo(ByVal ImpExpFlag As Boolean)
  Dim HighVar As Integer, LowVar As Integer
  Dim WhatTypes As String, CustAcc As String
  Dim UBFile As Integer, UBSenIOFile As Integer
  Dim UBSenRdRecLen As Integer, NumSenRdRecs As Integer
  Dim BookCnt As Integer, MtrCnt As Integer
  Dim RMCnt As Long, WhatRMRec As Long
  Dim Account As String, SensusType As String
  Dim Average As Double, LowRead As Double
  Dim MeterID As String, MRDate As String
  'Dim ReadLowI As String, PrevRead As String
  'Dim NCurRead As String, PrevDate As String
  Dim HighRead As Double, ILowRead As Double
  Dim UBSenGetRecLen As Integer, NumSenGetRecs As Integer
  Dim MeterReadDate As String
  Dim DashPos As Integer
  Dim CurReading As Double

  UBCustRecLen = Len(UBCustRec(1))

  ReDim UBSenRdRec(1) As UBSensusReadRecType

  ' Check For Device Type and Run Appropriate Program
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen

  HighVar = UBSetUpRec(1).HighRead
  LowVar = UBSetUpRec(1).LowRead

  If HighVar < 100 Then
    HighVar = 100           'make sure
  End If
  If LowVar > HighVar Then
    LowVar = HighVar
  End If

  If ImpExpFlag Then     'EXPORTING METER READINGS
    WhatTypes$ = Left$(Me.fpMtrType.Text, 1)
    GoSub SSendSensus
    Call cmdExit_Click
  Else
    GoSub SGetSensus
    Call cmdExit_Click
  End If

Exit Sub

SGetSensus:

  SensusIOFile = HHPathInOut + "exssi00" + QPTrim(Str$(InterrNum)) + ".DAT"

  MRDate$ = QPTrim$(fpReadDate.Text)

  MsgText(0) = "Import Sensus Reading File."
  MsgText(1) = ""
  MsgText(2) = "Import File:"
  MsgText(3) = SensusIOFile
  MsgText(4) = "Ready to Proceed?"
  MsgText(5) = ""

  Select Case GetOKorNot%(MsgText(), False, True, 1)
  Case False
    GoTo ErrorSGetSensusExit
    'Stop
  End Select

  ReDim UBSenGetRdRec(1) As UBSensusGetReadRecType
  UBSenGetRecLen = Len(UBSenGetRdRec(1))

  UBSenIOFile = FreeFile
  Open SensusIOFile For Random Shared As UBSenIOFile Len = UBSenGetRecLen

  NumSenGetRecs = LOF(UBSenIOFile) / UBSenGetRecLen

  UBFile = FreeFile
  Open UBPath + "UBCUST.DAT" For Random Shared As UBFile Len = UBCustRecLen

  If NumSenGetRecs = 0 Then
    Close
    MsgText(0) = "ERROR:"
    MsgText(1) = "IMPORT FILE NOT FOUND"
    MsgText(2) = "Make sure that: " + "'" + "exssi00" + QPTrim(Str$(InterrNum)) + ".DAT" + "'"
    MsgText(3) = "is in the Sensus directory!"
    MsgText(4) = "Please call Southern Software for"
    MsgText(5) = "additional Information."
    GetOKorNot% MsgText(), True, False
    GoTo ErrorSGetSensusExit
  End If

  FrmShowPctComp.Label1 = "Importing Meter Reading Information."
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show '1, Parent

  RMCnt& = 1                ' Initialize File Counter to 1
  Do
    Get UBSenIOFile, RMCnt&, UBSenGetRdRec(1)
    WhatRMRec = Val(QPTrim$(UBSenGetRdRec(1).Account))
    If WhatRMRec > 0 Then
      GoSub SExtractRecord
    End If
    RMCnt& = RMCnt& + 1
    FrmShowPctComp.ShowPctComp RMCnt&, NumSenGetRecs
  Loop Until RMCnt& > NumSenGetRecs

  Close

  'Done show import complete
  MsgText(0) = "Sensus Operation"
  MsgText(1) = "Sensus Import Complete."
  MsgText(2) = ""
  MsgText(3) = "Imported:" + Str$(FileSize(SensusIOFile) / UBSenGetRecLen) + " Readings."
  MsgText(4) = ""
  MsgText(5) = ""
  GetOKorNot% MsgText(), True, True

ErrorSGetSensusExit:

Return

SExtractRecord:

  Get UBFile, WhatRMRec, UBCustRec(1)
  DashPos = InStr(UBSenGetRdRec(1).Account, "-")
  MtrCnt = Val(Mid$(UBSenGetRdRec(1).Account, DashPos + 1))

  If MtrCnt = 0 Then MtrCnt = 1
  ' Check Meter Updated Flag
  ' Update Meter W/Reading
  If SaludaCoF Then
    CurReading# = Val(UBSenGetRdRec(1).Reading) * 1000#
  Else
    CurReading# = Val(UBSenGetRdRec(1).Reading)
  End If
  MeterReadDate$ = Left$(UBSenGetRdRec(1).DateRead, 2) + "/" + Mid$(UBSenGetRdRec(1).DateRead, 3, 2) + "/" + Mid$(MRDate$, 7, 2) + Right$(UBSenGetRdRec(1).DateRead, 2)
  If Date2Num(MeterReadDate$) < 0 Then
    MeterReadDate$ = MRDate$
  End If

  If UBCustRec(1).LocMeters(MtrCnt).ReadFlag = "Y" Then
    UBCustRec(1).LocMeters(MtrCnt).CurRead = CurReading#
    UBCustRec(1).LocMeters(MtrCnt).CurDate = Date2Num(MeterReadDate$)
  Else
    UBCustRec(1).LocMeters(MtrCnt).PrevRead = UBCustRec(1).LocMeters(MtrCnt).CurRead
    UBCustRec(1).LocMeters(MtrCnt).PastDate = UBCustRec(1).LocMeters(MtrCnt).CurDate
    UBCustRec(1).LocMeters(MtrCnt).ReadFlag = "Y"
    UBCustRec(1).LocMeters(MtrCnt).CurDate = Date2Num(MeterReadDate$)
    UBCustRec(1).LocMeters(MtrCnt).CurRead = CurReading#
  End If

  Put UBFile, WhatRMRec, UBCustRec(1)

Return

'************************* Send info to sensus
SSendSensus:

  If CycleCnt > 0 Then
    'build sensus output file name
    SensusIOFile = HHPathInOut + "SSI00" + QPTrim(Str$(InterrNum)) + ".RTE"

    KillFile SensusIOFile 'kill old if there

    If UBSetUpRec(1).UseSeq = "Y" Then 'if they are using sequence numbers
      IdxRecLen = 4         'we are using a integer
      MakeSequenceIndex "Sequence Number", Me
      IdxNumOfRecs = FileSize&("UBTEMP.IDX") \ 4
      ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
      FGetAH "UBTEMP.IDX", IdxBuff(), 4, IdxNumOfRecs
    Else                               'use location number index
      IdxRecLen = 4         'we are using a integer
      IdxFileSize& = FileSize&("UBCUSTBK.IDX")
      IdxNumOfRecs = IdxFileSize& \ IdxRecLen
      ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
      FGetAH "UBCUSTBK.IDX", IdxBuff(), IdxRecLen, IdxNumOfRecs            'load it
    End If

    FrmShowPctComp.Label1 = "Exporting Meter Reading Information."
    FrmShowPctComp.cmdCancel.Enabled = False
    FrmShowPctComp.Show '1, Parent

    UBSenRdRecLen = Len(UBSenRdRec(1))

    UBSenIOFile = FreeFile
    Open SensusIOFile For Random Shared As UBSenIOFile Len = UBSenRdRecLen

    NumSenRdRecs = LOF(UBSenIOFile) / UBSenRdRecLen

    UBFile = FreeFile
    Open UBPath + "UBCUST.DAT" For Random Shared As UBFile Len = UBCustRecLen

    RMCnt& = 1
    Do
      WhatRMRec& = IdxBuff(RMCnt&).RecNum
      If WhatRMRec& > 0 Then
        Get UBFile, WhatRMRec&, UBCustRec(1)
        For BookCnt = 1 To CycleCnt
          If Val(UBCustRec(1).Book) = Cycle(BookCnt) And (UBCustRec(1).Status <> "F") Then
            GoSub SWriteRecord
            Exit For
          End If
        Next
      End If
      RMCnt& = RMCnt& + 1
      FrmShowPctComp.ShowPctComp RMCnt&, IdxNumOfRecs
    Loop Until RMCnt& > IdxNumOfRecs

    Close   'done with output file.

    MsgText(0) = "Sensus Operation"
    MsgText(1) = "Sensus Export Complete."
    MsgText(2) = ""
    MsgText(3) = "Exported:" + Str$(FileSize(SensusIOFile) / UBSenRdRecLen) + " Readings."
    MsgText(4) = ""
    MsgText(5) = ""
    GetOKorNot% MsgText(), True, True

  End If

Return

SWriteRecord:
'*****
  Account$ = Str$(WhatRMRec&)
  For MtrCnt = 1 To 7     'look at all possiable meters

    Select Case UBCustRec(1).LocMeters(MtrCnt).MTRType
    Case "C", "S", "W", "T", "E", "D", "P", "I"   'here dale
      If (UBCustRec(1).LocMeters(MtrCnt).MTRType = "T" And Val(QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrNum)) = 0) Then
        GoTo ESkipEm
      Else
        ' Determine Sensus Meter Type
        Select Case UBCustRec(1).LocMeters(MtrCnt).MTRType
        Case "T", "I"
          SensusType$ = "B"
        Case "P"
          SensusType$ = "P"
        Case Else
          SensusType$ = "M"
        End Select
        'Determine High and Low Reading
        Average# = UBCustRec(1).LocMeters(MtrCnt).AvgUse
        If Average# < 0 Then
          Average# = 0
        End If

        ILowRead# = Val(QPTrim$(Str$(UBCustRec(1).LocMeters(MtrCnt).CurRead)))
        HighRead# = Fix(Average# * (HighVar / 100)) + UBCustRec(1).LocMeters(MtrCnt).CurRead

        If Fix(HighRead#) = ILowRead# Then
          HighRead# = HighRead# + 5
        End If

        MeterID$ = LTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrNum)
        MeterID$ = RTrim$(MeterID$)
        If Val(MeterID$) = 0 Then
          MeterID$ = UBCustRec(1).Book + UBCustRec(1).SEQNUMB
        End If
        If Len(MeterID$) < 8 Then
          MeterID$ = String$(8 - Len(MeterID$), "0") + MeterID$
        End If
        
        MeterID$ = Left$(MeterID$, 8)
'Set Record Fields and Put On Disk
'************ clear old info
        UBSenRdRec(1).ServAddress = ""
        UBSenRdRec(1).MeterID = ""
        UBSenRdRec(1).LowRead = ""
        UBSenRdRec(1).HighRead = ""
        UBSenRdRec(1).Account = ""
        UBSenRdRec(1).SensusType = ""
        UBSenRdRec(1).CustName = ""
        UBSenRdRec(1).SerialNumb = ""
'***************************

        UBSenRdRec(1).ServAddress = QPTrim$(UBCustRec(1).ServAddr)
        UBSenRdRec(1).MeterID = MeterID$
        UBSenRdRec(1).LowRead = Str$(ILowRead#)
        UBSenRdRec(1).HighRead = Str$(HighRead#)
        UBSenRdRec(1).Account = Account$ + "-" + QPTrim$(Str$(MtrCnt))
        UBSenRdRec(1).SensusType = SensusType$
        UBSenRdRec(1).CustName = QPTrim$(UBCustRec(1).CustName)
        UBSenRdRec(1).SerialNumb = UBCustRec(1).LocMeters(MtrCnt).MtrNum
'        UBSenRdRec(1).LocationNumber = QPTrim$(UBCustRec(1).Book + UBCustRec(1).SEQNUMB)
        'UBSenRdRec(1).Message = QPTrim$(UBCustRec(1).HHMSG1)
        'reuse the MeterID$ variable
        'MeterID$ = QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrIDNO)
        'If Len(MeterID$) > 0 Then
        '  UBSenRdRec(1).MtrIDMST = Left$(MeterID$, Len(MeterID$) - 1)
        '  UBSenRdRec(1).MtrIDNO = Right$(MeterID$, 1)
        'End If
        'UBSenRdRec(1).MtrLat = MakeExpCoordinate$(UBCustRec(1).LocMeters(MtrCnt).MtrLat)
        'UBSenRdRec(1).MtrLng = MakeExpCoordinate$(UBCustRec(1).LocMeters(MtrCnt).MtrLng)

        Put UBSenIOFile, , UBSenRdRec(1)
      End If
    Case Else
      'no meter in this slot.
    End Select
ESkipEm:
  Next
Return

End Sub

Private Sub ImpExpBadgerHHInfo(ByVal ImpExpFlag As Boolean)
  Dim HighVar As Integer, LowVar As Integer, IdxFile As Integer
  Dim WhatTypes As String, CustAcc As String, Prec As Long
  Dim UBFile As Integer, BadgerCFGFile As Integer, Multi As Long
  Dim UBBadgerRecLen As Integer, NumBadgerRec As Integer, OurInfo As String
  Dim BookCnt As Integer, MtrCnt As Integer, IdxNumRecs As Long
  Dim RMCnt As Long, WhatRMRec As Long, BadgerFile As String
  Dim Account As String, BadgerType As String, DoneCnt As Long
  Dim Average As Double, LowRead As Double, NumberofRoutes As Integer
  Dim MeterID As String, MRDate As String, PathWay As String
  Dim HighRead As Double, ILowRead As Double, IdxFileName As String
  Dim UBBdgrGetRecLen As Integer, NumBdgrGetRecs As Integer, CircleC As String
  Dim MeterReadDate As String, Z9 As String, FFName As String
  Dim DashPos As Integer, cnt As Long, UBBadRdRecLen As Integer
  Dim CurReading As Double, UBBadRdFile As Integer, NumBadRdRecs As Long
  Dim CRead As Double, IHiRead As String, ILoRead As String, MtNumb As String
  Dim CurRead As String, Acct As String, MTCnt As String, SEQNUMB As String
  Dim MT As String, MeterRecord As Integer, ReadingDate As Integer
  Dim OverDate As Integer, UseOverDate As Boolean, UpdCnt As Long
  PathWay$ = HHPathInOut
  ReDim Route(100)
  UBCustRecLen = Len(UBCustRec(1))
  Z9$ = "000000000"
  BadgerFile$ = "BADGERMR.DAT"
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen
  HighVar = UBSetUpRec(1).HighRead
  If HighVar > 0 Then
    HighVar = HighVar / 100
  Else
    HighVar = 1.5
  End If

  LowVar = UBSetUpRec(1).LowRead
  If LowVar > 0 Then
    LowVar = LowVar / 100
  Else
    LowVar = 0.75
  End If

'  ReDim UBBadgerRec(1) As UBSensusRecType
'  UBBadgerRecLen = Len(UBBadgerRec(1))
'  BadgerCFGFile = FreeFile
'  Open "UBBADGER.DAT" For Random Shared As BadgerCFGFile Len = UBBadgerRecLen
'  NumBadgerRec = LOF(BadgerCFGFile) / UBBadgerRecLen
'  If NumBadgerRec = 1 Then
'    Get BadgerCFGFile, 1, UBBadgerRec(1)
'    PathWay$ = UBBadgerRec(1).PathWay
'  End If

'  ReDim Choice$(3, 0)
'  Choice$(0, 0) = "1"
'  Choice$(1, 0) = "Send Info to Badger"
'  Choice$(2, 0) = "Get Info From Badger"

'  If NumBadgerRec = 1 Then
'    Form$(2, 0) = PathWay$
'  End If

'  Do
'    EditForm Form$(), Fld(), frm(1), Cnf, Action
'    Select Case frm(1).KeyCode
'    Case F10Key
'      PathWay$ = Form$(2, 0)
'      UBBadgerRec(1).PathWay = PathWay$
'      Put BadgerCFGFile, 1, UBBadgerRec(1)
'      Close BadgerCFGFile
'
'      Operation$ = Left$(QPTrim$(Form$(1, 0)), 1)
'      Select Case Operation$
'      Case "S"
'        GoSub BadgerSend
'        Done = True
'      Case "G"
'        GoSub BadgerGet
'        Done = True
'      Case Else
'        OK = MsgBox(LibName$, "NOOPERAT")
'        frm(1).FldNo = 1
'        Action = 1
'        Done = False
'      End Select
'    Case ESC
'      Done = True
'      Exit Sub
'    Case Else
'      Done = False
'    End Select
'  Loop Until Done
  If ImpExpFlag Then     'EXPORTING METER READINGS
    WhatTypes$ = Left$(Me.fpMtrType.Text, 1)
    GoSub BadgerSend
    Call cmdExit_Click
  Else
    GoSub BadgerGet
    Call cmdExit_Click
  End If

  
BadgerReadExit:
  Exit Sub

BadgerSend:
  NumberofRoutes = 0

        GoSub MakeOUTTFFileName
        GoSub BadgerOpenCust    'Open Customer Data File

        ReDim UBBadRdRec(1) As UBBadgerRecType
        UBBadRdRecLen = Len(UBBadRdRec(1))

        KillFile FFName$
        UBBadRdFile = FreeFile
        Open FFName$ For Random Shared As UBBadRdFile Len = UBBadRdRecLen
        NumBadRdRecs = LOF(UBBadRdFile) / UBBadRdRecLen

    HighVar = UBSetUpRec(1).HighRead
    If UBSetUpRec(1).UseSeq = "Y" Then
      MakeSequenceIndex "Sequence Number", Me
      IdxFileName = UBPath$ + "UBTEMP.IDX"
    Else
      IdxFileName = UBPath$ + "UBCUSTBK.IDX"
    End If
    IdxRecLen = 4         'we are using a integer
    IdxFileSize& = FileSize&(IdxFileName)
    IdxNumOfRecs = IdxFileSize& \ IdxRecLen
    ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
    FGetAH IdxFileName, IdxBuff(), IdxRecLen, IdxNumOfRecs            'load it
  
  FrmShowPctComp.Label1 = "Exporting Reading Information."
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show '1, Parent

        DoneCnt& = 1
        Do
          Prec& = IdxBuff(DoneCnt&).RecNum
          If Prec& > 0 Then
            Get UBFile, Prec&, UBCustRec(1)
            For BookCnt = 1 To CycleCnt
              If Val(UBCustRec(1).Book) = Cycle(BookCnt) And (UBCustRec(1).Status <> "F") Then
                
                GoSub BadgerPutRec
              End If
            Next
          End If
          DoneCnt& = DoneCnt& + 1
         FrmShowPctComp.ShowPctComp DoneCnt&, IdxNumOfRecs&
        Loop Until DoneCnt& > IdxNumOfRecs&
        Close
Return

BadgerPutRec:

  'modifyed
  'May Have Up to 7 Meters to Read
  MtrCnt = 1
  Account$ = Str$(Prec&)
  While MtrCnt < 8
    If (Asc(UBCustRec(1).LocMeters(MtrCnt).MTRType) > 32) Then
      Select Case UBCustRec(1).LocMeters(MtrCnt).MTRType
      Case "C", "S", "W", "T", "E", "D"         'here dale
        If UBCustRec(1).LocMeters(MtrCnt).MTRMulti <= 0 Then
          UBCustRec(1).LocMeters(MtrCnt).MTRMulti = 1
        End If
        Multi& = UBCustRec(1).LocMeters(MtrCnt).MTRMulti

        Average# = UBCustRec(1).LocMeters(MtrCnt).AvgUse
        CRead# = UBCustRec(1).LocMeters(MtrCnt).CurRead
        'make sure we have valid average & current readings
        If CRead# < 0 Then
          CRead# = 0
        End If
        If Average# <= 0 Then
          Average# = CRead#
        End If

        IHiRead$ = Right$((Z9$ + QPTrim$(Str$(CRead# + Fix(Average# * HighVar)))), 9)
        ILoRead$ = Right$((Z9$ + QPTrim$(Str$(CRead# + Fix(Average# * LowVar)))), 9)
        CurRead$ = Right$((Z9$ + QPTrim$(Str$(CRead#))), 9)

        Acct$ = UBCustRec(1).Book + UBCustRec(1).SEQNUMB
        If UBCustRec(1).Seq < 0 Then
          UBCustRec(1).Seq = 0
        End If
        MTCnt$ = QPTrim$(Str$(MtrCnt))
        SEQNUMB$ = QPTrim$(Str$(UBCustRec(1).Seq))
        OurInfo$ = QPTrim$(Str$(Prec&)) + "-" + MTCnt$

        MtNumb$ = QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrNum)
        'IF LEN(MtNumb$) > 7 THEN
        '  MtNumb$ = LEFT$(MtNumb$, 7)
        'END IF

        Select Case Multi&
        Case 1
          CircleC$ = "1"
        Case 10
          CircleC$ = "2"
        Case 100
          CircleC$ = "3"
        Case 1000
          CircleC$ = "4"
        End Select

        LSet UBBadRdRec(1).Fill1 = ""
        LSet UBBadRdRec(1).CustName = QPTrim$(UBCustRec(1).CustName)
        LSet UBBadRdRec(1).ServAddr = QPTrim$(UBCustRec(1).ServAddr)
        LSet UBBadRdRec(1).MtrNum1 = MtNumb$
        LSet UBBadRdRec(1).Multi = ""           'Multi&
        LSet UBBadRdRec(1).Status = UBCustRec(1).Status
        LSet UBBadRdRec(1).ReadCode = MTCnt$
        LSet UBBadRdRec(1).ServFreq = MTCnt$
        LSet UBBadRdRec(1).DNI = ""
        LSet UBBadRdRec(1).MtrNum2 = UBBadRdRec(1).MtrNum1
        LSet UBBadRdRec(1).NumDials = ""
        RSet UBBadRdRec(1).HiRead = IHiRead$
        RSet UBBadRdRec(1).LoRead = ILoRead$
        LSet UBBadRdRec(1).CurrRead = ""        'changed per SCS 'CurRead$
        LSet UBBadRdRec(1).ReadTime = ""
        LSet UBBadRdRec(1).ReadCode2 = ""
        LSet UBBadRdRec(1).CmntCode = ""
        LSet UBBadRdRec(1).Fill2 = ""
        LSet UBBadRdRec(1).Account = Acct$
        LSet UBBadRdRec(1).ReadDate = ""
        LSet UBBadRdRec(1).DevCode = "P"
        LSet UBBadRdRec(1).MMILat = ""
        LSet UBBadRdRec(1).MMILong = ""
        LSet UBBadRdRec(1).MMIChanl = ""
        LSet UBBadRdRec(1).CircleCode = CircleC$                'changed per SCS
        'this is a code from above
        RSet UBBadRdRec(1).SEQNUMB = SEQNUMB$
        LSet UBBadRdRec(1).MfgModel = ""
        LSet UBBadRdRec(1).UserField = OurInfo$
        LSet UBBadRdRec(1).ReadID = ""
        LSet UBBadRdRec(1).ReadCo1 = ""
        LSet UBBadRdRec(1).ReadCo2 = ""
        LSet UBBadRdRec(1).ReadCo3 = ""
        LSet UBBadRdRec(1).MMIReadCode = ""
        LSet UBBadRdRec(1).Pad = ""
        LSet UBBadRdRec(1).CrLf = CrLf$
        Put UBBadRdFile, , UBBadRdRec(1)
      End Select
    End If
ESkipEm:
    MtrCnt = MtrCnt + 1
  Wend

Return

MakeOUTTFFileName:
  PathWay$ = QPTrim$(PathWay$)
  If Len(PathWay$) > 0 Then
    If Right$(PathWay$, 1) <> "\" Then
      PathWay$ = PathWay$ + "\"
    End If
  End If
  FFName$ = PathWay$ + "CONNECT.IN3"
Return

MakeINTFFileName:
  PathWay$ = QPTrim$(PathWay$)
  If Len(PathWay$) > 0 Then
    If Right$(PathWay$, 1) <> "\" Then
      PathWay$ = PathWay$ + "\"
    End If
  End If
  FFName$ = PathWay$ + "CONNECT.OT3"
Return

BadgerGet:
  NumberofRoutes = 0

      OverDate = Date2Num(fpReadDate)
      If OverDate > 0 Then
        UseOverDate = True
      Else
        UseOverDate = False
      End If
      GoSub MakeINTFFileName      'Get Badger File
      GoSub BadgerOpenCust      'Open Customer Data File

      ReDim UBBadRdRec(1) As UBBadgerRecType
      UBBadRdRecLen = Len(UBBadRdRec(1))
      'Open meter reading information File
      UBBadRdFile = FreeFile
      Open FFName$ For Random Shared As UBBadRdFile Len = UBBadRdRecLen
      NumBadRdRecs& = LOF(UBBadRdFile) / UBBadRdRecLen

      If NumBadRdRecs& = 0 Then
        Close
        MsgBox "No Records Found to Import.", vbOKOnly, "Procedure Ended"
        GoTo BadgerReadExit
      End If

  FrmShowPctComp.Label1 = "Importing Reading Information."
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show '1, Parent

      DoneCnt& = 1              ' Initialize File Counter to 1
      Do
        Get UBBadRdFile, DoneCnt&, UBBadRdRec(1)
        Prec& = Val(QPTrim$(UBBadRdRec(1).UserField))
        If Prec& > 0 Then
          Get UBFile, Prec&, UBCustRec(1)
          GoSub EExtractRecord
        Else
          'Stop
        End If
        FrmShowPctComp.ShowPctComp DoneCnt&, NumBadRdRecs&
       DoneCnt& = DoneCnt& + 1
      Loop Until DoneCnt& > NumBadRdRecs&
      Close
      'Done = True

Return
EExtractRecord:
'this extracts the reading & date

 UpdCnt& = UpdCnt& + 1
  'QPrintRC " Updated Count:" + Str$(UpdCnt&), 11, 28, -1

  DashPos = InStr(UBBadRdRec(1).UserField, "-")
  MT$ = Mid$(UBBadRdRec(1).UserField, DashPos + 1)
  MeterRecord = Val(MT$)

  If MeterRecord = 0 Then MeterRecord = 1
  ' Check Meter Updated Flag
  ' Update Meter W/Reading
  CurReading# = Val(UBBadRdRec(1).CurrRead)

  If UBCustRec(1).LocMeters(MeterRecord).MTRMulti = 10 Then
    CurReading# = (CurReading# * 0.1)
  End If

  If UseOverDate Then
    ReadingDate = OverDate          'if they want to overide the read date
  Else
    MeterReadDate$ = Left$(UBBadRdRec(1).ReadDate, 2) + "/" + Mid$(UBBadRdRec(1).ReadDate, 3, 2) + "/" + Right$(UBBadRdRec(1).ReadDate, 4)
    ReadingDate = Date2Num(MeterReadDate$)
    If ReadingDate <= 0 Then
      ReadingDate = Date2Num(Date$) 'if the read date was bad then fix it
    End If
  End If

  If UBCustRec(1).LocMeters(MeterRecord).ReadFlag = "Y" Then
    UBCustRec(1).LocMeters(MeterRecord).CurRead = CurReading#
    UBCustRec(1).LocMeters(MeterRecord).CurDate = ReadingDate 'Date2Num(MeterReadDate$)
  Else
    UBCustRec(1).LocMeters(MeterRecord).PrevRead = UBCustRec(1).LocMeters(MeterRecord).CurRead
    UBCustRec(1).LocMeters(MeterRecord).PastDate = UBCustRec(1).LocMeters(MeterRecord).CurDate
    UBCustRec(1).LocMeters(MeterRecord).ReadFlag = "Y"
    UBCustRec(1).LocMeters(MeterRecord).CurDate = ReadingDate 'Date2Num(MeterReadDate$)
    UBCustRec(1).LocMeters(MeterRecord).CurRead = CurReading#
  End If
  'now update the customers record...
  Put UBFile, Prec&, UBCustRec(1)
Return
BadgerOpenCust:
'  REDIM UBCustRec(1) AS NewUBCustRecType
'  UBCustRecLen = LEN(UBCustRec(1))
  UBFile = FreeFile
  Open "UBCUST.DAT" For Random Shared As UBFile Len = UBCustRecLen
Return

End Sub
'this badger is for Y just a few diff from X added March 2010 by PS-
'ONE TIME IMPORT NEW METORS with old reads for meter 1 and new reads and info for meter 2

'Changed before Murphy so will use same type as G hh type does.
Private Sub ImpExpBadgerV3HHInfo(ByVal ImpExpFlag As Boolean)
  Dim HighVar As Integer, LowVar As Integer, IdxFile As Integer
  Dim WhatTypes As String, CustAcc As String, Prec As Long
  Dim UBFile As Integer, BadgerCFGFile As Integer, Multi As Long
  Dim UBBadgerRecLen As Integer, NumBadgerRec As Integer, OurInfo As String
  Dim BookCnt As Integer, MtrCnt As Integer, IdxNumRecs As Long
  Dim RMCnt As Long, WhatRMRec As Long, BadgerFile As String
  Dim Account As String, BadgerType As String, DoneCnt As Long
  Dim Average As Double, LowRead As Double, NumberofRoutes As Integer
  Dim MeterID As String, MRDate As String, PathWay As String
  Dim HighRead As Double, ILowRead As Double, IdxFileName As String
  Dim UBBdgrGetRecLen As Integer, NumBdgrGetRecs As Integer, CircleC As String
  Dim MeterReadDate As String, Z9 As String, FFName As String
  Dim DashPos As Integer, cnt As Long, UBBadRdRecLen As Integer
  Dim CurReading As Double, UBBadRdFile As Integer, NumBadRdRecs As Long
  Dim CRead As Double, IHiRead As String, ILoRead As String, MtNumb As String
  Dim CurRead As String, Acct As String, MTCnt As String, SEQNUMB As String
  Dim MT As String, MeterRecord As Integer, ReadingDate As Integer
  Dim OverDate As Integer, UseOverDate As Boolean, UpdCnt As Long
  PathWay$ = HHPathInOut
  ReDim Route(100)
  UBCustRecLen = Len(UBCustRec(1))
  Z9$ = "000000000"
  BadgerFile$ = "BADGERMR.DAT"
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen
  HighVar = UBSetUpRec(1).HighRead
  If HighVar > 0 Then
    HighVar = HighVar / 100
  Else
    HighVar = 1.5
  End If

  LowVar = UBSetUpRec(1).LowRead
  If LowVar > 0 Then
    LowVar = LowVar / 100
  Else
    LowVar = 0.75
  End If

    GoSub BadgerGet
    Call cmdExit_Click

  
BadgerReadExit:
  Exit Sub



MakeINTFFileName:
  PathWay$ = QPTrim$(PathWay$)
  If Len(PathWay$) > 0 Then
    If Right$(PathWay$, 1) <> "\" Then
      PathWay$ = PathWay$ + "\"
    End If
  End If
  FFName$ = PathWay$ + "CONNECT.OT3"
Return

BadgerGet:
  NumberofRoutes = 0

      OverDate = Date2Num(fpReadDate)
      If OverDate > 0 Then
        UseOverDate = True
      Else
        UseOverDate = False
      End If
      GoSub MakeINTFFileName      'Get Badger File
      GoSub BadgerOpenCust      'Open Customer Data File

      ReDim UBBadRdRec(1) As UBBadgerV3RecType
      UBBadRdRecLen = Len(UBBadRdRec(1))
      'Open meter reading information File
      UBBadRdFile = FreeFile
      Open FFName$ For Random Shared As UBBadRdFile Len = UBBadRdRecLen
      NumBadRdRecs& = LOF(UBBadRdFile) / UBBadRdRecLen

      If NumBadRdRecs& = 0 Then
        Close
        MsgBox "No Records Found to Import.", vbOKOnly, "Procedure Ended"
        GoTo BadgerReadExit
      End If

  FrmShowPctComp.Label1 = "Importing Reading Information."
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show '1, Parent

      DoneCnt& = 1              ' Initialize File Counter to 1
      Do
        Get UBBadRdFile, DoneCnt&, UBBadRdRec(1)
        Prec& = Val(QPTrim$(UBBadRdRec(1).UserUse))
        If Prec& > 0 Then
          Get UBFile, Prec&, UBCustRec(1)
          GoSub EExtractRecord
        Else
          'Stop
        End If
        FrmShowPctComp.ShowPctComp DoneCnt&, NumBadRdRecs&
       DoneCnt& = DoneCnt& + 1
      Loop Until DoneCnt& > NumBadRdRecs&
      Close
      'Done = True

Return
EExtractRecord:
'this extracts the reading & date

 UpdCnt& = UpdCnt& + 1
  'QPrintRC " Updated Count:" + Str$(UpdCnt&), 11, 28, -1

  DashPos = InStr(UBBadRdRec(1).UserUse, "-")
  MT$ = Mid$(UBBadRdRec(1).UserUse, DashPos + 1)
  MeterRecord = Val(MT$)

  If MeterRecord = 0 Then MeterRecord = 1
  ' Check Meter Updated Flag
  ' Update Meter W/Reading
  CurReading# = Val(UBBadRdRec(1).CurrRead)

  If UBCustRec(1).LocMeters(MeterRecord).MTRMulti = 10 Then
    CurReading# = (CurReading# * 0.1)
  End If

  If UseOverDate Then
    ReadingDate = OverDate          'if they want to overide the read date
  Else
    MeterReadDate$ = Left$(UBBadRdRec(1).ReadDate, 2) + "/" + Mid$(UBBadRdRec(1).ReadDate, 3, 2) + "/" + Right$(UBBadRdRec(1).ReadDate, 4)
    ReadingDate = Date2Num(MeterReadDate$)
    If ReadingDate <= 0 Then
      ReadingDate = Date2Num(Date$) 'if the read date was bad then fix it
    End If
  End If

  If UBCustRec(1).LocMeters(MeterRecord).ReadFlag = "Y" Then
    UBCustRec(1).LocMeters(MeterRecord).CurRead = CurReading#
    UBCustRec(1).LocMeters(MeterRecord).CurDate = ReadingDate 'Date2Num(MeterReadDate$)
  Else
    UBCustRec(1).LocMeters(MeterRecord).PrevRead = UBCustRec(1).LocMeters(MeterRecord).CurRead
    UBCustRec(1).LocMeters(MeterRecord).PastDate = UBCustRec(1).LocMeters(MeterRecord).CurDate
    UBCustRec(1).LocMeters(MeterRecord).ReadFlag = "Y"
    UBCustRec(1).LocMeters(MeterRecord).CurDate = ReadingDate 'Date2Num(MeterReadDate$)
    UBCustRec(1).LocMeters(MeterRecord).CurRead = CurReading#
  End If
  
      Select Case UBBadRdRec(1).MtrRes
        Case "1"
          UBCustRec(1).LocMeters(MeterRecord + 1).MTRMulti = 1
        Case "2"
          UBCustRec(1).LocMeters(MeterRecord + 1).MTRMulti = 10
        Case "3"
          UBCustRec(1).LocMeters(MeterRecord + 1).MTRMulti = 100
        Case "4"
          UBCustRec(1).LocMeters(MeterRecord + 1).MTRMulti = 1000
       End Select
     
  UBCustRec(1).LocMeters(MeterRecord + 1).AvgUse = 0
  UBCustRec(1).LocMeters(MeterRecord + 1).InsDate = ReadingDate
  UBCustRec(1).LocMeters(MeterRecord + 1).MtrNum = "Z" + UBBadRdRec(1).MtrSerN
  UBCustRec(1).LocMeters(MeterRecord + 1).MtrIDNO = UBBadRdRec(1).MtrNum1
  UBCustRec(1).LocMeters(MeterRecord + 1).MtrLat = Val(QPTrim$(UBBadRdRec(1).Latitude))
  UBCustRec(1).LocMeters(MeterRecord + 1).MtrLng = Val(QPTrim$(UBBadRdRec(1).Longitude))
  If Len(QPTrim(UBCustRec(1).LocMeters(MeterRecord).MTRType)) > 0 Then
  UBCustRec(1).LocMeters(MeterRecord + 1).MTRType = UBCustRec(1).LocMeters(MeterRecord).MTRType
  End If
  UBCustRec(1).LocMeters(MeterRecord + 1).MtrUnit = UBCustRec(1).LocMeters(MeterRecord).MtrUnit
  UBCustRec(1).LocMeters(MeterRecord + 1).NumUser = UBCustRec(1).LocMeters(MeterRecord).NumUser
  UBCustRec(1).LocMeters(MeterRecord + 1).ReadFlag = UBCustRec(1).LocMeters(MeterRecord).ReadFlag
  UBCustRec(1).LocMeters(MeterRecord + 1).UseCnt = UBCustRec(1).LocMeters(MeterRecord).UseCnt
  
  UBCustRec(1).LocMeters(MeterRecord + 1).PrevRead = 0
  UBCustRec(1).LocMeters(MeterRecord + 1).PastDate = ReadingDate
  UBCustRec(1).LocMeters(MeterRecord + 1).CurDate = ReadingDate
  UBCustRec(1).LocMeters(MeterRecord + 1).CurRead = 0
  
  'now update the customers record...
  Put UBFile, Prec&, UBCustRec(1)
Return
BadgerOpenCust:
'  REDIM UBCustRec(1) AS NewUBCustRecType
'  UBCustRecLen = LEN(UBCustRec(1))
  UBFile = FreeFile
  Open "UBCUST.DAT" For Random Shared As UBFile Len = UBCustRecLen
Return

End Sub

  Private Sub ImpExpRadioReadHHInfo(ByVal ImpExpFlag As Boolean)
    'Const PathData$ = "UBBAGPTH.DAT"
    Dim Path As String, RDGLen As Integer, FileNam As String
    Dim UBFile As Integer, ReadFile As Integer, NumReadRecs As Long
    Dim zz As Integer, CustRec As Long, GotCust As Boolean
    Dim cnt As Long, RecordID As String, CustName As String
    Dim UpdCnt As Long, MeterRecord As Long, RDate As String
    Dim ReadDate As Integer, CurReading As Double, InfoOK As Boolean
    Dim NewPath As String, PathFile As Integer, q As String
    Dim cb As String, BroadFlag As Boolean, CYCLES As String
    Dim Route As String, TCnt As Long, MReading As String
    Dim TReading As String, KCnt As Long, KReading As String
    Dim RReading As String, Reading As String, Done As Boolean
    Dim HighReadPerc As Double, LowReadPerc As Double, CUSFile As Integer
    Dim MtrFile As Integer, RdgFile As Integer, RFFFile As Integer
    Dim WRRFile As Integer, Prec As Long, RecNumber As Long
    Dim DisplayAcct As String, MeterNumber As String, MNum2File As String
    Dim MeterSeq As String, Dials As Integer, DialsS As String
    Dim Decimals As String, ERTID As String, ReadMethod As String
    Dim PrevRead As String, HighRead As Double, HighReadS As String
    Dim LowRead As Double, LowReadS As String, FileHdrLine As String
    Dim CycleLine As String, RouteHeader As String, outfile As Integer, UBRpt As Integer
    Dim CusRecord As Long, CusLine As String, MtrLine As String
    Dim RdgLine As String, RffLine As String, WRRLine As String
    Dim Chkbook As Boolean, RType As String, MPIDNO As String, BluSerial As String
 
      '''' C:\mvrs50\xfer
    'PathData$ = "UBBAGPTH.DAT"
    On Local Error GoTo GiveItUp
    Chkbook = False
    PathData$ = UBHHPathWayFile
    If ImpExpFlag Then     'EXPORTING METER READINGS
      GoSub RadioReadIF
    Else
      GoSub GetBadgerReadings
    End If

     Close

AbortExit:

    Exit Sub

RadioReadIF:
  
    q$ = Chr$(34)
    cb$ = Space$(45)

    If InStr(UBSetUpRec(1).UTILNAME, "BROADWAY") > 0 Then
      BroadFlag = True
    End If
    Path$ = GetBadgerPath$

   ' ReDim Route(100)
    ReDim FHD(1) As FDRTYPE
    ReDim CHD(1) As CDRTYPE
    ReDim RHD(1) As RDRType
    ReDim CUS(1) As CUSType
    ReDim Mtr(1) As MTRType
    ReDim Rdg(1) As RDGType
    ReDim RFF(1) As RFFType
    ReDim WRR(1) As WRRType


BadDataRestart:

        GoSub CheckInfo
        If InfoOK Then
          FrmShowPctComp.Label1 = "Exporting Reading Information."
          FrmShowPctComp.cmdCancel.Enabled = False
          FrmShowPctComp.Show '1, Parent

          CYCLES$ = QPTrim$(fpMtrCycle)
          If Len(CYCLES$) = 1 Then CYCLES$ = "0" + CYCLES$
          Route$ = QPTrim$(fpMtrRoute)
          If Len(Route$) = 1 Then
            Route$ = "0" + Route$
            Chkbook = True
          Else
            Chkbook = False
          End If
          Route$ = "0000" + Route$
          GoSub HostDownLoad
          MReading$ = Str$(TCnt) '# of Readings RDG to Expect
          MReading$ = LTrim$(MReading$)
          TReading$ = MReading$
          KReading$ = Str$(KCnt) '# of Keyed Readings RDG to Expect
          KReading$ = LTrim$(KReading$)
          RReading$ = Str$(RCnt) '# of Radio Readings RDG to Expect
          RReading$ = LTrim$(RReading$)
          If Len(MReading$) < 4 Then
            MReading$ = String$(4 - Len(MReading$), "0") + MReading$
          End If
          If Len(TReading$) < 4 Then
            TReading$ = String$(4 - Len(TReading$), "0") + TReading$
          End If
          If Len(KReading$) < 4 Then
            KReading$ = String$(4 - Len(KReading$), "0") + KReading$
          End If
          If Len(RReading$) < 4 Then
            RReading$ = String$(4 - Len(RReading$), "0") + RReading$
          End If

          'Check the len of the reading here
          '
          Reading$ = LTrim$(Str$(TCnt))
          If Val(Reading$) < 1000 Then
            Reading$ = String$(4 - Len(Reading$), "0") + Reading$
          End If
          GoSub BuildDownLoad
          Done = True
        Else
          'GoTo BadDataRestart
        End If
    Close
  Return
HostDownLoad:

    KillFile "UBHHCUS.DAT"
    KillFile "UBHHMTR.DAT"
    KillFile "UBHHRDG.DAT"
    KillFile "UBHHRFF.DAT"
    KillFile "UBHHWRR.DAT"
    LoadUBSetUpFile UBSetUpRec(), UBSetupLen

    'Open Customer Data File
   ' ReDim UBCustRec(1) As NewUBCustRecType
     UBCustRecLen = Len(UBCustRec(1))
    UBFile = FreeFile
    Open "UBCUST.DAT" For Random Shared As UBFile Len = UBCustRecLen

    HighReadPerc = UBSetUpRec(1).HighRead
    LowReadPerc = UBSetUpRec(1).LowRead

    'Open the Correct Order for Reading
    If UBSetUpRec(1).UseSeq = "Y" Then
      IdxRecLen = 4               'we are using a integer
      MakeSequenceIndex "Sequence Number", Me
      IdxNumOfRecs = FileSize&("UBTEMP.IDX") \ IdxRecLen
      ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
      FGetAH "UBTEMP.IDX", IdxBuff(), IdxRecLen, IdxNumOfRecs
    Else
      IdxRecLen = 4               'we are using a integer
      IdxFileSize& = FileSize&("UBCUSTBK.IDX")
      IdxNumOfRecs = IdxFileSize& \ IdxRecLen
        ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
      FGetAH "UBCUSTBK.IDX", IdxBuff(), IdxRecLen, IdxNumOfRecs  'load it
    End If
    cnt = 1

    'Open HandHeld Work Files
    CUSFile = FreeFile
    Open "UBHHCUS.DAT" For Random Shared As CUSFile Len = Len(CUS(1))
    MtrFile = FreeFile
    Open "UBHHMTR.DAT" For Random Shared As MtrFile Len = Len(Mtr(1))
    RdgFile = FreeFile
    Open "UBHHRDG.DAT" For Random Shared As RdgFile Len = Len(Rdg(1))
    RFFFile = FreeFile
    Open "UBHHRFF.DAT" For Random Shared As RFFFile Len = Len(RFF(1))
    WRRFile = FreeFile
    Open "UBHHWRR.DAT" For Random Shared As WRRFile Len = Len(WRR(1))
    'added (2014/01) for Saluda(BluTower HH Company)
    KillFile (Path$ + Route$ + ".tmc")
    UBRpt = FreeFile
    Open Path$ + Route$ + ".tmc" For Output As UBRpt
  
    Do

      Prec& = IdxBuff(cnt).RecNum
      RecNumber = Prec&
       If Not (Prec&) = 0 Then
        Get UBFile, Prec&, UBCustRec(1)
        If UBCustRec(1).DelFlag = 0 Then
        DisplayAcct$ = UBCustRec(1).CustName
        'QPrintRC DisplayAcct$, 15, 35, -1
        'If Chkbook = True Then
          If Not Val(UBCustRec(1).Book) = Val(Route$) Then
            GoTo BadOnes
          End If
       ' End If
        If UBCustRec(1).BILLCYCL = Val(CYCLES$) And (UBCustRec(1).Status <> "F") Then
         'Do SomeThing Here
          'Build Customer Meter Reading & Radio Read File
          TCnt = TCnt + 1
          'Build Customer First

          MeterNumber$ = UBCustRec(1).LocMeters(1).MtrNum
          MNum2File$ = MeterNumber$
          'added (2014/01) the bluserial$ to use in the extra file exported for Saluda(BluTower HH Company)
          BluSerial$ = MNum2File$
          If Left$(MNum2File$, 2) = "WR" Then
            MNum2File$ = Mid$(MNum2File$, 3)
            BluSerial$ = Mid$(MNum2File$, 3)
          End If
          If Left$(MNum2File$, 2) = "RR" Then
            BluSerial$ = Mid$(MNum2File$, 3)
          End If
          MeterSeq$ = Str$(cnt * 10)
          If Len(MeterSeq$) < 8 Then
             MeterSeq$ = String$(8 - Len(MeterSeq$), "0") + MeterSeq$
          End If
          If CitySaluda = True Then MeterSeq$ = "00000000"
          Dials = Val(UBCustRec(1).USERCODE2)
          If Dials = 0 Then Dials = 7             'Default for Caldwell Cty
          DialsS$ = "0" + Right$(Str$(Dials), 1)
          Decimals$ = "00"
          If Perquim = True Or CitySaluda = True Then
            MPIDNO$ = QPTrim$(UBCustRec(1).LocMeters(1).MtrIDNO)
            If Left$(MeterNumber$, 2) = "RR" Then
              If Len(MPIDNO$) > 0 Then
                ERTID$ = MPIDNO$
              Else
                ERTID$ = Mid$(MeterNumber$, 3, 8)
              End If
              ReadMethod$ = "R"
              RCnt = RCnt + 1
            ElseIf Left$(MeterNumber$, 2) = "WR" Then
              If Len(MPIDNO$) > 0 Then
                ERTID$ = MPIDNO$
              Else
                ERTID$ = Mid$(MeterNumber$, 3, 8)
              End If
              ReadMethod$ = "W"
              RCnt = RCnt + 1
            Else
              ERTID$ = "00000000"
              ReadMethod$ = "K"
              KCnt = KCnt + 1
            End If

          'reuse the MeterID$ variable
          'MeterID$ = QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrIDNO)
          'If Len(MeterID$) > 0 Then
          '  UBSenRdRec(1).MtrIDMST = Left$(MeterID$, Len(MeterID$) - 1)
          '  UBSenRdRec(1).MtrIDNO = Right$(MeterID$, 1)
          'End If
          Else
            If Left$(MeterNumber$, 2) = "RR" Then
              ERTID$ = Mid$(MeterNumber$, 3, 8)
              ReadMethod$ = "R"
              RCnt = RCnt + 1
            ElseIf Left$(MeterNumber$, 2) = "WR" Then
              ERTID$ = Mid$(MeterNumber$, 3, 8)
              ReadMethod$ = "W"
              RCnt = RCnt + 1
            Else
              ERTID$ = "00000000"
              ReadMethod$ = "K"
              KCnt = KCnt + 1
            End If
          End If
   'broadway W ^^
          If UBCustRec(1).LocMeters(1).CurRead > 0 Then
            PrevRead$ = LTrim$(Str$(UBCustRec(1).LocMeters(1).CurRead))
          Else
            PrevRead$ = "0"
          End If
          PrevRead$ = String$(10 - Len(PrevRead$), "0") + PrevRead$

          'need to fix these

          If UBCustRec(1).LocMeters(1).AvgUse < 0 Then
            UBCustRec(1).LocMeters(1).AvgUse = 0
          End If
         'HighRead# = FIX(((0# + UBCustRec(1).LocMeters(1).AvgUse) * (HighReadPerc / 100) + UBCustRec(1).LocMeters(1).CurRead))
          HighRead# = Fix((0# + UBCustRec(1).LocMeters(1).AvgUse) * (HighReadPerc / 100))
          If HighRead# <= 1 Then
            HighRead# = 1
          End If

          HighReadS$ = LTrim$(Str$(HighRead#))
          If Len(HighReadS$) < 10 Then
            HighReadS$ = String$(10 - Len(HighReadS$), "0") + HighReadS$
          End If

          LowRead# = 1 'UBCustRec(1).LocMeters(1).CurRead
              LowReadS$ = LTrim$(Str$(LowRead#))
          If Len(LowReadS$) < 10 Then
            LowReadS$ = String$(10 - Len(LowReadS$), "0") + LowReadS$
          End If

'*********************************
          'added (2014/01) for Saluda(BluTower HH Company)

        Print #UBRpt, Str$(126);
        Print #UBRpt, ","; LTrim(RTrim(MPIDNO$));
        Print #UBRpt, ","; LTrim(RTrim(BluSerial$));
        Print #UBRpt, ","; Str$(0)

'********************************
          CUS(1).RecordID = "CUS"
          CUS(1).RouteNumber = CYCLES$ + Route$
          CUS(1).Meters = "001"
          CUS(1).AccountNumber = LTrim$(Str$(RecNumber))
          CUS(1).Name = LTrim$(Left$(UBCustRec(1).CustName, 20))
          CUS(1).Address1 = LTrim$(Left$(UBCustRec(1).ServAddr, 20))
          CUS(1).Address2 = String$(20, " ")
          CUS(1).Reserved = "  "
          CUS(1).Group = " "
          CUS(1).CustInfo = LTrim$(Left$(UBCustRec(1).HHMSG1, 20))
          CUS(1).ExCustRec = "N"
          CUS(1).Segment = "    "
          CUS(1).UtilityID = "WS"
          CUS(1).PassThur = " "
          CUS(1).Pad = " "

          Put CUSFile, TCnt, CUS(1)
          Mtr(1).RecordID = "MTR"
          Mtr(1).RouteNumber = CYCLES$ + Route$
          Mtr(1).Reads = "001"
          Mtr(1).Reserved = "00"
          Mtr(1).Group = "0"
          Mtr(1).Reserved1 = "0"
          Mtr(1).Reserved2 = "0000000"
          Mtr(1).Survey = "0"
          Mtr(1).Reserved3 = "00"
          Mtr(1).Survey2 = "0"
          If Val(UBCustRec(1).CUSTTYPE) = 1 Then
            Mtr(1).BillCode = "R"
          Else
            Mtr(1).BillCode = "C"
          End If
          Mtr(1).MtrStatus = "A"
          Mtr(1).OpticalProbe = String$(14, " ")
           If CitySaluda = True Then
              If Left$(MNum2File$, 2) = "RR" Then
                MNum2File$ = Mid$(MNum2File$, 3)
              End If
            End If
          Mtr(1).MtrNumber = MNum2File$
          Mtr(1).Reserved4 = "00"
          Mtr(1).MeterType = "01" ' 1=Water 2=Elec 3=Gas
          Mtr(1).MeterSeq = MeterSeq$
          Mtr(1).MeterInfo = QPTrim$(UBCustRec(1).USERCODE1)
          Mtr(1).Reserved5 = "0"
          Mtr(1).Location = "00"
          Mtr(1).Reserved6 = "0"
          Mtr(1).ReadInst1 = "0"
          Mtr(1).Reserved7 = "0"
          Mtr(1).ReadInst2 = "00"
          Mtr(1).SpecMessage = "0"
          Mtr(1).Reserved8 = "0"
          Mtr(1).SpecMessageYN = "N"
          Mtr(1).MtrCat = "W"
          Mtr(1).LocExtraMtr = "L"
          Mtr(1).TimeCode = "000"
          Mtr(1).MtrAudit1 = "  "
          Mtr(1).MtrAudit2 = "  "
          Mtr(1).MtrAudit3 = "  "
          Mtr(1).MtrAudit4 = "  "
          Mtr(1).Pad = String$(14, " ")

          Put MtrFile, TCnt, Mtr(1)

          Rdg(1).RecordID = "RDG"
          Rdg(1).RouteNumber = CYCLES$ + Route$
          Rdg(1).Text = "WATR"
          Rdg(1).Prompt = "Y"
          Rdg(1).ReadDir = "L"
          Rdg(1).Compare = "000"
          Rdg(1).Validation = "000"
          Rdg(1).Reserved = "      "   'was a "0" change to spaces on 6/14/06
          Rdg(1).Channel = "00"
          Rdg(1).Dials = DialsS$   ' STRING * 2
          Rdg(1).Decimals = Decimals$             ' STRING * 2
          Rdg(1).ReadMethod = ReadMethod$
          Rdg(1).PrevRead = PrevRead$
          Rdg(1).High1 = HighReadS$                '    AS STRING * 10
          Rdg(1).Low1 = LowReadS$  '    AS STRING * 10
          Rdg(1).MtrConstant = "000000"
          Rdg(1).ConstantFlag = "0"
          Rdg(1).HHFFlag = "0"
          Rdg(1).PosCreep = "00001"
          Rdg(1).Estimates = "0"
          Rdg(1).Reserved2 = "0"
          RType$ = QPTrim$(Mid$(UBCustRec(1).HHMSG3, 1, 2))
          If Len(RType$) > 0 Then
            Rdg(1).ReadType = RType$
          Else
            Rdg(1).ReadType = "01"
          End If
          'vvv Caldwell
          'Rdg(1).ReadType = Dials$

          'change to user2 field
          Rdg(1).MaxPercent = "000000"
          Rdg(1).MinPercent = "000000"
          Rdg(1).NegCreep = "00001"
          Rdg(1).Pad = String$(31, " ")

          Put RdgFile, TCnt, Rdg(1)

          RFF(1).RecordID = "RFF"
          RFF(1).RouteNumber = CYCLES$ + Route$
           If CitySaluda = True Then
            If Len(ERTID$) < 8 Then
              ERTID$ = String$(8 - Len(ERTID$), "0") + ERTID$
            End If
              RFF(1).RFERT = "126" + Mid$(ERTID$, 1, 5)
              RFF(1).Reserved = Mid$(ERTID$, 6, 3) + "000"
              RFF(1).RFProgram = "RMLX"
            Else
              RFF(1).RFERT = ERTID$
              RFF(1).Reserved = "      "   'changed on 6/14/06  from "000000"
              RFF(1).RFProgram = "    "    'change to 4 spaces 6/14/06 from "ERT "
            End If
          RFF(1).Reserved2 = "       "  'change to 7 spaces 6/14/06 from "0000000"
          RFF(1).GeoArea = "00"
          If Perquim = True Then
            RFF(1).RFFreq = "000956.44375" ' Hard Coded for Perquimans
            RFF(1).RFTone = "0007"         ' Hard Coded for Perquimans
          Else
            RFF(1).RFFreq = "000956.25625" ' Hard Coded from Badger
            RFF(1).RFTone = "0012"         ' Hard Coded from Badger  'calddwell
          End If
          RFF(1).Reserved3 = "0000000000"
          RFF(1).Tamper = "17"
          If Alberta = True Then RFF(1).Tamper = "16"
          RFF(1).ConcIndicator = "N"
          RFF(1).Pad = String$(59, " ")

          Put RFFFile, TCnt, RFF(1)

          WRR(1).RecordID = "WRR"
          WRR(1).RouteNumber = CYCLES$ + Route$
          If applachflag = True Then
            WRR(1).DeviceID = MNum2File$
            LSet WRR(1).WandProg = "VERS"
          Else
            WRR(1).DeviceID = ERTID$
            LSet WRR(1).WandProg = "MTR2"
          End If
          LSet WRR(1).Resv1 = ""
          LSet WRR(1).Resv2 = ""
          WRR(1).NodeNumb = "00"
          WRR(1).NoTamper = "Y"
          LSet WRR(1).Fill1 = ""
          Put WRRFile, TCnt, WRR(1)
        End If
BadOnes:
       End If
      End If
      cnt = cnt + 1
    Loop Until cnt > IdxNumOfRecs
    Close
  Return
BuildDownLoad:
    CUSFile = FreeFile
    Open "UBHHCUS.DAT" For Random Shared As CUSFile Len = Len(CUS(1))
    MtrFile = FreeFile
    Open "UBHHMTR.DAT" For Random Shared As MtrFile Len = Len(Mtr(1))
    RdgFile = FreeFile
    Open "UBHHRDG.DAT" For Random Shared As RdgFile Len = Len(Rdg(1))
    RFFFile = FreeFile
    Open "UBHHRFF.DAT" For Random Shared As RFFFile Len = Len(RFF(1))

 

    'File Header
    FHD(1).RecordID = "FHD"
    FHD(1).TABLE = "N"

  '  FHD(1).PROBE = "N"

  'broadway
    FHD(1).PROBE = "Y"

    FHD(1).VERS = "01.00"
    FHD(1).Reserved = "   "
    FHD(1).CYCLES = "01"
    FHD(1).RADIOREADYN = "Y"


  '  FHD(1).WANDYN = "N"

  'broadway
    FHD(1).WANDYN = "Y"

    FHD(1).EXTFORMATYN = "N"
    FHD(1).Pad = String$(108, 32)
    FileHdrLine$ = FHD(1).RecordID + FHD(1).TABLE + FHD(1).PROBE + FHD(1).VERS + FHD(1).Reserved
    FileHdrLine$ = FileHdrLine$ + FHD(1).CYCLES + FHD(1).RADIOREADYN + FHD(1).WANDYN + FHD(1).EXTFORMATYN + FHD(1).Pad

    'Cycle Header
    CHD(1).RecordID = "CHD"
    CHD(1).CycleNumber = CYCLES$
    CHD(1).NumberCycles = "0001"
    CHD(1).CycleDate = Left$(Date$, 2) + Mid$(Date$, 4, 2) + Mid$(Date$, 7, 4)
    CHD(1).Pad = String$(109, 32)
    CycleLine$ = CHD(1).RecordID + CHD(1).CycleNumber + CHD(1).NumberCycles + CHD(1).CycleDate + CHD(1).Pad

    'Route Header
    RHD(1).RecordID = "RHD"
    RHD(1).RouteNumber = CYCLES$ + Route$
    RHD(1).SurveyYN = "N"
    RHD(1).RouteMessYN = "N"
    RHD(1).Keys = "0000"
    RHD(1).Readings = Reading$
    RHD(1).Demands = "0000"
    RHD(1).Keyed = KReading$
    RHD(1).Probed = "0000"
    RHD(1).Radio = RReading$
    RHD(1).Customers = TReading$
    RHD(1).Meters = TReading$
    RHD(1).TimeAllowed = "000000"
    RHD(1).Gas = "0000"
    RHD(1).Water = TReading$
    RHD(1).Electric = "0000"
    RHD(1).Location = "0000"
    RHD(1).Extra = "0000"
    RHD(1).Region = "  "
    RHD(1).ZONE = "  "
    RHD(1).Office = "  "
     'was this
    ' RHD(1).BillCycle = Cycle$

    RHD(1).BillCycle = ""

    RHD(1).DropCycle = "  "
    RHD(1).WandReads = "0000"
    RHD(1).AMR = " "
    RHD(1).Pad = String$(32, " ")
    RouteHeader$ = RHD(1).RecordID + RHD(1).RouteNumber + RHD(1).SurveyYN
    RouteHeader$ = RouteHeader$ + RHD(1).RouteMessYN + RHD(1).Keys + RHD(1).Readings
    RouteHeader$ = RouteHeader$ + RHD(1).Demands + RHD(1).Keyed + RHD(1).Probed
    RouteHeader$ = RouteHeader$ + RHD(1).Radio + RHD(1).Customers + RHD(1).Meters
    RouteHeader$ = RouteHeader$ + RHD(1).TimeAllowed + RHD(1).Gas + RHD(1).Water
    RouteHeader$ = RouteHeader$ + RHD(1).Electric + RHD(1).Location
    RouteHeader$ = RouteHeader$ + RHD(1).Extra + RHD(1).Region + RHD(1).ZONE
    RouteHeader$ = RouteHeader$ + RHD(1).Office + RHD(1).BillCycle
    RouteHeader$ = RouteHeader$ + RHD(1).DropCycle + RHD(1).WandReads
    RouteHeader$ = RouteHeader$ + RHD(1).AMR + RHD(1).Pad

    'Start Build of Test File For Now
    outfile = FreeFile
    Open Path$ + "DOWNLOAD.DAT" For Output As #outfile
    Print #outfile, FileHdrLine$
    Print #outfile, CycleLine$
    Print #outfile, RouteHeader$

    'Now Do Individual Customers, Readings , and Radio Read if necessary

    CUSFile = FreeFile
    Open "UBHHCUS.DAT" For Random Shared As CUSFile Len = Len(CUS(1))
    MtrFile = FreeFile
    Open "UBHHMTR.DAT" For Random Shared As MtrFile Len = Len(Mtr(1))
    RdgFile = FreeFile
    Open "UBHHRDG.DAT" For Random Shared As RdgFile Len = Len(Rdg(1))
    RFFFile = FreeFile
    Open "UBHHRFF.DAT" For Random Shared As RFFFile Len = Len(RFF(1))
    WRRFile = FreeFile
    Open "UBHHWRR.DAT" For Random Shared As WRRFile Len = Len(WRR(1))

    For CusRecord = 1 To TCnt
      FrmShowPctComp.ShowPctComp CusRecord, TCnt&

      Get CUSFile, CusRecord, CUS(1)
      CusLine$ = CUS(1).RecordID + CUS(1).RouteNumber + CUS(1).Meters
      CusLine$ = CusLine$ + CUS(1).AccountNumber + CUS(1).Name
      CusLine$ = CusLine$ + CUS(1).Address1 + CUS(1).Address2 + CUS(1).Reserved
      CusLine$ = CusLine$ + CUS(1).Group + CUS(1).CustInfo + CUS(1).ExCustRec
      CusLine$ = CusLine$ + CUS(1).Segment + CUS(1).UtilityID + CUS(1).PassThur + CUS(1).Pad
      Print #outfile, CusLine$
      Get MtrFile, CusRecord, Mtr(1)
      MtrLine$ = Mtr(1).RecordID + Mtr(1).RouteNumber + Mtr(1).Reads + Mtr(1).Reserved
      MtrLine$ = MtrLine$ + Mtr(1).Group + Mtr(1).Reserved1 + Mtr(1).Reserved2
      MtrLine$ = MtrLine$ + Mtr(1).Survey + Mtr(1).Reserved3 + Mtr(1).Survey2
      MtrLine$ = MtrLine$ + Mtr(1).BillCode + Mtr(1).MtrStatus + Mtr(1).OpticalProbe
      MtrLine$ = MtrLine$ + Mtr(1).MtrNumber + Mtr(1).Reserved4 + Mtr(1).MeterType
      MtrLine$ = MtrLine$ + Mtr(1).MeterSeq + Mtr(1).MeterInfo + Mtr(1).Reserved5
      MtrLine$ = MtrLine$ + Mtr(1).Location + Mtr(1).Reserved6 + Mtr(1).ReadInst1
      MtrLine$ = MtrLine$ + Mtr(1).Reserved7 + Mtr(1).ReadInst2 + Mtr(1).SpecMessage + Mtr(1).Reserved8
      MtrLine$ = MtrLine$ + Mtr(1).SpecMessageYN + Mtr(1).MtrCat + Mtr(1).LocExtraMtr + Mtr(1).TimeCode
      MtrLine$ = MtrLine$ + Mtr(1).MtrAudit1 + Mtr(1).MtrAudit2 + Mtr(1).MtrAudit3
      MtrLine$ = MtrLine$ + Mtr(1).MtrAudit4 + Mtr(1).Pad
      Print #outfile, MtrLine$
      Get RdgFile, CusRecord, Rdg(1)
      RdgLine$ = Rdg(1).RecordID + Rdg(1).RouteNumber + Rdg(1).Text + Rdg(1).Prompt
      RdgLine$ = RdgLine$ + Rdg(1).ReadDir + Rdg(1).Compare + Rdg(1).Validation + Rdg(1).Reserved
      RdgLine$ = RdgLine$ + Rdg(1).Channel + Rdg(1).Dials + Rdg(1).Decimals
      RdgLine$ = RdgLine$ + Rdg(1).ReadMethod + Rdg(1).PrevRead + Rdg(1).High1
      RdgLine$ = RdgLine$ + Rdg(1).Low1 + Rdg(1).MtrConstant + Rdg(1).ConstantFlag
      RdgLine$ = RdgLine$ + Rdg(1).HHFFlag + Rdg(1).PosCreep + Rdg(1).Estimates
      RdgLine$ = RdgLine$ + Rdg(1).Reserved2 + Rdg(1).ReadType + Rdg(1).MaxPercent
      RdgLine$ = RdgLine$ + Rdg(1).MinPercent + Rdg(1).NegCreep + Rdg(1).Pad
      Print #outfile, RdgLine$
      If Rdg(1).ReadMethod = "R" Then
        Get RFFFile, CusRecord, RFF(1)
        RffLine$ = RFF(1).RecordID + RFF(1).RouteNumber + RFF(1).RFERT
        RffLine$ = RffLine$ + RFF(1).Reserved + RFF(1).RFProgram + RFF(1).Reserved2
        RffLine$ = RffLine$ + RFF(1).GeoArea + RFF(1).RFFreq + RFF(1).RFTone
        RffLine$ = RffLine$ + RFF(1).Reserved3 + RFF(1).Tamper + RFF(1).ConcIndicator
        RffLine$ = RffLine$ + RFF(1).Pad
        Print #outfile, RffLine$
      End If
      If Rdg(1).ReadMethod = "W" Then
        Get WRRFile, CusRecord, WRR(1)
        WRRLine$ = WRR(1).RecordID + WRR(1).RouteNumber + WRR(1).DeviceID
        WRRLine$ = WRRLine$ + WRR(1).WandProg + WRR(1).Resv1 + WRR(1).Resv2
        WRRLine$ = WRRLine$ + WRR(1).NodeNumb + WRR(1).NoTamper + WRR(1).Fill1
        Print #outfile, WRRLine$
      End If
     Next CusRecord
    If Not TCnt > 0 Then
      Unload FrmShowPctComp
      MsgBox "No Readings to Export", vbOKOnly, "No Reads Found"
    Else
      MsgBox "Readings Exported - " & Str(TCnt), vbOKOnly, "Procedure Complete"
    End If

    'File Trailor
    FHD(1).RecordID = "FTR"
    FHD(1).TABLE = "N"
    FHD(1).PROBE = "N"
    FHD(1).VERS = "01.00"
    FHD(1).Reserved = "   "
    FHD(1).CYCLES = "01"
    FHD(1).RADIOREADYN = "Y"
    FHD(1).WANDYN = "Y"
    FHD(1).EXTFORMATYN = "N"
    FHD(1).Pad = String$(108, 32)
    FileHdrLine$ = FHD(1).RecordID + FHD(1).TABLE + FHD(1).PROBE + FHD(1).VERS + FHD(1).Reserved
    FileHdrLine$ = FileHdrLine$ + FHD(1).CYCLES + FHD(1).RADIOREADYN + FHD(1).WANDYN + FHD(1).EXTFORMATYN + FHD(1).Pad

    'Cycle Trailor
    CHD(1).RecordID = "CTR"
    CHD(1).CycleNumber = CYCLES$
    CHD(1).NumberCycles = "0001"
    CHD(1).CycleDate = Left$(Date$, 2) + Mid$(Date$, 4, 2) + Mid$(Date$, 7, 4)
    CHD(1).Pad = String$(109, 32)
    CycleLine$ = CHD(1).RecordID + CHD(1).CycleNumber + CHD(1).NumberCycles + CHD(1).CycleDate + CHD(1).Pad

    'Route Trailor
    RHD(1).RecordID = "RTR"
    RHD(1).RouteNumber = CYCLES$ + Route$
    RHD(1).SurveyYN = "N"
    RHD(1).RouteMessYN = "N"
    RHD(1).Keys = "0000"
    RHD(1).Readings = Reading$
    RHD(1).Demands = "0000"
    RHD(1).Keyed = KReading$
    RHD(1).Probed = "0000"
    RHD(1).Radio = RReading$
    RHD(1).Customers = TReading$
    RHD(1).Meters = TReading$
    RHD(1).TimeAllowed = "000000"
    RHD(1).Gas = "0000"
    RHD(1).Water = TReading$
    RHD(1).Electric = "0000"
    RHD(1).Location = "0000"
    RHD(1).Extra = "0000"
     RHD(1).Region = "  "
    RHD(1).ZONE = "  "
    RHD(1).Office = "  "
    RHD(1).BillCycle = CYCLES$
    RHD(1).DropCycle = "  "
    RHD(1).WandReads = "0000"
    RHD(1).AMR = " "
    RHD(1).Pad = String$(32, " ")
    RouteHeader$ = RHD(1).RecordID + RHD(1).RouteNumber + RHD(1).SurveyYN
    RouteHeader$ = RouteHeader$ + RHD(1).RouteMessYN + RHD(1).Keys + RHD(1).Readings
    RouteHeader$ = RouteHeader$ + RHD(1).Demands + RHD(1).Keyed + RHD(1).Probed
    RouteHeader$ = RouteHeader$ + RHD(1).Radio + RHD(1).Customers + RHD(1).Meters
    RouteHeader$ = RouteHeader$ + RHD(1).TimeAllowed + RHD(1).Gas + RHD(1).Water
    RouteHeader$ = RouteHeader$ + RHD(1).Electric + RHD(1).Location
    RouteHeader$ = RouteHeader$ + RHD(1).Extra + RHD(1).Region + RHD(1).ZONE
    RouteHeader$ = RouteHeader$ + RHD(1).Office + RHD(1).BillCycle
    RouteHeader$ = RouteHeader$ + RHD(1).DropCycle + RHD(1).WandReads
    RouteHeader$ = RouteHeader$ + RHD(1).AMR + RHD(1).Pad
    Print #outfile, RouteHeader$       'Actually Trailors
    Print #outfile, CycleLine$
    Print #outfile, FileHdrLine$
    Close
  Return
CheckInfo:

    InfoOK = True

    NewPath$ = QPTrim$(fpPathOut.Text)
    Path$ = GetBadgerPath$

    If Right$(NewPath$, 1) <> "\" Then
      NewPath$ = NewPath$ + "\"
    End If

  'if this is a chage to the path to badger then update ubdata
    If NewPath$ <> Path$ Then
  '    PathFile = FreeFile
  '    Open PathData$ For Random Shared As PathFile
  '    PathFile.PathWay = NewPath$
  '    Put #PathFile, 1, PathFile
  '    Close PathFile
      Path$ = NewPath$
    End If
   ' Path$ = GetBadgerPath$
    If InStr(Path$, "BAD") > 0 Then
      'DisplaySUScrn "Invalid"
      InfoOK = False
      GoTo NoGoodRet
    End If

  '  If Not Val(fpMtrCycle) > 0 And Val(fpMtrRoute) > 0 Then
  ''    Color 12
  ''    LOCATE 11, 40
  ''    Print "Invalid Operation"
  ''    Beep
  ''    WaitForAction
  ''    Color 15
  ''    LOCATE 11, 40
  ''    Print String$(25, 32)
  ''    frm(1).FldNo = 1
  ''    Action = 1
  '    Done = False
  '    InfoOK = False
  '  End If

NoGoodRet:
  Return
  '**************************************************
GetBadgerReadings:
    ReDim Rdgg(1) As BADReadRecType
    RDGLen = Len(Rdgg(1))
    Path$ = GetBadgerPath$

    If Len(Path$) > 0 And Right$(Path$, 1) <> "\" Then
      Path$ = Path$ + "\"
    End If

    FileNam$ = Path$ + "UPLOAD.DAT"

        GoSub CheckBadgerPath
        If InfoOK = False Then
          Exit Sub
        End If
      FrmShowPctComp.Label1 = "Importing Reading Information."
      FrmShowPctComp.cmdCancel.Enabled = False
      FrmShowPctComp.Show '1, Parent

    UBCustRecLen = Len(UBCustRec(1))
    UBFile = FreeFile
    Open "UBCUST.DAT" For Random Shared As UBFile Len = UBCustRecLen

    ReadFile = FreeFile
    Open FileNam$ For Random Shared As ReadFile Len = RDGLen
    NumReadRecs = LOF(ReadFile) / RDGLen

    If NumReadRecs = 0 Then
      Close
  '    For zz = 8 To 16
  '      QPrintRC Space$(49), zz, 16, 15
  '    Next
  '    DisplaySUScrn "BADMTRUP"
  '    QPrintRC " ERROR: NO READINGS FOUND", 9, 27, -1
  '    QPrintRC " Check the Badger Path and ", 10, 24, -1
  '    QPrintRC " try the transfer again!!!", 11, 24, -1
  '    QPrintRC " Press any key to continue. ", 13, 24, 14
  '    Beep
      Unload FrmShowPctComp
      MsgBox "ERROR: NO READINGS FOUND", vbOKOnly, "Error"
    End If
  '    BlockClear
  '  DisplaySUScrn "BADMTRUP"
  '
  '  QPrintRC "  Processing Meter Readings.  ", 9, 25, 15
  '  QPrintRC "    Updated Count:            ", 11, 25, 11

    CustRec& = 0
    GotCust = False
    For cnt& = 1 To NumReadRecs
      FrmShowPctComp.ShowPctComp cnt&, NumReadRecs
      Get ReadFile, cnt&, Rdgg(1)
      RecordID$ = Left$(Rdgg(1).Fill, 3)
      Select Case RecordID$
      Case "CUS"
        CustRec& = GetAccount&(Mid$(Rdgg(1).Fill, 15, 20))
        CustName$ = QPTrim$(Mid$(Rdgg(1).Fill, 35, 20))
      Case "RDG"
        GoSub ProcessReading
        CustRec& = 0
      Case Else
        'ignore all others
      End Select
    Next
    Close
    MsgBox "Number of Records Retrieved - " & Str(UpdCnt), vbOKOnly, "Completed"
  Return


ProcessReading:
    If CustRec& <= 0 Then
      GoTo BadCustRet
    End If
    Get UBFile, CustRec&, UBCustRec(1)

    UpdCnt = UpdCnt + 1
   ' QPrintRC Str$(UpdCnt!), 11, 44, 15

  '  MeterRecord = VAL(RIGHT$((QPTrim$(UBPC3000GetRDRec(1).Account)), 1))
    MeterRecord = 1
    ' Check Meter Updated Flag
    ' Update Meter W/Reading
    RDate$ = Mid$(Rdgg(1).Fill, 48, 8)
    ReadDate = Date2Num(Left$(RDate$, 2) + "/" + Mid$(RDate$, 3, 2) + "/" + Right$(RDate$, 4))
  '  If Perquim = True Then
  '    CurReading# = GetReading&(Mid$(Rdgg(1).Fill, 32, 10))
  '  Else
      CurReading# = GetReading&(Mid$(Rdgg(1).Fill, 34, 10))
  '  End If
   ' CurReading# = CurReading# * 100
    'IF CurReading# > 0 THEN STOP

    If UBCustRec(1).LocMeters(MeterRecord).ReadFlag = "Y" Then
      UBCustRec(1).LocMeters(MeterRecord).CurRead = CurReading#
      UBCustRec(1).LocMeters(MeterRecord).CurDate = ReadDate
    Else
      UBCustRec(1).LocMeters(MeterRecord).PrevRead = UBCustRec(1).LocMeters(MeterRecord).CurRead
      UBCustRec(1).LocMeters(MeterRecord).PastDate = UBCustRec(1).LocMeters(MeterRecord).CurDate
      UBCustRec(1).LocMeters(MeterRecord).ReadFlag = "Y"
      UBCustRec(1).LocMeters(MeterRecord).CurDate = ReadDate
      UBCustRec(1).LocMeters(MeterRecord).CurRead = CurReading#
    End If

    Put UBFile, CustRec&, UBCustRec(1)

BadCustRet:

  Return

  '**************************************************
CheckBadgerPath:
    InfoOK = True
    NewPath$ = QPTrim$(fpPathOut.Text)
    Path$ = GetBadgerPath$

    If Right$(NewPath$, 1) <> "\" Then
      NewPath$ = NewPath$ + "\"
    End If

  'if this is a change to the path to badger then update ubdata
    If NewPath$ <> Path$ Then
  '    PathFile = FreeFile
  '    Open PathData$ For Random Shared As PathFile
  '    Print #PathFile, NewPath$
  '    Close PathFile
      Path$ = NewPath$
    End If
  '  Path$ = GetBadgerPath$

    If Right$(Path$, 1) <> "\" Then
      Path$ = Path$ + "\"
    End If
    FileNam$ = Path$ + "UPLOAD.DAT"
    If InStr(Path$, "BAD") > 0 Then
  '    BlockClear
  '    DisplaySUScrn "Invalid"
  '    WaitForAction
      InfoOK = False
    End If

  Return
GiveItUp:
    Close
    Unload FrmShowPctComp
    MsgBox "Errors were encountered, Please contact software support for assistance.", vbOKOnly, "ERRORS"
  End Sub
  'This code is if need multiple meters per customer tested 9/12/06 for Wadesboro
  'Changed 09/01/06 and 9/12/06 need the import and export changed for mult meter possibilities
Private Sub ImpExpRadioReadHHInfo2(ByVal ImpExpFlag As Boolean)
  'Const PathData$ = "UBBAGPTH.DAT"
  Dim Path As String, RDGLen As Integer, FileNam As String
  Dim UBFile As Integer, ReadFile As Integer, NumReadRecs As Long
  Dim zz As Integer, CustRec As Long, GotCust As Boolean
  Dim cnt As Long, RecordID As String, CustName As String
  Dim UpdCnt As Long, MeterRecord As Long, RDate As String
  Dim ReadDate As Integer, CurReading As Double, InfoOK As Boolean
  Dim NewPath As String, PathFile As Integer, q As String
  Dim cb As String, BroadFlag As Boolean, CYCLES As String
  Dim Route As String, TCnt As Long, MReading As String
  Dim TReading As String, KCnt As Long, KReading As String
  Dim RReading As String, Reading As String, Done As Boolean
  Dim HighReadPerc As Double, LowReadPerc As Double, CUSFile As Integer
  Dim MtrFile As Integer, RdgFile As Integer, RFFFile As Integer
  Dim WRRFile As Integer, Prec As Long, RecNumber As Long
  Dim DisplayAcct As String, MeterNumber As String, MNum2File As String
  Dim MeterSeq As String, Dials As Integer, DialsS As String
  Dim Decimals As String, ERTID As String, ReadMethod As String
  Dim PrevRead As String, HighRead As Double, HighReadS As String
  Dim LowRead As Double, LowReadS As String, FileHdrLine As String
  Dim CycleLine As String, RouteHeader As String, outfile As Integer, UBRpt As Integer
  Dim CusRecord As Long, CusLine As String, MtrLine As String
  Dim RdgLine As String, RffLine As String, WRRLine As String, MtrNumber$
  Dim Chkbook As Boolean, RType As String, MPIDNO As String, metercnt As Integer, BluSerial As String
    '''' C:\mvrs50\xfer
  'PathData$ = "UBBAGPTH.DAT"
  On Local Error GoTo GiveItUp
  Chkbook = False
   If InStr(TOWNNAME$, "BROADWAY") Then
    BroadFlag = True
  End If

  PathData$ = UBHHPathWayFile
  If ImpExpFlag Then     'EXPORTING METER READINGS
    GoSub RadioReadIF
  Else
    GoSub GetBadgerReadings
  End If

   Close

AbortExit:

  Exit Sub

RadioReadIF:
  q$ = Chr$(34)
  cb$ = Space$(45)

 
  Path$ = GetBadgerPath$

 ' ReDim Route(100)
  ReDim FHD(1) As FDRTYPE
  ReDim CHD(1) As CDRTYPE
  ReDim RHD(1) As RDRType
  ReDim CUS(1) As CUSType
  ReDim Mtr(1) As MTRType
  ReDim Rdg(1) As RDGType
  ReDim RFF(1) As RFFType
  ReDim WRR(1) As WRRType


BadDataRestart:

      GoSub CheckInfo
      If InfoOK Then
        FrmShowPctComp.Label1 = "Exporting Reading Information."
        FrmShowPctComp.cmdCancel.Enabled = False
        FrmShowPctComp.Show '1, Parent

        CYCLES$ = QPTrim$(fpMtrCycle)
        If Len(CYCLES$) = 1 Then CYCLES$ = "0" + CYCLES$
        Route$ = QPTrim$(fpMtrRoute)
        If Len(Route$) = 1 Then
          Route$ = "0" + Route$
          Chkbook = True
        Else
          Chkbook = False
        End If
        Route$ = "0000" + Route$
        GoSub HostDownLoad
        MReading$ = Str$(TCnt) '# of Readings RDG to Expect
        MReading$ = LTrim$(MReading$)
        TReading$ = MReading$
        KReading$ = Str$(KCnt) '# of Keyed Readings RDG to Expect
        KReading$ = LTrim$(KReading$)
        RReading$ = Str$(RCnt) '# of Radio Readings RDG to Expect
        RReading$ = LTrim$(RReading$)
        If Len(MReading$) < 4 Then
          MReading$ = String$(4 - Len(MReading$), "0") + MReading$
        End If
        If Len(TReading$) < 4 Then
          TReading$ = String$(4 - Len(TReading$), "0") + TReading$
        End If
        If Len(KReading$) < 4 Then
          KReading$ = String$(4 - Len(KReading$), "0") + KReading$
        End If
        If Len(RReading$) < 4 Then
          RReading$ = String$(4 - Len(RReading$), "0") + RReading$
        End If

        'Check the len of the reading here
        '
        Reading$ = LTrim$(Str$(TCnt))
        If Val(Reading$) < 1000 Then
          Reading$ = String$(4 - Len(Reading$), "0") + Reading$
        End If
        GoSub BuildDownLoad
        Done = True
      Else
        'GoTo BadDataRestart
      End If
  Close
Return
HostDownLoad:

  KillFile "UBHHCUS.DAT"
  KillFile "UBHHMTR.DAT"
  KillFile "UBHHRDG.DAT"
  KillFile "UBHHRFF.DAT"
  KillFile "UBHHWRR.DAT"
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen

  'Open Customer Data File
 ' ReDim UBCustRec(1) As NewUBCustRecType
   UBCustRecLen = Len(UBCustRec(1))
  UBFile = FreeFile
  Open "UBCUST.DAT" For Random Shared As UBFile Len = UBCustRecLen

  HighReadPerc = UBSetUpRec(1).HighRead
  LowReadPerc = UBSetUpRec(1).LowRead

  'Open the Correct Order for Reading
  If UBSetUpRec(1).UseSeq = "Y" Then
    IdxRecLen = 4               'we are using an integer
    MakeSequenceIndex "Sequence Number", Me
    IdxNumOfRecs = FileSize&("UBTEMP.IDX") \ IdxRecLen
    ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
    FGetAH "UBTEMP.IDX", IdxBuff(), IdxRecLen, IdxNumOfRecs
  Else
    IdxRecLen = 4               'we are using an integer
    IdxFileSize& = FileSize&("UBCUSTBK.IDX")
    IdxNumOfRecs = IdxFileSize& \ IdxRecLen
      ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
    FGetAH "UBCUSTBK.IDX", IdxBuff(), IdxRecLen, IdxNumOfRecs  'load it
  End If
  cnt = 1

  'Open HandHeld Work Files
  CUSFile = FreeFile
  Open "UBHHCUS.DAT" For Random Shared As CUSFile Len = Len(CUS(1))
  MtrFile = FreeFile
  Open "UBHHMTR.DAT" For Random Shared As MtrFile Len = Len(Mtr(1))
  RdgFile = FreeFile
  Open "UBHHRDG.DAT" For Random Shared As RdgFile Len = Len(Rdg(1))
  RFFFile = FreeFile
  Open "UBHHRFF.DAT" For Random Shared As RFFFile Len = Len(RFF(1))
  WRRFile = FreeFile
  Open "UBHHWRR.DAT" For Random Shared As WRRFile Len = Len(WRR(1))
     'added (2014/01) for Saluda(BluTower HH Company)
    KillFile (Path$ + CYCLES$ + Route$ + ".tmc")
    UBRpt = FreeFile
    Open Path$ + CYCLES$ + Route$ + ".tmc" For Output As UBRpt
  Do

    Prec& = IdxBuff(cnt).RecNum
    RecNumber = Prec&
     If Not (Prec&) = 0 Then
      Get UBFile, Prec&, UBCustRec(1)
      If UBCustRec(1).DelFlag = 0 Then
      DisplayAcct$ = UBCustRec(1).CustName
      'QPrintRC DisplayAcct$, 15, 35, -1
      'If Chkbook = True Then
        If Not Val(UBCustRec(1).Book) = Val(Route$) Then
          GoTo BadOnes
        End If
     ' End If
      If UBCustRec(1).BILLCYCL = Val(CYCLES$) And (UBCustRec(1).Status <> "F") Then
       'Do SomeThing Here
        'Build Customer Meter Reading & Radio Read File

        'Build Customer First
        CUS(1).RecordID = "CUS"
        CUS(1).RouteNumber = CYCLES$ + Route$
        CUS(1).Meters = "001"
        CUS(1).AccountNumber = LTrim$(Str$(RecNumber))
        CUS(1).Name = LTrim$(Left$(UBCustRec(1).CustName, 20))
        CUS(1).Address1 = LTrim$(Left$(UBCustRec(1).ServAddr, 20))
        CUS(1).Address2 = String$(20, " ")
        CUS(1).Reserved = "  "
        CUS(1).Group = " "
        CUS(1).CustInfo = LTrim$(Left$(UBCustRec(1).HHMSG1, 20))
        CUS(1).ExCustRec = "N"
        CUS(1).Segment = "    "
        CUS(1).UtilityID = "WS"
        CUS(1).PassThur = " "
        CUS(1).Pad = " "
        For metercnt = 1 To 7
          If Len(QPTrim$(UBCustRec(1).LocMeters(metercnt).MtrNum)) > 0 Then
            TCnt = TCnt + 1
            Put CUSFile, TCnt, CUS(1)
            MeterNumber$ = UBCustRec(1).LocMeters(metercnt).MtrNum
            MNum2File$ = MeterNumber$
             'added (2014/01) the bluserial$ to use in the extra file exported for Saluda(BluTower HH Company)
          BluSerial$ = MNum2File$
          If Left$(MNum2File$, 2) = "WR" Then
            MNum2File$ = Mid$(MNum2File$, 3)
            BluSerial$ = Mid$(MNum2File$, 3)
          End If
          If Left$(MNum2File$, 2) = "RR" Then
            BluSerial$ = Mid$(MNum2File$, 3)
          End If
       
            MeterSeq$ = Str$(cnt * 10)
            If Len(MeterSeq$) < 8 Then
               MeterSeq$ = String$(8 - Len(MeterSeq$), "0") + MeterSeq$
            End If
            If CitySaluda = True Then MeterSeq$ = "00000000"
            Dials = Val(UBCustRec(1).USERCODE2)
            If Dials = 0 Then Dials = 7             'Default for Caldwell Cty
            DialsS$ = "0" + Right$(Str$(Dials), 1)
            Decimals$ = "00"
            If Perquim = True Or CitySaluda = True Then
              MPIDNO$ = QPTrim$(UBCustRec(1).LocMeters(metercnt).MtrIDNO)
              If Left$(MeterNumber$, 2) = "RR" Then
                If Len(MPIDNO$) > 0 Then
                  ERTID$ = MPIDNO$
                Else
                  ERTID$ = Mid$(MeterNumber$, 3, 8)
                End If
                ReadMethod$ = "R"
                RCnt = RCnt + 1
              ElseIf Left$(MeterNumber$, 2) = "WR" Then
                If Len(MPIDNO$) > 0 Then
                  ERTID$ = MPIDNO$
                Else
                  ERTID$ = Mid$(MeterNumber$, 3, 8)
                End If
                ReadMethod$ = "W"
                RCnt = RCnt + 1
              Else
                ERTID$ = "00000000"
                ReadMethod$ = "K"
                KCnt = KCnt + 1
              End If

            'reuse the MeterID$ variable
            'MeterID$ = QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrIDNO)
            'If Len(MeterID$) > 0 Then
            '  UBSenRdRec(1).MtrIDMST = Left$(MeterID$, Len(MeterID$) - 1)
            '  UBSenRdRec(1).MtrIDNO = Right$(MeterID$, 1)
            'End If
            Else
            MPIDNO$ = QPTrim$(UBCustRec(1).LocMeters(metercnt).MtrIDNO)
              If Left$(MeterNumber$, 2) = "RR" Then
                ERTID$ = Mid$(MeterNumber$, 3, 8)
                ReadMethod$ = "R"
                RCnt = RCnt + 1
              ElseIf Left$(MeterNumber$, 2) = "WR" Then
                ERTID$ = Mid$(MeterNumber$, 3, 8)
                ReadMethod$ = "W"
                RCnt = RCnt + 1
              Else
                ERTID$ = "00000000"
                ReadMethod$ = "K"
                KCnt = KCnt + 1
              End If
            End If
     'broadway W ^^
            If UBCustRec(1).LocMeters(metercnt).CurRead > 0 Then
              PrevRead$ = LTrim$(Str$(UBCustRec(1).LocMeters(metercnt).CurRead))
            Else
              PrevRead$ = "0"
            End If
            PrevRead$ = String$(10 - Len(PrevRead$), "0") + PrevRead$

            'need to fix these

            If UBCustRec(1).LocMeters(metercnt).AvgUse < 0 Then
              UBCustRec(1).LocMeters(metercnt).AvgUse = 0
            End If
           'HighRead# = FIX(((0# + UBCustRec(1).LocMeters(1).AvgUse) * (HighReadPerc / 100) + UBCustRec(1).LocMeters(1).CurRead))
            HighRead# = Fix((0# + UBCustRec(1).LocMeters(metercnt).AvgUse) * (HighReadPerc / 100))
            If HighRead# <= 1 Then
              HighRead# = 1
            End If

            HighReadS$ = LTrim$(Str$(HighRead#))
            If Len(HighReadS$) < 10 Then
              HighReadS$ = String$(10 - Len(HighReadS$), "0") + HighReadS$
            End If

            LowRead# = 1 'UBCustRec(1).LocMeters(1).CurRead
                LowReadS$ = LTrim$(Str$(LowRead#))
            If Len(LowReadS$) < 10 Then
              LowReadS$ = String$(10 - Len(LowReadS$), "0") + LowReadS$
            End If

            Mtr(1).RecordID = "MTR"
            Mtr(1).RouteNumber = CYCLES$ + Route$
            Mtr(1).Reads = "001"
            Mtr(1).Reserved = "00"
            Mtr(1).Group = "0"
            Mtr(1).Reserved1 = "0"
            Mtr(1).Reserved2 = "0000000"
            Mtr(1).Survey = "0"
            Mtr(1).Reserved3 = "00"
            Mtr(1).Survey2 = "0"
            If Val(UBCustRec(1).CUSTTYPE) = 1 Then
              Mtr(1).BillCode = "R"
            Else
              Mtr(1).BillCode = "C"
            End If
            Mtr(1).MtrStatus = "A"
            Mtr(1).OpticalProbe = String$(14, " ")
            If CitySaluda = True Then
              If Left$(MNum2File$, 2) = "RR" Then
                MNum2File$ = Mid$(MNum2File$, 3)
              End If
            End If
            Mtr(1).MtrNumber = MNum2File$
            Mtr(1).Reserved4 = "00"
            Mtr(1).MeterType = "01" ' 1=Water 2=Elec 3=Gas
            Mtr(1).MeterSeq = MeterSeq$
            Mtr(1).MeterInfo = QPTrim$(UBCustRec(1).USERCODE1)
            Mtr(1).Reserved5 = "0"
            Mtr(1).Location = "00"
            Mtr(1).Reserved6 = "0"
            Mtr(1).ReadInst1 = "0"
            Mtr(1).Reserved7 = "0"
            Mtr(1).ReadInst2 = "00"
            Mtr(1).SpecMessage = "0"
            Mtr(1).Reserved8 = "0"
            Mtr(1).SpecMessageYN = "N"
            Mtr(1).MtrCat = "W"
            Mtr(1).LocExtraMtr = "L"
            Mtr(1).TimeCode = "000"
            Mtr(1).MtrAudit1 = "  "
            Mtr(1).MtrAudit2 = "  "
            Mtr(1).MtrAudit3 = "  "
            Mtr(1).MtrAudit4 = "  "
            Mtr(1).Pad = String$(14, " ")

            Put MtrFile, TCnt, Mtr(1)
            
 '*********************************
     'added (2014/01) for Saluda(BluTower HH Company)

        Print #UBRpt, Str$(126);
        Print #UBRpt, ","; LTrim(RTrim(MPIDNO$));
        Print #UBRpt, ","; LTrim(RTrim(BluSerial$));
        Print #UBRpt, ","; Str$(1)

'********************************

            Rdg(1).RecordID = "RDG"
            Rdg(1).RouteNumber = CYCLES$ + Route$
            Rdg(1).Text = "WATR"
            Rdg(1).Prompt = "Y"
            Rdg(1).ReadDir = "L"
            Rdg(1).Compare = "000"
            Rdg(1).Validation = "000"
            Rdg(1).Reserved = "      "   'was a "0" change to spaces on 6/14/06
            Rdg(1).Channel = "00"
            Rdg(1).Dials = DialsS$   ' STRING * 2
            Rdg(1).Decimals = Decimals$             ' STRING * 2
            Rdg(1).ReadMethod = ReadMethod$
            Rdg(1).PrevRead = PrevRead$
            Rdg(1).High1 = HighReadS$                '    AS STRING * 10
            Rdg(1).Low1 = LowReadS$  '    AS STRING * 10
            Rdg(1).MtrConstant = "000000"
            Rdg(1).ConstantFlag = "0"
            Rdg(1).HHFFlag = "0"
            Rdg(1).PosCreep = "00001"
            Rdg(1).Estimates = "0"
            Rdg(1).Reserved2 = "0"
            RType$ = QPTrim$(Mid$(UBCustRec(1).HHMSG3, 1, 2))
            If Len(RType$) > 0 Then
              Rdg(1).ReadType = RType$
            Else
              Rdg(1).ReadType = "01"
            End If
            'vvv Caldwell
            'Rdg(1).ReadType = Dials$

            'change to user2 field
            Rdg(1).MaxPercent = "000000"
            Rdg(1).MinPercent = "000000"
            Rdg(1).NegCreep = "00001"
            Rdg(1).Pad = String$(31, " ")

            Put RdgFile, TCnt, Rdg(1)

            RFF(1).RecordID = "RFF"
            RFF(1).RouteNumber = CYCLES$ + Route$
             If CitySaluda = True Then
            If Len(ERTID$) < 8 Then
              ERTID$ = String$(8 - Len(ERTID$), "0") + ERTID$
            End If
              RFF(1).RFERT = "126" + Mid$(ERTID$, 1, 5)
              RFF(1).Reserved = Mid$(ERTID$, 6, 3) + "000"
              RFF(1).RFProgram = "RMLX"
            Else
              RFF(1).RFERT = ERTID$
              RFF(1).Reserved = "      "   'changed on 6/14/06  from "000000"
              RFF(1).RFProgram = "    "    'change to 4 spaces 6/14/06 from "ERT "
            End If
            RFF(1).Reserved2 = "       "  'change to 7 spaces 6/14/06 from "0000000"
            RFF(1).GeoArea = "00"
            If Perquim = True Then
              RFF(1).RFFreq = "000956.44375" ' Hard Coded for Perquimans
              RFF(1).RFTone = "0007"         ' Hard Coded for Perquimans
            Else
              RFF(1).RFFreq = "000956.25625" ' Hard Coded from Badger
              RFF(1).RFTone = "0012"         ' Hard Coded from Badger  'calddwell
            End If
            RFF(1).Reserved3 = "0000000000"
            RFF(1).Tamper = "17"
            If Alberta = True Then RFF(1).Tamper = "16"
            RFF(1).ConcIndicator = "N"
            RFF(1).Pad = String$(59, " ")

            Put RFFFile, TCnt, RFF(1)

            WRR(1).RecordID = "WRR"
            WRR(1).RouteNumber = CYCLES$ + Route$
            If applachflag = True Then
              WRR(1).DeviceID = MNum2File$
              LSet WRR(1).WandProg = "VERS"
            Else
              WRR(1).DeviceID = ERTID$
              LSet WRR(1).WandProg = "MTR2"
            End If
            LSet WRR(1).Resv1 = ""
            LSet WRR(1).Resv1 = ""
            WRR(1).NodeNumb = "00"
            WRR(1).NoTamper = "Y"
            LSet WRR(1).Fill1 = ""
            Put WRRFile, TCnt, WRR(1)
        End If
        'End If
      Next metercnt
BadOnes:
     End If
     End If
    End If
    cnt = cnt + 1
  Loop Until cnt > IdxNumOfRecs
  Close
Return
BuildDownLoad:
  CUSFile = FreeFile
  Open "UBHHCUS.DAT" For Random Shared As CUSFile Len = Len(CUS(1))
  MtrFile = FreeFile
  Open "UBHHMTR.DAT" For Random Shared As MtrFile Len = Len(Mtr(1))
  RdgFile = FreeFile
  Open "UBHHRDG.DAT" For Random Shared As RdgFile Len = Len(Rdg(1))
  RFFFile = FreeFile
  Open "UBHHRFF.DAT" For Random Shared As RFFFile Len = Len(RFF(1))

  'File Header
  FHD(1).RecordID = "FHD"
  FHD(1).TABLE = "N"

'  FHD(1).PROBE = "N"

'broadway
  FHD(1).PROBE = "Y"

  FHD(1).VERS = "01.00"
  FHD(1).Reserved = "   "
  FHD(1).CYCLES = "01"
  FHD(1).RADIOREADYN = "Y"


'  FHD(1).WANDYN = "N"

'broadway
  FHD(1).WANDYN = "Y"

  FHD(1).EXTFORMATYN = "N"
  FHD(1).Pad = String$(108, 32)
  FileHdrLine$ = FHD(1).RecordID + FHD(1).TABLE + FHD(1).PROBE + FHD(1).VERS + FHD(1).Reserved
  FileHdrLine$ = FileHdrLine$ + FHD(1).CYCLES + FHD(1).RADIOREADYN + FHD(1).WANDYN + FHD(1).EXTFORMATYN + FHD(1).Pad

  'Cycle Header
  CHD(1).RecordID = "CHD"
  CHD(1).CycleNumber = CYCLES$
  CHD(1).NumberCycles = "0001"
  CHD(1).CycleDate = Left$(Date$, 2) + Mid$(Date$, 4, 2) + Mid$(Date$, 7, 4)
  CHD(1).Pad = String$(109, 32)
  CycleLine$ = CHD(1).RecordID + CHD(1).CycleNumber + CHD(1).NumberCycles + CHD(1).CycleDate + CHD(1).Pad

  'Route Header
  RHD(1).RecordID = "RHD"
  RHD(1).RouteNumber = CYCLES$ + Route$
  RHD(1).SurveyYN = "N"
  RHD(1).RouteMessYN = "N"
  RHD(1).Keys = "0000"
  RHD(1).Readings = Reading$
  RHD(1).Demands = "0000"
  RHD(1).Keyed = KReading$
  RHD(1).Probed = "0000"
  RHD(1).Radio = RReading$
  RHD(1).Customers = TReading$
  RHD(1).Meters = TReading$
  RHD(1).TimeAllowed = "000000"
  RHD(1).Gas = "0000"
  RHD(1).Water = TReading$
  RHD(1).Electric = "0000"
  RHD(1).Location = "0000"
  RHD(1).Extra = "0000"
  RHD(1).Region = "  "
  RHD(1).ZONE = "  "
  RHD(1).Office = "  "
   'was this
  ' RHD(1).BillCycle = Cycle$

  RHD(1).BillCycle = ""

  RHD(1).DropCycle = "  "
  RHD(1).WandReads = "0000"
  RHD(1).AMR = " "
  RHD(1).Pad = String$(32, " ")
  RouteHeader$ = RHD(1).RecordID + RHD(1).RouteNumber + RHD(1).SurveyYN
  RouteHeader$ = RouteHeader$ + RHD(1).RouteMessYN + RHD(1).Keys + RHD(1).Readings
  RouteHeader$ = RouteHeader$ + RHD(1).Demands + RHD(1).Keyed + RHD(1).Probed
  RouteHeader$ = RouteHeader$ + RHD(1).Radio + RHD(1).Customers + RHD(1).Meters
  RouteHeader$ = RouteHeader$ + RHD(1).TimeAllowed + RHD(1).Gas + RHD(1).Water
  RouteHeader$ = RouteHeader$ + RHD(1).Electric + RHD(1).Location
  RouteHeader$ = RouteHeader$ + RHD(1).Extra + RHD(1).Region + RHD(1).ZONE
  RouteHeader$ = RouteHeader$ + RHD(1).Office + RHD(1).BillCycle
  RouteHeader$ = RouteHeader$ + RHD(1).DropCycle + RHD(1).WandReads
  RouteHeader$ = RouteHeader$ + RHD(1).AMR + RHD(1).Pad

  'Start Build of Test File For Now
  outfile = FreeFile
  Open Path$ + "DOWNLOAD.DAT" For Output As #outfile
  Print #outfile, FileHdrLine$
  Print #outfile, CycleLine$
  Print #outfile, RouteHeader$

  'Now Do Individual Customers, Readings , and Radio Read if necessary

  CUSFile = FreeFile
  Open "UBHHCUS.DAT" For Random Shared As CUSFile Len = Len(CUS(1))
  MtrFile = FreeFile
  Open "UBHHMTR.DAT" For Random Shared As MtrFile Len = Len(Mtr(1))
  RdgFile = FreeFile
  Open "UBHHRDG.DAT" For Random Shared As RdgFile Len = Len(Rdg(1))
  RFFFile = FreeFile
  Open "UBHHRFF.DAT" For Random Shared As RFFFile Len = Len(RFF(1))
  WRRFile = FreeFile
  Open "UBHHWRR.DAT" For Random Shared As WRRFile Len = Len(WRR(1))

  For CusRecord = 1 To TCnt
    FrmShowPctComp.ShowPctComp CusRecord, TCnt&

    Get CUSFile, CusRecord, CUS(1)
    CusLine$ = CUS(1).RecordID + CUS(1).RouteNumber + CUS(1).Meters
    CusLine$ = CusLine$ + CUS(1).AccountNumber + CUS(1).Name
    CusLine$ = CusLine$ + CUS(1).Address1 + CUS(1).Address2 + CUS(1).Reserved
    CusLine$ = CusLine$ + CUS(1).Group + CUS(1).CustInfo + CUS(1).ExCustRec
    CusLine$ = CusLine$ + CUS(1).Segment + CUS(1).UtilityID + CUS(1).PassThur + CUS(1).Pad
    Print #outfile, CusLine$
    Get MtrFile, CusRecord, Mtr(1)
    MtrLine$ = Mtr(1).RecordID + Mtr(1).RouteNumber + Mtr(1).Reads + Mtr(1).Reserved
    MtrLine$ = MtrLine$ + Mtr(1).Group + Mtr(1).Reserved1 + Mtr(1).Reserved2
    MtrLine$ = MtrLine$ + Mtr(1).Survey + Mtr(1).Reserved3 + Mtr(1).Survey2
    MtrLine$ = MtrLine$ + Mtr(1).BillCode + Mtr(1).MtrStatus + Mtr(1).OpticalProbe
    MtrLine$ = MtrLine$ + Mtr(1).MtrNumber + Mtr(1).Reserved4 + Mtr(1).MeterType
    MtrLine$ = MtrLine$ + Mtr(1).MeterSeq + Mtr(1).MeterInfo + Mtr(1).Reserved5
    MtrLine$ = MtrLine$ + Mtr(1).Location + Mtr(1).Reserved6 + Mtr(1).ReadInst1
    MtrLine$ = MtrLine$ + Mtr(1).Reserved7 + Mtr(1).ReadInst2 + Mtr(1).SpecMessage + Mtr(1).Reserved8
    MtrLine$ = MtrLine$ + Mtr(1).SpecMessageYN + Mtr(1).MtrCat + Mtr(1).LocExtraMtr + Mtr(1).TimeCode
    MtrLine$ = MtrLine$ + Mtr(1).MtrAudit1 + Mtr(1).MtrAudit2 + Mtr(1).MtrAudit3
    MtrLine$ = MtrLine$ + Mtr(1).MtrAudit4 + Mtr(1).Pad
    Print #outfile, MtrLine$
    Get RdgFile, CusRecord, Rdg(1)
    RdgLine$ = Rdg(1).RecordID + Rdg(1).RouteNumber + Rdg(1).Text + Rdg(1).Prompt
    RdgLine$ = RdgLine$ + Rdg(1).ReadDir + Rdg(1).Compare + Rdg(1).Validation + Rdg(1).Reserved
    RdgLine$ = RdgLine$ + Rdg(1).Channel + Rdg(1).Dials + Rdg(1).Decimals
    RdgLine$ = RdgLine$ + Rdg(1).ReadMethod + Rdg(1).PrevRead + Rdg(1).High1
    RdgLine$ = RdgLine$ + Rdg(1).Low1 + Rdg(1).MtrConstant + Rdg(1).ConstantFlag
    RdgLine$ = RdgLine$ + Rdg(1).HHFFlag + Rdg(1).PosCreep + Rdg(1).Estimates
    RdgLine$ = RdgLine$ + Rdg(1).Reserved2 + Rdg(1).ReadType + Rdg(1).MaxPercent
    RdgLine$ = RdgLine$ + Rdg(1).MinPercent + Rdg(1).NegCreep + Rdg(1).Pad
    Print #outfile, RdgLine$
    If Rdg(1).ReadMethod = "R" Then
      Get RFFFile, CusRecord, RFF(1)
      RffLine$ = RFF(1).RecordID + RFF(1).RouteNumber + RFF(1).RFERT
      RffLine$ = RffLine$ + RFF(1).Reserved + RFF(1).RFProgram + RFF(1).Reserved2
      RffLine$ = RffLine$ + RFF(1).GeoArea + RFF(1).RFFreq + RFF(1).RFTone
      RffLine$ = RffLine$ + RFF(1).Reserved3 + RFF(1).Tamper + RFF(1).ConcIndicator
      RffLine$ = RffLine$ + RFF(1).Pad
      Print #outfile, RffLine$
    End If
    If Rdg(1).ReadMethod = "W" Then
      Get WRRFile, CusRecord, WRR(1)
      WRRLine$ = WRR(1).RecordID + WRR(1).RouteNumber + WRR(1).DeviceID
      WRRLine$ = WRRLine$ + WRR(1).WandProg + WRR(1).Resv1 + WRR(1).Resv2
      WRRLine$ = WRRLine$ + WRR(1).NodeNumb + WRR(1).NoTamper + WRR(1).Fill1
      Print #outfile, WRRLine$
    End If
   Next CusRecord
  If Not TCnt > 0 Then
    Unload FrmShowPctComp
    MsgBox "No Readings to Export", vbOKOnly, "No Reads Found"
  Else
    MsgBox "Readings Exported - " & Str(TCnt), vbOKOnly, "Procedure Complete"
  End If

  'File Trailor
  FHD(1).RecordID = "FTR"
  FHD(1).TABLE = "N"
  FHD(1).PROBE = "N"
  FHD(1).VERS = "01.00"
  FHD(1).Reserved = "   "
  FHD(1).CYCLES = "01"
  FHD(1).RADIOREADYN = "Y"
  FHD(1).WANDYN = "Y"
  FHD(1).EXTFORMATYN = "N"
  FHD(1).Pad = String$(108, 32)
  FileHdrLine$ = FHD(1).RecordID + FHD(1).TABLE + FHD(1).PROBE + FHD(1).VERS + FHD(1).Reserved
  FileHdrLine$ = FileHdrLine$ + FHD(1).CYCLES + FHD(1).RADIOREADYN + FHD(1).WANDYN + FHD(1).EXTFORMATYN + FHD(1).Pad

  'Cycle Trailor
  CHD(1).RecordID = "CTR"
  CHD(1).CycleNumber = CYCLES$
  CHD(1).NumberCycles = "0001"
  CHD(1).CycleDate = Left$(Date$, 2) + Mid$(Date$, 4, 2) + Mid$(Date$, 7, 4)
  CHD(1).Pad = String$(109, 32)
  CycleLine$ = CHD(1).RecordID + CHD(1).CycleNumber + CHD(1).NumberCycles + CHD(1).CycleDate + CHD(1).Pad

  'Route Trailor
  RHD(1).RecordID = "RTR"
  RHD(1).RouteNumber = CYCLES$ + Route$
  RHD(1).SurveyYN = "N"
  RHD(1).RouteMessYN = "N"
  RHD(1).Keys = "0000"
  RHD(1).Readings = Reading$
  RHD(1).Demands = "0000"
  RHD(1).Keyed = KReading$
  RHD(1).Probed = "0000"
  RHD(1).Radio = RReading$
  RHD(1).Customers = TReading$
  RHD(1).Meters = TReading$
  RHD(1).TimeAllowed = "000000"
  RHD(1).Gas = "0000"
  RHD(1).Water = TReading$
  RHD(1).Electric = "0000"
  RHD(1).Location = "0000"
  RHD(1).Extra = "0000"
   RHD(1).Region = "  "
  RHD(1).ZONE = "  "
  RHD(1).Office = "  "
  RHD(1).BillCycle = CYCLES$
  RHD(1).DropCycle = "  "
  RHD(1).WandReads = "0000"
  RHD(1).AMR = " "
  RHD(1).Pad = String$(32, " ")
  RouteHeader$ = RHD(1).RecordID + RHD(1).RouteNumber + RHD(1).SurveyYN
  RouteHeader$ = RouteHeader$ + RHD(1).RouteMessYN + RHD(1).Keys + RHD(1).Readings
  RouteHeader$ = RouteHeader$ + RHD(1).Demands + RHD(1).Keyed + RHD(1).Probed
  RouteHeader$ = RouteHeader$ + RHD(1).Radio + RHD(1).Customers + RHD(1).Meters
  RouteHeader$ = RouteHeader$ + RHD(1).TimeAllowed + RHD(1).Gas + RHD(1).Water
  RouteHeader$ = RouteHeader$ + RHD(1).Electric + RHD(1).Location
  RouteHeader$ = RouteHeader$ + RHD(1).Extra + RHD(1).Region + RHD(1).ZONE
  RouteHeader$ = RouteHeader$ + RHD(1).Office + RHD(1).BillCycle
  RouteHeader$ = RouteHeader$ + RHD(1).DropCycle + RHD(1).WandReads
  RouteHeader$ = RouteHeader$ + RHD(1).AMR + RHD(1).Pad
  Print #outfile, RouteHeader$       'Actually Trailors
  Print #outfile, CycleLine$
  Print #outfile, FileHdrLine$
  Close
Return
CheckInfo:

  InfoOK = True

  NewPath$ = QPTrim$(fpPathOut.Text)
  Path$ = GetBadgerPath$

  If Right$(NewPath$, 1) <> "\" Then
    NewPath$ = NewPath$ + "\"
  End If

'if this is a chage to the path to badger then update ubdata
  If NewPath$ <> Path$ Then
'    PathFile = FreeFile
'    Open PathData$ For Random Shared As PathFile
'    PathFile.PathWay = NewPath$
'    Put #PathFile, 1, PathFile
'    Close PathFile
    Path$ = NewPath$
  End If
 ' Path$ = GetBadgerPath$
  If InStr(Path$, "BAD") > 0 Then
    'DisplaySUScrn "Invalid"
    InfoOK = False
    GoTo NoGoodRet
  End If

'  If Not Val(fpMtrCycle) > 0 And Val(fpMtrRoute) > 0 Then
''    Color 12
''    LOCATE 11, 40
''    Print "Invalid Operation"
''    Beep
''    WaitForAction
''    Color 15
''    LOCATE 11, 40
''    Print String$(25, 32)
''    frm(1).FldNo = 1
''    Action = 1
'    Done = False
'    InfoOK = False
'  End If

NoGoodRet:
Return
'**************************************************
GetBadgerReadings:
  ReDim Rdgg(1) As BADReadRecType
  RDGLen = Len(Rdgg(1))
  Path$ = GetBadgerPath$

  If Len(Path$) > 0 And Right$(Path$, 1) <> "\" Then
    Path$ = Path$ + "\"
  End If

  FileNam$ = Path$ + "UPLOAD.DAT"

      GoSub CheckBadgerPath
      If InfoOK = False Then
        Exit Sub
      End If
    FrmShowPctComp.Label1 = "Importing Reading Information."
    FrmShowPctComp.cmdCancel.Enabled = False
    FrmShowPctComp.Show '1, Parent

  UBCustRecLen = Len(UBCustRec(1))
  UBFile = FreeFile
  Open "UBCUST.DAT" For Random Shared As UBFile Len = UBCustRecLen

  ReadFile = FreeFile
  Open FileNam$ For Random Shared As ReadFile Len = RDGLen
  NumReadRecs = LOF(ReadFile) / RDGLen

  If NumReadRecs = 0 Then
    Close
'    For zz = 8 To 16
'      QPrintRC Space$(49), zz, 16, 15
'    Next
'    DisplaySUScrn "BADMTRUP"
'    QPrintRC " ERROR: NO READINGS FOUND", 9, 27, -1
'    QPrintRC " Check the Badger Path and ", 10, 24, -1
'    QPrintRC " try the transfer again!!!", 11, 24, -1
'    QPrintRC " Press any key to continue. ", 13, 24, 14
'    Beep
    Unload FrmShowPctComp
    MsgBox "ERROR: NO READINGS FOUND", vbOKOnly, "Error"
  End If
'    BlockClear
'  DisplaySUScrn "BADMTRUP"
'
'  QPrintRC "  Processing Meter Readings.  ", 9, 25, 15
'  QPrintRC "    Updated Count:            ", 11, 25, 11

  CustRec& = 0
  GotCust = False
  For cnt& = 1 To NumReadRecs
    FrmShowPctComp.ShowPctComp cnt&, NumReadRecs
    
    Get ReadFile, cnt&, Rdgg(1)
    RecordID$ = Left$(Rdgg(1).Fill, 3)
    
    Select Case RecordID$
    Case "CUS"
      CustRec& = GetAccount&(Mid$(Rdgg(1).Fill, 15, 20))
      CustName$ = QPTrim$(Mid$(Rdgg(1).Fill, 35, 20))
    Case "RDG"
      GoSub ProcessReading
      CustRec& = 0
    Case "MTR"
      MtrNumber$ = QPTrim$(Mid$(Rdgg(1).Fill, 46, 11))
    Case Else
      'ignore all others
    End Select
  Next
  Close
  MsgBox "Number of Records Retrieved - " & Str(UpdCnt), vbOKOnly, "Completed"
Return


ProcessReading:
  If CustRec& <= 0 Then
    GoTo BadCustRet
  End If
  Get UBFile, CustRec&, UBCustRec(1)
    'If CustRec& = 2673 Then Stop
  RDate$ = Mid$(Rdgg(1).Fill, 48, 8)
  ReadDate = Date2Num(Left$(RDate$, 2) + "/" + Mid$(RDate$, 3, 2) + "/" + Right$(RDate$, 4))
  CurReading# = GetReading&(Mid$(Rdgg(1).Fill, 34, 10))
  For MeterRecord = 1 To 7
    MeterNumber$ = QPTrim$(UBCustRec(1).LocMeters(MeterRecord).MtrNum)
    If BroadFlag = True Then
      If Left$(MeterNumber$, 2) = "RR" Then
          MeterNumber$ = Mid$(MeterNumber$, 3, 8)
      ElseIf Left$(MeterNumber$, 2) = "WR" Then
          MeterNumber$ = Mid$(MeterNumber$, 3, 8)
      End If
    End If
     If InStr(MeterNumber$, QPTrim$(MtrNumber$)) > 0 Then
     'Changed 02072020
    'If MeterNumber$ = QPTrim$(MtrNumber$) Then
      If UBCustRec(1).LocMeters(MeterRecord).ReadFlag = "Y" Then
        UBCustRec(1).LocMeters(MeterRecord).CurRead = CurReading#
        UBCustRec(1).LocMeters(MeterRecord).CurDate = ReadDate
      Else
        UBCustRec(1).LocMeters(MeterRecord).PrevRead = UBCustRec(1).LocMeters(MeterRecord).CurRead
        UBCustRec(1).LocMeters(MeterRecord).PastDate = UBCustRec(1).LocMeters(MeterRecord).CurDate
        UBCustRec(1).LocMeters(MeterRecord).ReadFlag = "Y"
        UBCustRec(1).LocMeters(MeterRecord).CurDate = ReadDate
        UBCustRec(1).LocMeters(MeterRecord).CurRead = CurReading#
      End If
      UpdCnt = UpdCnt + 1
      Put UBFile, CustRec&, UBCustRec(1)
    End If
  Next
'look here
BadCustRet:

Return

'**************************************************
CheckBadgerPath:
  InfoOK = True
  NewPath$ = QPTrim$(fpPathOut.Text)
  Path$ = GetBadgerPath$

  If Right$(NewPath$, 1) <> "\" Then
    NewPath$ = NewPath$ + "\"
  End If

'if this is a change to the path to badger then update ubdata
  If NewPath$ <> Path$ Then
'    PathFile = FreeFile
'    Open PathData$ For Random Shared As PathFile
'    Print #PathFile, NewPath$
'    Close PathFile
    Path$ = NewPath$
  End If
'  Path$ = GetBadgerPath$

  If Right$(Path$, 1) <> "\" Then
    Path$ = Path$ + "\"
  End If
  FileNam$ = Path$ + "UPLOAD.DAT"
  If InStr(Path$, "BAD") > 0 Then
'    BlockClear
'    DisplaySUScrn "Invalid"
'    WaitForAction
    InfoOK = False
  End If

Return
GiveItUp:
  Close
  Unload FrmShowPctComp
  MsgBox "Errors were encountered, Please contact software support for assistance.", vbOKOnly, "ERRORS"
End Sub

Private Function GetBadgerPath$()
  Dim PathFile As Integer, Path As String, PathTest As Integer
  On Local Error GoTo BadFile

  If Not Exist(PathData$) Then
    PathFile = FreeFile
    Open PathData$ For Random Shared As PathFile
    Print #PathFile, "C:\mvrs50\xfer"
    Close PathFile
  End If
  PathFile = FreeFile
  Open PathData$ For Random Shared As PathFile
  Line Input #PathFile, Path$
  Close PathFile

  Path$ = QPTrim$(Path$)
  If Len(Path$) > 0 Then
    If Right$(Path$, 1) <> "\" Then
      Path$ = Path$ + "\"
    End If
  End If
'ok lets see if we can create a file in this path...
  PathTest = FreeFile
  Open Path$ + "UBTMPMES.MES" For Output As PathTest
  Close PathTest
'if we get to here the path is ok...
  KillFile Path$ + "UBTMPMES.MES"

  GetBadgerPath$ = Path$

BadExit:
  Exit Function

BadFile:
  GetBadgerPath$ = "BAD"
  Resume BadExit

End Function
Private Function GetAccount&(Account$)
  Account$ = QPTrim$(Account$)
  If Len(Account$) = 0 Then
    GetAccount& = -1
    GoTo AccountExit
  End If
  Do
    If Left$(Account$, 1) = "0" Then
      If Len(Account$) > 0 Then
        Account$ = Mid$(Account$, 2)
      Else
        GetAccount& = -1
        GoTo AccountExit
      End If
    Else
      GetAccount& = Val(Account$)
      Exit Do
    End If
  Loop
AccountExit:
End Function
Private Function GetReading&(Reading$)
  Reading$ = QPTrim$(Reading$)
  If Len(Reading$) = 0 Then
    GetReading& = -1
    GoTo ReadingExit
  End If
  Do
    If Left$(Reading$, 1) = "0" Then
      If Len(Reading$) > 0 Then
        Reading$ = Mid$(Reading$, 2)
      Else
        GetReading& = -1
        GoTo ReadingExit
      End If
    Else
      GetReading& = Val(Reading$)
      Exit Do
    End If
  Loop
ReadingExit:
End Function

'Type UBJettHHRecType         '
'' First two fields are required by CE File I/O dll.
'  CEVariant       As String * 2  'This can be blank or anything your heart desires
'  CEStrLen        As String * 2  'same as above
''**************************************************************
'  CustName      As String * 20
'  ServAddress   As String * 20
'  MeterID       As String * 8
'  LowRead       As String * 9
'  HighRead      As String * 9
'  Account       As String * 8
'  ReadFlag      As String * 1         'Y/N
'  MeterType     As String * 1
'  Book          As String * 2
'  CurRead       As String * 9
'  PastRead      As String * 9
'  ReadDate      As String * 8   ' fmt mmddyyyy
'  ReadTime      As String * 6   ' fmt hhmmss
'  Note1         As String * 20
'  Note2         As String * 20
'  Note3         As String * 20
'  NoteStatus    As String * 1  'T=Temp  P=Perm
'End Type
Private Sub ImpExpJettHHInfo(ByVal ImpExpFlag As Boolean)
  Dim CEReadFile As String, ReportFile As String
  Dim HighVar As Integer, LowVar As Integer
  Dim UBFile As Integer, HHFile As Integer
  Dim UBJettRecLen As Integer, NumJettRecs As Long
  Dim RMCnt As Long, WhatRMRec As Long
  Dim Ok2DoIt As Boolean
  Dim IdxFileName As String
  Dim BookCnt As Integer
  Dim MtrCnt As Integer
  Dim MeterOK As Boolean
  Dim Account As String, MTRType As String
  Dim WhatTypes As String, MeterID As String
  Dim Average As Double, LowRead As Double, HighRead As Double
  Dim bln4NewJett As Boolean
  ReDim UBJettRec(1) As UBJettHHRecType
  UBJettRecLen = Len(UBJettRec(1))
  
'  ReDim MsgText(0 To 5) As String
  Dim C As String, ThisDate As String, RptHandle As Integer
  Dim ReadingDate As Integer, ReadTime As String, Page As Integer
  Dim MMsg1 As String, MMsg2 As String, MMsg3 As String
  Dim CMsg1 As String, CMsg2 As String, CMsg3 As String
  Dim NumIntermecRdRecs As Long, sentcnt As Long
  Dim MeterRecord As Long
  Dim bln4OakFlag As Boolean
  Dim CurReading As Double
  ReportFile$ = "UBJettTR.RPT"
  MaxLines = 55
  
  If NewExist(UBPath$ + "UBNUJETT.DAT") Then
    bln4NewJett = True
  Else
    bln4NewJett = False
  End If
     
  CEReadFile$ = UBPath$ + "CEMTREAD.DAT"
  
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen
  If InStr(UCase(UBSetUpRec(1).DEFCITY), "FOUR OAK") > 0 Then
    bln4OakFlag = True
  End If

  HighVar = UBSetUpRec(1).HighRead
  LowVar = UBSetUpRec(1).LowRead
  
  If ImpExpFlag Then     'EXPORTING METER READINGS
    WhatTypes$ = Left$(Me.fpMtrType.Text, 1)
    GoSub ExportJett
      MsgText(0) = "Jett Export Operation."
      MsgText(1) = ""
      MsgText(2) = ""
      MsgText(3) = "Process Complete."
      MsgText(4) = ""
      MsgText(5) = ""
    Select Case GetOKorNot%(MsgText(), True, True)
    Case True
      Call cmdExit_Click
    End Select
  Else
    GoSub ImportJett
      MsgText(0) = "Jett Import Operation."
      MsgText(1) = ""
      MsgText(2) = ""
      MsgText(3) = "Process Complete."
      MsgText(4) = ""
      MsgText(5) = ""
    Select Case GetOKorNot%(MsgText(), True, True)
    Case True
      Call cmdExit_Click
    End Select

  End If

Exit Sub

ExportJett:
  sentcnt = 0
  
  If CycleCnt > 0 Then
    
    Call KillFile(CEReadFile$)
    
    UBJettRec(1).CEVariant = ""
    UBJettRec(1).CEStrLen = ""
    UBJettRecLen = Len(UBJettRec(1))
    
    UBCustRecLen = Len(UBCustRec(1))
    UBFile = FreeFile               'Open Customer Data File
    Open UBPath$ + "UBCUST.DAT" For Random Shared As UBFile Len = UBCustRecLen

    HHFile = FreeFile
    Open CEReadFile$ For Random Shared As HHFile Len = UBJettRecLen
        
    If UBSetUpRec(1).UseSeq = "Y" Then
      MakeSequenceIndex "Sequence Number", Me
      IdxFileName = UBPath$ + "UBTEMP.IDX"
    Else
      IdxFileName = UBPath$ + "UBCUSTBK.IDX"
    End If
    
    IdxRecLen = 4         'we are using a integer
    IdxFileSize& = FileSize&(IdxFileName)
    IdxNumOfRecs = IdxFileSize& \ IdxRecLen
    
    ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
    FGetAH IdxFileName, IdxBuff(), IdxRecLen, IdxNumOfRecs            'load it
    'Open the Correct Order for Reading
  End If
  
  FrmShowPctComp.Label1 = "Exporting Reading Information."
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show '1, Parent
  
  RMCnt = 1

  Do
    WhatRMRec& = IdxBuff(RMCnt).RecNum
    If Not (WhatRMRec&) = 0 Then
      Get UBFile, WhatRMRec&, UBCustRec(1)
      If InStr(UBCustRec(1).HHMSG1, "NOREAD") > 0 Then
        GoTo JWriteSkip
      End If
      If bln4OakFlag Then
        If UBCustRec(1).BILLCYCL = 99 Then
          GoTo JWriteSkip
        End If
      End If
      For BookCnt = 1 To CycleCnt
        If Val(UBCustRec(1).Book) = Cycle(BookCnt) And (UBCustRec(1).Status <> "F") Then
          GoSub JettWriteRec
          Exit For
        End If
      Next
    End If
JWriteSkip:
    FrmShowPctComp.ShowPctComp RMCnt&, IdxNumOfRecs
    RMCnt = RMCnt + 1
  Loop Until RMCnt > IdxNumOfRecs
  
  Close  'Close files

  MsgText(0) = "Jett Export Operation."
  MsgText(1) = "Created the 'CEMTREAD.DAT'"
  MsgText(2) = "with " + Str$(sentcnt&) + " records."
  MsgText(3) = "Make sure the Jett HH is connected."
  MsgText(4) = "On the Citipak File Transfer Utility"
  MsgText(5) = " Select Copy TO HH button."
  
  Select Case GetOKorNot%(MsgText(), False, True)
  Case True
    Shell "JettCopy.exe"
  End Select
  Return
  
ImportJett:
  sentcnt = 0
    MsgText(0) = "Jett Import Operation"
    MsgText(1) = "Make sure the JETT is connected"
    MsgText(2) = "then on the Citipak File Transfer Utility"
    MsgText(3) = "Select Copy FROM HH button."
    MsgText(4) = ""
    MsgText(5) = "IS THE JETT HH READY TO PROCEED?"
  
    DoEvents
 
    Select Case GetOKorNot%(MsgText(), False, True)
    Case Not True
      GoTo JettGetExit
    Case True
      Shell "JettCopy.exe"
    End Select
    MsgText(0) = "Jett Import Readings."
    MsgText(1) = "Continue with reading update?"
    MsgText(2) = ""
    MsgText(3) = ""
    MsgText(4) = ""
    MsgText(5) = ""
  Select Case GetOKorNot%(MsgText(), False, True)
  Case Not True
    GoTo JettGetExit
  End Select
  
  UBCustRecLen = Len(UBCustRec(1))
  UBFile = FreeFile
  Open UBPath + "UBCUST.DAT" For Random Shared As UBFile Len = UBCustRecLen

  UBJettRecLen = Len(UBJettRec(1))

  HHFile = FreeFile
  Open CEReadFile$ For Random Shared As HHFile Len = UBJettRecLen

  NumJettRecs = LOF(HHFile) / UBJettRecLen

  If NumJettRecs = 0 Then
    Close
    MsgText(0) = "ERROR:"
    MsgText(1) = ""
    MsgText(2) = "The file 'CEMTREAD.DAT' must be"
    MsgText(3) = "in the Citipak folder."
    MsgText(4) = "Please call Southern Software for"
    MsgText(5) = "additional Information."
    GetOKorNot% MsgText(), True, False
    GoTo JettGetExit:
  End If
  RptHandle = FreeFile
  Open ReportFile$ For Output As RptHandle
  GoSub ImpRptHeader

  FrmShowPctComp.Label1 = "Importing Meter Reading Information."
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show '1, Parent
  
  For RMCnt& = 1 To NumJettRecs
    Get HHFile, RMCnt, UBJettRec(1)
    WhatRMRec& = Val(QPTrim$(UBJettRec(1).Account))
    If Not (WhatRMRec&) = 0 Then
      Get UBFile, WhatRMRec&, UBCustRec(1)
      GoSub JettExtractRec
      
      If LineCnt > MaxLines Then
        Print #RptHandle, FF$
        GoSub ImpRptHeader
      End If

    End If
    FrmShowPctComp.ShowPctComp RMCnt&, NumJettRecs
  Next
  GoSub ImpRptTotal
  

  Close
  
  MsgText(0) = "Jett Import Readings."
  MsgText(1) = "Readings Updated Successfully."
  MsgText(2) = ""
  MsgText(3) = " UPDATED: " + Str$(sentcnt&) + " Readings"
  MsgText(4) = " of " + Str$(NumJettRecs) + " Accounts."
  MsgText(5) = ""
  GetOKorNot% MsgText(), True, True
  
  frmReportOpt.Show 1
  If rptopt = 1 Then
    'do the graphics
    Load frmLoadingRpt
    frmLoadingRpt.setwherefrom frmUBImpExpHHRead
    ARptLineRpt.GetName ReportFile$
    ARptLineRpt.startrpt
  ElseIf rptopt = 2 Then
    'do the text
    ViewPrint ReportFile$, "Jett Meter Readings Report"
  End If
  
JettGetExit:
  Return
ImpRptHeader:
  Page = Page + 1
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle, "Date: "; Date$; Tab(27); "Jett Imported Meter Reading Report"; Tab(70); "Page:"; Page
  Print #RptHandle, "Account   RecNum  Customer Name               Reading     ReadDate   ReadTime"
  Print #RptHandle, String$(80, "-")
  LineCnt = 6
Return

ImpRptTotal:
  Print #RptHandle, String$(80, "-")
  Print #RptHandle, "Read Count: " + Using("######", sentcnt&, False)
Return

JettExtractRec:

  If UBJettRec(1).ReadFlag = "Y" Then
    MeterRecord = Val(Right$((QPTrim$(UBJettRec(1).Account)), 1))
    sentcnt = sentcnt + 1
' Check Meter Updated Flag
' Update Meter W/Reading
''NOTE: New current reading stored in the
' pastreading field from CEMTREAD.DAT
    
    CurReading# = Val(UBJettRec(1).PastRead)
    ThisDate$ = Left$(UBJettRec(1).ReadDate, 2) + "/" + Mid$(UBJettRec(1).ReadDate, 3, 2) + "/" + Right$(UBJettRec(1).ReadDate, 4)
    ReadingDate = Date2Num(ThisDate$)

    If UBCustRec(1).LocMeters(MeterRecord).ReadFlag = "Y" Then
      UBCustRec(1).LocMeters(MeterRecord).CurRead = CurReading#
      UBCustRec(1).LocMeters(MeterRecord).CurDate = ReadingDate
    Else
      UBCustRec(1).LocMeters(MeterRecord).PrevRead = UBCustRec(1).LocMeters(MeterRecord).CurRead
      UBCustRec(1).LocMeters(MeterRecord).PastDate = UBCustRec(1).LocMeters(MeterRecord).CurDate
      UBCustRec(1).LocMeters(MeterRecord).ReadFlag = "Y"
      UBCustRec(1).LocMeters(MeterRecord).CurDate = ReadingDate
      UBCustRec(1).LocMeters(MeterRecord).CurRead = CurReading#
    End If

    MMsg1$ = QPTrim$(UBJettRec(1).Note1)
    CMsg1$ = QPTrim$(UBCustRec(1).HHMSG1)
    MMsg2$ = QPTrim$(UBJettRec(1).Note2)
    CMsg2$ = QPTrim$(UBCustRec(1).HHMSG2)
    MMsg3$ = QPTrim$(UBJettRec(1).Note3)
    CMsg3$ = QPTrim$(UBCustRec(1).HHMSG3)
    'ReadTime$ = QPTrim$(UBJettRec(1).ReadTime)
    ReadTime$ = Format$(UBJettRec(1).ReadTime, "##:##:##")
    Print #RptHandle, Str(WhatRMRec&); Tab(10); Using("######", RMCnt, False); Tab(19);
    Print #RptHandle, QPTrim$(UBJettRec(1).CustName); Tab(45); Using("#########", CurReading#, False);
    Print #RptHandle, Tab(58); ThisDate$; Tab(70); ReadTime$
    LineCnt = LineCnt + 1

    If MMsg1$ <> CMsg1$ Then
      GoSub UpDateNoteInfo
      GoTo DoneINNotes
    End If
    If MMsg2$ <> CMsg2$ Then
      GoSub UpDateNoteInfo
      GoTo DoneINNotes
    End If
    If MMsg3$ <> CMsg3$ Then
      GoSub UpDateNoteInfo
    End If

DoneINNotes:
    Put UBFile, WhatRMRec&, UBCustRec(1)
  End If
Return

UpDateNoteInfo:
  UBCustRec(1).NewNotes = True
  UBCustRec(1).HHMSG1 = MMsg1$
  UBCustRec(1).HHMSG2 = MMsg2$
  UBCustRec(1).HHMSG3 = MMsg3$
Return

JettWriteRec:  'May Have Up to 7 Meters to Read
  For MtrCnt = 1 To 7
    MeterOK = False
    Account$ = Str$(WhatRMRec&)
    Account$ = Left$(Account$, 6) + "-" + Right$(Str$(MtrCnt), 1)
      
    MTRType$ = UBCustRec(1).LocMeters(MtrCnt).MTRType
    If MTRType$ = "W" Or MTRType$ = "S" Or MTRType$ = "C" Or MTRType$ = "E" Or MTRType$ = "D" Or MTRType$ = "G" Then
      Select Case WhatTypes$
      Case "W"                'water/sewer
        If MTRType$ = "W" Or MTRType$ = "S" Or MTRType$ = "C" Then
          MeterOK = True
        End If
      Case "E"                'electric & demand elec.
        If MTRType$ = "E" Or MTRType$ = "D" Then
          MeterOK = True
        End If
      Case "G"                'gas
        If MTRType$ = "G" Then
          MeterOK = True
        End If
      Case "A", " "           'all meters
        MeterOK = True
      End Select

      If MeterOK = True Then
'          ' Determine High and Low Reading
        Average = UBCustRec(1).LocMeters(MtrCnt).AvgUse
        LowRead# = Fix(LowRead#)
        HighRead# = Fix(Average# * (HighVar / 100)) + UBCustRec(1).LocMeters(MtrCnt).CurRead

        MeterID$ = LTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrNum)
        MeterID$ = RTrim$(MeterID$)

        If Val(MeterID$) = 0 Then
          MeterID$ = UBCustRec(1).Book + UBCustRec(1).SEQNUMB
        End If
        If Len(MeterID$) < 8 Then
          MeterID$ = String$(8 - Len(MeterID$), "0") + MeterID$
        End If
        MeterID$ = Left$(MeterID$, 8)
'          'Set Record Fields and Put On Disk
        UBJettRec(1).CustName = UBCustRec(1).CustName
        UBJettRec(1).ServAddress = Left$(UBCustRec(1).ServAddr, 20)
        UBJettRec(1).ReadDate = ""
        UBJettRec(1).ReadTime = ""
        UBJettRec(1).MeterID = MeterID$
        UBJettRec(1).LowRead = QPTrim$(Str$(LowRead#))
        UBJettRec(1).HighRead = QPTrim$(Str$(HighRead#))
        UBJettRec(1).Account = Account$
        UBJettRec(1).MeterType$ = UBCustRec(1).LocMeters(MtrCnt).MTRType
        UBJettRec(1).Book = UBCustRec(1).Book
        UBJettRec(1).CurRead = QPTrim$(Str$(UBCustRec(1).LocMeters(MtrCnt).CurRead))
        UBJettRec(1).PastRead = "0" 'QPTrim$(Str$(UBCustRec(1).LocMeters(MtrCnt).PrevRead))
        'UBJettRec(1).PastRead = QPTrim$(Str$(UBCustRec(1).LocMeters(MtrCnt).PrevRead))
       
        UBJettRec(1).ReadFlag = "N"

        UBJettRec(1).Note1 = QPTrim$(UBCustRec(1).HHMSG1)
        UBJettRec(1).Note2 = QPTrim$(UBCustRec(1).HHMSG2)
        UBJettRec(1).Note3 = QPTrim$(UBCustRec(1).HHMSG3)
        '         ^^^
        UBJettRec(1).NoteStatus = ""
        sentcnt = sentcnt + 1
        Put HHFile, (LOF(HHFile) / UBJettRecLen) + 1, UBJettRec(1)
      End If
    End If
  Next
  
Return


End Sub
''Performance Meter HH
Private Sub ImpExpPMHHInfo(ByVal ImpExpFlag As Boolean)
  Dim Path As String, FileNam As String, Whattomult As Integer
  Dim UBFile As Integer, ReadFile As Integer, NumReadRecs As Long
  Dim zz As Integer, CustRec As Long, GotCust As Boolean
  Dim cnt As Long, RecordID As String, CustName As String
  Dim UpdCnt As Long, MeterRecord As Long, RDate As String
  Dim ReadDate As Integer, CurReading As Double, InfoOK As Boolean
  Dim NewPath As String, PathFile As Integer, q As String
  Dim cb As String, BroadFlag As Boolean, CYCLES As String
  Dim Route As String, TCnt As Long, MReading As String
  Dim TReading As String, KCnt As Long, KReading As String
  Dim RReading As String, Reading As String, Done As Boolean
  Dim HighReadPerc As Double, LowReadPerc As Double
  Dim Prec As Long, RecNumber As Long, Chkcycle As Boolean
  Dim DisplayAcct As String, MeterNumber As String, MNum2File As String
  Dim MeterSeq As String, Dials As Integer, DialsS As String
  Dim Decimals As String, ERTID As String, ReadMethod As String
  Dim PrevRead As String, HighRead As Double, HighReadS As String
  Dim LowRead As Double, LowReadS As String, FileHdrLine As String
  Dim CycleLine As String, RouteHeader As String, outfile As Integer
  Dim CusRecord As Long, MtrNumber$, PrevDate$, Numsent As Long
  Dim RtrFile As Integer, BookCnt As Integer, UBGetemLen As Integer
  Dim MtrFile As Integer, Mout As String, Hout As String
  Dim Chkbook As Boolean, RType As String, MPIDNO As String, metercnt As Integer
  Dim FirstRt As String, FirstCyc As String, UsesSeqN As Boolean
  Dim Prev#
  FirstRt = ""
  FirstCyc = ""
  Numsent = 0
  'On Local Error GoTo GiveItUp
  Chkbook = False
  Chkcycle = False
  PathData$ = UBHHPathWayFile
  
  If ImpExpFlag Then     'EXPORTING METER READINGS
    GoSub SendReads
  Else
    GoSub GetReads
  End If

   Close

AbortExit:

  Exit Sub

SendReads:
  q$ = Chr$(34)
  cb$ = Space$(45)

  
  
  ReDim Rtr(1) As UBPMHHRouteRecType
  ReDim Mtr(1) As UBPMHHSendRecType
  Path$ = HHPathInOut
  outfile = FreeFile
  Open Path$ + "import.txt" For Output As #outfile
  
BadDataRestart:

      GoSub CheckInfo
      If InfoOK Then
        FrmShowPctComp.Label1 = "Exporting Reading Information."
        FrmShowPctComp.cmdCancel.Enabled = False
        FrmShowPctComp.Show '1, Parent
        CYCLES$ = QPTrim$(fpMtrCycle)
        If Len(CYCLES$) = 1 And CYCLES$ <> "0" Then
          CYCLES$ = "0" + CYCLES$
          Chkcycle = True
        Else
          CYCLES$ = "  "
          Chkcycle = False
        End If
        If CycleCnt > 0 Then
          Chkbook = True
        Else
          Chkbook = False
        End If
        If Chkcycle = False And Chkbook = False Then
          Unload FrmShowPctComp
          MsgBox "You must select either a cycle or book."
          Exit Sub
        End If
        GoSub HostDownLoad

      'Open HandHeld Work Files
       RtrFile = FreeFile
       Open "UBHHRtr.DAT" For Random Shared As RtrFile Len = Len(Rtr(1))
       MtrFile = FreeFile
       Open "UBHHMtr.DAT" For Random Shared As MtrFile Len = Len(Mtr(1))
           If Chkbook = True Then
             For BookCnt = 1 To CycleCnt
                 Route$ = Cycle(BookCnt)
                  cnt = 1
                 GoSub Dothecust
             Next
             Print #outfile,
             Close
             MsgBox "Number of Records Sent - " & Str(Numsent), vbOKOnly, "Completed"
          Else
             Route$ = "      "
             GoSub Dothecust
             Print #outfile,
             Close
             MsgBox "Number of Records Sent - " & Str(Numsent), vbOKOnly, "Completed"
          End If
      Else
        'GoTo BadDataRestart
      End If
  Close
Return

HostDownLoad:

  KillFile "UBHHTHd.DAT"
  KillFile "UBHHTSr.DAT"
 
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen

  'Open Customer Data File
 ' ReDim UBCustRec(1) As NewUBCustRecType
   UBCustRecLen = Len(UBCustRec(1))
  UBFile = FreeFile
  Open "UBCUST.DAT" For Random Shared As UBFile Len = UBCustRecLen

  HighReadPerc = UBSetUpRec(1).HighRead
  LowReadPerc = UBSetUpRec(1).LowRead

  'Open the Correct Order for Reading
  If UBSetUpRec(1).UseSeq = "Y" Then
    UsesSeqN = True
    IdxRecLen = 4               'we are using an integer
    MakeSequenceIndex "Sequence Number", Me
    IdxNumOfRecs = FileSize&("UBTEMP.IDX") \ IdxRecLen
    ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
    FGetAH "UBTEMP.IDX", IdxBuff(), IdxRecLen, IdxNumOfRecs
  Else
    UsesSeqN = False
    IdxRecLen = 4               'we are using an integer
    IdxFileSize& = FileSize&("UBCUSTBK.IDX")
    IdxNumOfRecs = IdxFileSize& \ IdxRecLen
      ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
    FGetAH "UBCUSTBK.IDX", IdxBuff(), IdxRecLen, IdxNumOfRecs  'load it
  End If
  cnt = 1

 
Return
Dothecust:
Do

    Prec& = IdxBuff(cnt).RecNum
    RecNumber = Prec&
    
     If Not (Prec&) = 0 Then
      Get UBFile, Prec&, UBCustRec(1)
      FrmShowPctComp.ShowPctComp cnt, IdxNumOfRecs
  If (UBCustRec(1).Status <> "F") And Val(UBCustRec(1).Book) > 0 Then
  If UBCustRec(1).DelFlag = 0 Then
      If Chkbook = True Then
          If Not Val(UBCustRec(1).Book) = Cycle(BookCnt) Then
            GoTo BadOnes
          Else
            Route$ = Cycle(BookCnt)
            If Len(Route$) = 1 Then Route$ = "0" + Route$
            Route$ = Route$ + "    "
          End If
      End If
      If Chkcycle = True Then
        If Not UBCustRec(1).BILLCYCL = Val(CYCLES$) Then
        GoTo BadOnes
        End If
      
      End If
        
    
       'when to create ?????
       If FirstCyc <> CYCLES$ Or FirstRt <> Route$ Then
          FirstCyc = CYCLES$
          FirstRt = Route$
          Rtr(1).HeaderID = "R1"
          Rtr(1).CycleNum = CYCLES$
          Rtr(1).RouteNum = Route$
          Rtr(1).CrLf = Chr$(13) + Chr$(10)
          Hout$ = Rtr(1).HeaderID + Rtr(1).CycleNum + Rtr(1).RouteNum + Rtr(1).CrLf
          Print #outfile, Hout$;
        End If
       'Do SomeThing Here
        'Build Customer Meter Reading & Radio Read File
      Whattomult = Val(UBCustRec(1).USERCODE1)
      For metercnt = 1 To 7

        If Len(QPTrim$(UBCustRec(1).LocMeters(metercnt).MtrNum)) > 0 Then
          Mtr(1).RecordID = "M1"
          Mtr(1).AccountNum = LTrim$(Str$(RecNumber))
          Mtr(1).MeterNum = UBCustRec(1).LocMeters(metercnt).MtrNum
          If UsesSeqN = True Then
            If UBCustRec(1).Seq < 0 Then
              UBCustRec(1).Seq = 0
            End If
            Mtr(1).SeqNum = UBCustRec(1).Seq
          Else
            Mtr(1).SeqNum = UBCustRec(1).SEQNUMB
          End If
          Mtr(1).ReadTypeCode = UBCustRec(1).LocMeters(metercnt).MTRType
            If UBCustRec(1).LocMeters(metercnt).CurRead > 0 Then
             Prev# = UBCustRec(1).LocMeters(metercnt).CurRead
             'This was added for Ron Gallon so their software will report correct prev read
             'for Saluda City 8/26/2008
             If Whattomult > 0 Then
                Prev# = Prev# / Whattomult
              Else
                Prev# = Prev# / 10
              End If
              Prev# = Int(Prev#)
              PrevRead$ = LTrim$(Str$(Prev#))
            Else
              PrevRead$ = "0"
            End If
            
            PrevRead$ = PrevRead$ + String$(10 - Len(PrevRead$), " ")
            'need to fix these
            If UBCustRec(1).LocMeters(metercnt).AvgUse < 0 Then
              UBCustRec(1).LocMeters(metercnt).AvgUse = 0
            End If
            'do the first line instead of a consumption
           HighRead# = Fix(((0# + UBCustRec(1).LocMeters(1).AvgUse) * (HighReadPerc / 100) + UBCustRec(1).LocMeters(1).CurRead))
           ' HighRead# = Fix((0# + UBCustRec(1).LocMeters(metercnt).AvgUse) * (HighReadPerc / 100))
            If HighRead# <= 1 Then
              HighRead# = 1
            End If
            'added 5/10/2010 to fix hi read for Saluda
             If Whattomult > 0 Then
                HighRead# = HighRead# / Whattomult
              Else
                HighRead# = (HighRead# / 10)
              End If
              HighRead# = Int(HighRead#)
            HighReadS$ = LTrim$(Str$(HighRead#))
            If Len(HighReadS$) < 10 Then
              HighReadS$ = HighReadS$ + String$(10 - Len(HighReadS$), " ")
            End If
            LowRead# = PrevRead$ '1 'UBCustRec(1).LocMeters(1).CurRead
                LowReadS$ = LTrim$(Str$(LowRead#))
            If Len(LowReadS$) < 10 Then
              LowReadS$ = LowReadS$ + String$(10 - Len(LowReadS$), " ")
            End If
            PrevDate$ = Num2Date(UBCustRec(1).LocMeters(metercnt).CurDate)
            PrevDate$ = Left$(PrevDate$, 2) + Mid$(PrevDate$, 4, 2)
          Mtr(1).HighRead = HighReadS$
          Mtr(1).LowRead = LowReadS$
          Mtr(1).LocatnCode = String$(4, " ")
          Mtr(1).InstrCode = String$(4, " ")
          Mtr(1).AcctCode = UBCustRec(1).Status
          Mtr(1).Address = LTrim$(Left$(UBCustRec(1).ServAddr, 40))
          Mtr(1).Name = LTrim$(Left$(UBCustRec(1).CustName, 20))
          Mtr(1).PrevRead = PrevRead$
          Mtr(1).PReadDate = PrevDate$
            Dials = Val(UBCustRec(1).USERCODE2)
            If Dials = 0 Then Dials = 6
            DialsS$ = "0" + Right$(Str$(Dials), 1)
          Mtr(1).NumDials = DialsS$
          Mtr(1).DecimalPL = " "
          Mtr(1).Compound = String$(10, " ")
          Mtr(1).UID = QPTrim$(UBCustRec(1).LocMeters(metercnt).MtrIDNO)
          Mtr(1).Longitude = UBCustRec(1).LocMeters(metercnt).MtrLng
          Mtr(1).Latitude = UBCustRec(1).LocMeters(metercnt).MtrLat
          Mtr(1).Filler = String$(10, " ")
          Mtr(1).CrLf = Chr$(13) + Chr$(10)
          Mout$ = Mtr(1).RecordID + Mtr(1).AccountNum + Mtr(1).MeterNum + Mtr(1).SeqNum
          Mout$ = Mout$ + Mtr(1).ReadTypeCode + Mtr(1).HighRead + Mtr(1).LowRead
          Mout$ = Mout$ + Mtr(1).LocatnCode + Mtr(1).InstrCode + Mtr(1).AcctCode
          Mout$ = Mout$ + Mtr(1).Address + Mtr(1).Name + Mtr(1).PrevRead
          Mout$ = Mout$ + Mtr(1).PReadDate + Mtr(1).NumDials + Mtr(1).DecimalPL
          Mout$ = Mout$ + Mtr(1).Compound + Mtr(1).UID + Mtr(1).Longitude
          Mout$ = Mout$ + Mtr(1).Latitude + Mtr(1).Filler + Mtr(1).CrLf
          Numsent = Numsent + 1
          Print #outfile, Mout$;
          
        End If
        'End If
      Next metercnt

BadOnes:
     End If
     End If
    End If
    cnt = cnt + 1
  Loop Until cnt > IdxNumOfRecs
Return
CheckInfo:

  InfoOK = True

  NewPath$ = QPTrim$(fpPathOut.Text)
  

  
'if this is a chage to the path to hh then update ubdata
  If NewPath$ <> Path$ Then
    Path$ = NewPath$
    
  End If
  If InStr(Path$, "BAD") > 0 Then
    'DisplaySUScrn "Invalid"
    InfoOK = False
    GoTo NoGoodRet
  End If
NoGoodRet:
Return
'**************************************************
GetReads:
  Path$ = QPTrim$(fpPathOut.Text)
'Let user key path and file name in
  FileNam$ = Path$
      GoSub CheckPath
      If InfoOK = False Then
        Exit Sub
      End If
Dim TextLine As String
Dim xx As Integer, xx2 As Integer
xx = FreeFile
Open FileNam$ + "export.txt" For Input As #xx
xx2 = FreeFile
Open FileNam$ + "export1.txt" For Output As #xx2
Do While Not eof(xx)
   Line Input #xx, TextLine
   If Mid$(TextLine, 1, 2) = "M1" Then
      Print #xx2, TextLine
   End If
Loop
Close #xx
Close #xx2
  ReDim UBGetem(1) As UBPMHHGetRecType
  UBGetemLen = Len(UBGetem(1))
    FrmShowPctComp.Label1 = "Importing Reading Information."
    FrmShowPctComp.cmdCancel.Enabled = False
    FrmShowPctComp.Show '1, Parent

  UBCustRecLen = Len(UBCustRec(1))
  UBFile = FreeFile
  Open "UBCUST.DAT" For Random Shared As UBFile Len = UBCustRecLen

  ReadFile = FreeFile
  Open FileNam$ + "export1.txt" For Random Shared As ReadFile Len = UBGetemLen
  NumReadRecs = LOF(ReadFile) / UBGetemLen

  If NumReadRecs = 0 Then
    Close

    Unload FrmShowPctComp
    MsgBox "ERROR: NO READINGS FOUND", vbOKOnly, "Error"
  End If

  CustRec& = 0
  GotCust = False
  For cnt& = 1 To NumReadRecs
    FrmShowPctComp.ShowPctComp cnt&, NumReadRecs
    
    Get ReadFile, cnt&, UBGetem(1)
    RecordID$ = (UBGetem(1).RecordID)
    
    Select Case RecordID$
    Case "M1"
      CustRec& = GetAccount&(QPTrim$(UBGetem(1).AccountNum))
      CustName$ = QPTrim$(UBGetem(1).Name)
      MtrNumber$ = QPTrim$(UBGetem(1).MeterNum)
      GoSub ProcessReading
      
    Case Else
      'ignore all others
      CustRec& = 0
    End Select
  
  Next
  Close
  MsgBox "Number of Records Retrieved - " & Str(UpdCnt), vbOKOnly, "Completed"
Return


ProcessReading:
  If CustRec& <= 0 Then
    GoTo BadCustRet
  End If
  Get UBFile, CustRec&, UBCustRec(1)
  'from amr stored yyyymmdd
  RDate$ = Mid$(UBGetem(1).ReadDate, 5, 2) + "/" + Right$(UBGetem(1).ReadDate, 2) + "/" + Left$(UBGetem(1).ReadDate, 4)
  ReadDate = Date2Num(RDate$)
  CurReading# = GetReading&(UBGetem(1).Reading)
  'added 4/9/2008  so the cust would not have to add mult and mess the hist up
  Whattomult = Val(UBCustRec(1).USERCODE1)
  If Whattomult > 0 Then
    CurReading# = CurReading# * Whattomult
  Else
    CurReading# = CurReading# * 10
  End If
  For MeterRecord = 1 To 7
    If QPTrim$(UBCustRec(1).LocMeters(MeterRecord).MtrNum) = QPTrim$(MtrNumber$) Then
      If UBCustRec(1).LocMeters(MeterRecord).ReadFlag = "Y" Then
        UBCustRec(1).LocMeters(MeterRecord).CurRead = CurReading#
        UBCustRec(1).LocMeters(MeterRecord).CurDate = ReadDate
      Else
        UBCustRec(1).LocMeters(MeterRecord).PrevRead = UBCustRec(1).LocMeters(MeterRecord).CurRead
        UBCustRec(1).LocMeters(MeterRecord).PastDate = UBCustRec(1).LocMeters(MeterRecord).CurDate
        UBCustRec(1).LocMeters(MeterRecord).ReadFlag = "Y"
        UBCustRec(1).LocMeters(MeterRecord).CurDate = ReadDate
        UBCustRec(1).LocMeters(MeterRecord).CurRead = CurReading#
      End If
      UBCustRec(1).LocMeters(MeterRecord).MtrLat = Val(QPTrim$(UBGetem(1).Latitude))
      UBCustRec(1).LocMeters(MeterRecord).MtrLng = Val(QPTrim$(UBGetem(1).Longitude))

      UpdCnt = UpdCnt + 1
      Put UBFile, CustRec&, UBCustRec(1)
    End If
  Next

BadCustRet:

Return

'**************************************************
CheckPath:
  InfoOK = True
  NewPath$ = QPTrim$(fpPathOut.Text)

  If NewPath$ <> Path$ Then
    GetHandHeldPathWay%
    Path$ = NewPath$
  End If
'  Path$ = GetBadgerPath$

'  If Right$(Path$, 1) <> "\" Then
'    Path$ = Path$ + "\"
'  End If
  FileNam$ = Path$
  If InStr(Path$, "BAD") > 0 Then
'    BlockClear
'    DisplaySUScrn "Invalid"
'    WaitForAction
    InfoOK = False
  End If

Return
GiveItUp:
  Close
  Unload FrmShowPctComp
  MsgBox "Errors were encountered, Please contact software support for assistance.", vbOKOnly, "ERRORS"
End Sub

Private Sub ImpExpBadgerV2HHInfo(ByVal ImpExpFlag As Boolean)
  Dim HighVar As Integer, LowVar As Integer, IdxFile As Integer
  Dim WhatTypes As String, CustAcc As String, Prec As Long, Z6 As String
  Dim UBFile As Integer, BadgerCFGFile As Integer, Multi As Long
  Dim UBBadgerRecLen As Integer, NumBadgerRec As Integer, OurInfo As String
  Dim BookCnt As Integer, MtrCnt As Integer, IdxNumRecs As Long
  Dim RMCnt As Long, WhatRMRec As Long, BadgerFile As String
  Dim Account As String, BadgerType As String, DoneCnt As Long
  Dim Average As Double, LowRead As Double, NumberofRoutes As Integer
  Dim MeterID As String, MRDate As String, PathWay As String
  Dim HighRead As Double, ILowRead As Double, IdxFileName As String
  Dim UBBdgrGetRecLen As Integer, NumBdgrGetRecs As Integer, CircleC As String
  Dim MeterReadDate As String, Z9 As String, FFName As String, Mttype As String
  Dim DashPos As Integer, cnt As Long, UBBadRdRecLen As Integer
  Dim CurReading As Double, UBBadRdFile As Integer, NumBadRdRecs As Long
  Dim CRead As Double, IHiRead As String, ILoRead As String, MtNumb As String
  Dim CurRead As String, Acct As String, MTCnt As String, SEQNUMB As String
  Dim MT As String, MeterRecord As Integer, ReadingDate As Integer
  Dim OverDate As Integer, UseOverDate As Boolean, UpdCnt As Long, MtrNumb2 As String
  Dim cntnumsent As Long, CZCode As String
  PathWay$ = HHPathInOut
  ReDim Route(100)
  UBCustRecLen = Len(UBCustRec(1))
  Z9$ = "000000000"
  Z6$ = "000000"
  BadgerFile$ = "BADGERMR.DAT"
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen
  HighVar = UBSetUpRec(1).HighRead
  If HighVar > 0 Then
    HighVar = HighVar / 100
  Else
    HighVar = 1.5
  End If
  cntnumsent = 0
  LowVar = UBSetUpRec(1).LowRead
  If LowVar > 0 Then
    LowVar = LowVar / 100
  Else
    LowVar = 0.75
  End If
  If ImpExpFlag Then     'EXPORTING METER READINGS
    WhatTypes$ = Left$(Me.fpMtrType.Text, 1)
    GoSub BadgerSend
    MsgBox "Records sent to HH - " + Str(cntnumsent), vbOKOnly, "Export HH"
    Call cmdExit_Click
  Else
    GoSub BadgerGet
    MsgBox "Records retrieved from HH - " + Str(UpdCnt&), vbOKOnly, "Import HH"
    Call cmdExit_Click
  End If

  
BadgerReadExit:
  Exit Sub

BadgerSend:
  NumberofRoutes = 0

        GoSub MakeOUTTFFileName
        GoSub BadgerOpenCust    'Open Customer Data File

        ReDim UBBadRdRec(1) As UBBadgerV2RecType
        UBBadRdRecLen = Len(UBBadRdRec(1))

        KillFile FFName$
        UBBadRdFile = FreeFile
        Open FFName$ For Random Shared As UBBadRdFile Len = UBBadRdRecLen
        NumBadRdRecs = LOF(UBBadRdFile) / UBBadRdRecLen

    HighVar = UBSetUpRec(1).HighRead
    If UBSetUpRec(1).UseSeq = "Y" Then
      MakeSequenceIndex "Sequence Number", Me
      IdxFileName = UBPath$ + "UBTEMP.IDX"
    Else
      IdxFileName = UBPath$ + "UBCUSTBK.IDX"
    End If
    IdxRecLen = 4         'we are using a integer
    IdxFileSize& = FileSize&(IdxFileName)
    IdxNumOfRecs = IdxFileSize& \ IdxRecLen
    ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
    FGetAH IdxFileName, IdxBuff(), IdxRecLen, IdxNumOfRecs            'load it
  
  FrmShowPctComp.Label1 = "Exporting Reading Information."
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show '1, Parent

        DoneCnt& = 1
        Do
          Prec& = IdxBuff(DoneCnt&).RecNum
          If Prec& > 0 Then
            Get UBFile, Prec&, UBCustRec(1)
            For BookCnt = 1 To CycleCnt
              If Val(UBCustRec(1).Book) = Cycle(BookCnt) And (UBCustRec(1).Status <> "F") Then
                
                GoSub BadgerPutRec
              End If
            Next
          End If
          DoneCnt& = DoneCnt& + 1
         FrmShowPctComp.ShowPctComp DoneCnt&, IdxNumOfRecs&
        Loop Until DoneCnt& > IdxNumOfRecs&
        Close
Return

BadgerPutRec:

  'modifyed
  'May Have Up to 7 Meters to Read
  MtrCnt = 1
  Account$ = Str$(Prec&)
  While MtrCnt < 8
    If (Asc(UBCustRec(1).LocMeters(MtrCnt).MTRType) > 32) Then
      Select Case UBCustRec(1).LocMeters(MtrCnt).MTRType
      Case "C", "S", "W", "T", "E", "D", "G", "I"
        If UBCustRec(1).LocMeters(MtrCnt).MTRMulti <= 0 Then
          UBCustRec(1).LocMeters(MtrCnt).MTRMulti = 1
        End If
        Multi& = UBCustRec(1).LocMeters(MtrCnt).MTRMulti
        If UBCustRec(1).LocMeters(MtrCnt).MTRType = "G" Then
          Mttype = "G"
        Else
          Mttype = "W"
        End If
        Average# = UBCustRec(1).LocMeters(MtrCnt).AvgUse
        CRead# = UBCustRec(1).LocMeters(MtrCnt).CurRead
        'make sure we have valid average & current readings
        If CRead# < 0 Then
          CRead# = 0
        End If
        If Average# <= 0 Then
          Average# = CRead#
        End If
        IHiRead$ = Right$((Z9$ + QPTrim$(Str$(CRead# + Fix(Average# * HighVar)))), 9)
        ILoRead$ = Right$((Z9$ + QPTrim$(Str$(CRead# + Fix(Average# * LowVar)))), 9)
        CurRead$ = Right$((Z9$ + QPTrim$(Str$(0))), 9)

        Acct$ = UBCustRec(1).Book + UBCustRec(1).SEQNUMB
        If UBCustRec(1).Seq < 0 Then
          UBCustRec(1).Seq = 0
        End If
        MTCnt$ = QPTrim$(Str$(MtrCnt))
        SEQNUMB$ = Right$((Z6$ + QPTrim$(Str$(UBCustRec(1).Seq))), 6)
        OurInfo$ = QPTrim$(Str$(Prec&)) + "-" + MTCnt$

        MtNumb$ = QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrNum)
        MtrNumb2$ = QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrIDNO)
        If Left$(MtNumb$, 1) = "Z" Then
          MtNumb$ = Mid$(MtNumb$, 2, Len(MtNumb$))
          CZCode$ = "Z"
        ElseIf Left$(MtNumb$, 1) = "D" Then
          MtNumb$ = Mid$(MtNumb$, 2, Len(MtNumb$))
          CZCode$ = "D"
        Else
          CZCode$ = "C"
        End If

        Select Case Multi&
        Case 1
          CircleC$ = "01"
        Case 10
          CircleC$ = "02"
        Case 100
          CircleC$ = "03"
        Case 1000
          CircleC$ = "04"
        End Select

        LSet UBBadRdRec(1).Fill1 = ""
        LSet UBBadRdRec(1).CustName = QPTrim$(UBCustRec(1).CustName)
        LSet UBBadRdRec(1).ServAddr = QPTrim$(UBCustRec(1).ServAddr)
        LSet UBBadRdRec(1).MtrNum1 = MtrNumb2$
        LSet UBBadRdRec(1).Fill2 = ""
        LSet UBBadRdRec(1).Status = UBCustRec(1).Status
        LSet UBBadRdRec(1).MtrServ = MTCnt$
        LSet UBBadRdRec(1).MTRType = Mttype
        LSet UBBadRdRec(1).DNI = ""
        LSet UBBadRdRec(1).MtrSerN = MtNumb$
        LSet UBBadRdRec(1).Fill3 = ""
        RSet UBBadRdRec(1).HiRead = IHiRead$
        RSet UBBadRdRec(1).LoRead = ILoRead$
        LSet UBBadRdRec(1).CurrRead = CurRead$
        LSet UBBadRdRec(1).ReadTime = ""
        LSet UBBadRdRec(1).ReadType = ""
        LSet UBBadRdRec(1).TamperMode = ""
        LSet UBBadRdRec(1).CmntCode = ""
        LSet UBBadRdRec(1).RouteNum = UBCustRec(1).Book
        LSet UBBadRdRec(1).Account = Acct$  'location number
        LSet UBBadRdRec(1).ReadDate = ""
        LSet UBBadRdRec(1).ModType = CZCode$   'C for manual and Z for radio
        LSet UBBadRdRec(1).Latitude = Str(UBCustRec(1).LocMeters(MtrCnt).MtrLat)
        LSet UBBadRdRec(1).Fill3 = ""
        LSet UBBadRdRec(1).MtrRes = CircleC$  'this is set based on multiplier
        RSet UBBadRdRec(1).SEQNUMB = SEQNUMB$
        LSet UBBadRdRec(1).MfgModel = ""
        LSet UBBadRdRec(1).UserUse = OurInfo$  'this is cust acct# and number of meter
        LSet UBBadRdRec(1).ReadID = ""
        LSet UBBadRdRec(1).ReadCo1 = ""
        LSet UBBadRdRec(1).ReadCo2 = ""
        LSet UBBadRdRec(1).ReadCo3 = ""
        LSet UBBadRdRec(1).MtrLoc = ""
        LSet UBBadRdRec(1).MMIReadCode = ""
        LSet UBBadRdRec(1).Longitude = Str(UBCustRec(1).LocMeters(MtrCnt).MtrLng)
        LSet UBBadRdRec(1).Fill4 = ""
        LSet UBBadRdRec(1).CrLf = CrLf$
        Put UBBadRdFile, , UBBadRdRec(1)
        cntnumsent = cntnumsent + 1
      End Select
    End If
ESkipEm:
    MtrCnt = MtrCnt + 1
  Wend

Return

MakeOUTTFFileName:
  PathWay$ = QPTrim$(PathWay$)
  If Len(PathWay$) > 0 Then
    If Right$(PathWay$, 1) <> "\" Then
      PathWay$ = PathWay$ + "\"
    End If
  End If
  FFName$ = PathWay$ + "CONNECT.IN3"
Return

MakeINTFFileName:
  PathWay$ = QPTrim$(PathWay$)
  If Len(PathWay$) > 0 Then
    If Right$(PathWay$, 1) <> "\" Then
      PathWay$ = PathWay$ + "\"
    End If
  End If
  FFName$ = PathWay$ + "CONNECT.OT3"
Return

BadgerGet:
  NumberofRoutes = 0

      OverDate = Date2Num(fpReadDate)
      If OverDate > 0 Then
        UseOverDate = True
      Else
        UseOverDate = False
      End If
      GoSub MakeINTFFileName      'Get Badger File
      GoSub BadgerOpenCust      'Open Customer Data File

      ReDim UBBadRdRec(1) As UBBadgerV2RecType
      UBBadRdRecLen = Len(UBBadRdRec(1))
      'Open meter reading information File
      UBBadRdFile = FreeFile
      Open FFName$ For Random Shared As UBBadRdFile Len = UBBadRdRecLen
      NumBadRdRecs& = LOF(UBBadRdFile) / UBBadRdRecLen

      If NumBadRdRecs& = 0 Then
        Close
        MsgBox "No Records Found to Import.", vbOKOnly, "Procedure Ended"
        GoTo BadgerReadExit
      End If

  FrmShowPctComp.Label1 = "Importing Reading Information."
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show '1, Parent

      DoneCnt& = 1              ' Initialize File Counter to 1
      Do
        Get UBBadRdFile, DoneCnt&, UBBadRdRec(1)
        Prec& = Val(QPTrim$(UBBadRdRec(1).UserUse))
        If Prec& > 0 Then
          Get UBFile, Prec&, UBCustRec(1)
          GoSub EExtractRecord
        Else
          'Stop
        End If
        FrmShowPctComp.ShowPctComp DoneCnt&, NumBadRdRecs&
       DoneCnt& = DoneCnt& + 1
      Loop Until DoneCnt& > NumBadRdRecs&
      Close
      'Done = True

Return
EExtractRecord:
'this extracts the reading & date

  'QPrintRC " Updated Count:" + Str$(UpdCnt&), 11, 28, -1

  DashPos = InStr(UBBadRdRec(1).UserUse, "-")
  MT$ = Mid$(UBBadRdRec(1).UserUse, DashPos + 1)
  MeterRecord = Val(MT$)

  If MeterRecord = 0 Then MeterRecord = 1
  ' Check Meter Updated Flag
  ' Update Meter W/Reading
  CurReading# = Val(UBBadRdRec(1).CurrRead)

  If UBCustRec(1).LocMeters(MeterRecord).MTRMulti = 10 Then
    CurReading# = (CurReading# * 0.1)
  End If

  If UseOverDate Then
    ReadingDate = OverDate          'if they want to overide the read date
  Else
    MeterReadDate$ = Left$(UBBadRdRec(1).ReadDate, 2) + "/" + Mid$(UBBadRdRec(1).ReadDate, 3, 2) + "/" + Right$(UBBadRdRec(1).ReadDate, 4)
    ReadingDate = Date2Num(MeterReadDate$)
    If ReadingDate <= 0 Then
      ReadingDate = Date2Num(Date$) 'if the read date was bad then fix it
    End If
  End If

  If UBCustRec(1).LocMeters(MeterRecord).ReadFlag = "Y" Then
    UBCustRec(1).LocMeters(MeterRecord).CurRead = CurReading#
    UBCustRec(1).LocMeters(MeterRecord).CurDate = ReadingDate 'Date2Num(MeterReadDate$)
  Else
    UBCustRec(1).LocMeters(MeterRecord).PrevRead = UBCustRec(1).LocMeters(MeterRecord).CurRead
    UBCustRec(1).LocMeters(MeterRecord).PastDate = UBCustRec(1).LocMeters(MeterRecord).CurDate
    UBCustRec(1).LocMeters(MeterRecord).ReadFlag = "Y"
    UBCustRec(1).LocMeters(MeterRecord).CurDate = ReadingDate 'Date2Num(MeterReadDate$)
    UBCustRec(1).LocMeters(MeterRecord).CurRead = CurReading#
  End If
  UBCustRec(1).LocMeters(MeterRecord).MtrLat = Val(QPTrim$(UBBadRdRec(1).Latitude))
  UBCustRec(1).LocMeters(MeterRecord).MtrLng = Val(QPTrim$(UBBadRdRec(1).Longitude))
 'this was to fix switch nums during new meter retrieval using Y HH type
 'rem out after run once.
'  If QPTrim(UBBadRdRec(1).ModType) = "Z" Then
'    UBCustRec(1).LocMeters(MeterRecord).MtrNum = "Z" + UBBadRdRec(1).MtrSerN
'  Else
'    UBCustRec(1).LocMeters(MeterRecord).MtrNum = UBBadRdRec(1).MtrSerN
'  End If
'   UBCustRec(1).LocMeters(MeterRecord).MtrIDNO = UBBadRdRec(1).MtrNum1
 
  'now update the customers record...
  UpdCnt& = UpdCnt& + 1
  Put UBFile, Prec&, UBCustRec(1)
Return
BadgerOpenCust:
'  REDIM UBCustRec(1) AS NewUBCustRecType
'  UBCustRecLen = LEN(UBCustRec(1))
  UBFile = FreeFile
  Open "UBCUST.DAT" For Random Shared As UBFile Len = UBCustRecLen
Return

End Sub

'----------------------------------------------------------
Private Sub ImpExpBeaconV2HHInfo(ByVal ImpExpFlag As Boolean)
  Dim HighVar As Integer, LowVar As Integer
  Dim IdxFile As Integer, IdxNumRecs As Long, IdxFileName As String
  Dim WhatTypes As String, Z6 As String, Z9 As String, Account As String
  Dim CustAcc As String, Prec As Long, DoneCnt As Long
  Dim UBFile As Integer, BadgerCFGFile As Integer, Multi As Long
  Dim MeterID As String, MRDate As String, PathWay As String
  Dim MeterReadDate As String, FFName As String, Mttype As String
  Dim cntnumsent As Long, UpdCnt As Long, NumBadRdRecs As Long 'CZCode As String
  Dim UBBeaconRecLen As Integer, NumberofRoutes As Integer, UBBadRdFile As Integer
  Dim LocIDOut As String, tWrk1 As String, tWrk2 As String, tBld As String
  Dim BookCnt As Integer, MtrCnt As Integer, BeV2InRecCnt As Integer
  Dim Average As Double, CRead As Double ', LowRead As Double
  Dim IHiRead As String, ILoRead As String, CurRead As String, MtNumb As String
  Dim Acct As String, MTCnt As String, SEQNUMB As String, OurInfo As String
  Dim V$, eFlag As Boolean
  Dim BeaconTxtRec As String, lngCurrRead As Long
  Dim MT As String, MeterRecord As Integer, ReadingDate As Integer
  'Dim TTTCnt As Integer
  'Dim CurRead As String
'  Dim UBBadgerRecLen As Integer, NumBadgerRec As Integer, OurInfo As String
'  Dim RMCnt As Long, WhatRMRec As Long, BadgerFile As String
'  Dim Account As String, BadgerType As String
'  Dim HighRead As Double, ILowRead As Double
'  Dim UBBdgrGetRecLen As Integer, NumBdgrGetRecs As Integer, CircleC As String
'  Dim DashPos As Integer, cnt As Long, UBBadRdRecLen As Integer
'  Dim CurReading As Double, UBBadRdFile As Integer, NumBadRdRecs As Long

'  Dim MT As String, MeterRecord As Integer, ReadingDate As Integer
  Dim OverDate As Integer, UseOverDate As Boolean ', , MtrNumb2 As String
  PathWay$ = HHPathInOut
  ReDim Route(100)
  UBCustRecLen = Len(UBCustRec(1))
  V$ = "|"
  Z9$ = "000000000"
  Z6$ = "000000"
  'BadgerFile$ = "BADGERMR.DAT"
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen
  HighVar = UBSetUpRec(1).HighRead
  If HighVar > 0 Then
    HighVar = HighVar / 100
  Else
    HighVar = 1.5
  End If
  cntnumsent = 0
  LowVar = UBSetUpRec(1).LowRead
  If LowVar > 0 Then
    LowVar = LowVar / 100
  Else
    LowVar = 0.75
  End If
  If ImpExpFlag Then     'EXPORTING METER READINGS
    WhatTypes$ = Left$(Me.fpMtrType.Text, 1)
    GoSub BeaconV2Send
    MsgBox "Records sent to HH - " + Str(cntnumsent), vbOKOnly, "Export HH"
    Call cmdExit_Click
  Else
    GoSub BeaconV2Get
    MsgBox "  Records imported from HH -  " + Str(UpdCnt&), vbOKOnly, "Import HH"
    Call cmdExit_Click
  End If
 
BeaconV2ReadExit:
  Exit Sub

BeaconV2Send:
    NumberofRoutes = 0
'HERE
    GoSub MakeBeaconV2OutFileName
    GoSub BeaconV2OpenCust    'Open Customer Data File
    
    ReDim UBBeacRdRec(1) As UBBadgerBeaconRecType
    UBBeaconRecLen = Len(UBBeacRdRec(1))
    
    Prec& = LOF(UBFile) / Len(UBCustRec(1))

    KillFile FFName$
    UBBadRdFile = FreeFile
    Open FFName$ For Output As UBBadRdFile Len = UBBeaconRecLen
    
    Print #UBBadRdFile, BeaconHeader

    HighVar = UBSetUpRec(1).HighRead
    If UBSetUpRec(1).UseSeq = "Y" Then
      MakeSequenceIndex "Sequence Number", Me
      IdxFileName = UBPath$ + "UBTEMP.IDX"
    Else
      IdxFileName = UBPath$ + "UBCUSTBK.IDX"
    End If
    IdxRecLen = 4         'we are using a integer
    IdxFileSize& = FileSize&(IdxFileName)
    IdxNumOfRecs = IdxFileSize& \ IdxRecLen
    ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
    FGetAH IdxFileName, IdxBuff(), IdxRecLen, IdxNumOfRecs            'load it
  
    FrmShowPctComp.Label1 = "Exporting Reading Information."
    FrmShowPctComp.cmdCancel.Enabled = False
    FrmShowPctComp.Show '1, Parent

    DoneCnt& = 1
    Do
        Prec& = IdxBuff(DoneCnt&).RecNum
        If Prec& > 0 Then
            Get UBFile, Prec&, UBCustRec(1)
            For BookCnt = 1 To CycleCnt
                If Val(UBCustRec(1).Book) = Cycle(BookCnt) And (UBCustRec(1).Status <> "F") Then
                    GoSub BeaconV2PutRec
                    'TTTCnt = TTTCnt + 1
                End If
            Next
        End If
        DoneCnt& = DoneCnt& + 1
        FrmShowPctComp.ShowPctComp DoneCnt&, IdxNumOfRecs&
        'If TTTCnt > 9 Then Exit Do
    Loop Until DoneCnt& > IdxNumOfRecs&
    Close

Return

BeaconV2PutRec:
  tWrk1 = QPTrim(UBCustRec(1).Book)
  Select Case Len(tWrk1)
  Case 0
    tWrk1 = "00"
  Case 1
    tWrk1 = "0" + tWrk1
  Case 2
    tWrk1 = tWrk1
  End Select
  
  tWrk2 = QPTrim(UBCustRec(1).SEQNUMB)
      
  If Len(tWrk2) < 6 Then
    tWrk2 = String$(6 - Len(tWrk2), "0") + tWrk2
  End If
    
  tBld = Right$("0000" + QPTrim(CStr(Prec)), 4)
  
  LocIDOut = String$(40, "0") + tWrk1 + tWrk2 + tBld
  LocIDOut = Right$(LocIDOut, 40)
  
  MtrCnt = 1
  Account$ = Str$(Prec&)
'step thru 7 possiable meters
  While MtrCnt < 8
    If (Asc(UBCustRec(1).LocMeters(MtrCnt).MTRType) > 32) Then
        ReDim UBBeacRdRec(1) As UBBadgerBeaconRecType
        Select Case UBCustRec(1).LocMeters(MtrCnt).MTRType
        Case "C", "S", "W", "T", "E", "D"         'here dale
            If UBCustRec(1).LocMeters(MtrCnt).MTRMulti <= 0 Then
                UBCustRec(1).LocMeters(MtrCnt).MTRMulti = 1
            End If
            Multi& = UBCustRec(1).LocMeters(MtrCnt).MTRMulti
            Average# = UBCustRec(1).LocMeters(MtrCnt).AvgUse
            CRead# = UBCustRec(1).LocMeters(MtrCnt).CurRead
            'make sure we have valid average & current readings
            If CRead# < 0 Then
                CRead# = 0
            End If
            If Average# <= 0 Then
                Average# = CRead#
            End If
            IHiRead$ = Right$((Z9$ + QPTrim$(Str$(CRead# + Fix(Average# * HighVar)))), 9)
            ILoRead$ = Right$((Z9$ + QPTrim$(Str$(CRead# + Fix(Average# * LowVar)))), 9)
            CurRead$ = Right$((Z9$ + QPTrim$(Str$(CRead#))), 9)
            Acct$ = UBCustRec(1).Book + UBCustRec(1).SEQNUMB
            If UBCustRec(1).Seq < 0 Then
                UBCustRec(1).Seq = 0
            End If
            MTCnt$ = QPTrim$(Str$(MtrCnt))
            SEQNUMB$ = QPTrim$(Str$(UBCustRec(1).Seq))
            OurInfo$ = QPTrim$(Str$(Prec)) + "-" + MTCnt$
            MtNumb$ = QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrNum)
'insert the info into the beacon record
'Account
            LSet UBBeacRdRec(1).AcctID = QPTrim(Acct$)
            LSet UBBeacRdRec(1).AcctFName = ""
            LSet UBBeacRdRec(1).AcctLName = ""
            LSet UBBeacRdRec(1).AcctFullName = QPTrim$(UBCustRec(1).CustName)
            LSet UBBeacRdRec(1).AcctEMail = ""
            LSet UBBeacRdRec(1).AcctPhone = ""
            LSet UBBeacRdRec(1).AcctPhoneExt = ""
            LSet UBBeacRdRec(1).AcctAddr1 = QPTrim$(UBCustRec(1).ADDR1)
            LSet UBBeacRdRec(1).AcctAddr2 = QPTrim$(UBCustRec(1).ADDR2)
            LSet UBBeacRdRec(1).AcctAddr3 = ""
            LSet UBBeacRdRec(1).AcctCity = QPTrim$(UBCustRec(1).CITY)
            LSet UBBeacRdRec(1).AcctState = QPTrim$(UBCustRec(1).STATE)
            LSet UBBeacRdRec(1).AcctZip = QPTrim$(UBCustRec(1).ZIPCODE)
            LSet UBBeacRdRec(1).AcctCountry = "USA"
            LSet UBBeacRdRec(1).AcctPerID = ""
            LSet UBBeacRdRec(1).AcctStatus = QPTrim$(UBCustRec(1).Status)
            LSet UBBeacRdRec(1).AcctPortStat = ""
            LSet UBBeacRdRec(1).AcctBLCyc = ""           'not Citipak's billing cycle
            LSet UBBeacRdRec(1).AcctPaperless = ""
            LSet UBBeacRdRec(1).AcctAutoPay = ""
            LSet UBBeacRdRec(1).AcctBillerAP = ""
''Location
            LSet UBBeacRdRec(1).LocID = LocIDOut
            LSet UBBeacRdRec(1).LocName = QPTrim$(UBCustRec(1).CustName)
            LSet UBBeacRdRec(1).LocAddParity = ""
            LSet UBBeacRdRec(1).LocAddr1 = QPTrim$(UBCustRec(1).ServAddr)
            LSet UBBeacRdRec(1).LocAddr2 = ""
            LSet UBBeacRdRec(1).LocAddr3 = ""
            LSet UBBeacRdRec(1).LocCity = QPTrim$(UBCustRec(1).CITY)
            LSet UBBeacRdRec(1).LocState = QPTrim$(UBCustRec(1).STATE)
            LSet UBBeacRdRec(1).LocZip = QPTrim$(UBCustRec(1).ZIPCODE)
            LSet UBBeacRdRec(1).LocCountyName = ""
            LSet UBBeacRdRec(1).LocCountry = "USA"
            LSet UBBeacRdRec(1).LocLatitude = QPTrim$(Str$(UBCustRec(1).LocMeters(MtrCnt).MtrLat))
            LSet UBBeacRdRec(1).LocLongitude = QPTrim$(Str$(UBCustRec(1).LocMeters(MtrCnt).MtrLng))
'Tags
            LSet UBBeacRdRec(1).TagLocBldType = ""
            LSet UBBeacRdRec(1).TagLocBldNumb = ""
            LSet UBBeacRdRec(1).TagLocSite = ""
            LSet UBBeacRdRec(1).TagLocFunding = ""
            LSet UBBeacRdRec(1).TagLocMainUse = ""
            LSet UBBeacRdRec(1).TagLocWatType = ""
            LSet UBBeacRdRec(1).TagLocArea = ""
            LSet UBBeacRdRec(1).TagLocIrrArea = ""
            LSet UBBeacRdRec(1).TagLocPopu = ""
            LSet UBBeacRdRec(1).TagLocWFR = ""
            LSet UBBeacRdRec(1).TagLocIrr = ""
            LSet UBBeacRdRec(1).TagLocYearBuilt = ""
            LSet UBBeacRdRec(1).TagLocPool = ""
            LSet UBBeacRdRec(1).TagLocBathrooms = ""
            LSet UBBeacRdRec(1).TagLocDistrict = ""
            LSet UBBeacRdRec(1).TagLocDHSCode = ""
            LSet UBBeacRdRec(1).TagLocParcNumb = ""
            LSet UBBeacRdRec(1).TagLocETJan = ""
            LSet UBBeacRdRec(1).TagLocETFeb = ""
            LSet UBBeacRdRec(1).TagLocETMar = ""
            LSet UBBeacRdRec(1).TagLocETApr = ""
            LSet UBBeacRdRec(1).TagLocETMay = ""
            LSet UBBeacRdRec(1).TagLocETJun = ""
            LSet UBBeacRdRec(1).TagLocETJul = ""
            LSet UBBeacRdRec(1).TagLocETAug = ""
            LSet UBBeacRdRec(1).TagLocETSep = ""
            LSet UBBeacRdRec(1).TagLocETOct = ""
            LSet UBBeacRdRec(1).TagLocETNov = ""
            LSet UBBeacRdRec(1).TagLocETDec = ""
'Service Point
            'If Len(QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrIDNO)) > 0 Then
              'Stop
            'End If
            LSet UBBeacRdRec(1).SrvPntID = MTCnt$
            'QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrIDNO)
            LSet UBBeacRdRec(1).SrvPntType = "W"
            LSet UBBeacRdRec(1).SrvPntCycle = ""
            LSet UBBeacRdRec(1).SrvPntRoute = ""
            LSet UBBeacRdRec(1).SrvPntCsCode = ""
            LSet UBBeacRdRec(1).SrvPntCsCodeNm = ""
            LSet UBBeacRdRec(1).SrvPntLatitude = ""
            LSet UBBeacRdRec(1).SrvPntLongitude = ""
            LSet UBBeacRdRec(1).SrvPntTimeZone = ""
'Meter
            LSet UBBeacRdRec(1).MtrID = OurInfo$  'QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrIDNO)
            LSet UBBeacRdRec(1).MtrSerNo = QPTrim$(MtNumb$)
            LSet UBBeacRdRec(1).MtrManufact = ""
            LSet UBBeacRdRec(1).MtrModel = ""
            LSet UBBeacRdRec(1).MtrSize = ""
            LSet UBBeacRdRec(1).MtrSizeUnit = ""
            LSet UBBeacRdRec(1).MtrNote = ""
            LSet UBBeacRdRec(1).MtrContFlow = ""
            LSet UBBeacRdRec(1).MtrRegNumb = ""
            LSet UBBeacRdRec(1).MtrRegUOM = "GAL"
            LSet UBBeacRdRec(1).MtrRegReso = QPTrim(UBCustRec(1).USERCODE1)
            'LSet UBBeacRdRec(1).MtrRegReso = QPTrim(Str(Multi&))
            LSet UBBeacRdRec(1).MtrInstDate = ""
            LSet UBBeacRdRec(1).MtrInstStrRead = ""
            LSet UBBeacRdRec(1).MtrRemDate = ""
            LSet UBBeacRdRec(1).MtrRemRead = ""
'Service Agreement
            LSet UBBeacRdRec(1).SAStartDate = Num2Date$(UBCustRec(1).LocMeters(MtrCnt).PastDate)
            LSet UBBeacRdRec(1).SAEndDate = ""
'Endpoint Config
            LSet UBBeacRdRec(1).EndPointSerNo = QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrIDNO)  'OurInfo$ '
            LSet UBBeacRdRec(1).EndPointType = QPTrim(UBCustRec(1).USERCODE2) ' "1"  'Will be the UserCode1 field
            LSet UBBeacRdRec(1).EndPointInsDate = ""
            LSet UBBeacRdRec(1).EndPointRemDate = ""
''Manual/Mobile Readings
            LSet UBBeacRdRec(1).ManMobSequ = ""
            LSet UBBeacRdRec(1).ManMobAlertCode = ""
            LSet UBBeacRdRec(1).ManMobHighRead = ""
            LSet UBBeacRdRec(1).ManMobLowRead = ""
''Mics Info
            LSet UBBeacRdRec(1).MIUse1 = OurInfo$
            LSet UBBeacRdRec(1).MIUse2 = ""
'Write it to the file
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).AcctID) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).AcctFName) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).AcctLName) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).AcctFullName) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).AcctEMail) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).AcctPhone) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).AcctPhoneExt) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).AcctAddr1) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).AcctAddr2) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).AcctAddr3) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).AcctCity) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).AcctState) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).AcctZip) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).AcctCountry) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).AcctPerID) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).AcctStatus) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).AcctPortStat) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).AcctBLCyc) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).AcctPaperless) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).AcctAutoPay) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).AcctBillerAP) + V$;
'''Location
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).LocID) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).LocName) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).LocAddParity) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).LocAddr1) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).LocAddr2) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).LocAddr3) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).LocCity) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).LocState) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).LocZip) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).LocCountyName) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).LocCountry) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).LocLatitude) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).LocLongitude) + V$;
''Tags
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).TagLocBldType) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).TagLocBldNumb) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).TagLocSite) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).TagLocFunding) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).TagLocMainUse) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).TagLocWatType) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).TagLocArea) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).TagLocIrrArea) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).TagLocPopu) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).TagLocWFR) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).TagLocIrr) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).TagLocYearBuilt) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).TagLocPool) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).TagLocBathrooms) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).TagLocDistrict) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).TagLocDHSCode) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).TagLocParcNumb) + V$;
            
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).TagLocETJan) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).TagLocETFeb) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).TagLocETMar) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).TagLocETApr) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).TagLocETMay) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).TagLocETJun) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).TagLocETJul) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).TagLocETAug) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).TagLocETSep) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).TagLocETOct) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).TagLocETNov) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).TagLocETDec) + V$;
            
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).SrvPntID) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).SrvPntType) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).SrvPntCycle) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).SrvPntRoute) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).SrvPntCsCode) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).SrvPntCsCodeNm) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).SrvPntLatitude) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).SrvPntLongitude) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).SrvPntTimeZone) + V$;
''Meter
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).MtrID) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).MtrSerNo) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).MtrManufact) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).MtrModel) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).MtrSize) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).MtrSizeUnit) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).MtrNote) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).MtrContFlow) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).MtrRegNumb) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).MtrRegUOM) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).MtrRegReso) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).MtrInstDate) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).MtrInstStrRead) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).MtrRemDate) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).MtrRemRead) + V$;
''Service Agreement
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).SAStartDate) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).SAEndDate) + V$;
''Endpoint Config
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).EndPointSerNo) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).EndPointType) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).EndPointInsDate) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).EndPointRemDate) + V$;
'''Manual/Mobile Readings
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).ManMobSequ) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).ManMobAlertCode) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).ManMobHighRead) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).ManMobLowRead) + V$;
'''Mics Info
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).MIUse1) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).MIUse2)
            cntnumsent = cntnumsent + 1
'--------------------------------------------
        End Select
    End If

    MtrCnt = MtrCnt + 1
  Wend


Return ' End V2 Put

'--------- Beacon V2 Get
BeaconV2Get:
      NumberofRoutes = 0
      eFlag = False
      OverDate = Date2Num(fpReadDate)
      If OverDate > 0 Then
        UseOverDate = True
      Else
        UseOverDate = False
      End If
      
      GoSub MakeBeaconV2INFileName      'Get Badger File
      
      GoSub BeaconV2GetINRecCnt
      '"CONNECT.OT3"
      Open "fixed_ot3.txt" For Output As #50
      
'8-8-8-8-8-8-8-8-8-8-8-8-8-8-8-8-8-8-8
      If BeV2InRecCnt < 2 Then  '1st record is a header
        Close
        MsgText(0) = "Beacon Import Operation."
        MsgText(1) = "ERROR!"
        MsgText(2) = ""
        MsgText(3) = "No Records found to Import!"
        MsgText(4) = ""
        MsgText(5) = ""
        Call GetOKorNot%(MsgText(), True, True)
        'MsgBox "No Records Found to Import.", vbOKOnly, "Procedure Ended"
        GoTo BeaconV2ReadExit
        'no records error exit
      End If
      
      GoSub BeaconV2OpenCust      'Open Customer Data File
      
      UBBadRdFile = FreeFile
      Open FFName$ For Input As UBBadRdFile 'Len = UBBeaconRecLen

      FrmShowPctComp.Label1 = "Importing Reading Information."
      FrmShowPctComp.cmdCancel.Enabled = False
      FrmShowPctComp.Show '1, Parent

      DoneCnt& = 1              ' Initialize File Counter to 1
      Line Input #UBBadRdFile, BeaconTxtRec
      Print #50, BeaconTxtRec
      Do
    
        Line Input #UBBadRdFile, BeaconTxtRec
        
        GetBeaCustAndMtrID BeaconTxtRec, Prec&, MeterRecord
        
        If Prec& > 0 Then
          Get UBFile, Prec&, UBCustRec(1)
          GoSub BeaconV2ExtractRec
        Else
          eFlag = True
        End If
        FrmShowPctComp.ShowPctComp DoneCnt&, NumBadRdRecs&
        DoneCnt& = DoneCnt& + 1
      Loop Until eof(UBBadRdFile) 'DoneCnt& > BeInRecCnt
    Close
    eFlag = False
    
    If eFlag Then
      frmMsgDialog.Caption = "WARNING!"
      frmMsgDialog.RetLabel = "-2"
      frmMsgDialog.Label(0).Caption = ""
      frmMsgDialog.Label(1).Caption = "Invalid reading records "
      frmMsgDialog.Label(2).Caption = "found in import file!"
      frmMsgDialog.Label(3).Caption = "Please call Southern Software for"
      frmMsgDialog.Label(4).Caption = "additional Information."
      frmMsgDialog.Show vbModal
'    Else
'      frmDataUpdated.Show vbModal
    End If
'8-8-8-8-8-8-8-8-8-8-8-8-8-8-8-8-8-8-8

Return

BeaconV2ExtractRec:
'this extracts the reading & date

  UpdCnt& = UpdCnt& + 1

  If MeterRecord = 0 Then MeterRecord = 1
  
  GetBeaCurReadAndDate BeaconTxtRec, lngCurrRead, ReadingDate
'   If InStr(BeaconTxtRec, "311-1") > 0 Then Stop
  If ReadingDate < 0 Then
    eFlag = True
    GoTo BeaconV2ExtractErrorExit
  End If
  
  If UseOverDate Then
    ReadingDate = OverDate          'if they want to overide the read date
  End If

  If UBCustRec(1).LocMeters(MeterRecord).ReadFlag = "Y" Then
    UBCustRec(1).LocMeters(MeterRecord).CurRead = lngCurrRead
    UBCustRec(1).LocMeters(MeterRecord).CurDate = ReadingDate 'Date2Num(MeterReadDate$)
  Else
    UBCustRec(1).LocMeters(MeterRecord).PrevRead = UBCustRec(1).LocMeters(MeterRecord).CurRead
    UBCustRec(1).LocMeters(MeterRecord).PastDate = UBCustRec(1).LocMeters(MeterRecord).CurDate
    UBCustRec(1).LocMeters(MeterRecord).ReadFlag = "Y"
    UBCustRec(1).LocMeters(MeterRecord).CurDate = ReadingDate 'Date2Num(MeterReadDate$)
    UBCustRec(1).LocMeters(MeterRecord).CurRead = lngCurrRead
  End If
'  'now update the customers record...
  
  Put UBFile, Prec&, UBCustRec(1)

BeaconV2ExtractErrorExit:
Return


BeaconV2OpenCust:
  UBFile = FreeFile
  Open "UBCUST.DAT" For Random Shared As UBFile Len = UBCustRecLen
Return

MakeBeaconV2OutFileName:
  PathWay$ = QPTrim$(PathWay$)
  If Len(PathWay$) > 0 Then
    If Right$(PathWay$, 1) <> "\" Then
      PathWay$ = PathWay$ + "\"
    End If
  End If
  FFName$ = PathWay$ + "CONNECT.IN3"
Return

MakeBeaconV2INFileName:
  PathWay$ = QPTrim$(PathWay$)
  If Len(PathWay$) > 0 Then
    If Right$(PathWay$, 1) <> "\" Then
      PathWay$ = PathWay$ + "\"
    End If
  End If
  FFName$ = PathWay$ + "CONNECT.OT3"
Return

BeaconV2GetINRecCnt:
  BeV2InRecCnt = 0
  UBBadRdFile = FreeFile
  If Exist(FFName$) Then
    Open FFName$ For Input Shared As UBBadRdFile
    If LOF(UBBadRdFile) > 0 Then
      While Not eof(UBBadRdFile)
        Line Input #UBBadRdFile, BeaconTxtRec
        BeV2InRecCnt = BeV2InRecCnt + 1
      Wend
    End If
    Close UBBadRdFile
 End If
Return

End Sub
'----------------------------------------------------------

'**********************************************
Private Sub ImpExpBeaconHHInfo(ByVal ImpExpFlag As Boolean)    'BEACON
  
  Dim HighVar As Integer, LowVar As Integer, IdxFile As Integer
  Dim WhatTypes As String, CustAcc As String, Prec As Long
  Dim UBFile As Integer, BadgerCFGFile As Integer, Multi As Long
  Dim UBBadgerRecLen As Integer, NumBadgerRec As Integer, OurInfo As String
  Dim BookCnt As Integer, MtrCnt As Integer, IdxNumRecs As Long
  Dim RMCnt As Long, WhatRMRec As Long, BadgerFile As String
  Dim Account As String ', BadgerType As String, DoneCnt As Long
  Dim Average As Double, LowRead As Double, NumberofRoutes As Integer
  Dim MeterID As String, MRDate As String, PathWay As String
  Dim HighRead As Double, ILowRead As Double, IdxFileName As String
  Dim MeterReadDate As String, Z9 As String, FFName As String
  Dim ExpCnt As Long
  Dim UBBeaconRecLen As Integer
  Dim CurReading As Double, UBBadRdFile As Integer, NumBadRdRecs As Long
  Dim CRead As Double, IHiRead As String, ILoRead As String, MtNumb As String
  Dim CurRead As String, Acct As String, MTCnt As String, SEQNUMB As String
  Dim MT As String, MeterRecord As Integer, ReadingDate As Integer
  Dim OverDate As Integer, UseOverDate As Boolean, UpdCnt As Long
  Dim V$, eFlag As Boolean
  Dim LocIDOut As String, tWrk1 As String, tWrk2 As String, tBld As String
  Dim BeInRecCnt As Integer, DoneCnt As Long, BConCustRec As Integer
  Dim BeaconTxtRec As String, lngCurrRead As Long
  
  PathWay$ = HHPathInOut
  ReDim Route(100)
  UBCustRecLen = Len(UBCustRec(1))
  V$ = "|"
  Z9$ = "000000000"
  'BadgerFile$ = "BADGERMR.DAT"
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen
  HighVar = UBSetUpRec(1).HighRead
  If HighVar > 0 Then
    HighVar = HighVar / 100
  Else
    HighVar = 1.5
  End If

  LowVar = UBSetUpRec(1).LowRead
  If LowVar > 0 Then
    LowVar = LowVar / 100
  Else
    LowVar = 0.75
  End If
 
  If ImpExpFlag Then     'EXPORTING METER READINGS
    WhatTypes$ = Left$(Me.fpMtrType.Text, 1)
    GoSub BadgerBeaconSend
    MsgText(0) = "Beacon Export Operation."
    MsgText(1) = "Meter Reading Export"
    MsgText(2) = ""
    MsgText(3) = "Exported:" + Str$(ExpCnt) + " Readings."
    MsgText(4) = ""
    MsgText(5) = "Process Complete."
    Call GetOKorNot%(MsgText(), True, True)
    Call cmdExit_Click
  Else
    GoSub BadgerBeaconGet
    MsgText(0) = "Beacon Import Operation."
    MsgText(1) = ""
    MsgText(2) = "Meter Reading Import"
    MsgText(3) = ""
    MsgText(4) = "Process Complete."
    MsgText(5) = ""
    Call GetOKorNot%(MsgText(), True, True)
    Call cmdExit_Click
  End If
  
BadgerBeaconReadExit:
  Exit Sub

'Beacon Output records
BadgerBeaconSend:
    NumberofRoutes = 0
    ExpCnt = 0
    GoSub MakeBeaconOUTTFFileName
    GoSub BadgerBeaconOpenCust    'Open Customer Data File
    
    ReDim UBBeacRdRec(1) As UBBadgerBeaconRecType
    UBBeaconRecLen = Len(UBBeacRdRec(1))
    Prec& = LOF(UBFile) / Len(UBCustRec(1))
    KillFile FFName$
    UBBadRdFile = FreeFile
    Open FFName$ For Output As UBBadRdFile Len = UBBeaconRecLen
    
    Print #UBBadRdFile, BeaconHeader
  
    FrmShowPctComp.Label1 = "Exporting Reading Information."
    FrmShowPctComp.cmdCancel.Enabled = False
    FrmShowPctComp.Show '1, Parent
    DoEvents
    'DoneCnt& = 1
    For DoneCnt& = 1 To Prec&
      Get UBFile, DoneCnt&, UBCustRec(1)
        If UBCustRec(1).BILLCYCL = "99" And (UBCustRec(1).Status <> "F") Then
          If Len(QPTrim(UBCustRec(1).Book)) > 0 Then
            GoSub BadgerBeaconPutRec
            ExpCnt = ExpCnt + 1
          End If
        End If
        FrmShowPctComp.ShowPctComp DoneCnt&, Prec&
    Next
    Close
Return

BadgerBeaconPutRec:
  'May Have Up to 7 Meters to Read
  tWrk1 = QPTrim(UBCustRec(1).Book)
  
  Select Case Len(tWrk1)
  Case 0
    tWrk1 = "00"
  Case 1
    tWrk1 = "0" + tWrk1
  Case 2
    tWrk1 = tWrk1
  End Select
  
  tWrk2 = QPTrim(UBCustRec(1).SEQNUMB)
      
  If Len(tWrk2) < 6 Then
    tWrk2 = String$(6 - Len(tWrk2), "0") + tWrk2
  End If
    
  tBld = Right$("0000" + QPTrim(CStr(Prec)), 4)
  
  LocIDOut = String$(40, "0") + tWrk1 + tWrk2 + tBld
  LocIDOut = Right$(LocIDOut, 40)
  
  MtrCnt = 1
  Account$ = Str$(Prec&)
'step thru 7 possiable meters
  While MtrCnt < 8
    If (Asc(UBCustRec(1).LocMeters(MtrCnt).MTRType) > 32) Then
        ReDim UBBeacRdRec(1) As UBBadgerBeaconRecType
        Select Case UBCustRec(1).LocMeters(MtrCnt).MTRType
        Case "C", "S", "W", "T", "E", "D"         'here dale
            If UBCustRec(1).LocMeters(MtrCnt).MTRMulti <= 0 Then
                UBCustRec(1).LocMeters(MtrCnt).MTRMulti = 1
            End If
            Multi& = UBCustRec(1).LocMeters(MtrCnt).MTRMulti
            Average# = UBCustRec(1).LocMeters(MtrCnt).AvgUse
            CRead# = UBCustRec(1).LocMeters(MtrCnt).CurRead
            'make sure we have valid average & current readings
            If CRead# < 0 Then
                CRead# = 0
            End If
            If Average# <= 0 Then
                Average# = CRead#
            End If
            IHiRead$ = Right$((Z9$ + QPTrim$(Str$(CRead# + Fix(Average# * HighVar)))), 9)
            ILoRead$ = Right$((Z9$ + QPTrim$(Str$(CRead# + Fix(Average# * LowVar)))), 9)
            CurRead$ = Right$((Z9$ + QPTrim$(Str$(CRead#))), 9)
            Acct$ = UBCustRec(1).Book + UBCustRec(1).SEQNUMB
            If UBCustRec(1).Seq < 0 Then
                UBCustRec(1).Seq = 0
            End If
            MTCnt$ = QPTrim$(Str$(MtrCnt))
            SEQNUMB$ = QPTrim$(Str$(UBCustRec(1).Seq))
            OurInfo$ = QPTrim$(Str$(DoneCnt&)) + "-" + MTCnt$
            MtNumb$ = QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrNum)
'insert the info into the beacon record
'Account
            LSet UBBeacRdRec(1).AcctID = QPTrim(Acct$)
            LSet UBBeacRdRec(1).AcctFName = ""
            LSet UBBeacRdRec(1).AcctLName = ""
            LSet UBBeacRdRec(1).AcctFullName = QPTrim$(UBCustRec(1).CustName)
            LSet UBBeacRdRec(1).AcctEMail = ""
            LSet UBBeacRdRec(1).AcctPhone = ""
            LSet UBBeacRdRec(1).AcctPhoneExt = ""
            LSet UBBeacRdRec(1).AcctAddr1 = QPTrim$(UBCustRec(1).ADDR1)
            LSet UBBeacRdRec(1).AcctAddr2 = QPTrim$(UBCustRec(1).ADDR2)
            LSet UBBeacRdRec(1).AcctAddr3 = ""
            LSet UBBeacRdRec(1).AcctCity = QPTrim$(UBCustRec(1).CITY)
            LSet UBBeacRdRec(1).AcctState = QPTrim$(UBCustRec(1).STATE)
            LSet UBBeacRdRec(1).AcctZip = QPTrim$(UBCustRec(1).ZIPCODE)
            LSet UBBeacRdRec(1).AcctCountry = "USA"
            LSet UBBeacRdRec(1).AcctPerID = ""
            LSet UBBeacRdRec(1).AcctStatus = QPTrim$(UBCustRec(1).Status)
            LSet UBBeacRdRec(1).AcctPortStat = ""
            LSet UBBeacRdRec(1).AcctBLCyc = ""           'not Citipak's billing cycle
            LSet UBBeacRdRec(1).AcctPaperless = ""
            LSet UBBeacRdRec(1).AcctAutoPay = ""
            LSet UBBeacRdRec(1).AcctBillerAP = ""

''Location
            LSet UBBeacRdRec(1).LocID = LocIDOut
            LSet UBBeacRdRec(1).LocName = QPTrim$(UBCustRec(1).CustName)
            LSet UBBeacRdRec(1).LocAddParity = ""
            LSet UBBeacRdRec(1).LocAddr1 = QPTrim$(UBCustRec(1).ServAddr)
            LSet UBBeacRdRec(1).LocAddr2 = ""
            LSet UBBeacRdRec(1).LocAddr3 = ""
            LSet UBBeacRdRec(1).LocCity = QPTrim$(UBCustRec(1).CITY)
            LSet UBBeacRdRec(1).LocState = QPTrim$(UBCustRec(1).STATE)
            LSet UBBeacRdRec(1).LocZip = QPTrim$(UBCustRec(1).ZIPCODE)
            LSet UBBeacRdRec(1).LocCountyName = ""
            LSet UBBeacRdRec(1).LocCountry = "USA"
            LSet UBBeacRdRec(1).LocLatitude = QPTrim$(Str$(UBCustRec(1).LocMeters(MtrCnt).MtrLat))
            LSet UBBeacRdRec(1).LocLongitude = QPTrim$(Str$(UBCustRec(1).LocMeters(MtrCnt).MtrLng))
'Tags
            LSet UBBeacRdRec(1).TagLocBldType = ""
            LSet UBBeacRdRec(1).TagLocBldNumb = ""
            LSet UBBeacRdRec(1).TagLocSite = ""
            LSet UBBeacRdRec(1).TagLocFunding = ""
            LSet UBBeacRdRec(1).TagLocMainUse = ""
            LSet UBBeacRdRec(1).TagLocWatType = ""
            LSet UBBeacRdRec(1).TagLocArea = ""
            LSet UBBeacRdRec(1).TagLocIrrArea = ""
            LSet UBBeacRdRec(1).TagLocPopu = ""
            LSet UBBeacRdRec(1).TagLocWFR = ""
            LSet UBBeacRdRec(1).TagLocIrr = ""
            LSet UBBeacRdRec(1).TagLocYearBuilt = ""
            LSet UBBeacRdRec(1).TagLocPool = ""
            LSet UBBeacRdRec(1).TagLocBathrooms = ""
            LSet UBBeacRdRec(1).TagLocDistrict = ""
            LSet UBBeacRdRec(1).TagLocDHSCode = ""
            LSet UBBeacRdRec(1).TagLocParcNumb = ""
            LSet UBBeacRdRec(1).TagLocETJan = ""
            LSet UBBeacRdRec(1).TagLocETFeb = ""
            LSet UBBeacRdRec(1).TagLocETMar = ""
            LSet UBBeacRdRec(1).TagLocETApr = ""
            LSet UBBeacRdRec(1).TagLocETMay = ""
            LSet UBBeacRdRec(1).TagLocETJun = ""
            LSet UBBeacRdRec(1).TagLocETJul = ""
            LSet UBBeacRdRec(1).TagLocETAug = ""
            LSet UBBeacRdRec(1).TagLocETSep = ""
            LSet UBBeacRdRec(1).TagLocETOct = ""
            LSet UBBeacRdRec(1).TagLocETNov = ""
            LSet UBBeacRdRec(1).TagLocETDec = ""
'Service Point
            'If Len(QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrIDNO)) > 0 Then
              'Stop
            'End If
            LSet UBBeacRdRec(1).SrvPntID = "" 'QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrIDNO)
            LSet UBBeacRdRec(1).SrvPntType = "W"
            LSet UBBeacRdRec(1).SrvPntCycle = ""
            LSet UBBeacRdRec(1).SrvPntRoute = ""
            LSet UBBeacRdRec(1).SrvPntCsCode = ""
            LSet UBBeacRdRec(1).SrvPntCsCodeNm = ""
            LSet UBBeacRdRec(1).SrvPntLatitude = ""
            LSet UBBeacRdRec(1).SrvPntLongitude = ""
            LSet UBBeacRdRec(1).SrvPntTimeZone = ""
'Meter
            LSet UBBeacRdRec(1).MtrID = OurInfo$  'QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrIDNO)
            LSet UBBeacRdRec(1).MtrSerNo = QPTrim$(MtNumb$)
            LSet UBBeacRdRec(1).MtrManufact = ""
            LSet UBBeacRdRec(1).MtrModel = ""
            LSet UBBeacRdRec(1).MtrSize = ""
            LSet UBBeacRdRec(1).MtrSizeUnit = ""
            LSet UBBeacRdRec(1).MtrNote = ""
            LSet UBBeacRdRec(1).MtrContFlow = ""
            LSet UBBeacRdRec(1).MtrRegNumb = ""
            LSet UBBeacRdRec(1).MtrRegUOM = "GAL"
            LSet UBBeacRdRec(1).MtrRegReso = QPTrim(UBCustRec(1).USERCODE1)
            LSet UBBeacRdRec(1).MtrInstDate = ""
            LSet UBBeacRdRec(1).MtrInstStrRead = ""
            LSet UBBeacRdRec(1).MtrRemDate = ""
            LSet UBBeacRdRec(1).MtrRemRead = ""
'Service Agreement
            LSet UBBeacRdRec(1).SAStartDate = Num2Date$(UBCustRec(1).LocMeters(MtrCnt).PastDate)
            LSet UBBeacRdRec(1).SAEndDate = ""
'Endpoint Config
            LSet UBBeacRdRec(1).EndPointSerNo = QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrIDNO)  'OurInfo$ '
            LSet UBBeacRdRec(1).EndPointType = QPTrim(UBCustRec(1).USERCODE2) ' "1"  'Will be the UserCode1 field
            LSet UBBeacRdRec(1).EndPointInsDate = ""
            LSet UBBeacRdRec(1).EndPointRemDate = ""
''Manual/Mobile Readings
            LSet UBBeacRdRec(1).ManMobSequ = ""
            LSet UBBeacRdRec(1).ManMobAlertCode = ""
            LSet UBBeacRdRec(1).ManMobHighRead = ""
            LSet UBBeacRdRec(1).ManMobLowRead = ""
''Mics Info
            LSet UBBeacRdRec(1).MIUse1 = OurInfo$
            LSet UBBeacRdRec(1).MIUse2 = ""
'Write it to the file
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).AcctID) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).AcctFName) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).AcctLName) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).AcctFullName) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).AcctEMail) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).AcctPhone) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).AcctPhoneExt) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).AcctAddr1) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).AcctAddr2) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).AcctAddr3) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).AcctCity) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).AcctState) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).AcctZip) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).AcctCountry) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).AcctPerID) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).AcctStatus) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).AcctPortStat) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).AcctBLCyc) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).AcctPaperless) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).AcctAutoPay) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).AcctBillerAP) + V$;
'''Location
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).LocID) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).LocName) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).LocAddParity) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).LocAddr1) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).LocAddr2) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).LocAddr3) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).LocCity) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).LocState) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).LocZip) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).LocCountyName) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).LocCountry) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).LocLatitude) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).LocLongitude) + V$;
''Tags
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).TagLocBldType) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).TagLocBldNumb) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).TagLocSite) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).TagLocFunding) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).TagLocMainUse) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).TagLocWatType) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).TagLocArea) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).TagLocIrrArea) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).TagLocPopu) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).TagLocWFR) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).TagLocIrr) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).TagLocYearBuilt) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).TagLocPool) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).TagLocBathrooms) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).TagLocDistrict) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).TagLocDHSCode) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).TagLocParcNumb) + V$;
            
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).TagLocETJan) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).TagLocETFeb) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).TagLocETMar) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).TagLocETApr) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).TagLocETMay) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).TagLocETJun) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).TagLocETJul) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).TagLocETAug) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).TagLocETSep) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).TagLocETOct) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).TagLocETNov) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).TagLocETDec) + V$;
            
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).SrvPntID) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).SrvPntType) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).SrvPntCycle) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).SrvPntRoute) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).SrvPntCsCode) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).SrvPntCsCodeNm) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).SrvPntLatitude) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).SrvPntLongitude) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).SrvPntTimeZone) + V$;
''Meter
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).MtrID) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).MtrSerNo) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).MtrManufact) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).MtrModel) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).MtrSize) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).MtrSizeUnit) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).MtrNote) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).MtrContFlow) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).MtrRegNumb) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).MtrRegUOM) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).MtrRegReso) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).MtrInstDate) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).MtrInstStrRead) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).MtrRemDate) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).MtrRemRead) + V$;
''Service Agreement
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).SAStartDate) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).SAEndDate) + V$;
''Endpoint Config
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).EndPointSerNo) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).EndPointType) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).EndPointInsDate) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).EndPointRemDate) + V$;
'''Manual/Mobile Readings
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).ManMobSequ) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).ManMobAlertCode) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).ManMobHighRead) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).ManMobLowRead) + V$;
'''Mics Info
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).MIUse1) + V$;
            Print #UBBadRdFile, QPTrim(UBBeacRdRec(1).MIUse2)
'--------------------------------------------
        End Select
    End If

    MtrCnt = MtrCnt + 1
  Wend

Return
'Beacon END Output records

BadgerBeaconGet:
      NumberofRoutes = 0
      eFlag = False
      OverDate = Date2Num(fpReadDate)
      If OverDate > 0 Then
        UseOverDate = True
      Else
        UseOverDate = False
      End If
      
      GoSub MakeBeaconINTFFileName      'Get Badger File
      
      GoSub BeaconGetINRecCnt
      
      If BeInRecCnt < 2 Then  '1st record is a header
        Close
        MsgText(0) = "Beacon Import Operation."
        MsgText(1) = "ERROR!"
        MsgText(2) = ""
        MsgText(3) = "No Records found to Import!"
        MsgText(4) = ""
        MsgText(5) = ""
        Call GetOKorNot%(MsgText(), True, True)
        'MsgBox "No Records Found to Import.", vbOKOnly, "Procedure Ended"
        GoTo BadgerBeaconReadExit
        'no records error exit
      End If
      
      GoSub BadgerBeaconOpenCust      'Open Customer Data File
      
      UBBadRdFile = FreeFile
      Open FFName$ For Input As UBBadRdFile 'Len = UBBeaconRecLen

      FrmShowPctComp.Label1 = "Importing Reading Information."
      FrmShowPctComp.cmdCancel.Enabled = False
      FrmShowPctComp.Show '1, Parent

      DoneCnt& = 1              ' Initialize File Counter to 1
      Line Input #UBBadRdFile, BeaconTxtRec
      Do
    
        Line Input #UBBadRdFile, BeaconTxtRec
        
        GetBeaCustAndMtrID BeaconTxtRec, Prec&, MeterRecord
        
        If Prec& > 0 Then
          Get UBFile, Prec&, UBCustRec(1)
          GoSub BEEExtractRecord
        Else
          eFlag = True
        End If
        FrmShowPctComp.ShowPctComp DoneCnt&, NumBadRdRecs&
        DoneCnt& = DoneCnt& + 1
      Loop Until eof(UBBadRdFile) 'DoneCnt& > BeInRecCnt
    Close
    eFlag = False
    
    If eFlag Then
      frmMsgDialog.Caption = "WARNING!"
      frmMsgDialog.RetLabel = "-2"
      frmMsgDialog.Label(0).Caption = ""
      frmMsgDialog.Label(1).Caption = "Invalid reading records "
      frmMsgDialog.Label(2).Caption = "found in import file!"
      frmMsgDialog.Label(3).Caption = "Please call Southern Software for"
      frmMsgDialog.Label(4).Caption = "additional Information."
      frmMsgDialog.Show vbModal
    Else
      frmDataUpdated.Show vbModal
    End If

Return

BEEExtractRecord:
'this extracts the reading & date

  UpdCnt& = UpdCnt& + 1

  If MeterRecord = 0 Then MeterRecord = 1
  
  GetBeaCurReadAndDate BeaconTxtRec, lngCurrRead, ReadingDate

  If ReadingDate < 0 Then
    eFlag = True
    GoTo BEEExtractErrorExit
  End If
  
  If UseOverDate Then
    ReadingDate = OverDate          'if they want to overide the read date
  End If

  If UBCustRec(1).LocMeters(MeterRecord).ReadFlag = "Y" Then
    UBCustRec(1).LocMeters(MeterRecord).CurRead = lngCurrRead
    UBCustRec(1).LocMeters(MeterRecord).CurDate = ReadingDate 'Date2Num(MeterReadDate$)
  Else
    UBCustRec(1).LocMeters(MeterRecord).PrevRead = UBCustRec(1).LocMeters(MeterRecord).CurRead
    UBCustRec(1).LocMeters(MeterRecord).PastDate = UBCustRec(1).LocMeters(MeterRecord).CurDate
    UBCustRec(1).LocMeters(MeterRecord).ReadFlag = "Y"
    UBCustRec(1).LocMeters(MeterRecord).CurDate = ReadingDate 'Date2Num(MeterReadDate$)
    UBCustRec(1).LocMeters(MeterRecord).CurRead = lngCurrRead
  End If
'  'now update the customers record...
  Put UBFile, Prec&, UBCustRec(1)
BEEExtractErrorExit:
Return

BadgerBeaconOpenCust:
  UBFile = FreeFile
  Open "UBCUST.DAT" For Random Shared As UBFile Len = UBCustRecLen
Return

MakeBeaconOUTTFFileName:
  PathWay$ = QPTrim$(PathWay$)
  If Len(PathWay$) > 0 Then
    If Right$(PathWay$, 1) <> "\" Then
      PathWay$ = PathWay$ + "\"
    End If
  End If
  FFName$ = PathWay$ + "CONNECT.IN3"
Return

MakeBeaconINTFFileName:
  PathWay$ = QPTrim$(PathWay$)
  If Len(PathWay$) > 0 Then
    If Right$(PathWay$, 1) <> "\" Then
      PathWay$ = PathWay$ + "\"
    End If
  End If
  FFName$ = PathWay$ + "CONNECT.OT3"
Return

BeaconGetINRecCnt:
  BeInRecCnt = 0
  UBBadRdFile = FreeFile
  If Exist(FFName$) Then
    Open FFName$ For Input Shared As UBBadRdFile
    If LOF(UBBadRdFile) > 0 Then
      While Not eof(UBBadRdFile)
        Line Input #UBBadRdFile, BeaconTxtRec
        BeInRecCnt = BeInRecCnt + 1
      Wend
    End If
    Close UBBadRdFile
 End If
Return

End Sub     'BEACON
'**********************************************

